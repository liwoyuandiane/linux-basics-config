#!/bin/bash
# =============================================================================
# collect_hw_info.sh — Linux 硬件配置信息导出工具
# =============================================================================
# 用途  : 采集整机硬件配置，导出 Excel XML Spreadsheet 2003 (.xls)，供离线
#         资产盘点 / 巡检台账直接复制；重复执行在同一文件追加新页签。
# 环境  : Kylin / CentOS / Ubuntu / Debian / Fedora 等主流 Linux。
#         无需联网、无需安装软件包；缺失命令自动跳过，不中断导出。
# 交付  : 单个 .xls，6 列紧凑表格（列宽整表共用），彩色样式 + 合并单元格。
#         章节：基本信息 / 系统版本 / CPU(多路) / 内存 / 硬盘(品牌·SN) /
#         SMART 通电·出厂 / 网卡 / GPU / [可选性能测试] / 分区与挂载。
# SMART : 通电时长、通电次数、健康、温度、SN、出厂装配日期。四级后端回退
#         （详见 --help-full 第五节）：
#           ① 环境变量 SMARTCTL_BIN / 同目录静态 smartctl.<架构>
#           ② 系统 PATH 里的 smartctl
#           ③ 内嵌 Python 原生 ioctl 探针（零安装零携带）
#           ④ 全部不可用留空
#         随包静态 smartctl 可完全不依赖 python3；读取 SMART 需 root。
# 代码结构（按出现顺序）：
#   日志/错误 → 帮助与参数解析 → 工具检测 → XML 表格构造 → SMART 解析与后端
#   → 各硬件采集函数 → 分区表 → -d 性能测试 → XLS 生成 → main → 执行入口
# 注释约定：'# ====' 分章节，'# ----' 分小节，'#' 说明设计意图与已知坑。
# =============================================================================

SCRIPT_VERSION="5.21.1"
SCRIPT_NAME=$(basename "$0")

# ---- 硬盘性能测试配置 ----
DISK_TEST_SIZE="10G"          # 默认测试文件大小（可在此修改）
DISK_DEVICES=()              # 测试磁盘设备列表（-d 可多次指定，开跑前去重保序）
DISK_DEVICE_SET=0             # 用户是否显式给了 -d（决定是否做前置块设备校验）
TEST_FILE_DIR="/tmp"          # 测试文件写入目录
TEST_FILE_DIR_SET=0           # 用户是否显式给了 -p（默认 /tmp 不做前置校验）
# help 文案引用的"默认值"必须是内置值：show_help 在参数解析循环内执行，
# 若直接展开 DISK_TEST_SIZE，-s 5G -h 会把用户本次取值冒充成默认值
DISK_TEST_SIZE_DEFAULT="$DISK_TEST_SIZE"
TEST_FILE_DIR_DEFAULT="$TEST_FILE_DIR"

# ---- 全局工具检测标志 ----
TOOL_DMIDECODE=0
TOOL_LSPCI=0
TOOL_LSBLK=0
TOOL_LSCPU=0
TOOL_IP=0
TOOL_FREE=0

# ---- 错误/警告计数 ----
ERROR_COUNT=0
WARN_COUNT=0

# =============================================================================
# 日志与错误处理
# =============================================================================
# 日志全部走 stderr，颜色按 stderr 是否 tty 判定：
# stdout 重定向/进管道时日志仍保色，stdout 是 tty 而 stderr 进日志文件时不污染
if [ -t 2 ]; then
    RED='\033[0;31m'; GREEN='\033[0;32m'; YELLOW='\033[0;33m'; NC='\033[0m'
else
    RED=''; GREEN=''; YELLOW=''; NC=''
fi
log_info()  { echo -e "${GREEN}[INFO]${NC} $*" >&2; }
log_warn()  { echo -e "${YELLOW}[WARN]${NC} $*" >&2; }
log_error() { echo -e "${RED}[ERROR]${NC} $*" >&2; }

record_error() { ERROR_COUNT=$((ERROR_COUNT + 1)); log_error "$*"; }
record_warn()  { WARN_COUNT=$((WARN_COUNT + 1)); log_warn "$*"; }

# 去首尾空白。不要用 xargs 干这事：xargs 会把单引号当特殊字符，
# 遇到 "NVIDIA-SMI ... couldn't ..." 这类文本会报 unmatched single quote。
trim() { sed 's/^[[:space:]]*//; s/[[:space:]]*$//'; }

summary_errors() {
    if [ "$ERROR_COUNT" -gt 0 ] || [ "$WARN_COUNT" -gt 0 ]; then
        echo ""
        echo "=============================================="
        echo "  执行结果汇总"
        echo "=============================================="
        # 汇总行不加 [ERROR]/[WARN] 前缀，保证日志里带前缀的行数 === 计数
        [ "$ERROR_COUNT" -gt 0 ] && echo "  共发生 ${ERROR_COUNT} 个错误"
        [ "$WARN_COUNT" -gt 0 ]  && echo "  共发生 ${WARN_COUNT} 个警告"
    fi
}

# 检查文件是否可写
check_file_writable() {
    local file_path="$1"
    local dir_path; dir_path=$(dirname "$file_path")

    if [ ! -d "$dir_path" ]; then
        record_error "输出目录不存在: ${dir_path}"
        return 1
    fi
    # 共享目录下预埋符号链接会让 root 顺着链覆盖任意文件，一律拒收
    if [ -h "$file_path" ]; then
        record_error "输出路径是符号链接，拒绝写入: ${file_path}"
        return 1
    fi
    if [ -f "$file_path" ] && [ ! -w "$file_path" ]; then
        record_error "文件不可写: ${file_path}"
        return 1
    fi
    # mktemp 以 O_EXCL 创建测试文件，固定名 + touch 存在被预埋 symlink 劫持的窗口
    local test_file
    test_file=$(mktemp "${dir_path}/.wtest_XXXXXX" 2>/dev/null) || {
        record_error "无法在目录 ${dir_path} 创建文件（权限不足）"
        return 1
    }
    rm -f "$test_file" 2>/dev/null
    return 0
}

# =============================================================================
# 帮助信息与参数解析
# =============================================================================
# 简版帮助：只放"怎么调用"。GNU 编码规范要求 --help 输出简短用法后退出，
# clig.dev 也建议示例放在最前、帮助控制在几十行内；完整文档内置于 --help-full。
show_help() {
    cat << EOF
用法: ${SCRIPT_NAME} [选项] [输出目录]

单文件离线采集 Linux 硬件配置，导出 Excel XML (.xls)；重复执行追加页签。

示例:
  sudo ./${SCRIPT_NAME}                          # 当前目录生成 collect_hw_info.xls
  sudo ./${SCRIPT_NAME} /mnt/usb                 # 指定输出目录
  sudo ./${SCRIPT_NAME} -d /dev/sda -s 5G        # 附带硬盘性能测试
  sudo ./${SCRIPT_NAME} -o /tmp/机房A-01.xls     # 指定输出文件名

选项:
  -h, --help                    显示此帮助
  -H, --help-full               显示完整说明（离线部署 / SMART / 环境变量 / 排错）
  -v, --version                 显示版本号
  -d, --disk <设备>             待测硬盘，可多次指定（自动去重），如 -d /dev/sda -d /dev/sdb
  -s, --size <大小>             测试文件大小，默认 ${DISK_TEST_SIZE_DEFAULT}（K/M/G，大小写均可）
  -p, --test-path <目录>        测试文件写入目录，默认 ${TEST_FILE_DIR_DEFAULT}
  -o, --output <文件>           指定输出 .xls 文件（默认 <输出目录>/collect_hw_info.xls）
  --                            结束选项解析，其后仅作输出目录

输出:  <输出目录>/collect_hw_info.xls（-o 优先）
       参数错误 exit 2；完整说明用 --help-full（内置，单文件离线可查）
EOF
    exit 0
}

# 完整文档：离线环境常只拷一个 .sh，不能依赖外部 md，故内置一份。
show_help_full() {
    cat << EOF
${SCRIPT_NAME} v${SCRIPT_VERSION} — 完整说明
用法: ${SCRIPT_NAME} [选项] [输出目录]

一、输出与页签
  collect_hw_info.xls 为 Excel XML Spreadsheet，重复执行会追加新页签。
  页签名为 <3位序号>-<设备标识SN>，例如 001-ZFA1BFP9。整机 DMI 序列号读不到
  （Kylin/飞腾机器常回 "Not Set"）时自动改用第一块硬盘 SN，再退回主机名。
  序号前缀必须保留：Excel 不允许同名页签。
  -o 可直接指定输出文件名与路径，父目录不存在会自动创建。

二、报表章节
  一、基本信息        / 二、系统版本信息 / 三、CPU（含多路判定与双路插槽明细）
  四、内存（品牌、模组型号、SN、速率、Rank）
  五、硬盘信息（品牌·系列 / 型号 / SN 独立列，末尾「采集说明」写明 SN 与品牌来源）
  五（补充）SMART 通电：通电小时 / 次数 / 健康 / 温度 / 出厂装配日期；
             首部写明「SMART 后端程序」= 来源级别 + smartctl 版本 + 路径，
             版本 < 7.4 时另给「后端能力」行（不支持 FARM 出厂日志）
  六、网卡（已过滤 sit、ip6tnl、veth 等虚拟接口）
  七、GPU / 七（补充）驱动与软件版本 / 八、硬盘分区与挂载
  带 -d 时：八=硬盘详细信息与性能测试，九=分区与挂载（SMART 数值在
  「五（补充）」已有表格，详细段只给摘要 + 固件/扇区计数等独有项）
  分区表为紧凑 6 列：分区(树形 NAME) / 大小 / 文件系统 / 挂载点 /
  用量(已用/可用/率) / 标签·UUID，下方汇总压成两行。
  列宽为整表共用，不足 6 列的表格末格自动跨列，长文本不会被挤成多行。

三、驱动与软件版本
  网卡：ethtool -i 取 driver / version / firmware-version；无 ethtool 时读
  /sys/class/net/<iface>/device/driver 链接，再用 modinfo -F version 或
  /sys/module/<mod>/version 补版本。同驱动同版本的多个网口合并成一行。
  其他 PCI：按 /sys/bus/pci/devices/*/class 分出显示 / 存储 / USB 控制器，
  不依赖 lspci -k（CentOS 6 老 pciutils 可能没有 -k）。内核内置驱动常不声明
  version 字段，此时显示「未提供」，不是采集失败。
  软件版本：目前只有 sshd。注意 sshd -V 仅较新 OpenSSH 支持，老系统
  （CentOS 6 的 5.3p1、Kylin V10 的 8.2p1）会报 illegal option，脚本依次
  回退「服务端二进制版本串(grep -a)」→ ssh -V（同包客户端）→ rpm/dpkg 查询；
  「来源」列写明实际用了哪一级。服务端二进制串比客户端版本号更准。
  要加 nginx / openssl 等软件版本，在 collect_driver_versions() 里照
  sshd 那行继续 add_row 即可，表格结构不用改。

四、硬盘 SN 与出厂日期
  SN 按准确度依次尝试：smartctl -i → lsblk -o SERIAL → udevadm → sysfs
  (device/serial、VPD 0x80)。读不到显示「未提供」，以盘体标签为准；
  不用 WWN 冒充 SN。来源写在「五、硬盘信息」的「采集说明」一行。
  ATA / NVMe / SMART 标准【没有】生产日期字段：脚本尝试 ① Seagate FARM
  日志的 Assembly Date（需 smartctl 7.4+，且仅部分希捷盘支持，USB 桥可能
  不透传）；② 少数 SCSI/SAS 盘 smartctl -i 的 Manufactured 行。都拿不到
  显示「未提供」，可凭 SN 到厂商官网或盘体标签查询。装配日期是年周精度、
  不等于首次使用日期；FARM 通电小时会与 SMART 通电小时交叉核对，不一致时
  提示疑似计数器被重置。

五、离线 / 免安装（SMART 通电信息）
  不需要在目标机联网或安装软件包。当前部署只需三个文件：
    collect_hw_info.sh  smartctl.aarch64  smartctl.amd64
  四级后端回退（优先级）：
    1. 环境变量 SMARTCTL_BIN / 同目录静态二进制 smartctl.<架构>
       （含 bin/<架构>/smartctl、./smartctl，按 uname -m 自动选）
    2. 系统 PATH 里的 smartctl（例如 Kylin 自带的 7.1）
    3. 内嵌 Python 原生 ioctl 探针（仅依赖自带 python3，零安装零携带）
    4. 全部不可用时通电信息留空，其余硬件信息照常导出
  每个候选先做 --version 启动预检，架构不符或 glibc 过旧自动跳过并试下一个。
  实际命中的来源级别与版本会写进日志和 Excel，便于多机批量核对；
  smartctl < 7.4 不支持 FARM 日志，报表会明确提示而非静默留空。

  静态链接不看目标机 glibc，也不需要 python3；两个架构的静态文件即可通吃
  Kylin 桌面/服务器(aarch64) 与 CentOS 服务器(x86_64)。按 uname -m 依次探测
  smartctl.<架构>、bin/<架构>/smartctl、smartctl/<架构>/smartctl 与同目录裸名
  ./smartctl（均需 chmod +x，文件名不强制含架构关键字）；读取 SMART 需 root。
  纯 Bash/sh 无法发 ioctl，所以"零依赖"只能是脚本+解释器，或脚本+已编译 ELF。

六、环境变量
    SMARTCTL_BIN=/path/to/smartctl   指定随包 smartctl 路径
    SMART_PY=/usr/bin/python3        指定 python3 解释器路径
    SMART_FORCE_NATIVE=1             跳过所有 smartctl 候选，强制用内嵌 ioctl 探针
                                    （仅用于对照测试"机器上完全没装 smartctl"）

七、注意
  读取 SMART 需要 root（或 CAP_SYS_RAWIO），与是否安装工具无关。
  虚拟机磁盘、云块存储（如 Oracle BlockVolume）、部分 USB 桥接、硬件 RAID
  卡后的盘不提供物理通电计数器：显示「未获取到」属预期行为，脚本会识别
  虚拟环境并在 Excel 写明原因，且不计入警告；请在宿主机或物理机采集。
  健康状态显示 OK 仅为设备自评，不代表有完整属性表。物理机读不到才计警告。

八、命令行约定
  -h 只给简短用法，-H 给本说明，-v 只打印版本号。
  参数错误打印一行可操作提示并 exit 2（便于批量脚本判断成败），不会静默成功。
  日志口径：[WARN]/[ERROR] 行数与汇总计数一致；候选跳过、可选工具缺失等
  提示用 [INFO]，不计入警告；汇总行本身不带前缀。
  -o 优先于位置参数；输出父目录不存在时自动创建。
  -- 之后的参数一律按输出目录处理；输出目录重复指定会报错。
  -s 接受 数字+K/M/G（大小写均可），如 10G / 512m。
  -d 时章节「八」为性能测试、「九」为分区与挂载；不带 -d 时分区表即「八」。

九、示例
  ${SCRIPT_NAME}                                # 当前目录生成
  ${SCRIPT_NAME} -o /tmp/机房A-01.xls           # 指定输出文件名
  ${SCRIPT_NAME} -d /dev/sda -s 5G /tmp         # 5G 测试并输出到 /tmp
  ${SCRIPT_NAME} -H                             # 显示本说明
EOF
    exit 0
}

OUTPUT_DIR="."
OUTPUT_DIR_SET=0            # 位置参数只能给一次
OUTPUT_FILE_OVERRIDE=""     # -o 指定的输出文件（优先于 输出目录）
OUTPUT_FILE_OVERRIDE_SET=0  # -o 也只能指定一次，与位置参数口径一致
ONLY_POSITIONAL=0           # 遇到 -- 后，剩余参数一律按输出目录处理

# CLI 约定：参数错误要给出可操作的提示，并以非零码退出（便于批量脚本判断成败），
# 而不是打印整页帮助后 exit 0。
usage_error() {
    echo "错误: $*" >&2
    echo "用 ${SCRIPT_NAME} --help 查看用法，--help-full 查看完整说明" >&2
    exit 2
}

# 带值选项统一校验：缺值、或下一个参数本身是选项，都算误用。
# 不写成 "[ -z "$2" ] || [[ "$2" == -* ]] && usage_error"：依赖 ||/&& 左结合
# 虽能跑，但可读性差且 shellcheck 会报 SC2015；显式 if 一目了然。
need_arg() {
    local opt="$1" val="$2" hint="$3"
    if [ -z "$val" ] || [[ "$val" == -* ]]; then
        usage_error "${opt} 需要${hint}"
    fi
}

while [ $# -gt 0 ]; do
    if [ "$ONLY_POSITIONAL" -eq 1 ]; then
        [ "$OUTPUT_DIR_SET" -eq 1 ] && usage_error "输出目录只能指定一次（已指定: ${OUTPUT_DIR}）"
        [ -n "$1" ] || usage_error "输出目录不能为空字符串"
        OUTPUT_DIR="$1"; OUTPUT_DIR_SET=1; shift
        continue
    fi
    case "$1" in
        -h|--help) show_help ;;
        -H|--help-full) show_help_full ;;
        -v|--version) echo "${SCRIPT_NAME} v${SCRIPT_VERSION}"; exit 0 ;;
        --) ONLY_POSITIONAL=1; shift ;;
        -o|--output)
            need_arg "-o/--output" "$2" "指定输出 .xls 文件路径"
            [ "$OUTPUT_FILE_OVERRIDE_SET" -eq 1 ] \
                && usage_error "-o/--output 只能指定一次（已指定: ${OUTPUT_FILE_OVERRIDE}）"
            OUTPUT_FILE_OVERRIDE="$2"; OUTPUT_FILE_OVERRIDE_SET=1; shift 2 ;;
        -d|--disk)
            need_arg "-d/--disk" "$2" "指定磁盘设备路径"
            DISK_DEVICES+=("$2"); DISK_DEVICE_SET=1; shift 2 ;;
        -s|--size)
            need_arg "-s/--size" "$2" "指定测试文件大小"
            # 单位统一成大写，下游 parse_size_mb / fio / 报表文案口径一致
            DISK_TEST_SIZE=$(echo "$2" | sed 's/k$/K/; s/m$/M/; s/g$/G/')
            echo "$DISK_TEST_SIZE" | grep -qE '^[1-9][0-9]*[KMG]$' \
                || usage_error "无效的大小格式 '$2'，支持 数字+K/M/G (如 10G, 512M)；数字须为正整数（拒绝 0 与前导零）"
            shift 2 ;;
        -p|--test-path)
            need_arg "-p/--test-path" "$2" "指定测试文件目录"
            TEST_FILE_DIR="$2"; TEST_FILE_DIR_SET=1; shift 2 ;;
        -*) usage_error "未知参数 '$1'" ;;
        *)
            [ "$OUTPUT_DIR_SET" -eq 1 ] && usage_error "输出目录只能指定一次（已指定: ${OUTPUT_DIR}）"
            [ -n "$1" ] || usage_error "输出目录不能为空字符串"
            OUTPUT_DIR="$1"
            OUTPUT_DIR_SET=1; shift ;;
    esac
done

# -d/-p 真实性校验放在开跑前：性能测试参数错误应立即 exit 2，
# 而不是采集完全部硬件、跑完测试后才在报表里记一笔（批量脚本会误判成败）
if [ "$DISK_DEVICE_SET" -eq 1 ]; then
    # O(n²) 循环去重（保序）：兼容 bash 3.2，不依赖关联数组；
    # -d /dev/sda -d /dev/sda 只测一次
    _uniq_disks=()
    for _d in "${DISK_DEVICES[@]}"; do
        _dup=0
        for _u in "${_uniq_disks[@]}"; do
            [ "$_u" = "$_d" ] && { _dup=1; break; }
        done
        [ "$_dup" -eq 0 ] && _uniq_disks+=("$_d")
    done
    DISK_DEVICES=("${_uniq_disks[@]}")
    unset _uniq_disks _d _u _dup
    # 打印实际待测清单：批量核对"我写什么就测什么"，去重结果可观测
    log_info "待测硬盘(去重后): ${DISK_DEVICES[*]}"
    for _d in "${DISK_DEVICES[@]}"; do
        [ -b "$_d" ] || usage_error "-d 指定的不是块设备: ${_d}"
    done
    unset _d
fi
if [ "$TEST_FILE_DIR_SET" -eq 1 ]; then
    if [ ! -d "$TEST_FILE_DIR" ] || [ ! -w "$TEST_FILE_DIR" ]; then
        usage_error "-p 指定的目录不存在或不可写: ${TEST_FILE_DIR}"
    fi
fi

# -o 优先；未指定时用 <输出目录>/collect_hw_info.xls
if [ -n "$OUTPUT_FILE_OVERRIDE" ]; then
    OUTPUT_FILE="$OUTPUT_FILE_OVERRIDE"
else
    OUTPUT_FILE="${OUTPUT_DIR}/collect_hw_info.xls"
fi
# 常见 CLI 约定：目标父目录不存在时自动创建；创建失败（权限/路径非法）
# 立即按参数错误 exit 2，不能吞掉后等采集完才暴露
OUT_DIR_ACTUAL=$(dirname "$OUTPUT_FILE")
if [ ! -d "$OUT_DIR_ACTUAL" ]; then
    mkdir -p "$OUT_DIR_ACTUAL" 2>/dev/null \
        || usage_error "无法创建输出目录: ${OUT_DIR_ACTUAL}（权限不足或路径非法）"
fi

# =============================================================================
# 工具检测
# =============================================================================
check_tools() {
    command -v dmidecode &>/dev/null && TOOL_DMIDECODE=1
    command -v lspci     &>/dev/null && TOOL_LSPCI=1
    command -v lsblk     &>/dev/null && TOOL_LSBLK=1
    command -v lscpu     &>/dev/null && TOOL_LSCPU=1
    command -v ip        &>/dev/null && TOOL_IP=1
    command -v free      &>/dev/null && TOOL_FREE=1
    log_info "工具检测完成"
}

# =============================================================================
# XML 工具函数
# =============================================================================
xml_esc() {
    [ -z "$1" ] && echo "" && return
    # 删控制字符必须用八进制 \000-\037：GNU tr 不支持 \xHH 十六进制语法，
    # '\x00-\x1F' 会被解析成字符范围 0-x（0x30..0x78），把数字、大写字母、
    # a-w 和冒号一并删光（中文 UTF-8 字节与 y/z 幸存，单元格值只剩残骸）
    printf '%s\n' "$1" | sed 's/&/\&amp;/g; s/</\&lt;/g; s/>/\&gt;/g; s/"/\&quot;/g' | tr -d '\000-\037'
}

get_system_sn() {
    local sn=""
    if [ "$TOOL_DMIDECODE" -eq 1 ]; then
        sn=$(dmidecode -s system-serial-number 2>/dev/null | head -1)
    fi
    [ -z "$sn" ] && sn=$(cat /sys/class/dmi/id/product_serial 2>/dev/null | tr -d '\n\r')
    [ -z "$sn" ] && sn="Not Set"
    echo "$sn"
}

# =============================================================================
# 设备标识 SN（用于标题和页签名）
# 国产整机（飞腾/鲲鹏 + Kylin）的 dmidecode 常回 "Not Set" / 占位串，
# 但硬盘 SN 是真实可读的。整机 SN 无效时就用第一块物理盘的 SN，
# 这样页签名不再是 001-Not_Set，批量盘点时能直接看出是哪台机器。
# =============================================================================
DEVICE_SN=""
DEVICE_SN_SRC=""

first_physical_disk() {
    local n=""
    if [ "${TOOL_LSBLK:-0}" -eq 1 ]; then
        n=$(lsblk -d -n -o NAME 2>/dev/null \
            | grep -vE '^(loop|ram|zram|dm-|md|sr|fd|nbd|pmem)' | head -1)
    else
        local p bn
        for p in /sys/block/*; do
            [ -e "$p" ] || continue
            bn=$(basename "$p")
            case "$bn" in loop*|ram*|zram*|dm-*|md*|sr*|fd*|nbd*|pmem*) continue ;; esac
            n="$bn"; break
        done
    fi
    echo "$n" | trim
}

resolve_device_sn() {
    local sn; sn=$(get_system_sn)
    local src="整机 DMI"
    case "$sn" in
        ""|"Not Set"|"Not set"|"None"|"N/A"|"n/a"|"-"|"0"|"0000000000"|\
        "To be filled by O.E.M."|"To Be Filled By O.E.M."|\
        "System Serial Number"|"Chassis Serial Number"|"Not Specified"|\
        "Default string"|"Unknown"|"unknown")
            local name dsn=""
            name=$(first_physical_disk)
            if [ -n "$name" ]; then
                dsn=$(get_disk_serial "$name" | sed -n '1p')
                if [ -n "$dsn" ]; then
                    sn="$dsn"; src="第一块硬盘 SN（${name}）"
                fi
            fi
            # 虚拟机 / 云盘连硬盘 SN 都没有时，最后用主机名兜底，
            # 不留 "Not Set" 这种没有识别度的页签名。
            if [ -z "$dsn" ]; then
                local hn; hn=$(hostname 2>/dev/null | trim)
                if [ -n "$hn" ]; then sn="$hn"; src="主机名"; fi
            fi
            ;;
    esac
    DEVICE_SN="$sn"
    DEVICE_SN_SRC="$src"
}

# =============================================================================
# 数据行构建
# =============================================================================
DATA_ROWS=""

reset_rows() { DATA_ROWS=""; }

_str_cell() {
    local style="$1" value="$2" merge="$3"
    local s="" m=""
    [ -n "$style" ] && s=" ss:StyleID=\"${style}\""
    [ -n "$merge" ] && m=" ${merge}"
    echo "      <Cell${s}${m}><Data ss:Type=\"String\">${value}</Data></Cell>"
}

# 列宽是整张工作表共用的（118/150/100/128/150/205），没法给某节单独设。
# 不足 6 列的表，最后一格自动 MergeAcross 到剩余列，长文本就不会被挤在
# 100~150 宽的窄列里折成好几行 —— 网卡型号、出厂说明、分区汇总都受益。
TOTAL_COLS=6
_cell_list() {
    local style="$1"; shift
    local n=$# i=1 c merge="" out=""
    for c in "$@"; do
        merge=""
        if [ "$i" -eq "$n" ] && [ "$n" -lt "$TOTAL_COLS" ]; then
            merge="ss:MergeAcross=\"$((TOTAL_COLS - n))\""
        fi
        out="${out}
      $(_str_cell "$style" "$(xml_esc "$c")" "$merge")"
        i=$((i + 1))
    done
    printf '%s' "$out"
}

add_title() {
    local text; text=$(xml_esc "$1")
    DATA_ROWS="${DATA_ROWS}
    <Row>
      $(_str_cell "title" "$text" "ss:MergeAcross=\"5\"")
    </Row>"
}

add_blank() {
    DATA_ROWS="${DATA_ROWS}
    <Row/>"
}

add_section() {
    local text; text=$(xml_esc "$1")
    DATA_ROWS="${DATA_ROWS}
    <Row>
      $(_str_cell "section" "$text" "ss:MergeAcross=\"5\"")
    </Row>"
}

add_row() {
    # 两列 kv 行：值横跨第 2~6 列（约 733 点），系统版本、控制器等长文本一行放得下
    DATA_ROWS="${DATA_ROWS}
    <Row>
      $(_str_cell "label" "$(xml_esc "$1")")
      $(_str_cell "value" "$(xml_esc "$2")" "ss:MergeAcross=\"4\"")
    </Row>"
}

add_subheader() {
    DATA_ROWS="${DATA_ROWS}
    <Row>$(_cell_list "label" "$@")
    </Row>"
}

add_table_row() {
    DATA_ROWS="${DATA_ROWS}
    <Row>$(_cell_list "value" "$@")
    </Row>"
}

# 「五、硬盘信息」曾用的 4 列版本，保留以兼容旧调用
add_subheader4() {
    DATA_ROWS="${DATA_ROWS}
    <Row>$(_cell_list "label" "$@")
    </Row>"
}

add_table_row4() {
    DATA_ROWS="${DATA_ROWS}
    <Row>$(_cell_list "value" "$@")
    </Row>"
}

# 通用 N 列表格：最后一格自动补足到 6 列；正好 6 列时不合并。
add_hdrN() {
    DATA_ROWS="${DATA_ROWS}
    <Row>$(_cell_list "label" "$@")
    </Row>"
}

add_rowN() {
    DATA_ROWS="${DATA_ROWS}
    <Row>$(_cell_list "value" "$@")
    </Row>"
}

add_footer() {
    local text; text=$(xml_esc "$1")
    DATA_ROWS="${DATA_ROWS}
    <Row>
      $(_str_cell "" "$text" "ss:MergeAcross=\"5\"")
    </Row>"
}

# =============================================================================
# SMART 属性提取辅助函数
# =============================================================================
smart_get_attr() {
    local data="$1" attr_name="$2"
    # smartctl 7.x 的 RAW_VALUE 后面常跟 "(Min/Max ...)" 或 "(0 20 0 0 0)"，
    # 直接取 $NF 会拿到 "0)" 之类；先截掉括号注释再取最后一列。
    echo "$data" | grep -i "$attr_name" | sed 's/(.*//' | awk '{print $NF}' | tail -1
}

# 从 smartctl -A 输出中提取指定 SMART 值
# 兼容 ATA 属性表（Power_On_Hours / Power_Cycle_Count）
#   与 NVMe 文本格式（Power On Hours: 1,234 / Power Cycles: 56）
smart_extract_value() {
    local data="$1" pattern="$2" val=""
    # 形式 1: "Key: value"（NVMe / smartctl 文本输出）
    val=$(echo "$data" | grep -iE "^[[:space:]]*${pattern}[[:space:]]*:" \
          | head -1 | cut -d: -f2- | trim)
    # 形式 2: ATA 属性表 "  9 Power_On_Hours  ...  12345"
    # 截掉 RAW_VALUE 后的括号注释（Temperature_Celsius 常见 "28 (0 20 0 0 0)"）
    [ -z "$val" ] && val=$(echo "$data" \
          | grep -iE "^[[:space:]]*[0-9]+[[:space:]]+${pattern}[[:space:]]" \
          | sed 's/(.*//' | awk '{print $NF}' | tail -1)
    # 形式 3: 宽松兜底
    [ -z "$val" ] && val=$(echo "$data" | grep -iE "${pattern}" \
          | sed 's/(.*//' | awk '{print $NF}' | tail -1)
    # 去掉千分位逗号，便于 Excel 中排序/计算
    val=${val//,/}
    echo "$val"
}

# smartctl 7.4+ 对部分希捷企业盘把 Power_On_Hours 输出成
# "55h+11m+01.231s"；后面再拼 " 小时" 就成了 "55h+11m+01.231s 小时"。
# 统一收敛成纯小时数，Excel 里可直接排序/比较。
smart_normalize_hours() {
    local v="$1" n
    case "$v" in
        *[hH+]*)
            n=$(echo "$v" | grep -oE '^[0-9]+' | head -1)
            [ -n "$n" ] && { echo "$n"; return 0; }
            ;;
    esac
    echo "$v"
}

smart_get_health() {
    local disk="$1"
    "${SMARTCTL_BIN:-smartctl}" -H "$disk" 2>/dev/null \
        | grep -iE 'SMART overall-health|SMART Health Status' \
        | sed 's/.*:[[:space:]]*//' | head -1 | trim
}

smart_get_temp() {
    local data="$1" temp=""
    # NVMe: "Temperature: 32 Celsius"
    temp=$(echo "$data" | grep -iE '^[[:space:]]*Temperature:' \
           | head -1 | sed 's/.*:[[:space:]]*//' | trim)
    # ATA: "Temperature_Celsius" 属性行
    if [ -z "$temp" ]; then
        temp=$(smart_extract_value "$data" 'Temperature[_ ]Celsius')
        [ -n "$temp" ] && temp="${temp}°C"
    fi
    echo "$temp"
}

# =============================================================================
# 硬盘生产日期 / 装配日期
#
# 重要事实：ATA IDENTIFY、NVMe Identify、SMART 属性表都【没有】标准的
# "生产日期"字段。能拿到的只有厂商私有来源：
#   ① Seagate FARM 日志的 "Assembly Date (YYWW)"（smartctl -l farm，需 7.4+）
#   ② 少数 SCSI/SAS 设备 smartctl -i 输出的 "Manufactured:" 行
#   ③ 序列号里厂商自己编的日期码（规则各家不同，只作线索，不自动推断）
# 其余情况只能靠盘体标签、厂商工具或官网保修查询。
# =============================================================================

# 把 FARM 的 YYWW 原始值换成 "YYYY 年第 WW 周"。
# 希捷该字段两个字节各是反序 BCD：如 4254 -> 24 年 / 45 周；1283 -> 21 年 / 38 周。
# 年份世纪按当前年份回退，避免 2099 这种离谱结果。
smart_farm_date_fmt() {
    local raw="$1" yy ww year week y cy desc
    raw=$(echo "$raw" | grep -oE '[0-9]+' | head -1)
    [ "${#raw}" -lt 4 ] && { echo ""; return 1; }
    yy=${raw:0:2}; ww=${raw:2:2}
    year="${yy:1:1}${yy:0:1}"
    week="${ww:1:1}${ww:0:1}"
    case "$year" in ''|*[!0-9]*) echo ""; return 1 ;; esac
    case "$week" in ''|*[!0-9]*) echo ""; return 1 ;; esac
    y=$((2000 + 10#$year))
    cy=$(date +%Y 2>/dev/null); [ -z "$cy" ] && cy=2026
    [ "$y" -gt "$cy" ] && y=$((y - 100))
    w=$((10#$week))
    if [ "$w" -lt 1 ] || [ "$w" -gt 53 ]; then
        echo ""; return 1
    fi
    desc="${y} 年第 ${w} 周（原始值 ${raw}）"
    echo "$desc"
}

# 返回三行：日期 | 来源 | 补充说明（用换行分隔，由调用方拆分）
smart_get_disk_mfg() {
    local dev="$1" backend="$2" date="" src="" note=""

    if [ "$backend" != "smartctl" ] || [ -z "${SMARTCTL_BIN:-}" ]; then
        if [ "$backend" = "none" ]; then
            printf '%s\n%s\n%s\n' "未提供" "无可用 SMART 后端" \
                "请放静态 smartctl 到脚本同目录（或安装 python3），才能读取 FARM / 厂商日期信息"
        else
            printf '%s\n%s\n%s\n' "未提供" "内嵌 ioctl 探针" \
                "FARM / 厂商日期码需要 smartctl 7.4+；请放静态 smartctl 到脚本同目录"
        fi
        return 0
    fi

    local si model farm raw farm_hours smart_hours v
    si=$("$SMARTCTL_BIN" -i "$dev" 2>/dev/null)
    model=$(echo "$si" | grep -iE '^(Device Model|Model Number|Product):' \
            | head -1 | cut -d: -f2- | trim)

    # ① 少数 SCSI/SAS 设备会在 -i 里直接给 Manufactured / Date of Manufacture
    v=$(echo "$si" | grep -iE '^(Manufactured|Manufacture Date|Date of Manufacture|Mfg Date)[[:space:]]*:' \
        | head -1 | cut -d: -f2- | trim)
    if [ -n "$v" ]; then
        date="$v"; src="smartctl -i（厂商提供）"
    fi

    # ② Seagate FARM：只对疑似希捷盘尝试，避免每块盘多一次无谓调用
    if [ -z "$date" ] && echo "$model $si" | grep -qiE 'seagate|ST[0-9]{4}[A-Z]'; then
        farm=$("$SMARTCTL_BIN" -l farm "$dev" 2>/dev/null)
        raw=$(echo "$farm" | grep -iE '^[[:space:]]*Assembly Date' \
              | head -1 | cut -d: -f2- | trim)
        if [ -n "$raw" ]; then
            v=$(smart_farm_date_fmt "$raw")
            if [ -n "$v" ]; then
                date="$v"
                src="Seagate FARM（smartctl -l farm）"
                farm_hours=$(echo "$farm" | grep -iE '^[[:space:]]*Power on Hours:' \
                             | head -1 | cut -d: -f2- | trim)
            else
                date="格式未识别（原始值 ${raw}）"
                src="Seagate FARM（smartctl -l farm）"
            fi
        fi
    fi

    # 型号与 SN 已在「五、硬盘信息」单独成列，这里不再重复，只写来源与查询建议
    note=""
    if [ -z "$date" ]; then
        date="未提供"
        src="无标准字段"
        note="凭 SN 查厂商官网或盘体标签（型号/SN 见「五、硬盘信息」）"
        if [ "${SMARTCTL_FARM_OK:-0}" -eq 0 ] && [ "$backend" = "smartctl" ]; then
            note="${note}；当前后端不支持 FARM，换希捷盘也读不到装配日期（见上方「后端能力」行）"
        fi
    else
        note="装配日期不等于购买或首次使用日期，以盘体标签为准"
    fi
    # FARM 通电小时与 SMART 通电小时交叉核对（识别被清零的翻新盘）
    if [ -n "${farm_hours:-}" ]; then
        smart_hours=$(smart_extract_value "$("$SMARTCTL_BIN" -A "$dev" 2>/dev/null)" 'Power[_ ]On[_ ]Hours')
        smart_hours=$(smart_normalize_hours "$smart_hours")
        if [ -n "$smart_hours" ] && [ "$smart_hours" != "$farm_hours" ]; then
            note="${note}；⚠️ FARM 通电 ${farm_hours}h ≠ SMART 通电 ${smart_hours}h，疑似计数器被重置，建议复核"
        else
            note="${note}；FARM 通电 ${farm_hours}h（与 SMART 一致）"
        fi
    fi
    printf '%s\n%s\n%s\n' "$date" "$src" "$note"
}

# =============================================================================
# SMART 后端：随包静态 smartctl → 系统 smartctl → 内嵌原生 ioctl 探针
# 设计目标：离线批量分发时，目标机无需联网、无需安装任何软件包
# =============================================================================
SMART_BACKEND="none"                    # smartctl | native | none
SMARTCTL_BIN="${SMARTCTL_BIN:-}"        # 可由环境变量指定随包二进制
SMART_PY="${SMART_PY:-}"                # 可由环境变量指定 python3
SMART_PROBE_PY=""
SMART_PROBE_IS_TMP=0
SMARTCTL_VERSION=""             # 实际采用的 smartctl 版本（如 7.1 / 7.5）
SMARTCTL_SOURCE=""              # 来源级别：环境变量 / 随包静态 / 系统自带 / 内嵌探针
SMARTCTL_FARM_OK=0              # 版本是否 >= 7.4（支持 -l farm 读出厂日期）
SMART_ENV_IS_VM=0               # 是否虚拟机 / 云主机 / 虚拟块设备
SMART_ENV_DESC=""               # 运行环境描述
SMART_ENV_REASON=""             # 判定依据
SMART_NATIVE_FAIL_REASON=""     # 内嵌探针不可用的真实原因（无 python3 / 自检失败 / 写不了临时文件）
SCRIPT_DIR=$(cd "$(dirname "$0")" 2>/dev/null && pwd)
[ -z "$SCRIPT_DIR" ] && SCRIPT_DIR="."

cleanup_smart_runtime() {
    if [ "$SMART_PROBE_IS_TMP" -eq 1 ] && [ -n "$SMART_PROBE_PY" ] && [ -f "$SMART_PROBE_PY" ]; then
        rm -f "$SMART_PROBE_PY" 2>/dev/null
    fi
    # 锁文件只是 flock 的载体，进程退出即解锁；不删会永久残留在输出目录
    # （U盘交付场景下尤其显眼），必须由持有者自己清理
    if [ -n "${LOCK_FILE:-}" ]; then
        exec 200>&- 2>/dev/null
        rm -f "$LOCK_FILE" 2>/dev/null
    fi
}
trap cleanup_smart_runtime EXIT INT TERM

# 内嵌探针源码（单文件分发时用；与随包 smart_probe.py 内容完全一致）
write_embedded_probe() {
    local out="$1"
    cat > "$out" << 'SMART_PROBE_EOF'
#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
smart_probe.py —— 零依赖硬盘 SMART 探针（免安装 / 离线可用）

无需 smartmontools、无需编译、无需联网，仅用 Python 3 标准库：
  * SATA / SAS / 多数 USB 桥：SG_IO + ATA PASS-THROUGH(12/16) 读 SMART READ DATA
  * NVMe：NVME_IOCTL_ADMIN_CMD + Get Log Page (LID=0x02 SMART/Health)

运行要求：root（或 CAP_SYS_RAWIO）；设备需能透传 ATA/NVMe 命令。
用法：python3 smart_probe.py /dev/sda [/dev/nvme0n1 ...]
      python3 smart_probe.py --self-test   # 合成数据自检，不访问真实磁盘
输出：每行一条记录，字段用 US(\\x1f) 分隔，共 10 列：
      设备 | 后端 | 通电小时 | 通电次数 | 健康 | 温度 | 备注 | 重映射 | 待处理 | 不可修复
      读不到的字段留空；失败不抛异常、不影响调用方。
"""
import ctypes
import fcntl
import os
import sys

US = "\x1f"

# ---------------------------------------------------------------------------
# SCSI Generic / ATA PASS-THROUGH
# ---------------------------------------------------------------------------
SG_IO = 0x2285
SG_DXFER_FROM_DEV = -3


class SgIoHdr(ctypes.Structure):
    _fields_ = [
        ("interface_id", ctypes.c_int),
        ("dxfer_direction", ctypes.c_int),
        ("cmd_len", ctypes.c_ubyte),
        ("mx_sb_len", ctypes.c_ubyte),
        ("iovec_count", ctypes.c_ushort),
        ("dxfer_len", ctypes.c_uint),
        ("dxferp", ctypes.c_void_p),
        ("cmdp", ctypes.c_void_p),
        ("sbp", ctypes.c_void_p),
        ("timeout", ctypes.c_uint),
        ("flags", ctypes.c_uint),
        ("pack_id", ctypes.c_int),
        ("usr_ptr", ctypes.c_void_p),
        ("status", ctypes.c_ubyte),
        ("masked_status", ctypes.c_ubyte),
        ("msg_status", ctypes.c_ubyte),
        ("sb_len_wr", ctypes.c_ubyte),
        ("host_status", ctypes.c_ubyte),
        ("driver_status", ctypes.c_ubyte),
        ("resid", ctypes.c_int),
        ("duration", ctypes.c_uint),
        ("info", ctypes.c_uint),
    ]


# ATA SMART READ DATA: features=0xD0, count=1, LBA mid=0x4F, LBA high=0xC2, cmd=0xB0
ATA_CDBS = (
    # SAT ATA PASS-THROUGH(12)，Linux libata 上最常用
    bytes((0xA1, 0x08, 0x0E, 0xD0, 0x01, 0x00, 0x4F, 0xC2, 0x00, 0xB0, 0x00, 0x00)),
    # DEVICE=0xA0（部分 HBA / USB 桥要求 device 寄存器为 LBA 模式）
    bytes((0xA1, 0x08, 0x0E, 0xD0, 0x01, 0x00, 0x4F, 0xC2, 0xA0, 0xB0, 0x00, 0x00)),
    # 部分桥接芯片要求 count=0
    bytes((0xA1, 0x08, 0x0E, 0xD0, 0x00, 0x00, 0x4F, 0xC2, 0x00, 0xB0, 0x00, 0x00)),
    # ATA PASS-THROUGH(16) 兜底
    bytes((0x85, 0x08, 0x0E, 0xD0, 0x00, 0x01, 0x00, 0x00,
           0x00, 0x4F, 0x00, 0xC2, 0x00, 0x00, 0xB0, 0x00)),
)


def parse_ata(buf):
    """解析 512 字节 SMART READ DATA，返回 {属性ID: RAW_VALUE}"""
    attrs = {}
    for i in range(30):
        off = 2 + i * 12
        e = buf[off:off + 12]
        if len(e) < 12:
            break
        aid = e[0]
        if aid == 0:
            continue
        attrs[aid] = int.from_bytes(e[5:11], "little")
    return attrs


# 设备根本不支持该 ioctl 时立即放弃，不再逐个试 CDB 变体，
# 避免批量场景下每块盘白等若干秒。
FATAL_ERRNOS = {1, 13, 19, 22, 25}  # EPERM EACCES ENODEV EINVAL ENOTTY


def probe_ata(dev):
    """返回 (attrs, note)；失败时 attrs 为空 dict"""
    fd = None
    last_err = ""
    try:
        fd = os.open(dev, os.O_RDONLY | getattr(os, "O_NONBLOCK", 0))
    except OSError as ex:
        return {}, "打开设备失败: %s" % ex.strerror

    try:
        for cdb in ATA_CDBS:
            try:
                data = ctypes.create_string_buffer(512)
                sense = ctypes.create_string_buffer(32)
                cdb_buf = ctypes.create_string_buffer(cdb, len(cdb))

                h = SgIoHdr()
                h.interface_id = ord("S")
                h.dxfer_direction = SG_DXFER_FROM_DEV
                h.cmd_len = len(cdb)
                h.mx_sb_len = 32
                h.dxfer_len = 512
                h.dxferp = ctypes.addressof(data)
                h.cmdp = ctypes.addressof(cdb_buf)
                h.sbp = ctypes.addressof(sense)
                h.timeout = 3000
                h.flags = 0
                h.pack_id = 0

                buf = ctypes.create_string_buffer(ctypes.sizeof(h))
                ctypes.memmove(buf, ctypes.byref(h), ctypes.sizeof(h))
                out = fcntl.ioctl(fd, SG_IO, bytearray(buf.raw))
                ho = SgIoHdr.from_buffer_copy(out)
                if ho.status != 0 or (ho.info & 0x1) or ho.driver_status != 0:
                    last_err = "SG_IO 返回异常 (status=0x%02x info=0x%x)" % (ho.status, ho.info)
                    continue

                attrs = parse_ata(data.raw[:512])
                if not any(k in attrs for k in (9, 12, 194, 190, 5, 197)):
                    last_err = "返回数据不是有效的 SMART 属性表"
                    continue
                return attrs, ""
            except OSError as ex:
                last_err = "ioctl SG_IO: %s" % (ex.strerror or ex)
                if ex.errno in FATAL_ERRNOS:
                    break
                continue
            except Exception as ex:  # 结构体/对齐差异等，不致命
                last_err = "SG_IO 解析异常: %s" % ex
                continue
    finally:
        try:
            os.close(fd)
        except OSError:
            pass

    return {}, last_err or "设备未返回 SMART 数据"


# ---------------------------------------------------------------------------
# NVMe Get Log Page (SMART / Health, LID 0x02)
# ---------------------------------------------------------------------------
class NvmeAdminCmd(ctypes.Structure):
    _fields_ = [
        ("opcode", ctypes.c_uint8),
        ("flags", ctypes.c_uint8),
        ("rsvd1", ctypes.c_uint16),
        ("nsid", ctypes.c_uint32),
        ("cdw2", ctypes.c_uint32),
        ("cdw3", ctypes.c_uint32),
        ("metadata", ctypes.c_uint64),
        ("addr", ctypes.c_uint64),
        ("metadata_len", ctypes.c_uint32),
        ("data_len", ctypes.c_uint32),
        ("cdw10", ctypes.c_uint32),
        ("cdw11", ctypes.c_uint32),
        ("cdw12", ctypes.c_uint32),
        ("cdw13", ctypes.c_uint32),
        ("cdw14", ctypes.c_uint32),
        ("cdw15", ctypes.c_uint32),
        ("timeout_ms", ctypes.c_uint32),
        ("result", ctypes.c_uint32),
    ]


def _iowr(magic, nr, size):
    return (3 << 30) | (size << 16) | (ord(magic) << 8) | nr


NVME_IOCTL_ADMIN_CMD = _iowr("N", 0x41, ctypes.sizeof(NvmeAdminCmd))


def parse_nvme(log):
    """解析 512 字节 SMART/Health Log，返回 dict"""
    return {
        "critical": log[0],
        "temp_k": int.from_bytes(log[1:3], "little"),
        "power_cycles": int.from_bytes(log[112:120], "little"),
        "power_on_hours": int.from_bytes(log[128:136], "little"),
        "unsafe_shutdowns": int.from_bytes(log[144:152], "little"),
    }


def probe_nvme(dev, name):
    # 命名空间节点 /dev/nvme0n1 与控制器节点 /dev/nvme0 都试一遍
    import re
    cands = [dev]
    ctrl = os.path.join(os.path.dirname(dev) or "/dev",
                        re.sub(r"n\d+$", "", name))
    if ctrl != dev and os.path.exists(ctrl):
        cands.append(ctrl)

    nsids = [0xFFFFFFFF, 0]
    try:
        p = "/sys/block/%s/nsid" % name
        if os.path.isfile(p):
            nsids.insert(0, int(open(p).read().strip()))
    except Exception:
        pass

    last_err = ""
    for path in cands:
        fd = None
        try:
            fd = os.open(path, os.O_RDONLY | getattr(os, "O_NONBLOCK", 0))
        except OSError as ex:
            last_err = "打开设备失败: %s" % ex.strerror
            continue
        try:
            for nsid in nsids:
                try:
                    data = ctypes.create_string_buffer(512)
                    c = NvmeAdminCmd()
                    c.opcode = 0x02          # Get Log Page
                    c.nsid = nsid
                    c.addr = ctypes.addressof(data)
                    c.data_len = 512
                    c.cdw10 = 0x02 | (127 << 16)   # LID=0x02, NUMDL=127 (128 dwords)
                    c.cdw11 = 0
                    c.timeout_ms = 3000

                    raw = ctypes.string_at(ctypes.byref(c), ctypes.sizeof(c))
                    fcntl.ioctl(fd, NVME_IOCTL_ADMIN_CMD, bytearray(raw))
                    info = parse_nvme(data.raw[:512])
                    # 合理性校验，避免把未初始化的内存当作 SMART 日志
                    if not (250 <= info["temp_k"] <= 450):
                        last_err = "SMART/Health 日志校验失败"
                        continue
                    if not (0 <= info["power_on_hours"] <= 2000000):
                        last_err = "通电时长超出合理范围"
                        continue
                    if info["critical"] & 0xC0:
                        last_err = "Critical Warning 保留位非零"
                        continue
                    return info, ""
                except OSError as ex:
                    last_err = "ioctl NVME_IOCTL_ADMIN_CMD: %s" % (ex.strerror or ex)
                    if ex.errno in FATAL_ERRNOS:
                        break
                    continue
                except Exception as ex:
                    last_err = "NVMe 日志解析异常: %s" % ex
                    continue
        finally:
            try:
                os.close(fd)
            except OSError:
                pass

    return {}, last_err or "设备未返回 NVMe SMART/Health 日志"


# ---------------------------------------------------------------------------
# 统一输出
# ---------------------------------------------------------------------------
def clean(s):
    return str(s).replace(US, ";").replace("\n", " ").replace("\r", " ").replace("\t", " ")


def probe_one(dev):
    name = os.path.basename(dev)
    hours = cycles = health = temp = ""
    realloc = pending = offline = ""
    note = ""

    if name.startswith("nvme"):
        backend = "原生ioctl-NVMe"
        info, err = probe_nvme(dev, name)
        if info:
            hours = info["power_on_hours"]
            cycles = info["power_cycles"]
            tc = info["temp_k"] - 273
            temp = "%d°C" % tc if 0 < tc < 120 else ""
            health = "PASSED" if info["critical"] == 0 else "WARNING (critical=0x%02x)" % info["critical"]
        else:
            note = err
    else:
        backend = "原生ioctl-ATA/SAT"
        attrs, err = probe_ata(dev)
        if attrs:
            if 9 in attrs:
                hours = attrs[9]
            if 12 in attrs:
                cycles = attrs[12]
            if 194 in attrs:
                t = attrs[194] & 0xFF
                temp = "%d°C" % t if 0 < t < 120 else ""
            elif 190 in attrs:
                t = attrs[190] & 0xFF
                temp = "%d°C" % t if 0 < t < 120 else ""
            realloc = attrs.get(5, "")
            pending = attrs.get(197, "")
            offline = attrs.get(198, "")
            bad = [k for k in (5, 197, 198) if attrs.get(k, 0)]
            health = "属性正常" if not bad else "警告 (异常属性: %s)" % "/".join(str(x) for x in bad)
            if not hours and not cycles:
                note = "属性表中无 ID9/ID12（厂商自定义布局）"
        else:
            note = err

    fields = (dev, backend, hours, cycles, health, temp, note,
              realloc, pending, offline)
    return US.join(clean(f) for f in fields)


def self_test():
    """用合成数据自检解析器；不访问任何真实磁盘，可在任意离线机器上跑"""
    b = bytearray(512)

    def put(idx, aid, raw):
        off = 2 + idx * 12
        b[off] = aid
        b[off + 5:off + 11] = raw.to_bytes(6, "little")

    put(0, 9, 12345)
    put(1, 12, 234)
    put(2, 194, 32)
    put(3, 198, 3)
    a = parse_ata(bytes(b))
    assert a[9] == 12345 and a[12] == 234 and a[194] == 32 and a[198] == 3, a

    n = bytearray(512)
    n[1:3] = (304).to_bytes(2, "little")
    n[112:120] = (56).to_bytes(8, "little")
    n[128:136] = (1234).to_bytes(8, "little")
    i = parse_nvme(bytes(n))
    assert i["power_on_hours"] == 1234 and i["power_cycles"] == 56, i
    assert i["temp_k"] - 273 == 31, i

    # SgIoHdr 布局必须与 Linux sg_io_hdr.h 一致：偏移错位会让 ioctl 把
    # cmdp/dxferp 指向错误内存，轻则 EIO、重则踩内存
    if ctypes.sizeof(ctypes.c_void_p) == 8:
        expect_off = {"dxferp": 16, "cmdp": 24, "sbp": 32, "timeout": 40,
                      "status": 64}
        expect_size = 88
    else:
        expect_off = {"dxferp": 16, "cmdp": 20, "sbp": 24, "timeout": 28,
                      "status": 44}
        expect_size = 64
    for _fname, _off in expect_off.items():
        _got = getattr(SgIoHdr, _fname).offset
        assert _got == _off, "SgIoHdr.%s offset=%d 期望 %d" % (
            _fname, _got, _off)
    assert ctypes.sizeof(SgIoHdr) == expect_size, ctypes.sizeof(SgIoHdr)

    # CDB 关键字节：opcode、feature=SMART(0xD0)、LBA mid=0x4F、
    # LBA high=0xC2、command=0xB0——错一个字节设备直接返回 ILLEGAL REQUEST
    assert len(ATA_CDBS) == 4, len(ATA_CDBS)
    for _cdb in ATA_CDBS[:3]:
        assert len(_cdb) == 12 and _cdb[0] == 0xA1, _cdb
        assert (_cdb[3], _cdb[6], _cdb[7], _cdb[9]) == (0xD0, 0x4F, 0xC2, 0xB0), _cdb
    _c16 = ATA_CDBS[3]
    assert len(_c16) == 16 and _c16[0] == 0x85, _c16
    assert (_c16[3], _c16[9], _c16[11], _c16[14]) == (0xD0, 0x4F, 0xC2, 0xB0), _c16

    # SG_IO 冒烟（仅 Linux 生效）：对 /dev/null 发 ioctl，期望得到
    # FATAL_ERRNOS 里的「设备不支持」类错误；TypeError 说明参数/缓冲区
    # 类型错——不捕获，直接炸出来判失败
    try:
        _fd = os.open("/dev/null", os.O_RDONLY)
    except OSError:
        _fd = -1
    if _fd >= 0:
        try:
            _buf = ctypes.create_string_buffer(ctypes.sizeof(SgIoHdr))
            try:
                fcntl.ioctl(_fd, SG_IO, bytearray(_buf.raw))
                raise AssertionError("SG_IO(/dev/null) 意外成功")
            except OSError as _ex:
                assert _ex.errno in FATAL_ERRNOS, \
                    "SG_IO errno=%s 不在 FATAL_ERRNOS" % _ex.errno
        finally:
            os.close(_fd)

    # 对外输出契约：US 分隔 10 列，设备打不开也要干净返回、不抛异常
    _row = probe_one("/dev/this-device-does-not-exist-cd0")
    assert len(_row.split(US)) == 10, _row

    print("OK  ATA/NVMe 解析器自检通过 | NVME_IOCTL_ADMIN_CMD=0x%08X | "
          "nvme_admin_cmd=%d bytes | python=%s" %
          (NVME_IOCTL_ADMIN_CMD, ctypes.sizeof(NvmeAdminCmd),
           sys.version.split()[0]))
    return 0


def main():
    devs = sys.argv[1:]
    if devs and devs[0] == "--self-test":
        return self_test()
    if not devs:
        return 0
    for d in devs:
        try:
            print(probe_one(d))
        except Exception as ex:
            print(US.join((clean(d), "原生ioctl", "", "", "", "",
                           "探针异常: %s" % clean(ex), "", "", "")))
    return 0


if __name__ == "__main__":
    sys.exit(main())
SMART_PROBE_EOF
    chmod 600 "$out" 2>/dev/null
}

SMARTCTL_REJECT_REASON=""

# 当前机器架构 → 随包二进制的候选后缀（按优先级排列）
smart_arch_tags() {
    local m; m=$(uname -m 2>/dev/null)
    case "$m" in
        x86_64|amd64)            echo "amd64 x86_64 x64" ;;
        aarch64|arm64)           echo "arm64 aarch64" ;;
        armv7l|armv8l|armhf)     echo "armhf armv7 arm" ;;
        i386|i486|i586|i686)     echo "i386 x86" ;;
        loongarch64)             echo "loongarch64 loong64" ;;
        mips64el)                echo "mips64el" ;;
        ppc64le|ppc64el)         echo "ppc64el ppc64le" ;;
        riscv64)                 echo "riscv64" ;;
        *)                       echo "$m" ;;
    esac
}

# 启动预检：可执行文件不仅要存在，还得真能在本机跑起来。
# 拦下 "Exec format error"（架构不符）和 "GLIBC_x.xx not found"（glibc 过旧）。
smartctl_usable() {
    local bin="$1" out rc
    [ -x "$bin" ] || { SMARTCTL_REJECT_REASON="不可执行（请 chmod +x）"; return 1; }
    out=$("$bin" --version 2>&1)
    rc=$?
    if echo "$out" | grep -qiE 'GLIBC_|GLIBCXX_|Exec format error|cannot execute binary|error while loading shared libraries|Wrong ELF class'; then
        SMARTCTL_REJECT_REASON="架构或系统库不兼容: $(echo "$out" | head -1 | cut -c1-160)"
        return 1
    fi
    if [ "$rc" -eq 0 ] || echo "$out" | grep -qi 'smartctl'; then
        return 0
    fi
    SMARTCTL_REJECT_REASON="启动失败(退出码 ${rc}): $(echo "$out" | head -1 | cut -c1-160)"
    return 1
}

smart_use_binary() {
    local cand="$1" src="$2"
    SMARTCTL_BIN="$cand"
    SMART_BACKEND="smartctl"
    SMARTCTL_SOURCE="${src:-未标注}"
    local full; full=$("$cand" --version 2>/dev/null | head -1 | trim)
    SMARTCTL_VERSION=$(echo "$full" | grep -oE 'smartctl [0-9]+\.[0-9]+' | awk '{print $2}')
    [ -z "$SMARTCTL_VERSION" ] && SMARTCTL_VERSION=$(echo "$full" | grep -oE '[0-9]+\.[0-9]+' | head -1)
    # -l farm（希捷出厂装配日期）是 7.4 才加入的；7.1 等老版本调用会失败，
    # 报表里必须写明原因，否则以后换希捷盘会误以为"盘不支持"。
    SMARTCTL_FARM_OK=0
    [ -n "$SMARTCTL_VERSION" ] && ver_ge "$SMARTCTL_VERSION" "7.4" && SMARTCTL_FARM_OK=1
    return 0
}

# ③ 内嵌 ioctl 探针（独立成函数，便于 SMART_FORCE_NATIVE 强制走这条路做对照）
smart_try_native() {
    SMART_NATIVE_FAIL_REASON=""
    local probe_ok=0 py_found=0 c cand d
    if [ -f "${SCRIPT_DIR}/smart_probe.py" ]; then
        SMART_PROBE_PY="${SCRIPT_DIR}/smart_probe.py"
        probe_ok=1
    else
        for d in "${TMPDIR:-/tmp}" "${OUTPUT_DIR:-.}" "."; do
            [ -d "$d" ] && [ -w "$d" ] || continue
            # mktemp 以 O_EXCL 创建，防 /tmp 下预埋同名 symlink 劫持写入
            SMART_PROBE_PY=$(mktemp "${d}/.smart_probe_XXXXXXXX.py" 2>/dev/null) \
                || { SMART_PROBE_PY=""; continue; }
            if write_embedded_probe "$SMART_PROBE_PY" && [ -s "$SMART_PROBE_PY" ]; then
                SMART_PROBE_IS_TMP=1; probe_ok=1; break
            fi
            rm -f "$SMART_PROBE_PY"
            SMART_PROBE_PY=""
        done
        if [ "$probe_ok" -eq 0 ]; then
            SMART_BACKEND="none"
            SMART_NATIVE_FAIL_REASON="无法写入内嵌探针临时文件（检查 /tmp 与输出目录写权限）"
            return 1
        fi
    fi
    if [ -n "$SMART_PY" ] && [ -x "$SMART_PY" ]; then
        py_found=1
        if "$SMART_PY" "$SMART_PROBE_PY" --self-test >/dev/null 2>&1; then
            SMART_BACKEND="native"; return 0
        fi
    fi
    for c in python3 python; do
        if command -v "$c" &>/dev/null; then
            py_found=1
            cand=$(command -v "$c")
            if "$cand" "$SMART_PROBE_PY" --self-test >/dev/null 2>&1; then
                SMART_PY="$cand"; SMART_BACKEND="native"; return 0
            fi
        fi
    done
    SMART_BACKEND="none"
    # 「无解释器」和「有解释器但自检失败」处置方式不同，
    # 必须分别给结论，否则排错方向会被带偏（历史文案的教训）
    if [ "$py_found" -eq 0 ]; then
        SMART_NATIVE_FAIL_REASON="未找到可用的 python3 解释器，且无可用 smartctl"
    else
        SMART_NATIVE_FAIL_REASON="python3 存在但内嵌 ioctl 探针 --self-test 未通过（多为 Python 2 / ABI 不兼容），且无可用 smartctl"
    fi
    return 1
}

detect_smart_backend() {
    local tags t cand sysctl
    tags=$(smart_arch_tags)

    # 只用于对照测试：跳过所有 smartctl 候选，精确复现"机器上完全没装 smartctl"
    # 的场景，不必去改 /usr/sbin 或临时挪动系统文件。
    if [ "${SMART_FORCE_NATIVE:-0}" = "1" ]; then
        log_info "SMART_FORCE_NATIVE=1：跳过全部 smartctl 候选，强制使用内嵌 ioctl 探针"
        smart_try_native; return $?
    fi

    # ① 显式指定的二进制（环境变量 + 预解包文件）
    # 用数组收集候选：SCRIPT_DIR 或 SMARTCTL_BIN 含空格时，
    # `for cand in $cands` 的词分割会把路径拆碎丢弃
    local env_bin="$SMARTCTL_BIN" t
    local -a cands=()
    [ -n "$SMARTCTL_BIN" ] && cands+=("$SMARTCTL_BIN")
    for t in $tags; do
        cands+=("${SCRIPT_DIR}/smartctl.${t}")
        cands+=("${SCRIPT_DIR}/bin/${t}/smartctl")
        cands+=("${SCRIPT_DIR}/smartctl/${t}/smartctl")
    done
    cands+=("${SCRIPT_DIR}/smartctl")
    for cand in "${cands[@]}"; do
        [ -e "$cand" ] || continue
        local src="随包静态文件"
        [ -n "$env_bin" ] && [ "$cand" = "$env_bin" ] && src="环境变量 SMARTCTL_BIN"
        if smartctl_usable "$cand"; then
            smart_use_binary "$cand" "$src"; return 0
        fi
        log_info "跳过 ${cand}: ${SMARTCTL_REJECT_REASON}"
    done
    SMARTCTL_BIN=""

    # ② 系统 PATH 里的 smartctl
    if command -v smartctl &>/dev/null; then
        sysctl=$(command -v smartctl)
        if smartctl_usable "$sysctl"; then
            smart_use_binary "$sysctl" "系统自带"; return 0
        fi
        log_info "跳过系统 ${sysctl}: ${SMARTCTL_REJECT_REASON}"
    fi

    # ③ 回退：随包 / 内嵌 Python 原生 ioctl 探针
    smart_try_native; return $?
}

# =============================================================================
# 版本比较：smartctl 是否 >= 7.4（决定能否用 -l farm 读希捷出厂装配日期）
# =============================================================================
ver_ge() {  # a >= b ?
    if [ -z "$1" ] || [ -z "$2" ]; then return 1; fi
    if printf '%s\n' "$1" | sort -V >/dev/null 2>&1; then
        [ "$(printf '%s\n%s\n' "$1" "$2" | sort -V | head -1)" = "$2" ]
    else
        awk -v a="$1" -v b="$2" 'BEGIN{
            split(a,x,"."); split(b,y,".");
            for(i=1;i<=3;i++){
                if((x[i]+0)>(y[i]+0)) exit 0;
                if((x[i]+0)<(y[i]+0)) exit 1;
            }
            exit 0;
        }'
    fi
}

# =============================================================================
# 运行环境探测：虚拟磁盘 / 云块存储不提供 SMART 通电属性，这是预期行为，
# 不该记成警告，也不该提示"可能需要 root"误导用户。
# =============================================================================
smart_detect_env() {
    [ -n "${SMART_ENV_DETECTED:-}" ] && return 0
    SMART_ENV_IS_VM=0
    SMART_ENV_DESC="物理机（未检测到虚拟化特征）"
    SMART_ENV_REASON=""

    local vendor="" product="" virt="" disk_model="" blob="" ev="" hit="" p
    if [ "${TOOL_DMIDECODE:-0}" -eq 1 ]; then
        vendor=$(dmidecode -s system-manufacturer 2>/dev/null | head -1 | trim)
        product=$(dmidecode -s system-product-name 2>/dev/null | head -1 | trim)
    fi
    [ -z "$vendor" ]  && vendor=$(cat /sys/class/dmi/id/sys_vendor 2>/dev/null | head -1 | trim)
    [ -z "$product" ] && product=$(cat /sys/class/dmi/id/product_name 2>/dev/null | head -1 | trim)
    if command -v systemd-detect-virt &>/dev/null; then
        virt=$(systemd-detect-virt --vm 2>/dev/null | head -1 | trim)
    fi
    if [ "${TOOL_LSBLK:-0}" -eq 1 ]; then
        disk_model=$(lsblk -d -n -o VENDOR,MODEL 2>/dev/null \
                     | grep -viE '^[[:space:]]*$' | head -3 | tr '\n' ' ' | trim)
    fi
    blob="$vendor $product $virt $disk_model"

    # systemd-detect-virt 明确报告虚拟化类型时直接采信（kvm / qemu / xen / oracle ...）
    if [ -n "$virt" ] && [ "$virt" != "none" ]; then
        hit="systemd-detect-virt: ${virt}"
    fi

    if [ -z "$hit" ]; then
    for ev in "QEMU" "KVM" "OracleCloud" "Oracle Cloud" "VMware" "VirtualBox" \
              "Xen" "Parallels" "Bochs" "virtio" \
              "Virtio" "BlockVolume" "Virtual Machine" "Hyper-V" "Amazon EC2" \
              "Google Compute" "Alibaba Cloud" "OpenStack" "Proxmox" "libvirt" \
              "OpenTelekomCloud" "Huawei Cloud" "FusionCompute" "ZStack"; do
        case "$blob" in *"$ev"*) hit="$ev"; break ;; esac
    done
    fi
    if [ -z "$hit" ]; then
        for p in /sys/block/*/device; do
            [ -e "$p" ] || continue
            case "$p" in
                */loop*|*/ram*|*/zram*|*/dm-*|*/md*|*/pmem*|*/nbd*) continue ;;
            esac
            case "$(readlink -f "$p" 2>/dev/null)" in
                *virtio*) hit="virtio-blk"; break ;;
            esac
        done
    fi

    if [ -n "$hit" ]; then
        SMART_ENV_IS_VM=1
        SMART_ENV_DESC="虚拟机 / 云主机（特征: ${hit}）"
        SMART_ENV_REASON="制造商: ${vendor:-未知} | 产品: ${product:-未知} | 磁盘: ${disk_model:-未知} | 虚拟化: ${virt:-未识别}"
    fi
    SMART_ENV_DETECTED=1
}

# 统一 TSV（US 分隔，10 列）：
# 设备 | 后端 | 通电小时 | 通电次数 | 健康 | 温度 | 备注 | 重映射 | 待处理 | 不可修复
smart_query_smartctl() {
    local dev="$1" a hours cycles health temp note="" realloc pending offline raw_msg
    # 2>&1 才能拿到 "does not support SMART" / "Permission denied" 等诊断信息
    a=$("$SMARTCTL_BIN" -A "$dev" 2>&1)
    hours=$(smart_extract_value "$a" 'Power[_ ]On[_ ]Hours')
    hours=$(smart_normalize_hours "$hours")
    cycles=$(smart_extract_value "$a" 'Power[_ ]Cycle[s]?([_ ]Count)?')
    health=$(smart_get_health "$dev")
    temp=$(smart_get_temp "$a")
    realloc=$(smart_get_attr "$a" "Reallocated_Sector_Ct")
    pending=$(smart_get_attr "$a" "Current_Pending_Sector")
    offline=$(smart_get_attr "$a" "Offline_Uncorrectable")

    if [ -z "$hours" ] && [ -z "$cycles" ]; then
        if   echo "$a" | grep -qiE 'permission denied|operation not permitted'; then
            note="权限不足：请以 root 运行（或授予 CAP_SYS_RAWIO）"
        elif echo "$a" | grep -qiE 'does not support SMART|not supported|inappropriate ioctl|unsupported|SMART support is.*[Uu]navailable'; then
            note="设备不支持或未返回 SMART 属性表（虚拟盘 / 云盘 / USB 桥 / RAID 未透传时常见）"
        elif echo "$a" | grep -qiE 'device open failed|no such device|read failed|Input/output error'; then
            note="设备无法打开或读取：请确认设备路径与内核驱动"
        else
            note="设备未返回 ID 9 通电时长 / ID 12 通电次数属性"
        fi
        raw_msg=$(echo "$a" \
                  | grep -iE \
                        -e 'permission denied'  -e 'operation not permitted' \
                        -e 'does not support'   -e 'not supported' \
                        -e 'unsupported'        -e 'inappropriate ioctl' \
                        -e 'device open'        -e 'no such device' \
                        -e 'read failed'        -e 'SMART support is' \
                        -e 'error' -e 'warning' \
                  | head -2 | tr '\n' ';' | sed 's/;*$//' | cut -c1-140 | trim)
        [ -n "$raw_msg" ] && note="${note}｜smartctl: ${raw_msg}"
    fi
    printf '%s\037%s\037%s\037%s\037%s\037%s\037%s\037%s\037%s\037%s\n' \
        "$dev" "smartctl" "$hours" "$cycles" "$health" "$temp" "$note" \
        "$realloc" "$pending" "$offline"
}

smart_query() {
    local dev="$1"
    if [ "$SMART_BACKEND" = "smartctl" ]; then
        smart_query_smartctl "$dev"
    elif [ "$SMART_BACKEND" = "native" ]; then
        "$SMART_PY" "$SMART_PROBE_PY" "$dev" 2>/dev/null | head -1
    else
        printf '%s\037%s\037%s\037%s\037%s\037%s\037%s\037%s\037%s\037%s\n' \
            "$dev" "无可用后端" "" "" "" "" "${SMART_NATIVE_FAIL_REASON:-无可用 SMART 后端}" "" "" ""
    fi
}

# 解析一行 TSV 并写入全局变量，避免子 shell 丢失数据
smart_parse_line() {
    local line="$1"
    IFS=$'\037' read -r \
        Q_DEV Q_BACKEND Q_HOURS Q_CYCLES Q_HEALTH Q_TEMP Q_NOTE \
        Q_REALLOC Q_PENDING Q_OFFLINE <<< "$line"
}

# 把 SMART 结果写成一组明细行（-d 详细测试段复用）。
# 通电/健康/温度/出厂日期在「五（补充）」已有完整表格，这里压成摘要并指向该节，
# 避免 -d 段把同一组数字重复 8 行；只保留详细段独有的固件版本、关键扇区计数。
add_smart_detail_rows() {
    local disk="$1" line si sn fw
    line=$(smart_query "$disk")
    smart_parse_line "$line"

    if [ "$SMART_BACKEND" = "smartctl" ]; then
        si=$("$SMARTCTL_BIN" -i "$disk" 2>/dev/null)
        sn=$(echo "$si" | grep -i 'Serial Number'   | head -1 | cut -d: -f2- | trim)
        fw=$(echo "$si" | grep -i 'Firmware Version' | head -1 | cut -d: -f2- | trim)
        add_row "序列号 / 固件版本" "${sn:-未提供} / ${fw:-未提供}"
    fi

    local life="" extra=""
    [ -n "$Q_HOURS" ]  && life="${Q_HOURS} 小时"
    [ -n "$Q_CYCLES" ] && life="${life}${life:+ / }${Q_CYCLES} 次"
    [ -n "$Q_HEALTH" ] && extra="$Q_HEALTH"
    [ -n "$Q_TEMP" ]   && extra="${extra}${extra:+ / }${Q_TEMP}"
    add_row "通电与健康" "${life:-未获取到} / ${extra:-未获取到}（详见「五（补充）」）"

    local cnt=""
    [ -n "$Q_REALLOC" ] && cnt="重分配 ${Q_REALLOC}"
    [ -n "$Q_PENDING" ] && cnt="${cnt}${cnt:+ / }待处理 ${Q_PENDING}"
    [ -n "$Q_OFFLINE" ] && cnt="${cnt}${cnt:+ / }不可修复 ${Q_OFFLINE}"
    [ -n "$cnt" ] && add_row "关键扇区计数" "$cnt"
    [ -n "$Q_NOTE" ] && add_row "采集备注" "$Q_NOTE"

    local mfg md
    mfg=$(smart_get_disk_mfg "$disk" "$SMART_BACKEND")
    md=$(echo "$mfg" | sed -n '1p')
    add_row "出厂 / 装配日期" "${md:-未提供}（详见「五（补充）」）"

    smart_detect_env
    if [ "$SMART_ENV_IS_VM" -eq 1 ]; then
        add_row "运行环境" "$SMART_ENV_DESC"
        if [ -z "$Q_HOURS" ] && [ -z "$Q_CYCLES" ]; then
            add_row "通电信息说明" "虚拟磁盘 / 云块存储不提供物理通电计数器，属预期行为；如需该数据请在宿主机或物理机上采集"
        fi
    fi
    add_blank
}

# =============================================================================
# 数据收集
# =============================================================================
# 基本信息 / 系统版本：DMI、/etc/os-release、uname；主机名与采集时间放这里
collect_os_info() {
    log_info "正在收集系统信息..."
    local os_name="" os_hostname="" os_arch="" os_kernel=""
    os_hostname=$(hostname 2>/dev/null)
    [ -z "$os_hostname" ] && os_hostname="未知"
    os_arch=$(uname -m 2>/dev/null)
    [ -z "$os_arch" ] && os_arch="未知"
    os_kernel=$(uname -r 2>/dev/null)
    [ -z "$os_kernel" ] && os_kernel="未知"

    if [ -f /etc/os-release ]; then
        os_name=$(grep -E '^PRETTY_NAME=' /etc/os-release 2>/dev/null | cut -d'"' -f2)
        [ -z "$os_name" ] && os_name=$(grep -E '^PRETTY_NAME=' /etc/os-release 2>/dev/null | cut -d"'" -f2)
    fi
    [ -z "$os_name" ] && [ -f /etc/kylin-release ]  && os_name="$(cat /etc/kylin-release 2>/dev/null)"
    [ -z "$os_name" ] && [ -f /etc/centos-release ] && os_name="$(cat /etc/centos-release 2>/dev/null)"
    [ -z "$os_name" ] && [ -f /etc/redhat-release ]  && os_name="$(cat /etc/redhat-release 2>/dev/null)"
    [ -z "$os_name" ] && [ -f /etc/debian_version ]  && os_name="Debian $(cat /etc/debian_version 2>/dev/null)"
    [ -z "$os_name" ] && [ -f /etc/lsb-release ]     && os_name="$(grep '^DISTRIB_DESCRIPTION=' /etc/lsb-release 2>/dev/null | cut -d'"' -f2)"
    [ -z "$os_name" ] && os_name="未知"

    local now; now=$(date '+%Y-%m-%d %H:%M:%S')
    [ -z "$DEVICE_SN" ] && resolve_device_sn

    add_section "一、基本信息"
    add_row "主机名" "$os_hostname"
    add_row "采集时间" "$now"
    case "$DEVICE_SN_SRC" in
        "整机 DMI")   add_row "设备序列号" "$DEVICE_SN" ;;
        "主机名")     add_row "设备序列号（取主机名）" "$DEVICE_SN" ;;
        *)            add_row "设备序列号（取硬盘 SN）" "$DEVICE_SN" ;;
    esac
    add_blank

    add_section "二、系统版本信息"
    add_row "操作系统" "$os_name"
    add_row "系统架构" "$os_arch"
    add_row "内核版本" "$os_kernel"

    if [ -f /etc/.kyinfo ]; then
        local kv km
        kv=$(grep '^version=' /etc/.kyinfo 2>/dev/null | cut -d'=' -f2)
        km=$(grep '^milestone=' /etc/.kyinfo 2>/dev/null | cut -d'=' -f2)
        [ -n "$kv" ] && add_row "版本标识" "$kv"
        [ -n "$km" ] && add_row "系统版本" "$km"
    fi

    if [ "$TOOL_DMIDECODE" -eq 1 ]; then
        local sm sp
        sm=$(dmidecode -s system-manufacturer 2>/dev/null)
        sp=$(dmidecode -s system-product-name 2>/dev/null)
        [ -n "$sm" ] && add_row "系统制造商" "$sm"
        [ -n "$sp" ] && add_row "产品名称" "$sp"

        local bios_vendor bios_ver bios_date
        bios_vendor=$(dmidecode -s bios-vendor 2>/dev/null | head -1)
        bios_ver=$(dmidecode -s bios-version 2>/dev/null | head -1)
        bios_date=$(dmidecode -s bios-release-date 2>/dev/null | head -1)
        [ -n "$bios_vendor" ] && add_row "BIOS 厂商" "$bios_vendor"
        [ -n "$bios_ver" ]    && add_row "BIOS 版本" "$bios_ver"
        [ -n "$bios_date" ]   && add_row "BIOS 日期" "$bios_date"

        local board_mfr board_name
        board_mfr=$(dmidecode -s baseboard-manufacturer 2>/dev/null | head -1)
        board_name=$(dmidecode -s baseboard-product-name 2>/dev/null | head -1)
        [ -n "$board_mfr" ]  && add_row "主板厂商" "$board_mfr"
        [ -n "$board_name" ] && add_row "主板型号" "$board_name"
    fi
}

# 从 "Key: value" 文本里取值（行首匹配、忽略大小写），供 lscpu / smartctl -i 复用
kv_get() {
    # 用括号把调用方传入的交替式包起来，保证 ^[[:space:]]* 作用于整个模式；
    # 否则 "A|B" 会被解析成 "^...A" 或 "任意位置 B"，匹配到无关行。
    echo "$1" | grep -iE "^[[:space:]]*(${2})[[:space:]]*:" \
        | head -1 | cut -d: -f2- | trim
}

# 解析 dmidecode -t processor：按 Handle 分块，只保留已插 CPU 的插槽。
# 输出 TSV：插槽 | 厂商 | 型号/版本 | 核心数 | 线程数 | 最大频率
parse_dmi_cpu_sockets() {
    echo "$1" | awk '
    function sock_ok(s) {
        # 国产 BIOS 有时把 Socket Designation 填成 CPU 型号（如 "Phytium D2000/8 E8C"），
        # 不像插槽名就统一编号，避免第一列出现型号、双路机器两行同名。
        return (tolower(s) ~ /(cpu|socket|proc|^s[0-9]|^p[0-9]|^u[0-9])/)
    }
    function flush() {
        if (sock != "" && blk !~ /Status:.*[Uu]npopulated/ && blk !~ /Status:.*Empty/ && blk !~ /No Module Installed/) {
            if (!sock_ok(sock)) sock = "CPU" idx
            printf "%s\t%s\t%s\t%s\t%s\t%s\n", sock, mfr, ver, core, thr, spd
            idx++
        }
    }
    BEGIN { idx=0 }
    /^Handle .*DMI type 4/ { flush(); blk=""; sock=""; mfr=""; ver=""; core=""; thr=""; spd=""; next }
    {
        line=$0; blk = blk "\n" line
        k=line; sub(/^[[:space:]]*/,"",k); sub(/:.*$/,"",k)
        v=line; sub(/^[^:]*:[[:space:]]*/,"",v)
        if      (k=="Socket Designation") sock=v
        else if (k=="Manufacturer")       mfr=v
        else if (k=="Version")            ver=v
        else if (k=="Core Count")         core=v
        else if (k=="Thread Count")       thr=v
        else if (k=="Max Speed")          spd=v
    }
    END { flush() }
    ' | grep -v '^[[:space:]]*$'
}

# CPU：lscpu Socket(s) + dmidecode type 4 + /proc/cpuinfo 三源交叉判路数；
# ARM 上 physical id 常恒为 0，不拿它做差异报警；末尾附每颗插槽明细
collect_cpu_info() {
    log_info "正在收集 CPU 信息..."
    add_section "三、CPU 信息"

    local out=""; [ "$TOOL_LSCPU" -eq 1 ] && out=$(LC_ALL=C lscpu 2>/dev/null)
    local model arch mhz numa sockets cores_per threads_per cpus
    # ARM64 的 lscpu 同时有 "Model: 3"(CPU variant 编号) 和 "Model name: xxx"，
    # 且 Model 常排在前面。必须分步取，绝不能把两者放进同一个交替式里。
    model=$(kv_get "$out" 'Model name')
    [ -z "$model" ] && model=$(kv_get "$out" 'Model')
    # 部分国产 BIOS 把 lscpu 的 Model name 写成 "Phytium,D2000/8 E8C"，
    # 与 DMI 的 "Phytium D2000/8 E8C" 不一致，统一成空格
    model=$(echo "$model" | tr ',' ' ' | tr -s ' ' | trim)
    arch=$(kv_get "$out" 'Architecture')
    mhz=$(kv_get "$out" 'CPU max MHz')
    numa=$(kv_get "$out" 'NUMA node\(s\)')
    sockets=$(kv_get "$out" 'Socket\(s\)')
    cores_per=$(kv_get "$out" 'Core\(s\) per socket')
    threads_per=$(kv_get "$out" 'Thread\(s\) per core')
    cpus=$(kv_get "$out" 'CPU\(s\)')

    # 无 lscpu（CentOS 6 老机）或从 ARM 定制内核拿不到时，退回 /proc/cpuinfo
    if [ -z "$model" ]; then
        model=$(grep -m1 -iE '^(model name|Hardware|Processor)' /proc/cpuinfo 2>/dev/null \
                | cut -d: -f2- | trim)
    fi
    [ -z "$model" ] && model="未知"
    [ -z "$arch" ]  && arch=$(uname -m 2>/dev/null)
    [ -z "$cpus" ]  && cpus=$(grep -c '^processor' /proc/cpuinfo 2>/dev/null)
    [ -z "$cpus" ]  && cpus="未知"

    # ---- 多路（双 CPU）判定：三源交叉 ----
    local dmi_rows="" dmi_socks=0 phys_ids=0 sock_src=""
    if [ "$TOOL_DMIDECODE" -eq 1 ]; then
        dmi_rows=$(parse_dmi_cpu_sockets "$(dmidecode -t processor 2>/dev/null)")
        [ -n "$dmi_rows" ] && dmi_socks=$(echo "$dmi_rows" | wc -l | trim)
    fi
    if grep -q '^physical id' /proc/cpuinfo 2>/dev/null; then
        phys_ids=$(grep '^physical id' /proc/cpuinfo 2>/dev/null \
                   | awk '{print $NF}' | sort -u | wc -l | trim)
    fi
    if [[ "$sockets" =~ ^[0-9]+$ ]] && [ "$sockets" -gt 0 ]; then
        sock_src="lscpu Socket(s)"
    elif [ "$dmi_socks" -gt 0 ]; then
        sockets="$dmi_socks"; sock_src="dmidecode 已插 CPU 插槽数"
    elif [ "$phys_ids" -gt 0 ]; then
        sockets="$phys_ids"; sock_src="/proc/cpuinfo physical id 去重"
    else
        sockets=1; sock_src="未读到路数信息，按单路处理"
    fi

    local sock_desc="$sockets"
    case "$sockets" in
        1) sock_desc="1（单路）" ;;
        2) sock_desc="2（双路 / 双 CPU）" ;;
        4) sock_desc="4（四路）" ;;
        8) sock_desc="8（八路）" ;;
        *) [ "$sockets" -gt 2 ] && sock_desc="${sockets}（多路 / ${sockets} 颗 CPU）" ;;
    esac

    add_row "型号" "$model"
    # 「架构」已在「二、系统版本信息」给出，这里不再重复
    add_row "CPU 路数（物理颗数）" "$sock_desc"
    [ -n "$cores_per" ]  && add_row "每路物理核心数" "$cores_per"
    if [[ "$sockets" =~ ^[0-9]+$ ]] && [[ "$cores_per" =~ ^[0-9]+$ ]]; then
        add_row "物理核心总数" "$((sockets * cores_per)) 核（${sockets} 路 × ${cores_per} 核）"
    fi
    [ -n "$threads_per" ] && add_row "每核线程数" "$threads_per"
    add_row "总逻辑核心数" "$cpus"
    if [ -n "$threads_per" ] && [ "$threads_per" -gt 1 ] 2>/dev/null; then
        add_row "并行技术" "超线程 / SMT 已启用（每核 ${threads_per} 线程）"
    else
        add_row "并行技术" "每核 1 线程（未启用超线程 / SMT）"
    fi
    [ -n "$numa" ] && add_row "NUMA 节点数" "$numa"
    [ -n "$mhz" ]  && add_row "主频（最大）" "${mhz} MHz" || add_row "主频（最大）" "未知"

    # 交叉核对：三源不一致时明确提示，避免双路被误报成单路。
    # /proc/cpuinfo 的 physical id 只有 x86 可靠，ARM 上常恒为 0，不参与差异判断。
    local x86=0 xcheck=""
    case "${arch:-}" in x86_64|i686|i386|i586|amd64) x86=1 ;; esac
    [ "$dmi_socks" -gt 0 ] && [ "$dmi_socks" -ne "$sockets" ] && xcheck="${xcheck}dmidecode=${dmi_socks} "
    [ "$x86" -eq 1 ] && [ "$phys_ids" -gt 0 ] && [ "$phys_ids" -ne "$sockets" ] \
        && xcheck="${xcheck}physical_id=${phys_ids} "
    if [ -n "$xcheck" ]; then
        local xnote="依据 ${sock_src}=${sockets}；其他来源不一致：${xcheck}"
        xnote="${xnote}（ARM 平台常无 physical id，属正常；若确为双路请核对 BIOS）"
        add_row "路数交叉核对" "$xnote"
    else
        local phys_note=""
        [ "$x86" -eq 1 ] && [ "$phys_ids" -gt 0 ] && phys_note="；physical id 去重 ${phys_ids}"
        add_row "路数判定依据" "${sock_src}${dmi_socks:+；dmidecode 已插插槽 ${dmi_socks}}${phys_note}"
    fi

    # 每颗 CPU 明细：单路时型号/核心/线程/频率与上方汇总完全重复，省略；
    # 双路/多路才展开，便于确认两颗 CPU 是否同型号。
    if [ "$sockets" -gt 1 ] && [ -n "$dmi_rows" ]; then
        add_hdrN "CPU 插槽" "厂商" "型号 / 版本" "核心数" "线程数" "最大频率"
        local IFS_BAK="$IFS"
        while IFS=$'\t' read -r sl mf vr co th sp; do
            [ -z "$sl" ] && continue
            add_rowN "$sl" "${mf:-未知}" "${vr:-未知}" "${co:-未知}" "${th:-未知}" "${sp:-未知}"
        done <<< "$dmi_rows"
        IFS="$IFS_BAK"
    fi
}

# 内存品牌：优先 DMI Manufacturer；BIOS 填 Unknown 时按模组型号前缀轻推断，
# 明确标注「推断」，不把猜测当官方数据。JEDEC 厂商码也做几个常见映射。
mem_brand() {
    local manu="$1" part="$2" b=""
    case "$manu" in
        ""|"-"|"—"|"Unknown"|"unknown"|"UNKNOWN"|"Not Specified"|\
        "0x0000"|"0x00000000"|"To be filled by O.E.M."|"Default string") manu="" ;;
    esac
    [ "$part" = "—" ] && part=""
    if [ -n "$manu" ]; then
        case "$manu" in
            0x802C|0x2C|0x2c00)  manu="Micron 镁光（DMI 厂商码）" ;;
            0x80CE|0xCE|0xce00)  manu="Samsung 三星（DMI 厂商码）" ;;
            0x80AD|0xAD|0xad00)  manu="SK hynix 海力士（DMI 厂商码）" ;;
            0x8098|0x98)         manu="Kioxia / Toshiba（DMI 厂商码）" ;;
        esac
        echo "$manu"
        return
    fi
    case "$part" in
        KVR*|KHX*|HX[34]*|KF[34]*) b="Kingston 金士顿" ;;
        M39*|M47*|M32*|M37*)       b="Samsung 三星" ;;
        HMA*|HMT*|HMC*)            b="SK hynix 海力士" ;;
        MTA*|MTF*)                 b="Micron 镁光" ;;
        CT*)                        b="Crucial 英睿达" ;;
        CM[KTWZ]*|CMS*)            b="Corsair 海盗船" ;;
        F4-*|F3-*)                 b="G.Skill 芝奇" ;;
    esac
    if [ -n "$b" ]; then
        echo "${b}（按模组型号前缀推断）"
    else
        echo "未提供"
    fi
}

# 国产 BIOS 常把 SMBIOS 字段填成占位垃圾，直接进报表很难看，统一归一成 —
mem_clean() {
    local v="$1"
    case "$v" in
        ""|"-"|"Unknown"|"unknown"|"UNKNOWN"|"Not Specified"|"Not Set"|\
        "<BAD INDEX>"|"<BAD INDEX"|"BAD INDEX"|"N/A"|"None"|"none"|\
        "0x0000"|"0x00000000"|"To be filled by O.E.M."|"Default string")
            echo "—"; return ;;
    esac
    case "$v" in
        *[!0Ff]*) echo "$v" ;;      # 含非 0/F 字符 → 正常值
        *) echo "—" ;;              # 00000000 / FFFFFFFF → 占位
    esac
}

# 内存：dmidecode -t memory；品牌优先 DMI Manufacturer，缺失时按模组型号前缀
# 推断并明确标注；同规格多条合并成一行标数量，坏值(<BAD INDEX>/00000000)归一为 —
collect_memory_info() {
    log_info "正在收集内存信息..."
    add_section "四、内存信息"

    if [ "$TOOL_FREE" -eq 1 ]; then
        # LC_ALL=C 已挡住标签被 gettext 翻译（Mem:→内存：）；
        # free -h 要求 procps-ng ≥3.3，老系统 -h 直接报错，依次回退 -k 与 /proc/meminfo
        local tm; tm=$(LC_ALL=C free -h 2>/dev/null | awk '/^Mem:/{print $2; exit}')
        if [ -z "$tm" ]; then
            local kb; kb=$(LC_ALL=C free -k 2>/dev/null | awk '/^Mem:/{print $2; exit}')
            [ -z "$kb" ] && kb=$(awk '/^MemTotal:/{print $2}' /proc/meminfo 2>/dev/null)
            [ -n "$kb" ] && [ "$kb" -gt 0 ] 2>/dev/null && tm="$((kb / 1024 / 1024))G"
        fi
        add_row "总内存大小" "${tm:-未知}"
    else
        local mkb; mkb=$(grep 'MemTotal' /proc/meminfo 2>/dev/null | awk '{print $2}')
        [ -n "$mkb" ] && add_row "总内存大小" "$((mkb / 1024 / 1024))G" || add_row "总内存大小" "未知"
    fi

    local dimm_count=0
    local -a dloc=() dbrand=() dsize=() dtype=() dspeed=() dpart=() dsn=() drank=() dkey=()
    if [ "$TOOL_DMIDECODE" -eq 1 ]; then
        # 末尾补一个假的 "Memory Device" 头，让循环统一冲刷最后一根内存
        local slots; slots=$(dmidecode -t memory 2>/dev/null; printf '\nMemory Device\n')
        if [ -n "$slots" ]; then
            local in_slot=0 loc="" manu="" part="" sn="" size="" type="" speed="" conf="" rank=""
            while IFS= read -r line; do
                if echo "$line" | grep -q 'Memory Device$'; then
                    if [ "$in_slot" -eq 1 ] && [ -n "$size" ] \
                       && [ "$size" != "No Module Installed" ] && [ "$size" != "Unknown" ]; then
                        local idx=$(( ${#dloc[@]} + 1 ))
                        # 清洗占位垃圾（<BAD INDEX> / 00000000 / Unknown …）
                        local cmanu cpart csn crank ctype cspeed cconf
                        cmanu=$(mem_clean "$manu"); cpart=$(mem_clean "$part")
                        csn=$(mem_clean "$sn");     crank=$(mem_clean "$rank")
                        ctype=$(mem_clean "$type"); cspeed=$(mem_clean "$speed")
                        cconf=$(mem_clean "$conf")
                        dloc+=("${loc:-插槽${idx}}")
                        dbrand+=("$(mem_brand "$cmanu" "$cpart")")
                        # Speed=Unknown 而 Configured Memory Speed 有效时直接用后者，
                        # 别显示成 "Unknown / 配置 2133 MT/s"
                        local sp="$cspeed"
                        if [ "$cspeed" = "—" ] && [ "$cconf" != "—" ]; then
                            sp="$cconf"
                        elif [ "$cspeed" != "—" ] && [ "$cconf" != "—" ] && [ "$cconf" != "$cspeed" ]; then
                            sp="${cspeed} / 配置 ${cconf}"
                        fi
                        dsize+=("$size"); dtype+=("$ctype"); dspeed+=("$sp")
                        dpart+=("$cpart"); dsn+=("$csn"); drank+=("$crank")
                        # SN 不参与去重：同规格多条合并成一行，SN 全部列出
                        dkey+=("${loc}|${cmanu}|${cpart}|${size}|${ctype}|${cspeed}|${crank}")
                    fi
                    in_slot=1; loc=""; manu=""; part=""; sn=""; size=""
                    type=""; speed=""; conf=""; rank=""
                    continue
                fi
                [ "$in_slot" -eq 0 ] && continue
                local k; k=$(echo "$line" | sed 's/^[[:space:]]*//' | cut -d':' -f1 | trim)
                local v; v=$(echo "$line" | sed 's/^[[:space:]]*//' | cut -d':' -f2- | trim)
                case "$k" in
                    Locator)                   loc="$v" ;;
                    Manufacturer)              manu="$v" ;;
                    "Part Number")             part="$v" ;;
                    "Serial Number")           sn="$v" ;;
                    Size)                      size="$v" ;;
                    Type)                      type="$v" ;;
                    Speed)                     speed="$v" ;;
                    "Configured Memory Speed") conf="$v" ;;
                    Rank)                      rank="$v" ;;
                esac
            done <<< "$slots"
        fi
    fi

    dimm_count=${#dloc[@]}
    add_row "内存条数量" "$dimm_count"
    if [ "$dimm_count" -gt 0 ]; then
        local brands; brands=$(printf '%s\n' "${dbrand[@]}" | sort -u | grep -v '^$' \
                              | tr '\n' ',' | sed 's/,$//; s/,/, /g' | trim)
        local types;  types=$(printf '%s\n' "${dtype[@]}" | sort -u | grep -v '^$' \
                             | tr '\n' ',' | sed 's/,$//; s/,/, /g' | trim)
        local total_mb=0 n
        for n in "${dsize[@]}"; do
            # dmidecode 3.x 输出带单位（"16 GB"），2.x 老版本为 "8192 MB"；
            # 不换算单位会把 GB 当 MB，合计偏小 1024 倍
            local num unit
            num=$(echo "$n" | grep -oE '^[0-9]+' | head -1)
            [ -z "$num" ] && continue
            unit=$(echo "$n" | grep -oE '[KkMmGgTt][Bb]' | head -1 | tr '[:lower:]' '[:upper:]')
            case "$unit" in
                KB) total_mb=$((total_mb + (num + 1023) / 1024)) ;;
                GB) total_mb=$((total_mb + num * 1024)) ;;
                TB) total_mb=$((total_mb + num * 1024 * 1024)) ;;
                *)  total_mb=$((total_mb + num)) ;;  # MB 或无单位按 MB
            esac
        done
        add_row "内存品牌" "${brands:-未知}"
        printf '%s\n' "${dbrand[@]}" | grep -q '未提供' \
            && add_row "品牌说明" "BIOS 未填 Manufacturer、模组型号也不可读；以内存条标签为准"
        add_row "内存类型" "${types:-未知}"
        [ "$total_mb" -gt 0 ] && add_row "插槽容量合计" "$((total_mb / 1024)) GiB（按 DMI 累加 ${total_mb} MB）"

        # 完全相同的内存条折叠成一行标「数量」，不再出现两行一模一样的内容
        local -a oloc=() obrand=() osize=() otype=() ospeed=() opart=() osn=() orank=() ocnt=() okey=()
        local i j k=-1
        for i in "${!dloc[@]}"; do
            k=-1
            for j in "${!oloc[@]}"; do
                [ "${dkey[$i]}" = "${okey[$j]}" ] && { k=$j; break; }
            done
            if [ "$k" -ge 0 ]; then
                ocnt[$k]=$(( ${ocnt[$k]} + 1 ))
                # 只追加有效 SN；BIOS 全填 00000000 时不出现 "SN —, —"
                if [ "${dsn[$i]}" != "—" ]; then
                    if [ -z "${osn[$k]}" ] || [ "${osn[$k]}" = "—" ]; then
                        osn[$k]="${dsn[$i]}"
                    else
                        osn[$k]="${osn[$k]}, ${dsn[$i]}"
                    fi
                fi
            else
                oloc+=("${dloc[$i]}");   obrand+=("${dbrand[$i]}")
                osize+=("${dsize[$i]}"); otype+=("${dtype[$i]}")
                ospeed+=("${dspeed[$i]}"); opart+=("${dpart[$i]}")
                osn+=("${dsn[$i]}");     orank+=("${drank[$i]}")
                okey+=("${dkey[$i]}");   ocnt+=(1)
            fi
        done
        add_hdrN "插槽" "数量" "品牌 / 制造商" "容量" "类型 / 速率" "模组型号 / SN / Rank"
        for j in "${!oloc[@]}"; do
            local label="${oloc[$j]}" dup=0
            for i in "${!oloc[@]}"; do
                [ "$i" -ne "$j" ] && [ "${oloc[$i]}" = "${oloc[$j]}" ] && dup=$((dup + 1))
            done
            [ "$dup" -gt 0 ] && label="${label} (#$((j + 1)))"
            add_rowN "$label" "${ocnt[$j]}" "${obrand[$j]}" "${osize[$j]}" \
                     "${otype[$j]} / ${ospeed[$j]}" \
                     "${opart[$j]} / SN ${osn[$j]} / Rank ${orank[$j]}"
        done
    else
        add_row "内存插槽明细" "dmidecode 未提供内存条信息（虚拟机 / 非 root / ARM 平台无 SMBIOS 内存表）"
    fi
}

# =============================================================================
# 硬盘品牌 / 产品系列
# sysfs 的 device/vendor 常是 "ATA"（传输层标识，不是品牌），必须排除。
# 优先用 smartctl -i 的 Model Family（厂商系列），再按型号前缀推断。
# 推断结果一律带来源标注，不冒充固件官方字段。
# =============================================================================
disk_brand_from_model() {
    local u; u=$(echo "$1" | tr '[:lower:]' '[:upper:]')
    case "$u" in
        ST[0-9]*|SEAGATE*|EXOS*|IRONWOLF*|BARRACUDA*|SKYHAWK*|ARCHIVE*|CONSTELLATION*)
            echo "Seagate 希捷" ;;
        WD[0-9A-Z]*|WDC*|WESTERN*|ULTRASTAR*|RED*|BLUE*|BLACK*|PURPLE*|GOLD*|SN[0-9]*)
            echo "Western Digital 西部数据" ;;
        HGST*|HUS*|HUA*|HITACHI*|HTS*|IBM*)
            echo "HGST 日立（现属西数）" ;;
        TOSHIBA*|MQ0*|MD0*|DT01*|AL1*|MG0*|THNS*|Q300*|P300*|TR200*|XG*)
            echo "Toshiba 东芝" ;;
        KIOXIA*|KBG*|KXG*|EXCERIA*)
            echo "Kioxia 铠侠" ;;
        SAMSUNG*|MZ[VNTQ7]*|PM[89]*|SM[89]*|HD[0-9]*|SPINPOINT*)
            echo "Samsung 三星" ;;
        INTEL*|SSDSC*|SSDPE*|OPTANE*|DCS*|P4*|S4*)
            echo "Intel" ;;
        MICRON*|MTFD*|M600*|CRUCIAL*|CT[0-9]*|P[12345]*)
            echo "Micron 镁光 / Crucial 英睿达" ;;
        HYNIX*|SK*HYNIX*|SHGP*|HFS*|PC[34]*)
            echo "SK hynix 海力士" ;;
        KINGSTON*|SA400*|SUV*|SKC*|KC[236]*|A400*|A2000*|NV[123]*)
            echo "Kingston 金士顿" ;;
        SANDISK*|SD[0-9]*|ULTRA*|EXTREME*)
            echo "SanDisk 闪迪（现属西数）" ;;
        LITEON*|LITE-ON*|CV[138]*|PLEXTOR*|PX-*)
            echo "LiteOn 建兴 / Plextor 浦科特" ;;
        ADATA*|XPG*|ASX*|SU[678]*)
            echo "ADATA 威刚" ;;
        TRANSCEND*|TS[0-9]*)
            echo "Transcend 创见" ;;
        LENOVO*|SX[0-9]*|SL[0-9]*)
            echo "Lenovo 联想" ;;
        FORESEE*|LONGYSYS*|LONGSYS*|UNIC*|BIWIN*|GLOWAY*|KINGBANK*|NETAC*|GIGASTONE*|ZHITAI*|TIANJI*|MAXSUN*|COLORFUL*|DATO*|FANXIANG*)
            echo "国产模组厂（江波龙 / 佰维 / 致钛 等）" ;;
        *) echo "" ;;
    esac
}

# 输出一行 "品牌|系列|识别依据"
get_disk_brand() {
    local name="$1" model="$2" sysv="$3" dev="/dev/${1}"
    local family="" vendor="" brand="" how="" si=""
    if [ "${SMART_BACKEND:-none}" = "smartctl" ] && [ -n "${SMARTCTL_BIN:-}" ] && [ -b "$dev" ]; then
        si=$("$SMARTCTL_BIN" -i "$dev" 2>/dev/null)
        family=$(kv_get "$si" 'Model Family')
        vendor=$(kv_get "$si" 'Vendor')
    fi
    local model_brand=""
    model_brand=$(disk_brand_from_model "$model")
    if [ -n "$family" ]; then
        brand=$(disk_brand_from_model "$family $model")
        [ -z "$brand" ] && brand=$(echo "$family" | awk '{print $1}')
        how="smartctl Model Family"
        # 系列库与设备 IDENTIFY 型号结论冲突时，以型号为准，避免张冠李戴
        if [ -n "$model_brand" ] && [ "$model_brand" != "$brand" ]; then
            brand="$model_brand"; family=""
            how="型号前缀（与 Model Family 不一致，已按型号优先）"
        fi
    fi
    if [ -z "$brand" ]; then
        brand=$(disk_brand_from_model "$model")
        [ -n "$brand" ] && how="型号前缀"
    fi
    if [ -z "$brand" ] && [ -n "$vendor" ]; then
        case "$vendor" in ATA|VENDOR|Unknown|unknown|"") ;; *) brand="$vendor"; how="smartctl Vendor" ;; esac
    fi
    if [ -z "$brand" ] && [ -n "$sysv" ]; then
        case "$sysv" in
            ATA|VENDOR|Unknown|unknown|"0x0000"|"") ;;
            *) brand="$sysv"; how="sysfs device/vendor" ;;
        esac
    fi
    [ -z "$brand" ] && { brand="未识别"; how="型号 ${model:-未知} 未命中品牌规则"; }
    printf '%s|%s|%s\n' "$brand" "${family:-—}" "$how"
}

# =============================================================================
# 硬盘序列号（SN）读取
# 输出两行：第 1 行 SN，第 2 行来源。读不到时第 1 行为空、第 2 行给排查建议。
# 顺序按准确度排：smartctl -i（和盘体标签一致）→ lsblk → udev → sysfs。
# 不把 WWN 当 SN：vpd_pg80 是 Unit Serial Number，vpd_pg83 是设备标识（含 WWN），不用。
# =============================================================================
get_disk_serial() {
    # 必须拆两条 local：同一 local 语句内 ${name} 在赋值生效前展开，
    # dev 会取到调用者作用域的旧 name（传 $d 时路径错，三个 SN 源全废）
    local name="$1"
    local dev="/dev/${name}" sn="" src="" u f

    # ① smartctl -i：希捷/西数/NVMe 都给 "Serial Number:"，与标签 SN 一致
    if [ "${SMART_BACKEND:-none}" = "smartctl" ] && [ -n "${SMARTCTL_BIN:-}" ] && [ -b "$dev" ]; then
        sn=$("$SMARTCTL_BIN" -i "$dev" 2>/dev/null \
             | grep -iE '^[[:space:]]*Serial[[:space:]]*[Nn]umber[[:space:]]*:' \
             | head -1 | cut -d: -f2- | trim)
        [ -n "$sn" ] && src="smartctl -i"
    fi

    # ② lsblk -o SERIAL（多数 Kylin / CentOS 7 上直接可用）
    if [ -z "$sn" ] && [ "${TOOL_LSBLK:-0}" -eq 1 ]; then
        sn=$(lsblk -d -n -o SERIAL "$dev" 2>/dev/null | head -1 | trim)
        [ -n "$sn" ] && src="lsblk"
    fi

    # ③ udev 属性（CentOS 6 等老 lsblk 没有 SERIAL 列时很管用）
    if [ -z "$sn" ] && command -v udevadm &>/dev/null; then
        u=$(udevadm info --query=property --name="$dev" 2>/dev/null)
        sn=$(echo "$u" | grep -iE '^(ID_SCSI_SERIAL|ID_SERIAL_SHORT)=' \
             | head -1 | cut -d= -f2- | trim)
        # ID_SERIAL 常是 "ATA_型号_SN"，取末段；仅在前面都没读到时用
        if [ -z "$sn" ]; then
            sn=$(echo "$u" | grep -iE '^ID_SERIAL=' | head -1 | cut -d= -f2- | trim)
            case "$sn" in *_*_*) sn="${sn##*_}" ;; esac
        fi
        [ -n "$sn" ] && src="udevadm"
    fi

    # ④ sysfs：NVMe / USB 盘的 device/serial
    if [ -z "$sn" ]; then
        for f in "/sys/block/${name}/device/serial" "/sys/class/block/${name}/device/serial"; do
            [ -r "$f" ] || continue
            sn=$(cat "$f" 2>/dev/null | trim)
            [ -n "$sn" ] && { src="sysfs"; break; }
        done
    fi

    # ⑤ SCSI Unit Serial Number VPD page 0x80（SATA 盘经 libata 常在这里）
    #    前 6 字节是页头+描述符头，跳过后再取可打印字符
    if [ -z "$sn" ]; then
        f="/sys/block/${name}/device/vpd_pg80"
        if [ -r "$f" ]; then
            sn=$(tail -c +7 "$f" 2>/dev/null | tr -cd '[:print:]' | trim)
            [ -n "$sn" ] && src="sysfs VPD 0x80"
        fi
    fi

    if [ -z "$sn" ]; then
        printf '%s\n%s\n' "" "未提供（USB 桥/RAID 卡不透传或内核未导出；以盘体标签为准）"
    else
        printf '%s\n%s\n' "$sn" "$src"
    fi
}

# 硬盘：lsblk 枚举物理盘；品牌按 Model Family / 型号前缀推断；SN 四级回退。
# 末尾「采集说明」一行汇总 SN 来源、品牌依据与存储控制器
collect_disk_info() {
    log_info "正在收集硬盘信息..."
    add_section "五、硬盘信息"
    add_hdrN "盘符" "大小" "品牌 / 产品系列" "型号 / 类型 / 接口" "序列号 (SN)"

    local disk_found=0 sn_sources="" brand_sources=""
    if [ "$TOOL_LSBLK" -eq 1 ]; then
        # -P 键值对输出：空列也保留 KEY=""，MODEL 含空格、TRAN 为空都不会错位；
        # 老 util-linux 没有 -P/TRAN 时按能力分级降级（TRAN 列 2.23 才引入，
        # 老版 -o 遇到未知列会整体失败），全部不可用时退回空格分档解析
        local disks="" use_pairs=0 have_tran=0
        if lsblk -d -n -P -o NAME >/dev/null 2>&1; then
            use_pairs=1
            local cols="NAME,SIZE,MODEL,ROTA"
            if lsblk -d -n -o TRAN >/dev/null 2>&1; then
                have_tran=1
                cols="${cols},TRAN"
            fi
            disks=$(lsblk -d -n -P -o "$cols" 2>/dev/null)
        else
            if lsblk -d -n -o TRAN >/dev/null 2>&1; then
                have_tran=1
                disks=$(lsblk -d -n -o NAME,SIZE,MODEL,ROTA,TRAN 2>/dev/null)
            else
                disks=$(lsblk -d -n -o NAME,SIZE,MODEL,ROTA 2>/dev/null)
            fi
        fi
        if [ -n "$disks" ]; then
            while IFS= read -r line; do
                [ -z "${line// /}" ] && continue
                local name size model rota tran
                if [ "$use_pairs" -eq 1 ]; then
                    name=$(lsblk_pval "$line" NAME)
                    size=$(lsblk_pval "$line" SIZE)
                    model=$(lsblk_pval "$line" MODEL)
                    rota=$(lsblk_pval "$line" ROTA)
                    tran=$(lsblk_pval "$line" TRAN)
                else
                    local nf; nf=$(echo "$line" | awk '{print NF}')
                    name=$(echo "$line" | awk '{print $1}')
                    size=$(echo "$line" | awk '{print $2}')
                    if [ "$have_tran" -eq 1 ]; then
                        # ROTA 恒为 0/1、TRAN 永不取值 0/1，先看末列是谁：
                        # 末列是 rota → TRAN 为空（NVMe 等），否则末列是 tran。
                        # 靠这个区分才能处理「型号带空格且 TRAN 为空」的错位
                        local last_col; last_col=$(echo "$line" | awk '{print $NF}')
                        if [ "$last_col" = "0" ] || [ "$last_col" = "1" ]; then
                            rota="$last_col"
                            tran=""
                            model=""
                            [ "$nf" -ge 4 ] && model=$(echo "$line" | awk '{for(i=3;i<=NF-1;i++) printf "%s ",$i}' | trim)
                        else
                            tran="$last_col"
                            model=""
                            if [ "$nf" -ge 4 ]; then
                                rota=$(echo "$line" | awk '{print $(NF-1)}')
                                [ "$nf" -ge 5 ] && model=$(echo "$line" | awk '{for(i=3;i<=NF-2;i++) printf "%s ",$i}' | trim)
                            else
                                rota=""
                            fi
                        fi
                        # NF<3 的残缺行（如只有 name+size）会被上面误判，
                        # 统一清空：宁可留空也不能把 size 当成 model/rota/tran
                        if [ "$nf" -lt 3 ]; then
                            model=""; rota=""; tran=""
                        fi
                    else
                        # 无 TRAN 列：rota 是末列，model 占 3..NF-1
                        model=""; rota=""; tran=""
                        if [ "$nf" -ge 4 ]; then
                            model=$(echo "$line" | awk '{for(i=3;i<=NF-1;i++) printf "%s ",$i}' | trim)
                        fi
                        [ "$nf" -ge 3 ] && rota=$(echo "$line" | awk '{print $NF}')
                    fi
                fi
                # -d 把 nslaves==0 的设备全吐出来：loop 回环、ram 虚拟盘、zram 压缩盘、
                # sr 光驱、fd 软驱、dm/md 虚拟卷、nbd 网络盘、pmem 持久内存都不是
                # 要盘点的物理硬盘，按 0/1 转速强行标成 SSD/HDD 会误导资产盘点
                case "$name" in
                    ""|loop*|ram*|zram*|sr*|fd*|dm-*|md*|nbd*|pmem*) continue ;;
                esac
                disk_found=1
                local dtype="SSD"; [ "$rota" = "1" ] && dtype="HDD"
                local parts_raw=""
                [ -f "/sys/block/${name}/device/vendor" ] && parts_raw=$(cat "/sys/block/${name}/device/vendor" 2>/dev/null | trim)
                local detail="${model:-未知} / ${dtype}"
                [ -n "$tran" ] && [ "$tran" != "-" ] && detail="${detail} / ${tran}"

                local gb brand family bhow
                gb=$(get_disk_brand "$name" "${model:-}" "$parts_raw")
                brand=$(echo "$gb" | cut -d'|' -f1)
                family=$(echo "$gb" | cut -d'|' -f2)
                bhow=$(echo "$gb" | cut -d'|' -f3)
                # Model Family 常带品牌前缀（如 "Hitachi/HGST Ultrastar 7K2"），
                # 品牌列已有品牌，去掉前缀只留系列，避免 "HGST / 日立 / Hitachi/HGST ..." 三段拼接
                local fam_short; fam_short=$(echo "$family" | sed \
                    -e 's#^Hitachi/HGST ##' -e 's#^HGST ##' -e 's#^Seagate ##' \
                    -e 's#^Western Digital ##' -e 's#^WDC ##' -e 's#^Samsung ##' \
                    -e 's#^Toshiba ##')
                local brand_col="$brand"
                [ "$fam_short" != "—" ] && [ -n "$fam_short" ] \
                    && brand_col="${brand_col} · ${fam_short}"
                brand_sources="${brand_sources}${brand_sources:+; }${name}=${bhow}"

                local gs sn snsrc
                gs=$(get_disk_serial "$name")
                sn=$(echo "$gs" | sed -n '1p')
                snsrc=$(echo "$gs" | sed -n '2p')
                [ -z "$sn" ] && sn="未提供"
                sn_sources="${sn_sources}${sn_sources:+; }${name}=${snsrc}"
                add_rowN "$name" "$size" "$brand_col" "$detail" "$sn"
            done <<< "$disks"
        fi
    else
        local parts; parts=$(awk 'NR>2 {print $4}' /proc/partitions 2>/dev/null \
                            | grep -E '^[[:alnum:]]+$' \
                            | grep -vE '^(loop|ram|zram|dm|md|sr|fd|nbd|pmem)' | head -20)
        if [ -n "$parts" ]; then
            for d in $parts; do
                # /proc/partitions 同时列出整盘和分区（sda 与 sda1），
                # 分区不建 /sys/block/<name> 条目，借此只留整盘
                [ -d "/sys/block/$d" ] || continue
                disk_found=1
                local sz="" gs sn snsrc gb2 b2 f2 h2
                [ -f "/sys/block/$d/size" ] && sz="$(( $(cat "/sys/block/$d/size" 2>/dev/null) * 512 / 1024 / 1024 / 1024 ))G"
                gs=$(get_disk_serial "$d")
                sn=$(echo "$gs" | sed -n '1p')
                snsrc=$(echo "$gs" | sed -n '2p')
                [ -z "$sn" ] && sn="未提供"
                sn_sources="${sn_sources}${sn_sources:+; }${d}=${snsrc}"
                gb2=$(get_disk_brand "$d" "" "")
                b2=$(echo "$gb2" | cut -d'|' -f1)
                f2=$(echo "$gb2" | cut -d'|' -f2)
                h2=$(echo "$gb2" | cut -d'|' -f3)
                brand_sources="${brand_sources}${brand_sources:+; }${d}=${h2}"
                add_rowN "$d" "${sz:-未知}" "$b2" "型号未提供（无 lsblk） / 类型未知" "$sn"
            done
        fi
    fi
    [ "$disk_found" -eq 0 ] && add_rowN "未检测到磁盘" "" "" "" ""
    # 三条来源注脚合并成一行，避免「序列号读取来源 / 品牌识别依据 / 存储控制器」占三行。
    # 只有一块盘时去掉 "sda=" 这种多余前缀（表格里已经是单盘一行了）。
    local sn_note="${sn_sources:-未检测到磁盘}" br_note="${brand_sources:-未检测到磁盘}"
    if ! echo "$sn_sources" | grep -q ';'; then sn_note="${sn_note#*=}"; fi
    if ! echo "$brand_sources" | grep -q ';'; then br_note="${br_note#*=}"; fi
    local disk_notes="SN: ${sn_note}；品牌: ${br_note}（推断值，以盘体标签为准）"
    if [ "$TOOL_LSPCI" -eq 1 ]; then
        local ctrls; ctrls=$(lspci 2>/dev/null | grep -iE 'SATA|AHCI|NVMe|SAS' | head -5 \
            | sed -e 's/^[0-9a-f:.]* //' \
                  -e 's/^SATA controller: //' -e 's/^AHCI controller: //' \
                  -e 's/^NVMe controller: //' -e 's/^SAS controller: //' \
                  -e 's/ Co\., Ltd\.//g' -e 's/ Technology//g' \
                  -e 's/ (rev [^)]*)//g' \
            | tr '\n' ';' | sed 's/;$//')
        [ -n "$ctrls" ] && disk_notes="${disk_notes}；控制器: ${ctrls}"
    fi
    add_row "采集说明" "$disk_notes"
}

# =============================================================================
# 硬盘 SMART 通电信息采集（默认对所有物理盘执行，无需 -d 参数）
# =============================================================================
collect_smart_info() {
    log_info "正在收集硬盘 SMART 通电信息（后端: ${SMART_BACKEND}）..."
    smart_detect_env
    add_section "五（补充）、硬盘 SMART 通电信息"
    add_row "运行环境" "$SMART_ENV_DESC"
    [ "$SMART_ENV_IS_VM" -eq 1 ] && add_row "环境判定依据" "$SMART_ENV_REASON"
    add_row "SMART 后端" "$SMART_BACKEND"
    if [ "$SMART_BACKEND" = "smartctl" ]; then
        local sver="${SMARTCTL_VERSION:-未知}"
        add_row "SMART 后端程序" "${SMARTCTL_SOURCE:-未标注} | smartctl ${sver} | ${SMARTCTL_BIN}"
        if [ "$SMARTCTL_FARM_OK" -eq 0 ]; then
            add_row "后端能力" "smartctl ${sver} < 7.4，不支持 FARM 日志（希捷出厂装配日期）；通电/健康/SN/温度不受影响"
        fi
    fi
    [ "$SMART_BACKEND" = "native" ]   && add_row "SMART 后端程序" "内嵌原生 ioctl 探针 (${SMART_PY})"
    if [ "$SMART_BACKEND" = "none" ]; then
        add_row "采集状态" "无可用后端：${SMART_NATIVE_FAIL_REASON:-未检测到可用的 smartctl 或 python3 探针}"
        local static_hint="把静态编译的 smartctl 放脚本同目录，命名 ./smartctl、./smartctl.arm64、"
        static_hint="${static_hint}./smartctl.aarch64、./bin/arm64/smartctl 之一（按本机架构），"
        static_hint="${static_hint}或设置 SMARTCTL_BIN=/path/to/smartctl；静态文件不依赖 glibc 版本"
        add_row "解决方式（免 Python，推荐）" "$static_hint"
        local py_hint="安装/启用 python3，或设置 SMART_PY 指向可用解释器"
        case "$SMART_NATIVE_FAIL_REASON" in
            *未通过*)     py_hint="当前 python3 探针自检未通过：优先改用静态 smartctl，或更换/升级 python3 后重试" ;;
            *无法写入*)   py_hint="先修复临时目录写权限，或改用静态 smartctl" ;;
        esac
        add_row "解决方式（用 Python）" "$py_hint"
        add_row "离线说明" "本脚本不需要在目标机安装 smartmontools；以上两种方式任选其一，无网络也可完成采集"
        record_warn "无可用 SMART 后端，未采集硬盘通电时长/通电次数"
        return 0
    fi

    add_subheader "盘符" "通电时长 / 通电次数" "SMART 健康 / 温度 / 备注"

    local devs=""
    if [ "$TOOL_LSBLK" -eq 1 ]; then
        devs=$(lsblk -d -n -o NAME 2>/dev/null \
               | grep -vE '^(loop|ram|zram|dm-|md|sr|fd|nbd|pmem)')
    else
        local p bn
        for p in /sys/block/*; do
            [ -e "$p" ] || continue
            bn=$(basename "$p")
            case "$bn" in
                loop*|ram*|zram*|dm-*|md*|sr*|fd*|nbd*|pmem*) continue ;;
            esac
            devs="$devs $bn"
        done
    fi

    local found=0 readable=0
    local name disk_path line life extra mfg d s n
    for name in $devs; do
        [ -z "${name// /}" ] && continue
        found=1
        disk_path="/dev/${name}"
        if [ ! -b "$disk_path" ]; then
            add_table_row "$name" "设备节点不存在" "-"
            continue
        fi

        line=$(smart_query "$disk_path")
        smart_parse_line "$line"

        life=""
        [ -n "$Q_HOURS" ]  && life="${Q_HOURS} 小时"
        [ -n "$Q_CYCLES" ] && life="${life}${life:+ / }${Q_CYCLES} 次"
        [ -n "$life" ] && readable=$((readable + 1))
        [ -z "$life" ] && life="未获取到"

        extra=""
        [ -n "$Q_HEALTH" ] && extra="$Q_HEALTH"
        [ -n "$Q_TEMP" ]   && extra="${extra}${extra:+ / }${Q_TEMP}"
        if [ -z "$life" ] || [ "$life" = "未获取到" ]; then
            [ -n "$Q_NOTE" ] && extra="${extra}${extra:+ / }${Q_NOTE}"
        fi
        [ -z "$extra" ] && extra="未获取到（可能需要 root 权限）"

        add_table_row "$name" "$life" "$extra"
    done

    # ---- 出厂 / 装配日期（固件提供时才有，没有标准字段）----
    add_blank
    add_subheader "盘符（出厂信息）" "出厂 / 装配日期" "日期来源 / 说明（型号与 SN 见上一节）"
    for name in $devs; do
        [ -z "${name// /}" ] && continue
        disk_path="/dev/${name}"
        [ -b "$disk_path" ] || continue
        mfg=$(smart_get_disk_mfg "$disk_path" "$SMART_BACKEND")
        # 函数输出三行：日期 / 来源 / 说明；逐行取，避免 read 只吃第一行
        d=$(echo "$mfg" | sed -n '1p')
        s=$(echo "$mfg" | sed -n '2p')
        n=$(echo "$mfg" | sed -n '3p')
        add_table_row "$name" "${d:-未提供}" "${s:-无标准字段} | ${n:-}"
    done

    [ "$found" -eq 0 ] && add_table_row "未检测到物理磁盘" "" ""
    if [ "$readable" -eq 0 ]; then
        if [ "$SMART_ENV_IS_VM" -eq 1 ]; then
            # 虚拟盘 / 云块存储没有物理通电计数器，属预期行为，不计警告
            local vm_note="虚拟磁盘 / 云块存储不提供 SMART 通电属性；"
            vm_note="${vm_note}通电时长与通电次数属于宿主机物理盘，客户机无法读取。"
            vm_note="${vm_note}此为预期行为，非脚本或权限故障"
            add_row "环境说明" "$vm_note"
            [ -n "${Q_HEALTH:-}" ] && add_row "健康状态说明" "SMART 健康状态 ${Q_HEALTH} 仅为设备自评响应，不代表存在完整 SMART 属性表"
            add_row "如需物理盘通电信息" "请在宿主机或物理机上运行本脚本；虚拟机内无法获取"
            log_info "虚拟磁盘不提供 SMART 通电信息（预期行为），未计入警告"
        else
            local phy_hint="未读取到通电信息：① 确认以 root 运行；② 部分 USB 硬盘盒需 SAT 透传；"
            phy_hint="${phy_hint}③ 硬件 RAID 卡后的盘需 megaraid,N 等专用通道；④ 少数厂商用自定义属性布局"
            add_row "提示" "$phy_hint"
            record_warn "未能从任何磁盘读取 SMART 通电时长/通电次数"
        fi
    fi
    case "$SMART_BACKEND" in
        smartctl) add_footer "通电时长/次数来自 SMART 属性 ID 9 / ID 12；装配日期见本节出厂信息表" ;;
        native)   add_footer "SMART 数据来源: 内嵌原生 ioctl 探针 | 免安装·离线·跨 CPU 架构 | ATA 健康状态由关键属性(5/197/198)评估" ;;
    esac
}

# 隧道 / 容器 / 桥接等虚拟接口没有 MAC、没有 PCI 设备，进表全是「未知」垃圾行
is_virtual_iface() {
    case "${1%%@*}" in
        lo|sit*|ip6tnl*|gre*|tunl*|teql*|veth*|docker*|br-*|virbr*|\
        tap*|tun*|ifb*|dummy*|bonding_masters|nlmon*|hwsim*) return 0 ;;
        *) return 1 ;;
    esac
}

# lspci 的型号描述往往 100+ 字符且每行重复，剥掉设备类别前缀与公司后缀
short_pci_desc() {
    echo "$1" | sed \
        -e 's/^Ethernet controller: //' \
        -e 's/^Network controller: //' \
        -e 's/^Wireless controller: //' \
        -e 's/^InfiniBand controller: //' \
        -e 's/^Display controller: //' \
        -e 's/^VGA compatible controller: //' \
        -e 's/^3D controller: //' \
        -e 's/ Co\., Ltd\.//g' \
        -e 's/ Corporation//g' \
        -e 's/ Technology//g' \
        -e 's/ Semiconductor//g' \
        -e 's/ Inc\.//g' \
        -e 's/ (rev [^)]*)//g' \
        -e 's/Gigabit Ethernet Controller/GbE/g' \
        -e 's/Ethernet Controller/Ethernet/g' \
        -e 's/  */ /g' | trim
}

# 网卡：ip link/addr + lspci；过滤 sit/ip6tnl/veth/lo 等虚拟接口。
# DOWN 口速率写「—」而非「未知」（未插线本就读不到，不是采集失败）
collect_network_info() {
    log_info "正在收集网卡信息..."
    add_section "六、网卡信息"
    add_subheader "接口名" "MAC 地址" "型号 / 状态 / 速率 / IP"

    local iface_found=0
    if [ "$TOOL_IP" -eq 1 ]; then
        local ifaces; ifaces=$(ip link show 2>/dev/null | grep -E '^[0-9]+:' | grep -v 'lo:' | awk -F': ' '{print $2}' | cut -d'@' -f1 | awk '{print $1}')
        if [ -n "$ifaces" ]; then
            while IFS= read -r iface; do
                [ -z "$iface" ] && continue
                is_virtual_iface "$iface" && continue
                iface_found=1
                local mac ip_addr state speed
                mac=$(ip link show "$iface" 2>/dev/null | grep -i 'link/ether' | awk '{print $2}')
                [ -z "$mac" ] && mac="无"
                ip_addr=$(ip -4 addr show "$iface" 2>/dev/null | grep 'inet ' | awk '{print $2}' | tr '\n' ' ' | trim)
                [ -z "$ip_addr" ] && ip_addr="未分配"
                state=$(ip link show "$iface" 2>/dev/null | sed -n 's/.*state \([^ ]*\).*/\1/p')
                # DOWN / 未插线时 sysfs speed 本就无效，写「未知」容易被误读成采集失败
                speed="—"
                if [ -f "/sys/class/net/${iface}/speed" ]; then
                    speed=$(cat "/sys/class/net/${iface}/speed" 2>/dev/null)
                    if [ -n "$speed" ] && [ "$speed" -gt 0 ] 2>/dev/null; then
                        speed="${speed}Mb/s"
                    else
                        speed="—"
                    fi
                fi

                local model=""
                local pci_path="/sys/class/net/${iface}/device"
                if [ -L "$pci_path" ] && [ "$TOOL_LSPCI" -eq 1 ]; then
                    local pci_addr; pci_addr=$(basename "$(readlink "${pci_path}" 2>/dev/null)" 2>/dev/null)
                    if [ -n "$pci_addr" ]; then
                        model=$(lspci -s "$pci_addr" 2>/dev/null | sed 's/^[0-9a-f:. ()]* //')
                    fi
                fi
                [ -z "$model" ] && model="未知" || model=$(short_pci_desc "$model")
                add_table_row "$iface" "$mac" "${model} / ${state:-未知} / ${speed} / ${ip_addr}"
            done <<< "$ifaces"
        fi
    else
        local ifaces2; ifaces2=$(ls /sys/class/net/ 2>/dev/null | grep -v lo)
        if [ -n "$ifaces2" ]; then
            for iface in $ifaces2; do
                is_virtual_iface "$iface" && continue
                iface_found=1
                local mac2=""
                [ -f "/sys/class/net/${iface}/address" ] && mac2=$(cat "/sys/class/net/${iface}/address" 2>/dev/null)
                add_table_row "$iface" "${mac2:-无}" "-"
            done
        fi
    fi
    [ "$iface_found" -eq 0 ] && add_table_row "未检测到网卡接口" "" ""
}

# GPU：lspci 显示控制器 + nvidia-smi。nvidia-smi 的英文错误文本不是数据行，
# 必须跳过，否则会被 cut 成 9 个相同字段写进表格（历史 bug）
collect_gpu_info() {
    log_info "正在收集 GPU 信息..."
    add_section "七、GPU 信息"

    local gpu_found=0
    local nvidia_pci_addrs=""

    # ---- NVIDIA GPU (通过 nvidia-smi，最详细) ----
    if command -v nvidia-smi &>/dev/null; then
        local nvidia_output
        nvidia_output=$(nvidia-smi \
            --query-gpu=index,name,driver_version,memory.total,memory.used,memory.free,temperature.gpu,utilization.gpu,utilization.memory \
            --format=csv,noheader,nounits 2>/dev/null)
        # 驱动未装 / 无 NVIDIA 卡时，nvidia-smi 会输出一行错误文本，例如
        # "NVIDIA-SMI has failed because it couldn't communicate with ..."
        # 该文本含单引号且没有逗号：旧逻辑下 cut 会把整行原样吐给每个字段，
        # 既污染报表又让 xargs 刷屏。这里只保留"首字段是数字、至少 9 列"的真数据行。
        local valid_csv
        valid_csv=$(echo "$nvidia_output" | awk -F',' 'NF>=9 && $1 ~ /^[[:space:]]*[0-9]+[[:space:]]*$/')
        if [ -z "$valid_csv" ]; then
            log_info "nvidia-smi 未返回有效 GPU 数据（无 NVIDIA 设备或驱动未加载），跳过 NVIDIA 明细"
        fi
        if [ -n "$valid_csv" ]; then
            while IFS= read -r line; do
                [ -z "$line" ] && continue
                gpu_found=1
                local idx name driver mem_total mem_used mem_free temp gpu_util mem_util
                idx=$(echo "$line" | cut -d',' -f1 | trim)
                name=$(echo "$line" | cut -d',' -f2 | trim)
                driver=$(echo "$line" | cut -d',' -f3 | trim)
                mem_total=$(echo "$line" | cut -d',' -f4 | trim)
                mem_used=$(echo "$line" | cut -d',' -f5 | trim)
                mem_free=$(echo "$line" | cut -d',' -f6 | trim)
                temp=$(echo "$line" | cut -d',' -f7 | trim)
                gpu_util=$(echo "$line" | cut -d',' -f8 | trim)
                mem_util=$(echo "$line" | cut -d',' -f9 | trim)

                add_row "GPU ${idx} 型号" "$name"
                add_row "GPU ${idx} 类型" "NVIDIA"
                add_row "GPU ${idx} 驱动版本" "$driver"
                add_row "GPU ${idx} 显存总量" "${mem_total} MiB"
                add_row "GPU ${idx} 显存已用" "${mem_used} MiB (${mem_util}%)"
                add_row "GPU ${idx} 显存可用" "${mem_free} MiB"
                add_row "GPU ${idx} GPU 利用率" "${gpu_util}%"
                add_row "GPU ${idx} 温度" "${temp}°C"
                add_blank
            done <<< "$valid_csv"

            # 记录 NVIDIA GPU 的 PCI 地址，避免 lspci 重复检测
            if [ "$TOOL_LSPCI" -eq 1 ]; then
                nvidia_pci_addrs=$(lspci 2>/dev/null | grep -i NVIDIA | awk '{print $1}')
            fi
        fi
    fi

    # ---- 通过 lspci 检测 GPU（排除已由 nvidia-smi 报告的设备） ----
    if [ "$TOOL_LSPCI" -eq 1 ]; then
        local gpus; gpus=$(lspci 2>/dev/null | grep -iE 'VGA|3D|Display')
        if [ -n "$gpus" ]; then
            while IFS= read -r line; do
                [ -z "$line" ] && continue
                local pci_addr; pci_addr=$(echo "$line" | awk '{print $1}')
                # 跳过已被 nvidia-smi 报告的 NVIDIA GPU
                if [ -n "$nvidia_pci_addrs" ] && echo "$nvidia_pci_addrs" | grep -q "$pci_addr"; then
                    continue
                fi
                gpu_found=1
                local gpu_desc; gpu_desc=$(echo "$line" | sed 's/^[0-9a-f:. ]* //')

                local gpu_type="其他/未识别"
                echo "$gpu_desc" | grep -qi "NVIDIA"           && gpu_type="NVIDIA"
                echo "$gpu_desc" | grep -qi "AMD\|ATI\|Radeon" && gpu_type="AMD"
                echo "$gpu_desc" | grep -qi "Intel"            && gpu_type="Intel"
                echo "$gpu_desc" | grep -qi "Phytium\|飞腾\|X100\|DC Controller" && gpu_type="飞腾/国产显示控制器"
                echo "$gpu_desc" | grep -qi "Loongson\|兆芯\|海光\|景嘉微\|JM9\|摩尔\|芯瞳\|格兰菲" && gpu_type="国产 GPU"

                add_row "GPU (${pci_addr})" "${gpu_type} · $(short_pci_desc "$gpu_desc")"
            done <<< "$gpus"
        fi
    fi

    # ---- 兜底：通过 /sys 检测 ----
    if [ "$gpu_found" -eq 0 ]; then
        local sys_gpus; sys_gpus=$(ls -d /sys/class/drm/card[0-9]* 2>/dev/null | head -5)
        if [ -n "$sys_gpus" ]; then
            for card in $sys_gpus; do
                gpu_found=1
                add_row "GPU ($(basename "$card"))" "通过 /sys 检测到"
            done
        fi
    fi

    [ "$gpu_found" -eq 0 ] && add_row "GPU 信息" "未检测到 GPU 设备"
}

# =============================================================================
# 七（补充）、驱动与软件版本
# 网卡驱动：ethtool -i（driver / version / firmware-version）→ sysfs driver 链接
#           → modinfo -F version / /sys/module/<mod>/version。
# 其他 PCI 驱动：/sys/bus/pci/devices/*/driver + class 分类，不依赖 lspci -k
#           （CentOS 6 老 pciutils 可能没有 -k）。
# 软件版本：本节先只放 sshd；以后要加 nginx / openssl 等，在 sshd 那行下面
#           照样 add_row 即可，无需改表格结构。
# 坑：sshd -V 只有较新 OpenSSH 才支持，老系统（CentOS 6 的 5.3p1、
#     Kylin V10 的 8.2p1）会报 illegal option -- V，必须多级回退。
# =============================================================================

# 版本值归一：空 / unknown / N/A / 占位串统一成「未提供」，避免报表出现假值
norm_version() {
    local v; v=$(echo "$1" | trim)
    case "$v" in
        ''|unknown|Unknown|UNKNOWN|N/A|n/a|none|None|'0'|'0.0'|'-'|'—') echo "未提供" ;;
        *) echo "$v" ;;
    esac
}

# 取模块版本：sysfs 优先，其次 modinfo -F version。
# 很多内核内置驱动不声明 version 字段，此时显示「未提供」而非留空。
module_version() {
    local mod="$1" v=""
    [ -z "$mod" ] && { echo "未提供"; return 0; }
    if [ -r "/sys/module/${mod}/version" ]; then
        v=$(cat "/sys/module/${mod}/version" 2>/dev/null | head -1 | trim)
    fi
    if [ -z "$v" ] && command -v modinfo &>/dev/null; then
        v=$(modinfo -F version "$mod" 2>/dev/null | head -1 | trim)
    fi
    norm_version "$v"
}

# sshd 版本：四级回退，返回 "版本|路径|来源"
# ① sshd -V ② 服务端二进制版本串 ③ ssh -V（同包客户端）④ rpm/dpkg 查询
sshd_version_info() {
    local path="" p out="" src="" ver=""
    for p in "$(command -v sshd 2>/dev/null)" /usr/sbin/sshd \
             /usr/local/sbin/sshd /usr/bin/sshd; do
        if [ -n "$p" ] && [ -x "$p" ]; then path="$p"; break; fi
    done
    if [ -n "$path" ]; then
        out=$("$path" -V 2>&1 | head -1 | trim)
        echo "$out" | grep -q 'OpenSSH_' && { ver="$out"; src="sshd -V"; }
    fi
    # 老版 sshd 不支持 -V，但服务端二进制自身带版本串；用 grep -a 直接读，
    # 比退到客户端版本号更准，也不依赖 strings（精简系统常没装 binutils）
    if [ -z "$ver" ] && [ -n "$path" ] && [ -r "$path" ]; then
        out=$(grep -a -m1 -oE 'OpenSSH_[0-9][^ ,";]*' "$path" 2>/dev/null | head -1 | trim)
        [ -n "$out" ] && { ver="$out"; src="sshd 二进制版本串 (grep -a)"; }
    fi
    if [ -z "$ver" ] && command -v ssh &>/dev/null; then
        out=$(ssh -V 2>&1 | head -1 | trim)
        echo "$out" | grep -q 'OpenSSH_' && {
            ver="$out"; src="ssh -V（同包客户端；若服务端单独升级过会不一致）"
        }
    fi
    if [ -z "$ver" ]; then
        # rpm 与 dpkg-query 都试：有的系统装了 rpm 但包却不归它管
        if command -v rpm &>/dev/null; then
            out=$(rpm -q --qf '%{VERSION}-%{RELEASE}' openssh-server 2>/dev/null | head -1 | trim)
            out=$(echo "$out" | sed 's/^openssh-server-//')
            case "$out" in
                ''|*'not installed'*|*'未安装'*) ;;
                *) ver="openssh-server ${out}"; src="rpm -q openssh-server" ;;
            esac
        fi
        if [ -z "$ver" ] && command -v dpkg-query &>/dev/null; then
            out=$(dpkg-query -W -f='${Version}' openssh-server 2>/dev/null | head -1 | trim)
            [ -n "$out" ] && { ver="openssh-server ${out}"; src="dpkg-query openssh-server"; }
        fi
    fi
    [ -z "$ver" ] && ver="未提供"
    [ -z "$src" ] && src="未获取到"
    printf '%s|%s|%s\n' "$ver" "$path" "$src"
}

# 按相邻相同驱动合并成一行：10 个同型号网口只占一行，符合紧凑偏好
merge_nic_records() {
    printf '%b' "$1" | awk -F'|' '
        NF < 5 { next }
        function emit(   label) {
            if (cnt == 0) return
            label = (cnt == 1) ? first : first " ~ " last " (" cnt " 口)"
            printf "%s|网卡|%s|%s|%s|%s\n", label, kdrv, kver, kfw, ksrc
        }
        {
            key = $2 "|" $3 "|" $4 "|" $5
            if (key == cur) { last = $1; cnt++ }
            else { emit(); cur = key; first = $1; last = $1; cnt = 1
                   kdrv = $2; kver = $3; kfw = $4; ksrc = $5 }
        }
        END { emit() }
    '
}

merge_pci_records() {
    printf '%b' "$1" | awk -F'|' '
        NF < 4 { next }
        function emit(   label) {
            if (cnt == 0) return
            label = (cnt == 1) ? first : kkind " ×" cnt " 设备"
            printf "%s|%s|%s|%s|—|sysfs PCI class + modinfo\n", label, kkind, kdrv, kver
        }
        {
            key = $1 "|" $2 "|" $3
            if (key == cur) { cnt++ }
            else { emit(); cur = key; first = $4; cnt = 1
                   kkind = $1; kdrv = $2; kver = $3 }
        }
        END { emit() }
    '
}

collect_driver_versions() {
    log_info "正在收集驱动与软件版本..."
    add_section "七（补充）、驱动与软件版本"
    add_hdrN "设备 / 接口" "类型" "驱动模块" "驱动版本" "固件版本" "来源"

    # ---- 网卡 ----
    local ifaces="" rec="" iface drv ver fw ei mod src="" bus
    if [ "$TOOL_IP" -eq 1 ]; then
        ifaces=$(ip link show 2>/dev/null | grep -E '^[0-9]+:' | grep -v 'lo:' \
                 | awk -F': ' '{print $2}' | awk '{print $1}')
    else
        ifaces=$(ls /sys/class/net/ 2>/dev/null | grep -v '^lo$')
    fi
    while IFS= read -r iface; do
        [ -z "$iface" ] && continue
        is_virtual_iface "$iface" && continue
        drv=""; ver=""; fw=""; src=""
        if command -v ethtool &>/dev/null; then
            ei=$(ethtool -i "$iface" 2>/dev/null)
            if [ -n "$ei" ]; then
                drv=$(echo "$ei" | sed -n 's/^driver: //p' | head -1 | trim)
                ver=$(echo "$ei" | sed -n 's/^version: //p' | head -1 | trim)
                fw=$(echo "$ei"  | sed -n 's/^firmware-version: //p' | head -1 | trim)
                bus=$(echo "$ei" | sed -n 's/^bus-info: //p' | head -1 | trim)
                [ -n "$drv" ] && src="ethtool -i"
            fi
        fi
        # ethtool 缺失 / 不支持时，从 sysfs 的 driver 软链接拿模块名
        if [ -z "$drv" ] && [ -e "/sys/class/net/${iface}/device/driver" ]; then
            mod=$(basename "$(readlink -f "/sys/class/net/${iface}/device/driver" 2>/dev/null)" 2>/dev/null)
            [ -n "$mod" ] && { drv="$mod"; src="sysfs + modinfo"; }
        fi
        drv=$(norm_version "$drv"); ver=$(norm_version "$ver")
        # ethtool 常给空 / unknown 的 version：再用模块元数据补一次
        if [ "$ver" = "未提供" ] && [ "$drv" != "未提供" ]; then
            local mv; mv=$(module_version "$drv")
            if [ "$mv" != "未提供" ]; then
                ver="$mv"; src="${src:+${src}; }modinfo"
            fi
        fi
        fw=$(norm_version "$fw"); [ "$fw" = "未提供" ] && fw="—"
        [ -z "$src" ] && src="未提供"
        rec="${rec}${iface}|${drv}|${ver}|${fw}|${src}\n"
    done <<< "$ifaces"

    local line rows=0
    while IFS='|' read -r c1 c2 c3 c4 c5 c6; do
        [ -z "$c1" ] && continue
        add_rowN "$c1" "$c2" "$c3" "$c4" "$c5" "$c6"
        rows=$((rows + 1))
    done <<< "$(merge_nic_records "$rec")"

    # ---- 其他 PCI 设备：显示 / 存储 / USB 控制器（网卡 class 0x02 跳过，避免重复） ----
    rec=""
    local p cls kind addr
    for p in /sys/bus/pci/devices/*; do
        [ -e "$p/driver" ] || continue
        cls=$(cat "$p/class" 2>/dev/null)
        case "$cls" in
            0x03*)                                  kind="显示控制器" ;;
            0x0100*|0x0101*|0x0104*|0x0106*|0x0108*) kind="存储控制器" ;;
            0x0c03*)                                kind="USB 控制器" ;;
            *) continue ;;
        esac
        mod=$(basename "$(readlink -f "$p/driver" 2>/dev/null)" 2>/dev/null)
        [ -z "$mod" ] && continue
        ver=$(module_version "$mod")
        addr=$(basename "$p")
        rec="${rec}${kind}|${mod}|${ver}|${addr}\n"
    done
    while IFS='|' read -r c1 c2 c3 c4 c5 c6; do
        [ -z "$c1" ] && continue
        add_rowN "$c1" "$c2" "$c3" "$c4" "$c5" "$c6"
        rows=$((rows + 1))
    done <<< "$(merge_pci_records "$rec")"

    [ "$rows" -eq 0 ] && add_rowN "未检测到驱动信息" "—" "—" "—" "—" "—"

    # ---- 软件版本（先只放 sshd；要加别的在此继续 add_row） ----
    local sshd_info sver spath ssrc
    sshd_info=$(sshd_version_info)
    IFS='|' read -r sver spath ssrc <<< "$sshd_info"
    local sloc=""
    [ -n "$spath" ] && sloc=" | ${spath}"
    add_row "sshd (SSH 服务端)" "${sver}${sloc} | 来源: ${ssrc}"
}

# =============================================================================
# 硬盘详细信息与性能测试
# =============================================================================
parse_size_mb() {
    local size="$1"
    # 去除空格
    size=$(echo "$size" | tr -d '[:space:]')
    # 提取数字部分（整数）
    local num=$(echo "$size" | sed -E 's/^([0-9]+).*/\1/')
    local unit=$(echo "$size" | sed -E 's/[0-9]+//I' | tr '[:lower:]' '[:upper:]')
    case "$unit" in
        G) echo $((num * 1024)) ;;
        M) echo "$num" ;;
        K) echo $(((num + 1023) / 1024)) ;;
        *) echo "$num" ;;
    esac
}

parse_dd_speed() {
    local output="$1"
    # 提取最后一行中的速率值（MB/s 或 GB/s）
    local speed
    speed=$(echo "$output" | tail -1 | grep -oE '[0-9]+\.?[0-9]*\s*(MB|GB|MiB|GiB)/s' | tail -1)
    [ -z "$speed" ] && speed=$(echo "$output" | tail -1 | grep -oE '[0-9]+\.?[0-9]*\s*(MB|GB|MiB|GiB)/秒' | tail -1)
    [ -z "$speed" ] && {
        # fallback: 从多行输出中提取最后一个速率值
        speed=$(echo "$output" | grep -oE '[0-9]+\.?[0-9]*\s*(MB|GB|MiB|GiB)/s' | tail -1)
    }
    echo "${speed:-未知}"
}

parse_fio_val() {
    local output="$1" field="$2"
    echo "$output" | grep -i "${field}=" | tail -1 | sed -E "s/.*${field}=([^,]*).*/\1/"
}

# fio normal 输出的延迟摘要在 `clat (usec): min=..., avg=234.56` 这种行里；
# 用 `lat.*usec` 宽松匹配会先命中百分位行，-A1 取到的下一行并不是摘要。
# 按 clat→lat→slat 优先级找真正的摘要行，抽 avg 值和单位。
parse_fio_lat() {
    local output="$1" k line v u
    for k in clat lat slat; do
        line=$(echo "$output" | grep -E "^[[:space:]]*${k} \((nsec|usec|msec|sec)\):" | head -1)
        [ -z "$line" ] && continue
        v=$(echo "$line" | sed -n 's/.*avg=\([0-9][0-9.]*\).*/\1/p')
        u=$(echo "$line" | sed -n 's/^[[:space:]]*[a-z]* (\([a-z]*\)):.*/\1/p')
        if [ -n "$v" ]; then
            echo "${v} ${u} (avg)"
            return 0
        fi
    done
    return 1
}

# 测试结果落表：空值或「未知」说明测试/解析失败，必须上抛计数，
# 不能静默写成空白单元格让人误以为测过且正常
row_or_fail() {  # $1=表行标签  $2=测试结果  $3=失败上下文
    if [ -z "$2" ] || [ "$2" = "未知" ]; then
        record_error "性能测试失败: ${3}"
        add_row "$1" "测试失败"
    else
        add_row "$1" "$2"
    fi
}

# =============================================================================
# 硬盘分区与挂载信息（报表最后一节）
# 紧凑 6 列：分区 / 设备 | 大小 | 文件系统 | 挂载点 | 用量 | 标签 / UUID
# 「用量」压成 13G/44G (23%)，「标签/UUID」只占一列，NAME 用 ├─/└─ 树形。
# 列宽由 make_worksheet_xml 统一收紧，全部 6 列一屏可见。
# 三级数据源：① lsblk -P  ② lsblk 原样文本  ③ /proc/partitions + /proc/mounts
# =============================================================================
lsblk_pval() {
    local line="$1" key="$2"
    printf ' %s\n' "$line" \
        | sed -n 's/.* '"${key}"'="\([^"]*\)".*/\1/p' | head -1
}

partition_pttype() {
    local dev="/dev/$1" pt="$2"
    if [ -z "$pt" ] && command -v blkid &>/dev/null; then
        pt=$(blkid -p -o value -s PTTYPE "$dev" 2>/dev/null | head -1 | trim)
    fi
    if [ -z "$pt" ] && command -v parted &>/dev/null; then
        pt=$(parted -s "$dev" print 2>/dev/null \
             | grep -i 'Partition Table' | head -1 | cut -d: -f2- | trim)
    fi
    echo "${pt:-未提供}"
}

# 老版本 lsblk 没有 PKNAME 列时，按分区名反推父盘名
guess_parent_disk() {
    local n="$1"
    case "$n" in
        nvme*n*p[0-9]*|mmcblk*p[0-9]*) echo "${n%p*}" ;;
        *) echo "$n" | sed 's/[0-9][0-9]*$//' ;;
    esac
}

# 把 lsblk -P 的结构化行还原成 lsblk 默认 7 列（TSV，NAME 带树形前缀）
lsblk_to_tree_rows() {
    local data="$1"
    local -a L=() P=() D=()
    local line
    while IFS= read -r line; do [ -n "$line" ] && L+=("$line"); done <<< "$data"
    local n=${#L[@]}
    [ "$n" -eq 0 ] && return 0
    local i j k pk nm tp depth last prefix mj rmv sz rov mnt
    for ((i=0;i<n;i++)); do
        nm=$(lsblk_pval "${L[$i]}" NAME)
        tp=$(lsblk_pval "${L[$i]}" TYPE)
        pk=$(lsblk_pval "${L[$i]}" PKNAME)
        if [ -z "$pk" ] && { [ "$tp" = "part" ] || [ "$tp" = "lvm" ]; }; then
            pk=$(guess_parent_disk "$nm")
        fi
        P[$i]="$pk"
    done
    for ((i=0;i<n;i++)); do
        line="${L[$i]}"
        nm=$(lsblk_pval "$line" NAME)
        mj=$(lsblk_pval "$line" "MAJ:MIN")
        rmv=$(lsblk_pval "$line" RM);   [ -z "$rmv" ] && rmv="0"
        sz=$(lsblk_pval "$line" SIZE);  [ -z "$sz" ]  && sz="未知"
        rov=$(lsblk_pval "$line" RO);   [ -z "$rov" ] && rov="0"
        tp=$(lsblk_pval "$line" TYPE);  [ -z "$tp" ]  && tp="—"
        mnt=$(lsblk_pval "$line" MOUNTPOINT); [ -z "$mnt" ] && mnt="—"
        pk="${P[$i]}"
        depth=0
        if [ -n "$pk" ]; then
            for ((j=i-1;j>=0;j--)); do
                if [ "$(lsblk_pval "${L[$j]}" NAME)" = "$pk" ]; then
                    depth=$(( ${D[$j]:-0} + 1 )); break
                fi
            done
            [ "$depth" -eq 0 ] && depth=1
        fi
        D[$i]=$depth
        last=1
        for ((j=i+1;j<n;j++)); do
            [ "${P[$j]}" = "$pk" ] && { last=0; break; }
        done
        prefix=""
        if [ "$depth" -gt 0 ]; then
            for ((k=1;k<depth;k++)); do prefix="${prefix}   "; done
            [ "$last" -eq 1 ] && prefix="${prefix}└─" || prefix="${prefix}├─"
        fi
        printf '%s\t%s\t%s\t%s\t%s\t%s\t%s\n' \
            "${prefix}${nm}" "$mj" "$rmv" "$sz" "$rov" "$tp" "$mnt"
    done
}

fmt_size_kb() {
    echo "$1" | awk '{
        if      ($1>=1073741824) printf "%.1fT\n", $1/1073741824;
        else if ($1>=1048576)    printf "%.1fG\n", $1/1048576;
        else if ($1>=1024)       printf "%.0fM\n", $1/1024;
        else                     printf "%dK\n", $1;
    }' | sed 's/\.0\([TGM]\)$/\1/'
}

# 用量压成一行：13G/44G (23%)；未挂载或读不到给 —
partition_usage_short() {
    local mnt="$1" dfo=""
    [ -z "$mnt" ] || [ "$mnt" = "—" ] && { echo "—"; return; }
    [ "$mnt" = "[SWAP]" ] && { echo "交换分区"; return; }
    dfo=$(df -P -h "$mnt" 2>/dev/null | tail -1)
    [ -z "$dfo" ] && dfo=$(df -h "$mnt" 2>/dev/null | tail -1)
    if [ -n "$dfo" ]; then
        echo "$dfo" | awk '{print $3 "/" $4 " (" $5 ")"}'
    else
        echo "已挂载"
    fi
}

# 老 lsblk / /proc 兜底路径用 blkid 补文件系统、UUID、标签（一次调用）
# 输出 TSV：FSTYPE | UUID | LABEL
blkid_fields() {
    local dev="/dev/$1" be=""
    [ -b "$dev" ] || { printf '\t\t\n'; return; }
    be=$(blkid -o export "$dev" 2>/dev/null)
    printf '%s\t%s\t%s\n' \
        "$(echo "$be" | grep '^TYPE='    | head -1 | cut -d= -f2- | trim)" \
        "$(echo "$be" | grep '^UUID='    | head -1 | cut -d= -f2- | trim)" \
        "$(echo "$be" | grep '^LABEL='   | head -1 | cut -d= -f2- | trim)"
}

# 标签与 UUID 合并成一格，尽量短；UUID 保留完整以便资产盘点
compact_tag() {
    local lb="$1" uu="$2" s=""
    [ -n "$lb" ] && s="$lb"
    if [ -n "$uu" ]; then
        [ -n "$s" ] && s="${s} / "
        s="${s}${uu}"
    fi
    echo "${s:-—}"
}

# 分区与挂载：lsblk -P 结构化 → lsblk 原样文本 + blkid → /proc 兜底；
# 输出紧凑 6 列树形表，用量只对已挂载项用 df 计算
collect_partition_info() {
    local no="${1:-八}"
    log_info "正在收集硬盘分区与挂载信息..."
    add_section "${no}、硬盘分区与挂载信息"
    add_hdrN "分区 / 设备" "大小" "文件系统" "挂载点" "用量 (已用/可用/率)" "标签 / UUID"

    local total=0 mounted=0 pttypes="" src="" out="" cols="" fslist=""
    local -a PL=() TN=() TSV=()
    local mode="" line t nm sz fs mnt uuid label pttype typ clean usage tag
    local bfs buu blb i

    if [ "$TOOL_LSBLK" -eq 1 ]; then
        # ① 结构化 -P：按列分级回退（老 util-linux 可能没有 PTTYPE / UUID / PKNAME）
        for cols in \
            "NAME,MAJ:MIN,RM,SIZE,RO,TYPE,MOUNTPOINT,PKNAME,FSTYPE,UUID,LABEL,PTTYPE" \
            "NAME,MAJ:MIN,RM,SIZE,RO,TYPE,MOUNTPOINT,PKNAME,FSTYPE,UUID,LABEL" \
            "NAME,MAJ:MIN,RM,SIZE,RO,TYPE,MOUNTPOINT,PKNAME,FSTYPE" \
            "NAME,MAJ:MIN,RM,SIZE,RO,TYPE,MOUNTPOINT,PKNAME" \
            "NAME,SIZE,TYPE,MOUNTPOINT,PKNAME"; do
            out=$(lsblk -P -n -o "$cols" 2>/dev/null)
            [ -n "$out" ] && break
            out=""
        done
        [ -n "$out" ] && out=$(echo "$out" \
            | grep -vE 'NAME="(loop|ram|zram|dm-|md|sr|fd|nbd|pmem)')
        src="lsblk -P（结构化）"

        # ② -P 不被支持的老系统：拿 lsblk 原样文本（自带树形），细节用 blkid 补
        if [ -z "$out" ]; then
            out=$(lsblk -n -o NAME,MAJ:MIN,RM,SIZE,RO,TYPE,MOUNTPOINT 2>/dev/null \
                  | grep -vE '^[[:space:]]*(loop|ram|zram|dm-|md|sr|fd|nbd|pmem)')
            src="lsblk 原样文本 + blkid"
        fi
    fi

    if [ -n "$out" ]; then
        if echo "$out" | head -1 | grep -q 'NAME="'; then
            mode="pairs"
            while IFS= read -r line; do [ -n "$line" ] && PL+=("$line"); done <<< "$out"
            while IFS= read -r t; do TN+=("$t"); done < <(lsblk_to_tree_rows "$out" | cut -f1)
        else
            mode="plain"
            while IFS= read -r t; do TSV+=("$t"); done < <(echo "$out" | awk '{
                nm=$1; sz=$4; tp=$6;
                mnt=""; for(i=7;i<=NF;i++) mnt=mnt (i>7?" ":"") $i;
                printf "%s\t%s\t%s\t%s\n", nm, sz, tp, mnt
            }')
        fi
    fi

    # 三段数据源共用同一个落表函数，保证列数与格式一致
    emit_part_row() {
        local nm="$1" sz="$2" fs="$3" mnt="$4" uuid="$5" label="$6" \
              clean="$7" typ="$8" pt="$9"
        [ -z "$nm" ] && return
        [ -z "$sz" ]  && sz="—"
        [ -z "$mnt" ] && mnt="—"
        # 整盘行没有文件系统，blkid 兜底有时会误报成 ext4，统一留空
        if [ "$typ" = "disk" ]; then
            fs=""; uuid=""; label=""
        fi
        [ -z "$fs" ] && fs="—"
        [ "$fs" != "—" ]  && fslist="${fslist}${fslist:+, }${fs}"
        [ "$mnt" != "—" ] && mounted=$((mounted + 1))
        usage=$(partition_usage_short "$mnt")
        tag=$(compact_tag "$label" "$uuid")
        if [ "$typ" = "disk" ]; then
            pttypes="${pttypes}${pttypes:+; }${clean}=$(partition_pttype "$clean" "$pt")"
        fi
        total=$((total + 1))
        add_rowN "$nm" "$sz" "$fs" "$mnt" "$usage" "$tag"
    }

    if [ "$mode" = "pairs" ]; then
        for ((i=0;i<${#PL[@]};i++)); do
            line="${PL[$i]}"
            nm="${TN[$i]:-}"
            clean=$(lsblk_pval "$line" NAME)
            typ=$(lsblk_pval "$line" TYPE)
            fs=$(lsblk_pval "$line" FSTYPE)
            uuid=$(lsblk_pval "$line" UUID)
            label=$(lsblk_pval "$line" LABEL)
            # 老 lsblk 不支持 FSTYPE/UUID 列（分级降到最小列集）时用 blkid 补
            if [ -z "$fs" ] && [ -z "$uuid" ] && [ "$typ" != "disk" ]; then
                IFS=$'\t' read -r bfs buu blb <<< "$(blkid_fields "$clean")"
                fs="$bfs"; uuid="$buu"; label="$blb"
            fi
            emit_part_row "$nm" \
                "$(lsblk_pval "$line" SIZE)" "$fs" \
                "$(lsblk_pval "$line" MOUNTPOINT)" "$uuid" "$label" \
                "$clean" "$typ" "$(lsblk_pval "$line" PTTYPE)"
        done
    elif [ "$mode" = "plain" ]; then
        for t in "${TSV[@]}"; do
            IFS=$'\t' read -r nm sz typ mnt <<< "$t"
            clean=$(echo "$nm" | sed 's/^[^A-Za-z0-9]*//')
            IFS=$'\t' read -r bfs buu blb <<< "$(blkid_fields "$clean")"
            emit_part_row "$nm" "$sz" "$bfs" "$mnt" "$buu" "$blb" "$clean" "$typ" ""
        done
    fi

    # ③ 无 lsblk：/proc/partitions + /proc/mounts 兜底（列表格式，无树形）
    if [ "$total" -eq 0 ]; then
        src="/proc/partitions + /proc/mounts（无 lsblk）"
        local raw mj minor blocks dname sz2 mntline mnt2 typ2
        raw=$(awk 'NR>2 {print $1, $2, $3, $4}' /proc/partitions 2>/dev/null \
              | grep -vE ' (loop|ram|zram|dm-|md|sr|fd|nbd|pmem)')
        while read -r mj minor blocks dname; do
            [ -z "$dname" ] && continue
            sz2=$(fmt_size_kb "$blocks")
            mntline=$(grep -E "^/dev/${dname}[[:space:]]" /proc/mounts 2>/dev/null | head -1)
            mnt2=$(echo "$mntline" | awk '{print $2}')
            typ2="part"; [ -d "/sys/block/${dname}" ] && typ2="disk"
            IFS=$'\t' read -r bfs buu blb <<< "$(blkid_fields "$dname")"
            [ -z "$bfs" ] && bfs=$(echo "$mntline" | awk '{print $3}')
            emit_part_row "$dname" "$sz2" "$bfs" "$mnt2" "$buu" "$blb" "$dname" "$typ2" ""
        done <<< "$raw"
    fi

    [ "$total" -eq 0 ] && add_rowN "未检测到分区" "" "" "" "" ""
    local fs_uniq=""
    if [ -n "$fslist" ]; then
        fs_uniq=$(echo "$fslist" | tr ',' '\n' | sed 's/^[[:space:]]*//' \
                  | grep -v '^$' | sort -u | tr '\n' ',' \
                  | sed 's/,$//; s/,/, /g' | trim)
    fi
    # 汇总压成两行，不再堆一长串说明行
    add_row "分区汇总" "共 ${total} 条 / 已挂载 ${mounted} 个 / 文件系统 ${fs_uniq:-未识别} / 分区表 ${pttypes:-未提供}"
    add_row "数据来源" "${src}；用量取自 df -P -h（仅已挂载项）；已排除 loop/ram/zram 等虚拟设备"
}

# -d 详细段：sysfs 块设备参数 + SMART 摘要 + hdparm/dd/fio/ioping 性能测试。
# SMART 完整表格在「五（补充）」，这里只给摘要与固件/扇区计数等独有项
collect_disk_perf_info() {
    local disk="$1"
    local dn; dn=$(basename "$disk")
    log_info "正在收集硬盘详细信息并进行性能测试: ${disk}"

    # 检查磁盘设备
    if [ ! -b "$disk" ]; then
        record_error "磁盘设备不存在或不是块设备: ${disk}"
        add_section "八、硬盘详细信息与性能测试"
        add_row "状态" "磁盘设备检查失败: ${disk}"
        return 1
    fi

    # 检查测试目录
    if [ ! -d "$TEST_FILE_DIR" ] || [ ! -w "$TEST_FILE_DIR" ]; then
        record_error "测试目录不存在或不可写: ${TEST_FILE_DIR}"
        add_section "八、硬盘详细信息与性能测试"
        add_row "状态" "测试目录检查失败: ${TEST_FILE_DIR}"
        return 1
    fi

    add_section "八、硬盘详细信息与性能测试 - ${disk}"
    add_row "测试开始时间" "$(date '+%Y-%m-%d %H:%M:%S')"
    add_blank

    local test_size_mb; test_size_mb=$(parse_size_mb "$DISK_TEST_SIZE")
    local dd_count=$test_size_mb
    local missing_tools=""
    local used_tools="sysfs + blockdev"
    local err0=$ERROR_COUNT
    local write_ok=1

    # 空间校验：写测试要在 TEST_FILE_DIR 预留约 test_size_mb 的空间，
    # 空间不足还继续跑会把分区写满，撑爆系统；读测试只读被测盘，不受影响
    local avail_kb
    avail_kb=$(df -Pk "$TEST_FILE_DIR" 2>/dev/null | awk 'NR==2{print $4}')
    if [ -n "$avail_kb" ] && [ "$avail_kb" -lt $((test_size_mb * 1024)) ] 2>/dev/null; then
        write_ok=0
        record_error "测试目录剩余空间不足: ${TEST_FILE_DIR} 可用 $((avail_kb / 1024))M < ${DISK_TEST_SIZE}，跳过全部写测试"
    fi

    # 落盘校验：-p 目录可能落在 tmpfs/根盘上，与 -d 被测盘不是同一块，
    # 测出的「写速度」是内存或另一块盘的速度，必须在结果里标出来
    if [ "$write_ok" -eq 1 ]; then
        local wsrc wname
        wsrc=$(df -P "$TEST_FILE_DIR" 2>/dev/null | awk 'NR==2{print $1}')
        wname=$(basename "$wsrc" 2>/dev/null)
        case "$wsrc" in
            /dev/*)
                if [ ! -b "$wsrc" ]; then
                    log_warn "测试目录落盘设备不存在（${wsrc}），写入速度仅供参考"
                else
                    # PKNAME 逐级回溯到整盘根（sda1→sda），与 -d 的设备名比对；
                    # 老 util-linux 无 PKNAME 列时 lsblk 失败，回溯不到就跳过比对
                    local cur="$wname" parent="" hops=0
                    while [ -n "$cur" ] && [ "$hops" -lt 5 ]; do
                        parent=$(lsblk -n -o PKNAME "/dev/$cur" 2>/dev/null | head -1 | trim)
                        hops=$((hops + 1))
                        [ -z "$parent" ] && break
                        cur="$parent"
                    done
                    if { [ -n "$parent" ] || [ -e "/sys/block/$cur" ]; } && [ "$cur" != "$dn" ]; then
                        log_warn "测试目录落盘于 ${wname}，与被测盘 ${dn} 不同，写入速度不代表 ${dn}"
                        add_row "写入测试路径提示" "目录落盘 ${wname} ≠ 被测 ${dn}，写速度不可比"
                    fi
                fi
                ;;
            *)
                [ -n "$wsrc" ] && log_warn "测试目录不在块设备上（${wsrc}），写入速度可能是内存/网络盘，不代表 ${dn}"
                ;;
        esac
    fi

    # ==========================================================================
    # A. 磁盘基础参数
    # ==========================================================================
    add_row "磁盘设备" "$disk"

    local size_bytes="" rotational="" scheduler="" sector_size="" read_ahead=""
    local nr_requests="" max_sectors_kb="" phys_block_size="" log_block_size="" discard=""
    local device_vendor="" device_model=""
    [ -f "/sys/block/${dn}/size" ]                     && size_bytes=$(( $(cat "/sys/block/${dn}/size" 2>/dev/null) * 512 ))
    [ -f "/sys/block/${dn}/queue/rotational" ]         && rotational=$(cat "/sys/block/${dn}/queue/rotational" 2>/dev/null)
    [ -f "/sys/block/${dn}/queue/scheduler" ]          && scheduler=$(cat "/sys/block/${dn}/queue/scheduler" 2>/dev/null | sed 's/.*\[\([^]]*\)\].*/\1/')
    [ -f "/sys/block/${dn}/queue/hw_sector_size" ]     && sector_size=$(cat "/sys/block/${dn}/queue/hw_sector_size" 2>/dev/null)
    [ -f "/sys/block/${dn}/queue/read_ahead_kb" ]      && read_ahead=$(cat "/sys/block/${dn}/queue/read_ahead_kb" 2>/dev/null)
    [ -f "/sys/block/${dn}/queue/nr_requests" ]        && nr_requests=$(cat "/sys/block/${dn}/queue/nr_requests" 2>/dev/null)
    [ -f "/sys/block/${dn}/queue/max_sectors_kb" ]     && max_sectors_kb=$(cat "/sys/block/${dn}/queue/max_sectors_kb" 2>/dev/null)
    [ -f "/sys/block/${dn}/queue/physical_block_size" ] && phys_block_size=$(cat "/sys/block/${dn}/queue/physical_block_size" 2>/dev/null)
    [ -f "/sys/block/${dn}/queue/logical_block_size" ]  && log_block_size=$(cat "/sys/block/${dn}/queue/logical_block_size" 2>/dev/null)
    [ -f "/sys/block/${dn}/queue/discard_max_bytes" ]   && discard=$(cat "/sys/block/${dn}/queue/discard_max_bytes" 2>/dev/null)
    [ -f "/sys/block/${dn}/device/vendor" ] && device_vendor=$(cat "/sys/block/${dn}/device/vendor" 2>/dev/null | trim)
    [ -f "/sys/block/${dn}/device/model" ]  && device_model=$(cat "/sys/block/${dn}/device/model" 2>/dev/null | trim)

    if command -v blockdev &>/dev/null; then
        [ -z "$sector_size" ] && sector_size=$(blockdev --getbsz "$disk" 2>/dev/null)
        if [ -z "$read_ahead" ]; then
            local ra_sectors; ra_sectors=$(blockdev --getra "$disk" 2>/dev/null)
            [ -n "$ra_sectors" ] && read_ahead=$(( ra_sectors * 512 / 1024 ))
        fi
    fi

    # 容量优先取 lsblk，与「五、硬盘信息」同口径：sysfs size*512 再 numfmt 会得到
    # 932G，而 lsblk 给 931.5G，同一块盘在两节里数字不一致会让人以为换过盘。
    local size_hr=""
    if [ "$TOOL_LSBLK" -eq 1 ]; then
        size_hr=$(lsblk -n -d -o SIZE "$disk" 2>/dev/null | head -1 | trim)
    fi
    if [ -z "$size_hr" ] && [ -n "$size_bytes" ] && [ "$size_bytes" -gt 0 ] 2>/dev/null; then
        if command -v numfmt &>/dev/null; then
            size_hr=$(numfmt --to=iec "$size_bytes" 2>/dev/null)
        else
            size_hr="$(( size_bytes / 1024 / 1024 / 1024 ))G"
        fi
    fi
    [ -n "$size_hr" ] && add_row "磁盘容量" "$size_hr"

    local dtype="SSD"; [ "$rotational" = "1" ] && dtype="HDD (机械)"
    add_row "磁盘类型" "$dtype"

    # sysfs device/vendor 常是 "ATA"（传输层标识，不是品牌）；device/model 还会被
    # 截断（真机出现过 "HGST HUS722T1TAL" 少了尾部 A604）。优先用 smartctl -i 的
    # 完整型号 + 型号前缀推断品牌，两者都没有才退回 sysfs 的值。
    local full_model="$device_model" brand="" si_m sm lm
    # lsblk MODEL 多数系统给完整型号（无 smartctl 时用它兜底，避免 sysfs 截断值）
    if [ "$TOOL_LSBLK" -eq 1 ]; then
        lm=$(lsblk -n -d -o MODEL "$disk" 2>/dev/null | head -1 | trim)
        [ -n "$lm" ] && full_model="$lm"
    fi
    if [ "$SMART_BACKEND" = "smartctl" ] && [ -n "${SMARTCTL_BIN:-}" ]; then
        si_m=$("$SMARTCTL_BIN" -i "$disk" 2>/dev/null)
        sm=$(echo "$si_m" | grep -iE '^(Device Model|Model Number|Product):' \
             | head -1 | cut -d: -f2- | trim)
        [ -n "$sm" ] && full_model="$sm"
    fi
    [ -n "$full_model" ] && brand=$(disk_brand_from_model "$full_model")
    if [ -n "$full_model" ] && [ "$full_model" != "MODEL" ]; then
        add_row "型号 / 品牌" "${full_model}${brand:+ · ${brand}}"
    fi
    # 7 个 sysfs 参数原来是 7 行，全是短值，占了大半屏。按语义并成 3 行，
    # 值列横跨 5 列，一行放得下，字段一个不少。
    local io_line="调度器 ${scheduler:-未知} / 队列 ${nr_requests:-未知}"
    [ -n "$max_sectors_kb" ] && io_line="${io_line} / 最大 IO ${max_sectors_kb} KB"
    add_row "IO 参数" "$io_line"
    add_row "扇区大小" "物理 ${phys_block_size:-${sector_size:-未知}} / 逻辑 ${log_block_size:-${sector_size:-未知}} 字节"
    local discard_str="不支持"
    [ -n "$discard" ] && [ "$discard" != "0" ] && discard_str="支持"
    add_row "预读 / TRIM" "${read_ahead:-未知} KB / ${discard_str}"
    add_blank

    # ==========================================================================
    # B. SMART / 设备信息（统一后端，支持免安装原生 ioctl 探针）
    # ==========================================================================
    add_smart_detail_rows "$disk"
    case "$SMART_BACKEND" in
        smartctl) used_tools="${used_tools} + smartctl" ;;
        native)   used_tools="${used_tools} + 原生ioctl探针(免安装)" ;;
        none)     missing_tools="$missing_tools smartmontools" ;;
    esac

    # udevadm 设备路径信息
    if command -v udevadm &>/dev/null; then
        local udev_info; udev_info=$(udevadm info --query=all --name="$disk" 2>/dev/null)
        if [ -n "$udev_info" ]; then
            local id_type wwn dev_id=""
            id_type=$(echo "$udev_info" | grep "ID_ATA_ROTATION_RATE_RPM=" | cut -d= -f2- | head -1)
            wwn=$(echo "$udev_info"     | grep "ID_WWN=" | cut -d= -f2- | head -1)
            [ -n "$id_type" ] && [ "$id_type" != "0" ] && dev_id="转速 ${id_type} RPM"
            [ -n "$wwn" ] && dev_id="${dev_id}${dev_id:+ / }WWN ${wwn}"
            [ -n "$dev_id" ] && add_row "设备标识" "$dev_id"
        fi
    fi

    # ==========================================================================
    # C. hdparm 缓存/缓冲读取测试
    # ==========================================================================
    if command -v hdparm &>/dev/null; then
        log_info "  [hdparm] 正在测试 ${disk} 读取速度..."
        local hdparm_out; hdparm_out=$(hdparm -tT "$disk" 2>/dev/null)
        if [ -n "$hdparm_out" ]; then
            local cached_read buffered_read
            cached_read=$(echo "$hdparm_out"  | grep "Timing cached reads" | sed 's/^[[:space:]]*Timing cached reads:[[:space:]]*//')
            buffered_read=$(echo "$hdparm_out" | grep "Timing buffered disk reads" | sed 's/^[[:space:]]*Timing buffered disk reads:[[:space:]]*//')
            [ -n "$cached_read" ]   && add_row "hdparm 缓存读取" "$cached_read"
            [ -n "$buffered_read" ] && add_row "hdparm 缓冲读取" "$buffered_read"
            used_tools="${used_tools} + hdparm"
        else
            add_row "hdparm 测试" "无输出（可能需要 root 权限）"
        fi
        add_blank
    fi

    # ==========================================================================
    # D. dd 顺序读取测试
    # ==========================================================================
    log_info "  [dd] 读取测试中..."
    # coreutils <8.24 的 dd 不认识 status= 会直接报错退出，先探测能力
    local dd_status_args=""
    dd --help 2>&1 | grep -q 'status=' && dd_status_args="status=progress"
    local dd_read_out dd_read_speed
    # 不加引号是故意的：0 或 2 个词都必须原样展开为 0/2 个参数
    dd_read_out=$(dd if="$disk" of=/dev/null bs=1M count="$dd_count" iflag=direct $dd_status_args 2>&1)
    dd_read_speed=$(parse_dd_speed "$dd_read_out")
    row_or_fail "dd 顺序读取 (bs=1M)" "$dd_read_speed" "dd 读取 ${disk}: $(echo "$dd_read_out" | tail -1)"
    used_tools="${used_tools} + dd"
    add_blank

    # ==========================================================================
    # E. dd 顺序写入测试
    # ==========================================================================
    log_info "  [dd] 写入测试中..."
    # mktemp O_EXCL：固定名 + $$ 存在被同目录低权用户预埋符号链接劫持的窗口
    local test_file=""
    if [ "$write_ok" -eq 1 ]; then
        test_file=$(mktemp "${TEST_FILE_DIR}/.disk_write_test_XXXXXX" 2>/dev/null)
        if [ -z "$test_file" ]; then
            write_ok=0
            record_error "无法在 ${TEST_FILE_DIR} 创建写测试文件，跳过后续写测试"
        fi
    fi
    if [ -z "$test_file" ]; then
        add_row "dd 顺序写入 (bs=1M)" "跳过（写入测试不可用）"
    else
        local dd_write_out dd_write_speed
        dd_write_out=$(dd if=/dev/zero of="$test_file" bs=1M count="$dd_count" oflag=direct conv=fdatasync $dd_status_args 2>&1)
        dd_write_speed=$(parse_dd_speed "$dd_write_out")
        row_or_fail "dd 顺序写入 (bs=1M)" "$dd_write_speed" "dd 写入 ${test_file}: $(echo "$dd_write_out" | tail -1)"
        rm -f "$test_file" 2>/dev/null
    fi
    add_blank

    # ==========================================================================
    # F. fio 综合基准测试
    # ==========================================================================
    if command -v fio &>/dev/null; then
        local fio_out
        used_tools="${used_tools} + fio"

        log_info "  [fio] 正在测试顺序读取 (seq read, ${DISK_TEST_SIZE})..."
        fio_out=$(fio --name=seq_read --filename="$disk" --direct=1 \
            --rw=read --bs=1M --size="$DISK_TEST_SIZE" --runtime=30 \
            --time_based --output-format=normal 2>&1)
        local fio_ctx; fio_ctx=$(echo "$fio_out" | grep -iE 'error|fail' | tail -1)
        row_or_fail "fio 顺序读取带宽"  "$(parse_fio_val "$fio_out" 'BW')" "fio seq_read: ${fio_ctx:-带宽值为空}"
        row_or_fail "fio 顺序读取 IOPS" "$(parse_fio_val "$fio_out" 'IOPS')" "fio seq_read: ${fio_ctx:-IOPS 值为空}"
        add_blank

        # mktemp O_EXCL 防同目录预埋符号链接；写测试统一受 write_ok 开关控制
        local test_file2=""
        if [ "$write_ok" -eq 1 ]; then
            test_file2=$(mktemp "${TEST_FILE_DIR}/.fio_seq_write_XXXXXX" 2>/dev/null)
            [ -z "$test_file2" ] && record_error "无法在 ${TEST_FILE_DIR} 创建 fio 顺序写测试文件"
        fi
        if [ -z "$test_file2" ]; then
            add_row "fio 顺序写入" "跳过（写入测试不可用）"
            add_blank
        else
            log_info "  [fio] 正在测试顺序写入 (seq write, ${DISK_TEST_SIZE})..."
            fio_out=$(fio --name=seq_write --filename="$test_file2" --direct=1 \
                --size="$DISK_TEST_SIZE" --rw=write --bs=1M --runtime=30 \
                --time_based --output-format=normal --fallocate=none 2>&1)
            fio_ctx=$(echo "$fio_out" | grep -iE 'error|fail' | tail -1)
            row_or_fail "fio 顺序写入带宽"  "$(parse_fio_val "$fio_out" 'BW')" "fio seq_write: ${fio_ctx:-带宽值为空}"
            row_or_fail "fio 顺序写入 IOPS" "$(parse_fio_val "$fio_out" 'IOPS')" "fio seq_write: ${fio_ctx:-IOPS 值为空}"
            rm -f "$test_file2" 2>/dev/null
            add_blank
        fi

        log_info "  [fio] 正在测试 4K 随机读取 (randread, ${DISK_TEST_SIZE})..."
        fio_out=$(fio --name=rand_read --filename="$disk" --direct=1 \
            --rw=randread --bs=4k --size="$DISK_TEST_SIZE" --iodepth=64 \
            --runtime=20 --time_based --output-format=normal 2>&1)
        fio_ctx=$(echo "$fio_out" | grep -iE 'error|fail' | tail -1)
        row_or_fail "fio 4K 随机读取带宽"  "$(parse_fio_val "$fio_out" 'BW')" "fio rand_read: ${fio_ctx:-带宽值为空}"
        row_or_fail "fio 4K 随机读取 IOPS" "$(parse_fio_val "$fio_out" 'IOPS')" "fio rand_read: ${fio_ctx:-IOPS 值为空}"
        local lat; lat=$(parse_fio_lat "$fio_out")
        row_or_fail "fio 4K 随机读取延迟" "$lat" "fio rand_read: ${fio_ctx:-延迟行缺失}"
        add_blank

        local test_file3=""
        if [ "$write_ok" -eq 1 ]; then
            test_file3=$(mktemp "${TEST_FILE_DIR}/.fio_rand_write_XXXXXX" 2>/dev/null)
            [ -z "$test_file3" ] && record_error "无法在 ${TEST_FILE_DIR} 创建 fio 随机写测试文件"
        fi
        if [ -z "$test_file3" ]; then
            add_row "fio 4K 随机写入" "跳过（写入测试不可用）"
            add_blank
        else
            log_info "  [fio] 正在测试 4K 随机写入 (randwrite, ${DISK_TEST_SIZE})..."
            fio_out=$(fio --name=rand_write --filename="$test_file3" \
                --size="$DISK_TEST_SIZE" --rw=randwrite --bs=4k --direct=1 --iodepth=64 \
                --runtime=20 --time_based --output-format=normal --fallocate=none 2>&1)
            fio_ctx=$(echo "$fio_out" | grep -iE 'error|fail' | tail -1)
            row_or_fail "fio 4K 随机写入带宽"  "$(parse_fio_val "$fio_out" 'BW')" "fio rand_write: ${fio_ctx:-带宽值为空}"
            row_or_fail "fio 4K 随机写入 IOPS" "$(parse_fio_val "$fio_out" 'IOPS')" "fio rand_write: ${fio_ctx:-IOPS 值为空}"
            lat=$(parse_fio_lat "$fio_out")
            row_or_fail "fio 4K 随机写入延迟" "$lat" "fio rand_write: ${fio_ctx:-延迟行缺失}"
            rm -f "$test_file3" 2>/dev/null
            add_blank
        fi
    else
        add_row "fio 基准测试" "fio 未安装 — 无法测试 IOPS/延迟/随机读写性能"
        missing_tools="$missing_tools fio"
        add_blank
    fi

    # ==========================================================================
    # G. ioping 延迟测试
    # ==========================================================================
    if command -v ioping &>/dev/null; then
        log_info "  [ioping] 正在测试 ${disk} IO 延迟..."
        local ioping_out ioping_res
        ioping_out=$(ioping -c 10 "$disk" 2>&1)
        ioping_res=$(echo "$ioping_out" | grep "avg" | sed 's/.*avg=\([^ ]*\).*/\1/')
        row_or_fail "ioping 读取延迟" "$ioping_res" "ioping 读取 ${disk}: $(echo "$ioping_out" | tail -1)"

        # mktemp O_EXCL 防预埋符号链接；写测试受 write_ok 开关控制
        local ioping_w_file=""
        if [ "$write_ok" -eq 1 ]; then
            ioping_w_file=$(mktemp "${TEST_FILE_DIR}/.ioping_write_XXXXXX" 2>/dev/null)
            [ -z "$ioping_w_file" ] && record_error "无法在 ${TEST_FILE_DIR} 创建 ioping 写测试文件"
        fi
        if [ -z "$ioping_w_file" ]; then
            add_row "ioping 写入延迟" "跳过（写入测试不可用）"
        else
            ioping_out=$(ioping -c 10 -W "$ioping_w_file" 2>&1)
            ioping_res=$(echo "$ioping_out" | grep "avg" | sed 's/.*avg=\([^ ]*\).*/\1/')
            row_or_fail "ioping 写入延迟" "$ioping_res" "ioping 写入: $(echo "$ioping_out" | tail -1)"
            rm -f "$ioping_w_file" 2>/dev/null
        fi
        used_tools="${used_tools} + ioping"
        add_blank
    else
        add_row "ioping 延迟测试" "ioping 未安装 — 无法测试 IO 延迟"
        missing_tools="$missing_tools ioping"
        add_blank
    fi

    if [ "$ERROR_COUNT" -gt "$err0" ]; then
        add_footer "测试结束（有失败项，见上方各节「测试失败/跳过」行及日志错误） | 测试数据大小: ${DISK_TEST_SIZE} | 工具: ${used_tools}"
    else
        add_footer "测试完成 | 测试数据大小: ${DISK_TEST_SIZE} | 工具: ${used_tools}"
    fi

    # 缺失工具提示：离线机器上 apt/yum 根本装不了，给「可忽略」的明确结论，
    # 不要打印一串在气隙环境里无法执行的安装命令。
    if [ -n "$missing_tools" ]; then
        local mt; mt=$(echo "$missing_tools" | tr ' ' '\n' | grep -v '^$' | tr '\n' ' ' | trim)
        echo ""
        log_info "可选性能测试工具未安装（不影响导出与 SMART 采集）: ${mt}"
        log_info "离线批量场景可忽略；如需 IOPS/延迟，请在有软件源的机器补齐或随包携带"
        echo ""
    fi
}

# =============================================================================
# XLS 文件生成
# =============================================================================
make_xls_styles() {
    cat << 'XLSSTYLE'
  <Styles>
    <Style ss:ID="Default" ss:Name="Normal">
      <Font ss:Size="10" ss:FontName="Calibri"/>
    </Style>
    <Style ss:ID="title">
      <Font ss:Bold="1" ss:Size="14" ss:Color="#1F4E79"/>
    </Style>
    <Style ss:ID="section">
      <Font ss:Bold="1" ss:Size="10"/>
      <Interior ss:Color="#D6E4F0" ss:Pattern="Solid"/>
      <Borders>
        <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>
      </Borders>
    </Style>
    <Style ss:ID="label">
      <Font ss:Bold="1" ss:Size="10"/>
      <Alignment ss:Vertical="Top" ss:WrapText="1"/>
      <Borders>
        <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>
        <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>
        <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>
        <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>
      </Borders>
    </Style>
    <Style ss:ID="value">
      <Font ss:Size="10"/>
      <Alignment ss:Vertical="Top" ss:WrapText="1"/>
      <Borders>
        <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>
        <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>
        <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>
        <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>
      </Borders>
    </Style>
  </Styles>
XLSSTYLE
}

make_worksheet_xml() {
    local sheet_name="$1" data_rows="$2"
    cat << XLSWORKSHEET
  <Worksheet ss:Name="${sheet_name}">
    <Table>
      <Column ss:AutoFitWidth="0" ss:Width="118"/>
      <Column ss:AutoFitWidth="0" ss:Width="150"/>
      <Column ss:AutoFitWidth="0" ss:Width="100"/>
      <Column ss:AutoFitWidth="0" ss:Width="128"/>
      <Column ss:AutoFitWidth="0" ss:Width="150"/>
      <Column ss:AutoFitWidth="0" ss:Width="205"/>
${data_rows}
    </Table>
  </Worksheet>
XLSWORKSHEET
}

# =============================================================================
# 页签序号管理
# =============================================================================
get_max_sheet_num() {
    local file="$1"
    [ ! -f "$file" ] && echo "0" && return
    local nums
    if echo "test" | grep -oP 'test' &>/dev/null; then
        nums=$(grep -oP 'ss:Name="\K[0-9]+' "$file" 2>/dev/null | sort -n)
    else
        nums=$(grep 'ss:Name=' "$file" 2>/dev/null | sed 's/.*ss:Name="\([0-9]*\).*/\1/' | sort -n)
    fi
    [ -n "$nums" ] && echo "$(echo "$nums" | tail -1)" || echo "0"
}

# =============================================================================
# 主流程
# =============================================================================
main() {
    echo ""
    echo "=============================================="
    echo "  Linux 硬件配置信息导出工具 v${SCRIPT_VERSION}"
    echo "  交付: Excel XML Spreadsheet (.xls) · 6 列紧凑表格 · 多页签追加"
    echo "  采集: CPU多路 / 内存 / 硬盘品牌·SN / 通电·出厂 / 网卡 / GPU / 驱动·软件版本 / 分区挂载"
    echo "  离线: 单文件免安装，SMART 四级后端回退；-d 可附硬盘性能测试"
    echo "=============================================="
    echo ""

    check_tools
    detect_smart_backend
    case "$SMART_BACKEND" in
        smartctl)
            local sv="${SMARTCTL_VERSION:-未知}"
            log_info "SMART 后端: ${SMARTCTL_SOURCE:-未标注} | smartctl ${sv} | ${SMARTCTL_BIN}"
            [ "$SMARTCTL_FARM_OK" -eq 0 ] \
                && log_info "  该版本 < 7.4，不支持 FARM 日志（希捷出厂装配日期）；通电/健康/SN 不受影响"
            ;;
        native)   log_info "SMART 后端: 内嵌原生 ioctl 探针 (${SMART_PY}) — 免安装、离线可用" ;;
        none)     log_info "SMART 后端: 不可用（${SMART_NATIVE_FAIL_REASON:-未检测到可用后端}），通电信息将留空" ;;
    esac

    echo ""
    log_info "开始收集系统配置数据..."

    resolve_device_sn
    local sys_sn="$DEVICE_SN"
    local hostname_str; hostname_str=$(hostname 2>/dev/null)
    [ -z "$hostname_str" ] && hostname_str="未知"

    # 页签名计算移到写锁之后（锁内读页签号），此处只留不依赖文件状态的标题
    local title_text="硬件配置信息 - ${hostname_str} (SN: ${sys_sn})"

    reset_rows

    add_title "$title_text"
    add_blank

    collect_os_info;      add_blank
    collect_cpu_info;     add_blank
    collect_memory_info;  add_blank
    collect_disk_info;    add_blank
    collect_smart_info;   add_blank
    collect_network_info; add_blank
    collect_gpu_info;     add_blank
    collect_driver_versions; add_blank

    local _perf_disk
    for _perf_disk in "${DISK_DEVICES[@]}"; do
        collect_disk_perf_info "$_perf_disk"
        add_blank
    done
    # 分区表是用户要的"最后底下那一份"：放在可选性能测试段之后、页脚之前
    local part_no="八"
    [ "${#DISK_DEVICES[@]}" -gt 0 ] && part_no="九"
    collect_partition_info "$part_no"
    add_blank

    add_footer "导出时间: $(date '+%Y-%m-%d %H:%M:%S')  |  导出工具: ${SCRIPT_NAME} v${SCRIPT_VERSION}"

    local sheet_data="$DATA_ROWS"
    # 并发写同一 .xls 会互相踩，flock 可用时上互斥锁；缺失则降级为无锁
    # （老系统/精简环境可能没 flock，不能因此拦住导出）
    if command -v flock &>/dev/null; then
        exec 200>"${OUTPUT_FILE}.lock"
        if ! flock -n 200; then
            # 失败者绝不能设 LOCK_FILE：否则 EXIT trap 会把持有者的锁文件从
            # 磁盘提前删掉，后续新实例将创建新 inode 加锁成功，互斥失效
            exec 200>&-
            record_error "无法获取写锁: ${OUTPUT_FILE} 正被其他实例写入"
            summary_errors
            return 1
        fi
        # 故意不用 local：cleanup_smart_runtime 由 EXIT trap 触发时 main 已返回，
        # local 变量届时已销毁，锁文件就删不掉了
        LOCK_FILE="${OUTPUT_FILE}.lock"
    fi

    # 页签名（Excel 上限 31 字符）：必须在锁内读页签号，否则两个并发实例
    # 都读到同一个 N、都写 N+1，后写的覆盖先写的；N>999 时 printf 自然输出
    # 4 位数，配合 max_sn_len 截断仍不会超 31 字符，不做回绕（回绕会撞重名）
    local curr_num; curr_num=$(get_max_sheet_num "$OUTPUT_FILE")
    local new_num=$((curr_num + 1))
    local sheet_prefix; sheet_prefix=$(printf "%03d" "$new_num")
    local safe_sn; safe_sn=$(echo "$sys_sn" | sed 's/[^a-zA-Z0-9_-]/_/g')
    local max_sn_len=$((27 - ${#sheet_prefix}))
    [ ${#safe_sn} -gt "$max_sn_len" ] && safe_sn="${safe_sn:0:$max_sn_len}"
    local sheet_name="${sheet_prefix}-${safe_sn}"
    log_info "页签名称: ${sheet_name}"

    echo ""
    log_info "正在生成 Excel 文件..."

    local ws_xml; ws_xml=$(make_worksheet_xml "$sheet_name" "$sheet_data")

    if ! check_file_writable "$OUTPUT_FILE"; then
        summary_errors
        return 1
    fi

    local out_dir; out_dir=$(dirname "$OUTPUT_FILE")

    if [ ! -f "$OUTPUT_FILE" ]; then
        # mktemp 同目录临时文件 + mv 原子落位：直接 cat > 目标路径时中途失败
        # （磁盘满等）会留下半截文件，下次运行又把它当「已有文件」去追加，
        # 损坏被固化
        local tmp_out
        tmp_out=$(mktemp "${out_dir}/.$(basename "$OUTPUT_FILE").new_XXXXXX" 2>/dev/null)
        if [ -z "$tmp_out" ]; then
            record_error "无法在 ${out_dir} 创建临时文件"
            summary_errors
            return 1
        fi
        # 与普通重定向一致：最终权限按当前 umask 计算（0666 & ~umask）
        chmod "$(printf '%o' "$((0666 & ~0$(umask)))")" "$tmp_out" 2>/dev/null
        cat > "$tmp_out" << XLSHEAD
<?xml version="1.0" encoding="UTF-8"?>
<?mso-application progid="Excel.Sheet"?>
<Workbook xmlns="urn:schemas-microsoft-com:office:spreadsheet"
          xmlns:ss="urn:schemas-microsoft-com:office:spreadsheet">
$(make_xls_styles)
${ws_xml}
</Workbook>
XLSHEAD
        if [ $? -ne 0 ] || ! mv -f "$tmp_out" "$OUTPUT_FILE"; then
            rm -f "$tmp_out" 2>/dev/null
            record_error "创建文件失败: ${OUTPUT_FILE}"
            summary_errors
            return 1
        fi
        log_info "文件已创建: ${OUTPUT_FILE}"
    else
        # 追加前校验结构：非空、含 <Workbook、最后一个非空行恰为 </Workbook>。
        # 被外部改坏/半截的文件直接拒收，否则 sed 删行会把损坏固化
        local tail_line struct_ok=1
        [ -s "$OUTPUT_FILE" ] || struct_ok=0
        grep -q '<Workbook' "$OUTPUT_FILE" 2>/dev/null || struct_ok=0
        tail_line=$(awk 'NF{l=NR} END{print l+0}' "$OUTPUT_FILE" 2>/dev/null)
        if [ "$tail_line" -eq 0 ] 2>/dev/null \
            || [ "$(sed -n "${tail_line}p" "$OUTPUT_FILE" 2>/dev/null)" != "</Workbook>" ]; then
            struct_ok=0
        fi
        if [ "$struct_ok" -ne 1 ]; then
            record_error "输出文件结构异常，拒绝追加（可能已损坏或被外部修改）: ${OUTPUT_FILE}"
            summary_errors
            return 1
        fi

        local backup_file
        backup_file=$(mktemp "${out_dir}/.$(basename "$OUTPUT_FILE").bak_XXXXXX" 2>/dev/null)
        if [ -z "$backup_file" ] || ! cp "$OUTPUT_FILE" "$backup_file" 2>/dev/null; then
            record_error "无法创建备份，拒绝追加: ${OUTPUT_FILE}"
            rm -f "$backup_file" 2>/dev/null
            summary_errors
            return 1
        fi

        # 按校验出的行号删到行尾：'$ d' 只删物理最后一行，若 </Workbook> 后
        # 还有空行，删空行会残留旧闭合标签，追加后结构错乱
        sed -i "${tail_line},\$d" "$OUTPUT_FILE"
        if [ $? -ne 0 ]; then
            record_error "修改文件失败: ${OUTPUT_FILE}"
            cp "$backup_file" "$OUTPUT_FILE" 2>/dev/null
            rm -f "$backup_file" 2>/dev/null
            summary_errors
            return 1
        fi

        # 两次追加必须整体成功才算更新：任一失败时文件缺 </Workbook> 已是损坏状态，
        # 此时要回滚备份并报错，绝不能删备份 + 宣称成功
        if ! { echo "${ws_xml}" && echo "</Workbook>"; } >> "$OUTPUT_FILE"; then
            record_error "追加页签失败: ${OUTPUT_FILE}（磁盘满 / 配额 / IO 错误？）"
            cp "$backup_file" "$OUTPUT_FILE" 2>/dev/null
            rm -f "$backup_file" 2>/dev/null
            summary_errors
            return 1
        fi

        rm -f "$backup_file" 2>/dev/null
        log_info "文件已更新，页签已追加: ${OUTPUT_FILE}"
    fi

    echo ""
    log_info "导出完成！页签: ${sheet_name}"
    echo ""

    summary_errors
    # 有错误就不算成功：批量脚本按退出码判断成败（与 --help 的承诺一致）
    [ "$ERROR_COUNT" -gt 0 ] && return 1
    return 0
}

# =============================================================================
# 执行入口
# 参数已在全局解析完毕，main 不接收位置参数
main
exit $?
