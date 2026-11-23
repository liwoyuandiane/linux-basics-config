# collect_hw_info.sh

**一个 Bash 脚本 + 两个静态二进制，在完全离线的 Linux 机器上采集硬件配置，导出一张能直接粘进资产台账的 Excel 表。**

无需联网、无需安装包、无需编译、不写系统目录。插上 U 盘敲一行命令，得到 `collect_hw_info.xls`。

---

## 它解决什么问题

机房里常有一批**断网、不能装软件、系统版本还各不相同**的机器：银河麒麟桌面 V10、麒麟服务器、CentOS 6/7 混着，CPU 有飞腾、鲲鹏，也有 x86。要盘点 CPU 路数、内存条、硬盘序列号、通电时长、分区挂载，通常得挨台敲十几条命令再手工汇总。

这个脚本把常用数据源（`dmidecode` / `lscpu` / `lsblk` / `smartctl` / `ethtool` / `/proc` / `sysfs`）收进一个 Bash 文件，输出 **Excel XML Spreadsheet 2003 (.xls)**：6 列紧凑表格、彩色样式、合并单元格，重复执行追加新页签；**单文件最多 255 个页签，达到上限会报错并提示备份删除后重试**，不自动分卷。

关键点是 **SMART 数据**：通电时长和通电次数锁在硬盘固件里，内核 sysfs 不暴露，必须发 `ioctl` 才能读到。所以仓库里随附两个静态编译的 `smartctl`，运行时不依赖目标机的 glibc 版本——这正好绕开了"CentOS 6 是 glibc 2.12、Kylin 桌面是 2.31、服务器是 2.28"这种版本错配。

---

## 仓库文件（U 盘里只放这 3 个）

| 文件 | 作用 | 目标架构 |
|---|---|---|
| `collect_hw_info.sh` | 主脚本，采集 + 导出 | 通用 Bash |
| `smartctl.aarch64` | 静态编译的 smartctl 7.5 | ARM64：飞腾 FT-2000/4、腾锐 D2000、鲲鹏 920 |
| `smartctl.amd64` | 静态编译的 smartctl 7.5 | x86_64：CentOS / 普通 PC 服务器 |

> `README.md` 是仓库文档，**不需要拷进 U 盘**。脚本把完整使用说明内置在 `--help-full` 里，单文件离线也能查。
>
> 也不要再带 `.deb` / `.rpm` 安装包或额外的 Python 文件：探针源码已内嵌在 `.sh` 中，运行时生成临时文件、退出即清理；两个静态文件覆盖全部目标架构，命中后完全不需要 python3。

### 已验证的机型

| 操作系统 | CPU | 使用文件 |
|---|---|---|
| 银河麒麟桌面 V10 2101 / 2203 / 2303 | 飞腾 FT-2000/4、腾锐 D2000（`aarch64`） | `smartctl.aarch64` |
| 银河麒麟服务器 V10 2303 / 2403 | 鲲鹏 920（`aarch64`） | `smartctl.aarch64` |
| CentOS 6.7 / 7.6 | x86_64 | `smartctl.amd64` |

其他 Kylin / Ubuntu / Debian / CentOS / Fedora 等主流发行版同样可用；缺失命令会自动跳过，不中断导出。

---

## 两个静态二进制的来源（重要）

**这两个文件不是本项目编译的，也不是 smartmontools 官方发布的**，来自第三方静态构建项目 [perryflynn/static-binaries](https://github.com/perryflynn/static-binaries)。

### 直接下载链接

| 本仓库文件名 | 架构 | 下载直链 | 大小 |
|---|---|---|---|
| `smartctl.aarch64` | ARM64v8 (aarch64) | <https://files.serverless.industries/bin/smartctl.aarch64> | 约 1.9 MB |
| `smartctl.amd64` | x86_64 (amd64) | <https://files.serverless.industries/bin/smartctl.amd64> | 约 1.9 MB |

```bash
curl -fLO https://files.serverless.industries/bin/smartctl.aarch64
curl -fLO https://files.serverless.industries/bin/smartctl.amd64
chmod +x smartctl.aarch64 smartctl.amd64
```

### 相关链接

| 用途 | 链接 |
|---|---|
| 构建源码与 CI 配置 | <https://github.com/perryflynn/static-binaries> |
| 二进制下载目录（含全部架构与 `info.txt`） | <https://files.serverless.industries/bin/> |
| 该项目的 Releases（**sha256 + GPG 校验文件在这里**） | <https://github.com/perryflynn/static-binaries/releases> |
| 项目说明 | <https://serverless.industries/2023/07/02/static-binaries.en.html> |
| ARM 架构恢复构建的公告（2025-11-23） | <https://serverless.industries/2025/11/23/static-binaries-update.en.html> |
| smartmontools 上游官网 | <https://www.smartmontools.org/> |
| smartmontools 官方 GitHub（2025-06-01 起为官方上游仓库） | <https://github.com/smartmontools/smartmontools> |
| 官方下载与安装说明 | <https://www.smartmontools.org/wiki/Download> |
| 官方 Releases / 源码 tarball | <https://github.com/smartmontools/smartmontools/releases> |

### 版本与校验

- 二进制对应 **smartmontools 7.5**（上游 2025-04-30 发布）。
- 该项目为每个工具提供 **sha256 + GPG** 校验，见上表 Releases 页面；下载目录另有 `info.txt` 记录构建信息。
- **本 README 不写死 sha256 值**：下载站会重新构建并覆盖文件（目录时间戳会变），写死的哈希只会过期误导人。建议每次下载后核对，并把确认过的值记进自己的发布台账或 `SHA256SUMS` 文件：

```bash
sha256sum smartctl.aarch64 smartctl.amd64     # 与上游 Releases / info.txt 对照
file smartctl.aarch64 smartctl.amd64          # 应显示 statically linked
ldd  smartctl.amd64                           # 应显示 not a dynamic executable
./smartctl.amd64 --version                    # 能打印版本即可用
```

> 涉密或生产环境建议在受控机器上**自行静态编译**，见下方「自行构建」。

### 为什么不用发行版安装包

[smartmontools 官方](https://www.smartmontools.org/wiki/Download) 对 Linux 只提供源码 tarball，预编译 `.deb` / `.rpm` 是各发行版自己打包的，可按 [pkgs.org](https://pkgs.org/download/smartmontools)、[archive.ubuntu.com](https://archive.ubuntu.com/ubuntu/)（x86）、[ports.ubuntu.com](https://ports.ubuntu.com/ubuntu-ports/)（ARM）等站点检索。

这些包都是**动态链接**，有各自的 glibc 基线：Ubuntu 24.04 的包要 `libc6 >= 2.38`，Ubuntu 22.04 要 `>= 2.34`，RHEL 9 系要 `GLIBC_2.34`。而 Kylin 桌面 V10 是 glibc 2.31、服务器版 2.28、CentOS 7 是 2.17、CentOS 6 是 2.12 —— 高基线编的程序在低版本系统上直接报 `GLIBC_x.xx not found`。

静态链接把依赖打进二进制自身，同一个 `aarch64` 文件就能通吃 2.28 / 2.31 / 更新版本，所以本项目只带两个文件。脚本本身也已移除 deb/rpm 自动解包逻辑，不再需要找包。

### 自行构建（可选，最可控）

```bash
# 在对应架构的构建机上，用官方 7.5 源码
tar xf smartmontools-7.5.tar.gz && cd smartmontools-7.5
./configure LDFLAGS="-static" --disable-shared --disable-nls
make -j"$(nproc)"
ldd src/smartctl        # not a dynamic executable
cp src/smartctl ../smartctl.aarch64     # 或 ../smartctl.amd64
```

构建机需要 `glibc-static libstdc++-static gcc-c++`（RHEL 系）或 `libc6-dev libstdc++-dev g++`（Debian 系）。官方源码 tarball 用 OpenPGP key `FF3AEFF5` 签名，指纹 `0C95 77FD 2C4C FCB4 B9A5 9964 0A30 812E FF3A EFF5`。

**文件名不能随意改**：脚本按 `uname -m` 查表匹配文件名里的架构关键字（`aarch64` / `arm64`、`amd64` / `x86_64` 等），改成 `smartctl-arm` 之类会识别不到。

### 许可说明

- `collect_hw_info.sh` 为本仓库脚本，许可见仓库根目录的 `LICENSE` 文件（MIT）。
- 随附的两个 `smartctl` 二进制由第三方从 [smartmontools](https://github.com/smartmontools/smartmontools) 源码静态编译，smartmontools 使用 **GNU GPL v2**（见上游 [COPYING](https://github.com/smartmontools/smartmontools/blob/main/COPYING)）。本仓库仅做未经修改的再分发。
- 若你二次分发这些二进制，请保留上游许可信息并提供源码获取途径（上面的官方仓库 / 源码 tarball 链接）。

---

## 快速开始

```bash
# 三个文件放同一目录，插到目标机
sudo bash /media/usb/collect_hw_info.sh /media/usb
```

输出 `collect_hw_info.xls`，用 Excel / WPS / LibreOffice 直接打开。

**四条必守**：

1. **必须 `sudo` / root** —— SMART ioctl 需要 `CAP_SYS_RAWIO`，与带不带工具无关。
2. **`chmod +x`** 两个静态文件。
3. **文件名别改**：`smartctl.aarch64` / `smartctl.amd64`。
4. **正常巡检不用删旧 `.xls`** —— 脚本是追加页签；单文件最多 255 个页签，达到上限会报错并提示「备份 → 删除 → 重试」。只有想从 `001` 重新编号、或已触发上限时才主动删。

页签名为 `<3位序号>-<硬盘SN>`，例如 `001-WCC6M84U7JA3`。整机 DMI 序列号读不到时（Kylin/飞腾常回 `Not Set`），自动改用第一块硬盘的 SN，再退回主机名。序号按十进制递增，`008 → 009 → 010 → … → 1000` 均正常（页签名始终 ≤ 31 字符，超出时自动截断 SN 部分）。

### 单文件 255 页签上限（v5.24 起）

脚本只维护**一份** `.xls`（默认 `collect_hw_info.xls`，`-o` 指定时以该名字为准），最多 **255 个页签**（`001~255`）。**不会自动分卷**，不会新建 `collect_hw_info1.xls` 之类的文件。

第 256 次执行时脚本直接**报错退出（exit 1）**，不改原文件、不新建任何文件，日志给出可操作提示：

```
[ERROR] 输出文件页签数已达上限 255（当前 255 个），拒绝继续追加: /media/usb/collect_hw_info.xls
[INFO]  处理办法：先备份 /media/usb/collect_hw_info.xls，再删除 /media/usb/collect_hw_info.xls 文件并重试（新文件从 001 开始编号）
```

按提示**备份 → 删除 → 重跑**即可，新文件从 `001` 重新编号。旧数据留在备份里，不会丢。

为什么设 255：微软官方规格里工作簿的工作表数是「受可用内存限制」，没有硬性个数上限；255 是**本脚本设定的硬上限**，给批量巡检一个明确、可预期的边界，避免单个 `.xls` 无限膨胀（每次追加都要 `cp` 整份文件做备份 + `sed -i` 重写，文件越大越慢、中断时风险越集中）。不分卷是刻意选择：同目录出现多份编号文件、或前部卷被删后序号回退/重号，都会让资产台账错位——宁可明确报错，由人决定何时归档重开。

并发跑同一份 `.xls` 时多个实例会有限等待（最多约 60 秒）后依次写入，不会抢同一个页签号。`.lock` 是 `flock` 的文件载体：进程退出时内核自动释放锁，但文件本身不会被脚本删除（避免 TOCTOU 导致两进程同时进临界区）。残留的 `.xls.lock` 只是个空文件，不影响交付、也不影响下次执行。

> 验收口径（v5.24.3 起）：旧记录里「`.lock` 残留 0」的断言已作废，应改为「允许存在空的 `.xls.lock` 载体，且并发无重号、无半成品 xls」。

> **从 v5.23.x 升级的注意**：目录里若残留 `collect_hw_info1.xls` / `collect_hw_info2.xls`，本版**不会读取、不会修改、也不会删除**，请人工归档。若想继续沿用其中的数据，把要保留的那份改名为基础文件名再执行即可（例如 `mv collect_hw_info2.xls collect_hw_info.xls`）。`HW_MAX_SHEETS` 环境变量已移除，上限固定 255。

### 批量巡检模板

```bash
OUT=/media/usb/$(hostname)
mkdir -p "$OUT"
rm -f "$OUT/hw.xls"
sudo bash /media/usb/collect_hw_info.sh -o "$OUT/hw.xls"
```

---

## 输出内容

**v5.22 起为七节结构：一个模块一个序号，不再有「五（补充）」「七（补充）」这类补充章节。** 系统版本并入第一节、SMART 并入第四节、驱动与软件版本并入第六节；**三个合并节的标题按各节实际采集内容命名**，不沿用合并前的旧名。

| 章节 | 内容 |
|---|---|
| 一、基本信息与系统版本 | 主机名、采集时间、设备序列号；操作系统、架构、内核、BIOS、主板、制造商（原「二、系统版本信息」并入，两组行之间保留空行分隔） |
| 二、CPU 信息 | 路数（单/双/多路）、每路核心、物理核心总数、SMT、NUMA、每颗插槽明细 |
| 三、内存信息 | **厂商（DMI 原文，不推断）**、总容量、内存条数量、**「内存条序列号 (SN)」按插槽逐条汇总行**、类型、插槽容量合计、模组型号 / SN / Rank 明细（同规格合并标数量） |
| 四、硬盘与 SMART 通电信息 | **厂商 / 系列（设备报告英文原文）** / 型号 / 类型·接口 / **SN 独立列**（不做品牌推断；含 NVMe / SATA SSD 六级 SN 回退，来源写在「采集说明」）；同节续 **通电时长 / 通电次数 / 健康 / 温度**、出厂装配日期、SMART 后端来源与版本 |
| 五、网卡信息 | MAC、型号、状态、速率、IP（已过滤虚拟接口） |
| 六、GPU、驱动与软件版本 | 显示控制器 / NVIDIA 卡与驱动版本；**驱动与软件版本**：网卡 / 显示 / 存储 / USB 控制器驱动版本、固件版本、`sshd` 版本 |
| 七、硬盘分区与挂载信息 | 紧凑 6 列树形表（分区 / 大小 / 文件系统 / 挂载点 / 用量 / 标签·UUID），下方汇总压成两行 |
| 七…（`-d` 时） | 每块盘一个「硬盘详细信息与性能测试」章节，从「七」起按 `-d` 顺序顺延（hdparm / dd / fio / ioping） |
| 顺延（`-d` 时） | 分区与挂载表排在最后一块盘之后；单盘为「八」、双盘为「九」 |

**编号约定**：Excel 页签内的报表章节用中文数字「一~七」，`--help-full` 内置说明用 `1.~9.`，两套互不相干；上表章节清单里的「一、」「四、」等均指报表章节。

**标题命名口径**：三个合并节按实际内容取名——第一节含设备身份与系统版本、第四节含硬盘标识与 SMART 通电/健康/出厂、第六节含 GPU 与各控制器驱动及 `sshd` 软件版本；其余四节是单一模块，标题直接对应模块名，不额外加字。`-d` 时新增的「硬盘详细信息与性能测试」按盘逐章顺延，分区表始终排最后。

### 数据边界（避免误读）

- **厂商 / 系列不做自动识别**：旧版有一张「型号前缀 → 中文品牌」规则表（`ST*`→希捷、`HUA*`→日立…），会把国产盘 / OEM 盘张冠李戴（例如「华电通信 HUADIAN COMMUNICATION」的 `H10C-Y` SATA SSD，型号 `DGD67T1S3C27`）。v5.21.8 起**整张规则表已删除**，这一列只显示设备自己报出来的英文原文：`smartctl -i` 的 `Model Family` / `Vendor`，或 `sysfs device/vendor`（排除 `ATA` 这类传输层标识）。型号列本来就是 `lsblk MODEL` / `smartctl Device Model` 的原文，不再去前缀、不翻译成中文。**设备固件没报厂商时只能显示「—」，以盘体标签为准**——这个信息任何程序都读不出来，只能人工录入台账。
- **硬盘 SN（含固态盘 / NVMe）**：六级回退 ① `smartctl -i`（与 SMART 后端解耦：旧版后端走内嵌探针时即使有 smartctl 也不读标识，这是 SATA / NVMe 固态盘漏 SN 的主因；NVMe 另试控制器节点 `/dev/nvmeN`）→ ② `lsblk -o SERIAL`（`-d` 组合失败时去掉 `-d` 重试）→ ③ `udevadm` → ④ `sysfs`（`device/serial`；NVMe **只取本盘推导出的** `/sys/class/nvme/nvmeN/serial`，不遍历其他控制器，避免多 NVMe 串号）→ ⑤ SCSI VPD 0x80 → ⑥ `nvme id-ctrl`。**④⑥ 只对 NVMe 生效；SATA SSD 主要靠 ①②③⑤**。不用 WWN 冒充；读不到显示「未提供」，以盘体标签为准，并按「采集说明」提示手工核对 `smartctl -i` / `smartctl -d sat -i`。
- **出厂日期**：ATA / NVMe / SMART 标准里**没有**生产日期字段。FARM 是**希捷厂商私有码**（私有日志页 `0xA6`，跨厂商的 Device Statistics 日志 `0xB1` 里没有 Assembly Date），仅部分希捷盘可读 `Assembly Date (YYWW)` 原始值（需 smartctl ≥ 7.4；Exos 2018 年后基本支持，消费级看年份，USB 桥可能不透传），其余凭 SN 查厂商官网或看盘体标签。装配日期 ≠ 购买日期 ≠ 首次使用日期。
  **v5.24.2 起原始值原样写入、不做任何解读**：该私有码编码方式不统一，脚本不猜年周、不换算世纪、不做任何判断，读到什么写什么（如 `2602` 就显示 `2602`）。年周含义与准确生产日期以盘体标签 `DOM`（Date of Manufacture，精确到日）为准，标签始终比 FARM 权威。
- **虚拟机 / 云块存储**（如 Oracle BlockVolume）不提供物理通电计数器，属预期行为，Excel 会写明原因且**不计警告**。
- **驱动版本**：内核内置驱动常不声明 `version` 字段，显示「未提供」不是采集失败；形如 `name: xxx` 的 `modinfo` 字段行（来源未定案）也会被形态校验拦下，不会当成版本号。
- **内存厂商不做自动识别**：v5.22.0 起删除旧版「模组型号前缀 → 中文品牌」推断（`KVR*`→金士顿、`M39*`→三星、JEDEC 厂商码映射），这一节只读 `dmidecode -t memory` 的 `Manufacturer` 原文——含 `0x802C` 这类厂商码也原样保留、不翻译、不映射成中文品牌。BIOS 没填或填 `Unknown` / `<BAD INDEX>` 时写「未提供（BIOS 未填）」，以内存条物理标签为准。与硬盘厂商/系列同一口径。
- **内存 SN**：`dmidecode -t memory` 的 `Serial Number`，「内存条序列号 (SN)」行按插槽逐条列出（明细表第 6 列仍保留 `模组型号 / SN / Rank`）。BIOS 填 `00000000` / `<BAD INDEX>` 等占位值时归一为「—」并说明「以内存条标签为准」，不是脚本没采。
- **日志口径**：`[WARN]`/`[ERROR]` 行数与「执行结果汇总」计数严格一致；写测试路径提示等说明性信息走 `[INFO]`，不计入警告。
- **sshd 版本**：老版 OpenSSH（CentOS 6 的 5.3p1、Kylin V10 的 8.2p1）不支持 `sshd -V`，脚本依次回退「服务端二进制版本串 → `ssh -V` → `rpm`/`dpkg-query`」，并在「来源」列写明用了哪一级。
- **运行环境口径**：脚本按 **GNU coreutils / GNU grep / GNU sed** 编写与验证（目标机型银河麒麟、CentOS、Ubuntu、Debian、Fedora 均为 GNU 环境）。BusyBox / 嵌入式精简环境**未在支持范围内**，其 `grep` 的词边界与 `sed` 的换行处理差异可能让 GPU 类型识别或单元格文本出现偏差；本版不对这些环境做代码级兼容承诺。

---

## 命令行

```
用法: collect_hw_info.sh [选项] [输出目录]

  -h, --help                    简短用法
  -H, --help-full               完整说明（内置，单文件离线可查）
  -v, --version                 版本号
  -d, --disk <设备>             附带硬盘性能测试，可多次指定（自动去重），如 -d /dev/sda -d /dev/sdb
  -s, --size <大小>             测试文件大小，默认 10G（K/M/G，大小写均可）
  -p, --test-path <目录>        测试文件写入目录，默认 /tmp
  -o, --output <文件>           指定输出 .xls（优先于位置参数，只能给一次）
  --                            结束选项解析
```

示例：

```bash
sudo ./collect_hw_info.sh                          # 当前目录生成
sudo ./collect_hw_info.sh /mnt/usb                 # 指定输出目录
sudo ./collect_hw_info.sh -o /tmp/机房A-01.xls     # 指定输出文件名
sudo ./collect_hw_info.sh -d /dev/sda -d /dev/sdb -s 5G   # 多盘性能测试（每盘单独成章）
./collect_hw_info.sh -H                            # 查看完整文档，无需联网
```

参数 / 路径类型错误统一 `exit 2`（含 `-o` 指向目录、设备、FIFO、socket、符号链接），方便批量脚本判断成败。

### 环境变量

| 变量 | 作用 |
|---|---|
| `SMARTCTL_BIN=/path/to/smartctl` | 指定任意路径的 smartctl，优先级最高 |
| `SMART_PY=/usr/bin/python3` | 指定 python3 解释器路径；不可执行或自检未通过时自动回退系统 python3 |
| `SMART_FORCE_NATIVE=1` | 跳过全部 smartctl 候选，强制用内嵌 ioctl 探针（仅用于对照测试） |

---

## SMART 四级后端回退

| 优先级 | 来源 | 说明 |
|---|---|---|
| ① | `SMARTCTL_BIN` / 同目录 `smartctl.<架构>` | **当前部署命中这一级** |
| ② | 系统 PATH 里的 `smartctl` | 静态文件缺失时兜底，例如 Kylin 自带的 7.1 |
| ③ | 内嵌 Python 原生 ioctl 探针 | 零安装零携带，仅依赖系统自带 python3 |
| ④ | 全部不可用 | 通电信息留空，其余硬件信息照常导出 |

每个候选都先跑 `--version` 启动预检，架构不符或 glibc 过旧自动跳过试下一个。实际命中的**来源级别 + 版本 + 路径**会写进日志和 Excel，便于多机批量核对；版本 < 7.4 时会另给一行提示（不支持 FARM 出厂日志）。

日志与退出码：`[WARN]` / `[ERROR]` 行数与「执行结果汇总」计数严格一致；候选跳过、可选工具缺失、写测试路径提示等用 `[INFO]`，不计入警告（详见「数据边界」）。参数 / 路径类型错误（`-o` 指向目录、设备、FIFO、socket、符号链接）`exit 2`、采集/写盘/加锁等硬错误 `exit 1`、性能项未测成 `exit 0`。

---

## 常见问题

**Q：不装 `smartmontools` 真的能读通电时间吗？**
能。两个静态文件就是 `smartctl` 本尊，绿色免运行安装；脚本只是把它从 U 盘调起来，不改系统、不动包管理器数据库。

**Q：为什么必须 root？**
SMART 数据要发 `ioctl`（ATA PASS-THROUGH / NVMe Get Log Page），普通用户会被内核拒绝。这与装不装工具无关。

**Q：厂商 / 系列列显示「—」，是脚本没采到吗？**
不是。多数国产盘 / OEM 盘的固件只报型号、SN、固件版本，不报厂商或系列字段，程序无从读出。查看「四、硬盘与 SMART 通电信息」末尾「采集说明」里的「厂商/系列」来源即可确认；品牌请以盘体标签为准（例如标签上印的「华电通信 HUADIAN COMMUNICATION」）。脚本刻意不按型号前缀猜品牌，避免误判。

**Q：固态盘显示「未提供」SN？**
先看「四、硬盘与 SMART 通电信息」末尾「采集说明」里的 SN 来源，再手工核对 `sudo smartctl -i /dev/xxx`（经 USB/SATA 桥接时加 `-d sat`）。
- **SATA SSD**（如国产固态盘）：v5.21.6 起 `smartctl -i` 已与 SMART 后端解耦，旧版「后端走内嵌探针 → 完全不读标识」的漏采已修；另补了 `lsblk` 无 `-d` 重试、`udevadm` 属性扩展、VPD 0x80 多路径。若仍读不到，多为固件/桥接不导出 SN，以盘体标签为准。
- **NVMe SSD**：补了控制器节点 `/dev/nvmeN`、`/sys/class/nvme/nvmeN/serial` 与 `nvme id-ctrl`。多 NVMe 机器上只取本盘推导出的控制器，不会拿别的盘的 SN。

**Q：连续采集几十台后会不会生成一个巨大的 .xls？**
会一直追加到 **255 个页签**为止，不会自动拆成多个文件。第 256 次执行时脚本报错退出并提示「备份 → 删除 → 重试」，按提示做即可，新文件从 `001` 重新编号。详见上方「单文件 255 页签上限」。已有的 `001~008` 旧文件直接换新版继续跑即可，不用删、不用重采。

**Q：为什么我手工把 .xls 压成单行后，脚本说「输出文件结构异常，拒绝追加」？**
追加前会校验「最后一个非空行恰为 `</Workbook>`」，被外部压缩成单行的 XML 不满足这个条件，脚本保守拒收、不损坏原文件。脚本自己产出的始终是多行 XML，正常巡检不会触发；遇到这种情况把文件恢复成多行、或删除后重新生成即可。

**Q：报表里出厂日期显示「1962 年第 20 周（原始值 2602）」，但盘体标签写的是 2026 年 1 月？**
那是 **v5.24.0 及以前脚本的解析 bug，不是这块盘的数据**：旧版对 FARM 私有码做了自行换算，得出了明显错误的年份。**v5.24.2 起不再解读**，这一行只显示设备报出的原始值 `2602`。**准确生产日期请看盘体标签 `DOM`**：它精确到日（例如 `DOM: 08JAN2026`），始终比 FARM 权威；标签不可见时凭 SN 到厂商官网查询。

**Q：连续采集到第 9 台，报 `008: value too great for base` 然后脚本直接退出？**
这是 v5.22.3 及以前的 bug：页签序号是零填充字符串（`008`），旧版把它直接交给 Bash 算术 `$((...))`，`008` / `009` 被按八进制解析，8/9 不是合法八进制数字 → 致命错误，脚本当场退出、那台机器没有产出；`010` 之后虽不报错但会被当成八进制值（010→8），序号错乱。
**v5.22.4 已修**：页签序号统一按十进制归一后再递增。**已有的 `001~008` 页签数据不会丢** —— 直接换上新脚本、对同一份 `.xls` 继续跑即可，下一页签就是 `009`，无需删文件重建。

**Q：用新版往旧版生成的 .xls 里追加页签，章节号会不会乱？**
会。v5.21.x 及以前是十节编号（系统版本单独成章、SMART 与驱动各占一个「补充」章节），v5.22 起是七节，同一文件里两套编号并存容易误读。脚本检测到旧版页签时会打一行 `[INFO]`，并在新页签顶部写「版本提示」行；按提示**删除旧 .xls 重新生成**，或用 `-o` 另存新文件即可。不会静默混编。

**Q：某块盘显示「未提供」SN 或「未获取到」通电信息？**
常见原因：USB 硬盘盒 / 硬件 RAID 卡不透传 SMART 命令、虚拟磁盘、非 root 运行。以盘体标签为准。

**Q：CentOS 6.7 会不会跑不起来？**
它的内核 2.6.32 比其他机器老得多。铺开前先单跑一次 `./smartctl.amd64 --version`，能打印版本号即可；万一失败，只这台改用系统自带 smartctl 或内嵌探针，其余机器不受影响。

**Q：想加其他软件版本（nginx / openssl）怎么办？**
在 `collect_driver_versions()` 末尾照 `sshd` 那行继续 `add_rowN`（六列对齐）即可，表格结构不用改。

**Q：为什么输出是 `.xls` 不是 `.xlsx`？**
Excel XML Spreadsheet 2003 是纯 XML，Bash 直接拼字符串就能生成，不需要任何库，也不需要目标机装 Python 的 Excel 模块。Excel / WPS / LibreOffice 都能直接打开。

---

## 致谢

- [smartmontools](https://github.com/smartmontools/smartmontools) — SMART 数据读取与硬盘数据库，GNU GPL v2
- [perryflynn/static-binaries](https://github.com/perryflynn/static-binaries) — 多架构静态 `smartctl` 构建
- `dmidecode` / `lscpu` / `lsblk` / `ethtool` / `iproute2` / `pciutils` 等系统工具

---

当前脚本版本：**v5.24.3**（随时可用 `./collect_hw_info.sh -v` 确认）
