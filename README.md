<div align="center">

# vivid/stasis 全成绩导出工具

**把存档里的全部成绩，一键变成能筛选、能统计、能分享的表格。**

[![Platform](https://img.shields.io/badge/platform-Windows%2010%20%7C%2011-0078D6?style=flat-square)](#兼容性)
[![PowerShell](https://img.shields.io/badge/PowerShell-5.1-5391FE?style=flat-square)](#兼容性)
[![Dependencies](https://img.shields.io/badge/dependencies-none-brightgreen?style=flat-square)](#它是怎么做到的)
[![License](https://img.shields.io/badge/license-GPL--3.0-blue?style=flat-square)](#许可证)

[快速开始](#快速开始) · [输出说明](#excel-里有什么) · [常见问题](#常见问题) · [English](#english)

</div>

---

## 这是什么

一个**纯 PowerShell** 写的导出工具，读取 [vivid/stasis](https://store.steampowered.com/app/2093940/) 的本地存档，把**全部谱面成绩**导出成 Excel / CSV / JSON。

不需要安装 Python、.NET SDK、Office，也不需要管理员权限：Windows 10 / 11 自带的 Windows PowerShell 5.1 就能跑。

**只读存档，不改数据，不联网。**

---

## 功能特性

- **一键导出**：双击 `导出成绩.bat`，几秒内得到 `.xlsx` + `.csv` + `.json` 三份文件。
- **全量成绩**：官方谱面、自制谱面、段位挑战（Course）一次导出，不遗漏。
- **自动补全**：曲名、曲师、难度、等级、定数、谱师、音符数、BPM 全部对上，不用自己查表。
- **双计分体系**：同时给出游戏内「分数」（满分 1,010,000）与「EX 分数 / EX 率」，并自动换算评级。
- **可统计**：`概况` 汇总 Rating、总分数、平均 EX、FC/AC 数量；`难度统计` 按难度聚合。
- **零依赖**：不依赖 Excel COM、不依赖第三方库，xlsx 由脚本直接生成。
- **零安装**：解压即用，删除即卸载，不写注册表、不装服务。
- **路径友好**：中文、空格等路径都能正常处理；CSV 带 UTF-8 BOM，Excel 打开不乱码。
- **失败降级**：找不到游戏目录时仍可导出分数，只是曲名等资料留空并给出提示。

---

## 快速开始

### 1. 下载

```bash
git clone https://github.com/<your-name>/vivid-stasis-score-exporter.git
```

或直接下载仓库 ZIP 并解压到任意位置。

### 2. 运行

关闭游戏（建议），双击：

```
导出成绩.bat
```

### 3. 拿结果

脚本会自动定位存档与游戏目录，并在同目录生成 `导出结果` 文件夹：

```
导出结果/
├─ vivid-stasis-成绩_20260601_213000.xlsx    ← Excel，6 个工作表
├─ vivid-stasis-成绩_20260601_213000.csv     ← UTF-8 BOM，可直接用 Excel 打开
└─ vivid-stasis-成绩_20260601_213000.json    ← 结构化数据，方便二次开发
```

文件名带时间戳，多次导出不会互相覆盖。

> 首次运行若弹出 SmartScreen「已保护你的电脑」，点「更多信息 → 仍要运行」即可。脚本没有任何联网行为，可自行用记事本审查。

---

## 命令行用法

```powershell
# 默认：自动查找存档与游戏目录，导出到脚本旁的“导出结果”
.\Export-VividStasisScores.ps1

# 指定导出目录并自动打开
.\Export-VividStasisScores.ps1 -OutDir D:\Scores -Open

# 只导官方谱面，不生成 JSON
.\Export-VividStasisScores.ps1 -NoCustomSongs -NoJson

# 游戏或存档不在默认位置
.\Export-VividStasisScores.ps1 `
  -SaveDir "D:\Games\VIVIDSTASIS" `
  -GameDir "D:\SteamLibrary\steamapps\common\vividstasis"
```

| 参数 | 说明 |
| --- | --- |
| `-SaveDir <路径>` | 存档目录，默认 `%LOCALAPPDATA%\VIVIDSTASIS` |
| `-GameDir <路径>` | 游戏安装目录，默认通过 Steam 注册表与库文件夹自动查找 |
| `-OutDir <路径>` | 导出目录，默认脚本所在目录下的 `导出结果` |
| `-NoCustomSongs` | 不导出自制谱面成绩 |
| `-NoJson` | 不生成 JSON 文件 |
| `-Open` | 导出完成后打开导出目录 |

---

## Excel 里有什么

### 工作表

| 工作表 | 内容 |
| --- | --- |
| **概况** | 玩家昵称、Rating、游戏版本、成绩总数、总分数、平均 EX 率、FC/AC 数量、最高分等 |
| **全部成绩** | 官方 + 自制全部成绩，可按表头筛选、排序 |
| **官方谱面** | 仅官方谱面 |
| **自制谱面** | 仅自制谱面 |
| **难度统计** | 各难度的谱面数、平均分数、平均 EX 率 |
| **段位挑战** | Course 模式各段位最高总分 |

### 字段

| 列 | 含义 |
| --- | --- |
| 类型 | `官方` / `自制` |
| 曲名 / 曲师 | 曲目信息（官方取自 `song_information.bin`，自制取自 `info.json`） |
| 难度 | `OPENING` / `MIDDLE` / `FINALE` / `ENCORE`，以及自制或追加难度 |
| 等级 / 定数 / 谱师 | 谱面等级显示值、定数、谱师名义 |
| **分数** | 游戏内「准确度得分」，满分 **1,010,000** |
| **评级** | VS / V+ / V / SS+ / SS / S+ / S / AA / A / B / C / D / E |
| **EX 分数** | 每音符最多 3 分累加得到的 EX 分 |
| **EX 率** | EX 分数 ÷ (3 × 音符数) × 100% |
| **EX 评级** | MAX / PLATINUM / GOLD / SILVER / BRONZE / GREAT / GOOD |
| 最大连击 / 音符数 | 该成绩的最大连击、谱面音符总数 |
| 通关 | `AC` / `FC` / `Cleared` / `EX Only`（仅记录 EX、分数为 0，常见于带 mod 的谱面） |
| 谱面 ID / 歌曲 ID / BPM | 内部标识，便于核对与二次处理 |

### 示例（字段预览）

| 类型 | 曲名 | 难度 | 等级 | 定数 | 分数 | 评级 | EX 率 | 通关 |
| --- | --- | --- | --- | --- | --- | --- | --- | --- |
| 官方 | … | MIDDLE | 12 | 12.3 | 1,003,957 | SS | 96.12% | FC |
| 自制 | … | FINALE | 13 | 13.4 | 998,120 | S+ | 93.40% | Cleared |

### 分数评级阈值

| 评级 | 分数 |
| --- | --- |
| VS | 1,010,000 |
| V+ | 1,009,000 – 1,009,999 |
| V | 1,008,000 – 1,008,999 |
| SS+ | 1,004,000 – 1,007,999 |
| SS | 1,000,000 – 1,003,999 |
| S+ | 990,000 – 999,999 |
| S | 980,000 – 989,999 |
| AA | 950,000 – 979,999 |
| A | 900,000 – 949,999 |
| B | 850,000 – 899,999 |
| C | 800,000 – 849,999 |
| D | 600,000 – 799,999 |
| E | 0 – 599,999 |

### EX 评级阈值

`MAX` = 100%，`PLATINUM` ≥ 99%，`GOLD` ≥ 98%，`SILVER` ≥ 96%，`BRONZE` ≥ 93%，`GREAT` ≥ 90%，`GOOD` ≥ 80%，低于 80% 记为 `—`。

---

## 数据来源

| 数据 | 来源 |
| --- | --- |
| 官方谱面成绩 | `%LOCALAPPDATA%\VIVIDSTASIS\highscore_table` |
| 自制谱面成绩 | `%LOCALAPPDATA%\VIVIDSTASIS\custom_highscore` |
| 昵称 / Rating | `%LOCALAPPDATA%\VIVIDSTASIS\profile` |
| 曲名 / 曲师 / 等级 / 定数 / 谱师 | `<游戏目录>\song_information.bin` |
| 音符数等谱面统计 | 存档 `Charts\*\*.stats`（优先）与 `<游戏目录>\Charts\*\*.stats` |
| 自制曲目资料 | `<游戏目录>\Custom Songs\**\info.json`、`shatterinfo.json` |
| 自制谱面音符数 | 存档 `Custom Songs\**\*.stats_custom` |

以下细节值得留意：

- 游戏只为每张谱面保存**最高分**，因此导出的是最佳成绩，而非每一局的历史记录。
- 「分数」与「EX 率」是两套独立计分：分数上限 1,010,000，EX 率按每音符最多 3 分归一化。
- `FC` / `AC` 由存档中的 `COMBO` 字段（0 / 1 / 2）推断，仅供参考；若与游戏内显示不一致，以游戏为准。
- 极少数没有曲目资料的曲目会显示为 `未知曲目 #编号`，但成绩本身仍会完整导出。

---

## 它是怎么做到的

- 存档是 GameMaker 风格的 INI 文本，直接按行解析即可，无需修改。
- `song_information.bin` 是自定义二进制格式（`VSD` 魔数 + token 流），脚本内置了轻量解析器，提取曲名、曲师、等级、定数与谱师。
- 谱面音符数来自 `.stats` / `.stats_custom`，存档目录的版本优先（游戏更新后会重新生成）。
- `.xlsx` 由脚本按照 OOXML 规范直接拼装 zip 包生成，**不依赖 Excel 或 COM 组件**。
- 所有文件以只读方式打开；输出统一写入 `导出结果`，不触碰游戏存档。

> 在 v6.2.2.2 上实测：导出结果与存档原始数据逐条一致，数百条成绩约 5 秒完成。

---

## 兼容性

| 项目 | 要求 |
| --- | --- |
| 操作系统 | Windows 10 / 11 |
| 运行环境 | Windows PowerShell 5.1（系统自带） |
| 游戏版本 | Steam 版 vivid/stasis（新版本若改动存档结构，可能需要更新解析逻辑） |
| 其他依赖 | 无 |

**已知限制**

- 不导出在线数据：Rating 榜、Steam 排行榜、VS Online 成绩不在本地 `highscore_table` 中。
- 不支持 `otherside`：它是独立程序与数据，存档位置与格式不同。
- 不保留每一局成绩，游戏本身只存最高分。

---

## 常见问题

<details>
<summary><b>提示“未找到 vivid/stasis 存档目录”？</b></summary>

请确认游戏至少运行过一次，并存在 `%LOCALAPPDATA%\VIVIDSTASIS\highscore_table`；也可以用 `-SaveDir` 手动指定。
</details>

<details>
<summary><b>提示“未找到游戏安装目录”？</b></summary>

工具仍能导出分数，但曲名、等级、谱师等信息会缺失（显示 `未知曲目 #编号`）。用 `-GameDir` 指向 `...\steamapps\common\vividstasis` 即可补全。
</details>

<details>
<summary><b>CSV 用 Excel 打开乱码？</b></summary>

`.xlsx` 不会乱码。若坚持用 Excel 打开 `.csv`，请用「数据 → 从文本/CSV」并选择 **UTF-8** 编码。
</details>

<details>
<summary><b>双击 bat 后窗口一闪而过？</b></summary>

说明脚本报错了。在该文件夹按住 Shift 右键 →「在此处打开 PowerShell 窗口」运行：

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File .\Export-VividStasisScores.ps1
```

即可看到完整错误信息。
</details>

<details>
<summary><b>会改动我的存档吗？</b></summary>

不会。脚本对所有存档文件只读；建议仍按习惯定期备份 `%LOCALAPPDATA%\VIVIDSTASIS`。
</details>

<details>
<summary><b>杀毒软件报毒 / 被 SmartScreen 拦截？</b></summary>

脚本只做「读文件 + 生成 xlsx/csv/json」，无网络访问、无注册表写入、无进程注入。可直接用记事本打开 `.ps1` 审查，或对文件「右键 → 属性 → 解除锁定」。
</details>

---

## 项目结构

```
vivid-stasis-score-exporter/
├─ Export-VividStasisScores.ps1   主程序（UTF-8 with BOM）
├─ 导出成绩.bat                   双击入口
├─ README.md                      本说明
├─ LICENSE                        GPL-3.0 许可证全文
└─ 导出结果/                      运行后自动生成
```

## 参与贡献

欢迎 Issue 与 PR，尤其是：

- 新游戏版本导致存档结构变化时的兼容修复；
- 自制谱面 / 特殊难度的曲目资料补全；
- 新的导出格式（如 Markdown 成绩单、图片卡片）。

提交前请说明游戏版本，并附上一小段可复现的存档样例（注意隐去个人信息）。提交 PR 即表示你同意以 GPL-3.0 授权你的贡献。

## 许可证

本项目采用 [GNU General Public License v3.0](https://www.gnu.org/licenses/gpl-3.0.html)（GPL-3.0）发布，完整条款见仓库根目录的 [`LICENSE`](LICENSE)。

这意味着你可以自由地使用、修改和分发本工具（包括商业使用），但在分发修改后的版本时需要遵守 copyleft 要求：

- 保留版权声明与许可证声明；
- 提供对应的完整源代码；
- 衍生作品同样以 GPL-3.0 授权。

如果你只是自己用、或在视频/直播里展示导出的成绩，不受任何额外限制。

## 免责声明

本工具为玩家自制，与 vivid/stasis 开发组及发行商无关。请自行承担使用风险，并定期备份存档。

Read-only and offline. Windows only. Licensed under [GPL-3.0](LICENSE).
