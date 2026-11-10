# 文件源数据编辑

本地 Windows 桌面工具，用于查看、修改或清除常见文件的嵌入元数据。文件处理在本机完成，不会把文件内容上传到网络。

## 下载

打开 [GitHub Releases](https://github.com/Abysstutu/file-metadata-editor/releases)，选择适合的 Windows x64 包：

- **v4.1（推荐）**：安装版，沿用 v4.0 功能，安装后可从桌面快捷方式快速启动。
- **v4.0**：黑白桌面界面；单文件与批量修改历史统一保存在本地记录中，清除记录可在文件哈希匹配时恢复清除前的可编辑属性。
- **v3.5**：沿用 v3.4 的格式支持，改进桌面界面排布，并换用黑白直线 M 图标。
- **v3.4**：增强 Word、Excel、PowerPoint、图片、音视频等格式的元数据编辑，并提供各格式测试报告。
- **v3.3**：支持拖放、批量处理和旧版 Office，并随包提供所需运行组件。
- **v2.0**：支持旧版 `.doc`、`.xls`、`.ppt`，内置 Java 运行时。
- **v1.0**：仅支持新版 Office 格式与 PDF。

从 v4.1 开始，GitHub Releases 只发布 `FileMetadataEditor-vX.Y-Windows-x64-Setup.exe` 安装程序。它采用快速启动版的文件夹结构，安装时创建桌面和开始菜单快捷方式，并可从 Windows“已安装的应用”中卸载。默认安装在当前用户目录，无需管理员权限；卸载时会保留用户的修改历史记录。

历史版本 v4.0 提供两个 Windows x64 下载包：

- `FileMetadataEditor-vX.Y-Windows-x64.exe`：单文件便携版，下载后可直接运行，但首次启动需要解包，启动速度较慢。
- `FileMetadataEditor-vX.Y-Windows-x64-FastStart.zip`：快速启动版，解压后运行文件夹中的 EXE；不要单独移动 EXE，需保留同目录运行组件。

v4.0 的两种下载包均无需单独安装 Python、Java、FFmpeg 或 ffprobe。其他版本包含的组件请以对应 Release 说明为准。Windows SmartScreen 或杀毒软件可能对未签名的个人项目程序发出提示；请仅从本仓库 Releases 下载，并在运行前核对发布版本与文件名。

## 支持范围

| 类型 | 扩展名 | 说明 |
| --- | --- | --- |
| Office Open XML | `.docx`、`.xlsx`、`.pptx` 及常见宏/模板变体 | 核心、应用、自定义属性和文件系统时间 |
| 旧版 Office | `.doc`、`.xls`、`.ppt` | 通过 Apache POI 修改标准属性流，不转换文档内容；EXE 内置 Java 运行时 |
| PDF | `.pdf` | Info 与 XMP 元数据 |
| Markdown | `.md`、`.markdown` 等 | YAML Front Matter |
| 图片 | JPEG、TIFF、WebP | 常用 EXIF 字段 |
| 音频 | `.mp3`、`.flac`、`.m4a`、`.wav` | 常用标签 |
| 视频 | `.mp4`、`.mov`、`.mkv` | 发布版内置 FFmpeg 与 ffprobe |
| ZIP | `.zip` | 压缩包注释 |

支持单文件编辑、批量清除、按类别批量清除、批量修改、拖放、修改历史记录，以及文件只读或占用时另存为可写副本。清除前可预览将处理的属性，处理后可逐项查看原值和结果；单文件和批量操作都会重新读取文件，区分验证通过、检测到残留和无法验证。批量修改可统一填写标题、作者、关键词，或用文件名作为标题；不同格式不支持的字段会跳过。清除和修改操作会改变文件，请自行保留重要文件的原件。

Word `.doc/.docx/.docm/.dotx/.dotm` 的 18 个预设字段及命名自定义属性已完成写入、读回和清除测试。详见 [Word 能力表](docs/word-capability/metadata-capability-report.md) 和 [50 份样本回归结果](docs/word-corpus-results.json)。

Excel `.xls/.xlsx/.xlsm/.xltx/.xltm` 也完成相同范围的验证。详见 [Excel 能力表](docs/excel-capability/metadata-capability-report.md) 和 [50 份样本回归结果](docs/excel-corpus-results.json)。

PowerPoint 的 `.ppt/.pptx/.pptm/.potx/.potm/.ppsx/.ppsm` 完成 [逐项能力测试](docs/powerpoint-capability/metadata-capability-report.md) 和 [70 份样本回归](docs/powerpoint-corpus-results.json)。图片常用 EXIF 字段支持中文标题、作者、描述与 GPS 坐标；MP3 描述、视频专辑和类型已可修改。PDF Info 自定义键、Markdown 顶层普通标量自定义键也支持逐项编辑。查看 [35 种格式的完整能力表](docs/full-capability/metadata-capability-report.md) 与 [350 份样本的清除回归结果](docs/full-regression/regression-summary.json)。

输出可选择直接修改原文件，或另存副本；副本可指定输出文件夹和命名规则，`{name}` 表示原文件名。若目标文件名已存在，程序会自动添加序号，避免覆盖。开启“记录修改历史”后，每次操作都会在成功写入后向同一个本地 `edit-history.jsonl` 文件追加一条记录，包含操作前和操作后的元数据、结果文件 SHA-256、文件名与路径、验证状态。Windows 默认位置为 `%LOCALAPPDATA%\FileMetadataEditor\edit-history.jsonl`；点击“查看修改历史”可搜索记录、查看前后属性并复制哈希，也可删除选中记录或清空全部记录。删除历史不会改动原文件，但被删除的清除记录无法再用于恢复。清除记录可点击“恢复清除前属性”：程序先以清除后的 SHA-256 确认目标文件仍是同一份，再写回清除前的可编辑属性与文件系统时间，并重新读取验证。旧版仅保存操作后数据的历史记录无法恢复；文件若在清除后又被修改，也会因哈希不同而被拒绝。历史记录不包含文件正文，也不能恢复不受本程序编辑能力覆盖的私有或结构属性。程序不再自动创建 `.metadata-backup` 完整文件副本。

## 限制与故障处理

- 加密 PDF 不支持修改；检测到数字签名的 PDF 会拒绝修改，避免破坏签名。
- 旧版 Office 加密、损坏或受数字签名保护的文件会拒绝处理。
- 只读或被 Office/WPS、同步软件占用时，程序可能无法覆盖原件；请关闭占用程序，或使用自动生成的可写副本。
- 从源码运行时，视频功能需要系统提供 `ffmpeg` 和 `ffprobe`；也可分别设置 `FILE_METADATA_FFMPEG`、`FILE_METADATA_FFPROBE` 为可执行文件路径。发布版已内置，无需另外安装。
- WAV 支持 RIFF INFO 标签的读取、编辑和清除。清除后的验证会排除旧式 Office 编辑时长的零值、TIFF 必需的结构标签以及视频容器自动生成的技术标识。
- ZIP 文件数量和成员信息是压缩包结构，不提供属性编辑；图片非预设 EXIF、音视频流级与厂商私有标签仍需按具体格式继续适配。Markdown 多行及嵌套 YAML 保留原样，不支持逐项修改。
- 修改后可重新打开文件检查属性；不同软件可能缓存属性，关闭后重开或复制到本地目录再确认。

## 从源码运行

需要 Python 3.10+。旧版 Office 文件还需 Java 11+。

```powershell
python -m pip install -r requirements.txt
python app.py
```

## 第三方组件

视频运行组件采用 [BtbN 的 LGPL shared 版本](https://github.com/BtbN/FFmpeg-Builds/releases/tag/autobuild-2026-10-31-12-59)，下载包内保留其 `LICENSE.txt`。对应的 [FFmpeg 源码](https://github.com/FFmpeg/FFmpeg)和[打包脚本](https://github.com/BtbN/FFmpeg-Builds)可在各自仓库查看。

## 项目结构

```text
app.py                 程序入口与 Windows 高 DPI 配置
metadata_editor/       元数据读写与 Tkinter 界面
legacy_helper/         Apache POI 旧版 Office 属性兼容组件
scripts/               Windows 视频运行组件下载与校验脚本
requirements.txt       源码运行依赖
run.cmd                Windows 快捷启动脚本
high-dpi.manifest      Windows 高 DPI 清单
```

仓库不保存构建输出或源码 ZIP；源码以 Git 历史维护，可执行文件发布在 Releases。

## 许可

暂未声明开源许可证。
