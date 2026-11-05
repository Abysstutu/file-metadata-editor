# 文件源数据编辑

本地 Windows 桌面工具，用于查看、修改或清除常见文件的嵌入元数据。文件处理在本机完成，不会把文件内容上传到网络。

## 下载

打开 [GitHub Releases](https://github.com/Abysstutu/file-metadata-editor/releases)，选择适合的 Windows x64 包：

- **v3.x（推荐）**：新版功能与拖放清除，支持旧版 Office；应用运行组件随 EXE 提供。
- **v2.0**：支持旧版 `.doc`、`.xls`、`.ppt`，内置 Java 运行时。
- **v1.0**：仅支持新版 Office 格式与 PDF。

新版本会同时提供两个构建：

- `FileMetadataEditor-vX.Y-Windows-x64.exe`：单文件便携版，下载后可直接运行，但首次启动需要解包，启动速度较慢。
- `FileMetadataEditor-vX.Y-Windows-x64-FastStart.zip`：快速启动版，解压后运行文件夹中的 EXE；不要单独移动 EXE，需保留同目录运行组件。

两种版本均无需单独安装 Python 或 Java。Windows SmartScreen 或杀毒软件可能对未签名的个人项目程序发出提示；请仅从本仓库 Releases 下载，并在运行前核对发布版本与文件名。

## 支持范围

| 类型 | 扩展名 | 说明 |
| --- | --- | --- |
| Office Open XML | `.docx`、`.xlsx`、`.pptx` 及常见宏/模板变体 | 核心、应用、自定义属性和文件系统时间 |
| 旧版 Office | `.doc`、`.xls`、`.ppt` | 通过 Apache POI 修改标准属性流，不转换文档内容；EXE 内置 Java 运行时 |
| PDF | `.pdf` | Info 与 XMP 元数据 |
| Markdown | `.md`、`.markdown` 等 | YAML Front Matter |
| 图片 | JPEG、TIFF、WebP | 常用 EXIF 字段 |
| 音频 | `.mp3`、`.flac`、`.m4a`、`.wav` | 常用标签 |
| 视频 | `.mp4`、`.mov`、`.mkv` | 需系统可用的 FFmpeg |
| ZIP | `.zip` | 压缩包注释 |

支持单文件编辑、批量清除、按类别批量清除、批量修改、拖放、保存前备份，以及文件只读或占用时另存为可写副本。清除前可预览将处理的属性，处理后可逐项查看原值和结果；单文件和批量操作都会重新读取文件，区分验证通过、检测到残留和无法验证。批量修改可统一填写标题、作者、关键词，或用文件名作为标题；不同格式不支持的字段会跳过。清除和修改操作会改变文件，请先备份重要文件。

输出可选择直接修改原文件，或另存副本；副本可指定输出文件夹和命名规则，`{name}` 表示原文件名。若目标文件名已存在，程序会自动添加序号，避免覆盖。勾选“保存前备份”后，可在本次运行中使用“撤销上次操作”；也可打开目标文件后使用“从备份恢复”选择历史 `.metadata-backup` 文件。恢复前程序还会保存当前状态的备份。关闭备份选项的操作不能通过“撤销上次操作”恢复。

## 限制与故障处理

- 加密 PDF 不支持修改；检测到数字签名的 PDF 会拒绝修改，避免破坏签名。
- 旧版 Office 加密、损坏或受数字签名保护的文件会拒绝处理。
- 只读或被 Office/WPS、同步软件占用时，程序可能无法覆盖原件；请关闭占用程序，或使用自动生成的可写副本。
- 视频属性编辑需要安装 FFmpeg 并确保 `ffmpeg` 可从命令行访问。其他格式和功能不依赖 FFmpeg。
- 修改后可重新打开文件检查属性；不同软件可能缓存属性，关闭后重开或复制到本地目录再确认。

## 从源码运行

需要 Python 3.10+。旧版 Office 文件还需 Java 11+。

```powershell
python -m pip install -r requirements.txt
python app.py
```

## 自动构建

推送形如 `v3.1` 的版本标签后，GitHub Actions 会自动在 Windows x64 runner 上检查代码、生成精简 Java 运行时、构建并冒烟测试单文件 EXE 与快速启动 ZIP，然后创建对应的 GitHub Release 并附上两个下载包。也可以手动发布已有 Release，工作流会构建并上传附件；手动触发构建则只保存为 14 天有效的 Actions 构建产物。运行时随程序提供，下载用户无需单独安装 Java。

## 项目结构

```text
app.py                 程序入口与 Windows 高 DPI 配置
metadata_editor/       元数据读写与 Tkinter 界面
legacy_helper/         Apache POI 旧版 Office 属性兼容组件
requirements.txt       源码运行依赖
run.cmd                Windows 快捷启动脚本
high-dpi.manifest      Windows 高 DPI 清单
.github/workflows/     自动检查、打包与 Release 上传
```

仓库不保存构建输出或源码 ZIP；源码以 Git 历史维护，可执行文件发布在 Releases。

## 许可

暂未声明开源许可证。
