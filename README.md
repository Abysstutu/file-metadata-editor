# 文件源数据编辑

本地 Windows 桌面工具，用于查看、修改或清除新版 Office 文档与 PDF 的嵌入元数据。文件处理在本机完成，不会上传到网络。

## 支持范围

- Office Open XML：`.docx`、`.xlsx`、`.pptx`，以及常见宏与模板变体
  - 核心属性、应用属性、自定义属性
  - Windows 创建时间、修改时间
- PDF：Info 与 XMP 元数据
- 批量清除嵌入属性
- 保存前自动备份；文件只读或被占用时另存为可修改副本

不支持旧版二进制格式 `.doc`、`.xls`、`.ppt`。请先在 Word、Excel、PowerPoint 或 WPS 中另存为新版格式。

## 使用方式

### 直接运行 Windows 版

从发布包中运行 `文件源数据编辑.exe`，无需安装 Python。

### 从源码运行

```powershell
python -m pip install -r requirements.txt
python app.py
```

## 功能说明

### 单文件编辑

选择一个文件后可查看与修改常用属性，点击“保存修改”写入。选择“清除嵌入属性”会清空文档内嵌属性与自定义属性，但不会修改 Windows 文件系统时间。

### 批量清除

点击“批量清除”可选择多个 Office 或 PDF 文件。程序会逐项显示成功或失败原因；单个文件失败不会中断其余任务。

### 备份、只读与占用

- 备份采用带时间戳的文件名，并保留原扩展名，能够直接打开。
- 无法直接覆盖的文件会提示另存副本；批量清除会自动生成名称含“已清除属性”的副本。
- 自动创建的备份与副本会移除只读属性，然后继续处理。

## PDF 限制

- 加密 PDF 不支持修改。
- 检测到数字签名的 PDF 会拒绝修改，避免签名失效。
- PDF 保存会重写元数据。建议保留“保存前备份”。

## 项目结构

```text
app.py                 程序入口与 Windows 高 DPI 配置
metadata_editor/       元数据读写与 Tkinter 界面
requirements.txt       源码运行依赖
run.cmd                Windows 快捷启动脚本
high-dpi.manifest      Windows 高 DPI 清单
```

## 许可

暂未声明开源许可证。
