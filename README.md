# 文件属性元数据编辑器

本地 Windows 桌面工具，用于查看与编辑新版 Office 文档和 PDF 的常用文件属性。

## 运行

```powershell
python app.py
```

无需网络连接或第三方 Python 包。

## 当前支持

- Office Open XML：`.docx`、`.xlsx`、`.pptx` 及其宏/模板变体
  - 核心属性、应用属性、自定义属性
  - Windows 创建/修改时间
- PDF：常用 Info 元数据及文件系统时间（不支持加密 PDF）

旧式 `.doc`、`.xls`、`.ppt` 为二进制复合文档，当前只读提示；请另存为新版格式后编辑。

保存前可选择创建 `.metadata-backup` 备份；“清除嵌入属性”不会修改 Windows 文件系统时间。
