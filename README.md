# political-meeting

政治学习会议材料整理 skill，用于把学校政治学习材料按固定版式归入总文档，并持续追踪缺失材料。

这个 skill 面向“材料异步到达”的真实工作流：老师可能先发学习内容，之后再补签到表或会议照片。脚本会按既定顺序归档，并在每次处理后提示该周还缺什么。

## 功能特性

- 自动维护政治学习记录顺序：`会议记录`、`会议签到表`、周标题、简要介绍、`会议照片：`、学习内容。
- 支持学习内容、签到表、会议照片异步到达。
- 支持连续周次合并，例如第 9 周和第 10 周合并为“第九、十周政治学习”。
- 简要介绍只保留学校、学期、周次和学习材料标题，不把正文摘要塞进简介。
- 会议照片和签到表使用真实图片检测，避免把占位行误判为已上传。
- 总文档被 Word 打开时，会提示老师先关闭文档，而不是输出 Python 堆栈。
- 首次使用时，如果不知道总文档路径，会主动询问老师发送总文档或提供完整路径。

## 目录结构

```text
political-meeting/
├── SKILL.md
├── README.md
├── config.example.json
├── scripts/
│   ├── check_status.py
│   ├── config.py
│   ├── format_convert.py
│   ├── insert_material.py
│   ├── officecli_helper.py
│   └── requirements.txt
└── tests/
    ├── test_config_resolution.py
    └── test_ordered_layout.py
```

## 依赖

需要先安装 OfficeCLI，用于稳定读写 Word 文档：

```powershell
irm https://raw.githubusercontent.com/iOfficeAI/OfficeCLI/main/install.ps1 | iex
```

安装 Python 依赖：

```powershell
pip install -r scripts/requirements.txt
```

如果要运行测试：

```powershell
pip install pytest
```

## 配置

复制示例配置：

```powershell
Copy-Item config.example.json config.json
```

编辑 `config.json`：

```json
{
  "总表路径": "C:\\path\\to\\政治学习会议记录.docx",
  "学校名": "广州市番禺区番广附万博学校",
  "学期": "2025学年第二学期"
}
```

`config.json` 包含本机路径，默认不提交到 Git。

## 使用方式

插入学习内容：

```powershell
python scripts/insert_material.py --week-title "2025学年第二学期第九、十周政治学习" --type "学习内容" --file "C:\path\to\第9周学习内容.docx"
```

插入会议照片：

```powershell
python scripts/insert_material.py --week-title "2025学年第二学期第九、十周政治学习" --type "会议照片" --file "C:\path\to\会议照片.jpg"
```

插入签到表：

```powershell
python scripts/insert_material.py --week-title "2025学年第二学期第九、十周政治学习" --type "签到表" --file "C:\path\to\签到表.jpg"
```

查询某周还缺什么：

```powershell
python scripts/check_status.py --week-title "第九、十周政治学习"
```

查询所有周次状态：

```powershell
python scripts/check_status.py
```

## 老师对话示例

老师发送两个文档并说：

```text
把这两周的学习材料归入总文档
```

XiaoBa 应识别第 9 周、第 10 周学习内容，按合并周次执行：

```text
2025学年第二学期第九、十周政治学习
```

归档完成后，必须基于 `check_status.py` 的结果回复，例如：

```text
已将第九、十周学习内容归入总文档。
当前还缺：签到表、会议照片。
```

## 测试

```powershell
python -m pytest tests -q
```

当前测试覆盖：

- 总文档路径为空、失效或首次使用时的提示。
- 学习内容先到时创建固定结构。
- `会议签到表` 和 `会议照片：` 占位不误判为已上传。
- 签到表或照片后到时插回正确位置。
- 简要介绍只使用材料标题。
- 学习内容正文使用统一字号。
- OfficeCLI 图片插入参数兼容当前版本。

## 发布到 GitHub

建议仓库名：

```text
political-meeting-skill
```

在 GitHub 页面创建新仓库后，可以执行：

```powershell
git init
git add .
git commit -m "Initial political meeting skill"
git branch -M main
git remote add origin https://github.com/buildsense-ai/political-meeting-skill.git
git push -u origin main
```

上传前确认不要提交：

- `config.json`
- `__pycache__/`
- `.pytest_cache/`
- 真实学校材料、会议照片或签到表

## License

按仓库需要补充许可证。若没有特殊要求，建议使用 MIT License。
