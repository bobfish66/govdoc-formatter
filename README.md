# govdoc-formatter

按 **GB/T 9704-2012《党政机关公文格式》** 与 **《党政机关公文处理工作条例》** 规范化 Word 公文格式的 Skill / 命令行工具。

把任意 `.doc` / `.docx` 一键排成合规公文：版心、字体字号、行数字数、层次序数、成文日期、附件说明、页码镜像、表格格式全部自动处理，并输出格式体检与真实页数。

## 功能

| 项目 | 处理结果 |
|---|---|
| 页面与版心 | A4；天头 37mm、下白边 35mm、订口 28mm、右白边 26mm；版心 156mm×225mm |
| 标题 | 2 号方正小标宋，居中，与下段同页 |
| 层次序数 | 一级「一、」黑体 3 号；二级「（一）」楷体 3 号；三级「1.」「（1）」仿宋 3 号 |
| 正文 | 3 号仿宋，首行缩进 2 字（字符单位），固定行距 28.99pt，即每面 22 行、每行 28 字 |
| 附件说明 | 左空二字，「附件：」名称不加标点 |
| 成文日期 | 右空四字，阿拉伯数字 |
| 页码 | 4 号半角宋体「— 1 —」，单页居右空一字、双页居左空一字，距版心下边缘 7mm |
| 表格 | 宋体小四，表头加粗并跨页重复，行不跨页断开 |

## 安装为 Marvis Skill

把本目录导入 Marvis 的自定义 Skill 目录即可，或通过 `skill_install` 以本地路径安装：

```
skill_install(local_paths=["<本目录绝对路径>"])
```

安装后，说出「按国标规范这份文件的格式」即可自动命中。

## 命令行用法

依赖：Python 3.8+、`python-docx`

```bash
pip install python-docx
```

```bash
# 体检：只看当前版心、字号字体分布与页数，不写文件
python scripts/format_govdoc.py 报告.docx --check

# 规范化：输出 报告_国标版.docx
python scripts/format_govdoc.py 报告.docx

# 指定输出路径（支持 .doc 自动转换）
python scripts/format_govdoc.py 报告.doc -o 报告_国标.docx

# 其它
python scripts/format_govdoc.py 报告.docx --no-page-number      # 不加页码
python scripts/format_govdoc.py 报告.docx --char-spacing -0.2   # 正文字符间距
```

`.doc` 转换与页数统计依赖本机 Word（COM）或 LibreOffice；均缺失时仅法规化照常执行，页数需用 Word 打开确认。

## 目录结构

```
govdoc-formatter/
├── SKILL.md                    # Skill 定义：触发条件、执行流程、要素核对清单、常见坑
├── scripts/
│   └── format_govdoc.py        # 格式化引擎
└── references/
    └── gb9704-spec.md          # GB/T 9704-2012 与条例条款速查
```

## 关于页数

3 号字、每面 22 行的合规版式比常规小四、1.5 倍行距明显吃页数，同样内容通常多出 30%–50%。若同时有页数上限，正确做法是**精简内容**，而不是缩小字号——缩字号必然不符合国标。

## 字体说明

标准字体为方正小标宋简体、仿宋_GB2312、楷体_GB2312、黑体。本机未安装时 Word 预览会替代显示，但文档内保留标准字体名，在已安装字体的电脑上打开即恢复标准效果。

## 适用范围

仅覆盖版式自动化。版头（份号、密级、紧急程度、发文机关标志、发文字号、签发人）、主体（主送机关、署名、印章、附注）、版记（抄送、印发机关、分隔线）的要素齐备性需按 `references/gb9704-spec.md` 人工核对。

不适用于参考文献著录格式（那是 GB/T 7714）。

## License

MIT
