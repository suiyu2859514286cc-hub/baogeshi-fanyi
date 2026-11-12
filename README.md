# 保格式翻译 / Format-Preserving Translator

> 把合同、报价单、工程量清单翻成另一种语言，**排版原样保留**。
> 支持 `.docx` 和 `.xlsx`，英↔中双向翻译。

一个给非程序员用的本地翻译工具：拖文件进去，出来还是原来那个文件，只是文字换成另一种语言。字体、边框、表格、合并单元格、公式——一个都不动。

---

## 它解决什么问题

用通用翻译工具翻合同，最烦的不是翻译质量，是**排版全乱了**：表格错位、字号乱跳、页眉页脚跑掉、公式变成乱码，最后还得手工排版一遍。

这个工具换个思路：**不重建文档，只换文字**。

- `.xlsx`：只改 `xl/sharedStrings.xml` 里的文本节点，其余二十几个内部文件一个字节不碰
- `.docx`：只改 `word/document.xml` 里段落级的文字，样式继承链、表格结构、编号、页眉页脚全保留

翻完的字还是那个字，格式还是那个格式。

---

## 主要功能

| 功能 | 说明 |
|---|---|
| **保格式翻译** | Word / Excel 原地翻，格式零改动 |
| **术语锁定** | 一张术语表决定所有专业词怎么译，全篇一致 |
| **歧义词处理** | `panel` 在 LED 屏上是"面板"、在配电里是"屏"——按语境自动选 |
| **占位符保护** | 数字、型号、品牌、日期先挖出来，译完原样放回 |
| **规则挑刺** | 翻完自动检查：数字对不对、术语用没用、有没有漏译 |
| **回译体检** | 把译文再翻回去跟原文比，找出翻歪的句子 |
| **逐条校对** | 左边原文右边译文，直接改字，改完重下 |
| **双语对照** | 原文和译文放在同一格里（Excel） |
| **文件转换** | 内置 [Filee](https://github.com/KnifeLemon/Filee)，PDF → Word 再翻 |
| **本地模型** | 接 Ollama，断网或涉密文件不出内网 |

---

## 快速开始

### 用现成的

1. 下载发布包，解压，双击 `保格式翻译.exe`
2. 第一次打开会让你填 API Key（[DeepSeek](https://platform.deepseek.com) 最便宜，注册送 6 元）
3. 把文件拖进去，选方向，等进度条，下载

详细步骤见 [`docs/使用教程.md`](docs/使用教程.md)（[English](docs/使用教程_en.md)）。

### 从源码跑

```bash
pip install -r requirements.txt
python app.py
```

浏览器打开 `http://127.0.0.1:8770`。

---

## 工作原理

```
文件
  ↓ 解析（按格式读文本层）
待译文本 + 占位符
  ↓ 按需注入术语表 + 歧义词规则 + 语域提示
LLM（DeepSeek / OpenAI / 智谱 / Kimi / 通义 / Claude / 自定义）
  ↓ 还原占位符 + 校验完整性
写回原文件（其余部分字节级不动）
  ↓
成品 + 可疑清单
```

**为什么能做到格式零改动**：因为从头到尾没有"重建文档"这一步。xlsx 是 OOXML 的压缩包，文本全在 `sharedStrings.xml`；docx 的正文在 `document.xml`。改文本节点，打包回去，其他部件原样拷贝——天然保格式。

---

## 术语表

用的是人工审过的对照表，不是让模型自由发挥。

`术语表/glossary.json` 格式：

```json
{
  "items": [
    { "en": ["MCB"], "zh": "微型断路器", "note": "表头缩写" },
    { "en": ["Distribution Board", "DB"], "zh": "配电箱", "note": "" }
  ]
}
```

`术语表/ambiguous.json` 存多义词的判据：

```json
[
  {
    "en": "panel",
    "senses": [
      { "zh": "面板", "when": "指 LED 视频墙、显示屏的显示面板单元" },
      { "zh": "屏", "when": "指电气控制屏、AMF/ATS 自动切换屏" }
    ]
  }
]
```

**术语表里的每一条都应该经过懂行的人确认。** 错的术语比没有术语更危险——它会让后面所有文件都跟着错。

---

## 支持的模型供应商

| 供应商 | 注册地址 | 备注 |
|---|---|---|
| DeepSeek | https://platform.deepseek.com | 默认，最便宜 |
| 智谱 GLM | https://open.bigmodel.cn | 国内，有免费模型 |
| Moonshot Kimi | https://platform.moonshot.cn | 国内 |
| 阿里通义千问 | https://bailian.console.aliyun.com | 国内 |
| OpenAI | https://platform.openai.com | 需国际信用卡 |
| Anthropic Claude | https://console.anthropic.com | 较贵 |
| 自定义 | 任意 OpenAI 兼容接口 | 填地址 + 模型名 |

本地模型走 [Ollama](https://ollama.com)，默认 `qwen2.5:7b`，效果弱于云端，适合断网或数据不能外发的场景。

---

## 开发

```bash
# 打包成 exe
pip install pyinstaller
pyinstaller --onedir --windowed --name 保格式翻译 app.py

# 打包成安装包（需要 Inno Setup）
iscc build/setup.iss
```

主要文件：

```
app.py                    本地服务 + 面板接口
panel/index.html          全部前端
脚本/translate_xlsx.py    Excel 保格式翻译 + 术语/占位符/歧义词注入
脚本/translate_docx.py    Word 保格式翻译
脚本/quality_check.py     规则挑刺（数字 / 术语 / 长度 / 漏译）
脚本/suggest_terms.py     自动发现新术语候选
脚本/update.py            自动更新
```

---

## 已知限制

- **PDF 不直接支持**——先用内置的文件转换转成 Word
- **扫描件没做 OCR**——图片型 PDF 转出来是空的
- **复杂排版有损**：文本框、艺术字、SmartArt 转出来可能变形
- **一段里多种字体样式**（半句加粗）会统一成第一段的样式
- **法律语义等价无法自动验证**——所以才有"挑刺"和"回译体检"，它们是线索不是判决

---

## 许可证

[MIT](LICENSE)

文件转换引擎是 [Filee](https://github.com/KnifeLemon/Filee)（MIT），单独分发时请遵守它自己的许可证。

---

## 致谢

- [Filee](https://github.com/KnifeLemon/Filee) — 270 种格式的离线转换引擎
- [python-docx](https://github.com/python-openxml/python-docx) / [openpyxl](https://openpyxl.readthedocs.io/) / [lxml](https://lxml.de/) — 文档解析
- [pywebview](https://pywebview.flowrl.com/) — 把网页装进原生窗口
- DeepSeek 等模型供应商
