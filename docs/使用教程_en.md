# Format-Preserving Translator · User Guide

> A tool that translates contracts, quotations and bills of quantities into another language.
> **Layout stays exactly as it was**: fonts, borders, tables and formulas are untouched — only the text changes.
> Supports `.xlsx` (Excel) and `.docx` (Word), both English → Chinese and Chinese → English.

---

## 1. Getting an API Key (one time only)

Translation is powered by AI and billed by usage. **A 300-row bill of quantities costs roughly 20 fen (0.2 CNY)** — even dozens of documents cost only a few yuan.

The app supports several AI providers. **Any one of them will work** — you only need a key from one. DeepSeek is the cheapest and is described first; the others are listed below.

### Step 1 — Open the website

Go to: **https://platform.deepseek.com**

(This is DeepSeek's developer platform, not the chat website.)

### Step 2 — Sign up and log in

Register with your phone number or email, the same as any other website.

### Step 3 — Top-up (probably not needed yet)

**New accounts receive 6 CNY (¥6) of free credit.**

That sounds small, but for translation it goes a long way — an 88-row BOQ costs about 5 fen, so ¥6 covers well over a hundred documents. **Use the free credit first; decide about topping up later.**

When the credit runs out:

- To keep going → top up 10–20 CNY, which lasts several months
- **If you would rather not pay** → two options:
  1. Switch to **Local translation** (files never leave the company network, but quality is one notch lower — see section 5)
  2. **Ask the developer for a key** and use their quota temporarily

> Whether the free credit expires depends on the platform — check the balance page after signing up.

### Step 4 — Create an API Key

- In the left menu, find **API Keys**
- Click **Create API Key**
- Give it any name, e.g. `translate`
- Confirm

### Step 5 — Copy it and save it

**The key is displayed only once.** It is the string starting with `sk-`.

⚠️ Copy it immediately and paste it into a text file. Once you close the page you cannot see it again — you would have to create a new one.

### Step 6 — Paste it into the app

At the bottom of the app window there is a small link: **"更换翻译供应商 / 填 API Key"** (Change provider / enter API key). Click it.

1. Choose **DeepSeek** from the dropdown
2. Paste the key
3. Click Save

**You only do this once.** The app remembers it.

> Keep your own key to yourself. **Never post a key in a group chat** — anyone who has it can spend your credit.

---

### Using a different provider

Seven providers are supported. **Each key is stored separately**, so switching back and forth does not overwrite anything. To switch: pick it in the dropdown, paste that provider's key, save.

| Provider | Console / sign-up | Price level | Notes |
|---|---|---|---|
| **DeepSeek** | https://platform.deepseek.com | Cheapest | Default; recommended starting point |
| **Zhipu GLM** | https://open.bigmodel.cn | Cheap, some models free | China-based, easy sign-up |
| **Moonshot Kimi** | https://platform.moonshot.cn | Medium | China-based |
| **Alibaba Qwen** | https://bailian.console.aliyun.com | Medium, new-user credit | Log in with an Alibaba Cloud account |
| **OpenAI** | https://platform.openai.com | Cheap (billed in USD) | Needs an OpenAI account and an international card |
| **Anthropic Claude** | https://console.anthropic.com | **Expensive — roughly 3–4× DeepSeek** | International card required |
| **Custom** | You supply it | Depends | Any OpenAI-compatible endpoint |

**Getting a key works the same way everywhere**: sign up → find "API Keys" → create → copy the string.

⚠️ **Two things you must know:**

1. **Prices change — always check the provider's own page.** The table above is only a rough guide. DeepSeek is currently the best value and is good enough for contract work.
2. **Switching providers degrades the terminology and ambiguity tables.** Those were tuned around DeepSeek, which follows instructions closely — when told to use a fixed glossary it actually does. Some models rewrite freely, and term consistency drops. **Test with one document before trusting a new provider with an important contract.**

---

## 2. Installation

### Unzip

Extract the archive anywhere you like, e.g. `D:\translation`.

> ⚠️ **Keep the whole folder together.** The `_internal` folder is the program itself; dragging the .exe out on its own will not work.

### Double-click

Double-click **`保格式翻译.exe`** (the file name is Chinese, meaning "format-preserving translator").

A window titled 保格式翻译 opens. The first launch may take a few seconds.

> If no window appears, check the taskbar — it may be behind another window. If it still does not appear, restart the computer and try again.

---

## 3. Translating a file

1. **Drag the file into the dashed box** (or click the box to browse). Several files at once is fine.
2. Choose the direction: **英 → 中** (English → Chinese) or **中 → 英** (Chinese → English).
3. Wait for the progress bar (it shows "翻译中 42 / 88 条" = translating item 42 of 88).
4. Click **下载译文** (Download translation).

After that, extra buttons appear under the file:

| Button | What it does |
|---|---|
| **下载译文** | Get the finished file |
| **回译体检** | Back-translates the result and compares it with the original to flag sentences that may have drifted |
| **逐条改** | Side-by-side editor: original on the left, translation on the right — edit and re-download |
| **看可疑的** | Shows only the suspicious items (number mismatch, glossary not followed, untranslated) |
| **发现新词** | Finds terms in this document that are not yet in the glossary; tick the ones to add |
| **一键重译** | Re-translate the flagged items |

### Two switches

**双语对照 (bilingual)** — puts the original and the translation in the same cell (original on top, translation below). Useful for reviewing or for showing a manager. Not available for Word.

**Engine:**

- **云端翻译 (cloud)** — default, best quality, about 0.2 CNY per document
- **本地翻译 (local)** — use this for confidential files; quality is slightly lower (see section 5)

---

## 4. File conversion (the second tab)

There is a second tab at the top: **文件转换** (file conversion). It converts PDF to Word, Word to PDF, changes image formats and so on — 270 formats in total.

### Why it is here

Because contracts from clients often arrive as **PDF**, which cannot be translated directly. Convert it to Word first, then translate.

```
PDF (from the client) → File conversion → Word → Translate → Chinese Word, layout intact
```

### How to use it

1. Click **文件转换** at the top
2. Drag files into the box (several at once is fine)
3. Choose the target format below: PDF / DOCX / XLSX / PPTX / PNG / JPG / WEBP / TXT / CSV / HTML / EPUB / MD
4. When it finishes, click the download button

### Things worth knowing

- **Conversion happens on your own computer** — nothing is uploaded, it costs nothing, and it works offline
- It is fast: Word to PDF takes about 1–5 seconds
- The engine is the open-source project [Filee](https://github.com/KnifeLemon/Filee) (MIT licence), installed together with this program
- Unusual content (formulas inside a PDF, scanned pages) may not come out perfectly — it is best at ordinary documents

---

## 5. Local translation (keeping files inside the company)

This mode **never sends the file outside the company network**. Use it for contracts that must not go to an external server.

### When to use it

- The contract contains real party names, amounts or penalty clauses that policy forbids sending externally
- There is no internet connection

### How to use it

1. Select **本地翻译 (local)**
2. A field labelled **本地模型地址** (local model address) appears
3. Enter the address your administrator gives you (for example `http://192.168.1.100:11434`)
4. Click Save

From then on, translation runs on that machine.

### Keep in mind

- **Quality is one notch lower** — long sentences may break up. Have someone review the result, or run 回译体检 on it
- **That machine must be switched on**; if it is off, local translation will not work
- **Only one person can translate at a time** — if someone else is using it, you wait
- No address? Then your computer cannot reach the internal machine — use the cloud for now

---

## 6. Cost

| Item | Roughly |
|---|---|
| An 88-row BOQ | about 5 fen (0.05 CNY) |
| A 300-row BOQ | about 20 fen (0.2 CNY) |
| A contract of a few dozen pages | a few jiao to one yuan |

Pay-as-you-go; nothing is charged when you are not translating.

**The ¥6 sign-up credit covers well over a hundred BOQs** — normally one or two months of use.

### When the ¥6 runs out and you would rather not pay

Two ways to keep working:

1. **Switch to local translation** — free, but your computer must reach the company's internal machine, and quality is one notch lower
2. **Ask the developer for a key** — use their quota for a while, just tell them

> Do not register multiple accounts to farm the sign-up credit. Platforms identify users by phone number and real-name verification; repeated abuse gets accounts banned, and the work you already did is lost.

---

## 7. FAQ

**Q: I entered the key but it still says "insufficient balance".**
Check your balance at the provider's console. After topping up, wait a minute or two and try again.

**Q: It says "port is already in use".**
A previous instance did not close properly. Open Task Manager, find 保格式翻译, end it, and start again.

**Q: Antivirus software warns about the .exe.**
The program is self-built and has no digital signature, so some antivirus products are cautious. Add it to the trusted list.

**Q: The layout came out wrong.**
It should not — the program only changes text; every other byte is untouched. If it really is broken, send the file to the developer.

**Q: Can it translate PDFs?**
Not yet. Convert to Word first.

**Q: Can I add terms to the glossary?**
Yes. After translating, click 发现新词 (find new terms). It lists words not yet in the glossary; tick them, adjust the translation, and add them. Every future document follows the updated glossary.

**Q: How do I know which sentence is wrong?**
Click 回译体检 (back-translation check). It translates the result back and compares it with the original, listing the sentences that drifted. **Note: these are leads, not verdicts** — a listed sentence is not necessarily wrong, it is just worth a look.

---

## 8. Reporting a problem

Send the developer these three things:

1. The file that caused the problem
2. The exact message shown (a screenshot is fine)
3. The file `运行日志.txt` (run log) in the program folder

With those three, most problems can be located quickly.
