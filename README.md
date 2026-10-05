# 🧠 N-Back Test — Face Stimuli Edition

> A **free** cognitive task for measuring **working memory** using human
> face images. Built with Python for **psychology and cognitive science
> students** worldwide.

[![Download](https://img.shields.io/github/v/release/Hossein-aliian/N_Back?label=Download&style=for-the-badge&color=blue)](https://github.com/Hossein-aliian/N_Back/releases/latest)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](LICENSE)
[![Python](https://img.shields.io/badge/Python-3.10%2B-blue.svg)](https://www.python.org/)

🌐 **English:** README.md | **فارسی:** [README.fa.md](README.fa.md)

---

## ✨ What does it do?

The **N-Back** task is one of the most widely used paradigms in cognitive
psychology for assessing **working memory capacity**. In this version:

- Face images are presented sequentially
- The participant must decide whether the current face matches the one
  from **N steps earlier**
- Three stages in random order: **Attractive / Neutral / Unattractive**
- Results are automatically saved to a **Persian-headed Excel** file
- Full **Persian text input** support (name & field of study)

---

## 🚀 Quick Start (no Python required)

1. Go to the [**Releases**](../../releases/latest) page
2. Download the latest **`N_Back.rar`** file
3. Extract the archive
4. Double-click **`N_Back.exe`**
5. Done! ✅

> ⚠️ Windows Defender may warn you on the first run. Click
> **More info → Run anyway**. This is a **false positive** caused by
> PyInstaller — the full source code is in this repository.

---

## 🖼️ Add Your Own Images

Inside the `images/` folder there are three subfolders:

```
images/
├── high/       → attractive faces
├── neutral/    → neutral faces
└── low/        → unattractive faces
```

Replace the sample images with your own (JPG or PNG).

- **Minimum:** 2 images per folder
- **Recommended:** 20–40 images per folder

---

## ⚙️ Settings

Click the **gear icon** (top-left corner of the participant form) to
adjust:

| Setting | Default |
|---|---|
| Image duration | 2 s |
| Rest between images | 1 s |
| Trials per stage | 22 |
| N level | 2 |
| Match rate | 20% |
| Min match gap | 3 |
| **Stage Mode** | **Random** |

### 🎛️ Stage Mode

- **Random** (default): classic N-Back — stages shuffled randomly, 50/50 target/filler
- **Manual**: researcher configures each stage individually:
  - Which folder is **Target** (Attractive / Neutral / Unattractive)
  - Which folder is **Filler**
  - What **% of trials** come from the Target folder (10% – 90%)

In Manual mode, click **Configure Stages...** to open the configuration screen.

---

## 📊 Output Files

After each session, these files are created next to `N_Back.exe`:

| File | Description |
|---|---|
| `participants.xlsx` | One row per participant (Persian headers, 47 statistical columns) |
| `results/nback_*.csv` | Full trial-by-trial log for analysis |
| `settings.json` | Current configuration (auto-managed) |

---

## 💻 For Developers

```bash
git clone https://github.com/Hossein-aliian/N_Back.git
cd N_Back
pip install -r requirements.txt
python N_Back.py
```

**Dependencies:** `pygame`, `openpyxl`, `arabic-reshaper`, `python-bidi`

**Build a standalone EXE:**

```bash
pip install pyinstaller
pyinstaller --onefile --windowed --name "N_Back" --clean --icon="icon.ico" N_Back.py
```

---

## 💚 Free for Students

Released under the **MIT License**. Free for **academic research, theses,
and class projects** — no fees, no restrictions. If you use it in your
research, I'd love to hear about it.

---

## 📄 Citation

If you use this tool in a scientific publication, please cite it as:

```bibtex
@software{nback_face_2026,
  author    = {Hossein Alian},
  title     = {N-Back Test — Face Stimuli Edition},
  year      = {2026},
  publisher = {GitHub},
  url       = {https://github.com/Hossein-aliian/N_Back}
}
```

---

## 👤 Author

**Hossein Alian**
- GitHub: [@Hossein-aliian](https://github.com/Hossein-aliian)
- Email: mr.alian1997@yahoo.com

---

⭐ If this project helped you, please give it a **Star** so others can find it!