# Google Search Console SEO Report Generator 📈

A Python script I wrote and use at work to generate monthly web traffic reports using the Google Search Console API and the xlsxwriter library.

## 🔧 What It Does

- Creates a formatted Excel spreadsheet with two tabs:
  - **Query View**: Top-performing monthly keywords, with clicks, impressions, CTR, month-over-month % change, and totals.
  - **Page View**: Top-ranking pages with similar stats (minus MoM % change due to variability).
- Designed for a clean two-page printable layout
- Auto-calculates deltas and aggregates
- Custom formatting using company brand colors

## 💡 Why I Built It

This tool helps track changes in organic web traffic as I implemented SEO-related updates. It was designed to give clear, month-to-month visibility into which keywords and pages are performing and shifting.

> *In retrospect, I could’ve abstracted more of the repeated logic into functions. But as the scope expanded, copy-paste happened.*

## 🚀 How to Use

- Plug in your own GSC credentials
- Edit / Comment out lines 310-312 & 399-401 for http + https combining if necessary
- Run the script and get a professional web traffic report instantly!

---

## 🛠️ Tech Stack

- Python
- xlsxwriter
- Google Search Console API

