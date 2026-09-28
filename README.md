# 📧 Openhousefx Lead List MX Verifier

An open-source, multi-threaded bulk email domain verifier built with **Streamlit**, **Pandas**, and **dnspython**. 

This application cleans massive lead databases (CSV or Excel files up to 100,000+ rows) for **$0 cost** by performing parallel DNS MX (Mail Exchange) record lookups. It automatically filters out dead, inactive, or non-existent domains before uploading campaigns to cold email platforms like Instantly.ai or Smartlead, protecting your sending domain health and keeping bounce rates under 2%.

---

## 🚀 Key Features

* **Drag-and-Drop File Support:** Upload raw `.csv`, `.xlsx`, or `.xls` lead sheets directly.
* **Auto-Column Detection:** Automatically detects column headers containing "email" (case-insensitive).
* **High-Speed Multithreading:** Uses Python's `ThreadPoolExecutor` (50 parallel workers) to process tens of thousands of emails per minute.
* **Smart Encoding & File Handling:** Native support for `UTF-8`, `latin-1`, and Excel binary formats.
* **Detailed Analytics & Export:** View instant counts of original leads, valid leads, and filtered dead domains, with a 1-click download button for your cleaned CSV.

---

## 📂 Project Structure

```text
openhousefx-verifier/
│
├── app.py              # Main Streamlit web application script
├── requirements.txt    # Python package dependencies
└── README.md           # Project documentation
