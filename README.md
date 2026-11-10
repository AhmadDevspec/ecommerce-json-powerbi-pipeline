# Enterprise-Grade JSON to Excel Flattener (Python)

A production-ready Python script designed to recursively flatten highly nested JSON structures (APIs, NoSQL exports) and compile them into beautifully styled, corporate-ready Excel reports.

## 🚀 Key Features
- **Deep Recursive Flattening:** Automatically unpacks infinitely nested objects and lists into distinct relational columns.
- **Corporate Styling Engine:** Applies dynamic column auto-fitting, automated datatype alignments (text vs numbers), and professional navy/gray presentation styling.
- **Data Standardization:** Cleanly normalizes system keys (`user_profile_email`) into client-friendly headers (`User Profile Email`).

## 📦 Installation & Setup

1. Clone the repository:
   ```bash
   git clone https://github.com
   cd enterprise-json-excel-flattener
   ```

2. Install the necessary transformation engines:
   ```bash
   pip install pandas openpyxl
   ```

## 💻 Quick Start Usage

Run the script directly to see how it automatically parses the provided mock dataset into a styled report:
```bash
python json_to_excel_pro.py
```

## 📊 Visual Transformation Pipeline

### 1. Input (Complex, Nested JSON)
```json
[
  {
    "id": 101,
    "product": "Enterprise Cloud Server",
    "pricing": {"monthly": 89.99, "yearly": 950.00},
    "specs": {"ram": "32GB", "storage": ["1TB NVMe", "2TB HDD"]}
  }
]
```

### 2. Output (Formatted Excel)

| Id | Product | Pricing Monthly | Pricing Yearly | Specs Ram | Specs Storage 0 | Specs Storage 1 |
| :--- | :--- | :--- | :--- | :--- | :--- | :--- |
| 101 | Enterprise Cloud Server | 89.99 | 950.00 | 32GB | 1TB NVMe | 2TB HDD |
*(Note: The actual Excel file outputs auto-sized widths with a corporate navy blue header theme.)*

## 🛠️ Custom Fiverr Automations Available
Looking for a custom business integration? I build bespoke scripts tailored to your architecture:
- Live folder-monitoring automations (hot-folders)
- Scheduled background automation (cron-jobs/Task Scheduler)
- Custom header-mapping arrays and brand-specific formatting themes
