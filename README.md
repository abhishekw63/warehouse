# RENEE Cosmetics - GT Mass Dump Automation

A Django-based web utility that automates the extraction of **BC Code** and **Order Quantity** from multiple Sales Order Excel files and generates a consolidated dump file.

This tool removes the need for repetitive manual copying from Excel sheets and significantly speeds up operational workflows.

---

# Problem

Operations teams often receive **30–40 Sales Order Excel files daily**.

To prepare a consolidated dump, the manual process requires:

1. Opening each Excel file
2. Locating **BC Code** and **Order Qty**
3. Copying item numbers and quantities
4. Combining them into a single sheet

This process is repetitive, time-consuming, and prone to human error.

---

# Solution

**GT Mass Dump Automation** processes multiple Sales Order files automatically and produces a consolidated dump within seconds.

The web application:

* Reads multiple Excel files simultaneously via drag-and-drop or file selector
* Extracts **BC Code** and **Order Qty**
* Cleans and validates data
* Reconstructs **SO number format**
* Generates a ready-to-use dump file and immediately downloads it to your browser

---

# Features

* Clean, responsive **Web UI** built with HTML, CSS, and JS (AJAX)
* Select multiple Excel files at once
* Automatic header detection
* Ignores summary rows
* Cleans numeric formats (`1,000 → 1000`)
* Converts filenames to original SO format
* Immediate in-browser file download via blob responses
* Date-based dump file naming

---

# Example Output

| SO Number   | Item No | Qty |
| ----------- | ------- | --- |
| SO/GTM/5954 | 200173  | 120 |
| SO/GTM/5954 | 200165  | 30  |
| SO/GTM/5955 | 201249  | 240 |

---

# Installation & Setup

## 1. Clone Repository

```
git clone <repository-url>
cd project
```

---

## 2. Create Virtual Environment

```
python -m venv .venv
```

Activate environment:

Windows:
```
.venv\Scripts\activate
```

Mac/Linux:
```
source .venv/bin/activate
```

---

## 3. Install Dependencies

```
pip install -r requirements.txt
```

Required libraries:

* Django
* pandas
* openpyxl

---

## 4. Run Migrations

Before running the server, make sure to apply the initial Django migrations:

```
python manage.py migrate
```

---

## 5. Start Server

Start the local Django web server:

```
python manage.py runserver
```

Open your browser and navigate to `http://127.0.0.1:8000/`.

---

# Application Workflow

1. Navigate to the web application URL.
2. Click **Select Excel Files**
3. Choose multiple Sales Order files
4. Click **Generate Dump**
5. The processed dump file will be automatically downloaded to your machine.

---

# Excel Template Requirements

Each Sales Order file must contain these columns:

| Column    | Description       |
| --------- | ----------------- |
| BC Code   | Product Item Code |
| Order Qty | Ordered quantity  |

The script automatically detects the header row and extracts the required data.

---

# Data Processing Rules

The automation performs the following cleaning steps:

### BC Code Handling

Excel values such as `200453.0` are converted to `200453`.

---

### Quantity Cleaning

The script converts values like:

```
1,000 → 1000
- → 0
```

Rows with **quantity ≤ 0** are ignored.

---

# Sales Order Format

Excel filenames typically appear as: `SOGTM5985.xlsx`

The script converts them back to the original format: `SO/GTM/5985`

---

# Technologies Used

| Technology | Purpose                   |
| ---------- | ------------------------- |
| Python     | Core language             |
| Django     | Web framework             |
| pandas     | Excel data processing     |
| HTML/JS    | Web Interface & AJAX      |

---

# License

Internal automation tool for operational efficiency.

---

# Author

Developed to automate repetitive Excel workflows and improve operational productivity for RENEE Cosmetics.
