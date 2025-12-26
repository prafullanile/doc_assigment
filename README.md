# PDF to MS Word Exact Replica (Django + Python)

##  Project Overview
This project recreates a provided **PDF form into an MS Word (.docx) document** as an **exact visual replica**, maintaining the same **layout, spacing, alignment, headings, tables, and structure**.

The application is built using **Django** and **python-docx**, with an HTML form interface to dynamically enter data and generate a print-ready Word document.

---

##  Objective
- Convert a government-style PDF form into an MS Word document
- Ensure **100% visual accuracy** with the original PDF
- Preserve layout integrity and readability
- Automate document generation via a web interface

---

##  Tech Stack
- **Python 3.x**
- **Django**
- **python-docx**
- **HTML/CSS**
- **Gunicorn (for deployment)**

---

##  Key Features
- Exact table structure with calibrated column widths
- Fixed row heights for consistent spacing
- Merged cells for section headers
- Address sections formatted exactly like the PDF
- Dynamic form-based data input
- Auto-generated `.docx` file download
- Print-safe and readable output



##  Layout Accuracy Highlights
- Column widths manually calibrated to A4 printable dimensions
- Address blocks rendered with bold headings and values below
- Labels and values correctly placed in separate columns
- Section headers merged exactly as per original PDF

---

## 📂 Project Structure

assigment/
│
├── pdf_to_doc/
│ ├── settings.py
│ ├── urls.py
│ ├── wsgi.py
│
├── generator/
│ ├── templates/
│ │ └── form.html
│ ├── views.py
│ ├── urls.py
│
├── manage.py
└── requirements.txt


---

##  How to Run Locally

### 1Clone Repository
```bash
git clone https://github.com/yourusername/pdf_to_doc.git
cd pdf_to_doc

python -m venv venv
venv\Scripts\activate   # Windows
source venv/bin/activate # macOS/Linux

pip install -r requirements.txt
python manage.py runserver

