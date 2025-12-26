from django.shortcuts import render
from django.http import HttpResponse

from docx import Document
from docx.shared import Cm, Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn


# ================= HELPERS =================

def set_col_widths(table):
    widths = [Cm(1.2), Cm(5.0), Cm(10.8)]  # PDF-accurate columns
    for row in table.rows:
        for i, w in enumerate(widths):
            row.cells[i].width = w


def set_row_height(row, height_cm):
    tr = row._tr
    trPr = tr.get_or_add_trPr()
    trHeight = OxmlElement("w:trHeight")
    trHeight.set(qn("w:val"), str(int(height_cm * 567)))
    trHeight.set(qn("w:hRule"), "exact")
    trPr.append(trHeight)


def bold_cell(cell):
    for p in cell.paragraphs:
        for r in p.runs:
            r.bold = True


# ================= MAIN VIEW =================

def generate_doc(request):

    if request.method == "POST":

        # ---------- Applicant ----------
        client_name = request.POST.get("client_name")
        branch_address = request.POST.get("branch_address")
        correspondence_address = request.POST.get("correspondence_address")
        telephone_no = request.POST.get("telephone_no")
        mobile = request.POST.get("mobile")
        email = request.POST.get("email")

        # ---------- Opposite Party ----------
        customer_name = request.POST.get("customer_name")
        op_registered = request.POST.get("op_registered_address") or "________________"
        op_correspondence = request.POST.get("op_correspondence_address") or "________________"
        op_telephone = request.POST.get("op_telephone")
        op_mobile = request.POST.get("op_mobile")
        op_email = request.POST.get("op_email")

        # ---------- Dispute ----------
        dispute_nature = request.POST.get("dispute_nature")

        # ================= DOCUMENT =================

        doc = Document()

        # ---------- Title ----------
        title = doc.add_paragraph()
        title.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r = title.add_run(
            "FORM ‘A’\n"
            "MEDIATION APPLICATION FORM\n"
            "[REFER RULE 3(1)]\n"
        )
        r.bold = True
        r.font.size = Pt(14)

        head = doc.add_paragraph(
            "Mumbai District Legal Services Authority\n"
            "City Civil Court, Mumbai"
        )
        head.alignment = WD_ALIGN_PARAGRAPH.CENTER

        doc.add_paragraph("")

        # ================= TABLE =================

        table = doc.add_table(rows=0, cols=3)
        table.style = "Table Grid"

        # ---------- DETAILS OF PARTIES ----------
        row = table.add_row()
        set_row_height(row, 0.9)
        row.cells[0].merge(row.cells[2])
        row.cells[0].text = "DETAILS OF PARTIES:"
        bold_cell(row.cells[0])

        # ---------- Applicant Name ----------
        row = table.add_row()
        set_row_height(row, 0.9)
        row.cells[0].text = "1"
        row.cells[1].text = "Name of Applicant"
        row.cells[2].text = client_name
        bold_cell(row.cells[1])

        # ---------- Applicant Address (EXACT PDF STYLE) ----------
        row = table.add_row()
        set_row_height(row, 3.0)

        row.cells[1].text = "Address"
        bold_cell(row.cells[1])

        cell = row.cells[2]
        p = cell.paragraphs[0]

        r = p.add_run("REGISTERED ADDRESS:\n")
        r.bold = True
        p.add_run(f"{branch_address}\n\n")

        r = p.add_run("CORRESPONDENCE BRANCH ADDRESS:\n")
        r.bold = True
        p.add_run(f"{correspondence_address}")

        # ---------- Applicant Contacts ----------
        row = table.add_row()
        set_row_height(row, 0.8)
        row.cells[1].text = "Telephone No."
        row.cells[2].text = telephone_no
        bold_cell(row.cells[1])

        row = table.add_row()
        set_row_height(row, 0.8)
        row.cells[1].text = "Mobile No."
        row.cells[2].text = mobile
        bold_cell(row.cells[1])

        row = table.add_row()
        set_row_height(row, 0.8)
        row.cells[1].text = "Email ID"
        row.cells[2].text = email
        bold_cell(row.cells[1])

        # ---------- Opposite Party Header ----------
        row = table.add_row()
        set_row_height(row, 0.9)
        row.cells[0].merge(row.cells[2])
        row.cells[0].text = "2  Name, Address and Contact details of Opposite Party:"
        bold_cell(row.cells[0])

        # ---------- Opposite Party Name ----------
        row = table.add_row()
        set_row_height(row, 0.9)
        row.cells[1].text = "Name"
        row.cells[2].text = customer_name
        bold_cell(row.cells[1])

        # ---------- Opposite Party Address (EXACT PDF STYLE) ----------
        row = table.add_row()
        set_row_height(row, 3.0)

        row.cells[1].text = "Address"
        bold_cell(row.cells[1])

        cell = row.cells[2]
        p = cell.paragraphs[0]

        r = p.add_run("REGISTERED ADDRESS:\n")
        r.bold = True
        p.add_run(f"{op_registered}\n\n")

        r = p.add_run("CORRESPONDENCE ADDRESS:\n")
        r.bold = True
        p.add_run(f"{op_correspondence}")

        # ---------- Opposite Party Contacts ----------
        row = table.add_row()
        set_row_height(row, 0.8)
        row.cells[1].text = "Telephone No."
        row.cells[2].text = op_telephone
        bold_cell(row.cells[1])

        row = table.add_row()
        set_row_height(row, 0.8)
        row.cells[1].text = "Mobile No."
        row.cells[2].text = op_mobile
        bold_cell(row.cells[1])

        row = table.add_row()
        set_row_height(row, 0.8)
        row.cells[1].text = "Email ID"
        row.cells[2].text = op_email
        bold_cell(row.cells[1])

        # ---------- Dispute ----------
        row = table.add_row()
        set_row_height(row, 0.9)
        row.cells[0].merge(row.cells[2])
        row.cells[0].text = "DETAILS OF DISPUTE:"
        bold_cell(row.cells[0])

        row = table.add_row()
        set_row_height(row, 0.9)
        row.cells[0].merge(row.cells[2])
        row.cells[0].text = (
            "THE COMM. COURTS (PRE-INSTITUTION SETTLEMENT) RULES, 2018"
        )
        bold_cell(row.cells[0])

        row = table.add_row()
        set_row_height(row, 1.4)
        row.cells[0].merge(row.cells[2])
        row.cells[0].text = dispute_nature

        # ---------- Apply column widths LAST ----------
        set_col_widths(table)

        # ================= RESPONSE =================

        response = HttpResponse(
            content_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )
        response["Content-Disposition"] = (
            "attachment; filename=FORM_A_Mediation_Application.docx"
        )
        doc.save(response)
        return response

    return render(request, "form.html")
