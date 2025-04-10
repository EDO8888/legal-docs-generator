from flask import Flask, request, send_file, jsonify
from docxtpl import DocxTemplate
from docx2pdf import convert
from datetime import date
import os
import smtplib
from email.message import EmailMessage
from dotenv import load_dotenv

load_dotenv()

app = Flask(__name__)

SMTP_HOST = "smtp.gmail.com"
SMTP_PORT = 587
EMAIL_USER = os.getenv("EMAIL_USER")
EMAIL_PASS = os.getenv("EMAIL_PASS")

def get_template(language, doc_type):
    return f"templates/{language}/{doc_type}.docx"

@app.route('/generate', methods=['POST'])
def generate():
    try:
        data = request.json
        print("🔔 בקשה התקבלה מ:", data.get("sender_name"))
        print("📨 שליחה ל:", data.get("email"))

        language = data.get("language", "he")
        doc_type = data.get("doc_type", "legal_warning")
        output_format = data.get("output_format", "docx")

        template_path = get_template(language, doc_type)
        doc = DocxTemplate(template_path)

        context = {
            "שם הנמען": data["recipient_name"],
            "נושא": data["subject"],
            "תאריך": date.today().strftime("%d/%m/%Y"),
            "תאריך_הסכם": data["agreement_date"],
            "סכום": data["amount"],
            "תאריך_סופי": data["due_date"],
            "שם השולח": data["sender_name"],
            "תפקיד": data["sender_role"],
            "חתימה": data.get("sender_signature", "")
        }

        os.makedirs("output", exist_ok=True)
        filename_base = f"generated_letter_{date.today()}"
        docx_path = f"output/{filename_base}.docx"
        pdf_path = f"output/{filename_base}.pdf"

        doc.render(context)
        doc.save(docx_path)
        print("✅ נוצר מסמך DOCX:", docx_path)

        final_path = pdf_path if output_format == "pdf" else docx_path
        if output_format == "pdf":
            convert(docx_path, pdf_path)
            print("📄 הומר ל־PDF:", pdf_path)

        if data.get("send_email"):
            try:
                send_email_with_attachment(data["email"], final_path)
                print("✉️ מייל נשלח בהצלחה ל:", data["email"])
            except Exception as e:
                print("❌ שגיאה בשליחת מייל:", str(e))
                return jsonify({"error": str(e)}), 500

        return send_file(final_path, as_attachment=True)

    except Exception as e:
        print("❌ שגיאה כללית:", str(e))
        return jsonify({"error": str(e)}), 500

def send_email_with_attachment(recipient_email, filepath):
    msg = EmailMessage()
    msg["Subject"] = "המכתב המשפטי שלך מוכן"
    msg["From"] = EMAIL_USER
    msg["To"] = recipient_email
    msg.set_content("מצורף המכתב המשפטי שלך.")

    with open(filepath, 'rb') as f:
        file_data = f.read()
        file_name = os.path.basename(filepath)

    maintype, subtype = ("application", "pdf") if filepath.endswith(".pdf") else \
                        ("application", "vnd.openxmlformats-officedocument.wordprocessingml.document")

    msg.add_attachment(file_data, maintype=maintype, subtype=subtype, filename=file_name)

    with smtplib.SMTP(SMTP_HOST, SMTP_PORT) as server:
        server.starttls()
        server.login(EMAIL_USER, EMAIL_PASS)
        server.send_message(msg)

if __name__ == '__main__':
    port = int(os.environ.get("PORT", 5000))
    app.run(host='0.0.0.0', port=port)
