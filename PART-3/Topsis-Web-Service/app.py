from flask import Flask, render_template, request
from topsis_logic import run_topsis
import os
import smtplib
from email.message import EmailMessage

app = Flask(__name__)

UPLOAD_FOLDER = "uploads"
os.makedirs(UPLOAD_FOLDER, exist_ok=True)

EMAIL = "aaryan_be23@thapar.edu"
PASSWORD = "ovwr mvri fswl rirf"

@app.route("/", methods=["GET", "POST"])
def index():
    if request.method == "POST":
        file = request.files["file"]
        weights = request.form["weights"]
        impacts = request.form["impacts"]
        email = request.form["email"]

        if not file or not weights or not impacts or not email:
            return "All fields are required"

        input_path = os.path.join(UPLOAD_FOLDER, file.filename)
        output_path = os.path.join(UPLOAD_FOLDER, "result.csv")
        file.save(input_path)

        try:
            run_topsis(input_path, weights, impacts, output_path)
            send_email(email, output_path)
            return "Result sent to email successfully"
        except Exception as e:
            return str(e)

    return render_template("index.html")

def send_email(to_email, attachment):
    msg = EmailMessage()
    msg["Subject"] = "TOPSIS Result"
    msg["From"] = EMAIL
    msg["To"] = to_email
    msg.set_content("Please find attached TOPSIS result.")

    with open(attachment, "rb") as f:
        msg.add_attachment(f.read(), maintype="application", subtype="octet-stream", filename="result.csv")

    with smtplib.SMTP_SSL("smtp.gmail.com", 465) as server:
        server.login(EMAIL, PASSWORD)
        server.send_message(msg)

if __name__ == "__main__":
    app.run(debug=True)
