from flask import Flask, render_template, request, send_from_directory
import os

app = Flask(__name__)

UPLOAD_FOLDER = "uploads"
RESULT_FOLDER = "results"
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(RESULT_FOLDER, exist_ok=True)

app.config["UPLOAD_FOLDER"] = UPLOAD_FOLDER
app.config["RESULT_FOLDER"] = RESULT_FOLDER

uploaded_files = []

@app.route("/", methods=["GET", "POST"])
def index():
    global uploaded_files
    if request.method == "POST":
        files = request.files.getlist("files")
        uploaded_files = []
        for file in files:
            filename = file.filename
            path = os.path.join(app.config["UPLOAD_FOLDER"], filename)
            file.save(path)
            uploaded_files.append(filename)
    return render_template("index.html", uploaded=uploaded_files, result_ready=False)

@app.route("/verarbeiten", methods=["POST"])
def verarbeiten():
    excel_output = "Verbandsmeldung_2025.xlsx"
    txt_output = "Fehlerprotokoll_Verbandsmeldung_2025.txt"

    with open(os.path.join(RESULT_FOLDER, excel_output), "w") as f:
        f.write("Beispiel-Ausgabe-Excel")
    with open(os.path.join(RESULT_FOLDER, txt_output), "w") as f:
        f.write("Keine Fehler festgestellt.")

    return render_template(
        "index.html",
        uploaded=uploaded_files,
        result_ready=True,
        excel_file=excel_output,
        txt_file=txt_output,
    )

@app.route("/results/<filename>")
def download_file(filename):
    return send_from_directory(app.config["RESULT_FOLDER"], filename, as_attachment=True)

@app.route("/uploads/<filename>")
def uploaded_file(filename):
    return send_from_directory(app.config["UPLOAD_FOLDER"], filename)

if __name__ == "__main__":
    app.run(debug=False, host="0.0.0.0", port=int(os.environ.get("PORT", 5000)))
