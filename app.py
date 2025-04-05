# app.py
from flask import Flask, render_template, request, redirect, send_from_directory, url_for
import os
from werkzeug.utils import secure_filename
import pandas as pd
from datetime import datetime
from collections import defaultdict

UPLOAD_FOLDER = "uploads"
RESULT_FOLDER = "results"
ALLOWED_EXTENSIONS = {"xlsx"}

app = Flask(__name__)
app.config["UPLOAD_FOLDER"] = UPLOAD_FOLDER
app.config["RESULT_FOLDER"] = RESULT_FOLDER

os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(RESULT_FOLDER, exist_ok=True)

def allowed_file(filename):
    return "." in filename and filename.rsplit(".", 1)[1].lower() in ALLOWED_EXTENSIONS

@app.route("/", methods=["GET", "POST"])
def index():
    if request.method == "POST":
        uploaded_files = request.files.getlist("files")
        saved_files = []

        for file in uploaded_files:
            if file and allowed_file(file.filename):
                filename = secure_filename(file.filename)
                filepath = os.path.join(app.config["UPLOAD_FOLDER"], filename)
                file.save(filepath)
                saved_files.append(filename)

        return render_template("index.html", uploaded=saved_files, result_ready=False)

    return render_template("index.html", uploaded=None, result_ready=False)

@app.route("/verarbeiten", methods=["POST"])
def verarbeiten():
    stichtag = datetime(2025, 1, 1)
    zeilen = []
    fehler = defaultdict(list)
    lfd_nr = 1

    upload_folder = app.config["UPLOAD_FOLDER"]
    result_folder = app.config["RESULT_FOLDER"]

    files = [f for f in os.listdir(upload_folder) if f.endswith(".xlsx")]

    for fname in files:
        try:
            pfad = os.path.join(upload_folder, fname)
            vereinssatz = pd.read_excel(pfad, sheet_name="Vereinssatz")
            personensatz = pd.read_excel(pfad, sheet_name="Personensatz")

            personen = personensatz.copy()
            personen["Geburtsdatum"] = pd.to_datetime(personen["Geburtsdatum"], errors="coerce", dayfirst=True)
            personen["Alter"] = personen["Geburtsdatum"].apply(lambda x: (stichtag - x).days // 365 if pd.notnull(x) else None)
            personen["GeschlechtLabel"] = personen["Geschlecht"].apply(lambda x: {1: "w", 2: "m", 3: "d"}.get(x, None))

            def zaehle(min_alt, max_alt=None):
                if max_alt:
                    df = personen[(personen["Alter"] >= min_alt) & (personen["Alter"] < max_alt)]
                else:
                    df = personen[personen["Alter"] >= min_alt]
                return {g: int((df["GeschlechtLabel"] == g).sum()) for g in ["w", "m", "d"]}

            def get_wert(bezeichnung):
                z = vereinssatz[vereinssatz["Bezeichnung"] == bezeichnung]["Inhalt"]
                return z.iloc[0] if not z.empty and pd.notnull(z.iloc[0]) else ""

            def ersatz(feld):
                p = personen[personen["Funktion1"] == 1]
                return p.iloc[0][feld] if not p.empty and pd.notnull(p.iloc[0][feld]) else ""

            u18, u27, a18, a27 = zaehle(7,18), zaehle(18,27), zaehle(18), zaehle(27)
            gesamt = personen["GeschlechtLabel"].value_counts().reindex(["w","m","d"], fill_value=0)

            verein_name = get_wert("Verein") or fname

            zeile = {
                "lfd. Nr.": lfd_nr,
                "Verbandsnummer": get_wert("Verbandsnummer"),
                "Verein/Verband": verein_name,
                "Titel": get_wert("Titel"),
                "Vorname": get_wert("Vorname") or ersatz("Vorname"),
                "Name": get_wert("Name") or ersatz("Name"),
                "Straße/Postfach": get_wert("Straße") or ersatz("Straße/Postfach"),
                "Land": "D",
                "PLZ": get_wert("PLZ") or ersatz("PLZ"),
                "Ort": get_wert("Ort") or ersatz("Ort"),
                "E-Mail": get_wert("E-Mail") or ersatz("E-Mail"),
                "Jugendorchester": int(personen["Jugendkapelle"].dropna().shape[0] > 0),
                "Erwachsenenorchester": 1,
                "Seniorenorchester": int("Senioren" in personen.columns and personen["Senioren"].dropna().shape[0] > 0),
                "Orchester gesamt": 0,
                "Aktive < 18 w": u18["w"], "Aktive < 18 m": u18["m"], "Aktive < 18 d": u18["d"], "Aktive < 18": sum(u18.values()),
                "Aktive < 27 w": u27["w"], "Aktive < 27 m": u27["m"], "Aktive < 27 d": u27["d"], "Aktive < 27": sum(u27.values()),
                "Aktive ab 18 w": a18["w"], "Aktive ab 18 m": a18["m"], "Aktive ab 18 d": a18["d"], "Aktive ab 18": sum(a18.values()),
                "Aktive ab 27 w": a27["w"], "Aktive ab 27 m": a27["m"], "Aktive ab 27 d": a27["d"], "Aktive ab 27": sum(a27.values()),
                "Aktive w": int(gesamt["w"]), "Aktive m": int(gesamt["m"]), "Aktive d": int(gesamt["d"]), "Aktive": int(gesamt.sum()),
                "Fördernde": int(get_wert("Fördernde Mitglieder") or 0)
            }

            zeile["Orchester gesamt"] = zeile["Jugendorchester"] + zeile["Erwachsenenorchester"] + zeile["Seniorenorchester"]
            zeilen.append(zeile)
            lfd_nr += 1

        except Exception as e:
            fehler[fname].append(str(e))

    timestamp = datetime.now().strftime("%Y_%m_%d_%H-%M")
    excel_name = f"{timestamp}_Verbandsmeldung_BVN_2025.xlsx"
    txt_name = f"Fehlerprotokoll_{timestamp}.txt"

    excel_path = os.path.join(result_folder, excel_name)
    txt_path = os.path.join(result_folder, txt_name)

    pd.DataFrame(zeilen).to_excel(excel_path, index=False)
    with open(txt_path, "w", encoding="utf-8") as f:
        if fehler:
            for verein in sorted(fehler):
                f.write(f"{verein}\n")
                for msg in fehler[verein]:
                    f.write(f"{msg}\n")
                f.write("\n")
        else:
            f.write("Keine Fehler festgestellt.")

    return render_template("index.html", uploaded=files, result_ready=True,
                           excel_file=excel_name, txt_file=txt_name)

@app.route("/results/<filename>")
def download_file(filename):
    return send_from_directory(app.config["RESULT_FOLDER"], filename, as_attachment=True)

@app.route("/uploads/<filename>")
def uploaded_file(filename):
    return send_from_directory(app.config["UPLOAD_FOLDER"], filename)

if __name__ == "__main__":
    app.run(debug=False, host="0.0.0.0", port=int(os.environ.get("PORT", 5000)))

