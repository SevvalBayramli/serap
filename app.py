from flask import Flask, render_template, request, send_file
import pandas as pd
import os

app = Flask(__name__)
UPLOAD_FOLDER = "uploads"
os.makedirs(UPLOAD_FOLDER, exist_ok=True)

@app.route("/", methods=["GET", "POST"])
def index():
    if request.method == "POST":
        file = request.files["file"]

        if not file or not file.filename.endswith(".xlsx"):
            return "Lütfen Excel (.xlsx) dosyası yükleyin"

        filepath = os.path.join(UPLOAD_FOLDER, file.filename)
        file.save(filepath)

        # === SENİN KODUN ===
        dosya = filepath
        genel_sayfa_adi = "GENEL_TOPLAM"

        excel = pd.ExcelFile(dosya)
        tum_urunler = []

        for sheet in excel.sheet_names:
            if sheet == genel_sayfa_adi:
                continue

            df = excel.parse(sheet, header=None)

            urun_adi_row = None
            urun_kod_row = None
            toplam_row = None

            for i in range(len(df)):
                hucre = str(df.iloc[i, 0]).strip().upper()
                if hucre == "ÜRÜN ADI":
                    urun_adi_row = i
                elif hucre in ["ÜRÜN KOD", "ÜRÜN KODU"]:
                    urun_kod_row = i
                elif hucre == "TOPLAM":
                    toplam_row = i

            if urun_adi_row is None or urun_kod_row is None or toplam_row is None:
                continue

            for col in range(1, len(df.columns)):
                urun_adi = df.iloc[urun_adi_row, col]
                kodu = df.iloc[urun_kod_row, col]
                toplam = df.iloc[toplam_row, col]

                if pd.isna(urun_adi) or pd.isna(kodu) or pd.isna(toplam):
                    continue

                tum_urunler.append({
                    "ÜRÜN ADI": str(urun_adi).strip(),
                    "KODU": str(kodu).strip(),
                    "TOPLAM": float(toplam)
                })

        genel_df = pd.DataFrame(tum_urunler)

        with pd.ExcelWriter(
            dosya,
            engine="openpyxl",
            mode="a",
            if_sheet_exists="replace"
        ) as writer:
            genel_df.to_excel(writer, sheet_name=genel_sayfa_adi, index=False)

        # kullanıcıya dosyayı geri ver
        return send_file(dosya, as_attachment=True)

    return render_template("index.html")


if __name__ == "__main__":
    app.run(debug=True)
