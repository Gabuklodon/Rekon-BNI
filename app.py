from flask import Flask, render_template, request, send_from_directory
import os
import pandas as pd
import re
import io

app = Flask(__name__)
app.secret_key = "secret_key"

# Folder untuk file upload dan output
UPLOAD_FOLDER = os.path.join(os.getcwd(), 'uploads')
OUTPUT_FOLDER = os.path.join(os.getcwd(), 'outputs')
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(OUTPUT_FOLDER, exist_ok=True)

########################################
# Fungsi Helper untuk memproses file TXT
########################################

def parse_line(line):
    """
    Memecah baris berdasarkan '|' dengan aturan:
      - Kolom terakhir = Balance
      - Kolom kedua terakhir = DB/CR
      - Kolom ketiga terakhir = Amount
      - Kolom 0 hingga 3: No, PostDate, Branch, JournalNo
      - Sisanya digabung sebagai Description.
    """
    parts = line.split("|")
    if len(parts) < 8:
        return None
    balance = parts[-1].strip()
    dbcr = parts[-2].strip()
    amount = parts[-3].strip()
    no_ = parts[0].strip()
    post_date = parts[1].strip()
    branch = parts[2].strip()
    journal_no = parts[3].strip()
    description = "|".join(parts[4:-3]).strip()
    return {
        "No": no_,
        "PostDate": post_date,
        "Branch": branch,
        "JournalNo": journal_no,
        "Description": description,
        "Amount": amount,
        "DBCR": dbcr,
        "Balance": balance
    }

def extract_branch_code(text):
    """
    Mencoba mengekstrak kode cabang dari teks:
    1) Coba cari pola VA 16 digit (98822222XXX00000) dan ambil XXX.
    2) Jika tidak ditemukan, cari 3 digit (word boundary) di teks.
    3) Jika tidak ada, return "000".
    """
    # Step 1: Regex untuk 16 digit
    match_16 = re.search(r"(98822222\d{3}00000)", text)
    if match_16:
        return match_16.group(1)[8:11]
    # Step 2: Cari 3 digit
    match_3 = re.search(r"\b(\d{3})\b", text)
    if match_3:
        return match_3.group(1)
    return "000"

def summarize_description_by_segments(cleaned_desc, num_segments=2):
    """
    Merangkum deskripsi dengan mengambil 'num_segments' segmen terakhir
    dari deskripsi yang dipisahkan oleh delimiter '|'.
    """
    segments = cleaned_desc.split("|")
    segments = [seg.strip() for seg in segments if seg.strip() != ""]
    if not segments:
        return ""
    if len(segments) <= num_segments:
        return " | ".join(segments)
    else:
        return " | ".join(segments[-num_segments:])

########################################
# Fungsi untuk memproses file dan matching
########################################

def process_files(txt_path, bni_path):
    # Proses file TXT → df_main
    data_lines = []
    header_found = False
    with io.open(txt_path, "r", encoding="utf-8") as f:
        for line in f:
            line = line.strip()
            if not header_found:
                if line.startswith("No.|Post Date|Branch"):
                    header_found = True
                continue
            if line and line[0].isdigit():
                row_dict = parse_line(line)
                if row_dict and row_dict["DBCR"].strip().upper() == "C":
                    desc = row_dict["Description"]
                    amt_str = row_dict["Amount"]
                    branch_code = extract_branch_code(desc)
                    try:
                        amt_val = float(amt_str)
                    except:
                        amt_val = None
                    summary_desc = summarize_description_by_segments(desc, num_segments=2)
                    data_lines.append({
                        "kode cabang": branch_code,
                        "keterangan": summary_desc,
                        "nominal": amt_val
                    })
    df_main = pd.DataFrame(data_lines)
    
    # Proses file DATA TRANSAKSI BNI dari sheet "BNI180225"
    df_bni = pd.read_excel(bni_path, sheet_name="BNI180225")
    if "AP_INVOICE_AMOUNT" in df_bni.columns:
        df_bni.rename(columns={"AP_INVOICE_AMOUNT": "nominal"}, inplace=True)
    elif "AMOUNT" in df_bni.columns:
        df_bni.rename(columns={"AMOUNT": "nominal"}, inplace=True)
    if "AP_BRANCH_ID" not in df_bni.columns:
        if "BRANCH_ID" in df_bni.columns:
            df_bni.rename(columns={"BRANCH_ID": "AP_BRANCH_ID"}, inplace=True)
    df_bni["nominal"] = pd.to_numeric(df_bni["nominal"], errors="coerce")
    
    # Buat key di df_main dan df_bni (format key: "kode cabang_nominal" tanpa desimal)
    df_main["kode cabang"] = df_main["kode cabang"].astype(str).str.strip()
    df_main["nominal_int"] = df_main["nominal"].round(0).astype(int)
    df_main["match_key"] = df_main["kode cabang"] + "_" + df_main["nominal_int"].astype(str)
    
    df_bni["AP_BRANCH_ID"] = df_bni["AP_BRANCH_ID"].astype(str).str.strip()
    df_bni["short_branch"] = df_bni["AP_BRANCH_ID"].str[:3]
    df_bni["nominal_int"] = df_bni["nominal"].round(0).astype(int)
    df_bni["match_key"] = df_bni["short_branch"] + "_" + df_bni["nominal_int"].astype(str)
    
    # Debug: cetak key
    print("Contoh key df_main:")
    print(df_main[["kode cabang", "nominal_int", "match_key"]].head())
    print("\nContoh key df_bni:")
    print(df_bni[["AP_BRANCH_ID", "short_branch", "nominal_int", "match_key"]].head())
    
    # Lakukan merge berdasarkan match_key
    df_merge = pd.merge(df_main, df_bni[["match_key"]], on="match_key", how="left", indicator=True)
    df_merge["status"] = df_merge["_merge"].apply(lambda x: "match" if x=="both" else "unmatch")
    df_match = df_merge[["kode cabang", "nominal", "status"]].copy()
    
    # Tulis output ke file Excel
    output_path = os.path.join(OUTPUT_FOLDER, "final_output.xlsx")
    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        df_main.drop(columns=["match_key", "nominal_int"], errors="ignore").to_excel(writer, sheet_name="Processed Output", index=False)
        df_match.to_excel(writer, sheet_name="Matching Results", index=False)
    
    return output_path

########################################
# ROUTES FLASK
########################################

@app.route("/", methods=["GET", "POST"])
def index():
    if request.method == "POST":
        if "txt_file" not in request.files or "bni_file" not in request.files:
            return "File tidak ditemukan. Pastikan Anda mengupload file TXT dan file DATA TRANSAKSI BNI."
        txt_file = request.files["txt_file"]
        bni_file = request.files["bni_file"]
        if txt_file.filename == "" or bni_file.filename == "":
            return "Tidak ada file yang dipilih."
        
        # Simpan file ke folder uploads
        txt_path = os.path.join(UPLOAD_FOLDER, txt_file.filename)
        bni_path = os.path.join(UPLOAD_FOLDER, bni_file.filename)
        txt_file.save(txt_path)
        bni_file.save(bni_path)
        
        # Proses file
        output_path = process_files(txt_path, bni_path)
        download_link = f"/download/{os.path.basename(output_path)}"
        return render_template("index.html", download_link=download_link)
    return render_template("index.html", download_link=None)

@app.route("/download/<filename>")
def download_file(filename):
    return send_from_directory(OUTPUT_FOLDER, filename, as_attachment=True)

if __name__ == "__main__":
    app.run(debug=True)
