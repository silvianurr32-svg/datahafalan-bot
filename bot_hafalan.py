import logging
import gspread
from oauth2client.service_account import ServiceAccountCredentials
from telegram import Update
from telegram.ext import Application, CommandHandler, MessageHandler, filters, ContextTypes
from datetime import datetime, time
import asyncio

# --- KONFIGURASI ---
SHEET_NAME = 'Data Hafalan Santri'
CREDENTIALS_FILE = 'credentials.json'

# DAFTAR NAMA LENGKAP (Sesuai Kolom A di Excel)
DAFTAR_SANTRI = [
    "Angzil Wahyu Setiawan",
    "Sajjadul Aziz Arif",
    "Sakti Abdul Gani Agin",
    "M Jovan Ardiansyah Putra",
    "Royyan Mumtazan",
    "Altamis Abi Jaya",
    "A. Zaidan Ramadhanis",
    "Adesta solichul azka",
    "Ibnu dzaky",
    "Galih saputra",
    "Maftuh basyurani",
    "Dhiyaulhaq nazif hajid",
    "Ilham reifansyah",
    "Umar Maksum Abdul ",
    "Letda Devano Pratama",
    "Khairul Azzam",
    "M. Hafizh Baihaqi"
]

# Logging
logging.basicConfig(format='%(asctime)s - %(name)s - %(levelname)s - %(message)s', level=logging.INFO)
logger = logging.getLogger(__name__)

def connect_sheets():
    """Menghubungkan ke Google Sheets"""
    scope = ["https://spreadsheets.google.com/feeds", 'https://www.googleapis.com/auth/spreadsheets',
             "https://www.googleapis.com/auth/drive.file", "https://www.googleapis.com/auth/drive"]
    creds = ServiceAccountCredentials.from_json_keyfile_name(CREDENTIALS_FILE, scope)
    client = gspread.authorize(creds)
    sheet = client.open(SHEET_NAME).sheet1
    return sheet

def cari_nama_lengkap(keyword):
    """Mencari nama lengkap dari potongan nama"""
    keyword = keyword.lower().strip()
    for nama_asli in DAFTAR_SANTRI:
        if keyword in nama_asli.lower():
            return nama_asli
    return None

# --- FUNGSI AUTO-FILL ALPA ---
async def proses_auto_alpa(context: ContextTypes.DEFAULT_TYPE, chat_id=None):
    """Fungsi inti untuk mengisi slot kosong dengan Alpa"""
    try:
        sheet = connect_sheets()
        tanggal_header = datetime.now().strftime("%d %b")
        
        # 1. Cek Kolom Tanggal Hari Ini
        header_values = sheet.row_values(1)
        if tanggal_header in header_values:
            col_index = header_values.index(tanggal_header) + 1
        else:
            col_index = len(header_values) + 1
            sheet.update_cell(1, col_index, tanggal_header)

        # 2. Ambil Semua Data
        all_values = sheet.get_all_values()
        requests = []
        count_alpa = 0
        
        # 3. Loop setiap nama santri
        for i, nama_santri in enumerate(DAFTAR_SANTRI):
            # Cari baris nama ini di Excel
            row_index = -1
            for r_idx, row_val in enumerate(all_values):
                if row_val[0].strip() == nama_santri.strip(): 
                    row_index = r_idx + 1
                    break
            
            if row_index == -1:
                continue 

            current_row_data = all_values[row_index-1]
            
            isi_cell = ""
            if len(current_row_data) >= col_index:
                isi_cell = current_row_data[col_index-1]
            
            # Jika KOSONG, maka tandai ALPA
            if not isi_cell.strip():
                count_alpa += 1
                sheet.update_cell(row_index, col_index, "Alpa")
                
                requests.append({
                    "repeatCell": {
                        "range": {
                            "sheetId": sheet.id,
                            "startRowIndex": row_index - 1, "endRowIndex": row_index,
                            "startColumnIndex": col_index - 1, "endColumnIndex": col_index
                        },
                        "cell": {
                            "userEnteredFormat": {
                                "backgroundColor": {"red": 1.0, "green": 0.0, "blue": 0.0},
                                "textFormat": {"foregroundColor": {"red": 1.0, "green": 1.0, "blue": 1.0}, "bold": True},
                                "wrapStrategy": "WRAP"
                            }
                        },
                        "fields": "userEnteredFormat(backgroundColor,textFormat,wrapStrategy)"
                    }
                })

        if requests:
            sheet.client.batch_update(sheet.spreadsheet.id, {"requests": requests})
            msg = f"✅ **REKAP HARIAN (OTOMATIS)**\nDitemukan {count_alpa} santri kosong.\nOtomatis diisi **Alpa** (Merah)."
        else:
            msg = "✅ **REKAP HARIAN (OTOMATIS)**\nSemua santri aman (sudah ada data)."

        # Kirim laporan ke chat terakhir yg aktif (opsional, karena otomatis)
        if chat_id:
            await context.bot.send_message(chat_id=chat_id, text=msg, parse_mode='Markdown')

    except Exception as e:
        logger.error(e)

# --- HANDLER ---

async def command_selesai(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    """Trigger manual lewat perintah /selesai"""
    await update.message.reply_text("⏳ Sedang memeriksa data kosong...")
    await proses_auto_alpa(context, chat_id=update.effective_chat.id)

async def start(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    welcome_msg = (
        "🤖 **BOT ABSENSI (SERVER) SIAP**\n\n"
        "Cara Pakai:\n`Nama, Kategori, Hafalan`\n\n"
        "Jadwal Auto-Alpa: **00:00 WIB**"
    )
    await update.message.reply_text(welcome_msg, parse_mode='Markdown')

async def proses_input(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    pesan = update.message.text
    parts = pesan.split(',')
    
    if len(parts) < 2: return 

    input_nama = parts[0].strip()
    input_kategori = parts[1].strip()
    input_hafalan = parts[2].strip() if len(parts) > 2 else "-"
    
    nama_lengkap = cari_nama_lengkap(input_nama)
    
    if not nama_lengkap:
        await update.message.reply_text(f"❌ Nama **'{input_nama}'** tidak ditemukan.", parse_mode='Markdown')
        return

    status_msg = await update.message.reply_text(f"⏳ Memproses: **{nama_lengkap}**...", parse_mode='Markdown')

    try:
        sheet = connect_sheets()
        kategori_kapital = input_kategori.title()
        
        if input_kategori.lower() == "alpa":
            isi_baru = "Alpa"
        else:
            isi_baru = f"{kategori_kapital}: {input_hafalan}"
            
        tanggal_header = datetime.now().strftime("%d %b")

        try:
            cell_nama = sheet.find(nama_lengkap, in_column=1)
            row_index = cell_nama.row
        except gspread.exceptions.CellNotFound:
            await context.bot.edit_message_text(chat_id=update.effective_chat.id, message_id=status_msg.message_id, text=f"❌ Error: Nama '{nama_lengkap}' belum dimasukkan ke Excel Kolom A.")
            return

        header_values = sheet.row_values(1)
        if tanggal_header in header_values:
            col_index = header_values.index(tanggal_header) + 1
        else:
            col_index = len(header_values) + 1
            sheet.update_cell(1, col_index, tanggal_header)

        isi_lama = sheet.cell(row_index, col_index).value
        
        if isi_lama:
            if isi_baru in isi_lama:
                await context.bot.edit_message_text(chat_id=update.effective_chat.id, message_id=status_msg.message_id, text=f"⚠️ Data '{isi_baru}' sudah ada.")
                return
            isi_final = f"{isi_lama}\n{isi_baru}"
            teks_balasan = f"✅ Data **{nama_lengkap}** DITAMBAHKAN (Menumpuk)."
        else:
            isi_final = isi_baru
            teks_balasan = f"✅ Data **{nama_lengkap}** BERHASIL disimpan!"

        sheet.update_cell(row_index, col_index, isi_final)
        
        requests = []
        fmt = {"wrapStrategy": "WRAP"}
        if isi_baru == "Alpa":
            fmt["backgroundColor"] = {"red": 1.0, "green": 0.0, "blue": 0.0}
            fmt["textFormat"] = {"foregroundColor": {"red": 1.0, "green": 1.0, "blue": 1.0}, "bold": True}
            
        requests.append({
            "repeatCell": {
                "range": {
                    "sheetId": sheet.id,
                    "startRowIndex": row_index - 1, "endRowIndex": row_index,
                    "startColumnIndex": col_index - 1, "endColumnIndex": col_index
                },
                "cell": {"userEnteredFormat": fmt},
                "fields": "userEnteredFormat(backgroundColor,textFormat,wrapStrategy)"
            }
        })
        sheet.client.batch_update(sheet.spreadsheet.id, {"requests": requests})
        
        await context.bot.edit_message_text(chat_id=update.effective_chat.id, message_id=status_msg.message_id, text=teks_balasan, parse_mode='Markdown')

    except Exception as e:
        logger.error(e)
        await context.bot.edit_message_text(chat_id=update.effective_chat.id, message_id=status_msg.message_id, text=f"❌ Error: {e}")

def main():
    print("==================================================")
    print("   BOT PYTHONANYWHERE - SCHEDULED UTC 17:00       ")
    print("==================================================")
    
    # --- PASTE TOKEN DI BAWAH INI (Di dalam tanda petik) ---
    TOKEN_FINAL = "8408957968:AAEKJDk049Fq9VmH0cXVlYK1gXCGsWUZY4A"
    # -------------------------------------------------------

    application = Application.builder().token(TOKEN_FINAL).build()
    
    application.add_handler(CommandHandler("start", start))
    application.add_handler(CommandHandler("selesai", command_selesai))
    application.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND, proses_input))

    if application.job_queue:
        # SET WAKTU UTC: Jam 17:00 UTC = Jam 00:00 WIB
        t = time(17, 0, 0) 
        application.job_queue.run_daily(proses_auto_alpa, t)
        print("⏰ Jadwal Auto-Alpa diset ke jam 17:00 UTC (00:00 WIB).")

    print(f"✅ Bot sedang berjalan...")
    
    try:
        application.run_polling()
    except Exception as e:
        print(f"❌ Error: {e}")

if __name__ == '__main__':
    main()