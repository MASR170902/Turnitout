# Turnitout - Self-Hosted Academic AI Humanizer & De-AI Evader

Aplikasi web self-hosted berbasis Streamlit untuk membersihkan indikator pola tulisan AI pada dokumen akademis, jurnal ilmiah, dan esai secara otomatis tanpa mengubah data teknis, formula, angka riset, maupun merusak tata letak dokumen Word (`.docx`).

---

## ✨ Fitur Utama

* **100% Programmatic Engine (Bebas Halusinasi):** Bekerja murni menggunakan aturan linguistik deterministik dan regex tanpa bergantung pada model LLM generatif. Nilai akurasi, metodologi riset, persamaan, dan sitasi (contoh: `[1]`, `[10]`) tidak akan diubah atau dikarang bebas.
* **In-Place DOCX XML Rewriter:** Mengubah teks langsung pada hierarki XML berkas `.docx` (termasuk paragraf, tabel bertingkat, dan *textbox/caption*). Format dokumen dua kolom standar IEEE, jenis font, ukuran, margin, dan gambar tetap utuh 100% setelah diunduh.
* **Multi-Standard AI Detection Filter:** Mengintegrasikan basis data pola AI dari berbagai rujukan industri:
* **Wikipedia Signs of AI Writing:** Memangkas kata klise (*delve into, tapestry, testament, crucial, realm*).
* **GPTZero & Copyleaks Patterns:** Menghapus transisi kaku (*Moreover, Furthermore, In conclusion, At its core*).
* **SearchAtlas & Grammarly Updates:** Membersihkan kata kerja analitis berlebih (*illuminate, bolster, streamline*), buzzwords teknologi (*cutting-edge, revolutionize*), serta frasa penahan sikap (*generally speaking, typically*).
* **blader/humanizer:** Menerapkan variasi ritme sintaksis alami.


* **Burstiness & Rhythm Enhancer:** Memecah kalimat majemuk panjang yang monoton (>26 kata) secara alami untuk meningkatkan metrik *burstiness* dokumen ke standar tulisan manusia (>15–20).
* **Em-Dash (`—`) Cleaner:** Merapikan penggunaan tanda pisah panjang berlebih ciri khas ChatGPT tanpa merusak penulisan format resmi IEEE (seperti `Abstract—`).
* **Kapasitas Tanpa Batas (Unlimited Words):** Sanggup memproses puluhan ribu kata sekaligus secara instan tanpa limit token atau kuota API.

---

## 🛠️ Tech Stack

* **Frontend & Web UI:** [Streamlit](https://streamlit.io/)
* **Document Processing:** [python-docx](https://python-docx.readthedocs.io/)
* **PDF Extraction:** [PyMuPDF](https://pymupdf.readthedocs.io/)
* **Runtime:** Python 3.10+

---

## 🚀 Panduan Instalasi Lokal (macOS / Linux / Windows)

1. **Clone repositori:**
```bash
git clone https://github.com/MASR170902/humanizer-app.git
cd humanizer-app

```


2. **Buat dan aktifkan virtual environment:**
```bash
python3 -m venv .venv
source .venv/bin/activate  # Di Windows: .venv\Scripts\activate

```


3. **Install dependensi:**
```bash
pip install --upgrade pip
pip install streamlit>=1.32.0 python-docx>=1.1.0 PyMuPDF>=1.24.0

```


4. **Jalankan aplikasi:**
```bash
streamlit run app.py

```


Aplikasi akan otomatis terbuka di browser pada alamat `http://localhost:8501`.

---

## 🌐 Panduan Deployment di Server / Proxmox VE (LXC Ubuntu/Debian)

1. **Siapkan direktori & instal paket:**
```bash
mkdir -p /opt/humanizer-app
cd /opt/humanizer-app
apt update && apt install -y python3 python3-pip python3-venv git
python3 -m venv .venv
/opt/humanizer-app/.venv/bin/pip install streamlit python-docx PyMuPDF

```


2. **Buat file service systemd (`/etc/systemd/system/humanizer.service`):**
```ini
[Unit]
Description=Streamlit Humanizer App
After=network.target

[Service]
Type=simple
User=root
WorkingDirectory=/opt/humanizer-app
Environment="OMP_NUM_THREADS=1"
Environment="ONNXRUNTIME_DISABLE_THREAD_AFFINITY=1"
Environment="ORT_LOG_SEVERITY_LEVEL=3"
ExecStart=/opt/humanizer-app/.venv/bin/streamlit run /opt/humanizer-app/app.py --server.port 8501 --server.address 0.0.0.0 --browser.gatherUsageStats false --server.enableCORS false --server.enableXsrfProtection false
Restart=always
RestartSec=5

[Install]
WantedBy=multi-user.target

```


3. **Aktifkan dan jalankan service:**
```bash
systemctl daemon-reload
systemctl enable humanizer
systemctl start humanizer

```


Akses melalui IP server di port `8501` (misal: `[http://192.168.](http://192.168.)x.x:8501`).

---

## 📄 Cara Penggunaan

1. Buka antarmuka aplikasi di browser.
2. Unggah dokumen berformat `.docx`, `.pdf`, atau `.txt` (atau gunakan tab **Paste Teks Langsung**).
3. Sesuaikan opsi filter di sidebar (pengaturan bawaan sudah dioptimalkan untuk naskah ilmiah).
4. Klik tombol **🚀 Proses & Humanize Dokumen**.
5. Periksa metrik perbandingan (skor indikator AI, burstiness, dan daftar kata yang disesuaikan).
6. Unduh berkas hasil pemrosesan via tombol **Download Microsoft Word (.docx)** dengan tata letak asli yang tetap terjaga.

---

## 📜 Lisensi

Proyek ini dirilis di bawah lisensi [MIT](https://www.google.com/search?q=LICENSE). Bebas digunakan, dimodifikasi, dan disebarluaskan untuk kebutuhan riset akademik maupun personal.