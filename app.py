import streamlit as st
import pandas as pd
import numpy as np
import datetime
from PIL import Image
import io

# --- PAGE CONFIGURATION (MOBILE OPTIMIZED) ---
st.set_page_config(
    page_title="App Presensi & LMS Les Bahasa Inggris (Mobile V4)",
    page_icon="📱",
    layout="wide",
    initial_sidebar_state="collapsed"
)

# --- CUSTOM MOBILE-FRIENDLY CSS ---
st.markdown(""" <style> .main-title { font-size: 22px; font-weight: 800; color: #1E3A8A; text-align: center; margin-bottom: 5px; } .role-badge { background-color: #DBEAFE; color: #1E40AF; padding: 6px 12px; border-radius: 20px; font-size: 13px; font-weight: bold; text-align: center; margin-bottom: 15px; } .mobile-card { background-color: #FFFFFF; border: 1px solid #E5E7EB; padding: 12px; border-radius: 12px; box-shadow: 0 2px 4px rgba(0,0,0,0.05); margin-bottom: 10px; } .stButton>button { width: 100%; border-radius: 10px; height: 45px; font-weight: bold; } .download-btn-card { background-color: #F0FDF4; border: 1px solid #BBF7D0; padding: 15px; border-radius: 12px; margin-top: 10px; margin-bottom: 15px; } </style> """, unsafe_allow_html=True)

# --- HELPER FUNCTION: EXCEL EXPORTER ---
def generate_excel_recap(df_presensi, df_jurnal, df_nilai, df_siswa):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df_presensi.to_excel(writer, sheet_name='Rekap Presensi', index=False)
        df_jurnal.to_excel(writer, sheet_name='Jurnal Mengajar', index=False)
        df_nilai.to_excel(writer, sheet_name='Daftar Nilai', index=False)
        df_siswa.to_excel(writer, sheet_name='Data Siswa', index=False)
    return output.getvalue()

# --- INITIALIZE SESSION STATE ---
if 'data_siswa' not in st.session_state:
    st.session_state['data_siswa'] = pd.DataFrame([
        {"ID": "S01", "Nama": "Andi Pratama", "Kelas": "General Speaking - A", "Tipe": "Les Online", "Status": "Aktif"},
        {"ID": "S02", "Nama": "Siti Rahmawati", "Kelas": "IELTS Prep - B", "Tipe": "Tatap Muka", "Status": "Aktif"},
        {"ID": "S03", "Nama": "Budi Santoso", "Kelas": "Private Conversation", "Tipe": "Privat Rumah", "Status": "Aktif"},
        {"ID": "S04", "Nama": "Citra Dewi", "Kelas": "General Speaking - A", "Tipe": "Les Online", "Status": "Aktif"},
        {"ID": "S05", "Nama": "Doni Wijaya", "Kelas": "IELTS Prep - B", "Tipe": "Tatap Muka", "Status": "Aktif"},
    ])

if 'data_presensi' not in st.session_state:
    st.session_state['data_presensi'] = [
        {"Tanggal": datetime.date(2026, 10, 1), "ID Siswa": "S01", "Nama": "Andi Pratama", "Kelas": "General Speaking - A", "Modus": "Les Online", "Status Kehadiran": "Hadir", "Lokasi/GPS": "Zoom Meeting (Online)", "Foto Bukti": None, "Catatan": "Koneksi lancar"},
        {"Tanggal": datetime.date(2026, 10, 1), "ID Siswa": "S02", "Nama": "Siti Rahmawati", "Kelas": "IELTS Prep - B", "Modus": "Tatap Muka", "Status Kehadiran": "Hadir", "Lokasi/GPS": "Gedung Les R.01", "Foto Bukti": None, "Catatan": "Datang tepat waktu"},
        {"Tanggal": datetime.date(2026, 10, 2), "ID Siswa": "S03", "Nama": "Budi Santoso", "Kelas": "Private Conversation", "Modus": "Privat Rumah", "Status Kehadiran": "Hadir", "Lokasi/GPS": "-6.2088, 106.8456 (GPS Verified)", "Foto Bukti": None, "Catatan": "Check-in rumah siswa"},
    ]

if 'jurnal_mengajar' not in st.session_state:
    st.session_state['jurnal_mengajar'] = pd.DataFrame([
        {"Pertemuan Ke": 1, "Tanggal": datetime.date(2026, 10, 1), "Kelas": "General Speaking - A", "Modus": "Les Online", "Pengajar": "Miss Sarah", "Topik/Materi": "Introducing Yourself & Small Talk", "Ringkasan Pembelajaran": "Siswa berlatih roleplay perkenalan."},
        {"Pertemuan Ke": 1, "Tanggal": datetime.date(2026, 10, 1), "Kelas": "IELTS Prep - B", "Modus": "Tatap Muka", "Pengajar": "Mr. John", "Topik/Materi": "IELTS Speaking Part 1 Practice", "Ringkasan Pembelajaran": "Latihan topik hobi & pekerjaan."},
        {"Pertemuan Ke": 1, "Tanggal": datetime.date(2026, 10, 2), "Kelas": "Private Conversation", "Modus": "Privat Rumah", "Pengajar": "Miss Sarah", "Topik/Materi": "Business Vocabulary & Negotiation", "Ringkasan Pembelajaran": "Latihan negosiasi bahasa Inggris."},
    ])

if 'materi_pelajaran' not in st.session_state:
    st.session_state['materi_pelajaran'] = pd.DataFrame([
        {"Kode Modul": "MOD-01", "Judul": "Mastering English Pronunciation", "Kelas": "General Speaking - A", "Format": "PDF & Audio MP3", "Link/Akses": "https://drive.google.com/materi-01"},
        {"Kode Modul": "MOD-02", "Judul": "IELTS Speaking Cue Cards 2026", "Kelas": "IELTS Prep - B", "Format": "E-Book PDF", "Link/Akses": "https://drive.google.com/materi-02"},
        {"Kode Modul": "MOD-03", "Judul": "Daily Conversation & Idioms", "Kelas": "Private Conversation", "Format": "Audio & Video", "Link/Akses": "https://drive.google.com/materi-03"},
    ])

if 'daftar_nilai' not in st.session_state:
    st.session_state['daftar_nilai'] = pd.DataFrame([
        {"ID Siswa": "S01", "Nama": "Andi Pratama", "Kelas": "General Speaking - A", "Speaking Test": 85, "Listening Quiz": 90, "Grammar & Vocab": 80, "Tugas Mandiri": 88, "Nilai Akhir": 85.75},
        {"ID Siswa": "S02", "Nama": "Siti Rahmawati", "Kelas": "IELTS Prep - B", "Speaking Test": 78, "Listening Quiz": 82, "Grammar & Vocab": 85, "Tugas Mandiri": 80, "Nilai Akhir": 81.25},
        {"ID Siswa": "S03", "Nama": "Budi Santoso", "Kelas": "Private Conversation", "Speaking Test": 92, "Listening Quiz": 88, "Grammar & Vocab": 90, "Tugas Mandiri": 95, "Nilai Akhir": 91.25},
    ])

if 'daftar_penugasan' not in st.session_state:
    st.session_state['daftar_penugasan'] = pd.DataFrame([
        {"ID Tugas": "TUG-01", "Judul Tugas": "Audio Recording: Self Introduction", "Kelas": "General Speaking - A", "Tenggat Waktu": datetime.date(2026, 10, 15), "Status": "Terbit", "Instruksi": "Kirim rekaman suara (MP3/M4A) via HP"},
        {"ID Tugas": "TUG-02", "Judul Tugas": "IELTS Speaking Part 2 Mock Record", "Kelas": "IELTS Prep - B", "Tenggat Waktu": datetime.date(2026, 10, 18), "Status": "Terbit", "Instruksi": "Topik: Describe a memorable journey"},
    ])

# --- TOP MOBILE HEADER & ROLE SWITCHER ---
st.markdown("<div class='main-title'>📱 Presensi & LMS English Class</div>", unsafe_allow_html=True)

role = st.selectbox(
    "📲 Pilih Mode Peran Pengguna (Tampilan HP):",
    ["👑 Mode Admin / Pemilik Les", "👨‍🏫 Mode Guru / Tutor", "👤 Mode Siswa / Murid", "📝 Form Pendaftaran Mandiri (Public)"],
    index=0
)

st.markdown(f"<div class='role-badge'>Aktif: {role}</div>", unsafe_allow_html=True)


# ==============================================================================
# MODE 1: ADMIN / PEMILIK LES (Dashboard Rekap, Export Excel, Foto Bukti, Master Data)
# ==============================================================================
if role == "👑 Mode Admin / Pemilik Les":
    tab_a1, tab_a2, tab_a3, tab_a4 = st.tabs([
        "📊 Dashboard & Monitoring",
        "📥 Ekspor & Download Laporan",
        "📷 Rekap Presensi & Foto GPS",
        "⚙️ Kelola Master Data"
    ])

    with tab_a1:
        st.subheader("📊 Ringkasan Operasional Kursus")
        m1, m2, m3 = st.columns(3)
        with m1:
            st.metric("Total Siswa", len(st.session_state['data_siswa']))
        with m2:
            st.metric("Record Absensi", len(st.session_state['data_presensi']))
        with m3:
            st.metric("Pertemuan Selesai", len(st.session_state['jurnal_mengajar']))

        st.markdown("---")
        st.write("📈 **Distribusi Modus Pembelajaran Siswa:**")
        st.bar_chart(st.session_state['data_siswa']['Tipe'].value_counts())

    with tab_a2:
        st.subheader("📥 Ekspor Laporan Otomatis ke HP")
        st.caption("Unduh laporan lengkap dalam format **Excel (.xlsx)** atau **CSV** langsung ke galeri/file manager HP Anda.")

        # Prepare export DataFrame (removing raw photo bytes for clean tabular output)
        presensi_export_data = []
        for p in st.session_state['data_presensi']:
            p_copy = p.copy()
            p_copy['Ada Foto Bukti'] = "Ya" if p['Foto Bukti'] is not None else "Tidak"
            p_copy.pop('Foto Bukti', None)
            presensi_export_data.append(p_copy)
        
        df_p_export = pd.DataFrame(presensi_export_data)

        # Date Range & Filter
        col_f1, col_f2 = st.columns(2)
        with col_f1:
            st_date = st.date_input("Dari Tanggal", datetime.date(2026, 10, 1))
        with col_f2:
            end_date = st.date_input("Sampai Tanggal", datetime.date.today())

        # Filtered Export Data
        if not df_p_export.empty and 'Tanggal' in df_p_export.columns:
            df_p_filtered = df_p_export[(df_p_export['Tanggal'] >= st_date) & (df_p_export['Tanggal'] <= end_date)]
        else:
            df_p_filtered = df_p_export

        st.markdown("<div class='download-btn-card'>", unsafe_allow_html=True)
        st.write("📊 **1. Laporan Terpadu (Excel - Multi Sheet)**")
        st.caption("Mencakup Sheet: Rekap Presensi, Jurnal Mengajar, Daftar Nilai, & Master Data Siswa.")
        
        excel_bytes = generate_excel_recap(
            df_p_filtered,
            st.session_state['jurnal_mengajar'],
            st.session_state['daftar_nilai'],
            st.session_state['data_siswa']
        )
        
        st.download_button(
            label="📥 Download Laporan Lengkap (.xlsx Excel)",
            data=excel_bytes,
            file_name=f"Rekap_Lengkap_Les_Bahasa_Inggris_{datetime.date.today()}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )
        st.markdown("</div>", unsafe_allow_html=True)

        st.markdown("<div class='download-btn-card'>", unsafe_allow_html=True)
        st.write("📄 **2. Ekspor Khusus Presensi Kehadiran (.csv)**")
        csv_presensi = df_p_filtered.to_csv(index=False).encode('utf-8')
        st.download_button(
            label="📄 Download Rekap Absensi (.csv)",
            data=csv_presensi,
            file_name=f"Rekap_Presensi_{datetime.date.today()}.csv",
            mime="text/csv",
            use_container_width=True
        )
        st.markdown("</div>", unsafe_allow_html=True)

        st.markdown("<div class='download-btn-card'>", unsafe_allow_html=True)
        st.write("💯 **3. Ekspor Laporan Nilai Siswa (.csv)**")
        csv_nilai = st.session_state['daftar_nilai'].to_csv(index=False).encode('utf-8')
        st.download_button(
            label="💯 Download Laporan Nilai (.csv)",
            data=csv_nilai,
            file_name=f"Laporan_Nilai_Siswa_{datetime.date.today()}.csv",
            mime="text/csv",
            use_container_width=True
        )
        st.markdown("</div>", unsafe_allow_html=True)

    with tab_a3:
        st.subheader("📷 Galeri Rekap Presensi & Bukti Selfie / GPS")
        for idx, item in enumerate(reversed(st.session_state['data_presensi'])):
            with st.expander(f"📍 {item['Tanggal']} | {item['Nama']} - {item['Modus']} ({item['Status Kehadiran']})"):
                st.write(f"**Kelas:** {item['Kelas']}")
                st.write(f"**Lokasi / Titik GPS:** {item['Lokasi/GPS']}")
                st.write(f"**Catatan:** {item['Catatan']}")
                
                if item['Foto Bukti'] is not None:
                    try:
                        image = Image.open(io.BytesIO(item['Foto Bukti']))
                        st.image(image, caption=f"Foto Selfie / Verifikasi Lokasi ({item['Nama']})", width=300)
                    except Exception as e:
                        st.caption("📷 Foto tersimpan dalam format file.")
                else:
                    st.caption("ℹ️ Tidak ada lampiran foto (Presensi Manual/Online).")

    with tab_a4:
        st.subheader("⚙️ Kelola Data Siswa & Kursus")
        st.dataframe(st.session_state['data_siswa'], use_container_width=True)

        with st.expander("➕ Tambah Siswa Baru"):
            with st.form("form_tambah_siswa_admin"):
                s_id = f"S0{len(st.session_state['data_siswa'])+1}"
                s_n = st.text_input("Nama Siswa Baru")
                s_k = st.selectbox("Kelas", ["General Speaking - A", "IELTS Prep - B", "Private Conversation", "Kids English"])
                s_t = st.selectbox("Tipe Pembelajaran", ["Privat Rumah", "Tatap Muka", "Les Online"])

                btn_add_siswa = st.form_submit_button("➕ Daftarkan Siswa")
                if btn_add_siswa:
                    new_s_data = {"ID": s_id, "Nama": s_n, "Kelas": s_k, "Tipe": s_t, "Status": "Aktif"}
                    st.session_state['data_siswa'] = pd.concat([st.session_state['data_siswa'], pd.DataFrame([new_s_data])], ignore_index=True)
                    
                    # Add to score table too
                    new_n_data = {"ID Siswa": s_id, "Nama": s_n, "Kelas": s_k, "Speaking Test": 0, "Listening Quiz": 0, "Grammar & Vocab": 0, "Tugas Mandiri": 0, "Nilai Akhir": 0.0}
                    st.session_state['daftar_nilai'] = pd.concat([st.session_state['daftar_nilai'], pd.DataFrame([new_n_data])], ignore_index=True)
                    st.success(f"✅ Siswa {s_n} berhasil didaftarkan!")


# ==============================================================================
# MODE 2: GURU / TUTOR (Absensi Kamera Selfie, GPS, Jurnal, Penilaian, Download Rekap)
# ==============================================================================
elif role == "👨‍🏫 Mode Guru / Tutor":
    tab_g1, tab_g2, tab_g3, tab_g4, tab_g5 = st.tabs([
        "📸 Absensi & Selfie",
        "📖 Jurnal Mengajar",
        "💯 Input Nilai",
        "📌 Terbit Tugas",
        "📥 Download Rekap"
    ])

    with tab_g1:
        st.subheader("📸 Input Presensi Siswa (Kamera HP)")
        st.caption("Gunakan kamera HP untuk verifikasi foto selfie / lokasi saat kelas privat atau tatap muka.")

        with st.form("form_presensi_kamera", clear_on_submit=True):
            s_nama = st.selectbox("Pilih Siswa", options=st.session_state['data_siswa']['Nama'].tolist())
            selected_student = st.session_state['data_siswa'][st.session_state['data_siswa']['Nama'] == s_nama].iloc[0]
            
            col1, col2 = st.columns(2)
            with col1:
                tgl_sesi = st.date_input("Tanggal", datetime.date.today())
                modus = st.selectbox("Modus Pembelajaran", ["Privat Rumah", "Tatap Muka", "Les Online"], index=["Privat Rumah", "Tatap Muka", "Les Online"].index(selected_student['Tipe']))
            with col2:
                status_kehadiran = st.selectbox("Status Kehadiran", ["Hadir", "Izin", "Sakit", "Alpha"])
                lokasi_gps = st.text_input("Koordinat GPS / Alamat / Platform", value="-6.2088, 106.8456 (Verified GPS HP)" if modus=="Privat Rumah" else ("Ruang Kelas 1" if modus=="Tatap Muka" else "Zoom Meeting"))

            st.write("📷 **Ambil Foto Selfie / Bukti Lokasi via HP:**")
            img_file_buffer = st.camera_input("Buka Kamera HP untuk Ambil Foto Bukti")
            upload_backup = st.file_uploader("Atau Unggah Foto dari Galeri HP", type=['jpg', 'png', 'jpeg'])

            catatan_sesi = st.text_area("Catatan Evaluasi Sesi", placeholder="Contoh: Pronunciation and vocabulary extended well.")

            btn_simpan_presensi = st.form_submit_button("🚀 Simpan Presensi & Bukti Foto")

            if btn_simpan_presensi:
                foto_final = None
                if img_file_buffer is not None:
                    foto_final = img_file_buffer.getvalue()
                elif upload_backup is not None:
                    foto_final = upload_backup.getvalue()

                new_record = {
                    "Tanggal": tgl_sesi,
                    "ID Siswa": selected_student['ID'],
                    "Nama": s_nama,
                    "Kelas": selected_student['Kelas'],
                    "Modus": modus,
                    "Status Kehadiran": status_kehadiran,
                    "Lokasi/GPS": lokasi_gps,
                    "Foto Bukti": foto_final,
                    "Catatan": catatan_sesi
                }
                st.session_state['data_presensi'].append(new_record)
                st.success(f"✅ Presensi {s_nama} berhasil dicatat dengan foto bukti!")

    with tab_g2:
        st.subheader("📖 Isi Jurnal Mengajar Harian")
        with st.form("form_jurnal_guru"):
            p_ke = st.number_input("Pertemuan Ke-", min_value=1, value=len(st.session_state['jurnal_mengajar']) + 1)
            tgl_j = st.date_input("Tanggal Pertemuan", datetime.date.today())
            kls_j = st.selectbox("Kelas Target", st.session_state['data_siswa']['Kelas'].unique().tolist())
            mds_j = st.selectbox("Modus Kelas", ["Privat Rumah", "Tatap Muka", "Les Online"])
            pengajar_n = st.text_input("Nama Pengajar / Tutor", value="Miss Sarah")
            topik_m = st.text_input("Topik Bahasa Inggris Hari Ini", placeholder="Contoh: Phrasal Verbs & Speaking Practice")
            ringkasan_m = st.text_area("Ringkasan Materi & Evaluasi Siswa")

            btn_jurnal_save = st.form_submit_button("💾 Simpan Jurnal Mengajar")
            if btn_jurnal_save:
                new_j = {
                    "Pertemuan Ke": p_ke,
                    "Tanggal": tgl_j,
                    "Kelas": kls_j,
                    "Modus": mds_j,
                    "Pengajar": pengajar_n,
                    "Topik/Materi": topik_m,
                    "Ringkasan Pembelajaran": ringkasan_m
                }
                st.session_state['jurnal_mengajar'] = pd.concat([st.session_state['jurnal_mengajar'], pd.DataFrame([new_j])], ignore_index=True)
                st.success("✅ Jurnal mengajar berhasil disimpan!")

    with tab_g3:
        st.subheader("💯 Penilaian Siswa")
        df_n = st.session_state['daftar_nilai']
        s_dinilai = st.selectbox("Pilih Nama Siswa", df_n['Nama'].tolist())
        
        with st.form("form_nilai_guru"):
            spk = st.slider("Speaking Test", 0, 100, 85)
            lis = st.slider("Listening Quiz", 0, 100, 80)
            grm = st.slider("Grammar & Vocab", 0, 100, 85)
            tgs = st.slider("Tugas Mandiri", 0, 100, 90)

            btn_save_nilai = st.form_submit_button("💾 Simpan Nilai")
            if btn_save_nilai:
                n_akhir = (spk * 0.35) + (lis * 0.25) + (grm * 0.20) + (tgs * 0.20)
                idx = st.session_state['daftar_nilai'][st.session_state['daftar_nilai']['Nama'] == s_dinilai].index[0]
                st.session_state['daftar_nilai'].at[idx, 'Speaking Test'] = spk
                st.session_state['daftar_nilai'].at[idx, 'Listening Quiz'] = lis
                st.session_state['daftar_nilai'].at[idx, 'Grammar & Vocab'] = grm
                st.session_state['daftar_nilai'].at[idx, 'Tugas Mandiri'] = tgs
                st.session_state['daftar_nilai'].at[idx, 'Nilai Akhir'] = n_akhir
                st.success(f"✅ Nilai {s_dinilai} berhasil diupdate: {n_akhir:.2f}")

    with tab_g4:
        st.subheader("📌 Terbitkan Tugas Baru")
        with st.form("form_tugas_guru"):
            t_judul = st.text_input("Judul Tugas", placeholder="Contoh: Record Speaking Audio 2 Menit")
            t_kelas = st.selectbox("Untuk Kelas", st.session_state['data_siswa']['Kelas'].unique().tolist())
            t_dl = st.date_input("Deadline", datetime.date.today() + datetime.timedelta(days=7))
            t_inst = st.text_area("Instruksi Tugas Mandiri")

            btn_tugas_pub = st.form_submit_button("🚀 Diterbitkan ke Siswa")
            if btn_tugas_pub:
                new_t = {
                    "ID Tugas": f"TUG-0{len(st.session_state['daftar_penugasan'])+1}",
                    "Judul Tugas": t_judul,
                    "Kelas": t_kelas,
                    "Tenggat Waktu": t_dl,
                    "Status": "Terbit",
                    "Instruksi": t_inst
                }
                st.session_state['daftar_penugasan'] = pd.concat([st.session_state['daftar_penugasan'], pd.DataFrame([new_t])], ignore_index=True)
                st.success("✅ Tugas berhasil diterbitkan!")

    with tab_g5:
        st.subheader("📥 Download Rekap untuk Pengajar")
        st.caption("Unduh rekap jurnal mengajar dan nilai kelas langsung ke HP.")
        
        jurnal_csv = st.session_state['jurnal_mengajar'].to_csv(index=False).encode('utf-8')
        st.download_button(
            label="📖 Download Jurnal Mengajar (.csv)",
            data=jurnal_csv,
            file_name=f"Jurnal_Mengajar_{datetime.date.today()}.csv",
            mime="text/csv",
            use_container_width=True
        )

        nilai_csv = st.session_state['daftar_nilai'].to_csv(index=False).encode('utf-8')
        st.download_button(
            label="💯 Download Daftar Nilai Class (.csv)",
            data=nilai_csv,
            file_name=f"Daftar_Nilai_Class_{datetime.date.today()}.csv",
            mime="text/csv",
            use_container_width=True
        )


# ==============================================================================
# MODE 3: SISWA / MURID (Akses Materi, Pengumpulan Rekaman Suara, Nilai)
# ==============================================================================
elif role == "👤 Mode Siswa / Murid":
    siswa_login = st.selectbox("Login Sebagai Siswa:", st.session_state['data_siswa']['Nama'].tolist())
    st.info(f"Welcome back, **{siswa_login}**! 👋")

    tab_s1, tab_s2, tab_s3, tab_s4 = st.tabs([
        "📚 Modul Materi",
        "🎙️ Pengumpulan Tugas",
        "📊 Nilai Saya",
        "📅 Riwayat Kehadiran"
    ])

    with tab_s1:
        st.subheader("📚 Materi & Modul Pelajaran")
        df_mat = st.session_state['materi_pelajaran']
        for _, m in df_mat.iterrows():
            with st.container():
                st.markdown(f""" <div class='mobile-card'> <b>{m['Kode Modul']} - {m['Judul']}</b><br/> <small>Kelas: {m['Kelas']} | Format: {m['Format']}</small><br/> <a href="{m['Link/Akses']}" target="_blank">🔗 Buka / Download Materi (Google Drive)</a> </div> """, unsafe_allow_html=True)

    with tab_s2:
        st.subheader("🎙️ Pengumpulan Tugas Voice / Speaking HP")
        df_tug = st.session_state['daftar_penugasan']
        st.dataframe(df_tug[['Judul Tugas', 'Tenggat Waktu', 'Instruksi']], use_container_width=True)

        with st.form("form_submit_tugas_siswa"):
            tugas_pilih = st.selectbox("Pilih Tugas yang Dikumpulkan", df_tug['Judul Tugas'].tolist())
            audio_file = st.file_uploader("Upload File Rekaman Suara / Tugas (MP3, WAV, M4A, PDF)", type=['mp3', 'wav', 'm4a', 'pdf'])
            catatan_siswa = st.text_area("Catatan Tambahan untuk Tutor")

            btn_sub_tugas = st.form_submit_button("📤 Kirim Tugas Lengkap")
            if btn_sub_tugas:
                st.success(f"✅ Tugas '{tugas_pilih}' berhasil dikirim ke tutor!")

    with tab_s3:
        st.subheader("📊 Transkrip & Laporan Nilai")
        df_my_nilai = st.session_state['daftar_nilai'][st.session_state['daftar_nilai']['Nama'] == siswa_login]
        if not df_my_nilai.empty:
            row_n = df_my_nilai.iloc[0]
            col_m1, col_m2 = st.columns(2)
            with col_m1:
                st.metric("Speaking Test", f"{row_n['Speaking Test']}")
                st.metric("Grammar & Vocab", f"{row_n['Grammar & Vocab']}")
            with col_m2:
                st.metric("Listening Quiz", f"{row_n['Listening Quiz']}")
                st.metric("Nilai Akhir", f"{row_n['Nilai Akhir']:.2f}")
        else:
            st.info("Nilai belum diinput oleh pengajar.")

    with tab_s4:
        st.subheader("📅 Riwayat Kehadiran Anda")
        my_pres = [p for p in st.session_state['data_presensi'] if p['Nama'] == siswa_login]
        if my_pres:
            df_my_p = pd.DataFrame(my_pres)
            st.dataframe(df_my_p[['Tanggal', 'Modus', 'Status Kehadiran', 'Lokasi/GPS', 'Catatan']], use_container_width=True)
        else:
            st.info("Belum ada catatan presensi.")


# ==============================================================================
# MODE 4: FORM PENDAFTARAN MANDIRI (PUBLIC)
# ==============================================================================
elif role == "📝 Form Pendaftaran Mandiri (Public)":
    st.subheader("📝 Formulir Pendaftaran Siswa Baru")
    st.caption("Formulir pendaftaran online untuk calon siswa les bahasa Inggris.")

    with st.form("form_pendaftaran_public"):
        reg_nama = st.text_input("Nama Lengkap Calon Siswa *")
        reg_wa = st.text_input("Nomor WhatsApp Aktif (Contoh: 6281234567890) *", value="6281234567890")
        reg_kelas = st.selectbox("Program Les yang Diminati *", ["General Speaking - A", "IELTS Prep - B", "Private Conversation", "Kids English"])
        reg_tipe = st.selectbox("Tipe Pembelajaran Utama *", ["Les Online", "Tatap Muka", "Privat Rumah"])
        reg_jadwal = st.multiselect("Pilihan Hari & Waktu yang Diinginkan:", ["Senin & Rabu (16:00 WIB)", "Selasa & Kamis (16:00 WIB)", "Sabtu (09:00 WIB)", "Fleksibel (Privat)"], default=["Senin & Rabu (16:00 WIB)"])
        reg_alamat = st.text_area("Alamat Lengkap (Wajib Diisi jika Memilih Privat Rumah):", placeholder="Jl. Anggrek No. 15, Jakarta...")

        btn_reg_submit = st.form_submit_button("📋 Kirim Pendaftaran Baru")

        if btn_reg_submit:
            if not reg_nama:
                st.error("Mohon isi Nama Lengkap terlebih dahulu!")
            else:
                new_reg_id = f"S0{len(st.session_state['data_siswa'])+1}"
                new_s_row = {"ID": new_reg_id, "Nama": reg_nama, "No WA": reg_wa, "Kelas": reg_kelas, "Tipe": reg_tipe, "Status": "Pending / Reg"}
                st.session_state['data_siswa'] = pd.concat([st.session_state['data_siswa'], pd.DataFrame([new_s_row])], ignore_index=True)

                new_n_row = {"ID Siswa": new_reg_id, "Nama": reg_nama, "Kelas": reg_kelas, "Speaking Test": 0, "Listening Quiz": 0, "Grammar & Vocab": 0, "Tugas Mandiri": 0, "Nilai Akhir": 0.0}
                st.session_state['daftar_nilai'] = pd.concat([st.session_state['daftar_nilai'], pd.DataFrame([new_n_row])], ignore_index=True)

                st.success("🎉 Pendaftaran Anda Berhasil Diterima!")

                st.markdown(f""" <div class='download-btn-card'> <h4>🎟️ Bukti Pendaftaran Siswa Baru</h4> <b>ID Registrasi:</b> {new_reg_id}<br/> <b>Nama Siswa:</b> {reg_nama}<br/> <b>Program:</b> {reg_kelas} ({reg_tipe})<br/> <b>No WA:</b> {reg_wa}<br/> <b>Status:</b> Terdaftar di Sistem<br/> <small>Tim admin kami akan segera menghubungi nomor WhatsApp Anda untuk konfirmasi jadwal & pembayaran.</small> </div> """, unsafe_allow_html=True)
