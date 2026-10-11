import streamlit as st
import pandas as pd
import numpy as np
import datetime
from PIL import Image
import io

# --- PAGE CONFIGURATION (MOBILE OPTIMIZED) ---
st.set_page_config(
    page_title="App Presensi & LMS Les Bahasa Inggris",
    page_icon="📱",
    layout="wide",
    initial_sidebar_state="collapsed"
)

# --- CUSTOM MOBILE-FRIENDLY CSS ---
st.markdown("""
    <style>
    .main-title {
        font-size: 22px;
        font-weight: 800;
        color: #1E3A8A;
        text-align: center;
        margin-bottom: 5px;
    }
    .role-badge {
        background-color: #DBEAFE;
        color: #1E40AF;
        padding: 6px 12px;
        border-radius: 20px;
        font-size: 13px;
        font-weight: bold;
        text-align: center;
        margin-bottom: 15px;
    }
    .mobile-card {
        background-color: #FFFFFF;
        border: 1px solid #E5E7EB;
        padding: 12px;
        border-radius: 12px;
        box-shadow: 0 2px 4px rgba(0,0,0,0.05);
        margin-bottom: 10px;
    }
    .stButton>button {
        width: 100%;
        border-radius: 10px;
        height: 45px;
        font-weight: bold;
    }
    .download-btn-card {
        background-color: #F0FDF4;
        border: 1px solid #BBF7D0;
        padding: 15px;
        border-radius: 12px;
        margin-top: 10px;
        margin-bottom: 15px;
    }
    .auth-card {
        background-color: #FEF3C7;
        border: 1px solid #F59E0B;
        padding: 15px;
        border-radius: 12px;
        margin-bottom: 15px;
    }
    </style>
""", unsafe_allow_html=True)

# --- PIN / PASSWORD KONFIGURASI ---
PIN_ADMIN_DEFAULT = "rasiya21"
PIN_GURU_DEFAULT = "3v@semok"

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
    st.session_state['data_siswa'] = pd.DataFrame(columns=["ID", "Nama", "No WA", "Kelas", "Tipe", "Status"])

if 'data_presensi' not in st.session_state:
    st.session_state['data_presensi'] = []

if 'jurnal_mengajar' not in st.session_state:
    st.session_state['jurnal_mengajar'] = pd.DataFrame(columns=["Pertemuan Ke", "Tanggal", "Kelas", "Modus", "Pengajar", "Topik/Materi", "Ringkasan Pembelajaran"])

if 'materi_pelajaran' not in st.session_state:
    st.session_state['materi_pelajaran'] = pd.DataFrame(columns=["Kode Modul", "Judul", "Kelas", "Format", "Link/Akses"])

if 'daftar_nilai' not in st.session_state:
    st.session_state['daftar_nilai'] = pd.DataFrame(columns=["ID Siswa", "Nama", "Kelas", "Speaking Test", "Listening Quiz", "Grammar & Vocab", "Tugas Mandiri", "Nilai Akhir"])

if 'daftar_penugasan' not in st.session_state:
    st.session_state['daftar_penugasan'] = pd.DataFrame(columns=["ID Tugas", "Judul Tugas", "Kelas", "Tenggat Waktu", "Status", "Instruksi"])

if 'admin_auth' not in st.session_state:
    st.session_state['admin_auth'] = False

if 'guru_auth' not in st.session_state:
    st.session_state['guru_auth'] = False

# --- TOP MOBILE HEADER & ROLE SWITCHER ---
st.markdown("<div class='main-title'>📱 Presensi & LMS English Class</div>", unsafe_allow_html=True)

# Default index is set to Mode Siswa so public users open directly as Siswa
role = st.selectbox(
    "📲 Pilih Mode Akses:",
    ["👤 Mode Siswa / Murid", "📝 Form Pendaftaran Mandiri (Public)", "👨‍🏫 Mode Guru / Tutor (🔒 PIN Required)", "👑 Mode Admin / Pemilik Les (🔒 PIN Required)"],
    index=0
)

st.markdown(f"<div class='role-badge'>Akses: {role}</div>", unsafe_allow_html=True)

# ==============================================================================
# MODE 1: SISWA / MURID (TIDAK PERLU PIN - DIBUKA LANGSUNG UNTUK SISWA)
# ==============================================================================
if role == "👤 Mode Siswa / Murid":
    if not st.session_state['data_siswa'].empty:
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
            if not df_mat.empty:
                for _, m in df_mat.iterrows():
                    with st.container():
                        st.markdown(f"""
                        <div class='mobile-card'>
                            <b>{m['Kode Modul']} - {m['Judul']}</b><br/>
                            <small>Kelas: {m['Kelas']} | Format: {m['Format']}</small><br/>
                            <a href="{m['Link/Akses']}" target="_blank">🔗 Buka / Download Materi (Google Drive)</a>
                        </div>
                        """, unsafe_allow_html=True)
            else:
                st.info("ℹ️ Belum ada modul materi yang diunggah pengajar.")

        with tab_s2:
            st.subheader("🎙️ Pengumpulan Tugas Voice / Speaking HP")
            df_tug = st.session_state['daftar_penugasan']
            if not df_tug.empty:
                st.dataframe(df_tug[['Judul Tugas', 'Tenggat Waktu', 'Instruksi']], use_container_width=True)

                with st.form("form_submit_tugas_siswa"):
                    tugas_pilih = st.selectbox("Pilih Tugas yang Dikumpulkan", df_tug['Judul Tugas'].tolist())
                    audio_file = st.file_uploader("Upload File Rekaman Suara / Tugas (MP3, WAV, M4A, PDF)", type=['mp3', 'wav', 'm4a', 'pdf'])
                    catatan_siswa = st.text_area("Catatan Tambahan untuk Tutor")

                    btn_sub_tugas = st.form_submit_button("📤 Kirim Tugas Lengkap")
                    if btn_sub_tugas:
                        st.success(f"✅ Tugas '{tugas_pilih}' berhasil dikirim ke tutor!")
            else:
                st.info("ℹ️ Belum ada penugasan aktif.")

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
    else:
        st.warning("⚠️ Belum ada siswa terdaftar di sistem. Silakan mendaftar via 'Form Pendaftaran Mandiri'.")


# ==============================================================================
# MODE 2: FORM PENDAFTARAN MANDIRI (PUBLIC - TIDAK PERLU PIN)
# ==============================================================================
elif role == "📝 Form Pendaftaran Mandiri (Public)":
    st.subheader("📝 Formulir Pendaftaran Siswa Baru")
    st.caption("Formulir pendaftaran online untuk calon siswa les bahasa Inggris.")

    with st.form("form_pendaftaran_public"):
        reg_nama = st.text_input("Nama Lengkap Calon Siswa *")
        reg_wa = st.text_input("Nomor WhatsApp Aktif (Contoh: 6281234567890) *", value="6281234567890")
        reg_kelas = st.selectbox("Program Les *", ["Reguler", "Publik"])
        reg_tipe = st.selectbox("Tipe Pembelajaran Utama *", ["Les Online", "Tatap Muka", "Privat Rumah"])
        reg_hari = st.multiselect("Pilihan Hari yang Diinginkan *", ["Senin", "Selasa", "Rabu", "Kamis", "Jumat", "Sabtu", "Minggu"], default=["Senin"])
        reg_waktu = st.text_input("Waktu / Jam yang Diinginkan (Ketik Manual) *", placeholder="Contoh: 16:00 WIB / Jam 4 Sore")
        reg_alamat = st.text_area("Alamat Lengkap (Wajib Diisi jika Memilih Privat Rumah):", placeholder="Jl. Anggrek No. 15, Jakarta...")

        btn_reg_submit = st.form_submit_button("📋 Kirim Pendaftaran Baru")

        if btn_reg_submit:
            if not reg_nama:
                st.error("Mohon isi Nama Lengkap terlebih dahulu!")
            else:
                new_reg_id = f"S0{len(st.session_state['data_siswa'])+1}"
                new_s_row = {"ID": new_reg_id, "Nama": reg_nama, "No WA": reg_wa, "Kelas": reg_kelas, "Tipe": reg_tipe, "Status": "Aktif"}
                st.session_state['data_siswa'] = pd.concat([st.session_state['data_siswa'], pd.DataFrame([new_s_row])], ignore_index=True)

                new_n_row = {"ID Siswa": new_reg_id, "Nama": reg_nama, "Kelas": reg_kelas, "Speaking Test": 0, "Listening Quiz": 0, "Grammar & Vocab": 0, "Tugas Mandiri": 0, "Nilai Akhir": 0.0}
                st.session_state['daftar_nilai'] = pd.concat([st.session_state['daftar_nilai'], pd.DataFrame([new_n_row])], ignore_index=True)

                st.success("🎉 Pendaftaran Anda Berhasil Diterima!")

                st.markdown(f"""
                <div class='download-btn-card'>
                    <h4>🎟️ Bukti Pendaftaran Siswa Baru</h4>
                    <b>ID Registrasi:</b> {new_reg_id}<br/>
                    <b>Nama Siswa:</b> {reg_nama}<br/>
                    <b>Program Les:</b> {reg_kelas} ({reg_tipe})<br/>
                    <b>Hari & Waktu:</b> {", ".join(reg_hari) if reg_hari else "Belum dipilih"} ({reg_waktu})<br/>
                    <b>No WA:</b> {reg_wa}<br/>
                    <b>Status:</b> Terdaftar di Sistem<br/>
                    <small>Tim admin kami akan segera menghubungi nomor WhatsApp Anda untuk konfirmasi jadwal & pembayaran.</small>
                </div>
                """, unsafe_allow_html=True)


# ==============================================================================
# MODE 3: GURU / TUTOR (🔒 DILINDUNGI PIN GURU)
# ==============================================================================
elif role == "👨‍🏫 Mode Guru / Tutor (🔒 PIN Required)":
    if not st.session_state['guru_auth']:
        st.markdown("<div class='auth-card'>", unsafe_allow_html=True)
        st.subheader("🔒 Keamanan Akses Guru / Tutor")
        st.caption("Akses khusus Pengajar & Tutor. Masukkan PIN untuk membuka menu absensi dan input nilai.")
        
        pin_input_guru = st.text_input("🔑 Masukkan PIN Guru:", type="password")
        btn_login_guru = st.button("🔓 Masuk Mode Guru")
        
        if btn_login_guru:
            if pin_input_guru == PIN_GURU_DEFAULT:
                st.session_state['guru_auth'] = True
                st.success("✅ PIN Benar! Mengakses Mode Guru...")
                st.rerun()
            else:
                st.error("❌ PIN Guru Salah! Akses ditolak.")
        st.markdown("</div>", unsafe_allow_html=True)

    else:
        col_logout_g, _ = st.columns([1, 3])
        with col_logout_g:
            if st.button("🔒 Keluar Mode Guru"):
                st.session_state['guru_auth'] = False
                st.rerun()

        tab_g1, tab_g2, tab_g3, tab_g4 = st.tabs([
            "📸 Absensi & Selfie",
            "📖 Jurnal Mengajar",
            "💯 Input Nilai",
            "📌 Terbit Tugas"
        ])

        with tab_g1:
            st.subheader("📸 Input Presensi Siswa (Kamera HP)")
            st.caption("Gunakan kamera HP untuk verifikasi foto selfie / lokasi saat kelas privat atau tatap muka.")

            if not st.session_state['data_siswa'].empty:
                with st.form("form_presensi_kamera", clear_on_submit=True):
                    s_nama = st.selectbox("Pilih Siswa", options=st.session_state['data_siswa']['Nama'].tolist())
                    selected_student = st.session_state['data_siswa'][st.session_state['data_siswa']['Nama'] == s_nama].iloc[0]
                    
                    col1, col2 = st.columns(2)
                    with col1:
                        tgl_sesi = st.date_input("Tanggal", datetime.date.today())
                        modus = st.selectbox("Modus Pembelajaran", ["Privat Rumah", "Tatap Muka", "Les Online"], index=["Privat Rumah", "Tatap Muka", "Les Online"].index(selected_student['Tipe']) if selected_student['Tipe'] in ["Privat Rumah", "Tatap Muka", "Les Online"] else 0)
                    with col2:
                        status_kehadiran = st.selectbox("Status Kehadiran", ["Hadir", "Izin", "Sakit", "Alpha"])
                        lokasi_gps = st.text_input("Koordinat GPS / Alamat / Platform", value="-6.2088, 106.8456 (Verified GPS HP)" if modus=="Privat Rumah" else ("Ruang Kelas" if modus=="Tatap Muka" else "Zoom Meeting"))

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
            else:
                st.warning("⚠️ Belum ada siswa terdaftar. Tambahkan siswa terlebih dahulu di Mode Admin atau lewat Form Pendaftaran Mandiri.")

        with tab_g2:
            st.subheader("📖 Isi Jurnal Mengajar Harian")
            with st.form("form_jurnal_guru"):
                p_ke = st.number_input("Pertemuan Ke-", min_value=1, value=len(st.session_state['jurnal_mengajar']) + 1)
                tgl_j = st.date_input("Tanggal Pertemuan", datetime.date.today())
                kls_j = st.selectbox("Kelas Target", ["Reguler", "Publik"])
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
            if not st.session_state['daftar_nilai'].empty:
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
            else:
                st.info("ℹ️ Belum ada siswa untuk dinilai.")

        with tab_g4:
            st.subheader("📌 Terbitkan Tugas Baru")
            with st.form("form_tugas_guru"):
                t_judul = st.text_input("Judul Tugas", placeholder="Contoh: Record Speaking Audio 2 Menit")
                t_kelas = st.selectbox("Untuk Kelas", ["Reguler", "Publik"])
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


# ==============================================================================
# MODE 4: ADMIN / PEMILIK LES (🔒 DILINDUNGI PIN ADMIN)
# ==============================================================================
elif role == "👑 Mode Admin / Pemilik Les (🔒 PIN Required)":
    if not st.session_state['admin_auth']:
        st.markdown("<div class='auth-card'>", unsafe_allow_html=True)
        st.subheader("🔒 Keamanan Akses Admin / Pemilik Les")
        st.caption("Area rahasia Admin. Masukkan PIN Admin untuk mengakses rekapitulasi, ekspor data, dan pengelolaan sistem.")
        
        pin_input_admin = st.text_input("🔑 Masukkan PIN Admin:", type="password")
        btn_login_admin = st.button("🔓 Masuk Mode Admin")
        
        if btn_login_admin:
            if pin_input_admin == PIN_ADMIN_DEFAULT:
                st.session_state['admin_auth'] = True
                st.success("✅ PIN Benar! Mengakses Mode Admin...")
                st.rerun()
            else:
                st.error("❌ PIN Admin Salah! Akses ditolak.")
        st.markdown("</div>", unsafe_allow_html=True)

    else:
        col_logout_a, _ = st.columns([1, 3])
        with col_logout_a:
            if st.button("🔒 Keluar Mode Admin"):
                st.session_state['admin_auth'] = False
                st.rerun()

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
            if not st.session_state['data_siswa'].empty:
                st.write("📈 **Distribusi Modus Pembelajaran Siswa:**")
                st.bar_chart(st.session_state['data_siswa']['Tipe'].value_counts())
            else:
                st.info("ℹ️ Belum ada data siswa terdaftar.")

        with tab_a2:
            st.subheader("📥 Ekspor Laporan Otomatis ke HP")
            st.caption("Unduh laporan lengkap dalam format **Excel (.xlsx)** atau **CSV** langsung ke HP Anda.")

            presensi_export_data = []
            for p in st.session_state['data_presensi']:
                p_copy = p.copy()
                p_copy['Ada Foto Bukti'] = "Ya" if p.get('Foto Bukti') is not None else "Tidak"
                p_copy.pop('Foto Bukti', None)
                presensi_export_data.append(p_copy)
            
            df_p_export = pd.DataFrame(presensi_export_data)

            col_f1, col_f2 = st.columns(2)
            with col_f1:
                st_date = st.date_input("Dari Tanggal", datetime.date.today() - datetime.timedelta(days=30))
            with col_f2:
                end_date = st.date_input("Sampai Tanggal", datetime.date.today())

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

        with tab_a3:
            st.subheader("📷 Galeri Rekap Presensi & Bukti Selfie / GPS")
            if st.session_state['data_presensi']:
                for idx, item in enumerate(reversed(st.session_state['data_presensi'])):
                    with st.expander(f"📍 {item['Tanggal']} | {item['Nama']} - {item['Modus']} ({item['Status Kehadiran']})"):
                        st.write(f"**Kelas:** {item['Kelas']}")
                        st.write(f"**Lokasi / Titik GPS:** {item['Lokasi/GPS']}")
                        st.write(f"**Catatan:** {item['Catatan']}")
                        
                        if item.get('Foto Bukti') is not None:
                            try:
                                image = Image.open(io.BytesIO(item['Foto Bukti']))
                                st.image(image, caption=f"Foto Selfie / Verifikasi Lokasi ({item['Nama']})", width=300)
                            except Exception as e:
                                st.caption("📷 Foto tersimpan.")
                        else:
                            st.caption("ℹ️ Tidak ada lampiran foto.")
            else:
                st.info("ℹ️ Belum ada riwayat presensi yang dicatat.")

        with tab_a4:
            st.subheader("⚙️ Kelola Data Siswa & Kursus")
            st.dataframe(st.session_state['data_siswa'], use_container_width=True)

            with st.expander("➕ Tambah Siswa Baru (Manual oleh Admin)"):
                with st.form("form_tambah_siswa_admin"):
                    s_id = f"S0{len(st.session_state['data_siswa'])+1}"
                    s_n = st.text_input("Nama Siswa Baru")
                    s_wa = st.text_input("Nomor WhatsApp (Contoh: 6281234567890)", value="6281234567890")
                    s_k = st.selectbox("Kelas / Program", ["Reguler", "Publik"])
                    s_t = st.selectbox("Tipe Pembelajaran", ["Privat Rumah", "Tatap Muka", "Les Online"])

                    btn_add_siswa = st.form_submit_button("➕ Daftarkan Siswa")
                    if btn_add_siswa:
                        if s_n:
                            new_s_data = {"ID": s_id, "Nama": s_n, "No WA": s_wa, "Kelas": s_k, "Tipe": s_t, "Status": "Aktif"}
                            st.session_state['data_siswa'] = pd.concat([st.session_state['data_siswa'], pd.DataFrame([new_s_data])], ignore_index=True)
                            
                            new_n_data = {"ID Siswa": s_id, "Nama": s_n, "Kelas": s_k, "Speaking Test": 0, "Listening Quiz": 0, "Grammar & Vocab": 0, "Tugas Mandiri": 0, "Nilai Akhir": 0.0}
                            st.session_state['daftar_nilai'] = pd.concat([st.session_state['daftar_nilai'], pd.DataFrame([new_n_data])], ignore_index=True)
                            st.success(f"✅ Siswa {s_n} berhasil didaftarkan!")
                        else:
                            st.error("Mohon isi nama siswa.")
