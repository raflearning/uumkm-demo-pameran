import streamlit as st
import pandas as pd
import time
import google.generativeai as genai
from google.api_core.exceptions import InternalServerError
from dotenv import load_dotenv
from vis_interpret import (
    visualize_pelanggan, visualize_produk, visualize_transaksi_penjualan, 
    visualize_lokasi_penjualan, visualize_staf_penjualan, visualize_inventaris, 
    visualize_promosi_pemasaran, visualize_feedback_pengembalian, 
    visualize_analisis_penjualan, visualize_lainnya
)

st.title("🚀Pantau Kinerja Bisnis Kamu!")
st.write(
    "Memudahkan kamu untuk mengambil informasi bisnis dan rekomendasi pengambilan keputusan dengan Artificial Intelligence!"
)

# Section divider
st.markdown("---")

# Ambil API key dari variabel lingkungan
API_KEY = st.secrets["general"]["API_KEY"]
genai.configure(api_key=API_KEY)
model = genai.GenerativeModel(model_name='gemini-1.5-pro-latest')  # Model untuk interpretasi dan chatbot

# Function to load data from all sheets
def load_data(uploaded_file):
    if uploaded_file is not None:
        data = pd.read_excel(uploaded_file, sheet_name=None)
        return data
    else:
        return None

# Function to get business info options based on selected sheet
def get_business_options(sheet_name):
    options = {
        'Pelanggan': [
            'Analisis demografi pelanggan',
            'Distribusi usia dan jenis kelamin pelanggan',
            'Segmentasi pelanggan berdasarkan preferensi'
        ],
        'Produk': [
            'Kinerja penjualan produk dan stok',
            'Distribusi penjualan berdasarkan kategori produk',
            'Analisis harga produk dan trend penjualan'
        ],
        'Transaksi Penjualan': [
            'Jumlah penjualan, pendapatan, dan metode pembayaran',
            'Tren penjualan',
            'Analisis status penjualan dan metode pembayaran'
        ],
        'Lokasi Penjualan': [
            'Kinerja penjualan di berbagai lokasi',
            'Distribusi penjualan berdasarkan kota/provinsi',
            'Analisis lokasi dengan penjualan tertinggi/rendah'
        ],
        'Staf Penjualan': [
            'Kinerja dan komisi staf penjualan',
            'Analisis penilaian kinerja staf',
            'Distribusi staf berdasarkan posisi/jabatan'
        ],
        'Inventaris': [
            'Manajemen stok produk',
            'Tren stok masuk dan keluar',
            'Analisis produk dengan stok terbanyak/terkecil'
        ],
        'Promosi dan Pemasaran': [
            'Efektivitas kampanye promosi',
            'Distribusi penjualan berdasarkan media promosi',
            'Analisis kode diskon promosi'
        ],
        'Feedback dan Pengembalian': [
            'Masalah dan kepuasan pelanggan',
            'Distribusi alasan pengembalian produk',
            'Status pengembalian produk'
        ],
        'Analisis Penjualan': [
            'Penjualan agregat dan tren',
            'Analisis penjualan berdasarkan produk/kategori',
            'Tren penjualan/tahunan'
        ],
        'Lainnya': [
            'Analisis tambahan dan faktor eksternal',
            'Tren penjualan berdasarkan faktor ekonomi',
            'Distribusi biaya operasional terkait penjualan'
        ]
    }
    return options.get(sheet_name, [])

# Streamlit app
st.sidebar.header("Unggah Data Penjualan Bisnis Kamu")
uploaded_file = st.sidebar.file_uploader("Unggah file Excel", type=["xlsx"])
data = load_data(uploaded_file)

if data is not None:
    st.sidebar.success("Data berhasil diunggah!")
    sheet_names = list(data.keys())
else:
    st.sidebar.warning("Silakan unggah file Excel untuk melanjutkan.")
    sheet_names = []

# Navigation bar to select sheet
selected_sheet = st.sidebar.selectbox("Pilih Kategori Data", [""] + sheet_names)

if data is not None and selected_sheet:
    if selected_sheet != "":
        st.write(f"### 📶Dashboard Analitik - {selected_sheet}")
        sheet_data = data[selected_sheet]
        st.write("##### Data yang Diunggah")
        st.dataframe(sheet_data)

        st.write("##### 👇Pilih Informasi Bisnis yang Kamu Inginkan")
        business_options = get_business_options(selected_sheet)
        selected_business_info = st.selectbox("", [""] + business_options)

        if selected_business_info and selected_business_info != "":
            # Function to get visualization and interpretation
            def get_visualization_and_interpretation(sheet_data, selected_business_info, selected_sheet, model):
                if selected_sheet == 'Pelanggan':
                    return visualize_pelanggan(sheet_data, selected_business_info, model)
                elif selected_sheet == 'Produk':
                    return visualize_produk(sheet_data, selected_business_info, model)
                elif selected_sheet == 'Transaksi Penjualan':
                    return visualize_transaksi_penjualan(sheet_data, selected_business_info, model)
                elif selected_sheet == 'Lokasi Penjualan':
                    return visualize_lokasi_penjualan(sheet_data, selected_business_info, model)
                elif selected_sheet == 'Staf Penjualan':
                    return visualize_staf_penjualan(sheet_data, selected_business_info, model)
                elif selected_sheet == 'Inventaris':
                    return visualize_inventaris(sheet_data, selected_business_info, model)
                elif selected_sheet == 'Promosi dan Pemasaran':
                    return visualize_promosi_pemasaran(sheet_data, selected_business_info, model)
                elif selected_sheet == 'Feedback dan Pengembalian':
                    return visualize_feedback_pengembalian(sheet_data, selected_business_info, model)
                elif selected_sheet == 'Analisis Penjualan':
                    return visualize_analisis_penjualan(sheet_data, selected_business_info, model)
                elif selected_sheet == 'Lainnya':
                    return visualize_lainnya(sheet_data, selected_business_info, model)
                else:
                    return [], ""

            # Get visualization and interpretation
            try:
                charts, interpretation = get_visualization_and_interpretation(sheet_data, selected_business_info, selected_sheet, model)
            except InternalServerError as e:
                st.error("Terjadi kesalahan pada server saat mencoba mendapatkan interpretasi. Silakan coba lagi nanti.")
                st.stop()

            # Function to display charts one by one
            def display_charts(charts):
                for chart in charts:
                    figure = chart.get('figure')
                    try:
                        st.plotly_chart(figure)
                    except Exception as e:
                        st.write(f"### Error: Could not display Plotly figure. Error: {e}")

            # Display charts
            display_charts(charts)

            # Function to display interpretation one character at a time
            def display_interpretation_one_by_one(interpretation):
                if interpretation:
                    interpretation_text = ""
                    interpretation_box = st.empty()
                    for i in range(len(interpretation)):
                        interpretation_text += interpretation[i]
                        interpretation_box.markdown(interpretation_text)
                        time.sleep(0.004)  # Adjust the speed of typing effect
                    return interpretation_text
                return ""

            # Display interpretation
            interpretation_text = display_interpretation_one_by_one(interpretation)
            st.markdown("---")

            # Create a container for the chatbot section that appears after interpretation
            def display_chatbot(model, charts, interpretation_text):
                st.write("### 💬Chatbot AI")
                st.write("Kamu masih punya pertanyaan terkait hasil visualisasinya? Tanyakan di bawah ya!")

                # Input box for user questions
                user_question = st.text_input("Ajukan pertanyaan kamu di sini...")

                if user_question:
                    try:
                        response = model.generate_content(
                            f"Pertanyaan: {user_question}\n"
                            f"Chart yang telah divisualkan: {charts}\n"
                            f"Hasil interpretasi: {interpretation_text}\n"
                            "Jawab dalam konteks bisnis."
                        )
                        chatbot_response = response.text

                        # Display the response as typing effect
                        st.write("#### Jawaban Chatbot:")
                        typing_response = ""
                        typing_box = st.empty()
                        for i in range(len(chatbot_response)):
                            typing_response += chatbot_response[i]
                            typing_box.markdown(typing_response)
                            time.sleep(0.004)  # Adjust the speed of typing effect
                    except Exception as e:
                        st.write(f"### Error: {e}")

            # Display chatbot
            display_chatbot(model, charts, interpretation_text)
