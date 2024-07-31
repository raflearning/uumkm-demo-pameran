import streamlit as st
import pandas as pd
import time
import google.generativeai as genai
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
selected_sheet = st.sidebar.selectbox("Pilih Kategori Data", sheet_names)

if data is not None and selected_sheet:
    st.write(f"### 📶Dashboard Analitik - {selected_sheet}")
    sheet_data = data[selected_sheet]
    st.write("##### Data yang Diunggah")
    st.dataframe(sheet_data)

    st.write("##### 👇Pilih Informasi Bisnis yang Kamu Inginkan")
    business_options = get_business_options(selected_sheet)
    selected_business_info = st.selectbox("", business_options)

    if selected_business_info:
        # Initialize variables for storing charts and interpretation
        charts = []
        interpretation = ""
        model_interpretasi = genai.GenerativeModel(model_name='gemini-1.5-pro-latest') # Model untuk interpretasi

        # Call appropriate visualization function based on the selected sheet
        if selected_sheet == 'Pelanggan':
            charts, interpretation = visualize_pelanggan(sheet_data, selected_business_info, model_interpretasi)
        elif selected_sheet == 'Produk':
            charts, interpretation = visualize_produk(sheet_data, selected_business_info, model_interpretasi)
        elif selected_sheet == 'Transaksi Penjualan':
            charts, interpretation = visualize_transaksi_penjualan(sheet_data, selected_business_info, model_interpretasi)
        elif selected_sheet == 'Lokasi Penjualan':
            charts, interpretation = visualize_lokasi_penjualan(sheet_data, selected_business_info, model_interpretasi)
        elif selected_sheet == 'Staf Penjualan':
            charts, interpretation = visualize_staf_penjualan(sheet_data, selected_business_info, model_interpretasi)
        elif selected_sheet == 'Inventaris':
            charts, interpretation = visualize_inventaris(sheet_data, selected_business_info, model_interpretasi)
        elif selected_sheet == 'Promosi dan Pemasaran':
            charts, interpretation = visualize_promosi_pemasaran(sheet_data, selected_business_info, model_interpretasi)
        elif selected_sheet == 'Feedback dan Pengembalian':
            charts, interpretation = visualize_feedback_pengembalian(sheet_data, selected_business_info, model_interpretasi)
        elif selected_sheet == 'Analisis Penjualan':
            charts, interpretation = visualize_analisis_penjualan(sheet_data, selected_business_info, model_interpretasi)
        elif selected_sheet == 'Lainnya':
            charts, interpretation = visualize_lainnya(sheet_data, selected_business_info, model_interpretasi)

        # Save interpretation and charts in session state
        st.session_state['interpretation_text'] = interpretation
        st.session_state['charts'] = charts

        # Display all the relevant charts first
        if charts:
            for chart in st.session_state['charts']:
                figure = chart.get('figure')
                try:
                    # Directly display the Plotly figure
                    st.plotly_chart(figure)
                except Exception as e:
                    st.write(f"### Error: Could not display Plotly figure. Error: {e}")

            # Display interpretation after the charts
            interpretation_text = ""
            interpretation_box = st.empty()
            for i in range(len(st.session_state['interpretation_text'])):
                interpretation_text += st.session_state['interpretation_text'][i]
                interpretation_box.markdown(interpretation_text)
                time.sleep(0.005)  # Adjust the speed of typing effect

    # Create a container for the chatbot section that appears after interpretation
    with st.container():
        st.markdown("---")
        st.write("### 💬Chatbot AI")
        st.write("Kamu masih punya pertanyaan terkait hasil visualisasinya? Tanyakan di bawah ya!")

        # Input box for user questions
        user_question = st.text_input("Ajukan pertanyaan kamu di sini...")

        if user_question:
            try:
                model_chatbot = genai.GenerativeModel(model_name='gemini-1.5-pro-latest')  # Model terpisah untuk chatbot
                general_chatbot_prompt = (
                    f"""
                    Kamu adalah seorang data analyst dan business intelligence handal dan profesional. Tugas kamu adalah menjawab pertanyaan dari user terkait hasil interpretasi pada. Gunakan bahasa yang lumayan santai, mudah dipahami, beginner hingga expert friendly, dan tetap bercirikhas bisnis.
                    Interpretasikan secara spesifik dan mendalam dalam konteks bisnis yang sesuai dan memberikan rekomendasi yang dapat membangun bisnis untuk ke depannya.
                    Perhatikan chart dengan detail, jelaskan data-datanya, dan sampaikan semua informasi yang bermanfaat kepada pelaku UMKM.
                    Tekankan kalimat atau kata yang penting dengan **bold**/underline/italic. Buatkan poin-poin atau tabel jika perlu.
                    Berikan judul yang sesuai dengan topik dan juga 1 emoji di depan judul yang sesuai dengan yang Kamu interpretasikan supaya user UMKM paham akan data yang dibahas.
                    """
                )
                response = model_chatbot.generate_content(f"Prompt: {general_chatbot_prompt}\nPertanyaan: {user_question}\nData: {st.session_state['interpretation_text']}")

                chatbot_response = response.text

                # Display the response as typing effect
                st.write("#### Jawaban Chatbot:")
                typing_response = ""
                typing_box = st.empty()
                for i in range(len(chatbot_response)):
                    typing_response += chatbot_response[i]
                    typing_box.markdown(typing_response)
                    time.sleep(0.005)  # Adjust the speed of typing effect
            except Exception as e:
                st.write(f"### Error: {e}")
