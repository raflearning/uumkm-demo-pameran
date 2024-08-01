from io import BytesIO
from PIL import Image
import plotly.express as px
import pandas as pd

def fig_to_pil_image(fig):
    buf = BytesIO()
    fig.write_image(buf, format='png')
    buf.seek(0)
    image = Image.open(buf)
    return image

def interpret_chart(sheet_name, charts, model):
    interpretations = []
    
    for chart in charts:
        general_prompt = (
            f"""
            Kamu adalah seorang data analyst dan business intelligence handal dan profesional. Tugas Kamu adalah menginterpretasikan data 
            penjualan UMKM dari sheet {sheet_name}. Gunakan bahasa yang lumayan santai, mudah dipahami, beginner hingga expert friendly, dan tetap bercirikhas bisnis.
            Interpretasikan secara spesifik dan mendalam dalam konteks bisnis yang sesuai dan memberikan rekomendasi yang dapat membangun bisnis untuk ke depannya.
            Perhatikan chart dengan detail, jelaskan data-datanya, dan sampaikan semua informasi yang bermanfaat kepada pelaku UMKM.
            Tekankan kalimat atau kata yang penting dengan bold/underline/italic, sesuaikan saja. Buatkan poin-poin atau tabel jika perlu.
            Setiap chart yang Kamu interpretasikan, berikan judul dengan ukuran font yang lebih besar yang sesuai dengan nama data dan juga 1 emoji di depan judul yang sesuai dengan yang Kamu interpretasikan supaya user UMKM paham akan data yang dibahas.

            Berikut adalah visualisasi yang tersedia:
            """
        )
        
        chart_image = fig_to_pil_image(chart['figure'])
        chart_prompt = f"Tipe Visualisasi: {chart['type']}. Interpretasikan data berikut:"
        combined_prompt = f"{general_prompt}\n{chart_prompt}"
        response = model.generate_content([combined_prompt, chart_image])
        chart_description = response.text.strip()
        interpretations.append(chart_description)
    
    return "\n\n".join(interpretations)


def convert_to_date(df, columns):
    for col in columns:
        if (col in df.columns) and (df[col].dtype != 'datetime64[ns]'):
            df[col] = pd.to_datetime(df[col], errors='coerce')
    return df

def visualize_pelanggan(df, selected_business_info, model):
    df = convert_to_date(df, ['Tanggal'])
    charts = []

    if selected_business_info == 'Analisis demografi pelanggan':
        if 'Jenis Kelamin Pelanggan' in df.columns:
            gender_counts = df['Jenis Kelamin Pelanggan'].value_counts().reset_index()
            gender_counts.columns = ['Jenis Kelamin Pelanggan', 'Jumlah']
            charts.append({
                'type': 'Analisis demografi pelanggan',
                'figure': px.bar(data_frame=gender_counts, 
                                 x='Jenis Kelamin Pelanggan', 
                                 y='Jumlah', 
                                 labels={'Jenis Kelamin Pelanggan': 'Jenis Kelamin Pelanggan', 'Jumlah': 'Jumlah'})
            })
        if 'Umur Pelanggan' in df.columns:
            age_counts = df['Umur Pelanggan'].value_counts().reset_index()
            age_counts.columns = ['Umur Pelanggan', 'Jumlah']
            charts.append({
                'type': 'Analisis demografi pelanggan',
                'figure': px.bar(data_frame=age_counts, 
                                 x='Umur Pelanggan', 
                                 y='Jumlah', 
                                 labels={'Umur Pelanggan': 'Umur Pelanggan', 'Jumlah': 'Jumlah'})
            })
        if 'Segmentasi Pelanggan' in df.columns:
            segmentation_counts = df['Segmentasi Pelanggan'].value_counts().reset_index()
            segmentation_counts.columns = ['Segmentasi Pelanggan', 'Jumlah']
            charts.append({
                'type': 'Analisis demografi pelanggan',
                'figure': px.pie(data_frame=segmentation_counts, 
                                 names='Segmentasi Pelanggan', 
                                 values='Jumlah')
            })

    elif selected_business_info == 'Distribusi usia dan jenis kelamin pelanggan':
        if 'Umur Pelanggan' in df.columns and 'Jenis Kelamin Pelanggan' in df.columns:
            age_gender_counts = df.groupby(['Umur Pelanggan', 'Jenis Kelamin Pelanggan']).size().reset_index(name='Jumlah')
            charts.append({
                'type': 'Distribusi usia dan jenis kelamin pelanggan',
                'figure': px.histogram(data_frame=age_gender_counts, 
                                       x='Umur Pelanggan', 
                                       y='Jumlah', 
                                       color='Jenis Kelamin Pelanggan', 
                                       barmode='group')
            })

    elif selected_business_info == 'Segmentasi pelanggan berdasarkan preferensi':
        if 'Preferensi Pembelian' in df.columns and 'Segmentasi Pelanggan' in df.columns:
            pref_segment_counts = df.groupby(['Preferensi Pembelian', 'Segmentasi Pelanggan']).size().reset_index(name='Jumlah')
            charts.append({
                'type': 'Segmentasi pelanggan berdasarkan preferensi',
                'figure': px.sunburst(data_frame=pref_segment_counts, 
                                      path=['Preferensi Pembelian', 'Segmentasi Pelanggan'], 
                                      values='Jumlah')
            })

    interpretation = interpret_chart('Pelanggan', charts, model)
    return charts, interpretation

def visualize_produk(df, selected_business_info, model):
    df = convert_to_date(df, ['Tanggal'])
    charts = []

    if selected_business_info == 'Kinerja penjualan produk dan stok':
        if 'Produk' in df.columns and 'Jumlah Terjual' in df.columns:
            product_sales = df.groupby('Produk')['Jumlah Terjual'].sum().reset_index()
            charts.append({
                'type': 'Kinerja penjualan produk dan stok',
                'figure': px.bar(data_frame=product_sales, 
                                 x='Produk', 
                                 y='Jumlah Terjual', 
                                 labels={'Produk': 'Produk', 'Jumlah Terjual': 'Jumlah Terjual'})
            })
    elif selected_business_info == 'Distribusi penjualan berdasarkan kategori produk':
        if 'Kategori Produk' in df.columns and 'Jumlah Terjual' in df.columns:
            category_sales = df.groupby('Kategori Produk')['Jumlah Terjual'].sum().reset_index()
            charts.append({
                'type': 'Distribusi penjualan berdasarkan kategori produk',
                'figure': px.pie(data_frame=category_sales, 
                                 names='Kategori Produk', 
                                 values='Jumlah Terjual')
            })

    elif selected_business_info == 'Analisis harga produk dan trend penjualan':
        if 'Tanggal' in df.columns and 'Harga Produk' in df.columns and 'Jumlah Terjual' in df.columns:
            price_trends = df.groupby(['Tanggal', 'Harga Produk'])['Jumlah Terjual'].sum().reset_index()
            charts.append({
                'type': 'Analisis harga produk dan trend penjualan',
                'figure': px.line(data_frame=price_trends, 
                                  x='Tanggal', 
                                  y='Jumlah Terjual', 
                                  color='Harga Produk', 
                                  labels={'Tanggal': 'Tanggal', 'Jumlah Terjual': 'Jumlah Terjual', 'Harga Produk': 'Harga Produk'})
            })

    interpretation = interpret_chart('Produk', charts, model)
    return charts, interpretation

def visualize_transaksi_penjualan(df, selected_business_info, model):
    df = convert_to_date(df, ['Tanggal'])
    charts = []

    if selected_business_info == 'Jumlah penjualan, pendapatan, dan metode pembayaran':
        if 'Metode Pembayaran' in df.columns and 'Pendapatan' in df.columns:
            payment_sales = df.groupby('Metode Pembayaran')['Pendapatan'].sum().reset_index()
            charts.append({
                'type': 'Jumlah penjualan, pendapatan, dan metode pembayaran',
                'figure': px.bar(data_frame=payment_sales, 
                                 x='Metode Pembayaran', 
                                 y='Pendapatan', 
                                 labels={'Metode Pembayaran': 'Metode Pembayaran', 'Pendapatan': 'Pendapatan'})
            })

    elif selected_business_info == 'Tren penjualan':
        if 'Tanggal' in df.columns and 'Pendapatan' in df.columns:
            daily_trends = df.groupby('Tanggal')['Pendapatan'].sum().reset_index()
            charts.append({
                'type': 'Tren penjualan',
                'figure': px.line(data_frame=daily_trends, 
                                  x='Tanggal', 
                                  y='Pendapatan', 
                                  labels={'Tanggal': 'Tanggal', 'Pendapatan': 'Pendapatan'})
            })

    elif selected_business_info == 'Analisis status penjualan dan metode pembayaran':
        if 'Status Penjualan' in df.columns and 'Metode Pembayaran' in df.columns:
            status_payment_sales = df.groupby(['Status Penjualan', 'Metode Pembayaran']).size().reset_index(name='Jumlah')
            charts.append({
                'type': 'Analisis status penjualan dan metode pembayaran',
                'figure': px.bar(data_frame=status_payment_sales, 
                                 x='Status Penjualan', 
                                 y='Jumlah', 
                                 color='Metode Pembayaran', 
                                 barmode='group')
            })

    interpretation = interpret_chart('Transaksi Penjualan', charts, model)
    return charts, interpretation

def visualize_lokasi_penjualan(df, selected_business_info, model):
    df = convert_to_date(df, ['Tanggal'])
    charts = []

    if selected_business_info == 'Kinerja penjualan di berbagai lokasi':
        if 'Lokasi' in df.columns and 'Pendapatan' in df.columns:
            location_sales = df.groupby('Lokasi')['Pendapatan'].sum().reset_index()
            charts.append({
                'type': 'Kinerja penjualan di berbagai lokasi',
                'figure': px.bar(data_frame=location_sales, 
                                 x='Lokasi', 
                                 y='Pendapatan', 
                                 labels={'Lokasi': 'Lokasi', 'Pendapatan': 'Pendapatan'})
            })

    elif selected_business_info == 'Distribusi penjualan berdasarkan kota/provinsi':
        if 'Kota/Provinsi' in df.columns and 'Pendapatan' in df.columns:
            city_province_sales = df.groupby('Kota/Provinsi')['Pendapatan'].sum().reset_index()
            charts.append({
                'type': 'Distribusi penjualan berdasarkan kota/provinsi',
                'figure': px.pie(data_frame=city_province_sales, 
                                 names='Kota/Provinsi', 
                                 values='Pendapatan')
            })

    elif selected_business_info == 'Analisis lokasi dengan penjualan tertinggi/rendah':
        if 'Lokasi' in df.columns and 'Pendapatan' in df.columns:
            location_sales = df.groupby('Lokasi')['Pendapatan'].sum().reset_index()
            top_location_sales = location_sales.sort_values('Pendapatan', ascending=False).head(10)
            charts.append({
                'type': 'Analisis lokasi dengan penjualan tertinggi/rendah',
                'figure': px.bar(data_frame=top_location_sales, 
                                 x='Lokasi', 
                                 y='Pendapatan', 
                                 labels={'Lokasi': 'Lokasi', 'Pendapatan': 'Pendapatan'})
            })

    interpretation = interpret_chart('Lokasi Penjualan', charts, model)
    return charts, interpretation

def visualize_staf_penjualan(df, selected_business_info, model):
    df = convert_to_date(df, ['Tanggal'])
    charts = []

    if selected_business_info == 'Kinerja dan komisi staf penjualan':
        if 'Nama Staf' in df.columns and 'Komisi' in df.columns:
            staff_commission = df.groupby('Nama Staf')['Komisi'].sum().reset_index()
            charts.append({
                'type': 'Kinerja dan komisi staf penjualan',
                'figure': px.bar(data_frame=staff_commission, 
                                 x='Nama Staf', 
                                 y='Komisi', 
                                 labels={'Nama Staf': 'Nama Staf', 'Komisi': 'Komisi'})
            })

    elif selected_business_info == 'Analisis penilaian kinerja staf':
        if 'Nama Staf' in df.columns and 'Penilaian Kinerja' in df.columns:
            staff_performance = df.groupby('Nama Staf')['Penilaian Kinerja'].mean().reset_index()
            charts.append({
                'type': 'Analisis penilaian kinerja staf',
                'figure': px.bar(data_frame=staff_performance, 
                                 x='Nama Staf', 
                                 y='Penilaian Kinerja', 
                                 labels={'Nama Staf': 'Nama Staf', 'Penilaian Kinerja': 'Penilaian Kinerja'})
            })

    elif selected_business_info == 'Distribusi staf berdasarkan posisi/jabatan':
        if 'Posisi/Jabatan' in df.columns:
            position_counts = df['Posisi/Jabatan'].value_counts().reset_index()
            position_counts.columns = ['Posisi/Jabatan', 'Jumlah']
            charts.append({
                'type': 'Distribusi staf berdasarkan posisi/jabatan',
                'figure': px.pie(data_frame=position_counts, 
                                 names='Posisi/Jabatan', 
                                 values='Jumlah')
            })

    interpretation = interpret_chart('Staf Penjualan', charts, model)
    return charts, interpretation

def visualize_inventaris(df, selected_business_info, model):
    df = convert_to_date(df, ['Tanggal'])
    charts = []

    if selected_business_info == 'Manajemen stok produk':
        if 'Produk' in df.columns and 'Stok' in df.columns:
            product_stock = df.groupby('Produk')['Stok'].sum().reset_index()
            charts.append({
                'type': 'Manajemen stok produk',
                'figure': px.bar(data_frame=product_stock, 
                                 x='Produk', 
                                 y='Stok', 
                                 labels={'Produk': 'Produk', 'Stok': 'Stok'})
            })

    elif selected_business_info == 'Tren stok masuk dan keluar':
        if 'Tanggal' in df.columns and 'Stok Masuk' in df.columns and 'Stok Keluar' in df.columns:
            stock_trends = df.groupby('Tanggal').agg({'Stok Masuk': 'sum', 'Stok Keluar': 'sum'}).reset_index()
            stock_trends = stock_trends.melt(id_vars='Tanggal', value_vars=['Stok Masuk', 'Stok Keluar'], var_name='Tipe', value_name='Jumlah')
            charts.append({
                'type': 'Tren stok masuk dan keluar',
                'figure': px.line(data_frame=stock_trends, 
                                  x='Tanggal', 
                                  y='Jumlah', 
                                  color='Tipe', 
                                  labels={'Tanggal': 'Tanggal', 'Jumlah': 'Jumlah', 'Tipe': 'Tipe'})
            })

    elif selected_business_info == 'Analisis produk dengan stok terbanyak/terkecil':
        if 'Produk' in df.columns and 'Stok' in df.columns:
            top_stock_products = df.groupby('Produk')['Stok'].sum().reset_index().sort_values('Stok', ascending=False).head(10)
            charts.append({
                'type': 'Analisis produk dengan stok terbanyak/terkecil',
                'figure': px.bar(data_frame=top_stock_products, 
                                 x='Produk', 
                                 y='Stok', 
                                 labels={'Produk': 'Produk', 'Stok': 'Stok'})
            })

    interpretation = interpret_chart('Inventaris', charts, model)
    return charts, interpretation

def visualize_promosi_pemasaran(df, selected_business_info, model):
    df = convert_to_date(df, ['Tanggal'])
    charts = []

    if selected_business_info == 'Efektivitas kampanye promosi':
        if 'Kampanye Promosi' in df.columns and 'Pendapatan' in df.columns:
            campaign_sales = df.groupby('Kampanye Promosi')['Pendapatan'].sum().reset_index()
            charts.append({
                'type': 'Efektivitas kampanye promosi',
                'figure': px.bar(data_frame=campaign_sales, 
                                 x='Kampanye Promosi', 
                                 y='Pendapatan', 
                                 labels={'Kampanye Promosi': 'Kampanye Promosi', 'Pendapatan': 'Pendapatan'})
            })

    elif selected_business_info == 'Distribusi penjualan berdasarkan media promosi':
        if 'Media Promosi' in df.columns and 'Pendapatan' in df.columns:
            media_sales = df.groupby('Media Promosi')['Pendapatan'].sum().reset_index()
            charts.append({
                'type': 'Distribusi penjualan berdasarkan media promosi',
                'figure': px.pie(data_frame=media_sales, 
                                 names='Media Promosi', 
                                 values='Pendapatan')
            })

    elif selected_business_info == 'Analisis kode diskon promosi':
        if 'Kode Diskon' in df.columns and 'Pendapatan' in df.columns:
            discount_code_sales = df.groupby('Kode Diskon')['Pendapatan'].sum().reset_index()
            charts.append({
                'type': 'Analisis kode diskon promosi',
                'figure': px.bar(data_frame=discount_code_sales, 
                                 x='Kode Diskon', 
                                 y='Pendapatan', 
                                 labels={'Kode Diskon': 'Kode Diskon', 'Pendapatan': 'Pendapatan'})
            })

    interpretation = interpret_chart('Promosi dan Pemasaran', charts, model)
    return charts, interpretation

def visualize_feedback_pengembalian(df, selected_business_info, model):
    df = convert_to_date(df, ['Tanggal'])
    charts = []

    if selected_business_info == 'Masalah dan kepuasan pelanggan':
        if 'Masalah Pelanggan' in df.columns:
            issue_counts = df['Masalah Pelanggan'].value_counts().reset_index()
            issue_counts.columns = ['Masalah Pelanggan', 'Jumlah']
            charts.append({
                'type': 'Masalah dan kepuasan pelanggan',
                'figure': px.bar(data_frame=issue_counts, 
                                 x='Masalah Pelanggan', 
                                 y='Jumlah', 
                                 labels={'Masalah Pelanggan': 'Masalah Pelanggan', 'Jumlah': 'Jumlah'})
            })

    elif selected_business_info == 'Distribusi alasan pengembalian produk':
        if 'Alasan Pengembalian' in df.columns:
            return_reason_counts = df['Alasan Pengembalian'].value_counts().reset_index()
            return_reason_counts.columns = ['Alasan Pengembalian', 'Jumlah']
            charts.append({
                'type': 'Distribusi alasan pengembalian produk',
                'figure': px.pie(data_frame=return_reason_counts, 
                                 names='Alasan Pengembalian', 
                                 values='Jumlah')
            })

    elif selected_business_info == 'Status pengembalian produk':
        if 'Status Pengembalian' in df.columns:
            return_status_counts = df['Status Pengembalian'].value_counts().reset_index()
            return_status_counts.columns = ['Status Pengembalian', 'Jumlah']
            charts.append({
                'type': 'Status pengembalian produk',
                'figure': px.bar(data_frame=return_status_counts, 
                                 x='Status Pengembalian', 
                                 y='Jumlah', 
                                 labels={'Status Pengembalian': 'Status Pengembalian', 'Jumlah': 'Jumlah'})
            })

    interpretation = interpret_chart('Feedback dan Pengembalian', charts, model)
    return charts, interpretation

def visualize_analisis_penjualan(df, selected_business_info, model):
    df = convert_to_date(df, ['Tanggal'])
    charts = []

    if selected_business_info == 'Penjualan agregat dan tren':
        if 'Tanggal' in df.columns and 'Pendapatan' in df.columns:
            sales_trends = df.groupby('Tanggal')['Pendapatan'].sum().reset_index()
            charts.append({
                'type': 'Penjualan agregat dan tren',
                'figure': px.line(data_frame=sales_trends, 
                                  x='Tanggal', 
                                  y='Pendapatan', 
                                  labels={'Tanggal': 'Tanggal', 'Pendapatan': 'Pendapatan'})
            })

    elif selected_business_info == 'Analisis penjualan berdasarkan produk/kategori':
        if 'Produk' in df.columns and 'Pendapatan' in df.columns:
            product_sales = df.groupby('Produk')['Pendapatan'].sum().reset_index()
            charts.append({
                'type': 'Analisis penjualan berdasarkan produk/kategori',
                'figure': px.bar(data_frame=product_sales, 
                                 x='Produk', 
                                 y='Pendapatan', 
                                 labels={'Produk': 'Produk', 'Pendapatan': 'Pendapatan'})
            })

    elif selected_business_info == 'Tren penjualan/tahunan':
        if 'Tanggal' in df.columns and 'Pendapatan' in df.columns:
            df['Bulan'] = df['Tanggal'].dt.to_period('M')
            df['Tahun'] = df['Tanggal'].dt.year
            monthly_trends = df.groupby('Bulan')['Pendapatan'].sum().reset_index()
            yearly_trends = df.groupby('Tahun')['Pendapatan'].sum().reset_index()
            charts.append({
                'type': 'Tren penjualan/tahunan',
                'figure': px.line(data_frame=monthly_trends, 
                                  x='Bulan', 
                                  y='Pendapatan', 
                                  labels={'Bulan': 'Bulan', 'Pendapatan': 'Pendapatan'})
            })
            charts.append({
                'type': 'Tren penjualan/tahunan',
                'figure': px.line(data_frame=yearly_trends, 
                                  x='Tahun', 
                                  y='Pendapatan', 
                                  labels={'Tahun': 'Tahun', 'Pendapatan': 'Pendapatan'})
            })

    interpretation = interpret_chart('Analisis Penjualan', charts, model)
    return charts, interpretation

def visualize_lainnya(df, selected_business_info, model):
    df = convert_to_date(df, ['Tanggal'])
    charts = []

    if selected_business_info == 'Analisis tambahan dan faktor eksternal':
        if 'Faktor Eksternal' in df.columns and 'Pendapatan' in df.columns:
            external_factor_sales = df.groupby('Faktor Eksternal')['Pendapatan'].sum().reset_index()
            charts.append({
                'type': 'Analisis tambahan dan faktor eksternal',
                'figure': px.bar(data_frame=external_factor_sales, 
                                 x='Faktor Eksternal', 
                                 y='Pendapatan', 
                                 labels={'Faktor Eksternal': 'Faktor Eksternal', 'Pendapatan': 'Pendapatan'})
            })

    elif selected_business_info == 'Tren penjualan berdasarkan faktor ekonomi':
        if 'Tanggal' in df.columns and 'Faktor Ekonomi' in df.columns and 'Pendapatan' in df.columns:
            economic_factor_trends = df.groupby(['Tanggal', 'Faktor Ekonomi'])['Pendapatan'].sum().reset_index()
            charts.append({
                'type': 'Tren penjualan berdasarkan faktor ekonomi',
                'figure': px.line(data_frame=economic_factor_trends, 
                                  x='Tanggal', 
                                  y='Pendapatan', 
                                  color='Faktor Ekonomi', 
                                  labels={'Tanggal': 'Tanggal', 'Pendapatan': 'Pendapatan', 'Faktor Ekonomi': 'Faktor Ekonomi'})
            })

    elif selected_business_info == 'Analisis faktor lingkungan dan penjualan':
        if 'Faktor Lingkungan' in df.columns and 'Pendapatan' in df.columns:
            environmental_factor_sales = df.groupby('Faktor Lingkungan')['Pendapatan'].sum().reset_index()
            charts.append({
                'type': 'Analisis faktor lingkungan dan penjualan',
                'figure': px.bar(data_frame=environmental_factor_sales, 
                                 x='Faktor Lingkungan', 
                                 y='Pendapatan', 
                                 labels={'Faktor Lingkungan': 'Faktor Lingkungan', 'Pendapatan': 'Pendapatan'})
            })

    interpretation = interpret_chart('Lainnya', charts, model)
    return charts, interpretation
