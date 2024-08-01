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

def interpret_chart(sheet_name, charts, model, existing_interpretations):
    interpretations = []
    
    for chart in charts:
        if chart['type'] in existing_interpretations:
            chart_description = existing_interpretations[chart['type']]
        else:
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
            existing_interpretations[chart['type']] = chart_description

        interpretations.append(chart_description)
    
    return "\n\n".join(interpretations)

def convert_to_date(df, columns):
    for col in columns:
        if (col in df.columns) and (df[col].dtype != 'datetime64[ns]'):
            df[col] = pd.to_datetime(df[col], errors='coerce')
    return df

def visualize_pelanggan(df, selected_business_info, model, existing_interpretations):
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

    interpretation = interpret_chart('Pelanggan', charts, model, existing_interpretations)
    return charts, interpretation

# Fungsi visualisasi lainnya tetap sama, tambahkan parameter `existing_interpretations` dan panggil `interpret_chart` dengan parameter tersebut
def visualize_produk(df, selected_business_info, model, existing_interpretations):
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

    interpretation = interpret_chart('Produk', charts, model, existing_interpretations)
    return charts, interpretation

def visualize_transaksi_penjualan(df, selected_business_info, model, existing_interpretations):
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

    interpretation = interpret_chart('Transaksi Penjualan', charts, model, existing_interpretations)
    return charts, interpretation

def visualize_transaksi_pembelian(df, selected_business_info, model, existing_interpretations):
    df = convert_to_date(df, ['Tanggal'])
    charts = []

    if selected_business_info == 'Jumlah pembelian, pengeluaran, dan metode pembayaran':
        if 'Metode Pembayaran' in df.columns and 'Pengeluaran' in df.columns:
            payment_purchases = df.groupby('Metode Pembayaran')['Pengeluaran'].sum().reset_index()
            charts.append({
                'type': 'Jumlah pembelian, pengeluaran, dan metode pembayaran',
                'figure': px.bar(data_frame=payment_purchases, 
                                 x='Metode Pembayaran', 
                                 y='Pengeluaran', 
                                 labels={'Metode Pembayaran': 'Metode Pembayaran', 'Pengeluaran': 'Pengeluaran'})
            })

    elif selected_business_info == 'Tren pembelian':
        if 'Tanggal' in df.columns and 'Pengeluaran' in df.columns:
            daily_purchase_trends = df.groupby('Tanggal')['Pengeluaran'].sum().reset_index()
            charts.append({
                'type': 'Tren pembelian',
                'figure': px.line(data_frame=daily_purchase_trends, 
                                  x='Tanggal', 
                                  y='Pengeluaran', 
                                  labels={'Tanggal': 'Tanggal', 'Pengeluaran': 'Pengeluaran'})
            })

    elif selected_business_info == 'Analisis status pembelian dan metode pembayaran':
        if 'Status Pembelian' in df.columns and 'Metode Pembayaran' in df.columns:
            status_payment_purchases = df.groupby(['Status Pembelian', 'Metode Pembayaran']).size().reset_index(name='Jumlah')
            charts.append({
                'type': 'Analisis status pembelian dan metode pembayaran',
                'figure': px.bar(data_frame=status_payment_purchases, 
                                 x='Status Pembelian', 
                                 y='Jumlah', 
                                 color='Metode Pembayaran', 
                                 barmode='group')
            })

    interpretation = interpret_chart('Transaksi Pembelian', charts, model, existing_interpretations)
    return charts, interpretation

def visualize_transaksi_operasional(df, selected_business_info, model, existing_interpretations):
    df = convert_to_date(df, ['Tanggal'])
    charts = []

    if selected_business_info == 'Jumlah pengeluaran operasional':
        if 'Pengeluaran Operasional' in df.columns and 'Jumlah' in df.columns:
            operational_expenses = df.groupby('Pengeluaran Operasional')['Jumlah'].sum().reset_index()
            charts.append({
                'type': 'Jumlah pengeluaran operasional',
                'figure': px.bar(data_frame=operational_expenses, 
                                 x='Pengeluaran Operasional', 
                                 y='Jumlah', 
                                 labels={'Pengeluaran Operasional': 'Pengeluaran Operasional', 'Jumlah': 'Jumlah'})
            })

    elif selected_business_info == 'Tren pengeluaran operasional':
        if 'Tanggal' in df.columns and 'Jumlah' in df.columns:
            daily_operational_trends = df.groupby('Tanggal')['Jumlah'].sum().reset_index()
            charts.append({
                'type': 'Tren pengeluaran operasional',
                'figure': px.line(data_frame=daily_operational_trends, 
                                  x='Tanggal', 
                                  y='Jumlah', 
                                  labels={'Tanggal': 'Tanggal', 'Jumlah': 'Jumlah'})
            })

    interpretation = interpret_chart('Transaksi Operasional', charts, model, existing_interpretations)
    return charts, interpretation

def visualize_karyawan(df, selected_business_info, model, existing_interpretations):
    df = convert_to_date(df, ['Tanggal'])
    charts = []

    if selected_business_info == 'Analisis kinerja karyawan berdasarkan penjualan':
        if 'Karyawan' in df.columns and 'Jumlah Penjualan' in df.columns:
            employee_sales = df.groupby('Karyawan')['Jumlah Penjualan'].sum().reset_index()
            charts.append({
                'type': 'Analisis kinerja karyawan berdasarkan penjualan',
                'figure': px.bar(data_frame=employee_sales, 
                                 x='Karyawan', 
                                 y='Jumlah Penjualan', 
                                 labels={'Karyawan': 'Karyawan', 'Jumlah Penjualan': 'Jumlah Penjualan'})
            })

    elif selected_business_info == 'Analisis absensi karyawan':
        if 'Karyawan' in df.columns and 'Jumlah Hari Hadir' in df.columns:
            employee_attendance = df.groupby('Karyawan')['Jumlah Hari Hadir'].sum().reset_index()
            charts.append({
                'type': 'Analisis absensi karyawan',
                'figure': px.bar(data_frame=employee_attendance, 
                                 x='Karyawan', 
                                 y='Jumlah Hari Hadir', 
                                 labels={'Karyawan': 'Karyawan', 'Jumlah Hari Hadir': 'Jumlah Hari Hadir'})
            })

    elif selected_business_info == 'Analisis produktivitas karyawan':
        if 'Karyawan' in df.columns and 'Jumlah Penjualan' in df.columns and 'Jumlah Hari Hadir' in df.columns:
            employee_productivity = df.groupby('Karyawan').apply(lambda x: x['Jumlah Penjualan'].sum() / x['Jumlah Hari Hadir'].sum()).reset_index(name='Produktivitas')
            charts.append({
                'type': 'Analisis produktivitas karyawan',
                'figure': px.bar(data_frame=employee_productivity, 
                                 x='Karyawan', 
                                 y='Produktivitas', 
                                 labels={'Karyawan': 'Karyawan', 'Produktivitas': 'Produktivitas'})
            })

    interpretation = interpret_chart('Karyawan', charts, model, existing_interpretations)
    return charts, interpretation
