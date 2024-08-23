import streamlit as st
import pandas as pd
import re
from reportlab.lib.pagesizes import letter
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, Image, PageBreak
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import inch
import matplotlib.pyplot as plt
from datetime import datetime

# Variables globales para mantener estado
users = {}
attempts = {}
current_user = None
files = {"nomina": None, "asistencia": None, "productividad": None}

def main():
    global current_user

    st.title("Papeles de trabajo Auditoría en sistemas Chalen, Luo y Palau")

    # Pantalla de login
    if 'authenticated' not in st.session_state or not st.session_state.authenticated:
        login()
    else:
        main_app()

def login():
    global current_user

    st.subheader("Iniciar Sesión")
    username = st.text_input("Usuario:")
    password = st.text_input("Contraseña:", type="password")

    if st.button("Iniciar Sesión"):
        if username in users and users[username] == password:
            current_user = username
            st.session_state.authenticated = True
            st.success("Inicio de sesión exitoso")
        else:
            attempts[username] = attempts.get(username, 0) + 1
            if attempts[username] >= 3:
                st.error("Usuario bloqueado. Demasiados intentos fallidos.")
            else:
                st.error("Usuario o contraseña incorrectos")

    st.write("¿No tienes cuenta?")
    if st.button("Crear Usuario"):
        create_user_interface()

def create_user_interface():
    st.subheader("Crear Usuario")
    new_username = st.text_input("Nuevo Usuario:")
    new_password = st.text_input("Nueva Contraseña:", type="password")

    if st.button("Crear Usuario"):
        if not new_username or not new_password:
            st.error("Usuario y contraseña no pueden estar vacíos.")
        elif len(new_password) < 8 or not re.search(r'[A-Z]', new_password) or not re.search(r'[0-9]', new_password):
            st.error("La contraseña debe tener al menos 8 caracteres, incluir mayúsculas y números.")
        else:
            users[new_username] = new_password
            attempts[new_username] = 0
            st.success("Usuario creado exitosamente.")

def main_app():
    st.sidebar.title(f"Bienvenido {current_user}")

    option = st.sidebar.selectbox("Selecciona una opción", ["Carga de Archivos", "Generación de Análisis", "Papel de Trabajo", "Cerrar Sesión"])

    if option == "Carga de Archivos":
        upload_files()
    elif option == "Generación de Análisis":
        generate_analysis()
    elif option == "Papel de Trabajo":
        generate_work_paper()
    elif option == "Cerrar Sesión":
        st.session_state.authenticated = False
        st.experimental_rerun()

def upload_files():
    st.subheader("Carga de Archivos")
    files['nomina'] = st.file_uploader("Carga archivo de nómina", type="xlsx")
    files['asistencia'] = st.file_uploader("Carga archivo de asistencia", type="xlsx")
    files['productividad'] = st.file_uploader("Carga archivo de productividad", type="xlsx")

    if files['nomina']:
        st.success("Archivo de nómina cargado correctamente.")
    if files['asistencia']:
        st.success("Archivo de asistencia cargado correctamente.")
    if files['productividad']:
        st.success("Archivo de productividad cargado correctamente.")

def generate_analysis():
    st.subheader("Generación de Análisis")
    if st.button("Análisis de Nómina"):
        analyze_nomina()
    if st.button("Análisis de Asistencia"):
        analyze_asistencia()
    if st.button("Análisis de Productividad"):
        analyze_productividad()

def generate_work_paper():
    st.subheader("Papel de Trabajo")
    if st.button("Generar Papel de Trabajo de Nómina"):
        generate_papel_trabajo_nomina()
    if st.button("Generar Papel de Trabajo de Asistencia"):
        generate_papel_trabajo_asistencia()
    if st.button("Generar Papel de Trabajo de Productividad"):
        generate_papel_trabajo_productividad()

def analyze_nomina():
    if not files['nomina']:
        st.error("No se ha cargado ningún archivo de nómina.")
        return

    df_nomina = pd.read_excel(files['nomina'])
    duplicados_nombre = df_nomina[df_nomina.duplicated(subset='Nombre', keep=False)]
    duplicados_cuenta = df_nomina[df_nomina.duplicated(subset='Cuenta Bancaria', keep=False)]

    total_rows = len(df_nomina)
    anomalías = len(duplicados_nombre) + len(duplicados_cuenta)
    porcentaje_anomalías = (anomalías / total_rows) * 100 if total_rows > 0 else 0

    st.write(f"Anomalías identificadas:\n\nNúmero de Empleados duplicados: {len(duplicados_nombre)}\nNúmero de cuenta bancaria Empleados duplicadas: {len(duplicados_cuenta)}")

    st.session_state.nomina_data = {
        'df_nomina': df_nomina,
        'duplicados_nombre': duplicados_nombre,
        'duplicados_cuenta': duplicados_cuenta,
        'total_rows': total_rows,
        'porcentaje_anomalías': porcentaje_anomalías
    }

    st.success("Análisis de nómina completado.")
    create_nomina_pdf_report("nomina_report.pdf")

def create_nomina_pdf_report(pdf_path):
    doc = SimpleDocTemplate(pdf_path, pagesize=letter)
    elements = []

    styles = getSampleStyleSheet()
    styles.add(ParagraphStyle(name='CustomTitle', fontName='Helvetica', fontSize=14, leading=24, spaceAfter=12))
    styles.add(ParagraphStyle(name='CustomText', fontName='Helvetica', fontSize=12, leading=24))

    title = Paragraph("Análisis de Nómina - Anomalías Identificadas", styles['CustomTitle'])
    elements.append(title)

    auditor_paragraph = Paragraph(f"Auditor/es: {current_user}", styles['CustomText'])
    elements.append(auditor_paragraph)

    summary = Paragraph(f"<br/>Número de Empleados duplicados: {len(st.session_state.nomina_data['duplicados_nombre'])}<br/>"
                        f"Número de cuentas bancarias duplicadas: {len(st.session_state.nomina_data['duplicados_cuenta'])}<br/><br/>", 
                        styles['CustomText'])
    elements.append(summary)

    elements.append(Paragraph("Listado de Empleados Duplicados:", styles['CustomTitle']))

    filtered_data_1 = st.session_state.nomina_data['duplicados_nombre'][['ID de Empleado', 'Nombre']].sort_values(by='Nombre')
    data_table_1 = [filtered_data_1.columns.tolist()] + filtered_data_1.values.tolist()

    table_1 = Table(data_table_1, colWidths=[doc.width / len(data_table_1[0])] * len(data_table_1[0]), repeatRows=1)
    
    table_1.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.grey),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
        ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
        ('BACKGROUND', (0, 1), (-1, -1), colors.beige),
        ('GRID', (0, 0), (-1, -1), 1, colors.black),
    ]))
    elements.append(table_1)

    elements.append(Spacer(1, 12))
    elements.append(Paragraph("Listado de Cuentas Bancarias Duplicadas:", styles['CustomTitle']))

    filtered_data_2 = st.session_state.nomina_data['duplicados_cuenta'][['ID de Empleado', 'Nombre', 'Cuenta Bancaria']].sort_values(by='Nombre')
    data_table_2 = [filtered_data_2.columns.tolist()] + filtered_data_2.values.tolist()

    table_2 = Table(data_table_2, colWidths=[doc.width / len(data_table_2[0])] * len(data_table_2[0]), repeatRows=1)
    
    table_2.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.grey),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
        ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
        ('BACKGROUND', (0, 1), (-1, -1), colors.beige),
        ('GRID', (0, 0), (-1, -1), 1, colors.black),
    ]))
    elements.append(table_2)

    elements.append(Spacer(1, 12))
    create_pie_chart(len(st.session_state.nomina_data['duplicados_nombre']), st.session_state.nomina_data['total_rows'], "nombres_duplicados.png", "Porcentaje de Nombres Duplicados")
    create_pie_chart(len(st.session_state.nomina_data['duplicados_cuenta']), st.session_state.nomina_data['total_rows'], "cuentas_duplicadas.png", "Porcentaje de Cuentas Bancarias Duplicadas")

    elements.append(Spacer(1, 12))
    elements.append(Paragraph("Gráfico de Nombres Duplicados:", styles['CustomTitle']))
    elements.append(Image("nombres_duplicados.png", 4 * inch, 4 * inch))

    elements.append(Spacer(1, 12))
    elements.append(Paragraph("Gráfico de Cuentas Bancarias Duplicadas:", styles['CustomTitle']))
    elements.append(Image("cuentas_duplicadas.png", 4 * inch, 4 * inch))

    doc.build(elements)
    st.success(f"Reporte PDF generado exitosamente en {pdf_path}")

def create_pie_chart(count, total, filename, title):
    if total == 0:
        return
    labels = ['Duplicados', 'Únicos']
    sizes = [count, total - count]
    colors = ['#ff9999','#66b3ff']
    fig1, ax1 = plt.subplots()
    ax1.pie(sizes, labels=labels, autopct='%1.1f%%', colors=colors, startangle=90)
    ax1.axis('equal')
    plt.title(title)
    plt.savefig(filename)
    plt.close()

def analyze_asistencia():
    if not files['asistencia']:
        st.error("No se ha cargado ningún archivo de asistencia.")
        return

    try:
        df_asistencia = pd.read_excel(files['asistencia'])
        meses = ['Julio', 'Agosto', 'Septiembre', 'Octubre', 'Noviembre', 'Diciembre']
        anomalías_mensuales = {}
        all_anomalías = pd.DataFrame()

        for mes in meses:
            columna = f'Días Trabajados en {mes}'
            if columna in df_asistencia.columns:
                anomalías_mensuales[mes] = df_asistencia[df_asistencia[columna] <= 20]
                all_anomalías = pd.concat([all_anomalías, anomalías_mensuales[mes]])
            else:
                st.error(f"No se encuentra la columna para {mes} en el archivo.")
                return

        all_anomalías = all_anomalías.drop_duplicates()
        total_anomalías = df_asistencia[df_asistencia['Total Días Trabajados'] < 120]
        all_anomalías = pd.concat([all_anomalías, total_anomalías]).drop_duplicates()

        st.write(f"Análisis de Asistencia:\n\nNúmero de empleados con anomalías en días trabajados por mes: {len(all_anomalías)}\n")

        st.session_state.asistencia_data = {
            'df_asistencia': df_asistencia,
            'anomalías_mensuales': anomalías_mensuales,
            'total_anomalías': total_anomalías,
            'all_anomalías': all_anomalías
        }

        st.success("Análisis de asistencia completado.")
        create_asistencia_pdf_report("asistencia_report.pdf")
    except Exception as e:
        st.error(f"Se produjo un error durante el análisis de asistencia: {str(e)}")

def create_asistencia_pdf_report(pdf_path):
    doc = SimpleDocTemplate(pdf_path, pagesize=letter)
    elements = []
    styles = getSampleStyleSheet()

    title = Paragraph("Análisis de Asistencia - Anomalías Identificadas", styles['Title'])
    elements.append(title)

    auditor_paragraph = Paragraph(f"Auditor/es: {current_user}", styles['Normal'])
    elements.append(auditor_paragraph)

    summary = Paragraph(f"<br/>Número de empleados con menos de 20 días trabajados en al menos un mes: {len(st.session_state.asistencia_data['all_anomalías'])}<br/>"
                        f"Empleados con menos de 120 días trabajados en total: {len(st.session_state.asistencia_data['total_anomalías'])}<br/><br/>", styles['Normal'])
    elements.append(summary)

    elements.append(Paragraph("Datos de Asistencia:", styles['Heading2']))
    data_table = [st.session_state.asistencia_data['df_asistencia'].columns.tolist()] + st.session_state.asistencia_data['df_asistencia'].values.tolist()
    
    table = Table(data_table, colWidths=[doc.width / len(data_table[0])] * len(data_table[0]), repeatRows=1)
    
    table.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.grey),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
        ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
        ('BACKGROUND', (0, 1), (-1, -1), colors.beige),
        ('GRID', (0, 0), (-1, -1), 1, colors.black),
    ]))
    elements.append(table)

    elements.append(Spacer(1, 12))
    create_pie_chart(len(st.session_state.asistencia_data['all_anomalías']), len(st.session_state.asistencia_data['df_asistencia']), "dias_trabajados.png", "Porcentaje de Anomalías en Días Trabajados")

    elements.append(Spacer(1, 12))
    elements.append(Paragraph("Gráfico de Días Trabajados:", styles['Heading2']))
    elements.append(Image("dias_trabajados.png", 4 * inch, 4 * inch))

    doc.build(elements)
    st.success(f"Reporte PDF generado exitosamente en {pdf_path}")

def analyze_productividad():
    if not files['productividad']:
        st.error("No se ha cargado ningún archivo de productividad.")
        return

    df_productividad = pd.read_excel(files['productividad'])
    anomalías_mensuales = {}
    anomalías_totales = []

    for mes in ['Julio', 'Agosto', 'Septiembre', 'Octubre', 'Noviembre', 'Diciembre']:
        anomalías_mensuales[mes] = df_productividad[df_productividad[f'Tareas Realizadas en {mes}'] < 17]

    total_anomalías = df_productividad[df_productividad['Productividad (Tareas - 6 meses)'] < 102]

    for mes in anomalías_mensuales:
        anomalías_totales.append(anomalías_mensuales[mes])

    anomalías_totales.append(total_anomalías)
    all_anomalías = pd.concat(anomalías_totales).drop_duplicates()

    st.write(f"Análisis de Productividad:\n\nNúmero de empleados con anomalías en tareas realizadas por mes: {len(all_anomalías)}\n")

    st.session_state.productividad_data = {
        'df_productividad': df_productividad,
        'anomalías_mensuales': anomalías_mensuales,
        'total_anomalías': total_anomalías,
        'all_anomalías': all_anomalías
    }

    st.success("Análisis de productividad completado.")
    create_productividad_pdf_report("productividad_report.pdf")

def create_productividad_pdf_report(pdf_path):
    doc = SimpleDocTemplate(pdf_path, pagesize=letter)
    elements = []
    styles = getSampleStyleSheet()

    title = Paragraph("Análisis de Productividad - Anomalías Identificadas", styles['Title'])
    elements.append(title)

    auditor_paragraph = Paragraph(f"Auditor/es: {current_user}", styles['Normal'])
    elements.append(auditor_paragraph)

    summary = Paragraph(f"<br/>Número de empleados con menos de 17 tareas realizadas en al menos un mes: {len(st.session_state.productividad_data['all_anomalías'])}<br/>"
                        f"Empleados con menos de 102 tareas realizadas en total: {len(st.session_state.productividad_data['total_anomalías'])}<br/><br/>", styles['Normal'])
    elements.append(summary)

    elements.append(Paragraph("Datos de Productividad:", styles['Heading2']))
    data_table = [st.session_state.productividad_data['df_productividad'].columns.tolist()] + st.session_state.productividad_data['df_productividad'].values.tolist()
    
    table = Table(data_table, colWidths=[doc.width / len(data_table[0])] * len(data_table[0]), repeatRows=1)
    
    table.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.grey),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
        ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
        ('BACKGROUND', (0, 1), (-1, -1), colors.beige),
        ('GRID', (0, 0), (-1, -1), 1, colors.black),
    ]))
    elements.append(table)

    elements.append(Spacer(1, 12))
    create_pie_chart(len(st.session_state.productividad_data['all_anomalías']), len(st.session_state.productividad_data['df_productividad']), "tareas_realizadas.png", "Porcentaje de Anomalías en Tareas Realizadas")

    elements.append(Spacer(1, 12))
    elements.append(Paragraph("Gráfico de Tareas Realizadas:", styles['Heading2']))
    elements.append(Image("tareas_realizadas.png", 4 * inch, 4 * inch))

    doc.build(elements)
    st.success(f"Reporte PDF generado exitosamente en {pdf_path}")

if __name__ == "__main__":
    main()
