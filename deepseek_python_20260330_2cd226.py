"""
Dashboard Financiero Completo
Este script genera un dashboard interactivo con datos de mercados financieros
y lo guarda como un archivo HTML autónomo listo para GitHub Pages.
"""

import pandas as pd
import plotly.graph_objects as go
from plotly.subplots import make_subplots
import numpy as np
import os
import json
from datetime import datetime

# ============================================================================
# 1. CARGA Y LIMPIEZA DE DATOS
# ============================================================================

print("=" * 80)
print("📊 GENERANDO DASHBOARD FINANCIERO")
print("=" * 80)
print("\n1. Cargando datos desde Excel...")

# Leer el archivo Excel
try:
    df = pd.read_excel('Mercury 1.0.xlsx', sheet_name='Sheet1')
    print("   ✓ Archivo Excel cargado correctamente")
except FileNotFoundError:
    print("   ✗ Error: No se encuentra el archivo 'Mercury 1.0.xlsx'")
    print("   Asegúrate de que el archivo está en la misma carpeta que este script")
    exit(1)

# Convertir fecha
df['Date'] = pd.to_datetime(df['Date'], errors='coerce')
df = df.dropna(subset=['Date'])
df = df.sort_values('Date').reset_index(drop=True)

# Convertir columnas numéricas
numeric_columns = ['S&P 500', 'Nasdaq 100', 'Dow Jones', 'MSCI world', 
                   'US 10YT', 'US2YT (%)', '10-2 years spread', 
                   'VIX riesgo de los mercados', 'DXY index', 'USD/EUR (', 
                   'Petroleo WTI $/Bar', 'Oro $/oz', 'Bs/Usd ( tasa oficial)', 
                   'SIlver USD/Oz']

for col in numeric_columns:
    if col in df.columns:
        df[col] = pd.to_numeric(df[col], errors='coerce')

# Eliminar filas vacías
df = df.dropna(subset=numeric_columns, how='all')

print(f"   ✓ Datos cargados: {len(df)} registros del {df['Date'].min().strftime('%Y-%m-%d')} "
      f"al {df['Date'].max().strftime('%Y-%m-%d')}")

# Renombrar columnas
columns = {
    'Date': 'Date',
    'S&P 500': 'S&P 500',
    'Nasdaq 100': 'Nasdaq 100',
    'Dow Jones': 'Dow Jones',
    'MSCI world': 'MSCI World',
    'US 10YT': 'US 10YT',
    'US2YT (%)': 'US 2YT',
    '10-2 years spread': '10-2 Spread',
    'VIX riesgo de los mercados': 'VIX',
    'DXY index': 'DXY',
    'USD/EUR (': 'USD/EUR',
    'Petroleo WTI $/Bar': 'Oil WTI',
    'Oro $/oz': 'Gold',
    'Bs/Usd ( tasa oficial)': 'Bs/USD',
    'SIlver USD/Oz': 'Silver',
}

df.columns = [columns.get(col, col) for col in df.columns]

# ============================================================================
# 2. CREACIÓN DEL DASHBOARD CON PLOTLY
# ============================================================================

print("\n2. Generando visualizaciones...")

# Crear subplots con 8 filas y 2 columnas
fig = make_subplots(
    rows=8, cols=2,
    subplot_titles=(
        '📈 S&P 500', '🏦 US Treasury Yields',
        '💵 DXY Index', '📊 VIX Volatility',
        '🛢️ Oil WTI (Crude Oil)', '🥇 Gold (XAU/USD)',
        '🥈 Silver (XAG/USD)', '📉 10-2 Year Spread',
        '💶 USD/EUR Exchange Rate', '🌍 MSCI World',
        '🏛️ Dow Jones', '📱 Nasdaq 100',
        '🇻🇪 Bs/USD (Venezuelan Official Rate)', ''
    ),
    vertical_spacing=0.05,
    horizontal_spacing=0.1,
    row_heights=[0.125, 0.125, 0.125, 0.125, 0.125, 0.125, 0.125, 0.125]
)

# ==================== FILA 1 ====================
# S&P 500
if 'S&P 500' in df.columns:
    fig.add_trace(
        go.Scatter(x=df['Date'], y=df['S&P 500'], mode='lines', name='S&P 500',
                   line=dict(color='#1f77b4', width=2),
                   hovertemplate='<b>S&P 500</b><br>Date: %{x}<br>Value: %{y:.2f}<extra></extra>'),
        row=1, col=1
    )

# US Treasury Yields
if 'US 10YT' in df.columns:
    fig.add_trace(
        go.Scatter(x=df['Date'], y=df['US 10YT'], mode='lines', name='US 10YR',
                   line=dict(color='#ff7f0e', width=2)),
        row=1, col=2
    )
if 'US 2YT' in df.columns:
    fig.add_trace(
        go.Scatter(x=df['Date'], y=df['US 2YT'], mode='lines', name='US 2YR',
                   line=dict(color='#2ca02c', width=2)),
        row=1, col=2
    )

# ==================== FILA 2 ====================
# DXY Index
if 'DXY' in df.columns:
    fig.add_trace(
        go.Scatter(x=df['Date'], y=df['DXY'], mode='lines', name='DXY',
                   line=dict(color='#d62728', width=2),
                   fill='tozeroy', fillcolor='rgba(214, 39, 40, 0.1)'),
        row=2, col=1
    )

# VIX
if 'VIX' in df.columns:
    fig.add_trace(
        go.Scatter(x=df['Date'], y=df['VIX'], mode='lines', name='VIX',
                   line=dict(color='#9467bd', width=2)),
        row=2, col=2
    )
    # Línea de referencia (20 = umbral de volatilidad)
    fig.add_hline(y=20, line_dash="dash", line_color="red", opacity=0.5, row=2, col=2,
                  annotation_text="Volatility Threshold", annotation_position="bottom right")

# ==================== FILA 3 ====================
# Oil WTI
if 'Oil WTI' in df.columns:
    fig.add_trace(
        go.Scatter(x=df['Date'], y=df['Oil WTI'], mode='lines', name='Oil WTI',
                   line=dict(color='#8c564b', width=2),
                   fill='tozeroy', fillcolor='rgba(140, 86, 75, 0.1)'),
        row=3, col=1
    )
    fig.add_hline(y=70, line_dash="dash", line_color="gray", opacity=0.5, row=3, col=1)

# Gold
if 'Gold' in df.columns:
    fig.add_trace(
        go.Scatter(x=df['Date'], y=df['Gold'], mode='lines', name='Gold',
                   line=dict(color='#e377c2', width=2)),
        row=3, col=2
    )

# ==================== FILA 4 ====================
# Silver
if 'Silver' in df.columns:
    fig.add_trace(
        go.Scatter(x=df['Date'], y=df['Silver'], mode='lines', name='Silver',
                   line=dict(color='#7f7f7f', width=2)),
        row=4, col=1
    )

# 10-2 Year Spread
if '10-2 Spread' in df.columns:
    fig.add_trace(
        go.Scatter(x=df['Date'], y=df['10-2 Spread'], mode='lines', name='10-2 Spread',
                   line=dict(color='#bcbd22', width=2, dash='dash')),
        row=4, col=2
    )
    fig.add_hline(y=0, line_dash="dash", line_color="red", opacity=0.5, row=4, col=2,
                  annotation_text="Inversion Signal", annotation_position="bottom right")

# ==================== FILA 5 ====================
# USD/EUR
if 'USD/EUR' in df.columns:
    fig.add_trace(
        go.Scatter(x=df['Date'], y=df['USD/EUR'], mode='lines', name='USD/EUR',
                   line=dict(color='#17becf', width=2)),
        row=5, col=1
    )

# MSCI World
if 'MSCI World' in df.columns:
    fig.add_trace(
        go.Scatter(x=df['Date'], y=df['MSCI World'], mode='lines', name='MSCI World',
                   line=dict(color='#ff9896', width=2)),
        row=5, col=2
    )

# ==================== FILA 6 ====================
# Dow Jones
if 'Dow Jones' in df.columns:
    fig.add_trace(
        go.Scatter(x=df['Date'], y=df['Dow Jones'], mode='lines', name='Dow Jones',
                   line=dict(color='#c5b0d5', width=2)),
        row=6, col=1
    )

# Nasdaq 100
if 'Nasdaq 100' in df.columns:
    fig.add_trace(
        go.Scatter(x=df['Date'], y=df['Nasdaq 100'], mode='lines', name='Nasdaq 100',
                   line=dict(color='#f7b6d2', width=2)),
        row=6, col=2
    )

# ==================== FILA 7 ====================
# Bs/USD (Venezuelan Official Rate)
if 'Bs/USD' in df.columns:
    fig.add_trace(
        go.Scatter(x=df['Date'], y=df['Bs/USD'], mode='lines', name='Bs/USD',
                   line=dict(color='#98df8a', width=2, dash='dot'),
                   fill='tozeroy', fillcolor='rgba(152, 223, 138, 0.2)'),
        row=7, col=1
    )

# ==================== CONFIGURACIÓN DEL LAYOUT ====================

# Actualizar layout principal
fig.update_layout(
    title=dict(
        text="📊 Financial Markets Dashboard - Complete View",
        font=dict(size=28, family="Arial Black"),
        x=0.5,
        xanchor='center'
    ),
    showlegend=True,
    legend=dict(
        orientation="h",
        yanchor="bottom",
        y=1.02,
        xanchor="right",
        x=1,
        font=dict(size=10)
    ),
    hovermode='x unified',
    template='plotly_white',
    height=2200,
    width=1400,
    margin=dict(l=50, r=50, t=100, b=50)
)

# Actualizar ejes X (solo los últimos dos tienen título)
fig.update_xaxes(title_text="Date", row=8, col=1)
fig.update_xaxes(title_text="Date", row=8, col=2)

# Agregar range sliders solo en la última fila
fig.update_xaxes(rangeslider_visible=True, row=8, col=1)
fig.update_xaxes(rangeslider_visible=True, row=8, col=2)

# Actualizar etiquetas de ejes Y
fig.update_yaxes(title_text="Index Value", row=1, col=1)
fig.update_yaxes(title_text="Yield (%)", row=1, col=2)
fig.update_yaxes(title_text="DXY Index", row=2, col=1)
fig.update_yaxes(title_text="VIX Level", row=2, col=2)
fig.update_yaxes(title_text="USD/Barrel", row=3, col=1)
fig.update_yaxes(title_text="USD/Oz", row=3, col=2)
fig.update_yaxes(title_text="USD/Oz", row=4, col=1)
fig.update_yaxes(title_text="Basis Points (bps)", row=4, col=2)
fig.update_yaxes(title_text="USD/EUR", row=5, col=1)
fig.update_yaxes(title_text="Index Value", row=5, col=2)
fig.update_yaxes(title_text="Index Value", row=6, col=1)
fig.update_yaxes(title_text="Index Value", row=6, col=2)
fig.update_yaxes(title_text="Bs/USD", row=7, col=1)

print("   ✓ Visualizaciones generadas correctamente")

# ============================================================================
# 3. CALCULAR MÉTRICAS Y ESTADÍSTICAS
# ============================================================================

print("\n3. Calculando métricas de rendimiento...")

# Diccionario para almacenar métricas
metrics = {}

if len(df) > 0:
    latest_data = df.iloc[-1]
    first_data = df.iloc[0]
    
    metrics['period_start'] = first_data['Date'].strftime('%Y-%m-%d') if pd.notna(first_data['Date']) else 'N/A'
    metrics['period_end'] = latest_data['Date'].strftime('%Y-%m-%d') if pd.notna(latest_data['Date']) else 'N/A'
    
    # Calcular métricas para cada variable
    if 'S&P 500' in df.columns and pd.notna(first_data['S&P 500']) and pd.notna(latest_data['S&P 500']):
        metrics['S&P 500'] = {
            'start': first_data['S&P 500'],
            'end': latest_data['S&P 500'],
            'change': ((latest_data['S&P 500']/first_data['S&P 500'])-1)*100
        }
    
    if 'US 10YT' in df.columns and pd.notna(first_data['US 10YT']) and pd.notna(latest_data['US 10YT']):
        metrics['US 10YR'] = {
            'start': first_data['US 10YT'],
            'end': latest_data['US 10YT'],
            'change': latest_data['US 10YT'] - first_data['US 10YT']
        }
    
    if 'DXY' in df.columns and pd.notna(first_data['DXY']) and pd.notna(latest_data['DXY']):
        metrics['DXY'] = {
            'start': first_data['DXY'],
            'end': latest_data['DXY'],
            'change': ((latest_data['DXY']/first_data['DXY'])-1)*100
        }
    
    if 'Gold' in df.columns and pd.notna(first_data['Gold']) and pd.notna(latest_data['Gold']):
        metrics['Gold'] = {
            'start': first_data['Gold'],
            'end': latest_data['Gold'],
            'change': ((latest_data['Gold']/first_data['Gold'])-1)*100
        }
    
    if 'Silver' in df.columns and pd.notna(first_data['Silver']) and pd.notna(latest_data['Silver']):
        metrics['Silver'] = {
            'start': first_data['Silver'],
            'end': latest_data['Silver'],
            'change': ((latest_data['Silver']/first_data['Silver'])-1)*100
        }
    
    if 'Oil WTI' in df.columns and pd.notna(first_data['Oil WTI']) and pd.notna(latest_data['Oil WTI']):
        metrics['Oil WTI'] = {
            'start': first_data['Oil WTI'],
            'end': latest_data['Oil WTI'],
            'change': ((latest_data['Oil WTI']/first_data['Oil WTI'])-1)*100
        }
    
    if 'Bs/USD' in df.columns and pd.notna(first_data['Bs/USD']) and pd.notna(latest_data['Bs/USD']):
        metrics['Bs/USD'] = {
            'start': first_data['Bs/USD'],
            'end': latest_data['Bs/USD'],
            'change': ((latest_data['Bs/USD']/first_data['Bs/USD'])-1)*100
        }

print("   ✓ Métricas calculadas")

# ============================================================================
# 4. GUARDAR COMO HTML CON DATOS INCRUSTADOS
# ============================================================================

print("\n4. Guardando dashboard como archivo HTML...")

# Convertir datos a JSON para incrustar
df_json = df.to_json(orient='records', date_format='iso')
metrics_json = json.dumps(metrics, default=str)

# Guardar HTML inicial
html_filename = "index.html"
fig.write_html(html_filename, include_plotlyjs='cdn', full_html=True)

# Leer y modificar HTML para incrustar datos
with open(html_filename, "r", encoding='utf-8') as f:
    html_content = f.read()

# Crear script con datos incrustados
script_data = f"""
<script>
// Datos financieros incrustados
var financialData = {df_json};
var metricsData = {metrics_json};
console.log("📊 Dashboard cargado correctamente");
console.log("   Registros:", financialData.length);
console.log("   Período:", metricsData.period_start, "→", metricsData.period_end);
</script>

<style>
/* Estilos personalizados para mejor visualización */
body {{
    font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
    background-color: #f8f9fa;
}}
.header-info {{
    background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
    color: white;
    padding: 20px;
    border-radius: 10px;
    margin: 20px;
    box-shadow: 0 4px 6px rgba(0,0,0,0.1);
}}
.metrics-grid {{
    display: grid;
    grid-template-columns: repeat(auto-fit, minmax(250px, 1fr));
    gap: 15px;
    margin: 20px;
}}
.metric-card {{
    background: white;
    padding: 15px;
    border-radius: 10px;
    box-shadow: 0 2px 4px rgba(0,0,0,0.1);
    text-align: center;
    transition: transform 0.2s;
}}
.metric-card:hover {{
    transform: translateY(-2px);
    box-shadow: 0 4px 8px rgba(0,0,0,0.15);
}}
.metric-name {{
    font-size: 14px;
    color: #666;
    text-transform: uppercase;
    letter-spacing: 1px;
}}
.metric-value {{
    font-size: 20px;
    font-weight: bold;
    margin: 10px 0;
}}
.metric-change {{
    font-size: 14px;
}}
.positive {{
    color: #28a745;
}}
.negative {{
    color: #dc3545;
}}
.footer {{
    text-align: center;
    padding: 20px;
    color: #666;
    font-size: 12px;
    border-top: 1px solid #ddd;
    margin-top: 20px;
}}
</style>
"""

# Agregar panel de métricas al HTML
metrics_html = '<div class="header-info"><h2>📈 Resumen de Mercado</h2>'
metrics_html += f'<p>Período: {metrics.get("period_start", "N/A")} → {metrics.get("period_end", "N/A")}</p></div>'
metrics_html += '<div class="metrics-grid">'

# Definir emojis para cada métrica
emojis = {
    'S&P 500': '📈',
    'US 10YR': '🏦',
    'DXY': '💵',
    'Gold': '🥇',
    'Silver': '🥈',
    'Oil WTI': '🛢️',
    'Bs/USD': '🇻🇪'
}

for key, value in metrics.items():
    if key not in ['period_start', 'period_end'] and isinstance(value, dict):
        emoji = emojis.get(key, '📊')
        change_class = 'positive' if value['change'] > 0 else 'negative'
        change_symbol = '+' if value['change'] > 0 else ''
        
        if 'YR' in key:
            change_text = f"{change_symbol}{value['change']:.2f} bps"
        else:
            change_text = f"{change_symbol}{value['change']:.2f}%"
        
        metrics_html += f'''
        <div class="metric-card">
            <div class="metric-name">{emoji} {key}</div>
            <div class="metric-value">{value['end']:.2f}</div>
            <div class="metric-change {change_class}">{change_text}</div>
            <div style="font-size: 10px; color: #999;">desde {value['start']:.2f}</div>
        </div>
        '''

metrics_html += '</div><div class="footer">'
metrics_html += '📊 Dashboard generado automáticamente | Datos actualizados: ' + datetime.now().strftime('%Y-%m-%d %H:%M:%S')
metrics_html += '<br>🔗 Visualización interactiva - Haz clic y arrastra para explorar los datos'
metrics_html += '</div>'

# Insertar todos los elementos en el HTML
html_content = html_content.replace("</head>", script_data + "</head>")
html_content = html_content.replace("<body>", "<body>\n" + metrics_html)

# Guardar HTML final
with open(html_filename, "w", encoding='utf-8') as f:
    f.write(html_content)

print(f"   ✓ Dashboard guardado como '{html_filename}'")
print(f"   ✓ Tamaño del archivo: {os.path.getsize(html_filename) / 1024:.2f} KB")

# ============================================================================
# 5. IMPRIMIR RESUMEN FINAL
# ============================================================================

print("\n" + "=" * 80)
print("✅ DASHBOARD GENERADO EXITOSAMENTE")
print("=" * 80)

print(f"\n📁 Archivo generado: {html_filename}")
print(f"📊 Registros incluidos: {len(df)}")
print(f"📅 Período: {metrics.get('period_start', 'N/A')} → {metrics.get('period_end', 'N/A')}")

print("\n📈 Resumen de rendimiento:")
for key, value in metrics.items():
    if key not in ['period_start', 'period_end'] and isinstance(value, dict):
        if 'YR' in key:
            print(f"   {key}: {value['start']:.2f} → {value['end']:.2f} ({value['change']:+.2f} bps)")
        else:
            print(f"   {key}: {value['start']:.2f} → {value['end']:.2f} ({value['change']:+.2f}%)")

print("\n" + "=" * 80)
print("🚀 PARA SUBIR A GITHUB PAGES:")
print("=" * 80)
print("\n1. Abre tu navegador y ve a https://github.com")
print("2. Crea un nuevo repositorio (público) llamado 'financial-dashboard'")
print("3. Sube los siguientes archivos:")
print(f"   - {html_filename} (este archivo)")
print("   - Mercury 1.0.xlsx (opcional, los datos ya están incrustados)")
print("4. Ve a Settings → Pages → Branch: main → Save")
print("5. Espera 2-3 minutos y visita:")
print("   https://TU_USUARIO.github.io/financial-dashboard/")
print("\n" + "=" * 80)

# Opcional: abrir el archivo automáticamente en el navegador
try:
    import webbrowser
    webbrowser.open(html_filename)
    print("\n🌐 Abriendo dashboard en tu navegador...")
except:
    pass

print("\n✨ ¡Proceso completado!")