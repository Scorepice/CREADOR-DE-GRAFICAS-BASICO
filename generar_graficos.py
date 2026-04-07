import io
import requests
import pandas as pd
import matplotlib.pyplot as plt
import matplotlib.ticker as ticker

# ==========================================
# 1. ENLACE DE INVITADO
# ==========================================
url_invitado = "https://wcisneros-my.sharepoint.com/:x:/g/personal/operaciones_wecisneros_com/IQAoRUfU-XYMR5if6iuQ6N7ZAfJgZjH0DEBkcUZtQkAwVuc?download=1"

# ==========================================
# 2. PROCESAMIENTO DINÁMICO DE DATOS
# ==========================================
def obtener_datos_de_enlace(url):
    print("Conectando al enlace de invitado...")
    headers = {'User-Agent': 'Mozilla/5.0'}
    try:
        respuesta = requests.get(url, headers=headers)
        respuesta.raise_for_status() 
        return io.BytesIO(respuesta.content)
    except Exception as e:
        raise Exception(f"Error al descargar el archivo: {e}")

def procesar_datos_excel(archivo_memoria):
    print("Accediendo a la segunda hoja y extrayendo totales...")
    
    # Leemos la hoja 2 (sheet_name=1) y le decimos que los títulos están en la fila 2 (header=1)
    df = pd.read_excel(archivo_memoria, sheet_name=1, header=1)
    
    # "Bala de plata": Limpiamos espacios invisibles al principio o final de los nombres de las columnas
    df.rename(columns=lambda x: str(x).strip(), inplace=True)
    
    # Limpieza de filas 100% vacías al final
    df = df.dropna(how='all')
    
    # Capturamos la ÚLTIMA fila de la tabla (donde están tus totales numéricos)
    datos_resumen = df.iloc[-1].to_dict()

    # Agrupación ESTRICTA de columnas (Ignoramos Proveedor, Comentarios, RDM, etc.)
    # Nota: Usamos la ortografía exacta que arrojó tu terminal para "Analísis"
    cols_en_proceso = [
        'Aprobación Inicio', 
        'Envío de Solicitud de Ofertas', 
        'Ofertas Recibidas', 
        'Consultas Adicionales', 
        'Analísis de Ofertas' 
    ]
    
    cols_procesadas = [
        'Presentación al Comité', 
        'Adjudicado', 
        'En Espera de Pago', 
        'Pagado', 
        'Recibido'
    ]
    
    # Convertimos los valores a números (si dice 'NaN' o hay texto por error, lo vuelve 0)
    resumen_numerico = {k: pd.to_numeric(v, errors='coerce') for k, v in datos_resumen.items()}
    for k in resumen_numerico:
        if pd.isna(resumen_numerico[k]):
            resumen_numerico[k] = 0

    # Cálculos de subtotales
    rdms_en_proceso = sum(resumen_numerico.get(col, 0) for col in cols_en_proceso)
    rdms_procesadas = sum(resumen_numerico.get(col, 0) for col in cols_procesadas)
    
    # El Total lo calculamos nosotros mismos sumando todo, sin depender de la columna "Compromiso o Cierre"
    total_rdm = rdms_en_proceso + rdms_procesadas

    # Estructura para las gráficas
    datos_finales = {
        'Individuales': {
            'Aprobación Inicio': resumen_numerico.get('Aprobación Inicio', 0),
            'Envío de Solicitud': resumen_numerico.get('Envío de Solicitud de Ofertas', 0),
            'Consultas Adicionales': resumen_numerico.get('Consultas Adicionales', 0),
            'Análisis de Ofertas': resumen_numerico.get('Analísis de Ofertas', 0), # Etiqueta gráfica correcta vs Nombre raro en Excel
            'Presentación al Comité': resumen_numerico.get('Presentación al Comité', 0),
            'Adjudicado': resumen_numerico.get('Adjudicado', 0),
            'En Espera de Pago': resumen_numerico.get('En Espera de Pago', 0),
            'Pagado': resumen_numerico.get('Pagado', 0),
            'Recibidas': resumen_numerico.get('Recibido', 0)
        },
        'Resumen': {
            'Procesadas': rdms_procesadas,
            'En Proceso': rdms_en_proceso
        },
        'Total': total_rdm
    }
            
    return datos_finales

# ==========================================
# 3. GENERACIÓN DE IMÁGENES (PNG)
# ==========================================
def generar_grafico_distribucion(datos):
    print("Dibujando Gráfico de Distribución...")
    titulos = list(datos['Individuales'].keys())
    valores = list(datos['Individuales'].values())
    colores = ['#FF9999', '#66B2FF', '#99FF99', '#D3D3D3', '#FFCC99', '#FFD700', '#C6C6F6', '#FFB6C1', '#80C680']
    
    fig, ax = plt.figure(figsize=(12, 7)), plt.gca()
    bars = ax.bar(titulos, valores, color=colores, edgecolor='black', width=0.7)
    
    for bar in bars:
        height = bar.get_height()
        ax.text(bar.get_x() + bar.get_width() / 2, height + 0.2, str(int(height)), ha='center', fontsize=12, fontweight='bold')
        
    ax.set_title(f'Distribución de RDMs por Proceso (Total: {int(datos["Total"])})', fontsize=15)
    ax.set_ylabel('Cantidad')
    plt.xticks(rotation=45, ha='right')
    ax.set_ylim(bottom=0)
    ax.yaxis.set_major_locator(ticker.MaxNLocator(integer=True))
    ax.spines['top'].set_visible(False)
    ax.spines['right'].set_visible(False)
    
    plt.tight_layout()
    plt.savefig('grafico_distribucion_rdm.png', dpi=300)
    plt.close()

def generar_grafico_estatus(datos):
    print("Dibujando Gráfico de Estatus...")
    titulos = list(datos['Resumen'].keys())
    valores = list(datos['Resumen'].values())
    
    fig, ax = plt.figure(figsize=(10, 6)), plt.gca()
    bars = ax.bar(titulos, valores, color=['#80C680', '#FFAF4C'], edgecolor='black', width=0.5)
    
    for bar in bars:
        height = bar.get_height()
        ax.text(bar.get_x() + bar.get_width() / 2, height + 0.5, str(int(height)), ha='center', fontsize=14, fontweight='bold')
        
    ax.set_title(f'Estatus de RDMs Recibidas (Total: {int(datos["Total"])})', fontsize=15)
    ax.set_ylim(bottom=0)
    ax.yaxis.set_major_locator(ticker.MaxNLocator(integer=True))
    ax.spines['top'].set_visible(False)
    ax.spines['right'].set_visible(False)
    
    plt.tight_layout()
    plt.savefig('grafico_estatus_rdm.png', dpi=300)
    plt.close()

# ==========================================
# 4. EJECUCIÓN PRINCIPAL
# ==========================================
if __name__ == "__main__":
    try:
        archivo = obtener_datos_de_enlace(url_invitado)
        datos = procesar_datos_excel(archivo)
        
        if datos['Total'] > 0:
            generar_grafico_distribucion(datos)
            generar_grafico_estatus(datos)
            print("\n¡ÉXITO TOTAL! Revisa tus nuevas gráficas en el panel izquierdo.")
        else:
            print("\nEl total es 0. Asegúrate de que la última fila de la tabla contenga los números.")
            
    except Exception as e:
        print(f"\nError en la ejecución: {e}")