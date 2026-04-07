import tkinter as tk
from tkinter import colorchooser, messagebox
import io
import requests
import pandas as pd
import matplotlib.pyplot as plt
import matplotlib.ticker as ticker

# ==========================================
# CONFIGURACIÓN DE FUENTE ESTILIZADA
# Intenta usar Poppins. Si la PC no la tiene instalada, usa alternativas limpias.
# ==========================================
plt.rcParams['font.family'] = 'sans-serif'
plt.rcParams['font.sans-serif'] = ['Poppins', 'Century Gothic', 'Segoe UI', 'Arial']

class AplicacionGraficas:
    def __init__(self, root):
        self.root = root
        self.root.title("Generador Automático de Gráficas RDM")
        self.root.geometry("400x350")
        self.root.configure(padx=20, pady=20, bg="#F3F4F6")

        # Colores por defecto para las categorías principales
        self.color_procesadas = '#80C680' # Verde
        self.color_en_proceso = '#FFAF4C' # Naranja

        # Título de la interfaz
        tk.Label(root, text="Panel de Control RDM", font=("Segoe UI", 16, "bold"), bg="#F3F4F6").pack(pady=(0, 20))

        # Sección de personalización de colores
        tk.Label(root, text="Personalizar colores del Estatus:", font=("Segoe UI", 10), bg="#F3F4F6").pack()

        self.btn_col_proc = tk.Button(root, text="Color: Procesadas", bg=self.color_procesadas, 
                                      command=self.elegir_color_proc, width=25, relief="flat", borderwidth=1)
        self.btn_col_proc.pack(pady=5)

        self.btn_col_proceso = tk.Button(root, text="Color: En Proceso", bg=self.color_en_proceso, 
                                         command=self.elegir_color_proceso, width=25, relief="flat", borderwidth=1)
        self.btn_col_proceso.pack(pady=5)

        # Botón principal de acción
        self.btn_generar = tk.Button(root, text="Generar y Descargar Gráficas", command=self.ejecutar_proceso, 
                                     bg="#2563EB", fg="white", font=("Segoe UI", 12, "bold"), 
                                     relief="flat", pady=10, padx=20, cursor="hand2")
        self.btn_generar.pack(pady=30)

        # Firma
        tk.Label(root, text="v1.0 - Departamento de Sistemas", font=("Segoe UI", 8, "italic"), bg="#F3F4F6", fg="#6B7280").pack(side="bottom")

    def elegir_color_proc(self):
        color = colorchooser.askcolor(title="Elige color para Procesadas")[1]
        if color:
            self.color_procesadas = color
            self.btn_col_proc.configure(bg=color)

    def elegir_color_proceso(self):
        color = colorchooser.askcolor(title="Elige color para En Proceso")[1]
        if color:
            self.color_en_proceso = color
            self.btn_col_proceso.configure(bg=color)

    def ejecutar_proceso(self):
        self.btn_generar.configure(text="Procesando...", state="disabled", bg="#9CA3AF")
        self.root.update()
        
        try:
            url_invitado = "https://wcisneros-my.sharepoint.com/:x:/g/personal/operaciones_wecisneros_com/IQAoRUfU-XYMR5if6iuQ6N7ZAfJgZjH0DEBkcUZtQkAwVuc?download=1"
            archivo = self.obtener_datos_de_enlace(url_invitado)
            datos = self.procesar_datos_excel(archivo)
            
            if datos['Total'] > 0:
                self.generar_grafico_distribucion(datos)
                self.generar_grafico_estatus(datos)
                messagebox.showinfo("¡Éxito!", "Las gráficas se han generado y guardado en esta carpeta con éxito.\n\nYa puedes insertarlas en tu presentación.")
            else:
                messagebox.showwarning("Advertencia", "El archivo se leyó, pero el total numérico es 0.")
                
        except Exception as e:
            messagebox.showerror("Error", f"Ocurrió un error en el proceso:\n{str(e)}")
            
        finally:
            self.btn_generar.configure(text="Generar y Descargar Gráficas", state="normal", bg="#2563EB")

    # --- LÓGICA INTERNA DE DATOS ---
    def obtener_datos_de_enlace(self, url):
        headers = {'User-Agent': 'Mozilla/5.0'}
        respuesta = requests.get(url, headers=headers)
        respuesta.raise_for_status() 
        return io.BytesIO(respuesta.content)

    def procesar_datos_excel(self, archivo_memoria):
        df = pd.read_excel(archivo_memoria, sheet_name=1, header=1)
        df.rename(columns=lambda x: str(x).strip(), inplace=True)
        df = df.dropna(how='all')
        datos_resumen = df.iloc[-1].to_dict()

        cols_en_proceso = ['Aprobación Inicio', 'Envío de Solicitud de Ofertas', 'Ofertas Recibidas', 'Consultas Adicionales', 'Analísis de Ofertas']
        cols_procesadas = ['Presentación al Comité', 'Adjudicado', 'En Espera de Pago', 'Pagado', 'Recibido']
        
        resumen_numerico = {k: pd.to_numeric(v, errors='coerce') for k, v in datos_resumen.items()}
        for k in resumen_numerico:
            if pd.isna(resumen_numerico[k]):
                resumen_numerico[k] = 0

        rdms_en_proceso = sum(resumen_numerico.get(col, 0) for col in cols_en_proceso)
        rdms_procesadas = sum(resumen_numerico.get(col, 0) for col in cols_procesadas)
        total_rdm = rdms_en_proceso + rdms_procesadas

        return {
            'Individuales': {
                'Aprobación Inicio': resumen_numerico.get('Aprobación Inicio', 0),
                'Envío de Solicitud': resumen_numerico.get('Envío de Solicitud de Ofertas', 0),
                'Consultas Adicionales': resumen_numerico.get('Consultas Adicionales', 0),
                'Análisis de Ofertas': resumen_numerico.get('Analísis de Ofertas', 0),
                'Presentación al Comité': resumen_numerico.get('Presentación al Comité', 0),
                'Adjudicado': resumen_numerico.get('Adjudicado', 0),
                'En Espera de Pago': resumen_numerico.get('En Espera de Pago', 0),
                'Pagado': resumen_numerico.get('Pagado', 0),
                'Recibidas': resumen_numerico.get('Recibido', 0)
            },
            'Resumen': { 'Procesadas': rdms_procesadas, 'En Proceso': rdms_en_proceso },
            'Total': total_rdm
        }

    # --- GENERACIÓN DE IMÁGENES ---
    def generar_grafico_distribucion(self, datos):
        titulos = list(datos['Individuales'].keys())
        valores = list(datos['Individuales'].values())
        colores = ['#FF9999', '#66B2FF', '#99FF99', '#D3D3D3', '#FFCC99', '#FFD700', '#C6C6F6', '#FFB6C1', '#80C680']
        
        fig, ax = plt.figure(figsize=(12, 7)), plt.gca()
        bars = ax.bar(titulos, valores, color=colores, edgecolor='black', width=0.7)
        
        for bar in bars:
            height = bar.get_height()
            ax.text(bar.get_x() + bar.get_width() / 2, height + 0.2, str(int(height)), ha='center', fontsize=12, fontweight='bold')
            
        ax.set_title(f'Distribución de RDMs por Proceso (Total: {int(datos["Total"])})', fontsize=18, fontweight='bold')
        ax.set_ylabel('Cantidad', fontsize=12)
        plt.xticks(rotation=45, ha='right', fontsize=10)
        ax.set_ylim(bottom=0)
        ax.yaxis.set_major_locator(ticker.MaxNLocator(integer=True))
        ax.spines['top'].set_visible(False)
        ax.spines['right'].set_visible(False)
        
        plt.tight_layout()
        plt.savefig('grafico_distribucion_rdm.png', dpi=300, transparent=True) # transparent=True hace que el fondo sea transparente ideal para presentaciones
        plt.close()

    def generar_grafico_estatus(self, datos):
        titulos = list(datos['Resumen'].keys())
        valores = list(datos['Resumen'].values())
        
        fig, ax = plt.figure(figsize=(10, 6)), plt.gca()
        # Usamos los colores seleccionados en la interfaz visual
        bars = ax.bar(titulos, valores, color=[self.color_procesadas, self.color_en_proceso], edgecolor='black', width=0.5)
        
        for bar in bars:
            height = bar.get_height()
            ax.text(bar.get_x() + bar.get_width() / 2, height + 0.5, str(int(height)), ha='center', fontsize=14, fontweight='bold')
            
        ax.set_title(f'Estatus de RDMs Recibidas (Total: {int(datos["Total"])})', fontsize=18, fontweight='bold')
        ax.set_ylim(bottom=0)
        ax.yaxis.set_major_locator(ticker.MaxNLocator(integer=True))
        ax.spines['top'].set_visible(False)
        ax.spines['right'].set_visible(False)
        plt.xticks(fontsize=12)
        
        plt.tight_layout()
        plt.savefig('grafico_estatus_rdm.png', dpi=300, transparent=True)
        plt.close()

if __name__ == "__main__":
    ventana = tk.Tk()
    app = AplicacionGraficas(ventana)
    ventana.mainloop()