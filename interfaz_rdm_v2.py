import os
from dotenv import load_dotenv

# Cargar las variables del archivo .env oculto
load_dotenv()

import tkinter as tk
# ... el resto de tus imports ...
from tkinter import colorchooser, messagebox, filedialog
import io
import os
import re
from datetime import date, datetime
import shutil
import unicodedata
import requests
import pandas as pd
import matplotlib.pyplot as plt
import matplotlib.ticker as ticker
import tkinter.font as tkfont
from tkinter import ttk
from matplotlib.figure import Figure
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
from crear_grafica_personalizada import VentanaGraficaPersonalizada
import os
from dotenv import load_dotenv

# Carga el archivo secreto
load_dotenv()

# ==========================================
# CONFIGURACIÓN DE FUENTE ESTILIZADA
# Intenta usar Poppins. Si la PC no la tiene instalada, usa alternativas limpias.
# ==========================================
# Valores seguros por defecto. Se sobreescriben en __init__ si Poppins existe.
plt.rcParams['font.family'] = 'sans-serif'
plt.rcParams['font.sans-serif'] = ['Segoe UI', 'Arial', 'Century Gothic']
font_interface = ("Segoe UI", 10)
font_title = ("Segoe UI", 16, "bold")
font_button = ("Segoe UI", 11, "bold")

class AplicacionGraficasAvanzada:
    def __init__(self, root):
        self.root = root
        self.root.title("Generador Automático de Gráficas RDM - v2.0")
        self.root.geometry("1100x900") # Ventana más grande para la vista previa
        # Fondo blanco para la interfaz
        self.root.configure(padx=20, pady=20, bg="white")
        self.usa_poppins = self._configurar_fuentes()
        self.url_invitado = os.getenv("URL_SHAREPOINT")
        self.output_dir = os.path.dirname(os.path.abspath(__file__))
        self.ruta_grafico_distribucion = os.path.join(self.output_dir, 'grafico_distribucion_rdm.png')
        self.ruta_grafico_estatus = os.path.join(self.output_dir, 'grafico_estatus_rdm.png')

        # Colores por defecto para las categorías individuales
        # Estos colores son similares a los de tu imagen original
        self.colores_distribucion = {
            'Aprobación Inicio': '#FF9999',       # Rojo claro
            'Envío de Solicitud': '#66B2FF',      # Azul cielo
            'Consultas Adicionales': '#99FF99',  # Verde claro
            'Análisis de Ofertas': '#D3D3D3',     # Gris claro
            'Presentación al Comité': '#FFCC99', # Naranja peachy
            'Adjudicado': '#FFD700',            # Amarillo dorado
            'En Espera de Pago': '#C6C6F6',       # Lavanda
            'Pagado': '#FFB6C1',               # Rosa suave
            'Recibidas': '#80C680'              # Verde medio
        }

        # Colores para la gráfica de resumen (Procesadas vs En Proceso)
        self.colores_resumen = {
            'Procesadas': '#80C680',
            'En Proceso': '#FFAF4C'
        }

        self.categorias_individuales = [
            'Aprobación Inicio',
            'Envío de Solicitud',
            'Consultas Adicionales',
            'Análisis de Ofertas',
            'Presentación al Comité',
            'Adjudicado',
            'En Espera de Pago',
            'Pagado',
            'Recibidas'
        ]
        self.categorias_en_proceso = [
            'Aprobación Inicio',
            'Envío de Solicitud',
            'Consultas Adicionales',
            'Análisis de Ofertas'
        ]
        self.categorias_procesadas = [
            'Presentación al Comité',
            'Adjudicado',
            'En Espera de Pago',
            'Pagado',
            'Recibidas'
        ]

        self.datos_ultima_ejecucion = None

        self.meses = [
            "Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio",
            "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre"
        ]
        hoy = date.today()
        self.mes_var = tk.StringVar(value=self.meses[hoy.month - 1])
        self.anio_var = tk.StringVar(value=str(hoy.year))
        self.usar_periodo_excel = tk.BooleanVar(value=True)
        self.tipo_grafico_desglose = tk.StringVar(value="Barras")
        self.tipo_grafico_estatus = tk.StringVar(value="Barras")
        self.mostrar_porcentajes = tk.BooleanVar(value=True)
        self.preview_window = None
        self.preview_figure_popup = None
        self.preview_canvas_popup = None

        # Título de la interfaz
        self.label_titulo = tk.Label(root, text="", font=font_title, bg="white", fg="black")
        self.label_titulo.pack(pady=(0, 12))

        # Selector de periodo para títulos
        frame_periodo = tk.Frame(root, bg="white")
        frame_periodo.pack(pady=(0, 10))
        tk.Label(frame_periodo, text="Mes:", font=font_interface, bg="white", fg="black").grid(row=0, column=0, padx=(0, 6))
        self.combo_mes = ttk.Combobox(frame_periodo, textvariable=self.mes_var, values=self.meses, state="readonly", width=12)
        self.combo_mes.grid(row=0, column=1, padx=(0, 14))
        tk.Label(frame_periodo, text="Año:", font=font_interface, bg="white", fg="black").grid(row=0, column=2, padx=(0, 6))
        self.combo_anio = ttk.Combobox(
            frame_periodo,
            textvariable=self.anio_var,
            values=[str(y) for y in range(hoy.year - 3, hoy.year + 6)],
            state="readonly",
            width=7
        )
        self.combo_anio.grid(row=0, column=3)
        self.combo_mes.bind("<<ComboboxSelected>>", lambda _e: self._al_cambiar_periodo_manual())
        self.combo_anio.bind("<<ComboboxSelected>>", lambda _e: self._al_cambiar_periodo_manual())

        self.chk_periodo_excel = tk.Checkbutton(
            frame_periodo,
            text="Usar periodo automático del Excel",
            variable=self.usar_periodo_excel,
            bg="white",
            fg="black",
            selectcolor="white",
            activebackground="white",
            command=self._al_cambiar_periodo
        )
        self.chk_periodo_excel.grid(row=0, column=4, padx=(14, 0))
        self._actualizar_titulo_panel()

        frame_tipos = tk.Frame(root, bg="white")
        frame_tipos.pack(pady=(0, 10))
        tk.Label(frame_tipos, text="Tipo Desglose:", font=font_interface, bg="white", fg="black").grid(row=0, column=0, padx=(0, 6))
        self.combo_tipo_desglose = ttk.Combobox(
            frame_tipos,
            textvariable=self.tipo_grafico_desglose,
            values=["Barras", "Pastel"],
            state="readonly",
            width=10
        )
        self.combo_tipo_desglose.grid(row=0, column=1, padx=(0, 14))

        tk.Label(frame_tipos, text="Tipo Estatus:", font=font_interface, bg="white", fg="black").grid(row=0, column=2, padx=(0, 6))
        self.combo_tipo_estatus = ttk.Combobox(
            frame_tipos,
            textvariable=self.tipo_grafico_estatus,
            values=["Barras", "Pastel"],
            state="readonly",
            width=10
        )
        self.combo_tipo_estatus.grid(row=0, column=3)

        self.chk_mostrar_porcentajes = tk.Checkbutton(
            frame_tipos,
            text="Mostrar porcentajes",
            variable=self.mostrar_porcentajes,
            bg="white",
            fg="black",
            selectcolor="white",
            activebackground="white",
            command=self.actualizar_vista_previa
        )
        self.chk_mostrar_porcentajes.grid(row=0, column=4, padx=(14, 0))
        self.combo_tipo_desglose.bind("<<ComboboxSelected>>", lambda _e: self.actualizar_vista_previa())
        self.combo_tipo_estatus.bind("<<ComboboxSelected>>", lambda _e: self.actualizar_vista_previa())

        # Sección de personalización de colores individuales
        tk.Label(root, text="Colores de Distribución (personalizar por categoría):", font=font_interface, bg="white", fg="black").pack()

        # Marco para los botones de color individuales, organizado en dos columnas
        frame_colors = tk.Frame(root, bg="white")
        frame_colors.pack(pady=10)

        self.btn_colors = {}
        for i, (categoria, color) in enumerate(self.colores_distribucion.items()):
            col = 0 if i < 5 else 1
            row = i if i < 5 else i - 5
            
            # Botón con el color actual como fondo y texto en negro
            btn = tk.Button(frame_colors, text=f"Color: {categoria}", bg=color, fg="black",
                            command=lambda cat=categoria: self.elegir_color_distribucion(cat),
                            width=28, relief="flat", borderwidth=1)
            btn.grid(row=row, column=col, padx=5, pady=3)
            self.btn_colors[categoria] = btn

        # Colores para la gráfica resumen
        tk.Label(root, text="Colores de Estatus (Procesadas vs En Proceso):", font=font_interface, bg="white", fg="black").pack(pady=(10, 5))
        frame_colors_resumen = tk.Frame(root, bg="white")
        frame_colors_resumen.pack(pady=(0, 8))

        self.btn_colors_resumen = {}
        for i, (categoria, color) in enumerate(self.colores_resumen.items()):
            btn = tk.Button(
                frame_colors_resumen,
                text=f"Color: {categoria}",
                bg=color,
                fg="black",
                command=lambda cat=categoria: self.elegir_color_resumen(cat),
                width=28,
                relief="flat",
                borderwidth=1
            )
            btn.grid(row=0, column=i, padx=5, pady=3)
            self.btn_colors_resumen[categoria] = btn

        # Botón principal de acción
        self.btn_generar = tk.Button(root, text="Generar Gráficas", command=self.ejecutar_proceso, 
                                     bg="#2563EB", fg="white", font=font_button, 
                                     relief="flat", pady=12, padx=25, cursor="hand2")
        self.btn_generar.pack(pady=(16, 8))

        self.btn_vista_previa = tk.Button(
            root,
            text="Vista Previa",
            command=self.abrir_ventana_preview,
            bg="#E5E7EB",
            fg="black",
            font=font_interface,
            relief="flat",
            pady=8,
            padx=14,
            cursor="hand2"
        )
        self.btn_vista_previa.pack(pady=(0, 8))
        frame_manual_personalizada = tk.Frame(root, bg="white")
        frame_manual_personalizada.pack(pady=(0, 10))

        self.btn_manual = tk.Button(
            frame_manual_personalizada,
            text="Ingresar Datos Manuales",
            command=self.abrir_captura_manual,
            bg="#E5E7EB",
            fg="black",
            font=font_interface,
            relief="flat",
            pady=8,
            padx=14,
            cursor="hand2"
        )
        self.btn_manual.grid(row=0, column=0, padx=(0, 8))

        self.btn_crear_grafica = tk.Button(
            frame_manual_personalizada,
            text="Crear Gráfica",
            command=self.abrir_creador_grafica,
            bg="#E5E7EB",
            fg="black",
            font=font_interface,
            relief="flat",
            pady=8,
            padx=14,
            cursor="hand2"
        )
        self.btn_crear_grafica.grid(row=0, column=1)

        frame_analisis = tk.Frame(root, bg="white")
        frame_analisis.pack(pady=(0, 8))
        tk.Label(frame_analisis, text="Analisis porcentual:", font=font_interface, bg="white", fg="black").pack(anchor="w")
        self.label_analisis_resumen = tk.Label(
            frame_analisis,
            text="Procesadas: --% | En Proceso: --%",
            font=font_interface,
            bg="white",
            fg="#374151"
        )
        self.label_analisis_resumen.pack(anchor="w")

        frame_descarga_individual = tk.Frame(root, bg="white")
        frame_descarga_individual.pack(pady=(0, 8))
        self.btn_descargar_desglose = tk.Button(
            frame_descarga_individual,
            text="Descargar Desglose",
            command=lambda: self.descargar_grafica_individual("desglose"),
            bg="#F3F4F6",
            fg="black",
            relief="flat",
            borderwidth=1,
            width=22
        )
        self.btn_descargar_desglose.grid(row=0, column=0, padx=8)

        self.btn_descargar_estatus = tk.Button(
            frame_descarga_individual,
            text="Descargar Estatus",
            command=lambda: self.descargar_grafica_individual("estatus"),
            bg="#F3F4F6",
            fg="black",
            relief="flat",
            borderwidth=1,
            width=22
        )
        self.btn_descargar_estatus.grid(row=0, column=1, padx=8)
        # Firma
        tk.Label(root, text="v2.0", font=("Poppins", 8, "italic") if self.usa_poppins else ("Segoe UI", 8, "italic"), bg="white", fg="#6B7280").pack(side="bottom")
        self.actualizar_vista_previa()
        self.root.after(150, self.cargar_datos_reales_para_preview)

    def _configurar_fuentes(self):
        global font_interface, font_title, font_button
        try:
            # Verificamos fuentes cuando Tk ya tiene una ventana creada.
            available_fonts = tkfont.families(root=self.root)
            if "Poppins" in available_fonts:
                plt.rcParams['font.family'] = 'Poppins'
                font_interface = ("Poppins", 10)
                font_title = ("Poppins", 16, "bold")
                font_button = ("Poppins", 11, "bold")
                return True
        except Exception:
            pass
        return False

    def elegir_color_distribucion(self, categoria):
        # Abre el selector de color de Windows
        color = colorchooser.askcolor(title=f"Elige color para {categoria}")[1]
        if color:
            self.colores_distribucion[categoria] = color
            # Actualiza el botón con el nuevo color
            self.btn_colors[categoria].configure(bg=color)
            self.actualizar_vista_previa()

    def elegir_color_resumen(self, categoria):
        color = colorchooser.askcolor(title=f"Elige color para {categoria}")[1]
        if color:
            self.colores_resumen[categoria] = color
            self.btn_colors_resumen[categoria].configure(bg=color)
            self.actualizar_vista_previa()

    def _periodo_titulo(self):
        return f"{self.mes_var.get()} {self.anio_var.get()}"

    def _actualizar_titulo_panel(self):
        self.label_titulo.configure(text=f"Panel de Control RDM - {self._periodo_titulo()}")

    def _al_cambiar_periodo(self):
        self._actualizar_titulo_panel()
        self.actualizar_vista_previa()

    def _al_cambiar_periodo_manual(self):
        # Si el usuario cambia mes/año manualmente, desactivamos el modo automático.
        self.usar_periodo_excel.set(False)
        self._al_cambiar_periodo()

    def _normalizar_texto(self, texto):
        texto = str(texto)
        texto = unicodedata.normalize('NFKD', texto).encode('ascii', 'ignore').decode('ascii')
        return texto.lower()

    def _detectar_periodo_desde_excel(self, archivo_memoria):
        meses_norm = {
            'enero': 'Enero',
            'febrero': 'Febrero',
            'marzo': 'Marzo',
            'abril': 'Abril',
            'mayo': 'Mayo',
            'junio': 'Junio',
            'julio': 'Julio',
            'agosto': 'Agosto',
            'septiembre': 'Septiembre',
            'setiembre': 'Septiembre',
            'octubre': 'Octubre',
            'noviembre': 'Noviembre',
            'diciembre': 'Diciembre'
        }

        try:
            archivo_memoria.seek(0)
            df_muestra = pd.read_excel(archivo_memoria, sheet_name=1, header=None, nrows=12)
            celdas = [str(v) for v in df_muestra.values.flatten() if pd.notna(v)]
        except Exception:
            return None

        mes_detectado = None
        anio_detectado = None

        for celda in celdas:
            texto = self._normalizar_texto(celda)
            if not mes_detectado:
                for mes_norm, mes_titulo in meses_norm.items():
                    if mes_norm in texto:
                        mes_detectado = mes_titulo
                        break

            if not anio_detectado:
                match = re.search(r'(20\d{2})', texto)
                if match:
                    anio_detectado = int(match.group(1))

            if mes_detectado and anio_detectado:
                break

        if mes_detectado and anio_detectado:
            return mes_detectado, anio_detectado
        return None

    def _formatear_valor_barra(self, valor):
        try:
            numero = float(valor)
            if numero.is_integer():
                return str(int(numero))
            return f"{numero:.2f}".rstrip('0').rstrip('.')
        except (TypeError, ValueError):
            return str(valor)

    def _abreviar_categoria_preview(self, categoria):
        abreviaciones = {
            'Aprobación Inicio': 'Aprob. Inicio',
            'Envío de Solicitud': 'Envio Solicitud',
            'Consultas Adicionales': 'Consultas Adic.',
            'Análisis de Ofertas': 'Analisis Ofertas',
            'Presentación al Comité': 'Present. Comite',
            'Adjudicado': 'Adjudicado',
            'En Espera de Pago': 'Espera de Pago',
            'Pagado': 'Pagado',
            'Recibidas': 'Recibidas'
        }
        return abreviaciones.get(categoria, categoria)

    def _calcular_porcentajes(self, valores):
        total = sum(float(v) for v in valores) if valores else 0
        if total <= 0:
            return [0 for _ in valores]
        return [(float(v) / total) * 100 for v in valores]

    def _etiqueta_con_porcentaje(self, valor, porcentaje):
        if self.mostrar_porcentajes.get():
            return f"{self._formatear_valor_barra(valor)} ({porcentaje:.1f}%)"
        return self._formatear_valor_barra(valor)

    def _label_pie(self, titulo, valor, porcentaje):
        if self.mostrar_porcentajes.get():
            return f"{titulo}\n{self._formatear_valor_barra(valor)} ({porcentaje:.1f}%)"
        return f"{titulo}\n{self._formatear_valor_barra(valor)}"

    def _label_pie_legend(self, titulo, valor, porcentaje):
        if self.mostrar_porcentajes.get():
            return f"{titulo}: {self._formatear_valor_barra(valor)} ({porcentaje:.1f}%)"
        return f"{titulo}: {self._formatear_valor_barra(valor)}"

    def _labels_pie_sin_choque(self, titulos, valores, porcentajes, umbral=4.0):
        labels_visibles = []
        labels_legend = []
        for i, (t, v) in enumerate(zip(titulos, valores)):
            p = porcentajes[i]
            labels_legend.append(self._label_pie_legend(t, v, p))
            if p >= umbral:
                labels_visibles.append(self._label_pie(t, v, p))
            else:
                labels_visibles.append("")
        return labels_visibles, labels_legend

    def _actualizar_analisis_porcentual(self, datos):
        if not datos:
            self.label_analisis_resumen.configure(text="Procesadas: --% | En Proceso: --%")
            return

        valores_resumen = list(datos['Resumen'].values())
        porcentajes = self._calcular_porcentajes(valores_resumen)
        if self.mostrar_porcentajes.get():
            self.label_analisis_resumen.configure(
                text=f"Procesadas: {porcentajes[0]:.1f}% | En Proceso: {porcentajes[1]:.1f}%"
            )
        else:
            self.label_analisis_resumen.configure(
                text="Procesadas: porcentaje oculto | En Proceso: porcentaje oculto"
            )

    def _ruta_salida_unica(self, prefijo):
        periodo = self._periodo_titulo().replace(" ", "_")
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        base = os.path.join(self.output_dir, f"{prefijo}_{periodo}_{timestamp}.png")
        if not os.path.exists(base):
            return base

        n = 2
        while True:
            candidato = os.path.join(self.output_dir, f"{prefijo}_{periodo}_{timestamp}_{n}.png")
            if not os.path.exists(candidato):
                return candidato
            n += 1

    def _construir_datos_desde_individuales(self, individuales):
        en_proceso = sum(float(individuales.get(c, 0)) for c in self.categorias_en_proceso)
        procesadas = sum(float(individuales.get(c, 0)) for c in self.categorias_procesadas)
        total = en_proceso + procesadas
        return {
            'Individuales': individuales,
            'Resumen': {
                'Procesadas': procesadas,
                'En Proceso': en_proceso
            },
            'Total': total
        }

    def _guardar_graficas_en_archivos_unicos(self, datos):
        self.ruta_grafico_distribucion = self._ruta_salida_unica("grafico_distribucion_rdm")
        self.ruta_grafico_estatus = self._ruta_salida_unica("grafico_estatus_rdm")
        self.generar_grafico_distribucion(datos)
        self.generar_grafico_estatus(datos)

    def descargar_grafica_individual(self, tipo):
        if self.datos_ultima_ejecucion is None:
            messagebox.showwarning("Sin datos", "Primero genera gráficas o carga datos manuales.")
            return

        if tipo == "desglose":
            if not os.path.exists(self.ruta_grafico_distribucion):
                self._guardar_graficas_en_archivos_unicos(self.datos_ultima_ejecucion)
            origen = self.ruta_grafico_distribucion
            nombre = os.path.basename(origen)
        else:
            if not os.path.exists(self.ruta_grafico_estatus):
                self._guardar_graficas_en_archivos_unicos(self.datos_ultima_ejecucion)
            origen = self.ruta_grafico_estatus
            nombre = os.path.basename(origen)

        destino = filedialog.asksaveasfilename(
            title="Guardar gráfica",
            defaultextension=".png",
            initialfile=nombre,
            filetypes=[("Imagen PNG", "*.png")]
        )
        if not destino:
            return

        shutil.copy2(origen, destino)
        messagebox.showinfo("Descarga completada", "La gráfica se guardó correctamente.")

    def abrir_captura_manual(self):
        ventana_manual = tk.Toplevel(self.root)
        ventana_manual.title("Carga Manual de Datos")
        ventana_manual.geometry("1180x760")
        ventana_manual.configure(bg="white", padx=16, pady=16)
        ventana_manual.resizable(True, True)

        tk.Label(
            ventana_manual,
            text="Ingresa los valores para cada barra",
            font=font_title,
            bg="white",
            fg="black"
        ).pack(pady=(0, 10))

        frame_contenido = tk.Frame(ventana_manual, bg="white")
        frame_contenido.pack(fill="both", expand=True)

        frame_entries = tk.Frame(frame_contenido, bg="white")
        frame_entries.pack(side="left", fill="y", padx=(0, 14))

        entries = {}
        for i, categoria in enumerate(self.categorias_individuales):
            tk.Label(frame_entries, text=categoria + ":", font=font_interface, bg="white", fg="black", anchor="w", width=24).grid(row=i, column=0, sticky="w", padx=(0, 8), pady=2)
            ent = tk.Entry(frame_entries, width=12)
            ent.insert(0, "0")
            ent.grid(row=i, column=1, pady=2)
            entries[categoria] = ent

        frame_preview_manual = tk.Frame(frame_contenido, bg="white", relief="sunken", borderwidth=1)
        frame_preview_manual.pack(side="left", fill="both", expand=True)
        fig_manual = Figure(figsize=(7.8, 5.2), dpi=100, facecolor='white')
        canvas_manual = FigureCanvasTkAgg(fig_manual, master=frame_preview_manual)
        canvas_manual.get_tk_widget().pack(fill="both", expand=True, padx=8, pady=8)

        def leer_individuales_desde_entries():
            individuales = {}
            for categoria, ent in entries.items():
                texto = ent.get().strip().replace(',', '.')
                if not texto:
                    valor = 0.0
                else:
                    try:
                        valor = float(texto)
                    except ValueError:
                        valor = 0.0
                if valor < 0:
                    valor = 0.0
                individuales[categoria] = valor
            return individuales

        def refrescar_preview_manual(*_args):
            individuales = leer_individuales_desde_entries()
            datos = self._construir_datos_desde_individuales(individuales)
            fig_manual.clear()

            gs = fig_manual.add_gridspec(1, 2, width_ratios=[1.2, 1.0])
            ax1 = fig_manual.add_subplot(gs[0, 0])
            ax2 = fig_manual.add_subplot(gs[0, 1])

            titulos_ind = list(datos['Individuales'].keys())
            valores_ind = list(datos['Individuales'].values())
            porcentajes_ind = self._calcular_porcentajes(valores_ind)
            colores_ind = [self.colores_distribucion.get(t, '#D3D3D3') for t in titulos_ind]
            titulos_ind_preview = [self._abreviar_categoria_preview(t) for t in titulos_ind]
            x_ind = list(range(len(titulos_ind_preview)))

            if self.tipo_grafico_desglose.get() == "Pastel":
                labels_ind, legend_ind = self._labels_pie_sin_choque(titulos_ind_preview, valores_ind, porcentajes_ind, umbral=4.0)
                ax1.pie(
                    valores_ind,
                    labels=labels_ind,
                    colors=colores_ind,
                    startangle=90,
                    labeldistance=1.14,
                    textprops={'fontsize': 8},
                    wedgeprops={'edgecolor': 'white'}
                )
                ax1.legend(legend_ind, loc='center left', bbox_to_anchor=(1.0, 0.5), fontsize=7, frameon=False)
            else:
                ax1.bar(x_ind, valores_ind, color=colores_ind, edgecolor='black', width=0.7)
                ax1.set_xticks(x_ind)
                ax1.set_xticklabels(titulos_ind_preview)
                ax1.tick_params(axis='x', rotation=28, labelsize=8)
                ax1.tick_params(axis='y', labelsize=8)
                ax1.yaxis.set_major_locator(ticker.MaxNLocator(nbins=5, integer=True))
                for i, v in enumerate(valores_ind):
                    ax1.text(i, v + 0.1, self._etiqueta_con_porcentaje(v, porcentajes_ind[i]), ha='center', va='bottom', fontsize=7, fontweight='bold')

            ax1.set_title('Desglose (Manual)', fontsize=11, fontweight='bold')
            ax1.set_facecolor('white')

            titulos_res = list(datos['Resumen'].keys())
            valores_res = list(datos['Resumen'].values())
            porcentajes_res = self._calcular_porcentajes(valores_res)
            colores_res = [self.colores_resumen.get(t, '#D3D3D3') for t in titulos_res]

            if self.tipo_grafico_estatus.get() == "Pastel":
                labels_res, legend_res = self._labels_pie_sin_choque(titulos_res, valores_res, porcentajes_res, umbral=4.0)
                ax2.pie(
                    valores_res,
                    labels=labels_res,
                    colors=colores_res,
                    startangle=90,
                    labeldistance=1.12,
                    textprops={'fontsize': 9},
                    wedgeprops={'edgecolor': 'white'}
                )
                ax2.legend(legend_res, loc='center left', bbox_to_anchor=(1.0, 0.5), fontsize=8, frameon=False)
            else:
                ax2.bar(titulos_res, valores_res, color=colores_res, edgecolor='black', width=0.6)
                max_res = max(valores_res) if valores_res else 1
                ax2.set_ylim(0, max_res * 1.25 if max_res > 0 else 1)
                ax2.yaxis.set_major_locator(ticker.MaxNLocator(nbins=5, integer=True))
                ax2.tick_params(axis='x', labelsize=9)
                ax2.tick_params(axis='y', labelsize=8)
                for i, v in enumerate(valores_res):
                    ax2.text(i, v + max(0.15, max_res * 0.04), self._etiqueta_con_porcentaje(v, porcentajes_res[i]), ha='center', va='bottom', fontsize=8, fontweight='bold')

            ax2.set_title('Estatus (Manual)', fontsize=11, fontweight='bold')
            ax2.set_facecolor('white')

            fig_manual.suptitle(f"Vista previa manual - {self._periodo_titulo()}", fontsize=12, fontweight='bold')
            fig_manual.tight_layout(rect=[0.02, 0.05, 0.98, 0.93])
            canvas_manual.draw()

        def aplicar_datos_manuales():
            try:
                individuales = leer_individuales_desde_entries()

                datos = self._construir_datos_desde_individuales(individuales)
                self.datos_ultima_ejecucion = datos
                self._actualizar_analisis_porcentual(datos)
                self.actualizar_vista_previa()
                self._guardar_graficas_en_archivos_unicos(datos)
                messagebox.showinfo("Éxito", "Se generaron las gráficas con los datos manuales.")
                ventana_manual.destroy()
            except ValueError as e:
                messagebox.showerror("Dato inválido", f"Revisa los valores ingresados.\n\nDetalle: {str(e)}")

        frame_botones = tk.Frame(ventana_manual, bg="white")
        frame_botones.pack(pady=(12, 0))
        tk.Button(frame_botones, text="Generar con datos manuales", command=aplicar_datos_manuales, bg="#2563EB", fg="white", relief="flat", padx=12, pady=6).grid(row=0, column=0, padx=6)
        tk.Button(frame_botones, text="Cancelar", command=ventana_manual.destroy, bg="#E5E7EB", fg="black", relief="flat", padx=12, pady=6).grid(row=0, column=1, padx=6)

        for ent in entries.values():
            ent.bind("<KeyRelease>", refrescar_preview_manual)

        refrescar_preview_manual()

    def abrir_creador_grafica(self):
        VentanaGraficaPersonalizada(self.root, font_interface, font_title)

    def abrir_ventana_preview(self):
        if self.preview_window is not None and self.preview_window.winfo_exists():
            self.preview_window.lift()
            self.preview_window.focus_force()
            self.actualizar_vista_previa()
            return

        self.preview_window = tk.Toplevel(self.root)
        self.preview_window.title("Vista Previa de Gráficas")
        self.preview_window.geometry("1200x760")
        self.preview_window.configure(bg="white")

        frame = tk.Frame(self.preview_window, bg="white")
        frame.pack(fill="both", expand=True, padx=12, pady=12)

        self.preview_figure_popup = Figure(figsize=(12, 7), dpi=100, facecolor='white')
        self.preview_canvas_popup = FigureCanvasTkAgg(self.preview_figure_popup, master=frame)
        self.preview_canvas_popup.get_tk_widget().pack(fill="both", expand=True)

        self.preview_window.bind("<Destroy>", self._cerrar_ventana_preview)
        self.actualizar_vista_previa()

    def _cerrar_ventana_preview(self, event):
        if event.widget == self.preview_window:
            self.preview_window = None
            self.preview_figure_popup = None
            self.preview_canvas_popup = None

    def cargar_datos_reales_para_preview(self):
        try:
            archivo = self.obtener_datos_de_enlace(self.url_invitado)
            if self.usar_periodo_excel.get():
                periodo = self._detectar_periodo_desde_excel(archivo)
                if periodo:
                    mes_detectado, anio_detectado = periodo
                    self.mes_var.set(mes_detectado)
                    self.anio_var.set(str(anio_detectado))
                    self._actualizar_titulo_panel()

            datos = self.procesar_datos_excel(archivo)
            if datos['Total'] > 0:
                self.datos_ultima_ejecucion = datos
                self._actualizar_analisis_porcentual(datos)
                self.actualizar_vista_previa()
        except Exception:
            # Si falla la carga inicial, dejamos la UI usable sin bloquear al usuario.
            pass

    def actualizar_vista_previa(self):
        datos = self.datos_ultima_ejecucion
        self._actualizar_analisis_porcentual(datos)

        if self.preview_window is None or not self.preview_window.winfo_exists() or self.preview_figure_popup is None:
            return

        self.preview_figure_popup.clear()

        if not datos:
            ax_msg = self.preview_figure_popup.add_subplot(1, 1, 1)
            ax_msg.axis('off')
            ax_msg.text(
                0.5,
                0.5,
                "Genera las graficas para ver\nlos numeros reales de la exportacion",
                ha='center',
                va='center',
                fontsize=12,
                fontweight='bold'
            )
            self.preview_figure_popup.patch.set_facecolor('white')
            self.preview_figure_popup.suptitle(f"Vista Previa - {self._periodo_titulo()}", fontsize=12, fontweight='bold')
            self.preview_figure_popup.tight_layout(rect=[0.02, 0.02, 0.98, 0.95])
            self.preview_canvas_popup.draw()
            return

        gs = self.preview_figure_popup.add_gridspec(1, 2, width_ratios=[1.25, 1.0])
        ax1 = self.preview_figure_popup.add_subplot(gs[0, 0])
        ax2 = self.preview_figure_popup.add_subplot(gs[0, 1])
        tam_titulo = 13
        tam_y = 10
        tam_lbl_desglose = 9
        tam_x_res = 12
        tam_lbl_res = 11
        hpad = 1.5

        titulos_ind = list(datos['Individuales'].keys())
        valores_ind = list(datos['Individuales'].values())
        porcentajes_ind = self._calcular_porcentajes(valores_ind)
        colores_ind = [self.colores_distribucion.get(t, '#D3D3D3') for t in titulos_ind]
        titulos_ind_preview = [self._abreviar_categoria_preview(t) for t in titulos_ind]
        x_ind = list(range(len(titulos_ind_preview)))

        if self.tipo_grafico_desglose.get() == "Pastel":
            labels_ind, legend_ind = self._labels_pie_sin_choque(titulos_ind_preview, valores_ind, porcentajes_ind, umbral=4.0)
            ax1.pie(
                valores_ind,
                labels=labels_ind,
                colors=colores_ind,
                startangle=90,
                labeldistance=1.14,
                textprops={'fontsize': 9},
                wedgeprops={'edgecolor': 'white'}
            )
            ax1.legend(legend_ind, loc='center left', bbox_to_anchor=(1.0, 0.5), fontsize=8, frameon=False)
        else:
            ax1.bar(x_ind, valores_ind, color=colores_ind, edgecolor='black', width=0.7)
            ax1.set_xticks(x_ind)
            ax1.set_xticklabels(titulos_ind_preview)
        ax1.set_title('Desglose', fontsize=tam_titulo, fontweight='bold')
        if self.tipo_grafico_desglose.get() != "Pastel":
            ax1.tick_params(axis='x', rotation=25, labelsize=9)
            ax1.tick_params(axis='y', labelsize=tam_y)
            ax1.yaxis.set_major_locator(ticker.MaxNLocator(nbins=6, integer=True))
            ax1.margins(x=0.04)
        ax1.set_facecolor('white')
        if self.tipo_grafico_desglose.get() != "Pastel":
            for i, v in enumerate(valores_ind):
                ax1.text(
                    i,
                    v + 0.18,
                    self._etiqueta_con_porcentaje(v, porcentajes_ind[i]),
                    ha='center',
                    va='bottom',
                    fontsize=tam_lbl_desglose,
                    fontweight='bold'
                )

        titulos_res = list(datos['Resumen'].keys())
        valores_res = list(datos['Resumen'].values())
        porcentajes_res = self._calcular_porcentajes(valores_res)
        colores_res = [self.colores_resumen.get(t, '#D3D3D3') for t in titulos_res]
        if self.tipo_grafico_estatus.get() == "Pastel":
            labels_res, legend_res = self._labels_pie_sin_choque(titulos_res, valores_res, porcentajes_res, umbral=4.0)
            ax2.pie(
                valores_res,
                labels=labels_res,
                colors=colores_res,
                startangle=90,
                labeldistance=1.12,
                textprops={'fontsize': 10},
                wedgeprops={'edgecolor': 'white'}
            )
            ax2.legend(legend_res, loc='center left', bbox_to_anchor=(1.0, 0.5), fontsize=9, frameon=False)
        else:
            ax2.bar(titulos_res, valores_res, color=colores_res, edgecolor='black', width=0.6)
        ax2.set_title('Procesadas vs En Proceso', fontsize=tam_titulo, fontweight='bold')
        if self.tipo_grafico_estatus.get() != "Pastel":
            max_res = max(valores_res) if valores_res else 1
            ax2.set_ylim(0, max_res * 1.25 if max_res > 0 else 1)
            ax2.yaxis.set_major_locator(ticker.MaxNLocator(nbins=6, integer=True))
            ax2.tick_params(axis='x', labelsize=tam_x_res)
            ax2.tick_params(axis='y', labelsize=tam_y)
        ax2.set_facecolor('white')
        if self.tipo_grafico_estatus.get() != "Pastel":
            max_res = max(valores_res) if valores_res else 1
            for i, v in enumerate(valores_res):
                ax2.text(
                    i,
                    v + max(0.3, max_res * 0.04),
                    self._etiqueta_con_porcentaje(v, porcentajes_res[i]),
                    ha='center',
                    va='bottom',
                    fontsize=tam_lbl_res,
                    fontweight='bold'
                )

        self.preview_figure_popup.patch.set_facecolor('white')
        self.preview_figure_popup.suptitle(f"Vista Previa - {self._periodo_titulo()}", fontsize=tam_titulo + 2, fontweight='bold')
        self.preview_figure_popup.tight_layout(rect=[0.02, 0.06, 0.98, 0.93], h_pad=hpad)
        self.preview_canvas_popup.draw()

    def ejecutar_proceso(self):
        # Cambia el estado del botón durante el procesamiento
        self.btn_generar.configure(text="Procesando...", state="disabled", bg="#9CA3AF")
        self.root.update()
        
        try:
            archivo = self.obtener_datos_de_enlace(self.url_invitado)

            if self.usar_periodo_excel.get():
                periodo = self._detectar_periodo_desde_excel(archivo)
                if periodo:
                    mes_detectado, anio_detectado = periodo
                    self.mes_var.set(mes_detectado)
                    self.anio_var.set(str(anio_detectado))
                    self._al_cambiar_periodo()

            datos = self.procesar_datos_excel(archivo)
            
            if datos['Total'] > 0:
                self._guardar_graficas_en_archivos_unicos(datos)
                self.datos_ultima_ejecucion = datos
                self._actualizar_analisis_porcentual(datos)
                self.actualizar_vista_previa()
                # Alerta de éxito con tipografía limpia
                messagebox.showinfo("¡Éxito!", "Las gráficas se generaron sin sobrescribir archivos anteriores.")
            else:
                messagebox.showwarning("Advertencia", "El archivo se leyó, pero el total numérico es 0.")
                
        except Exception as e:
            # Alerta de error
            messagebox.showerror("Error", f"Ocurrió un error en el proceso:\n{str(e)}")
            
        finally:
            # Restaura el botón
            self.btn_generar.configure(text="Generar Gráficas", state="normal", bg="#2563EB")

    # --- LÓGICA INTERNA DE DATOS ---
    def obtener_datos_de_enlace(self, url):
        headers = {'User-Agent': 'Mozilla/5.0'}
        respuesta = requests.get(url, headers=headers)
        respuesta.raise_for_status() 
        return io.BytesIO(respuesta.content)

    def procesar_datos_excel(self, archivo_memoria):
        # sheet_name=1 -> Siempre lee la SEGUNDA pestaña del Excel.
        # header=1     -> Siempre busca los títulos en la FILA 2 de Excel.
        archivo_memoria.seek(0)
        df = pd.read_excel(archivo_memoria, sheet_name=1, header=1)
        # Limpieza de nombres de columnas
        df.rename(columns=lambda x: str(x).strip(), inplace=True)
        # Limpieza de filas vacías
        df = df.dropna(how='all')
        
        # Capturamos la última fila de datos (dinámico para cualquier mes)
        datos_resumen = df.iloc[-1].to_dict()

        # Agrupación de columnas según tu lógica de procesos
        cols_en_proceso = ['Aprobación Inicio', 'Envío de Solicitud de Ofertas', 'Ofertas Recibidas', 'Consultas Adicionales', 'Analísis de Ofertas']
        cols_procesadas = ['Presentación al Comité', 'Adjudicado', 'En Espera de Pago', 'Pagado', 'Recibido']
        
        # Conversión a números para evitar errores de formato en Excel
        resumen_numerico = {k: pd.to_numeric(v, errors='coerce') for k, v in datos_resumen.items()}
        for k in resumen_numerico:
            if pd.isna(resumen_numerico[k]):
                resumen_numerico[k] = 0

        # Cálculos de subtotales
        rdms_en_proceso = sum(resumen_numerico.get(col, 0) for col in cols_en_proceso)
        rdms_procesadas = sum(resumen_numerico.get(col, 0) for col in cols_procesadas)
        total_rdm = rdms_en_proceso + rdms_procesadas

        # Estructura para las gráficas
        datos_finales = {
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
            'Resumen': {
                'Procesadas': rdms_procesadas,
                'En Proceso': rdms_en_proceso
            },
            'Total': total_rdm
        }
        return datos_finales

    # --- GENERACIÓN DE IMÁGENES (CON FUENTE POPIINS/LIMPIA) ---
    def generar_grafico_distribucion(self, datos):
        titulos = list(datos['Individuales'].keys())
        valores = list(datos['Individuales'].values())
        porcentajes = self._calcular_porcentajes(valores)
        
        # Usamos los colores seleccionados en la interfaz visual
        colores = [self.colores_distribucion.get(t, '#D3D3D3') for t in titulos]
        
        fig, ax = plt.figure(figsize=(12, 7)), plt.gca()
        fig.patch.set_facecolor('white')
        ax.set_facecolor('white')
        if self.tipo_grafico_desglose.get() == "Pastel":
            labels, legend_items = self._labels_pie_sin_choque(titulos, valores, porcentajes, umbral=4.0)
            ax.pie(
                valores,
                labels=labels,
                colors=colores,
                startangle=90,
                labeldistance=1.16,
                wedgeprops={'edgecolor': 'white'}
            )
            ax.legend(legend_items, loc='center left', bbox_to_anchor=(1.02, 0.5), fontsize=10, frameon=False)
        else:
            bars = ax.bar(titulos, valores, color=colores, edgecolor='black', width=0.7)
            for i, bar in enumerate(bars):
                height = bar.get_height()
                ax.text(
                    bar.get_x() + bar.get_width() / 2,
                    height + 0.2,
                    self._etiqueta_con_porcentaje(height, porcentajes[i]),
                    ha='center',
                    fontsize=12,
                    fontweight='bold'
                )
            
        ax.set_title(f'Distribución de RDMs por Proceso - {self._periodo_titulo()} (Total: {int(datos["Total"])})', fontsize=18, fontweight='bold')
        if self.tipo_grafico_desglose.get() != "Pastel":
            ax.set_ylabel('Cantidad', fontsize=12)
            plt.xticks(rotation=45, ha='right', fontsize=10)
            ax.set_ylim(bottom=0)
            ax.yaxis.set_major_locator(ticker.MaxNLocator(integer=True))
            ax.spines['top'].set_visible(False)
            ax.spines['right'].set_visible(False)
        
        plt.tight_layout()
        plt.savefig(self.ruta_grafico_distribucion, dpi=300, facecolor='white', transparent=False)
        plt.close()

    def generar_grafico_estatus(self, datos):
        titulos = list(datos['Resumen'].keys())
        valores = list(datos['Resumen'].values())
        porcentajes = self._calcular_porcentajes(valores)
        colores = [self.colores_resumen.get(t, '#D3D3D3') for t in titulos]
        
        fig, ax = plt.figure(figsize=(10, 6)), plt.gca()
        fig.patch.set_facecolor('white')
        ax.set_facecolor('white')
        if self.tipo_grafico_estatus.get() == "Pastel":
            labels, legend_items = self._labels_pie_sin_choque(titulos, valores, porcentajes, umbral=4.0)
            ax.pie(
                valores,
                labels=labels,
                colors=colores,
                startangle=90,
                labeldistance=1.14,
                wedgeprops={'edgecolor': 'white'}
            )
            ax.legend(legend_items, loc='center left', bbox_to_anchor=(1.02, 0.5), fontsize=10, frameon=False)
        else:
            bars = ax.bar(titulos, valores, color=colores, edgecolor='black', width=0.5)
            for i, bar in enumerate(bars):
                height = bar.get_height()
                ax.text(
                    bar.get_x() + bar.get_width() / 2,
                    height + 0.5,
                    self._etiqueta_con_porcentaje(height, porcentajes[i]),
                    ha='center',
                    fontsize=14,
                    fontweight='bold'
                )
            
        ax.set_title(f'Estatus de RDMs Recibidas - {self._periodo_titulo()} (Total: {int(datos["Total"])})', fontsize=18, fontweight='bold')
        if self.tipo_grafico_estatus.get() != "Pastel":
            ax.set_ylim(bottom=0)
            ax.yaxis.set_major_locator(ticker.MaxNLocator(integer=True))
            ax.spines['top'].set_visible(False)
            ax.spines['right'].set_visible(False)
            plt.xticks(fontsize=12)
        
        plt.tight_layout()
        plt.savefig(self.ruta_grafico_estatus, dpi=300, facecolor='white', transparent=False)
        plt.close()

if __name__ == "__main__":
    ventana = tk.Tk()
    app = AplicacionGraficasAvanzada(ventana)
    ventana.mainloop()