import os
import sys
from dotenv import load_dotenv

def obtener_base_app():
    if getattr(sys, 'frozen', False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))


BASE_APP = obtener_base_app()

# Cargar las variables del archivo .env oculto
load_dotenv(dotenv_path=os.path.join(BASE_APP, '.env'))

import tkinter as tk
# ... el resto de tus imports ...
from tkinter import colorchooser, messagebox, filedialog
import io
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

APP_BG = "#F3F6FB"
CARD_BG = "#FFFFFF"
TEXT_PRIMARY = "#111827"
TEXT_SECONDARY = "#6B7280"
BORDER_COLOR = "#D9E2EC"
PREVIEW_BORDER_DARK = "#1E3A8A"
ACCENT_COLOR = "#1D4ED8"
ACCENT_HOVER = "#1E40AF"
SECONDARY_BUTTON = "#E8EEF8"
SECONDARY_BUTTON_HOVER = "#D7E2F4"
MUTED_BUTTON = "#F3F4F6"
MUTED_BUTTON_HOVER = "#E5E7EB"

# Carga el archivo secreto
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
        self.root.geometry("1180x940")
        self.root.configure(padx=18, pady=18, bg=APP_BG)
        self.usa_poppins = self._configurar_fuentes()
        self._configurar_estilos()
        self.url_invitado = os.getenv("URL_SHAREPOINT")
        self.output_dir = BASE_APP
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
        self.preview_figure_inline = None
        self.preview_canvas_inline = None
        panel_superior = tk.Frame(root, bg=CARD_BG, bd=1, relief="solid", highlightthickness=1, highlightbackground=BORDER_COLOR)
        panel_superior.pack(fill="x", pady=(0, 12))

        tk.Label(
            panel_superior,
            text="Generador Automático de Gráficas RDM",
            font=("Poppins", 18, "bold") if self.usa_poppins else ("Segoe UI", 18, "bold"),
            bg=CARD_BG,
            fg=TEXT_PRIMARY
        ).pack(anchor="w", padx=18, pady=14)
        self.label_titulo = None

        panel_config = tk.Frame(root, bg=CARD_BG, bd=1, relief="solid", highlightthickness=1, highlightbackground=BORDER_COLOR)
        panel_config.pack(fill="x", pady=(0, 12))
        tk.Label(panel_config, text="Configuración", font=("Poppins", 12, "bold") if self.usa_poppins else ("Segoe UI", 12, "bold"), bg=CARD_BG, fg=TEXT_PRIMARY).pack(anchor="w", padx=16, pady=(12, 8))

        frame_config_main = tk.Frame(panel_config, bg=CARD_BG)
        frame_config_main.pack(fill="x", padx=16, pady=(0, 12))
        frame_config_izq = tk.Frame(frame_config_main, bg=CARD_BG)
        frame_config_izq.pack(side="left", fill="both", expand=True, padx=(0, 12))
        frame_config_der = tk.Frame(frame_config_main, bg=CARD_BG)
        frame_config_der.pack(side="left", fill="both", expand=True)

        frame_periodo = tk.Frame(frame_config_izq, bg=CARD_BG)
        frame_periodo.pack(anchor="w", pady=(0, 10))
        tk.Label(frame_periodo, text="Mes:", font=font_interface, bg=CARD_BG, fg=TEXT_PRIMARY).grid(row=0, column=0, padx=(0, 6), pady=4)
        self.combo_mes = ttk.Combobox(frame_periodo, textvariable=self.mes_var, values=self.meses, state="readonly", width=12, style="Corporate.TCombobox")
        self.combo_mes.grid(row=0, column=1, padx=(0, 14), pady=4)
        tk.Label(frame_periodo, text="Año:", font=font_interface, bg=CARD_BG, fg=TEXT_PRIMARY).grid(row=0, column=2, padx=(0, 6), pady=4)
        self.combo_anio = ttk.Combobox(frame_periodo, textvariable=self.anio_var, values=[str(y) for y in range(hoy.year - 3, hoy.year + 6)], state="readonly", width=7, style="Corporate.TCombobox")
        self.combo_anio.grid(row=0, column=3, pady=4)
        self.combo_mes.bind("<<ComboboxSelected>>", lambda _e: self._al_cambiar_periodo_manual())
        self.combo_anio.bind("<<ComboboxSelected>>", lambda _e: self._al_cambiar_periodo_manual())

        self.chk_periodo_excel = tk.Checkbutton(frame_config_izq, text="Usar periodo automático del Excel", variable=self.usar_periodo_excel, bg=CARD_BG, fg=TEXT_PRIMARY, selectcolor=CARD_BG, activebackground=CARD_BG, command=self._al_cambiar_periodo)
        self.chk_periodo_excel.pack(anchor="w", pady=(0, 10))
        self._actualizar_titulo_panel()

        frame_tipos = tk.Frame(frame_config_izq, bg=CARD_BG)
        frame_tipos.pack(anchor="w", pady=(0, 14))
        tk.Label(frame_tipos, text="Tipo Desglose:", font=font_interface, bg=CARD_BG, fg=TEXT_PRIMARY).grid(row=0, column=0, padx=(0, 6), pady=4)
        self.combo_tipo_desglose = ttk.Combobox(frame_tipos, textvariable=self.tipo_grafico_desglose, values=["Barras", "Pastel"], state="readonly", width=10, style="Corporate.TCombobox")
        self.combo_tipo_desglose.grid(row=0, column=1, padx=(0, 14), pady=4)
        tk.Label(frame_tipos, text="Tipo Estatus:", font=font_interface, bg=CARD_BG, fg=TEXT_PRIMARY).grid(row=0, column=2, padx=(0, 6), pady=4)
        self.combo_tipo_estatus = ttk.Combobox(frame_tipos, textvariable=self.tipo_grafico_estatus, values=["Barras", "Pastel"], state="readonly", width=10, style="Corporate.TCombobox")
        self.combo_tipo_estatus.grid(row=0, column=3, pady=4)
        self.chk_mostrar_porcentajes = tk.Checkbutton(frame_tipos, text="Mostrar porcentajes", variable=self.mostrar_porcentajes, bg=CARD_BG, fg=TEXT_PRIMARY, selectcolor=CARD_BG, activebackground=CARD_BG, command=self.actualizar_vista_previa)
        self.chk_mostrar_porcentajes.grid(row=0, column=4, padx=(14, 0), pady=4)
        self.combo_tipo_desglose.bind("<<ComboboxSelected>>", lambda _e: self.actualizar_vista_previa())
        self.combo_tipo_estatus.bind("<<ComboboxSelected>>", lambda _e: self.actualizar_vista_previa())

        tk.Label(frame_config_izq, text="Accesos rápidos", font=font_interface, bg=CARD_BG, fg=TEXT_PRIMARY).pack(anchor="w", pady=(0, 6))
        frame_acciones_rapidas_1 = tk.Frame(frame_config_izq, bg=CARD_BG)
        frame_acciones_rapidas_1.pack(anchor="w", pady=(0, 6))

        self.btn_generar = tk.Button(frame_acciones_rapidas_1, text="Generar gráficas", command=self.ejecutar_proceso, bg=ACCENT_COLOR, fg="white", font=font_button, relief="flat", padx=16, pady=8, cursor="hand2", activebackground=ACCENT_HOVER)
        self.btn_generar.pack(side="left")
        self._aplicar_hover(self.btn_generar, ACCENT_COLOR, ACCENT_HOVER)

        self.btn_vista_previa = tk.Button(frame_acciones_rapidas_1, text="Vista previa", command=self.abrir_ventana_preview, bg=SECONDARY_BUTTON, fg=TEXT_PRIMARY, font=font_interface, relief="flat", padx=14, pady=8, cursor="hand2", activebackground=SECONDARY_BUTTON_HOVER)
        self.btn_vista_previa.pack(side="left", padx=(8, 0))
        self._aplicar_hover(self.btn_vista_previa, SECONDARY_BUTTON, SECONDARY_BUTTON_HOVER)

        self.btn_manual = tk.Button(frame_acciones_rapidas_1, text="Ingresar datos manuales", command=self.abrir_captura_manual, bg=MUTED_BUTTON, fg=TEXT_PRIMARY, font=font_interface, relief="flat", padx=14, pady=8, cursor="hand2", activebackground=MUTED_BUTTON_HOVER)
        self.btn_manual.pack(side="left", padx=(8, 0))
        self._aplicar_hover(self.btn_manual, MUTED_BUTTON, MUTED_BUTTON_HOVER)

        frame_acciones_rapidas_2 = tk.Frame(frame_config_izq, bg=CARD_BG)
        frame_acciones_rapidas_2.pack(anchor="w", pady=(0, 4))

        self.btn_crear_grafica = tk.Button(frame_acciones_rapidas_2, text="Crear gráfica", command=self.abrir_creador_grafica, bg=MUTED_BUTTON, fg=TEXT_PRIMARY, font=font_interface, relief="flat", padx=14, pady=8, cursor="hand2", activebackground=MUTED_BUTTON_HOVER)
        self.btn_crear_grafica.pack(side="left")
        self._aplicar_hover(self.btn_crear_grafica, MUTED_BUTTON, MUTED_BUTTON_HOVER)

        self.btn_descargar_desglose = tk.Button(frame_acciones_rapidas_2, text="Descargar desglose", command=lambda: self.descargar_grafica_individual("desglose"), bg=SECONDARY_BUTTON, fg=TEXT_PRIMARY, relief="flat", borderwidth=0, padx=14, pady=8, font=font_interface, activebackground=SECONDARY_BUTTON_HOVER, cursor="hand2")
        self.btn_descargar_desglose.pack(side="left", padx=(8, 0))
        self._aplicar_hover(self.btn_descargar_desglose, SECONDARY_BUTTON, SECONDARY_BUTTON_HOVER)

        self.btn_descargar_estatus = tk.Button(frame_acciones_rapidas_2, text="Descargar estatus", command=lambda: self.descargar_grafica_individual("estatus"), bg=SECONDARY_BUTTON, fg=TEXT_PRIMARY, relief="flat", borderwidth=0, padx=14, pady=8, font=font_interface, activebackground=SECONDARY_BUTTON_HOVER, cursor="hand2")
        self.btn_descargar_estatus.pack(side="left", padx=(8, 0))
        self._aplicar_hover(self.btn_descargar_estatus, SECONDARY_BUTTON, SECONDARY_BUTTON_HOVER)

        tk.Label(frame_config_der, text="Vista previa de datos (Excel)", font=font_interface, bg=CARD_BG, fg=TEXT_PRIMARY).pack(anchor="w", pady=(2, 6))
        frame_preview_inline = tk.Frame(frame_config_der, bg=CARD_BG, relief="solid", borderwidth=1, highlightthickness=2, highlightbackground=PREVIEW_BORDER_DARK)
        frame_preview_inline.pack(fill="both", expand=True)
        self.preview_figure_inline = Figure(figsize=(6.2, 3.0), dpi=100, facecolor=CARD_BG)
        self.preview_canvas_inline = FigureCanvasTkAgg(self.preview_figure_inline, master=frame_preview_inline)
        self.preview_canvas_inline.get_tk_widget().pack(fill="both", expand=True, padx=6, pady=6)

        panel_colores = tk.Frame(root, bg=CARD_BG, bd=1, relief="solid", highlightthickness=1, highlightbackground=BORDER_COLOR)
        panel_colores.pack(fill="x", pady=(0, 12))
        tk.Label(panel_colores, text="Personalización de colores", font=("Poppins", 12, "bold") if self.usa_poppins else ("Segoe UI", 12, "bold"), bg=CARD_BG, fg=TEXT_PRIMARY).pack(anchor="w", padx=16, pady=(12, 2))
        frame_estado_superior = tk.Frame(panel_colores, bg=CARD_BG)
        frame_estado_superior.pack(fill="x", padx=16, pady=(0, 8))
        tk.Label(frame_estado_superior, text="Estatus general", font=font_interface, bg=CARD_BG, fg=TEXT_PRIMARY).pack(anchor="w", pady=(0, 6))

        frame_colors_resumen = tk.Frame(frame_estado_superior, bg=CARD_BG)
        frame_colors_resumen.pack(anchor="center")
        self.btn_colors_resumen = {}
        for i, (categoria, color) in enumerate(self.colores_resumen.items()):
            btn = tk.Button(
                frame_colors_resumen,
                text=f"{categoria}",
                bg=color,
                fg=self._color_texto_contrastante(color),
                command=lambda cat=categoria: self.elegir_color_resumen(cat),
                relief="solid",
                borderwidth=1,
                highlightthickness=1,
                highlightbackground="#94A3B8",
                width=18,
                padx=8,
                pady=6,
                font=font_interface,
                cursor="hand2"
            )
            btn.grid(row=0, column=i, padx=8)
            self._aplicar_hover(btn, color)
            self.btn_colors_resumen[categoria] = btn

        tk.Label(panel_colores, text="Desglose por proceso", font=font_interface, bg=CARD_BG, fg=TEXT_PRIMARY).pack(anchor="w", padx=16, pady=(0, 6))

        frame_colors = tk.Frame(panel_colores, bg=CARD_BG)
        frame_colors.pack(fill="x", padx=16, pady=(0, 8))
        self.btn_colors = {}
        categorias_distribucion = list(self.colores_distribucion.items())
        for i, (categoria, color) in enumerate(categorias_distribucion):
            row = i // 3
            col = i % 3
            btn = tk.Button(frame_colors, text=f"{categoria}", bg=color, fg=self._color_texto_contrastante(color), command=lambda cat=categoria: self.elegir_color_distribucion(cat), relief="solid", borderwidth=1, highlightthickness=1, highlightbackground="#94A3B8", padx=10, pady=6, font=font_interface, cursor="hand2")
            btn.grid(row=row, column=col, sticky="ew", padx=6, pady=5)
            frame_colors.grid_columnconfigure(col, weight=1)
            self._aplicar_hover(btn, color)
            self.btn_colors[categoria] = btn

        frame_analisis = tk.Frame(root, bg=CARD_BG, bd=1, relief="solid", highlightthickness=1, highlightbackground=BORDER_COLOR)
        frame_analisis.pack(fill="x", pady=(0, 12))
        tk.Label(frame_analisis, text="Análisis porcentual", font=("Poppins", 12, "bold") if self.usa_poppins else ("Segoe UI", 12, "bold"), bg=CARD_BG, fg=TEXT_PRIMARY).pack(anchor="w", padx=16, pady=(12, 2))
        tk.Label(frame_analisis, text="Resumen rápido de la distribución procesada.", font=font_interface, bg=CARD_BG, fg=TEXT_SECONDARY).pack(anchor="w", padx=16, pady=(0, 6))
        self.label_analisis_resumen = tk.Label(frame_analisis, text="Procesadas: --% | En Proceso: --%", font=font_interface, bg=CARD_BG, fg=TEXT_PRIMARY)
        self.label_analisis_resumen.pack(anchor="w", padx=16, pady=(0, 14))

        tk.Label(root, text="v2.0", font=("Poppins", 8, "italic") if self.usa_poppins else ("Segoe UI", 8, "italic"), bg=APP_BG, fg=TEXT_SECONDARY).pack(side="bottom")
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

    def _configurar_estilos(self):
        style = ttk.Style(self.root)
        try:
            style.theme_use("clam")
        except Exception:
            pass

        style.configure(
            "Corporate.TCombobox",
            fieldbackground="white",
            background="white",
            foreground=TEXT_PRIMARY,
            bordercolor=BORDER_COLOR,
            arrowcolor=TEXT_SECONDARY,
            padding=6,
            relief="flat"
        )
        style.map(
            "Corporate.TCombobox",
            fieldbackground=[("readonly", "white")],
            foreground=[("readonly", TEXT_PRIMARY)]
        )

    def _color_texto_contrastante(self, color_hex):
        try:
            color_hex = color_hex.lstrip("#")
            rojo = int(color_hex[0:2], 16)
            verde = int(color_hex[2:4], 16)
            azul = int(color_hex[4:6], 16)
            luminosidad = (0.299 * rojo + 0.587 * verde + 0.114 * azul) / 255
            return "#111827" if luminosidad > 0.62 else "white"
        except Exception:
            return "#111827"

    def _aclarar_color(self, color_hex, factor=0.12):
        try:
            color_hex = color_hex.lstrip("#")
            rojo = int(color_hex[0:2], 16)
            verde = int(color_hex[2:4], 16)
            azul = int(color_hex[4:6], 16)
            rojo = int(rojo + (255 - rojo) * factor)
            verde = int(verde + (255 - verde) * factor)
            azul = int(azul + (255 - azul) * factor)
            return f"#{rojo:02x}{verde:02x}{azul:02x}"
        except Exception:
            return color_hex

    def _aplicar_hover(self, button, color_base, color_hover=None):
        if color_hover is None:
            color_hover = self._aclarar_color(color_base, 0.1)

        def entrar(_event):
            button.configure(bg=color_hover)

        def salir(_event):
            button.configure(bg=color_base)

        button.bind("<Enter>", entrar)
        button.bind("<Leave>", salir)

    def elegir_color_distribucion(self, categoria):
        # Abre el selector de color de Windows
        color = colorchooser.askcolor(title=f"Elige color para {categoria}")[1]
        if color:
            self.colores_distribucion[categoria] = color
            # Actualiza el botón con el nuevo color
            self.btn_colors[categoria].configure(bg=color)
            self.btn_colors[categoria].configure(fg=self._color_texto_contrastante(color))
            self._aplicar_hover(self.btn_colors[categoria], color)
            self.actualizar_vista_previa()

    def elegir_color_resumen(self, categoria):
        color = colorchooser.askcolor(title=f"Elige color para {categoria}")[1]
        if color:
            self.colores_resumen[categoria] = color
            self.btn_colors_resumen[categoria].configure(bg=color)
            self.btn_colors_resumen[categoria].configure(fg=self._color_texto_contrastante(color))
            self._aplicar_hover(self.btn_colors_resumen[categoria], color)
            self.actualizar_vista_previa()

    def _periodo_titulo(self):
        return f"{self.mes_var.get()} {self.anio_var.get()}"

    def _actualizar_titulo_panel(self):
        if self.label_titulo is not None:
            self.label_titulo.configure(text=f"Panel de Control RDM - {self._periodo_titulo()}")

    def _render_preview_figure(self, figura, datos, modo="popup"):
        figura.clear()

        if not datos:
            ax_msg = figura.add_subplot(1, 1, 1)
            ax_msg.axis('off')
            ax_msg.text(
                0.5,
                0.5,
                "Genera las graficas para ver\nlos numeros reales de la exportacion",
                ha='center',
                va='center',
                fontsize=10 if modo == "inline" else 12,
                fontweight='bold'
            )
            figura.patch.set_facecolor('white')
            if modo == "popup":
                figura.suptitle(f"Vista Previa - {self._periodo_titulo()}", fontsize=12, fontweight='bold')
                figura.tight_layout(rect=[0.02, 0.02, 0.98, 0.95])
            else:
                figura.tight_layout(rect=[0.02, 0.03, 0.98, 0.97])
            return

        if modo == "inline":
            tam_titulo = 10
            tam_y = 8
            tam_lbl_desglose = 7
            tam_x_res = 8
            tam_lbl_res = 8
            hpad = 1.0
        else:
            tam_titulo = 13
            tam_y = 10
            tam_lbl_desglose = 9
            tam_x_res = 12
            tam_lbl_res = 11
            hpad = 1.5

        gs = figura.add_gridspec(1, 2, width_ratios=[1.25, 1.0])
        ax1 = figura.add_subplot(gs[0, 0])
        ax2 = figura.add_subplot(gs[0, 1])

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
                textprops={'fontsize': 7 if modo == "inline" else 9},
                wedgeprops={'edgecolor': 'white'}
            )
        else:
            ax1.bar(x_ind, valores_ind, color=colores_ind, edgecolor='black', width=0.7)
            ax1.set_xticks(x_ind)
            ax1.set_xticklabels(titulos_ind_preview)
        ax1.set_title('Desglose', fontsize=tam_titulo, fontweight='bold')
        if self.tipo_grafico_desglose.get() != "Pastel":
            ax1.tick_params(axis='x', rotation=22 if modo == "inline" else 25, labelsize=7 if modo == "inline" else 9)
            ax1.tick_params(axis='y', labelsize=tam_y)
            ax1.yaxis.set_major_locator(ticker.MaxNLocator(nbins=4 if modo == "inline" else 6, integer=True))
            ax1.margins(x=0.04)
        ax1.set_facecolor('white')
        if self.tipo_grafico_desglose.get() != "Pastel":
            max_ind = max(valores_ind) if valores_ind else 1
            for i, v in enumerate(valores_ind):
                ax1.text(
                    i,
                    v + (max(0.12, max_ind * 0.03) if modo == "inline" else 0.18),
                    self._formatear_valor_barra(v) if modo == "inline" else self._etiqueta_con_porcentaje(v, porcentajes_ind[i]),
                    ha='center',
                    va='bottom',
                    fontsize=7 if modo == "inline" else tam_lbl_desglose,
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
                textprops={'fontsize': 8 if modo == "inline" else 10},
                wedgeprops={'edgecolor': 'white'}
            )
        else:
            ax2.bar(titulos_res, valores_res, color=colores_res, edgecolor='black', width=0.6)
        ax2.set_title('Procesadas vs En Proceso', fontsize=tam_titulo, fontweight='bold')
        if self.tipo_grafico_estatus.get() != "Pastel":
            max_res = max(valores_res) if valores_res else 1
            ax2.set_ylim(0, max_res * 1.25 if max_res > 0 else 1)
            ax2.yaxis.set_major_locator(ticker.MaxNLocator(nbins=4 if modo == "inline" else 6, integer=True))
            ax2.tick_params(axis='x', labelsize=tam_x_res)
            ax2.tick_params(axis='y', labelsize=tam_y)
        ax2.set_facecolor('white')
        if self.tipo_grafico_estatus.get() != "Pastel":
            max_res = max(valores_res) if valores_res else 1
            for i, v in enumerate(valores_res):
                ax2.text(
                    i,
                    v + (max(0.12, max_res * 0.03) if modo == "inline" else max(0.3, max_res * 0.04)),
                    self._formatear_valor_barra(v) if modo == "inline" else self._etiqueta_con_porcentaje(v, porcentajes_res[i]),
                    ha='center',
                    va='bottom',
                    fontsize=8 if modo == "inline" else tam_lbl_res,
                    fontweight='bold'
                )

        for ax in (ax1, ax2):
            color_borde = PREVIEW_BORDER_DARK if modo == "inline" else '#94A3B8'
            for spine in ax.spines.values():
                spine.set_visible(True)
                spine.set_color(color_borde)
                spine.set_linewidth(1.0)

        figura.patch.set_facecolor('white')
        if modo == "popup":
            figura.suptitle(f"Vista Previa - {self._periodo_titulo()}", fontsize=tam_titulo + 2, fontweight='bold')
            figura.tight_layout(rect=[0.02, 0.06, 0.98, 0.93], h_pad=hpad)
        else:
            figura.tight_layout(rect=[0.02, 0.05, 0.98, 0.95], h_pad=hpad)

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

    def descargar_graficas_generales(self):
        if self.datos_ultima_ejecucion is None:
            messagebox.showwarning("Sin datos", "Primero genera gráficas o carga datos manuales.")
            return

        self._guardar_graficas_en_archivos_unicos(self.datos_ultima_ejecucion)
        messagebox.showinfo("Descarga completada", "Las gráficas se regeneraron y quedaron listas en la carpeta del proyecto.")

    def abrir_captura_manual(self):
        ventana_manual = tk.Toplevel(self.root)
        ventana_manual.title("Carga Manual de Datos")
        ventana_manual.geometry("1180x760")
        ventana_manual.configure(bg=APP_BG, padx=16, pady=16)
        ventana_manual.resizable(True, True)

        panel_manual = tk.Frame(ventana_manual, bg=CARD_BG, bd=1, relief="solid", highlightthickness=1, highlightbackground=BORDER_COLOR)
        panel_manual.pack(fill="both", expand=True)

        tk.Label(
            panel_manual,
            text="Ingresa los valores para cada barra",
            font=font_title,
            bg=CARD_BG,
            fg=TEXT_PRIMARY
        ).pack(anchor="w", padx=16, pady=(14, 4))

        tk.Label(
            panel_manual,
            text="Completa los valores y revisa la vista previa antes de generar.",
            font=font_interface,
            bg=CARD_BG,
            fg=TEXT_SECONDARY
        ).pack(anchor="w", padx=16, pady=(0, 10))

        tk.Label(panel_manual, text="Opciones de gráfica", font=("Poppins", 10, "bold") if self.usa_poppins else ("Segoe UI", 10, "bold"), bg=CARD_BG, fg=TEXT_PRIMARY).pack(anchor="w", padx=16, pady=(0, 4))
        frame_controles_superiores = tk.Frame(panel_manual, bg=CARD_BG)
        frame_controles_superiores.pack(fill="x", padx=16, pady=(0, 10))

        frame_opciones_manual_top = tk.Frame(frame_controles_superiores, bg=CARD_BG)
        frame_opciones_manual_top.pack(side="left")

        tk.Label(frame_opciones_manual_top, text="Tipo Desglose:", font=font_interface, bg=CARD_BG, fg=TEXT_PRIMARY).grid(row=0, column=0, padx=(0, 6), pady=3, sticky="w")
        combo_manual_desglose = ttk.Combobox(
            frame_opciones_manual_top,
            textvariable=self.tipo_grafico_desglose,
            values=["Barras", "Pastel"],
            state="readonly",
            width=12,
            style="Corporate.TCombobox"
        )
        combo_manual_desglose.grid(row=0, column=1, padx=(0, 14), pady=3, sticky="w")

        tk.Label(frame_opciones_manual_top, text="Tipo Estatus:", font=font_interface, bg=CARD_BG, fg=TEXT_PRIMARY).grid(row=0, column=2, padx=(0, 6), pady=3, sticky="w")
        combo_manual_estatus = ttk.Combobox(
            frame_opciones_manual_top,
            textvariable=self.tipo_grafico_estatus,
            values=["Barras", "Pastel"],
            state="readonly",
            width=12,
            style="Corporate.TCombobox"
        )
        combo_manual_estatus.grid(row=0, column=3, padx=(0, 14), pady=3, sticky="w")

        chk_manual_porcentaje = tk.Checkbutton(
            frame_opciones_manual_top,
            text="Mostrar porcentajes",
            variable=self.mostrar_porcentajes,
            bg=CARD_BG,
            fg=TEXT_PRIMARY,
            selectcolor=CARD_BG,
            activebackground=CARD_BG,
            command=lambda: None
        )
        chk_manual_porcentaje.grid(row=0, column=4, padx=(0, 0), pady=3, sticky="w")

        frame_superior_manual = tk.Frame(frame_controles_superiores, bg=CARD_BG)
        frame_superior_manual.pack(side="right")
        frame_acciones_manual_top = tk.Frame(frame_superior_manual, bg=CARD_BG)
        frame_acciones_manual_top.pack()

        frame_contenido = tk.Frame(panel_manual, bg=CARD_BG)
        frame_contenido.pack(fill="both", expand=True, padx=16, pady=(0, 10))

        frame_entries = tk.Frame(frame_contenido, bg=CARD_BG)
        frame_entries.pack(side="left", fill="y", padx=(0, 14))

        frame_preview_column = tk.Frame(frame_contenido, bg=CARD_BG)
        frame_preview_column.pack(side="left", fill="both", expand=True)

        entries = {}
        for i, categoria in enumerate(self.categorias_individuales):
            tk.Label(frame_entries, text=categoria + ":", font=font_interface, bg=CARD_BG, fg=TEXT_PRIMARY, anchor="w", width=24).grid(row=i, column=0, sticky="w", padx=(0, 8), pady=3)
            ent = tk.Entry(frame_entries, width=12)
            ent.insert(0, "0")
            ent.grid(row=i, column=1, pady=2)
            entries[categoria] = ent

        frame_preview_manual = tk.Frame(frame_preview_column, bg=CARD_BG, relief="solid", borderwidth=1, highlightthickness=1, highlightbackground=BORDER_COLOR)
        frame_preview_manual.pack(fill="both", expand=True)
        fig_manual = Figure(figsize=(7.0, 4.4), dpi=100, facecolor=CARD_BG)
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

        combo_manual_desglose.bind("<<ComboboxSelected>>", refrescar_preview_manual)
        combo_manual_estatus.bind("<<ComboboxSelected>>", refrescar_preview_manual)
        chk_manual_porcentaje.configure(command=refrescar_preview_manual)

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

        def descargar_graficas_manuales():
            try:
                individuales = leer_individuales_desde_entries()
                datos = self._construir_datos_desde_individuales(individuales)
                self.datos_ultima_ejecucion = datos
                self._actualizar_analisis_porcentual(datos)
                self.actualizar_vista_previa()
                self._guardar_graficas_en_archivos_unicos(datos)
                messagebox.showinfo("Descarga completada", "Las gráficas manuales se guardaron correctamente.")
            except ValueError as e:
                messagebox.showerror("Dato inválido", f"Revisa los valores ingresados.\n\nDetalle: {str(e)}")

        btn_generar_manual = tk.Button(frame_acciones_manual_top, text="Generar con datos manuales", command=aplicar_datos_manuales, bg=ACCENT_COLOR, fg="white", relief="flat", padx=14, pady=8, cursor="hand2", activebackground=ACCENT_HOVER)
        btn_generar_manual.grid(row=0, column=0, padx=(0, 8))
        self._aplicar_hover(btn_generar_manual, ACCENT_COLOR, ACCENT_HOVER)

        btn_descargar_manual = tk.Button(frame_acciones_manual_top, text="Guardar", command=descargar_graficas_manuales, bg=ACCENT_COLOR, fg="white", relief="flat", padx=14, pady=8, cursor="hand2", activebackground=ACCENT_HOVER)
        btn_descargar_manual.grid(row=0, column=1, padx=(0, 8))
        self._aplicar_hover(btn_descargar_manual, ACCENT_COLOR, ACCENT_HOVER)

        btn_cancelar_manual = tk.Button(frame_acciones_manual_top, text="Cancelar", command=ventana_manual.destroy, bg=MUTED_BUTTON, fg=TEXT_PRIMARY, relief="flat", padx=14, pady=8, cursor="hand2", activebackground=MUTED_BUTTON_HOVER)
        btn_cancelar_manual.grid(row=0, column=2)
        self._aplicar_hover(btn_cancelar_manual, MUTED_BUTTON, MUTED_BUTTON_HOVER)

        # Colores y opciones manuales se mantienen a la izquierda, debajo de las entradas.
        fila_base = len(self.categorias_individuales) + 1
        tk.Label(frame_entries, text="Colores de desglose", font=font_interface, bg=CARD_BG, fg=TEXT_PRIMARY, anchor="w").grid(row=fila_base, column=0, columnspan=2, sticky="w", pady=(8, 2))
        frame_colores_desglose_manual = tk.Frame(frame_entries, bg=CARD_BG)
        frame_colores_desglose_manual.grid(row=fila_base + 1, column=0, columnspan=2, sticky="w")

        btn_colores_desglose_manual = {}
        for i, categoria in enumerate(self.categorias_individuales):
            color_actual = self.colores_distribucion.get(categoria, '#D3D3D3')
            btn = tk.Button(
                frame_colores_desglose_manual,
                text=categoria,
                bg=color_actual,
                fg=self._color_texto_contrastante(color_actual),
                relief="solid",
                borderwidth=1,
                highlightthickness=1,
                highlightbackground="#94A3B8",
                padx=10,
                pady=6,
                font=(font_interface[0], 9),
                width=14,
                cursor="hand2"
            )
            btn.grid(row=i // 3, column=i % 3, sticky="ew", padx=2, pady=2)
            frame_colores_desglose_manual.grid_columnconfigure(i % 3, weight=1)
            self._aplicar_hover(btn, color_actual)
            btn_colores_desglose_manual[categoria] = btn

        tk.Label(frame_entries, text="Colores de estatus", font=font_interface, bg=CARD_BG, fg=TEXT_PRIMARY, anchor="w").grid(row=fila_base + 2, column=0, columnspan=2, sticky="w", pady=(8, 2))
        btn_colores_estatus_manual = {}
        frame_colores_estatus_top = tk.Frame(frame_entries, bg=CARD_BG)
        frame_colores_estatus_top.grid(row=fila_base + 3, column=0, columnspan=2, sticky="w")
        for i, categoria in enumerate(self.colores_resumen.keys()):
            color_actual = self.colores_resumen[categoria]
            btn = tk.Button(
                frame_colores_estatus_top,
                text=categoria,
                bg=color_actual,
                fg=self._color_texto_contrastante(color_actual),
                relief="solid",
                borderwidth=1,
                highlightthickness=1,
                highlightbackground="#94A3B8",
                padx=10,
                pady=6,
                font=(font_interface[0], 10),
                width=14,
                cursor="hand2"
            )
            btn.grid(row=0, column=i, padx=4, pady=2)
            self._aplicar_hover(btn, color_actual)
            btn_colores_estatus_manual[categoria] = btn

        def actualizar_botones_color_manual():
            for categoria, btn in btn_colores_estatus_manual.items():
                nuevo_color = self.colores_resumen.get(categoria, '#D3D3D3')
                btn.configure(bg=nuevo_color, fg=self._color_texto_contrastante(nuevo_color))
                self._aplicar_hover(btn, nuevo_color)

            for categoria, btn in btn_colores_desglose_manual.items():
                nuevo_color = self.colores_distribucion.get(categoria, '#D3D3D3')
                btn.configure(bg=nuevo_color, fg=self._color_texto_contrastante(nuevo_color))
                self._aplicar_hover(btn, nuevo_color)

        for categoria, btn in btn_colores_estatus_manual.items():
            btn.configure(command=lambda cat=categoria: (self.elegir_color_resumen(cat), actualizar_botones_color_manual(), refrescar_preview_manual()))

        for categoria, btn in btn_colores_desglose_manual.items():
            btn.configure(command=lambda cat=categoria: (self.elegir_color_distribucion(cat), actualizar_botones_color_manual(), refrescar_preview_manual()))

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
        self.preview_window.configure(bg=APP_BG)

        frame = tk.Frame(self.preview_window, bg=CARD_BG, bd=1, relief="solid", highlightthickness=1, highlightbackground=BORDER_COLOR)
        frame.pack(fill="both", expand=True, padx=12, pady=12)

        self.preview_figure_popup = Figure(figsize=(12, 7), dpi=100, facecolor=CARD_BG)
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

        if self.preview_figure_inline is not None and self.preview_canvas_inline is not None:
            self._render_preview_figure(self.preview_figure_inline, datos, modo="inline")
            self.preview_canvas_inline.draw()

        if self.preview_window is None or not self.preview_window.winfo_exists() or self.preview_figure_popup is None:
            return

        self._render_preview_figure(self.preview_figure_popup, datos, modo="popup")
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