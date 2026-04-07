import os
from datetime import datetime
import tkinter as tk
from tkinter import ttk, messagebox, filedialog, colorchooser

import matplotlib.ticker as ticker
from matplotlib.figure import Figure
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg


class VentanaGraficaPersonalizada:
    def __init__(self, parent, font_interface, font_title):
        self.parent = parent
        self.font_interface = font_interface
        self.font_title = font_title

        self.ventana = tk.Toplevel(parent)
        self.ventana.title("Crear Grafica Personalizada")
        self.ventana.geometry("1100x720")
        self.ventana.configure(bg="white", padx=14, pady=14)

        self.tipo_grafico = tk.StringVar(value="Barras")
        self.titulo_grafica = tk.StringVar(value="Grafica Personalizada")
        self.mostrar_porcentajes = tk.BooleanVar(value=True)
        self.filas = []
        self.paleta_default = [
            "#FF9999", "#66B2FF", "#99FF99", "#D3D3D3", "#FFCC99",
            "#FFD700", "#C6C6F6", "#FFB6C1", "#80C680", "#F59E0B"
        ]

        self._construir_ui()
        self._agregar_fila("Campo 1", "10")
        self._agregar_fila("Campo 2", "8")
        self._agregar_fila("Campo 3", "6")
        self.actualizar_preview()

    def _construir_ui(self):
        tk.Label(
            self.ventana,
            text="Crear Grafica Personalizada",
            font=self.font_title,
            bg="white",
            fg="black"
        ).pack(pady=(0, 10))

        frame_superior = tk.Frame(self.ventana, bg="white")
        frame_superior.pack(fill="x", pady=(0, 10))

        tk.Label(frame_superior, text="Tipo de grafica:", font=self.font_interface, bg="white", fg="black").pack(side="left", padx=(0, 6))
        combo = ttk.Combobox(
            frame_superior,
            textvariable=self.tipo_grafico,
            values=["Barras", "Pastel"],
            state="readonly",
            width=12
        )
        combo.pack(side="left")
        combo.bind("<<ComboboxSelected>>", lambda _e: self.actualizar_preview())

        chk_porcentaje = tk.Checkbutton(
            frame_superior,
            text="Mostrar porcentajes",
            variable=self.mostrar_porcentajes,
            bg="white",
            fg="black",
            selectcolor="white",
            activebackground="white",
            command=self.actualizar_preview
        )
        chk_porcentaje.pack(side="left", padx=(14, 0))

        tk.Label(frame_superior, text="Titulo:", font=self.font_interface, bg="white", fg="black").pack(side="left", padx=(16, 6))
        ent_titulo = tk.Entry(frame_superior, textvariable=self.titulo_grafica, width=32)
        ent_titulo.pack(side="left")
        ent_titulo.bind("<KeyRelease>", lambda _e: self.actualizar_preview())

        frame_principal = tk.Frame(self.ventana, bg="white")
        frame_principal.pack(fill="both", expand=True)

        self.frame_campos = tk.Frame(frame_principal, bg="white")
        self.frame_campos.pack(side="left", fill="y", padx=(0, 14))

        tk.Label(self.frame_campos, text="Nombre", font=self.font_interface, bg="white", fg="black", width=20, anchor="w").grid(row=0, column=0, padx=(0, 8), pady=(0, 6))
        tk.Label(self.frame_campos, text="Valor", font=self.font_interface, bg="white", fg="black", width=9, anchor="w").grid(row=0, column=1, padx=(0, 8), pady=(0, 6))
        tk.Label(self.frame_campos, text="Color", font=self.font_interface, bg="white", fg="black", width=8, anchor="w").grid(row=0, column=2, padx=(0, 8), pady=(0, 6))

        self.frame_rows = tk.Frame(self.frame_campos, bg="white")
        self.frame_rows.grid(row=1, column=0, columnspan=3)

        frame_row_buttons = tk.Frame(self.frame_campos, bg="white")
        frame_row_buttons.grid(row=2, column=0, columnspan=3, pady=(8, 0), sticky="w")

        tk.Button(
            frame_row_buttons,
            text="Agregar campo",
            command=lambda: self._agregar_fila(),
            bg="#E5E7EB",
            fg="black",
            relief="flat",
            padx=10,
            pady=4
        ).pack(side="left", padx=(0, 6))

        tk.Button(
            frame_row_buttons,
            text="Quitar ultimo",
            command=self._quitar_fila,
            bg="#E5E7EB",
            fg="black",
            relief="flat",
            padx=10,
            pady=4
        ).pack(side="left")

        self.frame_preview = tk.Frame(frame_principal, bg="white", relief="sunken", borderwidth=1)
        self.frame_preview.pack(side="left", fill="both", expand=True)

        self.figure = Figure(figsize=(7.8, 5.2), dpi=100, facecolor="white")
        self.canvas = FigureCanvasTkAgg(self.figure, master=self.frame_preview)
        self.canvas.get_tk_widget().pack(fill="both", expand=True, padx=8, pady=8)

        frame_actions = tk.Frame(self.ventana, bg="white")
        frame_actions.pack(fill="x", pady=(10, 0))

        tk.Button(
            frame_actions,
            text="Guardar grafica",
            command=self.guardar_grafica,
            bg="#2563EB",
            fg="white",
            relief="flat",
            padx=12,
            pady=6
        ).pack(side="left")

        tk.Button(
            frame_actions,
            text="Cerrar",
            command=self.ventana.destroy,
            bg="#E5E7EB",
            fg="black",
            relief="flat",
            padx=12,
            pady=6
        ).pack(side="left", padx=(8, 0))

    def _agregar_fila(self, nombre="", valor="0"):
        idx = len(self.filas)
        frame = tk.Frame(self.frame_rows, bg="white")
        frame.grid(row=idx, column=0, sticky="w", pady=2)

        ent_nombre = tk.Entry(frame, width=24)
        ent_nombre.insert(0, nombre if nombre else f"Campo {idx + 1}")
        ent_nombre.grid(row=0, column=0, padx=(0, 8))

        ent_valor = tk.Entry(frame, width=12)
        ent_valor.insert(0, valor)
        ent_valor.grid(row=0, column=1)

        color_inicial = self.paleta_default[idx % len(self.paleta_default)]
        color_var = tk.StringVar(value=color_inicial)

        btn_color = tk.Button(
            frame,
            text="Color",
            bg=color_inicial,
            fg="black",
            width=8,
            relief="flat",
            borderwidth=1,
            command=lambda v=color_var, b=None: self._elegir_color_fila(v, b)
        )
        btn_color.grid(row=0, column=2, padx=(8, 0))
        btn_color.configure(command=lambda v=color_var, b=btn_color: self._elegir_color_fila(v, b))

        ent_nombre.bind("<KeyRelease>", lambda _e: self.actualizar_preview())
        ent_valor.bind("<KeyRelease>", lambda _e: self.actualizar_preview())

        self.filas.append((frame, ent_nombre, ent_valor, color_var, btn_color))
        self.actualizar_preview()

    def _quitar_fila(self):
        if len(self.filas) <= 1:
            return
        frame, _, _, _, _ = self.filas.pop()
        frame.destroy()
        self.actualizar_preview()

    def _elegir_color_fila(self, color_var, btn_color):
        color = colorchooser.askcolor(title="Selecciona un color")[1]
        if not color:
            return
        color_var.set(color)
        btn_color.configure(bg=color)
        self.actualizar_preview()

    def _leer_datos(self):
        nombres = []
        valores = []
        colores = []
        for _, ent_nombre, ent_valor, color_var, _ in self.filas:
            nombre = ent_nombre.get().strip()
            if not nombre:
                continue

            texto = ent_valor.get().strip().replace(",", ".")
            if not texto:
                valor = 0.0
            else:
                try:
                    valor = float(texto)
                except ValueError:
                    valor = 0.0

            if valor < 0:
                valor = 0.0

            nombres.append(nombre)
            valores.append(valor)
            colores.append(color_var.get())

        return nombres, valores, colores

    def _titulo_actual(self):
        titulo = self.titulo_grafica.get().strip()
        if not titulo:
            titulo = "Grafica"
        return titulo

    def _porcentajes(self, valores):
        total = sum(valores)
        if total <= 0:
            return [0 for _ in valores]
        return [(v / total) * 100 for v in valores]

    def actualizar_preview(self):
        nombres, valores, colores = self._leer_datos()
        porcentajes = self._porcentajes(valores)
        self.figure.clear()
        ax = self.figure.add_subplot(1, 1, 1)

        if not nombres:
            ax.axis("off")
            ax.text(0.5, 0.5, "Agrega al menos un campo", ha="center", va="center", fontsize=12, fontweight="bold")
            self.canvas.draw()
            return

        if self.tipo_grafico.get() == "Pastel":
            total = sum(valores)
            if total <= 0:
                ax.axis("off")
                ax.text(0.5, 0.5, "Los valores deben ser mayores que 0", ha="center", va="center", fontsize=12, fontweight="bold")
            else:
                if self.mostrar_porcentajes.get():
                    labels = [f"{n}\n{v:g} ({porcentajes[i]:.1f}%)" for i, (n, v) in enumerate(zip(nombres, valores))]
                else:
                    labels = [f"{n}\n{v:g}" for n, v in zip(nombres, valores)]
                ax.pie(
                    valores,
                    labels=labels,
                    colors=colores,
                    startangle=90,
                    wedgeprops={"edgecolor": "white"},
                    textprops={"fontsize": 9}
                )
                ax.set_title(self._titulo_actual(), fontsize=12, fontweight="bold")
        else:
            x = list(range(len(nombres)))
            barras = ax.bar(x, valores, color=colores, edgecolor="black")
            ax.set_xticks(x)
            ax.set_xticklabels(nombres, rotation=25, ha="right", fontsize=9)
            ax.yaxis.set_major_locator(ticker.MaxNLocator(integer=True))
            for i, b in enumerate(barras):
                etiqueta = f"{valores[i]:g}"
                if self.mostrar_porcentajes.get():
                    etiqueta = f"{etiqueta} ({porcentajes[i]:.1f}%)"
                ax.text(b.get_x() + b.get_width() / 2, b.get_height() + 0.1, etiqueta, ha="center", va="bottom", fontsize=9, fontweight="bold")
            ax.set_title(self._titulo_actual(), fontsize=12, fontweight="bold")

        self.figure.patch.set_facecolor("white")
        self.figure.tight_layout()
        self.canvas.draw()

    def guardar_grafica(self):
        nombres, valores, _ = self._leer_datos()
        if not nombres:
            messagebox.showwarning("Sin datos", "Agrega al menos un campo para guardar la grafica.")
            return

        if self.tipo_grafico.get() == "Pastel" and sum(valores) <= 0:
            messagebox.showwarning("Valores invalidos", "Para grafica de pastel, los valores deben sumar mas de 0.")
            return

        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        nombre_sugerido = f"grafica_personalizada_{timestamp}.png"
        ruta = filedialog.asksaveasfilename(
            title="Guardar grafica personalizada",
            defaultextension=".png",
            initialfile=nombre_sugerido,
            filetypes=[("Imagen PNG", "*.png")]
        )
        if not ruta:
            return

        self.figure.savefig(ruta, dpi=300, facecolor="white", transparent=False)
        messagebox.showinfo("Guardado", "La grafica personalizada se guardo correctamente.")
