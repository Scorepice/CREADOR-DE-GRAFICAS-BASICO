# Sistema Presentacion - Generador de graficas RDM

Proyecto en Python con interfaz grafica para generar imagenes PNG de indicadores RDM a partir de un archivo Excel remoto (SharePoint) o con captura manual.

## Objetivo

Automatizar la creacion de graficas para presentaciones operativas:
- Desglose por proceso
- Estatus general (Procesadas vs En Proceso)

## Estructura del proyecto

- `interfaz_rdm_v2.py`: Aplicacion principal (version recomendada).
- `crear_grafica_personalizada.py`: Modulo reutilizable para crear graficas personalizadas desde la UI.
- `interfaz_rdm.py`: Version anterior de la interfaz (legacy).
- `generar_graficos.py`: Script de consola sin interfaz (legacy/alternativo).
- `DOCUMENTACION_PROYECTO.md`: Nota corta de limpieza y ejecucion.
- `.env`: Variables de entorno (por ejemplo URL de SharePoint).

## Flujo de ejecucion

### 1) Punto de entrada principal

Abrir y ejecutar:

- `interfaz_rdm_v2.py`

Este archivo levanta la ventana principal Tkinter y coordina todo el flujo:
- Descarga de Excel remoto
- Procesamiento de hoja y columnas
- Render de graficas
- Exportacion a PNG
- Vista previa y carga manual

### 2) Modulo de apoyo

`interfaz_rdm_v2.py` importa `VentanaGraficaPersonalizada` desde `crear_grafica_personalizada.py`.

### 3) Archivos alternativos

- `interfaz_rdm.py`: interfaz v1 funcional pero antigua.
- `generar_graficos.py`: ejecucion por consola, util para pruebas rapidas.

## Requisitos

Python recomendado: 3.10+

Paquetes:
- python-dotenv
- requests
- pandas
- matplotlib
- openpyxl

Instalacion sugerida:

```powershell
pip install python-dotenv requests pandas matplotlib openpyxl
```

## Configuracion

Crear o validar el archivo `.env` en la raiz del proyecto con:

```env
URL_SHAREPOINT=https://tu-enlace-sharepoint?download=1
```

Notas:
- `interfaz_rdm_v2.py` usa `URL_SHAREPOINT` desde `.env`.
- Si la URL no existe o expira, fallara la descarga del Excel.

## Uso rapido

1. Abrir `interfaz_rdm_v2.py`.
2. Ejecutar con Play en VS Code.
3. Opcional: ajustar mes/anio o usar deteccion automatica del Excel.
4. Elegir tipo de grafica (Barras o Pastel).
5. Personalizar colores.
6. Pulsar "Generar graficas".
7. Exportar/guardar PNG desde los botones de descarga.

## Portable por USB

### En la PC donde vas a preparar el paquete

1. Verifica que exista el archivo `.env` en la raiz del proyecto con `URL_SHAREPOINT`.
2. Ejecuta `build_portable.bat` para generar `dist\SistemaPresentacionRDM.exe`.
3. Ejecuta `preparar_paquete.bat` para crear `release\SistemaPresentacionRDM\` y el ZIP `release\SistemaPresentacionRDM.zip`.
4. Copia ese ZIP al USB.

### En la otra PC

1. Descomprime `SistemaPresentacionRDM.zip` en una carpeta local o directamente en el USB.
2. Ejecuta `Ejecutar_RDM.bat` o `SistemaPresentacionRDM.exe`.
3. Mantén el archivo `.env` junto al ejecutable si necesitas que la app lea el enlace de SharePoint.

Notas:
- No hace falta instalar Python en la PC destino si usas el `.exe` portable.
- Si la pantalla es baja o la resolución es reducida, la ventana principal ahora permite desplazamiento vertical para acceder a la parte inferior.

## Logica de datos (resumen)

- Se lee la hoja 2 del Excel (`sheet_name=1`) con encabezado en fila 2 (`header=1`).
- Se limpia nombre de columnas y filas vacias.
- Se usa la ultima fila como resumen numerico.
- Se agrupan columnas en:
  - En Proceso
  - Procesadas
- Se calcula total como suma de ambos grupos.

## Salidas

Se generan archivos PNG en la carpeta del proyecto con nombres unicos por periodo y timestamp, por ejemplo:
- `grafico_distribucion_rdm_<periodo>_<timestamp>.png`
- `grafico_estatus_rdm_<periodo>_<timestamp>.png`

## Mantenimiento recomendado

1. Mantener como base:
- `interfaz_rdm_v2.py`
- `crear_grafica_personalizada.py`

2. Archivar o eliminar si ya no se usan:
- `interfaz_rdm.py`
- `generar_graficos.py`

3. Validar periodicamente:
- Vigencia del enlace de SharePoint
- Nombres reales de columnas en Excel
- Dependencias instaladas en el entorno

## Solucion de problemas

### Error al descargar archivo

Causas comunes:
- URL invalida o expirada
- Falta de internet o bloqueo de red

Acciones:
- Probar la URL en navegador
- Actualizar `URL_SHAREPOINT` en `.env`

### Error al leer Excel

Causas comunes:
- Cambio de estructura de la hoja
- Encabezados diferentes a los esperados

Acciones:
- Revisar que la informacion este en hoja 2
- Validar nombres de columnas contra los usados en codigo

### Graficas vacias o total en 0

Causas comunes:
- Ultima fila sin numericos
- Columnas de procesos no coinciden

Acciones:
- Confirmar fila de totales en el Excel
- Revisar columnas agrupadas en la funcion de procesamiento

## Evolucion sugerida

- Crear `requirements.txt` para versionado de dependencias.
- Agregar logging basico para diagnostico.
- Separar logica de negocio y UI en modulos para pruebas unitarias.
- Crear carpeta `output/` para separar codigo y archivos PNG generados.
