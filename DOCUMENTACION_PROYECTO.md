# Documentacion del proyecto

## Que archivo se ejecuta al dar Play

En VS Code, el boton Play de Python ejecuta el archivo que tengas abierto en ese momento.

- Si abres `interfaz_rdm_v2.py` y das Play, se ejecuta la app principal (UI v2).
- Si abres `interfaz_rdm.py` y das Play, se ejecuta la app antigua (UI v1).
- Si abres `generar_graficos.py` y das Play, se ejecuta el script por consola sin interfaz.
- `crear_grafica_personalizada.py` esta pensado como modulo de apoyo para v2.

## Dependencias entre archivos

- `interfaz_rdm_v2.py` importa `VentanaGraficaPersonalizada` desde `crear_grafica_personalizada.py`.
- `interfaz_rdm.py` no importa archivos locales del proyecto.
- `generar_graficos.py` no importa archivos locales del proyecto.

## Recomendacion para limpiar

Si tu flujo actual es usar solo la interfaz moderna (v2):

1. Mantener:
- `interfaz_rdm_v2.py`
- `crear_grafica_personalizada.py`
- `.env` (si sigues leyendo la URL desde variables de entorno)

2. Opcional archivar o borrar (si ya no los usas):
- `interfaz_rdm.py` (version anterior)
- `generar_graficos.py` (version consola)

## Antes de borrar

1. Ejecuta `interfaz_rdm_v2.py`.
2. Prueba botones clave: Generar graficas, Vista previa, Crear grafica.
3. Si todo funciona, entonces puedes borrar los archivos legacy.

## Dependencias Python

Paquetes usados en el proyecto:
- `python-dotenv`
- `requests`
- `pandas`
- `matplotlib`
- `openpyxl` (normalmente requerido por `pandas` para leer xlsx)

Ejemplo de instalacion:

```powershell
pip install python-dotenv requests pandas matplotlib openpyxl
```

## Generar un EXE portable

1. Ejecuta `build_portable.bat` desde la raiz del proyecto.
2. El ejecutable queda en `dist\SistemaPresentacionRDM.exe`.
3. Para moverlo a otra PC, copia el `.exe` junto con `.env`.

## Instalar en otra PC mediante USB

### Preparar en la PC de origen

1. Asegura que el archivo `.env` exista en la raiz del proyecto y tenga `URL_SHAREPOINT` configurado.
2. Ejecuta `build_portable.bat` para generar el ejecutable portable.
3. Ejecuta `preparar_paquete.bat` para crear la carpeta `release\SistemaPresentacionRDM` y el archivo `release\SistemaPresentacionRDM.zip`.
4. Copia el ZIP al USB.

### Instalar en la PC destino

1. Copia o descomprime `SistemaPresentacionRDM.zip` en una carpeta local de la PC destino.
2. Abre `Ejecutar_RDM.bat` o ejecuta `SistemaPresentacionRDM.exe`.
3. Conserva `.env` junto al `.exe` si la app debe leer el enlace de SharePoint.
4. Si la ventana se ve recortada por la resolución del monitor, usa el scroll vertical de la pantalla principal para bajar hasta la sección inferior.

### Ventaja del flujo portable

- No requiere instalar Python en la PC destino.
- El archivo ZIP ya incluye el ejecutable y el lanzador.
- El scroll vertical evita perder acceso a la parte inferior en monitores de escritorio más bajos.

## Generar instalador de Windows

1. Ejecuta `crear_instalador.bat`.
2. Este paso requiere Inno Setup Compiler (`ISCC.exe`) instalado en Windows.
3. El script usa [instalador_sistema_presentacion.iss](instalador_sistema_presentacion.iss) para crear el setup.

Notas:
- Si Inno Setup no está instalado, el instalador no se puede compilar en esta PC.
- El paquete portable sí quedó generado en `dist\SistemaPresentacionRDM.exe`.

## Preparar paquete para otra PC

1. Ejecuta `preparar_paquete.bat`.
2. Se crea una carpeta `release\SistemaPresentacionRDM` con el `.exe`, `.env` y un lanzador.
3. También se genera `release\SistemaPresentacionRDM.zip` para compartir o copiar.
