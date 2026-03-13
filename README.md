# RVTools Infrastructure Analyzer

Script automatizado de análisis de reportes RVTools que detecta anomalías en la infraestructura de virtualización.

## 📋 Descripción

Este proyecto procesa reportes de RVTools (vHealth y vPartition) para analizar el estado de tu infraestructura virtualizada. Detecta automáticamente problemas como:

- ⚠️ Bajo espacio en disco
- 🔧 VMware Tools desactualizadas o no instaladas
- 💾 CDROMs conectados
- 👻 VMs zombies
- Otros problemas de configuración

El script compara el estado actual con el estado anterior para identificar nuevos problemas e **genera reportes en Excel y Markdown** para fácil consulta.

## 🛠️ Requisitos Previos

- **Windows 10/11** (el sistema operativo)
- **Python 3.12+**
- **uv** - Gestor de paquetes Python moderno (opcional pero recomendado)

## 📦 Instalación

### Opción 1: Usando `uv` (Recomendado)

Si no tienes `uv` instalado, instálalo primero con:

```powershell
winget install --id=astral-sh.uv -e
```

Luego, sincroniza las dependencias del proyecto:

```powershell
uv sync
```

### Opción 2: Sin `uv`

Si prefieres usar `pip` directamente:

```powershell
python -m pip install -r requirements.txt
```

*Nota: Se recomienda usar `uv` para mejor compatibilidad y automatización.*

## 🚀 Uso

### Con `uv`

```powershell
uv run main.py
```

### Sin `uv`

```powershell
python main.py
```

## 📊 Entrada

El script requiere un archivo Excel con reportes de RVTools que contenga las siguientes hojas:

- **vHealth** - Información de salud de las máquinas virtuales
- **vPartition** - Información de particiones y almacenamiento

## 📁 Salida

El script genera reportes en la carpeta `rvtools_reports/`:

- **rvtools_report_YYYYMMDD.xlsx** - Reporte en Excel con detalles de anomalías
- **rvtools_previous_state.json** - Historial del estado anterior (para comparaciones futuras)
- **current_issues.md** - Resumen de problemas actuales en formato Markdown

## 📝 Estructura del Proyecto

``` text
Scripts/
├── main.py                          # Script principal
├── pyproject.toml                   # Configuración del proyecto
├── README.md                        # Este archivo
└── rvtools_reports/                # Carpeta de salida
    ├── rvtools_report_*.xlsx        # Reportes generados
    ├── rvtools_previous_state.json  # Historial de estado
    └── current_issues.md            # Resumen de problemas
```

## 🔧 Configuración

Modifica el archivo `main.py` para ajustar:

- Archivos de entrada personalizados
- Configuración de anomalías a detectar
- Path del reporte Excel

## 📚 Dependencias

- **pandas** - Análisis y manipulación de datos
- **openpyxl** - Lectura/escritura de archivos Excel
- **numpy** - Operaciones numéricas

## 💡 Ejemplo de Flujo de Trabajo

```powershell
# 1. Instalar dependencias
uv sync

# 2. Colocar el reporte RVTools en el proyecto

# 3. Ejecutar el análisis
uv run main.py

# 4. Revisar los reportes generados en rvtools_reports/
```

## 🐛 Solución de Problemas

### **Error: "FileNotFoundError"**

- Verifica que el archivo Excel con reportes RVTools esté en el directorio del proyecto

### **Error: "No module named 'pandas'"**

- Ejecuta `uv sync` o `pip install -r requirements.txt`

### **Error: "Permission denied"**

- Asegúrate de tener permisos de escritura en la carpeta del proyecto

## 📧 Notas

- Los reportes se actualizan diariamente basándose en la fecha actual
- El historial de estado permite seguimiento de cambios a lo largo del tiempo
- Se recomienda ejecutar este script regularmente para monitoreo continuo

---

**Última actualización**: Marzo 2026
