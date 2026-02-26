# PDF a Excel

Convierte archivos PDF a Excel (.xlsx) desde la línea de comandos. Auto-detecta tablas y texto, con soporte batch.

## Requisitos

- **Python 3.10+**
- **Java JRE/JDK** (necesario para `tabula-py`)

```bash
python -m venv .venv
.venv\Scripts\Activate.ps1   # Windows
# source .venv/bin/activate  # Linux/Mac
pip install -r requirements.txt
```

---

## Uso

```bash
# Modo automático (detecta tablas, si no hay extrae texto)
python pdf_a_excel.py documento.pdf

# Forzar extracción de tablas
python pdf_a_excel.py factura.pdf -o factura.xlsx --modo tablas

# Forzar extracción de texto
python pdf_a_excel.py libro.pdf -o libro.xlsx --modo texto

# Solo páginas específicas
python pdf_a_excel.py reporte.pdf --paginas 1,3,5-10

# Tablas con bordes visibles
python pdf_a_excel.py tabla.pdf --modo tablas --lattice

# Texto separado por ";" en columnas
python pdf_a_excel.py datos.pdf --modo texto --separador ";"

# Procesar todos los PDFs de una carpeta
python -u pdf_a_excel.py pdfs/ --batch -o salida/

# Modo verbose
python pdf_a_excel.py doc.pdf -v
```

## Argumentos

| Argumento | Descripción |
|---|---|
| `entrada` | Ruta al PDF (o carpeta con `--batch`) |
| `-o, --output` | Ruta del Excel de salida |
| `-m, --modo` | `auto`, `tablas`, `texto` (default: `auto`) |
| `-p, --paginas` | `all`, `1`, `1,3,5`, `1-5` (default: `all`) |
| `--lattice` | Tablas con bordes visibles |
| `--stream` | Tablas sin bordes (por espacios) |
| `--multiple-tablas` | Detectar múltiples tablas por página |
| `--separar-hojas` | Guardar cada tabla en una hoja separada (default: todo en una) |
| `--separador` | Separador para dividir texto en columnas |
| `--sin-vacias` | Omitir líneas vacías |
| `--modo-texto` | `linea` o `pagina` (default: `linea`) |
| `--batch` | Procesar todos los PDFs de una carpeta |
| `--encoding` | Codificación (default: `utf-8`) |
| `-v, --verbose` | Información detallada |

## Ejemplos

```bash
# Factura con tablas
python pdf_a_excel.py factura.pdf -o factura.xlsx --modo tablas --lattice

# Reporte con texto tabulado por ";"
python pdf_a_excel.py reporte.pdf --modo texto --separador ";"

# Batch de 50 PDFs
python -u pdf_a_excel.py mis_pdfs/ --batch -o resultados/

# Solo las primeras 5 páginas
python pdf_a_excel.py grande.pdf --paginas 1-5
```

## Cómo funciona

1. **Modo auto** (por defecto): intenta extraer tablas con `tabula-py`. Si falla, prueba con `pdfplumber`. Si no encuentra tablas, extrae texto línea por línea.
2. **Modo tablas**: fuerza extracción de tablas (tabula → pdfplumber como fallback).
3. **Modo texto**: extrae texto directamente con `pdfplumber`.

## Notas

- `tabula-py` requiere Java y es ideal para PDFs con tablas bien definidas.
- `pdfplumber` no requiere Java y sirve como fallback para tablas y para extracción de texto.
- Para PDFs escaneados (imágenes) se necesitaría OCR adicional (no incluido).
