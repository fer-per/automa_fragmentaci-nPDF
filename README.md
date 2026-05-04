# Sistema de Automatización Archivística

Automatiza la extracción de documentos individuales desde un PDF escaneado de protocolos notariales, usando un inventario de registros en formato Excel (`.xlsx`).

---

## Instalación

```bash
pip install -r requirements.txt
```

---

## Configuración (`config.py`)

Antes de ejecutar, edita `config.py` y ajusta las rutas:

```python
EXCEL_PATH = BASE_DIR / "inventario.xlsx"   # ← tu archivo Excel real
PDF_PATH   = BASE_DIR / "documento.pdf"     # ← tu PDF escaneado real
```

---

## Uso Rápido

### 1. Generar un PDF de prueba sintético

```bash
python generate_test_pdf.py --pages 40
```

Esto crea `documento.pdf` con 40 páginas numeradas para probar el sistema sin el PDF real.

### 2. Ejecutar el proceso completo

```bash
python main.py
```

### 3. Modo simulación (sin escribir archivos)

```bash
python main.py --dry-run
```

### 4. Sin detección de saltos de secuencia

```bash
python main.py --no-gap-check
```

---

## Estructura de Salida

```
output/
└── PORTUGAL, Cesar/
    └── 1567 - Protocolo N° 1/
        └── Obligacion/
            └── Julio/
                └── Diego de Aramburu/
                    └── Antonio de Oviedo.pdf
```

---

## Logs

| Archivo | Contenido |
|---|---|
| `logs/process.log` | Log completo de cada operación |
| `logs/pendientes.csv` | Registros no procesados con motivo |

---

## Notación de Folios

| Formato | Ejemplo | Páginas PDF |
|---|---|---|
| Rango completo | `1r-1v` | 1, 2 |
| Multi-hoja | `4r-6v` | 7, 8, 9, 10, 11, 12 |
| Solo recto | `7r` | 13 |
| Cruzado | `7v-8r` | 14, 15 |
| Rango vuelta | `8v-12r` | 16–23 |

---

## Estructura del Proyecto

```
fracmen_auto/
├── main.py                  # Punto de entrada
├── config.py                # Configuración centralizada
├── generate_test_pdf.py     # Genera PDF sintético de prueba
├── requirements.txt
├── modules/
│   ├── excel_reader.py      # Lectura del Excel .xlsx
│   ├── folio_parser.py      # Algoritmo r/v → páginas PDF
│   ├── pdf_extractor.py     # Extracción y escritura de PDFs
│   ├── folder_builder.py    # Construcción de estructura de carpetas
│   └── validator.py         # Validación de registros
├── output/                  # PDFs generados (se crea automáticamente)
└── logs/                    # Logs y pendientes (se crea automáticamente)
```
