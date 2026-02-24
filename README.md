# Generador de Vouchers WiFi

Aplicación en Python con interfaz gráfica (Tkinter) que extrae códigos WiFi desde un PDF y permite generar un PDF con diseño listo para imprimir o un Excel con los códigos organizados.

## Funcionalidades

- Lee códigos en formato `XXXXX-XXXXX` desde un PDF
- **Opción 1:** Genera un PDF con el diseño del voucher (4x4 = 16 vouchers por página, 500x500px cada uno) con líneas guía para recortar
- **Opción 2:** Genera un Excel con los códigos numerados y formato visual

## Requisitos

- Python 3.9+
- Poppler (necesario para `pdf2image`) → [Descargar para Windows](https://github.com/oschwartz10612/poppler-windows/releases)

## Instalación

```bash
pip install -r requirements.txt
```

Agrega Poppler al PATH o especifica su ruta en el script:
```python
poppler_path=r"C:\poppler\Library\bin"
```

## Uso

```bash
python generar_vouchers.py
```

1. Selecciona el **PDF con los códigos**
2. Selecciona el **PDF con el diseño del voucher** *(solo para generar PDF)*
3. Elige una acción:
   - `📄 Generar PDF con diseño` — crea el PDF con vouchers listos para imprimir
   - `📊 Generar Excel de códigos` — exporta solo los códigos a un archivo .xlsx

El archivo generado se guarda automáticamente en la carpeta **Descargas**.

## Dependencias

```
pdfplumber
pdf2image
Pillow
openpyxl
```