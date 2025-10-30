# MPR - Gera Pedido de Compras

Windows desktop app (Tkinter) to generate purchase order files from spreadsheets, exporting XML/Text as required by the internal flow.

- Author: MPR Labs
- Platform: Windows (Tkinter)
- Language: Python 3.11+
- Build: PyInstaller
- Assets: `mprIco.ico`, `mprLabs4sml.png`

## Version

- Current version: 011025
- Latest binary: `dist/MPR-GeraPedidoCompras-v11025.exe`

## Requirements

Install dependencies:

```bash
pip install -r requirements.txt
```

Main dependencies:
- requests
- pandas
- openpyxl
- xlrd (for .xls files)
- pyinstaller (build only)

## Run from source

```bash
python MPR-GeraPedidoCompras.py
```

A Tkinter UI will open to select files and generate the order.

## Build executable (PyInstaller)

```bash
pyinstaller --noconfirm --onefile --windowed --icon mprIco.ico MPR-GeraPedidoCompras.py
```

The executable will be available in `dist/`.

## Logs

Logs are written to `logs/` with supplier/code/timestamp in the filename.

## Changelog

- 011025
  - Documentation (README) created/updated and requirements reviewed.
  - Binary published `MPR-GeraPedidoCompras-v11025.exe`.

## Support

If you face issues, attach the relevant file from `logs/` and describe the steps performed.
