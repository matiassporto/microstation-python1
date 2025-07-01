# Microstation-Python1

This repository contains a Python script that interacts with Bentley MicroStation through its Python API. The script extracts formatted text labels from a DGN model, filters out those containing the word "SOCKET", and exports the results into an Excel file.

## Features

- Connects to the active MicroStation DGN model.
- Iterates through all graphic elements, extracting text from "Text Node" elements.
- Filters out text labels containing "SOCKET".
- Splits the remaining labels into parts and writes each part to a separate Excel cell in a new row.
- Saves the results in an Excel file named `etichette_formattate.xlsx` inside an `Excel` folder.

## Prerequisites

- Bentley MicroStation with Python integration enabled.
- The following MicroStation Python modules:
  - `MSPyBentley`
  - `MSPyBentleyGeom`
  - `MSPyECObjects`
  - `MSPyDgnPlatform`
  - `MSPyDgnView`
  - `MSPyMstnPlatform`
- Python libraries:
  - `openpyxl`

## How it works

1. **Setup**: The script creates an `Excel` folder (if it doesn't exist) one directory up from the script location.
2. **Excel Preparation**: Initializes a new Excel workbook and worksheet.
3. **Element Processing**: For each graphic element in the active DGN model:
    - If the element is a Text Node, extract its text.
    - If the text does not contain "SOCKET", split it into parts and write to Excel.
4. **Save**: The Excel file is saved in the `Excel` folder.

## Usage

1. Place the script inside your MicroStation Python environment.
2. Ensure all required Bentley and Python modules are available.
3. Run the script within the MicroStation environment.

The resulting file will be located at:

```
<repository-root>/Excel/etichette_formattate.xlsx
```

## Code Structure

- **aggiungi_etichetta(stringa, ws, riga_corrente)**: Adds a label to the worksheet if it doesn't contain "SOCKET".
- **salva_excel(nome_file)**: Saves the current workbook to the specified file.
- **main()**: Orchestrates the scanning, extraction, and saving process.

## Customization

- To change the Excel output path or filename, adjust the `cartella_excel` and `nome_file` variables.
- To filter on a different word, modify the check in `aggiungi_etichetta`.

## License

Specify your license here.

## Author

Matias Porto
