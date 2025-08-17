# Script-SGD: Document Processor

A Python script that automates the processing of documents from a local HTML file. It extracts document information, downloads the main PDF and any associated attachments, generates a routing slip (Hoja Ruta) for specific documents, and creates a consolidated Excel report.

## Features

-   **HTML Parsing**: Parses a local HTML file (`recibidos.html`) to extract a list of documents.
-   **File Downloading**: For each document, it downloads the main PDF and any associated attachments ("anexos").
-   **PDF Generation**: Generates a "Hoja Ruta" (routing slip) in PDF format for documents that have been "reasignado" (reassigned) or "informado" (informed).
-   **Structured Output**: Creates a well-organized output with a separate folder for each document, containing all its related files.
-   **Excel Reporting**: Generates a summary Excel report (`.xlsx`) with details and logs for all processed documents.

## Prerequisites

-   Python 3.x
-   pip (Python package installer)

## Installation

1.  **Clone the repository:**
    ```bash
    git clone https://github.com/tu-usuario/Script-SGD.git
    cd Script-SGD
    ```

2.  **Create and activate a virtual environment:**
    ```bash
    # Create the virtual environment
    python -m venv .venv

    # Activate the environment
    # On Windows:
    .venv\Scripts\Activate.ps1
    # On macOS/Linux:
    source .venv/bin/activate
    ```

3.  **Install the required dependencies:**
    ```bash
    pip install -r requirements.txt
    ```

## Configuration

**IMPORTANT:** Before running the script, you must configure the file paths in `main_recibidos_optimizado.py`.

1.  Open the `main_recibidos_optimizado.py` file in a text editor.

2.  Modify the following variables with the absolute paths on your system:

    -   `ruta_html`: Set this to the full path of your `recibidos.html` file.
        ```python
        # Example:
        ruta_html = r"C:\Users\YourUser\Documents\SGD\recibidos.html"
        ```

    -   `carpeta_destino`: Set this to the full path of the folder where you want to save the processed documents and the final report.
        ```python
        # Example:
        carpeta_destino = r"C:\Users\YourUser\Documents\SGD\Processed"
        ```

## Usage

Once you have configured the paths, run the script from your terminal:

```bash
python main_recibidos_optimizado.py
```

The script will show a progress bar as it processes the documents.

## Output

The script will generate the following:

-   A main output folder at the path specified in `carpeta_destino`.
-   Inside the output folder, a subfolder for each processed document (e.g., `01_MEMORANDO-001-2023`). Each of these folders will contain:
    -   The main PDF of the document.
    -   Any associated attachments ("anexos").
    -   A `Hoja Ruta_...pdf` file if the document had a tracking history.
-   An Excel file named `doc._recibidos_extraidos_YYYY-MM-DD.xlsx` in the main output folder, containing a summary of all processed documents and logs.

## Dependencies

This project relies on the following main libraries:

-   [Beautiful Soup](https://www.crummy.com/software/BeautifulSoup/): For parsing HTML.
-   [Pandas](https://pandas.pydata.org/): For data manipulation and creating the Excel report.
-   [ReportLab](https://www.reportlab.com/): For generating PDFs.
-   [tqdm](https://github.com/tqdm/tqdm): For displaying progress bars.
