# Auto PDF Filling

Automated PDF form filling tool with GUI support. Generate placeholders, define field mappings, and fill PDFs in bulk or through an intuitive interface.

## Features

✨ **4 Main Components:**

1. **Placeholder Generator** - Automatically generate placeholder templates and save them as JSON files
2. **GUI Field Mapper** - Interactive GUI to select PDF files and define field locations and values
3. **Batch PDF Filler** - Command-line tool to fill PDF files using JSON configuration files
4. **GUI PDF Filler** - User-friendly interface for filling PDFs without command-line knowledge

## Installation

```bash
# Clone the repository
git clone https://github.com/Soroushmehr110/autopdffilling.git
cd autopdffilling

# Install dependencies
pip install -r requirements.txt
```

## Usage

### 1. Generate Placeholder Template
```bash
python placeholder_generator.py
```
Creates a JSON file with field placeholders ready to be filled with your data.

### 2. GUI Field Mapper
```bash
python gui_field_mapper.py
```
Opens an interactive GUI where you can:
- Select a PDF file
- Specify field locations
- Define values for each field
- Save the configuration as JSON

### 3. Batch Fill PDFs (CLI)
```bash
python pdf_filler.py --pdf input.pdf --json config.json --output output.pdf
```
Fills a PDF file based on a JSON configuration file from the command line.

### 4. GUI PDF Filler
```bash
python gui_pdf_filler.py
```
Launches a graphical interface to:
- Select PDF and JSON files
- Preview field mappings
- Fill and save the output PDF

## Requirements

- Python 3.7+
- Required packages (see `requirements.txt`)

## Example Workflow

```
1. Run placeholder_generator.py → Get template.json
2. Edit template.json with your field data
3. Run gui_pdf_filler.py → Select PDF & JSON → Fill & Save
```

Or use CLI:
```
python pdf_filler.py --pdf form.pdf --json data.json --output filled_form.pdf
```

## Contributing

Feel free to submit issues and enhancement requests!

## License

MIT License