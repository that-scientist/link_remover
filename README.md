# Link Remover

A Python tool that removes hyperlinks from `.docx` and PDF files while preserving all human-readable text. The tool also removes hyperlink-associated font color changes, ensuring all text appears in black.

## Features

- ✅ Removes hyperlinks from `.docx` files
- ✅ Removes hyperlinks from PDF files
- ✅ Preserves all text content unchanged
- ✅ Removes hyperlink color formatting (sets text to black)
- ✅ Automatically detects and uses virtual environment
- ✅ Processes files in batch from an input folder
- ✅ Organizes processed files into output and done folders

## Requirements

- Python 3.6 or higher
- Virtual environment (recommended)

## Installation

1. Clone the repository:
```bash
git clone <repository-url>
cd link_remover
```

2. Create a virtual environment (recommended):
```bash
python3 -m venv venv
```

3. Activate the virtual environment:
```bash
# On macOS/Linux:
source venv/bin/activate

# On Windows:
venv\Scripts\activate
```

4. Install dependencies:
```bash
pip install -r requirements.txt
```

## Usage

1. Place your `.docx` or `.pdf` files in the `input/` folder

2. Run the script:

   **Option A: Double-click (macOS)**
   - Double-click `link_remover.command` in Finder
   - A Terminal window will open and run the script automatically

   **Option B: Command line**
   ```bash
   ./link_remover.py
   ```

   Or if the virtual environment is activated:
   ```bash
   python link_remover.py
   ```

   Or using Python directly:
   ```bash
   python3 link_remover.py
   ```

3. Results:
   - **Processed files** (without hyperlinks) → `out/` folder
   - **Original files** → `done/` folder
   - **Input folder** → emptied after processing

## How It Works

### For .docx Files
- Scans all paragraphs and table cells for hyperlinks
- Extracts text content from hyperlink elements
- Removes hyperlink formatting and color
- Sets text color to black
- Preserves all other text formatting

### For PDF Files
- Removes annotation objects (which include hyperlinks)
- Preserves all text content and formatting
- Maintains document structure

## Project Structure

```
link_remover/
├── link_remover.py      # Main script
├── link_remover.command # macOS double-click launcher
├── requirements.txt      # Python dependencies
├── README.md            # This file
├── input/               # Place files to process here
├── out/                 # Processed files (without hyperlinks)
└── done/                # Original files after processing
```

## Dependencies

- `python-docx` (>=1.1.0) - For processing .docx files
- `pypdf` (>=3.0.0) - For processing PDF files

## Notes

- The script automatically detects and uses a virtual environment if one exists in the project directory
- Files are moved (not copied) from `input/` to `done/` after successful processing
- If processing fails, the original file remains in the `input/` folder
- The script processes all `.docx` and `.pdf` files found in the `input/` folder

## Troubleshooting

### "Library not installed" error
If you see an error about missing libraries:
1. Make sure the virtual environment is activated
2. Or install dependencies: `pip install -r requirements.txt`

### Script doesn't find files
- Ensure files are placed in the `input/` folder
- Check that files have `.docx` or `.pdf` extensions (case-sensitive)

## License

[Add your license here]

## Contributing

[Add contribution guidelines if applicable]

