# PDF to Word Converter

A user-friendly GUI application for converting PDF files to Microsoft Word documents (.docx format).

## Features

✨ **Easy-to-Use Interface**
- Clean, modern GUI built with Python tkinter
- Drag-and-drop friendly file selection
- Real-time conversion progress

📁 **Flexible Output Options**
- Save converted files to the same folder as the input PDF
- Choose a custom output folder
- Batch conversion support (coming soon)

⚙️ **Conversion Options**
- Preserve PDF formatting in the Word document
- Support for multi-page PDFs
- Detailed conversion logging

## Requirements

- Python 3.6 or higher
- tkinter (usually included with Python)
- pdf2docx
- python-docx

## Installation

1. **Clone or download this project:**
   ```bash
   cd pdf-word
   ```

2. **Install dependencies:**
   ```bash
   pip install -r requirements.txt
   ```

   Or install manually:
   ```bash
   pip install pdf2docx python-docx
   ```

## Usage

### GUI Application

Run the GUI application:
```bash
python main.py
```

Then:
1. Click "Browse PDF" to select your PDF file
2. Choose output location (same folder or custom folder)
3. Click "Convert to Word"
4. Wait for the conversion to complete

### Command Line Usage

You can also use the converter module in your own Python scripts:

```python
from converter import PDFToWordConverter

converter = PDFToWordConverter()

# Single file conversion
success, message = converter.convert(
    pdf_path="path/to/file.pdf",
    output_path="path/to/output.docx"
)

print(message)

# Batch conversion
converted_count, message = converter.batch_convert(
    pdf_folder="path/to/pdf/folder",
    output_folder="path/to/output/folder"
)

print(f"Converted {converted_count} files")
```

## File Structure

```
pdf-word/
├── main.py              # GUI application
├── converter.py         # PDF conversion logic
├── requirements.txt     # Python dependencies
└── README.md           # This file
```

## Troubleshooting

### Module not found error
Make sure all dependencies are installed:
```bash
pip install -r requirements.txt
```

### "No module named 'pdf2docx'"
Install pdf2docx:
```bash
pip install pdf2docx
```

### Conversion fails
- Ensure the PDF file is not corrupted
- Try with a different PDF file
- Check that you have write permissions to the output folder

## Limitations

- Works best with text-based PDFs (image-based PDFs may have limited conversion quality)
- Complex PDF layouts may not convert perfectly to Word format
- Very large PDFs may take longer to convert

## Future Enhancements

- [ ] Batch conversion for multiple files
- [ ] OCR support for image-based PDFs
- [ ] Advanced formatting options
- [ ] Drag-and-drop support
- [ ] Command-line interface improvements

## License

This project is open source and available under the MIT License.

## Support

For issues or questions, please check the troubleshooting section above or ensure all dependencies are properly installed.

---

**Enjoy converting your PDFs to Word documents! 📄→📝**
