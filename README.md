# Document Processing Tool

This tool processes Word documents by replacing values from text files into corresponding DOCX templates.

## Setup

1. Install the required dependencies:
```bash
pip3 install python-docx
pip3 install git+https://github.com/ryanpierson/merge_docx.git
```

2. Prepare your files:
   - Place your text files with values in the `valori/` directory
   - Place your DOCX templates in the `sample/` directory
   - The text files should contain key-value pairs in the format: `KEY = value`

## Usage

1. Run the script:
```bash
python3 app.py
```

The script will:
- Process each text file and its corresponding DOCX template
- Save intermediate results in `/tmp`
- Merge all documents into a final result
- Clean up temporary files
- Save the final merged document as `rez_final.docx` in the root directory

## File Structure

```
.
├── app.py
├── valori/
│   ├── n1.txt
│   ├── n2.txt
│   └── n3.txt
├── sample/
│   ├── n1.docx
│   ├── n2.docx
│   └── n3.docx
└── rez_final.docx (output)
```

## Notes

- The script processes files in pairs (text file + DOCX template)
- Intermediate files are stored in `/tmp` and automatically cleaned up
- The final merged document is saved as `rez_final.docx` in the root directory