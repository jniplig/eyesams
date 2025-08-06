# EyeSAMS - Excel Processing Tool

A Python tool for merging multiple Excel files with multiple worksheets into a consolidated dataset with comprehensive error handling and logging.

## Features

- ✅ **Multi-file Processing**: Handles multiple Excel files in a directory
- ✅ **Multi-sheet Support**: Processes all worksheets within each file
- ✅ **Data Validation**: Checks for empty sheets, insufficient data, corrupted files
- ✅ **Metadata Extraction**: Extracts teacher codes and worksheet names
- ✅ **Error Handling**: Comprehensive error catching with detailed reporting
- ✅ **Progress Tracking**: Real-time processing feedback and final statistics
- ✅ **Version Control**: Automatic output file versioning to prevent overwrites

## Installation

### 1. Clone the repository
```bash
git clone https://github.com/yourusername/eyesams.git
cd eyesams
```

### 2. Create a virtual environment
```bash
python -m venv venv
source venv/bin/activate  # On Windows: venv\Scripts\activate
```

### 3. Install dependencies
```bash
pip install -r requirements.txt
```

## Usage

### Basic Usage (Google Colab/Default Paths)
```python
from src.excel_processor import process_excel_files

# Process files with default Colab paths
result = process_excel_files()

if result:
    print(f"Success! Output saved to: {result}")
else:
    print("Processing failed - check error messages")
```

### Custom Directory Usage
```python
from src.excel_processor import process_excel_files

# Process files with custom paths
result = process_excel_files(
    input_directory="/path/to/excel/files",
    output_directory="/path/to/output"
)
```

### Command Line Usage
```bash
# Run with default settings
python src/excel_processor.py
```

## File Structure Requirements

Your Excel files should follow this structure:
- **Row 1**: Teacher information (last 5 characters will be extracted)
- **Row 2**: Column headers
- **Rows 3-N**: Data rows
- **Last 2 rows**: Footer information (will be removed)

## Output

The tool creates a consolidated Excel file with:
- All data from all processed worksheets
- **Teacher column**: Extracted teacher codes
- **Set column**: Original worksheet names
- Automatic version numbering (`merged_sets_1.xlsx`, `merged_sets_2.xlsx`, etc.)

## Error Handling

The processor handles various error conditions gracefully:
- Missing or inaccessible directories
- Corrupted or locked Excel files
- Empty worksheets or insufficient data
- Invalid teacher data
- File permission issues

## Project Structure

```
eyesams/
├── README.md
├── requirements.txt
├── src/
│   └── excel_processor.py    # Main processing module
├── tests/
│   └── test_data/            # Test Excel files (if any)
├── docs/                     # Documentation
└── examples/                 # Usage examples
```

## Development

### Setting up for development
```bash
# Clone and setup
git clone https://github.com/yourusername/eyesams.git
cd eyesams
python -m venv venv
source venv/bin/activate
pip install -r requirements.txt

# Test the installation
python -c "from src.excel_processor import process_excel_files; print('Import successful!')"
```

### Adding new features
1. Create a new branch: `git checkout -b feature-name`
2. Make your changes
3. Test thoroughly
4. Commit: `git commit -m "Add: feature description"`
5. Push: `git push origin feature-name`
6. Create a Pull Request

## Google Colab Integration

To use in Google Colab:

```python
# Clone the repository
!git clone https://github.com/yourusername/eyesams.git
%cd eyesams

# Install dependencies
!pip install -r requirements.txt

# Mount Google Drive
from google.colab import drive
drive.mount('/content/drive')

# Use the processor
from src.excel_processor import process_excel_files
result = process_excel_files()
```

## Contributing

1. Fork the repository
2. Create your feature branch (`git checkout -b feature/amazing-feature`)
3. Commit your changes (`git commit -m 'Add: amazing feature'`)
4. Push to the branch (`git push origin feature/amazing-feature`)
5. Open a Pull Request

## License

This project is licensed under the MIT License - see the LICENSE file for details.

## Author

Your Name - [your.email@example.com](mailto:your.email@example.com)

## Acknowledgments

- Built with [pandas](https://pandas.pydata.org/) for data manipulation
- Inspired by the need for efficient educational data processing