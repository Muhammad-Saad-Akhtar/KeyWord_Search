# Keyword Searcher

## Overview
Keyword Searcher is a document search utility designed to efficiently find and highlight keywords in `.docx` files within a specified folder. The tool features a user-friendly GUI built with `DearPyGui` and supports real-time search progress tracking.

## Features
- **Keyword Search in `.docx` files**: Scans all Word documents in a given folder.
- **Exclusion of Words**: Allows users to specify words to be omitted from the search.
- **Progress Tracking**: Displays search progress with a progress bar.
- **Result Highlighting**: Option to highlight search results in output documents.
- **Graphical User Interface**: Built using `DearPyGui` for an interactive experience.

## Installation
### Prerequisites
Ensure you have the following installed:
- Python 3.x
- Required dependencies:
  ```bash
  pip install dearpygui python-docx tqdm
  ```

## Usage
### Running the Tool
1. Clone the repository:
   ```bash
   git clone https://github.com/yourusername/Keyword-Searcher.git
   ```
2. Navigate to the project directory:
   ```bash
   cd Keyword-Searcher
   ```
3. Run the UI script:
   ```bash
   python ui.py
   ```

### How It Works
- Enter a keyword in the search box.
- Click "Start Search" to begin scanning `.docx` files in the configured folder.
- The progress bar updates as files are scanned.
- Results are displayed in the UI, showing filename, paragraph number, and matched text.
- If configured, results can be saved with highlighted text.

## File Structure
```
Keyword-Searcher/
│── ui.py            # Graphical User Interface
│── Pandora.py       # Search logic and highlighting
│── README.md        # Project documentation
```

## Configuration
- Update `FOLDER_PATH` in `ui.py` with the correct path to your `.docx` files.
- Modify `highlight_and_save_results` in `Pandora.py` to change output settings.

## Contributing
Feel free to fork the repository and submit pull requests for improvements!
