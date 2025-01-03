# Dynamic Word Document Creation with Save Dialog

This Python script dynamically creates a Microsoft Word document (`.docx`) using the `python-docx` library and allows the user to save the file at a desired location with a custom name through a graphical "Save As" dialog, powered by the `tkinter` library.

---

## Features
- **Dynamic File Saving**: Prompts the user to select a location and specify a file name for the document.
- **Custom Content**: Prepares a sample curriculum vitae (CV) with placeholders for professional and educational details.
- **User-Friendly Interface**: Utilizes a graphical interface for file-saving.
- **Default File Type**: Ensures the file is saved as a `.docx` document.

---

## Prerequisites
Ensure you have the following installed on your system:
1. **Python 3.6+**
2. **Required Libraries**:
   - `python-docx`
   - `tkinter` (comes pre-installed with Python on most platforms)

To install `python-docx`, run:
```bash
pip install python-docx
```

---

## How to Run the Script
1. Save the script as `dynamic_save_cv.py`.
2. Open a terminal or command prompt and navigate to the directory containing the script.
3. Run the script using:
   ```bash
   python dynamic_save_cv.py
   ```
4. A "Save As" dialog will appear. Choose the desired location and file name.
5. The script will save the Word document to the specified location and notify you of the saved file path.

---

## Code Overview
### Key Libraries
- **`python-docx`**: For creating and editing Word documents.
- **`tkinter`**: For opening a graphical "Save As" dialog.

### Highlights of the Code
#### Document Content Creation
The script prepares a sample CV with headings and paragraphs:
```python
from docx import Document
doc = Document()
doc.add_heading('Jatin Chandani', level=0)
doc.add_paragraph('Colchester, UK | jatinchandani8@gmail.com | +44 7407 022519')
```
#### Dynamic Save Dialog
The `tkinter` library's `asksaveasfilename` function is used to open the dialog:
```python
from tkinter.filedialog import asksaveasfilename
file_path = asksaveasfilename(defaultextension=".docx",
                               filetypes=[("Word Documents", "*.docx")],
                               title="Save Your CV",
                               initialfile="Jatin_Chandani_Updated_CV.docx")
```

---

## Example Output
When the script is executed, the following happens:
1. A "Save As" dialog box appears.
2. The user selects a file location and enters a file name.
3. The script generates a Word document with the specified content and saves it.
4. A success message is displayed in the terminal:
   ```bash
   File saved at: /path/to/your/directory/Your_File_Name.docx
   ```

---

## Customization
You can modify the document content by editing the `doc.add_heading` and `doc.add_paragraph` lines in the script to include your own data.

---

## Troubleshooting
- **Error: `No module named 'docx'`**:
  - Ensure `python-docx` is installed using:
    ```bash
    pip install python-docx
    ```
- **Save Dialog Doesn't Open**:
  - Verify `tkinter` is installed. On Linux, you may need to install it via your package manager (e.g., `sudo apt-get install python3-tk`).

---

## License
Copyright (c) 2025 Jatin Chandani

Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the "Software"), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all
copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN
THE SOFTWARE.

---

## Author
**Jatin Chandani**  
[LinkedIn Profile](https://www.linkedin.com/in/jatinchandani28)

