Abdulrahman Mohamed Eltahan 21100958

## Project Description
This project involves preprocessing unstructured data from various document types (PDF, Word, Excel, PowerPoint) and generating normalized JSON files suitable for LLM (Large Language Model) applications. The preprocessing pipeline extracts text, cleans it, and saves it in a consistent JSON format, ready for use in RAG (Retrieval-Augmented Generation) systems or other NLP tasks.

## Folder Structure
```

AIE314-Tutorial1/
│
├── data/               # Folder containing the original documents
├── output/             # Folder containing the normalized JSON files
├── preprocessing.py    # Python script with preprocessing pipeline
└── README.md           # Project details and instructions

````

## How to Run the Code
1. Place your documents (PDF, Word, Excel, PPTX) in the `data/` folder.
2. Install the required Python libraries:
````
   pip install pymupdf python-docx pandas python-pptx openpyxl
````

3. Run the preprocessing script:

   ```
   python preprocessing.py
   ```
4. Check the `output/` folder for the generated JSON files.

## Notes

* The pipeline automatically skips unsupported file types.
* Text is cleaned by removing extra whitespace.
* Each JSON file includes:

  * `title`: file name without extension
  * `content`: extracted and cleaned text
  * `metadata`: source file name, file type, and text length

## Example Output

```json
{
    "title": "Ahmed Abdelbaset Badir",
    "content": "Objective: Seeking a position as an Aluminum Fabrication Foreman or Assistant supervisor...",
    "metadata": {
        "source_file": "Ahmed Abdelbaset Badir.docx",
        "file_type": "docx"
    }
}
