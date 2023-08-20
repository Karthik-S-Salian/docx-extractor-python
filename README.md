# docx-extractor-python
DocxExtractor - Extract Text and Images from DOCX Files

This module provides a class for extracting text and content from DOCX files. The class can handle extracting text from the document body, headers, footers, as well as extracting images and unzipping the contents of the DOCX file.

Usage Example:
--------------
# Import the module
from docx_extractor import DocxExtractor

# Initialize a DocxExtractor instance with the path to the DOCX file
docx_path = "path/to/your/file.docx"
with DocxExtractor(docx_path) as extractor:
    # Extract and print the text content of the document body
    body_text = extractor.extractDocumentBodyText()
    print("Document Body Text:")
    print(body_text)

    # Unzip the DOCX contents to a directory
    # exists ok = True to rewrite the folder if exists
    extractor.unzipDocxToFolder("unzipped_folder", exist_ok=True)

    # Extract text content from headers
    header_text = extractor.extractHeaderText()
    print("Header Text:")
    print(header_text)

    # Extract text content from footers
    footer_text = extractor.extractFooterText()
    print("Footer Text:")
    print(footer_text)

    # Extract images and save them to a directory
    extractor.extractImages("image_folder", exist_ok=True)
    print("Images extracted and saved to 'image_folder'.")

    # Get the hierarchical file tree of the DOCX contents
    file_tree = extractor.getFileTree()
    print("File Tree:")
    print(file_tree)

# The DOCX file will be automatically closed when exiting the context

Note: Make sure to replace "path/to/your/file.docx" with the actual path of your DOCX file.

Class Methods:
--------------
- `__init__(self, docx: str) -> None`:
  Constructor method. Initializes the DocxExtractor instance with the provided DOCX file path.

- `extractDocumentBodyText(self) -> str`:
  Extracts and returns the text content of the document body.

- `unzipDocxToFolder(self, store_dir: str = "", exist_ok: bool = False) -> None`:
  Unzips the contents of the DOCX file to the specified directory.

- `extractHeaderText(self) -> str`:
  Extracts and returns the text content from headers in the DOCX file.

- `extractFooterText(self) -> str`:
  Extracts and returns the text content from footers in the DOCX file.

- `extractImages(self, store_dir: str = "", exist_ok: bool = False) -> None`:
  Extracts images from the DOCX file and saves them to the specified directory.

- `getFileTree(self) -> Dict[str, Any]`:
  Returns a hierarchical dictionary representing the file tree of the DOCX contents.

- `closeDocx(self) -> None`:
  Closes the underlying DOCX(zip) file in the memory.

Usage Notes:
------------
- The `DocxExtractor` class is used to extract text and content from DOCX files.
- The class methods provide various functionalities for extracting different types of content.
- The class is intended to be used as a context manager (using the `with` statement) to ensure proper resource management.

