# Document Converter API

A Django REST API for converting Word documents to PDF with variable replacement.

## Features

- Convert .doc and .docx files to PDF
- Replace variables in documents using two different modes
- Store converted files in a dedicated directory
- Return PDF as downloadable file

## Installation

1. Clone the repository and navigate to the directory:

```bash
git clone <repository-url>
cd <repository-directory>
```

2. Set up a virtual environment:

```bash
python -m venv venv
source venv/bin/activate  # On Windows: venv\Scripts\activate
```

3. Install dependencies:

```bash
pip install -r requirements.txt
```

4. Run the development server:

```bash
python manage.py runserver
```

## API Usage

### Convert Document to PDF

**Endpoint:** `POST /api/convert/`

**Request Body:**
- `file`: The .doc or .docx file to convert (required)
- `variables`: JSON object or string with variables to replace in the document
- `replacement_mode`: Mode for variable replacement (optional, defaults to "dynamic")
  - `dynamic`: Uses Jinja-like placeholders `{{variable_name}}` in the document
  - `static`: Replaces exact word matches in the document



## Variable Replacement Modes

### Dynamic Mode (Default)
In dynamic mode, your document should contain placeholders in the format `{{variable_name}}`.

Example document content:
```
Hello, my name is {{name}} and I am {{age}} years old.
```

### Static Mode
In static mode, the API looks for exact word matches in the document and replaces them.



With variables `{"Introduction": "Custom intro", "Notes": "Custom notes"}`, the words "Introduction" and "Notes" will be replaced with the provided values.

## Error Handling

The API returns appropriate HTTP status codes and error messages:

- `400 Bad Request`: File not provided, unsupported file format, or invalid variables format
- `500 Internal Server Error`: Error during conversion process

