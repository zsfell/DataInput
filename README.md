# Data Input Form with Excel Integration

## Project Overview
This project is a web-based application that allows users to input data into a form and store the data in an Excel spreadsheet. The application is built using Python and Flask for the server-side logic and HTML with JavaScript for the client-side form. The primary goal is to ensure that new numeric inputs are added to existing values in the same row of the Excel sheet, without creating new rows unnecessarily, and to prevent the submission of empty forms.

## Technology Stack
- **Python 3**
- **Flask**: A lightweight WSGI web application framework.
- **Openpyxl**: A Python library to read/write Excel xlsx/xlsm/xltx/xltm files.
- **HTML**: Markup language for creating web pages.
- **JavaScript**: Programming language to enhance web pages with client-side validation.


## Setup Instructions
1. **Install Dependencies**:
   Ensure you have Python 3 and pip installed. Install the required libraries using:
   ```sh
   pip install flask openpyxl
   
2. **Run the Flask Application**:
   Start the Flask server by running the following command in the project directory:
   ```sh
   python app.py
3. **Access the Application**:

   Open your web browser and navigate to http://127.0.0.1:5000 to access the data input form.
