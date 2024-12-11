# Certificate Generator App

This application allows you to generate personalized certificates from an Excel file containing student names and achievements. The certificates are saved as PDF files and compressed into a ZIP file for easy download.

## Features
- Upload an Excel file containing `Name` and `Surname` columns.
- Input the achievement to be displayed on the certificates.
- Generate PDF certificates for each student.
- Download all certificates as a single ZIP file.

---

## Requirements

### MacOS
- Python 3.7 or later
- Required Python packages:
  - `tkinter`
  - `pandas`
  - `fpdf`

### Windows
- Python 3.7 or later
- Required Python packages:
  - `tkinter`
  - `pandas`
  - `fpdf`

---

## Setup and Installation

### Step 1: Clone the Repository

1. Open your terminal (MacOS) or command prompt (Windows).
2. Clone the repository or download the project files.
   ```bash
   git clone <repository-url>
   cd tkinter_certificate_app
   ```

### Step 2: Set Up a Virtual Environment

1. Create a virtual environment:
   ```bash
   python -m venv env
   ```

2. Activate the virtual environment:
   - **MacOS/Linux**:
     ```bash
     source env/bin/activate
     ```
   - **Windows**:
     ```bash
     .\env\Scripts\activate
     ```

3. Verify the virtual environment is activated:
   - Your terminal prompt should show `(env)` at the beginning.

4. Install required packages:
   ```bash
   pip install pandas fpdf
   ```

   For MacOS:
   - Ensure `tkinter` is installed by running:
     ```bash
     brew install python-tk
     ```

   For Windows:
   - `tkinter` is bundled with Python. No additional installation is needed.

---

## How to Run the Application

1. Navigate to the project directory:
   ```bash
   cd tkinter_certificate_app
   ```

2. Activate the virtual environment if not already active:
   - **MacOS/Linux**:
     ```bash
     source env/bin/activate
     ```
   - **Windows**:
     ```bash
     .\env\Scripts\activate
     ```

3. Run the application:
   ```bash
   python app.py
   ```

4. The application window will open. Follow these steps to use the app:

   - Click "Browse" to select an Excel file containing the `Name` and `Surname` columns.
   - Enter the achievement name.
   - Click "Generate Certificates."
   - Choose an output folder to save the certificates.
   - Download the ZIP file containing all certificates.

---

## Creating a Sample Excel File
To test the application, create an Excel file (`.xlsx`) with the following structure:

| Name      | Surname    |
|-----------|------------|
| John      | Doe        |
| Jane      | Smith      |

Save this file and upload it in the app.


## License
This project is licensed under the MIT License.
