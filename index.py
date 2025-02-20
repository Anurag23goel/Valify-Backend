import os
import json
import tempfile
import requests
import base64
from datetime import datetime

import firebase_admin
from firebase_admin import credentials, firestore, storage, auth

from flask import Flask, request, jsonify, send_file
from flask_cors import CORS

from dotenv import load_dotenv
from openpyxl import load_workbook
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter


# Load environment variables
if os.getenv("RENDER") is None:  # Render automatically sets this variable
    if not load_dotenv():
        print("Warning: .env file not found. Ensure environment variables are set!")

# Function to verify Firebase Authentication Token
def verify_token(id_token):
    try:
        decoded_token = auth.verify_id_token(id_token)
        return decoded_token
    except auth.ExpiredIdTokenError:
        raise Exception("Token has expired")
    except auth.RevokedIdTokenError:
        raise Exception("Token has been revoked")
    except Exception as e:
        raise Exception(f"Error verifying token: {str(e)}")


# Initialize Firebase
firebase_credentials = os.getenv("FIREBASE_CREDENTIALS")

if not firebase_credentials:
    raise ValueError("Missing FIREBASE_CREDENTIALS environment variable!")

try:
    cred_dict = json.loads(firebase_credentials)
    cred = credentials.Certificate(cred_dict)
    firebase_admin.initialize_app(cred, {"storageBucket": "valify-7e530.appspot.com"})
    db = firestore.client()
except json.JSONDecodeError as e:
    raise ValueError(f"Invalid FIREBASE_CREDENTIALS JSON: {e}")


# Initialize Flask app
app = Flask(__name__)
CORS(app)

# Define paths
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
TEMPLATE_PATH = os.path.join(BASE_DIR, "dynamic_excel.xlsx")  # Ensure this file exists
OUTPUT_DIR = os.path.join(BASE_DIR, "output")
os.makedirs(OUTPUT_DIR, exist_ok=True)

# JSON to Excel mapping
json_to_excel_mapping = {
    "Inputs": {
        "valuerType": "E18",
        "clientName": "E19",
        "valuerName": "E20",
        "purpose": "E21",
        "premise": "E22",
        "draftNote": "E23",
        "projectTitle": "E26",
        "subjectCompanyName": "E24",
        "shortName": "E25",
        "nextFiscalYearEndDate": "E30",
        "valuationDate": "E29",
        "ytd": "E33",
        "ytgApproach": "E36",
        "informationCurrency": "E43",
        "presentationCurrency": "E44",
        "units": "E46",
        "industryPrimaryBusiness": "E52",
        "subindustryPrimaryBusiness": "E53",
        "primaryBusiness": "E54",
        "primaryBusinessDescription": "E55",
        "primaryRegions": "E56",
        "industrySecondaryBusiness": "E63",
        "subindustrySecondaryBusiness": "E64",
        "secondaryBusiness": "E65",
        "secondaryBusinessDescription": "E66",
        "secondaryRegions": "E67",
        "avgAnnualRevenue": "E49",
        "developmentPhase": "E50",
    }
}


# Route to remove formulas from an Excel file
@app.route('/remove-formulas', methods=['POST'])
def remove_formulas_route():
    if 'file' not in request.files:
        return jsonify({"error": "No file uploaded"}), 400

    excel_file = request.files['file']

    with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as temp_input:
        excel_file.save(temp_input.name)
        temp_output = temp_input.name.replace(".xlsx", "_no_formulas.xlsx")

        remove_formulas_from_excel(temp_input.name, temp_output)
        return send_file(temp_output, as_attachment=True)


# Route to generate an Excel file with Firestore data
@app.route('/generate-excel', methods=['GET'])
def generate_excel():
    try:
        uid = request.args.get('uid')
        project_id = request.args.get('project_id')

        if not uid or not project_id:
            return jsonify({"error": "uid and project_id are required"}), 400

        doc_ref = db.collection("users").document(uid).collection("projects").document(project_id)
        doc = doc_ref.get()

        if not doc.exists:
            return jsonify({"error": "Document not found"}), 404

        res = doc.to_dict()
        data = res.get("answers", {})

        workbook = load_workbook(TEMPLATE_PATH, keep_vba=True)

        if "Inputs" not in workbook.sheetnames:
            return jsonify({"error": "Excel template is missing 'Inputs' sheet"}), 500

        worksheet = workbook["Inputs"]

        for field, cell_location in json_to_excel_mapping["Inputs"].items():
            value = data.get(field, None)
            if value is not None:
                worksheet[cell_location].value = value

        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        output_filename = f"final_invoice_{timestamp}.xlsx"
        output_path = os.path.join(OUTPUT_DIR, output_filename)

        workbook.save(output_path)
        return send_file(output_path, as_attachment=True)

    except Exception as e:
        return jsonify({"error": str(e)}), 500


# Function to remove formulas from an Excel file
def remove_formulas_from_excel(input_file: str, output_file: str):
    wb = load_workbook(input_file, data_only=True)

    for sheet in wb.worksheets:
        for row in sheet.iter_rows():
            for cell in row:
                if cell.data_type == 'f':  
                    cell.value = cell.value  
                    cell.data_type = 'n'  

    wb.save(output_file)
    print(f"Processed file saved as: {output_file}")


# Route to convert an Excel file to PDF
@app.route('/convert-to-pdf', methods=['POST'])
def convert_to_pdf_route():
    if 'file' not in request.files:
        return jsonify({"error": "No file uploaded"}), 400

    excel_file = request.files['file']

    with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as temp_input:
        excel_file.save(temp_input.name)
        temp_output = temp_input.name.replace(".xlsx", ".pdf")

        convert_excel_to_pdf(temp_input.name, temp_output)
        return send_file(temp_output, as_attachment=True)


# Function to convert an Excel file to PDF using ConvertAPI
CONVERT_API_KEY = os.getenv("CONVERT_API_KEY")
CONVERT_API_URL = "https://v2.convertapi.com/convert/xls/to/pdf"

if not CONVERT_API_KEY:
    raise ValueError("Missing CONVERT_API_KEY in environment variables!")

def convert_excel_to_pdf(input_file: str, output_file: str):
    headers = {"Authorization": f"Bearer {CONVERT_API_KEY}"}
    files = {"File": open(input_file, "rb")}
    data = {"StoreFile": "false", "WorksheetActive": "true", "PageOrientation": "landscape"}

    try:
        response = requests.post(CONVERT_API_URL, headers=headers, files=files, data=data)
        response.raise_for_status()
        response_data = response.json()
        file_data_base64 = response_data["Files"][0]["FileData"]
        file_data_bytes = base64.b64decode(file_data_base64)

        with open(output_file, "wb") as pdf_file:
            pdf_file.write(file_data_bytes)

        print(f"PDF successfully saved as '{output_file}'")
    except requests.exceptions.RequestException as e:
        print(f"Error during API request: {e}")


@app.route('/health', methods=['GET'])
def health_check():
    return jsonify({"status": "ok", "message": "Flask app is running"}), 200


# Run the Flask app
if __name__ == '__main__':
    app.run(debug=True, host='0.0.0.0', port=5000)



