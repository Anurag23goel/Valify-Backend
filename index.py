from flask import Flask, request, jsonify, send_file
from openpyxl import load_workbook
from pycel.excelcompiler import ExcelCompiler
from datetime import datetime
import os
import firebase_admin
from firebase_admin import credentials, firestore, storage
from flask_cors import CORS
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter
import requests
import base64
import tempfile
from firebase_admin import auth
import os
import os
import json
import firebase_admin
from firebase_admin import credentials, firestore, storage



def verify_token(id_token):
    try:
        decoded_token = auth.verify_id_token(id_token)
        return decoded_token  # Token is valid, proceed with the request
    except auth.InvalidIdTokenError:
        raise Exception('Invalid or expired token')
    except Exception as e:
        raise Exception(f'Error verifying token: {str(e)}')


# Initialize Firebase
if not firebase_admin._apps:
    # Read Firebase credentials from environment variable
    firebase_credentials = os.getenv("FIREBASE_CREDENTIALS")

    if not firebase_credentials:
        raise ValueError("Missing FIREBASE_CREDENTIALS environment variable!")

    # Parse the JSON string
    cred_dict = json.loads(firebase_credentials)
    cred = credentials.Certificate(cred_dict)

    # Initialize Firebase
    firebase_admin.initialize_app(cred, {"storageBucket": "valify-7e530.appspot.com"})


# Initialize Firestore DB
db = firestore.client()

# Flask app
app = Flask(__name__)
CORS(app)

# Define template path
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
TEMPLATE_PATH = os.path.join(BASE_DIR, "dynamic_excel.xlsx")  # Ensure this file exists in the same directory
OUTPUT_DIR = "output"
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
        "existingStream1": "E77",
        "existingStream2": "E78",
        "existingStream3": "E79",
        "existingStream4": "E80",
        "pipelineStream1": "E82",
        "pipelineStream2": "E83",
        "pipelineStream3": "E84",
        "pipelineStream4": "E85",
        "potentialStream1": "E87",
        "potentialStream2": "E88",
        "potentialStream3": "E89",
        "potentialStream4": "E90",
        "potentialStream1Probability": "E93",
        "potentialStream2Probability": "E94",
        "potentialStream3Probability": "E95",
        "potentialStream4Probability": "E96",
    }
}


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


@app.route('/generate-excel', methods=['GET'])
def generate_excel():
    try:
        # Get UID and project_id from query parameters
        uid = request.args.get('uid')
        project_id = request.args.get('project_id')

        if not uid or not project_id:
            return jsonify({"error": "uid and project_id are required parameters"}), 400

        # Fetch Firestore data
        doc_ref = db.collection("users").document(uid).collection("projects").document(project_id)
        doc = doc_ref.get()

        if not doc.exists:
            return jsonify({"error": "Document not found"}), 404  # Handle missing document

        res = doc.to_dict()
        if not res or "answers" not in res:
            return jsonify({"error": "No answers found in the document"}), 404

        data = res["answers"]
        print(data)

        # Load the Excel template
        workbook = load_workbook(TEMPLATE_PATH, keep_vba=True)

        # Update the Excel sheet with Firestore data
        for sheet_name, field_map in json_to_excel_mapping.items():
            worksheet = workbook[sheet_name]
            for field, cell_location in field_map.items():
                value = data.get(field, None)
                if value is not None:
                    worksheet[cell_location].value = value

        # Save the file locally
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        output_filename = f"final_invoice_{timestamp}.xlsx"
        output_path = os.path.join("output", output_filename)
        os.makedirs("output", exist_ok=True)
        workbook.save(output_path)

        # Return the file as a download
        return send_file(output_path, as_attachment=True)

    except Exception as e:
        return jsonify({"error": str(e)}), 500





# Supporting functions
def remove_formulas_from_excel(input_file: str, output_file: str):
    # Load the workbook with data_only=True to retain computed values
    wb = load_workbook(input_file, data_only=True)

    for sheet in wb.worksheets:
        for row in sheet.iter_rows():
            for cell in row:
                if cell.data_type == 'f':  # 'f' indicates a formula
                    cell.value = cell.value  # Replace formula with computed value
                    cell.data_type = 'n'  # Set data type to number

    # Save the static version of the workbook
    wb.save(output_file)
    print(f"Processed file saved as: {output_file}")


def convert_excel_to_pdf(input_file: str, output_file: str):
    API_KEY = "secret_5CEgfbil9AyOMKxw"  # Replace with your actual API key
    CONVERT_API_URL = "https://v2.convertapi.com/convert/xls/to/pdf"

    headers = {
        "Authorization": f"Bearer {API_KEY}"
    }
    files = {
        "File": open(input_file, "rb")
    }
    data = {
        "StoreFile": "false",
        "WorksheetActive": "true",
        "PageOrientation": "landscape",
    }

    try:
        # Make the POST request to the conversion API
        response = requests.post(CONVERT_API_URL, headers=headers, files=files, data=data)
        response.raise_for_status()  # Raise an exception for HTTP errors
        response_data = response.json()
        file_data_base64 = response_data['Files'][0]['FileData']
        file_data_bytes = base64.b64decode(file_data_base64)

        # Save the PDF if the request was successful
        with open(output_file, "wb") as pdf_file:
            pdf_file.write(file_data_bytes)

        print(f"PDF successfully downloaded as '{output_file}'")
    except requests.exceptions.RequestException as e:
        print(f"Error during API request: {e}")



if __name__ == '__main__':
    app.run(debug=True, host='0.0.0.0', port=5000)