import azure.functions as func
import io
from docx import Document
import json
import base64
import logging

_ALLOWED_HTTP_METHOD = "POST"

def main(req: func.HttpRequest) -> func.HttpResponse:
    logging.info('Python HTTP trigger function processed a request.')

    try:
        # Check if the request method is POST
        if req.method != _ALLOWED_HTTP_METHOD:
            return func.HttpResponse(
                "This function only supports POST requests.",
                status_code=405
            )

        # Attempt to parse the JSON body
        req_body = req.get_json()
        logging.info(f"Request Body: {req_body}")

        # Extract the file content and content-type
        base64_file = req_body.get("$content")
        content_type = req_body.get("$content-type")

        if not base64_file or not content_type:
            return func.HttpResponse(
                "File or content-type not provided in the request",
                status_code=400
            )

        # Decode the base64 file
        file_bytes = base64.b64decode(base64_file)
        doc = Document(io.BytesIO(file_bytes))  # Read Word document

        # Extract data from the document
        extracted_data = extract_data_from_controls(doc)

        # Convert the extracted data to JSON
        json_data = json.dumps(extracted_data)

        return func.HttpResponse(json_data, mimetype="application/json")

    except Exception as e:
        logging.error(f"Error: {e}")
        return func.HttpResponse(
            f"Error processing request: {str(e)}", 
            status_code=500
        )

# Ensure your extract_data_from_controls function is defined here or imported if it's defined in another module
