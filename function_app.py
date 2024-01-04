import azure.functions as func
import datetime
import json
import logging
import io
from docx import Document
import base64
import traceback

app = func.FunctionApp()

@app.route(route="extract", auth_level=func.AuthLevel.ANONYMOUS)
def extract(req: func.HttpRequest) -> func.HttpResponse:
    logging.info('Python HTTP trigger function processed a request.')

    try:
        # Log the entire raw request body and headers
        logging.info(f"Raw request body: {req.get_body().decode()}")
        logging.info(f"Request headers: {req.headers}")

        # Parse the JSON body of the request
        req_body = req.get_json()
        logging.info(f"Parsed JSON body: {req_body}")

        # Extract the file content and content-type
        base64_file = req_body.get("$content")
        content_type = req_body.get("$content-type")
        logging.info(f"Base64 file content: {base64_file}")
        logging.info(f"Content type: {content_type}")

        if not base64_file or not content_type:
            return func.HttpResponse(
                "File or content-type not provided in the request",
                status_code=400
            )

        # Optionally, validate the content type here
        if content_type != "application/vnd.openxmlformats-officedocument.wordprocessingml.document":
            return func.HttpResponse(
                "Unsupported content type",
                status_code=415
            )

        # Decode the base64 file
        file_bytes = base64.b64decode(base64_file)
        logging.info("Successfully decoded base64 file")
        doc = Document(io.BytesIO(file_bytes))  # Read Word document from decoded bytes

        # Extract data from the document
        extracted_data = extract_data_from_controls(doc)
        logging.info(f"Extracted data: {extracted_data}")

        # Convert the extracted data to JSON
        json_data = json.dumps(extracted_data)
        return func.HttpResponse(json_data, mimetype="application/json")

    except Exception as e:
        # Detailed exception logging
        logging.error("Exception type: " + str(type(e)))
        logging.error("Error processing request: " + str(e))
        logging.error("Traceback: " + traceback.format_exc())
        return func.HttpResponse(
            "Error processing request: " + str(e),
            status_code=500
        )
        
def extract_data_from_controls(doc):
    data = {}
    for part in doc.element.body.iter():
        if part.tag.endswith('sdt'):
            sdtPr = part.find('.//{http://schemas.openxmlformats.org/wordprocessingml/2006/main}sdtPr')
            tagElem = sdtPr.find('.//{http://schemas.openxmlformats.org/wordprocessingml/2006/main}tag') if sdtPr is not None else None
            tag = tagElem.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val') if tagElem is not None else None
            
            texts = part.xpath('.//w:t', namespaces={'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'})
            if texts:
                content = ''.join(t.text for t in texts)
                if tag:
                    if tag not in data:
                        data[tag] = []
                    data[tag].append(content)

    logging.info(f"Data extracted from document controls: {data}")
    return data
