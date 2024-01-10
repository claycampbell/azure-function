import azure.functions as func
import json
import logging
import base64
import io
from docx import Document
import traceback
from docx.oxml.ns import qn
import pandas as pd
import tempfile
import logging
import os
app = func.FunctionApp()


@app.route(route="process_document", auth_level=func.AuthLevel.ANONYMOUS, methods=['POST'])
def process_document(req: func.HttpRequest) -> func.HttpResponse:
    logging.info('Python HTTP trigger function processed a request.')

    try:
        # Get the raw body as bytes
        file_bytes = req.get_body()

        # Check if the request body is empty
        if not file_bytes:
            return func.HttpResponse("Request body is empty", status_code=400)

        # Read the Word document from the binary data
        doc = Document(io.BytesIO(file_bytes))

        # Process the Word document for control data extraction
        control_data = extract_data_from_controls(doc)
        logging.info(f"Control Data: {control_data}")

        # Process the Word document for track changes extraction
        track_changes_data = extract_data_with_track_changes(doc)
        logging.info(f"Track Changes Data: {track_changes_data}")

        # Combine both results
        combined_result = {
            "Entity Extracts": control_data,
            "Tracked Changes": track_changes_data
        }

        # Return the combined result
        return func.HttpResponse(json.dumps(combined_result), mimetype="application/json", status_code=200)

    except Exception as e:
        logging.error("Exception type: " + str(type(e)))
        logging.error("Error processing request: " + str(e))
        logging.error("Traceback: " + traceback.format_exc())
        return func.HttpResponse("Error processing request: " + str(e), status_code=500)
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
def extract_track_changes(p_element):
    """
    Extracts text including proposed insertions, deletions, and modifications in track changes.
    """
    text = ''
    for child in p_element.iterchildren():
        if child.tag == qn('w:t'):  # Regular text
            text += child.text
        elif child.tag == qn('w:ins'):  # Inserted text
            inserted_text = ''.join(node.text for node in child.iterdescendants(tag=qn('w:t')))
            text += f" [Inserted: {inserted_text}]"
        elif child.tag == qn('w:del'):  # Deleted text
            deleted_text = ''.join(node.text for node in child.iterdescendants(tag=qn('w:t')))
            text += f" [Deleted: {deleted_text}]"
    return text
def extract_data_with_track_changes(doc):
    data = []
    for p in doc.paragraphs:
        text = extract_track_changes(p._element)
        data.append(text)
    return [element for element in data if element != ""]
