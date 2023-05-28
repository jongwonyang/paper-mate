from azure.core.exceptions import ResourceNotFoundError
from azure.core.credentials import AzureKeyCredential
from azure.ai.formrecognizer import FormRecognizerClient, DocumentAnalysisClient
import json

credentials = json.load(open('paper2slide/credentials.json'))
API_KEY = credentials["API_KEY"]
ENDPOINT = credentials["ENDPOINT"]

def extract_data(path):
    source = path
    form_recognizer_client = DocumentAnalysisClient(ENDPOINT, AzureKeyCredential(API_KEY))
    with open(source, "rb") as f:
        data_bytes = f.read()
    poller = form_recognizer_client.begin_analyze_document("prebuilt-layout",document=data_bytes)#,content_type='application/pdf')

    form_result = poller.result().to_dict()
    return form_result
