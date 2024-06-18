import json
import base64
import io
from email.parser import Parser
from email.message import EmailMessage
import time
from openai import OpenAI
import boto3
import uuid


def lambda_handler(event, context):
    # Check if the Headers and Body are there
    print(f"Headers: {event['headers']}")
    original_headers = event["headers"]
    headers = {k.lower(): v for k, v in original_headers.items()}
    debug_mode = headers.get("debug_mode", "false").lower() == "true"
    upload_id = uuid.uuid4()
    print(f"Debug Mode: {debug_mode}")
    if debug_mode:
        response = {
            "uploadId": "2f951640-24fc-46ba-93d8-c7558bd6d0e2",
            "message": "CV uploaded successfully and initiated processing",
        }
        return generate_response(200, response)
    if not ("body" in event and "headers" in event and "content-type" in headers):
        if not "body" in event:
            return generate_response(400, "Bad Request - No Body found in request")
        return generate_response(400, "Bad Request - No Headers found in request")

    # Exctratct pdf from Request and Save it as a Bytestream
    try:
        pdf = extract_pdf_from_formdata(
            body64=event["body"],
            content_type=headers["content-type"],
            upload_id=f"{upload_id}",
        )
    except Exception as e:
        return generate_response(400, f"Bad Request - No PDF found in request: {e}")

    # Upload the extracted pdf file to the the OpenAI API
    try:
        openai_file, vectorstore = upload_file_to_openai_v2(
            pdf=pdf, upload_id=upload_id
        )
    except Exception as e:
        return generate_response(
            500, f"Internal Server Error - Error uploading PDF to OpenAI API: {e}"
        )

    # Initialize a database entry for this CV
    try:
        primaryKey = initialize_database_entry(
            openai_file=openai_file, vectorstore=vectorstore, upload_id=upload_id
        )
    except Exception as e:
        return generate_response(
            500,
            f"Internal Server Error - Error Initializing the Database entry for this CV - DynamoDB Error: {e}",
        )

    # Start the Processing of the CV in the Step Function
    try:
        startStateMachineProcessing(openai_file, primaryKey, vectorstore)
    except Exception as e:
        return generate_response(
            500,
            f"Internal Server Error - Error Initializing the processing of the CV - StepFunctions Error: {e}",
        )

    response = {
        "uploadId": primaryKey,
        "message": "CV uploaded successfully and initiated processing",
    }
    return generate_response(200, response)


def startStateMachineProcessing(oai_file, uploadId, vectorstore):
    statemachine_client = boto3.client("stepfunctions")
    state_machine_arn = (
        "arn:aws:states:eu-central-1:891376982948:stateMachine:MyStateMachine-h1zdun91x"
    )
    assistant_id = "asst_ab8KCfa3TRFd5MbN0iGXs9bj"
    state_input_data = {
        "upload_id": uploadId,
        "assistant_id": assistant_id,
        "file_id": oai_file.id,
        "vectorstore_ids": [vectorstore.id],
        "file_ids": [oai_file.id],
        "prompt": f"In der PDF-Datei {oai_file.id} befindet sich der Lebenslauf aus dem du die Daten extrahieren sollst",
        "mode": "extraction",
    }
    print(state_input_data)
    input_str = json.dumps(state_input_data)
    try:
        # Start the execution of the state machine
        response = statemachine_client.start_execution(
            stateMachineArn=state_machine_arn,
            input=input_str,
        )
        print(f"Started state machine execution: {response}")
        return response
    except Exception as e:
        print(f"Error starting state machine execution: {e}")
        raise e


def initialize_database_entry(openai_file, vectorstore, upload_id):
    print(f"Upload ID: {upload_id}")
    data = {
        "process_status": "in_progress",
        "upload_id": str(upload_id),
        "vector_store_id": vectorstore.id,
        "file_id": openai_file.id,
    }
    database_client = boto3.resource("dynamodb")
    table = database_client.Table("cv_uploads")
    try:
        table.put_item(Item=data)
    except Exception as e:
        raise e
    return str(upload_id)


def upload_file_to_openai(pdf):
    client = OpenAI()
    try:
        expiration_time = (
            int(time.time()) + 1800
        )  # Current time + 30 Minutes (in seconds)
        file = client.files.create(file=pdf, purpose="assistants")
    except Exception as e:
        raise e
    print(f"File {file.id} uploaded to OpenAI API")
    return file


def upload_file_to_openai_v2(pdf, upload_id):
    client = OpenAI()
    try:
        file = client.files.create(
            file=open("/tmp/temp_cv.pdf", "rb"), purpose="assistants"
        )
        vector_store = client.beta.vector_stores.create(
            name=str(upload_id),
            file_ids=[file.id],
            expires_after={"anchor": "last_active_at", "days": 1},
        )
        print(vector_store.id)
    except Exception as e:
        raise e
    print(f"File {file.id} uploaded to OpenAI API")
    return file, vector_store


def upload_pdf_to_s3(pdf_bytesio, bucket_name, object_key):
    s3_client = boto3.client("s3")
    s3_client.upload_fileobj(pdf_bytesio, bucket_name, object_key)


def extract_pdf_from_formdata(body64, content_type, upload_id):
    # Decode Body
    form_data = base64.b64decode(body64)
    raw_data = form_data.decode("iso-8859-1")

    headers = f"content-type: {content_type}"
    full_message = headers + "\n\n" + raw_data
    pdf = io.BytesIO()

    # Use the email parser to parse the multipart data
    parser = Parser()
    msg = parser.parsestr(full_message)

    # Initialize an empty bytes object to hold the PDF content
    pdf_content = b""

    # Iterate through the parts to find the PDF file part
    if msg.is_multipart():
        for part in msg.walk():
            content_disposition = part.get("Content-Disposition", None)
            if content_disposition and "filename" in content_disposition:
                # Assuming the part with a filename is your PDF
                pdf_content = part.get_payload(decode=True)  # decode=True to get bytes
                break

    # Write the PDF content to a file, if we found any
    if pdf_content:
        temp_pdf_path = "/tmp/temp_cv.pdf"
        with open(temp_pdf_path, "wb") as pdf_file:
            pdf_file.write(pdf_content)
        print(f"PDF file extracted and saved to {temp_pdf_path}.")
        s3_client = boto3.client("s3")
        s3_client.upload_file(temp_pdf_path, "cv-uploaded-resumes", upload_id)
        return pdf_content
    else:
        print("No PDF file found in the multipart data.")
        raise Exception("No PDF in Request")


def generate_response(statusCode, body):
    return {
        "statusCode": statusCode,
        "body": json.dumps(body),
        "headers": {
            "Access-Control-Allow-Origin": "*",
            "Access-Control-Allow-Credentials": True,
            "Access-Control-Allow-Methods": "GET, POST, OPTIONS",
            "Access-Control-Allow-Headers": "Content-Type,X-Amz-Date,Authorization,X-Api-Key,X-Amz-Security-Token, debug_mode",
        },
    }
