import json
from openai import OpenAI
import boto3


def lambda_handler(event, context):

    # Extract the necessary data from the event and set other variables
    thread_id = event["threadCreationOutput"]["thread_id"]
    upload_id = event["upload_id"]
    oai_file_id = event["file_id"]
    vectorstore_ids = event["vectorstore_ids"]
    state_machine_arn = (
        "arn:aws:states:eu-central-1:891376982948:stateMachine:MyStateMachine-h1zdun91x"
    )
    correction_assistant_id = "asst_uPkzE0iGVonUn6cWg3uzlCQr"
    # Retreives the latest message from the assistant and removes JSON Headers
    response = retreive_json_answer_from_assistant(thread_id)
    state_input_data = {
        "upload_id": upload_id,
        "assistant_id": correction_assistant_id,
        "file_id": oai_file_id,
        "file_ids": [oai_file_id],
        "vectorstore_ids": vectorstore_ids,
        "prompt": f"In der PDF-Datei {oai_file_id} befindet sich der Lebenslauf eines Bewerbers. Im Folgenden ist der JSON-Output des Lebenslaufs dargestellt. Korrigiere den JSON-Output.\n{response}",
        "mode": "correction",
    }
    statemachine_client = boto3.client("stepfunctions")
    input_str = json.dumps(
        state_input_data
    )  # Convert the input data (JSON Format) to a JSON string
    try:
        # Start the execution of the state machine
        response = statemachine_client.start_execution(
            stateMachineArn=state_machine_arn,
            input=input_str,
        )
        print(f"Started state machine execution: {response}")
        return json.dumps(
            {"statusCode": 200, "body": "State machine execution started"}
        )
    except Exception as e:
        print(f"Error starting state machine execution: {e}")
        return json.dumps({"statusCode": 400, "body": "Error starting state machine"})


def retreive_json_answer_from_assistant(thread_id):
    openai_client = OpenAI()
    messages = openai_client.beta.threads.messages.list(thread_id=thread_id)
    assistant_messages = [
        msg.content for msg in messages.data if msg.role == "assistant"
    ]
    response = assistant_messages[0][0].text.value
    response = response.replace("json", "", 1)
    response = response.replace("`", "")
    return response
