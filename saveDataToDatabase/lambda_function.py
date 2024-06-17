import json
import boto3
from openai import OpenAI


def lambda_handler(event, context):
    thread_id = event["threadCreationOutput"]["thread_id"]
    database_client = boto3.resource("dynamodb")
    table = database_client.Table("cv_uploads")
    answer = retreive_json_answer_from_assistant(thread_id)
    try:
        print(answer)
        cv_data = json.loads(answer)
        print(cv_data)
        table.update_item(
            Key={"upload_id": event["upload_id"]},
            UpdateExpression="SET cv_data = :cvData , process_status = :sta",
            ExpressionAttributeValues={
                ":cvData": cv_data,
                ":sta": "ready_to_retrieve",
            },
            ReturnValues="UPDATED_NEW",
        )
    except Exception as e:
        print(e)
        table.update_item(
            Key={"upload_id": event["upload_id"]},
            UpdateExpression="SET process_status = :sta",
            ExpressionAttributeValues={
                ":sta": "failed",
            },
            ReturnValues="UPDATED_NEW",
        )
        return {"statusCode": 400, "body": json.dumps("Error saving data to database")}
    return {"statusCode": 200, "body": json.dumps("Data saved to database")}


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
