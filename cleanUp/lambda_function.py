import json
import boto3
from openai import OpenAI


def lambda_handler(event, context):
    upload_id = ""
    file_id = ""
    thread_id = ""
    stage = 0  # 0: File not uploaded to OpenAI yet, 1:File uploaded to OpenAI, 2:Database entry initialized, 3:Created Thread on OpenAI API
    openai_client = OpenAI()
    database_client = boto3.resource("dynamodb")
    table = database_client.Table("cv_uploads")
    item_key = {"upload_id": upload_id}

    if stage > 0:
        # Delete the OpenAi API File
        try:
            openai_client.files.delete(file_id)
        except:
            raise

        if stage > 1:
            # Delete the DynamoDB entry
            try:
                table.delete_item(Key=item_key)
            except:
                raise

            if stage > 2:
                # Delete the OpenAi API Thread
                try:
                    openai_client.beta.threads.delete(thread_id)
                except:
                    raise

    return {"statusCode": 200, "body": json.dumps("Everything is cleaned up")}
