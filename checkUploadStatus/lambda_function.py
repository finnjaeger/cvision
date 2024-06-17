import json
import boto3


def lambda_handler(event, context):
    query_string_params = event.get("queryStringParameters", {})
    upload_id = query_string_params.get("upload_id", "Default Value")
    print(upload_id)
    database_client = boto3.resource("dynamodb")
    table = database_client.Table("cv_uploads")

    try:
        response = table.get_item(Key={"upload_id": upload_id})
    except:
        raise

    status = response.get("Item", {}).get("process_status", "failed")
    data = response.get("Item", {}).get("cv_data", "No Data Yet")
    if status == "failed":
        data = "Upload Failed"
    server_response = {"process_status": status, "data": data}
    return {
        "statusCode": 200,
        "body": json.dumps(server_response),
        "headers": {
            "Access-Control-Allow-Origin": "*",
            "Access-Control-Allow-Credentials": True,
            "Access-Control-Allow-Methods": "GET, POST, OPTIONS",
            "Access-Control-Allow-Headers": "Content-Type,X-Amz-Date,Authorization,X-Api-Key,X-Amz-Security-Token",
        },
    }
