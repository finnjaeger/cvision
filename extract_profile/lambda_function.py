import fitz
import cv2
import boto3


def save_profile_picture_to_s3(upload_id):
    print("Extracting profile pictures from CV")
    image_paths = extract_images_from_pdf("/tmp/temp_cv.pdf")
    print("Extracted profile pictures from CV: ", image_paths)
    profile_pictures = [img for img in image_paths if is_profile_picture(img)]
    if len(profile_pictures) > 0:  # Checking if there is at least one profile picture
        s3_client = boto3.client("s3")
        bucket_name = "cv-profile-pictures"

        # Get the first profile picture
        first_profile_picture = profile_pictures[0]

        # Upload the first profile picture
        s3_client.upload_file(first_profile_picture, bucket_name, f"{upload_id}.png")
        print(f"Uploaded {upload_id} to S3 bucket {bucket_name}")


def extract_images_from_pdf(pdf_path):
    doc = fitz.open(pdf_path)
    images = []
    images_filenames = []
    for i in range(len(doc)):  # Loop through each page
        for img_index, img in enumerate(doc.get_page_images(i)):
            xref = img[0]
            base_image = doc.extract_image(xref)
            image_bytes = base_image["image"]
            images.append(image_bytes)  # Store the image bytes in list if needed

            # Create a filename and save the image
            image_filename = f"/tmp/image_page_{i}_index_{img_index}_xref_{xref}.png"
            with open(image_filename, "wb") as img_file:
                img_file.write(image_bytes)
            images_filenames.append(image_filename)
            print(
                f"Saved: {image_filename}"
            )  # Optional: print out filename of saved image
    doc.close()
    return images_filenames


def is_profile_picture(image_path):
    # Load the Haar Cascade for face detection
    face_cascade = cv2.CascadeClassifier(
        cv2.data.haarcascades + "haarcascade_frontalface_default.xml"
    )

    # Read the image
    image = cv2.imread(image_path)
    gray = cv2.cvtColor(image, cv2.COLOR_BGR2GRAY)  # Convert to grayscale

    # Detect faces
    faces = face_cascade.detectMultiScale(
        gray, scaleFactor=1.1, minNeighbors=5, minSize=(30, 30)
    )

    # If faces are detected, this could be a profile picture
    return len(faces) > 0


def download_pdf_to_tmp(bucket_name, object_key, local_path):
    s3_client = boto3.client("s3")
    with open(local_path, "wb") as file:
        s3_client.download_fileobj(bucket_name, object_key, file)


def lambda_handler(event, context):
    upload_id = event["Records"][0]["s3"]["object"]["key"]
    bucket_name = "cv-uploaded-resumes"
    path = "/tmp/temp_cv.pdf"
    print("Starting CV Download from S3")
    download_pdf_to_tmp(bucket_name, upload_id, path)
    print("Downloaded CV from S3")
    save_profile_picture_to_s3(upload_id)
    print("Done!")
    return {"statusCode": 200, "body": "Profile picture extracted and uploaded to S3"}
