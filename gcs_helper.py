import os
import datetime
from google.cloud import storage
import logging

logger = logging.getLogger(__name__)

def get_gcs_client():
    try:
        return storage.Client()
    except Exception as e:
        logger.warning(f"Failed to initialize GCS client: {e}")
        return None

def upload_to_gcs(local_file_path, destination_blob_name):
    bucket_name = os.getenv('GCS_BUCKET_NAME')
    if not bucket_name:
        return False
        
    client = get_gcs_client()
    if not client:
        return False
        
    try:
        bucket = client.bucket(bucket_name)
        blob = bucket.blob(destination_blob_name)
        blob.upload_from_filename(local_file_path)
        logger.info(f"Uploaded {local_file_path} to gs://{bucket_name}/{destination_blob_name}")
        return True
    except Exception as e:
        logger.error(f"Error uploading to GCS: {e}")
        return False

def generate_download_signed_url_v4(blob_name, expiration_minutes=60):
    bucket_name = os.getenv('GCS_BUCKET_NAME')
    if not bucket_name:
        return None
        
    client = get_gcs_client()
    if not client:
        return None
        
    try:
        bucket = client.bucket(bucket_name)
        blob = bucket.blob(blob_name)
        
        # Check if blob exists
        if not blob.exists():
            return None
            
        return f"https://storage.googleapis.com/{bucket_name}/{blob_name}"
    except Exception as e:
        logger.error(f"Failed to generate public url: {e}")
        return None
