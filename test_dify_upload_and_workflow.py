import requests
import json
import os
from dotenv import load_dotenv

load_dotenv()

# Configuration
API_KEY = os.environ["DIFY_API_KEY"]
WORKFLOW_ID = os.environ["DIFY_WORKFLOW_ID"]
BASE_URL = os.environ.get("DIFY_BASE_URL", "https://ai-dev.adata.com")
FILE_UPLOAD_URL = f"{BASE_URL}/v1/files/upload"
WORKFLOW_RUN_URL = f"{BASE_URL}/v1/workflows/{WORKFLOW_ID}/run"

# File path
FILE_PATH = "scream-cat.png"
USER_ID = os.environ.get("DIFY_USER_ID", "test_user1")

def upload_file(file_path):
    """Upload file using the file upload API"""
    print(f"Uploading file: {file_path}")
    
    if not os.path.exists(file_path):
        print(f"Error: File not found: {file_path}")
        return None
    
    headers = {
        "Authorization": f"Bearer {API_KEY}"
    }
    
    # Prepare the file for upload
    with open(file_path, 'rb') as f:
        files = {
            'file': (os.path.basename(file_path), f, 'image/png'),
            'user': (None, USER_ID)
        }
        
        try:
            response = requests.post(FILE_UPLOAD_URL, headers=headers, files=files)
            response.raise_for_status()
            
            result = response.json()
            print("File uploaded successfully!")
            print(json.dumps(result, indent=2, ensure_ascii=False))
            
            # Extract file_id from response
            file_id = result.get('id')
            return file_id
            
        except requests.exceptions.RequestException as e:
            print(f"Error uploading file: {e}")
            if hasattr(e.response, 'text'):
                print(f"Response: {e.response.text}")
            return None

def call_workflow_with_file(file_id):
    """Call dify workflow with the uploaded file"""
    print(f"\nCalling workflow with file_id: {file_id}")

    headers = {
        "Authorization": f"Bearer {API_KEY}",
        "Content-Type": "application/json"
    }

    # Request body - include the file_id in inputs
    payload = {
        "inputs": {
            "img": 
                {
                    "transfer_method": "local_file",
                    "upload_file_id": file_id,
                    "type": "image",
                }

        },
        "response_mode": "blocking",
        "user": USER_ID,
    }

    try:
        response = requests.post(WORKFLOW_RUN_URL, headers=headers, json=payload)
        response.raise_for_status()

        result = response.json()
        print("Workflow executed successfully!")
        print(json.dumps(result, indent=2, ensure_ascii=False))

        return result

    except requests.exceptions.RequestException as e:
        print(f"Error calling workflow: {e}")
        if hasattr(e.response, 'text'):
            print(f"Response: {e.response.text}")
        return None

def main():
    """Main function: Upload file, then call workflow"""
    print("=" * 60)
    print("Step 1: Upload img file")
    print("=" * 60)
    
    file_id = upload_file(FILE_PATH)
    
    if file_id:
        print("\n" + "=" * 60)
        print("Step 2: Call Dify Workflow")
        print("=" * 60)
        
        call_workflow_with_file(file_id)
    else:
        print("File upload failed. Workflow not called.")

if __name__ == "__main__":
    main()
