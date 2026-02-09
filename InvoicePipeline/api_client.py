import requests
import time
import json
from pathlib import Path
from datetime import datetime
from config import (
    API_BASE, HEADERS, CUSTOMER_ID, ASSET_ID,
    JSON_RESULTS_DIR, INITIAL_WAIT, MAX_WAIT_MINUTES, POLL_INTERVAL
)

def extract_invoice(pdf_path: Path) -> Path:
    """
    Complete pipeline: Submit PDF, poll for result, save to JSON file
    Returns path to the saved JSON file
    """
    print(f"📄 Processing: {pdf_path.name}")
    
    # Step 1: Submit job
    print("🔄 Submitting to API...")
    job_id = submit_job(pdf_path)
    
    # Step 2: Poll for result
    print("⏳ Waiting for extraction...")
    result_data = poll_job(job_id)
    
    # Step 3: Save JSON
    print("💾 Saving results...")
    json_path = save_json_result(result_data, pdf_path)
    
    print(f"✅ JSON saved: {json_path}")
    return json_path


def submit_job(pdf_path: Path) -> str:
    """Submit PDF for extraction and return job ID"""
    url = f"{API_BASE}/extract"
    
    payload = {
        "customer_id": CUSTOMER_ID,
        "asset_id": ASSET_ID
    }
    
    with open(pdf_path, "rb") as f:
        files = [
            ("files", (pdf_path.name, f, "application/pdf"))
        ]
        
        response = requests.post(
            url,
            headers=HEADERS,
            data=payload,
            files=files,
            timeout=60
        )
    
    response.raise_for_status()
    data = response.json()
    
    job_id = data.get("job_id") or data.get("id")
    if not job_id:
        raise ValueError(f"Job ID missing in response: {data}")
    
    return job_id


def poll_job(job_id: str) -> dict:
    """Poll job status until completion"""
    url = f"{API_BASE}/status/{job_id}"
    
    print(f"⏳ Waiting {INITIAL_WAIT} seconds before first check...")
    time.sleep(INITIAL_WAIT)
    
    start_time = time.time()
    max_wait_seconds = MAX_WAIT_MINUTES * 60
    poll_count = 0
    
    while True:
        poll_count += 1
        print(f"🔄 Poll #{poll_count}...")
        
        response = requests.get(url, headers=HEADERS, timeout=30)
        
        if response.status_code != 200:
            raise RuntimeError(f"API failed [{response.status_code}]: {response.text}")
        
        data = response.json()
        status = data.get("status")
        
        if status == "SUCCESS":
            print(f"✅ Extraction complete after {poll_count} polls")
            return data
        elif status == "FAILED":
            raise RuntimeError(f"❌ Job failed: {data}")
        
        # Timeout check
        if time.time() - start_time > max_wait_seconds:
            raise TimeoutError(f"⏰ Timeout after {MAX_WAIT_MINUTES} minutes")
        
        time.sleep(POLL_INTERVAL)
        

def save_json_result(result_data: dict, pdf_path: Path) -> Path:
    """Save extraction result to JSON file"""
    # Generate filename
    pdf_name = pdf_path.stem.replace(" ", "_")
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    json_filename = f"{pdf_name}_{timestamp}.json"
    json_path = JSON_RESULTS_DIR / json_filename
    
    # Save JSON
    with open(json_path, "w", encoding="utf-8") as f:
        json.dump(result_data, f, indent=2, ensure_ascii=False)
    
    return json_path

