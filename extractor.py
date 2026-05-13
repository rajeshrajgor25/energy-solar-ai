import json
import re
import base64
import fitz  
from pathlib import Path
from PIL import Image
import io


EXTRACTION_PROMPT = """You are an expert at reading Indian electricity bills (MSEDCL/Maharashtra DISCOM).

Extract the following fields from the bill text/image provided.
Return ONLY a valid JSON object — no explanation, no markdown, no backticks.

Fields to extract:
{
  "consumer_name": "Full name of consumer",
  "consumer_number": "Consumer/account number",
  "meter_number": "Meter number",
  "billing_month": "Month and year e.g. March 2024",
  "billing_address": "Full address",
  "division": "MSEDCL division/subdivision office",
  "units_consumed": <number, kWh consumed this month>,
  "sanctioned_load": <number in kW>,
  "connected_load": <number in kW>,
  "contract_demand": <number in kVA or null if not present>,
  "tariff_category": "e.g. LT-I Residential, LT-II Commercial, HT etc.",
  "per_unit_rate": <average rate in ₹/kWh as float>,
  "fixed_charges": <monthly fixed/demand charges in ₹ as float>,
  "total_monthly_bill": <total bill amount in ₹ as float>
}

Rules:
- Use null for any field not found in the bill
- units_consumed: look for 'Units', 'kWh', 'Energy Charges', present reading minus previous reading
- sanctioned_load: look for 'Sanctioned Load', 'Contract Demand', 'Connected Load' in kW
- per_unit_rate: calculate as (energy charges) / (units consumed) if not directly stated
- total_monthly_bill: the final payable amount
- Return only the JSON object, nothing else
"""


def extract_text_from_pdf(file_bytes: bytes) -> str:
    """Extract text from PDF using PyMuPDF."""
    doc = fitz.open(stream=file_bytes, filetype="pdf")
    text = ""
    for page in doc:
        text += page.get_text()
    doc.close()
    return text


def extract_text_from_image(file_bytes: bytes) -> str:
    """Convert image to base64 for vision model."""
    return base64.b64encode(file_bytes).decode("utf-8")


def get_bill_data_from_groq(content: str, is_image: bool, api_key: str, mime_type: str = "image/jpeg") -> dict:
    """Send bill content to Groq and get structured JSON back."""
    from groq import Groq
    
    client = Groq(api_key=api_key)
    
    if is_image:
        messages = [
            {
                "role": "user",
                "content": [
                    {
                        "type": "image_url",
                        "image_url": {
                            "url": f"data:{mime_type};base64,{content}"
                        }
                    },
                    {
                        "type": "text",
                        "text": EXTRACTION_PROMPT
                    }
                ]
            }
        ]
        model = "meta-llama/llama-4-scout-17b-16e-instruct"
    else:
        # Use text model for extracted PDF text
        messages = [
            {
                "role": "user",
                "content": f"{EXTRACTION_PROMPT}\n\nBILL TEXT:\n{content[:8000]}"
            }
        ]
        model = "llama-3.3-70b-versatile"
    
    response = client.chat.completions.create(
        model=model,
        messages=messages,
        temperature=0,
        max_tokens=1024,
    )
    
    raw = response.choices[0].message.content.strip()
    
    # Strip markdown fences if present
    raw = re.sub(r"```json\s*", "", raw)
    raw = re.sub(r"```\s*", "", raw)
    raw = raw.strip()
    
    try:
        return json.loads(raw)
    except json.JSONDecodeError:
        match = re.search(r'\{.*\}', raw, re.DOTALL)
        if match:
            return json.loads(match.group())
        raise ValueError(f"Could not parse JSON from Groq response:\n{raw}")


def extract_bill_data(file_bytes: bytes, filename: str, api_key: str) -> dict:
    """
    Main extraction function.
    Accepts file bytes + filename, returns structured dict.
    """
    filename_lower = filename.lower()
    
    if filename_lower.endswith(".pdf"):
        text = extract_text_from_pdf(file_bytes)
        if len(text.strip()) > 100:
            data = get_bill_data_from_groq(text, is_image=False, api_key=api_key)
        else:
            doc = fitz.open(stream=file_bytes, filetype="pdf")
            page = doc[0]
            mat = fitz.Matrix(2, 2) 
            pix = page.get_pixmap(matrix=mat)
            img_bytes = pix.tobytes("jpeg")
            doc.close()
            b64 = base64.b64encode(img_bytes).decode("utf-8")
            data = get_bill_data_from_groq(b64, is_image=True, api_key=api_key, mime_type="image/jpeg")
    
    elif filename_lower.endswith((".png", ".jpg", ".jpeg", ".webp")):
        mime_map = {".png": "image/png", ".jpg": "image/jpeg",
                    ".jpeg": "image/jpeg", ".webp": "image/webp"}
        ext = Path(filename_lower).suffix
        mime = mime_map.get(ext, "image/jpeg")
        b64 = extract_text_from_image(file_bytes)
        data = get_bill_data_from_groq(b64, is_image=True, api_key=api_key, mime_type=mime)
    
    else:
        raise ValueError(f"Unsupported file type: {filename}")
    
    return data


if __name__ == "__main__":
    # Quick test
    import os
    key = os.environ.get("GROQ_API_KEY", "")
    print("Extractor module loaded OK")
