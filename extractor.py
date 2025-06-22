import pdfplumber
import json
import re
import pytesseract
from PIL import Image
import io
import fitz
import subprocess

## setting the path for the tesseract
pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'

# creating the function for the ocr on pli image
def perform_ocr(pil_image):
    """Perform OCR on a PIL image"""
    try:
        page_text = pytesseract.image_to_string(  ## we can use this image to extract text from the image
            pil_image,
            lang='eng',
            config='--psm 6 --oem 3 -c tessedit_char_whitelist=ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789@.,&()-/'
        )
        return page_text
    ## if it fails, we can return an empty string
    except Exception as e:
        print("Error OCR:", e)
        return ""

## implementing a method to read the pdf data through ocr scan
def read_pdf(pdf_path):
    """Extract text from PDF file using OCR scan"""
    try:
        doc = fitz.open(pdf_path)
        # we can initialize an empty string to store all the text data
        text = ""

        for page_num in range(len(doc)):
            page = doc[page_num]
            page_text = page.get_text() or ""

            # should the text can not be extracted, use ocr
            # now we can convert pdf page to image
            if not page_text.strip():
                pix = page.get_pixmap()
                image_data = pix.tobytes("png")
                pil_image = Image.open(io.BytesIO(image_data))
                page_text = perform_ocr(pil_image)

            if page_text:
                text += page_text + "\n"

        doc.close()
        return text
    except Exception as e:
        print("Error reading PDF: ", e)
        return None

## we can create a method to extract data from the form using regex pattern
def extract_form(pdf_text):
    """Extract the data from Form ADT-1 using regex patterns"""
    data = {
        "company_name": "",
        "cin": "",
        "registered_office": "",
        "appointment_date": "",
        "auditor_name": "",
        "auditor_address": "",
        "auditor_frn_or_membership": "",
        "appointment_type": ""
    }

    # cin format for all the 21 characters
    cin_number = re.search(r'\b[A-Z0-9]{21}\b', pdf_text)
    if cin_number:
        data["cin"] = cin_number.group()

    # extracting the company name from the pdf
    company_name_found = re.search(r'\b[A-Za-z\s&.,]+PRIVATE LIMITED\b', pdf_text)
    if company_name_found:
        data["company_name"] = company_name_found.group().strip()
    else:
        company_name_found = re.search(r'\b[A-Za-z\s&.,]+(LIMITED|LTD|PVT)\b', pdf_text) # checking for other compnay types
        if company_name_found:
            data["company_name"] = company_name_found.group().strip()

    # registered office name extracting
    if data["company_name"]:
        # we can simply find the company name from the text
        company_index = pdf_text.find(data["company_name"])
        # we need to make sure to get text that comes after the company name
        if company_index != -1:
            after_company = pdf_text[company_index + len(data["company_name"]):]
            ## address patterns 6 digit pin locate that from the text
            address = re.search(r'([A-Za-z0-9,\s\-]+)\s+(\d{6})', after_company)
            if address:
                data["registered_office"] = f"{address.group(1).strip()}, {address.group(2)}"
                # combining the address and the pincode

    # appointment date
    ## formatting date dd/mm/yyyy
    appoint_date = re.search(r'\b(\d{2}/\d{2}/\d{4})\b', pdf_text)
    if appoint_date: ## if appointment date exits then
        data["appointment_date"] = appoint_date.group(1)

    # extracting the audiotr name
    ## we can first check for any partnership pattern and
    ## afterwards if it does not exits we can lookup for any firms
    ## should give us the firm or the auditor name(s)
    auditor_name = re.search(r'\b([A-Za-z\s]+&[A-Za-z\s]+)\b', pdf_text)
    if auditor_name:
        data["auditor_name"] = auditor_name.group(1).strip()
    else:
        auditor_name = re.search(r'\b([A-Z][A-Za-z\s]+(ASSOCIATES|PARTNERS|& CO))\b', pdf_text)
        if auditor_name:
            data["auditor_name"] = auditor_name.group(1).strip()

    # locate the auditor's address
    if data["auditor_name"]:
        auditor_index = pdf_text.find(data["auditor_name"])
        ## like before, we just check for after the name
        if auditor_index != -1:
            auditor_name_after = pdf_text[auditor_index + len(data["auditor_name"]):]
            address = re.search(r'([A-Za-z0-9,\s\-]+)\s+(\d{6})', auditor_name_after)
            if address:
                data["auditor_address"] = f"{address.group(1).strip()}, {address.group(2)}"

    ## extracting the fin or membership
    auditor_frn = re.search(r'\b([A-Z0-9]{6,})\b', pdf_text)
    if auditor_frn:
        data["auditor_frn_or_membership"] = auditor_frn.group(1)

    # appoinment type
    ## appointment types can be 3 types i guess.
    ## first check for appointment/re-appointment
    ## if not found check for re-appointment or new appointment
    if re.search(r'Appointment/Re-appointment in AGM', pdf_text, re.IGNORECASE):
        data["appointment_type"] = "Appointment/Re-appointment in AGM"
    elif re.search(r'Re-appointment', pdf_text, re.IGNORECASE):
        data["appointment_type"] = "Re-appointment"
    elif re.search(r'New Appointment', pdf_text, re.IGNORECASE):
        data["appointment_type"] = "New Appointment"
    else:
        data["appointment_type"] = "Not specified"

    return data

# model summarizing using local ollama
def summarize_with_ollama(data, model="llama3.2"):
    """Uses Ollama to generate a 3-4 line summary of the extracted data."""
    # we need to check if ollama is available or not
    # check for other available models
    try:
        result = subprocess.run(['ollama', 'list'], capture_output=True, text=True)
        if result.returncode != 0:
            print("Ollama not found or not running.")
            return None

        # now we need to check if it exits or not
        available_models = result.stdout
        if model not in available_models:
            print(f"Model '{model}' not found. Trying alternatives...")
            if "llama3" in available_models:
                model = "llama3"
            elif "llama2" in available_models:
                model = "llama2"
            elif "mistral" in available_models:
                model = "mistral"
            else:
                print("No suitable model found.")
                return None

        prompt = f"""Summarize this Form ADT-1 data in exactly 3-4 lines:

{json.dumps(data, indent=2)}

Keep it concise and focus on: company name, auditor details, appointment type, and key dates."""

        print(f"Generating summary with Ollama ({model})...")

        # Call Ollama
        result = subprocess.run(
            ['ollama', 'run', model],
            input=prompt,
            capture_output=True,
            text=True,
            timeout=60
        )

        if result.returncode == 0:
            summary = result.stdout.strip()
            return summary
        else:
            print(f"Ollama error: {result.stderr}")
            return None

    except Exception as e:
        print(f"Error using Ollama: {e}")
        return None

# lastly we need to create a summary.txt file to save it
def save_summary_to_file(summary, filename="summary.txt"):
    """Saves the summary text into a .txt file"""
    try:
        with open(filename, "w", encoding="utf-8") as f:
            f.write(summary)
        print(f"Summary saved to {filename}")
    except Exception as e:
        print("Error saving summary to file:", e)

def main():
    print("Performing data extraction from Form ADT-1")
    print("\n\n")

    pdf_path = "Form ADT-1-29092023_signed.pdf"

    ## raeding the pdf, if no text exist then return empty
    pdf_text = read_pdf(pdf_path)
    if not pdf_text:
        print("Failed to read the PDF file")
        return

    print("PDF read successfully")

    ## data extracting method call
    print("Trying to extract the data")
    extracted_data = extract_form(pdf_text)

    # JSON
    print("JSON file!!")
    with open("output.json", "w", encoding="utf-8") as f:
        json.dump(extracted_data, f, indent=2, ensure_ascii=False)

    # summary from the JSON data
    print("Generating summary...")
    summary = summarize_with_ollama(extracted_data)
    if summary:
        print("\nSummary:")
        print("\n")
        print(summary)
        save_summary_to_file(summary)
    else:
        print("Failed to generate summary")

    print("\nExtracted Data:")
    print("\n\n")
    for key, value in extracted_data.items():
        print(f"{key.replace('_', ' ').title()}: {value}")

    print("\nExtraction Complete!")

if __name__ == "__main__":
    main()
