import os
import pandas as pd
from pathlib import Path
import json
from openpyxl import load_workbook
from google import genai
class ProposalExtractorGemini:
    def __init__(self, api_key, use_pro_model=True):
        """Initialize with Gemini API key
        
        Args:
            api_key: Your Gemini API key
            use_pro_model: If True, uses gemini-2.5-pro (highest accuracy, slower)
                          If False, uses gemini-2.5-flash (fast, good accuracy)
        """
        self.client = genai.Client(api_key=api_key)
        self.model_name = 'gemini-2.5-pro' if use_pro_model else 'gemini-2.5-flash'
        print(f"Using model: {self.model_name}")
        
        # Define the extraction prompt
        self.extraction_prompt = """
You are analyzing a proposal PDF for quantum technologies lab setup. Extract the following information and respond ONLY with a JSON object.
THIS IS A VERY IMPORTANT TASK. FOLLOW THE INSTRUCTIONS CAREFULLY.
IMPORTANT OUTPUT REQUIREMENTS:
- For Yes/No fields: Return ONLY "Yes" or "No" or "Not Mentioned"(if you cannot find the information). In few case Yes/No may not be given explicitly instead might be in checkbox do this check if yes/no is not mentioned in the field.
- For text/qualitative fields: Extract the actual text/numbers from the document
- For ranking fields: Extract only the rank number (e.g., "45" not "Rank 45")

Extract these criteria:

COLUMN L: Central Government funded/aided institution - Yes/No only
COLUMN M: State Government funded/aided institution - Yes/No only  
COLUMN N: Other funding sources - Extract text describing other fundings that the institue is receiving(Not the funding for project), or write "None"
COLUMN O: Institution of Eminence by AICTE/UGC - Yes/No only
COLUMN P: NIRF 8th Edition rank number - Extract only the number (e.g., "45") or "Not Ranked"
COLUMN Q: QS World Ranking Asia 2024 rank number - Extract only the number (e.g., "120") or "Not Ranked"
COLUMN R: NBA accreditation for 30%+ courses (valid till April 2025) - Yes/No only
COLUMN S: NAAC score - Extract actual score number (e.g., "3.25") or "Not Available"
COLUMN T: Autonomous Status by UGC/AICTE(some propsal have not mentioned expicitly in that case check AICTE/UGC/Other Accreditation Status) - Return "UGC" or "AICTE" or "Both" or "None"
COLUMN U: More than 80% admission for last 5 years.(If not provided explicitly return Not Mentioned) - Yes/No only
COLUMN V: Two faculty members specified (1/3 time dedicated) - Yes/No only
COLUMN W: Lab space of at least 2000 sq.ft specified - Yes/No only
COLUMN X: Full-time lab technician specified - Yes/No only
COLUMN Y: Has the proposal specified that the institution has received Senate/Governing Body or Main Board of Institute approval for launching the UG Minor Programme in Quantum Technologies? - Yes/No only
COLUMN Z: Has the proposal submitted a formal written commitment indicating intent to start the programme upon obtaining the necessary approval?  - Yes/No only
COLUMN AA: Mention the page number of supporting document for COLUMN Y with comment on why yes/no(within 1 line)- Page Number(Comments) only
COLUMN AB: Mention the page number of supporting document for COLUMN Z with comment on why yes/no(within 1 line)- Page Number(Comments) only

Respond ONLY with this exact JSON format (no additional text):
{
    "col_l": "Yes or No",
    "col_m": "Yes or No",
    "col_n": "Text or None",
    "col_o": "Yes or No",
    "col_p": "Number or Not Ranked",
    "col_q": "Number or Not Ranked",
    "col_r": "Yes or No",
    "col_s": "Number or Not Available",
    "col_t": "UGC or AICTE or Both or None",
    "col_u": "Yes or No",
    "col_v": "Yes or No",
    "col_w": "Yes or No",
    "col_x": "Yes or No",
    "col_y": "Yes or No",
    "col_z": "Yes or No"
    "col_aa": "123(comment)"
    "col_ab": "123(comment)"

}
"""
    
    def extract_from_pdf(self, pdf_path, max_retries=2):
        """Extract criteria from PDF using Gemini with validation
            This function uploads the PDF, sends it to Gemini for analysis,
            and retries if there are transient errors or JSON parsing issues.
        
        Args:
            pdf_path: Path to PDF file
            max_retries: Number of retries if extraction fails validation
        """
        import time
        from random import random

        print(f"  Uploading and analyzing PDF...")

        # Retry upload in case of transient network issues
        upload_attempts = 0
        pdf_file = None
        while upload_attempts <= max_retries:
            try:
                pdf_file = self.client.files.upload(file=pdf_path)
                break
            except Exception as e:
                upload_attempts += 1
                wait = (2 ** upload_attempts) + random()
                print(f"  ⚠ Upload failed (attempt {upload_attempts}/{max_retries}). Error: {e}. Retrying in {wait:.1f}s...")
                time.sleep(wait)

        if pdf_file is None:
            print(f"  ⚠ Failed to upload PDF after {max_retries} attempts.")
            return None

        last_exception = None
        for attempt in range(max_retries + 1):
            try:
                response = self.client.models.generate_content(
                    model=self.model_name,
                    contents=[pdf_file, self.extraction_prompt]
                    )

                # Provide robust extraction of text portion
                response_text = getattr(response, 'text', '')
                if not response_text and hasattr(response, 'content'):
                    response_text = str(response.content)
                response_text = response_text.strip()

                # Remove markdown code blocks if present
                if response_text.startswith('```json'):
                    response_text = response_text.split('```json', 1)[1]
                if response_text.startswith('```'):
                    response_text = response_text.split('```', 1)[1]
                if response_text.endswith('```'):
                    response_text = response_text.rsplit('```', 1)[0]

                response_text = response_text.strip()

                # Parse JSON
                data = json.loads(response_text)

                # Clean up: delete the uploaded file from Gemini
                try:
                    pdf_file.delete()
                except Exception:
                    pass

                return data

            except json.JSONDecodeError as e:
                last_exception = e
                print(f"  ⚠ Error parsing JSON response: {str(e)}")
                snippet = (response_text[:200] + '...') if 'response_text' in locals() else ''
                print(f"  Raw response (truncated): {snippet}")
            except Exception as e:
                last_exception = e
                # Inspect message to detect typical timeout codes
                err_msg = str(e)
                print(f"  ⚠ Error during analysis: {err_msg}")
                if attempt < max_retries:
                    wait = (2 ** (attempt + 1)) + random()
                    print(f"  Retrying (attempt {attempt +1}/{max_retries + 1}) in {wait:.1f}s...")
                    time.sleep(wait)
                    continue
                else:
                    try:
                        pdf_file.delete()
                    except Exception:
                        pass
                    print(f"  ⚠ Final failure after {max_retries + 1} attempts: {last_exception}")
                    return None
    
    def update_excel_with_proposals(self, excel_path, pdf_folder, output_path=None):
        """Update an Excel file using TPN values from the sheet to find matching PDFs.

        The Excel is expected to have headers on row 5 (header=4 for pandas).
        """
        if output_path is None:
            output_path = excel_path.replace('.xlsx', '_updated.xlsx')

        print(f"Loading Excel file: {excel_path}")
        wb = load_workbook(excel_path)
        ws = wb.active

        df = pd.read_excel(excel_path, header=4)
        print(f"Found {len(df)} rows in Excel")

        pdf_folder_path = Path(pdf_folder)
        pdf_files = list(pdf_folder_path.glob('*/*.pdf'))
        pdf_files += list(pdf_folder_path.glob(f'*/*rop*/*.pdf')) # Include nested proposal folders

        if not pdf_files:
            print(f"No PDF files found in {pdf_folder}")
            return

        print(f"Found {len(pdf_files)} PDF files to process\n")

        # Build lookup map from TPN to pdf paths
        pdf_map = {}
        for pf in pdf_files:            
            tpn = pf.parent.__str__().split('/')[1]  
            if tpn:
                pdf_map.setdefault(str(tpn), []).append(pf)
            else:
                pdf_map.setdefault(pf.name, []).append(pf)
        for row_idx, row in df.iterrows():
            excel_row = row_idx + 6
            if not pd.isna(row.iloc[20]): # To perform check whether the row is already filled
                
                print(f"Row {excel_row}: already filled, skipping.")
                continue

            tpn_str = str(row.get('TPN No.'))
            matching_pdfs = pdf_map.get(tpn_str)
            if not matching_pdfs: # Try partial match
                matches = [pf for pf in pdf_files if tpn_str in pf.name]
                if matches:
                    matching_pdfs = matches

            if not matching_pdfs:
                print(f"  ⚠ Could not find PDF for TPN {tpn_str} (Excel row {excel_row})")
                continue
            print(matching_pdfs)
            
            pdf_file = matching_pdfs[0]
            print(f"Processing Excel row {excel_row} - TPN {tpn_str}: {pdf_file.name}")

            extracted_data = self.extract_from_pdf(pdf_file)

            if extracted_data:
                # Update columns L..AB
                for col in range(11, 28):  # Columns L (12) to AB (28)
                    col_letter = chr(ord('A') + col)
                    if col>25:
                        col_letter = 'A' + chr(ord('A') + (col - 26))
                    ws[f'{col_letter}{excel_row}'] = extracted_data.get(f'col_{col_letter.lower()}', '-')
                print(f"  ✓ Updated row {excel_row} in Excel, TPN: {tpn_str}")
            else:
                print(f"  ⚠ Could not extract data from {pdf_file.name}")

            wb.save(output_path)
            print()

        print(f"\n{'='*60}")
        print(f"✓ Excel file updated and saved to: {output_path}")
        print(f"{'='*60}")

    

if __name__ == "__main__":
   
    API_KEY = os.getenv('GENAI_API_KEY')  # Set your Gemini API key in environment variable
    
    # Path to your existing Excel file (with headers in row 5)
    EXCEL_FILE = "./proposals_sheet.xlsx"  # Change to your Excel file path
    
    # Path to folder containing PDF proposals
    PDF_FOLDER = "./proposals"  # Change to your PDF folder path
    
    # Output file path (will create new file, won't overwrite original)
    OUTPUT_FILE = "./proposals_sheet_updated.xlsx"
    
    # Use Pro model for highest accuracy
    USE_PRO_MODEL = False
    print(API_KEY)
    print("="*60)
    print("PDF Proposal Extractor - Excel Updater")
    print("="*60)
    print()
    
    # Create extractor instance
    extractor = ProposalExtractorGemini(api_key=API_KEY, use_pro_model=USE_PRO_MODEL)
    
    # Update Excel with extracted data from PDFs
    extractor.update_excel_with_proposals(
        excel_path=EXCEL_FILE,
        pdf_folder=PDF_FOLDER,
        output_path=OUTPUT_FILE
    )