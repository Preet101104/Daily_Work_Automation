import os
import json
from dotenv import load_dotenv
from google import genai
from google.genai import types
from pydantic import BaseModel, Field
from typing import List
from docxtpl import DocxTemplate

# Load environment variables from .env file
load_dotenv()

class NewsDetail(BaseModel):
    title: str = Field(description="News headline")
    content: str = Field(description="Short paragraph explaining the news")

class FinancialUpdates(BaseModel):
    date: str = Field(description="The date of the news or current date")
    executive_overview: List[str] = Field(
        description="List of bullet points summarizing the overall deal activity, institutional capital, and macro repositioning context based on the news."
    )
    details: List[NewsDetail] = Field(
        description="The detailed news items for M&A and ECM."
    )

def create_word_document(updates: FinancialUpdates, template_filename="template.docx", output_filename="newsletter_output.docx"):
    """Creates a styled Word Document using docxtpl."""
    try:
        doc = DocxTemplate(template_filename)
        context = updates.model_dump()
        doc.render(context)
        doc.save(output_filename)
        print(f"Document successfully saved to {output_filename}")
    except Exception as e:
        print(f"Error rendering template: {e}")

def extract_financial_updates(pdf_path: str):
    """
    Analyzes a newspaper PDF using Gemini 2.5 Flash to extract
    M&A and ECM updates.
    """
    api_key = os.getenv("GEMINI_API_KEY")
    if not api_key or api_key == "your_api_key_here":
        print("Error: Please set your GEMINI_API_KEY in the .env file.")
        return

    # Initialize the Gemini GenAI client
    client = genai.Client()

    print(f"Uploading {pdf_path} to Gemini...")
    try:
        # Upload the PDF file directly to Gemini's File API
        uploaded_file = client.files.upload(file=pdf_path)
    except Exception as e:
        print(f"Failed to upload the file: {e}")
        return

    prompt = """
    You are an expert financial analyst. Please carefully read the provided newspaper PDF.
    Identify and extract news and updates strictly related to:
    1. Mergers and Acquisitions (M&A)
    2. Equity Capital Markets (ECM)
    """

    print("Analyzing the document with Gemini 2.5 Flash... This may take a few moments.")
    try:
        response = client.models.generate_content(
            model='gemini-2.5-flash',
            contents=[uploaded_file, prompt],
            config=types.GenerateContentConfig(
                response_mime_type="application/json",
                response_schema=FinancialUpdates,
            ),
        )
        
        # Parse output into pydantic model
        structured_data = FinancialUpdates.model_validate_json(response.text)
        
        # Pass to the docxtpl generator
        create_word_document(structured_data, template_filename="template.docx", output_filename="newsletter_output.docx")
        
    except Exception as e:
        print(f"Error generating content: {e}")


if __name__ == "__main__":
    # Ensure you have a sample PDF to test with
    sample_pdf = "newspaper.pdf"
    
    if os.path.exists(sample_pdf):
        extract_financial_updates(sample_pdf)
    else:
        print(f"Please place a PDF named '{sample_pdf}' in the current directory,")
        print("or update the 'sample_pdf' variable in this script with your PDF path.")
