import os
from dotenv import load_dotenv
from google import genai
from google.genai import types
from pydantic import BaseModel, Field
from typing import List
from docxtpl import DocxTemplate

# Load environment variables
load_dotenv()

# -----------------------------
# DATA SCHEMA
# -----------------------------

class NewsDetail(BaseModel):
    title: str = Field(
        description="Short professional headline (max 12 words)"
    )
    content: str = Field(
        description="Two sentence summary explaining the transaction or event (40–60 words)"
    )


class FinancialUpdates(BaseModel):
    date: str = Field(
        description="Newsletter date"
    )

    executive_overview: List[str] = Field(
        description="4–6 concise bullet points summarizing key M&A and ECM trends"
    )

    details: List[NewsDetail] = Field(
        description="6–10 news items related to deals, IPOs, acquisitions, fundraises, or regulatory developments"
    )


# -----------------------------
# WORD DOCUMENT GENERATOR
# -----------------------------

def create_word_document(
    updates: FinancialUpdates,
    template_filename="template.docx",
    output_filename="newsletter_output.docx"
):
    """Render the Word newsletter using docxtpl."""

    try:
        doc = DocxTemplate(template_filename)

        context = updates.model_dump(exclude_none=True)

        doc.render(context)
        doc.save(output_filename)

        print(f"\n✅ Newsletter successfully generated: {output_filename}")

    except Exception as e:
        print(f"\n❌ Error generating Word document: {e}")


# -----------------------------
# SYSTEM PROMPT
# -----------------------------

SYSTEM_PROMPT = """
You are a senior investment banking analyst preparing a weekly institutional newsletter.

Your task is to extract and summarize ONLY news related to:
1. Mergers & Acquisitions (M&A)
2. Equity Capital Markets (ECM)

Output must be structured for a professional financial newsletter.

CONTENT RULES:

Executive Overview:
• Provide 3-4 concise bullet points
• Each bullet must be 10-20 words
• Focus on trends, deal activity, regulatory shifts, capital flows, or market sentiment
• Do NOT repeat specific deal headlines here

Details Section:
• Extract 5-7 relevant news items
• Each item must contain:
  - A clear headline (max 12 words)
  - A short summary (2 sentences maximum, 20-40 words total)

Focus on:
• acquisitions
• IPOs
• fundraises
• strategic divestitures
• regulatory changes affecting capital markets

WRITING STYLE:

• Professional institutional tone
• Concise and factual
• Similar to Goldman Sachs or McKinsey newsletters
• No speculation
• Avoid generic commentary

FORMATTING REQUIREMENTS:

• Headlines must be short
• Do NOT include bullet symbols
• Do NOT include numbering
• Return JSON matching the schema only
"""


# -----------------------------
# GEMINI ANALYSIS
# -----------------------------

def extract_financial_updates(pdf_path: str):
    """
    Analyze a newspaper PDF and extract M&A / ECM news.
    """

    api_key = os.getenv("GEMINI_API_KEY")

    if not api_key:
        print("❌ GEMINI_API_KEY not found in .env file")
        return

    try:
        print("\nUploading PDF to Gemini...")

        client = genai.Client(api_key=api_key)

        uploaded_file = client.files.upload(file=pdf_path)

    except Exception as e:
        print(f"\n❌ File upload failed: {e}")
        return

    try:

        print("Analyzing document with Gemini 2.5 Flash...\n")

        response = client.models.generate_content(
            model="gemini-2.5-flash",
            contents=[
                uploaded_file,
                SYSTEM_PROMPT
            ],
            config=types.GenerateContentConfig(
                temperature=0.2,
                response_mime_type="application/json",
                response_schema=FinancialUpdates,
            ),
        )

        structured_data = FinancialUpdates.model_validate_json(response.text)

        print("✅ Extraction completed")

        create_word_document(structured_data)

    except Exception as e:
        print(f"\n❌ Gemini processing error: {e}")


# -----------------------------
# MAIN
# -----------------------------

if __name__ == "__main__":

    pdf_file = "newspaper.pdf"

    if not os.path.exists(pdf_file):

        print("\n⚠️ Please place your PDF in the same folder and name it:")
        print("newspaper.pdf\n")

    else:

        extract_financial_updates(pdf_file)