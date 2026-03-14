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
        description="Short professional headline, maximum 12 words"
    )

    content: str = Field(
        description="Two sentence summary explaining the news event. Maximum 40 words."
    )


class FinancialUpdates(BaseModel):

    date: str = Field(
        description="Newsletter date"
    )

    theme: str = Field(
        description="Short report theme summarizing the entire report. Maximum 6 words."
    )

    executive_overview: List[str] = Field(
        description="3–4 concise bullet points summarizing trends in M&A and Equity Capital Markets"
    )

    details: List[NewsDetail] = Field(
        description="5–7 relevant news items related to M&A, IPOs, fundraises, acquisitions, or capital market regulations"
    )


# -----------------------------
# WORD DOCUMENT GENERATOR
# -----------------------------

def create_word_document(
    updates: FinancialUpdates,
    template_filename="template.docx",
    output_filename="newsletter_output.docx"
):

    try:
        doc = DocxTemplate(template_filename)

        context = updates.model_dump(exclude_none=True)

        doc.render(context)
        doc.save(output_filename)

        print(f"\n✅ Newsletter generated successfully: {output_filename}")

    except Exception as e:
        print(f"\n❌ Word document generation error: {e}")


# -----------------------------
# SYSTEM PROMPT
# -----------------------------

SYSTEM_PROMPT = """
You are a senior investment banking analyst preparing a weekly capital markets newsletter.

Carefully read the provided newspaper and extract ONLY relevant information about:

1. Mergers & Acquisitions (M&A)
2. Equity Capital Markets (ECM)
3. IPOs
4. Fundraises
5. Strategic investments
6. Capital market regulations affecting financing

Ignore unrelated macro news, politics, or general business news.

OUTPUT STRUCTURE:

Theme:
• Generate a short theme summarizing the overall report
• Maximum 6 words
• Example:
  - REIT Growth & Market Volatility
  - IPO Momentum & Regulatory Shifts
  - Capital Markets & Deal Activity

Executive Overview:
• Provide 3–4 bullet points
• Each bullet must be 10–20 words
• Summarize overall trends across the extracted news

Details Section:
• Extract 5–7 important news items
• Each item must contain:
  - headline (max 12 words)
  - summary (max 40 words)

WRITING STYLE:

• Institutional tone
• Professional financial language
• Similar to Goldman Sachs or McKinsey newsletters
• Concise and factual
• No speculation

FORMATTING RULES:

• Do NOT include bullet symbols
• Do NOT include numbering
• Return structured JSON matching the schema exactly
"""


# -----------------------------
# GEMINI ANALYSIS
# -----------------------------

def extract_financial_updates(pdf_path: str):

    api_key = os.getenv("GEMINI_API_KEY")

    if not api_key:
        print("❌ GEMINI_API_KEY not found in .env file")
        return

    try:

        print("\n📄 Uploading PDF to Gemini...")

        client = genai.Client(api_key=api_key)

        uploaded_file = client.files.upload(file=pdf_path)

    except Exception as e:

        print(f"\n❌ File upload failed: {e}")
        return

    try:

        print("🔎 Analyzing newspaper with Gemini 2.5 Flash...\n")

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

        print("✅ News extraction completed")

        create_word_document(structured_data)

    except Exception as e:

        print(f"\n❌ Gemini processing error: {e}")


# -----------------------------
# MAIN
# -----------------------------

if __name__ == "__main__":

    pdf_file = "newspaper.pdf"

    if not os.path.exists(pdf_file):

        print("\n⚠️ Please place your newspaper PDF in this folder and name it:")
        print("newspaper.pdf\n")

    else:

        extract_financial_updates(pdf_file)