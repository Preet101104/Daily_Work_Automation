import os
import asyncio
import pythoncom
from typing import List, Dict
from dotenv import load_dotenv
from google import genai
from google.genai import types
from pydantic import BaseModel, Field
from docxtpl import DocxTemplate
from docx2pdf import convert

from telegram import Update
from telegram.ext import Application, CommandHandler, MessageHandler, filters, ContextTypes

# Load environment variables
load_dotenv(dotenv_path=os.path.join(os.getcwd(), ".env"))


# -----------------------------
# DATA SCHEMA
# -----------------------------

class NewsDetail(BaseModel):
    title: str = Field(description="Short professional headline, maximum 12 words")
    content: str = Field(description="Two sentence summary explaining the news event. Maximum 40 words.")

class FinancialUpdates(BaseModel):
    date: str = Field(description="Newsletter date")
    theme: str = Field(description="Short report theme summarizing the entire report. Maximum 6 words.")
    executive_overview: List[str] = Field(description="3-4 concise bullet points summarizing trends in M&A and Equity Capital Markets")
    details: List[NewsDetail] = Field(description="5-7 relevant news items related to M&A, IPOs, fundraises, acquisitions, or capital market regulations")


# -----------------------------
# WORD DOCUMENT GENERATOR
# -----------------------------

def create_word_document(
    updates: FinancialUpdates,
    template_filename="template.docx",
    output_filename="newsletter_output.docx",
    pdf_filename="newsletter_output.pdf"
) -> bool:
    try:
        # Initialize COM for the background thread
        pythoncom.CoInitialize()
        
        doc = DocxTemplate(template_filename)
        context = updates.model_dump(exclude_none=True)
        doc.render(context)
        doc.save(output_filename)
        print(f"\n✅ Newsletter DOCX generated successfully: {output_filename}")
        
        print("\n📄 Converting Word document to PDF...")
        convert(output_filename, pdf_filename)
        print(f"✅ Newsletter PDF generated successfully: {pdf_filename}")
        return True
    except Exception as e:
        print(f"\n❌ Document generation error: {e}")
        return False
    finally:
        pythoncom.CoUninitialize()


# -----------------------------
# SYSTEM PROMPT
# -----------------------------

SYSTEM_PROMPT = """
You are a senior investment banking analyst preparing a weekly capital markets newsletter.

Carefully read ALL the provided newspapers and extract ONLY relevant information about:

1. Mergers & Acquisitions (M&A)
2. Equity Capital Markets (ECM)
3. IPOs
4. Fundraises
5. Strategic investments
6. Capital market regulations affecting financing

Ignore unrelated macro news, politics, or general business news.
CRITICAL: Since there are multiple sources, merge duplicate news stories and AVOID repetitive news items. Only include unique, distinct events.

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

async def extract_financial_updates(pdf_paths: List[str], base_filename: str) -> str:
    api_key = os.getenv("GEMINI_API_KEY")
    if not api_key:
        raise ValueError("GEMINI_API_KEY not found in .env file")

    client = genai.Client(api_key=api_key)
    uploaded_files = []

    try:
        print(f"\n📄 Uploading {len(pdf_paths)} PDFs to Gemini...")
        for pdf_path in pdf_paths:
            print(f"Uploading {pdf_path}")
            uploaded_file = await asyncio.to_thread(client.files.upload, file=pdf_path)
            uploaded_files.append(uploaded_file)
            
        print("🔎 Analyzing newspapers with Gemini 2.5 Flash...\n")
        response = await asyncio.to_thread(
            client.models.generate_content,
            model="gemini-2.5-flash",
            contents=[*uploaded_files, SYSTEM_PROMPT],
            config=types.GenerateContentConfig(
                temperature=0.2,
                response_mime_type="application/json",
                response_schema=FinancialUpdates,
            )
        )

        structured_data = FinancialUpdates.model_validate_json(response.text)
        print("✅ News extraction completed")

        docx_path = f"{base_filename}.docx"
        pdf_path_output = f"{base_filename}.pdf"

        # Export via a thread because COM via docx2pdf needs to avoid blocking bot async loop
        success = await asyncio.to_thread(
            create_word_document,
            structured_data,
            "template.docx",
            docx_path,
            pdf_path_output
        )

        if success:
            return pdf_path_output
        return None

    except Exception as e:
        print(f"\n❌ Gemini processing error: {e}")
        return None


# -----------------------------
# TELEGRAM BOT Logic
# -----------------------------

user_sessions: Dict[int, List[str]] = {}

async def start(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    welcome_text = (
        "👋 Welcome to the Daily Work Automation Bot!\n\n"
        "Send me one or more Newspaper PDFs. Once you're done mapping all your PDFs, send the command /generate and I will create your summarized Financial Newsletter."
    )
    await update.message.reply_text(welcome_text)
    user_sessions[update.message.chat_id] = []

async def receive_document(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    chat_id = update.message.chat_id
    document = update.message.document

    if not document.file_name.lower().endswith('.pdf'):
        await update.message.reply_text("⚠️ Please send only PDF files.")
        return

    if chat_id not in user_sessions:
        user_sessions[chat_id] = []

    file_id = document.file_id
    new_file = await context.bot.get_file(file_id)
    
    os.makedirs("downloads", exist_ok=True)
    file_path = f"downloads/{file_id}_{document.file_name}"
    await new_file.download_to_drive(custom_path=file_path)

    user_sessions[chat_id].append(file_path)

    await update.message.reply_text(
        f"📄 Received `{document.file_name}`.\n\n"
        f"Total PDFs queued: {len(user_sessions[chat_id])}\n"
        "Keep sending PDFs, or reply with /generate to create your newsletter."
    )

async def generate_newsletter(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    chat_id = update.message.chat_id
    pdfs = user_sessions.get(chat_id, [])

    if not pdfs:
        await update.message.reply_text("⚠️ No PDFs received. Please send at least one PDF first.")
        return

    await update.message.reply_text("⏳ Processing your documents with Gemini AI... Please wait, this may take a minute!")

    try:
        base_filename = f"downloads/newsletter_{chat_id}"
        
        output_pdf = await extract_financial_updates(pdfs, base_filename)

        if output_pdf and os.path.exists(output_pdf):
            with open(output_pdf, 'rb') as doc:
                await update.message.reply_document(
                    document=doc, 
                    filename="Financial_Newsletter.pdf", 
                    caption="✅ Here is your summarized newsletter!"
                )
        else:
            await update.message.reply_text("❌ Failed to generate the document. There might be an issue with processing.")
    
    except Exception as e:
        await update.message.reply_text(f"❌ An error occurred: {e}")
        
    finally:
        # Cleanup routine
        for pdf in pdfs:
            if os.path.exists(pdf):
                os.remove(pdf)
        
        user_sessions[chat_id] = []

def main():
    token = os.getenv("TELEGRAM_BOT_TOKEN")
    
    if not token:
        print("❌ TELEGRAM_BOT_TOKEN not found in .env file")
        print(f"📍 Current working directory: {os.getcwd()}")
        print(f"📍 .env path: {os.path.join(os.getcwd(), '.env')}")
        print(f"📍 .env exists: {os.path.exists('.env')}")
        return
        
    print("🤖 Starting Telegram Bot...")
    application = Application.builder().token(token).build()

    application.add_handler(CommandHandler("start", start))
    application.add_handler(CommandHandler("generate", generate_newsletter))
    application.add_handler(MessageHandler(filters.Document.ALL, receive_document))

    print("🚀 Bot is now running. Give it a try on Telegram!")
    application.run_polling()

if __name__ == "__main__":
    main()
