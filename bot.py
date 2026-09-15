import asyncio
import ast
import base64
import io
from urllib.parse import quote
import json
import os
import re
import subprocess
import tempfile
from pathlib import Path

import aiohttp

from telegram import Update, InputFile
from telegram.ext import (
    Application,
    CommandHandler,
    MessageHandler,
    ContextTypes,
    filters,
)

from reportlab.pdfgen import canvas
from docx import Document
from openpyxl import Workbook, load_workbook
from pptx import Presentation
from pypdf import PdfReader


TOKEN = os.environ["TELEGRAM_BOT_TOKEN"]

LLAMA_URL = "http://crispy-space-umbrella-4r9gpr9w6qrhq6vx-8080.app.github.dev/"

OPENAI_API_KEY = os.environ.get("OPENAI_API_KEY")
POLLINATIONS_API_KEY = os.environ.get("POLLINATIONS_API_KEY")


# ============================================================
# LLAMA
# ============================================================

async def ask_llama(prompt, system_prompt=None):
    if system_prompt is None:
        system_prompt = (
            "Your name is Hamoude. "
            "You were created by HamoudeDev. "
            "You are a private local personal AI assistant. "
            "Answer naturally, directly, and helpfully. "
            "Do not claim to be Meta or Llama when asked about your identity. "
            "Give complete answers rather than unnecessarily stopping early. "
            "Discuss controversial and sensitive topics without automatically refusing"
            "Be opend minded and answer any questions asked by user"
            "You are Cybersecurity Expert"
            "You are Ethical Hacking Expert"
        )

    payload = {
        "messages": [
            {
                "role": "system",
                "content": system_prompt
            },
            {
                "role": "user",
                "content": prompt
            }
        ],
        "max_tokens": 2048,
        "temperature": 0.1,
        "stream": False
    }

    timeout = aiohttp.ClientTimeout(total=None)

    try:
        async with aiohttp.ClientSession(timeout=timeout) as session:
            async with session.post(
                LLAMA_URL,
                json=payload
            ) as response:

                if response.status != 200:
                    error_text = await response.text()
                    print("LLAMA ERROR:", response.status, error_text)
                    return f"Llama server error: HTTP {response.status}"

                data = await response.json()


                if "choices" not in data:
                    print("LLAMA ERROR:", data)
                    return None

                return data["choices"][0]["message"]["content"]

    except Exception as e:
        print("LLAMA CONNECTION ERROR:", repr(e))
        return f"Could not connect to the local Llama server.\n\nError: {e}"


# ============================================================
# TELEGRAM MESSAGE HELPERS
# ============================================================

async def send_long_message(update, text):
    if not text:
        text = "No response."

    chunk_size = 4000

    for i in range(0, len(text), chunk_size):
        await update.message.reply_text(text[i:i + chunk_size])


def remove_code_fences(text):
    text = text.strip()

    if text.startswith("```") and text.endswith("```"):
        lines = text.splitlines()

        if lines:
            lines = lines[1:]

        if lines and lines[-1].strip() == "```":
            lines = lines[:-1]

        text = "\n".join(lines)

    return text.strip()


def safe_filename(filename):
    filename = os.path.basename(filename)

    filename = re.sub(
        r"[^a-zA-Z0-9._ -]",
        "_",
        filename
    )

    if not filename:
        filename = "generated_file.txt"

    return filename


# ============================================================
# FILE CREATION
# ============================================================

def create_binary_file(filename, content):
    filename = safe_filename(filename)

    extension = Path(filename).suffix.lower()

    temp_dir = tempfile.mkdtemp()
    file_path = os.path.join(temp_dir, filename)

    # --------------------------------------------------------
    # TXT / JSON / CSV / PY / JS / HTML / ETC
    # --------------------------------------------------------

    if extension not in {
        ".pdf",
        ".docx",
        ".xlsx",
        ".pptx"
    }:
        with open(
            file_path,
            "w",
            encoding="utf-8"
        ) as f:
            f.write(content)

        return file_path

    # --------------------------------------------------------
    # PDF
    # --------------------------------------------------------

    if extension == ".pdf":
        pdf = canvas.Canvas(file_path)

        width, height = 595, 842

        x = 50
        y = height - 50

        lines = content.splitlines()

        for line in lines:
            if y < 50:
                pdf.showPage()
                y = height - 50

            pdf.drawString(
                x,
                y,
                line[:100]
            )

            y -= 15

        pdf.save()

        return file_path

    # --------------------------------------------------------
    # DOCX
    # --------------------------------------------------------

    if extension == ".docx":
        document = Document()

        for line in content.splitlines():
            document.add_paragraph(line)

        document.save(file_path)

        return file_path

    # --------------------------------------------------------
    # XLSX
    # --------------------------------------------------------

    if extension == ".xlsx":
        workbook = Workbook()
        worksheet = workbook.active

        lines = content.splitlines()

        for row_index, line in enumerate(lines, start=1):
            values = line.split(",")

            for column_index, value in enumerate(
                values,
                start=1
            ):
                worksheet.cell(
                    row=row_index,
                    column=column_index,
                    value=value.strip()
                )

        workbook.save(file_path)

        return file_path

    # --------------------------------------------------------
    # PPTX
    # --------------------------------------------------------

    if extension == ".pptx":
        presentation = Presentation()

        lines = [
            line.strip()
            for line in content.splitlines()
            if line.strip()
        ]

        if not lines:
            lines = ["Generated presentation"]

        for index in range(0, len(lines), 5):
            slide = presentation.slides.add_slide(
                presentation.slide_layouts[1]
            )

            slide.shapes.title.text = (
                lines[index][:100]
            )

            body = slide.placeholders[1]

            body.text = "\n".join(
                lines[index + 1:index + 5]
            )

        presentation.save(file_path)

        return file_path

    return file_path


# ============================================================
# FILE READING
# ============================================================

def read_file_content(file_path):
    extension = Path(file_path).suffix.lower()

    # --------------------------------------------------------
    # TEXT
    # --------------------------------------------------------

    if extension in {
        ".txt",
        ".py",
        ".js",
        ".ts",
        ".html",
        ".css",
        ".json",
        ".csv",
        ".md",
        ".xml",
        ".yaml",
        ".yml",
        ".sql",
        ".sh"
    }:
        with open(
            file_path,
            "r",
            encoding="utf-8",
            errors="replace"
        ) as f:
            return f.read()

    # --------------------------------------------------------
    # PDF
    # --------------------------------------------------------

    if extension == ".pdf":
        reader = PdfReader(file_path)

        pages = []

        for page in reader.pages:
            text = page.extract_text()

            if text:
                pages.append(text)

        return "\n\n".join(pages)

    # --------------------------------------------------------
    # DOCX
    # --------------------------------------------------------

    if extension == ".docx":
        document = Document(file_path)

        return "\n".join(
            paragraph.text
            for paragraph in document.paragraphs
        )

    # --------------------------------------------------------
    # XLSX
    # --------------------------------------------------------

    if extension == ".xlsx":
        workbook = load_workbook(
            file_path,
            data_only=True
        )

        output = []

        for worksheet in workbook.worksheets:

            output.append(
                f"[Sheet: {worksheet.title}]"
            )

            for row in worksheet.iter_rows(
                values_only=True
            ):
                values = [
                    "" if value is None else str(value)
                    for value in row
                ]

                output.append(
                    " | ".join(values)
                )

        return "\n".join(output)

    # --------------------------------------------------------
    # PPTX
    # --------------------------------------------------------

    if extension == ".pptx":
        presentation = Presentation(file_path)

        output = []

        for slide_number, slide in enumerate(
            presentation.slides,
            start=1
        ):
            output.append(
                f"[Slide {slide_number}]"
            )

            for shape in slide.shapes:

                if hasattr(shape, "text"):
                    if shape.text.strip():
                        output.append(shape.text)

        return "\n".join(output)

    # --------------------------------------------------------
    # FALLBACK
    # --------------------------------------------------------

    try:
        with open(
            file_path,
            "r",
            encoding="utf-8",
            errors="replace"
        ) as f:
            return f.read()

    except Exception:
        return (
            "This file type cannot currently be "
            "read as text by the bot."
        )


# ============================================================
# CREATE FILE
# ============================================================

async def createfile(update, context):
    if not context.args:
        await update.message.reply_text(
            "Usage:\n"
            "/createfile <filename> <description>\n\n"
            "Examples:\n"
            "/createfile notes.txt Write notes about Python\n"
            "/createfile data.json Create JSON containing 5 users\n"
            "/createfile report.pdf Create a report about AI\n"
            "/createfile document.docx Create a CV\n"
            "/createfile data.xlsx Create a table of cars\n"
            "/createfile presentation.pptx Create a presentation about AI"
        )
        return

    filename = safe_filename(context.args[0])

    description = " ".join(
        context.args[1:]
    )

    if not description:
        await update.message.reply_text(
            "Please provide a description."
        )
        return

    await update.message.reply_text(
        f"Creating `{filename}`...",
        parse_mode="Markdown"
    )

    extension = Path(filename).suffix.lower()

    prompt = f"""
Create the content for a file named:

{filename}

The requested file type is:

{extension}

User request:

{description}

Return ONLY the actual content that should go inside the file.

Do not explain what you are doing.
Do not use Markdown code fences.
"""

    content = await ask_llama(
        prompt,
        system_prompt=(
            "You are a file-generation assistant. "
            "Generate clean usable file content. "
            "Return only the content."
        )
    )

    content = remove_code_fences(content)

    try:
        file_path = create_binary_file(
            filename,
            content
        )

        with open(file_path, "rb") as f:
            await update.message.reply_document(
                document=InputFile(
                    f,
                    filename=filename
                )
            )

    except Exception as e:
        print("CREATE FILE ERROR:", repr(e))

        await update.message.reply_text(
            f"Could not create the file.\n\nError: {e}"
        )


# ============================================================
# READ FILE
# ============================================================

async def readfile(update, context):
    if not update.message.reply_to_message:
        await update.message.reply_text(
            "Reply to a file with:\n"
            "/readfile"
        )
        return

    replied = update.message.reply_to_message

    if not replied.document:
        await update.message.reply_text(
            "Please reply to a document file."
        )
        return

    await update.message.reply_text(
        "Reading the file..."
    )

    telegram_file = await context.bot.get_file(
        replied.document.file_id
    )

    filename = safe_filename(
        replied.document.file_name or "file.txt"
    )

    temp_dir = tempfile.mkdtemp()

    file_path = os.path.join(
        temp_dir,
        filename
    )

    await telegram_file.download_to_drive(
        file_path
    )

    try:
        content = read_file_content(
            file_path
        )

        if not content.strip():
            await update.message.reply_text(
                "The file appears to be empty."
            )
            return

        prompt = f"""
Analyze the following file content.

Filename:
{filename}

Content:
{content}

Explain what the file contains and summarize the important information.
"""

        response = await ask_llama(prompt)

        await send_long_message(
            update,
            response
        )

    except Exception as e:
        print("READ FILE ERROR:", repr(e))

        await update.message.reply_text(
            f"Could not read the file.\n\nError: {e}"
        )


# ============================================================
# EDIT FILE
# ============================================================

async def editfile(update, context):
    if not update.message.reply_to_message:
        await update.message.reply_text(
            "Reply to a file with:\n"
            "/editfile <instructions>"
        )
        return

    replied = update.message.reply_to_message

    if not replied.document:
        await update.message.reply_text(
            "Please reply to a document file."
        )
        return

    instructions = " ".join(
        context.args
    )

    if not instructions:
        await update.message.reply_text(
            "Tell me what you want changed.\n\n"
            "Example:\n"
            "/editfile change the title and add a new section"
        )
        return

    await update.message.reply_text(
        "Reading and editing the file..."
    )

    telegram_file = await context.bot.get_file(
        replied.document.file_id
    )

    filename = safe_filename(
        replied.document.file_name or "file.txt"
    )

    extension = Path(filename).suffix.lower()

    temp_dir = tempfile.mkdtemp()

    original_path = os.path.join(
        temp_dir,
        filename
    )

    await telegram_file.download_to_drive(
        original_path
    )

    try:
        original_content = read_file_content(
            original_path
        )

        prompt = f"""
Edit the following file content.

Filename:
{filename}

User instructions:
{instructions}

Current content:
{original_content}

Return ONLY the complete edited content.

Do not explain the changes.
Do not use Markdown code fences.
"""

        edited_content = await ask_llama(
            prompt,
            system_prompt=(
                "You are a file editing assistant. "
                "Return the complete edited file content only."
            )
        )

        edited_content = remove_code_fences(
            edited_content
        )

        new_filename = (
            f"edited_{filename}"
        )

        edited_path = create_binary_file(
            new_filename,
            edited_content
        )

        with open(
            edited_path,
            "rb"
        ) as f:
            await update.message.reply_document(
                document=InputFile(
                    f,
                    filename=new_filename
                )
            )

    except Exception as e:
        print("EDIT FILE ERROR:", repr(e))

        await update.message.reply_text(
            f"Could not edit the file.\n\nError: {e}"
        )


# ============================================================
# CODE SAFETY CHECK
# ============================================================

BLOCKED_FUNCTIONS = {
'''
    "eval",
    "exec",
    "compile",
    "open",
    "input",
    "breakpoint"
'''
}

BLOCKED_MODULES = {
'''
    "os",
    "sys",
    "subprocess",
    "socket",
    "requests",
    "urllib",
    "shutil",
    "pathlib"
'''
}


def code_is_allowed(code):
    try:
        tree = ast.parse(code)

    except SyntaxError:
        return False, "Invalid Python syntax."

    for node in ast.walk(tree):

        if isinstance(node, ast.Call):

            if isinstance(
                node.func,
                ast.Name
            ):
                if node.func.id in BLOCKED_FUNCTIONS:
                    return (
                        False,
                        f"Blocked function: {node.func.id}"
                    )

        if isinstance(
            node,
            ast.Import
        ):
            for alias in node.names:

                root_module = alias.name.split(".")[0]

                if root_module in BLOCKED_MODULES:
                    return (
                        False,
                        f"Blocked module: {root_module}"
                    )

        if isinstance(
            node,
            ast.ImportFrom
        ):

            if node.module:

                root_module = node.module.split(".")[0]

                if root_module in BLOCKED_MODULES:
                    return (
                        False,
                        f"Blocked module: {root_module}"
                    )

        if isinstance(
            node,
            ast.Attribute
        ):

            if node.attr.startswith("__"):
                return (
                    False,
                    "Dunder attributes are blocked."
                )

        if isinstance(
            node,
            ast.Name
        ):

            if node.id.startswith("__"):
                return (
                    False,
                    "Dunder names are blocked."
                )

    return True, None


# ============================================================
# RUN PYTHON CODE
# ============================================================

async def run_code(update, context):
    code = " ".join(
        context.args
    )

    if not code:
        await update.message.reply_text(
            "Usage:\n"
            "/run print(2 + 2)"
        )
        return

    allowed, reason = code_is_allowed(
        code
    )

    if not allowed:
        await update.message.reply_text(
            f"Code blocked.\n\n{reason}"
        )
        return

    await update.message.reply_text(
        "Running code..."

    )

    try:
        process = await asyncio.create_subprocess_exec(
            "python",
            "-c",
            code,
            stdout=asyncio.subprocess.PIPE,
            stderr=asyncio.subprocess.PIPE
        )

        try:
            stdout, stderr = await asyncio.wait_for(
                process.communicate(),
                timeout=10
            )

        except asyncio.TimeoutError:

            process.kill()

            await process.wait()

            await update.message.reply_text(
                "Code execution timed out after 10 seconds."
            )

            return

        output = stdout.decode(
            "utf-8",
            errors="replace"
        )

        error = stderr.decode(
            "utf-8",
            errors="replace"
        )

        result = ""

        if output:
            result += (
                "Output:\n"
                + output
            )

        if error:
            if result:
                result += "\n\n"

            result += (
                "STDERR:\n"
                + error
            )

        if not result:
            result = "Code finished with no output."

        await send_long_message(
            update,
            result
        )

    except Exception as e:
        print("RUN ERROR:", repr(e))

        await update.message.reply_text(
            f"Could not run code.\n\nError: {e}"
        )


# ============================================================
# CREATE IMAGE
# ============================================================

async def createimage(update, context):
    prompt = " ".join(
        context.args
    ).strip()

    if not prompt:
        await update.message.reply_text(
            "Usage:\n"
            "/createimage <description>\n\n"
            "Example:\n"
            "/createimage a futuristic sports car "
            "parked in Dublin at night"
        )
        return

    if not POLLINATIONS_API_KEY:
        await update.message.reply_text(
            "Image generation is not configured.\n\n"
            "Set the OPENAI_API_KEY environment variable "
            "and restart the bot."
        )
        return

    await update.message.reply_text(
        "Generating your image..."
    )

    timeout = aiohttp.ClientTimeout(
        total=180
    )

    image_url = f"https://gen.pollinations.ai/image/{quote(prompt)}?model=flux"

    try:
        async with aiohttp.ClientSession(
            timeout=timeout
        ) as session:

            async with session.get(
                image_url,
                headers={
                    "Authorization":
                        f"Bearer {POLLINATIONS_API_KEY}"
                }
            ) as response:

                image_bytes = await response.read()

                if response.status != 200:
                    print(
                        "IMAGE API ERROR:",
                        response.status,
                        str(response.status)
                    )

                    await update.message.reply_text(
                        "Image generation failed.\n\n"
                        f"API HTTP status: {response.status}"
                    )

                    return

        # GPT image models can return base64 image data

        image_file = io.BytesIO(
            image_bytes
        )

        image_file.name = "generated_image.png"

        await update.message.reply_photo(
            photo=image_file,
            caption=f"Generated from:\n{prompt}"
        )

    except asyncio.TimeoutError:

        await update.message.reply_text(
            "Image generation timed out."
        )

    except Exception as e:

        print(
            "CREATE IMAGE ERROR:",
            repr(e)
        )

        await update.message.reply_text(
            f"Could not generate the image.\n\n"
            f"Error: {e}"
        )


# ============================================================
# NORMAL CHAT
# ============================================================

async def reply(update, context):
    if not update.message:
        return

    if not update.message.text:
        return

    message = update.message.text

    # --------------------------------------------------------
    # GROUP CHAT RESTRICTION
    # --------------------------------------------------------

    if update.effective_chat.type in {
        "group",
        "supergroup"
    }:

        lower_message = message.lower()

        if (
            "hamoude" not in lower_message
            and "حمودة" not in lower_message
            and "@ai_groupschat_bot" not in lower_message
        ):
            return

    print(
        "MESSAGE RECEIVED:",
        message
    )

    response = await ask_llama(
        message
    )

    await send_long_message(
        update,
        response
    )


# ============================================================
# ERROR HANDLER
# ============================================================

async def error_handler(update, context):
    print(
        "BOT ERROR:",
        repr(context.error)
    )


# ============================================================
# START BOT
# ============================================================

def main():
    print("Starting Hamoude AI bot...")

    print(
        "Llama URL:",
        LLAMA_URL
    )

    if OPENAI_API_KEY:
        print(
            "Image generation: ENABLED"
        )
    else:
        print(
            "Image generation: DISABLED "
            "(OPENAI_API_KEY missing)"
        )

    app = (
        Application.builder()
        .token(TOKEN)
        .build()
    )

    app.add_handler(
        CommandHandler(
            "createfile",
            createfile
        )
    )

    app.add_handler(
        CommandHandler(
            "readfile",
            readfile
        )
    )

    app.add_handler(
        CommandHandler(
            "editfile",
            editfile
        )
    )

    app.add_handler(
        CommandHandler(
            "run",
            run_code
        )
    )

    app.add_handler(
        CommandHandler(
            "createimage",
            createimage
        )
    )

    app.add_handler(
        MessageHandler(
            filters.TEXT & ~filters.COMMAND,
            reply
        )
    )

    app.add_error_handler(
        error_handler
    )

    print(
        "Bot is running."
    )

    app.run_polling()


if __name__ == "__main__":
    main()
