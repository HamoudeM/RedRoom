FROM python:3.13-slim
WORKDIR /app
COPY . .
RUN pip install --no-cache-dir python-telegram-bot aiohttp reportlab python-docx openpyxl
CMD ["python", "bot.py"]
