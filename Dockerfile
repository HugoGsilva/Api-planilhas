FROM python:3.13-slim

ENV PYTHONDONTWRITEBYTECODE=1
ENV PYTHONUNBUFFERED=1

WORKDIR /app

COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

COPY src ./src
COPY scripts ./scripts

RUN mkdir -p storage/jobs

EXPOSE 8000

CMD ["python", "-m", "uvicorn", "api_planilhas.web:app", "--app-dir", "src", "--host", "0.0.0.0", "--port", "8000"]
