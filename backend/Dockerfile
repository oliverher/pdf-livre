FROM python:3.12-slim

ENV PYTHONDONTWRITEBYTECODE=1 PYTHONUNBUFFERED=1
WORKDIR /srv

# LibreOffice (converte PPTX->PDF), Tesseract + portugues, Poppler e libs do OpenCV
RUN apt-get update && apt-get install -y --no-install-recommends \
      libreoffice-impress \
      tesseract-ocr tesseract-ocr-por \
      poppler-utils \
      libglib2.0-0 libsm6 libxext6 libxrender1 libgl1 \
    && rm -rf /var/lib/apt/lists/*

COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

COPY app ./app

# usuario sem privilegio: o servico manipula arquivo de terceiro
RUN useradd -m -u 10001 servico && chown -R servico:servico /srv
USER servico

# limite de tamanho ajustavel; chave de visao fica vazia por padrao (modo convencional)
ENV PPTLIVRE_MAX_MB=60
EXPOSE 8000
# ${PORT} e injetada por servicos de hospedagem (Render, Railway, Fly.io...);
# 8000 e o padrao para rodar local/Docker Compose sem essa variavel.
CMD ["sh", "-c", "uvicorn app.main:app --host 0.0.0.0 --port ${PORT:-8000}"]
