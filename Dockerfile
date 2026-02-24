FROM python:3.11-slim

ENV PYTHONDONTWRITEBYTECODE=1
ENV PYTHONUNBUFFERED=1

WORKDIR /app

# -------------------------------------------------
# System deps + font support
# -------------------------------------------------
RUN apt-get update && apt-get install -y \
    build-essential \
    fontconfig \
    && rm -rf /var/lib/apt/lists/*

# -------------------------------------------------
# Install Rupee Foradian font
# -------------------------------------------------
COPY fonts/Rupee_Foradian.ttf /usr/share/fonts/truetype/

RUN fc-cache -f -v

# -------------------------------------------------
# Python deps
# -------------------------------------------------
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# -------------------------------------------------
# App code
# -------------------------------------------------
COPY . .

CMD ["bash"]
