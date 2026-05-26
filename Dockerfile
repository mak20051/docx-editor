FROM python:3.11-slim
# antiword: lightweight .doc reader (replaces LibreOffice, ~200KB)
RUN apt-get update && apt-get install -y --no-install-recommends antiword \
    && rm -rf /var/lib/apt/lists/* \
    && mkdir -p /root/.antiword \
    && cp /usr/share/antiword/* /root/.antiword/
WORKDIR /app
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt
COPY . .
ENV PORT=8000
ENV DOCKER=1
EXPOSE 8000
CMD ["python", "main.py"]
