FROM node:24-bookworm-slim

WORKDIR /app

RUN apt-get update \
  && apt-get install -y --no-install-recommends ca-certificates python3 python3-venv \
  && rm -rf /var/lib/apt/lists/*

COPY package.json ./
RUN npm install --omit=dev --package-lock=false

COPY requirements.txt ./
RUN python3 -m venv /venv
ENV PATH="/venv/bin:${PATH}"
RUN pip install --no-cache-dir -r requirements.txt

COPY app.py workpaper_proof.mjs README.md ./

ENV GRADIO_SERVER_NAME=0.0.0.0
ENV GRADIO_SERVER_PORT=7860

EXPOSE 7860

CMD ["python", "app.py"]
