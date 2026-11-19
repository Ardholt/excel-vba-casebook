# syntax=docker/dockerfile:1
FROM ubuntu:24.04
LABEL org.opencontainers.image.source="https://github.com/Ardholt/excel-vba-casebook"
ARG DEBIAN_FRONTEND=noninteractive
RUN apt-get update && apt-get install -y --no-install-recommends \
    libreoffice-calc python3 python3-uno python3-venv python3-pip \
    python3-openpyxl python3-xlsxwriter python3-pandas python3-pytest python3-pil \
    imagemagick ffmpeg poppler-utils git rsync fonts-dejavu fonts-liberation \
    tesseract-ocr tesseract-ocr-eng curl \
    fonts-crosextra-carlito fonts-crosextra-caladea ca-certificates \
    && rm -rf /var/lib/apt/lists/*
RUN python3 -m venv --system-site-packages /opt/vba \
    && /opt/vba/bin/pip install --no-cache-dir --no-deps \
       'https://files.pythonhosted.org/packages/1c/c3/21de4f0e1013f25a0bd8ef2fb6408c6a3d23cfd7ee4a5c627aeb34d498c1/pyopenvba-6.6.0-py3-none-any.whl#sha256=3c1f370d21f0c57d1f11a69d78896209a14b32da389e7c12b88016ff809b6a47'
ENV SAL_USE_VCLPLUGIN=svp PYTHONDONTWRITEBYTECODE=1 OPENBLAS_NUM_THREADS=1 OMP_NUM_THREADS=2
ENV PLAYWRIGHT_BROWSERS_PATH=/opt/playwright
RUN python3 -m venv /opt/page \
    && /opt/page/bin/pip install --no-cache-dir playwright==1.56.0 markdown-it-py==4.0.0 mdit-py-plugins==0.5.0 linkify-it-py==2.0.3 \
    && /opt/page/bin/playwright install --with-deps chromium \
    && curl -fsSL https://raw.githubusercontent.com/sindresorhus/github-markdown-css/v5.8.1/github-markdown-light.css -o /opt/page/github-markdown.css \
    && curl -fsSL https://raw.githubusercontent.com/sindresorhus/github-markdown-css/v5.8.1/license -o /opt/page/github-markdown-css-LICENSE \
    && rm -rf /var/lib/apt/lists/*
COPY lo_session.py xltool.py pagerender xlentry /usr/local/bin/
RUN chmod 755 /usr/local/bin/xltool.py /usr/local/bin/pagerender /usr/local/bin/xlentry \
    && ln -s xltool.py /usr/local/bin/xlrecalc \
    && ln -s xltool.py /usr/local/bin/xlrender
ENV HOME=/tmp PYTHONUNBUFFERED=1
WORKDIR /work
USER 65534:65534
ENTRYPOINT ["/usr/local/bin/xlentry"]
CMD ["python3", "--version"]
