FROM python:3.13-slim AS builder

COPY --from=ghcr.io/astral-sh/uv:latest /uv /usr/local/bin/uv

WORKDIR /app
COPY pyproject.toml uv.lock ./
RUN uv sync --frozen --no-dev

FROM python:3.13-slim

WORKDIR /app
COPY --from=builder /app/.venv /app/.venv

COPY app.py config.yaml ./
COPY src/ src/

ENV PATH="/app/.venv/bin:$PATH"

EXPOSE 7860

CMD ["python", "app.py"]
