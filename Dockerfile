FROM python:3.11-slim

# Makes sure that logs are shown immediately
ENV PYTHONUNBUFFERED=1

COPY ./requirements-fastapi.txt .
RUN pip install --no-cache-dir -r requirements-fastapi.txt
# Watchfiles from uvicorn[standard] breaks reload inside Docker
RUN pip uninstall watchfiles -y

COPY ./app /app

EXPOSE 8000

WORKDIR /app

# This is for single-container deployments (multiple-workers)
CMD ["gunicorn", "server_fastapi:cors_app", \
     "--bind", "0.0.0.0:8000", \
     "--access-logfile", "-", \
     "--workers", "2", \
     "--worker-class", "uvicorn.workers.UvicornWorker"]
