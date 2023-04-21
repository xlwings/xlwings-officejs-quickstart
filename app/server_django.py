# Based on https://arunrocks.com/django-application-in-one-file/
# Run with: python app/server_django.py runserver
# pip install django-sslserver
# python app/server_django.py runsslserver --certificate certs/localhost+2.pem --key certs/localhost+2-key.pem
from pathlib import Path
import sys
import json

from django.conf import settings
from django.urls import path
from django.http import JsonResponse
from django.conf.urls.static import static

import xlwings as xw

BASE_DIR = Path(__file__).resolve().parent

settings.configure(
    DEBUG=True,
    SECRET_KEY="changeme",
    ROOT_URLCONF=__name__,
    INSTALLED_APPS=["sslserver"],
    STATICFILES_DIRS=[BASE_DIR],
)


def hello(request):
    # Instantiate a book object with the parsed request body
    data = json.loads(request.body.decode("utf-8"))
    book = xw.Book(json=data)

    # Use xlwings as usual
    sheet = book.sheets[0]
    cell = sheet["A1"]
    if cell.value == "Hello xlwings!":
        cell.value = "Bye xlwings!"
    else:
        cell.value = "Hello xlwings!"

    # Return a JSON response
    return JsonResponse(book.json())


urlpatterns = [
    path("", lambda request: JsonResponse({"status": "ok"})),
    path("hello", hello),
]  + static('/', document_root=BASE_DIR)

if __name__ == "__main__":
    from django.core.management import execute_from_command_line

    execute_from_command_line(sys.argv)
