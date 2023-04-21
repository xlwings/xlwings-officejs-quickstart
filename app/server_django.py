# Based on https://arunrocks.com/django-application-in-one-file/
# Run with: python app/server_django.py runserver
import sys
import json

from django.conf import settings
from django.urls import path
from django.http import JsonResponse

import xlwings as xw

settings.configure(
	DEBUG=True,
	SECRET_KEY="changeme",
	ROOT_URLCONF=__name__,
)


def hello(request):
    # Instantiate a book object with the parsed request body
    data = json.loads(request.body.decode('utf-8'))
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
	path("hello", hello),
]

if __name__ == "__main__":
	from django.core.management import execute_from_command_line

	execute_from_command_line(sys.argv)
