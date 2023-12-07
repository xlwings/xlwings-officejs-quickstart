import json
import sys
from pathlib import Path

import custom_functions
import markupsafe
import xlwings as xw
from django.conf import settings
from django.conf.urls.static import static
from django.http import HttpResponse, JsonResponse
from django.template import loader
from django.urls import path

BASE_DIR = Path(__file__).resolve().parent

settings.configure(
    DEBUG=True,
    SECRET_KEY="change_me",
    ROOT_URLCONF=__name__,
    TEMPLATES=[
        {
            "BACKEND": "django.template.backends.jinja2.Jinja2",
            "DIRS": [Path(xw.__file__).parent / "html"],
        },
    ],
    # Office Scripts and custom functions in Excel on the web require CORS
    INSTALLED_APPS=["sslserver", "corsheaders"],
    CORS_ALLOW_ALL_ORIGINS=True,
    MIDDLEWARE=[
        "corsheaders.middleware.CorsMiddleware",
        __name__ + ".ErrorHandlerMiddleware",
    ],
)


def root(request):
    return JsonResponse({"status": "ok"})


def hello(request):
    # Instantiate a book object with the parsed request body
    data = json.loads(request.body.decode("utf-8"))
    with xw.Book(json=data) as book:
        # Use xlwings as usual
        sheet = book.sheets[0]
        cell = sheet["A1"]
        if cell.value == "Hello xlwings!":
            cell.value = "Bye xlwings!"
        else:
            cell.value = "Hello xlwings!"

        # Return a JSON response
        return JsonResponse(book.json())


def capitalize_sheet_names_prompt(request):
    data = json.loads(request.body.decode("utf-8"))
    with xw.Book(json=data) as book:
        book.app.alert(
            prompt="This will capitalize all sheet names!",
            title="Are you sure?",
            buttons="ok_cancel",
            # this is the JS function name that gets called when the user clicks a button
            callback="capitalizeSheetNames",
        )

        return JsonResponse(book.json())


def capitalize_sheet_names(request):
    data = json.loads(request.body.decode("utf-8"))
    with xw.Book(json=data) as book:
        for sheet in book.sheets:
            sheet.name = sheet.name.upper()

        return JsonResponse(book.json())


def alert(request):
    """Boilerplate required by book.app.alert() and to show unhandled exceptions"""
    prompt = request.GET.get("prompt")
    title = request.GET.get("title")
    buttons = request.GET.get("buttons")
    mode = request.GET.get("mode")
    callback = request.GET.get("callback")
    template = loader.get_template("xlwings-alert.html")
    context = {
        "prompt": markupsafe.Markup(prompt.replace("\n", "<br>")),
        "title": title,
        "buttons": buttons,
        "mode": mode,
        "callback": callback,
    }
    return HttpResponse(template.render(context, request))


def custom_functions_meta(request):
    """Boilerplate required by custom functions"""
    return JsonResponse(xw.server.custom_functions_meta(custom_functions))


def custom_functions_code(request):
    """Boilerplate required by custom functions"""
    return HttpResponse(
        xw.server.custom_functions_code(custom_functions),
        content_type="application/javascript",
    )


async def custom_functions_call(request):
    """Boilerplate required by custom functions"""
    data = json.loads(request.body.decode("utf-8"))
    rv = await xw.server.custom_functions_call(data, custom_functions)
    return JsonResponse({"result": rv})


class ErrorHandlerMiddleware:
    """Show a short error message in an alert"""

    def __init__(self, get_response):
        self.get_response = get_response

    def __call__(self, request):
        response = self.get_response(request)
        return response

    def process_exception(self, request, exception):
        # This handles all exceptions, so you may want to make this more restrictive
        return HttpResponse(repr(exception), status=500, content_type="text/plain")


urlpatterns = [
    path("", root),
    path("hello", hello),
    path("capitalize-sheet-names-prompt", capitalize_sheet_names_prompt),
    path("capitalize-sheet-names", capitalize_sheet_names),
    path("xlwings/alert", alert),
    path("xlwings/custom-functions-meta", custom_functions_meta),
    path("xlwings/custom-functions-code", custom_functions_code),
    path("xlwings/custom-functions-call", custom_functions_call),
] + static("/", document_root=BASE_DIR)

if __name__ == "__main__":
    from django.core.management import execute_from_command_line

    execute_from_command_line(sys.argv)
