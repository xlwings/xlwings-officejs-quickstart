from pathlib import Path

import custom_functions
import jinja2
import markupsafe
import xlwings as xw
from flask import Flask, Response, request, send_from_directory
from flask.templating import render_template
from flask_cors import CORS

app = Flask(__name__)
CORS(app)

this_dir = Path(__file__).resolve().parent


@app.route("/")
def root():
    return {"status": "ok"}


@app.route("/hello", methods=["POST"])
def hello():
    # Instantiate a Book object with the deserialized request body
    book = xw.Book(json=request.json)

    # Use xlwings as usual
    sheet = book.sheets[0]
    cell = sheet["A1"]
    if cell.value == "Hello xlwings!":
        cell.value = "Bye xlwings!"
    else:
        cell.value = "Hello xlwings!"

    # Pass the following back as the response
    return book.json()


@app.route("/capitalize-sheet-names-prompt", methods=["POST"])
def capitalize_sheet_names_prompt():
    book = xw.Book(json=request.json)
    book.app.alert(
        prompt="This will capitalize all sheet names!",
        title="Are you sure?",
        buttons="ok_cancel",
        # this is the JS function name that gets called when the user clicks a button
        callback="capitalizeSheetNames",
    )
    return book.json()


@app.route("/capitalize-sheet-names", methods=["POST"])
def capitalize_sheet_names():
    book = xw.Book(json=request.json)
    for sheet in book.sheets:
        sheet.name = sheet.name.upper()
    return book.json()


@app.route("/xlwings/alert")
def alert():
    prompt = request.args.get("prompt")
    title = request.args.get("title")
    buttons = request.args.get("buttons")
    mode = request.args.get("mode")
    callback = request.args.get("callback")
    """Boilerplate required by book.app.alert() and to show unhandled exceptions"""
    return render_template(
        "xlwings-alert.html",
        prompt=markupsafe.Markup(prompt.replace("\n", "<br>")),
        title=title,
        buttons=buttons,
        mode=mode,
        callback=callback,
    )


@app.route("/xlwings/custom-functions-meta")
def custom_functions_meta():
    return xw.pro.custom_functions_meta(custom_functions)


@app.route("/xlwings/custom-functions-code")
def custom_functions_code():
    return xw.pro.custom_functions_code(custom_functions)


@app.route("/xlwings/custom-functions-call", methods=["POST"])
async def custom_functions_call():
    rv = await xw.pro.custom_functions_call(request.json, custom_functions)
    return {"result": rv}


# Add xlwings.html as additional source for templates so the /xlwings/alert endpoint
# will find xlwings-alert.html. "mytemplates" can be a dummy if the app doesn't use
# own templates
loader = jinja2.ChoiceLoader(
    [
        jinja2.FileSystemLoader(str(this_dir / "mytemplates")),
        jinja2.PackageLoader("xlwings", "html"),
    ]
)
app.jinja_loader = loader


# Serve static files (HTML and icons)
# This could also be handled by an external web server such as nginx, etc.
@app.route("/<path:path>")
def static_proxy(path):
    return send_from_directory(this_dir, path)


@app.errorhandler(Exception)
def xlwings_exception_handler(error):
    # This handles all exceptions, so you may want to make this more restrictive
    return Response(str(error), status=500)


if __name__ == "__main__":
    app.run(
        port=8000,
        debug=True,
        ssl_context=(
            this_dir.parent / "certs" / "localhost+2.pem",
            this_dir.parent / "certs" / "localhost+2-key.pem",
        ),
    )
