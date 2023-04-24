from pathlib import Path

import custom_functions
import jinja2
import markupsafe
import xlwings as xw
from fastapi import Body, FastAPI, Request, status
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import HTMLResponse, PlainTextResponse
from fastapi.staticfiles import StaticFiles
from fastapi.templating import Jinja2Templates

app = FastAPI()

this_dir = Path(__file__).resolve().parent


@app.get("/")
async def root():
    return {"status": "ok"}


@app.post("/hello")
async def hello(data: dict = Body):
    # Instantiate a Book object with the deserialized request body
    book = xw.Book(json=data)

    # Use xlwings as usual
    sheet = book.sheets[0]
    cell = sheet["A1"]
    if cell.value == "Hello xlwings!":
        cell.value = "Bye xlwings!"
    else:
        cell.value = "Hello xlwings!"

    # Pass the following back as the response
    return book.json()


@app.post("/capitalize-sheet-names-prompt")
async def capitalize_sheet_names_prompt(data: dict = Body):
    book = xw.Book(json=data)
    book.app.alert(
        prompt="This will capitalize all sheet names!",
        title="Are you sure?",
        buttons="ok_cancel",
        # this is the JS function name that gets called when the user clicks a button
        callback="capitalizeSheetNames",
    )
    return book.json()


@app.post("/capitalize-sheet-names")
async def capitalize_sheet_names(data: dict = Body):
    book = xw.Book(json=data)
    for sheet in book.sheets:
        sheet.name = sheet.name.upper()
    return book.json()


@app.get("/xlwings/alert", response_class=HTMLResponse)
async def alert(
    request: Request, prompt: str, title: str, buttons: str, mode: str, callback: str
):
    """Boilerplate required by book.app.alert() and to show unhandled exceptions"""
    return templates.TemplateResponse(
        "xlwings-alert.html",
        {
            "request": request,
            "prompt": markupsafe.Markup(prompt.replace("\n", "<br>")),
            "title": title,
            "buttons": buttons,
            "mode": mode,
            "callback": callback,
        },
    )


@app.get("/xlwings/custom-functions-meta")
async def custom_functions_meta():
    return xw.pro.custom_functions_meta(custom_functions)


@app.get("/xlwings/custom-functions-code")
async def custom_functions_code():
    return PlainTextResponse(xw.pro.custom_functions_code(custom_functions))


@app.post("/xlwings/custom-functions-call")
async def custom_functions_call(data: dict = Body):
    rv = await xw.pro.custom_functions_call(data, custom_functions)
    return {"result": rv}


# Add xlwings.html as additional source for templates so the /xlwings/alert endpoint
# will find xlwings-alert.html. "mytemplates" can be a dummy if the app doesn't use
# own templates
loader = jinja2.ChoiceLoader(
    [
        jinja2.FileSystemLoader(this_dir / "mytemplates"),
        jinja2.PackageLoader("xlwings", "html"),
    ]
)
templates = Jinja2Templates(directory=this_dir / "mytemplates", loader=loader)

# Serve static files (HTML and icons)
# This could also be handled by an external web server such as nginx, etc.
app.mount("/", StaticFiles(directory=this_dir), name="home")
# Never cache static files
StaticFiles.is_not_modified = lambda *args, **kwargs: False


@app.exception_handler(Exception)
async def xlwings_exception_handler(request, exception):
    # This handles all exceptions, so you may want to make this more restrictive
    return PlainTextResponse(
        str(exception), status_code=status.HTTP_500_INTERNAL_SERVER_ERROR
    )


# Office Scripts and custom functions in Excel on the web require CORS
# Using app.add_middleware won't add the CORS headers if you handle the root "Exception"
# in an exception handler (it would require a specific exception type). Note that
# cors_app is used in the uvicorn.run() call below.
cors_app = CORSMiddleware(
    app=app,
    allow_origins="*",
    allow_methods=["POST"],
    allow_headers=["*"],
)


if __name__ == "__main__":
    import uvicorn

    uvicorn.run(
        "server_fastapi:cors_app",
        host="127.0.0.1",
        port=8000,
        reload=True,
        ssl_keyfile=this_dir.parent / "certs" / "localhost+2-key.pem",
        ssl_certfile=this_dir.parent / "certs" / "localhost+2.pem",
    )
