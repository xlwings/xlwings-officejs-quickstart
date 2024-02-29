"""
This is a minimal Dash app (taken from https://dash.plotly.com/minimal-app) with xlwings
Server integrated. If you want to build Office.js add-ins, you need additional
endpoints, see: server_flask.py
"""

import pandas as pd
import plotly.express as px
import xlwings as xw
from dash import Dash, Input, Output, callback, dcc, html
from flask import Flask, Response, request

server = Flask(__name__)

# Dash app
app = Dash(server=server)


df = pd.read_csv(
    "https://raw.githubusercontent.com/plotly/datasets/master/gapminder_unfiltered.csv"
)


app.layout = html.Div(
    [
        html.H1(children="Title of Dash App", style={"textAlign": "center"}),
        dcc.Dropdown(df.country.unique(), "Canada", id="dropdown-selection"),
        dcc.Graph(id="graph-content"),
    ]
)


@callback(Output("graph-content", "figure"), Input("dropdown-selection", "value"))
def update_graph(value):
    dff = df[df.country == value]
    return px.line(dff, x="year", y="pop")


# xlwings Server
@server.route("/hello", methods=["POST"])
def hello():
    with xw.Book(json=request.json) as book:
        sheet = book.sheets[0]
        cell = sheet["A1"]
        if cell.value == "Hello xlwings!":
            cell.value = "Bye xlwings!"
        else:
            cell.value = "Hello xlwings!"
        return book.json()


@server.errorhandler(Exception)
def xlwings_exception_handler(error):
    # This handles all exceptions, so you may want to make this more restrictive
    return Response(str(error), status=500)


if __name__ == "__main__":
    # Run the Flask app, not the Dash app
    server.run(debug=True, port=8000)
