from flask import Flask, jsonify, request
from util.writer import SalesReport

app = Flask(__name__)


@app.post("/generate_report")
def entryPoint():
    SalesReport().generate()

if __name__ == "__main__":
    app.run(debug=True)


# python3 -m flask run
# FLASK_APP=app.py FLASK_ENV=development python3 -m flask run
    