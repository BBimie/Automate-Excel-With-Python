from flask import Flask, jsonify, request
from util.writer import SalesReport

app = Flask(__name__)


@app.get("/generate_report")
def Generate():
    try:
        SalesReport().generate()
        return jsonify(status=200,
                    message='Report generated, check project folder')
    except Exception as e:
        return jsonify(status=400,
                       message=f'Could not generate report')

if __name__ == "__main__":
    app.run(debug=True)


# python3 -m flask run
# FLASK_APP=app.py FLASK_ENV=development python3 -m flask run
    