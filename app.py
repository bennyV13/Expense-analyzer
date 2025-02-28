from flask import Flask, jsonify
from expense_analyzer import 

app = Flask(__name__)

def app_logic():
    # Replace this with your application's functionality.
    return "Hello from your Python app!"

@app.route('/')
def index():
    result = app_logic()
    return jsonify({"message": result})

if __name__ == '__main__':
    # Run the app on all available interfaces, port 5000.
    app.run(host='0.0.0.0', port=5000)
