from flask import Flask

app = Flask(__name__)

@app.before_first_request
def startup():
    print("Before first request works!")

@app.route('/')
def index():
    return "Hello, Flask!"

if __name__ == '__main__':
    app.run(debug=True)
