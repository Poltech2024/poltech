from flask import Flask

app = Flask(__name__)

@app.route("/")
def inicio():
    return "<h1>¡Hola! Bienvenido a Poltech 🚀</h1><p>Mi primera página web con Python.</p>"

if __name__ == "__main__":
    app.run(debug=True)