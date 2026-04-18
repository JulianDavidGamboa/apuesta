from livereload import Server
from app import app

server = Server(app.wsgi_app)

# Vigila cambios en templates y código Python
server.watch('templates/')
server.watch('app.py')
server.watch('static/', ignore=None)

server.serve(port=5000, debug=True)
