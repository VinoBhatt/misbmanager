"""Authentication and static hosting shared by local Flask and Workers WSGI."""
import hmac
import os
from pathlib import Path
from flask import Response, jsonify, request, redirect
from itsdangerous import URLSafeTimedSerializer, BadSignature, SignatureExpired
from werkzeug.exceptions import HTTPException
from storage import cloud


def setting(name):
    return str(getattr(cloud(), name, '') or '') if cloud() else os.environ.get(name, '')


def local_preview():
    return not cloud() and os.environ.get('MISB_LOCAL_PREVIEW') == '1'


def authenticated():
    if local_preview():
        return True
    secret = setting('SESSION_SECRET')
    if len(secret) < 32:
        return False
    try:
        return URLSafeTimedSerializer(secret).loads(request.cookies.get('misb_session',''), max_age=28800) == 'manager'
    except (BadSignature, SignatureExpired):
        return False


def asset(path):
    if cloud():
        from pyodide.ffi import run_sync
        res = run_sync(cloud().ASSETS.fetch('https://assets.local/' + path))
        body = run_sync(res.bytes())
        return Response(body, status=res.status, headers=res.headers)
    from flask import send_from_directory
    return send_from_directory(Path(__file__).resolve().parent.parent / 'public', path)


def configure_web(app):
    @app.before_request
    def protect():
        public = request.path in ('/login','/api/login','/static/style.css','/static/login.js')
        if not public and not authenticated():
            if request.path.startswith('/api/'):
                return jsonify(error='Please sign in to continue.'),401
            return redirect('/login')
        if request.method not in ('GET','HEAD','OPTIONS'):
            origin = request.headers.get('Origin')
            if origin and origin.rstrip('/') != request.host_url.rstrip('/'):
                return jsonify(error='Cross-origin requests are not allowed.'),403
            if request.headers.get('Sec-Fetch-Site') == 'cross-site':
                return jsonify(error='Cross-origin requests are not allowed.'),403
        if request.is_json and request.method in ('POST','PUT','PATCH'):
            if not isinstance(request.get_json(silent=True),dict):
                return jsonify(error='A JSON object is required.'),400

    @app.after_request
    def headers(response):
        response.headers['Cache-Control']='no-store'
        response.headers['X-Content-Type-Options']='nosniff'
        response.headers['X-Frame-Options']='DENY'
        response.headers['Referrer-Policy']='same-origin'
        response.headers['Content-Security-Policy']="default-src 'self'; script-src 'self'; style-src 'self' 'unsafe-inline'; img-src 'self' data:; object-src 'none'; base-uri 'self'; frame-ancestors 'none'; form-action 'self'"
        return response

    @app.get('/login')
    def login_page():
        return asset('login.html')

    @app.post('/api/login')
    def login():
        password, secret = setting('APP_PASSWORD'), setting('SESSION_SECRET')
        if len(password)<12 or len(secret)<32:
            return jsonify(error='The administrator must configure APP_PASSWORD and SESSION_SECRET before sign-in.'),503
        candidate = (request.get_json(silent=True) or {}).get('password','')
        if not isinstance(candidate,str) or not hmac.compare_digest(candidate.encode(),password.encode()):
            return jsonify(error='Incorrect password.'),401
        response=jsonify(ok=True)
        response.set_cookie('misb_session', URLSafeTimedSerializer(secret).dumps('manager'),
                            max_age=28800, httponly=True, secure=request.is_secure, samesite='Strict')
        return response

    @app.post('/api/logout')
    def logout():
        response=jsonify(ok=True)
        response.delete_cookie('misb_session')
        return response

    @app.errorhandler(HTTPException)
    def http_error(error):
        return jsonify(error=error.description),error.code

    @app.errorhandler(Exception)
    def server_error(error):
        app.logger.exception('Request failed')
        return jsonify(error='The request could not be completed. Check the source files or contact the administrator.'),500

    @app.get('/')
    def index():
        return asset('index.html')

    @app.get('/static/<path:path>')
    def static(path):
        return asset('static/'+path)
