import msal
import uuid
from flask import Flask, redirect, request, session, url_for
import requests
from os import environ
from config.settings import load_defaults

app = Flask(__name__)
app.secret_key = environ.get('SECRET_KEY', 'fallback_secret_key')
config = load_defaults()

def _build_msal_app(cache=None, authority=None):
    return msal.ConfidentialClientApplication(
        config.CLIENT_ID, authority=authority or config.AUTHORITY,
        client_credential=config.CLIENT_SECRET, token_cache=cache)

def _build_auth_url(authority=None, scopes=None, state=None):
    return _build_msal_app(authority=authority).get_authorization_request_url(
        scopes or config.SCOPE,
        state=state or str(uuid.uuid4()),
        redirect_uri=url_for("authorized", _external=True))

@app.route("/")
def index():
    if not session.get("user"):
        return redirect(url_for("login"))
    return "Logged in successfully!"

@app.route("/login")
def login():
    session["flow"] = _build_auth_url(scopes=config.SCOPE)
    return redirect(session["flow"])

@app.route("/getAToken")
def authorized():
    if request.args.get("error"):
        return f"Error: {request.args['error']} - {request.args.get('error_description')}"
    if request.args.get("code"):
        cache = msal.SerializableTokenCache()
        app = _build_msal_app(cache=cache)
        result = app.acquire_token_by_authorization_code(
            request.args["code"],
            scopes=config.SCOPE,
            redirect_uri=url_for("authorized", _external=True))
        if "error" in result:
            return f"Error acquiring token: {result['error']} - {result.get('error_description')}"
        session["user"] = result.get("id_token_claims")
        session["token_cache"] = cache.serialize()
        return redirect(url_for("index"))

    return "Unknown error."

if __name__ == "__main__":
    app.run(ssl_context='adhoc')  # Enable HTTPS for security if exposed to the web
