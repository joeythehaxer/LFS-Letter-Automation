import msal
from flask import Flask, redirect, request, session, url_for
import requests
from config.settings import load_defaults

app = Flask(__name__)
app.secret_key = 'your_secret_key'
config = load_defaults()

def _build_msal_app(cache=None, authority=None):
    return msal.ConfidentialClientApplication(
        config.CLIENT_ID, authority=authority or config.AUTHORITY,
        client_credential=config.CLIENT_SECRET, token_cache=cache)

def _build_auth_url(authority=None, scopes=None, state=None):
    return _build_msal_app(authority=authority).get_authorization_request_url(
        scopes or [],
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
        return request.args["error"] + ": " + request.args["error_description"]
    if request.args.get("code"):
        cache = msal.SerializableTokenCache()
        result = _build_msal_app(cache=cache).acquire_token_by_authorization_code(
            request.args["code"],
            scopes=config.SCOPE,
            redirect_uri=url_for("authorized", _external=True))
        if "error" in result:
            return result["error"] + ": " + result.get("error_description")
        session["user"] = result.get("id_token_claims")
        session["token_cache"] = cache.serialize()
    return redirect(url_for("index"))

if __name__ == "__main__":
    app.run()
