from flask import Flask, redirect, request, session, url_for
from flask_session import Session
import msal
import os
import config
import redis

# Load configuration
app_config = config.load_defaults()

CLIENT_ID = app_config['CLIENT_ID']
CLIENT_SECRET = app_config['CLIENT_SECRET']
TENANT_ID = app_config['TENANT_ID']
REDIRECT_URI = app_config['REDIRECT_URI']
AUTHORITY = f'https://login.microsoftonline.com/{TENANT_ID}'
SCOPE = app_config['SCOPE']

app = Flask(__name__)
app.secret_key = os.urandom(24)

# Configure server-side session storage
app.config["SESSION_TYPE"] = "redis"
app.config["SESSION_PERMANENT"] = False
app.config["SESSION_USE_SIGNER"] = True
app.config["SESSION_REDIS"] = redis.from_url("redis://localhost:6379")
Session(app)

# Initialize MSAL ConfidentialClientApplication
client = msal.ConfidentialClientApplication(
    CLIENT_ID,
    authority=AUTHORITY,
    client_credential=CLIENT_SECRET,
)

@app.route('/')
def index():
    if not session.get("user"):
        app.logger.debug("User not authenticated, redirecting to login.")
        return redirect(url_for("login"))
    app.logger.debug("User authenticated.")
    return 'Logged in successfully!'

@app.route('/login')
def login():
    auth_url = client.get_authorization_request_url(SCOPE, redirect_uri=REDIRECT_URI)
    app.logger.debug(f"Authorization URL: {auth_url}")
    return redirect(auth_url)

@app.route('/callback')
def callback():
    code = request.args.get('code')
    if code:
        result = client.acquire_token_by_authorization_code(code, scopes=SCOPE, redirect_uri=REDIRECT_URI)
        if "access_token" in result:
            session["user"] = {
                "access_token": result["access_token"],
                "id_token_claims": result.get("id_token_claims", {}),
                "account": result["account"]
            }
            app.logger.debug("Access token acquired, user authenticated.")
            return redirect(url_for("index"))
        app.logger.error(f"Login failure: {result.get('error')} - {result.get('error_description')}")
        return "Login failure: " + str(result.get("error")) + " - " + str(result.get("error_description"))
    app.logger.error("Missing authorization code.")
    return "Missing authorization code."

@app.route('/logout')
def logout():
    session.clear()
    return redirect(AUTHORITY + "/oauth2/v2.0/logout" +
                    "?post_logout_redirect_uri=" + url_for("index", _external=True))

def get_token():
    if "user" in session:
        return session["user"]["access_token"]
    return None

if __name__ == "__main__":
    app.run(debug=True)
