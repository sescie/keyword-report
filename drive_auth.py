"""
drive_auth.py — "Sign in with Google" for the Streamlit app.

No credentials are ever stored on disk or in Streamlit secrets on a
per-user basis — only the app's own OAuth Client ID/Secret (which
identifies the APP to Google, the same way any "Sign in with Google"
button on any website works) lives in Streamlit secrets. Each person's
actual access token lives only in that browser session's
st.session_state, exactly like being logged into any other website —
closing the tab forgets it; they sign in again next time.

Setup (one-time, in Google Cloud Console):
    1. Create a project (or use an existing one).
    2. APIs & Services -> Enable "Google Drive API".
    3. APIs & Services -> OAuth consent screen -> configure it (External
       is fine for a small trusted group; add the people who'll use this
       as "Test users" if the app stays in Testing mode, which avoids
       Google's full verification review).
    4. APIs & Services -> Credentials -> Create Credentials -> OAuth
       client ID -> Application type: "Web application".
    5. Under "Authorized redirect URIs", add your deployed Streamlit
       app's URL exactly, e.g. https://yourname-seo-reports.streamlit.app
       (and http://localhost:8501 too, for local testing).
    6. Copy the Client ID and Client Secret into Streamlit's Secrets
       (Settings -> Secrets, on Streamlit Community Cloud):
           [google_oauth]
           client_id = "....apps.googleusercontent.com"
           client_secret = "...."
           redirect_uri = "https://yourname-seo-reports.streamlit.app"
"""

import streamlit as st
from google_auth_oauthlib.flow import Flow
from google.oauth2.credentials import Credentials

SCOPES = ["https://www.googleapis.com/auth/drive"]


def _client_config():
    cfg = st.secrets["google_oauth"]
    return {
        "web": {
            "client_id": cfg["client_id"],
            "client_secret": cfg["client_secret"],
            "auth_uri": "https://accounts.google.com/o/oauth2/auth",
            "token_uri": "https://oauth2.googleapis.com/token",
            "redirect_uris": [cfg["redirect_uri"]],
        }
    }, cfg["redirect_uri"]


def get_credentials():
    """Returns valid google.oauth2.credentials.Credentials for the
    current session, or None if the person hasn't signed in yet.
    Automatically refreshes an expired token using the stored refresh
    token, so a long review session doesn't suddenly fail partway."""
    creds_dict = st.session_state.get("drive_credentials")
    if not creds_dict:
        return None
    creds = Credentials(**creds_dict)
    if creds.expired and creds.refresh_token:
        import google.auth.transport.requests
        creds.refresh(google.auth.transport.requests.Request())
        st.session_state["drive_credentials"] = _creds_to_dict(creds)
    return creds


def _creds_to_dict(creds):
    return {
        "token": creds.token, "refresh_token": creds.refresh_token,
        "token_uri": creds.token_uri, "client_id": creds.client_id,
        "client_secret": creds.client_secret, "scopes": creds.scopes,
    }


def render_sign_in():
    """Call this when get_credentials() returns None. Shows a "Sign in
    with Google" link; handles the redirect back (Google appends a
    ?code=... query param to the app's own URL) by exchanging that code
    for a real token, entirely within one Streamlit rerun cycle."""
    client_config, redirect_uri = _client_config()
    flow = Flow.from_client_config(client_config, scopes=SCOPES, redirect_uri=redirect_uri)

    query_params = st.query_params
    if "code" in query_params:
        try:
            flow.fetch_token(code=query_params["code"])
            st.session_state["drive_credentials"] = _creds_to_dict(flow.credentials)
            st.query_params.clear()
            st.rerun()
        except Exception as e:
            st.error(f"Sign-in failed: {e}")
            st.query_params.clear()
            return

    auth_url, _ = flow.authorization_url(access_type="offline", include_granted_scopes="true",
                                          prompt="consent")
    st.title("📊 SEO Monthly Report Builder")
    st.write("Sign in with the Google account that has access to your report folders in Drive.")
    st.link_button("🔐 Sign in with Google", auth_url, type="primary")


def render_sign_out_button():
    if st.sidebar.button("Sign out"):
        st.session_state.pop("drive_credentials", None)
        st.rerun()
