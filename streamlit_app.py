"""
Daily email tracker — Streamlit Cloud UI.

Signs the visitor into Google, pulls their sent Gmail messages within a
chosen time window, and lets them download an Excel tracker with columns:
Date | Employee Name | Task Type | Task/Ticket Description or Reference No. | Start Time | End Time

Nothing is written to a server-side file: the workbook is built in memory
per visitor and offered as a download, since this app is shared publicly.

Required secrets (set in Streamlit Cloud's "Secrets" panel, or locally in
.streamlit/secrets.toml — see that file for the template):
    [google]
    client_id = "..."
    client_secret = "..."
    redirect_uri = "https://your-app-name.streamlit.app"
"""

import datetime as dt
import io
import re
from zoneinfo import ZoneInfo

import pandas as pd
import streamlit as st
from google.auth.transport.requests import Request as GoogleRequest
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import Flow
from googleapiclient.discovery import build
from openpyxl import Workbook

SCOPES = ["https://www.googleapis.com/auth/gmail.readonly"]

# Streamlit Cloud runs its servers in UTC, but the date/time inputs in this
# app are wall-clock times entered by IST users. Gmail's internalDate is a
# UTC timestamp, so both sides must be anchored to the same zone or the
# window comparison silently misses everything (see get_sent_emails_in_range).
LOCAL_TZ = ZoneInfo("Asia/Kolkata")

HEADERS = [
    "Date",
    "Employee Name",
    "Task Type",
    "Task/Ticket Description or Reference No.",
    "Start Time",
    "End Time",
]

st.set_page_config(page_title="Daily Email Tracker", page_icon="\U0001F4E7")


def get_flow():
    client_config = {
        "web": {
            "client_id": st.secrets["google"]["client_id"],
            "client_secret": st.secrets["google"]["client_secret"],
            "auth_uri": "https://accounts.google.com/o/oauth2/auth",
            "token_uri": "https://oauth2.googleapis.com/token",
            "redirect_uris": [st.secrets["google"]["redirect_uri"]],
        }
    }
    # PKCE is skipped: the code_verifier it relies on lives in
    # st.session_state, which does not survive the full-page redirect to
    # Google and back. This is a confidential client (has client_secret),
    # so PKCE isn't required for the token exchange to be secure.
    return Flow.from_client_config(
        client_config,
        scopes=SCOPES,
        redirect_uri=st.secrets["google"]["redirect_uri"],
        autogenerate_code_verifier=False,
    )


def credentials_from_session():
    stored = st.session_state["credentials"]
    creds = Credentials(**stored)
    if creds.expired and creds.refresh_token:
        creds.refresh(GoogleRequest())
        st.session_state["credentials"]["token"] = creds.token
    return creds


def ensure_authenticated():
    if "credentials" in st.session_state:
        return credentials_from_session()

    query_params = st.query_params
    if "code" in query_params:
        flow = get_flow()
        flow.fetch_token(code=query_params["code"])
        creds = flow.credentials
        st.session_state["credentials"] = {
            "token": creds.token,
            "refresh_token": creds.refresh_token,
            "token_uri": creds.token_uri,
            "client_id": creds.client_id,
            "client_secret": creds.client_secret,
            "scopes": creds.scopes,
        }
        st.query_params.clear()
        st.rerun()

    flow = get_flow()
    auth_url, _ = flow.authorization_url(access_type="offline", prompt="consent")
    st.title("Daily Email Tracker")
    st.write("Sign in with Google to pull your sent emails into a tracker sheet.")
    st.link_button("Sign in with Google", auth_url)
    st.caption(
        "Your Gmail data is read only for this session and never stored on the server — "
        "the tracker file is generated in memory and offered as a download."
    )
    st.stop()


def get_display_name(service):
    profile = service.users().getProfile(userId="me").execute()
    email_address = profile["emailAddress"]
    results = service.users().messages().list(userId="me", q="in:sent", maxResults=1).execute()
    messages = results.get("messages", [])
    if messages:
        msg = service.users().messages().get(
            userId="me", id=messages[0]["id"], format="metadata", metadataHeaders=["From"]
        ).execute()
        for header in msg["payload"].get("headers", []):
            if header["name"] == "From":
                match = re.match(r'^"?([^"<]+)"?\s*<', header["value"])
                if match:
                    return match.group(1).strip()
    return email_address


def extract_header(headers, name):
    for header in headers:
        if header["name"].lower() == name.lower():
            return header["value"]
    return ""


def parse_recipient_name(to_header):
    match = re.match(r'^"?([^"<]+)"?\s*<', to_header)
    if match:
        return match.group(1).strip()
    match = re.match(r"^([^@<>]+)@", to_header)
    if match:
        return match.group(1).strip()
    return to_header.strip()


def get_sent_emails_in_range(service, start_dt, end_dt):
    # Widen by a day on each side since Gmail's after:/before: date operators
    # aren't guaranteed to use LOCAL_TZ boundaries; the precise window is
    # enforced below with tz-aware comparison.
    query_after = (start_dt.date() - dt.timedelta(days=1)).strftime("%Y/%m/%d")
    query_before = (end_dt.date() + dt.timedelta(days=2)).strftime("%Y/%m/%d")
    query = f"in:sent after:{query_after} before:{query_before}"

    emails = []
    request = service.users().messages().list(userId="me", q=query)
    while request is not None:
        response = request.execute()
        for msg_ref in response.get("messages", []):
            msg = service.users().messages().get(
                userId="me",
                id=msg_ref["id"],
                format="metadata",
                metadataHeaders=["Subject", "To", "Date"],
            ).execute()
            sent_dt = dt.datetime.fromtimestamp(
                int(msg["internalDate"]) / 1000, tz=dt.timezone.utc
            ).astimezone(LOCAL_TZ)
            if not (start_dt <= sent_dt <= end_dt):
                continue
            headers = msg["payload"].get("headers", [])
            subject = extract_header(headers, "Subject") or "(no subject)"
            to_header = extract_header(headers, "To") or "(unknown recipient)"
            recipient = parse_recipient_name(to_header)
            emails.append(
                {
                    "time": sent_dt,
                    "description": f"{subject} — to {recipient}",
                }
            )
        request = service.users().messages().list_next(previous_request=request, previous_response=response)

    emails.sort(key=lambda e: e["time"])
    return emails


def build_workbook(rows):
    wb = Workbook()
    ws = wb.active
    ws.title = "Tracker"
    ws.append(HEADERS)
    for row in rows:
        ws.append(row)
    buffer = io.BytesIO()
    wb.save(buffer)
    buffer.seek(0)
    return buffer


def main():
    creds = ensure_authenticated()
    service = build("gmail", "v1", credentials=creds)
    default_name = get_display_name(service)

    st.title("Daily Email Tracker")
    st.caption(f"Signed in as **{default_name}**")

    col1, col2 = st.columns(2)
    with col1:
        start_date = st.date_input("Start date", dt.date.today())
        start_time = st.time_input("Start time", dt.time(9, 0))
    with col2:
        end_date = st.date_input("End date", dt.date.today())
        end_time = st.time_input("End time", dt.time(17, 0))

    employee_name = st.text_input("Employee name", value=default_name)

    if st.button("Fetch sent emails", type="primary"):
        start_dt = dt.datetime.combine(start_date, start_time, tzinfo=LOCAL_TZ)
        end_dt = dt.datetime.combine(end_date, end_time, tzinfo=LOCAL_TZ)
        if end_dt <= start_dt:
            st.error("End must be after start.")
            st.stop()

        with st.spinner("Fetching sent emails..."):
            emails = get_sent_emails_in_range(service, start_dt, end_dt)

        if not emails:
            st.warning("No sent emails found in that window.")
            st.stop()

        rows = []
        for email in emails:
            time_str = email["time"].strftime("%m/%d/%Y %H:%M:%S")
            date_str = email["time"].strftime("%m/%d/%Y")
            rows.append([date_str, employee_name, "Email Ticket", email["description"], time_str, time_str])

        st.success(f"Found {len(rows)} sent emails.")
        st.dataframe(pd.DataFrame(rows, columns=HEADERS), use_container_width=True)

        buffer = build_workbook(rows)
        st.download_button(
            "Download tracker (.xlsx)",
            data=buffer,
            file_name=f"tracker_{start_date}_{end_date}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

    if st.button("Sign out"):
        st.session_state.pop("credentials", None)
        st.rerun()


if __name__ == "__main__":
    main()
