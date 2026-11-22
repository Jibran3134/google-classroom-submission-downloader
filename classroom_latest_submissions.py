"""Download currently attached Google Classroom assignment files, by student.

Setup:
  1. In Google Cloud Console enable Classroom API and Google Drive API.
  2. Configure OAuth consent (add your account as a test user if applicable).
  3. Create OAuth 2.0 Desktop App credentials, save as credentials.json here.
  4. pip install --upgrade google-api-python-client google-auth-httplib2 google-auth-oauthlib
  5. python classroom_latest_submissions.py

This script is read-only. It downloads the *current* attachments of submissions
whose Classroom state is TURNED_IN or RETURNED. It does NOT infer the final file
from Drive folder modification timestamps; removed/replaced old attachments are
not included. Google Workspace documents are exported to editable Office formats.
"""

import csv
import re
from datetime import datetime, timezone
from pathlib import Path

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
from googleapiclient.http import MediaIoBaseDownload

BASE = Path(__file__).resolve().parent
CREDENTIALS = BASE / "credentials.json"
TOKEN = BASE / "token.json"
SCOPES = [
    "https://www.googleapis.com/auth/classroom.courses.readonly",
    "https://www.googleapis.com/auth/classroom.coursework.students.readonly",
    "https://www.googleapis.com/auth/classroom.rosters.readonly",
    "https://www.googleapis.com/auth/classroom.profile.emails",
    "https://www.googleapis.com/auth/drive.readonly",
]

EXPORTS = {
    "application/vnd.google-apps.document": (
        "application/vnd.openxmlformats-officedocument.wordprocessingml.document", ".docx"
    ),
    "application/vnd.google-apps.spreadsheet": (
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", ".xlsx"
    ),
    "application/vnd.google-apps.presentation": (
        "application/vnd.openxmlformats-officedocument.presentationml.presentation", ".pptx"
    ),
    "application/vnd.google-apps.drawing": ("application/pdf", ".pdf"),
}


def clean(value, limit=100):
    value = re.sub(r'[<>:"/\\|?*\x00-\x1f]', "_", str(value or "unknown"))
    value = re.sub(r"\s+", " ", value).strip(" .")
    return (value[:limit].rstrip(" .") or "unknown")


def auth():
    if not CREDENTIALS.exists():
        raise FileNotFoundError(
            "credentials.json not found beside this script. Follow setup instructions "
            "at https://developers.google.com/workspace/classroom/quickstart/python"
        )
    credentials = None
    if TOKEN.exists():
        try:
            credentials = Credentials.from_authorized_user_file(str(TOKEN), SCOPES)
        except ValueError:
            print("Existing token is invalid; opening Google login again.")
    if not credentials or not credentials.valid:
        if credentials and credentials.expired and credentials.refresh_token:
            try:
                credentials.refresh(Request())
            except Exception:
                credentials = None
        if not credentials or not credentials.valid:
            flow = InstalledAppFlow.from_client_secrets_file(str(CREDENTIALS), SCOPES)
            credentials = flow.run_local_server(port=0)
        TOKEN.write_text(credentials.to_json(), encoding="utf-8")
    return credentials


def get_pages(call, result_key, **params):
    items = []
    token = None
    while True:
        kw = dict(params)
        if token:
            kw["pageToken"] = token
        result = call(**kw).execute()
        items.extend(result.get(result_key, []))
        token = result.get("nextPageToken")
        if not token:
            return items


def choose(label, items, title):
    if not items:
        raise RuntimeError(f"No {label} available to select.")
    print(f"\nChoose a {label}:")
    for i, item in enumerate(items, start=1):
        print(f" {i:>3}. {title(item)}")
    while True:
        entered = input(f"Enter number (1-{len(items)}): ").strip()
        if entered.isdigit() and 1 <= int(entered) <= len(items):
            return items[int(entered) - 1]
        print("Please enter a valid number.")


def latest_turn_in(submission):
    timestamps = []
    for entry in submission.get("submissionHistory", []):
        state = entry.get("stateHistory", {})
        if state.get("state") == "TURNED_IN" and state.get("stateTimestamp"):
            timestamps.append(state["stateTimestamp"])
    return max(timestamps, default="")


def folder_label(student, uid):
    profile = student.get("profile", {})
    name = profile.get("name", {}).get("fullName", "Student")
    email = profile.get("emailAddress", "")
    local = email.split("@", 1)[0]
    # Use a roll number when it can be recognized; otherwise use email or user ID.
    m = re.search(r"(?:^|\W)(\d{2})i[-_ ]?(\d{4})(?:$|\W)", local, re.I)
    n = re.search(r"(?:^|\W)i(\d{2})(\d{4})(?:$|\W)", local, re.I)
    if m or n:
        match = m or n
        identity = f"{match.group(1)}i-{match.group(2)}"
    else:
        identity = local or uid
    return clean(f"{identity} - {name}"), name, email


def nonconflicting_path(folder, filename):
    filename = clean(filename, 130)
    original = folder / filename
    if not original.exists():
        return original
    stem, suffix = original.stem, original.suffix
    index = 2
    while True:
        proposal = folder / f"{stem} ({index}){suffix}"
        if not proposal.exists():
            return proposal
        index += 1


def download_file(drive, file_id, folder):
    metadata = drive.files().get(
        fileId=file_id,
        fields="id,name,mimeType,capabilities(canDownload),webViewLink,shortcutDetails",
        supportsAllDrives=True,
    ).execute()
    mime = metadata.get("mimeType", "")
    if mime == "application/vnd.google-apps.shortcut":
        info = metadata.get("shortcutDetails", {})
        target = info.get("targetId")
        if not target:
            raise RuntimeError("Google Drive shortcut has no accessible target")
        return download_file(drive, target, folder)
    if metadata.get("capabilities", {}).get("canDownload") is False:
        raise PermissionError("Downloading is restricted by the owner")
    name = metadata.get("name") or file_id
    if mime in EXPORTS:
        export_mime, extension = EXPORTS[mime]
        if not name.lower().endswith(extension):
            name += extension
        request = drive.files().export_media(fileId=file_id, mimeType=export_mime)
    elif mime.startswith("application/vnd.google-apps."):
        raise RuntimeError(f"No download/export format configured for {mime}")
    else:
        request = drive.files().get_media(fileId=file_id, supportsAllDrives=True)
    path = nonconflicting_path(folder, name)
    # Partial files are removed when a download fails.
    try:
        with path.open("wb") as f:
            downloader = MediaIoBaseDownload(f, request, chunksize=1024 * 1024)
            done = False
            while not done:
                _, done = downloader.next_chunk(num_retries=3)
    except Exception:
        path.unlink(missing_ok=True)
        raise
    return path


def main():
    print("Google Classroom — current submissions downloader (READ ONLY)")
    creds = auth()
    classroom = build("classroom", "v1", credentials=creds, cache_discovery=False)
    drive = build("drive", "v3", credentials=creds, cache_discovery=False)

    courses = get_pages(
        classroom.courses().list, "courses", teacherId="me", pageSize=100
    )
    courses.sort(key=lambda c: c.get("name", "").casefold())
    course = choose("course", courses, lambda c: c.get("name", c["id"]))

    work = get_pages(
        classroom.courses().courseWork().list, "courseWork",
        courseId=course["id"], pageSize=100,
    )
    assignments = [x for x in work if x.get("workType") == "ASSIGNMENT"
                   and x.get("state") == "PUBLISHED"]
    assignments.sort(key=lambda a: a.get("creationTime", ""), reverse=True)
    assignment = choose("assignment", assignments,
                        lambda a: a.get("title", a["id"]))

    print("\nReading the class roster and submission records...")
    roster = get_pages(
        classroom.courses().students().list, "students",
        courseId=course["id"], pageSize=100,
    )
    students = {r.get("userId", ""): r for r in roster}
    submissions = get_pages(
        classroom.courses().courseWork().studentSubmissions().list,
        "studentSubmissions", courseId=course["id"],
        courseWorkId=assignment["id"], pageSize=100,
    )

    snapshot = datetime.now(timezone.utc).strftime("%Y%m%d_%H%M%S_%f_UTC")
    root = (BASE / "Classroom_Submissions" / clean(course["name"]) /
            clean(assignment["title"]) / snapshot)
    root.mkdir(parents=True, exist_ok=True)
    report = []
    successes, failures, ignored = 0, 0, 0
    print(f"\nChecking {len(submissions)} student records...")

    for submission in sorted(submissions, key=lambda x: x.get("userId", "")):
        uid = submission.get("userId", "")
        student_name, name, email = folder_label(students.get(uid, {}), uid)
        state = submission.get("state", "UNKNOWN")
        timestamp = latest_turn_in(submission)
        common = dict(student=name, student_email=email, user_id=uid,
                      classroom_status=state, last_turn_in_utc=timestamp,
                      submission_updated_utc=submission.get("updateTime", ""),
                      submission_url=submission.get("alternateLink", ""))
        if state not in ("TURNED_IN", "RETURNED"):
            ignored += 1
            report.append(dict(**common, filename="", attachment_url="",
                               download_status="not currently submitted"))
            continue
        attachments = submission.get("assignmentSubmission", {}).get("attachments", [])
        if not attachments:
            report.append(dict(**common, filename="", attachment_url="",
                               download_status="no current attachments"))
            continue

        folder = root / student_name
        folder.mkdir(parents=True, exist_ok=True)
        print(f" - {student_name} ({state}): {len(attachments)} attachment(s)")
        for attachment in attachments:
            drive_file = attachment.get("driveFile", {})
            if drive_file.get("id"):
                fid = drive_file["id"]
                try:
                    path = download_file(drive, fid, folder)
                    successes += 1
                    report.append(dict(**common, filename=str(path.relative_to(root)),
                                       attachment_url=drive_file.get("alternateLink", ""),
                                       download_status="downloaded"))
                except Exception as e:
                    failures += 1
                    print(f"    FAILED to download file {fid}: {e}")
                    report.append(dict(**common, filename=drive_file.get("title", fid),
                                       attachment_url=drive_file.get("alternateLink", ""),
                                       download_status=f"failed: {str(e)[:180]}"))
            else:
                link = attachment.get("link", {}).get("url", "")
                if not link:
                    link = attachment.get("youTubeVideo", {}).get("alternateLink", "")
                report.append(dict(**common, filename="", attachment_url=link,
                                   download_status="link/non-file attachment: recorded only"))

    columns = ["student", "student_email", "user_id", "classroom_status",
               "last_turn_in_utc", "submission_updated_utc", "submission_url",
               "filename", "attachment_url", "download_status"]
    output_csv = root / "submission_report.csv"
    with output_csv.open("w", encoding="utf-8-sig", newline="") as file:
        writer = csv.DictWriter(file, fieldnames=columns)
        writer.writeheader()
        writer.writerows(report)

    print(f"\nCompleted: {successes} files downloaded, {failures} failures, "
          f"{ignored} not currently submitted.")
    print(f"Files and CSV report saved to:\n  {root}\n")
    if failures:
        print("Check the report CSV for any restricted or failed downloads.")


if __name__ == "__main__":
    try:
        main()
    except (HttpError, FileNotFoundError, RuntimeError) as exc:
        print(f"\nError: {exc}")
        print("Check Google account access, enabled APIs, and OAuth setup.")
        raise SystemExit(1)
    except KeyboardInterrupt:
        print("\nCancelled.")
        raise SystemExit(130)
