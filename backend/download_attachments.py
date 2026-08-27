import os
import re
import json
import imaplib
import email
from email.header import decode_header
from email.utils import parseaddr
from pathlib import Path

# Configuration
WORKSPACE_DIR = Path("D:/SRI BALAJI TRADERS")
ENV_PATH = WORKSPACE_DIR / ".env"
PROCESSED_DB_PATH = WORKSPACE_DIR / "processed_emails.json"
IMAP_SERVER = "imap.gmail.com"

# Root labels to watch (case-insensitive)
TARGET_ROOT_LABELS = ["CORTEVA", "NEW GEN", "FMC"]

def load_env(path):
    """Load environment variables from a .env file."""
    env_vars = {}
    if os.path.exists(path):
        with open(path, "r", encoding="utf-8") as f:
            for line in f:
                line = line.strip()
                if not line or line.startswith("#"):
                    continue
                if "=" in line:
                    key, val = line.split("=", 1)
                    env_vars[key.strip()] = val.strip().strip('"').strip("'")
    return env_vars

def decode_mime_words(s):
    """Decode MIME encoded words in email headers (e.g. filenames)."""
    if not s:
        return ""
    try:
        decoded_parts = decode_header(s)
        decoded_str = ""
        for part, encoding in decoded_parts:
            if isinstance(part, bytes):
                if encoding:
                    try:
                        decoded_str += part.decode(encoding)
                    except Exception:
                        decoded_str += part.decode('utf-8', errors='replace')
                else:
                    decoded_str += part.decode('utf-8', errors='replace')
            else:
                decoded_str += part
        return decoded_str
    except Exception as e:
        print(f"Error decoding header {s}: {e}")
        return str(s)

def sanitize_filename(filename):
    """Sanitize filename to be safe for Windows file systems."""
    if not filename:
        return ""
    filename = os.path.basename(filename)
    filename = re.sub(r'[\r\n\t\x00-\x1f]', ' ', filename)
    invalid_chars = '<>:"/\\|?*'
    for c in invalid_chars:
        filename = filename.replace(c, '_')
    filename = re.sub(r'\s+', ' ', filename).strip()
    return filename

def load_processed_db():
    """Load the processed emails database."""
    if PROCESSED_DB_PATH.exists():
        try:
            with open(PROCESSED_DB_PATH, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception as e:
            print(f"Warning: Failed to load processed emails database: {e}. Starting fresh.")
    return {}

def save_processed_db(db):
    """Save the processed emails database."""
    try:
        with open(PROCESSED_DB_PATH, "w", encoding="utf-8") as f:
            json.dump(db, f, indent=4)
    except Exception as e:
        print(f"Error saving processed emails database: {e}")

def parse_mailbox_name(mailbox_bytes):
    """Parse mailbox name from IMAP LIST response."""
    try:
        line = mailbox_bytes.decode('utf-8', errors='replace')
        # Format is typically: (\HasNoChildren) "/" "CORTEVA/SURYAPET/SURYAPET 2026-2027"
        # Or: (\HasNoChildren) "/" CORTEVA
        match = re.search(r'\((.*?)\)\s+"(.*?)"\s+(.*)', line)
        if match:
            name_part = match.group(3).strip()
            if name_part.startswith('"') and name_part.endswith('"'):
                name_part = name_part[1:-1].replace('\\"', '"').replace('\\\\', '\\')
            return name_part
    except Exception as e:
        print(f"Error parsing mailbox line {mailbox_bytes}: {e}")
    return None

def process_mailbox(mail, mailbox_name, processed_db):
    """Select mailbox, find attachments in new emails, and download them."""
    # Split mailbox path by '/'
    parts = [p.strip() for p in mailbox_name.split('/') if p.strip()]
    if not parts:
        return

    root_label = parts[0]
    if root_label.upper() not in TARGET_ROOT_LABELS:
        return

    # Check if this mailbox contains TBM bills or POs
    is_tbm_bills = any('tbm' in p.lower() for p in parts)

    # Directly mirror the Gmail label structure into local workspace directory
    target_folder = WORKSPACE_DIR.joinpath(*parts)
    target_folder.mkdir(parents=True, exist_ok=True)

    print(f"\nProcessing Gmail label: {mailbox_name}")
    print(f"Target local folder: {target_folder}")

    # Select the mailbox (readonly so we don't change read status of emails)
    try:
        status, data = mail.select(f'"{mailbox_name}"', readonly=True)
        if status != 'OK':
            print(f"Error selecting mailbox '{mailbox_name}': {status}")
            return
    except Exception as e:
        print(f"Exception selecting mailbox '{mailbox_name}': {e}")
        return

    # Search for all UIDs in the mailbox
    try:
        status, data = mail.uid('search', None, 'ALL')
        if status != 'OK':
            print(f"Error searching mailbox: {status}")
            return
    except Exception as e:
        print(f"Exception searching mailbox: {e}")
        return

    uids = data[0].split()
    if not uids:
        print("No emails found in this label.")
        return

    # Initialize processed list for this mailbox if not exists
    if mailbox_name not in processed_db:
        processed_db[mailbox_name] = []

    new_downloads_count = 0
    skipped_count = 0

    for uid_bytes in uids:
        uid = uid_bytes.decode('utf-8')
        
        # Check if already processed
        if uid in processed_db[mailbox_name]:
            skipped_count += 1
            continue

        print(f"  Fetching email UID: {uid}...")
        try:
            # First fetch only Subject and From headers to verify sender and thread
            status, header_data = mail.uid('fetch', uid, '(BODY[HEADER.FIELDS (SUBJECT FROM)])')
            if status != 'OK' or not header_data:
                print(f"    Failed to fetch headers for email UID {uid}: {status}")
                continue

            header_bytes = None
            for response_part in header_data:
                if isinstance(response_part, tuple):
                    header_bytes = response_part[1]
                    break

            if not header_bytes:
                print(f"    No header data found for UID {uid}")
                continue

            header_msg = email.message_from_bytes(header_bytes)
            from_header = decode_mime_words(header_msg.get('From', ''))
            subject = decode_mime_words(header_msg.get('Subject', '(No Subject)'))

            # Parse sender email
            from_name, from_email = parseaddr(from_header)
            from_email = from_email.strip().lower()

            # Check if sender is correct based on company (Bypassed for TBM bills)
            if not is_tbm_bills:
                company = root_label.upper()
                if company == "FMC":
                    target_sender = "newgen.fmc@gmail.com"
                else:
                    target_sender = "ordersender-prod@ansmtp.ariba.com"
                    
                if from_email != target_sender:
                    # Skip and mark as processed so we don't query it again
                    processed_db[mailbox_name].append(uid)
                    save_processed_db(processed_db)
                    continue

            # Fetch the full email to download attachments
            print(f"    Processing email UID: {uid} | Subject: {subject}")
            status, msg_data = mail.uid('fetch', uid, '(RFC822)')
            if status != 'OK' or not msg_data:
                print(f"      Failed to fetch full email UID {uid}: {status}")
                continue

            raw_email = None
            for response_part in msg_data:
                if isinstance(response_part, tuple):
                    raw_email = response_part[1]
                    break

            if not raw_email:
                print(f"      No email data found for UID {uid}")
                continue

            msg = email.message_from_bytes(raw_email)

            # Walk through email parts to find attachments
            has_attachments = False
            for part in msg.walk():
                if part.get_content_maintype() == 'multipart':
                    continue

                raw_filename = part.get_filename() or part.get_param('name', header='content-type') or part.get_param('filename', header='content-disposition')
                if not raw_filename:
                    continue

                filename = sanitize_filename(decode_mime_words(raw_filename))
                if not filename:
                    continue

                payload = part.get_payload(decode=True)
                if payload:
                    filepath = target_folder / filename
                    
                    # Resolve naming conflicts if a file with the same name already exists
                    counter = 1
                    base_path = filepath
                    while filepath.exists():
                        stem = Path(base_path).stem
                        suffix = Path(base_path).suffix
                        filepath = target_folder / f"{stem}_{counter}{suffix}"
                        counter += 1

                    try:
                        with open(filepath, 'wb') as f:
                            f.write(payload)
                        print(f"      Downloaded: {filepath.name}")
                        has_attachments = True
                        new_downloads_count += 1
                    except Exception as e:
                        print(f"      Error saving attachment '{filename}': {e}")

            # Mark as processed
            processed_db[mailbox_name].append(uid)
            save_processed_db(processed_db)

        except Exception as e:
            print(f"    Error processing email UID {uid}: {e}")

    print(f"  Done. Downloaded {new_downloads_count} new attachment(s), skipped {skipped_count} already processed email(s).")

def main():
    print("==================================================")
    print("  Gmail Attachment Downloader & Organizer Started")
    print("==================================================")
    
    # Load credentials
    env_vars = load_env(ENV_PATH)
    email_user = env_vars.get("GMAIL_EMAIL")
    email_pass = env_vars.get("GMAIL_APP_PASSWORD")

    if not email_user or not email_pass:
        print(f"Error: Credentials not found in {ENV_PATH}")
        print("Please set GMAIL_EMAIL and GMAIL_APP_PASSWORD in your .env file.")
        print("Format:")
        print("GMAIL_EMAIL=your_email@gmail.com")
        print("GMAIL_APP_PASSWORD=xxxx xxxx xxxx xxxx")
        return

    # Load processed DB
    processed_db = load_processed_db()

    # Connect and log in
    print(f"Connecting to {IMAP_SERVER}...")
    try:
        mail = imaplib.IMAP4_SSL(IMAP_SERVER)
        print("Logging in...")
        mail.login(email_user, email_pass)
        print("Login successful!")
    except Exception as e:
        print(f"Connection/Authentication failed: {e}")
        print("Please verify your email address and App Password.")
        print("Ensure 'IMAP Access' is enabled in your Gmail Settings -> Forwarding and POP/IMAP.")
        return

    # List all mailboxes
    print("Retrieving labels list from Gmail...")
    try:
        status, mailboxes_list = mail.list()
        if status != 'OK':
            print(f"Failed to retrieve mailboxes: {status}")
            return
    except Exception as e:
        print(f"Error listing mailboxes: {e}")
        return

    # Find matching labels
    matching_mailboxes = []
    for m in mailboxes_list:
        mailbox_name = parse_mailbox_name(m)
        if mailbox_name:
            parts = [p.strip() for p in mailbox_name.split('/') if p.strip()]
            if parts:
                root_label = parts[0]
                if root_label.upper() in TARGET_ROOT_LABELS:
                    matching_mailboxes.append(mailbox_name)

    if not matching_mailboxes:
        print("No matching folders/labels found in Gmail.")
        print(f"Looked for label patterns starting with: {', '.join(TARGET_ROOT_LABELS)}")
        print("Example structure expected: CORTEVA/SURYAPET/SURYAPET 2026-2027")
    else:
        print(f"Found {len(matching_mailboxes)} matching label(s) to scan:")
        for m in matching_mailboxes:
            print(f"  - {m}")
        
        # Process each mailbox
        for m in matching_mailboxes:
            process_mailbox(mail, m, processed_db)

    # Logout
    try:
        print("\nLogging out...")
        mail.logout()
        print("Logged out successfully.")
    except Exception:
        pass

    print("==================================================")
    print("  Processing Complete.")
    print("==================================================")

if __name__ == "__main__":
    main()
