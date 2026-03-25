"""
Reads today's Outlook calendar appointments and sends emails via Outlook.
Handles UTC timezone conversion for local time display.
"""

import datetime
import time
import pythoncom
import win32com.client
from typing import List, Optional
from utils.logger import get_logger

logger = get_logger(__name__)


def _utc_offset() -> datetime.timedelta:
    import time as _time
    if _time.daylight and _time.localtime().tm_isdst:
        offset_seconds = -_time.altzone
    else:
        offset_seconds = -_time.timezone
    return datetime.timedelta(seconds=offset_seconds)


def _get_outlook(retries=3, delay=2.0):
    pythoncom.CoInitialize()
    for attempt in range(retries):
        try:
            outlook = win32com.client.Dispatch("Outlook.Application")
            return outlook
        except Exception as e:
            if attempt < retries - 1:
                time.sleep(delay)
            else:
                logger.error(f"Could not connect to Outlook: {e}")
                return None


def _get_meetings_from_folder(folder, today, offset):
    meetings = []
    try:
        items = folder.Items
        items.IncludeRecurrences = True
        items.Sort("[Start]")

        yesterday = today - datetime.timedelta(days=1)
        tomorrow  = today + datetime.timedelta(days=2)
        start_str = f"{yesterday.month}/{yesterday.day}/{yesterday.year} 12:00 AM"
        end_str   = f"{tomorrow.month}/{tomorrow.day}/{tomorrow.year} 12:00 AM"
        restriction = f"[Start] >= '{start_str}' AND [Start] <= '{end_str}'"

        try:
            source = items.Restrict(restriction)
        except Exception:
            source = items

        count = 0
        for item in source:
            if count > 500:
                break
            count += 1
            try:
                start = item.Start
                end   = item.End
                start_utc = datetime.datetime(
                    start.year, start.month, start.day,
                    start.hour, start.minute, start.second)
                end_utc = datetime.datetime(
                    end.year, end.month, end.day,
                    end.hour, end.minute, end.second)
                start_local = start_utc + offset
                end_local   = end_utc   + offset

                if start_local.date() != today:
                    continue

                meetings.append({
                    "subject":  item.Subject or "Untitled Meeting",
                    "start":    start_local,
                    "end":      end_local,
                    "location": getattr(item, "Location", "") or "",
                    "duration": max(1, int((end_local - start_local).seconds / 60)),
                })
            except Exception as e:
                logger.warning(f"Skipping item: {e}")
                continue
    except Exception as e:
        logger.warning(f"Could not read folder: {e}")
    return meetings


def _scan_for_calendars(folder, today, offset, seen, results, depth=0):
    """Recursively scan folders for calendars, max depth 4."""
    if depth > 4:
        return
    try:
        if folder.DefaultItemType == 1:
            name = folder.Name
            # Skip noise calendars
            skip = ["birthday", "holiday", "contacts", "birthdays"]
            if not any(s in name.lower() for s in skip):
                logger.info(f"  Scanning calendar: {name}")
                meetings = _get_meetings_from_folder(folder, today, offset)
                for m in meetings:
                    key = (m["subject"], m["start"])
                    if key not in seen:
                        seen.add(key)
                        results.append(m)
        for sub in folder.Folders:
            _scan_for_calendars(sub, today, offset, seen, results, depth + 1)
    except Exception:
        pass


def get_todays_meetings() -> List[dict]:
    outlook = _get_outlook()
    if not outlook:
        return []

    try:
        ns     = outlook.GetNamespace("MAPI")
        time.sleep(0.5)
        today  = datetime.datetime.now().date()
        offset = _utc_offset()
        logger.info(f"Local UTC offset: {offset}")

        all_meetings = []
        seen = set()

        for store in ns.Stores:
            try:
                logger.info(f"Checking store: {store.DisplayName}")
                root = store.GetRootFolder()
                # Scan direct children of root first
                for folder in root.Folders:
                    _scan_for_calendars(folder, today, offset, seen, all_meetings)
            except Exception as e:
                logger.warning(f"Could not read store: {e}")
                continue

        logger.info(f"Found {len(all_meetings)} meetings today")
        return sorted(all_meetings, key=lambda m: m["start"])

    except Exception as e:
        logger.error(f"Failed to read calendar: {e}")
        return []
    finally:
        pythoncom.CoUninitialize()


def send_summary_email(
    subject: str,
    meeting_title: str,
    meeting_date: str,
    summary_text: str,
    transcript_path: Optional[str] = None,
) -> bool:
    outlook = _get_outlook()
    if not outlook:
        return False

    try:
        ns   = outlook.GetNamespace("MAPI")
        mail = outlook.CreateItem(0)

        mail.Subject    = f"Meeting Notes: {meeting_title} ({meeting_date})"
        mail.BodyFormat = 2

        transcript_line = ""
        if transcript_path:
            transcript_line = (
                f"<p style='font-size:12px;color:#888;margin-top:16px;'>"
                f"Transcript saved to: {transcript_path}</p>"
            )

        html = f"""
<html><body style="font-family: Segoe UI, sans-serif; color: #1a1a1a; max-width: 680px;">
<div style="background:#003a57; padding:20px 24px; border-radius:8px; margin-bottom:20px;">
  <h2 style="color:#4fc3f7; margin:0; font-size:18px;">Meeting Recorder Summary</h2>
  <p style="color:#90caf9; margin:4px 0 0; font-size:13px;">
    {meeting_title} &mdash; {meeting_date}
  </p>
</div>
<div style="background:#f5f5f5; padding:20px 24px; border-radius:8px;
            white-space:pre-wrap; font-size:14px; line-height:1.7; color:#222;">
{summary_text}
</div>
{transcript_line}
<p style="font-size:11px; color:#aaa; margin-top:24px;
          border-top:1px solid #eee; padding-top:12px;">
  Sent automatically by Meeting Recorder
</p>
</body></html>
"""
        mail.HTMLBody = html
        mail.To       = ns.CurrentUser.Address
        mail.Send()

        logger.info(f"Summary email sent for: {meeting_title}")
        return True

    except Exception as e:
        logger.error(f"Failed to send email: {e}")
        return False
    finally:
        pythoncom.CoUninitialize()


def make_session_name(meeting: dict) -> str:
    date_str = meeting["start"].strftime("%Y-%m-%d")
    time_str = meeting["start"].strftime("%H%M")
    subject  = meeting["subject"]
    safe = "".join(c if c.isalnum() or c in " -_" else "" for c in subject)
    safe = safe.strip().replace("  ", " ")[:48]
    return f"{date_str} {time_str} {safe}"
