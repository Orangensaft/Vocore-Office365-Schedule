import datetime
import time
from datetime import timedelta
from typing import Optional
from zoneinfo import ZoneInfo
from environs import Env
import unidecode
from O365 import Account, FileSystemTokenBackend
from O365.calendar import Schedule, Event
from vocore_screen.screen import VocoreScreen

env = Env()
env.read_env()

"""
MS App needs the following scopes:
* offline_access
* User.Read
* Calendars.Read
* Calendars.Read.Shared
"""
MS_APP_CLIENT_ID = env.str("MS_APP_CLIENT_ID")
MS_CLIENT_SECRET_VALUE = env.str("MS_CLIENT_SECRET_VALUE")

DEFAULT_COLOR = "#d0f4de"
ACTIVE_COLOR = "#ff99c8"
NEXT_COLOR = "#fcf6bd"
ACCENT_COLOR = "#a9def9"


def get_current_time() -> datetime.datetime:
    return datetime.datetime.now(tz=ZoneInfo("Europe/Berlin"))


def login_o365() -> Account:
    token_backend = FileSystemTokenBackend(token_filename='my_token.txt')
    acc = Account((MS_APP_CLIENT_ID, MS_CLIENT_SECRET_VALUE), token_backend=token_backend)
    if acc.is_authenticated:
        print("Already logged in. Nice.")
        return acc
    if acc.authenticate(scopes=["basic", "calendar", "calendar_shared"]):
        return acc
    raise Exception("Could not auth :(")


def get_todays_meetings(acc: Account) -> [Event]:
    schedule: Schedule = acc.schedule()
    cal = schedule.get_default_calendar()
    today = get_current_time().date()
    tomorrow = today + timedelta(days=1)
    q = cal.new_query("start").greater_equal(today).chain("and").on_attribute("end").less(tomorrow)
    meetings = cal.get_events(query=q)
    return sorted(list(meetings), key=lambda i: i.start)


def get_next_meetings(meetings: [Event]) -> [Event]:
    out = []
    now = get_current_time()
    for m in meetings:
        if m.start >= now:
            out.append(m)
    return out


def get_next_meeting(meetings: [Event]) -> Optional[Event]:
    next_meetings = get_next_meetings(meetings)
    if len(next_meetings):
        return next_meetings[0]
    return None


def get_current_meeting(meetings: [Event]) -> Optional[Event]:
    for m in meetings:
        if is_currently_active(m):
            return m
    return None


def format_meeting(meeting: Event):
    subject = unidecode.unidecode(meeting.subject)
    timebox = f"{meeting.start.strftime('%H:%M')}-{meeting.end.strftime('%H:%M')}"
    if meeting.is_all_day:
        timebox="ALL DAY"
    out = f"[{timebox}] {subject}"
    return out


def is_currently_active(meeting: Event):
    if meeting.is_all_day:
        return False
    now = get_current_time()
    if meeting.start <= now <= meeting.end:
        return True
    return False


def draw_next_meeting_timer(meetings: [Event], screen: VocoreScreen):
    next_meeting = get_next_meeting(meetings)
    if next_meeting is not None:
        now = get_current_time()
        until = next_meeting.start - now
        total_secs = until.seconds
        hours = total_secs // 3600
        total_secs -= hours*3600
        minutes = total_secs // 60
        total_secs -= minutes*60
        seconds = total_secs
        if until.seconds <= 5*60:
            color = ACTIVE_COLOR
        elif until.seconds <= 30*60:
            color = NEXT_COLOR
        else:
            color = ACCENT_COLOR
        screen.draw_string(50, 400, f"Next meeting in: {hours}h, {minutes}m, {seconds}s", color, size=3)
    else:
        screen.draw_string(50, 400, "No more meetings!", "#00FF00", size=3)


def render(screen: VocoreScreen, acc: Account):
    screen.clear()
    meetings = get_todays_meetings(acc)
    screen.clear()
    cur_y = 480 - 100
    now = get_current_time()
    screen.draw_string(50, 430, f"Time: {now.strftime('%d.%m.%Y - %H:%M:%S')}", DEFAULT_COLOR, size=3)
    next_meeting = get_next_meeting(meetings)
    draw_next_meeting_timer(meetings, screen)

    for i, meeting in enumerate(meetings):
        try:
            color = ACTIVE_COLOR if is_currently_active(meeting) else DEFAULT_COLOR
            if meeting is next_meeting:
                color = NEXT_COLOR
            screen.draw_string(50, cur_y, format_meeting(meeting), color, size=2)
            cur_y -= 16
        except IndexError:
            print(f"End reached at {cur_y}")
            break  # Reached and of screen
    screen.blit()


if __name__ == '__main__':
    screen = VocoreScreen()
    screen.set_brightness(255)
    screen.draw_string(50, 480-50, "Logging in to office 365...", "#00FFFF", True)
    acc = login_o365()
    screen.draw_string(50, 480 - 50 - 16, "Success!", "#00FF00", True)
    while True:
        try:
            render(screen, acc)
            time.sleep(.5)
        except KeyboardInterrupt:
            print("Bye!")
            screen.clear(True)
            break
