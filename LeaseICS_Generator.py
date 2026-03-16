"""
LeaseICS_Generator.py
─────────────────────
Interactive script that collects critical lease dates and generates
an Outlook-compatible .ics calendar file.

Usage:
    python LeaseICS_Generator.py

Output:
    {TenantName}_LeaseDates.ics  (saved in the same folder as this script)

Leave any optional field blank and press Enter to skip it.
"""

import uuid
from datetime import date, timedelta


# ── Color / formatting helpers ─────────────────────────────────────────────────

RESET  = "\033[0m"
BOLD   = "\033[1m"
CYAN   = "\033[96m"
YELLOW = "\033[93m"
GREEN  = "\033[92m"
DIM    = "\033[2m"

def header(text):
    print(f"\n{BOLD}{CYAN}{'─' * 60}{RESET}")
    print(f"{BOLD}{CYAN}  {text}{RESET}")
    print(f"{BOLD}{CYAN}{'─' * 60}{RESET}")

def prompt(label, required=True, default=None):
    suffix = f" [{default}]" if default else (" (required)" if required else " (optional, press Enter to skip)")
    while True:
        val = input(f"  {BOLD}{label}{RESET}{DIM}{suffix}:{RESET} ").strip()
        if not val and default:
            return default
        if not val and required:
            print(f"  {YELLOW}⚠  This field is required.{RESET}")
            continue
        return val or None


# ── Date helpers ───────────────────────────────────────────────────────────────

def parse_date(s: str) -> date:
    """Accept YYYY-MM-DD or MM/DD/YYYY."""
    s = s.strip()
    for fmt in ("%Y-%m-%d", "%m/%d/%Y"):
        try:
            return date(*[int(x) for x in __import__("re").split(r"[-/]", s)][::-1]
                        ) if False else date.fromisoformat(
                            __import__("re").sub(r"(\d{2})/(\d{2})/(\d{4})", r"\3-\1-\2", s)
                        )
        except Exception:
            continue
    raise ValueError(f"Cannot parse date: '{s}'. Use YYYY-MM-DD or MM/DD/YYYY.")

def prompt_date(label, required=True):
    while True:
        val = prompt(label, required=required)
        if val is None:
            return None
        try:
            return parse_date(val)
        except ValueError as e:
            print(f"  {YELLOW}⚠  {e}{RESET}")

def months_before(d: date, months: int) -> date:
    """Return a date approximately N months before d."""
    m = d.month - months
    y = d.year + m // 12
    m = m % 12
    if m == 0:
        m = 12
        y -= 1
    try:
        return date(y, m, d.day)
    except ValueError:
        # Handle end-of-month edge cases
        import calendar
        last_day = calendar.monthrange(y, m)[1]
        return date(y, m, last_day)


# ── iCalendar helpers ──────────────────────────────────────────────────────────

def uid():
    return str(uuid.uuid4()).upper()

def dt(d: date) -> str:
    return d.strftime("%Y%m%d")

def fold(text: str) -> str:
    """RFC 5545 line folding at 75 octets."""
    lines = []
    while len(text) > 75:
        lines.append(text[:75])
        text = " " + text[75:]
    lines.append(text)
    return "\r\n".join(lines)

def make_event(summary: str, start: date, description: str, alarms_days: list = []) -> str:
    lines = [
        "BEGIN:VEVENT",
        f"UID:{uid()}",
        f"DTSTART;VALUE=DATE:{dt(start)}",
        f"DTEND;VALUE=DATE:{dt(date.fromordinal(start.toordinal() + 1))}",
        fold(f"SUMMARY:{summary}"),
        fold(f"DESCRIPTION:{description.replace(chr(10), chr(92) + 'n')}"),
    ]
    for days in alarms_days:
        lines += [
            "BEGIN:VALARM",
            "ACTION:DISPLAY",
            fold(f"DESCRIPTION:{summary}"),
            f"TRIGGER:-P{days}D",
            "END:VALARM",
        ]
    lines.append("END:VEVENT")
    return "\r\n".join(lines)


# ── Input collection ───────────────────────────────────────────────────────────

def collect_inputs():
    print(f"\n{BOLD}{'=' * 60}")
    print("  Lease Critical Dates — .ics Generator")
    print(f"{'=' * 60}{RESET}")
    print(f"{DIM}  Generates Outlook calendar events from lease critical dates.")
    print(f"  Use YYYY-MM-DD or MM/DD/YYYY for all dates.{RESET}\n")

    data = {}

    # ── Basics ──
    header("1. Lease Basics")
    data["tenant"]   = prompt("Tenant name")
    data["property"] = prompt("Property name / address")
    data["prefix"]   = prompt("Short property prefix for calendar (e.g. 'Centre Plaza')", required=False)
    data["suite"]    = prompt("Suite / unit", required=False)
    data["rsf"]      = prompt("Square footage (RSF)", required=False)

    # ── Term ──
    header("2. Lease Term")
    data["commencement"] = prompt_date("Commencement date")
    data["expiration"]   = prompt_date("Expiration date")

    # ── Renewal Option ──
    header("3. Renewal Option  (press Enter to skip if none)")
    has_renewal = prompt("Does the lease have a renewal option? (y/n)", required=False, default="n")
    data["has_renewal"] = has_renewal and has_renewal.lower().startswith("y")
    if data["has_renewal"]:
        data["renewal_notice_months"] = int(prompt("Notice required (months before expiration)", default="6"))
        data["renewal_options"]       = prompt("How many options, and how long each? (e.g. 2 x 5-year)", required=False)
        data["renewal_rate_note"]     = prompt("Renewal rate description (e.g. Prevailing Rental Rate)", required=False, default="To be negotiated")

    # ── ROFR ──
    header("4. Right of First Refusal  (press Enter to skip if none)")
    has_rofr = prompt("Does the lease have a ROFR? (y/n)", required=False, default="n")
    data["has_rofr"] = has_rofr and has_rofr.lower().startswith("y")
    if data["has_rofr"]:
        data["rofr_space"]          = prompt("ROFR space description (e.g. up to 7,500 RSF on Floors 4 & 6)", required=False)
        data["rofr_exercise_days"]  = prompt("Exercise window (business days after Landlord notice)", default="10")
        data["rofr_min_remaining"]  = int(prompt("Minimum months remaining required to exercise ROFR", default="24"))

    # ── Additional dates ──
    header("5. Other Important Dates  (all optional)")
    data["rent_commencement"] = prompt_date("Rent commencement date (if different from commencement)", required=False)
    data["abatement_end"]     = prompt_date("Rent abatement end date", required=False)
    data["option_to_purchase"]= prompt_date("Purchase option deadline", required=False)
    data["custom_date_label"] = prompt("Custom date label (e.g. 'Co-Tenancy Notice Deadline')", required=False)
    data["custom_date"]       = prompt_date("Custom date", required=False) if data.get("custom_date_label") else None

    return data


# ── Event builder ──────────────────────────────────────────────────────────────

def build_events(d):
    events   = []
    tenant   = d["tenant"]
    prop     = d["property"]
    prefix   = f"{d['prefix']} | " if d.get("prefix") else ""
    suite    = f", {d['suite']}" if d.get("suite") else ""
    rsf_info = f", {d['rsf']} RSF" if d.get("rsf") else ""
    location = f"{prop}{suite}{rsf_info}"

    # 1. Commencement
    events.append(make_event(
        summary     = f"{prefix}{tenant} — Lease Commencement",
        start       = d["commencement"],
        description = (
            f"Tenant: {tenant}\n"
            f"Premises: {location}\n"
            f"Lease commences: {d['commencement'].strftime('%B %d, %Y')}\n"
            f"Lease expires: {d['expiration'].strftime('%B %d, %Y')}"
        ),
        alarms_days = [7, 1],
    ))

    # 2. Rent commencement (if different)
    if d.get("rent_commencement") and d["rent_commencement"] != d["commencement"]:
        events.append(make_event(
            summary     = f"{prefix}{tenant} — Rent Commencement",
            start       = d["rent_commencement"],
            description = (
                f"First day rent is due under the lease.\n"
                f"Tenant: {tenant}\n"
                f"Premises: {location}"
            ),
            alarms_days = [14, 1],
        ))

    # 3. Abatement end
    if d.get("abatement_end"):
        events.append(make_event(
            summary     = f"{prefix}{tenant} — Rent Abatement Ends",
            start       = d["abatement_end"],
            description = (
                f"Rent abatement period ends. Full rent due starting the following month.\n"
                f"Tenant: {tenant}\n"
                f"Premises: {location}"
            ),
            alarms_days = [30, 7],
        ))

    # 4. ROFR 24-month warning
    if d.get("has_rofr"):
        min_months = d["rofr_min_remaining"]
        rofr_warning = months_before(d["expiration"], min_months)
        space   = d.get("rofr_space") or "as specified in lease"
        ex_days = d.get("rofr_exercise_days", "10")
        events.append(make_event(
            summary     = f"{prefix}{tenant} — ROFR {min_months}-Month Warning",
            start       = rofr_warning,
            description = (
                f"After this date fewer than {min_months} months remain on the lease.\n"
                f"The Right of First Refusal can no longer be exercised.\n\n"
                f"ROFR space: {space}\n"
                f"Exercise window: {ex_days} business days after Landlord notice\n"
                f"Lease expiration: {d['expiration'].strftime('%B %d, %Y')}"
            ),
            alarms_days = [30, 7],
        ))

    # 5. Begin renewal discussions
    if d.get("has_renewal"):
        notice_months = d["renewal_notice_months"]
        deadline      = months_before(d["expiration"], notice_months)
        discuss_start = months_before(d["expiration"], notice_months + 6)
        options_note  = d.get("renewal_options") or "See lease"
        rate_note     = d.get("renewal_rate_note") or "To be negotiated"
        events.append(make_event(
            summary     = f"{prefix}{tenant} — Begin Renewal Discussions",
            start       = discuss_start,
            description = (
                f"Recommended start of renewal negotiations.\n"
                f"Renewal notice deadline: {deadline.strftime('%B %d, %Y')} "
                f"({notice_months} months before expiration)\n\n"
                f"Options: {options_note}\n"
                f"Rate: {rate_note}\n\n"
                f"ACTION: Begin market analysis and draft renewal terms."
            ),
            alarms_days = [14],
        ))

        # 6. Renewal deadline
        events.append(make_event(
            summary     = f"{prefix}{tenant} — Renewal Option Deadline",
            start       = deadline,
            description = (
                f"LAST DAY to deliver written notice exercising the renewal option.\n"
                f"Lease requires notice at least {notice_months} months before "
                f"the {d['expiration'].strftime('%B %d, %Y')} expiration.\n\n"
                f"Options: {options_note}\n"
                f"Rate: {rate_note}\n\n"
                f"If notice is not delivered today, all renewal options are forfeited."
            ),
            alarms_days = [180, 90, 30, 7],
        ))

    # 7. Purchase option deadline
    if d.get("option_to_purchase"):
        events.append(make_event(
            summary     = f"{prefix}{tenant} — Purchase Option Deadline",
            start       = d["option_to_purchase"],
            description = (
                f"Last day to exercise the purchase option under the lease.\n"
                f"Tenant: {tenant}\n"
                f"Premises: {location}"
            ),
            alarms_days = [180, 90, 30, 7],
        ))

    # 8. Custom date
    if d.get("custom_date") and d.get("custom_date_label"):
        events.append(make_event(
            summary     = f"{prefix}{tenant} — {d['custom_date_label']}",
            start       = d["custom_date"],
            description = (
                f"{d['custom_date_label']}\n"
                f"Tenant: {tenant}\n"
                f"Premises: {location}"
            ),
            alarms_days = [30, 7],
        ))

    # 9. Annual lease anniversaries (30-day advance reminder each year)
    start     = d["commencement"]
    expiry    = d["expiration"]
    year_num  = 1
    ann_date  = date(start.year + 1, start.month, start.day) if start.month != 2 or start.day != 29 \
                else date(start.year + 1, 3, 1)
    while ann_date < expiry:
        # Calculate years remaining at this anniversary
        years_left = round((expiry - ann_date).days / 365.25, 1)
        events.append(make_event(
            summary     = f"{prefix}{tenant} — Lease Anniversary (Year {year_num})",
            start       = ann_date,
            description = (
                f"Year {year_num} lease anniversary — reach out to the tenant.\n\n"
                f"Tenant: {tenant}\n"
                f"Premises: {location}\n"
                f"Lease commenced: {start.strftime('%B %d, %Y')}\n"
                f"Lease expires: {expiry.strftime('%B %d, %Y')} "
                f"(approx. {years_left} years remaining)\n\n"
                f"ACTION: Property management team to contact tenant — check satisfaction, "
                f"address any outstanding issues, and strengthen the relationship."
            ),
            alarms_days = [30],
        ))
        year_num += 1
        try:
            ann_date = date(start.year + year_num, start.month, start.day)
        except ValueError:
            ann_date = date(start.year + year_num, start.month + 1, 1)

    # 10. Lease expiration — always last
    events.append(make_event(
        summary     = f"{prefix}{tenant} — Lease Expiration",
        start       = d["expiration"],
        description = (
            f"Final day of the lease term.\n\n"
            f"Tenant: {tenant}\n"
            f"Premises: {location}\n\n"
            f"Confirm vacation, key return, and space condition walkthrough.\n"
            f"Review holdover provisions in the lease."
        ),
        alarms_days = [365, 180, 90, 30],
    ))

    return events


# ── Main ───────────────────────────────────────────────────────────────────────

def main():
    try:
        data   = collect_inputs()
        events = build_events(data)

        # Build filename from tenant name (strip special chars)
        safe_name = "".join(c for c in data["tenant"] if c.isalnum() or c in " _-").strip().replace(" ", "_")
        output_file = f"{safe_name}_LeaseDates.ics"

        calendar = "\r\n".join([
            "BEGIN:VCALENDAR",
            "VERSION:2.0",
            f"PRODID:-//Spire Realty Group//Lease Critical Dates//EN",
            "CALSCALE:GREGORIAN",
            "METHOD:PUBLISH",
            f"X-WR-CALNAME:{data['tenant']} — Lease Critical Dates",
            *events,
            "END:VCALENDAR",
        ])

        with open(output_file, "w", encoding="utf-8") as f:
            f.write(calendar)

        print(f"\n{GREEN}{BOLD}✅  Done!{RESET}")
        print(f"   File:   {BOLD}{output_file}{RESET}")
        print(f"   Events: {len(events)}")
        print(f"\n{DIM}   Double-click the .ics file to import all events into Outlook.{RESET}\n")

    except KeyboardInterrupt:
        print(f"\n\n{YELLOW}Cancelled.{RESET}\n")


if __name__ == "__main__":
    main()
