import os
import sqlite3
from datetime import date

import resend
from dotenv import load_dotenv

# --------------------------------------------------------------------------
# Load environment variables from a local .env file (never commit this file
# or share its contents — it holds your real API key).
# --------------------------------------------------------------------------
load_dotenv()

resend.api_key = os.environ.get("RESEND_API_KEY")

if not resend.api_key:
    raise RuntimeError(
        "RESEND_API_KEY is not set. Create a .env file in this folder with:\n"
        "RESEND_API_KEY=your_actual_key_here"
    )


def run_due_date_check():
    """Find everyone whose book is due today OR overdue, and email them."""
    today = date.today()

    connection = sqlite3.connect("library.db")
    cursor = connection.cursor()

    # due_date <= today catches both "due today" and "overdue" books.
    cursor.execute(
        """
        SELECT id, employee_email, employee_name, book_title, due_date
        FROM issued_books
        WHERE status = 'Issued'
        AND due_date <= ?
        AND reminder_sent = 0
        """,
        (today.isoformat(),),
    )

    records = cursor.fetchall()

    if not records:
        print("No due or overdue books — nothing to send.")
        connection.close()
        return

    for record_id, email, name, book, due_date_text in records:
        due_date = date.fromisoformat(due_date_text)
        days_overdue = (today - due_date).days

        if days_overdue == 0:
            subject = f"Reminder: '{book}' is Due Today!"
            status_line = "is due <strong>today</strong>"
        else:
            subject = f"Overdue Notice: '{book}' is {days_overdue} day(s) late"
            status_line = (
                f"was due on <strong>{due_date_text}</strong> and is now "
                f"<strong>{days_overdue} day(s) overdue</strong>"
            )

        try:
            params = {
                "from": "Library System <onboarding@resend.dev>",  # swap for your verified domain later
                "to": [email],
                "subject": subject,
                "html": f"""
                    <p>Hi {name},</p>
                    <p>This is a reminder that your borrowed book,
                    <strong>{book}</strong>, {status_line}.</p>
                    <p>Please return it to the library as soon as possible to avoid any late fees.</p>
                    <br>
                    <p>Best regards,<br>Library Management Team</p>
                """,
            }

            resend.Emails.send(params)

            # Mark as sent so it's never emailed twice.
            cursor.execute(
                "UPDATE issued_books SET reminder_sent = 1 WHERE id = ?",
                (record_id,),
            )
            connection.commit()
            print(f"Sent {'due-today' if days_overdue == 0 else 'overdue'} reminder to {email} for '{book}'")

        except Exception as exc:
            print(f"Failed to send to {email}: {exc}")

    connection.close()


if __name__ == "__main__":
    run_due_date_check()