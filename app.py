import streamlit as st
import sqlite3
from datetime import date, timedelta
import pandas as pd

from database import (
    create_tables,
    get_connection,
    import_excel_to_database,
    export_database_to_excel
)


# =========================================================
# PAGE CONFIG
# =========================================================

st.set_page_config(
    page_title="Company Library Management",
    page_icon="📚",
    layout="wide",
    initial_sidebar_state="expanded"
)


# =========================================================
# PROFESSIONAL CSS
# =========================================================

st.markdown("""
<style>

/* ---------- FONT ---------- */

html, body, [class*="css"] {
    font-family: 'Inter', -apple-system, BlinkMacSystemFont,
                 'Segoe UI', Roboto, Helvetica, Arial, sans-serif;
}


/* ---------- MAIN PAGE ---------- */

.stApp {
    background: #f6f8fb;
}

.block-container {
    padding-top: 3rem;
    padding-bottom: 2rem;
    max-width: 1400px;
}

/* Streamlit's fixed top toolbar (Deploy button etc.) sits above the
   content; make sure it doesn't visually collide with our page title. */
header[data-testid="stHeader"] {
    background: transparent;
}


/* ---------- SIDEBAR ---------- */

section[data-testid="stSidebar"] {
    background: #111827;
}

section[data-testid="stSidebar"] * {
    color: #f9fafb;
}

.sidebar-brand {
    padding: 10px 5px 20px 5px;
}

.sidebar-brand-title {
    font-size: 22px;
    font-weight: 700;
}

.sidebar-brand-subtitle {
    font-size: 12px;
    color: #9ca3af !important;
    margin-top: 3px;
}


/* ---------- SIDEBAR NAVIGATION (button-based) ---------- */

/* We use st.button (not st.radio) for nav items so we have full
   control over the look — no native radio circle to fight with. */

section[data-testid="stSidebar"] div[data-testid="stVerticalBlock"] div[data-testid="stButton"] {
    margin-bottom: 2px;
}

section[data-testid="stSidebar"] .stButton > button {
    width: 100%;
    text-align: left;
    justify-content: flex-start;
    background: transparent;
    border: none;
    color: #d1d5db !important;
    font-weight: 500;
    padding: 10px 14px;
    border-radius: 8px;
    min-height: 0;
    box-shadow: none;
}

section[data-testid="stSidebar"] .stButton > button:hover {
    background: #1f2937;
    color: #ffffff !important;
    border: none;
}

section[data-testid="stSidebar"] .stButton > button:focus:not(:active) {
    background: #1f2937;
    color: #ffffff !important;
    box-shadow: none;
}

/* Active page — Streamlit renders the primary button with a literal
   kind="primary" attribute on the <button> element itself. */
section[data-testid="stSidebar"] .stButton > button[kind="primary"] {
    background: #2563eb !important;
    color: #ffffff !important;
    font-weight: 600;
    border: none;
}

section[data-testid="stSidebar"] .stButton > button[kind="primary"]:hover,
section[data-testid="stSidebar"] .stButton > button[kind="primary"]:focus {
    background: #2563eb !important;
    color: #ffffff !important;
}


/* ---------- KPI COLOR ACCENTS ---------- */
/* Numbers stay black for readability; only the left border
   carries the semantic color. */

.kpi-overdue {
    border-left: 4px solid #dc2626;
}

.kpi-due-today {
    border-left: 4px solid #d97706;
}

.kpi-available {
    border-left: 4px solid #059669;
}

.kpi-issued {
    border-left: 4px solid #2563eb;
}


/* ---------- PAGE HEADER ---------- */

.page-title {
    font-size: 30px;
    font-weight: 700;
    color: #111827;
    margin-bottom: 3px;
}

.page-subtitle {
    color: #6b7280;
    font-size: 14px;
    margin-bottom: 25px;
}


/* ---------- CARDS ---------- */

.card {
    background: white;
    border: 1px solid #e5e7eb;
    border-radius: 14px;
    padding: 22px;
    margin-bottom: 20px;
    box-shadow: 0 2px 8px rgba(0,0,0,0.03);
}

.card-title {
    font-size: 18px;
    font-weight: 650;
    color: #111827;
    margin-bottom: 4px;
}

.card-subtitle {
    color: #6b7280;
    font-size: 13px;
    margin-bottom: 18px;
}


/* ---------- KPI CARDS ---------- */

.kpi {
    background: white;
    border: 1px solid #e5e7eb;
    border-radius: 14px;
    padding: 20px;
    min-height: 125px;
    box-shadow: 0 2px 8px rgba(0,0,0,0.03);
    transition: box-shadow 0.15s ease, transform 0.15s ease;
}

.kpi:hover {
    box-shadow: 0 6px 16px rgba(0,0,0,0.07);
    transform: translateY(-2px);
}

.kpi-label {
    color: #6b7280;
    font-size: 12px;
    font-weight: 600;
    letter-spacing: 0.04em;
    text-transform: uppercase;
}

.kpi-value {
    color: #111827;
    font-size: 30px;
    font-weight: 700;
    margin-top: 8px;
}

.kpi-small {
    color: #6b7280;
    font-size: 12px;
    margin-top: 3px;
}


/* ---------- STATUS ---------- */

.status-available {
    color: #047857;
    font-weight: 600;
}

.status-issued {
    color: #dc2626;
    font-weight: 600;
}


/* ---------- BUTTONS ---------- */

.stButton > button {
    border-radius: 8px;
    font-weight: 600;
    min-height: 42px;
}


/* ---------- INPUTS ---------- */

/* Hide the "Press Enter to submit form" hint Streamlit shows under
   every text input inside a form - it's just visual clutter here
   since the Save/submit button is always right there. */
[data-testid="InputInstructions"] {
    display: none;
}

.stTextInput input,
.stSelectbox div[data-baseweb="select"],
.stDateInput input {
    border-radius: 8px;
}


/* ---------- TABLE ---------- */

[data-testid="stDataFrame"] {
    border-radius: 10px;
}


/* ---------- ALERT ---------- */

div[data-testid="stAlert"] {
    border-radius: 10px;
}


/* ---------- HORIZONTAL LINE ---------- */

hr {
    border: none;
    border-top: 1px solid #e5e7eb;
    margin: 25px 0;
}

</style>
""", unsafe_allow_html=True)


# =========================================================
# DATABASE INITIALIZATION
# =========================================================

create_tables()

# IMPORTANT: only import from Excel once per session (first load), not
# on every rerun. Streamlit reruns this whole script on every click, and
# re-importing every time meant any in-app change (issuing a book,
# deleting a row, etc.) could get silently overwritten back to whatever
# was last successfully written to Excel — which is exactly what caused
# statuses/deletions to "revert" after further clicks.
if "excel_imported" not in st.session_state:

    try:
        import_excel_to_database()

        # The import step also self-heals book availability (derived
        # from IssuedBooks, not the Excel Available column). Push that
        # corrected state back out to Excel immediately, otherwise the
        # fix only ever lives in the database and the spreadsheet keeps
        # showing the old, stale Available/Issued values.
        export_database_to_excel()

    except Exception as e:
        st.warning(f"Excel synchronization warning: {e}")

    st.session_state.excel_imported = True


# Version counters used to force st.data_editor widgets to reload
# fresh data instead of showing a stale cached snapshot whenever the
# underlying books/employees data changes from elsewhere in the app
# (e.g. issuing/returning a book).
if "books_version" not in st.session_state:
    st.session_state.books_version = 0

if "employees_version" not in st.session_state:
    st.session_state.employees_version = 0


# =========================================================
# HELPER FUNCTIONS
# =========================================================

def get_all_members():

    connection = get_connection()
    cursor = connection.cursor()

    cursor.execute("""
        SELECT
            employee_id,
            name,
            email,
            department
        FROM employees
        ORDER BY name
    """)

    data = cursor.fetchall()
    connection.close()

    return data


def get_all_books():

    connection = get_connection()
    cursor = connection.cursor()

    cursor.execute("""
        SELECT
            book_id,
            title,
            author,
            category,
            available,
            condition
        FROM books
        ORDER BY title
    """)

    data = cursor.fetchall()
    connection.close()

    return data


def get_currently_issued_books():

    connection = get_connection()
    cursor = connection.cursor()

    cursor.execute("""
        SELECT
            id,
            employee_id,
            employee_name,
            employee_email,
            book_id,
            book_title,
            issue_date,
            due_date,
            return_date,
            status,
            reminder_sent
        FROM issued_books
        WHERE status = 'Issued'
        ORDER BY due_date
    """)

    data = cursor.fetchall()
    connection.close()

    return data


def get_waiting_list(book_id=None):

    connection = get_connection()
    cursor = connection.cursor()

    if book_id:

        cursor.execute("""
            SELECT id, book_id, book_title, employee_id, employee_name,
                   employee_email, request_date, notified
            FROM waiting_list
            WHERE book_id = ?
            ORDER BY request_date
        """, (book_id,))

    else:

        cursor.execute("""
            SELECT id, book_id, book_title, employee_id, employee_name,
                   employee_email, request_date, notified
            FROM waiting_list
            ORDER BY request_date
        """)

    data = cursor.fetchall()
    connection.close()

    return data


def add_to_waiting_list(
    book_id,
    book_title,
    employee_id,
    employee_name,
    employee_email
):

    connection = get_connection()
    cursor = connection.cursor()

    # Don't let the same employee queue twice for the same book.
    cursor.execute("""
        SELECT id FROM waiting_list
        WHERE book_id = ? AND employee_id = ?
    """, (book_id, employee_id))

    if cursor.fetchone():
        connection.close()
        raise ValueError(
            "This employee is already on the waiting list for this book."
        )

    cursor.execute("""
        INSERT INTO waiting_list
        (book_id, book_title, employee_id, employee_name, employee_email, request_date)
        VALUES (?, ?, ?, ?, ?, ?)
    """, (
        book_id,
        book_title,
        employee_id,
        employee_name,
        employee_email,
        date.today().isoformat()
    ))

    connection.commit()
    connection.close()


def remove_from_waiting_list(entry_id):

    connection = get_connection()
    cursor = connection.cursor()

    cursor.execute(
        "DELETE FROM waiting_list WHERE id = ?",
        (entry_id,)
    )

    connection.commit()
    connection.close()


def add_employee(
    employee_id,
    employee_name,
    employee_email,
    department
):

    connection = get_connection()
    cursor = connection.cursor()

    cursor.execute("""
        INSERT INTO employees
        (
            employee_id,
            name,
            email,
            department
        )
        VALUES (?, ?, ?, ?)
    """, (
        employee_id,
        employee_name,
        employee_email,
        department
    ))

    connection.commit()
    connection.close()

    st.session_state.employees_version += 1

    try:
        excel_path = export_database_to_excel()
        st.toast(f"Synced to Excel: {excel_path}", icon="✅")
    except Exception as e:
        st.warning(f"Employee saved, but couldn't sync to Excel: {e}")


def add_book(
    book_id,
    title,
    author,
    category,
    available=1
):

    connection = get_connection()
    cursor = connection.cursor()

    cursor.execute("""
        INSERT INTO books
        (
            book_id,
            title,
            author,
            category,
            available
        )
        VALUES (?, ?, ?, ?, ?)
    """, (
        book_id,
        title,
        author,
        category,
        available
    ))

    connection.commit()
    connection.close()

    st.session_state.books_version += 1

    try:
        excel_path = export_database_to_excel()
        st.toast(f"Synced to Excel: {excel_path}", icon="✅")
    except Exception as e:
        st.warning(f"Book saved, but couldn't sync to Excel: {e}")


def update_employee(
    original_employee_id,
    employee_id,
    employee_name,
    employee_email,
    department
):

    connection = get_connection()
    cursor = connection.cursor()

    cursor.execute("""
        UPDATE employees
        SET
            employee_id = ?,
            name = ?,
            email = ?,
            department = ?
        WHERE employee_id = ?
    """, (
        employee_id,
        employee_name,
        employee_email,
        department,
        original_employee_id
    ))

    connection.commit()
    connection.close()

    st.session_state.employees_version += 1

    try:
        excel_path = export_database_to_excel()
        st.toast(f"Synced to Excel: {excel_path}", icon="✅")
    except Exception as e:
        st.warning(f"Employee updated, but couldn't sync to Excel: {e}")


def update_book(
    original_book_id,
    book_id,
    title,
    author,
    category
):

    connection = get_connection()
    cursor = connection.cursor()

    cursor.execute("""
        UPDATE books
        SET
            book_id = ?,
            title = ?,
            author = ?,
            category = ?
        WHERE book_id = ?
    """, (
        book_id,
        title,
        author,
        category,
        original_book_id
    ))

    connection.commit()
    connection.close()

    st.session_state.books_version += 1

    try:
        excel_path = export_database_to_excel()
        st.toast(f"Synced to Excel: {excel_path}", icon="✅")
    except Exception as e:
        st.warning(f"Book updated, but couldn't sync to Excel: {e}")


def delete_employee(employee_id):
    """
    Delete an employee. Blocked if they currently have any book
    issued to them (status = 'Issued') so we never lose track of
    an outstanding loan.
    """

    connection = get_connection()
    cursor = connection.cursor()

    cursor.execute(
        "SELECT COUNT(*) FROM issued_books "
        "WHERE employee_id = ? AND status = 'Issued'",
        (employee_id,)
    )

    outstanding = cursor.fetchone()[0]

    if outstanding > 0:
        connection.close()
        raise ValueError(
            f"Cannot delete: this employee still has {outstanding} "
            "book(s) issued. Return them first."
        )

    cursor.execute(
        "DELETE FROM employees WHERE employee_id = ?",
        (employee_id,)
    )

    connection.commit()
    connection.close()

    st.session_state.employees_version += 1

    try:
        excel_path = export_database_to_excel()
        st.toast(f"Synced to Excel: {excel_path}", icon="✅")
    except Exception as e:
        st.warning(f"Employee deleted, but couldn't sync to Excel: {e}")


def delete_book(book_id):
    """
    Delete a book. Blocked if it's currently issued (available = 0)
    so we never delete a book that's still out with an employee.
    """

    connection = get_connection()
    cursor = connection.cursor()

    cursor.execute(
        "SELECT available FROM books WHERE book_id = ?",
        (book_id,)
    )

    row = cursor.fetchone()

    if row is not None and row[0] == 0:
        connection.close()
        raise ValueError(
            "Cannot delete: this book is currently issued. "
            "Return it first."
        )

    cursor.execute(
        "DELETE FROM books WHERE book_id = ?",
        (book_id,)
    )

    connection.commit()
    connection.close()

    st.session_state.books_version += 1

    try:
        excel_path = export_database_to_excel()
        st.toast(f"Synced to Excel: {excel_path}", icon="✅")
    except Exception as e:
        st.warning(f"Book deleted, but couldn't sync to Excel: {e}")


# =========================================================
# PAGE HEADER FUNCTION
# =========================================================

def page_header(title, subtitle):

    st.markdown(
        f"""
        <div class="page-title">{title}</div>
        <div class="page-subtitle">{subtitle}</div>
        """,
        unsafe_allow_html=True
    )


# =========================================================
# SIDEBAR
# =========================================================

with st.sidebar:

    st.markdown(
        """
        <div class="sidebar-brand">
            <div class="sidebar-brand-title">
                📚 Company Library
            </div>
            <div class="sidebar-brand-subtitle">
                Library Management System
            </div>
        </div>
        """,
        unsafe_allow_html=True
    )

    st.divider()

    NAV_ITEMS = [
        "🏠 Dashboard",
        "👥 Members",
        "📚 Books",
        "📖 Issue Book",
        "↩️ Return Book"
    ]

    if "current_page" not in st.session_state:
        st.session_state.current_page = NAV_ITEMS[0]

    for nav_label in NAV_ITEMS:

        is_active = st.session_state.current_page == nav_label

        if st.button(
            nav_label,
            key=f"nav_{nav_label}",
            use_container_width=True,
            type="primary" if is_active else "secondary"
        ):
            st.session_state.current_page = nav_label
            st.rerun()

    page = st.session_state.current_page

    st.divider()

  

    st.caption("Library Management System")
   


# =========================================================
# DASHBOARD
# =========================================================

if page == "🏠 Dashboard":

    page_header(
        "Library Dashboard",
        "Overview of books, employees and current library activity"
    )

    members = get_all_members()
    books = get_all_books()
    issued_books = get_currently_issued_books()

    total_members = len(members)
    total_books = len(books)

    available_books = sum(
        1 for book in books
        if book[4] == 1
    )

    issued_count = len(issued_books)

    today = date.today()

    due_today_count = 0
    overdue_count = 0

    for book in issued_books:

        due_date = date.fromisoformat(
            str(book[7])[:10]
        )

        if due_date == today:

            due_today_count += 1

        elif due_date < today:

            overdue_count += 1


    # -----------------------------------------------------
    # KPI CARDS
    # -----------------------------------------------------

    col1, col2, col3, col4 = st.columns(4)

    with col1:

        st.markdown(
            f"""
            <div class="kpi">
                <div class="kpi-label">TOTAL BOOKS</div>
                <div class="kpi-value">{total_books}</div>
                <div class="kpi-small">Books in library</div>
            </div>
            """,
            unsafe_allow_html=True
        )

    with col2:

        st.markdown(
            f"""
            <div class="kpi kpi-available">
                <div class="kpi-label">AVAILABLE</div>
                <div class="kpi-value">{available_books}</div>
                <div class="kpi-small">Ready to issue</div>
            </div>
            """,
            unsafe_allow_html=True
        )

    with col3:

        st.markdown(
            f"""
            <div class="kpi kpi-issued">
                <div class="kpi-label">ISSUED</div>
                <div class="kpi-value">{issued_count}</div>
                <div class="kpi-small">Currently issued</div>
            </div>
            """,
            unsafe_allow_html=True
        )

    with col4:

        st.markdown(
            f"""
            <div class="kpi">
                <div class="kpi-label">EMPLOYEES</div>
                <div class="kpi-value">{total_members}</div>
                <div class="kpi-small">Registered members</div>
            </div>
            """,
            unsafe_allow_html=True
        )


    st.write("")


    # -----------------------------------------------------
    # DUE / OVERDUE
    # -----------------------------------------------------

    col1, col2 = st.columns(2)

    with col1:

        st.markdown(
            f"""
            <div class="kpi kpi-due-today">
                <div class="kpi-label">DUE TODAY</div>
                <div class="kpi-value">{due_today_count}</div>
                <div class="kpi-small">
                    Books that must be returned today
                </div>
            </div>
            """,
            unsafe_allow_html=True
        )

    with col2:

        st.markdown(
            f"""
            <div class="kpi kpi-overdue">
                <div class="kpi-label">OVERDUE</div>
                <div class="kpi-value">{overdue_count}</div>
                <div class="kpi-small">
                    Books past their due date
                </div>
            </div>
            """,
            unsafe_allow_html=True
        )


    st.write("")


    # -----------------------------------------------------
    # SEARCH
    # -----------------------------------------------------

    st.markdown(
        '<div class="card">',
        unsafe_allow_html=True
    )

    st.markdown(
        '<div class="card-title">🔍 Search Issued Books</div>',
        unsafe_allow_html=True
    )

    st.markdown(
        '<div class="card-subtitle">Search currently issued books by book or by employee</div>',
        unsafe_allow_html=True
    )

    search_col1, search_col2 = st.columns([1, 2])

    with search_col1:

        search_mode = st.selectbox(
            "Search by",
            ["Book", "Employee"],
            label_visibility="collapsed"
        )

    with search_col2:

        search_query = st.text_input(
            "Search",
            placeholder=(
                "Search by book title or ID..."
                if search_mode == "Book"
                else "Search by employee name or ID..."
            ),
            label_visibility="collapsed"
        )

    filtered_issued_books = issued_books

    if search_query:

        query_lower = search_query.lower()

        if search_mode == "Book":

            filtered_issued_books = [
                book for book in issued_books
                if query_lower in str(book[4]).lower()
                or query_lower in str(book[5]).lower()
            ]

        else:

            filtered_issued_books = [
                book for book in issued_books
                if query_lower in str(book[1]).lower()
                or query_lower in str(book[2]).lower()
            ]

    st.markdown("</div>", unsafe_allow_html=True)


    st.write("")


    # -----------------------------------------------------
    # DUE TODAY / OVERDUE / UPCOMING (merged from Due Books)
    # -----------------------------------------------------

    due_today_list = []
    overdue_list = []
    upcoming_list = []

    for book in filtered_issued_books:

        due_date_value = date.fromisoformat(
            str(book[7])[:10]
        )

        if due_date_value == today:

            due_today_list.append(book)

        elif due_date_value < today:

            overdue_list.append(book)

        else:

            upcoming_list.append(book)

    def render_issued_table(book_list):

        if not book_list:
            return False

        data = []

        for book in book_list:

            data.append([
                book[1],
                book[2],
                book[3],
                book[4],
                book[5],
                book[7]
            ])

        df = pd.DataFrame(
            data,
            columns=[
                "Employee ID",
                "Employee Name",
                "Email",
                "Book ID",
                "Book Title",
                "Due Date"
            ]
        )

        st.dataframe(
            df,
            use_container_width=True,
            hide_index=True
        )

        return True

    st.markdown("### 🔔 Due Today")

    if not render_issued_table(due_today_list):
        st.success("No books are due today.")

    st.markdown("### ⚠️ Overdue")

    if not render_issued_table(overdue_list):
        st.success("No overdue books.")

    st.markdown("### 📆 Upcoming")

    if not render_issued_table(upcoming_list):
        st.info("No upcoming books.")


# =========================================================
# MEMBERS
# =========================================================

elif page == "👥 Members":

    page_header(
        "Members",
        "Manage employees registered with the company library"
    )

    # -----------------------------------------------------
    # ADD EMPLOYEE
    # -----------------------------------------------------

    st.markdown(
        '<div class="card">',
        unsafe_allow_html=True
    )

    st.markdown(
        '<div class="card-title">➕ Add Employee</div>',
        unsafe_allow_html=True
    )

    st.markdown(
        '<div class="card-subtitle">Enter employee information</div>',
        unsafe_allow_html=True
    )

    with st.form("add_employee_form"):

        col1, col2 = st.columns(2)

        with col1:

            employee_id_input = st.text_input(
                "Employee ID"
            )

            employee_name_input = st.text_input(
                "Employee Name"
            )

        with col2:

            employee_email_input = st.text_input(
                "Email"
            )

            department_input = st.text_input(
                "Department"
            )

        submit = st.form_submit_button(
            "💾 Save",
            use_container_width=True
        )

        if submit:

            if not employee_id_input.strip():

                st.error("Please enter Employee ID.")

            elif not employee_name_input.strip():

                st.error("Please enter Employee Name.")

            elif not employee_email_input.strip():

                st.error("Please enter Email.")

            else:

                try:

                    add_employee(
                        employee_id_input.strip(),
                        employee_name_input.strip(),
                        employee_email_input.strip(),
                        department_input.strip()
                    )

                    st.success(
                        "Employee added successfully."
                    )

                except sqlite3.IntegrityError:

                    st.error(
                        "Employee ID already exists."
                    )

                except Exception as e:

                    st.error(
                        f"Error: {e}"
                    )

    st.markdown("</div>", unsafe_allow_html=True)


    # -----------------------------------------------------
    # EMPLOYEE DIRECTORY (inline editable + deletable rows)
    # -----------------------------------------------------

    members = get_all_members()

    if members:

        search = st.text_input(
            "🔎 Search employee",
            placeholder="Search by ID, name or department..."
        )

        filtered_members = members

        if search:

            search_lower = search.lower()

            filtered_members = [
                member
                for member in members
                if search_lower in str(member[0]).lower()
                or search_lower in str(member[1]).lower()
                or search_lower in str(member[3]).lower()
            ]

        data = []

        for member in filtered_members:

            data.append([
                member[0],
                member[1],
                member[2],
                member[3]
            ])

        df = pd.DataFrame(
            data,
            columns=[
                "Employee ID",
                "Name",
                "Email",
                "Department"
            ]
        )

        st.caption(
            "Select a row's checkbox and press the 🗑️ icon (or Delete key) "
            "to remove it — then click Save Changes."
        )

        edited_df = st.data_editor(
            df,
            use_container_width=True,
            hide_index=True,
            num_rows="dynamic",
            key=f"employee_directory_editor_{st.session_state.employees_version}"
        )

        if st.button(
            "💾 Save Changes",
            use_container_width=True,
            key="save_employee_directory"
        ):

            errors = []

            original_ids = set(df["Employee ID"])
            edited_ids = set(edited_df["Employee ID"])

            # ---- deleted rows ----
            for _, original_row in df.iterrows():

                if original_row["Employee ID"] not in edited_ids:

                    try:

                        delete_employee(original_row["Employee ID"])

                    except ValueError as e:

                        errors.append(str(e))

                    except Exception as e:

                        errors.append(f"Error deleting row: {e}")

            # ---- new rows added via the editor's "+" row ----
            for _, edited_row in edited_df.iterrows():

                if edited_row["Employee ID"] in original_ids:
                    continue

                new_employee_id = str(edited_row["Employee ID"]).strip()
                new_name = str(edited_row["Name"]).strip()
                new_email = str(edited_row["Email"]).strip()
                new_department = str(edited_row["Department"]).strip()

                if not new_employee_id and not new_name and not new_email:
                    # blank placeholder row - ignore
                    continue

                if not new_employee_id or not new_name or not new_email:
                    errors.append(
                        "New row: Employee ID, Name and Email cannot be empty."
                    )
                    continue

                try:

                    add_employee(
                        new_employee_id,
                        new_name,
                        new_email,
                        new_department
                    )

                except sqlite3.IntegrityError:

                    errors.append(
                        f"Employee ID '{new_employee_id}' already exists."
                    )

                except Exception as e:

                    errors.append(f"Error adding row: {e}")

            # ---- edited existing rows ----
            for _, original_row in df.iterrows():

                if original_row["Employee ID"] not in edited_ids:
                    continue

                edited_row = edited_df[
                    edited_df["Employee ID"] == original_row["Employee ID"]
                ].iloc[0]

                if original_row.equals(edited_row):
                    continue

                new_employee_id = str(edited_row["Employee ID"]).strip()
                new_name = str(edited_row["Name"]).strip()
                new_email = str(edited_row["Email"]).strip()
                new_department = str(edited_row["Department"]).strip()

                if not new_employee_id or not new_name or not new_email:
                    errors.append(
                        f"Row for '{original_row['Name']}': "
                        "Employee ID, Name and Email cannot be empty."
                    )
                    continue

                try:

                    update_employee(
                        original_row["Employee ID"],
                        new_employee_id,
                        new_name,
                        new_email,
                        new_department
                    )

                except sqlite3.IntegrityError:

                    errors.append(
                        f"Employee ID '{new_employee_id}' is already "
                        "used by another employee."
                    )

                except Exception as e:

                    errors.append(f"Error updating row: {e}")

            if errors:

                for error_message in errors:
                    st.error(error_message)

            else:

                st.success("Changes saved successfully.")
                st.rerun()

    else:

        st.info(
            "No employees found."
        )


# =========================================================
# BOOKS
# =========================================================

elif page == "📚 Books":

    page_header(
        "Books",
        "Manage and monitor the company library catalogue"
    )

    # -----------------------------------------------------
    # ADD BOOK
    # -----------------------------------------------------

    st.markdown(
        '<div class="card">',
        unsafe_allow_html=True
    )

    st.markdown(
        '<div class="card-title">➕ Add New Book</div>',
        unsafe_allow_html=True
    )

    st.markdown(
        '<div class="card-subtitle">Add a new book to the library</div>',
        unsafe_allow_html=True
    )

    BOOK_CATEGORIES = [
        "Fiction",
        "Non-Fiction",
        "Self Help",
        "Programming",
        "Biography",
        "Fantasy",
        "Science",
        "History",
        "Business",
        "Other"
    ]

    with st.form("add_book_form"):

        col1, col2 = st.columns(2)

        with col1:

            book_id_input = st.text_input(
                "Book ID"
            )

            title_input = st.text_input(
                "Book Title"
            )

            author_input = st.text_input(
                "Author"
            )

        with col2:

            category_choice = st.selectbox(
                "Category",
                BOOK_CATEGORIES
            )

            category_other_input = st.text_input(
                "Category (if 'Other')",
                placeholder="Type a custom category"
            )

            status_choice = st.selectbox(
                "Status",
                ["Available", "Issued"]
            )

        submit = st.form_submit_button(
            "💾 Save",
            use_container_width=True
        )

        if submit:

            category_input = (
                category_other_input.strip()
                if category_choice == "Other"
                else category_choice
            )

            if not book_id_input.strip():

                st.error(
                    "Please enter Book ID."
                )

            elif not title_input.strip():

                st.error(
                    "Please enter Book Title."
                )

            elif category_choice == "Other" and not category_other_input.strip():

                st.error(
                    "Please type a custom category."
                )

            else:

                try:

                    add_book(
                        book_id_input.strip(),
                        title_input.strip(),
                        author_input.strip(),
                        category_input,
                        available=1 if status_choice == "Available" else 0
                    )

                    st.success(
                        "Book added successfully."
                    )

                except sqlite3.IntegrityError:

                    st.error(
                        "Book ID already exists."
                    )

                except Exception as e:

                    st.error(
                        f"Error: {e}"
                    )

    st.markdown("</div>", unsafe_allow_html=True)


    # -----------------------------------------------------
    # BOOK CATALOGUE (inline editable + deletable rows)
    # -----------------------------------------------------

    books = get_all_books()

    if books:

        search = st.text_input(
            "🔎 Search books",
            placeholder="Search by book ID, title, author or category..."
        )

        filtered_books = books

        if search:

            search_lower = search.lower()

            filtered_books = [
                book
                for book in books
                if search_lower in str(book[0]).lower()
                or search_lower in str(book[1]).lower()
                or search_lower in str(book[2]).lower()
                or search_lower in str(book[3]).lower()
            ]

        data = []

        for book in filtered_books:

            status = (
                "Available"
                if book[4] == 1
                else "Issued"
            )

            data.append([
                book[0],
                book[1],
                book[2],
                book[3],
                status,
                book[5]
            ])

        df = pd.DataFrame(
            data,
            columns=[
                "Book ID",
                "Title",
                "Author",
                "Category",
                "Status",
                "Condition"
            ]
        )

        st.caption(
            "Select a row's checkbox and press the 🗑️ icon (or Delete key) "
            "to remove it — then click Save Changes."
        )

        edited_df = st.data_editor(
            df,
            use_container_width=True,
            hide_index=True,
            num_rows="dynamic",
            key=f"book_catalogue_editor_{st.session_state.books_version}",
            column_config={
                "Status": st.column_config.TextColumn(
                    "Status",
                    disabled=True,
                    help="Managed automatically via Issue Book / Return Book"
                ),
                "Condition": st.column_config.TextColumn(
                    "Condition",
                    disabled=True,
                    help="Managed automatically via Return Book"
                )
            }
        )

        if st.button(
            "💾 Save Changes",
            use_container_width=True,
            key="save_book_catalogue"
        ):

            errors = []

            original_ids = set(df["Book ID"])
            edited_ids = set(edited_df["Book ID"])

            # ---- deleted rows ----
            for _, original_row in df.iterrows():

                if original_row["Book ID"] not in edited_ids:

                    try:

                        delete_book(original_row["Book ID"])

                    except ValueError as e:

                        errors.append(str(e))

                    except Exception as e:

                        errors.append(f"Error deleting row: {e}")

            # ---- new rows added via the editor's "+" row ----
            for _, edited_row in edited_df.iterrows():

                if edited_row["Book ID"] in original_ids:
                    continue

                new_book_id = str(edited_row["Book ID"]).strip()
                new_title = str(edited_row["Title"]).strip()
                new_author = str(edited_row["Author"]).strip()
                new_category = str(edited_row["Category"]).strip()

                if not new_book_id and not new_title:
                    # blank placeholder row - ignore
                    continue

                if not new_book_id or not new_title:
                    errors.append(
                        "New row: Book ID and Title cannot be empty."
                    )
                    continue

                try:

                    add_book(
                        new_book_id,
                        new_title,
                        new_author,
                        new_category
                    )

                except sqlite3.IntegrityError:

                    errors.append(
                        f"Book ID '{new_book_id}' already exists."
                    )

                except Exception as e:

                    errors.append(f"Error adding row: {e}")

            # ---- edited existing rows ----
            for _, original_row in df.iterrows():

                if original_row["Book ID"] not in edited_ids:
                    continue

                edited_row = edited_df[
                    edited_df["Book ID"] == original_row["Book ID"]
                ].iloc[0]

                if original_row.equals(edited_row):
                    continue

                new_book_id = str(edited_row["Book ID"]).strip()
                new_title = str(edited_row["Title"]).strip()
                new_author = str(edited_row["Author"]).strip()
                new_category = str(edited_row["Category"]).strip()

                if not new_book_id or not new_title:
                    errors.append(
                        f"Row for '{original_row['Title']}': "
                        "Book ID and Title cannot be empty."
                    )
                    continue

                try:

                    update_book(
                        original_row["Book ID"],
                        new_book_id,
                        new_title,
                        new_author,
                        new_category
                    )

                except sqlite3.IntegrityError:

                    errors.append(
                        f"Book ID '{new_book_id}' is already "
                        "used by another book."
                    )

                except Exception as e:

                    errors.append(f"Error updating row: {e}")

            if errors:

                for error_message in errors:
                    st.error(error_message)

            else:

                st.success("Changes saved successfully.")
                st.rerun()

    else:

        st.info(
            "No books found."
        )


    # -----------------------------------------------------
    # WAITING LIST
    # -----------------------------------------------------

    st.markdown(
        '<div class="card">',
        unsafe_allow_html=True
    )

    st.markdown(
        '<div class="card-title">🕒 Waiting List</div>',
        unsafe_allow_html=True
    )

    st.markdown(
        '<div class="card-subtitle">Queue an employee for a book that\'s currently issued</div>',
        unsafe_allow_html=True
    )

    all_books_for_waitlist = get_all_books()
    all_members_for_waitlist = get_all_members()

    unavailable_books = [
        book for book in all_books_for_waitlist
        if book[4] == 0
    ]

    if not unavailable_books:

        st.info("No books are currently issued, so there's nothing to wait for.")

    elif not all_members_for_waitlist:

        st.info("Add an employee first before adding them to a waiting list.")

    else:

        wl_col1, wl_col2 = st.columns(2)

        with wl_col1:

            wl_book_options = {
                f"{book[0]} — {book[1]}": book
                for book in unavailable_books
            }

            wl_selected_book_label = st.selectbox(
                "Book (currently issued)",
                list(wl_book_options.keys()),
                key="waitlist_book_select"
            )

            wl_selected_book = wl_book_options[wl_selected_book_label]

        with wl_col2:

            wl_employee_options = {
                f"{member[0]} — {member[1]}": member
                for member in all_members_for_waitlist
            }

            wl_selected_employee_label = st.selectbox(
                "Employee",
                list(wl_employee_options.keys()),
                key="waitlist_employee_select"
            )

            wl_selected_employee = wl_employee_options[wl_selected_employee_label]

        if st.button(
            "🕒 Add to Waiting List",
            use_container_width=True,
            key="add_to_waitlist_btn"
        ):

            try:

                add_to_waiting_list(
                    wl_selected_book[0],
                    wl_selected_book[1],
                    wl_selected_employee[0],
                    wl_selected_employee[1],
                    wl_selected_employee[2]
                )

                st.success(
                    f"{wl_selected_employee[1]} added to the waiting list "
                    f"for '{wl_selected_book[1]}'."
                )
                st.rerun()

            except ValueError as e:

                st.error(str(e))

            except Exception as e:

                st.error(f"Error: {e}")

    current_waiting_list = get_waiting_list()

    if current_waiting_list:

        st.markdown("#### Current Waiting List")

        wl_data = []

        for entry in current_waiting_list:

            wl_data.append([
                entry[1],
                entry[2],
                entry[3],
                entry[4],
                entry[6]
            ])

        wl_df = pd.DataFrame(
            wl_data,
            columns=[
                "Book ID",
                "Book Title",
                "Employee ID",
                "Employee Name",
                "Requested On"
            ]
        )

        st.dataframe(
            wl_df,
            use_container_width=True,
            hide_index=True
        )

        remove_wl_options = {
            f"{entry[4]} waiting for '{entry[2]}'": entry[0]
            for entry in current_waiting_list
        }

        remove_col1, remove_col2 = st.columns([3, 1])

        with remove_col1:

            selected_remove_label = st.selectbox(
                "Remove an entry",
                list(remove_wl_options.keys()),
                key="waitlist_remove_select",
                label_visibility="collapsed"
            )

        with remove_col2:

            if st.button(
                "🗑️ Remove",
                use_container_width=True,
                key="remove_from_waitlist_btn"
            ):

                remove_from_waiting_list(
                    remove_wl_options[selected_remove_label]
                )
                st.success("Removed from waiting list.")
                st.rerun()

    st.markdown("</div>", unsafe_allow_html=True)


# =========================================================
# ISSUE BOOK
# =========================================================

elif page == "📖 Issue Book":

    page_header(
        "Issue Book",
        "Issue an available book to an employee"
    )

    members = get_all_members()
    books = get_all_books()

    available_books = [
        book
        for book in books
        if book[4] == 1
    ]

    if not members:

        st.warning(
            "Please add employees before issuing a book."
        )

    elif not available_books:

        st.warning(
            "There are currently no available books."
        )

    else:

        st.markdown(
            '<div class="card">',
            unsafe_allow_html=True
        )

        st.markdown(
            '<div class="card-title">📖 Issue Book</div>',
            unsafe_allow_html=True
        )

        st.markdown(
            '<div class="card-subtitle">Select employee, book and issue dates</div>',
            unsafe_allow_html=True
        )

        # -------------------------------------------------
        # EMPLOYEE
        # -------------------------------------------------

        employee_options = {
            f"{member[0]} — {member[1]}": member
            for member in members
        }

        selected_employee_text = st.selectbox(
            "Select Employee",
            list(employee_options.keys())
        )

        employee = employee_options[
            selected_employee_text
        ]

        employee_id = employee[0]
        employee_name = employee[1]
        employee_email = employee[2]

        # -------------------------------------------------
        # BOOK
        # -------------------------------------------------

        book_options = {
            f"{book[0]} — {book[1]}": book
            for book in available_books
        }

        selected_book_text = st.selectbox(
            "Select Book",
            list(book_options.keys())
        )

        book = book_options[
            selected_book_text
        ]

        book_id = book[0]
        book_title = book[1]

        # -------------------------------------------------
        # DATES
        # -------------------------------------------------

        col1, col2 = st.columns(2)

        with col1:

            issue_date = st.date_input(
                "Issue Date",
                value=date.today()
            )

        with col2:

            due_date = st.date_input(
                "Due Date",
                value=date.today() + timedelta(days=14)
            )

        st.divider()

        # -------------------------------------------------
        # PREVIEW
        # -------------------------------------------------

        st.markdown("### Issue Summary")

        col1, col2, col3 = st.columns(3)

        with col1:

            st.write("**Employee**")
            st.write(employee_name)

        with col2:

            st.write("**Book**")
            st.write(book_title)

        with col3:

            st.write("**Due Date**")
            st.write(due_date)

        st.write("")

        # -------------------------------------------------
        # ISSUE
        # -------------------------------------------------

        if st.button(
            "📖 Issue Book",
            use_container_width=True
        ):

            if due_date < issue_date:

                st.error(
                    "Due Date cannot be before Issue Date."
                )

            else:

                connection = get_connection()
                cursor = connection.cursor()

                try:

                    cursor.execute("""
                        INSERT INTO issued_books
                        (
                            employee_id,
                            employee_name,
                            employee_email,
                            book_id,
                            book_title,
                            issue_date,
                            due_date,
                            return_date,
                            status,
                            reminder_sent
                        )
                        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                    """, (
                        employee_id,
                        employee_name,
                        employee_email,
                        book_id,
                        book_title,
                        issue_date.isoformat(),
                        due_date.isoformat(),
                        None,
                        "Issued",
                        0
                    ))

                    cursor.execute("""
                        UPDATE books
                        SET available = 0
                        WHERE book_id = ?
                    """, (
                        book_id,
                    ))

                    connection.commit()

                    st.session_state.books_version += 1

                    st.success(
                        f"'{book_title}' has been issued to {employee_name}."
                    )

                    try:
                        export_database_to_excel()
                    except Exception as sync_error:
                        st.warning(
                            f"Book issued, but couldn't sync to Excel: {sync_error}"
                        )

                except Exception as e:

                    connection.rollback()

                    st.error(
                        f"Error issuing book: {e}"
                    )

                finally:

                    connection.close()

        st.markdown("</div>", unsafe_allow_html=True)


# =========================================================
# RETURN BOOK
# =========================================================

elif page == "↩️ Return Book":

    page_header(
        "Return Book",
        "Process a book returned by an employee"
    )

    issued_books = get_currently_issued_books()

    if not issued_books:

        st.success(
            "There are currently no issued books."
        )

    else:

        st.markdown(
            '<div class="card">',
            unsafe_allow_html=True
        )

        st.markdown(
            '<div class="card-title">↩️ Return Book</div>',
            unsafe_allow_html=True
        )

        st.markdown(
            '<div class="card-subtitle">Select the book being returned</div>',
            unsafe_allow_html=True
        )

        # -------------------------------------------------
        # SELECT BOOK
        # -------------------------------------------------

        book_options = {
            f"{book[5]} — {book[2]} (Due: {book[7]})": book
            for book in issued_books
        }

        selected_text = st.selectbox(
            "Select Issued Book",
            list(book_options.keys())
        )

        selected_book = book_options[
            selected_text
        ]

        # -------------------------------------------------
        # GET DETAILS
        # -------------------------------------------------

        issue_record_id = selected_book[0]
        employee_id = selected_book[1]
        employee_name = selected_book[2]
        employee_email = selected_book[3]
        book_id = selected_book[4]
        book_title = selected_book[5]
        issue_date = selected_book[6]
        due_date = selected_book[7]

        # -------------------------------------------------
        # EMPLOYEE DETAILS
        # -------------------------------------------------

        st.markdown("### 👤 Employee Details")

        col1, col2, col3 = st.columns(3)

        with col1:

            st.caption("Employee ID")
            st.write(employee_id)

        with col2:

            st.caption("Employee Name")
            st.write(employee_name)

        with col3:

            st.caption("Email")
            st.write(employee_email)

        st.divider()

        # -------------------------------------------------
        # BOOK DETAILS
        # -------------------------------------------------

        st.markdown("### 📚 Book Details")

        col1, col2, col3 = st.columns(3)

        with col1:

            st.caption("Book ID")
            st.write(book_id)

        with col2:

            st.caption("Book Title")
            st.write(book_title)

        with col3:

            st.caption("Issue Date")
            st.write(issue_date)

        col1, col2 = st.columns(2)

        with col1:

            st.caption("Due Date")
            st.write(due_date)

        with col2:

            return_date = st.date_input(
                "Return Date",
                value=date.today()
            )

        book_condition = st.selectbox(
            "Book Condition",
            ["Good", "Fair", "Damaged", "Lost"]
        )

        confirm_return = st.checkbox(
            f"I confirm '{book_title}' has been returned by {employee_name} "
            f"in {book_condition.lower()} condition."
        )

        st.divider()

        # -------------------------------------------------
        # RETURN BUTTON
        # -------------------------------------------------

        if st.button(
            "↩️ Confirm Return",
            use_container_width=True,
            disabled=not confirm_return
        ):

            connection = get_connection()
            cursor = connection.cursor()

            try:

                cursor.execute("""
                    UPDATE issued_books
                    SET
                        return_date = ?,
                        status = 'Returned'
                    WHERE id = ?
                """, (
                    return_date.isoformat(),
                    issue_record_id
                ))

                # A lost/damaged book shouldn't go straight back into
                # the "Available" pool for others to borrow.
                new_available = 1 if book_condition in ("Good", "Fair") else 0

                cursor.execute("""
                    UPDATE books
                    SET available = ?,
                        condition = ?
                    WHERE book_id = ?
                """, (
                    new_available,
                    book_condition,
                    book_id,
                ))

                connection.commit()

                st.session_state.books_version += 1

                st.toast(
                    f"'{book_title}' returned by {employee_name} "
                    f"({book_condition}).",
                    icon="✅"
                )

                waiting_for_this_book = get_waiting_list(book_id)

                if waiting_for_this_book and new_available == 1:
                    next_in_line = waiting_for_this_book[0]
                    st.toast(
                        f"📋 {next_in_line[4]} is next on the waiting list "
                        f"for '{book_title}'.",
                        icon="🕒"
                    )

                try:
                    export_database_to_excel()
                except Exception as sync_error:
                    st.warning(
                        f"Book returned, but couldn't sync to Excel: {sync_error}"
                    )

                connection.close()

                # Refresh the list so the just-returned book drops out
                # of the "Select Issued Book" dropdown immediately.
                st.rerun()

            except Exception as e:

                connection.rollback()
                connection.close()

                st.error(
                    f"Error returning book: {e}"
                )

        st.markdown("</div>", unsafe_allow_html=True)
