"""
=====================================================================
 RJK Student Management System - Streamlit Web Edition
 Developer/Owner: Rai Jahanzaib

 Converted from the original single-file Tkinter/SQLite desktop
 application to a single-file Streamlit web application so it can be
 deployed on Streamlit Cloud (Tkinter cannot run on cloud servers,
 since there is no display / windowing system there).

 FIRST RUN CREDENTIALS (change these after logging in!):
     Admin:      admin / admin123
     Management: management / management123
     Teacher:    teacher / teacher123
     Staff:      staff / staff123

 ACTIVATION KEY (basic local check, NOT secure commercial licensing):
     RJK-278-LEF-1185

 Run locally with:
     streamlit run streamlit_student_management_system.py

 Requires Python 3.8+ and:
     pip install streamlit psycopg2-binary
 Optional packages (the app still runs without them, just with Excel
 / PDF export features disabled):
     pip install openpyxl reportlab

 DATA IS STORED IN SUPABASE (Postgres), NOT SQLite. Before running,
 you must set SUPABASE_DB_URL in Streamlit secrets -- see the setup
 instructions provided alongside this file (secrets.toml.example).
=====================================================================

WHAT CHANGED VS THE TKINTER VERSION (for maintainers)
---------------------------------------------------------------------
- All Tkinter/ttk windows, frames and the mainloop are gone. Every
  "screen" is now a function that renders Streamlit widgets.
- tkinter.messagebox.* -> st.success / st.error / st.warning / st.info
- Tkinter Entry/Combobox/Treeview -> st.text_input / st.selectbox /
  st.dataframe (and st.data_editor for the editable attendance grid)
- filedialog.askopenfilename (student photo) -> st.file_uploader
- filedialog.asksaveasfilename (Excel/PDF export) -> files are now
  built entirely in memory (io.BytesIO) and handed to the user with
  st.download_button, since a Streamlit Cloud server has no concept
  of "the user's Downloads folder" the way a desktop app does.
- tkinter class-variable state (current user, current view, form
  data while a dialog is open) -> st.session_state.
- The DATABASE (SQLite), ALL table schemas, ALL business rules /
  permission checks, password hashing, and the grade calculation are
  kept exactly as they were in the desktop app - only the layer that
  talks to the user changed.
=====================================================================
"""

# =====================================================================
# 1. IMPORTS
# =====================================================================
import os
import io
import shutil
import hashlib
import binascii
import traceback
import uuid
from datetime import datetime, date

import streamlit as st
import psycopg2
import psycopg2.extras

# ---- Optional third-party libraries -------------------------------------
try:
    import openpyxl
    from openpyxl.styles import Font, PatternFill, Alignment
    from openpyxl.utils import get_column_letter
    OPENPYXL_AVAILABLE = True
except ImportError:
    OPENPYXL_AVAILABLE = False

try:
    from reportlab.lib.pagesizes import A4
    from reportlab.lib import colors
    from reportlab.lib.units import mm
    from reportlab.platypus import (
        SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, Image as RLImage
    )
    from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
    from reportlab.lib.enums import TA_CENTER
    REPORTLAB_AVAILABLE = True
except ImportError:
    REPORTLAB_AVAILABLE = False


# =====================================================================
# 2. CONFIGURATION / CONSTANTS  (unchanged from the desktop app)
# =====================================================================
APP_NAME = "RJK Student Management System"
APP_OWNER = "Rai Jahanzaib"
PHOTOS_DIR = "student_photos"  # still local disk -- see note in DatabaseManager below
ERROR_LOG_FILE = "app_errors.log"
ACTIVATION_KEY = "RJK-278-LEF-1185"

ROLES = ["admin", "management", "teacher", "staff"]
ATTENDANCE_STATUSES = ["Present", "Absent", "Leave"]

GRADE_BOUNDARIES = [
    (90, "A+"),
    (80, "A"),
    (70, "B"),
    (60, "C"),
    (50, "D"),
    (40, "E"),
    (0, "F"),
]
PASS_PERCENTAGE = 40.0


def calculate_grade(percentage):
    """Return the letter grade for a given percentage. Easy to edit."""
    for boundary, grade in GRADE_BOUNDARIES:
        if percentage >= boundary:
            return grade
    return "F"


def log_error_to_file(context, exc):
    """Write unexpected errors to a log file so a beginner can inspect them."""
    try:
        with open(ERROR_LOG_FILE, "a", encoding="utf-8") as f:
            f.write(f"\n[{datetime.now()}] Context: {context}\n")
            f.write(traceback.format_exc())
            f.write("\n")
    except Exception:
        pass  # Never let logging itself crash the app


def show_exception(context, exc):
    """Streamlit replacement for the old messagebox.showerror error popup."""
    log_error_to_file(context, exc)
    st.error(
        f"Something went wrong while: {context}\n\n"
        f"Details: {exc}\n\n"
        f"(A full log was saved to {ERROR_LOG_FILE})"
    )


# =====================================================================
# 3 & 4. DATABASE INITIALIZATION + PASSWORD HASHING  (unchanged logic)
# =====================================================================
class DatabaseManager:
    """
    Owns the single Postgres (Supabase) connection and every raw
    table-creation / low level query. Higher-level permission checks
    live in Session (see below) -- this class does NOT know about
    "current user"; it just executes whatever query it is given with
    parameters, always using parameterized SQL (never string
    concatenation).

    NOTE ON PHOTOS: student photos are still written to a local
    PHOTOS_DIR folder on whatever machine runs the app. That's fine
    for local use, but on Streamlit Cloud that folder is wiped on
    every redeploy/restart -- only the DATA in this class (students,
    classes, attendance, marks, users) is now permanent in Supabase.
    If you also want photos to survive redeploys, move them to a
    Supabase Storage bucket -- ask and it can be added the same way.

    This class reads the ORIGINAL queries written for SQLite (using
    "?" placeholders and INTEGER PRIMARY KEY AUTOINCREMENT). Rather
    than rewrite every single query throughout Session, this class
    translates "?" -> "%s" and returns dict-like rows, so all the
    business logic in Session works completely unchanged.
    """

    def __init__(self):
        # Supabase connection details come from Streamlit secrets, never
        # hardcoded in the file. See secrets.toml.example / the setup
        # instructions for exactly what to put in .streamlit/secrets.toml
        # (locally) or the app's "Secrets" settings (Streamlit Cloud).
     db_url = st.secrets.get("SUPABASE_DB_URL")
if not db_url:
    st.error("⚠️ SUPABASE_DB_URL missing hai! Pehle Streamlit ke Secrets mein database link dalein.")
    st.stop() # Yeh app ko crash nahi hone dega, bas yahin rok dega

        self.conn = psycopg2.connect(db_url, cursor_factory=psycopg2.extras.RealDictCursor)
        self.conn.autocommit = False
        self._create_tables()
        self._create_defaults_if_empty()
        os.makedirs(PHOTOS_DIR, exist_ok=True)

    # ---------------------------------------------------------------
    @staticmethod
    def _translate(sql):
        """SQLite used '?' placeholders; psycopg2/Postgres uses '%s'."""
        return sql.replace("?", "%s")

    # ---------------------------------------------------------------
    def _create_tables(self):
        cur = self.conn.cursor()
        cur.execute("""
        CREATE TABLE IF NOT EXISTS users (
            id SERIAL PRIMARY KEY,
            username TEXT UNIQUE NOT NULL,
            password_hash TEXT NOT NULL,
            salt TEXT NOT NULL,
            full_name TEXT NOT NULL,
            role TEXT NOT NULL CHECK(role IN ('admin','management','teacher','staff')),
            email TEXT,
            phone TEXT,
            is_active INTEGER NOT NULL DEFAULT 1,
            created_at TEXT NOT NULL,
            updated_at TEXT NOT NULL
        );

        CREATE TABLE IF NOT EXISTS classes (
            id SERIAL PRIMARY KEY,
            class_name TEXT NOT NULL,
            section TEXT,
            academic_year TEXT,
            is_active INTEGER NOT NULL DEFAULT 1,
            created_at TEXT NOT NULL,
            UNIQUE(class_name, section, academic_year)
        );

        CREATE TABLE IF NOT EXISTS teacher_classes (
            id SERIAL PRIMARY KEY,
            teacher_id INTEGER NOT NULL,
            class_id INTEGER NOT NULL,
            FOREIGN KEY(teacher_id) REFERENCES users(id) ON DELETE CASCADE,
            FOREIGN KEY(class_id) REFERENCES classes(id) ON DELETE CASCADE,
            UNIQUE(teacher_id, class_id)
        );

        CREATE TABLE IF NOT EXISTS students (
            id SERIAL PRIMARY KEY,
            student_id TEXT UNIQUE NOT NULL,
            roll_number TEXT,
            first_name TEXT NOT NULL,
            last_name TEXT,
            father_name TEXT,
            date_of_birth TEXT,
            gender TEXT,
            phone TEXT,
            address TEXT,
            class_id INTEGER,
            photo_path TEXT,
            admission_date TEXT,
            status TEXT NOT NULL DEFAULT 'Active',
            created_at TEXT NOT NULL,
            updated_at TEXT NOT NULL,
            FOREIGN KEY(class_id) REFERENCES classes(id) ON DELETE SET NULL
        );

        CREATE TABLE IF NOT EXISTS attendance (
            id SERIAL PRIMARY KEY,
            student_id INTEGER NOT NULL,
            attendance_date TEXT NOT NULL,
            status TEXT NOT NULL CHECK(status IN ('Present','Absent','Leave')),
            FOREIGN KEY(student_id) REFERENCES students(id) ON DELETE CASCADE,
            UNIQUE(student_id, attendance_date)
        );

        CREATE TABLE IF NOT EXISTS marks (
            id SERIAL PRIMARY KEY,
            student_id INTEGER NOT NULL,
            subject TEXT NOT NULL,
            exam_name TEXT NOT NULL,
            total_marks REAL NOT NULL,
            obtained_marks REAL NOT NULL,
            created_at TEXT NOT NULL,
            updated_at TEXT NOT NULL,
            FOREIGN KEY(student_id) REFERENCES students(id) ON DELETE CASCADE
        );

        CREATE TABLE IF NOT EXISTS settings (
            key TEXT PRIMARY KEY,
            value TEXT
        );

        CREATE TABLE IF NOT EXISTS audit_log (
            id SERIAL PRIMARY KEY,
            user_id INTEGER,
            action TEXT NOT NULL,
            description TEXT,
            timestamp TEXT NOT NULL
        );
        """)
        self.conn.commit()

    # ---------------------------------------------------------------
    def _create_defaults_if_empty(self):
        """Runs once: seeds the four demo accounts + default settings
        the very first time this Supabase database is used (instead of
        the old 'file didn't exist yet' check, which doesn't apply to
        a remote database)."""
        cur = self.conn.cursor()
        cur.execute("SELECT COUNT(*) AS c FROM users")
        if cur.fetchone()["c"] > 0:
            return  # already seeded

        now = datetime.now().isoformat(timespec="seconds")
        defaults = [
            ("admin", "admin123", "Administrator", "admin"),
            ("management", "management123", "Management User", "management"),
            ("teacher", "teacher123", "Demo Teacher", "teacher"),
            ("staff", "staff123", "Demo Staff", "staff"),
        ]
        for username, password, full_name, role in defaults:
            pwd_hash, salt = hash_password(password)
            cur.execute(
                """INSERT INTO users
                   (username, password_hash, salt, full_name, role, email,
                    phone, is_active, created_at, updated_at)
                   VALUES (%s, %s, %s, %s, %s, '', '', 1, %s, %s)""",
                (username, pwd_hash, salt, full_name, role, now, now)
            )
        cur.execute("INSERT INTO settings (key, value) VALUES (%s, %s) ON CONFLICT DO NOTHING",
                     ("activated", "0"))
        cur.execute("INSERT INTO settings (key, value) VALUES (%s, %s) ON CONFLICT DO NOTHING",
                     ("school_name", "RJK School / College"))
        cur.execute("INSERT INTO settings (key, value) VALUES (%s, %s) ON CONFLICT DO NOTHING",
                     ("school_address", ""))
        cur.execute("INSERT INTO settings (key, value) VALUES (%s, %s) ON CONFLICT DO NOTHING",
                     ("school_phone", ""))
        self.conn.commit()

    # ---------------------------------------------------------------
    # Generic helpers -- ALWAYS parameterized, never string-built SQL.
    def query(self, sql, params=()):
        sql = self._translate(sql)
        cur = self.conn.cursor()
        try:
            cur.execute(sql, params)
            return cur.fetchall()
        except Exception:
            self.conn.rollback()
            raise

    def query_one(self, sql, params=()):
        sql = self._translate(sql)
        cur = self.conn.cursor()
        try:
            cur.execute(sql, params)
            return cur.fetchone()
        except Exception:
            self.conn.rollback()
            raise

    def execute(self, sql, params=()):
        sql = self._translate(sql)
        cur = self.conn.cursor()
        # Every app table's primary key is called "id" except settings
        # (keyed by "key") -- auto-append RETURNING id on plain INSERTs
        # so callers get the new row's id back, the same way SQLite's
        # cursor.lastrowid used to work.
        wants_id = (
            sql.strip().upper().startswith("INSERT")
            and "RETURNING" not in sql.upper()
            and "settings" not in sql.lower()
        )
        if wants_id:
            sql += " RETURNING id"
        try:
            cur.execute(sql, params)
            new_id = None
            if wants_id:
                row = cur.fetchone()
                new_id = row["id"] if row else None
            self.conn.commit()
            return new_id
        except Exception:
            self.conn.rollback()
            raise

    def get_setting(self, key, default=None):
        row = self.query_one("SELECT value FROM settings WHERE key=?", (key,))
        return row["value"] if row else default

    def set_setting(self, key, value):
        self.execute(
            "INSERT INTO settings (key, value) VALUES (?, ?) "
            "ON CONFLICT(key) DO UPDATE SET value=excluded.value",
            (key, value)
        )


# ---------------------------------------------------------------------
# Password hashing: PBKDF2-HMAC-SHA256 with a random salt per user.
# ---------------------------------------------------------------------
def hash_password(password, salt=None):
    if salt is None:
        salt = binascii.hexlify(os.urandom(16)).decode("utf-8")
    dk = hashlib.pbkdf2_hmac(
        "sha256", password.encode("utf-8"), salt.encode("utf-8"), 100_000
    )
    return binascii.hexlify(dk).decode("utf-8"), salt


def verify_password(password, stored_hash, salt):
    test_hash, _ = hash_password(password, salt)
    return test_hash == stored_hash


# =====================================================================
# 5. CENTRALIZED PERMISSION SYSTEM  (unchanged from the desktop app)
# =====================================================================
PERMISSIONS = {
    "admin": {
        "view_dashboard", "manage_users_add", "manage_users_edit",
        "manage_users_delete", "manage_users_disable", "manage_users_reset_pw",
        "manage_classes_add", "manage_classes_edit", "manage_classes_delete",
        "manage_classes_view", "assign_teacher", "remove_teacher_assignment",
        "students_view_all", "students_add", "students_edit", "students_delete",
        "attendance_view", "attendance_edit",
        "marks_view", "marks_edit",
        "reports_view", "export_excel", "export_pdf",
        "view_audit_log", "manage_settings",
    },
    "management": {
        "view_dashboard",
        "manage_users_add_teacher",  # can ONLY create teacher accounts
        "manage_users_edit_teacher",
        "manage_classes_view",
        "assign_teacher_limited",
        "students_view_all", "students_add", "students_edit",
        "attendance_view", "attendance_edit",
        "marks_view", "marks_edit",
        "reports_view", "export_excel", "export_pdf",
    },
    "teacher": {
        "view_dashboard",
        "manage_classes_view_assigned",
        "students_view_assigned", "students_add_assigned", "students_edit_assigned",
        "attendance_view_assigned", "attendance_edit_assigned",
        "marks_view_assigned", "marks_edit_assigned",
        "reports_view_assigned", "export_excel_assigned", "export_pdf_assigned",
    },
    "staff": {
        "view_dashboard",
        "manage_classes_view",
        "students_view_all",
        "attendance_view",
        "marks_view",
        "reports_view_limited", "export_limited",
    },
}


def has_permission(role, permission):
    """Central permission check. Returns True/False, never raises."""
    return permission in PERMISSIONS.get(role, set())


class PermissionError_(Exception):
    """Raised internally when a backend check fails (not just a UI hide)."""
    pass


# =====================================================================
# 6. SESSION / AUTHENTICATION / BUSINESS LOGIC  (unchanged behaviour;
#    only the two student-photo methods were adapted to take raw
#    uploaded bytes instead of a local filesystem path, since a
#    Streamlit file_uploader never gives you a path on disk).
# =====================================================================
class Session:
    """
    Holds the DatabaseManager and the currently logged-in user, and
    exposes every "business logic" operation (add student, save
    attendance, etc). Each method re-checks permissions and, for
    teachers, re-checks class ownership DIRECTLY IN SQL -- so even if
    the UI were bypassed, unauthorized data can never be returned or
    modified.
    """

    def __init__(self, db: DatabaseManager):
        self.db = db
        self.current_user = None  # dict-like row (RealDictRow) once logged in

    # -----------------------------------------------------------------
    # AUTHENTICATION
    # -----------------------------------------------------------------
    def login(self, username, password):
        user = self.db.query_one("SELECT * FROM users WHERE username = ?", (username,))
        if user is None:
            return None, "Invalid username or password."
        if not verify_password(password, user["password_hash"], user["salt"]):
            return None, "Invalid username or password."
        if not user["is_active"]:
            return None, "This account has been disabled. Contact the Admin."
        self.current_user = user
        self.log_audit("login", f"User '{username}' logged in.")
        return user, None

    def logout(self):
        if self.current_user:
            self.log_audit("logout", f"User '{self.current_user['username']}' logged out.")
        self.current_user = None

    def log_audit(self, action, description=""):
        uid = self.current_user["id"] if self.current_user else None
        self.db.execute(
            "INSERT INTO audit_log (user_id, action, description, timestamp) "
            "VALUES (?, ?, ?, ?)",
            (uid, action, description, datetime.now().isoformat(timespec="seconds"))
        )

    # -----------------------------------------------------------------
    def require(self, permission):
        """Raise if current user lacks a permission. Call at top of every
        sensitive method."""
        if self.current_user is None:
            raise PermissionError_("Not logged in.")
        if not has_permission(self.current_user["role"], permission):
            raise PermissionError_(
                f"Role '{self.current_user['role']}' does not have permission "
                f"'{permission}'."
            )

    def teacher_owns_class(self, class_id):
        """True only if the current (teacher) user is assigned to class_id."""
        if self.current_user["role"] != "teacher":
            return False
        row = self.db.query_one(
            "SELECT 1 FROM teacher_classes WHERE teacher_id=? AND class_id=?",
            (self.current_user["id"], class_id)
        )
        return row is not None

    def teacher_owns_student(self, student_id):
        row = self.db.query_one("SELECT s.class_id FROM students s WHERE s.id=?", (student_id,))
        if row is None or row["class_id"] is None:
            return False
        return self.teacher_owns_class(row["class_id"])

    # ===================================================================
    # USERS (admin, + limited management)
    # ===================================================================
    def list_users(self):
        return self.db.query("SELECT * FROM users ORDER BY full_name")

    def require_any(self, permissions):
        if self.current_user is None:
            raise PermissionError_("Not logged in.")
        if not any(has_permission(self.current_user["role"], p) for p in permissions):
            raise PermissionError_("You do not have permission to perform this action.")

    def add_user(self, username, password, full_name, role, email, phone):
        if self.current_user["role"] == "management":
            self.require("manage_users_add_teacher")
            if role != "teacher":
                raise PermissionError_("Management can only create Teacher accounts.")
        else:
            self.require("manage_users_add")
            if role not in ROLES:
                raise ValueError("Invalid role.")

        existing = self.db.query_one("SELECT id FROM users WHERE username=?", (username,))
        if existing:
            raise ValueError(f"Username '{username}' already exists.")

        pwd_hash, salt = hash_password(password)
        now = datetime.now().isoformat(timespec="seconds")
        new_id = self.db.execute(
            """INSERT INTO users (username, password_hash, salt, full_name, role,
               email, phone, is_active, created_at, updated_at)
               VALUES (?, ?, ?, ?, ?, ?, ?, 1, ?, ?)""",
            (username, pwd_hash, salt, full_name, role, email, phone, now, now)
        )
        self.log_audit("add_user", f"Added user '{username}' with role '{role}'.")
        return new_id

    def edit_user(self, user_id, full_name, email, phone, role=None):
        target = self.db.query_one("SELECT * FROM users WHERE id=?", (user_id,))
        if target is None:
            raise ValueError("User not found.")

        if self.current_user["role"] == "management":
            self.require("manage_users_edit_teacher")
            if target["role"] != "teacher":
                raise PermissionError_("Management may only edit Teacher accounts.")
            role = "teacher"  # cannot change a teacher's role away from teacher
        else:
            self.require("manage_users_edit")
            if target["role"] == "admin" and role is not None and role != "admin":
                raise PermissionError_("Cannot change an Admin's role.")

        final_role = role if role else target["role"]
        now = datetime.now().isoformat(timespec="seconds")
        self.db.execute(
            "UPDATE users SET full_name=?, email=?, phone=?, role=?, updated_at=? WHERE id=?",
            (full_name, email, phone, final_role, now, user_id)
        )
        self.log_audit("edit_user", f"Edited user id={user_id}.")

    def set_user_active(self, user_id, active):
        self.require("manage_users_disable")
        target = self.db.query_one("SELECT * FROM users WHERE id=?", (user_id,))
        if target is None:
            raise ValueError("User not found.")
        if target["id"] == self.current_user["id"]:
            raise PermissionError_("You cannot disable your own account.")
        if target["role"] == "admin" and not active:
            admin_count = self.db.query_one(
                "SELECT COUNT(*) c FROM users WHERE role='admin' AND is_active=1"
            )["c"]
            if admin_count <= 1:
                raise PermissionError_("Cannot disable the last remaining active Admin.")
        now = datetime.now().isoformat(timespec="seconds")
        self.db.execute(
            "UPDATE users SET is_active=?, updated_at=? WHERE id=?",
            (1 if active else 0, now, user_id)
        )
        self.log_audit("disable_user" if not active else "enable_user",
                        f"User id={user_id} set active={active}.")

    def delete_user(self, user_id):
        self.require("manage_users_delete")
        target = self.db.query_one("SELECT * FROM users WHERE id=?", (user_id,))
        if target is None:
            raise ValueError("User not found.")
        if target["id"] == self.current_user["id"]:
            raise PermissionError_("You cannot delete your own account while logged in.")
        if target["role"] == "admin":
            admin_count = self.db.query_one("SELECT COUNT(*) c FROM users WHERE role='admin'")["c"]
            if admin_count <= 1:
                raise PermissionError_("Cannot delete the last remaining Admin account.")
        self.db.execute("DELETE FROM users WHERE id=?", (user_id,))
        self.log_audit("delete_user", f"Deleted user '{target['username']}' (id={user_id}).")

    def reset_password(self, user_id, new_password):
        self.require("manage_users_reset_pw")
        pwd_hash, salt = hash_password(new_password)
        now = datetime.now().isoformat(timespec="seconds")
        self.db.execute(
            "UPDATE users SET password_hash=?, salt=?, updated_at=? WHERE id=?",
            (pwd_hash, salt, now, user_id)
        )
        self.log_audit("reset_password", f"Password reset for user id={user_id}.")

    def change_own_password(self, old_password, new_password):
        user = self.current_user
        if not verify_password(old_password, user["password_hash"], user["salt"]):
            raise ValueError("Current password is incorrect.")
        pwd_hash, salt = hash_password(new_password)
        now = datetime.now().isoformat(timespec="seconds")
        self.db.execute(
            "UPDATE users SET password_hash=?, salt=?, updated_at=? WHERE id=?",
            (pwd_hash, salt, now, user["id"])
        )
        self.current_user = self.db.query_one("SELECT * FROM users WHERE id=?", (user["id"],))
        self.log_audit("change_password", "User changed their own password.")

    # ===================================================================
    # CLASSES
    # ===================================================================
    def list_classes(self):
        role = self.current_user["role"]
        if role in ("admin", "management", "staff"):
            return self.db.query("SELECT * FROM classes ORDER BY class_name, section")
        elif role == "teacher":
            return self.db.query(
                """SELECT c.* FROM classes c
                   JOIN teacher_classes tc ON tc.class_id = c.id
                   WHERE tc.teacher_id = ?
                   ORDER BY c.class_name, c.section""",
                (self.current_user["id"],)
            )
        return []

    def add_class(self, class_name, section, academic_year):
        self.require("manage_classes_add")
        now = datetime.now().isoformat(timespec="seconds")
        try:
            return self.db.execute(
                """INSERT INTO classes (class_name, section, academic_year, is_active, created_at)
                   VALUES (?, ?, ?, 1, ?)""",
                (class_name, section, academic_year, now)
            )
        except psycopg2.IntegrityError:
            raise ValueError("This class/section/year combination already exists.")

    def edit_class(self, class_id, class_name, section, academic_year, is_active):
        self.require("manage_classes_edit")
        self.db.execute(
            "UPDATE classes SET class_name=?, section=?, academic_year=?, is_active=? WHERE id=?",
            (class_name, section, academic_year, 1 if is_active else 0, class_id)
        )
        self.log_audit("edit_class", f"Edited class id={class_id}.")

    def delete_class(self, class_id):
        self.require("manage_classes_delete")
        student_count = self.db.query_one(
            "SELECT COUNT(*) c FROM students WHERE class_id=?", (class_id,)
        )["c"]
        if student_count > 0:
            raise PermissionError_(
                f"Cannot delete: {student_count} student(s) are still assigned to this class. "
                "Move or remove them first."
            )
        self.db.execute("DELETE FROM classes WHERE id=?", (class_id,))
        self.log_audit("delete_class", f"Deleted class id={class_id}.")

    # ===================================================================
    # TEACHER ASSIGNMENTS
    # ===================================================================
    def list_teachers(self):
        return self.db.query("SELECT * FROM users WHERE role='teacher' ORDER BY full_name")

    def list_assignments_for_teacher(self, teacher_id):
        return self.db.query(
            """SELECT c.* FROM classes c
               JOIN teacher_classes tc ON tc.class_id = c.id
               WHERE tc.teacher_id=? ORDER BY c.class_name, c.section""",
            (teacher_id,)
        )

    def assign_teacher(self, teacher_id, class_id):
        role = self.current_user["role"]
        if role == "management":
            self.require("assign_teacher_limited")
        else:
            self.require("assign_teacher")
        teacher = self.db.query_one("SELECT * FROM users WHERE id=? AND role='teacher'", (teacher_id,))
        if teacher is None:
            raise ValueError("Selected user is not a valid teacher.")
        try:
            self.db.execute(
                "INSERT INTO teacher_classes (teacher_id, class_id) VALUES (?, ?)",
                (teacher_id, class_id)
            )
        except psycopg2.IntegrityError:
            raise ValueError("This teacher is already assigned to that class.")
        self.log_audit("assign_teacher", f"Assigned teacher id={teacher_id} to class id={class_id}.")

    def remove_teacher_assignment(self, teacher_id, class_id):
        role = self.current_user["role"]
        if role == "management":
            self.require("assign_teacher_limited")
        else:
            self.require("remove_teacher_assignment")
        self.db.execute(
            "DELETE FROM teacher_classes WHERE teacher_id=? AND class_id=?",
            (teacher_id, class_id)
        )
        self.log_audit("remove_teacher_assignment",
                        f"Removed teacher id={teacher_id} from class id={class_id}.")

    # ===================================================================
    # STUDENTS  (teacher queries are ALWAYS filtered by SQL join, not
    # loaded fully then hidden)
    # ===================================================================
    def search_students(self, keyword="", class_id=None, status=None):
        role = self.current_user["role"]
        base = """SELECT s.*, c.class_name, c.section FROM students s
                   LEFT JOIN classes c ON c.id = s.class_id WHERE 1=1"""
        params = []

        if role == "teacher":
            base += """ AND s.class_id IN
                        (SELECT class_id FROM teacher_classes WHERE teacher_id = ?)"""
            params.append(self.current_user["id"])

        if keyword:
            base += """ AND (s.student_id LIKE ? OR s.roll_number LIKE ?
                              OR s.first_name LIKE ? OR s.last_name LIKE ?
                              OR s.father_name LIKE ? OR s.phone LIKE ?)"""
            like = f"%{keyword}%"
            params.extend([like, like, like, like, like, like])

        if class_id:
            base += " AND s.class_id = ?"
            params.append(class_id)

        if status:
            base += " AND s.status = ?"
            params.append(status)

        base += " ORDER BY s.first_name, s.last_name"
        return self.db.query(base, params)

    def get_student(self, student_id):
        role = self.current_user["role"]
        row = self.db.query_one(
            """SELECT s.*, c.class_name, c.section FROM students s
               LEFT JOIN classes c ON c.id = s.class_id WHERE s.id=?""",
            (student_id,)
        )
        if row is None:
            return None
        if role == "teacher" and not self.teacher_owns_student(student_id):
            raise PermissionError_("You cannot access a student outside your assigned class(es).")
        return row

    def add_student(self, data, photo_bytes=None, photo_ext=None):
        role = self.current_user["role"]
        if role == "teacher":
            self.require("students_add_assigned")
            if not self.teacher_owns_class(data["class_id"]):
                raise PermissionError_("You can only add students to your own assigned class(es).")
        else:
            self.require("students_add")

        existing = self.db.query_one("SELECT id FROM students WHERE student_id=?", (data["student_id"],))
        if existing:
            raise ValueError(f"Student ID '{data['student_id']}' already exists.")

        photo_path = None
        if photo_bytes:
            photo_path = self._save_student_photo(data["student_id"], photo_bytes, photo_ext)

        now = datetime.now().isoformat(timespec="seconds")
        new_id = self.db.execute(
            """INSERT INTO students
               (student_id, roll_number, first_name, last_name, father_name,
                date_of_birth, gender, phone, address, class_id, photo_path,
                admission_date, status, created_at, updated_at)
               VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)""",
            (data["student_id"], data["roll_number"], data["first_name"],
             data["last_name"], data["father_name"], data["date_of_birth"],
             data["gender"], data["phone"], data["address"], data["class_id"],
             photo_path, data["admission_date"], data.get("status", "Active"),
             now, now)
        )
        self.log_audit("add_student", f"Added student '{data['student_id']}'.")
        return new_id

    def edit_student(self, student_row_id, data, photo_bytes=None, photo_ext=None):
        role = self.current_user["role"]
        current = self.db.query_one("SELECT * FROM students WHERE id=?", (student_row_id,))
        if current is None:
            raise ValueError("Student not found.")

        if role == "teacher":
            self.require("students_edit_assigned")
            if not self.teacher_owns_student(student_row_id):
                raise PermissionError_("You can only edit students in your assigned class(es).")
            if not self.teacher_owns_class(data["class_id"]):
                raise PermissionError_("You cannot move a student to a class you do not teach.")
        else:
            self.require("students_edit")

        photo_path = current["photo_path"]
        if photo_bytes:
            photo_path = self._save_student_photo(data["student_id"], photo_bytes, photo_ext)

        now = datetime.now().isoformat(timespec="seconds")
        self.db.execute(
            """UPDATE students SET student_id=?, roll_number=?, first_name=?, last_name=?,
               father_name=?, date_of_birth=?, gender=?, phone=?, address=?, class_id=?,
               photo_path=?, admission_date=?, status=?, updated_at=? WHERE id=?""",
            (data["student_id"], data["roll_number"], data["first_name"], data["last_name"],
             data["father_name"], data["date_of_birth"], data["gender"], data["phone"],
             data["address"], data["class_id"], photo_path, data["admission_date"],
             data.get("status", "Active"), now, student_row_id)
        )
        self.log_audit("edit_student", f"Edited student id={student_row_id}.")

    def delete_student(self, student_row_id):
        role = self.current_user["role"]
        if role == "teacher":
            raise PermissionError_("Teachers cannot delete students.")
        self.require("students_delete")
        student = self.db.query_one("SELECT * FROM students WHERE id=?", (student_row_id,))
        if student is None:
            raise ValueError("Student not found.")
        self.db.execute("DELETE FROM students WHERE id=?", (student_row_id,))
        self.log_audit("delete_student", f"Deleted student '{student['student_id']}'.")

    def _save_student_photo(self, student_id, data_bytes, ext):
        os.makedirs(PHOTOS_DIR, exist_ok=True)
        ext = (ext or ".png").lower()
        if ext not in (".jpg", ".jpeg", ".png"):
            raise ValueError("Only JPG, JPEG and PNG images are supported.")
        unique = uuid.uuid4().hex[:8]
        dest_name = f"{student_id}_{unique}{ext}"
        dest_path = os.path.join(PHOTOS_DIR, dest_name)
        with open(dest_path, "wb") as f:
            f.write(data_bytes)
        return dest_path

    # ===================================================================
    # ATTENDANCE
    # ===================================================================
    def get_students_for_class(self, class_id):
        role = self.current_user["role"]
        if role == "teacher" and not self.teacher_owns_class(class_id):
            raise PermissionError_("You are not assigned to this class.")
        return self.db.query(
            "SELECT * FROM students WHERE class_id=? AND status='Active' ORDER BY roll_number",
            (class_id,)
        )

    def load_attendance(self, class_id, att_date):
        role = self.current_user["role"]
        if role == "teacher":
            self.require("attendance_view_assigned")
            if not self.teacher_owns_class(class_id):
                raise PermissionError_("You are not assigned to this class.")
        else:
            self.require("attendance_view")
        rows = self.db.query(
            """SELECT s.id as student_row_id, s.roll_number, s.first_name, s.last_name,
                      a.status
               FROM students s
               LEFT JOIN attendance a ON a.student_id = s.id AND a.attendance_date = ?
               WHERE s.class_id = ? AND s.status='Active'
               ORDER BY s.roll_number""",
            (att_date, class_id)
        )
        return rows

    def save_attendance(self, class_id, att_date, status_by_student_row_id):
        role = self.current_user["role"]
        if role == "teacher":
            self.require("attendance_edit_assigned")
            if not self.teacher_owns_class(class_id):
                raise PermissionError_("You are not assigned to this class.")
        else:
            self.require("attendance_edit")

        for student_row_id, status in status_by_student_row_id.items():
            if status not in ATTENDANCE_STATUSES:
                continue
            self.db.execute(
                """INSERT INTO attendance (student_id, attendance_date, status)
                   VALUES (?, ?, ?)
                   ON CONFLICT(student_id, attendance_date)
                   DO UPDATE SET status=excluded.status""",
                (student_row_id, att_date, status)
            )
        self.log_audit("save_attendance", f"Saved attendance for class id={class_id} on {att_date}.")

    def attendance_report(self, class_id=None, student_row_id=None, date_from=None, date_to=None):
        role = self.current_user["role"]
        if role == "teacher":
            self.require("reports_view_assigned")
        else:
            self.require("reports_view") if role != "staff" else self.require("reports_view_limited")

        base = """SELECT s.id as student_row_id, s.student_id, s.roll_number,
                          s.first_name, s.last_name, s.class_id,
                          SUM(CASE WHEN a.status='Present' THEN 1 ELSE 0 END) as present_days,
                          SUM(CASE WHEN a.status='Absent' THEN 1 ELSE 0 END) as absent_days,
                          SUM(CASE WHEN a.status='Leave' THEN 1 ELSE 0 END) as leave_days,
                          COUNT(a.id) as total_days
                   FROM students s
                   LEFT JOIN attendance a ON a.student_id = s.id"""
        conditions = []
        params = []

        if date_from:
            conditions.append("(a.attendance_date IS NULL OR a.attendance_date >= ?)")
            params.append(date_from)
        if date_to:
            conditions.append("(a.attendance_date IS NULL OR a.attendance_date <= ?)")
            params.append(date_to)

        if role == "teacher":
            conditions.append(
                "s.class_id IN (SELECT class_id FROM teacher_classes WHERE teacher_id=?)"
            )
            params.append(self.current_user["id"])

        if class_id:
            conditions.append("s.class_id = ?")
            params.append(class_id)
        if student_row_id:
            conditions.append("s.id = ?")
            params.append(student_row_id)

        if conditions:
            base += " WHERE " + " AND ".join(conditions)
        base += " GROUP BY s.id ORDER BY s.first_name"

        rows = self.db.query(base, params)
        results = []
        for r in rows:
            total = r["total_days"] or 0
            present = r["present_days"] or 0
            pct = (present / total * 100.0) if total > 0 else 0.0
            results.append({
                "student_row_id": r["student_row_id"],
                "student_id": r["student_id"],
                "roll_number": r["roll_number"],
                "name": f"{r['first_name']} {r['last_name'] or ''}".strip(),
                "total_days": total,
                "present": present,
                "absent": r["absent_days"] or 0,
                "leave": r["leave_days"] or 0,
                "percentage": round(pct, 2),
            })
        return results

    # ===================================================================
    # MARKS
    # ===================================================================
    def list_marks(self, student_row_id):
        role = self.current_user["role"]
        if role == "teacher":
            self.require("marks_view_assigned")
            if not self.teacher_owns_student(student_row_id):
                raise PermissionError_("You cannot view marks outside your assigned class(es).")
        else:
            self.require("marks_view")
        return self.db.query(
            "SELECT * FROM marks WHERE student_id=? ORDER BY exam_name, subject",
            (student_row_id,)
        )

    def save_mark(self, student_row_id, subject, exam_name, total_marks, obtained_marks, mark_id=None):
        role = self.current_user["role"]
        if role == "teacher":
            self.require("marks_edit_assigned")
            if not self.teacher_owns_student(student_row_id):
                raise PermissionError_("You cannot edit marks outside your assigned class(es).")
        else:
            self.require("marks_edit")

        total_marks = float(total_marks)
        obtained_marks = float(obtained_marks)
        if total_marks <= 0:
            raise ValueError("Total marks must be greater than 0.")
        if obtained_marks < 0:
            raise ValueError("Obtained marks cannot be negative.")
        if obtained_marks > total_marks:
            raise ValueError("Obtained marks cannot exceed total marks.")

        now = datetime.now().isoformat(timespec="seconds")
        if mark_id:
            self.db.execute(
                """UPDATE marks SET subject=?, exam_name=?, total_marks=?,
                   obtained_marks=?, updated_at=? WHERE id=?""",
                (subject, exam_name, total_marks, obtained_marks, now, mark_id)
            )
        else:
            self.db.execute(
                """INSERT INTO marks (student_id, subject, exam_name, total_marks,
                   obtained_marks, created_at, updated_at) VALUES (?, ?, ?, ?, ?, ?, ?)""",
                (student_row_id, subject, exam_name, total_marks, obtained_marks, now, now)
            )
        self.log_audit("save_marks", f"Saved marks for student id={student_row_id}, subject={subject}.")

    def delete_mark(self, mark_id, student_row_id):
        role = self.current_user["role"]
        if role == "teacher":
            self.require("marks_edit_assigned")
            if not self.teacher_owns_student(student_row_id):
                raise PermissionError_("You cannot delete marks outside your assigned class(es).")
        else:
            self.require("marks_edit")
        self.db.execute("DELETE FROM marks WHERE id=?", (mark_id,))
        self.log_audit("delete_marks", f"Deleted mark id={mark_id}.")

    def result_summary(self, student_row_id, exam_name=None):
        """Returns list of mark rows plus totals/percentage/grade/pass-fail
        for one student (optionally for a single exam)."""
        marks = self.list_marks(student_row_id)
        if exam_name:
            marks = [m for m in marks if m["exam_name"] == exam_name]
        total = sum(m["total_marks"] for m in marks)
        obtained = sum(m["obtained_marks"] for m in marks)
        pct = (obtained / total * 100.0) if total > 0 else 0.0
        grade = calculate_grade(pct)
        result = "PASS" if pct >= PASS_PERCENTAGE else "FAIL"
        return {
            "marks": marks, "total": total, "obtained": obtained,
            "percentage": round(pct, 2), "grade": grade, "result": result,
        }

    # ===================================================================
    # DASHBOARD STATISTICS
    # ===================================================================
    def dashboard_stats(self):
        role = self.current_user["role"]
        stats = {}
        if role == "admin":
            stats["Total Students"] = self.db.query_one("SELECT COUNT(*) c FROM students")["c"]
            stats["Total Teachers"] = self.db.query_one("SELECT COUNT(*) c FROM users WHERE role='teacher'")["c"]
            stats["Total Staff"] = self.db.query_one("SELECT COUNT(*) c FROM users WHERE role='staff'")["c"]
            stats["Total Management"] = self.db.query_one("SELECT COUNT(*) c FROM users WHERE role='management'")["c"]
            stats["Total Classes"] = self.db.query_one("SELECT COUNT(*) c FROM classes")["c"]
            stats["Active Users"] = self.db.query_one("SELECT COUNT(*) c FROM users WHERE is_active=1")["c"]
        elif role == "management":
            stats["Total Students"] = self.db.query_one("SELECT COUNT(*) c FROM students")["c"]
            stats["Total Teachers"] = self.db.query_one("SELECT COUNT(*) c FROM users WHERE role='teacher'")["c"]
            stats["Total Classes"] = self.db.query_one("SELECT COUNT(*) c FROM classes")["c"]
        elif role == "teacher":
            stats["Assigned Classes"] = self.db.query_one(
                "SELECT COUNT(*) c FROM teacher_classes WHERE teacher_id=?",
                (self.current_user["id"],)
            )["c"]
            stats["My Students"] = self.db.query_one(
                """SELECT COUNT(*) c FROM students
                   WHERE class_id IN (SELECT class_id FROM teacher_classes WHERE teacher_id=?)""",
                (self.current_user["id"],)
            )["c"]
            today = date.today().isoformat()
            stats["Today's Attendance Marked"] = self.db.query_one(
                """SELECT COUNT(*) c FROM attendance a
                   JOIN students s ON s.id = a.student_id
                   WHERE a.attendance_date=? AND s.class_id IN
                   (SELECT class_id FROM teacher_classes WHERE teacher_id=?)""",
                (today, self.current_user["id"])
            )["c"]
        elif role == "staff":
            stats["Total Students"] = self.db.query_one("SELECT COUNT(*) c FROM students")["c"]
            stats["Total Classes"] = self.db.query_one("SELECT COUNT(*) c FROM classes")["c"]
        return stats


# =====================================================================
# EXCEL EXPORT HELPERS -- build entirely in memory, return bytes so the
# caller can hand them to st.download_button (no server-side filedialog
# exists on Streamlit Cloud).
# =====================================================================
def _style_excel_header(ws, num_cols):
    header_font = Font(bold=True, color="FFFFFF")
    header_fill = PatternFill(start_color="2563EB", end_color="2563EB", fill_type="solid")
    for col in range(1, num_cols + 1):
        cell = ws.cell(row=1, column=col)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal="center")
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions


def export_students_to_excel(students):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Students"
    headers = ["Student ID", "Roll No", "Name", "Father Name", "Class", "Phone", "Status"]
    ws.append(headers)
    for s in students:
        class_label = f"{s['class_name']}-{s['section']}" if s["class_name"] and s["section"] else (s["class_name"] or "")
        ws.append([
            s["student_id"], s["roll_number"] or "",
            f"{s['first_name']} {s['last_name'] or ''}".strip(), s["father_name"] or "",
            class_label, s["phone"] or "", s["status"]
        ])
    _style_excel_header(ws, len(headers))
    for i, col in enumerate(headers, start=1):
        ws.column_dimensions[get_column_letter(i)].width = max(14, len(col) + 4)
    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf.getvalue()


def export_attendance_report_to_excel(report_rows):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Attendance Report"
    headers = ["Student ID", "Roll No", "Name", "Total Days", "Present", "Absent", "Leave", "Attendance %"]
    ws.append(headers)
    for r in report_rows:
        ws.append([r["student_id"], r["roll_number"] or "", r["name"], r["total_days"],
                   r["present"], r["absent"], r["leave"], r["percentage"]])
    _style_excel_header(ws, len(headers))
    for i, col in enumerate(headers, start=1):
        ws.column_dimensions[get_column_letter(i)].width = max(14, len(col) + 4)
    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf.getvalue()


# =====================================================================
# PDF RESULT CARD (ReportLab) -- returns bytes instead of a file path.
# =====================================================================
def generate_result_card_pdf(student_row, result, school_name):
    buf = io.BytesIO()
    doc = SimpleDocTemplate(buf, pagesize=A4, topMargin=25 * mm, bottomMargin=20 * mm)
    styles = getSampleStyleSheet()
    title_style = ParagraphStyle("TitleCenter", parent=styles["Title"], alignment=TA_CENTER, fontSize=16)
    sub_style = ParagraphStyle("SubCenter", parent=styles["Normal"], alignment=TA_CENTER, fontSize=11)

    elements = [Paragraph(school_name, title_style), Paragraph("Result Card", sub_style), Spacer(1, 10)]

    if student_row["photo_path"] and os.path.exists(student_row["photo_path"]):
        try:
            elements.append(RLImage(student_row["photo_path"], width=30 * mm, height=30 * mm))
            elements.append(Spacer(1, 8))
        except Exception:
            pass

    info_data = [
        ["Student Name:", f"{student_row['first_name']} {student_row['last_name'] or ''}".strip(),
         "Student ID:", student_row["student_id"]],
        ["Father Name:", student_row["father_name"] or "-", "Roll Number:", student_row["roll_number"] or "-"],
        ["Class:", (f"{student_row['class_name']}-{student_row['section']}"
                    if student_row["class_name"] and student_row["section"] else (student_row["class_name"] or "-")),
         "Academic Year:", ""],
    ]
    info_table = Table(info_data, colWidths=[80, 150, 90, 150])
    info_table.setStyle(TableStyle([
        ("FONTNAME", (0, 0), (-1, -1), "Helvetica"),
        ("FONTSIZE", (0, 0), (-1, -1), 9),
        ("FONTNAME", (0, 0), (0, -1), "Helvetica-Bold"),
        ("FONTNAME", (2, 0), (2, -1), "Helvetica-Bold"),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 6),
    ]))
    elements.append(info_table)
    elements.append(Spacer(1, 15))

    marks_data = [["Subject", "Total Marks", "Obtained Marks", "Percentage", "Grade"]]
    for m in result["marks"]:
        pct = round(m["obtained_marks"] / m["total_marks"] * 100, 1) if m["total_marks"] else 0
        marks_data.append([m["subject"], str(m["total_marks"]), str(m["obtained_marks"]), f"{pct}%",
                            calculate_grade(pct)])
    marks_table = Table(marks_data, colWidths=[140, 90, 100, 80, 60])
    marks_table.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#2563EB")),
        ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
        ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
        ("GRID", (0, 0), (-1, -1), 0.5, colors.grey),
        ("ALIGN", (1, 0), (-1, -1), "CENTER"),
        ("FONTSIZE", (0, 0), (-1, -1), 9),
    ]))
    elements.append(marks_table)
    elements.append(Spacer(1, 15))

    summary_data = [
        ["Total Marks", str(result["total"])],
        ["Obtained Marks", str(result["obtained"])],
        ["Overall Percentage", f"{result['percentage']}%"],
        ["Overall Grade", result["grade"]],
        ["Result", result["result"]],
    ]
    summary_table = Table(summary_data, colWidths=[150, 150])
    summary_table.setStyle(TableStyle([
        ("FONTNAME", (0, 0), (0, -1), "Helvetica-Bold"),
        ("FONTSIZE", (0, 0), (-1, -1), 10),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 6),
    ]))
    elements.append(summary_table)
    elements.append(Spacer(1, 40))

    sig_data = [["______________________", "______________________"],
                ["Class Teacher", "Principal / Admin"]]
    sig_table = Table(sig_data, colWidths=[200, 200])
    sig_table.setStyle(TableStyle([("ALIGN", (0, 0), (-1, -1), "CENTER"), ("FONTSIZE", (0, 0), (-1, -1), 9)]))
    elements.append(sig_table)

    doc.build(elements)
    buf.seek(0)
    return buf.getvalue()


# =====================================================================
# 7. STREAMLIT APP - RESOURCE / STATE SETUP
# =====================================================================
@st.cache_resource(show_spinner=False)
def get_db() -> DatabaseManager:
    """One shared SQLite connection for the whole running app instance."""
    return DatabaseManager()


def get_session() -> Session:
    """One Session (holds the logged-in user) per browser tab / visitor."""
    if "session" not in st.session_state:
        st.session_state.session = Session(get_db())
    return st.session_state.session


def class_choices(session):
    """Returns (labels_list, {label: class_id}) for select boxes."""
    labels, mapping = [], {}
    for c in session.list_classes():
        label = f"{c['class_name']}-{c['section']}" if c["section"] else c["class_name"]
        labels.append(label)
        mapping[label] = c["id"]
    return labels, mapping


# =====================================================================
# 8. ACTIVATION SCREEN
# =====================================================================
def render_activation(db: DatabaseManager):
    st.title(APP_NAME)
    st.subheader("Activation Required")
    st.caption("(Basic local activation check - not a secure commercial license)")
    with st.form("activation_form"):
        key = st.text_input("Activation Key", placeholder="RJK-XXX-XXX-XXXX")
        submitted = st.form_submit_button("Activate", type="primary")
    if submitted:
        if key.strip() == ACTIVATION_KEY:
            db.set_setting("activated", "1")
            st.success("Software activated successfully!")
            st.rerun()
        else:
            st.error("Invalid activation key.")


# =====================================================================
# 9. LOGIN SCREEN
# =====================================================================
def render_login(session: Session):
    st.title(APP_NAME)
    st.subheader("Login")
    with st.expander("First-run demo credentials"):
        st.markdown(
            "- **Admin:** `admin` / `admin123`\n"
            "- **Management:** `management` / `management123`\n"
            "- **Teacher:** `teacher` / `teacher123`\n"
            "- **Staff:** `staff` / `staff123`\n\n"
            "Please change these after your first login."
        )
    with st.form("login_form"):
        username = st.text_input("Username")
        password = st.text_input("Password", type="password")
        submitted = st.form_submit_button("Login", type="primary")
    if submitted:
        user, err = session.login(username.strip(), password)
        if err:
            st.error(err)
        else:
            st.success(f"Welcome, {user['full_name']}!")
            st.rerun()


# =====================================================================
# 10. SIDEBAR NAVIGATION
# =====================================================================
NAV_BY_ROLE = {
    "admin": [
        ("Dashboard", "dashboard"), ("Students", "students"), ("Classes", "classes"),
        ("Teachers", "teachers"), ("Attendance", "attendance"), ("Marks", "marks"),
        ("Reports", "reports"), ("Users", "users"), ("Audit Log", "audit"),
    ],
    "management": [
        ("Dashboard", "dashboard"), ("Students", "students"), ("Teachers", "teachers"),
        ("Classes", "classes"), ("Attendance", "attendance"), ("Marks", "marks"),
        ("Reports", "reports"),
    ],
    "teacher": [
        ("Dashboard", "dashboard"), ("My Classes", "classes"), ("My Students", "students"),
        ("Attendance", "attendance"), ("Marks", "marks"), ("Reports", "reports"),
    ],
    "staff": [
        ("Dashboard", "dashboard"), ("Students", "students"), ("Reports", "reports"),
    ],
}


def render_sidebar(session: Session):
    user = session.current_user
    with st.sidebar:
        st.markdown(f"### {APP_NAME}")
        st.caption(f"{user['full_name']}  ({user['role'].capitalize()})")
        st.divider()
        nav_items = NAV_BY_ROLE.get(user["role"], [])
        labels = [label for label, _ in nav_items]
        keys = [key for _, key in nav_items]
        if "nav_key" not in st.session_state or st.session_state.nav_key not in keys:
            st.session_state.nav_key = keys[0]
        choice = st.radio("Navigate", labels, index=keys.index(st.session_state.nav_key),
                           label_visibility="collapsed")
        st.session_state.nav_key = keys[labels.index(choice)]

        st.divider()
        with st.expander("Change Password"):
            with st.form("change_pw_form", clear_on_submit=True):
                old = st.text_input("Current Password", type="password")
                new = st.text_input("New Password", type="password")
                confirm = st.text_input("Confirm New Password", type="password")
                if st.form_submit_button("Update Password"):
                    if new != confirm:
                        st.error("New passwords do not match.")
                    elif len(new) < 4:
                        st.error("New password must be at least 4 characters.")
                    else:
                        try:
                            session.change_own_password(old, new)
                            st.success("Password changed successfully.")
                        except Exception as exc:
                            st.error(str(exc))

        if st.button("Logout", use_container_width=True):
            session.logout()
            for k in ("nav_key",):
                st.session_state.pop(k, None)
            st.rerun()

    return st.session_state.nav_key


# =====================================================================
# 11. DASHBOARD VIEW
# =====================================================================
def view_dashboard(session: Session):
    user = session.current_user
    st.header(f"Welcome, {user['full_name']}!")
    stats = session.dashboard_stats()
    cols = st.columns(len(stats) if stats else 1)
    for col, (label, value) in zip(cols, stats.items()):
        col.metric(label, value)

    missing = []
    if not OPENPYXL_AVAILABLE:
        missing.append("openpyxl (Excel export)")
    if not REPORTLAB_AVAILABLE:
        missing.append("reportlab (PDF result cards)")
    if missing:
        st.warning(
            "Optional features unavailable - missing packages: " + ", ".join(missing) +
            "\n\nInstall with: `pip install " + " ".join(m.split()[0] for m in missing) + "`"
        )


# =====================================================================
# 12. CLASSES VIEW
# =====================================================================
def view_classes(session: Session):
    user = session.current_user
    st.header("Classes")
    can_manage = has_permission(user["role"], "manage_classes_add")

    classes = session.list_classes()
    rows = [{
        "ID": c["id"], "Class Name": c["class_name"], "Section": c["section"] or "",
        "Academic Year": c["academic_year"] or "", "Status": "Active" if c["is_active"] else "Inactive",
    } for c in classes]
    st.dataframe(rows, use_container_width=True, hide_index=True)

    if can_manage:
        with st.expander("Add Class"):
            with st.form("add_class_form", clear_on_submit=True):
                cname = st.text_input("Class Name")
                section = st.text_input("Section")
                year = st.text_input("Academic Year")
                if st.form_submit_button("Add Class"):
                    if not cname.strip():
                        st.error("Class Name is required.")
                    else:
                        try:
                            session.add_class(cname.strip(), section.strip(), year.strip())
                            st.success("Class added.")
                            st.rerun()
                        except Exception as exc:
                            st.error(str(exc))

        if classes:
            with st.expander("Edit / Delete Class"):
                labels = {f"{c['class_name']}-{c['section']} (ID {c['id']})" if c["section"]
                          else f"{c['class_name']} (ID {c['id']})": c["id"] for c in classes}
                pick = st.selectbox("Select a class", list(labels.keys()), key="class_pick")
                cid = labels[pick]
                c = session.db.query_one("SELECT * FROM classes WHERE id=?", (cid,))
                with st.form("edit_class_form"):
                    cname = st.text_input("Class Name", value=c["class_name"])
                    section = st.text_input("Section", value=c["section"] or "")
                    year = st.text_input("Academic Year", value=c["academic_year"] or "")
                    status = st.selectbox("Status", ["Active", "Inactive"],
                                           index=0 if c["is_active"] else 1)
                    col1, col2 = st.columns(2)
                    save = col1.form_submit_button("Save Changes")
                    delete = col2.form_submit_button("Delete Class", type="secondary")
                if save:
                    try:
                        session.edit_class(cid, cname.strip(), section.strip(), year.strip(),
                                            status == "Active")
                        st.success("Class updated.")
                        st.rerun()
                    except Exception as exc:
                        st.error(str(exc))
                if delete:
                    try:
                        session.delete_class(cid)
                        st.success("Class deleted.")
                        st.rerun()
                    except Exception as exc:
                        st.error(str(exc))

    if user["role"] in ("admin", "management"):
        render_teacher_assignment(session)


def render_teacher_assignment(session: Session):
    with st.expander("Assign Teacher to a Class"):
        teachers = session.list_teachers()
        if not teachers:
            st.info("No teacher accounts exist yet.")
            return
        teacher_map = {t["full_name"]: t["id"] for t in teachers}
        teacher_name = st.selectbox("Teacher", list(teacher_map.keys()), key="assign_teacher_pick")
        teacher_id = teacher_map[teacher_name]

        all_classes = session.db.query("SELECT * FROM classes ORDER BY class_name, section")
        assigned_ids = {c["id"] for c in session.list_assignments_for_teacher(teacher_id)}
        class_map = {(f"{c['class_name']}-{c['section']}" if c["section"] else c["class_name"]): c["id"]
                     for c in all_classes}

        unassigned_labels = [lbl for lbl, cid in class_map.items() if cid not in assigned_ids]
        assigned_labels = [lbl for lbl, cid in class_map.items() if cid in assigned_ids]

        col1, col2 = st.columns(2)
        with col1:
            st.caption("Not yet assigned")
            pick_assign = st.selectbox("Class to assign", unassigned_labels or ["(none)"], key="pick_assign")
            if st.button("Assign ->", disabled=not unassigned_labels):
                try:
                    session.assign_teacher(teacher_id, class_map[pick_assign])
                    st.success(f"Assigned {pick_assign} to {teacher_name}.")
                    st.rerun()
                except Exception as exc:
                    st.error(str(exc))
        with col2:
            st.caption("Currently assigned")
            pick_remove = st.selectbox("Class to remove", assigned_labels or ["(none)"], key="pick_remove")
            if st.button("<- Remove", disabled=not assigned_labels):
                try:
                    session.remove_teacher_assignment(teacher_id, class_map[pick_remove])
                    st.success(f"Removed {pick_remove} from {teacher_name}.")
                    st.rerun()
                except Exception as exc:
                    st.error(str(exc))


# =====================================================================
# 13. TEACHERS VIEW
# =====================================================================
def view_teachers(session: Session):
    user = session.current_user
    st.header("Teachers")
    teachers = session.list_teachers()
    rows = []
    for t in teachers:
        classes = session.list_assignments_for_teacher(t["id"])
        class_str = ", ".join(f"{c['class_name']}-{c['section']}" if c["section"] else c["class_name"]
                               for c in classes)
        rows.append({
            "ID": t["id"], "Full Name": t["full_name"], "Username": t["username"],
            "Email": t["email"] or "", "Phone": t["phone"] or "",
            "Assigned Classes": class_str, "Status": "Active" if t["is_active"] else "Disabled",
        })
    st.dataframe(rows, use_container_width=True, hide_index=True)

    can_add = user["role"] == "admin" or has_permission(user["role"], "manage_users_add_teacher")
    if can_add:
        with st.expander("Add Teacher"):
            with st.form("add_teacher_form", clear_on_submit=True):
                username = st.text_input("Username")
                password = st.text_input("Password", type="password")
                full_name = st.text_input("Full Name")
                email = st.text_input("Email")
                phone = st.text_input("Phone")
                if st.form_submit_button("Add Teacher"):
                    if not username.strip() or not password or not full_name.strip():
                        st.error("Username, Password and Full Name are required.")
                    else:
                        try:
                            session.add_user(username.strip(), password, full_name.strip(),
                                              "teacher", email.strip(), phone.strip())
                            st.success("Teacher account created.")
                            st.rerun()
                        except Exception as exc:
                            st.error(str(exc))

    if user["role"] in ("admin", "management"):
        render_teacher_assignment(session)


# =====================================================================
# 14. USERS VIEW (admin only)
# =====================================================================
def view_users(session: Session):
    user = session.current_user
    if not has_permission(user["role"], "manage_users_add"):
        st.error("Only Admin can manage all users.")
        return
    st.header("System Users")

    keyword = st.text_input("Search by username or full name", key="user_search")
    users = session.list_users()
    if keyword:
        kw = keyword.lower()
        users = [u for u in users if kw in u["username"].lower() or kw in u["full_name"].lower()]
    rows = [{
        "ID": u["id"], "Username": u["username"], "Full Name": u["full_name"], "Role": u["role"],
        "Email": u["email"] or "", "Phone": u["phone"] or "",
        "Status": "Active" if u["is_active"] else "Disabled", "Created": u["created_at"],
    } for u in users]
    st.dataframe(rows, use_container_width=True, hide_index=True)

    with st.expander("Add User"):
        with st.form("add_user_form", clear_on_submit=True):
            username = st.text_input("Username")
            password = st.text_input("Password", type="password")
            full_name = st.text_input("Full Name")
            role = st.selectbox("Role", ROLES)
            email = st.text_input("Email")
            phone = st.text_input("Phone")
            if st.form_submit_button("Add User"):
                if not username.strip() or not password or not full_name.strip():
                    st.error("Username, Password and Full Name are required.")
                else:
                    try:
                        session.add_user(username.strip(), password, full_name.strip(), role,
                                          email.strip(), phone.strip())
                        st.success("User created.")
                        st.rerun()
                    except Exception as exc:
                        st.error(str(exc))

    if users:
        with st.expander("Manage Existing User"):
            labels = {f"{u['username']} - {u['full_name']} (ID {u['id']})": u["id"] for u in users}
            pick = st.selectbox("Select a user", list(labels.keys()), key="user_pick")
            uid = labels[pick]
            u = session.db.query_one("SELECT * FROM users WHERE id=?", (uid,))

            with st.form("edit_user_form"):
                full_name = st.text_input("Full Name", value=u["full_name"])
                role = st.selectbox("Role", ROLES, index=ROLES.index(u["role"]) if u["role"] in ROLES else 0)
                email = st.text_input("Email", value=u["email"] or "")
                phone = st.text_input("Phone", value=u["phone"] or "")
                if st.form_submit_button("Save Changes"):
                    try:
                        session.edit_user(uid, full_name.strip(), email.strip(), phone.strip(), role)
                        st.success("User updated.")
                        st.rerun()
                    except Exception as exc:
                        st.error(str(exc))

            col1, col2, col3, col4 = st.columns(4)
            if col1.button("Enable", key="enable_user"):
                try:
                    session.set_user_active(uid, True)
                    st.success("User enabled.")
                    st.rerun()
                except Exception as exc:
                    st.error(str(exc))
            if col2.button("Disable", key="disable_user"):
                try:
                    session.set_user_active(uid, False)
                    st.success("User disabled.")
                    st.rerun()
                except Exception as exc:
                    st.error(str(exc))
            if col3.button("Delete User", key="delete_user"):
                try:
                    session.delete_user(uid)
                    st.success("User deleted.")
                    st.rerun()
                except Exception as exc:
                    st.error(str(exc))
            with col4.popover("Reset Password"):
                with st.form("reset_pw_form", clear_on_submit=True):
                    new_pw = st.text_input("New Password", type="password")
                    if st.form_submit_button("Reset"):
                        if not new_pw:
                            st.error("Enter a new password.")
                        else:
                            try:
                                session.reset_password(uid, new_pw)
                                st.success("Password reset.")
                            except Exception as exc:
                                st.error(str(exc))


# =====================================================================
# 15. STUDENTS VIEW
# =====================================================================
def _student_row_to_dict(s):
    class_label = f"{s['class_name']}-{s['section']}" if s["class_name"] and s["section"] else (s["class_name"] or "")
    return {
        "ID": s["id"], "Student ID": s["student_id"], "Roll No": s["roll_number"] or "",
        "Name": f"{s['first_name']} {s['last_name'] or ''}".strip(),
        "Father Name": s["father_name"] or "", "Class": class_label,
        "Phone": s["phone"] or "", "Status": s["status"],
    }


def view_students(session: Session):
    user = session.current_user
    st.header("Students")

    can_add = has_permission(user["role"], "students_add") or has_permission(user["role"], "students_add_assigned")
    can_edit = has_permission(user["role"], "students_edit") or has_permission(user["role"], "students_edit_assigned")
    can_delete = has_permission(user["role"], "students_delete")
    can_export = OPENPYXL_AVAILABLE and (has_permission(user["role"], "export_excel")
                                          or has_permission(user["role"], "export_excel_assigned"))

    col1, col2, col3 = st.columns(3)
    keyword = col1.text_input("Search", key="student_search")
    labels, class_map = class_choices(session)
    class_choice = col2.selectbox("Class", ["All"] + labels, key="student_class_filter")
    status_choice = col3.selectbox("Status", ["All", "Active", "Inactive"], key="student_status_filter")

    class_id = class_map.get(class_choice) if class_choice != "All" else None
    status = None if status_choice == "All" else status_choice

    try:
        students = session.search_students(keyword.strip(), class_id, status)
    except Exception as exc:
        show_exception("loading students", exc)
        students = []

    st.dataframe([_student_row_to_dict(s) for s in students], use_container_width=True, hide_index=True)

    if can_export and students:
        st.download_button(
            "Export Excel", data=export_students_to_excel(students),
            file_name=f"Students_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
    elif has_permission(user["role"], "export_excel") or has_permission(user["role"], "export_excel_assigned"):
        if not OPENPYXL_AVAILABLE:
            st.caption("Install `openpyxl` to enable Excel export.")

    if can_add:
        with st.expander("Add Student"):
            _student_form(session, mode="add")

    if students and can_edit:
        with st.expander("Edit Student"):
            labels_map = {f"{s['student_id']} - {s['first_name']} {s['last_name'] or ''}".strip(): s["id"]
                          for s in students}
            pick = st.selectbox("Select a student", list(labels_map.keys()), key="edit_student_pick")
            _student_form(session, mode="edit", student_row_id=labels_map[pick])

    if students:
        with st.expander("View Details / Result Card"):
            labels_map = {f"{s['student_id']} - {s['first_name']} {s['last_name'] or ''}".strip(): s["id"]
                          for s in students}
            pick = st.selectbox("Select a student", list(labels_map.keys()), key="detail_student_pick")
            render_student_detail(session, labels_map[pick])

    if students and can_delete:
        with st.expander("Delete Student"):
            labels_map = {f"{s['student_id']} - {s['first_name']} {s['last_name'] or ''}".strip(): s["id"]
                          for s in students}
            pick = st.selectbox("Select a student to delete", list(labels_map.keys()), key="delete_student_pick")
            confirm = st.checkbox("I understand this permanently deletes the student.", key="delete_student_confirm")
            if st.button("Delete Student", type="secondary", disabled=not confirm):
                try:
                    session.delete_student(labels_map[pick])
                    st.success("Student deleted.")
                    st.rerun()
                except Exception as exc:
                    st.error(str(exc))


def _student_form(session: Session, mode="add", student_row_id=None):
    labels, mapping = class_choices(session)
    if not labels:
        st.warning("Please add/be assigned a class before adding students.")
        return

    initial = {}
    if mode == "edit" and student_row_id:
        try:
            s = session.get_student(student_row_id)
        except Exception as exc:
            st.error(str(exc))
            return
        current_label = next((lbl for lbl, cid in mapping.items() if cid == s["class_id"]), labels[0])
        initial = {
            "student_id": s["student_id"], "roll_number": s["roll_number"] or "",
            "first_name": s["first_name"], "last_name": s["last_name"] or "",
            "father_name": s["father_name"] or "",
            "date_of_birth": s["date_of_birth"] or date.today().isoformat(),
            "gender": s["gender"] or "Male", "phone": s["phone"] or "", "address": s["address"] or "",
            "class_label": current_label,
            "admission_date": s["admission_date"] or date.today().isoformat(),
            "status": s["status"],
        }
    else:
        initial = {
            "student_id": "", "roll_number": "", "first_name": "", "last_name": "",
            "father_name": "", "date_of_birth": date.today().isoformat(), "gender": "Male",
            "phone": "", "address": "", "class_label": labels[0],
            "admission_date": date.today().isoformat(), "status": "Active",
        }

    form_key = f"student_form_{mode}_{student_row_id or 'new'}"
    with st.form(form_key):
        c1, c2 = st.columns(2)
        student_id = c1.text_input("Student ID *", value=initial["student_id"])
        roll_number = c2.text_input("Roll Number", value=initial["roll_number"])
        first_name = c1.text_input("First Name *", value=initial["first_name"])
        last_name = c2.text_input("Last Name", value=initial["last_name"])
        father_name = c1.text_input("Father Name", value=initial["father_name"])
        gender = c2.selectbox("Gender", ["Male", "Female", "Other"],
                               index=["Male", "Female", "Other"].index(initial["gender"])
                               if initial["gender"] in ["Male", "Female", "Other"] else 0)
        dob = c1.text_input("Date of Birth (YYYY-MM-DD)", value=initial["date_of_birth"])
        admission_date = c2.text_input("Admission Date (YYYY-MM-DD)", value=initial["admission_date"])
        phone = c1.text_input("Phone", value=initial["phone"])
        status = c2.selectbox("Status", ["Active", "Inactive"],
                               index=["Active", "Inactive"].index(initial["status"])
                               if initial["status"] in ["Active", "Inactive"] else 0)
        address = st.text_area("Address", value=initial["address"])
        class_label = st.selectbox("Class *", labels, index=labels.index(initial["class_label"])
                                    if initial["class_label"] in labels else 0)
        photo = st.file_uploader("Student Photo (optional)", type=["jpg", "jpeg", "png"], key=f"{form_key}_photo")
        submitted = st.form_submit_button("Save Student", type="primary")

    if submitted:
        if not student_id.strip() or not first_name.strip():
            st.error("Student ID and First Name are required.")
            return
        data = {
            "student_id": student_id.strip(), "roll_number": roll_number.strip(),
            "first_name": first_name.strip(), "last_name": last_name.strip(),
            "father_name": father_name.strip(), "date_of_birth": dob.strip(),
            "gender": gender, "phone": phone.strip(), "address": address.strip(),
            "class_id": mapping.get(class_label), "admission_date": admission_date.strip(),
            "status": status,
        }
        photo_bytes, photo_ext = None, None
        if photo is not None:
            photo_bytes = photo.getvalue()
            photo_ext = os.path.splitext(photo.name)[1].lower()
        try:
            if mode == "add":
                session.add_student(data, photo_bytes, photo_ext)
                st.success("Student added.")
            else:
                session.edit_student(student_row_id, data, photo_bytes, photo_ext)
                st.success("Student updated.")
            st.rerun()
        except Exception as exc:
            st.error(str(exc))


def render_student_detail(session: Session, student_row_id):
    try:
        s = session.get_student(student_row_id)
    except Exception as exc:
        st.error(str(exc))
        return
    if s is None:
        st.error("Student not found.")
        return

    col_photo, col_info = st.columns([1, 3])
    with col_photo:
        if s["photo_path"] and os.path.exists(s["photo_path"]):
            st.image(s["photo_path"], width=140)
        else:
            st.info("No Photo")
    with col_info:
        class_label = (f"{s['class_name']}-{s['section']}" if s["class_name"] and s["section"]
                       else (s["class_name"] or "-"))
        st.markdown(
            f"**Student ID:** {s['student_id']}  \n"
            f"**Roll Number:** {s['roll_number'] or '-'}  \n"
            f"**Name:** {s['first_name']} {s['last_name'] or ''}  \n"
            f"**Father Name:** {s['father_name'] or '-'}  \n"
            f"**Date of Birth:** {s['date_of_birth'] or '-'}  \n"
            f"**Gender:** {s['gender'] or '-'}  \n"
            f"**Phone:** {s['phone'] or '-'}  \n"
            f"**Address:** {s['address'] or '-'}  \n"
            f"**Class:** {class_label}  \n"
            f"**Admission Date:** {s['admission_date'] or '-'}  \n"
            f"**Status:** {s['status']}"
        )

    att = session.attendance_report(student_row_id=student_row_id)
    att_pct = att[0]["percentage"] if att else 0.0
    result = session.result_summary(student_row_id)

    st.markdown("**Academic Summary**")
    c1, c2, c3 = st.columns(3)
    c1.metric("Attendance", f"{att_pct}%")
    c2.metric("Overall Marks", f"{result['obtained']}/{result['total']} ({result['percentage']}%)")
    c3.metric("Grade / Result", f"{result['grade']} / {result['result']}")

    render_result_card(session, s, result, key_prefix=f"detail_{student_row_id}")


def render_result_card(session: Session, s, result, key_prefix="result"):
    school_name = session.db.get_setting("school_name", "RJK School / College")
    with st.expander("Result Card", expanded=False):
        st.markdown(f"#### {school_name}")
        st.caption("Result Card")
        rows = []
        for m in result["marks"]:
            pct = round(m["obtained_marks"] / m["total_marks"] * 100, 1) if m["total_marks"] else 0
            rows.append({
                "Subject": m["subject"], "Total Marks": m["total_marks"],
                "Obtained Marks": m["obtained_marks"], "%": pct, "Grade": calculate_grade(pct),
            })
        st.dataframe(rows, use_container_width=True, hide_index=True)
        st.markdown(
            f"**Total Marks:** {result['total']}  \n"
            f"**Obtained Marks:** {result['obtained']}  \n"
            f"**Overall Percentage:** {result['percentage']}%  \n"
            f"**Overall Grade:** {result['grade']}  \n"
            f"**Result:** {result['result']}"
        )
        if REPORTLAB_AVAILABLE:
            pdf_bytes = generate_result_card_pdf(s, result, school_name)
            st.download_button(
                "Download Result Card (PDF)", data=pdf_bytes,
                file_name=f"Result_{s['student_id']}_{datetime.now().year}.pdf",
                mime="application/pdf", key=f"{key_prefix}_pdf"
            )
        else:
            st.caption("Install `reportlab` to enable PDF result cards.")


# =====================================================================
# 16. ATTENDANCE VIEW
# =====================================================================
def view_attendance(session: Session):
    st.header("Attendance")
    labels, class_map = class_choices(session)
    if not labels:
        st.info("No classes available.")
        return

    col1, col2 = st.columns(2)
    class_label = col1.selectbox("Class", labels, key="att_class")
    att_date = col2.date_input("Date", value=date.today(), key="att_date").isoformat()
    class_id = class_map[class_label]

    try:
        rows = session.load_attendance(class_id, att_date)
    except Exception as exc:
        show_exception("loading attendance", exc)
        return

    if not rows:
        st.info("No active students in this class.")
        return

    table = [{
        "student_row_id": r["student_row_id"],
        "Roll No": r["roll_number"] or "-",
        "Name": f"{r['first_name']} {r['last_name'] or ''}".strip(),
        "Status": r["status"] or "Present",
    } for r in rows]

    col_a, col_b = st.columns(2)
    if col_a.button("Mark All Present"):
        for row in table:
            row["Status"] = "Present"
    if col_b.button("Mark All Absent"):
        for row in table:
            row["Status"] = "Absent"

    edited = st.data_editor(
        table, use_container_width=True, hide_index=True, key=f"att_editor_{class_id}_{att_date}",
        column_config={
            "student_row_id": None,  # hide internal id column
            "Status": st.column_config.SelectboxColumn("Status", options=ATTENDANCE_STATUSES, required=True),
            "Roll No": st.column_config.TextColumn("Roll No", disabled=True),
            "Name": st.column_config.TextColumn("Name", disabled=True),
        },
    )

    if st.button("Save Attendance", type="primary"):
        status_map = {row["student_row_id"]: row["Status"] for row in edited}
        try:
            session.save_attendance(class_id, att_date, status_map)
            st.success("Attendance saved successfully.")
        except Exception as exc:
            st.error(str(exc))


# =====================================================================
# 17. MARKS VIEW
# =====================================================================
def view_marks(session: Session):
    st.header("Marks")
    try:
        students = session.search_students("")
    except Exception:
        students = []
    if not students:
        st.info("No students available.")
        return

    student_map = {f"{s['student_id']} - {s['first_name']} {s['last_name'] or ''}".strip(): s["id"]
                   for s in students}
    pick = st.selectbox("Student", list(student_map.keys()), key="marks_student_pick")
    sid = student_map[pick]

    try:
        result = session.result_summary(sid)
    except Exception as exc:
        show_exception("loading marks", exc)
        return

    rows = []
    for m in result["marks"]:
        pct = round(m["obtained_marks"] / m["total_marks"] * 100, 1) if m["total_marks"] else 0
        rows.append({
            "ID": m["id"], "Exam": m["exam_name"], "Subject": m["subject"],
            "Total": m["total_marks"], "Obtained": m["obtained_marks"],
            "%": pct, "Grade": calculate_grade(pct),
        })
    st.dataframe(rows, use_container_width=True, hide_index=True)
    st.markdown(
        f"**Overall:** {result['obtained']}/{result['total']} ({result['percentage']}%) "
        f"- Grade {result['grade']} - {result['result']}"
    )

    with st.expander("Add / Update Mark"):
        with st.form("add_mark_form", clear_on_submit=True):
            exam_name = st.text_input("Exam Name")
            subject = st.text_input("Subject")
            total_marks = st.text_input("Total Marks")
            obtained_marks = st.text_input("Obtained Marks")
            if st.form_submit_button("Save Mark"):
                try:
                    total = float(total_marks)
                    obtained = float(obtained_marks)
                except ValueError:
                    st.error("Total and Obtained marks must be numbers.")
                else:
                    try:
                        session.save_mark(sid, subject.strip(), exam_name.strip(), total, obtained)
                        st.success("Mark saved.")
                        st.rerun()
                    except Exception as exc:
                        st.error(str(exc))

    if result["marks"]:
        with st.expander("Delete a Mark"):
            mark_map = {f"{m['exam_name']} - {m['subject']} (ID {m['id']})": m["id"] for m in result["marks"]}
            pick_mark = st.selectbox("Select mark entry", list(mark_map.keys()), key="delete_mark_pick")
            if st.button("Delete Mark", type="secondary"):
                try:
                    session.delete_mark(mark_map[pick_mark], sid)
                    st.success("Mark deleted.")
                    st.rerun()
                except Exception as exc:
                    st.error(str(exc))

    s = session.get_student(sid)
    render_result_card(session, s, result, key_prefix=f"marks_{sid}")


# =====================================================================
# 18. REPORTS VIEW
# =====================================================================
def view_reports(session: Session):
    st.header("Attendance Reports")
    labels, class_map = class_choices(session)

    col1, col2, col3 = st.columns(3)
    class_choice = col1.selectbox("Class", ["All"] + labels, key="report_class")
    use_dates = col2.checkbox("Filter by date range", key="report_use_dates")
    date_from = date_to = None
    if use_dates:
        with col3:
            date_from = st.date_input("From", key="report_from").isoformat()
        date_to = st.date_input("To", key="report_to").isoformat()

    class_id = class_map.get(class_choice) if class_choice != "All" else None

    if st.button("View Report", type="primary"):
        try:
            report = session.attendance_report(class_id=class_id, date_from=date_from, date_to=date_to)
        except Exception as exc:
            show_exception("generating report", exc)
            report = []
        st.session_state["_last_report"] = report

    report = st.session_state.get("_last_report", [])
    rows = [{
        "Student ID": r["student_id"], "Roll No": r["roll_number"] or "", "Name": r["name"],
        "Total Days": r["total_days"], "Present": r["present"], "Absent": r["absent"],
        "Leave": r["leave"], "Attendance %": f"{r['percentage']}%",
    } for r in report]
    st.dataframe(rows, use_container_width=True, hide_index=True)

    user_role = session.current_user["role"]
    can_export = OPENPYXL_AVAILABLE and (has_permission(user_role, "export_excel")
                                          or has_permission(user_role, "export_excel_assigned"))
    if can_export and report:
        st.download_button(
            "Export Excel", data=export_attendance_report_to_excel(report),
            file_name=f"Attendance_Report_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )


# =====================================================================
# 19. AUDIT LOG VIEW (admin only)
# =====================================================================
def view_audit(session: Session):
    user = session.current_user
    if not has_permission(user["role"], "view_audit_log"):
        st.error("Only Admin can view the audit log.")
        return
    st.header("Audit Log")
    rows = session.db.query(
        """SELECT al.*, u.username FROM audit_log al
           LEFT JOIN users u ON u.id = al.user_id
           ORDER BY al.id DESC LIMIT 500"""
    )
    table = [{
        "Date/Time": r["timestamp"], "User": r["username"] or "(unknown)",
        "Action": r["action"], "Description": r["description"] or "",
    } for r in rows]
    st.dataframe(table, use_container_width=True, hide_index=True)


# =====================================================================
# 20. APP CONTROLLER / ENTRY POINT
# =====================================================================
VIEW_DISPATCH = {
    "dashboard": view_dashboard,
    "classes": view_classes,
    "teachers": view_teachers,
    "users": view_users,
    "students": view_students,
    "attendance": view_attendance,
    "marks": view_marks,
    "reports": view_reports,
    "audit": view_audit,
}


def main():
    st.set_page_config(page_title=APP_NAME, page_icon="🎓", layout="wide")

    try:
        db = get_db()
    except Exception as exc:
        st.error(f"Failed to initialize the database: {exc}")
        log_error_to_file("initializing the database", exc)
        st.stop()

    session = get_session()

    if db.get_setting("activated", "0") != "1":
        render_activation(db)
        return

    if session.current_user is None:
        render_login(session)
        return

    nav_key = render_sidebar(session)
    view_fn = VIEW_DISPATCH.get(nav_key)
    try:
        if view_fn:
            view_fn(session)
        else:
            st.write("View not implemented.")
    except PermissionError_ as pe:
        st.error(f"Permission Denied: {pe}")
    except Exception as exc:
        show_exception(f"opening view '{nav_key}'", exc)


if __name__ == "__main__":
    main()
