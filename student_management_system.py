"""
=====================================================================
 RJK Student Management System
 Developer/Owner: Rai Jahanzaib
=====================================================================

A complete, offline, single-file Student Management System built with
Python, Tkinter/ttk and SQLite.

FIRST RUN CREDENTIALS (change these after logging in!):
    Admin:      admin / admin123
    Management: management / management123
    Teacher:    teacher / teacher123
    Staff:      staff / staff123

ACTIVATION KEY (basic local check, NOT secure commercial licensing):
    RJK-278-LEF-1185

Run with:
    python student_management_system.py

Requires Python 3.8+. Optional packages (the app still runs without
them, just with photo/PDF/Excel features disabled):
    pip install pillow openpyxl reportlab
=====================================================================
"""

# =====================================================================
# 1. IMPORTS
# =====================================================================
import os
import sys
import shutil
import sqlite3
import hashlib
import binascii
import traceback
import uuid
from datetime import datetime, date

import tkinter as tk
from tkinter import ttk, messagebox, filedialog, simpledialog

# ---- Optional third-party libraries -------------------------------------
try:
    from PIL import Image, ImageTk
    PIL_AVAILABLE = True
except ImportError:
    PIL_AVAILABLE = False

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
# 2. CONFIGURATION / CONSTANTS
# =====================================================================
APP_NAME = "RJK Student Management System"
APP_OWNER = "Rai Jahanzaib"
DB_FILE = "rjk_sms.db"
PHOTOS_DIR = "student_photos"
EXPORTS_DIR = "exports"
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

APP_BG = "#f0f2f5"
SIDEBAR_BG = "#1f2937"
SIDEBAR_FG = "#e5e7eb"
SIDEBAR_ACTIVE_BG = "#374151"
ACCENT_COLOR = "#2563eb"
HEADER_BG = "#111827"


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
    log_error_to_file(context, exc)
    messagebox.showerror(
        "Unexpected Error",
        f"Something went wrong while: {context}\n\n"
        f"Details: {exc}\n\n"
        f"(A full log was saved to {ERROR_LOG_FILE})"
    )


# =====================================================================
# 3 & 4. DATABASE INITIALIZATION + PASSWORD HASHING
# =====================================================================
class DatabaseManager:
    """
    Owns the single SQLite connection and every raw table-creation /
    low level query. Higher-level permission checks live in Session
    (see below) -- this class does NOT know about "current user"; it
    just executes whatever query it is given with parameters, always
    using parameterized SQL (never string concatenation).
    """

    def __init__(self, db_path=DB_FILE):
        self.db_path = db_path
        first_run = not os.path.exists(db_path)
        # check_same_thread=False is fine here because this is a simple
        # single-window desktop app driving one connection from the UI
        # thread only.
        self.conn = sqlite3.connect(db_path, check_same_thread=False)
        self.conn.execute("PRAGMA foreign_keys = ON;")
        self.conn.row_factory = sqlite3.Row
        self._create_tables()
        if first_run:
            self._create_defaults()
        os.makedirs(PHOTOS_DIR, exist_ok=True)
        os.makedirs(EXPORTS_DIR, exist_ok=True)

    # ---------------------------------------------------------------
    def _create_tables(self):
        cur = self.conn.cursor()
        cur.executescript("""
        CREATE TABLE IF NOT EXISTS users (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
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
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            class_name TEXT NOT NULL,
            section TEXT,
            academic_year TEXT,
            is_active INTEGER NOT NULL DEFAULT 1,
            created_at TEXT NOT NULL,
            UNIQUE(class_name, section, academic_year)
        );

        CREATE TABLE IF NOT EXISTS teacher_classes (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            teacher_id INTEGER NOT NULL,
            class_id INTEGER NOT NULL,
            FOREIGN KEY(teacher_id) REFERENCES users(id) ON DELETE CASCADE,
            FOREIGN KEY(class_id) REFERENCES classes(id) ON DELETE CASCADE,
            UNIQUE(teacher_id, class_id)
        );

        CREATE TABLE IF NOT EXISTS students (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
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
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            student_id INTEGER NOT NULL,
            attendance_date TEXT NOT NULL,
            status TEXT NOT NULL CHECK(status IN ('Present','Absent','Leave')),
            FOREIGN KEY(student_id) REFERENCES students(id) ON DELETE CASCADE,
            UNIQUE(student_id, attendance_date)
        );

        CREATE TABLE IF NOT EXISTS marks (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
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
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            user_id INTEGER,
            action TEXT NOT NULL,
            description TEXT,
            timestamp TEXT NOT NULL
        );
        """)
        self.conn.commit()

    # ---------------------------------------------------------------
    def _create_defaults(self):
        now = datetime.now().isoformat(timespec="seconds")
        defaults = [
            ("admin", "admin123", "Administrator", "admin"),
            ("management", "management123", "Management User", "management"),
            ("teacher", "teacher123", "Demo Teacher", "teacher"),
            ("staff", "staff123", "Demo Staff", "staff"),
        ]
        cur = self.conn.cursor()
        for username, password, full_name, role in defaults:
            pwd_hash, salt = hash_password(password)
            cur.execute(
                """INSERT INTO users
                   (username, password_hash, salt, full_name, role, email,
                    phone, is_active, created_at, updated_at)
                   VALUES (?, ?, ?, ?, ?, '', '', 1, ?, ?)""",
                (username, pwd_hash, salt, full_name, role, now, now)
            )
        cur.execute(
            "INSERT OR IGNORE INTO settings (key, value) VALUES (?, ?)",
            ("activated", "0")
        )
        cur.execute(
            "INSERT OR IGNORE INTO settings (key, value) VALUES (?, ?)",
            ("school_name", "RJK School / College")
        )
        cur.execute(
            "INSERT OR IGNORE INTO settings (key, value) VALUES (?, ?)",
            ("school_address", "")
        )
        cur.execute(
            "INSERT OR IGNORE INTO settings (key, value) VALUES (?, ?)",
            ("school_phone", "")
        )
        self.conn.commit()

    # ---------------------------------------------------------------
    # Generic helpers -- ALWAYS parameterized, never string-built SQL.
    def query(self, sql, params=()):
        cur = self.conn.cursor()
        cur.execute(sql, params)
        return cur.fetchall()

    def query_one(self, sql, params=()):
        cur = self.conn.cursor()
        cur.execute(sql, params)
        return cur.fetchone()

    def execute(self, sql, params=()):
        cur = self.conn.cursor()
        cur.execute(sql, params)
        self.conn.commit()
        return cur.lastrowid

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
# 5. CENTRALIZED PERMISSION SYSTEM
# =====================================================================
# Every sensitive action name used throughout the app is listed here,
# per role. This is checked in the BUSINESS LOGIC (Session class
# below) before any database write/read of sensitive data -- never
# only by hiding a button.

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
# 6. SESSION / AUTHENTICATION / BUSINESS LOGIC
# =====================================================================
class Session:
    """
    Holds the DatabaseManager and the currently logged-in user, and
    exposes every "business logic" operation (add student, save
    attendance, etc). Each method re-checks permissions and, for
    teachers, re-checks class ownership DIRECTLY IN SQL -- so even if
    the GUI were bypassed, unauthorized data can never be returned or
    modified.
    """

    def __init__(self, db: DatabaseManager):
        self.db = db
        self.current_user = None  # sqlite3.Row once logged in

    # -----------------------------------------------------------------
    # AUTHENTICATION
    # -----------------------------------------------------------------
    def login(self, username, password):
        user = self.db.query_one(
            "SELECT * FROM users WHERE username = ?", (username,)
        )
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
        row = self.db.query_one(
            """SELECT s.class_id FROM students s WHERE s.id=?""", (student_id,)
        )
        if row is None or row["class_id"] is None:
            return False
        return self.teacher_owns_class(row["class_id"])

    # ===================================================================
    # USERS (admin, + limited management)
    # ===================================================================
    def list_users(self):
        self.require_any(["manage_users_add", "manage_users_add_teacher", "view_dashboard"])
        return self.db.query(
            "SELECT * FROM users ORDER BY role, username"
        )

    def require_any(self, permissions):
        if self.current_user is None:
            raise PermissionError_("Not logged in.")
        if not any(has_permission(self.current_user["role"], p) for p in permissions):
            raise PermissionError_("You do not have permission to view this.")

    def add_user(self, username, password, full_name, role, email, phone):
        # Backend-enforced rule: Management may ONLY create teacher accounts.
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

        now = datetime.now().isoformat(timespec="seconds")
        pwd_hash, salt = hash_password(password)
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
        except sqlite3.IntegrityError:
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
        except sqlite3.IntegrityError:
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
        # admin, management, staff can see all -> no extra filter

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

    def add_student(self, data, photo_source_path=None):
        role = self.current_user["role"]
        if role == "teacher":
            self.require("students_add_assigned")
            if not self.teacher_owns_class(data["class_id"]):
                raise PermissionError_("You can only add students to your own assigned class(es).")
        else:
            self.require("students_add")

        existing = self.db.query_one(
            "SELECT id FROM students WHERE student_id=?", (data["student_id"],)
        )
        if existing:
            raise ValueError(f"Student ID '{data['student_id']}' already exists.")

        photo_path = None
        if photo_source_path:
            photo_path = self._save_student_photo(data["student_id"], photo_source_path)

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

    def edit_student(self, student_row_id, data, photo_source_path=None):
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
        if photo_source_path:
            photo_path = self._save_student_photo(data["student_id"], photo_source_path)

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

    def _save_student_photo(self, student_id, source_path):
        os.makedirs(PHOTOS_DIR, exist_ok=True)
        ext = os.path.splitext(source_path)[1].lower()
        if ext not in (".jpg", ".jpeg", ".png"):
            raise ValueError("Only JPG, JPEG and PNG images are supported.")
        unique = uuid.uuid4().hex[:8]
        dest_name = f"{student_id}_{unique}{ext}"
        dest_path = os.path.join(PHOTOS_DIR, dest_name)
        shutil.copy2(source_path, dest_path)
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
            self.require("marks_view") if role != "staff" else self.require("marks_view")
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
            "marks": marks,
            "total": total,
            "obtained": obtained,
            "percentage": round(pct, 2),
            "grade": grade,
            "result": result,
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
# 7. UTILITY - simple reusable dialog for text/number entry forms
# =====================================================================
class FormDialog(tk.Toplevel):
    """A generic modal form: pass a list of (label, kind, options) tuples.
    kind is one of: 'text', 'combo', 'date', 'password'.
    Result is stored in self.result (dict) or None if cancelled."""

    def __init__(self, parent, title, fields, initial=None):
        super().__init__(parent)
        self.title(title)
        self.resizable(False, False)
        self.result = None
        self.entries = {}
        self.transient(parent)
        self.grab_set()
        initial = initial or {}

        container = ttk.Frame(self, padding=15)
        container.pack(fill="both", expand=True)

        for i, (key, label, kind, options) in enumerate(fields):
            ttk.Label(container, text=label + ":").grid(row=i, column=0, sticky="w", pady=4, padx=(0, 8))
            if kind == "combo":
                var = tk.StringVar(value=str(initial.get(key, options[0] if options else "")))
                widget = ttk.Combobox(container, textvariable=var, values=options, state="readonly", width=27)
            elif kind == "password":
                var = tk.StringVar(value=str(initial.get(key, "")))
                widget = ttk.Entry(container, textvariable=var, show="*", width=30)
            else:
                var = tk.StringVar(value=str(initial.get(key, "")))
                widget = ttk.Entry(container, textvariable=var, width=30)
            widget.grid(row=i, column=1, sticky="ew", pady=4)
            self.entries[key] = var

        btn_frame = ttk.Frame(container)
        btn_frame.grid(row=len(fields), column=0, columnspan=2, pady=(12, 0))
        ttk.Button(btn_frame, text="Save", command=self._on_save).pack(side="left", padx=5)
        ttk.Button(btn_frame, text="Cancel", command=self.destroy).pack(side="left", padx=5)

        self.bind("<Return>", lambda e: self._on_save())
        self.update_idletasks()
        x = parent.winfo_rootx() + 60
        y = parent.winfo_rooty() + 60
        self.geometry(f"+{x}+{y}")

    def _on_save(self):
        self.result = {k: v.get() for k, v in self.entries.items()}
        self.destroy()

    @staticmethod
    def ask(parent, title, fields, initial=None):
        dlg = FormDialog(parent, title, fields, initial)
        parent.wait_window(dlg)
        return dlg.result


def make_scrollable_treeview(parent, columns, headings, widths=None):
    """Creates and returns a ttk.Treeview with vertical+horizontal scrollbars,
    packed inside a frame that is itself returned as `frame`."""
    frame = ttk.Frame(parent)
    tree = ttk.Treeview(frame, columns=columns, show="headings", selectmode="browse")
    for i, col in enumerate(columns):
        tree.heading(col, text=headings[i])
        tree.column(col, width=(widths[i] if widths else 120), anchor="w")
    vsb = ttk.Scrollbar(frame, orient="vertical", command=tree.yview)
    hsb = ttk.Scrollbar(frame, orient="horizontal", command=tree.xview)
    tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
    tree.grid(row=0, column=0, sticky="nsew")
    vsb.grid(row=0, column=1, sticky="ns")
    hsb.grid(row=1, column=0, sticky="ew")
    frame.grid_rowconfigure(0, weight=1)
    frame.grid_columnconfigure(0, weight=1)
    return frame, tree


# =====================================================================
# 8. ACTIVATION WINDOW
# =====================================================================
class ActivationWindow(tk.Toplevel):
    def __init__(self, master, db: DatabaseManager, on_activated):
        super().__init__(master)
        self.db = db
        self.on_activated = on_activated
        self.title(f"{APP_NAME} - Activation")
        self.geometry("460x260")
        self.resizable(False, False)
        self.protocol("WM_DELETE_WINDOW", self._exit_app)
        self.configure(bg=APP_BG)
        self.grab_set()

        ttk.Label(self, text=APP_NAME, font=("Segoe UI", 16, "bold")).pack(pady=(25, 5))
        ttk.Label(self, text="Please enter your activation key to continue.",
                  font=("Segoe UI", 10)).pack(pady=(0, 15))

        self.key_var = tk.StringVar()
        entry = ttk.Entry(self, textvariable=self.key_var, width=30, font=("Segoe UI", 11), justify="center")
        entry.pack(pady=5)
        entry.focus()

        btn_frame = ttk.Frame(self)
        btn_frame.pack(pady=20)
        ttk.Button(btn_frame, text="Activate", command=self._activate).pack(side="left", padx=10)
        ttk.Button(btn_frame, text="Exit", command=self._exit_app).pack(side="left", padx=10)

        ttk.Label(self, text="(Basic local activation check - not a secure commercial license)",
                  font=("Segoe UI", 8), foreground="gray").pack(side="bottom", pady=8)

        self.bind("<Return>", lambda e: self._activate())

    def _activate(self):
        key = self.key_var.get().strip()
        if key == ACTIVATION_KEY:
            self.db.set_setting("activated", "1")
            messagebox.showinfo("Activated", "Software activated successfully!", parent=self)
            self.destroy()
            self.on_activated()
        else:
            messagebox.showerror("Invalid Key", "Invalid activation key.", parent=self)

    def _exit_app(self):
        self.master.destroy()
        sys.exit(0)


# =====================================================================
# 9. LOGIN WINDOW
# =====================================================================
class LoginWindow(tk.Toplevel):
    def __init__(self, master, session: Session, on_login_success):
        super().__init__(master)
        self.session = session
        self.on_login_success = on_login_success
        self.title(f"{APP_NAME} - Login")
        self.geometry("420x340")
        self.resizable(False, False)
        self.protocol("WM_DELETE_WINDOW", self._exit_app)
        self.configure(bg=APP_BG)
        self.grab_set()

        ttk.Label(self, text=APP_NAME, font=("Segoe UI", 15, "bold")).pack(pady=(25, 2))
        ttk.Label(self, text=f"by {APP_OWNER}", font=("Segoe UI", 9), foreground="gray").pack(pady=(0, 20))

        form = ttk.Frame(self)
        form.pack(pady=5)

        ttk.Label(form, text="Username:").grid(row=0, column=0, sticky="e", padx=6, pady=8)
        self.username_var = tk.StringVar()
        user_entry = ttk.Entry(form, textvariable=self.username_var, width=25)
        user_entry.grid(row=0, column=1, pady=8)
        user_entry.focus()

        ttk.Label(form, text="Password:").grid(row=1, column=0, sticky="e", padx=6, pady=8)
        self.password_var = tk.StringVar()
        pw_entry = ttk.Entry(form, textvariable=self.password_var, show="*", width=25)
        pw_entry.grid(row=1, column=1, pady=8)

        btn_frame = ttk.Frame(self)
        btn_frame.pack(pady=20)
        ttk.Button(btn_frame, text="Login", command=self._do_login).pack(side="left", padx=10)
        ttk.Button(btn_frame, text="Exit", command=self._exit_app).pack(side="left", padx=10)

        ttk.Label(self, text="Default: admin/admin123, management/management123,\n"
                              "teacher/teacher123, staff/staff123",
                  font=("Segoe UI", 8), foreground="gray", justify="center").pack(side="bottom", pady=10)

        self.bind("<Return>", lambda e: self._do_login())

    def _do_login(self):
        username = self.username_var.get().strip()
        password = self.password_var.get()
        if not username or not password:
            messagebox.showwarning("Missing Info", "Please enter both username and password.", parent=self)
            return
        try:
            user, error = self.session.login(username, password)
        except Exception as exc:
            show_exception("logging in", exc)
            return
        if error:
            messagebox.showerror("Login Failed", error, parent=self)
            return
        self.destroy()
        self.on_login_success()

    def _exit_app(self):
        self.master.destroy()
        sys.exit(0)


# =====================================================================
# 10-13. MAIN APPLICATION (post-login) -- role-based dashboard
# =====================================================================
class MainApplication:
    """
    Builds the sidebar + content area inside the ALREADY-VISIBLE root
    window (root.deiconify() is called before this is constructed --
    see AppController below). Each role sees only the relevant nav
    buttons, but every action still goes through Session, which
    re-validates permissions server-side / in SQL.
    """

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

    def __init__(self, root, session: Session, on_logout):
        self.root = root
        self.session = session
        self.on_logout = on_logout
        self.frame = ttk.Frame(root)
        self.frame.pack(fill="both", expand=True)
        self._build_ui()
        self.show_view("dashboard")

    # -------------------------------------------------------------
    def _build_ui(self):
        # ----- Header -----
        header = tk.Frame(self.frame, bg=HEADER_BG, height=55)
        header.pack(side="top", fill="x")
        header.pack_propagate(False)
        tk.Label(header, text=APP_NAME, bg=HEADER_BG, fg="white",
                 font=("Segoe UI", 14, "bold")).pack(side="left", padx=15)
        user = self.session.current_user
        tk.Label(header, text=f"{user['full_name']}  ({user['role'].capitalize()})",
                 bg=HEADER_BG, fg="#9ca3af", font=("Segoe UI", 10)).pack(side="right", padx=15)

        body = ttk.Frame(self.frame)
        body.pack(fill="both", expand=True)

        # ----- Sidebar -----
        sidebar = tk.Frame(body, bg=SIDEBAR_BG, width=190)
        sidebar.pack(side="left", fill="y")
        sidebar.pack_propagate(False)

        self.nav_buttons = {}
        for label, key in self.NAV_BY_ROLE.get(user["role"], []):
            btn = tk.Button(
                sidebar, text=label, bg=SIDEBAR_BG, fg=SIDEBAR_FG,
                activebackground=SIDEBAR_ACTIVE_BG, activeforeground="white",
                relief="flat", anchor="w", padx=20, pady=10,
                font=("Segoe UI", 10),
                command=lambda k=key: self.show_view(k)
            )
            btn.pack(fill="x")
            self.nav_buttons[key] = btn

        tk.Frame(sidebar, bg=SIDEBAR_BG).pack(fill="both", expand=True)  # spacer
        tk.Button(sidebar, text="Change Password", bg=SIDEBAR_BG, fg=SIDEBAR_FG,
                  activebackground=SIDEBAR_ACTIVE_BG, activeforeground="white",
                  relief="flat", anchor="w", padx=20, pady=10, font=("Segoe UI", 10),
                  command=self._change_password).pack(fill="x")
        tk.Button(sidebar, text="Logout", bg="#7f1d1d", fg="white",
                  activebackground="#991b1b", activeforeground="white",
                  relief="flat", anchor="w", padx=20, pady=10, font=("Segoe UI", 10),
                  command=self._logout).pack(fill="x")
        tk.Button(sidebar, text="Exit", bg=SIDEBAR_BG, fg=SIDEBAR_FG,
                  activebackground=SIDEBAR_ACTIVE_BG, activeforeground="white",
                  relief="flat", anchor="w", padx=20, pady=10, font=("Segoe UI", 10),
                  command=self._exit_app).pack(fill="x")

        # ----- Content -----
        self.content = ttk.Frame(body, padding=15)
        self.content.pack(side="left", fill="both", expand=True)

    def _clear_content(self):
        for w in self.content.winfo_children():
            w.destroy()

    def show_view(self, key):
        self._clear_content()
        try:
            method = getattr(self, f"view_{key}", None)
            if method:
                method()
            else:
                ttk.Label(self.content, text="View not implemented.").pack()
        except PermissionError_ as pe:
            messagebox.showerror("Permission Denied", str(pe))
        except Exception as exc:
            show_exception(f"opening view '{key}'", exc)

    def _change_password(self):
        result = FormDialog.ask(
            self.root, "Change Password",
            [("old", "Current Password", "password", None),
             ("new", "New Password", "password", None),
             ("confirm", "Confirm New Password", "password", None)]
        )
        if not result:
            return
        if result["new"] != result["confirm"]:
            messagebox.showerror("Error", "New passwords do not match.")
            return
        if len(result["new"]) < 4:
            messagebox.showerror("Error", "New password must be at least 4 characters.")
            return
        try:
            self.session.change_own_password(result["old"], result["new"])
            messagebox.showinfo("Success", "Password changed successfully.")
        except Exception as exc:
            messagebox.showerror("Error", str(exc))

    def _logout(self):
        if messagebox.askyesno("Logout", "Are you sure you want to log out?"):
            self.session.logout()
            self.frame.destroy()
            self.on_logout()

    def _exit_app(self):
        if messagebox.askyesno("Exit", "Are you sure you want to exit?"):
            self.root.destroy()
            sys.exit(0)

    # ===============================================================
    # DASHBOARD VIEW
    # ===============================================================
    def view_dashboard(self):
        user = self.session.current_user
        ttk.Label(self.content, text=f"Welcome, {user['full_name']}!",
                  font=("Segoe UI", 16, "bold")).pack(anchor="w", pady=(0, 15))
        stats = self.session.dashboard_stats()
        cards_frame = ttk.Frame(self.content)
        cards_frame.pack(anchor="w", fill="x")
        col = 0
        for label, value in stats.items():
            card = tk.Frame(cards_frame, bg="white", relief="solid", bd=1, padx=20, pady=15)
            card.grid(row=0, column=col, padx=8, pady=8, sticky="nsew")
            tk.Label(card, text=str(value), font=("Segoe UI", 22, "bold"), bg="white",
                     fg=ACCENT_COLOR).pack()
            tk.Label(card, text=label, font=("Segoe UI", 10), bg="white", fg="#374151").pack()
            col += 1
        for c in range(col):
            cards_frame.grid_columnconfigure(c, weight=1)

        if not PIL_AVAILABLE or not OPENPYXL_AVAILABLE or not REPORTLAB_AVAILABLE:
            missing = []
            if not PIL_AVAILABLE:
                missing.append("pillow (photos)")
            if not OPENPYXL_AVAILABLE:
                missing.append("openpyxl (Excel export)")
            if not REPORTLAB_AVAILABLE:
                missing.append("reportlab (PDF result cards)")
            ttk.Label(
                self.content,
                text="Optional features unavailable - missing packages: " + ", ".join(missing) +
                     "\nInstall with: pip install pillow openpyxl reportlab",
                foreground="#b45309", wraplength=650, justify="left"
            ).pack(anchor="w", pady=(20, 0))

    # ===============================================================
    # CLASSES VIEW
    # ===============================================================
    def view_classes(self):
        user = self.session.current_user
        top = ttk.Frame(self.content)
        top.pack(fill="x", pady=(0, 10))
        ttk.Label(top, text="Classes", font=("Segoe UI", 14, "bold")).pack(side="left")

        can_manage = has_permission(user["role"], "manage_classes_add")
        if can_manage:
            ttk.Button(top, text="Add Class", command=self._add_class).pack(side="right", padx=3)
            ttk.Button(top, text="Edit Class", command=self._edit_class).pack(side="right", padx=3)
            ttk.Button(top, text="Delete Class", command=self._delete_class).pack(side="right", padx=3)
        if user["role"] in ("admin", "management"):
            ttk.Button(top, text="Assign Teacher", command=self._open_teacher_assignment).pack(side="right", padx=3)
        ttk.Button(top, text="Refresh", command=self.view_classes).pack(side="right", padx=3)

        frame, tree = make_scrollable_treeview(
            self.content,
            columns=("id", "class_name", "section", "year", "status"),
            headings=("ID", "Class Name", "Section", "Academic Year", "Status"),
            widths=(50, 150, 100, 130, 90)
        )
        frame.pack(fill="both", expand=True)
        self.classes_tree = tree

        for c in self.session.list_classes():
            tree.insert("", "end", iid=c["id"], values=(
                c["id"], c["class_name"], c["section"] or "", c["academic_year"] or "",
                "Active" if c["is_active"] else "Inactive"
            ))

    def _add_class(self):
        result = FormDialog.ask(self.root, "Add Class", [
            ("class_name", "Class Name", "text", None),
            ("section", "Section", "text", None),
            ("academic_year", "Academic Year", "text", None),
        ])
        if not result:
            return
        if not result["class_name"].strip():
            messagebox.showerror("Error", "Class Name is required.")
            return
        try:
            self.session.add_class(result["class_name"].strip(), result["section"].strip(),
                                    result["academic_year"].strip())
            messagebox.showinfo("Success", "Class added.")
            self.view_classes()
        except Exception as exc:
            messagebox.showerror("Error", str(exc))

    def _get_selected_class_id(self):
        sel = self.classes_tree.selection()
        if not sel:
            messagebox.showwarning("No Selection", "Please select a class first.")
            return None
        return int(sel[0])

    def _edit_class(self):
        class_id = self._get_selected_class_id()
        if class_id is None:
            return
        c = self.session.db.query_one("SELECT * FROM classes WHERE id=?", (class_id,))
        result = FormDialog.ask(self.root, "Edit Class", [
            ("class_name", "Class Name", "text", None),
            ("section", "Section", "text", None),
            ("academic_year", "Academic Year", "text", None),
            ("status", "Status", "combo", ["Active", "Inactive"]),
        ], initial={"class_name": c["class_name"], "section": c["section"] or "",
                    "academic_year": c["academic_year"] or "",
                    "status": "Active" if c["is_active"] else "Inactive"})
        if not result:
            return
        try:
            self.session.edit_class(class_id, result["class_name"].strip(), result["section"].strip(),
                                     result["academic_year"].strip(), result["status"] == "Active")
            messagebox.showinfo("Success", "Class updated.")
            self.view_classes()
        except Exception as exc:
            messagebox.showerror("Error", str(exc))

    def _delete_class(self):
        class_id = self._get_selected_class_id()
        if class_id is None:
            return
        if not messagebox.askyesno("Confirm", "Are you sure you want to permanently delete this class?"):
            return
        try:
            self.session.delete_class(class_id)
            messagebox.showinfo("Deleted", "Class deleted.")
            self.view_classes()
        except Exception as exc:
            messagebox.showerror("Error", str(exc))

    def _open_teacher_assignment(self):
        TeacherAssignmentWindow(self.root, self.session)

    # ===============================================================
    # TEACHERS VIEW (list of teachers + their assigned classes)
    # ===============================================================
    def view_teachers(self):
        user = self.session.current_user
        top = ttk.Frame(self.content)
        top.pack(fill="x", pady=(0, 10))
        ttk.Label(top, text="Teachers", font=("Segoe UI", 14, "bold")).pack(side="left")

        can_add = user["role"] == "admin" or has_permission(user["role"], "manage_users_add_teacher")
        if can_add:
            ttk.Button(top, text="Add Teacher", command=self._add_teacher).pack(side="right", padx=3)
        if user["role"] in ("admin", "management"):
            ttk.Button(top, text="Assign Classes", command=self._open_teacher_assignment).pack(side="right", padx=3)
        ttk.Button(top, text="Refresh", command=self.view_teachers).pack(side="right", padx=3)

        frame, tree = make_scrollable_treeview(
            self.content,
            columns=("id", "name", "username", "email", "phone", "classes", "status"),
            headings=("ID", "Full Name", "Username", "Email", "Phone", "Assigned Classes", "Status"),
            widths=(40, 150, 110, 150, 100, 200, 80)
        )
        frame.pack(fill="both", expand=True)

        for t in self.session.list_teachers():
            classes = self.session.list_assignments_for_teacher(t["id"])
            class_str = ", ".join(f"{c['class_name']}-{c['section']}" if c["section"] else c["class_name"]
                                   for c in classes)
            tree.insert("", "end", iid=t["id"], values=(
                t["id"], t["full_name"], t["username"], t["email"] or "", t["phone"] or "",
                class_str, "Active" if t["is_active"] else "Disabled"
            ))

    def _add_teacher(self):
        result = FormDialog.ask(self.root, "Add Teacher", [
            ("username", "Username", "text", None),
            ("password", "Password", "password", None),
            ("full_name", "Full Name", "text", None),
            ("email", "Email", "text", None),
            ("phone", "Phone", "text", None),
        ])
        if not result:
            return
        if not result["username"].strip() or not result["password"] or not result["full_name"].strip():
            messagebox.showerror("Error", "Username, Password and Full Name are required.")
            return
        try:
            self.session.add_user(result["username"].strip(), result["password"],
                                   result["full_name"].strip(), "teacher",
                                   result["email"].strip(), result["phone"].strip())
            messagebox.showinfo("Success", "Teacher account created.")
            self.view_teachers()
        except Exception as exc:
            messagebox.showerror("Error", str(exc))

    # ===============================================================
    # USERS VIEW (admin only)
    # ===============================================================
    def view_users(self):
        user = self.session.current_user
        if not has_permission(user["role"], "manage_users_add"):
            raise PermissionError_("Only Admin can manage all users.")

        top = ttk.Frame(self.content)
        top.pack(fill="x", pady=(0, 10))
        ttk.Label(top, text="System Users", font=("Segoe UI", 14, "bold")).pack(side="left")
        ttk.Button(top, text="Add User", command=self._add_user).pack(side="right", padx=3)
        ttk.Button(top, text="Edit User", command=self._edit_user).pack(side="right", padx=3)
        ttk.Button(top, text="Disable", command=lambda: self._toggle_user(False)).pack(side="right", padx=3)
        ttk.Button(top, text="Enable", command=lambda: self._toggle_user(True)).pack(side="right", padx=3)
        ttk.Button(top, text="Delete User", command=self._delete_user).pack(side="right", padx=3)
        ttk.Button(top, text="Reset Password", command=self._reset_password).pack(side="right", padx=3)
        ttk.Button(top, text="Refresh", command=self.view_users).pack(side="right", padx=3)

        search_frame = ttk.Frame(self.content)
        search_frame.pack(fill="x", pady=(0, 8))
        ttk.Label(search_frame, text="Search:").pack(side="left")
        self.user_search_var = tk.StringVar()
        entry = ttk.Entry(search_frame, textvariable=self.user_search_var, width=30)
        entry.pack(side="left", padx=5)
        ttk.Button(search_frame, text="Go", command=self.view_users).pack(side="left")

        frame, tree = make_scrollable_treeview(
            self.content,
            columns=("id", "username", "full_name", "role", "email", "phone", "status", "created"),
            headings=("ID", "Username", "Full Name", "Role", "Email", "Phone", "Status", "Created"),
            widths=(40, 110, 140, 90, 150, 100, 80, 130)
        )
        frame.pack(fill="both", expand=True)
        self.users_tree = tree

        keyword = getattr(self, "user_search_var", tk.StringVar()).get().strip().lower()
        for u in self.session.list_users():
            if keyword and keyword not in u["username"].lower() and keyword not in u["full_name"].lower():
                continue
            tree.insert("", "end", iid=u["id"], values=(
                u["id"], u["username"], u["full_name"], u["role"], u["email"] or "",
                u["phone"] or "", "Active" if u["is_active"] else "Disabled", u["created_at"]
            ))

    def _get_selected_user_id(self):
        sel = self.users_tree.selection()
        if not sel:
            messagebox.showwarning("No Selection", "Please select a user first.")
            return None
        return int(sel[0])

    def _add_user(self):
        result = FormDialog.ask(self.root, "Add User", [
            ("username", "Username", "text", None),
            ("password", "Password", "password", None),
            ("full_name", "Full Name", "text", None),
            ("role", "Role", "combo", ROLES),
            ("email", "Email", "text", None),
            ("phone", "Phone", "text", None),
        ])
        if not result:
            return
        if not result["username"].strip() or not result["password"] or not result["full_name"].strip():
            messagebox.showerror("Error", "Username, Password and Full Name are required.")
            return
        try:
            self.session.add_user(result["username"].strip(), result["password"], result["full_name"].strip(),
                                   result["role"], result["email"].strip(), result["phone"].strip())
            messagebox.showinfo("Success", "User created.")
            self.view_users()
        except Exception as exc:
            messagebox.showerror("Error", str(exc))

    def _edit_user(self):
        uid = self._get_selected_user_id()
        if uid is None:
            return
        u = self.session.db.query_one("SELECT * FROM users WHERE id=?", (uid,))
        result = FormDialog.ask(self.root, "Edit User", [
            ("full_name", "Full Name", "text", None),
            ("role", "Role", "combo", ROLES),
            ("email", "Email", "text", None),
            ("phone", "Phone", "text", None),
        ], initial={"full_name": u["full_name"], "role": u["role"],
                    "email": u["email"] or "", "phone": u["phone"] or ""})
        if not result:
            return
        try:
            self.session.edit_user(uid, result["full_name"].strip(), result["email"].strip(),
                                    result["phone"].strip(), result["role"])
            messagebox.showinfo("Success", "User updated.")
            self.view_users()
        except Exception as exc:
            messagebox.showerror("Error", str(exc))

    def _toggle_user(self, active):
        uid = self._get_selected_user_id()
        if uid is None:
            return
        try:
            self.session.set_user_active(uid, active)
            self.view_users()
        except Exception as exc:
            messagebox.showerror("Error", str(exc))

    def _delete_user(self):
        uid = self._get_selected_user_id()
        if uid is None:
            return
        if not messagebox.askyesno("Confirm", "Are you sure you want to permanently delete this user?"):
            return
        try:
            self.session.delete_user(uid)
            messagebox.showinfo("Deleted", "User deleted.")
            self.view_users()
        except Exception as exc:
            messagebox.showerror("Error", str(exc))

    def _reset_password(self):
        uid = self._get_selected_user_id()
        if uid is None:
            return
        result = FormDialog.ask(self.root, "Reset Password", [
            ("new_password", "New Password", "password", None),
        ])
        if not result or not result["new_password"]:
            return
        try:
            self.session.reset_password(uid, result["new_password"])
            messagebox.showinfo("Success", "Password reset.")
        except Exception as exc:
            messagebox.showerror("Error", str(exc))

    # ===============================================================
    # STUDENTS VIEW
    # ===============================================================
    def view_students(self):
        user = self.session.current_user
        top = ttk.Frame(self.content)
        top.pack(fill="x", pady=(0, 10))
        ttk.Label(top, text="Students", font=("Segoe UI", 14, "bold")).pack(side="left")

        can_add = has_permission(user["role"], "students_add") or has_permission(user["role"], "students_add_assigned")
        can_edit = has_permission(user["role"], "students_edit") or has_permission(user["role"], "students_edit_assigned")
        can_delete = has_permission(user["role"], "students_delete")

        if can_add:
            ttk.Button(top, text="Add Student", command=self._add_student).pack(side="right", padx=3)
        if can_edit:
            ttk.Button(top, text="Edit Student", command=self._edit_student).pack(side="right", padx=3)
        if can_delete:
            ttk.Button(top, text="Delete Student", command=self._delete_student).pack(side="right", padx=3)
        ttk.Button(top, text="View Details", command=self._view_student_details).pack(side="right", padx=3)
        if OPENPYXL_AVAILABLE and has_permission(user["role"], "export_excel") or has_permission(user["role"], "export_excel_assigned"):
            ttk.Button(top, text="Export Excel", command=self._export_students_excel).pack(side="right", padx=3)
        ttk.Button(top, text="Refresh", command=self.view_students).pack(side="right", padx=3)

        filt = ttk.Frame(self.content)
        filt.pack(fill="x", pady=(0, 8))
        ttk.Label(filt, text="Search:").pack(side="left")
        self.student_search_var = tk.StringVar()
        ttk.Entry(filt, textvariable=self.student_search_var, width=25).pack(side="left", padx=5)

        ttk.Label(filt, text="Class:").pack(side="left", padx=(15, 0))
        classes = self.session.list_classes()
        class_map = {"All": None}
        for c in classes:
            label = f"{c['class_name']}-{c['section']}" if c["section"] else c["class_name"]
            class_map[label] = c["id"]
        self.student_class_var = tk.StringVar(value="All")
        ttk.Combobox(filt, textvariable=self.student_class_var, values=list(class_map.keys()),
                     state="readonly", width=20).pack(side="left", padx=5)
        self._student_class_map = class_map

        ttk.Label(filt, text="Status:").pack(side="left", padx=(15, 0))
        self.student_status_var = tk.StringVar(value="All")
        ttk.Combobox(filt, textvariable=self.student_status_var, values=["All", "Active", "Inactive"],
                     state="readonly", width=12).pack(side="left", padx=5)

        ttk.Button(filt, text="Search", command=self.view_students).pack(side="left", padx=8)
        ttk.Button(filt, text="Clear", command=self._clear_student_filters).pack(side="left")

        frame, tree = make_scrollable_treeview(
            self.content,
            columns=("id", "student_id", "roll", "name", "father", "class", "phone", "status"),
            headings=("Row", "Student ID", "Roll No", "Name", "Father Name", "Class", "Phone", "Status"),
            widths=(0, 100, 80, 150, 150, 120, 100, 80)
        )
        tree.column("id", width=0, stretch=False)  # hidden internal row id column
        frame.pack(fill="both", expand=True)
        self.students_tree = tree

        keyword = getattr(self, "student_search_var", tk.StringVar()).get().strip()
        class_label = getattr(self, "student_class_var", tk.StringVar(value="All")).get()
        class_id = self._student_class_map.get(class_label) if hasattr(self, "_student_class_map") else None
        status = getattr(self, "student_status_var", tk.StringVar(value="All")).get()
        status = None if status == "All" else status

        try:
            students = self.session.search_students(keyword, class_id, status)
        except Exception as exc:
            show_exception("loading students", exc)
            students = []

        for s in students:
            class_label = f"{s['class_name']}-{s['section']}" if s["class_name"] and s["section"] else (s["class_name"] or "")
            tree.insert("", "end", iid=s["id"], values=(
                s["id"], s["student_id"], s["roll_number"] or "",
                f"{s['first_name']} {s['last_name'] or ''}".strip(), s["father_name"] or "",
                class_label, s["phone"] or "", s["status"]
            ))

    def _clear_student_filters(self):
        self.student_search_var.set("")
        self.student_class_var.set("All")
        self.student_status_var.set("All")
        self.view_students()

    def _get_selected_student_row_id(self):
        sel = self.students_tree.selection()
        if not sel:
            messagebox.showwarning("No Selection", "Please select a student first.")
            return None
        return int(sel[0])

    def _class_choices(self):
        classes = self.session.list_classes()
        labels = []
        mapping = {}
        for c in classes:
            label = f"{c['class_name']}-{c['section']}" if c["section"] else c["class_name"]
            labels.append(label)
            mapping[label] = c["id"]
        return labels, mapping

    def _student_form(self, title, initial=None):
        labels, mapping = self._class_choices()
        if not labels:
            messagebox.showwarning("No Classes", "Please add/be assigned a class before adding students.")
            return None, None
        fields = [
            ("student_id", "Student ID", "text", None),
            ("roll_number", "Roll Number", "text", None),
            ("first_name", "First Name", "text", None),
            ("last_name", "Last Name", "text", None),
            ("father_name", "Father Name", "text", None),
            ("date_of_birth", "Date of Birth (YYYY-MM-DD)", "text", None),
            ("gender", "Gender", "combo", ["Male", "Female", "Other"]),
            ("phone", "Phone", "text", None),
            ("address", "Address", "text", None),
            ("class_label", "Class", "combo", labels),
            ("admission_date", "Admission Date (YYYY-MM-DD)", "text", None),
            ("status", "Status", "combo", ["Active", "Inactive"]),
        ]
        result = FormDialog.ask(self.root, title, fields, initial)
        return result, mapping

    def _add_student(self):
        result, mapping = self._student_form("Add Student", initial={
            "date_of_birth": date.today().isoformat(),
            "admission_date": date.today().isoformat(),
            "status": "Active",
        })
        if not result:
            return
        if not result["student_id"].strip() or not result["first_name"].strip():
            messagebox.showerror("Error", "Student ID and First Name are required.")
            return
        photo_path = None
        if PIL_AVAILABLE:
            if messagebox.askyesno("Photo", "Do you want to upload a photo now?"):
                photo_path = filedialog.askopenfilename(
                    title="Select Student Photo",
                    filetypes=[("Image files", "*.jpg *.jpeg *.png")]
                )
        data = {
            "student_id": result["student_id"].strip(),
            "roll_number": result["roll_number"].strip(),
            "first_name": result["first_name"].strip(),
            "last_name": result["last_name"].strip(),
            "father_name": result["father_name"].strip(),
            "date_of_birth": result["date_of_birth"].strip(),
            "gender": result["gender"],
            "phone": result["phone"].strip(),
            "address": result["address"].strip(),
            "class_id": mapping.get(result["class_label"]),
            "admission_date": result["admission_date"].strip(),
            "status": result["status"],
        }
        try:
            self.session.add_student(data, photo_path if photo_path else None)
            messagebox.showinfo("Success", "Student added.")
            self.view_students()
        except Exception as exc:
            messagebox.showerror("Error", str(exc))

    def _edit_student(self):
        row_id = self._get_selected_student_row_id()
        if row_id is None:
            return
        try:
            s = self.session.get_student(row_id)
        except Exception as exc:
            messagebox.showerror("Error", str(exc))
            return
        labels, mapping = self._class_choices()
        current_label = None
        for label, cid in mapping.items():
            if cid == s["class_id"]:
                current_label = label
        initial = {
            "student_id": s["student_id"], "roll_number": s["roll_number"] or "",
            "first_name": s["first_name"], "last_name": s["last_name"] or "",
            "father_name": s["father_name"] or "", "date_of_birth": s["date_of_birth"] or "",
            "gender": s["gender"] or "Male", "phone": s["phone"] or "", "address": s["address"] or "",
            "class_label": current_label or (labels[0] if labels else ""),
            "admission_date": s["admission_date"] or "", "status": s["status"],
        }
        result, mapping = self._student_form("Edit Student", initial)
        if not result:
            return
        photo_path = None
        if PIL_AVAILABLE and messagebox.askyesno("Photo", "Do you want to change the photo?"):
            photo_path = filedialog.askopenfilename(
                title="Select Student Photo", filetypes=[("Image files", "*.jpg *.jpeg *.png")]
            )
        data = {
            "student_id": result["student_id"].strip(), "roll_number": result["roll_number"].strip(),
            "first_name": result["first_name"].strip(), "last_name": result["last_name"].strip(),
            "father_name": result["father_name"].strip(), "date_of_birth": result["date_of_birth"].strip(),
            "gender": result["gender"], "phone": result["phone"].strip(), "address": result["address"].strip(),
            "class_id": mapping.get(result["class_label"]), "admission_date": result["admission_date"].strip(),
            "status": result["status"],
        }
        try:
            self.session.edit_student(row_id, data, photo_path if photo_path else None)
            messagebox.showinfo("Success", "Student updated.")
            self.view_students()
        except Exception as exc:
            messagebox.showerror("Error", str(exc))

    def _delete_student(self):
        row_id = self._get_selected_student_row_id()
        if row_id is None:
            return
        if not messagebox.askyesno("Confirm", "Are you sure you want to permanently delete this student?"):
            return
        try:
            self.session.delete_student(row_id)
            messagebox.showinfo("Deleted", "Student deleted.")
            self.view_students()
        except Exception as exc:
            messagebox.showerror("Error", str(exc))

    def _view_student_details(self):
        row_id = self._get_selected_student_row_id()
        if row_id is None:
            return
        try:
            StudentDetailWindow(self.root, self.session, row_id)
        except Exception as exc:
            messagebox.showerror("Error", str(exc))

    def _export_students_excel(self):
        if not OPENPYXL_AVAILABLE:
            messagebox.showwarning("Missing Package",
                                    "Excel export requires openpyxl.\nInstall it using:\npip install openpyxl")
            return
        try:
            students = self.session.search_students(
                getattr(self, "student_search_var", tk.StringVar()).get().strip()
            )
            path = export_students_to_excel(students)
            messagebox.showinfo("Exported", f"Students exported to:\n{path}")
        except Exception as exc:
            show_exception("exporting students to Excel", exc)

    # ===============================================================
    # ATTENDANCE VIEW
    # ===============================================================
    def view_attendance(self):
        AttendancePanel(self.content, self.session)

    # ===============================================================
    # MARKS VIEW
    # ===============================================================
    def view_marks(self):
        MarksPanel(self.content, self.session, self.root)

    # ===============================================================
    # REPORTS VIEW
    # ===============================================================
    def view_reports(self):
        ReportsPanel(self.content, self.session)

    # ===============================================================
    # AUDIT LOG VIEW (admin only)
    # ===============================================================
    def view_audit(self):
        user = self.session.current_user
        if not has_permission(user["role"], "view_audit_log"):
            raise PermissionError_("Only Admin can view the audit log.")
        ttk.Label(self.content, text="Audit Log", font=("Segoe UI", 14, "bold")).pack(anchor="w", pady=(0, 10))
        frame, tree = make_scrollable_treeview(
            self.content,
            columns=("time", "user", "action", "description"),
            headings=("Date/Time", "User", "Action", "Description"),
            widths=(150, 130, 130, 350)
        )
        frame.pack(fill="both", expand=True)
        rows = self.session.db.query(
            """SELECT al.*, u.username FROM audit_log al
               LEFT JOIN users u ON u.id = al.user_id
               ORDER BY al.id DESC LIMIT 500"""
        )
        for r in rows:
            tree.insert("", "end", values=(r["timestamp"], r["username"] or "(unknown)",
                                            r["action"], r["description"] or ""))


# =====================================================================
# TEACHER ASSIGNMENT WINDOW
# =====================================================================
class TeacherAssignmentWindow(tk.Toplevel):
    def __init__(self, master, session: Session):
        super().__init__(master)
        self.session = session
        self.title("Teacher Class Assignment")
        self.geometry("560x420")
        self.transient(master)
        self.grab_set()

        top = ttk.Frame(self, padding=10)
        top.pack(fill="x")
        ttk.Label(top, text="Teacher:").pack(side="left")
        self.teacher_map = {}
        teachers = session.list_teachers()
        labels = []
        for t in teachers:
            labels.append(t["full_name"])
            self.teacher_map[t["full_name"]] = t["id"]
        self.teacher_var = tk.StringVar(value=labels[0] if labels else "")
        combo = ttk.Combobox(top, textvariable=self.teacher_var, values=labels, state="readonly", width=25)
        combo.pack(side="left", padx=8)
        combo.bind("<<ComboboxSelected>>", lambda e: self._refresh())

        body = ttk.Frame(self, padding=10)
        body.pack(fill="both", expand=True)

        left = ttk.LabelFrame(body, text="All Classes", padding=8)
        left.pack(side="left", fill="both", expand=True, padx=5)
        self.all_list = tk.Listbox(left, selectmode="single")
        self.all_list.pack(fill="both", expand=True)

        mid = ttk.Frame(body)
        mid.pack(side="left", padx=8)
        ttk.Button(mid, text="Assign ->", command=self._assign).pack(pady=10)
        ttk.Button(mid, text="<- Remove", command=self._remove).pack(pady=10)

        right = ttk.LabelFrame(body, text="Assigned Classes", padding=8)
        right.pack(side="left", fill="both", expand=True, padx=5)
        self.assigned_list = tk.Listbox(right, selectmode="single")
        self.assigned_list.pack(fill="both", expand=True)

        ttk.Button(self, text="Close", command=self.destroy).pack(pady=8)

        self._class_map = {}
        self._refresh()

    def _refresh(self):
        self.all_list.delete(0, "end")
        self.assigned_list.delete(0, "end")
        classes = self.session.db.query("SELECT * FROM classes ORDER BY class_name, section")
        self._class_map = {}
        teacher_name = self.teacher_var.get()
        teacher_id = self.teacher_map.get(teacher_name)
        assigned_ids = set()
        if teacher_id:
            for c in self.session.list_assignments_for_teacher(teacher_id):
                assigned_ids.add(c["id"])

        for c in classes:
            label = f"{c['class_name']}-{c['section']}" if c["section"] else c["class_name"]
            self._class_map[label] = c["id"]
            if c["id"] in assigned_ids:
                self.assigned_list.insert("end", label)
            else:
                self.all_list.insert("end", label)

    def _assign(self):
        sel = self.all_list.curselection()
        if not sel:
            return
        label = self.all_list.get(sel[0])
        class_id = self._class_map[label]
        teacher_id = self.teacher_map.get(self.teacher_var.get())
        if not teacher_id:
            messagebox.showwarning("No Teacher", "Please select a teacher first.")
            return
        try:
            self.session.assign_teacher(teacher_id, class_id)
            self._refresh()
        except Exception as exc:
            messagebox.showerror("Error", str(exc))

    def _remove(self):
        sel = self.assigned_list.curselection()
        if not sel:
            return
        label = self.assigned_list.get(sel[0])
        class_id = self._class_map[label]
        teacher_id = self.teacher_map.get(self.teacher_var.get())
        try:
            self.session.remove_teacher_assignment(teacher_id, class_id)
            self._refresh()
        except Exception as exc:
            messagebox.showerror("Error", str(exc))


# =====================================================================
# STUDENT DETAIL WINDOW
# =====================================================================
class StudentDetailWindow(tk.Toplevel):
    def __init__(self, master, session: Session, student_row_id):
        super().__init__(master)
        self.session = session
        self.student_row_id = student_row_id
        self.title("Student Details")
        self.geometry("500x560")
        self.transient(master)
        self.grab_set()

        s = session.get_student(student_row_id)
        if s is None:
            messagebox.showerror("Error", "Student not found.", parent=self)
            self.destroy()
            return

        top = ttk.Frame(self, padding=15)
        top.pack(fill="both", expand=True)

        photo_frame = ttk.Frame(top)
        photo_frame.pack(pady=(0, 10))
        self._photo_ref = None
        if PIL_AVAILABLE and s["photo_path"] and os.path.exists(s["photo_path"]):
            try:
                img = Image.open(s["photo_path"])
                img.thumbnail((140, 140))
                self._photo_ref = ImageTk.PhotoImage(img)
                tk.Label(photo_frame, image=self._photo_ref).pack()
            except Exception:
                tk.Label(photo_frame, text="[Photo unavailable]").pack()
        else:
            tk.Label(photo_frame, text="[No Photo]", width=18, height=8, relief="groove").pack()

        info_frame = ttk.Frame(top)
        info_frame.pack(fill="x")
        rows = [
            ("Student ID", s["student_id"]), ("Roll Number", s["roll_number"] or "-"),
            ("Name", f"{s['first_name']} {s['last_name'] or ''}".strip()),
            ("Father Name", s["father_name"] or "-"), ("Date of Birth", s["date_of_birth"] or "-"),
            ("Gender", s["gender"] or "-"), ("Phone", s["phone"] or "-"),
            ("Address", s["address"] or "-"),
            ("Class", f"{s['class_name']}-{s['section']}" if s["class_name"] and s["section"] else (s["class_name"] or "-")),
            ("Admission Date", s["admission_date"] or "-"), ("Status", s["status"]),
        ]
        for i, (label, value) in enumerate(rows):
            ttk.Label(info_frame, text=label + ":", font=("Segoe UI", 9, "bold")).grid(
                row=i, column=0, sticky="ne", padx=5, pady=3)
            ttk.Label(info_frame, text=str(value), wraplength=300, justify="left").grid(
                row=i, column=1, sticky="nw", padx=5, pady=3)

        # Attendance + marks summary
        att = session.attendance_report(student_row_id=student_row_id)
        att_pct = att[0]["percentage"] if att else 0.0
        result = session.result_summary(student_row_id)

        summary = ttk.LabelFrame(top, text="Academic Summary", padding=8)
        summary.pack(fill="x", pady=(12, 0))
        ttk.Label(summary, text=f"Attendance: {att_pct}%").pack(anchor="w")
        ttk.Label(summary, text=f"Overall Marks: {result['obtained']}/{result['total']} "
                                 f"({result['percentage']}%)").pack(anchor="w")
        ttk.Label(summary, text=f"Grade: {result['grade']}   Result: {result['result']}").pack(anchor="w")

        btns = ttk.Frame(top)
        btns.pack(pady=15)
        ttk.Button(btns, text="Generate Result Card", command=self._open_result_card).pack(side="left", padx=5)
        ttk.Button(btns, text="Close", command=self.destroy).pack(side="left", padx=5)

    def _open_result_card(self):
        ResultCardWindow(self, self.session, self.student_row_id)


# =====================================================================
# ATTENDANCE PANEL (embedded, not a popup, since it's a main nav view)
# =====================================================================
class AttendancePanel:
    def __init__(self, parent, session: Session):
        self.parent = parent
        self.session = session
        self.row_vars = {}
        self._build()

    def _build(self):
        top = ttk.Frame(self.parent)
        top.pack(fill="x", pady=(0, 10))
        ttk.Label(top, text="Attendance", font=("Segoe UI", 14, "bold")).pack(side="left")

        ttk.Label(top, text="Class:").pack(side="left", padx=(20, 3))
        classes = self.session.list_classes()
        self.class_map = {}
        labels = []
        for c in classes:
            label = f"{c['class_name']}-{c['section']}" if c["section"] else c["class_name"]
            labels.append(label)
            self.class_map[label] = c["id"]
        self.class_var = tk.StringVar(value=labels[0] if labels else "")
        ttk.Combobox(top, textvariable=self.class_var, values=labels, state="readonly", width=20).pack(side="left")

        ttk.Label(top, text="Date (YYYY-MM-DD):").pack(side="left", padx=(15, 3))
        self.date_var = tk.StringVar(value=date.today().isoformat())
        ttk.Entry(top, textvariable=self.date_var, width=14).pack(side="left")

        ttk.Button(top, text="Load", command=self._load).pack(side="left", padx=10)
        ttk.Button(top, text="Mark All Present", command=lambda: self._mark_all("Present")).pack(side="left", padx=3)
        ttk.Button(top, text="Mark All Absent", command=lambda: self._mark_all("Absent")).pack(side="left", padx=3)
        ttk.Button(top, text="Save Attendance", command=self._save).pack(side="right", padx=3)

        self.list_frame = ttk.Frame(self.parent)
        self.list_frame.pack(fill="both", expand=True, pady=10)

        if labels:
            self._load()

    def _load(self):
        for w in self.list_frame.winfo_children():
            w.destroy()
        self.row_vars = {}

        class_label = self.class_var.get()
        class_id = self.class_map.get(class_label)
        if not class_id:
            ttk.Label(self.list_frame, text="No classes available.").pack()
            return
        att_date = self.date_var.get().strip()
        try:
            datetime.strptime(att_date, "%Y-%m-%d")
        except ValueError:
            messagebox.showerror("Invalid Date", "Please use YYYY-MM-DD format.")
            return

        try:
            rows = self.session.load_attendance(class_id, att_date)
        except Exception as exc:
            show_exception("loading attendance", exc)
            return

        header = ttk.Frame(self.list_frame)
        header.pack(fill="x")
        ttk.Label(header, text="Roll No", width=10, font=("Segoe UI", 9, "bold")).pack(side="left")
        ttk.Label(header, text="Name", width=30, font=("Segoe UI", 9, "bold")).pack(side="left")
        ttk.Label(header, text="Status", font=("Segoe UI", 9, "bold")).pack(side="left")

        canvas = tk.Canvas(self.list_frame, borderwidth=0)
        scroll_frame = ttk.Frame(canvas)
        vsb = ttk.Scrollbar(self.list_frame, orient="vertical", command=canvas.yview)
        canvas.configure(yscrollcommand=vsb.set)
        vsb.pack(side="right", fill="y")
        canvas.pack(side="left", fill="both", expand=True)
        canvas.create_window((0, 0), window=scroll_frame, anchor="nw")
        scroll_frame.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))

        if not rows:
            ttk.Label(scroll_frame, text="No active students in this class.").pack(anchor="w")

        for r in rows:
            row_frame = ttk.Frame(scroll_frame)
            row_frame.pack(fill="x", pady=2)
            ttk.Label(row_frame, text=r["roll_number"] or "-", width=10).pack(side="left")
            ttk.Label(row_frame, text=f"{r['first_name']} {r['last_name'] or ''}".strip(), width=30).pack(side="left")
            var = tk.StringVar(value=r["status"] or "Present")
            ttk.Combobox(row_frame, textvariable=var, values=ATTENDANCE_STATUSES,
                         state="readonly", width=12).pack(side="left")
            self.row_vars[r["student_row_id"]] = var

        self._current_class_id = class_id
        self._current_date = att_date

    def _mark_all(self, status):
        if not self.row_vars:
            return
        for var in self.row_vars.values():
            var.set(status)

    def _save(self):
        if not self.row_vars:
            messagebox.showwarning("Nothing to Save", "Load a class's attendance first.")
            return
        status_map = {sid: var.get() for sid, var in self.row_vars.items()}
        try:
            self.session.save_attendance(self._current_class_id, self._current_date, status_map)
            messagebox.showinfo("Saved", "Attendance saved successfully.")
        except Exception as exc:
            messagebox.showerror("Error", str(exc))


# =====================================================================
# MARKS PANEL
# =====================================================================
class MarksPanel:
    def __init__(self, parent, session: Session, root):
        self.parent = parent
        self.session = session
        self.root = root
        self._build()

    def _build(self):
        top = ttk.Frame(self.parent)
        top.pack(fill="x", pady=(0, 10))
        ttk.Label(top, text="Marks", font=("Segoe UI", 14, "bold")).pack(side="left")

        ttk.Label(top, text="Student:").pack(side="left", padx=(20, 3))
        try:
            students = self.session.search_students("")
        except Exception:
            students = []
        self.student_map = {}
        labels = []
        for s in students:
            label = f"{s['student_id']} - {s['first_name']} {s['last_name'] or ''}".strip()
            labels.append(label)
            self.student_map[label] = s["id"]
        self.student_var = tk.StringVar(value=labels[0] if labels else "")
        combo = ttk.Combobox(top, textvariable=self.student_var, values=labels, state="readonly", width=30)
        combo.pack(side="left")
        combo.bind("<<ComboboxSelected>>", lambda e: self._load())

        ttk.Button(top, text="Load", command=self._load).pack(side="left", padx=10)
        ttk.Button(top, text="Add / Update Mark", command=self._add_mark).pack(side="right", padx=3)
        ttk.Button(top, text="Delete Selected", command=self._delete_mark).pack(side="right", padx=3)
        ttk.Button(top, text="Result Card", command=self._result_card).pack(side="right", padx=3)

        frame, tree = make_scrollable_treeview(
            self.parent,
            columns=("id", "exam", "subject", "total", "obtained", "pct", "grade"),
            headings=("ID", "Exam", "Subject", "Total", "Obtained", "%", "Grade"),
            widths=(0, 130, 130, 70, 80, 70, 60)
        )
        tree.column("id", width=0, stretch=False)
        frame.pack(fill="both", expand=True, pady=10)
        self.tree = tree

        self.summary_label = ttk.Label(self.parent, text="", font=("Segoe UI", 10, "bold"))
        self.summary_label.pack(anchor="w")

        if labels:
            self._load()

    def _current_student_row_id(self):
        return self.student_map.get(self.student_var.get())

    def _load(self):
        for i in self.tree.get_children():
            self.tree.delete(i)
        sid = self._current_student_row_id()
        if not sid:
            return
        try:
            result = self.session.result_summary(sid)
        except Exception as exc:
            show_exception("loading marks", exc)
            return
        for m in result["marks"]:
            pct = round(m["obtained_marks"] / m["total_marks"] * 100, 1) if m["total_marks"] else 0
            grade = calculate_grade(pct)
            self.tree.insert("", "end", iid=m["id"], values=(
                m["id"], m["exam_name"], m["subject"], m["total_marks"], m["obtained_marks"], pct, grade
            ))
        self.summary_label.config(
            text=f"Overall: {result['obtained']}/{result['total']} ({result['percentage']}%) "
                 f"- Grade {result['grade']} - {result['result']}"
        )

    def _add_mark(self):
        sid = self._current_student_row_id()
        if not sid:
            messagebox.showwarning("No Student", "Please select a student first.")
            return
        result = FormDialog.ask(self.root, "Add / Update Mark", [
            ("exam_name", "Exam Name", "text", None),
            ("subject", "Subject", "text", None),
            ("total_marks", "Total Marks", "text", None),
            ("obtained_marks", "Obtained Marks", "text", None),
        ])
        if not result:
            return
        try:
            total = float(result["total_marks"])
            obtained = float(result["obtained_marks"])
        except ValueError:
            messagebox.showerror("Error", "Total and Obtained marks must be numbers.")
            return
        try:
            self.session.save_mark(sid, result["subject"].strip(), result["exam_name"].strip(), total, obtained)
            self._load()
        except Exception as exc:
            messagebox.showerror("Error", str(exc))

    def _delete_mark(self):
        sel = self.tree.selection()
        if not sel:
            messagebox.showwarning("No Selection", "Please select a mark row first.")
            return
        sid = self._current_student_row_id()
        if not messagebox.askyesno("Confirm", "Delete this mark entry?"):
            return
        try:
            self.session.delete_mark(int(sel[0]), sid)
            self._load()
        except Exception as exc:
            messagebox.showerror("Error", str(exc))

    def _result_card(self):
        sid = self._current_student_row_id()
        if not sid:
            messagebox.showwarning("No Student", "Please select a student first.")
            return
        ResultCardWindow(self.root, self.session, sid)


# =====================================================================
# RESULT CARD WINDOW (preview + PDF export)
# =====================================================================
class ResultCardWindow(tk.Toplevel):
    def __init__(self, master, session: Session, student_row_id):
        super().__init__(master)
        self.session = session
        self.student_row_id = student_row_id
        self.title("Result Card")
        self.geometry("520x600")
        self.transient(master)
        self.grab_set()

        s = session.get_student(student_row_id)
        result = session.result_summary(student_row_id)
        school_name = session.db.get_setting("school_name", "RJK School / College")

        wrapper = ttk.Frame(self, padding=15)
        wrapper.pack(fill="both", expand=True)

        ttk.Label(wrapper, text=school_name, font=("Segoe UI", 15, "bold")).pack()
        ttk.Label(wrapper, text="Result Card", font=("Segoe UI", 12, "italic")).pack(pady=(0, 10))

        info = ttk.Frame(wrapper)
        info.pack(fill="x", pady=5)
        details = [
            ("Student Name", f"{s['first_name']} {s['last_name'] or ''}".strip()),
            ("Father Name", s["father_name"] or "-"),
            ("Student ID", s["student_id"]),
            ("Roll Number", s["roll_number"] or "-"),
            ("Class", f"{s['class_name']}-{s['section']}" if s["class_name"] and s["section"] else (s["class_name"] or "-")),
        ]
        for i, (label, value) in enumerate(details):
            ttk.Label(info, text=label + ":", font=("Segoe UI", 9, "bold")).grid(row=i // 2, column=(i % 2) * 2, sticky="w", padx=5, pady=2)
            ttk.Label(info, text=value).grid(row=i // 2, column=(i % 2) * 2 + 1, sticky="w", padx=5, pady=2)

        frame, tree = make_scrollable_treeview(
            wrapper,
            columns=("subject", "total", "obtained", "pct", "grade"),
            headings=("Subject", "Total Marks", "Obtained Marks", "%", "Grade"),
            widths=(150, 90, 100, 60, 60)
        )
        frame.pack(fill="both", expand=True, pady=10)
        for m in result["marks"]:
            pct = round(m["obtained_marks"] / m["total_marks"] * 100, 1) if m["total_marks"] else 0
            tree.insert("", "end", values=(m["subject"], m["total_marks"], m["obtained_marks"], pct,
                                            calculate_grade(pct)))

        bottom = ttk.Frame(wrapper)
        bottom.pack(fill="x", pady=10)
        ttk.Label(bottom, text=f"Total Marks: {result['total']}").pack(anchor="w")
        ttk.Label(bottom, text=f"Obtained Marks: {result['obtained']}").pack(anchor="w")
        ttk.Label(bottom, text=f"Overall Percentage: {result['percentage']}%").pack(anchor="w")
        ttk.Label(bottom, text=f"Overall Grade: {result['grade']}").pack(anchor="w")
        ttk.Label(bottom, text=f"Result: {result['result']}",
                  font=("Segoe UI", 11, "bold")).pack(anchor="w", pady=5)

        sig = ttk.Frame(wrapper)
        sig.pack(fill="x", pady=15)
        ttk.Label(sig, text="____________________\nClass Teacher").pack(side="left", padx=20)
        ttk.Label(sig, text="____________________\nPrincipal / Admin").pack(side="right", padx=20)

        btns = ttk.Frame(wrapper)
        btns.pack(pady=10)
        ttk.Button(btns, text="Save PDF", command=lambda: self._save_pdf(s, result, school_name)).pack(side="left", padx=5)
        ttk.Button(btns, text="Close", command=self.destroy).pack(side="left", padx=5)

    def _save_pdf(self, s, result, school_name):
        if not REPORTLAB_AVAILABLE:
            messagebox.showwarning("Missing Package",
                                    "PDF export requires ReportLab.\nInstall it using:\npip install reportlab",
                                    parent=self)
            return
        try:
            path = generate_result_card_pdf(s, result, school_name)
            messagebox.showinfo("Saved", f"Result card saved to:\n{path}", parent=self)
        except Exception as exc:
            show_exception("generating PDF result card", exc)


# =====================================================================
# REPORTS PANEL
# =====================================================================
class ReportsPanel:
    def __init__(self, parent, session: Session):
        self.parent = parent
        self.session = session
        self._build()

    def _build(self):
        top = ttk.Frame(self.parent)
        top.pack(fill="x", pady=(0, 10))
        ttk.Label(top, text="Attendance Reports", font=("Segoe UI", 14, "bold")).pack(side="left")

        ttk.Label(top, text="Class:").pack(side="left", padx=(20, 3))
        classes = self.session.list_classes()
        self.class_map = {"All": None}
        labels = ["All"]
        for c in classes:
            label = f"{c['class_name']}-{c['section']}" if c["section"] else c["class_name"]
            labels.append(label)
            self.class_map[label] = c["id"]
        self.class_var = tk.StringVar(value="All")
        ttk.Combobox(top, textvariable=self.class_var, values=labels, state="readonly", width=18).pack(side="left")

        ttk.Label(top, text="From:").pack(side="left", padx=(15, 3))
        self.from_var = tk.StringVar()
        ttk.Entry(top, textvariable=self.from_var, width=12).pack(side="left")
        ttk.Label(top, text="To:").pack(side="left", padx=(8, 3))
        self.to_var = tk.StringVar()
        ttk.Entry(top, textvariable=self.to_var, width=12).pack(side="left")

        ttk.Button(top, text="View Report", command=self._load).pack(side="left", padx=10)
        user_role = self.session.current_user["role"]
        if OPENPYXL_AVAILABLE and (has_permission(user_role, "export_excel") or has_permission(user_role, "export_excel_assigned")):
            ttk.Button(top, text="Export Excel", command=self._export_excel).pack(side="right", padx=3)

        frame, tree = make_scrollable_treeview(
            self.parent,
            columns=("student_id", "roll", "name", "total", "present", "absent", "leave", "pct"),
            headings=("Student ID", "Roll No", "Name", "Total Days", "Present", "Absent", "Leave", "Attendance %"),
            widths=(90, 70, 160, 80, 70, 70, 70, 100)
        )
        frame.pack(fill="both", expand=True, pady=10)
        self.tree = tree
        self._last_report = []

    def _load(self):
        for i in self.tree.get_children():
            self.tree.delete(i)
        class_label = self.class_var.get()
        class_id = self.class_map.get(class_label)
        date_from = self.from_var.get().strip() or None
        date_to = self.to_var.get().strip() or None
        try:
            report = self.session.attendance_report(class_id=class_id, date_from=date_from, date_to=date_to)
        except Exception as exc:
            show_exception("generating report", exc)
            return
        self._last_report = report
        for r in report:
            self.tree.insert("", "end", values=(
                r["student_id"], r["roll_number"] or "", r["name"], r["total_days"],
                r["present"], r["absent"], r["leave"], f"{r['percentage']}%"
            ))

    def _export_excel(self):
        if not OPENPYXL_AVAILABLE:
            messagebox.showwarning("Missing Package",
                                    "Excel export requires openpyxl.\nInstall it using:\npip install openpyxl")
            return
        if not self._last_report:
            messagebox.showwarning("No Data", "Please generate a report first (click View Report).")
            return
        try:
            path = export_attendance_report_to_excel(self._last_report)
            messagebox.showinfo("Exported", f"Report exported to:\n{path}")
        except Exception as exc:
            show_exception("exporting report to Excel", exc)


# =====================================================================
# EXCEL EXPORT HELPERS
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

    os.makedirs(EXPORTS_DIR, exist_ok=True)
    filename = os.path.join(EXPORTS_DIR, f"Students_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx")
    wb.save(filename)
    return os.path.abspath(filename)


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

    os.makedirs(EXPORTS_DIR, exist_ok=True)
    filename = os.path.join(EXPORTS_DIR, f"Attendance_Report_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx")
    wb.save(filename)
    return os.path.abspath(filename)


# =====================================================================
# PDF RESULT CARD (ReportLab)
# =====================================================================
def generate_result_card_pdf(student_row, result, school_name):
    os.makedirs(EXPORTS_DIR, exist_ok=True)
    filename = os.path.join(EXPORTS_DIR, f"Result_{student_row['student_id']}_{datetime.now().year}.pdf")

    doc = SimpleDocTemplate(filename, pagesize=A4, topMargin=25 * mm, bottomMargin=20 * mm)
    styles = getSampleStyleSheet()
    title_style = ParagraphStyle("TitleCenter", parent=styles["Title"], alignment=TA_CENTER, fontSize=16)
    sub_style = ParagraphStyle("SubCenter", parent=styles["Normal"], alignment=TA_CENTER, fontSize=11)

    elements = [Paragraph(school_name, title_style), Paragraph("Result Card", sub_style), Spacer(1, 10)]

    if PIL_AVAILABLE and student_row["photo_path"] and os.path.exists(student_row["photo_path"]):
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
    return os.path.abspath(filename)


# =====================================================================
# 14. STARTUP / APP CONTROLLER
# =====================================================================
class AppController:
    """
    Owns the single Tk root window and walks through the fixed
    startup sequence:

        DB init -> check activation -> Activation Window (if needed)
        -> Login Window -> (root.deiconify()) -> MainApplication

    The root window is created ONCE and reused for the whole app
    lifetime; it is withdrawn during activation/login and explicitly
    deiconified again right before the dashboard is built, so the
    "hidden main window after login" bug cannot happen.
    """

    def __init__(self):
        self.root = tk.Tk()
        self.root.title(APP_NAME)
        self.root.geometry("1050x650")
        self.root.minsize(900, 550)
        self._configure_style()

        try:
            self.db = DatabaseManager()
        except Exception as exc:
            show_exception("initializing the database", exc)
            self.root.destroy()
            sys.exit(1)

        self.session = Session(self.db)

        # Root stays hidden until a user is fully logged in.
        self.root.withdraw()

        if self.db.get_setting("activated", "0") != "1":
            ActivationWindow(self.root, self.db, self._show_login)
        else:
            self._show_login()

    def _configure_style(self):
        style = ttk.Style()
        try:
            style.theme_use("clam")
        except tk.TclError:
            pass
        style.configure(".", font=("Segoe UI", 10))
        style.configure("TButton", padding=6)

    def _show_login(self):
        LoginWindow(self.root, self.session, self._show_dashboard)

    def _show_dashboard(self):
        # CRITICAL: make the main window visible again before building
        # the dashboard inside it (see section 11/29 requirements).
        self.root.deiconify()
        self.root.state("normal")
        self.root.lift()
        self.root.focus_force()
        MainApplication(self.root, self.session, self._on_logout)

    def _on_logout(self):
        self.root.withdraw()
        self._show_login()

    def run(self):
        self.root.mainloop()


# =====================================================================
# ENTRY POINT
# =====================================================================
if __name__ == "__main__":
    try:
        app = AppController()
        app.run()
    except Exception as e:
        log_error_to_file("application startup", e)
        try:
            messagebox.showerror("Fatal Error", f"The application failed to start:\n{e}")
        except Exception:
            print("Fatal error:", e)
        sys.exit(1)
