"""
=====================================================================
 RJK Student Management System - Streamlit Web Edition (Complete)
 Converted from Tkinter Desktop App
 Developer/Owner: Rai Jahanzaib

 FIRST RUN CREDENTIALS (change these after logging in!):
     Admin:      admin / admin123
     Management: management / management123
     Teacher:    teacher / teacher123
     Staff:      staff / staff123

 ACTIVATION KEY:
     RJK-278-LEF-1185

 Run with:
     streamlit run streamlit_rjk_sms_complete.py

 Requires Python 3.8+ and:
     pip install streamlit psycopg2-binary openpyxl reportlab
 
 DATABASE: Uses Supabase (Postgres). Set SUPABASE_DB_URL in secrets.
=====================================================================
"""

import os
import io
import shutil
import hashlib
import binascii
import traceback
import uuid
import tempfile
import zipfile
from datetime import datetime, date
from typing import Optional, Dict, List, Any

import streamlit as st
import psycopg2
import psycopg2.extras

# Optional packages
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
    from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
    from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
    from reportlab.lib.enums import TA_CENTER
    REPORTLAB_AVAILABLE = True
except ImportError:
    REPORTLAB_AVAILABLE = False


# =====================================================================
# CONSTANTS & CONFIGURATION
# =====================================================================
APP_NAME = "RJK Student Management System"
APP_OWNER = "Rai Jahanzaib"
ACTIVATION_KEY = "RJK-278-LEF-1185"

ROLES = ["admin", "management", "teacher", "staff"]
ATTENDANCE_STATUSES = ["Present", "Absent", "Leave"]

GRADE_BOUNDARIES = [
    (90, "A+"), (80, "A"), (70, "B"), (60, "C"),
    (50, "D"), (40, "E"), (0, "F"),
]
PASS_PERCENTAGE = 40.0

PERMISSIONS = {
    "admin": {
        "view_dashboard", "manage_users_add", "manage_users_edit", "manage_users_delete",
        "manage_users_disable", "manage_users_reset_pw", "manage_classes_add", "manage_classes_edit",
        "manage_classes_delete", "manage_classes_view", "assign_teacher", "remove_teacher_assignment",
        "students_view_all", "students_add", "students_edit", "students_delete",
        "attendance_view", "attendance_edit", "marks_view", "marks_edit",
        "reports_view", "export_excel", "export_pdf", "view_audit_log",
        "manage_settings", "manage_subject_assignments",
    },
    "management": {
        "view_dashboard", "manage_users_add_teacher", "manage_users_edit_teacher",
        "manage_classes_view", "assign_teacher_limited", "students_view_all", "students_add",
        "students_edit", "attendance_view", "attendance_edit", "marks_view", "marks_edit",
        "reports_view", "export_excel", "export_pdf", "manage_subject_assignments",
    },
    "teacher": {
        "view_dashboard", "manage_classes_view_assigned", "students_view_assigned",
        "students_add_assigned", "students_edit_assigned", "attendance_view_assigned",
        "attendance_edit_assigned", "marks_view_assigned", "marks_edit_assigned",
        "reports_view_assigned", "export_excel_assigned", "export_pdf_assigned",
    },
    "staff": {
        "view_dashboard", "manage_classes_view", "students_view_all",
        "attendance_view", "marks_view", "reports_view_limited", "export_limited",
    },
}


def calculate_grade(percentage):
    for boundary, grade in GRADE_BOUNDARIES:
        if percentage >= boundary:
            return grade
    return "F"


def has_permission(role, permission):
    return permission in PERMISSIONS.get(role, set())


class PermissionError_(Exception):
    pass


# =====================================================================
# DATABASE MANAGER
# =====================================================================
class DatabaseManager:
    @staticmethod
    def _translate(sql):
        return sql.replace("?", "%s")

    def __init__(self):
        db_url = st.secrets.get("SUPABASE_DB_URL")
        if not db_url:
            st.error(
                "Missing SUPABASE_DB_URL in Streamlit secrets.\n\n"
                "Add to .streamlit/secrets.toml (local) or Settings → Secrets (Streamlit Cloud):\n"
                "`SUPABASE_DB_URL = \"postgresql://...\"`"
            )
            st.stop()
        self.conn = psycopg2.connect(db_url, cursor_factory=psycopg2.extras.RealDictCursor)
        self.conn.autocommit = False
        self._create_tables()
        self._create_defaults_if_empty()

    def _create_tables(self):
        cur = self.conn.cursor()
        cur.execute("""
        CREATE TABLE IF NOT EXISTS users (
            id SERIAL PRIMARY KEY, username TEXT UNIQUE NOT NULL, password_hash TEXT NOT NULL,
            salt TEXT NOT NULL, full_name TEXT NOT NULL,
            role TEXT NOT NULL CHECK(role IN ('admin','management','teacher','staff')),
            email TEXT, phone TEXT, is_active INTEGER NOT NULL DEFAULT 1,
            created_at TEXT NOT NULL, updated_at TEXT NOT NULL
        );
        CREATE TABLE IF NOT EXISTS classes (
            id SERIAL PRIMARY KEY, class_name TEXT NOT NULL, section TEXT,
            academic_year TEXT, is_active INTEGER NOT NULL DEFAULT 1, created_at TEXT NOT NULL,
            UNIQUE(class_name, section, academic_year)
        );
        CREATE TABLE IF NOT EXISTS teacher_classes (
            id SERIAL PRIMARY KEY, teacher_id INTEGER NOT NULL, class_id INTEGER NOT NULL,
            FOREIGN KEY(teacher_id) REFERENCES users(id) ON DELETE CASCADE,
            FOREIGN KEY(class_id) REFERENCES classes(id) ON DELETE CASCADE,
            UNIQUE(teacher_id, class_id)
        );
        CREATE TABLE IF NOT EXISTS subject_assignments (
            id SERIAL PRIMARY KEY, class_id INTEGER NOT NULL, subject TEXT NOT NULL COLLATE NOCASE,
            teacher_id INTEGER NOT NULL, created_at TEXT NOT NULL,
            FOREIGN KEY(class_id) REFERENCES classes(id) ON DELETE CASCADE,
            FOREIGN KEY(teacher_id) REFERENCES users(id) ON DELETE CASCADE,
            UNIQUE(class_id, subject, teacher_id)
        );
        CREATE TABLE IF NOT EXISTS students (
            id SERIAL PRIMARY KEY, student_id TEXT UNIQUE NOT NULL, roll_number TEXT,
            first_name TEXT NOT NULL, last_name TEXT, father_name TEXT, date_of_birth TEXT,
            gender TEXT, phone TEXT, address TEXT, class_id INTEGER, photo_path TEXT,
            admission_date TEXT, status TEXT NOT NULL DEFAULT 'Active',
            created_at TEXT NOT NULL, updated_at TEXT NOT NULL,
            FOREIGN KEY(class_id) REFERENCES classes(id) ON DELETE SET NULL
        );
        CREATE TABLE IF NOT EXISTS attendance (
            id SERIAL PRIMARY KEY, student_id INTEGER NOT NULL, attendance_date TEXT NOT NULL,
            status TEXT NOT NULL CHECK(status IN ('Present','Absent','Leave')),
            FOREIGN KEY(student_id) REFERENCES students(id) ON DELETE CASCADE,
            UNIQUE(student_id, attendance_date)
        );
        CREATE TABLE IF NOT EXISTS marks (
            id SERIAL PRIMARY KEY, student_id INTEGER NOT NULL, subject TEXT NOT NULL,
            exam_name TEXT NOT NULL, total_marks REAL NOT NULL, obtained_marks REAL NOT NULL,
            created_at TEXT NOT NULL, updated_at TEXT NOT NULL,
            FOREIGN KEY(student_id) REFERENCES students(id) ON DELETE CASCADE
        );
        CREATE TABLE IF NOT EXISTS settings (
            key TEXT PRIMARY KEY, value TEXT
        );
        CREATE TABLE IF NOT EXISTS audit_log (
            id SERIAL PRIMARY KEY, user_id INTEGER, action TEXT NOT NULL,
            description TEXT, timestamp TEXT NOT NULL
        );
        """)
        self.conn.commit()

    def _create_defaults_if_empty(self):
        cur = self.conn.cursor()
        cur.execute("SELECT COUNT(*) AS c FROM users")
        if cur.fetchone()["c"] > 0:
            return
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
                """INSERT INTO users(username, password_hash, salt, full_name, role, email,
                   phone, is_active, created_at, updated_at)
                   VALUES (%s, %s, %s, %s, %s, '', '', 1, %s, %s)""",
                (username, pwd_hash, salt, full_name, role, now, now)
            )
        for key, val in [("activated", "0"), ("school_name", "RJK School / College"),
                         ("school_address", ""), ("school_phone", "")]:
            cur.execute("INSERT INTO settings (key, value) VALUES (%s, %s) ON CONFLICT DO NOTHING",
                       (key, val))
        self.conn.commit()

    def query(self, sql, params=()):
        sql = self._translate(sql)
        cur = self.conn.cursor()
        cur.execute(sql, params)
        return cur.fetchall()

    def query_one(self, sql, params=()):
        sql = self._translate(sql)
        cur = self.conn.cursor()
        cur.execute(sql, params)
        return cur.fetchone()

    def execute(self, sql, params=()):
        sql = self._translate(sql)
        cur = self.conn.cursor()
        wants_id = (
            sql.strip().upper().startswith("INSERT") and "RETURNING" not in sql.upper()
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
        row = self.query_one("SELECT value FROM settings WHERE key=%s", (key,))
        return row["value"] if row else default

    def set_setting(self, key, value):
        self.execute(
            "INSERT INTO settings (key, value) VALUES (%s, %s) "
            "ON CONFLICT(key) DO UPDATE SET value=excluded.value",
            (key, value)
        )


def hash_password(password, salt=None):
    if salt is None:
        salt = binascii.hexlify(os.urandom(16)).decode("utf-8")
    dk = hashlib.pbkdf2_hmac("sha256", password.encode("utf-8"), salt.encode("utf-8"), 100_000)
    return binascii.hexlify(dk).decode("utf-8"), salt


def verify_password(password, stored_hash, salt):
    test_hash, _ = hash_password(password, salt)
    return test_hash == stored_hash


# =====================================================================
# SESSION MANAGER
# =====================================================================
class Session:
    def __init__(self, db: DatabaseManager):
        self.db = db
        self.current_user = None

    def login(self, username, password):
        user = self.db.query_one("SELECT * FROM users WHERE username = %s", (username,))
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
            "INSERT INTO audit_log (user_id, action, description, timestamp) VALUES (%s, %s, %s, %s)",
            (uid, action, description, datetime.now().isoformat(timespec="seconds"))
        )

    def require(self, permission):
        if self.current_user is None:
            raise PermissionError_("Not logged in.")
        if not has_permission(self.current_user["role"], permission):
            raise PermissionError_(f"Role '{self.current_user['role']}' lacks permission '{permission}'.")

    def teacher_owns_class(self, class_id):
        if self.current_user["role"] != "teacher":
            return False
        row = self.db.query_one(
            "SELECT 1 FROM teacher_classes WHERE teacher_id=%s AND class_id=%s",
            (self.current_user["id"], class_id)
        )
        return row is not None

    def teacher_owns_student(self, student_id):
        row = self.db.query_one("SELECT s.class_id FROM students s WHERE s.id=%s", (student_id,))
        if row is None or row["class_id"] is None:
            return False
        return self.teacher_owns_class(row["class_id"])

    # ===== USERS =====
    def list_users(self):
        return self.db.query("SELECT * FROM users ORDER BY role, username")

    def add_user(self, username, password, full_name, role, email, phone):
        if self.current_user["role"] == "management":
            self.require("manage_users_add_teacher")
            if role != "teacher":
                raise PermissionError_("Management can only create Teacher accounts.")
        else:
            self.require("manage_users_add")
        if role not in ROLES:
            raise ValueError("Invalid role.")
        existing = self.db.query_one("SELECT id FROM users WHERE username=%s", (username,))
        if existing:
            raise ValueError(f"Username '{username}' already exists.")
        pwd_hash, salt = hash_password(password)
        now = datetime.now().isoformat(timespec="seconds")
        new_id = self.db.execute(
            """INSERT INTO users(username, password_hash, salt, full_name, role, email, phone, is_active, created_at, updated_at)
               VALUES (%s, %s, %s, %s, %s, %s, %s, 1, %s, %s)""",
            (username, pwd_hash, salt, full_name, role, email, phone, now, now)
        )
        self.log_audit("add_user", f"Added user '{username}' with role '{role}'.")
        return new_id

    def edit_user(self, user_id, full_name, email, phone, role=None):
        target = self.db.query_one("SELECT * FROM users WHERE id=%s", (user_id,))
        if target is None:
            raise ValueError("User not found.")
        if self.current_user["role"] == "management":
            self.require("manage_users_edit_teacher")
            if target["role"] != "teacher":
                raise PermissionError_("Management may only edit Teacher accounts.")
            role = "teacher"
        else:
            self.require("manage_users_edit")
        final_role = role if role else target["role"]
        now = datetime.now().isoformat(timespec="seconds")
        self.db.execute(
            "UPDATE users SET full_name=%s, email=%s, phone=%s, role=%s, updated_at=%s WHERE id=%s",
            (full_name, email, phone, final_role, now, user_id)
        )
        self.log_audit("edit_user", f"Edited user id={user_id}.")

    def set_user_active(self, user_id, active):
        self.require("manage_users_disable")
        target = self.db.query_one("SELECT * FROM users WHERE id=%s", (user_id,))
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
            "UPDATE users SET is_active=%s, updated_at=%s WHERE id=%s",
            (1 if active else 0, now, user_id)
        )
        self.log_audit("disable_user" if not active else "enable_user",
                       f"User id={user_id} set active={active}.")

    def delete_user(self, user_id):
        self.require("manage_users_delete")
        target = self.db.query_one("SELECT * FROM users WHERE id=%s", (user_id,))
        if target is None:
            raise ValueError("User not found.")
        if target["id"] == self.current_user["id"]:
            raise PermissionError_("You cannot delete your own account while logged in.")
        if target["role"] == "admin":
            admin_count = self.db.query_one("SELECT COUNT(*) c FROM users WHERE role='admin'")["c"]
            if admin_count <= 1:
                raise PermissionError_("Cannot delete the last remaining Admin account.")
        self.db.execute("DELETE FROM users WHERE id=%s", (user_id,))
        self.log_audit("delete_user", f"Deleted user '{target['username']}'.")

    def reset_password(self, user_id, new_password):
        self.require("manage_users_reset_pw")
        pwd_hash, salt = hash_password(new_password)
        now = datetime.now().isoformat(timespec="seconds")
        self.db.execute(
            "UPDATE users SET password_hash=%s, salt=%s, updated_at=%s WHERE id=%s",
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
            "UPDATE users SET password_hash=%s, salt=%s, updated_at=%s WHERE id=%s",
            (pwd_hash, salt, now, user["id"])
        )
        self.current_user = self.db.query_one("SELECT * FROM users WHERE id=%s", (user["id"],))
        self.log_audit("change_password", "User changed their own password.")

    # ===== CLASSES =====
    def list_classes(self):
        role = self.current_user["role"]
        if role in ("admin", "management", "staff"):
            return self.db.query("SELECT * FROM classes ORDER BY class_name, section")
        elif role == "teacher":
            return self.db.query(
                """SELECT c.* FROM classes c
                   JOIN teacher_classes tc ON tc.class_id = c.id
                   WHERE tc.teacher_id = %s ORDER BY c.class_name, c.section""",
                (self.current_user["id"],)
            )
        return []

    def add_class(self, class_name, section, academic_year):
        self.require("manage_classes_add")
        now = datetime.now().isoformat(timespec="seconds")
        try:
            return self.db.execute(
                """INSERT INTO classes(class_name, section, academic_year, is_active, created_at)
                   VALUES (%s, %s, %s, 1, %s)""",
                (class_name, section, academic_year, now)
            )
        except psycopg2.IntegrityError:
            raise ValueError("This class/section/year combination already exists.")

    def edit_class(self, class_id, class_name, section, academic_year, is_active):
        self.require("manage_classes_edit")
        self.db.execute(
            "UPDATE classes SET class_name=%s, section=%s, academic_year=%s, is_active=%s WHERE id=%s",
            (class_name, section, academic_year, 1 if is_active else 0, class_id)
        )
        self.log_audit("edit_class", f"Edited class id={class_id}.")

    def delete_class(self, class_id):
        self.require("manage_classes_delete")
        student_count = self.db.query_one(
            "SELECT COUNT(*) c FROM students WHERE class_id=%s", (class_id,)
        )["c"]
        if student_count > 0:
            raise PermissionError_(f"Cannot delete: {student_count} student(s) are assigned to this class.")
        self.db.execute("DELETE FROM classes WHERE id=%s", (class_id,))
        self.log_audit("delete_class", f"Deleted class id={class_id}.")

    # ===== TEACHERS & ASSIGNMENTS =====
    def list_teachers(self):
        return self.db.query("SELECT * FROM users WHERE role='teacher' ORDER BY full_name")

    def list_assignments_for_teacher(self, teacher_id):
        return self.db.query(
            """SELECT c.* FROM classes c
               JOIN teacher_classes tc ON tc.class_id = c.id
               WHERE tc.teacher_id=%s ORDER BY c.class_name, c.section""",
            (teacher_id,)
        )

    def assign_teacher(self, teacher_id, class_id):
        role = self.current_user["role"]
        if role == "management":
            self.require("assign_teacher_limited")
        else:
            self.require("assign_teacher")
        teacher = self.db.query_one("SELECT * FROM users WHERE id=%s AND role='teacher'", (teacher_id,))
        if teacher is None:
            raise ValueError("Selected user is not a valid teacher.")
        try:
            self.db.execute(
                "INSERT INTO teacher_classes(teacher_id, class_id) VALUES (%s, %s)",
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
            "DELETE FROM subject_assignments WHERE teacher_id=%s AND class_id=%s",
            (teacher_id, class_id)
        )
        self.db.execute(
            "DELETE FROM teacher_classes WHERE teacher_id=%s AND class_id=%s",
            (teacher_id, class_id)
        )
        self.log_audit("remove_teacher_assignment", f"Removed teacher id={teacher_id} from class id={class_id}.")

    # ===== SUBJECTS =====
    def list_subject_assignments(self):
        self.require("manage_subject_assignments")
        return self.db.query("""
            SELECT sa.id, sa.class_id, sa.subject, sa.teacher_id,
                   c.class_name, c.section, u.full_name AS teacher_name
            FROM subject_assignments sa
            JOIN classes c ON c.id=sa.class_id
            JOIN users u ON u.id=sa.teacher_id
            ORDER BY c.class_name, c.section, sa.subject, u.full_name
        """)

    def assign_subject(self, class_id, subject, teacher_id):
        self.require("manage_subject_assignments")
        subject = (subject or "").strip()
        if not subject:
            raise ValueError("Subject name is required.")
        if not self.db.query_one(
            "SELECT id FROM users WHERE id=%s AND role='teacher' AND is_active=1", (teacher_id,)
        ):
            raise ValueError("Select an active teacher.")
        if not self.db.query_one(
            "SELECT 1 FROM teacher_classes WHERE teacher_id=%s AND class_id=%s", (teacher_id, class_id)
        ):
            raise ValueError("Assign this teacher to the class first.")
        try:
            self.db.execute(
                "INSERT INTO subject_assignments(class_id, subject, teacher_id, created_at) VALUES(%s, %s, %s, %s)",
                (class_id, subject, teacher_id, datetime.now().isoformat(timespec="seconds"))
            )
        except psycopg2.IntegrityError:
            raise ValueError("This subject is already assigned to that teacher for this class.")
        self.log_audit("assign_subject", f"Assigned {subject} in class id={class_id} to teacher id={teacher_id}.")

    def remove_subject_assignment(self, assignment_id):
        self.require("manage_subject_assignments")
        self.db.execute("DELETE FROM subject_assignments WHERE id=%s", (assignment_id,))
        self.log_audit("remove_subject_assignment", f"Removed subject assignment id={assignment_id}.")

    def list_subjects_for_class(self, class_id):
        if self.current_user["role"] == "teacher":
            self.require("marks_view_assigned")
            if not self.teacher_owns_class(class_id):
                raise PermissionError_("You are not assigned to this class.")
            rows = self.db.query(
                "SELECT DISTINCT subject FROM subject_assignments WHERE class_id=%s AND teacher_id=%s ORDER BY subject",
                (class_id, self.current_user["id"])
            )
        else:
            self.require("marks_view")
            rows = self.db.query(
                "SELECT DISTINCT subject FROM subject_assignments WHERE class_id=%s ORDER BY subject",
                (class_id,)
            )
        return [r["subject"] for r in rows]

    def load_class_marks(self, class_id, subject, exam_name):
        role = self.current_user["role"]
        if role == "teacher":
            self.require("marks_view_assigned")
            if not self.teacher_owns_class(class_id):
                raise PermissionError_("You are not assigned to this class.")
            ok = self.db.query_one(
                "SELECT 1 FROM subject_assignments WHERE class_id=%s AND subject=%s AND teacher_id=%s",
                (class_id, subject, self.current_user["id"])
            )
            if not ok:
                raise PermissionError_("You are not assigned to this subject.")
        else:
            self.require("marks_view")
        return self.db.query("""
            SELECT s.id AS student_row_id, s.student_id, s.roll_number, s.first_name, s.last_name,
                   (SELECT m.obtained_marks FROM marks m
                    WHERE m.student_id=s.id AND m.subject=%s AND m.exam_name=%s
                    ORDER BY m.id DESC LIMIT 1) AS obtained_marks
            FROM students s
            WHERE s.class_id=%s AND s.status='Active'
            ORDER BY CAST(COALESCE(s.roll_number, '0') AS INTEGER), s.roll_number, s.first_name
        """, (subject, exam_name, class_id))

    def save_marks_batch(self, class_id, subject, exam_name, total_marks, marks_by_student):
        role = self.current_user["role"]
        if role == "teacher":
            self.require("marks_edit_assigned")
            if not self.teacher_owns_class(class_id):
                raise PermissionError_("You are not assigned to this class.")
            ok = self.db.query_one(
                "SELECT 1 FROM subject_assignments WHERE class_id=%s AND subject=%s AND teacher_id=%s",
                (class_id, subject, self.current_user["id"])
            )
            if not ok:
                raise PermissionError_("You are not assigned to this subject.")
        else:
            self.require("marks_edit")
        subject = (subject or "").strip()
        exam_name = (exam_name or "").strip()
        total = float(total_marks)
        if not subject or not exam_name or total <= 0:
            raise ValueError("Subject, exam name, and positive total marks are required.")
        parsed = {}
        for sid, value in marks_by_student.items():
            sid = int(sid)
            if value is None or str(value).strip() == "":
                continue
            val = float(value)
            if val < 0 or val > total:
                raise ValueError(f"Marks for student ID {sid} must be between 0 and {total:g}.")
            parsed[sid] = val
        if not parsed:
            raise ValueError("Enter at least one mark before saving.")
        ids = list(parsed)
        marks_str = ",".join(["%s"] * len(ids))
        valid = self.db.query(
            f"SELECT id FROM students WHERE class_id=%s AND status='Active' AND id IN ({marks_str})",
            [class_id] + ids
        )
        if {r["id"] for r in valid} != set(ids):
            raise PermissionError_("The mark sheet includes a student outside this active class.")
        now = datetime.now().isoformat(timespec="seconds")
        try:
            self.db.conn.execute("BEGIN")
            for sid, val in parsed.items():
                old = self.db.conn.execute(
                    "SELECT id FROM marks WHERE student_id=%s AND subject=%s AND exam_name=%s ORDER BY id DESC LIMIT 1",
                    (sid, subject, exam_name)
                ).fetchone()
                if old:
                    self.db.conn.execute(
                        "UPDATE marks SET total_marks=%s, obtained_marks=%s, updated_at=%s WHERE id=%s",
                        (total, val, now, old["id"])
                    )
                else:
                    self.db.conn.execute(
                        "INSERT INTO marks(student_id, subject, exam_name, total_marks, obtained_marks, created_at, updated_at) VALUES(%s, %s, %s, %s, %s, %s, %s)",
                        (sid, subject, exam_name, total, val, now, now)
                    )
                keep = self.db.conn.execute(
                    "SELECT id FROM marks WHERE student_id=%s AND subject=%s AND exam_name=%s ORDER BY id DESC LIMIT 1",
                    (sid, subject, exam_name)
                ).fetchone()
                self.db.conn.execute(
                    "DELETE FROM marks WHERE student_id=%s AND subject=%s AND exam_name=%s AND id<>%s",
                    (sid, subject, exam_name, keep["id"])
                )
            self.db.conn.commit()
        except Exception:
            self.db.conn.rollback()
            raise
        self.log_audit("save_marks_batch",
                       f"Saved {len(parsed)} marks for {subject}, {exam_name}, class id={class_id}.")

    # ===== STUDENTS =====
    def search_students(self, keyword="", class_id=None, status=None):
        role = self.current_user["role"]
        base = """SELECT s.*, c.class_name, c.section FROM students s
                   LEFT JOIN classes c ON c.id = s.class_id WHERE 1=1"""
        params = []
        if role == "teacher":
            base += """ AND s.class_id IN (SELECT class_id FROM teacher_classes WHERE teacher_id = %s)"""
            params.append(self.current_user["id"])
        if keyword:
            base += """ AND (s.student_id LIKE %s OR s.roll_number LIKE %s
                              OR s.first_name LIKE %s OR s.last_name LIKE %s
                              OR s.father_name LIKE %s OR s.phone LIKE %s)"""
            like = f"%{keyword}%"
            params.extend([like] * 6)
        if class_id:
            base += " AND s.class_id = %s"
            params.append(class_id)
        if status:
            base += " AND s.status = %s"
            params.append(status)
        base += " ORDER BY s.first_name, s.last_name"
        return self.db.query(base, params)

    def get_student(self, student_id):
        role = self.current_user["role"]
        row = self.db.query_one(
            """SELECT s.*, c.class_name, c.section FROM students s
               LEFT JOIN classes c ON c.id = s.class_id WHERE s.id=%s""",
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
        existing = self.db.query_one("SELECT id FROM students WHERE student_id=%s", (data["student_id"],))
        if existing:
            raise ValueError(f"Student ID '{data['student_id']}' already exists.")
        photo_path = None
        if photo_bytes:
            photo_path = self._save_student_photo(data["student_id"], photo_bytes, photo_ext)
        now = datetime.now().isoformat(timespec="seconds")
        new_id = self.db.execute(
            """INSERT INTO students(student_id, roll_number, first_name, last_name, father_name,
               date_of_birth, gender, phone, address, class_id, photo_path, admission_date, status,
               created_at, updated_at)
               VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s)""",
            (data["student_id"], data["roll_number"], data["first_name"], data["last_name"],
             data["father_name"], data["date_of_birth"], data["gender"], data["phone"],
             data["address"], data["class_id"], photo_path, data["admission_date"],
             data.get("status", "Active"), now, now)
        )
        self.log_audit("add_student", f"Added student '{data['student_id']}'.")
        return new_id

    def edit_student(self, student_row_id, data, photo_bytes=None, photo_ext=None):
        role = self.current_user["role"]
        current = self.db.query_one("SELECT * FROM students WHERE id=%s", (student_row_id,))
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
            """UPDATE students SET student_id=%s, roll_number=%s, first_name=%s, last_name=%s,
               father_name=%s, date_of_birth=%s, gender=%s, phone=%s, address=%s, class_id=%s,
               photo_path=%s, admission_date=%s, status=%s, updated_at=%s WHERE id=%s""",
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
        student = self.db.query_one("SELECT * FROM students WHERE id=%s", (student_row_id,))
        if student is None:
            raise ValueError("Student not found.")
        self.db.execute("DELETE FROM students WHERE id=%s", (student_row_id,))
        self.log_audit("delete_student", f"Deleted student '{student['student_id']}'.")

    def _save_student_photo(self, student_id, data_bytes, ext):
        ext = (ext or ".png").lower()
        if ext not in (".jpg", ".jpeg", ".png"):
            raise ValueError("Only JPG, JPEG and PNG images are supported.")
        unique = uuid.uuid4().hex[:8]
        dest_name = f"{student_id}_{unique}{ext}"
        # Streamlit doesn't have persistent file storage - we'd need Supabase Storage for this
        # For now, we're storing the path but not the actual file on cloud deployment
        return f"student_photos/{dest_name}"

    # ===== ATTENDANCE =====
    def load_attendance(self, class_id, att_date):
        role = self.current_user["role"]
        if role == "teacher":
            self.require("attendance_view_assigned")
            if not self.teacher_owns_class(class_id):
                raise PermissionError_("You are not assigned to this class.")
        else:
            self.require("attendance_view")
        rows = self.db.query(
            """SELECT s.id as student_row_id, s.roll_number, s.first_name, s.last_name, a.status
               FROM students s
               LEFT JOIN attendance a ON a.student_id = s.id AND a.attendance_date = %s
               WHERE s.class_id = %s AND s.status='Active'
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
        try:
            datetime.strptime(att_date, "%Y-%m-%d")
        except (TypeError, ValueError):
            raise ValueError("Attendance date must use YYYY-MM-DD format.")
        if not status_by_student_row_id:
            raise ValueError("There are no students to save attendance for.")
        ids = [int(x) for x in status_by_student_row_id]
        marks_str = ",".join(["%s"] * len(ids))
        rows = self.db.query(
            f"SELECT id FROM students WHERE class_id=%s AND status='Active' AND id IN ({marks_str})",
            [class_id] + ids
        )
        if {r["id"] for r in rows} != set(ids):
            raise PermissionError_("Attendance includes a student outside this active class.")
        if any(v not in ATTENDANCE_STATUSES for v in status_by_student_row_id.values()):
            raise ValueError("Choose Present, Absent, or Leave for every student.")
        try:
            self.db.conn.execute("BEGIN")
            for student_id, status in status_by_student_row_id.items():
                self.db.conn.execute(
                    """INSERT INTO attendance (student_id, attendance_date, status) VALUES (%s, %s, %s)
                       ON CONFLICT(student_id, attendance_date) DO UPDATE SET status=excluded.status""",
                    (int(student_id), att_date, status)
                )
            self.db.conn.commit()
        except Exception:
            self.db.conn.rollback()
            raise
        self.log_audit("save_attendance", f"Saved attendance for {len(ids)} students in class id={class_id} on {att_date}.")

    def attendance_report(self, class_id=None, student_row_id=None, date_from=None, date_to=None):
        role = self.current_user["role"]
        if role == "teacher":
            self.require("reports_view_assigned")
        else:
            self.require("reports_view") if role != "staff" else self.require("reports_view_limited")

        base = """SELECT s.id as student_row_id, s.student_id, s.roll_number, s.first_name, s.last_name,
                          SUM(CASE WHEN a.status='Present' THEN 1 ELSE 0 END) as present_days,
                          SUM(CASE WHEN a.status='Absent' THEN 1 ELSE 0 END) as absent_days,
                          SUM(CASE WHEN a.status='Leave' THEN 1 ELSE 0 END) as leave_days,
                          COUNT(a.id) as total_days
                   FROM students s
                   LEFT JOIN attendance a ON a.student_id = s.id"""
        conditions = []
        params = []
        if date_from:
            conditions.append("(a.attendance_date IS NULL OR a.attendance_date >= %s)")
            params.append(date_from)
        if date_to:
            conditions.append("(a.attendance_date IS NULL OR a.attendance_date <= %s)")
            params.append(date_to)
        if role == "teacher":
            conditions.append("s.class_id IN (SELECT class_id FROM teacher_classes WHERE teacher_id=%s)")
            params.append(self.current_user["id"])
        if class_id:
            conditions.append("s.class_id = %s")
            params.append(class_id)
        if student_row_id:
            conditions.append("s.id = %s")
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

    # ===== MARKS =====
    def list_marks(self, student_row_id):
        role = self.current_user["role"]
        if role == "teacher":
            self.require("marks_view_assigned")
            if not self.teacher_owns_student(student_row_id):
                raise PermissionError_("You cannot view marks outside your assigned class(es).")
        else:
            self.require("marks_view")
        return self.db.query(
            "SELECT * FROM marks WHERE student_id=%s ORDER BY exam_name, subject",
            (student_row_id,)
        )

    def result_summary(self, student_row_id):
        marks = self.list_marks(student_row_id)
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

    # ===== DASHBOARD =====
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
                "SELECT COUNT(*) c FROM teacher_classes WHERE teacher_id=%s",
                (self.current_user["id"],)
            )["c"]
            stats["My Students"] = self.db.query_one(
                """SELECT COUNT(*) c FROM students
                   WHERE class_id IN (SELECT class_id FROM teacher_classes WHERE teacher_id=%s)""",
                (self.current_user["id"],)
            )["c"]
        elif role == "staff":
            stats["Total Students"] = self.db.query_one("SELECT COUNT(*) c FROM students")["c"]
            stats["Total Classes"] = self.db.query_one("SELECT COUNT(*) c FROM classes")["c"]
        return stats


# =====================================================================
# STREAMLIT UI
# =====================================================================
@st.cache_resource(show_spinner=False)
def get_db():
    return DatabaseManager()


def get_session():
    if "session" not in st.session_state:
        st.session_state.session = Session(get_db())
    return st.session_state.session


def main():
    st.set_page_config(page_title=APP_NAME, page_icon="🎓", layout="wide", initial_sidebar_state="expanded")

    try:
        db = get_db()
    except Exception as e:
        st.error(f"Database connection failed: {e}")
        st.stop()

    session = get_session()

    # Activation check
    if db.get_setting("activated", "0") != "1":
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
        return

    # Login check
    if session.current_user is None:
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
        return

    # Sidebar navigation
    with st.sidebar:
        st.title(APP_NAME)
        user = session.current_user
        st.caption(f"{user['full_name']} ({user['role'].upper()})")
        st.divider()

        NAV_BY_ROLE = {
            "admin": [
                ("🏠 Dashboard", "dashboard"), ("👥 Students", "students"),
                ("📚 Classes", "classes"), ("👨‍🏫 Teachers", "teachers"),
                ("📖 Subjects", "subjects"), ("📋 Attendance", "attendance"),
                ("📊 Marks", "marks"), ("📈 Reports", "reports"),
                ("🛠️ Data Tools", "data_tools"), ("👤 Users", "users"),
                ("🔍 Audit Log", "audit"),
            ],
            "management": [
                ("🏠 Dashboard", "dashboard"), ("👥 Students", "students"),
                ("👨‍🏫 Teachers", "teachers"), ("📚 Classes", "classes"),
                ("📖 Subjects", "subjects"), ("📋 Attendance", "attendance"),
                ("📊 Marks", "marks"), ("📈 Reports", "reports"),
            ],
            "teacher": [
                ("🏠 Dashboard", "dashboard"), ("📚 My Classes", "classes"),
                ("👥 My Students", "students"), ("📋 Attendance", "attendance"),
                ("📊 Marks", "marks"), ("📈 Reports", "reports"),
            ],
            "staff": [
                ("🏠 Dashboard", "dashboard"), ("👥 Students", "students"),
                ("📈 Reports", "reports"),
            ],
        }

        nav_items = NAV_BY_ROLE.get(user["role"], [])
        if "nav_page" not in st.session_state:
            st.session_state.nav_page = nav_items[0][1] if nav_items else "dashboard"

        for label, key in nav_items:
            if st.button(label, use_container_width=True,
                        type="primary" if st.session_state.nav_page == key else "secondary"):
                st.session_state.nav_page = key
                st.rerun()

        st.divider()

        with st.expander("🔑 Change Password"):
            with st.form("change_pw_form", clear_on_submit=True):
                old_pw = st.text_input("Current Password", type="password")
                new_pw = st.text_input("New Password", type="password")
                confirm_pw = st.text_input("Confirm New Password", type="password")
                if st.form_submit_button("Update Password"):
                    if new_pw != confirm_pw:
                        st.error("New passwords do not match.")
                    elif len(new_pw) < 4:
                        st.error("New password must be at least 4 characters.")
                    else:
                        try:
                            session.change_own_password(old_pw, new_pw)
                            st.success("Password changed successfully.")
                        except Exception as e:
                            st.error(str(e))

        if st.button("🚪 Logout", use_container_width=True, type="secondary"):
            session.logout()
            st.session_state.nav_page = "dashboard"
            st.rerun()

        if st.button("❌ Exit", use_container_width=True):
            st.session_state.clear()
            st.rerun()

    # Main content
    page = st.session_state.get("nav_page", "dashboard")

    try:
        if page == "dashboard":
            render_dashboard(session)
        elif page == "students":
            render_students(session)
        elif page == "classes":
            render_classes(session)
        elif page == "teachers":
            render_teachers(session)
        elif page == "subjects":
            render_subjects(session)
        elif page == "attendance":
            render_attendance(session)
        elif page == "marks":
            render_marks(session)
        elif page == "reports":
            render_reports(session)
        elif page == "data_tools":
            render_data_tools(session)
        elif page == "users":
            render_users(session)
        elif page == "audit":
            render_audit(session)
        else:
            st.error("Page not found.")
    except PermissionError_ as e:
        st.error(f"Permission Denied: {e}")
    except Exception as e:
        st.error(f"Error: {e}")
        with st.expander("Details"):
            st.code(traceback.format_exc())


# =====================================================================
# VIEW RENDERERS
# =====================================================================
def render_dashboard(session):
    user = session.current_user
    st.header(f"👋 Welcome, {user['full_name']}!")
    
    stats = session.dashboard_stats()
    cols = st.columns(len(stats) if stats else 1)
    for col, (label, value) in zip(cols, stats.items()):
        col.metric(label, value)

    if not OPENPYXL_AVAILABLE or not REPORTLAB_AVAILABLE:
        missing = []
        if not OPENPYXL_AVAILABLE:
            missing.append("openpyxl")
        if not REPORTLAB_AVAILABLE:
            missing.append("reportlab")
        st.warning(f"Optional packages missing: {', '.join(missing)}\n\nInstall with: `pip install {' '.join(missing)}`")


def render_classes(session):
    st.header("📚 Classes")
    user = session.current_user
    
    classes = session.list_classes()
    if classes:
        st.subheader("Class List")
        class_data = [{
            "ID": c["id"],
            "Class": c["class_name"],
            "Section": c["section"] or "—",
            "Year": c["academic_year"] or "—",
            "Status": "Active" if c["is_active"] else "Inactive"
        } for c in classes]
        st.dataframe(class_data, use_container_width=True, hide_index=True)
    else:
        st.info("No classes yet.")

    if has_permission(user["role"], "manage_classes_add"):
        with st.expander("➕ Add Class"):
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
                        except Exception as e:
                            st.error(str(e))

    if has_permission(user["role"], "assign_teacher"):
        with st.expander("👨‍🏫 Assign Teacher to Class"):
            teachers = session.list_teachers()
            if not teachers:
                st.info("No teacher accounts exist yet.")
            else:
                teacher_map = {f"{t['full_name']} ({t['username']})": t["id"] for t in teachers}
                teacher_name = st.selectbox("Teacher", list(teacher_map.keys()), key="assign_teacher_pick")
                teacher_id = teacher_map[teacher_name]

                all_classes = session.db.query("SELECT * FROM classes ORDER BY class_name, section")
                assigned_ids = {c["id"] for c in session.list_assignments_for_teacher(teacher_id)}
                class_map = {(f"{c['class_name']}-{c['section']}" if c["section"] else c["class_name"]): c["id"]
                             for c in all_classes}

                unassigned_labels = [l for l, cid in class_map.items() if cid not in assigned_ids]
                assigned_labels = [l for l, cid in class_map.items() if cid in assigned_ids]

                col1, col2 = st.columns(2)
                with col1:
                    st.caption("Available")
                    if unassigned_labels:
                        pick = st.selectbox("Class to assign", unassigned_labels, key="pick_assign")
                        if st.button("➜ Assign"):
                            try:
                                session.assign_teacher(teacher_id, class_map[pick])
                                st.success("Assigned.")
                                st.rerun()
                            except Exception as e:
                                st.error(str(e))
                with col2:
                    st.caption("Currently assigned")
                    if assigned_labels:
                        pick = st.selectbox("Class to remove", assigned_labels, key="pick_remove")
                        if st.button("⬅ Remove"):
                            try:
                                session.remove_teacher_assignment(teacher_id, class_map[pick])
                                st.success("Removed.")
                                st.rerun()
                            except Exception as e:
                                st.error(str(e))


def render_students(session):
    st.header("👥 Students")
    user = session.current_user

    col1, col2, col3 = st.columns(3)
    with col1:
        keyword = st.text_input("Search", key="student_search")
    with col2:
        classes = session.list_classes()
        class_map = {"All": None}
        for c in classes:
            label = f"{c['class_name']}-{c['section']}" if c["section"] else c["class_name"]
            class_map[label] = c["id"]
        class_choice = st.selectbox("Class", list(class_map.keys()), key="student_class")
    with col3:
        status_choice = st.selectbox("Status", ["All", "Active", "Inactive"], key="student_status")

    class_id = class_map.get(class_choice)
    status = None if status_choice == "All" else status_choice

    try:
        students = session.search_students(keyword, class_id, status)
    except Exception as e:
        st.error(f"Error loading students: {e}")
        students = []

    if students:
        student_data = [{
            "ID": s["id"],
            "Student ID": s["student_id"],
            "Roll": s["roll_number"] or "—",
            "Name": f"{s['first_name']} {s['last_name'] or ''}",
            "Father": s["father_name"] or "—",
            "Class": f"{s['class_name']}-{s['section']}" if s["class_name"] and s["section"] else (s["class_name"] or "—"),
            "Phone": s["phone"] or "—",
            "Status": s["status"]
        } for s in students]
        st.dataframe(student_data, use_container_width=True, hide_index=True)
    else:
        st.info("No students found.")

    if has_permission(user["role"], "students_add") or has_permission(user["role"], "students_add_assigned"):
        with st.expander("➕ Add Student"):
            _student_form(session, "add")

    if students and (has_permission(user["role"], "students_edit") or has_permission(user["role"], "students_edit_assigned")):
        with st.expander("✏️ Edit Student"):
            student_map = {f"{s['student_id']} - {s['first_name']} {s['last_name'] or ''}": s["id"]
                          for s in students}
            pick = st.selectbox("Select Student", list(student_map.keys()), key="edit_student")
            _student_form(session, "edit", student_map[pick])

    if has_permission(user["role"], "students_delete") and students:
        with st.expander("🗑️ Delete Student"):
            student_map = {f"{s['student_id']} - {s['first_name']} {s['last_name'] or ''}": s["id"]
                          for s in students}
            pick = st.selectbox("Select Student to delete", list(student_map.keys()), key="delete_student")
            confirm = st.checkbox("I understand this cannot be undone.")
            if st.button("Delete", type="secondary", disabled=not confirm):
                try:
                    session.delete_student(student_map[pick])
                    st.success("Student deleted.")
                    st.rerun()
                except Exception as e:
                    st.error(str(e))


def _student_form(session, mode, student_id=None):
    classes = session.list_classes()
    if not classes:
        st.warning("Please add/be assigned a class before adding students.")
        return

    labels = [f"{c['class_name']}-{c['section']}" if c["section"] else c["class_name"] for c in classes]
    class_map = {labels[i]: classes[i]["id"] for i in range(len(classes))}

    if mode == "edit" and student_id:
        try:
            s = session.get_student(student_id)
            current_label = next((l for l, cid in class_map.items() if cid == s["class_id"]), labels[0])
            initial = {
                "student_id": s["student_id"], "roll_number": s["roll_number"] or "",
                "first_name": s["first_name"], "last_name": s["last_name"] or "",
                "father_name": s["father_name"] or "", "date_of_birth": s["date_of_birth"] or "",
                "gender": s["gender"] or "Male", "phone": s["phone"] or "",
                "address": s["address"] or "", "class_label": current_label,
                "admission_date": s["admission_date"] or "", "status": s["status"],
            }
        except Exception as e:
            st.error(str(e))
            return
    else:
        initial = {
            "student_id": "", "roll_number": "", "first_name": "", "last_name": "",
            "father_name": "", "date_of_birth": date.today().isoformat(), "gender": "Male",
            "phone": "", "address": "", "class_label": labels[0],
            "admission_date": date.today().isoformat(), "status": "Active",
        }

    form_key = f"student_{mode}_{student_id or 'new'}"
    with st.form(form_key):
        col1, col2 = st.columns(2)
        with col1:
            student_id_val = st.text_input("Student ID", value=initial["student_id"], key=f"{form_key}_id")
            first_name = st.text_input("First Name", value=initial["first_name"], key=f"{form_key}_first")
            father_name = st.text_input("Father Name", value=initial["father_name"], key=f"{form_key}_father")
            dob = st.text_input("Date of Birth (YYYY-MM-DD)", value=initial["date_of_birth"], key=f"{form_key}_dob")
            phone = st.text_input("Phone", value=initial["phone"], key=f"{form_key}_phone")
        with col2:
            roll_number = st.text_input("Roll Number", value=initial["roll_number"], key=f"{form_key}_roll")
            last_name = st.text_input("Last Name", value=initial["last_name"], key=f"{form_key}_last")
            gender = st.selectbox("Gender", ["Male", "Female", "Other"],
                                 index=["Male", "Female", "Other"].index(initial["gender"]), key=f"{form_key}_gender")
            admission_date = st.text_input("Admission Date (YYYY-MM-DD)", value=initial["admission_date"], key=f"{form_key}_adm")
            status = st.selectbox("Status", ["Active", "Inactive"],
                                 index=["Active", "Inactive"].index(initial["status"]), key=f"{form_key}_status")

        class_label = st.selectbox("Class", labels, index=labels.index(initial["class_label"]), key=f"{form_key}_class")
        address = st.text_area("Address", value=initial["address"], key=f"{form_key}_address")

        if st.form_submit_button("Save Student", type="primary"):
            if not student_id_val.strip() or not first_name.strip():
                st.error("Student ID and First Name are required.")
            else:
                data = {
                    "student_id": student_id_val.strip(), "roll_number": roll_number.strip(),
                    "first_name": first_name.strip(), "last_name": last_name.strip(),
                    "father_name": father_name.strip(), "date_of_birth": dob.strip(),
                    "gender": gender, "phone": phone.strip(), "address": address.strip(),
                    "class_id": class_map[class_label], "admission_date": admission_date.strip(),
                    "status": status,
                }
                try:
                    if mode == "add":
                        session.add_student(data)
                        st.success("Student added.")
                    else:
                        session.edit_student(student_id, data)
                        st.success("Student updated.")
                    st.rerun()
                except Exception as e:
                    st.error(str(e))


def render_teachers(session):
    st.header("👨‍🏫 Teachers")
    user = session.current_user

    teachers = session.list_teachers()
    if teachers:
        teacher_data = []
        for t in teachers:
            classes = session.list_assignments_for_teacher(t["id"])
            class_str = ", ".join(f"{c['class_name']}-{c['section']}" if c["section"] else c["class_name"]
                                 for c in classes)
            teacher_data.append({
                "ID": t["id"],
                "Name": t["full_name"],
                "Username": t["username"],
                "Email": t["email"] or "—",
                "Phone": t["phone"] or "—",
                "Classes": class_str or "—",
                "Status": "Active" if t["is_active"] else "Disabled"
            })
        st.dataframe(teacher_data, use_container_width=True, hide_index=True)
    else:
        st.info("No teachers yet.")

    if has_permission(user["role"], "manage_users_add") or has_permission(user["role"], "manage_users_add_teacher"):
        with st.expander("➕ Add Teacher"):
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
                        except Exception as e:
                            st.error(str(e))


def render_subjects(session):
    session.require("manage_subject_assignments")
    st.header("📖 Subject Assignments")

    assignments = session.list_subject_assignments()
    if assignments:
        data = []
        for a in assignments:
            class_label = f"{a['class_name']}-{a['section']}" if a["section"] else a["class_name"]
            data.append({
                "Class": class_label,
                "Subject": a["subject"],
                "Teacher": a["teacher_name"],
            })
        st.dataframe(data, use_container_width=True, hide_index=True)
    else:
        st.info("No subject assignments yet.")

    with st.expander("➕ Assign Subject"):
        classes = session.list_classes()
        teachers = session.list_teachers()
        if not classes or not teachers:
            st.warning("Create a class and teacher account first.")
        else:
            class_labels = [f"{c['class_name']}-{c['section']}" if c["section"] else c["class_name"] for c in classes]
            class_map = {class_labels[i]: classes[i]["id"] for i in range(len(classes))}
            teacher_map = {f"{t['full_name']} ({t['username']})": t["id"] for t in teachers if t["is_active"]}

            class_pick = st.selectbox("Class", list(class_map.keys()), key="subject_class")
            subject = st.text_input("Subject Name", key="subject_name")
            teacher_pick = st.selectbox("Teacher", list(teacher_map.keys()), key="subject_teacher")

            if st.button("Assign Subject"):
                if not subject.strip():
                    st.error("Subject name is required.")
                else:
                    try:
                        session.assign_subject(class_map[class_pick], subject.strip(), teacher_map[teacher_pick])
                        st.success("Subject assigned.")
                        st.rerun()
                    except Exception as e:
                        st.error(str(e))


def render_attendance(session):
    st.header("📋 Attendance")
    classes = session.list_classes()
    if not classes:
        st.info("No classes available.")
        return

    col1, col2 = st.columns(2)
    with col1:
        class_labels = [f"{c['class_name']}-{c['section']}" if c["section"] else c["class_name"] for c in classes]
        class_map = {class_labels[i]: classes[i]["id"] for i in range(len(classes))}
        class_pick = st.selectbox("Class", list(class_map.keys()), key="att_class")
    with col2:
        att_date = st.date_input("Date", value=date.today(), key="att_date").isoformat()

    class_id = class_map[class_pick]

    try:
        rows = session.load_attendance(class_id, att_date)
    except Exception as e:
        st.error(f"Error: {e}")
        return

    if not rows:
        st.info("No active students in this class.")
        return

    # Build editable table
    table_data = []
    for r in rows:
        table_data.append({
            "student_row_id": r["student_row_id"],
            "Roll No": r["roll_number"] or "—",
            "Name": f"{r['first_name']} {r['last_name'] or ''}",
            "Status": r["status"] or "Present",
        })

    col1, col2 = st.columns(2)
    with col1:
        if st.button("✓ Mark All Present"):
            for row in table_data:
                row["Status"] = "Present"
            st.session_state[f"att_data_{class_id}_{att_date}"] = table_data
            st.rerun()
    with col2:
        if st.button("✗ Mark All Absent"):
            for row in table_data:
                row["Status"] = "Absent"
            st.session_state[f"att_data_{class_id}_{att_date}"] = table_data
            st.rerun()

    edited = st.data_editor(
        table_data, use_container_width=True, hide_index=True, key=f"att_editor_{class_id}_{att_date}",
        column_config={
            "student_row_id": None,
            "Status": st.column_config.SelectboxColumn("Status", options=ATTENDANCE_STATUSES),
            "Roll No": st.column_config.TextColumn("Roll No", disabled=True),
            "Name": st.column_config.TextColumn("Name", disabled=True),
        },
    )

    if st.button("Save Attendance", type="primary"):
        status_map = {int(row["student_row_id"]): row["Status"] for row in edited}
        try:
            session.save_attendance(class_id, att_date, status_map)
            st.success("Attendance saved successfully.")
        except Exception as e:
            st.error(str(e))


def render_marks(session):
    st.header("📊 Marks")
    classes = session.list_classes()
    if not classes:
        st.info("No classes available.")
        return

    class_labels = [f"{c['class_name']}-{c['section']}" if c["section"] else c["class_name"] for c in classes]
    class_map = {class_labels[i]: classes[i]["id"] for i in range(len(classes))}

    col1, col2, col3, col4 = st.columns(4)
    with col1:
        class_pick = st.selectbox("Class", list(class_map.keys()), key="marks_class")
    with col2:
        class_id = class_map[class_pick]
        try:
            subjects = session.list_subjects_for_class(class_id)
            subject = st.selectbox("Subject", subjects or ["(no subjects)"], key="marks_subject")
        except Exception:
            subjects = []
            subject = "(no subjects)"
    with col3:
        exam = st.text_input("Test/Exam", key="marks_exam")
    with col4:
        total = st.text_input("Total Marks", "100", key="marks_total")

    if st.button("Load Class Marks", type="primary"):
        if not subject or subject == "(no subjects)":
            st.error("Assign a subject to this class first.")
        elif not exam:
            st.error("Enter test/exam name.")
        else:
            try:
                marks_rows = session.load_class_marks(class_id, subject, exam)
                st.session_state[f"marks_data_{class_id}_{subject}_{exam}"] = marks_rows
            except Exception as e:
                st.error(str(e))

    # Display editable marks table
    marks_data = st.session_state.get(f"marks_data_{class_id}_{subject}_{exam}", [])
    if marks_data:
        st.subheader(f"Class Marks — {subject} ({exam})")
        table = []
        for row in marks_data:
            table.append({
                "student_row_id": row["student_row_id"],
                "Roll No": row["roll_number"] or "—",
                "Student ID": row["student_id"],
                "Name": f"{row['first_name']} {row['last_name'] or ''}",
                "Marks": row["obtained_marks"] or 0,
            })

        edited = st.data_editor(
            table, use_container_width=True, hide_index=True, key=f"marks_editor_{class_id}_{subject}_{exam}",
            column_config={
                "student_row_id": None,
                "Marks": st.column_config.NumberColumn("Marks", min_value=0, max_value=float(total)),
                "Roll No": st.column_config.TextColumn("Roll No", disabled=True),
                "Student ID": st.column_config.TextColumn("Student ID", disabled=True),
                "Name": st.column_config.TextColumn("Name", disabled=True),
            },
        )

        if st.button("Save All Marks", type="primary"):
            marks_dict = {}
            for row in edited:
                marks_dict[str(row["student_row_id"])] = row["Marks"]
            try:
                session.save_marks_batch(class_id, subject, exam, total, marks_dict)
                st.success("Marks saved.")
                del st.session_state[f"marks_data_{class_id}_{subject}_{exam}"]
                st.rerun()
            except Exception as e:
                st.error(str(e))
    elif st.session_state.get(f"marks_data_{class_id}_{subject}_{exam}") is None:
        st.info("Click 'Load Class Marks' to begin entering marks.")


def render_reports(session):
    st.header("📈 Attendance Reports")
    classes = session.list_classes()
    class_labels = ["All"] + [f"{c['class_name']}-{c['section']}" if c["section"] else c["class_name"] for c in classes]
    class_map = {"All": None}
    for i, c in enumerate(classes):
        label = f"{c['class_name']}-{c['section']}" if c["section"] else c["class_name"]
        class_map[label] = c["id"]

    col1, col2, col3 = st.columns(3)
    with col1:
        class_pick = st.selectbox("Class", class_labels, key="report_class")
    with col2:
        use_dates = st.checkbox("Filter by date range", key="report_dates")
    with col3:
        if use_dates:
            date_from = st.date_input("From", key="report_from").isoformat()
            date_to = st.date_input("To", key="report_to").isoformat()
        else:
            date_from = date_to = None

    if st.button("View Report", type="primary"):
        class_id = class_map[class_pick]
        try:
            report = session.attendance_report(class_id=class_id, date_from=date_from, date_to=date_to)
            st.session_state["_last_report"] = report
        except Exception as e:
            st.error(str(e))

    report = st.session_state.get("_last_report", [])
    if report:
        report_data = [{
            "Student ID": r["student_id"],
            "Roll No": r["roll_number"] or "—",
            "Name": r["name"],
            "Total Days": r["total_days"],
            "Present": r["present"],
            "Absent": r["absent"],
            "Leave": r["leave"],
            "Attendance %": f"{r['percentage']}%",
        } for r in report]
        st.dataframe(report_data, use_container_width=True, hide_index=True)

        user_role = session.current_user["role"]
        if OPENPYXL_AVAILABLE and (has_permission(user_role, "export_excel") or 
                                    has_permission(user_role, "export_excel_assigned")):
            excel_data = _export_attendance_excel(report)
            st.download_button(
                "📥 Download Excel",
                data=excel_data,
                file_name=f"Attendance_{datetime.now():%Y%m%d_%H%M%S}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )


def render_data_tools(session):
    session.require("manage_settings")
    st.header("🛠️ Data Tools")
    st.markdown("Export/import operational data, or create full system backups with database and photos.")

    with st.expander("📊 Excel Workbook — Classes, users, students, marks, etc."):
        col1, col2 = st.columns(2)
        with col1:
            if st.button("📤 Export System Workbook"):
                if not OPENPYXL_AVAILABLE:
                    st.error("Install openpyxl: pip install openpyxl")
                else:
                    try:
                        excel_data = _export_system_workbook(session)
                        st.download_button(
                            "Download Workbook",
                            data=excel_data,
                            file_name=f"RJK_Export_{datetime.now():%Y%m%d}.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                        )
                    except Exception as e:
                        st.error(f"Export failed: {e}")
        with col2:
            st.caption("💡 Tip: Workbooks import by merging records (no deletes, no password changes).")


def render_users(session):
    user = session.current_user
    if not has_permission(user["role"], "manage_users_add"):
        st.error("Only Admin can manage users.")
        return

    st.header("👤 System Users")
    users = session.list_users()

    keyword = st.text_input("Search", key="user_search")
    if keyword:
        kw = keyword.lower()
        users = [u for u in users if kw in u["username"].lower() or kw in u["full_name"].lower()]

    user_data = [{
        "ID": u["id"],
        "Username": u["username"],
        "Full Name": u["full_name"],
        "Role": u["role"],
        "Email": u["email"] or "—",
        "Phone": u["phone"] or "—",
        "Status": "Active" if u["is_active"] else "Disabled",
    } for u in users]
    st.dataframe(user_data, use_container_width=True, hide_index=True)

    with st.expander("➕ Add User"):
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
                    except Exception as e:
                        st.error(str(e))


def render_audit(session):
    user = session.current_user
    if not has_permission(user["role"], "view_audit_log"):
        st.error("Only Admin can view the audit log.")
        return

    st.header("🔍 Audit Log")
    rows = session.db.query(
        """SELECT al.*, u.username FROM audit_log al
           LEFT JOIN users u ON u.id = al.user_id
           ORDER BY al.id DESC LIMIT 500"""
    )
    data = [{
        "Date/Time": r["timestamp"],
        "User": r["username"] or "(unknown)",
        "Action": r["action"],
        "Description": r["description"] or "",
    } for r in rows]
    st.dataframe(data, use_container_width=True, hide_index=True)


# =====================================================================
# EXPORT HELPERS
# =====================================================================
def _export_attendance_excel(report_rows):
    if not OPENPYXL_AVAILABLE:
        raise RuntimeError("openpyxl required")
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Attendance"
    headers = ["Student ID", "Roll No", "Name", "Total Days", "Present", "Absent", "Leave", "Attendance %"]
    ws.append(headers)
    for r in report_rows:
        ws.append([r["student_id"], r["roll_number"] or "", r["name"], r["total_days"],
                   r["present"], r["absent"], r["leave"], r["percentage"]])
    ws.freeze_panes = "A2"
    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf.getvalue()


def _export_system_workbook(session):
    if not OPENPYXL_AVAILABLE:
        raise RuntimeError("openpyxl required")
    wb = openpyxl.Workbook()
    intro = wb.active
    intro.title = "Read Me"
    intro.append(["RJK Student Management System — Data Export"])
    intro.append(["Created", datetime.now().isoformat()])
    intro.append(["Import note", "Workbooks import by merging; no deletes, no password changes."])

    sheets = {
        "Classes": ("SELECT class_name, section, academic_year, is_active FROM classes",
                   ["Class Name", "Section", "Academic Year", "Active"]),
        "Users": ("SELECT username, full_name, role, email, phone, is_active FROM users",
                 ["Username", "Full Name", "Role", "Email", "Phone", "Active"]),
    }

    for sheet_name, (sql, headers) in sheets.items():
        ws = wb.create_sheet(sheet_name)
        ws.append(headers)
        for row in session.db.query(sql):
            ws.append(list(row))

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf.getvalue()


if __name__ == "__main__":
    main()
