"""
Enterprise Active Directory Security Auditor
------------------------------------------
Python desktop application for performing configurable security assessments against 
Microsoft Active Directory environments and exporting findings to structured Excel reports.

This repository demonstrates the application's architecture and software design. 
Proprietary enterprise detection logic has been removed from this portfolio release.
"""

from datetime import datetime, timedelta
import configparser
import os
import sys
import hashlib
import traceback
import threading
import tkinter as tk
from tkinter import ttk, messagebox, filedialog

import openpyxl
from openpyxl.styles import Font, Alignment

# ---------------------------------------------------------------------------
# Portfolio Mode Toggle
# ---------------------------------------------------------------------------
# When True, the application bypasses enterprise directory queries and returns 
# mock data to demonstrate the report generation and GUI flow for portfolio viewers.
PORTFOLIO_MODE = True


# NTLM authentication on modern Python builds requires the MD4 hash
# algorithm, which OpenSSL has dropped from some default builds.
# This registers a fallback implementation when the native one is
# unavailable.
try:
    hashlib.new('md4')
except ValueError:
    from Crypto.Hash import MD4 as _MD4
    _original_hashlib_new = hashlib.new

    def _patched_hashlib_new(name, data=b''):
        if name.lower() == 'md4':
            h = _MD4.new()
            h.update(data)
            return h
        return _original_hashlib_new(name, data)

    hashlib.new = _patched_hashlib_new

from ldap3 import Server, Connection, SUBTREE, ALL, NTLM


def get_base_path():
    """Resolve the application's base directory, accounting for whether
    the app is running as a frozen executable or as a plain script."""
    if getattr(sys, 'frozen', False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))


BASE_DIR = get_base_path()
CONFIG_PATH = os.path.join(BASE_DIR, 'config.ini')
REPORTS_DIR = os.path.join(BASE_DIR, 'Reports')

# ---------------------------------------------------------------------------
# Configuration
# ---------------------------------------------------------------------------
# Settings are loaded from config.ini if present, otherwise safe generic
# defaults are used.

config = configparser.ConfigParser()
if os.path.exists(CONFIG_PATH):
    config.read(CONFIG_PATH)
    LDAP_SERVER = config.get('Settings', 'LdapServer', fallback='localhost')
    DIRECTORY_BASE = config.get('Settings', 'DirectoryBase', fallback='DC=example,DC=local')
    DEFAULT_INACTIVE_DAYS = config.getint('Settings', 'DefaultInactiveDays', fallback=90)
    DEFAULT_NTLM_DOMAIN = config.get('Settings', 'DefaultNtlmDomain', fallback='')

    DEF_CHECK_KERBEROASTABLE = config.getboolean('Settings', 'CheckKerberoastable', fallback=False)
    DEF_CHECK_ASREP_ROASTABLE = config.getboolean('Settings', 'CheckASREPRoastable', fallback=False)
    DEF_CHECK_UNCONSTRAINED_DELEGATION = config.getboolean('Settings', 'CheckUnconstrainedDelegation', fallback=False)
    DEF_CHECK_REVERSIBLE_ENCRYPTION = config.getboolean('Settings', 'CheckReversibleEncryption', fallback=False)
    DEF_CHECK_PASSWORD_NOT_REQUIRED = config.getboolean('Settings', 'CheckPasswordNotRequired', fallback=False)
    DEF_CHECK_PLAINTEXT_IN_NOTES = config.getboolean('Settings', 'CheckPlaintextInNotes', fallback=False)
    DEF_CHECK_LOCKED_OUT = config.getboolean('Settings', 'CheckLockedOut', fallback=False)
else:
    LDAP_SERVER = "dc01.example.local"
    DIRECTORY_BASE = "DC=example,DC=local"
    DEFAULT_INACTIVE_DAYS = 90
    DEFAULT_NTLM_DOMAIN = ""
    DEF_CHECK_KERBEROASTABLE = DEF_CHECK_ASREP_ROASTABLE = False
    DEF_CHECK_UNCONSTRAINED_DELEGATION = DEF_CHECK_REVERSIBLE_ENCRYPTION = False
    DEF_CHECK_PASSWORD_NOT_REQUIRED = DEF_CHECK_PLAINTEXT_IN_NOTES = DEF_CHECK_LOCKED_OUT = False


# ---------------------------------------------------------------------------
# Directory connection
# ---------------------------------------------------------------------------

def get_ad_connection(domain, username, password):
    """Establish an authenticated LDAP connection to the directory server
    using NTLM credentials supplied via the login screen."""
    if PORTFOLIO_MODE:
        return None # Bypass actual connection in portfolio mode

    server = Server(LDAP_SERVER, get_info=ALL)
    ntlm_user = f"{domain}\\{username}"
    conn = Connection(
        server,
        user=ntlm_user,
        password=password,
        authentication=NTLM,
        auto_bind=True
    )
    return conn



# ---------------------------------------------------------------------------
# Audit checks
# ---------------------------------------------------------------------------

def audit_inactive_accounts(conn, days_inactive):
    if PORTFOLIO_MODE:
        return [
            {
                "Username": "j.doe",
                "Email": "j.doe@example.local",
                "Issue": f"Inactive > {days_inactive} days",
                "LastLogon": "01/14/2026 09:22 AM"
            }
        ]
    return []


def audit_never_expire_passwords(conn):
    if PORTFOLIO_MODE:
        return [
            {
                "Username": "svc_sql_backup",
                "Email": "N/A",
                "Issue": "Password Never Expires",
                "LastLogon": "N/A"
            }
        ]
    return []


def audit_advanced_security(conn, opts):
    if PORTFOLIO_MODE:
        findings = []
        if opts.get('kerberoastable'):
            findings.append({"Username": "svc_web", "Email": "N/A", "Issue": "Kerberoastable (SPN set on user)", "LastLogon": "N/A"})
        if opts.get('asrep_roastable'):
            findings.append({"Username": "t.smith", "Email": "t.smith@example.local", "Issue": "AS-REP Roastable (Pre-auth not required)", "LastLogon": "07/28/2026 14:10 PM"})
        if opts.get('unconstrained_delegation'):
            findings.append({"Username": "APPSERV01$", "Email": "N/A", "Issue": "Unconstrained Delegation Enabled (Computer)", "LastLogon": "08/01/2026 08:00 AM"})
        if opts.get('locked_out'):
            findings.append({"Username": "m.jones", "Email": "m.jones@example.local", "Issue": "Account Currently Locked Out", "LastLogon": "08/07/2026 09:15 AM"})
        
        # Fallback if toggles are checked but not mocked above
        if not findings and any(opts.values()):
            findings.append({
                "Username": "example.user",
                "Email": "example@example.local",
                "Issue": "Sample Portfolio Finding",
                "LastLogon": "N/A"
            })
        return findings
    
    return []


# ---------------------------------------------------------------------------
# GUI
# ---------------------------------------------------------------------------

class ADAuditApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Active Directory Security Auditor")
        self.geometry("450x550")
        self.resizable(False, False)

        # Center the window on screen
        self.eval('tk::PlaceWindow . center')

        # Application state
        self.domain_var = tk.StringVar(value=DEFAULT_NTLM_DOMAIN)
        self.username_var = tk.StringVar()
        self.password_var = tk.StringVar()

        self.opt_inactive = tk.BooleanVar(value=True)
        self.val_days = tk.IntVar(value=DEFAULT_INACTIVE_DAYS)
        self.opt_never_expire = tk.BooleanVar(value=True)
        self.opt_kerberoastable = tk.BooleanVar(value=DEF_CHECK_KERBEROASTABLE)
        self.opt_asrep_roastable = tk.BooleanVar(value=DEF_CHECK_ASREP_ROASTABLE)
        self.opt_unconstrained_delegation = tk.BooleanVar(value=DEF_CHECK_UNCONSTRAINED_DELEGATION)
        self.opt_reversible_encryption = tk.BooleanVar(value=DEF_CHECK_REVERSIBLE_ENCRYPTION)
        self.opt_password_not_required = tk.BooleanVar(value=DEF_CHECK_PASSWORD_NOT_REQUIRED)
        self.opt_plaintext_in_notes = tk.BooleanVar(value=DEF_CHECK_PLAINTEXT_IN_NOTES)
        self.opt_locked_out = tk.BooleanVar(value=DEF_CHECK_LOCKED_OUT)

        # Frames for the two-step workflow: login, then check selection
        self.login_frame = ttk.Frame(self, padding=20)
        self.options_frame = ttk.Frame(self, padding=20)

        self.build_login_screen()
        self.build_options_screen()

        self.login_frame.pack(fill="both", expand=True)

    def toggle_all_advanced(self):
        new_state = not self.opt_kerberoastable.get()
        self.opt_kerberoastable.set(new_state)
        self.opt_asrep_roastable.set(new_state)
        self.opt_unconstrained_delegation.set(new_state)
        self.opt_reversible_encryption.set(new_state)
        self.opt_password_not_required.set(new_state)
        self.opt_plaintext_in_notes.set(new_state)
        self.opt_locked_out.set(new_state)


    def build_login_screen(self):
        ttk.Label(self.login_frame, text="Active Directory Security Scan", font=("Helvetica", 14, "bold")).pack(pady=(0, 20))

        # Domain
        ttk.Label(self.login_frame, text="Domain (NetBIOS):").pack(anchor="w")
        ttk.Entry(self.login_frame, textvariable=self.domain_var, width=40).pack(pady=(0, 10))

        # Username
        ttk.Label(self.login_frame, text="Username:").pack(anchor="w")
        ttk.Entry(self.login_frame, textvariable=self.username_var, width=40).pack(pady=(0, 10))

        # Password
        ttk.Label(self.login_frame, text="Password:").pack(anchor="w")
        password_entry = ttk.Entry(self.login_frame, textvariable=self.password_var, show="*", width=40)
        password_entry.pack(pady=(0, 20))
        # Allow the user to submit credentials by pressing Enter.
        password_entry.bind("<Return>", lambda event: self.show_options())

        ttk.Button(self.login_frame, text="Next  ➔", command=self.show_options).pack(fill="x", pady=10)

    def show_options(self):
        if not self.username_var.get() or not self.password_var.get():
            messagebox.showwarning("Input Error", "Username and password are required.")
            return

        # Validate credentials before moving to the check-selection screen

        try:
            conn = get_ad_connection(
                self.domain_var.get(),
                self.username_var.get(),
                self.password_var.get()
            )
            if conn:
                conn.unbind()
        except Exception:
            messagebox.showerror(
                "Authentication Failed",
                "Invalid username or password.\n\nPlease check your credentials and try again."
            )
            return

        self.login_frame.pack_forget()
        self.options_frame.pack(fill="both", expand=True)

    def build_options_screen(self):
        ttk.Label(self.options_frame, text="Select Security Checks", font=("Helvetica", 14, "bold")).pack(pady=(0, 10))

        inactive_frame = ttk.Frame(self.options_frame)
        inactive_frame.pack(fill="x", anchor="w", pady=2)
        ttk.Checkbutton(inactive_frame, text="Inactive Accounts (Days: ", variable=self.opt_inactive).pack(side="left")
        ttk.Entry(inactive_frame, textvariable=self.val_days, width=5).pack(side="left")
        ttk.Label(inactive_frame, text=")").pack(side="left")

        ttk.Checkbutton(self.options_frame, text="Password Never Expires", variable=self.opt_never_expire).pack(fill="x", anchor="w", pady=2)

        ttk.Separator(self.options_frame, orient='horizontal').pack(fill='x', pady=10)
        ttk.Label(self.options_frame, text="Advanced Checks", font=("Helvetica", 10, "bold")).pack(anchor="w", pady=(0, 5))
        ttk.Button(self.options_frame, text="Toggle All Advanced Checks", command=self.toggle_all_advanced).pack(fill="x", pady=(0, 5))

        ttk.Checkbutton(self.options_frame, text="Kerberoastable Accounts (SPN)", variable=self.opt_kerberoastable).pack(fill="x", anchor="w", pady=2)
        ttk.Checkbutton(self.options_frame, text="AS-REP Roastable (No Pre-Auth)", variable=self.opt_asrep_roastable).pack(fill="x", anchor="w", pady=2)
        ttk.Checkbutton(self.options_frame, text="Unconstrained Delegation", variable=self.opt_unconstrained_delegation).pack(fill="x", anchor="w", pady=2)
        ttk.Checkbutton(self.options_frame, text="Reversible Password Encryption", variable=self.opt_reversible_encryption).pack(fill="x", anchor="w", pady=2)
        ttk.Checkbutton(self.options_frame, text="Password Not Required Flag", variable=self.opt_password_not_required).pack(fill="x", anchor="w", pady=2)
        ttk.Checkbutton(self.options_frame, text="Plaintext Credentials in Notes", variable=self.opt_plaintext_in_notes).pack(fill="x", anchor="w", pady=2)
        ttk.Checkbutton(self.options_frame, text="Currently Locked Out Accounts", variable=self.opt_locked_out).pack(fill="x", anchor="w", pady=2)

        self.status_label = ttk.Label(self.options_frame, text="", foreground="blue")
        self.status_label.pack(pady=(15, 5))

        self.btn_run = ttk.Button(self.options_frame, text="Start Security Scan", command=self.start_scan_thread)
        self.btn_run.pack(fill="x", pady=10)

        ttk.Button(self.options_frame, text="🡄 Back", command=self.show_login).pack(fill="x")

    def show_login(self):
        self.options_frame.pack_forget()
        self.login_frame.pack(fill="both", expand=True)

    def start_scan_thread(self):
        try:
            days = self.val_days.get()
            if days < 1:
                raise ValueError
        except Exception:
            messagebox.showwarning("Input error", "Please enter a valid number of days (ex. 90)")
            return

        default_filename = f"AD_Audit_Report_{self.username_var.get()}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"

        output_path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel Files", "*.xlsx"), ("All Files", "*.*")],
            initialfile=default_filename,
            title="Save security report as"
        )

        if not output_path:
            return

        self.btn_run.config(state="disabled")
        self.status_label.config(text="Connecting to Active Directory...", foreground="blue")

        threading.Thread(target=self.execute_scan, args=(output_path,), daemon=True).start()

    def execute_scan(self, output_path):
        try:
            conn = get_ad_connection(self.domain_var.get(), self.username_var.get(), self.password_var.get())
            all_findings = []

            if self.opt_inactive.get():
                self.update_status("Scanning for inactive accounts...")
                all_findings.extend(audit_inactive_accounts(conn, self.val_days.get()))

            if self.opt_never_expire.get():
                self.update_status("Scanning for passwords never expire...")
                all_findings.extend(audit_never_expire_passwords(conn))

            self.update_status("Running advanced security scans...")
            advanced_opts = {
                'kerberoastable': self.opt_kerberoastable.get(),
                'asrep_roastable': self.opt_asrep_roastable.get(),
                'unconstrained_delegation': self.opt_unconstrained_delegation.get(),
                'reversible_encryption': self.opt_reversible_encryption.get(),
                'password_not_required': self.opt_password_not_required.get(),
                'plaintext_in_notes': self.opt_plaintext_in_notes.get(),
                'locked_out': self.opt_locked_out.get(),
            }
            all_findings.extend(audit_advanced_security(conn, advanced_opts))

            self.update_status("Generating Excel report...")

            # Build the report workbook
            wb = openpyxl.Workbook()
            ws = wb.active
            ws.title = "AD Security Audit"

            headers = ['Username', 'Email', 'Issue', 'LastLogon']
            ws.append(headers)

            for cell in ws[1]:
                cell.font = Font(bold=True)

            ws.freeze_panes = 'A2'

            for finding in all_findings:
                ws.append([finding['Username'], finding['Email'], finding['Issue'], finding['LastLogon']])

            center_align = Alignment(horizontal='center', vertical='center')
            left_align = Alignment(horizontal='left', vertical='center')

            for col in ws.columns:
                max_length = 0
                column_letter = col[0].column_letter
                header_name = col[0].value


                for cell in col:
                    if header_name == 'Issue':
                        cell.alignment = center_align
                    else:
                        cell.alignment = left_align

                    try:
                        if cell.value:
                            cell_len = len(str(cell.value))
                            if cell_len > max_length:
                                max_length = cell_len
                    except Exception:
                        pass

                adjusted_width = (max_length * 1.05) + 2
                ws.column_dimensions[column_letter].width = adjusted_width

            wb.save(output_path)
            self.after(0, self.scan_complete, output_path, len(all_findings))
        except Exception as e:
            error_msg = str(e)
            traceback.print_exc()
            self.after(0, self.scan_error, error_msg)

    def update_status(self, message):
        self.after(0, lambda: self.status_label.config(text=message))

    def scan_complete(self, filepath, count):
        self.btn_run.config(state="normal")
        self.status_label.config(text=f"Scan complete. {count} findings.", foreground="green")
        messagebox.showinfo("Audit Complete", f"Scan finished successfully.\n\nFound {count} issues.\n\nReport saved to:\n{filepath}")

    def scan_error(self, error_msg):
        self.btn_run.config(state="normal")
        self.status_label.config(text="Scan failed. Check credentials/network.", foreground="red")
        messagebox.showerror("Audit Error", f"An error occurred:\n\n{error_msg}")


if __name__ == "__main__":
    app = ADAuditApp()
    app.mainloop()