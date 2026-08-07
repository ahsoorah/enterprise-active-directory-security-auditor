# Enterprise Active Directory Security Auditor

> A Python-based desktop application for automating Active Directory security assessments, identity security reviews, and structured reporting.

> **Note:** This public repository contains a sanitized portfolio edition of the application I developed during my internship. Organization-specific detection logic has been intentionally removed while preserving the application's architecture, workflow, and software design.

---

## Overview

Enterprise Active Directory Security Auditor is a desktop application developed in Python to streamline common identity security assessments in Microsoft Active Directory environments.

The application authenticates to an Active Directory domain using LDAP/NTLM, allows administrators to select configurable security assessment categories through a graphical interface, and generates a structured Excel report summarizing discovered findings.

This project was originally developed for use in an enterprise environment. The public repository preserves the application's architecture, user interface, reporting pipeline, and overall design while omitting proprietary audit implementation.

---

# Features

* Active Directory authentication using LDAP and NTLM
* Desktop GUI built with Tkinter
* Configurable security assessment options
* Multi-threaded scan execution to keep the interface responsive
* Configuration-driven deployment
* Automated Excel report generation
* Structured finding collection
* Portfolio Mode demonstration data for safe public use
* Comprehensive error handling and status reporting

---

# Security Assessment Categories

The application demonstrates workflows commonly used during enterprise identity security assessments, including:

* Inactive account identification within a certain timeframe
* Password policy review
* Authentication configuration review
* Service account review
* Delegation configuration review
* Account lockout review
* Identity security reporting

The public version returns representative sample findings to demonstrate the reporting workflow. Organization-specific detection logic has been removed.

---

# Application Workflow

```text
User Authentication
        │
        ▼
LDAP / NTLM Connection
        │
        ▼
Security Assessment Engine
        │
        ▼
Finding Collection
        │
        ▼
Excel Report Generation
```

---

# Technologies

* Python 3
* LDAP
* NTLM Authentication
* ldap3
* Tkinter
* OpenPyXL
* Multithreading
* ConfigParser
* Enterprise Identity Security
* Microsoft Active Directory

---

# Project Structure

```text
.
├── ad_security_auditor.py
├── config.ini.example
├── requirements.txt
├── README.md
└── Reports/
```

---

# Configuration

Create a `config.ini` file based on `config.ini.example`.

Example configuration:

```ini
[Settings]
LDAP_SERVER = dc01.example.local
DIRECTORY_BASE = DC=example,DC=local
DEFAULT_INACTIVE_DAYS = 90
DEFAULT_DOMAIN = EXAMPLE
```

---

# Report Output

After completing a security assessment, the application exports an Excel workbook containing representative findings.

Example report fields include:

* Username
* Email Address
* Security Finding
* Last Logon
* Severity
* Recommendation

The included findings are demonstration data generated in Portfolio Mode and are intended to illustrate the application's reporting pipeline.

---

# Skills Demonstrated

This project demonstrates practical experience with:

* Identity and Access Management (IAM)
* Active Directory administration
* LDAP authentication
* Security automation
* Python application development
* Desktop application development
* Multithreaded programming
* Enterprise reporting
* Defensive security tooling
* Configuration management
* Error handling and logging

---

# Why This Repository Is Sanitized

This project was originally developed as part of an enterprise cybersecurity environment.

To respect organizational confidentiality, the public version removes:

* Proprietary Active Directory detection logic
* Internal LDAP queries
* Organization-specific configuration
* Enterprise infrastructure details
* Internal reporting logic

The application's architecture, user interface, workflow, and engineering design have been preserved to demonstrate software development and cybersecurity engineering skills without exposing sensitive implementation details.

---

# Future Enhancements

Potential improvements include:

* HTML reporting
* CSV export support
* Additional identity security assessments
* Logging framework integration
* Report severity scoring
* Automated remediation recommendations

---

# Disclaimer

This repository is intended for portfolio purposes.

It demonstrates the design and architecture of an enterprise identity security auditing application while intentionally omitting proprietary implementation details from the original production version.

---

## Author

**Suriyah Saravanan**

Management Information Systems, Cybersecurity

Focused on Identity Security, Security Engineering, Active Directory, and Defensive Cybersecurity Automation.
