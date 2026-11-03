# Google Sheets Operations Automation

A workflow automation solution built using Google Sheets and Google Apps Script to streamline task management, recurring schedules, deadline monitoring, and automated escalation.

## Overview

This project demonstrates how operational workflows can be automated using Google Sheets and Google Apps Script.

It provides a structured system for managing employee tasks, tracking completion, monitoring deadlines, notifying managers, and recording escalation activities.

## Project Highlights

- Automated task tracking and status monitoring
- Recurring task scheduling
- Employee-wise dashboards
- Automatic next due date calculation
- Automated manager email notifications
- Manager acknowledgement tracking
- Leadership escalation notifications
- Duplicate alert prevention
- Task completion and escalation history
- Configurable SLA settings

## Technology Stack

- **Google Sheets** – Data storage and dashboards
- **Google Apps Script** – Workflow automation
- **JavaScript** – Automation logic
- **Gmail Services** – Email notifications

## System Structure

The solution is organized into six worksheets:

**1. Task Master**

Stores task details, employee assignments, schedules, and due dates.

**2. Employee Dashboard**

Displays employee-wise task information and completion status.

**3. Completion Log**

Maintains a history of completed tasks.

**4. Manager Dashboard**

Tracks manager notifications and acknowledgement status.

**5. Alert & Escalation Log**

Records manager alerts and leadership escalation activities.

**6. Settings**

Stores configurable SLA values and timezone information.

## Workflow

1. Task details and employee assignments are maintained in Google Sheets.
2. The automation checks task deadlines and completion status.
3. Employee dashboards are updated with task information.
4. Upon task completion, the next due date is calculated for recurring tasks.
5. Managers receive email notifications for overdue tasks.
6. Manager acknowledgement status is tracked.
7. Leadership escalation notifications are generated according to configured rules.
8. Completion and escalation activities are recorded in their respective logs.

## SLA Configuration

The system uses the following configured values:

- **Manager Alert Delay:** 30 minutes
- **Manager Acknowledgement SLA:** 2 hours
- **Leadership Escalation Delay:** 4 hours
- **Timezone:** Asia/Kolkata

## Testing & Validation

The following scenarios were tested successfully:

- Dashboard setup and data updates
- Recurring task completion
- Next due date calculation
- Manager alert generation
- Manager acknowledgement updates
- Duplicate alert prevention
- Leadership escalation notifications
- Completion and escalation log updates

### Test Report

[View the Operations Automation Test Report](operations_automation_test_report_portfolio.pdf)

### Live Demo

[View the Google Sheets Project](https://docs.google.com/spreadsheets/d/1rnmNBkvzoG376Hx8hPqfBQDYiZaqwWetCbQuMYXVdy4/edit?usp=sharing)

## Repository Contents

| File | Description |
|---|---|
| operations_automation.gs | Google Apps Script source code |
| operations_automation_test_report_portfolio.pdf | Project testing report |
| README.md | Project documentation |

## Key Learnings

- Business process and workflow automation
- JavaScript programming with Google Apps Script
- Google Sheets integration and data management
- Time-driven and event-driven triggers
- Dashboard development and task monitoring
- Automated email notifications
- Alert management and escalation workflows

---

**Project Type:** AI Operations & Automation

**Author:** Bhavika Vaishnav
