# README.md

Page Contents:

- Power Query Code
- Technical walkthrough

# Business Central Finance System – High-Level Workflow Guide

This document provides a simplified overview of how the new Finance system using **Microsoft Dynamics 365 Business Central** operates. It is intended for non-technical users to understand the key steps, tools, and processes involved.

---

## 1. Overview

The Finance system is designed to manage journal data efficiently, ensure accuracy, and maintain a clear audit trail. It combines:

- Data extraction (RE Data)
- Templates and validation checks
- Journal uploads into Business Central (BC)
- Approval and audit processes

---

## 2. Key Components Explained

### Templates

#### **Old Templates**
Previously, templates included:
- Income Stream tracking
- Multiple pivot tables for comparison
- Journal-based data checks

These often became complex and sometimes had issues such as:
- Missing data
- Gaps in journals
- Over-reliance on manual validation

---

### New Template Structure

#### **Blue Table (Do Not Edit)**
- This is a **Power Query-generated table**
- To be populated with RE data
- **Important:** Do not format or edit this table manually

#### **Lookups**
- Reference tables used to map and structure data correctly
- Help transform raw data into the required journal format

#### **RE Query**
- Data is extracted and prepared (currently by Peter)
- Imported into the template manually

---

## 3. Business Central (BC) Integration

### Analysis Mode
Within BC:
- Users can validate data by checking:
  - **Sum of Debits**
  - **Sum of Credits**

✅ If totals match → Journal is balanced and ready  
❌ If totals do not match → Import will fail until resolved

---

## 4. Journal Posting Methods

### Manual Import (Less Common)
Using the **Pusher Template**:
1. Select the correct journal
2. Copy and paste data into the template
3. Publish via Microsoft Dynamics connector

> ⚠️ Note: This method can occasionally fail due to system configuration issues.

---

### Standard Process (Most Common)

1. Journals are prepared and saved in: SPS → Finance → Business Central → PCUK folder
   
2. **Power Automate Flow (overnight):**
- Collects all journal files
- Consolidates them into a single journal batch in BC

3. Files are:
- Processed automatically
- Moved into a **"Processed" folder** for audit tracking

---

## 5. Evidence & Audit Trail

Each journal:
- Must include **supporting evidence links**
- Goes through an **approval process**
- Is stored for auditing purposes once processed

This ensures:
- Transparency
- Accountability
- Compliance with finance controls

---

## 6. Journal Requirements

For successful upload into Business Central:
- Journals must follow a **strict format**
- Any deviation may cause:
- Upload failures
- Data inconsistencies

---

## 7. Monthly Workflow Timeline

Approximately **60 journals per month** are processed.

### 🗓 Weekly Cycle:

- **Monday–Tuesday**
- Journal preparation

- **Wednesday–Thursday**
- Review and approval process

- **Friday**
- Final submission to Business Central

---

## 8. Known Challenges

- **RE Data Extract Differences**
- Occasionally, extracted data may not align perfectly
- Requires validation before upload

- **Template Sensitivity**
- Formatting changes (especially in the Blue Table) can break the process

- **Manual Upload Limitations**
- Pusher template can fail depending on system configurations

---

## 9. Summary

The new Finance system improves:
- Automation (via Power Query & Power Automate)
- Accuracy (through validation checks)
- Auditability (through evidence tracking and processed file storage)

By following the defined templates and processes, users can ensure smooth and reliable journal posting into Business Central.

---

## ✅ Key Takeaways

- Do **not** edit the Blue Table
- Ensure journals are **balanced (Debit = Credit)**
- Follow the **standard template format**
- Use the **automated process wherever possible**
- Always include **supporting evidence**

---
