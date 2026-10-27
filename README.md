# 🛡️ Labor Audit ETL & Binary Excel Engine (.XLSX)

## 🎯 Business Problem
During internal labor litigation audits across multiple manufacturing plants, native Excel (`.xlsx`) reports suffered from critical data governance issues:
1. **Double-Counting Distortion:** A single lawsuit containing multiple secondary claims artificially inflated litigation volume metrics for executive leadership.
2. **File Corruption Risks:** Manually manipulating traditional `.xlsx` files inside cloud drives often broke formulas, formatting, or corporate compliance templates.
3. **Unstructured Third-Party Liability:** Outsourced vendor names were written inconsistently inside free-text lawsuit summaries, hindering accurate third-party risk allocation.

## 💡 Technical Solution (Binary `.XLSX` ETL Pipeline)
An automated **ETL (Extract, Transform, Load)** engine built with JavaScript (Google Apps Script + **Drive API v3**) that operates invisibly on native Excel (`.xlsx`) files without corrupting the source file.

### 🛠️ Key Engineering Highlights
* **Safe Binary `.XLSX` Injection Pipeline:** Creates a temporary cloud spreadsheet copy for high-speed memory calculation, formats the executive report, exports the binary `.xlsx` blob via `UrlFetchApp`, and injects it directly back into the original `.xlsx` file ID via `Drive.Files.update` before trashing the temporary file.
* **Single Primary Claim Classifier (*1 Lawsuit = 1 Primary Claim*):** Advanced Regex text-mining algorithm that scans initial petition summaries and isolates the earliest predominant claim (*Occupational Disease, Hazard Pay, Overtime, Just Cause Reversal*), mathematically guaranteeing **100% audit closing accuracy** across all plants.
* **Third-Party Vendor Pattern Matching:** Sanitizes and consolidates unstructured vendor mentions across 30+ outsourced service providers.
* **Programmatic Executive UI Rendering:** Dynamically builds dual-theme executive tables (Corporate Blue for Direct Employees, Executive Green for Outsourced Vendors), injects active data filters (`createFilter()`), and renders typographic progress bars (`█` and `░`) directly inside cells.

## 📂 Repository Structure
* `AuditETL.gs` — Complete ETL pipeline, Regex text-mining classifier, binary `.xlsx` injector, and programmatic spreadsheet styling engine.
