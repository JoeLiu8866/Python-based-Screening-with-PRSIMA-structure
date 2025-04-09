# Python-based-Screening-with-PRSIMA-structure
# 📚 Computer-Aided Systematic Review Toolkit

This repository contains a modular Python-based toolkit designed to support systematic literature reviews, especially under the PRISMA framework. It helps researchers streamline large-scale screening processes through semi-automated data handling, filtering, and validation workflows.

---

## 🚀 Overview

The toolkit consists of the following components:

| Module | Description |
|--------|-------------|
| **Bib to Excel** | Parses `.bib` files and formats metadata (title, author, year, venue, abstract, link) into styled Excel sheets. |
| **Excel Deduplication** | Merges multiple Excel files and removes duplicate entries. |
| **Remove Incomplete Data** | Filters out entries missing essential fields like title, year, author, or abstract. |
| **Remove Irrelevant Papers (Automated Driving)** | Filters articles based on specific keywords related to Automated Driving (AD) or Highly Automated Driving (HAD). |
| **Sampling & Validation** | Supports systematic random sampling and calculates optimal sample size using Neyman’s formula for screening validation. |

---

## 🧭 Workflow (7 Steps)

This toolkit supports the following computer-aided review process:

1. **BibTeX Import**  
   Export `.bib` citation files from major databases (e.g., ACM, IEEE Xplore, Web of Science, ScienceDirect).

2. **Excel Conversion**  
   Use `Bib to Excel` to convert `.bib` files into styled `.xlsx` files.

3. **Deduplication**  
   Apply the `Excel Deduplication` module to combine and remove duplicate entries.

4. **Incomplete Entry Filtering**  
   Use the `Remove Incomplete Data` module to remove entries missing key metadata.

5. **Domain Keyword Filtering**  
   Run the `Remove Irrelevant Papers` scripts (focused on HAD or takeover terms) to exclude articles that do not match your research scope.

6. **Systematic Random Sampling**  
   Use `Sampling Excel`, `Systematic Sampling`, and `Sample Size Calculator` to validate exclusion accuracy with statistically sound methods.

7. **Manual Eligibility Screening**  
   Manually inspect remaining articles using the exported Excel file for final inclusion.

---

## 🛠️ Requirements

- Python 3.8 or later  
- Recommended libraries:
  - `pandas`
  - `openpyxl`
  - `bibtexparser`

Install all dependencies:

```bash
pip install pandas openpyxl bibtexparser
