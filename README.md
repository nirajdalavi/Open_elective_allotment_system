# 🧠 Open Elective Allotment System

A Python-based automated elective allocation system that fairly and efficiently assigns elective courses to students based on CGPA and preference rankings. Built to support academic institutions managing open elective selections.

---

## 📌 Features

- ✅ Automatically allocates electives based on student choices and CGPA
- ✅ Resolves tie-breaks using preference order
- ✅ Rejects electives with fewer than 30 students
- ✅ Prevents re-allotment of previously completed courses
- ✅ Exports final allotments and course summaries to Excel
- ✅ Handles reallocation dynamically when conflicts arise

---

## 📂 Project Structure

- `Student` & `Elective` classes for core logic
- Reads input from Excel (`RawDataOE1.xlsx`)
- Generates output in `Allot3.xlsx`:
  - Full student allotment list
  - Course-wise sheets
  - Summary of enrollments and cutoff CGPA

---

## 🧮 Allotment Logic

1. **Students sorted by CGPA (highest first)**
2. For each student:
   - Traverse their choices
   - Allocate based on available capacity
   - In case of tie (CGPA), resolve using preference index
   - Reallocate if a more deserving student appears
3. Courses with fewer than 30 students are **removed**
4. Previous electives are **filtered out** from choices

---

## 📊 Output Files

- `Allotment`: All students with elective allocations
- `Summary`: Each course with total students and minimum CGPA
- Individual course sheets with detailed student lists

---

## 🛠️ Tech Stack

- Python 3
- [openpyxl](https://pypi.org/project/openpyxl/) for Excel handling

---

## 🚀 How to Run

1. **Install dependencies**:
   ```bash
   pip install openpyxl

2.	**Update file paths:
  Change the input (RawDataOE1.xlsx) and output (Allot3.xlsx) paths in the script according to your system.

3.	**Run the script:
  python OE1.py
