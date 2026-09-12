# python-course
Command-line result management sysstem

# 🎓 Student Result Management System
A menu-driven command-line application developed using Python CLI to manage student academic results efficiently. The system allows users to enter student information and marks, automatically calculate results, search individual records, and maintain student data in an Excel file using OpenPyXL.

## 📌 About the Project
The Student Result Management System is a Python-based CLI project designed to simplify basic student result management.

The application collects a student's:
* Roll Number
* Name
* Class/Course
* Marks of 5 subjects

Based on the entered marks, the system automatically calculates:

* Total Marks
* Percentage
* Grade
* Pass/Fail Status

All student records are stored permanently in an Excel workbook, making the data easy to maintain and access.

The project also includes input validation and duplicate Roll Number checking to improve the reliability of the system.

## ✨ Key Features

### 👨‍🎓 Student Management

* Add new student result
* Enter student details and marks
* Prevent duplicate Roll Numbers
* Validate student inputs

### 📊 Result Calculation

* Automatically calculate total marks
* Calculate percentage
* Assign grade based on percentage
* Determine Pass/Fail status

### 🔍 Result Search

* Search for a student's result using Roll Number
* Display complete result details including subject-wise marks

### 📋 Student Records

* Display all stored student records in a formatted table
* Read records directly from the Excel file

### 📁 Excel Data Storage

* Automatically create the Excel file if it does not exist
* Store student records permanently
* Use **OpenPyXL** for reading and writing Excel data

### 🛡️ Input Validation

* Prevent invalid Roll Numbers
* Prevent empty student names
* Prevent empty Class/Course
* Validate marks within the allowed range
* Handle invalid numeric input using exception handling
* Prevent duplicate Roll Numbers


## 🖥️ Application Menu
The application provides a simple command-line menu:
==================================================
       STUDENT RESULT MANAGEMENT SYSTEM
==================================================
1. Add Student Result
2. Get Student Result
3. Show All Student Data
4. Exit
==================================================
Enter your choice:

### Option 1 — Add Student Result

Allows the user to enter student information and marks for five subjects.

The system then calculates:
Total Marks
Percentage
Grade
Pass/Fail Status

The complete record is saved to:
student_results.xlsx

### Option 2 — Get Student Result

The user enters a student's Roll Number.

The system searches the Excel file and displays the student's complete result.

Example:

==========================================
              STUDENT RESULT

==========================================

Roll No     : 101

Name        : Rahul

Class       : BCA

------------------------------------------
Subject 1   : 85

Subject 2   : 78

Subject 3   : 92

Subject 4   : 80

Subject 5   : 88

------------------------------------------

Total       : 423

Percentage  : 84.6 %

Grade       : A

Status      : PASS

==========================================

### Option 3 — Show All Student Data

Displays all stored students in a simple tabular format.

Roll No   Name       Class    Total     Percentage   Grade   Status

---------------------------------------------------------------------
101       Rahul      BCA      423       84.6         A       PASS
102       Priya      BCA      378       75.6         B       PASS
103       Amit       BCA      198       39.6         F       FAIL


### Option 4 — Exit
Closes the application safely.

## 🧮 Grading System

The project uses the following grading criteria:

| Percentage    | Grade |
| ------------- | ----- |
| 90% and above | A+    |
| 80% – 89%     | A     |
| 70% – 79%     | B     |
| 60% – 69%     | C     |
| 50% – 59%     | D     |
| Below 50%     | F     |

### Pass/Fail

Percentage >= 40%  → PASS
Percentage < 40%   → FAIL


## 📊 Excel Data Structure

Student records are stored in:
**`student_results.xlsx`**

The Excel sheet contains the following columns:

| Column     | Description                |
| ---------- | -------------------------- |
| Roll No    | Unique student roll number |
| Name       | Student name               |
| Class      | Class/Course               |
| Subject 1  | Marks                      |
| Subject 2  | Marks                      |
| Subject 3  | Marks                      |
| Subject 4  | Marks                      |
| Subject 5  | Marks                      |
| Total      | Total obtained marks       |
| Percentage | Calculated percentage      |
| Grade      | Assigned grade             |
| Status     | PASS / FAIL                |


## 🛠️ Technologies Used

* **Python 3**
* **OpenPyXL**
* **Microsoft Excel**
* **Command Line Interface (CLI)**


## 📚 Python Concepts Used

This project demonstrates practical use of:
* Variables
* Data Types
* `input()` and output
* Lists
* Functions
* Function parameters and return values
* `if-elif-else`
* `for` loop
* `while` loop
* `try-except` exception handling
* String formatting
* File handling
* Excel file operations
* Reading and writing Excel data
* Input validation

## 🏗️ Working Flow

                 ┌──────────────────────────┐
                 │          START           │
                 └────────────┬─────────────┘
                              ↓
                 ┌──────────────────────────┐
                 │ Create / Check Excel     │
                 │ student_results.xlsx     │
                 └────────────┬─────────────┘
                              ↓
                 ┌──────────────────────────┐
                 │      MAIN MENU           │
                 ├──────────────────────────┤
                 │ 1. Add Student Result    │
                 │ 2. Get Student Result    │
                 │ 3. Show All Student Data │
                 │ 4. Exit                  │
                 └────────────┬─────────────┘
                              ↓
                 ┌──────────────────────────┐
                 │     Enter Choice         │
                 └────────────┬─────────────┘
                              ↓
             ┌────────────────┼─────────────────┐
             ↓                ↓                 ↓
        ┌─────────┐      ┌─────────┐       ┌─────────┐
        │ Option 1│      │ Option 2│       │ Option 3│
        │Add Data │      │Get Data │       │Show All │
        └────┬────┘      └────┬────┘       └────┬────┘
             ↓                ↓                 ↓
      Enter Student     Enter Roll No.    Read Excel Data
      Information             ↓                 ↓
             ↓          Search Roll No.   Display Records
      Enter 5 Subjects       ↓
             ↓          Display Result
      Calculate Total,
      Percentage & Grade
             ↓
      PASS / FAIL
             ↓
      Save to Excel
             ↓
             └──────────────┬──────────────────┘
                            ↓
                       Return to Menu
                            ↓
                     ┌──────────────┐
                     │ Option 4 ?   │
                     └──────┬───────┘
                            ↓
                          EXIT

### student_results.xlsx

Stores all student result records.
