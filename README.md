# RandomExamGenerator
This script automates the process of generating randomized exam papers. It uses a pool of questions from both multiple-choice questions (MCQ) and open questions. For each generated exam, solutions are also created. Moreover, it has the capability to include LaTeX exam copies alongside PDF.

## Features:
- Random sampling of questions for variability across multiple exam copies.
- Shuffle MCQ options to further randomize each exam.
- Incorporate images associated with questions, if provided.
- Generate LaTeX formatted exams.

## Prerequisites:
Make sure you have fpdf, pandas, PIL, and jinja2 libraries installed. If not, install them using:

```bash
pip install fpdf pandas pillow jinja2 openpyxl
```

You need two Excel sheets:
- mcq_input1.xlsx: Contains MCQs.
- open_input1.xlsx: Contains open questions.
- (Optional) For LaTeX exam generation, ensure you have a LaTeX distribution installed to compile the .tex files.

## Structure of Pre-existing Files:
mcq_input1.xlsx:
  Should contain the following columns:
        - Question: Text of the MCQ.
        - Option a: First option.
        - Option b: Second option.
        - Option c: Third option.
        - Option d: Fourth option.
        - Correct Answer: Letter (a, b, c, or d) indicating the correct answer.
        - Image_Path: (Optional) Path to an image related to the question, if any.

open_input1.xlsx:
    Should contain the following columns:
        - Questions: Text of the open question.
        - Image_Path: (Optional) Path to an image related to the question, if any.

## How to use:
Make sure the two Excel sheets (mcq_input1.xlsx and open_input1.xlsx) are in the same directory as the script.
If you're using the LaTeX generation feature, ensure the exam_template.tex is in the current directory.

## Run the script:
Provide the required inputs:
- Number of exam copies you want to create.
- Number of MCQs in each exam.
- Number of open questions in each exam.

## The script will generate:
- PDFs for each exam copy.
- Solutions PDF for each exam copy.
- (If LaTeX is used) .tex files for each exam copy.

## Note:
Ensure the image paths provided in the Excel sheets are correct and the images are accessible. If the script can't find an image, it'll print a warning message.
Make sure you have sufficient questions in the Excel sheets for the number of questions you wish to include in the exams. If you ask the script to include more questions than are available, an error will occur.
