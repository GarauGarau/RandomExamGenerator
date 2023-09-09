from fpdf import FPDF
import pandas as pd
import random
import os

from PIL import Image
from jinja2 import Environment, FileSystemLoader


class PDF(FPDF):
    def header(self):
        self.set_font("Times", "B", 12)
    
    def footer(self):
        self.set_font("Times", "I", 8)


def shuffle_mcq_choices(row):
    choices = ["a", "b", "c", "d"]
    correct = row['Correct Answer']
    shuffled_choices = random.sample(choices, len(choices))
    shuffled_options = [row[f"Option {choice}"] for choice in shuffled_choices]
    correct_position = shuffled_choices.index(correct)
    return shuffled_options, correct_position

def get_image_height(image_path, width):
    with Image.open(image_path) as img:
        aspect_ratio = img.width / img.height
        return width / aspect_ratio

def space_left(pdf):
    return 297 - pdf.get_y()  # Assuming an A4 size in portrait orientation which is 297mm

def create_exam_and_solution(mcq_df, open_df, num_copies, mcq_count, open_count):
    # Per la generazione dei file .tex
    env = Environment(loader=FileSystemLoader('.'))
    template = env.get_template('exam_template.tex')

    for copy_num in range(1, num_copies + 1):
        mcq_sample = mcq_df.sample(mcq_count).reset_index(drop=True)
        open_sample = open_df.sample(open_count).reset_index(drop=True)
        
        pdf_exam = PDF()
        pdf_exam.add_page()
        
        # Title
        pdf_exam.set_font('Arial', 'B', 18)
        pdf_exam.cell(200, 10, txt="Esame di Macroeconomia", ln=True, align='C')
        
        # Subtitle
        pdf_exam.set_font('Arial', 'B', 14)
        pdf_exam.cell(200, 10, txt="Domande a risposta multipla - 95 minuti", ln=True, align='C')

        # Space for the student's details
        pdf_exam.ln(10)
        pdf_exam.cell(40, 10, txt="Nome: ", border='B', ln=0)
        pdf_exam.cell(60, 10, "", border='B', ln=True)
        pdf_exam.cell(40, 10, txt="Cognome: ", border='B', ln=0)
        pdf_exam.cell(60, 10, "", border='B', ln=True)
        pdf_exam.cell(40, 10, txt="Matricola: ", border='B', ln=0)
        pdf_exam.cell(60, 10, "", border='B', ln=True)
        pdf_exam.ln(10)  # Add some space after the details section
        
        # Questions start here
        pdf_exam.set_font('Arial', size=12)
        pdf_exam.cell(200, 10, txt="Exam {}".format(copy_num), ln=True, align='C')

        pdf_solution = PDF()
        pdf_solution.add_page()
        pdf_solution.set_font('Arial', size=12)
        pdf_solution.cell(200, 10, txt="Solution Copy {}".format(copy_num), ln=True, align='C')

        for index, row in mcq_sample.iterrows():
            shuffled_options, correct_position = shuffle_mcq_choices(row)
            question_text = f"{index + 1}. {row['Question']}"
            pdf_exam.multi_cell(0, 10, txt=question_text)

            image_path = row.get('Image_Path', None)
            if pd.notna(image_path):
                if os.path.exists(image_path):
                    img_width = 90
                    img_height = get_image_height(image_path, img_width)
                    if space_left(pdf_exam) < img_height:
                        pdf_exam.add_page()
                    pdf_exam.image(image_path, x=10, y=pdf_exam.get_y(), w=img_width)
                    pdf_exam.ln(img_height + 5)  # Added a 5mm padding after the image
                else:
                    print(f"Image not found for path {image_path}")


            for idx, opt in enumerate(shuffled_options):
                choice_text = f"    {chr(97 + idx)}. {opt}"
                pdf_exam.multi_cell(0, 10, txt=choice_text)
            
            solution_text = f"{index + 1}. Correct Answer: {chr(97 + correct_position)}"
            pdf_solution.multi_cell(0, 10, txt=solution_text)
            
        # Creazione dei file .tex
        questions = []
        for index, row in mcq_sample.iterrows():
            shuffled_options, correct_position = shuffle_mcq_choices(row)
            question_text = row['Question']
            choices = [{'letter': chr(97 + idx), 'text': opt} for idx, opt in enumerate(shuffled_options)]
            questions.append({
                'type': 'mcq',
                'text': question_text,
                'choices': choices,
                'solution': chr(97 + correct_position)
            })

        for index, row in open_sample.iterrows():
            question_text = row['Questions']
            questions.append({'type': 'open', 'text': question_text})

        with open(f"Exam_Copy_{copy_num}.tex", 'w') as f:
            f.write(template.render(questions=questions, copy_num=copy_num))

        for index, row in open_sample.iterrows():
                    question_text = f"{index + 18}. {row['Questions']}"
                    pdf_exam.multi_cell(0, 10, txt=question_text)
        
                    # Check and insert image if present
                    image_path = row.get('Image_Path', None)
                    if pd.notna(image_path) and os.path.exists(image_path):
                        img_width = 90
                        img_height = get_image_height(image_path, img_width)
                        if space_left(pdf_exam) < img_height:
                            pdf_exam.add_page()
                        pdf_exam.image(image_path, x=10, y=pdf_exam.get_y(), w=img_width)
                        pdf_exam.ln(img_height + 5)

        pdf_exam.output(f"Exam_Copy_{copy_num}.pdf")
        pdf_solution.output(f"Solution_Copy_{copy_num}.pdf")

def main():
    mcq_df = pd.read_excel("mcq_input1.xlsx", engine='openpyxl')
    open_df = pd.read_excel("open_input1.xlsx", engine='openpyxl')

    num_copies = int(input("Enter the number of exam copies you want to create: "))
    
    # Ask the user for the number of MCQs and open questions
    mcq_count = int(input("Enter the number of multiple-choice questions you want in each exam: "))
    open_count = int(input("Enter the number of open questions you want in each exam: "))

    create_exam_and_solution(mcq_df, open_df, num_copies, mcq_count, open_count)

if __name__ == "__main__":
    main()
