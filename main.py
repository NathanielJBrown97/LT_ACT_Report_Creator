import os
import csv
import textwrap
from collections import Counter

import pandas as pd
import matplotlib.pyplot as plt

from reportlab.lib.pagesizes import letter
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import inch, mm
from reportlab.platypus import (BaseDocTemplate, Frame, PageTemplate, Paragraph, Spacer, Table, TableStyle, Image, PageBreak)
from reportlab.lib import colors


def create_CSV_raw(input_file, output_file):
    with open(input_file, 'r', newline='') as csvfile:
        reader = list(csv.reader(csvfile))
    if reader:
        reader.pop(0)
    if len(reader) > 215:
        reader = reader[:215]
    for i in range(len(reader)):
        reader[i] = [field for field in reader[i] if field != '']
    with open(output_file, 'w', newline='') as csvfile:
        writer = csv.writer(csvfile)
        writer.writerows(reader)

def create_CSV_summary(input_file, output_file):
    with open(input_file, 'r', newline='') as csvfile:
        reader = list(csv.reader(csvfile))
    if len(reader) > 1:
        reader = reader[:1] + reader[216:]
    for i in range(len(reader)):
        reader[i] = [field for field in reader[i] if field != '']
    with open(output_file, 'w', newline='') as csvfile:
        writer = csv.writer(csvfile)
        writer.writerows(reader)

def load_summary_csv(file_path):
    with open(file_path, 'r') as csvfile:
        reader = csv.reader(csvfile)
        data = list(reader)

    student_info = data[0]
    student_summary = {
        "Name": student_info[0],
        "Form Number": student_info[1],
        "Date": student_info[2],
        "English Score": student_info[3],
        "Math Score": student_info[4],
        "Reading Score": student_info[5],
        "Science Score": student_info[6],
        "Composite Score": student_info[7]
    }

    sections = {}
    current_section = None
    current_status = None

    for row in data[1:-5]: 
        if row[1] == '--' and row[2] == '--':
            current_section = row[0].replace(' correct', '').replace(' wrong', '').strip().lower()
            sections.setdefault(current_section, {"correct": [], "wrong": []})
            current_status = 'correct' if 'correct' in row[0] else 'wrong'
        elif current_section and len(row) >= 3:
            sections[current_section][current_status].append({
                "Topic": row[0],
                "Number": int(row[1]),
                "Total": int(row[2])
            })

    final_rows = data[-5:]
    student_summary["English_Right"] = int(final_rows[0][0])
    student_summary["English_Total"] = int(final_rows[0][1])
    student_summary["Math_Right"] = int(final_rows[1][0])
    student_summary["Math_Total"] = int(final_rows[1][1])
    student_summary["Reading_Right"] = int(final_rows[2][0])
    student_summary["Reading_Total"] = int(final_rows[2][1])
    student_summary["Science_Right"] = int(final_rows[3][0])
    student_summary["Science_Total"] = int(final_rows[3][1])

    percentiles = {
        "English Percentile": final_rows[4][0],
        "Math Percentile": final_rows[4][1],
        "Reading Percentile": final_rows[4][2],
        "Science Percentile": final_rows[4][3],
        "Composite Percentile": final_rows[4][4]
    }

    summary_df = pd.DataFrame([student_summary])
    summary_df = summary_df.assign(**percentiles)

    section_dfs = {}
    for section, data in sections.items():
        correct_df = pd.DataFrame(data['correct'])
        wrong_df = pd.DataFrame(data['wrong'])
        section_dfs[section] = {'correct': correct_df, 'wrong': wrong_df}

    return summary_df

def process_student_csv(input_file, output_folder):
    student_name = os.path.splitext(os.path.basename(input_file))[0].replace('Results Data for Students - ', '').strip()
    raw_output_file = 'RAW_student.csv'
    summary_output_file = 'SUM_student.csv'
    pdf_file_name = os.path.join(output_folder, f"{student_name}Report.pdf")
    
    create_CSV_raw(input_file, raw_output_file)
    create_CSV_summary(input_file, summary_output_file)

    column_names = ['Section', 'Question Number', 'Correct Answer', 'Student Answer', 'Topic Tested', 'Topic Tested 2', 'Topic Tested 3', 'Topic Tested 4']
    df = pd.read_csv(raw_output_file, names=column_names, header=None, nrows=215, skip_blank_lines=True)
    summary_df = load_summary_csv(summary_output_file)

    pdf_path = pdf_file_name
    doc = BaseDocTemplate(pdf_path, pagesize=letter)
    elements = []

    elements.append(Spacer(1, 80))

    student_info_table_data = [
        ['Student:', summary_df.at[0, 'Name']],
        ['Test Type:', 'ACT'],
        ['Test Form:', summary_df.at[0, 'Form Number']],
        ['Date:', summary_df.at[0, 'Date']]
    ]
    student_info_table = Table(student_info_table_data, colWidths=[1.5*inch, 4*inch])
    student_info_table.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#b29600')),
        ('BACKGROUND', (0, 1), (-1, -1), colors.beige),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.HexColor('#0e1027')),
        ('TEXTCOLOR', (0, 1), (-1, -1), colors.black),
        ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
        ('FONTNAME', (0, 0), (-1, -1), 'Helvetica-Bold'),
        ('FONTSIZE', (0, 0), (-1, -1), 12),
        ('TOPPADDING', (0, 0), (-1, -1), 6),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 6),
        ('LEFTPADDING', (0, 0), (-1, -1), 12),
        ('RIGHTPADDING', (0, 0), (-1, -1), 12),
        ('GRID', (0, 0), (-1, -1), 1, colors.black),
    ]))

    styled_box = Table([[student_info_table]], colWidths=[7*inch])
    styled_box.setStyle(TableStyle([
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ('LEFTPADDING', (0, 0), (-1, -1), 12),
        ('RIGHTPADDING', (0, 0), (-1, -1), 12),
        ('TOPPADDING', (0, 0), (-1, -1), 12),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 12),
        ('BACKGROUND', (0, 0), (-1, -1), colors.beige),
        ('BOX', (0, 0), (-1, -1), 2, colors.HexColor('#b29600')),
        ('GRID', (0, 0), (-1, -1), 1, colors.HexColor('#0e1027')),
    ]))

    elements.append(Spacer(1, 180))  
    elements.append(styled_box)
    elements.append(Spacer(1, 48))


    composite_score = summary_df.at[0, 'Composite Score']

    composite_score_table_data = [
        ['Composite Score:', composite_score]
    ]
    composite_score_table = Table(composite_score_table_data, colWidths=[3*inch, 1*inch])
    composite_score_table.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#b29600')),
        ('BACKGROUND', (0, 1), (-1, -1), colors.beige),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.HexColor('#0e1027')),
        ('TEXTCOLOR', (0, 1), (-1, -1), colors.black),
        ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
        ('FONTNAME', (0, 0), (-1, -1), 'Helvetica-Bold'),
        ('FONTSIZE', (0, 0), (-1, -1), 12),
        ('TOPPADDING', (0, 0), (-1, -1), 6),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 6),
        ('LEFTPADDING', (0, 0), (-1, -1), 12),
        ('RIGHTPADDING', (0, 0), (-1, -1), 12),
        ('GRID', (0, 0), (-1, -1), 1, colors.black),
    ]))

    composite_score_box = Table([[composite_score_table]], colWidths=[4*inch])
    composite_score_box.setStyle(TableStyle([
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ('LEFTPADDING', (0, 0), (-1, -1), 12),
        ('RIGHTPADDING', (0, 0), (-1, -1), 12),
        ('TOPPADDING', (0, 0), (-1, -1), 12),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 12),
        ('BACKGROUND', (0, 0), (-1, -1), colors.beige),
        ('BOX', (0, 0), (-1, -1), 2, colors.HexColor('#b29600')),
        ('GRID', (0, 0), (-1, -1), 1, colors.HexColor('#0e1027')),
    ]))

    elements.append(Spacer(1, 24))
    elements.append(composite_score_box)
    elements.append(Spacer(1, 24))
    elements.append(PageBreak())

    total_english_questions = 75
    total_math_questions = 60
    total_reading_questions = 40
    total_science_questions = 40

    try:
        english_df = df[df['Section'] == 'English']
        math_df = df[df['Section'] == 'Math']
        reading_df = df[df['Section'] == 'Reading']
        science_df = df[df['Section'] == 'Science']

        english_score = english_df[english_df['Correct Answer'] == english_df['Student Answer']].shape[0]
        math_score = math_df[math_df['Correct Answer'] == math_df['Student Answer']].shape[0]
        reading_score = reading_df[reading_df['Correct Answer'] == reading_df['Student Answer']].shape[0]
        science_score = science_df[science_df['Correct Answer'] == science_df['Student Answer']].shape[0]

        english_percentage = (english_score / total_english_questions) * 100
        math_percentage = (math_score / total_math_questions) * 100
        reading_percentage = (reading_score / total_reading_questions) * 100
        science_percentage = (science_score / total_science_questions) * 100

        english_act_score = summary_df.at[0, "English Score"]
        math_act_score = summary_df.at[0, "Math Score"]
        reading_act_score = summary_df.at[0, "Reading Score"]
        science_act_score = summary_df.at[0, "Science Score"]

        generate_bar_chart(['English', 'Math'], [english_percentage, math_percentage], [english_score, math_score], 'test_results_chart1.png')
        generate_bar_chart(['Reading', 'Science'], [reading_percentage, science_percentage], [reading_score, science_score], 'test_results_chart2.png')

        add_section_to_pdf(
            elements,
            'English and Math',
            ("The first two sections of the test evaluate your knowledge of grammar rules and math. Learn the material and "
            "learn how to apply it to improve your score. Focusing on what you know and what you don't know "
            "will help you improve these scores!"),
            'test_results_chart1.png',
            [
                ['Section', 'Correct', 'Score'],
                ['English', f'{english_score}/{total_english_questions}', english_act_score],
                ['Math', f'{math_score}/{total_math_questions}', math_act_score]
            ],
            {
                'English Percentile': summary_df.at[0, 'English Percentile'],
                'Math Percentile': summary_df.at[0, 'Math Percentile']
            }
        )

        add_section_to_pdf(
            elements,
            'Reading and Science',
            ("The final two sections of the test are very similar. Reading is testing your reading comprehension, so take  "
            "the time to read and understand the passages! Science is not testing your knowledge of science. Focus on the "
            "questions and the graphs primarily. A small number of questions will require you to dig into the text and very "
            "few will actually require you to understand it. Lastly, outside knowledge is very rare, "
             "so don't go digging through your old biology or chemistry books!"),
            'test_results_chart2.png',
            [
                ['Section', 'Correct', 'Score'],
                ['Reading', f'{reading_score}/{total_reading_questions}', reading_act_score],
                ['Science', f'{science_score}/{total_science_questions}', science_act_score]
            ],
            {
                'Reading Percentile': summary_df.at[0, 'Reading Percentile'],
                'Science Percentile': summary_df.at[0, 'Science Percentile']
            }
        )
        elements.append(PageBreak())
        df = df.dropna(subset=['Topic Tested'])

        missed_topics_summary_by_section = {
            'English': Counter(english_df[english_df['Correct Answer'] != english_df['Student Answer']]['Topic Tested']),
            'Math': Counter(math_df[math_df['Correct Answer'] != math_df['Student Answer']]['Topic Tested']),
            'Reading': Counter(reading_df[reading_df['Correct Answer'] != reading_df['Student Answer']]['Topic Tested']),
            'Science': Counter(science_df[science_df['Correct Answer'] != science_df['Student Answer']]['Topic Tested'])
        }

        add_detailed_results_section(elements, 'English', english_df, english_act_score, missed_topics_summary_by_section)
        add_detailed_results_section(elements, 'Math', math_df, math_act_score, missed_topics_summary_by_section)
        add_detailed_results_section(elements, 'Reading', reading_df, reading_act_score, missed_topics_summary_by_section)
        add_detailed_results_section(elements, 'Science', science_df, science_act_score, missed_topics_summary_by_section)

        frame = Frame(doc.leftMargin, doc.bottomMargin, doc.width, doc.height, id='normal')

        doc.addPageTemplates([PageTemplate(id='PageNumbers', frames=frame, onPage=add_page_style)])

        doc.build(elements)
        print(f"PDF report generated at: {pdf_path}")

    except KeyError as e:
        print(f"KeyError: {e}")

def add_page_style(canvas, doc):
    page_num = canvas.getPageNumber()
    text = f"Page {page_num}"
    
    canvas.saveState()

    canvas.setFillColor("#f4f4f4")
    canvas.rect(0, 0, doc.pagesize[0], doc.pagesize[1], stroke=0, fill=1)

    canvas.restoreState()
    
    canvas.saveState()

    canvas.setFont('Helvetica-Bold', 10)
    if page_num == 1:
        banner_path = 'images/Linkedin Banner-01.jpg'
        banner_width = doc.width + 145
        desired_height_inch = 1.25
        banner_height = desired_height_inch * inch
        banner_x = 0
        banner_y = doc.pagesize[1] - banner_height
        canvas.drawImage(banner_path, banner_x, banner_y, width=banner_width, height=banner_height)

    if page_num > 1:
        canvas.drawString(doc.leftMargin, 10 * mm, text)

        logo_path = 'images/Lee Tutoring Logo classic.png'
        logo_x = doc.width + doc.rightMargin - 1.9 * inch
        logo_y = 10 * mm
        logo_width = 2.5 * inch
        logo_height = 0.5 * inch
        canvas.drawImage(logo_path, logo_x, logo_y, width=logo_width, height=logo_height, mask='auto')

    contact_info_email = "hello@leetutoring.com"
    contact_info_phone = "727.543.8217"
    canvas.setFont('Helvetica-Bold', 8)
    contact_x = doc.width / 2.0 + doc.leftMargin
    contact_y_email = 15 * mm
    contact_y_phone = 12 * mm
    canvas.drawCentredString(contact_x, contact_y_email, contact_info_email)
    canvas.drawCentredString(contact_x, contact_y_phone, contact_info_phone)

    canvas.restoreState()



def generate_bar_chart(sections, percentages, scores, file_name):
    plt.figure(figsize=(3, 1.5), dpi=300)

    sections_ordered = sections[::-1]
    percentages_ordered = percentages[::-1]
    scores_ordered = scores[::-1]

    fig, ax = plt.subplots(figsize=(3, 1.5), dpi=300)
    

    fig.patch.set_facecolor('#f4f4f4')
    ax.set_facecolor('#f4f4f4')

    ax.barh(sections_ordered, percentages_ordered, color=['#b29600', '#0e1027'], height=0.8)
    for index, value in enumerate(percentages_ordered):
        ax.text(value, index, f'{scores_ordered[index]}', va='center', fontsize=6)

    ax.set_xlabel('Percentage Correct', fontsize=6)
    ax.set_xlim(0, 100)
    ax.set_yticks(range(len(sections_ordered)))
    ax.set_yticklabels(sections_ordered, fontsize=6)
    ax.set_xticks(range(0, 101, 10))
    ax.tick_params(axis='x', labelsize=6)
    ax.tick_params(axis='y', labelsize=6)

    plt.subplots_adjust(left=0.2, right=0.95, top=0.85, bottom=0.15)
    plt.savefig(file_name, bbox_inches='tight')
    plt.close()

def add_section_to_pdf(elements, header, text, chart_file_path, data, percentiles):
    header_style = ParagraphStyle('Header1', fontName='Helvetica-Bold', fontSize=18, textColor=colors.HexColor('#0e1027'))
    paragraph_style = ParagraphStyle('BodyText', fontName='Helvetica', fontSize=12, textColor=colors.black)

    elements.append(Paragraph(header, header_style))
    elements.append(Spacer(1, 12))

    wrapped_text = textwrap.fill(text, width=80)
    elements.append(Paragraph(wrapped_text, paragraph_style))
    elements.append(Spacer(1, 12))

    im = Image(chart_file_path)
    im.drawHeight = 2.5 * inch
    im.drawWidth = 3.5 * inch

    data[0].append('Percentile')  # Add Percentile column header

    for row in data[1:]:
        section_name = row[0]
        percentile_key = f"{section_name} Percentile"
        if percentile_key in percentiles:
            row.append(percentiles[percentile_key])
        else:
            row.append("N/A")

    table = Table(data)
    table.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#b29600')),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.HexColor('#0e1027')),
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
        ('FONTSIZE', (0, 0), (-1, 0), 12),
        ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
        ('BACKGROUND', (0, 1), (-1, -1), colors.beige),
        ('GRID', (0, 0), (-1, -1), 1, colors.black)
    ]))

    combined = Table(
        [[im, table]],
        colWidths=[3.5 * inch, 3.5 * inch] 
    )
    combined.setStyle(TableStyle([
        ('VALIGN', (0, 0), (-1, -1), 'TOP'),
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('LEFTPADDING', (1, 0), (1, -1), -20),  
        ('RIGHTPADDING', (0, 0), (0, -1), -20),  
        ('TOPPADDING', (1, 0), (1, -1), 14.5), 
        ('BOTTOMPADDING', (0, 0), (-1, -1), 6),  
        ('BACKGROUND', (0, 0), (-1, -1), colors.HexColor('#f4f4f4')),
    ]))

    elements.append(combined)
    elements.append(Spacer(1, 24))
    



def add_missed_questions_table(elements, section_name, missed_topics_summary_by_section, df):
    styles = getSampleStyleSheet()
    header_style = styles['Heading1']
    header_style.textColor = colors.HexColor('#0e1027')
    paragraph_style = styles['BodyText']
    paragraph_style.textColor = colors.black

    detailed_table_data = [['Topic', '# Missed', 'Missed Questions']]
    missed_topics_summary = missed_topics_summary_by_section[section_name]

    sorted_missed_topics = sorted(missed_topics_summary.items(), key=lambda x: x[1], reverse=True)

    for topic, count in sorted_missed_topics:
        questions = [f"{df.loc[i, 'Question Number']}" for i in df[(df['Correct Answer'] != df['Student Answer']) & (df['Topic Tested'] == topic) & (df['Section'] == section_name)].index]
        for i in range(0, len(questions), 11):
            chunk = questions[i:i + 11]
            detailed_table_data.append([topic if i == 0 else '', count if i == 0 else '', ", ".join(chunk)])

    detailed_table = Table(detailed_table_data, colWidths=[2 * inch, 1 * inch, 4.5 * inch])
    detailed_table.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#b29600')),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.HexColor('#0e1027')),
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
        ('FONTSIZE', (0, 0), (-1, 0), 12),
        ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
        ('BACKGROUND', (0, 1), (-1, -1), colors.beige),
        ('GRID', (0, 0), (-1, -1), 1, colors.black),
        ('VALIGN', (0, 0), (-1, -1), 'TOP')
    ]))

    for row in detailed_table_data[1:]:
        row[2] = Paragraph(row[2], styles['BodyText'])

    elements.append(detailed_table)
    elements.append(Spacer(1, 24))


def add_points_loss_section(elements, missed_topics_summary_by_section, df):
    styles = getSampleStyleSheet()
    header_style = styles['Heading1']
    header_style.textColor = colors.HexColor('#0e1027')
    paragraph_style = styles['BodyText']
    paragraph_style.textColor = colors.black

    missed_questions_by_section = {
        'English': [],
        'Math': [],
        'Reading': [],
        'Science': []
    }

    section_mapping = {
        'English': '1',
        'Math': '2',
        'Reading': '3',
        'Science': '4'
    }

    for section, missed_topics_summary in missed_topics_summary_by_section.items():
        for topic, count in missed_topics_summary.items():
            questions = [f"{section_mapping[section]}.{row['Question Number']}" for _, row in df[(df['Correct Answer'] != df['Student Answer']) & (df['Topic Tested'] == topic) & (df['Section'] == section)].iterrows()]
            for i in range(0, len(questions), 11):
                chunk = questions[i:i + 11]
                missed_questions_by_section[section].append([topic if i == 0 else '', count if i == 0 else '', ", ".join(chunk)])

    for section in ['English', 'Math', 'Reading', 'Science']:
        if missed_questions_by_section[section]:
            section_title = section + " Missed Questions"
            elements.append(Paragraph(section_title, header_style))
            elements.append(Spacer(1, 12))

            detailed_table_data = [['Topic', '# Missed', 'Missed Questions']]
            detailed_table_data += missed_questions_by_section[section]

            detailed_table = Table(detailed_table_data, colWidths=[2 * inch, 1 * inch, 4.5 * inch])
            detailed_table.setStyle(TableStyle([
                ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#b29600')),
                ('TEXTCOLOR', (0, 0), (-1, 0), colors.HexColor('#0e1027')),
                ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
                ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
                ('FONTSIZE', (0, 0), (-1, 0), 12),
                ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
                ('BACKGROUND', (0, 1), (-1, -1), colors.beige),
                ('GRID', (0, 0), (-1, -1), 1, colors.black),
                ('VALIGN', (0, 0), (-1, -1), 'TOP')
            ]))

            for row in detailed_table_data[1:]:
                row[2] = Paragraph(row[2], styles['BodyText'])

            elements.append(detailed_table)
            elements.append(Spacer(1, 24))

    elements.append(PageBreak())


def add_where_did_i_lose_points_section(elements, section_name):
    styles = getSampleStyleSheet()
    header_style = styles['Heading1']
    header_style.textColor = colors.HexColor('#0e1027')
    paragraph_style = styles['BodyText']
    paragraph_style.textColor = colors.black

    text = (
        "Enhancing your score requires focusing on the areas where you made mistakes. "
        "The following section details each type of problem you missed, the points you lost, and the specific questions involved. "
        "Use this analysis to tailor your study sessions and improve your performance on future tests."
    )
    wrapped_text = textwrap.fill(text, width=80)
    elements.append(Paragraph(wrapped_text, paragraph_style))
    elements.append(Spacer(1, 12))

    smaller_title_style = ParagraphStyle('SmallerTitle', fontName='Helvetica-Bold', fontSize=14, textColor=colors.HexColor('#0e1027'), alignment=1)
    elements.append(Paragraph("Analysis by Topic", smaller_title_style))
    elements.append(Spacer(1, 12))

def add_detailed_results_section(elements, section_name, df, act_score, missed_topics_summary_by_section):
    add_section_overview(elements, section_name, df, act_score)
    add_where_did_i_lose_points_section(elements, section_name)
    add_missed_questions_table(elements, section_name, missed_topics_summary_by_section, df)
    
    df = df.fillna('')
    df['Combined Topics'] = df[['Topic Tested', 'Topic Tested 2', 'Topic Tested 3', 'Topic Tested 4']].apply(lambda x: ', '.join(filter(None, x)), axis=1)
    data = df[['Question Number', 'Correct Answer', 'Student Answer', 'Combined Topics']].values.tolist()
    max_rows_per_table = 80
    data_chunks = [data[i:i + max_rows_per_table] for i in range(0, len(data), max_rows_per_table)]
    
    highlight_style = ParagraphStyle('Highlight', textColor=colors.red, fontName='Helvetica-Bold', alignment=1)
    header_style = ParagraphStyle('Header', textColor=colors.black, fontName='Helvetica-Bold', alignment=1)
    smaller_title_style = ParagraphStyle('SmallerTitle', fontName='Helvetica-Bold', fontSize=14, textColor=colors.HexColor('#0e1027'), alignment=1)
    elements.append(Spacer(1, 24))

    elements.append(Paragraph("Question-by-Question Analysis", smaller_title_style))
    elements.append(Spacer(1, 12))

    for chunk in data_chunks:
        table_data = [[Paragraph('<b>#</b>', header_style), Paragraph('<b>K</b>', header_style), Paragraph('<b>S</b>', header_style), Paragraph('<b>Topics</b>', header_style)]]
        for row in chunk:
            question_number, correct_answer, student_answer, topics = row
            if correct_answer != student_answer:
                table_data.append([
                    Paragraph(f'<b>{question_number}</b>', highlight_style),
                    Paragraph(f'<b>{correct_answer}</b>', highlight_style),
                    Paragraph(f'<b>{student_answer}</b>', highlight_style),
                    Paragraph(f'<b>{topics}</b>', highlight_style)
                ])
            else:
                table_data.append(row)

        table = Table(table_data, colWidths=[0.8 * inch, 0.6 * inch, 0.6 * inch, 3.8 * inch])
        table.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#b29600')),
            ('TEXTCOLOR', (0, 0), (-1, 0), colors.HexColor('#0e1027')),
            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
            ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
            ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
            ('FONTSIZE', (0, 0), (-1, 0), 10),
            ('BOTTOMPADDING', (0, 0), (-1, 0), 6),
            ('BACKGROUND', (0, 1), (-1, -1), colors.beige),
            ('GRID', (0, 0), (-1, -1), 1, colors.black),
            ('LEFTPADDING', (0, 0), (-1, -1), 4),
            ('RIGHTPADDING', (0, 0), (-1, -1), 4),
            ('TOPPADDING', (0, 0), (-1, -1), 4),
            ('BOTTOMPADDING', (0, 0), (-1, -1), 4)
        ]))

        elements.append(table)
        elements.append(Spacer(1, 12))
    elements.append(PageBreak())




def add_section_overview(elements, section_name, df, act_score):
    correct_count = df[df['Correct Answer'] == df['Student Answer']].shape[0]
    incorrect_count = df.shape[0] - correct_count
    incorrect_df = df[df['Correct Answer'] != df['Student Answer']]
    topic_counts = Counter(incorrect_df['Topic Tested'])
    most_lost_topics = topic_counts.most_common(3)

    overview_data = [
        [f'{section_name} Result', act_score],
        ['Correct Questions', correct_count],
        ['Incorrect Questions', incorrect_count]
    ]

    overview_table = Table(overview_data)
    overview_table.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#b29600')),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.HexColor('#0e1027')),
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
        ('FONTSIZE', (0, 0), (-1, 0), 12),
        ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
        ('BACKGROUND', (0, 1), (-1, -1), colors.beige),
        ('GRID', (0, 0), (-1, -1), 1, colors.black)
    ]))

    topics_data = [['Top Problem Topics', 'Missed']] + [[topic, count] for topic, count in most_lost_topics]
    topics_table = Table(topics_data)
    topics_table.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#b29600')),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.HexColor('#0e1027')),
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
        ('FONTSIZE', (0, 0), (-1, 0), 12),
        ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
        ('BACKGROUND', (0, 1), (-1, -1), colors.beige),
        ('GRID', (0, 0), (-1, -1), 1, colors.black)
    ]))

    combined_table = Table([[overview_table, topics_table]])
    combined_table.setStyle(TableStyle([
        ('LEFTPADDING', (0, 0), (-1, -1), 6),
        ('RIGHTPADDING', (0, 0), (-1, -1), 6),
        ('TOPPADDING', (0, 0), (-1, -1), 6),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 6),
        ('VALIGN', (0, 0), (-1, -1), 'TOP')
    ]))

    header_style = ParagraphStyle('Header1', fontName='Helvetica-Bold', fontSize=18, textColor=colors.HexColor('#0e1027'))
    elements.append(Paragraph(f'{section_name} Section Analysis', header_style))
    elements.append(Spacer(1, 24))

    elements.append(combined_table)
    elements.append(Spacer(1, 12))


def main():
    input_folder = 'SensitiveStudentData'
    output_folder = 'StudentReports'
    
    if not os.path.exists(output_folder):
        os.makedirs(output_folder)

    for filename in os.listdir(input_folder):
        if filename.endswith('.csv'):
            input_file = os.path.join(input_folder, filename)
            process_student_csv(input_file, output_folder)

    temp_files = ['RAW_student.csv', 'SUM_student.csv', 'test_results_chart1.png', 'test_results_chart2.png']
    for temp_file in temp_files:
        if os.path.exists(temp_file):
            os.remove(temp_file)

main()
