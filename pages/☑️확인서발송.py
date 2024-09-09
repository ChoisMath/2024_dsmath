import pandas as pd
import numpy as np
import streamlit as st
from utils import client
from utils import send_mail2
from reportlab.pdfbase import pdfmetrics
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas
from reportlab.pdfbase.ttfonts import TTFont
from io import BytesIO

def data_load(sheet_name):
    all_data = pd.DataFrame(sheet_name.get_all_values())
    all_data.columns = list(all_data.iloc[0])
    all_data = all_data.iloc[1:]
    return all_data

def conditional_filter(data,
                       school_name = None,
                       grade_num = None,
                       student_ban = None,
                       student_id = None,
                       student_name =None):
    filtered_data = data
    if school_name is not None:
        filtered_data = filtered_data[filtered_data["학교명"]==school_name]
    if grade_num is not None:
        filtered_data = filtered_data[filtered_data['학년']==str(grade_num)]
    if student_ban is not None:
        filtered_data = filtered_data[filtered_data['반']==str(student_ban)]
    if student_id is not None:
        filtered_data = filtered_data[filtered_data['번호']==str(student_id)]
    if student_name is not None:
        filtered_data = filtered_data[filtered_data['이름']==student_name]

    return filtered_data

def create_certificate(school, balgub, grade, ban, bun, name):
    # Register the font
    font_path = "./malgun.ttf"
    pdfmetrics.registerFont(TTFont("맑은고딕", font_path))
    font_path2 = "./malgunbd.ttf"
    pdfmetrics.registerFont(TTFont("맑은고딕굵게", font_path2))

    # Create a BytesIO buffer
    buffer = BytesIO()

    # Create PDF
    pdf = canvas.Canvas(buffer, pagesize=A4)
    width, height = A4

    # Add title
    pdf.setFont("맑은고딕", 18)
    title = "2024년 대구수학페스티벌 체험 확인서"
    str_width = pdf.stringWidth(title, "맑은고딕", 18)
    pdf.drawString((width // 2) - (str_width // 2), 750, title)

    # Add table data
    data = [
        ["학교", school, "발급번호", balgub],
        ["학년", grade, "반", ban, "번호", bun],
        ["이름", name]
    ]

    # Table starting position
    row_height = 20
    col_widths = [50, 100, 70, 100, 50, 50]
    x_offset = (width // 2) - (sum(col_widths) // 2)
    y_offset = 700

    # Add BOX
    pdf.setLineWidth(0.5)

    #윗줄 - 위
    pdf.line(x_offset, y_offset + row_height,
             x_offset + sum(col_widths), y_offset + row_height)
    #윗줄 - 아래
    pdf.line(x_offset + 3, y_offset + row_height - 3,
             x_offset + sum(col_widths) - 3, y_offset + row_height - 3)
    #아랫줄 - 아래
    pdf.line(x_offset, y_offset - len(data) * row_height,
             x_offset + sum(col_widths), y_offset - len(data) * row_height)
    #아랫줄 - 위
    pdf.line(x_offset + 3, y_offset - len(data) * row_height + 3,
             x_offset + sum(col_widths) - 3, y_offset - len(data) * row_height + 3)
    #왼쪽 - 바깥
    pdf.line(x_offset, y_offset + row_height,
             x_offset, y_offset - len(data) * row_height)
    #왼쪽 - 안
    pdf.line(x_offset + 3, y_offset + row_height - 3,
             x_offset + 3, y_offset - len(data) * row_height + 3)
    #오른쪽 - 바깥
    pdf.line(x_offset + sum(col_widths), y_offset + row_height,
             x_offset + sum(col_widths), y_offset - len(data) * row_height)
    #오른쪽 - 안
    pdf.line(x_offset + sum(col_widths) - 3, y_offset + row_height - 3,
             x_offset + sum(col_widths) - 3, y_offset - len(data) * row_height + 3)

    # Draw table
    pdf.setFont("맑은고딕", 12)
    for row in data:
        for i, cell in enumerate(row):
            pdf.drawString(x_offset + 15 + sum(col_widths[:i]), y_offset, str(cell))
        y_offset -= 1.2 * row_height

    # Add description
    description1 = "위 학생은 2024년 대구수학페스티벌의 행사에 적극적으로 참여하여"
    description2 = "체험 프로그램을 성실하게 이수하였음을 증명합니다."
    y_offset -= row_height
    pdf.drawString(x_offset + 15, y_offset, description1)
    y_offset -= row_height
    pdf.drawString(x_offset + 15, y_offset, description2)

    # Add footer
    pdf.setFont('맑은고딕굵게', 20)
    footer = "대구중등수학연구회 회장 문 희 정"
    str_width = pdf.stringWidth(footer, "맑은고딕굵게", 20)
    x_offset = (width // 2) - (str_width // 2)
    pdf.drawString(x_offset, y_offset - 50, footer)

    # Add stamp image
    stamp_image_path = "./dojang.png"
    pdf.drawImage(stamp_image_path, 422, y_offset - 65, width=50, height=50)



    # Finalize the PDF and get the BytesIO buffer
    pdf.save()
    buffer.seek(0)
    pdf_output = buffer.getvalue()
    buffer.close()

    return pdf_output

def main():
    using_spreadsheet_id = '1DwMKa9x9mHZnKUFgylhgQahEoFaTmfHCr4yeCVNVpT4'
    spreadsheet = client.open_by_key(using_spreadsheet_id)
    chehumsheet = spreadsheet.worksheet('2024dgmath')

    st.subheader('2024. 대구수학페스티벌 참가자 확인')
    chehumdata = data_load(chehumsheet)
    approved_data = chehumdata[chehumdata['승인']=="TRUE"]
    # st.dataframe(approved_data)

    row1 = st.columns(2)
    school_name = row1[0].selectbox('학교명', options=np.unique(np.array(approved_data['학교명'])))
    if school_name == "":
        school_name = None

    student_name = row1[1].text_input("이름")
    if student_name == "":
        student_name = None

    row2 = st.columns(3)
    grade_num = row2[0].number_input("학년", min_value=0, max_value=6, value=0)
    if grade_num == 0:
        grade_num = None

    student_ban = row2[1].number_input("반", step=1, min_value=0, max_value=15, value=0, placeholder="선생님은 o반")
    if student_ban == 0:
        student_ban = None

    student_id = row2[2].number_input("번호", step=1, min_value=0, max_value=40, value=0, placeholder="선생님은 o번")
    if student_id == 0:
        student_id = None

    filtered_data = conditional_filter(approved_data,
                                       school_name=school_name,
                                       grade_num=grade_num,
                                       student_ban=student_ban,
                                       student_id=student_id,
                                       student_name=student_name)
    forstudata = filtered_data[["학교명", "발급번호", "학년", "반", "번호", "이름", "E-mail"]]#.values.tolist()
    st.dataframe(forstudata, use_container_width=True)

    row3 = st.columns([0.7, 1.2, 0.5])
    set_serial = row3[0].button("메일발송")
    if set_serial:
         # 일련번호에서 NaN 값을 제외한 후 최대값 계산
        serials = approved_data['일련번호'].tolist()
        serials = [int(x) for x in serials if x is not ""]
        max_serial = int(max(serials))
        insert_index = approved_data[approved_data['일련번호'] == ""].index
        for i in range(len(insert_index)):
            chehumsheet.update(range_name="J" + str(insert_index[i] + 1), values=[[max_serial + i + 1]])

        datalist = forstudata.values.tolist()
        if len(datalist) != 1:
            st.warning("검색된 대상이 1명이 아닙니다. 조건을 상세히 입력하세요.")
        else:
            school, balgub, grade, ban, bun, name, email = datalist[0]
            # using_spreadsheet_id = "1DwMKa9x9mHZnKUFgylhgQahEoFaTmfHCr4yeCVNVpT4"
            # gid = 37733787
            # sheet = client.open_by_key(using_spreadsheet_id).worksheet("확인서")
            # sheet.update_cell(row=2, col=2, value=school)
            # sheet.update_cell(row=2, col=6, value=balgub)
            # sheet.update_cell(row=3, col=2, value=grade)
            # sheet.update_cell(row=3, col=4, value=ban)
            # sheet.update_cell(row=3, col=6, value=bun)
            # sheet.update_cell(row=3, col=8, value=name)
            #
            # pdf_data = get_pdfdata_from_sheet(using_spreadsheet_id, gid)
            # send_mail(pdf_data, fromaddr="complete860127@gmail.com",
            #           toaddr=email, school_name=school_name, name=name)


            #고친거 해보자
            pdf_output = create_certificate(school, balgub, grade, ban, bun, name)
            send_mail2(pdf_output, email, school, name)
            st.success('Email sent successfully!')
    else:
        st.success("검색된 대상이 1명일 경우에만 발송됩니다.")


if __name__ == "__main__":
    main()
