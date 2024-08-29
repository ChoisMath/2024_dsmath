import pandas as pd
import streamlit as st
import datetime
import json
import re
from utils import chehumsheet
from utils import insert_row



def data_load(sheet_name):
    all_data = pd.DataFrame(sheet_name.get_all_values())
    all_data.columns = list(all_data.iloc[0])
    all_data = all_data.iloc[1:]
    return all_data


def detect_same_index(df, input_list):
    data = df.values[:,1:6].tolist()
    input_list = [str(x) for x in input_list]
    if input_list in data:
        return data.index(input_list)+2
    else:
        return None

if "email" not in st.session_state:
    st.session_state.email = ""
def update_email():
    st.session_state.email = st.session_state.email_input

def main():
    st.title("2024 대수페 체험활동 현장신청")
    row1 = st.columns([6,2,2,2,3])
    school_name = row1[0].text_input("학교명")  # selectbox로 변경
    grade_num = row1[1].number_input("학년", min_value=1, max_value=6, value=1)
    student_ban = row1[2].number_input("반", step=1, min_value=1, max_value=15, value=1)
    student_id = row1[3].number_input("번호", step=1, min_value=1, max_value=40, value=1)
    student_name = row1[4].text_input("이름")


    row2 = st.columns(2)

    birthdate = row2[0].date_input(label="생년월일", value=datetime.date(2007, 1, 1))
    json_birth = json.dumps(birthdate, default=str).strip("\"")
    email_addr = row2[1].text_input(label="E-mail(입력 후 꼭!! 확인)", placeholder="example@domain.com", key="email_input", on_change=update_email)
    email_confirm_re = r'^[a-zA-Z0-9+-_.]+@[a-zA-Z0-9-]+\.[a-zA-Z0-9-.]+$'
    p = re.compile(email_confirm_re)
    if p.match(st.session_state["email"]) is None and email_addr:
        row2[1].warning("example@domain.com")

    input_list = [school_name, grade_num, student_ban, student_id, student_name, json_birth, email_addr]

    st.markdown("※ **학교명, 학년, 반, 번호, 이름**을 정확히 입력하세요. 오류 시 체험활동 인정이 되지 않습니다.")
    st.markdown("※ 본인의 **생년월일**을 정확하게 입력하세요.")
    st.markdown("※ 본관 1층에서 체험활동 **승인** 후 **E-mail**주소로 체험활동 확인서가 발송됩니다.")

    submit = st.button("저장하시겠습니까? ")
    if submit:
        del_message = f"""{input_list[0]} {input_list[1]}학년 {input_list[2]}반 {input_list[3]}번 {input_list[4]}학생의 기존에 기록된 정보를 수정하였습니다.  
         E-mail주소는 {input_list[6]}입니다. 대구과학고등학교 본관 1층 로비로 가서 체험활동확인서를 제출하고 승인 받으세요."""
        message = f"""{input_list[0]} {input_list[1]}학년 {input_list[2]}반 {input_list[3]}번 {input_list[4]}학생의 정보를 저장하였습니다.  
         E-mail주소는 {input_list[6]}입니다. 꼭 확인하세요. 틀렸을 경우 수정 후 저장하세요. 
         대구과학고등학교 본관 1층 로비로 가서 체험활동확인서를 제출하고 승인 받으세요."""
        if any(item is None or item == "" for item in input_list):
            st.warning("모든 필드를 채워주세요.")
        else:
            chehum_df = data_load(chehumsheet)

            detected_same_index = detect_same_index(chehum_df, input_list[:5])

            if detected_same_index:
                index_TF = chehum_df['승인'].values.tolist()[detected_same_index-2]
                if index_TF == 'TRUE':
                    st.warning("해당 학생의 정보가 이미 승인되었습니다. [승인확인] 탭에서 발급번호를 확인하세요.")
                else:
                    insert_row(chehumsheet, input_list, detected_same_index, del_message)

            else:
                last_row = chehumsheet.row_count + 1
                chehumsheet.add_rows(1)
                insert_row(chehumsheet, input_list, last_row, message)

if __name__ == "__main__":
    main()