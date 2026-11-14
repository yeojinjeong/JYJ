import os
import re
import tempfile
import streamlit as st
import openpyxl
import pdfplumber
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont

# 페이지 설정
st.set_page_config(
    page_title="수출신고필증 엑셀 변환기",
    page_icon="📄",
    layout="wide"
)

# ---------------------------------------------------------
# [함수] 테스트용 수출신고필증 PDF 생성
# ---------------------------------------------------------
def create_sample_export_pdf(filename):
    font_path = "C:/Windows/Fonts/gulim.ttc"
    if os.path.exists(font_path):
        pdfmetrics.registerFont(TTFont('Gulim', font_path))
        font_name = 'Gulim'
    else:
        font_name = 'Helvetica'

    c = canvas.Canvas(filename, pagesize=A4)
    width, height = A4

    c.setFont(font_name, 16)
    c.drawString(50, height - 50, "수 출 신 고 필 증 (EXPORTS DECLARATION)")

    c.setFont(font_name, 10)
    y = height - 90
    
    headers = [
        "수출신고번호: 123-26-000001A",
        "수출자: (주)한국무역상사 (사업자번호: 123-81-12345)",
        "제조자: (주)테크팩토리",
        "수하인: ABC GLOBAL INC. (USA)",
        "결제금액: USD 50,000.00",
        "인도조건: FOB",
        "총 중량: 1,200 KG"
    ]
    
    for text in headers:
        c.drawString(50, y, text)
        y -= 20

    y -= 20
    c.drawString(50, y, "[란별 품목 상세 내역]")
    y -= 25

    c.drawString(50, y, "란번호    HS부호          품명 및 규격                 수량(단위)      단가(USD)     금액(USD)")
    y -= 15
    c.line(50, y, 550, y)
    y -= 20

    items = [
        ("001", "8517.62-1000", "Smart Gateway v2.0", "500 EA", "40.00", "20,000.00"),
        ("002", "8504.40-2000", "Power Supply Unit 12V", "1,000 PCS", "15.00", "15,000.00"),
        ("003", "3926.90-9000", "Plastic Enclosure Case", "2,000 PCS", "7.50", "15,000.00"),
    ]

    for item in items:
        row_str = f"{item[0]:<8} {item[1]:<15} {item[2]:<28} {item[3]:<12} {item[4]:<12} {item[5]}"
        c.drawString(50, y, row_str)
        y -= 20

    c.save()

# ---------------------------------------------------------
# [함수] PDF 데이터 파싱 (순수 파이썬 데이터 구조 사용)
# ---------------------------------------------------------
def parse_export_pdf_to_data(pdf_path):
    header_data = {}
    item_rows = []

    with pdfplumber.open(pdf_path) as pdf:
        full_text = ""
        for page in pdf.pages:
            full_text += page.extract_text() + "\n"

    # 헤더 정보 정규식 추출
    m_no = re.search(r"수출신고번호[:\s]+([A-Za-z0-9-]+)", full_text)
    m_exporter = re.search(r"수출자[:\s]+([^\n]+)", full_text)
    m_maker = re.search(r"제조자[:\s]+([^\n]+)", full_text)
    m_consignee = re.search(r"수하인[:\s]+([^\n]+)", full_text)
    m_amount = re.search(r"결제금액[:\s]+([^\n]+)", full_text)

    header_data["수출신고번호"] = m_no.group(1) if m_no else ""
    header_data["수출자"] = m_exporter.group(1).strip() if m_exporter else ""
    header_data["제조자"] = m_maker.group(1).strip() if m_maker else ""
    header_data["수하인"] = m_consignee.group(1).strip() if m_consignee else ""
    header_data["결제금액"] = m_amount.group(1).strip() if m_amount else ""

    # 품목 상세 내역 파싱
    lines = full_text.split("\n")
    is_item_section = False
    
    for line in lines:
        if "[란별 품목 상세 내역]" in line:
            is_item_section = True
            continue
        
        if is_item_section:
            parts = line.strip().split()
            if parts and parts[0].isdigit() and len(parts[0]) == 3:
                item_rows.append({
                    "수출신고번호": header_data["수출신고번호"],
                    "수출자": header_data["수출자"],
                    "제조자": header_data["제조자"],
                    "수하인": header_data["수하인"],
                    "란번호": parts[0],
                    "HS부호": parts[1],
                    "품명 및 규격": " ".join(parts[2:-3]),
                    "수량/단위": f"{parts[-3]} {parts[-2]}",
                    "총금액(USD)": parts[-1]
                })

    return item_rows, header_data

# ---------------------------------------------------------
# Streamlit UI 구현
# ---------------------------------------------------------
st.title("📦 수출신고필증 데이터 추출기")
st.markdown("수출신고필증 PDF 파일을 업로드하면 핵심 데이터와 품목 상세 내역을 자동으로 추출하여 엑셀로 변환해 드립니다.")

st.sidebar.header("⚙️ 테스트 및 설정")
if st.sidebar.button("테스트용 샘플 PDF 생성하기"):
    sample_filename = "sample_export_cert.pdf"
    create_sample_export_pdf(sample_filename)
    st.sidebar.success("테스트 PDF 생성 완료! 업로드 창에서 활용하세요.")
    with open(sample_filename, "rb") as f:
        st.sidebar.download_button(
            label="샘플 PDF 다운로드",
            data=f,
            file_name=sample_filename,
            mime="application/pdf"
        )

# 파일 업로드 위젯
uploaded_file = st.file_uploader("수출신고필증 PDF 파일을 선택해 주세요.", type=["pdf"])

if uploaded_file is not None:
    with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as tmp_file:
        tmp_file.write(uploaded_file.getvalue())
        tmp_path = tmp_file.name

    st.info("파일 분석 중입니다...")
    
    try:
        item_rows, header_info = parse_export_pdf_to_data(tmp_path)
        
        st.success("데이터 추출이 완료되었습니다!")
        
        # 1. 상단 기본 정보 표시
        st.subheader("📌 주요 헤더 정보")
        cols = st.columns(3)
        cols[0].metric("수출신고번호", header_info.get("수출신고번호", "-"))
        cols[1].metric("결제금액", header_info.get("결제금액", "-"))
        cols[2].metric("수하인", header_info.get("수하인", "-"))

        # 2. 품목 상세 테이블 표시 (스트림릿 네이티브 표 렌더링)
        st.subheader("📋 란별 품목 상세 내역")
        if item_rows:
            # Streamlit이 인식할 수 있는 리스트 오브 딕셔너리 형태 출력
            st.table(item_rows)
        else:
            st.warning("추출된 품목 상세 내역이 없습니다. 기본 헤더 정보만 저장됩니다.")
            st.json(header_info)

        # 3. openpyxl을 이용한 엑셀 파일 생성 및 다운로드 제공
        output_excel_path = "export_extracted_result.xlsx"
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "수출신고내역"

        if item_rows:
            # 헤더 작성
            headers_keys = list(item_rows[0].keys())
            ws.append(headers_keys)
            # 데이터 작성
            for row in item_rows:
                ws.append(list(row.values()))
        else:
            ws.append(list(header_info.keys()))
            ws.append(list(header_info.values()))

        wb.save(output_excel_path)
        
        with open(output_excel_path, "rb") as f:
            excel_bytes = f.read()
            
        st.download_button(
            label="📥 변환된 엑셀 파일 다운로드 (.xlsx)",
            data=excel_bytes,
            file_name="export_extracted_result.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
        
    except Exception as e:
        st.error(f"데이터를 처리하는 중 오류가 발생했습니다: {e}")
        
    finally:
        if os.path.exists(tmp_path):
            os.remove(tmp_path)