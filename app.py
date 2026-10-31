import io
import copy
from datetime import datetime
import streamlit as st
from pptx import Presentation
from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload, MediaIoBaseUpload

# 페이지 기본 설정
st.set_page_config(page_title="예배 PPT 자동 생성기", page_icon="⛪", layout="centered")

# -----------------------------------------------------------------------------
# 1. 보안 비밀번호 게이트 (1907)
# -----------------------------------------------------------------------------
if "authenticated" not in st.session_state:
    st.session_state.authenticated = False

if not st.session_state.authenticated:
    st.markdown("### 🔒 예배사역 인증")
    st.write("접속 비밀번호를 입력해 주세요.")
    pw_input = st.text_input("비밀번호", type="password", key="auth_pw")
    if st.button("입장하기", use_container_width=True):
        if pw_input == "1907":
            st.session_state.authenticated = True
            st.rerun()
        else:
            st.error("비밀번호가 일치하지 않습니다.")
    st.stop()

# -----------------------------------------------------------------------------
# 2. Google Drive 연동 함수
# -----------------------------------------------------------------------------
SCOPES = ['https://www.googleapis.com/auth/drive']

@st.cache_resource
def get_drive_service():
    sa_info = dict(st.secrets["gcp_service_account"])
    creds = service_account.Credentials.from_service_account_info(sa_info, scopes=SCOPES)
    return build('drive', 'v3', credentials=creds)

def search_drive_file(service, folder_id, query_condition):
    q = f"'{folder_id}' in parents and {query_condition} and trashed = false"
    res = service.files().list(q=q, fields="files(id, name)").execute()
    files = res.get('files', [])
    return files[0] if files else None

def download_file_stream(service, file_id):
    req = service.files().get_media(fileId=file_id)
    fh = io.BytesIO()
    downloader = MediaIoBaseDownload(fh, req)
    done = False
    while not done:
        _, done = downloader.next_chunk()
    fh.seek(0)
    return fh

def upload_final_ppt(service, folder_id, filename, file_stream):
    media = MediaIoBaseUpload(
        file_stream,
        mimetype='application/vnd.openxmlformats-officedocument.presentationml.presentation',
        resumable=True
    )
    meta = {'name': filename, 'parents': [folder_id]}
    file_obj = service.files().create(body=meta, media_body=media, fields='id, webViewLink').execute()
    return file_obj.get('webViewLink')

# -----------------------------------------------------------------------------
# 3. PPTX 병합 및 서식 유지 텍스트 치환
# -----------------------------------------------------------------------------
def replace_shape_text(shape, old_text, new_text):
    if not shape.has_text_frame or not new_text:
        return
    for p in shape.text_frame.paragraphs:
        for run in p.runs:
            if old_text in run.text:
                run.text = run.text.replace(old_text, new_text)

def duplicate_slide(target_prs, source_slide):
    blank_layout = target_prs.slide_layouts[6]
    new_slide = target_prs.slides.add_slide(blank_layout)
    for shape in source_slide.shapes:
        new_shape = copy.deepcopy(shape.element)
        new_slide.shapes._spTree.append(new_shape)
    return new_slide

def append_presentation(target_prs, source_stream):
    src_prs = Presentation(source_stream)
    for slide in src_prs.slides:
        duplicate_slide(target_prs, slide)

# -----------------------------------------------------------------------------
# 4. 메인 입력 UI
# -----------------------------------------------------------------------------
st.title("⛪ 예배 PPT 자동 생성기")
st.caption("찬양과 본문 정보를 입력하면 템플릿 서식을 유지하며 완성본을 만듭니다.")

# 세션 상태 초기화 (동적 곡 관리)
if "praises" not in st.session_state:
    st.session_state.praises = ["", ""]
if "hymns" not in st.session_state:
    st.session_state.hymns = ["", ""]

# 1. 예배 기본 정보
st.subheader("1. 예배 기본 정보")
worship_type = st.selectbox("예배 구분", ["금요예배", "수요예배", "주일예배"])
col_date, col_preacher = st.columns(2)
with col_date:
    worship_date = st.date_input("예배 일자", datetime.today())
with col_preacher:
    preacher = st.text_input("설교자명", placeholder="예: 배정귀").strip()

# 2. 찬양 목록 (새찬양집)
st.subheader("2. 경배 찬양 목록 (새찬양집)")
for i, _ in enumerate(st.session_state.praises):
    c1, c2 = st.columns([5, 1])
    with c1:
        st.session_state.praises[i] = st.text_input(f"찬양 {i+1}번 곡", value=st.session_state.praises[i], key=f"praise_{i}", placeholder="예: 335")
    with c2:
        st.write("")
        st.write("")
        if st.button("삭제", key=f"del_praise_{i}"):
            st.session_state.praises.pop(i)
            st.rerun()

if st.button("➕ 찬양 추가", use_container_width=True):
    st.session_state.praises.append("")
    st.rerun()

# 3. 찬송가 목록 (새찬송가)
st.subheader("3. 찬송가 목록 (새찬송가)")
for j, _ in enumerate(st.session_state.hymns):
    c1, c2 = st.columns([5, 1])
    with c1:
        st.session_state.hymns[j] = st.text_input(f"찬송가 {j+1}번 장", value=st.session_state.hymns[j], key=f"hymn_{j}", placeholder="예: 338")
    with c2:
        st.write("")
        st.write("")
        if st.button("삭제", key=f"del_hymn_{j}"):
            st.session_state.hymns.pop(j)
            st.rerun()

if st.button("➕ 찬송가 추가", use_container_width=True):
    st.session_state.hymns.append("")
    st.rerun()

# 찬송가 목록 바로 다음에 위치하는 주일 교독문 섹션
reading_num = ""
if worship_type == "주일예배":
    st.markdown("---")
    st.subheader("📖 주일 교독문")
    reading_num = st.text_input("교독문 번호", placeholder="예: 28 (시편 63편)").strip()

# 4. 성경 본문 및 설교
st.subheader("4. 성경 본문 및 설교")
scripture = st.text_input("성경 본문 (필수)", placeholder="예: 고린도전서 12장 27절~31절").strip()
sermon_title = st.text_input("설교 제목 (선택 사항)", placeholder="미입력 시 템플릿 기본값을 유지합니다").strip()

st.markdown("---")

# -----------------------------------------------------------------------------
# 5. 생성 실행 버튼
# -----------------------------------------------------------------------------
if st.button("🚀 예배 PPT 생성 및 드라이브 저장", type="primary", use_container_width=True):
    if not preacher:
        st.warning("⚠️ 설교자명을 입력해 주세요.")
        st.stop()
    if not scripture:
        st.warning("⚠️ 성경 본문을 입력해 주세요.")
        st.stop()

    with st.spinner("구글 드라이브에서 템플릿과 찬양 파일을 불러와 병합 중입니다..."):
        try:
            drive = get_drive_service()
            
            # 폴더 ID 로드
            f_template = st.secrets["drive_folders"]["template_folder_id"]
            f_hymn = st.secrets["drive_folders"]["hymn_folder_id"]
            f_praise = st.secrets["drive_folders"]["praise_folder_id"]
            f_output = st.secrets["drive_folders"]["output_folder_id"]

            # 파일명 조합 (YYMMDD_요일_설교자.pptx)
            days_kor = ["월", "화", "수", "목", "금", "토", "일"]
            day_kor = days_kor[worship_date.weekday()]
            yymmdd = worship_date.strftime("%y%m%d")
            output_filename = f"{yymmdd}_{day_kor}_{preacher}.pptx"

            # 1. 템플릿 파일 탐색 및 로드
            template_kw = "_금_" if "금" in worship_type else ("_수_" if "수" in worship_type else "주일")
            t_meta = search_drive_file(drive, f_template, f"name contains '{template_kw}'")
            if not t_meta:
                st.error(f"'{template_kw}' 키워드가 포함된 템플릿 PPTX를 찾을 수 없습니다.")
                st.stop()

            prs_stream = download_file_stream(drive, t_meta['id'])
            target_prs = Presentation(prs_stream)

            # 2. 텍스트 치환
            for slide in target_prs.slides:
                for shape in slide.shapes:
                    replace_shape_text(shape, "{성경본문}", scripture)
                    if sermon_title:
                        replace_shape_text(shape, "{설교제목}", sermon_title)
                    if reading_num:
                        replace_shape_text(shape, "{교독문}", f"교독문 {reading_num}번")

            # 3. 찬양 파일 탐색 및 병합 (새찬양집: 3자리 숫자 + '. ')
            for p_val in st.session_state.praises:
                if p_val.strip().isdigit():
                    prefix = f"{int(p_val.strip()):03d}. "
                    p_meta = search_drive_file(drive, f_praise, f"name starts with '{prefix}'")
                    if p_meta:
                        st.write(f"✓ 찬양 추가됨: `{p_meta['name']}`")
                        p_stream = download_file_stream(drive, p_meta['id'])
                        append_presentation(target_prs, p_stream)
                    else:
                        st.warning(f"⚠️ 찬양 {p_val}번 파일을 찾지 못해 건너뜁니다.")

            # 4. 찬송가 파일 탐색 및 병합 (새찬송가: 3자리 숫자 + '장.pptx')
            for h_val in st.session_state.hymns:
                if h_val.strip().isdigit():
                    exact_name = f"{int(h_val.strip()):03d}장.pptx"
                    h_meta = search_drive_file(drive, f_hymn, f"name = '{exact_name}'")
                    if h_meta:
                        st.write(f"✓ 찬송가 추가됨: `{h_meta['name']}`")
                        h_stream = download_file_stream(drive, h_meta['id'])
                        append_presentation(target_prs, h_stream)
                    else:
                        st.warning(f"⚠️ 찬송가 {h_val}장 파일을 찾지 못해 건너뜁니다.")

            # 5. 완성본 메모리 저장 및 드라이브 업로드
            out_stream = io.BytesIO()
            target_prs.save(out_stream)
            out_stream.seek(0)

            drive_link = upload_final_ppt(drive, f_output, output_filename, out_stream)

            st.success(f"🎉 '{output_filename}' 생성이 완료되었습니다!")
            st.markdown(f"👉 [**구글 드라이브 완성본 폴더에서 열기**]({drive_link})")

            # 브라우저 직접 다운로드 버튼
            out_stream.seek(0)
            st.download_button(
                label="📥 내 PC로 바로 다운로드",
                data=out_stream,
                file_name=output_filename,
                mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                use_container_width=True
            )

        except Exception as e:
            st.error(f"생성 중 오류가 발생했습니다: {str(e)}")
