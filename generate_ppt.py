import os
import io
import json
import copy
from datetime import datetime
from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload, MediaIoBaseUpload
from pptx import Presentation

# -------------------------------------------------------------------------
# 1. Google Drive API 설정
# -------------------------------------------------------------------------
SCOPES = ['https://www.googleapis.com/auth/drive']

def init_drive_service():
    sa_info_str = os.environ.get("GCP_SA_KEY")
    if not sa_info_str:
        raise ValueError("GCP_SA_KEY 환경 변수가 설정되지 않았습니다.")
    sa_info = json.loads(sa_info_str)
    creds = service_account.Credentials.from_service_account_info(sa_info, scopes=SCOPES)
    return build('drive', 'v3', credentials=creds)

def search_file(service, folder_id, query_condition):
    q = f"'{folder_id}' in parents and {query_condition} and trashed = false"
    res = service.files().list(q=q, fields="files(id, name)").execute()
    files = res.get('files', [])
    return files[0] if files else None

def download_bytes(service, file_id):
    req = service.files().get_media(fileId=file_id)
    fh = io.BytesIO()
    downloader = MediaIoBaseDownload(fh, req)
    done = False
    while not done:
        _, done = downloader.next_chunk()
    fh.seek(0)
    return fh

def upload_final_file(service, folder_id, filename, file_stream):
    media = MediaIoBaseUpload(
        file_stream,
        mimetype='application/vnd.openxmlformats-officedocument.presentationml.presentation',
        resumable=True
    )
    body = {'name': filename, 'parents': [folder_id]}
    f = service.files().create(body=body, media_body=media, fields='id, name').execute()
    print(f"드라이브 업로드 완료: {f.get('name')}")

# -------------------------------------------------------------------------
# 2. PPT 슬라이드 복제 및 텍스트 치환
# -------------------------------------------------------------------------
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

def replace_text_in_prs(prs, old_text, new_text):
    if not new_text:
        return
    for slide in prs.slides:
        for shape in slide.shapes:
            if shape.has_text_frame:
                for paragraph in shape.text_frame.paragraphs:
                    for run in paragraph.runs:
                        if old_text in run.text:
                            run.text = run.text.replace(old_text, new_text)

# -------------------------------------------------------------------------
# 3. 메인 실행 흐름
# -------------------------------------------------------------------------
def main():
    service = init_drive_service()

    # 환경변수로부터 폴더 ID 읽기
    folder_template = os.environ.get("DRIVE_FOLDER_TEMPLATES")
    folder_hymn = os.environ.get("DRIVE_FOLDER_HYMNS")
    folder_praise = os.environ.get("DRIVE_FOLDER_PRAISES")
    folder_output = os.environ.get("DRIVE_FOLDER_OUTPUT")

    # 페이로드 추출 (repository_dispatch 또는 workflow_dispatch)
    raw_payload = os.environ.get("CLIENT_PAYLOAD") or os.environ.get("INPUT_PAYLOAD")
    if not raw_payload or raw_payload == '{}':
        raise ValueError("전송된 예배 데이터(Payload)가 비어 있습니다.")
    data = json.loads(raw_payload)

    worship_type = data.get("worship_type", "금요예배")
    date_str = data.get("date", datetime.today().strftime("%Y-%m-%d"))
    preacher = data.get("preacher", "").strip()
    scripture = data.get("scripture", "").strip()
    sermon_title = data.get("title", "").strip()
    praises = data.get("praises", [])
    hymns = data.get("hymns", [])

    # 요일 약칭 및 날짜 포맷 (YYMMDD_요일_설교자)
    dt = datetime.strptime(date_str, "%Y-%m-%d")
    days_kor = ["월", "화", "수", "목", "금", "토", "일"]
    day_kor = days_kor[dt.weekday()]
    yymmdd = dt.strftime("%y%m%d")
    output_filename = f"{yymmdd}_{day_kor}_{preacher}.pptx"

    # 1. 템플릿 탐색 (수, 금, 주일 매칭)
    template_kw = "_금_" if "금" in worship_type else ("_수_" if "수" in worship_type else "주일")
    template_meta = search_file(service, folder_template, f"name contains '{template_kw}'")
    if not template_meta:
        raise FileNotFoundError(f"'{template_kw}' 키워드를 포함한 템플릿 파일을 찾을 수 없습니다.")

    prs_stream = download_bytes(service, template_meta['id'])
    target_prs = Presentation(prs_stream)

    # 2. 텍스트 치환 (본문, 설교제목)
    replace_text_in_prs(target_prs, "{성경본문}", scripture)
    if sermon_title:
        replace_text_in_prs(target_prs, "{설교제목}", sermon_title)

    # 3. 찬양 슬라이드 병합 (새찬양집: 3자리 숫자 + '. ')
    for p_num in praises:
        if not str(p_num).strip():
            continue
        p_clean = str(int(p_num)).zfill(3) + ". "
        p_file = search_file(service, folder_praise, f"name starts with '{p_clean}'")
        if p_file:
            stream = download_bytes(service, p_file['id'])
            append_presentation(target_prs, stream)

    # 4. 찬송가 슬라이드 병합 (새찬송가: 3자리 숫자 + '장.pptx')
    for h_num in hymns:
        if not str(h_num).strip():
            continue
        h_clean = f"{str(int(h_num)).zfill(3)}장.pptx"
        h_file = search_file(service, folder_hymn, f"name = '{h_clean}'")
        if h_file:
            stream = download_bytes(service, h_file['id'])
            append_presentation(target_prs, stream)

    # 5. 완성본 파일 저장 및 업로드
    out_stream = io.BytesIO()
    target_prs.save(out_stream)
    out_stream.seek(0)
    upload_final_file(service, folder_output, output_filename, out_stream)

if __name__ == "__main__":
    main()
