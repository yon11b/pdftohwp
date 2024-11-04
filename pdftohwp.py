# pdf 파일에 있는 내용을 txt 파일에 입력하는 코드
# 완성
# 2024.10.25
# =====================[프로세스]========================
# 1. pdf => png
# 2. png의 왼쪽, 오른쪽 다단을 위 아래로 합쳐서 새 png 파일로저장
# 3. png의 텍스트를 네이버 clover ocr을 사용하여 텍스로 변환 후 txt 파일에 저장
# ====================================================
import requests
import uuid
import time
import json
from PIL import Image
import fitz
import os
import win32com.client as win32
import re
# 사용자로부터 키워드 입력받기
def input_file_name():
    keyword = input('파일 이름의 일부를 입력하세요: ')
    # 파일 리스트 가져오기
    files = os.listdir(PROBLEM_PATH)
    # 키워드를 포함한 파일 찾기
    matching_files = [file for file in files if keyword in file]
    return matching_files
PATH=''
PROBLEM_PATH=os.path.join(PATH, '문제파일')
PDF_FILE_PATH=''
FILE_NAME=''
# 폴더가 없으면 생성
os.makedirs(PROBLEM_PATH, exist_ok=True)

def find_file():
    global PDF_FILE_PATH
    global FILE_NAME
    matching_files = input_file_name()
    if matching_files:
        print("<찾은 파일>")
        if len(matching_files) == 1:
            file = matching_files[0]
            print(file)
            PDF_FILE_PATH=os.path.join(PROBLEM_PATH, file)
            FILE_NAME = file
        else:
            print()
            for idx, file in enumerate(matching_files):
                print(f'{idx+1}. {file}')
            fileIdx = int(input("변환하고 싶은 파일의 번호를 입력하세요: "))
            PDF_FILE_PATH =os.path.join(PROBLEM_PATH, matching_files[fileIdx - 1])
            FILE_NAME = matching_files[fileIdx - 1]
    else:
        print("파일을 찾을 수 없습니다.")
        find_file()
find_file()
page=0
# 1. pdf파일을 png 파일로 변환하여 저장
print('PDF 파일 읽는 중...')
doc = fitz.open(PDF_FILE_PATH)
input_dir = os.path.join(PATH, 'input_image')
# 폴더가 없으면 생성
os.makedirs(input_dir, exist_ok=True)
p = 0
for i, page in enumerate(doc):
    img = page.get_pixmap()
    img.save(os.path.join(input_dir, f'{i}_{FILE_NAME}.png'))
    p+=1
page=p
# 추출한 텍스트 저장할 txt 파일 생성
# text 폴더 경로
text_dir = os.path.join(PATH, 'text')
# 폴더가 없으면 생성
os.makedirs(text_dir, exist_ok=True)
# 텍스트 파일 경로
text_dir = os.path.join(text_dir, f'{FILE_NAME}.txt')
# UTF-8 인코딩으로 파일 열기 및 작성
f = open(text_dir, "w", encoding="utf-8")
output_dir = os.path.join(PATH, 'output_dir')
# 폴더가 없으면 생성
os.makedirs(output_dir, exist_ok=True)
# 2. png 파일을 다단 두 개로 나누어 저장
# 이미지 파일 열기
print('이미지를 텍스트로 변환  중...')

for i in range(page):
    image_path = os.path.join(input_dir, f'{i}_{FILE_NAME}.png') # 원본 이미지 파일 경로
    image = Image.open(image_path)
    # 이미지 크기 가져오기
    width, height = image.size
    # 이미지 절반으로 나누기
    left_image = image.crop((0, 0, width // 2, height))  # 왼쪽 절반
    right_image = image.crop((width // 2, 0, width, height))  # 오른쪽 절반
    # 새로운 이미지 생성 (왼쪽 이미지 높이 + 오른쪽 이미지 높이)
    new_image = Image.new('RGB', (width // 2, height * 2))
    # 왼쪽 이미지 붙여넣기
    new_image.paste(left_image, (0, 0))
    # 오른쪽 이미지 붙여넣기
    new_image.paste(right_image, (0, height))
    # 결과 이미지 저장
    new_image.save(os.path.join(output_dir, f'{i}_{FILE_NAME}.png'))  # 결과 파일 경로
    # 3. png 파일에서 텍스트 추출하기(네이버 클로버 ocr)
    api_url = ''
    secret_key = ''
    image_file = os.path.join(output_dir, f'{i}_{FILE_NAME}.png')
    request_json = {
        'images': [
            {
                'format': 'png',
                'name': 'demo'
            }
        ],
        'requestId': str(uuid.uuid4()),
        'version': 'V2',
        'timestamp': int(round(time.time() * 1000))
    }
    payload = {'message': json.dumps(request_json).encode('UTF-8')}
    files = [
      ('file', open(image_file,'rb'))
    ]
    headers = {
      'X-OCR-SECRET': secret_key
    }
    response = requests.request("POST", api_url, headers=headers, data = payload, files = files)
 
    print(f'{i+1} 페이지 완료...')
    # 4. 추출한 텍스트를 txt 파일에 저장
    for i in response.json()['images'][0]['fields']:
        text = i['inferText']
        f.write(text+' ')
f.close()
# 텍스트 파일에 있는 내용을 한글 파일에 입력하는 코드
# 완성
# 2024.10.25
# =====================[프로세스]========================
# 1. (txt) 텍스트 파일 불러오기
# 2. (hwp) 한글 파일에서 누름틀 생성하기
# 3. (hwp) 누름틀 사이 적절한 위치에 불러온 텍스트 삽입
# 4. 위 과정을 문제번호 처음부터 끝까지 반복
# ====================================================
# HWP 애플리케이션 시작
hwp = win32.gencache.EnsureDispatch('HWPFrame.HwpObject')
hwp.Run("FileNew")
problems = [[] for _ in range(40)]
numbers = [str(i) for i in range(1, 40)]
digits= [str(i) for i in range(0,10)]
def is십사(strNum):
    if strNum == '1' or strNum=='2':
        return True

TEXT_PATH=os.path.join(PATH, 'text')
# 폴더가 없으면 생성
os.makedirs(TEXT_PATH, exist_ok=True)
FILE_NAME = os.path.join(TEXT_PATH, FILE_NAME)
FILE_NAME = FILE_NAME+'.txt'
def write(s):
    try:
        # UTF-8 인코딩을 고려하여 텍스트 삽입
        act = hwp.CreateAction("InsertText")
        set = act.CreateSet()
        act.GetDefault(set)
        set.SetItem("Text", s.encode('utf-8').decode('utf-8'))  # UTF-8로 인코딩하여 삽입
        act.Execute(set)
    except Exception as e:
        print(f"Error inserting text: {e}")
def problem_insert(num):
    text_path = os.path.join(PATH, 'text')
    doc = fitz.open(FILE_NAME)
    contentNum = 0
    length=0
    num1=num
    num2=0
    if num >=10:
        num1=num//10
        num2=num%10
    for page in doc:
        text = page.get_text()
        content = re.sub(r'(\b\w+\b)(\s+\1)+', r'\1', text)
        length += len(content)
        for i in range(length):
            isAnswer = ''.join(content[i : i+6])
            if isAnswer=='정답과 해설':
                return -1
            contentNum = 0
            try:
                if content[i] == str(num1) and content[i+1]=='.':
                    contentNum = int(content[i])
                if content[i] == str(num1) and content[i+1] == str(num2) and content[i+2]=='.':
                    contentNum = int(content[i])*10 + int(content[i+1])
                    i+=1
            except:
                break
            if i > 0 and contentNum == num:
                if (num // 10==0 and is십사(content[i-1])):
                    continue
                j=i+1
                while True:
                    j+=1
                    try:
                        if content[j] == str(num+1) and content[j+1]=='.':       # 문제번호가 한 자리 수일 때
                            break
                        isTwoDigit = ''.join(content[j:j+2])
                        if  isTwoDigit== str(num+1) and content[j+2] == '.':
                        #if content[j] in digits and content[j+1] in digits and content[j+2] == '.':      # 문제번호가 두 자리 수일 때
                            break
                    except:
                         break
                problems[num-1].append(content[i+2:j])
                text=problems[num-1][0]
                # 정규 표현식으로 불필요한 문자들 제거
                cleaned_text = re.sub(r'\b([1-9]|[1-3][0-9]|40)\)\s*', '', text)    # 1부터 40까지의 숫자를 제거
                cleaned_text = cleaned_text.replace('zb', '')
                cleaned_text = cleaned_text.replace('<<<<보기>>>>', '<보기>')
                cleaned_text = cleaned_text.replace('????', '?')
                cleaned_text = cleaned_text.replace(', , , ,', ',')
                cleaned_text = cleaned_text.replace('. . . . ', '.\n')
                cleaned_text = cleaned_text.replace('다.', '다. ')
                cleaned_text = cleaned_text.replace('))))', ')')
                cleaned_text = cleaned_text.replace('((((', '(')
                return cleaned_text
    return -1
def insert_fields():
    hwp.RegisterModule("FilePathCheckDLL", "FilePathCheckerModule")
    start = int(input('시작 id를 입력해주세요: '))
    count = int(input('몇 개의 id를 생성할까요?: '))
    start_idx=0
    wheretostart = int(input('처음부터 시작하시겠습니까? 맞으면 1, 아니면 2를 입력해주세요: '))
    if wheretostart==2:
        # (커스텀) 문제 시작 위치:  중간
        num = int(input('문제 시작 번호를 입력하세요: '))
        num-=1
    else:
        # (커스텀) 문제 시작 위치:  처음(기본)
        num = 0
    for i in range(start, start + count):
        num +=1
        # 필드 템플릿 삽입 (ID)
        hwp.HAction.GetDefault("InsertFieldTemplate", hwp.HParameterSet.HInsertFieldTemplate.HSet)
        hwp.HParameterSet.HInsertFieldTemplate.TemplateDirection = "ID"
        hwp.HAction.Execute("InsertFieldTemplate", hwp.HParameterSet.HInsertFieldTemplate.HSet)
        hwp.HAction.Run("BreakPara")
        # 현재 커서 위치에 텍스트 삽입
        hwp.HAction.GetDefault("InsertText", hwp.HParameterSet.HInsertText.HSet)
        hwp.HParameterSet.HInsertText.Text = str(i)  # 삽입할 텍스트 설정
        hwp.HAction.Execute("InsertText", hwp.HParameterSet.HInsertText.HSet)
        hwp.HAction.Run("BreakPara")
        # 필드 템플릿 삽입 (문제)
        hwp.HAction.GetDefault("InsertFieldTemplate", hwp.HParameterSet.HInsertFieldTemplate.HSet)
        hwp.HParameterSet.HInsertFieldTemplate.TemplateDirection = "문제"
        hwp.HAction.Execute("InsertFieldTemplate", hwp.HParameterSet.HInsertFieldTemplate.HSet)
        hwp.HAction.Run("BreakPara")
        content = problem_insert(num)   # 문제 추출하기
        content = f"{num}번 문제는 naver ocr 기능의 정확도가 떨어져서 불러오지 못했습니다. 직접 작성해주세요!!" if content == -1 else content
        write(content)                    # 문제 입력하기
        hwp.HAction.Run("BreakPara")
        hwp.HAction.Run("BreakPara")
        # 필드 템플릿 삽입 (해설)
        hwp.HAction.GetDefault("InsertFieldTemplate", hwp.HParameterSet.HInsertFieldTemplate.HSet)
        hwp.HParameterSet.HInsertFieldTemplate.TemplateDirection = "해설"
        hwp.HAction.Execute("InsertFieldTemplate", hwp.HParameterSet.HInsertFieldTemplate.HSet)
        hwp.HAction.Run("BreakPara")
        hwp.HAction.Run("BreakPara")
        # 필드 템플릿 삽입 (정답)
        hwp.HAction.GetDefault("InsertFieldTemplate", hwp.HParameterSet.HInsertFieldTemplate.HSet)
        hwp.HParameterSet.HInsertFieldTemplate.TemplateDirection = "정답"
        hwp.HAction.Execute("InsertFieldTemplate", hwp.HParameterSet.HInsertFieldTemplate.HSet)
        hwp.HAction.Run("BreakPara")
        hwp.HAction.Run("BreakPara")
        # 페이지 나누기
        hwp.HAction.Run("BreakPage")
# 함수 실행
insert_fields()
hwp.SaveAs(f'{FILE_NAME}.hwp')
hwp.Quit()
