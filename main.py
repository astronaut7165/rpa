from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support.ui import WebDriverWait, Select
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import StaleElementReferenceException
from webdriver_manager.chrome import ChromeDriverManager
from datetime import datetime
import pandas as pd
import time, os, re, json, sys, openpyxl


# ──────────────────────────────────────────────────────────────
# 기본 변수 및 세팅
# ──────────────────────────────────────────────────────────────
LOGIN_GW_URL = "https://office.ms-global.com/login"
LOGIN_HRMS_URL = "https://hrms.ms-global.com/login.htm"

#휴일근무, 시간외근무 패턴 로딩
try:
    with open("patterns.json", "r", encoding="utf-8") as f:
        pattern_data = json.load(f)

    holidaywork_patterns = pattern_data["holidaywork_patterns"]
    overtime_patterns = pattern_data["overtime_patterns"]
except Exception as e:
    print(f"❌ patterns.json 로딩 실패: {e}")
    sys.exit(1)

# ──────────────────────────────────────────────────────────────
# 공통 유틸
# ──────────────────────────────────────────────────────────────

#문자->시간 변환함수
def str_to_time(tstr):
    try:
        return datetime.strptime(tstr.strip(), "%H:%M").time()
    except:
        return None

# ✅ duration key <-> (엑셀컬럼명, HRMS savename) 매핑
DURATION_MAPPING = {
    "work_time":              ("평일정취",         "work_time"),
    "overtime":               ("평일연장",         "overtime"),
    "night_time":             ("평일심야연장",     "night_time"),
    "work_extra_time":        ("특근정취",         "work_extra_time"),
    "minuit_over_time":       ("특근심야정취",     "minuit_over_time"),
    "holiday_over_time":      ("특근연장",         "holiday_over_time"),
    "extra_minuit_over_time": ("특근심야연장",     "extra_minuit_over_time"),
    "work_support":           ("유급휴가",             "work_support"),
    "late_time":              ("지각",             "late_time")
}

# ──────────────────────────────────────────────────────────────
# 로그 설정
# ──────────────────────────────────────────────────────────────
class DualLogger:
    def __init__(self, file_path):
        self.terminal = sys.__stdout__  # 원래 cmd 출력
        self.log = open(file_path, 'w', encoding='utf-8')

    def write(self, message):
        self.terminal.write(message)
        self.log.write(message)

    def flush(self):
        pass  # for compatibility


# 로그 경로 설정
log_file = f"작업확인서_자동처리로그_{datetime.now().strftime('%Y%m%d_%H%M%S')}.txt"
sys.stdout = sys.stderr = DualLogger(log_file)

#엑셀에 로그 추가
def append_log_to_excel(log_path, excel_path, sheet_name="로그기록"):
    # 엑셀 파일 열기 (없으면 새로 만듬)
    if os.path.exists(excel_path):
        wb = openpyxl.load_workbook(excel_path)
    else:
        wb = openpyxl.Workbook()
        wb.remove(wb.active)

    # 로그파일 읽기
    with open(log_path, "r", encoding="utf-8") as f:
        log_lines = f.readlines()

    # 시트 추가
    if sheet_name in wb.sheetnames:
        del wb[sheet_name]  # 기존 로그 시트 제거
    ws = wb.create_sheet(title=sheet_name)

    # 로그 줄별로 기록
    for i, line in enumerate(log_lines, start=1):
        ws.cell(row=i, column=1, value=line.strip())

    # 저장
    wb.save(excel_path)
    print(f"📄 로그 시트 추가 완료: {excel_path} > [{sheet_name}]")

# ──────────────────────────────────────────────────────────────
# 로그인 정보 입력받기
# ──────────────────────────────────────────────────────────────
def get_login_info():
    global USERNAME,GW_PASSWORD, HRMS_PASSWORD
    
    while True:
        sys.__stdout__.write("사원번호를 입력하세요: ")
        sys.__stdout__.flush()
        USERNAME = input().strip()
        
        sys.__stdout__.write("그룹웨어 비밀번호를 입력하세요: ")
        sys.__stdout__.flush()
        GW_PASSWORD = input().strip()
        time.sleep(0.1)
        sys.__stdout__.write("통합인사시스템 비밀번호를 입력하세요: ")
        sys.__stdout__.flush()
        HRMS_PASSWORD = input().strip()
        

        
        print("\n📝 입력한 정보 확인")
        print(f"사원번호               : {USERNAME}")
        print(f"그룹웨어 비밀번호      : {GW_PASSWORD}")
        print(f"통합인사시스템 비밀번호: {HRMS_PASSWORD}")

        while True :
            sys.__stdout__.write("✅ 맞으면 1, 틀리면 0을 입력하세요: ")
            sys.__stdout__.flush()
            confirm = input().strip()

            if confirm == "1":
                print("🔐 로그인 정보가 저장되었습니다.\n")
                return
            elif confirm == "0":
                print("🔁 다시 입력해주세요.\n")
                break
                time.sleep(1)
            else:
                print("⚠️ 잘못된 입력입니다. 1 또는 0만 입력 가능합니다.")
                time.sleep(1)


# ──────────────────────────────────────────────────────────────
# 그룹웨어 자동화 관련 함수
# ──────────────────────────────────────────────────────────────
def login_groupware():
    driver.get(LOGIN_GW_URL)
    wait.until(EC.presence_of_element_located((By.ID, "username"))).send_keys(USERNAME)
    driver.find_element(By.ID, "password").send_keys(GW_PASSWORD)
    driver.find_element(By.ID, "login_submit").click()
    wait.until(lambda d: "dashboard" in d.current_url or "home" in d.current_url)
    print("✅ 그룹웨어 로그인 완료")

def go_to_received_documents():
    wait.until(EC.element_to_be_clickable((By.XPATH, "//a[@href='/app/approval']"))).click()
    wait.until(EC.element_to_be_clickable((By.XPATH, "//a[@data-navi='todoreception']"))).click()
    print("✅ '결재 수신 문서' 진입 완료")

def search_work_confirmation():
    dropdown = Select(wait.until(EC.presence_of_element_located((By.ID, "searchtype"))))
    dropdown.select_by_value("formName")
    search_box = driver.find_element(By.ID, "keyword")
    search_box.send_keys("작업 확인서")
    driver.find_element(By.CLASS_NAME, "btn_search2").click()
    time.sleep(3)
    print("✅ '작업확인서' 검색 완료")

def click_receipt_and_confirm():
    try:
        # "접수" 버튼 존재 확인
        receipt_btns = driver.find_elements(By.XPATH, "//span[text()='접수']")
        if not receipt_btns:
            # print("ℹ️ 접수 버튼 없음 (이미 접수되었거나 조건 미충족)")
            return

        # "접수" 클릭
        receipt_btn = WebDriverWait(driver, 5).until(
            EC.element_to_be_clickable((By.XPATH, "//span[text()='접수']"))
        )
        receipt_btn.click()
        # print("📥 접수 버튼 클릭됨")

        # "확인" 버튼이 등장하면 접수 성공으로 간주
        WebDriverWait(driver, 5).until(
            EC.presence_of_element_located((By.XPATH, "//span[text()='확인']"))
        )
        # print("✅ 확인 버튼 등장 → 접수 성공으로 간주")

        # "확인" 버튼 존재 확인 후 클릭
        confirm_elements = driver.find_elements(By.XPATH, "//span[text()='확인']")
        if confirm_elements:
            confirm_btn = WebDriverWait(driver, 5).until(
                EC.element_to_be_clickable((By.XPATH, "//span[text()='확인']"))
            )
            confirm_btn.click()
            # print("✅ 확인 버튼 클릭 시도")
            WebDriverWait(driver, 5).until_not(
                EC.presence_of_element_located((By.XPATH, "//span[text()='확인']"))
            )
            # print("✅ 확인 완료됨")
        else:
            # print("ℹ️ 확인 버튼 없음")
            pass

    except Exception as e:
        print(f"❌ 접수 또는 확인 단계 실패: {type(e).__name__} - {e}")

def click_back_to_list():
    for _ in range(5):
        try:
            # 목록 버튼 대기 → XPath만 미리 쓰고, 클릭 직전에 다시 조회
            WebDriverWait(driver, 5).until(
                EC.presence_of_element_located((By.XPATH, "//span[text()='목록']"))
            )

            # 반드시 새로 조회해서 클릭해야 Stale 방지됨
            list_btn = WebDriverWait(driver, 5).until(
                EC.element_to_be_clickable((By.XPATH, "//span[text()='목록']"))
            )
            list_btn.click()
            time.sleep(2)
            # 문서 목록 페이지 복귀 확인
            WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.XPATH, "//td[@class='subject']/a"))
            )
            print("✅ 목록 복귀 완료")
            return True

        except StaleElementReferenceException:
            print("⚠️ 목록 요소가 사라짐 → 재시도")
            time.sleep(1)

        except Exception as e:
            print(f"❌ 목록 복귀 실패: {type(e).__name__} - {e}")
            return print("시스템을 종료합니다.")

def safe_driver_back(max_retries=5, wait_seconds=2):
    """
    driver.back()을 안전하게 수행하며, 뒤로 가기 성공 시까지 재시도함.
    """
    for attempt in range(1, max_retries + 1):
        try:
            print(f"🔙 뒤로가기 시도 {attempt}회차...")
            driver.back()

            # 문서 목록 페이지의 제목 요소 대기
            WebDriverWait(driver, 5).until(
                EC.presence_of_element_located((By.XPATH, "//td[@class='subject']/a"))
            )
            print("✅ 뒤로가기 성공! 문서 목록 페이지로 복귀함.")
            return True
        except Exception as e:
            print(f"⚠️ 뒤로가기 실패 {attempt}회차 → 재시도 예정")
            time.sleep(wait_seconds)

    print("❌ 최대 재시도 횟수 초과 → 강제 이동 시도 필요")
    return False

def get_work_confirmation_documents():
    all_dataframes = []
    doc_numbers = []
    count = 0  # ✅ 문서 처리 카운터

    while True:
        try:
            # 문서 리스트 재조회 (항상 첫 번째 문서만 처리)
            document_elements = driver.find_elements(By.XPATH, "//td[@class='subject']/a")
            docnum_elements = driver.find_elements(By.XPATH, "//td[@class='doc_num']/span")

            if not document_elements:
                print("✅ 모든 문서 처리 완료")
                break

            doc_element = document_elements[0]
            doc_number = re.sub(r'[\\/*?:\[\]]', '_', docnum_elements[0].text.strip())[:31]
            doc_numbers.append(doc_number)

            count += 1
            print(f"\n📄 {count}번째 문서 처리 중: 문서번호 {doc_number}")

            # 문서 클릭
            driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", doc_element)
            time.sleep(1)
            driver.execute_script("arguments[0].click();", doc_element)
            wait.until(EC.presence_of_element_located((By.TAG_NAME, "table")))

            # 테이블 존재 여부 확인
            try:
                table = WebDriverWait(driver, 5).until(
                    EC.presence_of_element_located((By.XPATH, "//table[contains(., '작업신청서(확인서)신청결과')]"))
                )
            except:
                print("⚠️ 테이블 없음 → 접수만 진행")
                click_receipt_and_confirm()
                click_back_to_list()
                continue

            # 데이터 수집
            rows = table.find_elements(By.TAG_NAME, "tr")
            table_data = [[td.text.strip() for td in row.find_elements(By.TAG_NAME, "td")] for row in rows]
            df = pd.DataFrame(table_data)
            if len(df) > 1:
                df.columns = df.iloc[0]
                df = df[1:]
            all_dataframes.append(df)

            print(f"✅ 수집 완료: {doc_number}")
            click_receipt_and_confirm()                  
            time.sleep(1)

            # 마지막 문서 구별
            if len(document_elements) == 1:
                print(f"✅ 마지막 문서 처리 완료: {doc_number}")
                break

            click_back_to_list()
            time.sleep(1)
            document_elements = driver.find_elements(By.XPATH, "//td[@class='subject']/a")
            if not document_elements:
                print("✅ 모든 문서 처리 완료 (남은 문서 없음)")
                break

        except Exception as e:
            print(f"❌ 문서 처리 실패: {e}")
            driver.refresh(); time.sleep(3)
            continue

    return all_dataframes, doc_numbers


def save_all_to_excel(dataframes, sheet_names, filename="작업확인서_신청결과.xlsx"):
    with pd.ExcelWriter(filename, engine="xlsxwriter") as writer:
        for idx, df in enumerate(dataframes):
            df.to_excel(writer, sheet_name=sheet_names[idx][:31], index=False)
    print(f"✅ 모든 데이터 저장 완료: {filename}")

# ──────────────────────────────────────────────────────────────
# 엑셀 정리 함수
# ──────────────────────────────────────────────────────────────
def format_excel(input_path, output_path):
    if not os.path.exists(input_path):
        print(f"❌ 파일이 존재하지 않음: {input_path}")
        return

    xls = pd.ExcelFile(input_path)
    combined_df = pd.DataFrame()

    for sheet in xls.sheet_names:
        df_raw = xls.parse(sheet, header=None)
        result_rows = []
        row_idx = 3

        while row_idx + 1 < len(df_raw):
            if df_raw.iloc[row_idx].isnull().all():
                break

            upper = df_raw.iloc[row_idx]
            lower = df_raw.iloc[row_idx + 1]

            row_data = {
                "No": upper[0],
                "구분": upper[1],
                "소속": upper[2],
                "사번": lower[0],
                "성명": upper[3],
                "시작": upper[4],
                "종료": lower[1],
                "신청시간": upper[5],
                "출근": upper[6],
                "퇴근": lower[2],
                "근무일자": upper[7],
                "보상구분": upper[8],
                "작업내용": upper[9],
                "문서번호": sheet
            }

            result_rows.append(row_data)
            row_idx += 3

        df_sheet = pd.DataFrame(result_rows)

        # 통합 시 중복 칼럼 제거
        if combined_df.empty:
            combined_df = df_sheet
        else:
            combined_df = pd.concat([combined_df, df_sheet], ignore_index=True)

    # 🔍 중복 열 탐지 및 제거
    duplicated_cols = combined_df.columns[combined_df.columns.duplicated()].tolist()
    if duplicated_cols:
        print(f"⚠️ 중복된 열 제거됨: {duplicated_cols}")
    else:
        #print("✅ 중복 열 없음")
        pass

    combined_df = combined_df.loc[:, ~combined_df.columns.duplicated()]
    combined_df.to_excel(output_path, index=False)
    print(f"✅ 엑셀 정리 및 통합 완료: {output_path}")

    return combined_df

def is_special_pattern_exception(row, pattern_start, pattern_end):
    """
    특정 패턴의 시간외근무 또는 휴일근무가 '정상 출근'으로 인정되도록 예외 처리.

    조건:
    - 시간외근무: 시작 00:20, 종료 01:20/02:20, 출근이 15:40 이전
    - 휴일근무: 시작 00:20, 출근이 23:00 이후 (전날 출근 간주)
    """
    try:
        goto_time = str_to_time(row["출근"])
        getoff_time = str_to_time(row["퇴근"])
        if not goto_time or not getoff_time:
            return False

        # [1] 시간외근무: 야간 연장 예외
        if (
            row["구분"] == "시간외근무" and
            pattern_start == str_to_time("00:20") and
            pattern_end in [str_to_time("01:20"), str_to_time("02:20")] and
            goto_time <= str_to_time("15:40") and
            getoff_time >= pattern_end
        ):
            return True

        # [2] 휴일근무: 전날 출근 예외
        if (
            row["구분"] == "휴일근무" and
            pattern_start == str_to_time("00:20") and
            goto_time >= str_to_time("23:00") and
            getoff_time >= pattern_end
        ):
            return True

        return False
    except:
        return False

def precheck_and_save_attendance_possibility(excel_path: str, json_path: str, output_path: str):
    """
    '작업확인서_신청결과_정리자동.xlsx'의 구분(B열)에 따라
    holidaywork_patterns 또는 overtime_patterns와 비교하여
    '작업가능' 또는 '작업불가능'을 '작업여부' 컬럼(O열)에 삽입하고 저장함.
    """
    df = pd.read_excel(excel_path)

    for col in ["평일정취", "평일연장", "평일심야연장", "특근정취", "특근심야정취", "특근연장", "특근심야연장", "유급휴가", "지각"]:
        df[col] = ""

    with open(json_path, "r", encoding="utf-8") as f:
        pattern_data = json.load(f)

    holiday_patterns = pattern_data["holidaywork_patterns"]
    overtime_patterns = pattern_data["overtime_patterns"]

    def check_row(row):
        #예외설정(장태근, 김규환)
        if row["성명"] in["장태근", "김규환", "이법훈", "배종태", "천국식", "손성호"]:
            return "예외설정"

        # 출근/퇴근 필수 체크
        if pd.isna(row["출근"]) or pd.isna(row["퇴근"]) or str(row["출근"]).strip() == "" or str(row["퇴근"]).strip() == "":
            return "출/퇴근시간 공란"

        start_time = str_to_time(row["시작"])
        end_time = str_to_time(row["종료"])
        work_time = str(row["신청시간"]).strip().replace(":", "").replace(".", "").zfill(4)
        goto_time = str_to_time(row["출근"])
        getoff_time = str_to_time(row["퇴근"])

        # 근무유형에 따라 비교할 패턴 선택
        if row["구분"] == "휴일근무":
            patterns = holiday_patterns
        elif row["구분"] == "시간외근무":
            patterns = overtime_patterns
        else:
            return "휴일,시간외근무 외 패턴"

        matched = False
        failure_reasons = [] # 최종 실패 이유 수집

        # 패턴 비교
        for pattern in patterns:
            pattern_start = str_to_time(pattern["start"])
            pattern_end = str_to_time(pattern["end"])

            reasons = [] # 현재 패턴에 대한 실패 이유

            if start_time != pattern_start:
                reasons.append("시작시간 불일치")
            if end_time != pattern_end:
                reasons.append("종료시간 불일치")
            if work_time not in pattern["work_times"]:
                reasons.append("신청시간 불일치")
            if not (
                 (goto_time <= pattern_start and getoff_time >= pattern_end) or # 엑셀 '출근'이 패턴 '시작' 초과,  엑셀 '퇴근'이 패턴 '종료' 이상이거나
                is_special_pattern_exception(row, pattern_start, pattern_end) # 예외패턴이면
            ):
                reasons.append("지각,조퇴 기타사유")

            if not reasons:
                #모든 조건 통과!
                for key, value in pattern["duration"].items():
                    if key in DURATION_MAPPING:
                        excel_col, _ = DURATION_MAPPING[key]
                        df.at[row.name, excel_col] = value
                matched = True
                break
            else:
                failure_reasons.append(reasons) # 현재 패턴 실패이유 누적
                
        if matched:
            return "작업가능"
        else:
            if failure_reasons:
                shortest_reason = min(failure_reasons, key=lambda x: len(x))
                return f"패턴불일치({', '.join(shortest_reason)})"

    df["작업여부"] = df.apply(check_row, axis=1)
    df.to_excel(output_path, index=False)
    print(f"✅ 저장 완료: {output_path}")

    return df




# ──────────────────────────────────────────────────────────────
# HRMS 자동화 함수 (구현 예정 단계 포함)
# ──────────────────────────────────────────────────────────────
def login_hrms():
    print("🔄 [1] 통합인사시스템 로그인 중...")
    driver.get(LOGIN_HRMS_URL)
    wait.until(EC.presence_of_element_located((By.ID, "login_id"))).send_keys(USERNAME)
    driver.find_element(By.ID, "passwd").send_keys(HRMS_PASSWORD)
    login_button = driver.find_element(By.CLASS_NAME, "btn_login")
    driver.execute_script("arguments[0].click();", login_button)

    try:
        WebDriverWait(driver, 5).until(EC.alert_is_present())
        alert = driver.switch_to.alert
        print(f"❌ 로그인 실패: {alert.text}")
        alert.accept()
        return False
    except:
        pass

    wait.until(EC.presence_of_element_located((By.TAG_NAME, "body")))
    print("✅ 로그인 성공!")
    return True

def set_hrms_role_if_needed():
    print("🔄 [2] 로그인 설정 '업무담당자'로 변경 중...")
    # 업무담당자가 이미 선택되어 있는지 확인하고, 없으면 설정
    driver.execute_script("onLoginAuthority();")
    WebDriverWait(driver, 10).until(lambda d: len(d.window_handles) > 1)
    original_window = driver.current_window_handle
    new_window = driver.window_handles[-1]
    driver.switch_to.window(new_window)
    print("✅ 로그인 권한 설정 창으로 전환 완료!")

    WebDriverWait(driver, 10).until(EC.frame_to_be_available_and_switch_to_it((By.NAME, "bottomF")))
    print("✅ 'bottomF' 프레임 전환 완료!")
    WebDriverWait(driver, 10).until(lambda d: d.execute_script("return document.readyState") == "complete")
    time.sleep(5)

    row = wait.until(EC.presence_of_element_located((By.XPATH, "//td[contains(text(), '업무담당자')]/parent::tr")))
    print("✅ '업무담당자' 행 찾기 완료!")

    login_auth_td = row.find_element(By.XPATH, ".//td[contains(@class, 'GMBool')]")
    if "GMBool3" in login_auth_td.get_attribute("class"):
        print("✅ 이미 '업무담당자' 셀이 선택되어 있음. 창 닫고 원래 창으로 복귀.")
        driver.close()
        driver.switch_to.window(original_window)
        return

    driver.execute_script("arguments[0].click();", login_auth_td)
    ActionChains(driver).move_to_element(login_auth_td).click().perform()
    time.sleep(1)
    print("✅ 로그인권한 선택 클릭 및 선택 적용 완료!")

    save_button = wait.until(EC.element_to_be_clickable((By.XPATH, "//img[contains(@src, 'b_save.gif')]")))
    driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", save_button)
    driver.execute_script("arguments[0].click();", save_button)

    WebDriverWait(driver, 5).until(EC.alert_is_present()).accept()
    time.sleep(2)
    WebDriverWait(driver, 5).until(EC.alert_is_present()).accept()
    time.sleep(2)
    print("✅ 저장 완료!")

    driver.close()
    driver.switch_to.window(original_window)
    print("✅ 원래 창으로 복귀 완료!")

def go_to_attendance_management():
    print("🔄 [3] 근태관리화면 이동 중...")
    driver.execute_script("subMenu('HRM_ODM')")
    time.sleep(1)
    driver.execute_script("menuAction('/common/page/comm_menu_action.jsp?menu_id=279503&action_uri=/odm/page/odm_offdutyDay_01_f.jsp','02');")

    # ✅ frame 또는 iframe 중 하나가 뜰 때까지 대기
    try:
        WebDriverWait(driver, 10).until(
            lambda d: len(d.find_elements(By.TAG_NAME, "frame")) > 0 or
                      len(d.find_elements(By.TAG_NAME, "iframe")) > 0
        )
        print("✅ '일일근태관리' 메뉴로 이동 완료!")
    except Exception as e:
        print(f"❌ 프레임 로딩 실패: {e}")

def search_user_in_hrms(emp_no: str, base_date: str):
    try:
        # 📌 기준일자 숫자만 남기기
        numeric_date = re.sub(r"\D", "", base_date)
        if len(numeric_date) != 8:
            raise ValueError(f"잘못된 base_date 형식: {base_date}")

        # 📌 정산년월: 앞 6자리만
        jungsan_ym = numeric_date[:6]

        # ✅ 프레임 초기화
        driver.switch_to.default_content()

        # ✅ 1단계: 바깥 프레임 진입 (Iframe_myIBTab1_divIBTabItem_1_Content)
        WebDriverWait(driver, 10).until(
            EC.frame_to_be_available_and_switch_to_it((By.NAME, "Iframe_myIBTab1_divIBTabItem_1_Content"))
        )

        # ✅ 2단계: 내부 프레임 진입 (topF)
        WebDriverWait(driver, 10).until(
            EC.frame_to_be_available_and_switch_to_it((By.NAME, "topF"))
        )

        # 사번 입력 + onchange 트리거
        script = f"""
        var input = document.getElementById("per_no");
        input.value = "{emp_no}";
        input.dispatchEvent(new Event("change"));
        """
        driver.execute_script(script)

        # 기준일자: 점 없는 숫자 8자리로 강제 입력 + keyup 이벤트 발생
        numeric_date = base_date.replace(".", "")  # ex: "20250409"
        driver.execute_script("""
            var input = document.getElementById("base_date");
            input.value = arguments[0];
            input.dispatchEvent(new Event("keyup"));
        """, numeric_date)


        # 정산년월 입력
        month_input = driver.find_element(By.ID, "jungsan_ym")
        month_input.clear()
        month_input.send_keys(jungsan_ym)

        # 조회 버튼 클릭
        search_button = driver.find_element(By.NAME, "Search_button")
        driver.execute_script("arguments[0].click();", search_button)

        print(f"✅ 사용자 조회 성공: 사번={emp_no}, 기준일자={base_date}")
        return True

    except Exception as e:
        print(f"❌ 사용자 조회 실패: {e}")
        return False

def apply_attendance_type_code(code_name: str):
    try:
        # ✅ 프레임 초기화
        driver.switch_to.default_content()

        # ✅ 1단계: 바깥 프레임 진입 (Iframe_myIBTab1_divIBTabItem_1_Content)
        WebDriverWait(driver, 10).until(
            EC.frame_to_be_available_and_switch_to_it((By.NAME, "Iframe_myIBTab1_divIBTabItem_1_Content"))
        )

        # ✅ 2단계: 내부 프레임 진입 (topF)
        WebDriverWait(driver, 10).until(
            EC.frame_to_be_available_and_switch_to_it((By.NAME, "bottomF"))
        )

        # 코드 매핑표
        code_map = {
            "특근": "900",
            "철야": "800",
        }

        if code_name not in code_map:
            print(f"⚠️ 근태코드 '{code_name}'은(는) 지원되지 않음. 반영 생략.")
            return

        code = code_map[code_name]
        driver.execute_script(f"mySheet.SetCellValue(3, 'odm_cd', '{code}');")
        print(f"✅ 근태코드 '{code_name}' ({code}) 반영 완료!")

    except Exception as e:
        print(f"❌ 근태코드 반영 실패: {e}")

def apply_attendance_hours(row: pd.Series):
    # 프레임 진입
    try:
        driver.switch_to.default_content()
        WebDriverWait(driver, 10).until(
            EC.frame_to_be_available_and_switch_to_it((By.NAME, "Iframe_myIBTab1_divIBTabItem_1_Content"))
        )
        WebDriverWait(driver, 10).until(
            EC.frame_to_be_available_and_switch_to_it((By.NAME, "bottomF"))
        )
    except Exception as e:
        print(f"❌ 프레임 진입 실패: {e}")
        return

    # IBSheet 로드 대기
    try:
        WebDriverWait(driver, 10).until(
            lambda d: d.execute_script("return typeof mySheet !== 'undefined'")
        )
    except Exception as e:
        print(f"❌ IBSheet 로드 실패: {e}")
        return

    # 헤더 → HRMS 필드명 매핑
    column_to_savename = {
        "특근정취": "work_extra_time",
        "특근심야": "minuit_over_time",
        "특근심야연장": "extra_minuit_over_time",
        "평일연장": "overtime",
        "평일심야연장": "night_time"
    }

    if row["작업여부"] != "작업가능":
        print(f"⏭️ 작업불가: {row['성명']} / {row['사번']}")
        return

    for key, (col_name, savename) in DURATION_MAPPING.items():
        value = row.get(col_name)
        if pd.notna(value) and str(value).strip() != "":
            try:
                driver.execute_script(f"mySheet.SetCellValue(3, '{savename}', {value});")
                print(f"✅ 반영됨 → {col_name} = {value}")
            except Exception as e:
                print(f"❌ 반영 실패: {col_name} → {e}")

def save_attendance(df: pd.DataFrame, idx: int):
    try:
        # 프레임 진입 (이미 이전에 들어가 있었으면 생략 가능하지만, 안전하게 다시 진입)
        driver.switch_to.default_content()
        WebDriverWait(driver, 10).until(
            EC.frame_to_be_available_and_switch_to_it((By.NAME, "Iframe_myIBTab1_divIBTabItem_1_Content"))
        )
        WebDriverWait(driver, 10).until(
            EC.frame_to_be_available_and_switch_to_it((By.NAME, "bottomF"))
        )

        # 저장 버튼 클릭
        save_btn = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.ID, "btn_save"))
        )
        driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", save_btn)
        save_btn.click()

        # 저장 확인 알림 처리
        WebDriverWait(driver, 5).until(EC.alert_is_present())
        alert = driver.switch_to.alert
        #print(f"💬 1차 알림: {alert.text}")
        alert.accept()
        time.sleep(1)

        # 저장 완료 알림 처리
        WebDriverWait(driver, 5).until(EC.alert_is_present())
        alert = driver.switch_to.alert
        #print(f"💬 2차 알림: {alert.text}")
        alert.accept()
        time.sleep(1)

        df.at[idx, "완료여부"] = "성공"
        print("✅ 저장 완료")
        

    except Exception as e:
        df.at[idx, "완료여부"] = "실패"
        print(f"❌ 저장 실패: {e}")

# ──────────────────────────────────────────────────────────────
# 메인 실행 흐름
# ──────────────────────────────────────────────────────────────

#로그인 정보 입력
get_login_info()

service = Service(ChromeDriverManager().install())
driver = webdriver.Chrome(service=service)
wait = WebDriverWait(driver, 10)

login_groupware()
go_to_received_documents()
search_work_confirmation()
dataframes, docnames = get_work_confirmation_documents()
save_all_to_excel(dataframes, docnames)
format_excel("작업확인서_신청결과.xlsx", "작업확인서_신청결과_정리자동.xlsx")
precheck_and_save_attendance_possibility(
    "작업확인서_신청결과_정리자동.xlsx",
    "patterns.json",
    "작업확인서_선별결과.xlsx"
)
login_hrms()
set_hrms_role_if_needed()
go_to_attendance_management()

# 파일 읽기
file_path = "작업확인서_선별결과.xlsx"
df = pd.read_excel(file_path)
df["완료여부"] = ""

for idx, row in df.iterrows():
    # 2. 작업가능 여부 확인
    if row.get("작업여부") != "작업가능":
        print(f"⏭️ 건너뜀: {row.get('성명')} / {row.get('사번')} → 작업불가능")
        continue

    emp_no = str(row.get("사번")).strip()
    work_date = str(row.get("근무일자")).strip()

    if not emp_no or not work_date:
        print(f"⚠️ 사번 또는 근무일자 누락 → 행 {idx+2} 건너뜀")
        continue

    print(f"▶️ 처리 중: {row.get('성명')} ({emp_no}) / {work_date}")

    # 3. 사용자 조회
    if not search_user_in_hrms(emp_no, work_date):
        continue

    # 4-1. 근태코드 반영 (휴일근무인 경우만)
    if row.get("구분") == "휴일근무":
        apply_attendance_type_code("특근")

    # 4-2. 근태코드 반영 (철야인 경우만)
    if row.get("시작") == "07:00" and row.get("종료") == "00:20" and row.get("구분") in ["시간외근무", "특근"]:
        apply_attendance_type_code("철야")
        
    # 5. 근무시간 반영
    apply_attendance_hours(row)

    # 6. 저장 버튼 클릭 / 완료여부 기재
    save_attendance(df, idx)

#이전 파일 삭제
# for f in ["작업확인서_신청결과.xlsx", "작업확인서_신청결과_정리자동.xlsx"]:
#     if os.path.exists(f):
#         os.remove(f)
    
#최종 파일
df.to_excel("작업확인서_처리결과.xlsx", index=False)
print("📝 완료결과 저장됨 → 작업확인서_처리결과.xlsx")

#로그 저장
append_log_to_excel(log_file, "작업확인서_처리결과.xlsx")

# 폴더 생성 및 파일 이동
today_folder = datetime.now().strftime("%Y%m%d_%H%M")
if not os.path.exists(today_folder):
    os.makedirs(today_folder)

result_files = [
    "작업확인서_처리결과.xlsx",
    "작업확인서_신청결과.xlsx",
    "작업확인서_신청결과_정리자동.xlsx",
    "작업확인서_선별결과.xlsx",
    log_file  # 자동생성된 로그 파일
]

for file in result_files:
    if os.path.exists(file):
        shutil.move(file, os.path.join(today_folder, file))
        print(f"✅ {file} → {today_folder} 폴더로 이동 완료")

print(f"\n📁 모든 결과 파일이 '{today_folder}' 폴더로 정리되었습니다.")

input("🔹 종료하려면 Enter 키를 누르세요...")
driver.quit()
