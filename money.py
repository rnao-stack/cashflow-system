import streamlit as st
import pandas as pd
from datetime import datetime, timedelta
import re
import requests
import io
import json
import os

# ===========================================================
# ⚙️ [사장님 맞춤 기본값 설정 공간] ⚙️
# 대시보드가 잠들었다 깨어나도 절대 안 변하게 하려면
# 아래 숫자들을 원하시는 최신 값으로 직접 수정해 주세요!
# ===========================================================
DEFAULT_TARGET_CNY = 210.0       # 🎯 CNY 목표 환율 (예: 210원)
DEFAULT_TARGET_USD = 1500.0      # 🎯 USD 목표 환율 (예: 1450원)

DEFAULT_BASE_DATE = "2026-07-05" # 🏦 초기 잔고 기준 날짜
DEFAULT_BASE_CNY = 15102.45     # 🏦 초기 CNY 잔고
DEFAULT_BASE_USD = 1110.21     # 🏦 초기 USD 잔고
# ===========================================================

# -----------------------------------------------------------
# 1. 페이지 설정 및 메모장 세션 초기화
# -----------------------------------------------------------
st.set_page_config(page_title="자금 관리 대시보드 Pro", layout="wide")

if 'memo_df' not in st.session_state:
    st.session_state.memo_df = pd.DataFrame({
        "구분": ["총 물품대", "부족 물품대", "환전 필요액"],
        "상세": ["발주, 발주 예정 물품대 합계", "보유 물품대 제외 부족 금액", "보유 외화 제외 추가 환전 필요 금액"]
    })

def fmt_num(x):
    try:
        return f"{float(x):,.2f}"
    except:
        return x

def fmt_krw(x):
    try:
        return f"{float(x):,.0f}"
    except:
        return x

# -----------------------------------------------------------
# 2. 데이터 로딩 및 전처리 (초고속 캐싱)
# -----------------------------------------------------------
def clean_currency(x):
    if isinstance(x, str):
        clean_str = re.sub(r'[^\d.-]', '', x)
        try:
            return float(clean_str) if clean_str else 0.0
        except:
            return 0.0
    return float(x) if pd.notnull(x) else 0.0

def parse_excel_data(file_bytes):
    try:
        xls = pd.ExcelFile(file_bytes)
        
        if '다이렉트' in xls.sheet_names:
            df_d = pd.read_excel(xls, sheet_name='다이렉트')
            df_d.columns = df_d.columns.str.strip()
            if '잔금_금액' in df_d.columns:
                df_d['잔금_금액'] = df_d['잔금_금액'].apply(clean_currency).fillna(0)
            if '잔금_날짜' in df_d.columns:
                df_d['잔금_날짜'] = pd.to_datetime(df_d['잔금_날짜'], errors='coerce')
            if '실지급_날짜' in df_d.columns:
                df_d['실지급_날짜'] = pd.to_datetime(df_d['실지급_날짜'], errors='coerce')
            if '화폐단위' in df_d.columns:
                df_d['화폐단위'] = df_d['화폐단위'].astype(str).str.upper().str.strip()
            if '구분' not in df_d.columns:
                df_d['구분'] = 'Direct'
        else:
            df_d = pd.DataFrame()

        if 'YIWU' in xls.sheet_names:
            df_y = pd.read_excel(xls, sheet_name='YIWU')
            df_y.columns = df_y.columns.str.strip()
            if '잔금' in df_y.columns and '잔금_금액' not in df_y.columns:
                df_y.rename(columns={'잔금': '잔금_금액'}, inplace=True)
            if '잔금_금액' in df_y.columns:
                df_y['잔금_금액'] = df_y['잔금_금액'].apply(clean_currency).fillna(0)
                comm_col = next((c for c in df_y.columns if '수수료' in str(c)), None)
                if comm_col:
                    def apply_fee(row):
                        val = row['잔금_금액']
                        status = str(row[comm_col]).strip()
                        if '별도' in status:
                            return val * 1.1
                        return val
                    df_y['잔금_금액'] = df_y.apply(apply_fee, axis=1)
            if '잔금_날짜' in df_y.columns:
                df_y['잔금_날짜'] = pd.to_datetime(df_y['잔금_날짜'], errors='coerce')
        else:
            df_y = pd.DataFrame()

        yiwu_balance = 0.0
        df_l = pd.DataFrame()
        target_sheet = None
        for s_name in xls.sheet_names:
            clean_name = s_name.replace(" ", "").upper()
            if '송금' in clean_name and 'YIWU' in clean_name:
                target_sheet = s_name
                break
        
        if target_sheet:
            df_l = pd.read_excel(xls, sheet_name=target_sheet)
            if '잔고' not in str(list(df_l.columns)) and '잔고(CNY)' not in str(list(df_l.columns)):
                df_l = pd.read_excel(xls, sheet_name=target_sheet, header=1)
            df_l.columns = df_l.columns.astype(str).str.strip().str.replace('\n', '')
            bal_col = next((c for c in df_l.columns if '잔고' in c), None)
            if bal_col:
                raw_balances = df_l[bal_col]
                balances = raw_balances.apply(clean_currency)
                valid_mask = raw_balances.notna() & raw_balances.astype(str).str.strip().ne("")
                valid_balances = balances[valid_mask]
                if not valid_balances.empty:
                    yiwu_balance = valid_balances.iloc[-1] 
            if '날짜' in df_l.columns:
                df_l['날짜'] = pd.to_datetime(df_l['날짜'], errors='coerce')
                
        df_ex = pd.DataFrame()
        if '환전내역' in xls.sheet_names:
            df_ex = pd.read_excel(xls, sheet_name='환전내역')
            df_ex.columns = df_ex.columns.str.strip()
            date_col = next((c for c in df_ex.columns if '날짜' in str(c) or '일자' in str(c)), None)
            if date_col:
                df_ex['날짜'] = pd.to_datetime(df_ex[date_col], errors='coerce')
            curr_col = next((c for c in df_ex.columns if '화폐' in str(c) or '통화' in str(c) or '구분' in str(c)), None)
            if curr_col:
                df_ex.rename(columns={curr_col: '화폐'}, inplace=True)
            amt_col = next((c for c in df_ex.columns if '환전' in str(c) and '원화' not in str(c)), None)
            if not amt_col:
                amt_col = next((c for c in df_ex.columns if '외화' in str(c)), None)
            if amt_col:
                df_ex['환전금액'] = df_ex[amt_col].apply(clean_currency).fillna(0)
            
        return df_d, df_y, yiwu_balance, df_l, df_ex
    except Exception as e:
        st.error(f"데이터 로드 에러: {e}")
        return pd.DataFrame(), pd.DataFrame(), 0.0, pd.DataFrame(), pd.DataFrame()

@st.cache_data(ttl=600, show_spinner="☁️ 구글 드라이브에서 데이터를 불러오는 중입니다...")
def get_drive_data(url):
    file_id_match = re.search(r'/d/([a-zA-Z0-9_-]+)', url) or re.search(r'id=([a-zA-Z0-9_-]+)', url)
    if not file_id_match:
        return None
    file_id = file_id_match.group(1)
    download_url = f"https://drive.google.com/uc?id={file_id}&export=download"
    try:
        response = requests.get(download_url, timeout=20)
        if response.status_code == 200:
            return parse_excel_data(io.BytesIO(response.content))
    except Exception as e:
        st.error(f"구글 드라이브 연동 실패: {e}")
    return None

RATE_CACHE_FILE = "last_exchange_rates.json"

def save_last_rates(cny_rate, usd_rate, source):
    data = {
        "cny": cny_rate,
        "usd": usd_rate,
        "source": source,
        "saved_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    }

    try:
        with open(RATE_CACHE_FILE, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False)
    except:
        pass

def load_last_rates():
    if not os.path.exists(RATE_CACHE_FILE):
        return None, None, None, None

    try:
        with open(RATE_CACHE_FILE, "r", encoding="utf-8") as f:
            data = json.load(f)

        return (
            float(data.get("cny")),
            float(data.get("usd")),
            data.get("source"),
            data.get("saved_at")
        )
    except:
        return None, None, None, None

@st.cache_data(ttl=600, show_spinner="💱 실시간 환율 정보를 불러오는 중입니다...")
def get_live_exchange_rates():
    headers = {
        "User-Agent": "Mozilla/5.0",
        "Referer": "https://m.stock.naver.com/"
    }

    # 1차: 네이버 모바일 페이지
    try:
        def get_naver_mobile_rate(code):
            url = f"https://m.stock.naver.com/marketindex/exchange/{code}"
            res = requests.get(url, headers=headers, timeout=5)
            res.raise_for_status()

            match = re.search(r'([\d,]+\.\d+)\s*KRW', res.text)
            if match:
                return float(match.group(1).replace(",", ""))

            return None

        usd_rate = get_naver_mobile_rate("FX_USDKRW")
        cny_rate = get_naver_mobile_rate("FX_CNYKRW")

        if cny_rate and usd_rate:
            save_last_rates(cny_rate, usd_rate, "네이버 모바일")
            return cny_rate, usd_rate, "네이버 모바일", False

    except:
        pass

    # 2차: 네이버 JSON API
    try:
        def get_naver_json_rate(code):
            url = f"https://api.stock.naver.com/marketindex/exchange/{code}/prices?page=1&pageSize=1"
            res = requests.get(url, headers=headers, timeout=5)
            res.raise_for_status()
            data = res.json()

            if not data:
                return None

            return float(str(data[0]["closePrice"]).replace(",", ""))

        usd_rate = get_naver_json_rate("FX_USDKRW")
        cny_rate = get_naver_json_rate("FX_CNYKRW")

        if cny_rate and usd_rate:
            save_last_rates(cny_rate, usd_rate, "네이버 API")
            return cny_rate, usd_rate, "네이버 API", False

    except:
        pass

    # 3차: 네이버 기존 HTML
    try:
        url = "https://finance.naver.com/marketindex/exchangeList.naver"
        res = requests.get(url, headers=headers, timeout=5)
        res.encoding = "euc-kr"

        u_match = re.search(r'미국 USD.*?<td class="sale">([\d,.]+)</td>', res.text, re.DOTALL)
        c_match = re.search(r'중국 CNY.*?<td class="sale">([\d,.]+)</td>', res.text, re.DOTALL)

        usd_rate = float(u_match.group(1).replace(",", "")) if u_match else None
        cny_rate = float(c_match.group(1).replace(",", "")) if c_match else None

        if cny_rate and usd_rate:
            save_last_rates(cny_rate, usd_rate, "네이버 HTML")
            return cny_rate, usd_rate, "네이버 HTML", False

    except:
        pass

    # 4차: Investing.com
    try:
        usd_res = requests.get("https://kr.investing.com/currencies/usd-krw", headers=headers, timeout=5)
        usd_match = re.search(r'data-test="instrument-price-last"[^>]*>([\d,.]+)<', usd_res.text)
        usd_rate = float(usd_match.group(1).replace(",", "")) if usd_match else None

        cny_res = requests.get("https://kr.investing.com/currencies/cny-krw", headers=headers, timeout=5)
        cny_match = re.search(r'data-test="instrument-price-last"[^>]*>([\d,.]+)<', cny_res.text)
        cny_rate = float(cny_match.group(1).replace(",", "")) if cny_match else None

        if cny_rate and usd_rate:
            save_last_rates(cny_rate, usd_rate, "Investing.com")
            return cny_rate, usd_rate, "Investing.com", False

    except:
        pass

    # 5차: 마지막 정상 환율
    last_cny, last_usd, last_source, saved_at = load_last_rates()
    if last_cny and last_usd:
        return last_cny, last_usd, f"마지막 정상 환율 ({last_source}, {saved_at})", True

    return None, None, "조회 실패", True

def calculate_realtime_balances(df_d, df_ex, df_l, base_date, base_cny, base_usd):
    cny_bal = base_cny
    usd_bal = base_usd

    if not df_ex.empty and '날짜' in df_ex.columns:
        new_ex = df_ex[df_ex['날짜'] > base_date]
        for _, row in new_ex.iterrows():
            amt = clean_currency(row.get('환전금액', 0))
            curr = str(row.get('화폐', '')).upper()
            if 'CNY' in curr:
                cny_bal += amt
            elif 'USD' in curr:
                usd_bal += amt

    if not df_d.empty and '실지급_날짜' in df_d.columns:
        df_d_paid = df_d[(df_d['실지급_날짜'] > base_date) & (df_d['진행단계'].astype(str).str.contains('완료', na=False))]
        for _, row in df_d_paid.iterrows():
            gubun = str(row.get('구분', ''))
            curr = str(row.get('화폐단위', '')).upper()
            amt_paid = clean_currency(row.get('실지급_금액', 0))
            amt_usd_actual = clean_currency(row.get('실제출금(USD)', 0))

            if 'USD' in gubun and 'CNY' in curr and amt_usd_actual > 0:
                usd_bal -= amt_usd_actual
            else:
                if 'USD' in curr:
                    usd_bal -= amt_paid
                elif 'CNY' in curr:
                    cny_bal -= amt_paid

    if not df_l.empty and '날짜' in df_l.columns:
        df_l_paid = df_l[(df_l['날짜'] > base_date) & (df_l['구분'].astype(str).str.contains('송금', na=False))]
        for _, row in df_l_paid.iterrows():
            amt_usd = clean_currency(row.get('입금액(USD)', 0))
            usd_bal -= amt_usd

    return cny_bal, usd_bal

def get_date_range(today):
    start_week = today - timedelta(days=today.weekday())
    end_next_month = (today.replace(day=1) + timedelta(days=65)).replace(day=1) - timedelta(days=1)
    return {
        "this_week": (start_week, start_week + timedelta(days=6)),
        "next_week": (start_week + timedelta(days=7), start_week + timedelta(days=13)),
        "this_month": (today.replace(day=1), (today.replace(day=1) + timedelta(days=32)).replace(day=1) - timedelta(days=1)),
        "next_month": ((today.replace(day=1) + timedelta(days=32)).replace(day=1), end_next_month),
        "this_plus_next_month": (today.replace(day=1), end_next_month)
    }

def split_direct_data(df):
    mask_usd = (df['구분'].astype(str).str.contains('USD|결제', case=False) | (df['화폐단위'] == 'USD'))
    return df[~mask_usd].copy(), df[mask_usd].copy()

# -----------------------------------------------------------
# 4. 사이드바
# -----------------------------------------------------------
FIXED_GDRIVE_URL = "https://docs.google.com/spreadsheets/d/1Sj1BNjMpBocCxTQV8OW_cK4-2-eClm2W/edit?usp=sharing&ouid=107636013985223863985&rtpof=true&sd=true"

with st.sidebar:
    st.title("⚙️ 자금 설정")
    menu = st.radio("화면 이동", ["전체 자금 현황", "다이렉트 (CNY)", "다이렉트 (USD)", "이우 (YIWU)", "환전 내역"])
    st.markdown("---")
    
    if st.button("🔄 최신 데이터 불러오기"):
        st.cache_data.clear()
        st.success("데이터를 새로 불러옵니다!")
        
    st.markdown("---")
    
    live_cny, live_usd, rate_source, is_stale_rate = get_live_exchange_rates()

    if live_cny is None or live_usd is None:
        st.warning("실시간 환율을 불러오지 못했습니다. 환율을 직접 입력해 주세요.")
        live_cny = DEFAULT_TARGET_CNY
        live_usd = DEFAULT_TARGET_USD
    elif is_stale_rate:
        st.warning(f"실시간 환율 조회 실패로 {rate_source}을 사용 중입니다.")
    else:
        st.caption(f"환율 출처: {rate_source}")
    
    col_r1, col_r2 = st.columns(2)
    with col_r1:
        rate_cny = st.number_input("1 CNY (원)", value=float(live_cny), format="%.2f")
    with col_r2:
        rate_usd = st.number_input("1 USD (원)", value=float(live_usd), format="%.2f")

    cny_to_usd_rate = rate_cny / rate_usd if rate_usd > 0 else 0

    st.markdown("---")
    st.markdown("#### 🎯 목표 환율 설정")
    col_t1, col_t2 = st.columns(2)
    with col_t1:
        target_cny_rate = st.number_input("CNY 미만 알림", value=DEFAULT_TARGET_CNY, step=1.0, format="%.2f")
    with col_t2:
        target_usd_rate = st.number_input("USD 미만 알림", value=DEFAULT_TARGET_USD, step=1.0, format="%.2f")

    st.markdown("---")
    today = pd.Timestamp.now().normalize()
    custom_date = st.date_input("📅 사용자 지정 기간", (today, today + timedelta(days=14)))

    st.markdown("---")
    with st.expander("초기 잔고 기준점 세팅"):
        base_date = st.date_input("기준 날짜", value=pd.to_datetime(DEFAULT_BASE_DATE))
        base_cny = st.number_input("초기 CNY", value=DEFAULT_BASE_CNY)
        base_usd = st.number_input("초기 USD", value=DEFAULT_BASE_USD)

# -----------------------------------------------------------
# 5. 화면 로직
# -----------------------------------------------------------
data_tuple = get_drive_data(FIXED_GDRIVE_URL)

if not data_tuple:
    st.warning("구글 드라이브 링크를 읽어오는 데 실패했습니다. 파일 공유 상태를 다시 확인해주세요.")
else:
    df_d, df_y, yiwu_balance, df_l, df_ex = data_tuple
    
    base_date_ts = pd.to_datetime(base_date)
    my_cny, my_usd = calculate_realtime_balances(df_d, df_ex, df_l, base_date_ts, base_cny, base_usd)
    
    with st.sidebar:
        st.subheader("환전보유액")
        st.metric("CNY 보유액", fmt_num(my_cny))
        st.metric("USD 보유액", fmt_num(my_usd))
    
    df_d_active = df_d[~df_d['진행단계'].astype(str).str.contains('완료', na=False)].copy() if '진행단계' in df_d.columns else df_d.copy()
    df_y_active = df_y[~df_y['진행단계'].astype(str).str.contains('완료', na=False)].copy() if '진행단계' in df_y.columns else df_y.copy()

    dates = get_date_range(today)
    
    df_d_active = df_d_active.sort_values('잔금_날짜')
    df_y_active = df_y_active.sort_values('잔금_날짜')
    
    df_cny_only, df_usd_only = split_direct_data(df_d_active)

    fixed_periods = [
        ("1. 이번주", dates['this_week'][0], dates['this_week'][1]),
        ("2. 다음주", dates['next_week'][0], dates['next_week'][1]),
        ("3. 이번주+다음주", dates['this_week'][0], dates['next_week'][1]),
        ("4. 이번달", dates['this_month'][0], dates['this_month'][1]),
        ("5. 다음달", dates['next_month'][0], dates['next_month'][1]),
        ("6. 이번달+다음달", dates['this_plus_next_month'][0], dates['this_plus_next_month'][1]),
    ]
    
    all_periods = [("0. 전체 예정", None, None)] + fixed_periods
    if len(custom_date) == 2:
        all_periods.append(("7. 사용자 지정", pd.to_datetime(custom_date[0]), pd.to_datetime(custom_date[1])))
        
    # =======================================================
    # PAGE 1: 전체 자금 현황
    # =======================================================
    if menu == "전체 자금 현황":
        top_col1, top_col2 = st.columns([7, 3])
        
        with top_col1:
            st.header("📊 전체 자금 현황 대시보드")
            
            alert_cny = rate_cny < target_cny_rate
            alert_usd = rate_usd < target_usd_rate
            
            if alert_cny and alert_usd:
                st.markdown(f"""
                <div style='background-color: #fdf2f2; border-left: 5px solid #e74c3c; padding: 12px; margin-top: -5px; margin-bottom: 15px; border-radius: 4px;'>
                    <span style='color:#c0392b; font-size:15px; font-weight:bold;'>
                        🚨 [환전 알림] 현재 USD({fmt_num(rate_usd)}원) 및 CNY({fmt_num(rate_cny)}원) 환율이 모두 목표가에 도달했습니다. 환전을 검토해 주시기 바랍니다.
                    </span>
                </div>
                """, unsafe_allow_html=True)
            elif alert_usd:
                st.markdown(f"""
                <div style='background-color: #fdf2f2; border-left: 5px solid #e74c3c; padding: 12px; margin-top: -5px; margin-bottom: 15px; border-radius: 4px;'>
                    <span style='color:#c0392b; font-size:15px; font-weight:bold;'>
                        🚨 [환전 알림] 현재 USD 환율({fmt_num(rate_usd)}원)이 설정하신 목표가에 도달했습니다. 환전을 검토해 주시기 바랍니다.
                    </span>
                </div>
                """, unsafe_allow_html=True)
            elif alert_cny:
                st.markdown(f"""
                <div style='background-color: #fdf2f2; border-left: 5px solid #e74c3c; padding: 12px; margin-top: -5px; margin-bottom: 15px; border-radius: 4px;'>
                    <span style='color:#c0392b; font-size:15px; font-weight:bold;'>
                        🚨 [환전 알림] 현재 CNY 환율({fmt_num(rate_cny)}원)이 설정하신 목표가에 도달했습니다. 환전을 검토해 주시기 바랍니다.
                    </span>
                </div>
                """, unsafe_allow_html=True)
            
        with top_col2:
            st.markdown("##### 💡 항목명 설명")
            st.session_state.memo_df = st.data_editor(
                st.session_state.memo_df, 
                use_container_width=False, 
                hide_index=True,
                num_rows="dynamic" 
            )

        c1, c2, c3 = st.columns(3)
        c1.metric("CNY 환전보유액", fmt_num(my_cny), f"≈ {fmt_krw(my_cny * rate_cny)} 원")
        c2.metric("USD 환전보유액", fmt_num(my_usd), f"≈ {fmt_krw(my_usd * rate_usd)} 원")
        yiwu_balance_usd = yiwu_balance * cny_to_usd_rate
        c3.metric("허사장님 물품대 (USD)", fmt_num(yiwu_balance_usd), f"≈ {fmt_krw(yiwu_balance_usd * rate_usd)} 원")
        
        st.markdown("---")

        st.subheader("통화별 환전 필요액")
        
        summary_rows = []
        for label, s, e in all_periods:
            sub_cny = df_cny_only[(df_cny_only['잔금_날짜'] >= s) & (df_cny_only['잔금_날짜'] <= e)] if s and e else df_cny_only
            exp_cny = sub_cny['잔금_금액'].sum()
            
            sub_usd = df_usd_only[(df_usd_only['잔금_날짜'] >= s) & (df_usd_only['잔금_날짜'] <= e)] if s and e else df_usd_only
            val_pure = sub_usd[sub_usd['화폐단위'] == 'USD']['잔금_금액'].sum()
            val_conv = sub_usd[sub_usd['화폐단위'] == 'CNY']['잔금_금액'].sum() * cny_to_usd_rate
            exp_usd_direct = val_pure + val_conv
            
            sub_yiwu = df_y_active[(df_y_active['잔금_날짜'] >= s) & (df_y_active['잔금_날짜'] <= e)] if s and e else df_y_active
            exp_yiwu_cny = sub_yiwu['잔금_금액'].sum()
            
            # ⭐ 총 물품대(USD) 누적 계산 (다이렉트 USD 총액 + 이우 USD 총액)
            gross_usd = exp_usd_direct + (exp_yiwu_cny * cny_to_usd_rate)

            if label in ["2. 다음주", "5. 다음달"]:
                prev_e = dates['this_week'][1] if label == "2. 다음주" else dates['this_month'][1]
                
                past_cny_sum = df_cny_only[df_cny_only['잔금_날짜'] <= prev_e]['잔금_금액'].sum()
                past_usd_pure = df_usd_only[(df_usd_only['잔금_날짜'] <= prev_e) & (df_usd_only['화폐단위'] == 'USD')]['잔금_금액'].sum()
                past_usd_conv = df_usd_only[(df_usd_only['잔금_날짜'] <= prev_e) & (df_usd_only['화폐단위'] == 'CNY')]['잔금_금액'].sum() * cny_to_usd_rate
                past_yiwu_cny = df_y_active[df_y_active['잔금_날짜'] <= prev_e]['잔금_금액'].sum()
                
                rem_cny = max(my_cny - past_cny_sum, 0)
                
                past_yiwu_short_cny = max(past_yiwu_cny - yiwu_balance, 0)
                past_req_usd = past_usd_pure + past_usd_conv + (past_yiwu_short_cny * cny_to_usd_rate)
                rem_usd = max(my_usd - past_req_usd, 0)
                
                rem_yiwu_bal = max(yiwu_balance - past_yiwu_cny, 0)
                yiwu_short_cny = max(exp_yiwu_cny - rem_yiwu_bal, 0)
                yiwu_req_usd = yiwu_short_cny * cny_to_usd_rate
                
                total_req_cny = exp_cny
                total_req_usd = exp_usd_direct + yiwu_req_usd
                
                final_short_cny = max(total_req_cny - rem_cny, 0)
                final_short_usd = max(total_req_usd - rem_usd, 0)
                
            else:
                yiwu_short_cny = max(exp_yiwu_cny - yiwu_balance, 0)
                yiwu_req_usd = yiwu_short_cny * cny_to_usd_rate
                
                total_req_cny = exp_cny
                total_req_usd = exp_usd_direct + yiwu_req_usd
                
                final_short_cny = max(total_req_cny - my_cny, 0)
                final_short_usd = max(total_req_usd - my_usd, 0)
            
            summary_rows.append({
                "기간": label,
                "총 물품대(CNY)": fmt_num(total_req_cny),
                "총 물품대(KRW) ": fmt_krw(total_req_cny * rate_cny), 
                "환전 필요액(CNY)": fmt_num(final_short_cny),
                "환전 필요액(KRW) ": fmt_krw(final_short_cny * rate_cny), 
                "총 물품대(USD)": fmt_num(gross_usd),
                "총 물품대(KRW)": fmt_krw(gross_usd * rate_usd),
                "부족 물품대(USD)": fmt_num(total_req_usd),
                "부족 물품대(KRW)": fmt_krw(total_req_usd * rate_usd),
                "환전 필요액(USD)": fmt_num(final_short_usd),
                "환전 필요액(KRW)": fmt_krw(final_short_usd * rate_usd)
            })
            
        st.dataframe(pd.DataFrame(summary_rows), hide_index=True, use_container_width=True)

        st.markdown("---")

        c_h, c_b = st.columns([5, 2])
        with c_h: st.subheader("1️⃣ 다이렉트 (CNY) 현황")
        with c_b: st.markdown(f"**💰 CNY 보유액:** :green[{fmt_num(my_cny)}]")
        
        rows_cny = []
        for label, s, e in all_periods:
            sub = df_cny_only[(df_cny_only['잔금_날짜'] >= s) & (df_cny_only['잔금_날짜'] <= e)] if s and e else df_cny_only
            exp_cny = sub['잔금_금액'].sum()
            
            if label in ["2. 다음주", "5. 다음달"]:
                prev_e = dates['this_week'][1] if label == "2. 다음주" else dates['this_month'][1]
                past_cny_sum = df_cny_only[df_cny_only['잔금_날짜'] <= prev_e]['잔금_금액'].sum()
                rem_cny = max(my_cny - past_cny_sum, 0)
                need_cny = max(exp_cny - rem_cny, 0)
            else:
                need_cny = max(exp_cny - my_cny, 0)
                
            rows_cny.append({
                "기간": label, 
                "총 물품대(CNY)": fmt_num(exp_cny), 
                "총 물품대(KRW)": fmt_krw(exp_cny * rate_cny),
                "환전 필요액(CNY)": fmt_num(need_cny), 
                "환전 필요액(KRW)": fmt_krw(need_cny * rate_cny) 
            })
        st.dataframe(pd.DataFrame(rows_cny), hide_index=True, use_container_width=True)

        st.markdown("---")

        c_h, c_b = st.columns([5, 2])
        with c_h: st.subheader("2️⃣ 다이렉트 (USD) 현황")
        with c_b: st.markdown(f"**💰 USD 보유액:** :green[{fmt_num(my_usd)}]")

        rows_usd = []
        for label, s, e in all_periods:
            sub = df_usd_only[(df_usd_only['잔금_날짜'] >= s) & (df_usd_only['잔금_날짜'] <= e)] if s and e else df_usd_only
            val_pure = sub[sub['화폐단위'] == 'USD']['잔금_금액'].sum()
            val_conv = sub[sub['화폐단위'] == 'CNY']['잔금_금액'].sum() * cny_to_usd_rate
            exp_usd = val_pure + val_conv
            
            if label in ["2. 다음주", "5. 다음달"]:
                prev_e = dates['this_week'][1] if label == "2. 다음주" else dates['this_month'][1]
                past_usd_pure = df_usd_only[(df_usd_only['잔금_날짜'] <= prev_e) & (df_usd_only['화폐단위'] == 'USD')]['잔금_금액'].sum()
                past_usd_conv = df_usd_only[(df_usd_only['잔금_날짜'] <= prev_e) & (df_usd_only['화폐단위'] == 'CNY')]['잔금_금액'].sum() * cny_to_usd_rate
                rem_usd = max(my_usd - (past_usd_pure + past_usd_conv), 0)
                need_usd = max(exp_usd - rem_usd, 0)
            else:
                need_usd = max(exp_usd - my_usd, 0)
                
            rows_usd.append({
                "기간": label, 
                "총 물품대(USD)": fmt_num(exp_usd), 
                "총 물품대(KRW)": fmt_krw(exp_usd * rate_usd),
                "환전 필요액(USD)": fmt_num(need_usd), 
                "환전 필요액(KRW)": fmt_krw(need_usd * rate_usd) 
            })
        st.dataframe(pd.DataFrame(rows_usd), hide_index=True, use_container_width=True)

        st.markdown("---")

        c_h, c_b1, c_b2 = st.columns([4, 2, 2])
        with c_h: st.subheader("3️⃣ 이우 (YIWU) 현황")
        with c_b1: st.markdown(f"**📒 허사장님 물품대 (USD):** :blue[{fmt_num(yiwu_balance_usd)}]")
        with c_b2: st.markdown(f"**💰 USD 보유액:** :green[{fmt_num(my_usd)}]")
        
        rows_yiwu = []
        for label, s, e in all_periods:
            sub = df_y_active[(df_y_active['잔금_날짜'] >= s) & (df_y_active['잔금_날짜'] <= e)] if s and e else df_y_active
            exp_cny = sub['잔금_금액'].sum()
            
            if label in ["2. 다음주", "5. 다음달"]:
                prev_e = dates['this_week'][1] if label == "2. 다음주" else dates['this_month'][1]
                past_yiwu_cny = df_y_active[df_y_active['잔금_날짜'] <= prev_e]['잔금_금액'].sum()
                
                rem_yiwu_bal = max(yiwu_balance - past_yiwu_cny, 0)
                short_cny = max(exp_cny - rem_yiwu_bal, 0)
                short_usd = short_cny * cny_to_usd_rate
                
                past_yiwu_short_cny = max(past_yiwu_cny - yiwu_balance, 0)
                past_req_usd = past_yiwu_short_cny * cny_to_usd_rate
                rem_usd = max(my_usd - past_req_usd, 0)
                
                remit_usd = max(short_usd - rem_usd, 0)
            else:
                short_cny = max(exp_cny - yiwu_balance, 0)
                short_usd = short_cny * cny_to_usd_rate
                remit_usd = max(short_usd - my_usd, 0)
                
            rows_yiwu.append({
                "기간": label, 
                "총 물품대(USD)": fmt_num(exp_cny * cny_to_usd_rate), 
                "총 물품대(KRW)": fmt_krw(exp_cny * rate_cny),
                "부족 물품대(USD)": fmt_num(short_usd), 
                "부족 물품대(KRW)": fmt_krw(short_usd * rate_usd),
                "환전 필요액(USD)": fmt_num(remit_usd), 
                "환전 필요액(KRW)": fmt_krw(remit_usd * rate_usd)
            })
        st.dataframe(pd.DataFrame(rows_yiwu), hide_index=True, use_container_width=True)

    # =======================================================
    # PAGE: 환전 내역
    # =======================================================
    elif menu == "환전 내역":
        st.header("💱 환전 내역 관리")
        if df_ex.empty: 
            st.warning("데이터가 없습니다.")
        else:
            df_cny_ex = df_ex[df_ex['화폐'].astype(str).str.upper() == 'CNY'].copy()
            df_usd_ex = df_ex[df_ex['화폐'].astype(str).str.upper() == 'USD'].copy()
            
            col1, col2 = st.columns(2)
            
            with col1:
                st.subheader("🇨🇳 CNY 환전 내역")
                if not df_cny_ex.empty:
                    disp_cny = df_cny_ex.copy()
                    date_col = '일자' if '일자' in disp_cny.columns else '날짜'
                    if date_col in disp_cny.columns:
                        disp_cny['환전 일자'] = pd.to_datetime(disp_cny[date_col]).dt.strftime('%Y-%m-%d')
                    else:
                        disp_cny['환전 일자'] = ""
                        
                    disp_cny['외화금액(CNY)'] = disp_cny.get('외화금액', disp_cny.get('환전금액', 0)).apply(clean_currency).apply(fmt_num)
                    disp_cny['환율'] = disp_cny.get('환율', 0).apply(clean_currency).apply(fmt_num)
                    disp_cny['원화금액(KRW)'] = disp_cny.get('원화금액', 0).apply(clean_currency).apply(fmt_krw)
                    
                    show_cols = ['환전 일자', '외화금액(CNY)', '환율', '원화금액(KRW)']
                    valid_cols = [c for c in show_cols if c in disp_cny.columns]
                    st.dataframe(disp_cny[valid_cols].sort_values('환전 일자', ascending=False), hide_index=True, use_container_width=True)
                else:
                    st.info("CNY 환전 내역이 없습니다.")
                    
            with col2:
                st.subheader("🇺🇸 USD 환전 내역")
                if not df_usd_ex.empty:
                    disp_usd = df_usd_ex.copy()
                    date_col = '일자' if '일자' in disp_usd.columns else '날짜'
                    if date_col in disp_usd.columns:
                        disp_usd['환전 일자'] = pd.to_datetime(disp_usd[date_col]).dt.strftime('%Y-%m-%d')
                    else:
                        disp_usd['환전 일자'] = ""
                        
                    disp_usd['외화금액(USD)'] = disp_usd.get('외화금액', disp_usd.get('환전금액', 0)).apply(clean_currency).apply(fmt_num)
                    disp_usd['환율'] = disp_usd.get('환율', 0).apply(clean_currency).apply(fmt_num)
                    disp_usd['원화금액(KRW)'] = disp_usd.get('원화금액', 0).apply(clean_currency).apply(fmt_krw)
                    
                    show_cols = ['환전 일자', '외화금액(USD)', '환율', '원화금액(KRW)']
                    valid_cols = [c for c in show_cols if c in disp_usd.columns]
                    st.dataframe(disp_usd[valid_cols].sort_values('환전 일자', ascending=False), hide_index=True, use_container_width=True)
                else:
                    st.info("USD 환전 내역이 없습니다.")

    # =======================================================
    # PAGE: 다이렉트 CNY 
    # =======================================================
    elif menu == "다이렉트 (CNY)":
        st.header("다이렉트 관리 (CNY)")
        c1, c2 = st.columns(2)
        c1.metric("CNY 보유액", fmt_num(my_cny), f"≈ {fmt_krw(my_cny * rate_cny)} 원")
        
        df_cny_only, _ = split_direct_data(df_d_active)
        df_view = df_cny_only.copy()
        
        rows = []
        for label, s, e in all_periods:
            sub = df_view[(df_view['잔금_날짜'] >= s) & (df_view['잔금_날짜'] <= e)] if s and e else df_view
            exp_cny = sub['잔금_금액'].sum()
            
            if label in ["2. 다음주", "5. 다음달"]:
                prev_e = dates['this_week'][1] if label == "2. 다음주" else dates['this_month'][1]
                past_cny_sum = df_cny_only[df_cny_only['잔금_날짜'] <= prev_e]['잔금_금액'].sum()
                rem_cny = max(my_cny - past_cny_sum, 0)
                need_cny = max(exp_cny - rem_cny, 0)
            else:
                need_cny = max(exp_cny - my_cny, 0)
                
            rows.append({
                "기간": label, 
                "총 물품대(CNY)": fmt_num(exp_cny), 
                "총 물품대(KRW)": fmt_krw(exp_cny * rate_cny),
                "환전 필요액(CNY)": fmt_num(need_cny), 
                "환전 필요액(KRW)": fmt_krw(need_cny * rate_cny) 
            })
        st.subheader("📅 기간별 CNY 자금 계획")
        st.dataframe(pd.DataFrame(rows), hide_index=True, use_container_width=True)
        
        st.markdown("---")
        st.subheader("📋 상세 내역")
        
        df_view = df_view.sort_values('잔금_날짜').copy()
        df_view['잔금 금액(KRW)'] = df_view['잔금_금액'] * rate_cny
        
        df_view['누적_잔금(CNY)'] = df_view['잔금_금액'].cumsum()
        df_view['환전 필요액(CNY)'] = (df_view['누적_잔금(CNY)'] - my_cny).clip(lower=0) 
        df_view['환전 필요액(KRW)'] = df_view['환전 필요액(CNY)'] * rate_cny
        
        df_disp = df_view[['잔금_날짜', '품목', '거래처', '잔금_금액', '잔금 금액(KRW)', '환전 필요액(CNY)', '환전 필요액(KRW)', '진행단계']].copy()
        df_disp.columns = ['잔금 날짜', '상품명', '거래처', '총 물품대(CNY)', '총 물품대(KRW)', '환전 필요액(CNY)', '환전 필요액(KRW)', '진행단계']
        
        if '잔금 날짜' in df_disp.columns: df_disp['잔금 날짜'] = df_disp['잔금 날짜'].dt.strftime('%Y-%m-%d')
        if '총 물품대(CNY)' in df_disp.columns: df_disp['총 물품대(CNY)'] = df_disp['총 물품대(CNY)'].apply(fmt_num)
        if '총 물품대(KRW)' in df_disp.columns: df_disp['총 물품대(KRW)'] = df_disp['총 물품대(KRW)'].apply(fmt_krw)
        if '환전 필요액(CNY)' in df_disp.columns: df_disp['환전 필요액(CNY)'] = df_disp['환전 필요액(CNY)'].apply(fmt_num)
        if '환전 필요액(KRW)' in df_disp.columns: df_disp['환전 필요액(KRW)'] = df_disp['환전 필요액(KRW)'].apply(fmt_krw)
        
        st.dataframe(df_disp, hide_index=True, use_container_width=True)

    # =======================================================
    # PAGE: 다이렉트 USD 
    # =======================================================
    elif menu == "다이렉트 (USD)":
        st.header("다이렉트 관리 (USD)")
        c1, c2 = st.columns(2)
        c1.metric("USD 보유액", fmt_num(my_usd), f"≈ {fmt_krw(my_usd * rate_usd)} 원")
        
        _, df_usd_only = split_direct_data(df_d_active)
        df_view = df_usd_only.copy()
        
        rows = []
        for label, s, e in all_periods:
            sub = df_view[(df_view['잔금_날짜'] >= s) & (df_view['잔금_날짜'] <= e)] if s and e else df_view
            val_pure = sub[sub['화폐단위'] == 'USD']['잔금_금액'].sum()
            val_conv = sub[sub['화폐단위'] == 'CNY']['잔금_금액'].sum() * cny_to_usd_rate
            exp_usd = val_pure + val_conv
            
            if label in ["2. 다음주", "5. 다음달"]:
                prev_e = dates['this_week'][1] if label == "2. 다음주" else dates['this_month'][1]
                past_usd_pure = df_usd_only[(df_usd_only['잔금_날짜'] <= prev_e) & (df_usd_only['화폐단위'] == 'USD')]['잔금_금액'].sum()
                past_usd_conv = df_usd_only[(df_usd_only['잔금_날짜'] <= prev_e) & (df_usd_only['화폐단위'] == 'CNY')]['잔금_금액'].sum() * cny_to_usd_rate
                rem_usd = max(my_usd - (past_usd_pure + past_usd_conv), 0)
                need_usd = max(exp_usd - rem_usd, 0)
            else:
                need_usd = max(exp_usd - my_usd, 0)
                
            rows.append({
                "기간": label, 
                "총 물품대(USD)": fmt_num(exp_usd), 
                "총 물품대(KRW)": fmt_krw(exp_usd * rate_usd),
                "환전 필요액(USD)": fmt_num(need_usd), 
                "환전 필요액(KRW)": fmt_krw(need_usd * rate_usd) 
            })
        st.subheader("📅 기간별 USD 자금 계획")
        st.dataframe(pd.DataFrame(rows), hide_index=True, use_container_width=True)
        
        st.markdown("---")
        st.subheader("📋 상세 내역")
        
        df_view = df_view.sort_values('잔금_날짜').copy()
        df_view['잔금 금액(CNY)'] = df_view.apply(lambda r: r['잔금_금액'] if r['화폐단위'] == 'CNY' else 0, axis=1)
        df_view['잔금 금액(USD)'] = df_view.apply(lambda r: r['잔금_금액'] if r['화폐단위'] == 'USD' else r['잔금_금액'] * cny_to_usd_rate, axis=1)
        df_view['잔금 금액(KRW)'] = df_view['잔금 금액(USD)'] * rate_usd
        
        df_view['누적_잔금(USD)'] = df_view['잔금 금액(USD)'].cumsum()
        df_view['환전 필요액(USD)'] = (df_view['누적_잔금(USD)'] - my_usd).clip(lower=0) 
        df_view['환전 필요액(KRW)'] = df_view['환전 필요액(USD)'] * rate_usd
        
        df_disp = df_view[['잔금_날짜', '품목', '거래처', '잔금 금액(CNY)', '잔금 금액(USD)', '잔금 금액(KRW)', '환전 필요액(USD)', '환전 필요액(KRW)', '진행단계']].copy()
        df_disp.columns = ['잔금 날짜', '상품명', '거래처', '총 물품대(CNY)', '총 물품대(USD)', '총 물품대(KRW)', '환전 필요액(USD)', '환전 필요액(KRW)', '진행단계']
        
        if '잔금 날짜' in df_disp.columns: df_disp['잔금 날짜'] = df_disp['잔금 날짜'].dt.strftime('%Y-%m-%d')
        if '총 물품대(CNY)' in df_disp.columns: df_disp['총 물품대(CNY)'] = df_disp['총 물품대(CNY)'].apply(lambda x: fmt_num(x) if x > 0 else "")
        if '총 물품대(USD)' in df_disp.columns: df_disp['총 물품대(USD)'] = df_disp['총 물품대(USD)'].apply(fmt_num)
        if '총 물품대(KRW)' in df_disp.columns: df_disp['총 물품대(KRW)'] = df_disp['총 물품대(KRW)'].apply(fmt_krw)
        if '환전 필요액(USD)' in df_disp.columns: df_disp['환전 필요액(USD)'] = df_disp['환전 필요액(USD)'].apply(fmt_num)
        if '환전 필요액(KRW)' in df_disp.columns: df_disp['환전 필요액(KRW)'] = df_disp['환전 필요액(KRW)'].apply(fmt_krw)
        
        st.dataframe(df_disp, hide_index=True, use_container_width=True)

    # =======================================================
    # PAGE: 이우 (YIWU) 
    # =======================================================
    elif menu == "이우 (YIWU)":
        st.header("이우(YIWU) 자금 관리")
        c1, c2 = st.columns(2)
        
        yiwu_balance_usd = yiwu_balance * cny_to_usd_rate
        c1.metric("허사장님 물품대 (USD)", fmt_num(yiwu_balance_usd), f"≈ {fmt_krw(yiwu_balance_usd * rate_usd)} 원")
        c2.metric("USD 보유액", fmt_num(my_usd), f"≈ {fmt_krw(my_usd * rate_usd)} 원")
        
        rows = []
        for label, s, e in all_periods:
            sub = df_y_active[(df_y_active['잔금_날짜'] >= s) & (df_y_active['잔금_날짜'] <= e)] if s and e else df_y_active
            exp_cny = sub['잔금_금액'].sum()
            exp_usd = exp_cny * cny_to_usd_rate
            
            if label in ["2. 다음주", "5. 다음달"]:
                prev_e = dates['this_week'][1] if label == "2. 다음주" else dates['this_month'][1]
                past_yiwu_cny = df_y_active[df_y_active['잔금_날짜'] <= prev_e]['잔금_금액'].sum()
                
                rem_yiwu_bal = max(yiwu_balance - past_yiwu_cny, 0)
                short_cny = max(exp_cny - rem_yiwu_bal, 0)
                short_usd = short_cny * cny_to_usd_rate
                
                past_yiwu_short_cny = max(past_yiwu_cny - yiwu_balance, 0)
                rem_usd = max(my_usd - (past_yiwu_short_cny * cny_to_usd_rate), 0)
                remit_usd = max(short_usd - rem_usd, 0)
            else:
                short_cny = max(exp_cny - yiwu_balance, 0)
                short_usd = short_cny * cny_to_usd_rate
                remit_usd = max(short_usd - my_usd, 0)
                
            rows.append({
                "기간": label, 
                "총 물품대(USD)": fmt_num(exp_cny * cny_to_usd_rate), 
                "총 물품대(KRW)": fmt_krw(exp_cny * rate_cny),
                "부족 물품대(USD)": fmt_num(short_usd), 
                "부족 물품대(KRW)": fmt_krw(short_usd * rate_usd),
                "환전 필요액(USD)": fmt_num(remit_usd), 
                "환전 필요액(KRW)": fmt_krw(remit_usd * rate_usd)
            })
        st.subheader("📅 기간별 이우(YIWU) 자금 계획")
        st.dataframe(pd.DataFrame(rows), hide_index=True, use_container_width=True)
        
        st.markdown("---")
        st.subheader("📋 상세 내역")
        
        df_disp = df_y_active.sort_values('잔금_날짜').copy()
        
        df_disp['잔금 금액(USD)'] = df_disp['잔금_금액'] * cny_to_usd_rate
        df_disp['잔금 금액(KRW)'] = df_disp['잔금 금액(USD)'] * rate_usd
        
        df_disp['누적_잔금(CNY)'] = df_disp['잔금_금액'].cumsum()
        df_disp['부족 물품대(CNY)'] = (df_disp['누적_잔금(CNY)'] - yiwu_balance).clip(lower=0)
        
        df_disp['부족 물품대(USD)'] = df_disp['부족 물품대(CNY)'] * cny_to_usd_rate
        df_disp['부족 물품대(KRW)'] = df_disp['부족 물품대(USD)'] * rate_usd
        
        df_disp['환전 필요액(USD)'] = (df_disp['부족 물품대(USD)'] - my_usd).clip(lower=0)
        df_disp['환전 필요액(KRW)'] = df_disp['환전 필요액(USD)'] * rate_usd
        
        cols_to_show = ['잔금_날짜', '품목', '잔금 금액(USD)', '잔금 금액(KRW)', '부족 물품대(USD)', '부족 물품대(KRW)', '환전 필요액(USD)', '환전 필요액(KRW)', '진행단계']
        disp_final = df_disp[[c for c in cols_to_show if c in df_disp.columns]].copy()
        disp_final.rename(columns={
            '잔금_날짜': '잔금 날짜', 
            '품목': '상품명', 
            '진행단계': '진행 단계',
            '잔금 금액(USD)': '총 물품대(USD)',
            '잔금 금액(KRW)': '총 물품대(KRW)'
        }, inplace=True)
        
        if '잔금 날짜' in disp_final.columns: disp_final['잔금 날짜'] = disp_final['잔금 날짜'].dt.strftime('%Y-%m-%d')
        if '총 물품대(USD)' in disp_final.columns: disp_final['총 물품대(USD)'] = disp_final['총 물품대(USD)'].apply(fmt_num)
        if '총 물품대(KRW)' in disp_final.columns: disp_final['총 물품대(KRW)'] = disp_final['총 물품대(KRW)'].apply(fmt_krw)
        if '부족 물품대(USD)' in disp_final.columns: disp_final['부족 물품대(USD)'] = disp_final['부족 물품대(USD)'].apply(fmt_num)
        if '부족 물품대(KRW)' in disp_final.columns: disp_final['부족 물품대(KRW)'] = disp_final['부족 물품대(KRW)'].apply(fmt_krw)
        if '환전 필요액(USD)' in disp_final.columns: disp_final['환전 필요액(USD)'] = disp_final['환전 필요액(USD)'].apply(fmt_num)
        if '환전 필요액(KRW)' in disp_final.columns: disp_final['환전 필요액(KRW)'] = disp_final['환전 필요액(KRW)'].apply(fmt_krw)
        
        st.dataframe(disp_final, hide_index=True, use_container_width=True)
