import os
import time
import asyncio
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload
from playwright.async_api import async_playwright

# -----------------------------------------------------------------
# 🎯 專案參數與 API 設定
# -----------------------------------------------------------------
# Require read/write access so we can update other vessel spreadsheets (Step 11),
# and Drive read access so we can download PO attachments (Step 10).
# If you previously authorized with fewer scopes, delete `token.json` to re-authorize with the new scopes.
SCOPES = ['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/drive.readonly']
SPREADSHEET_ID = '1tMonCbumPejfQCYgC8IwmPqBTZ9Tz5yuK7c-2qIk_8Y'
RANGE_NAME = "'pending'!A2:K99"  # 讀取 Column A 至 Column K（包含 vessel_name）
D365_URL = "https://wahkwong.operations.dynamics.com/?cmp=WGI&mi=THK_PurchWorkspace"
CLIENT_SECRET_FILE = 'client_secret.json'
DRIVE_DOWNLOAD_DIR = "/Users/fungz/Downloads"
RELOAD_EVERY_N_PO = 8  # 每處理幾多筆 PO 就強制重新載入頁面一次，避免 D365 SPA 隨時間變慢導致 timeout

# 用戶 Browser Profile 儲存路徑（針對 macOS 環境，讓 Session/Cookie 持久化，避免重複 MFA）
USER_DATA_DIR = os.path.abspath("./playwright_user_data")

# -----------------------------------------------------------------
# 🎯 Step 0 設定：本機重現 n8n workflow「invoices pending upload」嘅邏輯，
# 喺 RPA 開始前更新 pending sheet 內容（唔再依賴 n8n）
# -----------------------------------------------------------------
PENDING_VESSEL_NAMES = [
    "CSSC Cape Town", "CSSC Gladstone", "CSSC Le Havre", "CSSC Immingham",
    "CSSC Liao Ning", "CS Hunan Venture", "Anne", "Gas Quest", "BH POWER", "Australis"
]
STEP0_MAX_RETRIES = 5
STEP0_RETRY_DELAY_SEC = 30  # 失敗（例如 Google API too many requests）後等待再重試

def safe_get(row, index):
    """安全擷取欄位資料（避免 Row 尾端陣列越界）"""
    return str(row[index]).strip() if len(row) > index and row[index] is not None else ""

def clean_amount(val_str):
    """清理金額字串並轉換為 float 以進行精準數值比對"""
    if not val_str:
        return 0.0
    cleaned = str(val_str).replace(',', '').replace('$', '').strip()
    try:
        return float(cleaned)
    except ValueError:
        return 0.0


def get_vessel_spreadsheet_id(vessel_name):
    """根據 vessel_name 取得對應的 vessel_po_list spreadsheet id。"""
    mapping = {
        "cssc le havre": "1iGYuhjwQaiwPLXHUgAthAQC_YuLL83K6ympzpbCjWWI",
        "cssc cape town": "19-xqXf3j9iliUvPJov_NOOCG2dC8Y0Nx6UuMToeumzg",
        "cssc immingham": "1QhB9qCBwYwCL5aSoltHYokoBX9-d78o6VEy9AcF1XiE",
        "cssc gladstone": "1B0XrDZjNcx4UTjajIGUPnnYp5alClKxELYZY8ye-JR0",
        "anne": "1dSi3BxHKVSCHAw-ztMI8CGNxARGKxmy6TN9PmfxR5zw",
        "gas quest": "1NC1lzsBD_NOP3LXSCRN1yvL6qdu42uUBgFkyAyxjz60",
        "bh power": "1ZJZ2pyTFOge7-inVlqlufS77VgV3V0mIAIBasf0tA0s",
        "australis": "1msV0Saagu-cB2WzTl9Kfa3CxcyohF9JW0La4ElZEGjU",
        "cs hunan venture": "1-DxjOaG_KGAhckjZB1bJfX7rUI-QwROrpufSo-l7IZI"
    }
    return mapping.get(vessel_name.lower()) if vessel_name else None


# account_code 前綴 ➔ credit_account 對應表 (Step 9 用)，方便日後修改
CREDIT_ACCOUNT_PREFIX_MAPPING = {
    "61004": "c5600604",
    "61007": "C5600608",
    "61006": "c5600604",
}
CREDIT_ACCOUNT_DEFAULT = "C5600602"


def get_credit_account(ac_code):
    """根據 account_code 前綴取得 Step 9 應輸入的 credit_account。"""
    ac_code = ac_code.strip()
    for prefix, credit_account in CREDIT_ACCOUNT_PREFIX_MAPPING.items():
        if ac_code.startswith(prefix):
            return credit_account
    return CREDIT_ACCOUNT_DEFAULT


def update_vessel_sheet_status(vessel_name, po_number, value):
    """更新對應 vessel spreadsheet 的 Column R，支援 draft / error_msg。"""
    vessel_po_list = get_vessel_spreadsheet_id(vessel_name)

    if not vessel_po_list:
        print(f"⚠️ No matching vessel_po_list found for vessel_name '{vessel_name}', skipping update.")
        return False

    try:
        creds = None
        if os.path.exists('token.json'):
            creds = Credentials.from_authorized_user_file('token.json', SCOPES)

        if not creds or not creds.valid:
            if creds and creds.expired and creds.refresh_token:
                creds.refresh(Request())
            else:
                print("⚠️ No valid credentials available. Please obtain token.json via OAuth first. Cannot update the external spreadsheet.")
                return False

        service = build('sheets', 'v4', credentials=creds)
        sheet = service.spreadsheets()

        read_range = 'B2:B'
        resp = sheet.values().get(spreadsheetId=vessel_po_list, range=read_range).execute()
        other_rows = resp.get('values', [])

        found = False
        for i, r in enumerate(other_rows):
            cell_val = str(r[0]).strip() if len(r) > 0 else ""
            if cell_val == po_number:
                row_number = i + 2
                update_range = f'R{row_number}'
                body = {'values': [[value]]}
                sheet.values().update(
                    spreadsheetId=vessel_po_list,
                    range=update_range,
                    valueInputOption='RAW',
                    body=body
                ).execute()
                print(f"✅ Updated {update_range} in spreadsheet {vessel_po_list} to '{value}' (PO={po_number})")
                found = True
                break

        if not found:
            print(f"⚠️ PO not found in spreadsheet {vessel_po_list}: {po_number}")
            return False

        return True

    except Exception as e:
        print(f"❌ Status update failed: {e}")
        return False


async def wait_for_new_row_after_add_line(page, max_wait_ms=10000, poll_interval_ms=200, settle_ms=1000):
    """點擊 Add line 後，偵測表格新增一行 row，偵測到後再等待 settle_ms 才繼續（取代固定等待 3 秒）。"""
    row_locator_candidates = [
        "div[role='row']",
        "tr[role='row']",
        "[role='grid'] [role='row']",
        "table tr",
    ]

    active_selector = None
    for sel in row_locator_candidates:
        if await page.locator(sel).count() > 0:
            active_selector = sel
            break

    if not active_selector:
        print(f"⚠️ Could not locate table row elements, falling back to a fixed {settle_ms / 1000:.0f}s wait...")
        await page.wait_for_timeout(settle_ms)
        return False

    row_count_before = await page.locator(active_selector).count()
    print(f"⏳ Waiting for a new row to appear (current row count: {row_count_before})...")

    elapsed = 0
    while elapsed < max_wait_ms:
        await page.wait_for_timeout(poll_interval_ms)
        elapsed += poll_interval_ms
        row_count_now = await page.locator(active_selector).count()
        if row_count_now > row_count_before:
            print(f"✅ New row detected (current row count: {row_count_now}), waiting {settle_ms / 1000:.0f}s before continuing...")
            await page.wait_for_timeout(settle_ms)
            return True

    print(f"⚠️ No new row detected after {max_wait_ms / 1000:.0f}s, waiting {settle_ms / 1000:.0f}s before continuing anyway...")
    await page.wait_for_timeout(settle_ms)
    return False


async def click_finance_and_operations(page):
    """精準定位並點擊 Finance and Operations，等待頁面刷新回到首頁。"""
    print("👉 Locating and clicking 'Finance and Operations'...")
    finance_locators = [
        page.get_by_role("link", name="Finance and Operations", exact=True),
        page.get_by_role("button", name="Finance and Operations", exact=True),
        page.get_by_text("Finance and Operations", exact=True)
    ]

    finance_target = None
    for locator in finance_locators:
        if await locator.count() > 0:
            finance_target = locator.first
            print("🎯 Matched 'Finance and Operations' element")
            break

    if finance_target:
        await finance_target.wait_for(state="visible", timeout=15000)
        await finance_target.click()
        print("✅ Clicked 'Finance and Operations', waiting for the page to refresh back to the home screen...")
        await page.wait_for_load_state("networkidle", timeout=20000)
        await page.wait_for_timeout(3000)
        print("✅ Page refreshed, back to the home screen.")
    else:
        print("⚠️ Could not precisely locate 'Finance and Operations', skipping click.")


def _get_drive_service():
    """建立 Google Drive API service，重用與 Sheets 相同的 token.json 憑證。"""
    creds = None
    if os.path.exists('token.json'):
        creds = Credentials.from_authorized_user_file('token.json', SCOPES)

    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            print("⚠️ No valid credentials available. Please re-authorize via OAuth to get a token.json with Drive access.")
            return None

    return build('drive', 'v3', credentials=creds)


def _find_drive_file_id(service, file_name):
    """在 Google Drive (含共用雲端硬碟) 依檔名精準搜尋檔案 id；若同名檔案多於一個，只取最新（modifiedTime 最新）的一個。"""
    safe_name = file_name.replace("'", "\\'")
    query = f"name = '{safe_name}' and trashed = false"
    resp = service.files().list(
        q=query,
        fields="files(id, name, modifiedTime)",
        spaces="drive",
        corpora="allDrives",
        includeItemsFromAllDrives=True,
        supportsAllDrives=True,
        orderBy="modifiedTime desc"
    ).execute()
    files = resp.get('files', [])
    if not files:
        return None
    if len(files) > 1:
        print(f"⚠️ Found {len(files)} files named '{file_name}' on Google Drive, using only the most recently modified one (modifiedTime={files[0].get('modifiedTime')})")
    return files[0]['id']


def download_po_attachments(po_number, need_credit_note):
    """從 Google Drive 依 PO 號碼下載 invoice / receipt / (credit note) 至 DRIVE_DOWNLOAD_DIR。"""
    os.makedirs(DRIVE_DOWNLOAD_DIR, exist_ok=True)

    service = _get_drive_service()
    if not service:
        return {}

    file_map = {
        "invoice": f"{po_number}_invoice",
        "receipt": f"{po_number}_receipt",
    }
    if need_credit_note:
        file_map["credit_note"] = f"{po_number}_credit note"

    downloaded = {}
    for key, file_name in file_map.items():
        print(f"🔍 Searching Google Drive for file: '{file_name}'...")
        file_id = _find_drive_file_id(service, file_name)
        if not file_id:
            print(f"❌ File not found on Google Drive: '{file_name}'")
            continue

        dest_path = os.path.join(DRIVE_DOWNLOAD_DIR, f"{file_name}.pdf")
        try:
            request = service.files().get_media(fileId=file_id, supportsAllDrives=True)
            with open(dest_path, 'wb') as f:
                downloader = MediaIoBaseDownload(f, request)
                done = False
                while not done:
                    _, done = downloader.next_chunk()
            print(f"✅ Downloaded '{file_name}' to {dest_path}")
            downloaded[key] = dest_path
        except Exception as e:
            print(f"❌ Failed to download '{file_name}': {e}")

    return downloaded


async def upload_attachment_via_new(page, file_path, doc_type_label):
    """在 Attachments 畫面點擊 New ➔ 選擇附件類型 ➔ Browse 上傳指定檔案。"""
    print(f"👉 Uploading attachment ({doc_type_label}): {file_path}")

    new_btn_candidates = [
        page.locator("span.button-label-dropDown[id$='_ctrlAdd_label']", has_text="New"),
        page.get_by_role("button", name="+New", exact=True),
        page.get_by_role("button", name="New", exact=True),
        page.get_by_text("+New", exact=True),
        page.get_by_text("New", exact=True),
    ]

    new_btn = None
    for candidate in new_btn_candidates:
        if await candidate.count() > 0:
            new_btn = candidate
            break

    if not new_btn:
        new_btn = page.get_by_role("button", name="New", exact=True)

    await new_btn.first.wait_for(state="visible", timeout=15000)
    await new_btn.first.click()
    print("✅ Clicked 'New', waiting for the dropdown menu to appear...")
    await page.wait_for_timeout(500)

    type_option = page.get_by_role("menuitem", name=doc_type_label, exact=True)
    if await type_option.count() == 0:
        type_option = page.get_by_text(doc_type_label, exact=True)
    await type_option.first.wait_for(state="visible", timeout=10000)
    await type_option.first.click()
    print(f"✅ Selected attachment type: {doc_type_label}")
    await page.wait_for_timeout(1000)

    browse_candidates = [
        page.locator("[id$='_UploadControlBrowseButton']"),
        page.get_by_role("button", name="Browse", exact=False),
        page.get_by_text("Browse", exact=False),
    ]

    browse_btn = None
    for candidate in browse_candidates:
        if await candidate.count() > 0:
            browse_btn = candidate
            break

    if not browse_btn:
        browse_btn = page.get_by_text("Browse", exact=False)

    await browse_btn.first.wait_for(state="visible", timeout=15000)

    async with page.expect_file_chooser() as fc_info:
        await browse_btn.first.click()
    file_chooser = await fc_info.value
    await file_chooser.set_files(file_path)
    print(f"✅ Selected file: {file_path}")
    await page.wait_for_timeout(1000)

    confirm_candidates = ["OK", "Save", "Attach", "Insert"]
    confirmed = False
    for label in confirm_candidates:
        btn = page.get_by_role("button", name=label, exact=True)
        if await btn.count() > 0 and await btn.first.is_visible():
            await btn.first.click()
            print(f"✅ Clicked '{label}', attachment upload complete.")
            await page.wait_for_timeout(1000)
            confirmed = True
            break

    if not confirmed:
        print("⚠️ Could not automatically locate the confirm-upload button, please check the D365 screen to confirm the attachment was uploaded.")


def get_data_via_sheets_api():
    """【純 API Call】呼叫 Google Sheets API v4 讀取資料"""
    creds = None

    # 讀取已保存的 Token
    if os.path.exists('token.json'):
        creds = Credentials.from_authorized_user_file('token.json', SCOPES)

    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            if not os.path.exists(CLIENT_SECRET_FILE):
                raise FileNotFoundError(
                    f"❌ [API Error] '{CLIENT_SECRET_FILE}' not found!\n"
                    "This OAuth credential file is required for pure API calls. Please make sure it is placed in the project root directory."
                )
            flow = InstalledAppFlow.from_client_secrets_file(CLIENT_SECRET_FILE, SCOPES)
            creds = flow.run_local_server(port=0)

        # 儲存憑證供下次 API 自動呼叫
        with open('token.json', 'w') as token:
            token.write(creds.to_json())

    print("🌐 Making Google Sheets API call...")
    service = build('sheets', 'v4', credentials=creds)
    sheet = service.spreadsheets()

    # API 讀取請求
    result = sheet.values().get(spreadsheetId=SPREADSHEET_ID, range=RANGE_NAME).execute()
    rows = result.get('values', [])

    records = []
    for row in rows:
        vessel_name = safe_get(row, 0)   # Column A
        po_number = safe_get(row, 1)     # Column B

        # 只保留以 WGI-PO 開頭的 PO，排除其他格式
        if po_number and po_number.upper().startswith("WGI-PO"):
            record = {
                "vessel_name": vessel_name,           # Column A
                "po_number": po_number,               # Column B
                "account_code": safe_get(row, 2),     # Column C
                "invoice_number": safe_get(row, 7),   # Column H
                "invoice_amount": safe_get(row, 8),   # Column I
                "credit_amount": safe_get(row, 9),    # Column J
                "difference": safe_get(row, 10)       # Column K
            }
            records.append(record)

    print(f"✅ [API Success] Retrieved {len(records)} pending record(s):")
    for idx, item in enumerate(records, 1):
        print(f"  [{idx}] Vessel: {item['vessel_name']} | PO: {item['po_number']} | AccountCode: {item['account_code']} | Inv#: {item['invoice_number']} | Amt: {item['invoice_amount']} | CreditAmt: {item['credit_amount']} | Diff: {item['difference']}")

    return records


def _get_sheets_service():
    """建立 Google Sheets API service，重用與 Drive 相同的 token.json 憑證。"""
    creds = None
    if os.path.exists('token.json'):
        creds = Credentials.from_authorized_user_file('token.json', SCOPES)

    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            print("⚠️ No valid credentials available. Please obtain token.json via OAuth first.")
            return None

    return build('sheets', 'v4', credentials=creds)


def _clear_pending_sheet(sheets_service):
    """清空 pending sheet 內既有資料列（保留 header）。"""
    sheets_service.spreadsheets().values().clear(
        spreadsheetId=SPREADSHEET_ID, range="'pending'!A2:K1000"
    ).execute()


def _read_vessel_pending_rows(sheets_service, drive_service, vessel_name):
    """
    為單一 vessel，喺 Google Drive 搵返 '{vessel_name}_status' spreadsheet，
    讀取其 'tech' 分頁，回傳符合條件（Receipt=Yes、有 Invoice Number、Invoice Status 未填）的 rows。
    """
    file_name = f"{vessel_name}_status"
    file_id = _find_drive_file_id(drive_service, file_name)
    if not file_id:
        print(f"⚠️ '{file_name}' not found, skipping {vessel_name}.")
        return []

    result = sheets_service.spreadsheets().values().get(
        spreadsheetId=file_id, range="'tech'!A1:Z9999"
    ).execute()
    rows = result.get('values', [])
    if not rows:
        return []

    header = [h.strip() for h in rows[0]]  # 'tech' 分頁部分 header（例如 ' Credit amount '）前後有多餘空格
    matched = []
    for row in rows[1:]:
        rec = {header[i]: (row[i] if i < len(row) else "") for i in range(len(header))}
        if (
            str(rec.get("Receipt", "")).strip() == "Yes"
            and str(rec.get("Invoice Number", "")).strip() != ""
            and str(rec.get("Invoice Status", "")).strip() == ""
        ):
            matched.append(rec)
    return matched


def _append_pending_rows(sheets_service, vessel_name, records):
    """將指定 vessel 已符合條件嘅 records 追加到 pending sheet。"""
    values = []
    for rec in records:
        values.append([
            vessel_name,
            rec.get("PO Number", ""),
            rec.get("Account Code", ""),
            rec.get("Description", ""),
            rec.get("Supplier", ""),
            rec.get("CUR", ""),
            rec.get("Remarks", ""),
            rec.get("Invoice Number", ""),
            rec.get("Invoice Amount", ""),
            rec.get("Credit amount", ""),
            rec.get("difference", ""),
        ])
    if not values:
        return

    sheets_service.spreadsheets().values().append(
        spreadsheetId=SPREADSHEET_ID,
        range="'pending'!A2:K2",
        valueInputOption="USER_ENTERED",
        insertDataOption="INSERT_ROWS",
        body={"values": values}
    ).execute()


def refresh_pending_invoices():
    """
    Step 0: 本機重現 n8n workflow「invoices pending upload」嘅邏輯——
    清空 pending sheet，然後逐個 vessel 喺其 '{vessel}_status' spreadsheet 嘅 'tech' 分頁
    搵出 Receipt=Yes、有 Invoice Number、Invoice Status 未填嘅列，寫返入 pending sheet。
    失敗（例如 Google API too many requests）時等待 STEP0_RETRY_DELAY_SEC 秒後重試，
    最多 STEP0_MAX_RETRIES 次。回傳 True/False 代表整體是否成功。
    """
    for attempt in range(1, STEP0_MAX_RETRIES + 1):
        print(f"👉 Step 0: Refreshing the pending PO list (attempt {attempt}/{STEP0_MAX_RETRIES})...")
        try:
            sheets_service = _get_sheets_service()
            drive_service = _get_drive_service()
            if not sheets_service or not drive_service:
                raise RuntimeError("Could not create Google API service, please check token.json.")

            _clear_pending_sheet(sheets_service)

            for vessel_name in PENDING_VESSEL_NAMES:
                records = _read_vessel_pending_rows(sheets_service, drive_service, vessel_name)
                if records:
                    _append_pending_rows(sheets_service, vessel_name, records)
                    print(f"✅ {vessel_name}: added {len(records)} pending PO(s).")
                else:
                    print(f"ℹ️ {vessel_name}: no PO matched the criteria.")

            print("✅ Step 0: pending sheet update complete.")
            return True

        except Exception as e:
            print(f"⚠️ Step 0 failed: {e}")

        if attempt < STEP0_MAX_RETRIES:
            print(f"⏳ Waiting {STEP0_RETRY_DELAY_SEC}s before retrying (possibly too many requests)...")
            time.sleep(STEP0_RETRY_DELAY_SEC)

    print(f"❌ Step 0: still failing after {STEP0_MAX_RETRIES} retries, aborting RPA.")
    return False


async def run_d365_rpa(records):
    """ Playwright 執行 D365 自動化 (單筆測試 & Debug 保持開啟模式) """
    if not records:
        print("⚠️ No pending data to process, exiting program.")
        return

    print(f"\n🧪 [Batch mode started] Will process all {len(records)} PO(s) in sequence")

    p = await async_playwright().start()

    print(f"📂 Using persistent browser data directory: {USER_DATA_DIR}")

    # 已加入最大化視窗參數 --start-maximized 與 no_viewport=True
    context = await p.chromium.launch_persistent_context(
        user_data_dir=USER_DATA_DIR,
        headless=False,
        viewport=None,
        no_viewport=True,
        args=["--start-maximized"]
    )

    page = context.pages[0] if context.pages else await context.new_page()

    results = []  # 記錄每筆 PO 的處理結果 (vessel_name / po_number / status)

    try:
        print(f"\nOpening D365: {D365_URL}")
        await page.goto(D365_URL)

        print("\n" + "="*60)
        print("[Note] If you've already logged in before, the session should be remembered automatically.")
        print("If this is your first time or the session has expired, please complete login in Chrome, then return to the Terminal and press Enter to continue!")
        print("="*60 + "\n")

        await asyncio.to_thread(input, "Once you've confirmed you're on the D365 home screen, press Enter to start the RPA: ")

        current_vessel = ""
        current_po = ""
        error_msg = ""

        for idx, item in enumerate(records, 1):
            target_po = item["po_number"]
            target_account_code = item["account_code"]
            target_inv_num = item["invoice_number"]
            target_inv_amt = item["invoice_amount"]
            target_credit_amt = item["credit_amount"]
            target_diff = item["difference"]
            current_vessel = item.get("vessel_name", "").strip()
            current_po = target_po
            error_msg = ""

            print(f"\n--------------------------------------------------")
            print(f"Progress [{idx}/{len(records)}] ➔ Processing PO: {target_po} | Invoice Number: {target_inv_num}")

            # =============================================================
            # 📌 每處理 8 筆 PO，強制重新載入頁面一次，釋放 D365 SPA 累積嘅記憶體/DOM 狀態
            # (短時間內重新載入通常唔需要再做 2FA，所以唔喺度暫停等 Enter)
            # =============================================================
            if idx > 1 and (idx - 1) % RELOAD_EVERY_N_PO == 0:
                print(f"🔄 Processed {idx - 1} PO(s), forcing a page reload to avoid the D365 SPA slowing down and timing out...")
                await page.goto(D365_URL)
                await page.wait_for_load_state("networkidle", timeout=30000)
                await page.wait_for_timeout(2000)
                print("✅ Page reloaded, continuing with the next PO.")

            # =============================================================
            # 📌 Step 1: Click "New pending invoice"
            # =============================================================
            print("👉 Step 1: Clicking 'New pending invoice'...")

            new_invoice_btn = page.get_by_role("button", name="New pending invoice")
            if await new_invoice_btn.count() == 0:
                new_invoice_btn = page.get_by_text("New pending invoice", exact=False)

            await new_invoice_btn.wait_for(state="visible", timeout=30000)
            await new_invoice_btn.click()
            print("✅ Clicked 'New pending invoice', waiting for the page to load...")

            await page.wait_for_timeout(3000)

            # =============================================================
            # 📌 Step 2: 捲動並精準定位 "Purchase order (If applicable)" 輸入欄位
            # =============================================================
            print("👉 Step 2: Looking for the 'Purchase order (If applicable)' input field...")

            print("📜 Scrolling the page down...")
            for _ in range(3):
                await page.mouse.move(600, 400)
                await page.mouse.wheel(0, 600)
                await page.keyboard.press("PageDown")
                await page.wait_for_timeout(500)

            candidates = [
                "xpath=//*[contains(text(), 'Purchase order (If applicable)')]/following::input[1]",
                "xpath=//label[contains(text(), 'Purchase order (If applicable)')]/following::input[1]",
                "input[aria-label*='Purchase order (If applicable)']",
                "xpath=//*[contains(text(), 'Purchase order')]/following::input[1]",
                "input[aria-label*='Purchase order']",
                "input[data-dyn-controlname*='Purch']",
                "input[data-dyn-controlname*='PurchaseOrder']"
            ]

            po_input = None
            for sel in candidates:
                loc = page.locator(sel).first
                if await loc.count() > 0:
                    po_input = loc
                    print(f"🎯 Matched PO locator: {sel}")
                    break

            if not po_input:
                po_input = page.get_by_label("Purchase order (If applicable)", exact=False).first

            print("⏳ Waiting for the PO input field to attach and scrolling it into view...")
            await po_input.wait_for(state="attached", timeout=30000)

            await po_input.evaluate("el => el.scrollIntoView({block: 'center', inline: 'nearest'})")
            await page.wait_for_timeout(1000)

            await po_input.wait_for(state="visible", timeout=10000)
            print("👇 Clicking the input box just below 'Purchase order (If applicable)'...")
            await po_input.click(force=True)
            await page.wait_for_timeout(500)

            await po_input.fill("")
            await po_input.fill(target_po)
            await po_input.press("Tab")
            print(f"✅ Entered PO: {target_po} and pressed Tab to trigger auto-fill...")

            # =============================================================
            # 📌 等待 Vendor Account 自動帶出
            # =============================================================
            print("⏳ Waiting for 'Vendor account' to auto-populate...")

            vendor_candidates = [
                "input[data-dyn-controlname*='OrderAccount']",
                "input[data-dyn-controlname*='InvoiceAccount']",
                "input[aria-label*='Vendor account']",
                "xpath=//*[contains(text(), 'Vendor account')]/following::input[1]"
            ]

            vendor_input = None
            for sel in vendor_candidates:
                loc = page.locator(sel).first
                if await loc.count() > 0:
                    vendor_input = loc
                    print(f"🎯 Matched Vendor Account locator: {sel}")
                    break

            if not vendor_input:
                vendor_input = page.get_by_label("Vendor account", exact=False).first

            for _ in range(15):
                val = await vendor_input.input_value()
                if val and len(val.strip()) > 0:
                    print(f"🎉 Success! Vendor Account auto-populated: {val}")
                    break
                await page.wait_for_timeout(1000)
            else:
                print("⚠️ Note: Vendor account field still empty after waiting 15 seconds.")

            # =============================================================
            # 📌 Step 3: 精準尋找 "Invoice number" 欄位
            # =============================================================
            print("\n👉 Step 3: Precisely locating the 'Invoice number' input field...")

            inv_num_candidates = [
                "xpath=//label[normalize-space(text())='Invoice number']/following::input[1]",
                "xpath=//*[normalize-space(text())='Invoice number']/following::input[1]",
                "input[aria-label='Invoice number']",
                "input[data-dyn-controlname='InvoiceNumber']"
            ]

            inv_input = None
            for sel in inv_num_candidates:
                loc = page.locator(sel).first
                if await loc.count() > 0:
                    inv_input = loc
                    print(f"🎯 Matched Invoice Number locator: {sel}")
                    break

            if not inv_input:
                inv_input = page.get_by_label("Invoice number", exact=True).first

            if inv_input and await inv_input.count() > 0:
                print("⏳ Waiting for the Invoice Number field to attach and scrolling it into view...")
                await inv_input.wait_for(state="attached", timeout=15000)
                await inv_input.evaluate("el => el.scrollIntoView({block: 'center', inline: 'nearest'})")
                await page.wait_for_timeout(500)

                await inv_input.wait_for(state="visible", timeout=10000)
                print(f"👇 Clicking the Invoice Number input box and entering: {target_inv_num}")
                await inv_input.click(force=True)
                await page.wait_for_timeout(300)
                await inv_input.fill("")
                await inv_input.fill(target_inv_num)
                await inv_input.press("Tab")
                print(f"✅ Entered Invoice Number: {target_inv_num} and pressed Tab to leave the field.")

                print("🔎 Checking for an 'already exists' warning, polling for 1 second...")
                already_exists_found = False
                for _ in range(10):
                    if await page.get_by_text("already exists", exact=False).count() > 0:
                        already_exists_found = True
                        break
                    await page.wait_for_timeout(100)

                if already_exists_found:
                    error_msg = "error: inv exist"
                    print("⚠️ Detected 'already exists', proceeding directly to error reporting.")
                    update_vessel_sheet_status(current_vessel, current_po, error_msg)
                    results.append({"vessel": current_vessel, "po": current_po, "status": "error", "error": error_msg})
                    await click_finance_and_operations(page)
                    print("⏭️ Error logged, moving on to the next PO.")
                    continue
            else:
                print("❌ Could not find a dedicated 'Invoice number' input field, please check the D365 layout label names.")

            # =============================================================
            # 📌 Step 4: 精準尋找 "Document source" 欄位
            # =============================================================
            print("\n👉 Step 4: Precisely locating the 'Document source' input field...")

            doc_source_candidates = [
                "xpath=//*[contains(translate(text(), 'DOCUMENT SOURCE', 'document source'), 'document source')]/following::input[1]",
                "xpath=//label[contains(translate(text(), 'DOCUMENT SOURCE', 'document source'), 'document source')]/following::input[1]",
                "input[aria-label*='Document source' i]",
                "input[data-dyn-controlname*='DocumentSource' i]"
            ]

            doc_source_input = None
            for sel in doc_source_candidates:
                loc = page.locator(sel).first
                if await loc.count() > 0:
                    doc_source_input = loc
                    print(f"🎯 Matched Document Source locator: {sel}")
                    break

            if not doc_source_input:
                doc_source_input = page.get_by_label("Document source", exact=False).first

            if doc_source_input and await doc_source_input.count() > 0:
                print("⏳ Waiting for the Document Source field to attach and scrolling it into view...")
                await doc_source_input.wait_for(state="attached", timeout=15000)
                await doc_source_input.evaluate("el => el.scrollIntoView({block: 'center', inline: 'nearest'})")
                await page.wait_for_timeout(500)

                await doc_source_input.wait_for(state="visible", timeout=10000)
                print("👇 Clicking the Document Source input box and entering: d365")
                await doc_source_input.click(force=True)
                await page.wait_for_timeout(300)
                await doc_source_input.fill("")
                await doc_source_input.fill("d365")
                await doc_source_input.press("Tab")
                print("✅ Entered Document Source: d365 and pressed Tab to leave the field.")
            else:
                print("❌ Could not find a dedicated 'Document source' input field, please check the D365 layout label names.")

            # =============================================================
            # 📌 Step 5: 完全精準尋找 "Invoice date" 欄位
            # =============================================================
            print("\n👉 Step 5: Precisely locating the 'Invoice date' input field (excluding 'Invoice dates')...")

            inv_date_candidates = [
                "xpath=//label[normalize-space(translate(text(), 'INVOICE DATE', 'invoice date'))='invoice date']/following::input[1]",
                "xpath=//*[normalize-space(translate(text(), 'INVOICE DATE', 'invoice date'))='invoice date']/following::input[1]",
                "xpath=//*[contains(translate(text(), 'INVOICE DATE', 'invoice date'), 'invoice date') and not(contains(translate(text(), 'INVOICE DATES', 'invoice dates'), 'invoice dates'))]/following::input[1]",
                "input[aria-label='Invoice date' i]",
                "input[data-dyn-controlname*='InvoiceDate' i][not(contains(@aria-label, 'dates'))]"
            ]

            inv_date_input = None
            for sel in inv_date_candidates:
                loc = page.locator(sel).first
                if await loc.count() > 0:
                    inv_date_input = loc
                    print(f"🎯 Precisely matched Invoice Date locator: {sel}")
                    break

            if not inv_date_input or await inv_date_input.count() == 0:
                inv_date_input = page.get_by_label("Invoice date", exact=True).first

            if inv_date_input and await inv_date_input.count() > 0:
                print("⏳ Waiting for the Invoice Date field to attach and scrolling it into view...")
                await inv_date_input.wait_for(state="attached", timeout=15000)
                await inv_date_input.evaluate("el => el.scrollIntoView({block: 'center', inline: 'nearest'})")
                await page.wait_for_timeout(500)

                await inv_date_input.wait_for(state="visible", timeout=10000)
                print("👇 Clicking the Invoice Date (singular, not 'dates') input box...")
                await inv_date_input.click(force=True)
                await page.wait_for_timeout(300)

                print("⌨️ Pressing Alt+Down (Alt+ArrowDown) to open the date picker...")
                await inv_date_input.press("Alt+ArrowDown")
                await page.wait_for_timeout(800)

                print("👇 Looking for and clicking the 'Today' button...")
                today_btn = page.get_by_role("button", name="Today", exact=False)
                if await today_btn.count() == 0:
                    today_btn = page.get_by_text("Today", exact=True)
                if await today_btn.count() == 0:
                    today_btn = page.locator("button:has-text('Today'), [aria-label*='Today' i]").first

                if await today_btn.count() > 0:
                    await today_btn.first.click()
                    print("✅ Clicked 'Today'")
                else:
                    print("⚠️ Could not automatically locate the 'Today' button, trying to send a Tab key instead...")

                await page.keyboard.press("Tab")
                print("✅ Invoice Date setup complete.")
            else:
                print("❌ Could not find a dedicated 'Invoice date' input field, please check the D365 layout label names.")

            # =============================================================
            # 📌 Step 6: 精準尋找 "Rebate type" 欄位
            # =============================================================
            print("\n👉 Step 6: Precisely locating the 'Rebate type' input field...")

            cleaned_credit = target_credit_amt.strip()
            if cleaned_credit and cleaned_credit not in ["0", "0.00", "0.0"]:
                rebate_text = "has rebate"
            else:
                rebate_text = "no rebate"

            print(f"📊 Credit Amount is '{target_credit_amt}' ➔ Rebate type should be: '{rebate_text}'")

            rebate_candidates = [
                "xpath=//label[normalize-space(translate(text(), 'REBATE TYPE', 'rebate type'))='rebate type']/following::input[1]",
                "xpath=//*[normalize-space(translate(text(), 'REBATE TYPE', 'rebate type'))='rebate type']/following::input[1]",
                "xpath=//*[contains(translate(text(), 'REBATE TYPE', 'rebate type'), 'rebate type')]/following::input[1]",
                "input[aria-label*='Rebate type' i]",
                "input[data-dyn-controlname*='RebateType' i]"
            ]

            rebate_input = None
            for sel in rebate_candidates:
                loc = page.locator(sel).first
                if await loc.count() > 0:
                    rebate_input = loc
                    print(f"🎯 Matched Rebate Type locator: {sel}")
                    break

            if not rebate_input or await rebate_input.count() == 0:
                rebate_input = page.get_by_label("Rebate type", exact=False).first

            if rebate_input and await rebate_input.count() > 0:
                print("⏳ Waiting for the Rebate type field to attach and scrolling it into view...")
                await rebate_input.wait_for(state="attached", timeout=15000)
                await rebate_input.evaluate("el => el.scrollIntoView({block: 'center', inline: 'nearest'})")
                await page.wait_for_timeout(500)

                await rebate_input.wait_for(state="visible", timeout=10000)
                print(f"👇 Clicking the Rebate type input box and entering: {rebate_text}")
                await rebate_input.click(force=True)
                await page.wait_for_timeout(300)
                await rebate_input.fill("")
                await rebate_input.fill(rebate_text)
                await rebate_input.press("Tab")
                print(f"✅ Entered Rebate type: {rebate_text} and pressed Tab to leave the field.")
            else:
                print("❌ Could not find a dedicated 'Rebate type' input field, please check the D365 layout label names.")

            # =============================================================
            # 📌 Step 7: 檢查 Invoice total (Original currency) 與比對金額 (Add line 流程)
            # =============================================================
            print("\n👉 Step 7: Reading the 'Invoice total (Original currency)' field...")

            inv_total_candidates = [
                "xpath=//*[contains(translate(text(), 'INVOICE TOTAL (ORIGINAL CURRENCY)', 'invoice total (original currency)'), 'invoice total (original currency)')]/following::input[1]",
                "xpath=//label[contains(translate(text(), 'INVOICE TOTAL (ORIGINAL CURRENCY)', 'invoice total (original currency)'), 'invoice total (original currency)')]/following::input[1]",
                "input[aria-label*='Invoice total (Original currency)' i]",
                "input[data-dyn-controlname*='InvoiceTotal' i]"
            ]

            inv_total_input = None
            for sel in inv_total_candidates:
                loc = page.locator(sel).first
                if await loc.count() > 0:
                    inv_total_input = loc
                    print(f"🎯 Matched Invoice total (Original currency) locator: {sel}")
                    break

            if not inv_total_input or await inv_total_input.count() == 0:
                inv_total_input = page.get_by_label("Invoice total (Original currency)", exact=False).first

            inv_total_val_str = ""
            if inv_total_input and await inv_total_input.count() > 0:
                await inv_total_input.wait_for(state="attached", timeout=10000)
                inv_total_val_str = await inv_total_input.input_value()
                print(f"📖 Read Invoice total (Original currency) from D365 = '{inv_total_val_str}'")
            else:
                print("⚠️ Could not automatically locate the Invoice total (Original currency) field, proceeding with a default comparison.")

            val_total = clean_amount(inv_total_val_str)
            val_target_amt = clean_amount(target_inv_amt)

            print(f"⚖️ Comparing values: D365 ({val_total}) vs Google Sheet ({val_target_amt})")

            # 🎯 條件分流：金額不相等才執行 Add line 流程
            if val_total != val_target_amt:
                print("⚠️ Values don't match! Starting the Add line process...")

                # 📜 執行 Scroll Up & Down 確保表格與 Add line Toolbar 載入
                print("📜 Scrolling up & down to make sure the table and buttons are fully rendered...")
                await page.mouse.wheel(0, -300)
                await page.wait_for_timeout(300)
                await page.mouse.wheel(0, 500)
                await page.wait_for_timeout(500)

                # 1. 嚴格搜尋以 'C' / 'c' 開頭的 Item number 欄位 (強制排除 Budget check 等雜訊)
                print("🔍 Precisely searching for the Item number (must start with 'C' or 'c')...")

                item_num_candidates = [
                    "input[aria-label*='Item number' i]",
                    "input[data-dyn-controlname*='ItemId' i]",
                    "xpath=//label[contains(translate(text(), 'ITEM NUMBER', 'item number'), 'item number')]/following::input",
                    "xpath=//*[contains(translate(text(), 'ITEM NUMBER', 'item number'), 'item number')]/following::input"
                ]

                recorded_item_num = ""

                # 遍歷所有候選元素，精準鎖定以 c/C 開頭的值
                for sel in item_num_candidates:
                    locs = page.locator(sel)
                    count = await locs.count()
                    for i in range(count):
                        loc = locs.nth(i)
                        try:
                            val = await loc.input_value()
                            if not val:
                                val = await loc.get_attribute("value") or ""
                            val = val.strip()
                            if val.lower().startswith('c'):
                                recorded_item_num = val
                                print(f"🎯 Verified and captured an Item number starting with 'C': '{recorded_item_num}'")
                                break
                        except Exception:
                            continue
                    if recorded_item_num:
                        break

                # 備用機制：若 input 框未找到，嘗試從網格 cell / span 中提取以 'C' 開頭的值
                if not recorded_item_num:
                    print("⚠️ Could not extract a 'C'-prefixed value from the input fields, trying to search the data grid cells instead...")
                    grid_cells = page.locator("xpath=//div[contains(@aria-label, 'Item number') or contains(@data-dyn-controlname, 'ItemId')]//input | //div[contains(@aria-label, 'Item number') or contains(@data-dyn-controlname, 'ItemId')]//span")
                    grid_count = await grid_cells.count()
                    for i in range(grid_count):
                        cell = grid_cells.nth(i)
                        try:
                            text_val = ""
                            tag_name = await cell.evaluate("el => el.tagName.toLowerCase()")
                            if tag_name == "input":
                                text_val = await cell.input_value()
                            else:
                                text_val = await cell.inner_text()

                            text_val = text_val.strip()
                            if text_val.lower().startswith('c'):
                                recorded_item_num = text_val
                                print(f"🎯 Captured Item number from the data grid: '{recorded_item_num}'")
                                break
                        except Exception:
                            continue

                if not recorded_item_num:
                    print("⚠️ Could not automatically extract an Item number starting with 'C', please check the D365 screen.")

                # 2. 精準尋找帶有 '+' 圖示/文字的 Add line 按鈕 (排除 Fixed 與 Percent)
                print("👉 Searching for the '+ Add line' button (must include the + symbol, excluding Fixed and Percent)...")

                add_line_candidates = [
                    "xpath=//button[contains(., 'Add line') and not(contains(., 'Fixed')) and not(contains(., 'Percent'))]",
                    "xpath=//button[contains(., 'Add Line') and not(contains(., 'Fixed')) and not(contains(., 'Percent'))]",
                    "xpath=//*[contains(text(), 'Add line') or contains(text(), 'Add Line') or contains(text(), '+ Add line') or contains(text(), '+ Add Line')][not(contains(., 'Fixed')) and not(contains(., 'Percent'))]",
                    "button[aria-label*='Add line' i]:not([aria-label*='Fixed']):not([aria-label*='Percent'])"
                ]

                add_line_btn = None
                for sel in add_line_candidates:
                    loc = page.locator(sel).first
                    if await loc.count() > 0:
                        add_line_btn = loc
                        print(f"🎯 Matched Add line button locator: {sel}")
                        break

                if not add_line_btn or await add_line_btn.count() == 0:
                    add_line_btn = page.get_by_role("button", name="Add line", exact=False).filter(has_not_text="Fixed").filter(has_not_text="Percent").first

                if add_line_btn and await add_line_btn.count() > 0:
                    await add_line_btn.evaluate("el => el.scrollIntoView({block: 'center', inline: 'nearest'})")
                    await page.wait_for_timeout(500)
                    print("👇 Clicking the '+ Add line' button...")
                    await add_line_btn.click(force=True)
                    print("✅ Clicked 'Add line'...")

                    # =============================================================
                    # 📌 新增欄位自動化 (根據 difference 正負數判斷分流)
                    # =============================================================
                    diff_val = clean_amount(target_diff)
                    print(f"\n🔍 Checking Difference value: {diff_val}")

                    if diff_val >= 0:
                        print("▶️ Difference is positive (or 0), running the positive-value fill flow...")

                        await wait_for_new_row_after_add_line(page)

                        if recorded_item_num:
                            print(f"⌨️ Typing the previously captured Item number: {recorded_item_num}")
                            await page.keyboard.type(recorded_item_num, delay=50)
                        else:
                            print("⚠️ Note: no Item number starting with 'C' was captured, skipping Item number entry.")

                        print("⌨️ Pressing Tab 3 times (100ms apart)...")
                        for _ in range(3):
                            await page.keyboard.press("Tab")
                            await page.wait_for_timeout(100)

                        print("⌨️ Typing 'delivery charge'...")
                        await page.keyboard.type("delivery charge", delay=50)

                        print("⌨️ Pressing Tab 8 times (100ms apart)...")
                        for _ in range(8):
                            await page.keyboard.press("Tab")
                            await page.wait_for_timeout(100)

                        print(f"⌨️ Typing Difference value: {target_diff}")
                        await page.keyboard.type(str(target_diff), delay=50)

                        print("💾 Pressing Alt+S to save...")
                        await page.keyboard.press("Alt+s")
                        await page.wait_for_timeout(1000)

                        print("✅ Positive-value flow complete, Alt+S save triggered!")

                    else:
                        print("▶️ Difference is negative, running the negative-value fill flow...")

                        await wait_for_new_row_after_add_line(page)

                        if recorded_item_num:
                            print(f"⌨️ Typing the previously captured Item number: {recorded_item_num}")
                            await page.keyboard.type(recorded_item_num, delay=50)
                        else:
                            print("⚠️ Note: no Item number starting with 'C' was captured, skipping Item number entry.")

                        print("⌨️ Pressing Tab 3 times (100ms apart)...")
                        for _ in range(3):
                            await page.keyboard.press("Tab")
                            await page.wait_for_timeout(100)

                        print("⌨️ Typing 'adjustment'...")
                        await page.keyboard.type("adjustment", delay=50)

                        print("⌨️ Pressing Tab 6 times (100ms apart)...")
                        for _ in range(6):
                            await page.keyboard.press("Tab")
                            await page.wait_for_timeout(100)

                        print("⌨️ Typing '-1'...")
                        await page.keyboard.type("-1", delay=50)

                        print("⌨️ Pressing Tab 2 times (100ms apart)...")
                        for _ in range(2):
                            await page.keyboard.press("Tab")
                            await page.wait_for_timeout(100)

                        abs_diff_str = str(abs(diff_val))
                        print(f"⌨️ Typing the absolute Difference value: {abs_diff_str}")
                        await page.keyboard.type(abs_diff_str, delay=50)

                        print("💾 Pressing Alt+S to save...")
                        await page.keyboard.press("Alt+s")
                        await page.wait_for_timeout(1000)

                        print("✅ Negative-value flow complete, Alt+S save triggered!")

                else:
                    print("❌ Could not precisely locate the '+ Add line' button!")

            else:
                # 🎯 金額相等，跳過 Step 7 直接到達這裡
                print("✅ Amounts match (D365 == Google Sheet), no need to add a line, skipping Step 7!")

            # =============================================================
            # 📌 Step 8: 重新讀取 Invoice total (Original currency) 並比對 invoice_amount
            # =============================================================
            print("\n👉 Step 8: Reading the number in the field just below 'Invoice total (Original currency)'...")
            print("⏳ Waiting 2 seconds for the screen to settle before running the Step 8 comparison...")
            await page.wait_for_timeout(2000)

            inv_total_candidates_step8 = [
                "xpath=//*[contains(translate(text(), 'INVOICE TOTAL (ORIGINAL CURRENCY)', 'invoice total (original currency)'), 'invoice total (original currency)')]/following::input[1]",
                "xpath=//label[contains(translate(text(), 'INVOICE TOTAL (ORIGINAL CURRENCY)', 'invoice total (original currency)'), 'invoice total (original currency)')]/following::input[1]",
                "input[aria-label*='Invoice total (Original currency)' i]",
                "input[data-dyn-controlname*='InvoiceTotal' i]"
            ]

            inv_total_input_step8 = None
            for sel in inv_total_candidates_step8:
                loc = page.locator(sel).first
                if await loc.count() > 0:
                    inv_total_input_step8 = loc
                    print(f"🎯 Step 8: Located the Invoice total (Original currency) field: {sel}")
                    break

            if not inv_total_input_step8 or await inv_total_input_step8.count() == 0:
                inv_total_input_step8 = page.get_by_label("Invoice total (Original currency)", exact=False).first

            step8_total_str = ""
            if inv_total_input_step8 and await inv_total_input_step8.count() > 0:
                await inv_total_input_step8.wait_for(state="attached", timeout=10000)
                step8_total_str = await inv_total_input_step8.input_value()
                print(f"📖 Step 8: Read value = '{step8_total_str}'")
            else:
                print("⚠️ Step 8: Could not locate the field.")

            val_step8_total = clean_amount(step8_total_str)
            val_step8_target = clean_amount(target_inv_amt)

            print(f"⚖️ Step 8 comparison: read value ({val_step8_total}) vs invoice_amount ({val_step8_target})")

            if val_step8_total == val_step8_target:
                print("✅ Values match (equal to invoice_amount), proceeding to Step 9!")
                run_step_9 = True
            else:
                error_msg = "error: amount mismatch"
                print(f"❌ Values don't match ({val_step8_total} != {val_step8_target}), proceeding directly to error reporting.")
                update_vessel_sheet_status(current_vessel, current_po, error_msg)
                results.append({"vessel": current_vessel, "po": current_po, "status": "error", "error": error_msg})
                await click_finance_and_operations(page)
                print("⏭️ Error logged, moving on to the next PO.")
                continue

            # =============================================================
            # 📌 Step 9: Credit Amount 判斷與處理流程 (只有 Step 8 條件成立才執行)
            # =============================================================
            if run_step_9:
                print("\n👉 Step 9: Checking whether Credit Amount exists...")

                credit_amt_val = clean_amount(target_credit_amt)

                if not target_credit_amt or credit_amt_val == 0.0:
                    print("ℹ️ credit_amount is missing or 0, skipping the rest of Step 9.")
                else:
                    print(f"✅ credit_amount detected (value: {target_credit_amt}), starting the Step 9 Add Credit Line process...")

                    # 1. 先判斷 credit_account 的值
                    ac_code = target_account_code.strip()
                    credit_account = get_credit_account(ac_code)

                    print(f"📊 account_code is '{ac_code}' ➔ credit_account should be: '{credit_account}'")

                    # 2. 尋找並點擊 + Add line 按鈕
                    print("🔍 Searching for the '+ Add line' button (excluding Fixed and Percent)...")
                    add_line_candidates = [
                        "xpath=//button[contains(., 'Add line') and not(contains(., 'Fixed')) and not(contains(., 'Percent'))]",
                        "xpath=//button[contains(., 'Add Line') and not(contains(., 'Fixed')) and not(contains(., 'Percent'))]",
                        "xpath=//*[contains(text(), 'Add line') or contains(text(), 'Add Line') or contains(text(), '+ Add line') or contains(text(), '+ Add Line')][not(contains(., 'Fixed')) and not(contains(., 'Percent'))]",
                        "button[aria-label*='Add line' i]:not([aria-label*='Fixed']):not([aria-label*='Percent'])"
                    ]

                    add_line_btn = None
                    for sel in add_line_candidates:
                        loc = page.locator(sel).first
                        if await loc.count() > 0:
                            add_line_btn = loc
                            print(f"🎯 Step 9: Matched Add line button: {sel}")
                            break

                    if not add_line_btn or await add_line_btn.count() == 0:
                        add_line_btn = page.get_by_role("button", name="Add line", exact=False).filter(has_not_text="Fixed").filter(has_not_text="Percent").first

                    if add_line_btn and await add_line_btn.count() > 0:
                        await add_line_btn.evaluate("el => el.scrollIntoView({block: 'center', inline: 'nearest'})")
                        await page.wait_for_timeout(500)
                        print("👇 Clicking the '+ Add line' button...")
                        await add_line_btn.click(force=True)

                        await wait_for_new_row_after_add_line(page)

                        # 3. 輸入 credit_account 並按 3 次 Tab
                        print(f"⌨️ Typing credit_account: {credit_account}")
                        await page.keyboard.type(credit_account, delay=50)

                        print("⌨️ Pressing Tab 3 times (100ms apart)...")
                        for _ in range(3):
                            await page.keyboard.press("Tab")
                            await page.wait_for_timeout(100)

                        # 4. 輸入 discount 並按 6 次 Tab
                        print("⌨️ Typing 'discount'...")
                        await page.keyboard.type("discount", delay=50)

                        print("⌨️ Pressing Tab 6 times (100ms apart)...")
                        for _ in range(6):
                            await page.keyboard.press("Tab")
                            await page.wait_for_timeout(100)

                        # 5. 輸入 -1 並按 2 次 Tab
                        print("⌨️ Typing '-1'...")
                        await page.keyboard.type("-1", delay=50)

                        print("⌨️ Pressing Tab 2 times (100ms apart)...")
                        for _ in range(2):
                            await page.keyboard.press("Tab")
                            await page.wait_for_timeout(100)

                        # 6. 輸入 credit_amount 的絕對值
                        abs_credit_str = str(abs(credit_amt_val))
                        print(f"⌨️ Typing the absolute credit_amount value: {abs_credit_str}")
                        await page.keyboard.type(abs_credit_str, delay=50)

                        # 7. 按 Alt+S 儲存
                        print("💾 Pressing Alt+S to save...")
                        await page.keyboard.press("Alt+s")
                        await page.wait_for_timeout(1000)

                        print("✅ Step 9 complete, Alt+S save triggered!")
                    else:
                        print("❌ Step 9: Could not locate the '+ Add line' button!")

            # =============================================================
            # 📌 Step 10: 從 Google Drive 下載附件並上傳至 D365 Attachments
            # =============================================================
            print("\n👉 Step 10: Downloading the PO's attachments from Google Drive...")

            need_credit_note = bool(target_credit_amt) and clean_amount(target_credit_amt) != 0.0
            downloaded_files = download_po_attachments(target_po, need_credit_note)

            missing_required = "invoice" not in downloaded_files or "receipt" not in downloaded_files
            missing_credit_note = need_credit_note and "credit_note" not in downloaded_files

            if missing_required or missing_credit_note:
                error_msg = "error: missing drive file"
                print("❌ Missing required attachment(s), proceeding directly to error reporting.")
                update_vessel_sheet_status(current_vessel, current_po, error_msg)
                results.append({"vessel": current_vessel, "po": current_po, "status": "error", "error": error_msg})
                await click_finance_and_operations(page)
                print("⏭️ Error logged, moving on to the next PO.")
                continue

            print("👉 Clicking 'Attachments'...")
            attachments_candidates = [
                page.get_by_role("tab", name="Attachments", exact=False),
                page.get_by_role("button", name="Attachments", exact=False),
                page.get_by_text("Attachments", exact=False),
            ]

            attachments_target = None
            for candidate in attachments_candidates:
                if await candidate.count() > 0:
                    attachments_target = candidate
                    break

            if not attachments_target:
                attachments_target = page.get_by_text("Attachments", exact=False)

            await attachments_target.first.wait_for(state="visible", timeout=15000)
            await attachments_target.first.click()
            print("✅ Clicked 'Attachments', waiting 1 second for the screen to start refreshing...")
            await page.wait_for_timeout(1000)
            print("⏳ Waiting for the page to reach networkidle and settle...")
            await page.wait_for_load_state("networkidle", timeout=20000)
            await page.wait_for_timeout(1500)

            upload_plan = [
                ("invoice", "Invoice"),
                ("receipt", "Delivery Note / Service Report"),
            ]
            if need_credit_note:
                upload_plan.append(("credit_note", "Credit Note"))

            for file_key, doc_type_label in upload_plan:
                await upload_attachment_via_new(page, downloaded_files[file_key], doc_type_label)

            print("✅ Step 10 attachment upload complete.")

            # =============================================================
            # 📌 Step 11: 根據 vessel_name 取得對應的 vessel_po_list (spreadsheet id)
            #          並在該 spreadsheet 搜尋 po_number 所在列，將 Column R 更新為 'draft'
            # =============================================================
            print("\n👉 Step 11: Entering the Step 11 process...")

            vessel = item.get("vessel_name", "").strip()
            update_vessel_sheet_status(vessel, target_po, "draft")
            results.append({"vessel": current_vessel, "po": current_po, "status": "success"})

            await click_finance_and_operations(page)
            print("✅ Step 11 complete.")

    except Exception as e:
        error_msg = error_msg or str(e) or "RPA error"
        if current_po and current_vessel:
            print(f"❌ RPA error detected for PO {current_po}: {error_msg}")
            update_vessel_sheet_status(current_vessel, current_po, error_msg)
            results.append({"vessel": current_vessel, "po": current_po, "status": "error", "error": error_msg})
            await click_finance_and_operations(page)
        else:
            print(f"❌ Unexpected error during execution: {e}")

    finally:
        # 📋 列印所有 PO 的處理結果總覽 (vessel_name / status)
        print("\n" + "="*60)
        print("📋 Processing summary:")
        print("="*60)
        if results:
            for i, r in enumerate(results, 1):
                if r["status"] == "success":
                    status_display = "✅ Success"
                else:
                    status_display = f"❌ Error ({r.get('error', '')})"
                print(f"  [{i}] Vessel: {r['vessel']} | PO: {r['po']} | {status_display}")
        else:
            print("  (No records processed)")
        print("="*60 + "\n")

        # 🎯 保持瀏覽器開啟以便進行 Debug
        print("\n" + "="*60)
        print("⏸️ [Debug mode] All POs processed! The browser will stay open for you to inspect and debug.")
        print("When you're done debugging, press Enter in the Terminal to end the script and close the browser...")
        print("="*60 + "\n")
        await asyncio.to_thread(input, "Press Enter to close the program: ")

async def main():
    step0_ok = refresh_pending_invoices()
    if not step0_ok:
        print("❌ Step 0 did not complete successfully, aborting RPA (no up-to-date pending PO data to process).")
        return

    records_data = get_data_via_sheets_api()
    await run_d365_rpa(records_data)


if __name__ == "__main__":
    asyncio.run(main())
