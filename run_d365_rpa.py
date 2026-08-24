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

# 用戶 Browser Profile 儲存路徑（針對 macOS 環境，讓 Session/Cookie 持久化，避免重複 MFA）
USER_DATA_DIR = os.path.abspath("./playwright_user_data")

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
        print(f"⚠️ 未為 vessel_name '{vessel_name}' 找到對應的 vessel_po_list，跳過更新。")
        return False

    try:
        creds = None
        if os.path.exists('token.json'):
            creds = Credentials.from_authorized_user_file('token.json', SCOPES)

        if not creds or not creds.valid:
            if creds and creds.expired and creds.refresh_token:
                creds.refresh(Request())
            else:
                print("⚠️ 無有效憑證可用，請先以 OAuth 取得 token.json。無法更新外部 spreadsheet。")
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
                print(f"✅ 已在 spreadsheet {vessel_po_list} 的 {update_range} 更新為 '{value}' (PO={po_number})")
                found = True
                break

        if not found:
            print(f"⚠️ 在 spreadsheet {vessel_po_list} 中未找到 PO: {po_number}")
            return False

        return True

    except Exception as e:
        print(f"❌ 狀態更新失敗: {e}")
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
        print(f"⚠️ 未能定位到表格 row 元素，改為固定等待 {settle_ms / 1000:.0f} 秒...")
        await page.wait_for_timeout(settle_ms)
        return False

    row_count_before = await page.locator(active_selector).count()
    print(f"⏳ 等待新一行 row 出現 (目前列數: {row_count_before})...")

    elapsed = 0
    while elapsed < max_wait_ms:
        await page.wait_for_timeout(poll_interval_ms)
        elapsed += poll_interval_ms
        row_count_now = await page.locator(active_selector).count()
        if row_count_now > row_count_before:
            print(f"✅ 偵測到新一行 row 出現 (目前列數: {row_count_now})，等待 {settle_ms / 1000:.0f} 秒後繼續...")
            await page.wait_for_timeout(settle_ms)
            return True

    print(f"⚠️ 等待 {max_wait_ms / 1000:.0f} 秒仍未偵測到新一行，改為等待 {settle_ms / 1000:.0f} 秒後繼續...")
    await page.wait_for_timeout(settle_ms)
    return False


async def click_finance_and_operations(page):
    """精準定位並點擊 Finance and Operations，等待頁面刷新回到首頁。"""
    print("👉 尋找並點擊 'Finance and Operations'...")
    finance_locators = [
        page.get_by_role("link", name="Finance and Operations", exact=True),
        page.get_by_role("button", name="Finance and Operations", exact=True),
        page.get_by_text("Finance and Operations", exact=True)
    ]

    finance_target = None
    for locator in finance_locators:
        if await locator.count() > 0:
            finance_target = locator.first
            print("🎯 成功匹配 'Finance and Operations' 元素")
            break

    if finance_target:
        await finance_target.wait_for(state="visible", timeout=15000)
        await finance_target.click()
        print("✅ 已點擊 'Finance and Operations'，等待頁面刷新回到首頁...")
        await page.wait_for_load_state("networkidle", timeout=20000)
        await page.wait_for_timeout(3000)
        print("✅ 頁面已刷新，返回首頁流程完成。")
    else:
        print("⚠️ 未能精準找到 'Finance and Operations'，跳過點擊。")


def _get_drive_service():
    """建立 Google Drive API service，重用與 Sheets 相同的 token.json 憑證。"""
    creds = None
    if os.path.exists('token.json'):
        creds = Credentials.from_authorized_user_file('token.json', SCOPES)

    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            print("⚠️ 無有效憑證可用，請重新以 OAuth 取得包含 Drive 權限的 token.json。")
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
        print(f"⚠️ 於 Google Drive 搜尋到 {len(files)} 個同名檔案 '{file_name}'，只取最新修改的一個 (modifiedTime={files[0].get('modifiedTime')})")
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
        print(f"🔍 正在 Google Drive 搜尋檔案: '{file_name}'...")
        file_id = _find_drive_file_id(service, file_name)
        if not file_id:
            print(f"❌ 未能於 Google Drive 找到檔案: '{file_name}'")
            continue

        dest_path = os.path.join(DRIVE_DOWNLOAD_DIR, f"{file_name}.pdf")
        try:
            request = service.files().get_media(fileId=file_id, supportsAllDrives=True)
            with open(dest_path, 'wb') as f:
                downloader = MediaIoBaseDownload(f, request)
                done = False
                while not done:
                    _, done = downloader.next_chunk()
            print(f"✅ 已下載 '{file_name}' 至 {dest_path}")
            downloaded[key] = dest_path
        except Exception as e:
            print(f"❌ 下載 '{file_name}' 失敗: {e}")

    return downloaded


async def upload_attachment_via_new(page, file_path, doc_type_label):
    """在 Attachments 畫面點擊 New ➔ 選擇附件類型 ➔ Browse 上傳指定檔案。"""
    print(f"👉 正在上傳附件 ({doc_type_label}): {file_path}")

    new_btn_candidates = [
        ("span[id$='_ctrlAdd_label'] (New dropdown 按鈕)", page.locator("span.button-label-dropDown[id$='_ctrlAdd_label']", has_text="New")),
        ("role=button name='+New'", page.get_by_role("button", name="+New", exact=True)),
        ("role=button name='New'", page.get_by_role("button", name="New", exact=True)),
        ("text='+New'", page.get_by_text("+New", exact=True)),
        ("text='New'", page.get_by_text("New", exact=True)),
    ]

    print("🔎 開始逐一檢查 'New' 按鈕候選 selector 的搜尋結果...")
    new_btn = None
    matched_label = None
    for label, candidate in new_btn_candidates:
        count = await candidate.count()
        print(f"   ➜ [{label}] 找到 {count} 個匹配元素")
        if count > 0:
            for i in range(min(count, 5)):
                el = candidate.nth(i)
                try:
                    is_visible = await el.is_visible()
                    tag = await el.evaluate("el => el.tagName")
                    text = (await el.inner_text()).strip()
                    box = await el.bounding_box()
                    print(f"      [{i}] tag={tag} visible={is_visible} text='{text}' box={box}")
                except Exception as e:
                    print(f"      [{i}] 讀取元素資訊失敗: {e}")
            if not new_btn:
                new_btn = candidate
                matched_label = label

    if not new_btn:
        print("⚠️ 四個候選 selector 都搵唔到 'New'，嘗試用更闊的 xpath 搜尋畫面上所有含 'New' 文字的元素...")
        broad_search = page.locator("xpath=//*[normalize-space(text())='New' or normalize-space(text())='+New']")
        broad_count = await broad_search.count()
        print(f"   ➜ [broad xpath] 找到 {broad_count} 個匹配元素")
        for i in range(min(broad_count, 10)):
            el = broad_search.nth(i)
            try:
                is_visible = await el.is_visible()
                tag = await el.evaluate("el => el.tagName")
                cls = await el.get_attribute("class")
                box = await el.bounding_box()
                print(f"      [{i}] tag={tag} visible={is_visible} class='{cls}' box={box}")
            except Exception as e:
                print(f"      [{i}] 讀取元素資訊失敗: {e}")
        new_btn = page.get_by_role("button", name="New", exact=True)
    else:
        print(f"🎯 採用 selector: {matched_label}")

    await new_btn.first.wait_for(state="visible", timeout=15000)
    await new_btn.first.click()
    print("✅ 已點擊 'New'，等待下拉式選單出現...")
    await page.wait_for_timeout(500)

    type_option = page.get_by_role("menuitem", name=doc_type_label, exact=True)
    if await type_option.count() == 0:
        type_option = page.get_by_text(doc_type_label, exact=True)
    await type_option.first.wait_for(state="visible", timeout=10000)
    await type_option.first.click()
    print(f"✅ 已選擇附件類型: {doc_type_label}")
    await page.wait_for_timeout(1000)

    browse_candidates = [
        ("id$='_UploadControlBrowseButton' (實際按鈕)", page.locator("[id$='_UploadControlBrowseButton']")),
        ("role=button name='Browse'", page.get_by_role("button", name="Browse", exact=False)),
        ("text='Browse'", page.get_by_text("Browse", exact=False)),
    ]

    print("🔎 開始逐一檢查 'Browse' 按鈕候選 selector 的搜尋結果...")
    browse_btn = None
    matched_browse_label = None
    for label, candidate in browse_candidates:
        count = await candidate.count()
        print(f"   ➜ [{label}] 找到 {count} 個匹配元素")
        if count > 0 and not browse_btn:
            browse_btn = candidate
            matched_browse_label = label

    if not browse_btn:
        browse_btn = page.get_by_text("Browse", exact=False)
    else:
        print(f"🎯 採用 selector: {matched_browse_label}")

    await browse_btn.first.wait_for(state="visible", timeout=15000)

    async with page.expect_file_chooser() as fc_info:
        await browse_btn.first.click()
    file_chooser = await fc_info.value
    await file_chooser.set_files(file_path)
    print(f"✅ 已選擇檔案: {file_path}")
    await page.wait_for_timeout(1000)

    confirm_candidates = ["OK", "Save", "Attach", "Insert"]
    confirmed = False
    for label in confirm_candidates:
        btn = page.get_by_role("button", name=label, exact=True)
        if await btn.count() > 0 and await btn.first.is_visible():
            await btn.first.click()
            print(f"✅ 已點擊 '{label}' 完成附件上傳。")
            await page.wait_for_timeout(1000)
            confirmed = True
            break

    if not confirmed:
        print("⚠️ 未能自動定位到確認上傳按鈕，請檢查 D365 畫面確認附件已上傳。")


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
                    f"❌ [API 錯誤] 找不到 '{CLIENT_SECRET_FILE}'！\n"
                    "進行純 API Call 必須有此 OAuth 憑證檔。請確保已放入專案根目錄。"
                )
            flow = InstalledAppFlow.from_client_secrets_file(CLIENT_SECRET_FILE, SCOPES)
            creds = flow.run_local_server(port=0)
        
        # 儲存憑證供下次 API 自動呼叫
        with open('token.json', 'w') as token:
            token.write(creds.to_json())

    print("🌐 發起 Google Sheets API 呼叫...")
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

    print(f"✅ [API 成功] 讀取到 {len(records)} 筆待處理資料：")
    for idx, item in enumerate(records, 1):
        print(f"  [{idx}] PO: {item['po_number']} | AccountCode: {item['account_code']} | Inv#: {item['invoice_number']} | Amt: {item['invoice_amount']} | CreditAmt: {item['credit_amount']} | Diff: {item['difference']}")
    
    return records


async def run_d365_rpa(records):
    """ Playwright 執行 D365 自動化 (單筆測試 & Debug 保持開啟模式) """
    if not records:
        print("⚠️ 無待處理資料，程式結束。")
        return

    print(f"\n🧪 [批次處理模式已啟動] 將依序處理所有 {len(records)} 筆 PO")

    p = await async_playwright().start()
    
    print(f"📂 使用持久化瀏覽器資料目錄: {USER_DATA_DIR}")
    
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
        print(f"\n開啟 D365: {D365_URL}")
        await page.goto(D365_URL)

        print("\n" + "="*60)
        print("【提示】如果你已完成登入，系統會自動記住 Session。")
        print("如果首次開啟或 Session 過期，請喺 Chrome 完成登入後返回 Terminal 按 Enter 繼續！")
        print("="*60 + "\n")

        await asyncio.to_thread(input, "確定已進入 D365 主畫面後，請按 Enter 開始執行 RPA: ")

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
            print(f"進度 [{idx}/{len(records)}] ➔ 處理 PO: {target_po} | Invoice Number: {target_inv_num}")

            # =============================================================
            # 📌 Step 1: Click "New pending invoice"
            # =============================================================
            print("👉 Step 1: 點擊 'New pending invoice'...")
            
            new_invoice_btn = page.get_by_role("button", name="New pending invoice")
            if await new_invoice_btn.count() == 0:
                new_invoice_btn = page.get_by_text("New pending invoice", exact=False)

            await new_invoice_btn.wait_for(state="visible", timeout=30000)
            await new_invoice_btn.click()
            print("✅ 已點擊 'New pending invoice'，等待版面載入...")
            
            await page.wait_for_timeout(3000)

            # =============================================================
            # 📌 Step 2: 捲動並精準定位 "Purchase order (If applicable)" 輸入欄位
            # =============================================================
            print("👉 Step 2: 正在尋找 'Purchase order (If applicable)' 輸入欄位...")

            print("📜 執行 Scroll Down 滾動頁面...")
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
                    print(f"🎯 成功匹配 PO 定位器: {sel}")
                    break

            if not po_input:
                po_input = page.get_by_label("Purchase order (If applicable)", exact=False).first

            print("⏳ 等待 PO 輸入欄位掛載並捲動至畫面中央...")
            await po_input.wait_for(state="attached", timeout=30000)

            await po_input.evaluate("el => el.scrollIntoView({block: 'center', inline: 'nearest'})")
            await page.wait_for_timeout(1000)

            await po_input.wait_for(state="visible", timeout=10000)
            print("👇 點擊 Purchase order (If applicable) 正下方的輸入框...")
            await po_input.click(force=True)
            await page.wait_for_timeout(500)

            await po_input.fill("")
            await po_input.fill(target_po)
            await po_input.press("Tab")
            print(f"✅ 已成功輸入 PO: {target_po} 並按 Tab 觸發連動...")

            # =============================================================
            # 📌 等待 Vendor Account 自動帶出
            # =============================================================
            print("⏳ 等待 'Vendor account' 自動帶出文字...")
            
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
                    print(f"🎯 成功匹配 Vendor Account 定位器: {sel}")
                    break

            if not vendor_input:
                vendor_input = page.get_by_label("Vendor account", exact=False).first

            for _ in range(15):
                val = await vendor_input.input_value()
                if val and len(val.strip()) > 0:
                    print(f"🎉 成功！Vendor Account 自動帶出文字: {val}")
                    break
                await page.wait_for_timeout(1000)
            else:
                print("⚠️ 注意: 等待 15 秒後 Vendor account 欄位仍未獲取到文字。")

            # =============================================================
            # 📌 Step 3: 精準尋找 "Invoice number" 欄位
            # =============================================================
            print("\n👉 Step 3: 正在精準尋找 'Invoice number' 輸入欄位...")

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
                    print(f"🎯 成功匹配 Invoice Number 定位器: {sel}")
                    break

            if not inv_input:
                inv_input = page.get_by_label("Invoice number", exact=True).first

            if inv_input and await inv_input.count() > 0:
                print("⏳ 等待 Invoice Number 欄位掛載並捲動至畫面...")
                await inv_input.wait_for(state="attached", timeout=15000)
                await inv_input.evaluate("el => el.scrollIntoView({block: 'center', inline: 'nearest'})")
                await page.wait_for_timeout(500)
                
                await inv_input.wait_for(state="visible", timeout=10000)
                print(f"👇 點擊 Invoice Number 輸入框並填入: {target_inv_num}")
                await inv_input.click(force=True)
                await page.wait_for_timeout(300)
                await inv_input.fill("")
                await inv_input.fill(target_inv_num)
                await inv_input.press("Tab")
                print(f"✅ 已成功輸入 Invoice Number: {target_inv_num} 並按 Tab 離開。")

                print("🔎 開始檢查是否出現 'already exists' 提示，持續搜尋 1 秒...")
                already_exists_found = False
                for _ in range(10):
                    if await page.get_by_text("already exists", exact=False).count() > 0:
                        already_exists_found = True
                        break
                    await page.wait_for_timeout(100)

                if already_exists_found:
                    error_msg = "error: inv exist"
                    print("⚠️ 偵測到 'already exists'，立即進入 error reporting。")
                    update_vessel_sheet_status(current_vessel, current_po, error_msg)
                    results.append({"vessel": current_vessel, "po": current_po, "status": "error", "error": error_msg})
                    await click_finance_and_operations(page)
                    print("⏭️ 已記錄錯誤並繼續下一筆 PO。")
                    continue
            else:
                print("❌ 找不到專門的 'Invoice number' 輸入欄位，請檢查 D365 版面標籤名稱。")

            # =============================================================
            # 📌 Step 4: 精準尋找 "Document source" 欄位
            # =============================================================
            print("\n👉 Step 4: 正在精準尋找 'Document source' 輸入欄位...")

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
                    print(f"🎯 成功匹配 Document Source 定位器: {sel}")
                    break

            if not doc_source_input:
                doc_source_input = page.get_by_label("Document source", exact=False).first

            if doc_source_input and await doc_source_input.count() > 0:
                print("⏳ 等待 Document Source 欄位掛載並捲動至畫面...")
                await doc_source_input.wait_for(state="attached", timeout=15000)
                await doc_source_input.evaluate("el => el.scrollIntoView({block: 'center', inline: 'nearest'})")
                await page.wait_for_timeout(500)
                
                await doc_source_input.wait_for(state="visible", timeout=10000)
                print("👇 點擊 Document Source 輸入框並填入: d365")
                await doc_source_input.click(force=True)
                await page.wait_for_timeout(300)
                await doc_source_input.fill("")
                await doc_source_input.fill("d365")
                await doc_source_input.press("Tab")
                print("✅ 已成功輸入 Document Source: d365 並按 Tab 離開。")
            else:
                print("❌ 找不到專門的 'Document source' 輸入欄位，請檢查 D365 版面標籤名稱。")

            # =============================================================
            # 📌 Step 5: 完全精準尋找 "Invoice date" 欄位
            # =============================================================
            print("\n👉 Step 5: 正在完全精準尋找 'Invoice date' 輸入欄位 (排除 'Invoice dates')...")

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
                    print(f"🎯 成功精準匹配 Invoice Date 定位器: {sel}")
                    break

            if not inv_date_input or await inv_date_input.count() == 0:
                inv_date_input = page.get_by_label("Invoice date", exact=True).first

            if inv_date_input and await inv_date_input.count() > 0:
                print("⏳ 等待 Invoice Date 欄位掛載並捲動至畫面...")
                await inv_date_input.wait_for(state="attached", timeout=15000)
                await inv_date_input.evaluate("el => el.scrollIntoView({block: 'center', inline: 'nearest'})")
                await page.wait_for_timeout(500)
                
                await inv_date_input.wait_for(state="visible", timeout=10000)
                print("👇 點擊 Invoice Date (無 s) 輸入框...")
                await inv_date_input.click(force=True)
                await page.wait_for_timeout(300)

                print("⌨️ 按下 Alt+Down (Alt+ArrowDown) 開啟日期選擇器 (Date Picker)...")
                await inv_date_input.press("Alt+ArrowDown")
                await page.wait_for_timeout(800)

                print("👇 尋找並點擊 'Today' 按鈕...")
                today_btn = page.get_by_role("button", name="Today", exact=False)
                if await today_btn.count() == 0:
                    today_btn = page.get_by_text("Today", exact=True)
                if await today_btn.count() == 0:
                    today_btn = page.locator("button:has-text('Today'), [aria-label*='Today' i]").first

                if await today_btn.count() > 0:
                    await today_btn.first.click()
                    print("✅ 已點擊 'Today'")
                else:
                    print("⚠️ 未能自動定位到 'Today' 按鈕，嘗試發送 Tab 鍵...")

                await page.keyboard.press("Tab")
                print("✅ 已成功完成 Invoice Date 設定。")
            else:
                print("❌ 找不到專門的 'Invoice date' 輸入欄位，請檢查 D365 版面標籤名稱。")

            # =============================================================
            # 📌 Step 6: 精準尋找 "Rebate type" 欄位
            # =============================================================
            print("\n👉 Step 6: 正在精準尋找 'Rebate type' 輸入欄位...")

            cleaned_credit = target_credit_amt.strip()
            if cleaned_credit and cleaned_credit not in ["0", "0.00", "0.0"]:
                rebate_text = "has rebate"
            else:
                rebate_text = "no rebate"

            print(f"📊 Credit Amount 為 '{target_credit_amt}' ➔ 判定 Rebate type 應輸入: '{rebate_text}'")

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
                    print(f"🎯 成功匹配 Rebate Type 定位器: {sel}")
                    break

            if not rebate_input or await rebate_input.count() == 0:
                rebate_input = page.get_by_label("Rebate type", exact=False).first

            if rebate_input and await rebate_input.count() > 0:
                print("⏳ 等待 Rebate type 欄位掛載並捲動至畫面...")
                await rebate_input.wait_for(state="attached", timeout=15000)
                await rebate_input.evaluate("el => el.scrollIntoView({block: 'center', inline: 'nearest'})")
                await page.wait_for_timeout(500)

                await rebate_input.wait_for(state="visible", timeout=10000)
                print(f"👇 點擊 Rebate type 輸入框並填入: {rebate_text}")
                await rebate_input.click(force=True)
                await page.wait_for_timeout(300)
                await rebate_input.fill("")
                await rebate_input.fill(rebate_text)
                await rebate_input.press("Tab")
                print(f"✅ 已成功輸入 Rebate type: {rebate_text} 並按 Tab 離開。")
            else:
                print("❌ 找不到專門的 'Rebate type' 輸入欄位，請檢查 D365 版面標籤名稱。")

            # =============================================================
            # 📌 Step 7: 檢查 Invoice total (Original currency) 與比對金額 (Add line 流程)
            # =============================================================
            print("\n👉 Step 7: 正在讀取 'Invoice total (Original currency)' 欄位...")

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
                    print(f"🎯 成功匹配 Invoice total (Original currency) 定位器: {sel}")
                    break

            if not inv_total_input or await inv_total_input.count() == 0:
                inv_total_input = page.get_by_label("Invoice total (Original currency)", exact=False).first

            inv_total_val_str = ""
            if inv_total_input and await inv_total_input.count() > 0:
                await inv_total_input.wait_for(state="attached", timeout=10000)
                inv_total_val_str = await inv_total_input.input_value()
                print(f"📖 讀取到 D365 上的 Invoice total (Original currency) = '{inv_total_val_str}'")
            else:
                print("⚠️ 未能自動定位到 Invoice total (Original currency) 欄位，將進行預設比較。")

            val_total = clean_amount(inv_total_val_str)
            val_target_amt = clean_amount(target_inv_amt)

            print(f"⚖️ 數值比較: D365 ({val_total}) vs Google Sheet ({val_target_amt})")

            # 🎯 條件分流：金額不相等才執行 Add line 流程
            if val_total != val_target_amt:
                print("⚠️ 兩者不相等！開始執行 Add line 新增列流程...")

                # 📜 執行 Scroll Up & Down 確保表格與 Add line Toolbar 載入
                print("📜 執行 Scroll Up & Down 確保表格與按鈕完全展開...")
                await page.mouse.wheel(0, -300)
                await page.wait_for_timeout(300)
                await page.mouse.wheel(0, 500)
                await page.wait_for_timeout(500)

                # 1. 嚴格搜尋以 'C' / 'c' 開頭的 Item number 欄位 (強制排除 Budget check 等雜訊)
                print("🔍 正在精準搜尋 Item number (必須以 'C' 或 'c' 開頭)...")
                
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
                                print(f"🎯 成功驗證並獲得以 'C' 開頭的 Item number: '{recorded_item_num}'")
                                break
                        except Exception:
                            continue
                    if recorded_item_num:
                        break

                # 備用機制：若 input 框未找到，嘗試從網格 cell / span 中提取以 'C' 開頭的值
                if not recorded_item_num:
                    print("⚠️ 未能在輸入框中提取到 'C' 開頭的值，嘗試搜尋 Data Grid 網格內單元格...")
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
                                print(f"🎯 成功從表格網格中獲得 Item number: '{recorded_item_num}'")
                                break
                        except Exception:
                            continue

                if not recorded_item_num:
                    print("⚠️ 未能自動提取到以 'C' 開頭的 Item number，請檢查 D365 畫面。")

                # 2. 精準尋找帶有 '+' 圖示/文字的 Add line 按鈕 (排除 Fixed 與 Percent)
                print("👉 正在搜尋 '+ Add line' 按鈕 (包含 + 符號，並排除 Fixed 及 Percent)...")
                
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
                        print(f"🎯 成功匹配 Add line 按鈕定位器: {sel}")
                        break

                if not add_line_btn or await add_line_btn.count() == 0:
                    add_line_btn = page.get_by_role("button", name="Add line", exact=False).filter(has_not_text="Fixed").filter(has_not_text="Percent").first

                if add_line_btn and await add_line_btn.count() > 0:
                    await add_line_btn.evaluate("el => el.scrollIntoView({block: 'center', inline: 'nearest'})")
                    await page.wait_for_timeout(500)
                    print("👇 點擊 '+ Add line' 按鈕...")
                    await add_line_btn.click(force=True)
                    print("✅ 已點擊 'Add line'...")

                    # =============================================================
                    # 📌 新增欄位自動化 (根據 difference 正負數判斷分流)
                    # =============================================================
                    diff_val = clean_amount(target_diff)
                    print(f"\n🔍 檢查 Difference 數值: {diff_val}")

                    if diff_val >= 0:
                        print("▶️ Difference 為正數（或 0），執行正數填寫流程...")
                        
                        await wait_for_new_row_after_add_line(page)

                        if recorded_item_num:
                            print(f"⌨️ 輸入剛才複製的 Item number: {recorded_item_num}")
                            await page.keyboard.type(recorded_item_num, delay=50)
                        else:
                            print("⚠️ 注意: 未能獲取到以 'C' 開頭的 Item number，將跳過輸入 Item number。")

                        print("⌨️ 按 3 次 Tab 鍵 (每 100ms 間隔)...")
                        for _ in range(3):
                            await page.keyboard.press("Tab")
                            await page.wait_for_timeout(100)

                        print("⌨️ 輸入 'delivery charge'...")
                        await page.keyboard.type("delivery charge", delay=50)

                        print("⌨️ 按 8 次 Tab 鍵 (每 100ms 間隔)...")
                        for _ in range(8):
                            await page.keyboard.press("Tab")
                            await page.wait_for_timeout(100)

                        print(f"⌨️ 輸入 Difference 數值: {target_diff}")
                        await page.keyboard.type(str(target_diff), delay=50)

                        print("💾 按下 Alt+S 進行儲存...")
                        await page.keyboard.press("Alt+s")
                        await page.wait_for_timeout(1000)
                        
                        print("✅ 正數流程已完成並已觸發 Alt+S 儲存！")

                    else:
                        print("▶️ Difference 為負數，執行負數填寫流程...")

                        await wait_for_new_row_after_add_line(page)

                        if recorded_item_num:
                            print(f"⌨️ 輸入剛才複製的 Item number: {recorded_item_num}")
                            await page.keyboard.type(recorded_item_num, delay=50)
                        else:
                            print("⚠️ 注意: 未能獲取到以 'C' 開頭的 Item number，將跳過輸入 Item number。")

                        print("⌨️ 按 3 次 Tab 鍵 (每 100ms 間隔)...")
                        for _ in range(3):
                            await page.keyboard.press("Tab")
                            await page.wait_for_timeout(100)

                        print("⌨️ 輸入 'adjustment'...")
                        await page.keyboard.type("adjustment", delay=50)

                        print("⌨️ 按 6 次 Tab 鍵 (每 100ms 間隔)...")
                        for _ in range(6):
                            await page.keyboard.press("Tab")
                            await page.wait_for_timeout(100)

                        print("⌨️ 輸入 '-1'...")
                        await page.keyboard.type("-1", delay=50)

                        print("⌨️ 按 2 次 Tab 鍵 (每 100ms 間隔)...")
                        for _ in range(2):
                            await page.keyboard.press("Tab")
                            await page.wait_for_timeout(100)

                        abs_diff_str = str(abs(diff_val))
                        print(f"⌨️ 輸入 Difference 絕對值: {abs_diff_str}")
                        await page.keyboard.type(abs_diff_str, delay=50)

                        print("💾 按下 Alt+S 進行儲存...")
                        await page.keyboard.press("Alt+s")
                        await page.wait_for_timeout(1000)

                        print("✅ 負數流程已完成並已觸發 Alt+S 儲存！")

                else:
                    print("❌ 未能精準定位到 '+ Add line' 按鈕！")

            else:
                # 🎯 金額相等，跳過 Step 7 直接到達這裡
                print("✅ 金額相等 (D365 == Google Sheet)，無需新增列，直接跳過 Step 7！")

            # =============================================================
            # 📌 Step 8: 重新讀取 Invoice total (Original currency) 並比對 invoice_amount
            # =============================================================
            print("\n👉 Step 8: 正在讀取 'Invoice total (Original currency)' 正下方的欄位數字...")
            print("⏳ 先等待 2 秒，讓畫面穩定後再執行 Step 8 比對...")
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
                    print(f"🎯 Step 8: 成功定位 Invoice total (Original currency) 欄位: {sel}")
                    break

            if not inv_total_input_step8 or await inv_total_input_step8.count() == 0:
                inv_total_input_step8 = page.get_by_label("Invoice total (Original currency)", exact=False).first

            step8_total_str = ""
            if inv_total_input_step8 and await inv_total_input_step8.count() > 0:
                await inv_total_input_step8.wait_for(state="attached", timeout=10000)
                step8_total_str = await inv_total_input_step8.input_value()
                print(f"📖 Step 8: 讀取到的數字為 = '{step8_total_str}'")
            else:
                print("⚠️ Step 8: 未能定位到欄位。")

            val_step8_total = clean_amount(step8_total_str)
            val_step8_target = clean_amount(target_inv_amt)

            print(f"⚖️ Step 8 比對: 讀取到的數字 ({val_step8_total}) vs invoice_amount ({val_step8_target})")

            if val_step8_total == val_step8_target:
                print("✅ 數字相符 (等於 invoice_amount)，將執行 Step 9！")
                run_step_9 = True
            else:
                error_msg = "error: amount mismatch"
                print(f"❌ 數字不相符 ({val_step8_total} != {val_step8_target})，立即進入 error reporting。")
                update_vessel_sheet_status(current_vessel, current_po, error_msg)
                results.append({"vessel": current_vessel, "po": current_po, "status": "error", "error": error_msg})
                await click_finance_and_operations(page)
                print("⏭️ 已記錄錯誤並繼續下一筆 PO。")
                continue

            # =============================================================
            # 📌 Step 9: Credit Amount 判斷與處理流程 (只有 Step 8 條件成立才執行)
            # =============================================================
            if run_step_9:
                print("\n👉 Step 9: 判斷 Credit Amount 是否存在...")

                credit_amt_val = clean_amount(target_credit_amt)

                if not target_credit_amt or credit_amt_val == 0.0:
                    print("ℹ️ credit_amount 不存在或為 0，跳過 Step 9 下半段處理。")
                else:
                    print(f"✅ 偵測到 credit_amount 存在 (值為: {target_credit_amt})，開始執行 Step 9 新增 Credit Line 流程...")

                    # 1. 先判斷 credit_account 的值
                    ac_code = target_account_code.strip()
                    credit_account = get_credit_account(ac_code)

                    print(f"📊 account_code 為 '{ac_code}' ➔ 判定 credit_account 應輸入: '{credit_account}'")

                    # 2. 尋找並點擊 + Add line 按鈕
                    print("🔍 正在尋找 '+ Add line' 按鈕 (排除 Fixed 及 Percent)...")
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
                            print(f"🎯 Step 9: 成功匹配 Add line 按鈕: {sel}")
                            break

                    if not add_line_btn or await add_line_btn.count() == 0:
                        add_line_btn = page.get_by_role("button", name="Add line", exact=False).filter(has_not_text="Fixed").filter(has_not_text="Percent").first

                    if add_line_btn and await add_line_btn.count() > 0:
                        await add_line_btn.evaluate("el => el.scrollIntoView({block: 'center', inline: 'nearest'})")
                        await page.wait_for_timeout(500)
                        print("👇 點擊 '+ Add line' 按鈕...")
                        await add_line_btn.click(force=True)

                        await wait_for_new_row_after_add_line(page)

                        # 3. 輸入 credit_account 並按 3 次 Tab
                        print(f"⌨️ 輸入 credit_account: {credit_account}")
                        await page.keyboard.type(credit_account, delay=50)

                        print("⌨️ 按 3 次 Tab 鍵 (每 100ms 間隔)...")
                        for _ in range(3):
                            await page.keyboard.press("Tab")
                            await page.wait_for_timeout(100)

                        # 4. 輸入 discount 並按 6 次 Tab
                        print("⌨️ 輸入 'discount'...")
                        await page.keyboard.type("discount", delay=50)

                        print("⌨️ 按 6 次 Tab 鍵 (每 100ms 間隔)...")
                        for _ in range(6):
                            await page.keyboard.press("Tab")
                            await page.wait_for_timeout(100)

                        # 5. 輸入 -1 並按 2 次 Tab
                        print("⌨️ 輸入 '-1'...")
                        await page.keyboard.type("-1", delay=50)

                        print("⌨️ 按 2 次 Tab 鍵 (每 100ms 間隔)...")
                        for _ in range(2):
                            await page.keyboard.press("Tab")
                            await page.wait_for_timeout(100)

                        # 6. 輸入 credit_amount 的絕對值
                        abs_credit_str = str(abs(credit_amt_val))
                        print(f"⌨️ 輸入 credit_amount 絕對值: {abs_credit_str}")
                        await page.keyboard.type(abs_credit_str, delay=50)

                        # 7. 按 Alt+S 儲存
                        print("💾 按下 Alt+S 進行儲存...")
                        await page.keyboard.press("Alt+s")
                        await page.wait_for_timeout(1000)

                        print("✅ Step 9 處理完成並已觸發 Alt+S 儲存！")
                    else:
                        print("❌ Step 9: 未能定位到 '+ Add line' 按鈕！")

            # =============================================================
            # 📌 Step 10: 從 Google Drive 下載附件並上傳至 D365 Attachments
            # =============================================================
            print("\n👉 Step 10: 正在從 Google Drive 下載 PO 對應附件...")

            need_credit_note = bool(target_credit_amt) and clean_amount(target_credit_amt) != 0.0
            downloaded_files = download_po_attachments(target_po, need_credit_note)

            missing_required = "invoice" not in downloaded_files or "receipt" not in downloaded_files
            missing_credit_note = need_credit_note and "credit_note" not in downloaded_files

            if missing_required or missing_credit_note:
                error_msg = "error: missing drive file"
                print("❌ 缺少必要附件，立即進入 error reporting。")
                update_vessel_sheet_status(current_vessel, current_po, error_msg)
                results.append({"vessel": current_vessel, "po": current_po, "status": "error", "error": error_msg})
                await click_finance_and_operations(page)
                print("⏭️ 已記錄錯誤並繼續下一筆 PO。")
                continue

            print("👉 正在點擊 'Attachments'...")
            attachments_candidates = [
                ("role=tab name='Attachments'", page.get_by_role("tab", name="Attachments", exact=False)),
                ("role=button name='Attachments'", page.get_by_role("button", name="Attachments", exact=False)),
                ("text='Attachments'", page.get_by_text("Attachments", exact=False)),
            ]

            attachments_target = None
            matched_attachments_label = None
            for label, candidate in attachments_candidates:
                count = await candidate.count()
                print(f"   ➜ [{label}] 找到 {count} 個匹配元素")
                if count > 0 and not attachments_target:
                    attachments_target = candidate
                    matched_attachments_label = label

            if not attachments_target:
                attachments_target = page.get_by_text("Attachments", exact=False)
            else:
                print(f"🎯 採用 selector: {matched_attachments_label}")

            await attachments_target.first.wait_for(state="visible", timeout=15000)
            await attachments_target.first.click()
            print("✅ 已點擊 'Attachments'，先等待 1 秒讓畫面開始 refresh...")
            await page.wait_for_timeout(1000)
            print("⏳ 再等待畫面 networkidle 及穩定...")
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

            print("✅ Step 10 附件上傳流程完成。")

            # =============================================================
            # 📌 Step 11: 根據 vessel_name 取得對應的 vessel_po_list (spreadsheet id)
            #          並在該 spreadsheet 搜尋 po_number 所在列，將 Column R 更新為 'draft'
            # =============================================================
            print("\n👉 Step 11: 進入 Step 11 處理流程...")

            vessel = item.get("vessel_name", "").strip()
            update_vessel_sheet_status(vessel, target_po, "draft")
            results.append({"vessel": current_vessel, "po": current_po, "status": "success"})

            await click_finance_and_operations(page)
            print("✅ Step 11 步驟處理完成。")

    except Exception as e:
        error_msg = error_msg or str(e) or "RPA error"
        if current_po and current_vessel:
            print(f"❌ RPA error detected for PO {current_po}: {error_msg}")
            update_vessel_sheet_status(current_vessel, current_po, error_msg)
            results.append({"vessel": current_vessel, "po": current_po, "status": "error", "error": error_msg})
            await click_finance_and_operations(page)
        else:
            print(f"❌ 執行中發生未預期錯誤: {e}")

    finally:
        # 📋 列印所有 PO 的處理結果總覽 (vessel_name / status)
        print("\n" + "="*60)
        print("📋 處理結果總覽:")
        print("="*60)
        if results:
            for i, r in enumerate(results, 1):
                if r["status"] == "success":
                    status_display = "✅ Success"
                else:
                    status_display = f"❌ Error ({r.get('error', '')})"
                print(f"  [{i}] Vessel: {r['vessel']} | PO: {r['po']} | {status_display}")
        else:
            print("  (無任何處理紀錄)")
        print("="*60 + "\n")

        # 🎯 保持瀏覽器開啟以便進行 Debug
        print("\n" + "="*60)
        print("⏸️ [Debug 模式] 全部 PO 處理完成！瀏覽器將保持開啟狀態供您檢查與 Debug。")
        print("當您完成 Debug 後，請在 Terminal 按 Enter 鍵結束腳本並關閉瀏覽器...")
        print("="*60 + "\n")
        await asyncio.to_thread(input, "按下 Enter 鍵關閉程式: ")

if __name__ == "__main__":
    records_data = get_data_via_sheets_api()
    asyncio.run(run_d365_rpa(records_data))