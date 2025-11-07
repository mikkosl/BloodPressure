// BloodPressure.cpp : Defines the entry point for the application.
//

#include "framework.h"
#include "BloodPressure.h"
#include "Database.h"

#include <commdlg.h>
#include <commctrl.h>
#pragma comment(lib, "Comdlg32.lib")
#pragma comment(lib, "Comctl32.lib")

#include <memory>
#include <string>
#include <vector>
#include <ctime>
#include <windowsx.h>
#include <windows.h>
#include <objbase.h> 
#pragma comment(lib, "Ole32.lib")

#ifndef DateTime_SetFormatW
#define DateTime_SetFormatW(hwndDP, pszFormat) \
    (void)SendMessageW((hwndDP), DTM_SETFORMATW, 0, (LPARAM)(LPCWSTR)(pszFormat))
#endif

// Replace your ListView_SetItemTextW macro with this safer wrapper:
#ifndef ListView_SetItemTextW
static inline void LV_SetItemTextW(HWND hwndLV, int i, int iSub, const wchar_t* text)
{
    LVITEMW lvi{};
    lvi.iSubItem = iSub;
    lvi.pszText = const_cast<LPWSTR>(text);
    SendMessageW(hwndLV, LVM_SETITEMTEXTW, (WPARAM)i, (LPARAM)&lvi);
}
#define ListView_SetItemTextW(hwndLV, i, iSub, pszText) LV_SetItemTextW((hwndLV), (i), (iSub), (pszText))
#endif

// Add this macro near the top of your file, after including <commctrl.h>:
#ifndef ListView_InsertItemW
#define ListView_InsertItemW(hwndLV, pitem) \
    (int)SendMessageW((hwndLV), LVM_INSERTITEMW, 0, (LPARAM)(const LVITEMW *)(pitem))
#endif

#define MAX_LOADSTRING 100

// Control IDs for dialog-like window
#define IDC_EDIT_SYSTOLIC  41001
#define IDC_EDIT_DIASTOLIC 41002
#define IDC_EDIT_PULSE     41003
#define IDC_EDIT_NOTE      41004
#define IDC_EDIT_ROWCOMBO  41005
#define IDC_BTN_DELETE     41006
#define IDM_PAGE_PREV      40005
#define IDM_PAGE_NEXT      40006
#define IDC_DATES_START    43010
#define IDC_DATES_END      43011
#define IDC_REPORTALL_SAVE   43012
#define IDC_REPORTDATES_SAVE 43013
#define IDC_REPORTALL_PRINT    43014
#define IDC_REPORTDATES_PRINT  43015

// Global Variables:
HINSTANCE hInst;                                // current instance
WCHAR szTitle[MAX_LOADSTRING];                  // The title bar text
WCHAR szWindowClass[MAX_LOADSTRING];            // the main window class name

static std::unique_ptr<Database> g_db;
static HWND g_mainWnd = nullptr;                // new: remember main window

static int g_pageIndex = 0;
static constexpr int kPageSize = 20;

// Forward declarations:
ATOM                MyRegisterClass(HINSTANCE hInstance);
BOOL                InitInstance(HINSTANCE, int);
LRESULT CALLBACK    WndProc(HWND, UINT, WPARAM, LPARAM);
INT_PTR CALLBACK    About(HWND, UINT, WPARAM, LPARAM);
static void         CreateDatabaseDialog(HWND owner);
static void         OpenDatabaseDialog(HWND owner);
static void         CloseDatabaseDialog(HWND owner);
static void         ShowReportAllWindow(HWND owner); // <-- add
static void         ShowReportDatesWindow(HWND owner);
static LRESULT CALLBACK ReportDatesWndProc(HWND, UINT, WPARAM, LPARAM);
static LRESULT CALLBACK AddReadingWndProc(HWND hWnd, UINT msg, WPARAM wParam, LPARAM lParam);
static void ShowAddReadingDialog(HWND owner);
static void ShowEditReadingDialog(HWND owner, const Reading& r);

static int CompareSystemTimes(const SYSTEMTIME& a, const SYSTEMTIME& b)
{
    FILETIME fa{}, fb{};
    SystemTimeToFileTime(&a, &fa);
    SystemTimeToFileTime(&b, &fb);
    return CompareFileTime(&fa, &fb); // <0 if a<b, 0 if equal, >0 if a>b
}

// Add near other helpers (after CompareSystemTimes)
static std::wstring FormatYmd(const SYSTEMTIME& st)
{
    wchar_t buf[16];
    swprintf_s(buf, L"%04u-%02u-%02u", st.wYear, st.wMonth, st.wDay);
    return buf;
}

// Helpers to compute local time from UTC ISO and averages
static bool TryParseUtcIsoToLocalTm(const std::wstring& isoUtc, std::tm& outLocal)
{
    int Y = 0, M = 0, D = 0, h = 0, m = 0, s = 0;
    if (swscanf_s(isoUtc.c_str(), L"%d-%d-%dT%d:%d:%d", &Y, &M, &D, &h, &m, &s) != 6)
        return false;

    std::tm tmUtc{};
    tmUtc.tm_year = Y - 1900;
    tmUtc.tm_mon = M - 1;
    tmUtc.tm_mday = D;
    tmUtc.tm_hour = h;
    tmUtc.tm_min = m;
    tmUtc.tm_sec = s;

    time_t t = _mkgmtime(&tmUtc);
    if (t == (time_t)-1) return false;

    std::tm tmLocal{};
    if (localtime_s(&tmLocal, &t) != 0) return false;
    outLocal = tmLocal;
    return true;
}

static std::wstring LocalDateYmdFromUtcIso(const std::wstring& isoUtc)
{
    std::tm local{};
    if (!TryParseUtcIsoToLocalTm(isoUtc, local)) return L"";
    wchar_t buf[16]{};
    if (wcsftime(buf, _countof(buf), L"%Y-%m-%d", &local) == 0) return L"";
    return buf;
}

// Add near other helpers (after FormatYmd / LocalDateYmdFromUtcIso)
static std::wstring BuildReportTitle(HWND owner, const wchar_t* fallback)
{
    // If this is the By Dates window, pull dates from its pickers.
    HWND hStart = GetDlgItem(owner, IDC_DATES_START);
    HWND hEnd = GetDlgItem(owner, IDC_DATES_END);
    if (hStart && hEnd) {
        SYSTEMTIME st{}, en{};
        if (DateTime_GetSystemtime(hStart, &st) == GDT_VALID &&
            DateTime_GetSystemtime(hEnd, &en) == GDT_VALID)
        {
            if (CompareSystemTimes(st, en) > 0) {
                // normalize
                en = st;
            }
            return L"By Dates: " + FormatYmd(st) + L" — " + FormatYmd(en);
        }
    }

    // Otherwise, this is the Averages window: use oldest/newest reading dates if possible.
    if (g_db) {
        std::vector<Reading> readings;
        if (g_db->GetAllReadings(readings) && !readings.empty()) {
            const std::wstring newest = LocalDateYmdFromUtcIso(readings.front().tsUtc);
            const std::wstring oldest = LocalDateYmdFromUtcIso(readings.back().tsUtc);
            if (!newest.empty() && !oldest.empty()) {
                return L"Averages: " + oldest + L" — " + newest;
            }
        }
    }

    return fallback ? fallback : L"Report";
}

static bool WriteUtf8File(const wchar_t* path, const std::wstring& text)
{
    int needed = WideCharToMultiByte(CP_UTF8, 0, text.c_str(), (int)text.size(), nullptr, 0, nullptr, nullptr);
    if (needed < 0) return false;

    std::vector<char> bytes;
    bytes.reserve(3 + needed);
    // UTF-8 BOM
    bytes.push_back('\xEF');
    bytes.push_back('\xBB');
    bytes.push_back('\xBF');
    bytes.resize(3 + needed);

    WideCharToMultiByte(CP_UTF8, 0, text.c_str(), (int)text.size(), bytes.data() + 3, needed, nullptr, nullptr);

    HANDLE h = CreateFileW(path, GENERIC_WRITE, 0, nullptr, CREATE_ALWAYS, FILE_ATTRIBUTE_NORMAL, nullptr);
    if (h == INVALID_HANDLE_VALUE) return false;

    DWORD written = 0;
    BOOL ok = WriteFile(h, bytes.data(), (DWORD)bytes.size(), &written, nullptr);
    CloseHandle(h);
    return ok && written == bytes.size();
}

// Locale-aware CSV list separator (Excel honors this)
static wchar_t GetCsvSeparator()
{
    wchar_t sep = L',';
    wchar_t ls[8] = L"";
    int n = GetLocaleInfoW(LOCALE_USER_DEFAULT, LOCALE_SLIST, ls, (int)_countof(ls));
    if (n > 1 && ls[0]) sep = ls[0];
    return sep;
}

// Escape per RFC 4180: wrap if contains sep, quotes, or CR/LF; double the quotes.
static std::wstring CsvEscape(const std::wstring& v, wchar_t sep)
{
    if (v.find(sep) != std::wstring::npos ||
        v.find(L'"') != std::wstring::npos ||
        v.find(L'\r') != std::wstring::npos ||
        v.find(L'\n') != std::wstring::npos)
    {
        std::wstring out;
        out.reserve(v.size() + 2);
        out.push_back(L'"');
        for (wchar_t ch : v) {
            if (ch == L'"') out.push_back(L'"');
            out.push_back(ch);
        }
        out.push_back(L'"');
        return out;
    }
    return v;
}

// In SaveListViewAsCsv(...), prepend a title line before the header
static bool SaveListViewAsCsv(HWND owner, HWND hList, const wchar_t* suggestedName, const wchar_t* dlgTitle)
{
    if (!IsWindow(hList)) {
        MessageBoxW(owner, L"Report view is not available.", szTitle, MB_OK | MB_ICONERROR);
        return false;
    }

    wchar_t sep = GetCsvSeparator();
    std::wstring output;

    // Title
    const std::wstring title = BuildReportTitle(owner, L"Report");
    output.append(CsvEscape(title, sep));
    output.push_back(L'\n');

    HWND hHeader = (HWND)SendMessageW(hList, LVM_GETHEADER, 0, 0);
    int colCount = hHeader ? (int)SendMessageW(hHeader, HDM_GETITEMCOUNT, 0, 0) : 0;
    if (colCount <= 0) colCount = 1;

    // Header
    for (int c = 0; c < colCount; ++c) {
        HDITEMW hd{}; hd.mask = HDI_TEXT;
        wchar_t hbuf[128] = L"";
        hd.pszText = hbuf; hd.cchTextMax = (int)_countof(hbuf);
        if (hHeader) SendMessageW(hHeader, HDM_GETITEMW, (WPARAM)c, (LPARAM)&hd);
        output.append(CsvEscape(hbuf, sep));
        output.push_back(c + 1 < colCount ? sep : L'\n');
    }

    // Rows
    int rowCount = (int)SendMessageW(hList, LVM_GETITEMCOUNT, 0, 0);
    for (int r = 0; r < rowCount; ++r) {
        for (int c = 0; c < colCount; ++c) {
            wchar_t buf[512] = L"";
            LVITEMW it{}; it.iSubItem = c; it.pszText = buf; it.cchTextMax = (int)_countof(buf);
            SendMessageW(hList, LVM_GETITEMTEXTW, (WPARAM)r, (LPARAM)&it);
            output.append(CsvEscape(buf, sep));
            output.push_back(c + 1 < colCount ? sep : L'\n');
        }
    }

    // Save File dialog (unchanged) ...
    wchar_t file[MAX_PATH] = L"";
    wcsncpy_s(file, suggestedName, _TRUNCATE);

    OPENFILENAMEW ofn{}; ofn.lStructSize = sizeof(ofn);
    ofn.hwndOwner = owner;
    ofn.lpstrFilter =
        L"CSV (Comma/Locale separated) (*.csv)\0*.csv\0"
        L"All Files (*.*)\0*.*\0";
    ofn.lpstrFile = file;
    ofn.nMaxFile = (DWORD)_countof(file);
    ofn.lpstrTitle = dlgTitle;
    ofn.Flags = OFN_OVERWRITEPROMPT | OFN_PATHMUSTEXIST;
    ofn.lpstrDefExt = L"csv";

    if (!GetSaveFileNameW(&ofn)) return false; // cancelled

    if (!WriteUtf8File(file, output)) {
        MessageBoxW(owner, L"Failed to write the file.", szTitle, MB_OK | MB_ICONERROR);
        return false;
    }

    MessageBoxW(owner, L"Report saved.", szTitle, MB_OK | MB_ICONINFORMATION);
    return true;
}

// In SaveListViewAsText(...), prepend a title line before the header
static bool SaveListViewAsText(HWND owner, HWND hList, const wchar_t* suggestedName, const wchar_t* dlgTitle)
{
    if (!IsWindow(hList)) {
        MessageBoxW(owner, L"Report view is not available.", szTitle, MB_OK | MB_ICONERROR);
        return false;
    }

    // Build text from ListView: title + header + rows (tab-separated)
    std::wstring output;

    // Title
    output.append(BuildReportTitle(owner, L"Report"));
    output.append(L"\r\n");

    HWND hHeader = (HWND)SendMessageW(hList, LVM_GETHEADER, 0, 0);
    int colCount = hHeader ? (int)SendMessageW(hHeader, HDM_GETITEMCOUNT, 0, 0) : 0;
    if (colCount <= 0) colCount = 1;

    // Header line
    for (int c = 0; c < colCount; ++c) {
        HDITEMW hd{}; hd.mask = HDI_TEXT;
        wchar_t hbuf[128] = L"";
        hd.pszText = hbuf; hd.cchTextMax = (int)_countof(hbuf);
        if (hHeader) SendMessageW(hHeader, HDM_GETITEMW, (WPARAM)c, (LPARAM)&hd);
        output.append(hbuf);
        output.append(c + 1 < colCount ? L"\t" : L"\r\n");
    }

    // Rows
    int rowCount = (int)SendMessageW(hList, LVM_GETITEMCOUNT, 0, 0);
    for (int r = 0; r < rowCount; ++r) {
        for (int c = 0; c < colCount; ++c) {
            wchar_t buf[512] = L"";
            LVITEMW it{}; it.iSubItem = c; it.pszText = buf; it.cchTextMax = (int)_countof(buf);
            SendMessageW(hList, LVM_GETITEMTEXTW, (WPARAM)r, (LPARAM)&it);
            output.append(buf);
            output.append(c + 1 < colCount ? L"\t" : L"\r\n");
        }
    }

    // Save dialog/write (unchanged) ...
    wchar_t file[MAX_PATH] = L"";
    wcsncpy_s(file, suggestedName, _TRUNCATE);

    OPENFILENAMEW ofn{}; ofn.lStructSize = sizeof(ofn);
    ofn.hwndOwner = owner;
    ofn.lpstrFilter =
        L"Tab-separated (*.tsv)\0*.tsv\0"
        L"Text (*.txt)\0*.txt\0"
        L"All Files (*.*)\0*.*\0";
    ofn.lpstrFile = file;
    ofn.nMaxFile = (DWORD)_countof(file);
    ofn.lpstrTitle = dlgTitle;
    ofn.Flags = OFN_OVERWRITEPROMPT | OFN_PATHMUSTEXIST;
    ofn.lpstrDefExt = L"tsv";

    if (!GetSaveFileNameW(&ofn)) return false;

    if (!WriteUtf8File(file, output)) {
        MessageBoxW(owner, L"Failed to write the file.", szTitle, MB_OK | MB_ICONERROR);
        return false;
    }

    MessageBoxW(owner, L"Report saved.", szTitle, MB_OK | MB_ICONINFORMATION);
    return true;
}

// In CopyListViewAsTsvToClipboard(...), prepend a title line before the header
static bool CopyListViewAsTsvToClipboard(HWND owner, HWND hList)
{
    if (!IsWindow(hList)) return false;

    std::wstring output;

    // Title
    output.append(BuildReportTitle(owner, L"Report"));
    output.append(L"\r\n");

    HWND hHeader = (HWND)SendMessageW(hList, LVM_GETHEADER, 0, 0);
    int colCount = hHeader ? (int)SendMessageW(hHeader, HDM_GETITEMCOUNT, 0, 0) : 0;
    if (colCount <= 0) colCount = 1;

    // Header
    for (int c = 0; c < colCount; ++c) {
        HDITEMW hd{}; hd.mask = HDI_TEXT;
        wchar_t hbuf[128] = L"";
        hd.pszText = hbuf; hd.cchTextMax = (int)_countof(hbuf);
        if (hHeader) SendMessageW(hHeader, HDM_GETITEMW, (WPARAM)c, (LPARAM)&hd);
        output.append(hbuf);
        output.append(c + 1 < colCount ? L"\t" : L"\r\n");
    }
    // Rows
    int rowCount = (int)SendMessageW(hList, LVM_GETITEMCOUNT, 0, 0);
    for (int r = 0; r < rowCount; ++r) {
        for (int c = 0; c < colCount; ++c) {
            wchar_t buf[512] = L"";
            LVITEMW it{}; it.iSubItem = c; it.pszText = buf; it.cchTextMax = (int)_countof(buf);
            SendMessageW(hList, LVM_GETITEMTEXTW, (WPARAM)r, (LPARAM)&it);
            output.append(buf);
            output.append(c + 1 < colCount ? L"\t" : L"\r\n");
        }
    }

    const SIZE_T bytes = (output.size() + 1) * sizeof(wchar_t);
    HGLOBAL hMem = GlobalAlloc(GMEM_MOVEABLE, bytes);
    if (!hMem) return false;

    void* p = GlobalLock(hMem);
    if (!p) { GlobalFree(hMem); return false; }
    memcpy(p, output.c_str(), bytes);
    GlobalUnlock(hMem);

    if (!OpenClipboard(owner)) { GlobalFree(hMem); return false; }
    EmptyClipboard();
    SetClipboardData(CF_UNICODETEXT, hMem);
    CloseClipboard();
    return true;
}

// Add this helper near the other helpers (below Save/Copy helpers)
static bool PrintListView(HWND owner, HWND hList, const wchar_t* docTitle)
{
    if (!IsWindow(hList)) {
        MessageBoxW(owner, L"Report view is not available.", szTitle, MB_OK | MB_ICONERROR);
        return false;
    }

    // 1) Show printer selection dialog and get a printer DC
    PRINTDLGW pd{};
    pd.lStructSize = sizeof(pd);
    pd.Flags = PD_RETURNDC | PD_NOSELECTION | PD_NOPAGENUMS | PD_USEDEVMODECOPIESANDCOLLATE;
    pd.hwndOwner = owner;

    if (!PrintDlgW(&pd)) {
        // canceled or failed
        return false;
    }

    HDC hdc = pd.hDC;
    if (!hdc) {
        if (pd.hDevMode) GlobalFree(pd.hDevMode);
        if (pd.hDevNames) GlobalFree(pd.hDevNames);
        MessageBoxW(owner, L"Failed to get printer device context.", szTitle, MB_OK | MB_ICONERROR);
        return false;
    }

    // 2) Gather data from the ListView
    HWND hHeader = (HWND)SendMessageW(hList, LVM_GETHEADER, 0, 0);
    int colCount = hHeader ? (int)SendMessageW(hHeader, HDM_GETITEMCOUNT, 0, 0) : 0;
    if (colCount <= 0) colCount = 1;

    // Header texts
    std::vector<std::wstring> headers;
    headers.reserve(colCount);
    for (int c = 0; c < colCount; ++c) {
        HDITEMW hd{}; hd.mask = HDI_TEXT;
        wchar_t hbuf[256] = L"";
        hd.pszText = hbuf; hd.cchTextMax = (int)_countof(hbuf);
        if (hHeader) SendMessageW(hHeader, HDM_GETITEMW, (WPARAM)c, (LPARAM)&hd);
        headers.emplace_back(hbuf);
    }

    // Rows
    int rowCount = (int)SendMessageW(hList, LVM_GETITEMCOUNT, 0, 0);
    std::vector<std::vector<std::wstring>> rows;
    rows.resize(rowCount, std::vector<std::wstring>(colCount));
    for (int r = 0; r < rowCount; ++r) {
        for (int c = 0; c < colCount; ++c) {
            wchar_t buf[512] = L"";
            LVITEMW it{}; it.iSubItem = c; it.pszText = buf; it.cchTextMax = (int)_countof(buf);
            SendMessageW(hList, LVM_GETITEMTEXTW, (WPARAM)r, (LPARAM)&it);
            rows[r][c] = buf;
        }
    }

    // 3) Set up page metrics
    const int dpiX = GetDeviceCaps(hdc, LOGPIXELSX);
    const int dpiY = GetDeviceCaps(hdc, LOGPIXELSY);
    const int physW = GetDeviceCaps(hdc, PHYSICALWIDTH);
    const int physH = GetDeviceCaps(hdc, PHYSICALHEIGHT);
    const int offsetX = GetDeviceCaps(hdc, PHYSICALOFFSETX);
    const int offsetY = GetDeviceCaps(hdc, PHYSICALOFFSETY);

    // 1 inch margins
    const int marginX = dpiX;
    const int marginY = dpiY;

    const int printableLeft = max(0, marginX - offsetX);
    const int printableTop = max(0, marginY - offsetY);
    const int printableRight = physW - max(0, marginX - offsetX);
    const int printableBottom = physH - max(0, marginY - offsetY);
    const int printableW = max(0, printableRight - printableLeft);
    const int printableH = max(0, printableBottom - printableTop);

    // Fonts
    const int ptBody = 10;
    const int ptHeader = 11;
    HFONT hBody = CreateFontW(-MulDiv(ptBody, dpiY, 72), 0, 0, 0, FW_NORMAL, FALSE, FALSE, FALSE,
        DEFAULT_CHARSET, OUT_DEFAULT_PRECIS, CLIP_DEFAULT_PRECIS, CLEARTYPE_QUALITY,
        DEFAULT_PITCH | FF_DONTCARE, L"Consolas");
    if (!hBody) hBody = (HFONT)GetStockObject(SYSTEM_FONT);

    HFONT hHdrFont = CreateFontW(-MulDiv(ptHeader, dpiY, 72), 0, 0, 0, FW_SEMIBOLD, FALSE, FALSE, FALSE,
        DEFAULT_CHARSET, OUT_DEFAULT_PRECIS, CLIP_DEFAULT_PRECIS, CLEARTYPE_QUALITY,
        DEFAULT_PITCH | FF_DONTCARE, L"Consolas");
    if (!hHdrFont) hHdrFont = hBody;

    TEXTMETRICW tm{};
    HGDIOBJ oldFont = SelectObject(hdc, hBody);
    GetTextMetricsW(hdc, &tm);
    const int lineH = max((int)(tm.tmHeight + tm.tmExternalLeading), MulDiv(14, dpiY, 96)); // ~14px at 96dpi

    // Column widths: evenly divide printable width
    std::vector<int> colLeft(colCount + 1, 0);
    int colW = (colCount > 0) ? printableW / colCount : printableW;
    for (int c = 0; c < colCount; ++c) colLeft[c] = printableLeft + c * colW;
    colLeft[colCount] = printableLeft + printableW;

    // 4) Start document
    DOCINFOW di{}; di.cbSize = sizeof(di); di.lpszDocName = docTitle;
    if (StartDocW(hdc, &di) <= 0) {
        SelectObject(hdc, oldFont);
        if (hBody && hBody != GetStockObject(SYSTEM_FONT)) DeleteObject(hBody);
        if (hHdrFont && hHdrFont != hBody) DeleteObject(hHdrFont);
        DeleteDC(hdc);
        if (pd.hDevMode) GlobalFree(pd.hDevMode);
        if (pd.hDevNames) GlobalFree(pd.hDevNames);
        MessageBoxW(owner, L"Failed to start the print job.", szTitle, MB_OK | MB_ICONERROR);
        return false;
    }

    // Compute rows per page (reserve 2 lines: title + header + small spacing)
    const int usableH = printableH;
    const int headerLines = 2; // title + column header
    const int rowsPerPage = max(1, (usableH / lineH) - headerLines);

    int printed = 0;
    int pageNo = 0;

    while (printed < rowCount || (rowCount == 0 && pageNo == 0)) {
        if (StartPage(hdc) <= 0) break;
        int y = printableTop;

        // Title
        SelectObject(hdc, hHdrFont);
        RECT rcTitle{ printableLeft, y, printableLeft + printableW, y + lineH };
        DrawTextW(hdc, docTitle, -1, &rcTitle, DT_LEFT | DT_NOPREFIX | DT_SINGLELINE | DT_END_ELLIPSIS);
        y += lineH;

        // Column headers
        SelectObject(hdc, hHdrFont);
        for (int c = 0; c < colCount; ++c) {
            RECT rc{ colLeft[c], y, colLeft[c + 1], y + lineH };
            DrawTextW(hdc, headers[c].c_str(), -1, &rc, DT_LEFT | DT_NOPREFIX | DT_SINGLELINE | DT_END_ELLIPSIS | DT_VCENTER);
        }
        y += lineH;

        // Rows on this page
        SelectObject(hdc, hBody);
        int onThisPage = min(rowsPerPage, rowCount - printed);
        for (int i = 0; i < onThisPage; ++i) {
            for (int c = 0; c < colCount; ++c) {
                RECT rc{ colLeft[c], y, colLeft[c + 1], y + lineH };
                const std::wstring& cell = rows[printed + i][c];
                DrawTextW(hdc, cell.c_str(), -1, &rc, DT_LEFT | DT_NOPREFIX | DT_SINGLELINE | DT_END_ELLIPSIS | DT_VCENTER);
            }
            y += lineH;
        }

        // Footer: page number (optional)
        {
            wchar_t pg[64];
            swprintf_s(pg, L"Page %d", pageNo + 1);
            RECT rcPg{ printableLeft, printableBottom - lineH, printableLeft + printableW, printableBottom };
            DrawTextW(hdc, pg, -1, &rcPg, DT_RIGHT | DT_NOPREFIX | DT_SINGLELINE);
        }

        if (EndPage(hdc) <= 0) break;
        printed += onThisPage;
        ++pageNo;

        // Handle empty report: still print one page with just header/title
        if (rowCount == 0) break;
    }

    EndDoc(hdc);

    // Cleanup
    SelectObject(hdc, oldFont);
    if (hBody && hBody != GetStockObject(SYSTEM_FONT)) DeleteObject(hBody);
    if (hHdrFont && hHdrFont != hBody) DeleteObject(hHdrFont);
    DeleteDC(hdc);
    if (pd.hDevMode) GlobalFree(pd.hDevMode);
    if (pd.hDevNames) GlobalFree(pd.hDevNames);

    return true;
}

// Add near other helpers (below Save/Copy helpers)
static void ListView_AutoSizeToHeaderAndContent(HWND hList)
{
    if (!IsWindow(hList)) return;
    // Get column count from header
    HWND hHdr = (HWND)SendMessageW(hList, LVM_GETHEADER, 0, 0);
    int colCount = hHdr ? (int)SendMessageW(hHdr, HDM_GETITEMCOUNT, 0, 0) : 0;
    if (colCount <= 0) return;

    for (int i = 0; i < colCount; ++i) {
        // Size to content
        ListView_SetColumnWidth(hList, i, LVSCW_AUTOSIZE);
        int wContent = ListView_GetColumnWidth(hList, i);
        // Size to header
        ListView_SetColumnWidth(hList, i, LVSCW_AUTOSIZE_USEHEADER);
        int wHeader = ListView_GetColumnWidth(hList, i);
        // Take the larger one
        ListView_SetColumnWidth(hList, i, (wContent > wHeader) ? wContent : wHeader);
    }
}

static void ShowGettingStarted(HWND owner)
{
    const wchar_t* msg =
        L"Getting Started\n"
        L"\n"
        L"1) File -> Create DB to create a new database, or File -> Open DB to open an existing one.\n"
        L"2) Reading -> Add to enter your first blood pressure reading.\n"
        L"3) Use PageUp/PageDown to navigate pages.\n"
        L"4) Right-click the main window for a quick context menu.\n"
        L"5) Reports -> Averages and Reports -> By Dates to view summaries.\n"
        L"\n"
        L"Tips:\n"
        L"- In the By Dates report, press Enter or Esc to close the window.\n"
        L"- Use the Close button at bottom-right to dismiss report windows.";
    MessageBoxW(owner, msg, L"Getting Started", MB_OK | MB_ICONINFORMATION);
}

// Add near other helpers
inline void EnsureWindowOnScreen(HWND hWnd)
{
    RECT rc;
    if (!GetWindowRect(hWnd, &rc)) return;

    HMONITOR mon = MonitorFromRect(&rc, MONITOR_DEFAULTTONEAREST);
    MONITORINFO mi{};
    mi.cbSize = sizeof(mi);
    if (!GetMonitorInfoW(mon, &mi)) return;

    const int w = rc.right - rc.left;
    const int h = rc.bottom - rc.top;

    int x = rc.left;
    int y = rc.top;

    if (x < mi.rcWork.left) x = mi.rcWork.left;
    if (y < mi.rcWork.top)  y = mi.rcWork.top;
    if (x + w > mi.rcWork.right)  x = mi.rcWork.right - w;
    if (y + h > mi.rcWork.bottom) y = mi.rcWork.bottom - h;

    SetWindowPos(hWnd, nullptr, x, y, 0, 0, SWP_NOSIZE | SWP_NOZORDER);
}

// Add near other helpers (below EnsureWindowOnScreen)
static HFONT CreateUiFontForWindow(HWND hWnd, int pointSize, int weight = FW_SEMIBOLD, LPCWSTR face = L"Segoe UI")
{
    UINT dpiY = 96;
    // Try GetDpiForWindow if available (runtime lookup for compatibility)
    HMODULE hUser = GetModuleHandleW(L"user32.dll");
    using GetDpiForWindow_t = UINT(WINAPI*)(HWND);
    GetDpiForWindow_t pGetDpiForWindow = hUser ? (GetDpiForWindow_t)GetProcAddress(hUser, "GetDpiForWindow") : nullptr;
    if (pGetDpiForWindow) {
        dpiY = pGetDpiForWindow(hWnd);
    }
    else {
        HDC hdc = GetDC(hWnd);
        if (hdc) { dpiY = GetDeviceCaps(hdc, LOGPIXELSY); ReleaseDC(hWnd, hdc); }
    }
    const int height = -MulDiv(pointSize, dpiY, 72); // point size to pixels
    return CreateFontW(height, 0, 0, 0, weight, FALSE, FALSE, FALSE, DEFAULT_CHARSET,
                       OUT_OUTLINE_PRECIS, CLIP_DEFAULT_PRECIS, CLEARTYPE_QUALITY,
                       VARIABLE_PITCH, face);
}

// Helper: compute a good control height for the given font (handles DPI and leading)
static int IdealCtlHeightFromFont(HWND hwndRef, HFONT hFont, int minH = 28, int padding = 8)
{
    HDC hdc = GetDC(hwndRef);
    int h = minH;
    if (hdc) {
        HGDIOBJ old = SelectObject(hdc, hFont);
        TEXTMETRICW tm{};
        if (GetTextMetricsW(hdc, &tm)) {
            h = max(minH, (int)(tm.tmHeight + tm.tmExternalLeading + padding));
        }
        SelectObject(hdc, old);
        ReleaseDC(hwndRef, hdc);
    }
    return h;
}

static int RoundAvg(int sum, int count)
{
    return (count > 0) ? (sum + (count / 2)) / count : 0; // integer round to nearest
}

// Parse "YYYY-MM-DDTHH:MM:SSZ" (UTC) and format as local time "YYYY-MM-DD HH:MM"
static std::wstring UtcIsoToLocalDisplay(const std::wstring& isoUtc)
{
    int Y=0,M=0,D=0,h=0,m=0,s=0;
    if (swscanf_s(isoUtc.c_str(), L"%d-%d-%dT%d:%d:%d", &Y, &M, &D, &h, &m, &s) != 6)
        return isoUtc; // fallback

    std::tm tmUtc{};
    tmUtc.tm_year = Y - 1900;
    tmUtc.tm_mon  = M - 1;
    tmUtc.tm_mday = D;
    tmUtc.tm_hour = h;
    tmUtc.tm_min  = m;
    tmUtc.tm_sec  = s;

    time_t t = _mkgmtime(&tmUtc); // interpret as UTC
    if (t == (time_t)-1) return isoUtc;

    std::tm tmLocal{};
    if (localtime_s(&tmLocal, &t) != 0) return isoUtc;

    wchar_t buf[32];
    if (wcsftime(buf, 32, L"%Y-%m-%d %H:%M", &tmLocal) == 0)
        return isoUtc;

    return buf;
}

static int DateKeyFromSystemTime(const SYSTEMTIME& s)
{
    return (int)(s.wYear * 10000 + s.wMonth * 100 + s.wDay);
}

static int DateKeyFromTm(const std::tm& t)
{
    return (t.tm_year + 1900) * 10000 + (t.tm_mon + 1) * 100 + t.tm_mday;
}

int APIENTRY wWinMain(_In_ HINSTANCE hInstance,
                     _In_opt_ HINSTANCE hPrevInstance,
                     _In_ LPWSTR    lpCmdLine,
                     _In_ int       nCmdShow)
{
    UNREFERENCED_PARAMETER(hPrevInstance);
    UNREFERENCED_PARAMETER(lpCmdLine);

    // Init Common Controls (include date-time pickers)
    INITCOMMONCONTROLSEX icc{};
    icc.dwSize = sizeof(icc);
    icc.dwICC = ICC_LISTVIEW_CLASSES | ICC_STANDARD_CLASSES | ICC_DATE_CLASSES;
    InitCommonControlsEx(&icc);

    // Initialize COM (for SHGetKnownFolderPath)
    HRESULT hrCoInit = CoInitializeEx(nullptr, COINIT_APARTMENTTHREADED);
    if (FAILED(hrCoInit)) {
        MessageBoxW(nullptr, L"Failed to initialize COM.", szTitle, MB_ICONERROR | MB_OK);
        return FALSE;
    }

    // Initialize global strings
    LoadStringW(hInstance, IDS_APP_TITLE, szTitle, MAX_LOADSTRING);
    LoadStringW(hInstance, IDC_BLOODPRESSURE, szWindowClass, MAX_LOADSTRING);
    MyRegisterClass(hInstance);

    // Perform application initialization:
    if (!InitInstance (hInstance, nCmdShow))
    {
        CoUninitialize();
        return FALSE;
    }

    // Do not open any database at startup. User can Create/Open from the menu.

    HACCEL hAccelTable = LoadAccelerators(hInstance, MAKEINTRESOURCE(IDC_BLOODPRESSURE));

    MSG msg;

    // Main message loop:
    while (GetMessage(&msg, nullptr, 0, 0))
    {
        if (g_mainWnd && !TranslateAccelerator(g_mainWnd, hAccelTable, &msg))
        {
            TranslateMessage(&msg);
            DispatchMessage(&msg);
        }
        else if (!g_mainWnd)
        {
            TranslateMessage(&msg);
            DispatchMessage(&msg);
        }
    }

    g_db.reset();
    CoUninitialize();

    return (int) msg.wParam;
}

ATOM MyRegisterClass(HINSTANCE hInstance)
{
    WNDCLASSEXW wcex;

    wcex.cbSize = sizeof(WNDCLASSEX);

    wcex.style          = CS_HREDRAW | CS_VREDRAW;
    wcex.lpfnWndProc    = WndProc;
    wcex.cbClsExtra     = 0;
    wcex.cbWndExtra     = 0;
    wcex.hInstance      = hInstance;
    wcex.hIcon          = LoadIcon(hInstance, MAKEINTRESOURCE(IDI_BLOODPRESSURE));
    wcex.hCursor        = LoadCursor(nullptr, IDC_ARROW);
    wcex.hbrBackground  = (HBRUSH)(COLOR_WINDOW+1);
    wcex.lpszMenuName   = MAKEINTRESOURCEW(IDC_BLOODPRESSURE);
    wcex.lpszClassName  = szWindowClass;
    wcex.hIconSm        = LoadIcon(wcex.hInstance, MAKEINTRESOURCE(IDI_SMALL));

    return RegisterClassExW(&wcex);
}

BOOL InitInstance(HINSTANCE hInstance, int nCmdShow)
{
   hInst = hInstance; // Store instance handle in our global variable

   HWND hWnd = CreateWindowW(szWindowClass, szTitle, WS_OVERLAPPEDWINDOW,
      CW_USEDEFAULT, 0, CW_USEDEFAULT, 0, nullptr, nullptr, hInstance, nullptr);

   if (!hWnd)
   {
      return FALSE;
   }

   g_mainWnd = hWnd; // remember main window

   ShowWindow(hWnd, nCmdShow);
   UpdateWindow(hWnd);

   return TRUE;
}

static std::wstring TruncateForDisplay(const std::wstring& s, size_t maxChars)
{
    if (s.size() <= maxChars) return s;
    if (maxChars <= 1) return L"…";
    return s.substr(0, maxChars - 1) + L"…";
}

LRESULT CALLBACK WndProc(HWND hWnd, UINT message, WPARAM wParam, LPARAM lParam)
{
    switch (message)
    {
    case WM_COMMAND:
        {
            int wmId = LOWORD(wParam);
            switch (wmId)
            {
            case IDM_GETTINGSTARTED:
                ShowGettingStarted(hWnd);
                break;
            case IDM_ABOUT:
                DialogBox(hInst, MAKEINTRESOURCE(IDD_ABOUTBOX), hWnd, About);
                break;
            case IDM_CREATE:
                CreateDatabaseDialog(hWnd);
                break;
            case IDM_OPEN:
                OpenDatabaseDialog(hWnd);
                break;
            case IDM_CLOSE:
                CloseDatabaseDialog(hWnd);
                break;
            case IDM_DELETEALL: // <-- add
                if (!g_db) {
                    MessageBoxW(hWnd, L"No database is currently open.", szTitle, MB_OK | MB_ICONINFORMATION);
                    break;
                }
                if (MessageBoxW(hWnd,
                    L"Delete ALL readings? This cannot be undone.",
                    szTitle, MB_ICONWARNING | MB_OKCANCEL | MB_DEFBUTTON2) != IDOK)
                {
                    break;
                }
                if (!g_db->DeleteAllReadings()) {
                    MessageBoxW(hWnd, L"Failed to delete all readings.", szTitle, MB_OK | MB_ICONERROR);
                    break;
                }
                g_pageIndex = 0;
                InvalidateRect(hWnd, nullptr, TRUE);
                UpdateWindow(hWnd);
                MessageBoxW(hWnd, L"All readings deleted.", szTitle, MB_OK | MB_ICONINFORMATION);
                break;
            case IDM_ADD:
                ShowAddReadingDialog(hWnd);
                InvalidateRect(hWnd, nullptr, TRUE);
                break;
            case IDM_EDIT:
            {
                // Open edit dialog in "picker" mode (no preselected reading)
                Reading r{};
                ShowEditReadingDialog(hWnd, r);
                InvalidateRect(hWnd, nullptr, TRUE);
            }
                break;
            case IDM_EXIT:
                DestroyWindow(hWnd);
                break;
            case IDM_PAGE_PREV:
            {
                if (g_pageIndex > 0) {
                    --g_pageIndex;
                    InvalidateRect(hWnd, nullptr, TRUE);
                }
            }
            break;
            case IDM_PAGE_NEXT:
            {
                int total = 0;
                if (g_db && g_db->GetReadingCount(total)) {
                    if ((g_pageIndex + 1) * kPageSize < total) {
                        ++g_pageIndex;
                        InvalidateRect(hWnd, nullptr, TRUE);
                    }
                }
            }
            break;
            case IDM_REPORTALL: // <-- add
                ShowReportAllWindow(hWnd);
                break;
            case IDM_REPORTDATES: // <-- add
                ShowReportDatesWindow(hWnd);
                break;
            default:
                return DefWindowProc(hWnd, message, wParam, lParam);
            }
        }
        break;
    case WM_RBUTTONUP:
        {
            // Context menu mirrors the main menu command
            HMENU hMenu = CreatePopupMenu();
            if (hMenu)
            {
                AppendMenuW(hMenu, MF_STRING, IDM_ADD, L"Add Reading...");
                AppendMenuW(hMenu, MF_STRING, IDM_EDIT, L"Edit Reading...");
                // Paging controls
                AppendMenuW(hMenu, MF_SEPARATOR, 0, nullptr);
                int total = 0;
                bool hasTotal = g_db && g_db->GetReadingCount(total);
                bool canPrev = g_pageIndex > 0;
                bool canNext = hasTotal && ((g_pageIndex + 1) * kPageSize < total);

                AppendMenuW(hMenu, MF_STRING | (canPrev ? 0 : MF_GRAYED), IDM_PAGE_PREV, L"Previous Page");
                AppendMenuW(hMenu, MF_STRING | (canNext ? 0 : MF_GRAYED), IDM_PAGE_NEXT, L"Next Page");
                POINT pt{ GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam) };
                AppendMenuW(hMenu, MF_SEPARATOR, 0, nullptr);
                AppendMenuW(hMenu, MF_STRING, IDM_GETTINGSTARTED, L"Getting Started...");
                ClientToScreen(hWnd, &pt);
                TrackPopupMenu(hMenu, TPM_RIGHTBUTTON | TPM_LEFTALIGN, pt.x, pt.y, 0, hWnd, nullptr);
                DestroyMenu(hMenu);
            }
        }
        break;
    case WM_KEYDOWN:
        if (wParam == VK_PRIOR) { // PageUp
            SendMessage(hWnd, WM_COMMAND, IDM_PAGE_PREV, 0);
            return 0;
        }
        if (wParam == VK_NEXT) { // PageDown
            SendMessage(hWnd, WM_COMMAND, IDM_PAGE_NEXT, 0);
            return 0;
        }
        if (wParam == VK_F1) { // Getting Started
            ShowGettingStarted(hWnd);
            return 0;
        }
        break;
    case WM_PAINT:
        {
            PAINTSTRUCT ps;
            HDC hdc = BeginPaint(hWnd, &ps);

            int y = 10;

            if (g_db)
            {
                //Header
                HFONT hMono = (HFONT)GetStockObject(SYSTEM_FIXED_FONT);
                HGDIOBJ oldFont = SelectObject(hdc, hMono);
                const wchar_t* header = L"No  Date (Local)      Sys/Dia Pul  Note";
                TextOutW(hdc, 10, y, header, lstrlenW(header));
                y += 17;

                // Divider
                const wchar_t* div = L"----------------------------------------------";
                TextOutW(hdc, 10, y, div, lstrlenW(div));
                y += 17;

                // Rows:
                std::vector<Reading> rows;
                if (g_db && g_db->GetRecentReadingsPage(kPageSize, g_pageIndex * kPageSize, rows))
                {
                    int idx = 1;
                    for (const auto& r : rows)
                    {
                        std::wstring note = TruncateForDisplay(r.note, 60);
                        std::wstring tsLocal = UtcIsoToLocalDisplay(r.tsUtc);

                        wchar_t line[512];

                        if (r.diastolic >= 100) {
                            swprintf_s(line, L"%3d %-17s %3d/%3d%3d  %s",
                                idx, tsLocal.c_str(),
                                r.systolic, r.diastolic, r.pulse,
                                note.c_str());
                        }
                        else {
                            swprintf_s(line, L"%3d %-17s %3d/%2d %3d  %s",
                                idx, tsLocal.c_str(),
                                r.systolic, r.diastolic, r.pulse,
                                note.c_str());
                        }

                        TextOutW(hdc, 10, y, line, (int)wcslen(line));
                        y += 17;
                        ++idx;
                        if (y > ps.rcPaint.bottom - 20) break; // avoid drawing off-screen
                    }
                    if (rows.empty())
                    {
                        const wchar_t* none = L"(No readings yet. Use Reading -> Add to create one.)";
                        TextOutW(hdc, 10, y, none, lstrlenW(none));
                        y += 17;
                    }
                    // Page indicator
                    int total = 0;
                    if (g_db->GetReadingCount(total))
                    {
                        std::wstring pageInfo = L"Page " + std::to_wstring(g_pageIndex + 1) +
                                                L" of " + std::to_wstring((total + kPageSize - 1) / kPageSize) +
                                                L" (" + std::to_wstring(total) + L" total)";
                        TextOutW(hdc, 10, y, pageInfo.c_str(), (int)pageInfo.size());
                        y += 17;
                    }
                }
                else
                {
                    const wchar_t* errRows = L"Failed to load readings.";
                    TextOutW(hdc, 10, y, errRows, lstrlenW(errRows));
                    y += 17;
                }

                // Restore font
                SelectObject(hdc, oldFont);
            }
            else
            {
                // No DB open yet: show a friendly hint
                const wchar_t* msg1 = L"No database is open.";
                const wchar_t* msg2 = L"Use File -> Create DB or File -> Open DB to begin.";
                const wchar_t* gs1 = L"Getting Started:";
                const wchar_t* gs2 = L"1) Create/Open a DB, then use Reading -> Add to enter a reading.";
                const wchar_t* gs3 = L"2) PageUp/PageDown to navigate pages.";
                const wchar_t* gs4 = L"3) Reports -> Averages / By Dates for summaries.";
                const wchar_t* gs5 = L"4) Press F1 anytime to view these steps.";

                TextOutW(hdc, 10, y, msg1, lstrlenW(msg1)); y += 17;
                TextOutW(hdc, 10, y, msg2, lstrlenW(msg2)); y += 25;

                TextOutW(hdc, 10, y, gs1, lstrlenW(gs1)); y += 17;
                TextOutW(hdc, 20, y, gs2, lstrlenW(gs2)); y += 17;
                TextOutW(hdc, 20, y, gs3, lstrlenW(gs3)); y += 17;
                TextOutW(hdc, 20, y, gs4, lstrlenW(gs4)); y += 17;
                TextOutW(hdc, 20, y, gs5, lstrlenW(gs5)); y += 17;
            }

            EndPaint(hWnd, &ps);
        }
        break;
    case WM_DESTROY:
    {
        // If this window had an owner (rare for the main window), re-enable it.
        HWND owner = GetWindow(hWnd, GW_OWNER);
        if (owner && IsWindow(owner)) {
            EnableWindow(owner, TRUE);
            ShowWindow(owner, SW_RESTORE);
            SetForegroundWindow(owner);
        }
        PostQuitMessage(0); // <-- make the message loop exit so the process can terminate
        return 0;
    }
    default:
        return DefWindowProc(hWnd, message, wParam, lParam);
    }
    return 0;
}

// Message handler for about box.
INT_PTR CALLBACK About(HWND hDlg, UINT message, WPARAM wParam, LPARAM lParam)
{
    UNREFERENCED_PARAMETER(lParam);
    switch (message)
    {
    case WM_INITDIALOG:
        return (INT_PTR)TRUE;

    case WM_COMMAND:
        if (LOWORD(wParam) == IDOK || LOWORD(wParam) == IDCANCEL)
        {
            EndDialog(hDlg, LOWORD(wParam));
            return (INT_PTR)TRUE;
        }
        break;
    }
    return (INT_PTR)FALSE;
}

// -------------------------
// Add Reading dialog window
// -------------------------
struct DialogInitData
{
    HWND owner{};
    bool editMode{};
    Reading reading{}; // filled when editMode==true
};

struct AddReadingState
{
    HWND owner{};
    HWND hEditSys{};
    HWND hEditDia{};
    HWND hEditPulse{};
    HWND hEditNote{};
    HFONT hFont{};
    bool editMode{};
    int editId{};
    int result{-1};

    // Row picker (edit mode without preselected row)
    HWND hRowCombo{};
    std::vector<Reading> pageRows;

    // New: delete button (only in edit mode)
    HWND hBtnDelete{};
};

static void CenterToOwner(HWND hWnd, HWND owner)
{
    RECT rcOwner{}, rcDlg{};
    GetWindowRect(owner, &rcOwner);
    GetWindowRect(hWnd, &rcDlg);
    int width = rcDlg.right - rcDlg.left;
    int height = rcDlg.bottom - rcDlg.top;
    int x = rcOwner.left + ((rcOwner.right - rcOwner.left) - width) / 2;
    int y = rcOwner.top + ((rcOwner.bottom - rcOwner.top) - height) / 2;
    SetWindowPos(hWnd, nullptr, x, y, 0, 0, SWP_NOZORDER | SWP_NOSIZE);
}

static HWND CreateLabeledEdit(HWND parent, int x, int y, int wLabel, int wEdit, int idEdit, const wchar_t* label, DWORD editStyle)
{
    CreateWindowExW(0, L"STATIC", label,
        WS_CHILD | WS_VISIBLE,
        x, y + 3, wLabel, 20, parent, nullptr, hInst, nullptr);

    HWND hEdit = CreateWindowExW(WS_EX_CLIENTEDGE, L"EDIT", L"",
        WS_CHILD | WS_VISIBLE | WS_TABSTOP | editStyle,
        x + wLabel + 8, y, wEdit, 24, parent, (HMENU)(INT_PTR)idEdit, hInst, nullptr);
    return hEdit;
}

static void LoadReadingIntoFields(AddReadingState* st, const Reading& r)
{
    if (!st) return;
    st->editId = r.id;

    wchar_t buf[32];
    swprintf_s(buf, L"%d", r.systolic);
    SetWindowTextW(st->hEditSys, buf);
    swprintf_s(buf, L"%d", r.diastolic);
    SetWindowTextW(st->hEditDia, buf);
    swprintf_s(buf, L"%d", r.pulse);
    SetWindowTextW(st->hEditPulse, buf);
    SetWindowTextW(st->hEditNote, r.note.c_str());
}

static LRESULT CALLBACK AddReadingWndProc(HWND hWnd, UINT msg, WPARAM wParam, LPARAM lParam)
{
    AddReadingState* st = reinterpret_cast<AddReadingState*>(GetWindowLongPtrW(hWnd, GWLP_USERDATA));
    switch (msg)
    {
    case WM_CREATE:
        {
            st = new AddReadingState();
            auto cs = reinterpret_cast<LPCREATESTRUCT>(lParam);
            auto init = reinterpret_cast<const DialogInitData*>(cs->lpCreateParams);
            if (init)
            {
                st->owner = init->owner;
                st->editMode = init->editMode;
                if (st->editMode)
                    st->editId = init->reading.id;
            }
            SetWindowLongPtrW(hWnd, GWLP_USERDATA, (LONG_PTR)st);

            st->hFont = (HFONT)GetStockObject(DEFAULT_GUI_FONT);

            const int margin = 12;
            const int x = margin;
            int y = margin;

            // Optional row selector (only when edit mode and no preselected reading id)
            if (st->editMode && st->editId == 0)
            {
                CreateWindowExW(0, L"STATIC", L"Row:",
                    WS_CHILD | WS_VISIBLE, x, y + 3, 40, 20, hWnd, nullptr, hInst, nullptr);

                st->hRowCombo = CreateWindowExW(0, L"COMBOBOX", L"",
                    WS_CHILD | WS_VISIBLE | WS_TABSTOP | CBS_DROPDOWNLIST,
                    x + 45, y, 260, 300, hWnd, (HMENU)(INT_PTR)IDC_EDIT_ROWCOMBO, hInst, nullptr);
                SendMessageW(st->hRowCombo, WM_SETFONT, (WPARAM)st->hFont, TRUE);

                // Load current page rows
                if (g_db)
                {
                    g_db->GetRecentReadingsPage(kPageSize, g_pageIndex * kPageSize, st->pageRows);

                    int idx = 1;
                    for (const auto& r : st->pageRows)
                    {
                        std::wstring label =
                            std::to_wstring(idx) + L"  " +
                            UtcIsoToLocalDisplay(r.tsUtc) + L"  " +
                            std::to_wstring(r.systolic) + L"/" + std::to_wstring(r.diastolic) + L"  " +
                            std::to_wstring(r.pulse);
                        int pos = (int)SendMessageW(st->hRowCombo, CB_ADDSTRING, 0, (LPARAM)label.c_str());
                        SendMessageW(st->hRowCombo, CB_SETITEMDATA, pos, (LPARAM)idx - 1);
                        ++idx;
                    }
                    if (!st->pageRows.empty())
                    {
                        SendMessageW(st->hRowCombo, CB_SETCURSEL, 0, 0);
                        LoadReadingIntoFields(st, st->pageRows[0]);
                    }
                }
                y += 34;
            }

            st->hEditSys = CreateLabeledEdit(hWnd, x, y, 80, 80, IDC_EDIT_SYSTOLIC, L"Systolic", ES_NUMBER);
            y += 30;
            st->hEditDia = CreateLabeledEdit(hWnd, x, y, 80, 80, IDC_EDIT_DIASTOLIC, L"Diastolic", ES_NUMBER);
            y += 30;
            st->hEditPulse = CreateLabeledEdit(hWnd, x, y, 80, 80, IDC_EDIT_PULSE, L"Pulse", ES_NUMBER);
            y += 34;

            CreateWindowExW(0, L"STATIC", L"Note",
                WS_CHILD | WS_VISIBLE,
                x, y + 3, 80, 20, hWnd, nullptr, hInst, nullptr);

            st->hEditNote = CreateWindowExW(WS_EX_CLIENTEDGE, L"EDIT", L"",
                WS_CHILD | WS_VISIBLE | WS_TABSTOP | ES_AUTOVSCROLL | ES_MULTILINE,
                x + 80 + 8, y, 220, 60, hWnd, (HMENU)(INT_PTR)IDC_EDIT_NOTE, hInst, nullptr);
            y += 70;

            HWND hOk = CreateWindowExW(0, L"BUTTON", L"OK",
                WS_CHILD | WS_VISIBLE | WS_TABSTOP | BS_DEFPUSHBUTTON,
                x + 80 + 8, y, 80, 26, hWnd, (HMENU)IDOK, hInst, nullptr);

            CreateWindowExW(0, L"BUTTON", L"Cancel",
                WS_CHILD | WS_VISIBLE | WS_TABSTOP,
                x + 80 + 8 + 90, y, 80, 26, hWnd, (HMENU)IDCANCEL, hInst, nullptr);

            // New: Delete button in edit mode
            if (st->editMode) {
                st->hBtnDelete = CreateWindowExW(0, L"BUTTON", L"Delete",
                    WS_CHILD | WS_VISIBLE | WS_TABSTOP,
                    x + 80 + 8 + 90 + 90, y, 80, 26, hWnd, (HMENU)IDC_BTN_DELETE, hInst, nullptr);
                SendMessageW(st->hBtnDelete, WM_SETFONT, (WPARAM)st->hFont, TRUE);
                EnableWindow(st->hBtnDelete, st->editId > 0);
            }

            HWND edits[] = { st->hEditSys, st->hEditDia, st->hEditPulse, st->hEditNote, hOk };
            for (HWND e : edits) { if (e) SendMessageW(e, WM_SETFONT, (WPARAM)st->hFont, TRUE); }

            // Prefill only when explicit reading provided
            if (st->editMode && init && init->reading.id > 0)
            {
                LoadReadingIntoFields(st, init->reading);
            }

            SetFocus(st->hEditSys);

            RECT rc{ 0,0, 360, y + 50 };
            AdjustWindowRectEx(&rc, WS_CAPTION | WS_SYSMENU, FALSE, WS_EX_DLGMODALFRAME);
            SetWindowPos(hWnd, nullptr, 0, 0, rc.right - rc.left, rc.bottom - rc.top, SWP_NOMOVE | SWP_NOZORDER);
            if (st->owner) CenterToOwner(hWnd, st->owner);
        }
        return 0;

        // Inside AddReadingWndProc(...), add:
    case WM_KEYDOWN:
        if (wParam == VK_RETURN) { SendMessageW(hWnd, WM_COMMAND, IDOK, 0); return 0; }
        if (wParam == VK_ESCAPE) { SendMessageW(hWnd, WM_COMMAND, IDCANCEL, 0); return 0; }
        break;

    case WM_COMMAND:
            // Handle row selection change
        if (HIWORD(wParam) == CBN_SELCHANGE && LOWORD(wParam) == IDC_EDIT_ROWCOMBO)
        {
            if (st && st->hRowCombo)
            {
                int sel = (int)SendMessageW(st->hRowCombo, CB_GETCURSEL, 0, 0);
                if (sel >= 0)
                {
                    int vecIndex = (int)SendMessageW(st->hRowCombo, CB_GETITEMDATA, sel, 0);
                    if (vecIndex >= 0 && vecIndex < (int)st->pageRows.size())
                    {
                        LoadReadingIntoFields(st, st->pageRows[vecIndex]);
                        if (st->hBtnDelete) EnableWindow(st->hBtnDelete, st->editId > 0);
                    }
                }
            }
            return 0;
        }

        switch (LOWORD(wParam))
        {
        case IDOK:
            if (st && g_db)
            {
                wchar_t buf[64]{};
                GetWindowTextW(st->hEditSys, buf, 64);
                int sys = (int)wcstol(buf, nullptr, 10);

                GetWindowTextW(st->hEditDia, buf, 64);
                int dia = (int)wcstol(buf, nullptr, 10);

                GetWindowTextW(st->hEditPulse, buf, 64);
                int pulse = (int)wcstol(buf, nullptr, 10);

                std::wstring note;
                {
                    const int len = GetWindowTextLengthW(st->hEditNote);
                    if (len > 0) {
                        std::wstring tmp(len + 1, L'\0');
                        GetWindowTextW(st->hEditNote, tmp.data(), len + 1);
                        tmp.resize(wcslen(tmp.c_str()));
                        note = std::move(tmp);
                    }
                }

                if (sys <= 0 || dia <= 0 || pulse <= 0)
                {
                    MessageBoxW(hWnd, L"Please enter valid positive numbers for systolic, diastolic and pulse.", L"Validation", MB_OK | MB_ICONWARNING);
                    return 0;
                }

                bool ok = false;
                if (st->editMode)
                    ok = g_db->UpdateReading(st->editId, sys, dia, pulse, note.c_str());
                else
                    ok = g_db->AddReading(sys, dia, pulse, note.c_str());

                if (!ok)
                {
                    MessageBoxW(hWnd, L"Failed to save the reading.", L"Error", MB_OK | MB_ICONERROR);
                    return 0;
                }

                // Refresh the main window behind the dialog
                if (st->owner) {
                    InvalidateRect(st->owner, nullptr, TRUE);
                    UpdateWindow(st->owner);
                }

                if (st->editMode) {
                    DestroyWindow(hWnd); // close after edit
                } else {
                    // Clear inputs to allow entering another reading, keep dialog open
                    SetWindowTextW(st->hEditSys, L"");
                    SetWindowTextW(st->hEditDia, L"");
                    SetWindowTextW(st->hEditPulse, L"");
                    SetWindowTextW(st->hEditNote, L"");
                    SetFocus(st->hEditSys);
                }
            }
            return 0;

        case IDCANCEL:
            DestroyWindow(hWnd);
            return 0;

        case IDC_BTN_DELETE:
            if (st && g_db && st->editMode)
            {
                // Confirm before deletion
                if (MessageBoxW(hWnd, L"Delete this reading? This cannot be undone.",
                    szTitle, MB_ICONWARNING | MB_OKCANCEL | MB_DEFBUTTON2) != IDOK)
                {
                    return 0;
                }

                if (!g_db->DeleteReading(st->editId))
                {
                    MessageBoxW(hWnd, L"Failed to delete the reading.", szTitle, MB_OK | MB_ICONERROR);
                    return 0;
                }

                // Refresh main window and close dialog
                if (st->owner) {
                    InvalidateRect(st->owner, nullptr, TRUE);
                    UpdateWindow(st->owner);
                }
                DestroyWindow(hWnd);
                return 0;
            }
            return 0;
        }
        break;

    case WM_CLOSE:
        DestroyWindow(hWnd);
        return 0;

    case WM_NCDESTROY:
        if (st)
        {
            if (st->owner) EnableWindow(st->owner, TRUE);
            SetWindowLongPtrW(hWnd, GWLP_USERDATA, 0);
            delete st;
        }
        return 0;
        InvalidateRect(hWnd, nullptr, TRUE);
    }
    return DefWindowProcW(hWnd, msg, wParam, lParam);
}

static void ShowAddReadingDialog(HWND owner)
{
    static ATOM s_atom = 0;
    if (!s_atom)
    {
        WNDCLASSEXW wc{};
        wc.cbSize = sizeof(wc);
        wc.style = CS_DBLCLKS;
        wc.lpfnWndProc = AddReadingWndProc;
        wc.hInstance = hInst;
        wc.hCursor = LoadCursor(nullptr, IDC_ARROW);
        wc.hIcon = LoadIcon(nullptr, IDI_APPLICATION);
        wc.hbrBackground = (HBRUSH)(COLOR_BTNFACE + 1);
        wc.lpszClassName = L"BP_AddReadingDialog";
        s_atom = RegisterClassExW(&wc);
    }

    DialogInitData init{};
    init.owner = owner;
    init.editMode = false;

    HWND hDlg = CreateWindowExW(WS_EX_DLGMODALFRAME,
        L"BP_AddReadingDialog", L"Add Blood Pressure Reading",
        WS_CAPTION | WS_SYSMENU,
        CW_USEDEFAULT, CW_USEDEFAULT, 360, 240,
        owner, nullptr, hInst, &init);
    if (!hDlg) return;

    EnableWindow(owner, FALSE);
    ShowWindow(hDlg, SW_SHOW);
    UpdateWindow(hDlg);

    MSG msg;
    while (IsWindow(hDlg) && GetMessageW(&msg, nullptr, 0, 0))
    {
        if (msg.message == WM_KEYDOWN &&
            (msg.hwnd == hDlg || IsChild(hDlg, msg.hwnd)))
        {
            if (msg.wParam == VK_RETURN) {
                HWND hFocus = GetFocus();
                if (!hFocus || !(GetWindowLongPtrW(hFocus, GWL_STYLE) & ES_MULTILINE)) {
                    SendMessageW(hDlg, WM_COMMAND, IDOK, 0);
                    continue;
                }
            } else if (msg.wParam == VK_ESCAPE) {
                DestroyWindow(hDlg);
                continue;
            }
        }

        if (!IsDialogMessageW(hDlg, &msg))
        {
            TranslateMessage(&msg);
            DispatchMessageW(&msg);
        }
    }

    // Ensure owner is restored and activated
    if (owner && IsWindow(owner)) {
        EnableWindow(owner, TRUE);
        ShowWindow(owner, SW_RESTORE);
        SetForegroundWindow(owner);
    }
}

static void ShowEditReadingDialog(HWND owner, const Reading& r)
{
    // Register the dialog window class (same as add)
    static ATOM s_atom = 0;
    if (!s_atom)
    {
        WNDCLASSEXW wc{};
        wc.cbSize = sizeof(wc);
        wc.style = CS_DBLCLKS;
        wc.lpfnWndProc = AddReadingWndProc;
        wc.hInstance = hInst;
        wc.hCursor = LoadCursor(nullptr, IDC_ARROW);
        wc.hIcon = LoadIcon(nullptr, IDI_APPLICATION);
        wc.hbrBackground = (HBRUSH)(COLOR_BTNFACE + 1);
        wc.lpszClassName = L"BP_AddReadingDialog";
        s_atom = RegisterClassExW(&wc);
    }

    DialogInitData init{};
    init.owner = owner;
    init.editMode = true;
    init.reading = r;

    HWND hDlg = CreateWindowExW(WS_EX_DLGMODALFRAME,
        L"BP_AddReadingDialog", L"Edit Blood Pressure Reading",
        WS_CAPTION | WS_SYSMENU,
        CW_USEDEFAULT, CW_USEDEFAULT, 360, 240,
        owner, nullptr, hInst, &init);
    if (!hDlg) return;

    EnableWindow(owner, FALSE);
    ShowWindow(hDlg, SW_SHOW);
    UpdateWindow(hDlg);

    MSG msg;
    while (IsWindow(hDlg) && GetMessageW(&msg, nullptr, 0, 0))
    {
        if (msg.message == WM_KEYDOWN &&
            (msg.hwnd == hDlg || IsChild(hDlg, msg.hwnd)))
        {
            if (msg.wParam == VK_RETURN) {
                HWND hFocus = GetFocus();
                if (!hFocus || !(GetWindowLongPtrW(hFocus, GWL_STYLE) & ES_MULTILINE)) {
                    SendMessageW(hDlg, WM_COMMAND, IDOK, 0);
                    continue;
                }
            } else if (msg.wParam == VK_ESCAPE) {
                DestroyWindow(hDlg);
                continue;
            }
        }
        if (!IsDialogMessageW(hDlg, &msg))
        {
            TranslateMessage(&msg);
            DispatchMessageW(&msg);
        }
    }

    // Ensure owner is restored and activated
    if (owner && IsWindow(owner)) {
        EnableWindow(owner, TRUE);
        ShowWindow(owner, SW_RESTORE);
        SetForegroundWindow(owner);
    }
}

// Simple "Create DB" using a Save File dialog and (re)initializing the Database
static void CreateDatabaseDialog(HWND owner)
{
    wchar_t file[MAX_PATH] = L"BloodPressure.db";
    OPENFILENAMEW ofn{};
    ofn.lStructSize = sizeof(ofn);
    ofn.hwndOwner = owner;
    ofn.lpstrFilter = L"Database (*.db)\0*.db\0All Files (*.*)\0*.*\0";
    ofn.lpstrFile = file;
    ofn.nMaxFile = MAX_PATH;
    ofn.lpstrTitle = L"Create Database";
    ofn.Flags = OFN_OVERWRITEPROMPT | OFN_PATHMUSTEXIST;
    ofn.lpstrDefExt = L"db";

    if (!GetSaveFileNameW(&ofn))
        return; // user cancelled

    // (Re)create/open the DB at the chosen path; Initialize will create file/schema as needed
    g_db.reset();
    g_db = std::make_unique<Database>(file);
    if (!g_db->Initialize())
    {
        MessageBoxW(owner, L"Failed to create or initialize the database.", szTitle, MB_OK | MB_ICONERROR);
        g_db.reset();
        return;
    }

    g_pageIndex = 0;
    if (g_mainWnd) InvalidateRect(g_mainWnd, nullptr, TRUE);
    MessageBoxW(owner, L"Database created and ready.", szTitle, MB_OK | MB_ICONINFORMATION);
}

// Simple "Open DB" using an Open File dialog and (re)initializing the Database
static void OpenDatabaseDialog(HWND owner)
{
    wchar_t file[MAX_PATH] = L"";
    OPENFILENAMEW ofn{};
    ofn.lStructSize = sizeof(ofn);
    ofn.hwndOwner = owner;
    ofn.lpstrFilter = L"Database (*.db)\0*.db\0All Files (*.*)\0*.*\0";
    ofn.lpstrFile = file;
    ofn.nMaxFile = MAX_PATH;
    ofn.lpstrTitle = L"Open Database";
    ofn.Flags = OFN_FILEMUSTEXIST | OFN_PATHMUSTEXIST;
    ofn.lpstrDefExt = L"db";

    if (!GetOpenFileNameW(&ofn))
        return; // user cancelled

    // Open the selected DB; Initialize also checks schema
    g_db.reset();
    g_db = std::make_unique<Database>(file);
    if (!g_db->Initialize())
    {
        MessageBoxW(owner, L"Failed to open or initialize the database.", szTitle, MB_OK | MB_ICONERROR);
        g_db.reset();
        return;
    }

    g_pageIndex = 0;
    if (g_mainWnd) InvalidateRect(g_mainWnd, nullptr, TRUE);
    MessageBoxW(owner, L"Database opened successfully.", szTitle, MB_OK | MB_ICONINFORMATION);
}

// Simple "Close DB" confirmation and teardown
static void CloseDatabaseDialog(HWND owner)
{
    if (!g_db)
    {
        MessageBoxW(owner, L"No database is currently open.", szTitle, MB_OK | MB_ICONINFORMATION);
        return;
    }

    const int res = MessageBoxW(owner,
        L"Close the current database?",
        szTitle,
        MB_ICONQUESTION | MB_OKCANCEL | MB_DEFBUTTON2);

    if (res != IDOK)
        return;

    g_db.reset();           // closes the SQLite connection
    g_pageIndex = 0;

    if (g_mainWnd)
        InvalidateRect(g_mainWnd, nullptr, TRUE);

    MessageBoxW(owner, L"Database closed.", szTitle, MB_OK | MB_ICONINFORMATION);
}

// ------------------------------
// Report All Readings window
// ------------------------------
struct ReportAllState
{
    HWND hwnd{};
    HFONT hFont{};
    bool ownsFont{}; // track if we created the font
    HWND hList{};    // ListView (table)
    HWND hClose{};   // Close button
    std::vector<std::wstring> lines; // fallback text to paint (no data)
    HWND dtpH{};     // <-- Add this line to fix C2039 error
    HWND hSave{}; // <-- Save button
    HWND hPrint{}; // <-- Print button
};

// ------------------------------
// Report All Readings window (averages view)
// ------------------------------

// New time buckets : Morning = 00 : 00–11 : 59, Evening = 12 : 00–23 : 59 (local time)
static bool IsMorningHour(int hour) { return hour >= 0 && hour <= 11; }
static bool IsEveningHour(int hour) { return hour >= 12 && hour <= 23; }

static LRESULT CALLBACK ReportAllWndProc(HWND hWnd, UINT msg, WPARAM wParam, LPARAM lParam)
{
    ReportAllState* st = reinterpret_cast<ReportAllState*>(GetWindowLongPtrW(hWnd, GWLP_USERDATA));
    switch (msg)
    {
    case WM_CREATE:
    {
        st = new ReportAllState();
        st->hwnd = hWnd;
        SetWindowLongPtrW(hWnd, GWLP_USERDATA, (LONG_PTR)st);

        // Larger, clearer UI font (DPI-aware), fallback to DEFAULT_GUI_FONT
        st->hFont = CreateUiFontForWindow(hWnd, 14, FW_SEMIBOLD, L"Segoe UI");
        st->ownsFont = (st->hFont != nullptr);
        if (!st->hFont) st->hFont = (HFONT)GetStockObject(DEFAULT_GUI_FONT);

        // Compute averages
        bool hasData = false;
        int cntM = 0, cntE = 0, cntO = 0;
        int avgSysM = 0, avgDiaM = 0, avgPulM = 0;
        int avgSysE = 0, avgDiaE = 0, avgPulE = 0;
        int avgSysO = 0, avgDiaO = 0, avgPulO = 0;

        st->lines.clear();
        if (!g_db) {
            st->lines.push_back(L"No database is open.");
        }
        else {
            int totalCount = 0;
            if (!g_db->GetReadingCount(totalCount) || totalCount == 0) {
                st->lines.push_back(L"No readings yet.");
            }
            else {
                std::vector<Reading> readings;
                if (g_db->GetAllReadings(readings)) {
                    long sumSysM = 0, sumDiaM = 0, sumPulM = 0;
                    long sumSysE = 0, sumDiaE = 0, sumPulE = 0;
                    long sumSysO = 0, sumDiaO = 0, sumPulO = 0;

                    for (const auto& r : readings) {
                        std::tm local{};
                        if (!TryParseUtcIsoToLocalTm(r.tsUtc, local)) continue;

                        sumSysO += r.systolic; sumDiaO += r.diastolic; sumPulO += r.pulse; ++cntO;

                        if (IsMorningHour(local.tm_hour)) {
                            sumSysM += r.systolic; sumDiaM += r.diastolic; sumPulM += r.pulse; ++cntM;
                        }
                        else if (IsEveningHour(local.tm_hour)) {
                            sumSysE += r.systolic; sumDiaE += r.diastolic; sumPulE += r.pulse; ++cntE;
                        }
                    }

                    avgSysM = RoundAvg((int)sumSysM, cntM);
                    avgDiaM = RoundAvg((int)sumDiaM, cntM);
                    avgPulM = RoundAvg((int)sumPulM, cntM);

                    avgSysE = RoundAvg((int)sumSysE, cntE);
                    avgDiaE = RoundAvg((int)sumDiaE, cntE);
                    avgPulE = RoundAvg((int)sumPulE, cntE);

                    avgSysO = RoundAvg((int)sumSysO, cntO);
                    avgDiaO = RoundAvg((int)sumDiaO, cntO);
                    avgPulO = RoundAvg((int)sumPulO, cntO);

                    hasData = true;
                }
                else {
                    st->lines.push_back(L"Failed to load readings.");
                }
            }
        }

        // Close button
        const int margin = 10;
        st->hClose = CreateWindowExW(0, L"BUTTON", L"Close",
            WS_CHILD | WS_VISIBLE | WS_TABSTOP | BS_DEFPUSHBUTTON,
            margin, 0, 80, 26, hWnd, (HMENU)IDOK, hInst, nullptr);
        if (!st->hClose) {
            DWORD err = GetLastError();
            wchar_t msg[128];
            swprintf_s(msg, L"Close button creation failed (err=%lu).", err);
            MessageBoxW(hWnd, msg, szTitle, MB_OK | MB_ICONERROR);
        } else {
            SendMessageW(st->hClose, WM_SETFONT, (WPARAM)st->hFont, TRUE);
            ShowWindow(st->hClose, SW_SHOW);
        }

        // Save button
        st->hSave = CreateWindowExW(0, L"BUTTON", L"Save...",
            WS_CHILD | WS_VISIBLE | WS_TABSTOP,
            margin, 0, 80, 26, hWnd, (HMENU)IDC_REPORTALL_SAVE, hInst, nullptr);
        if (st->hSave) {
            SendMessageW(st->hSave, WM_SETFONT, (WPARAM)st->hFont, TRUE);
            ShowWindow(st->hSave, SW_SHOW);
        }

        // Print button
        st->hPrint = CreateWindowExW(0, L"BUTTON", L"Print...",
            WS_CHILD | WS_VISIBLE | WS_TABSTOP,
            margin, 0, 80, 26, hWnd, (HMENU)IDC_REPORTALL_PRINT, hInst, nullptr);
        if (st->hPrint) {
            SendMessageW(st->hPrint, WM_SETFONT, (WPARAM)st->hFont, TRUE);
            ShowWindow(st->hPrint, SW_SHOW);
        }

        // ListView (table) for the averages
        st->hList = CreateWindowExW(WS_EX_CLIENTEDGE, WC_LISTVIEWW, L"",
            WS_CHILD | WS_VISIBLE | WS_TABSTOP | LVS_REPORT | LVS_SINGLESEL | LVS_SHOWSELALWAYS,
            margin, margin, 100, 100, hWnd, (HMENU)42001, hInst, nullptr);
        if (!st->hList) {
            DWORD err = GetLastError();
            wchar_t msg[128];
            swprintf_s(msg, L"ListView creation failed (err=%lu).", err);
            MessageBoxW(hWnd, msg, szTitle, MB_OK | MB_ICONERROR);
        } else {
            SendMessageW(st->hList, WM_SETFONT, (WPARAM)st->hFont, TRUE);
            ListView_SetExtendedListViewStyle(st->hList,
                LVS_EX_FULLROWSELECT | LVS_EX_GRIDLINES | LVS_EX_DOUBLEBUFFER);

            // Columns: Bucket | N | Avg Sys | Avg Dia | Avg Pulse
            LVCOLUMNW col{};
            col.mask = LVCF_TEXT | LVCF_WIDTH | LVCF_SUBITEM | LVCF_FMT;

            struct ColDef { const wchar_t* text; int width; int fmt; } cols[] = {
                { L"Bucket",              160, LVCFMT_LEFT   },
                { L"N",                    80, LVCFMT_RIGHT  },
                { L"Avg Sys/Avg Dia",     120, LVCFMT_CENTER},
                { L"Avg Pulse",           100, LVCFMT_RIGHT  },
            };
            for (int i = 0; i < (int)(sizeof(cols) / sizeof(cols[0])); ++i) {
                col.pszText = const_cast<wchar_t*>(cols[i].text);
                col.cx = cols[i].width;
                col.iSubItem = i;
                col.fmt = cols[i].fmt;
                ListView_InsertColumn(st->hList, i, &col);
            }

            if (hasData) {
                auto addRow = [&](int row, const wchar_t* bucket, int n, int s, int d, int p) {
                    LVITEMW it{};
                    it.mask = LVIF_TEXT;
                    it.iItem = row;
                    it.iSubItem = 0;
                    it.pszText = const_cast<wchar_t*>(bucket);
                    ListView_InsertItemW(st->hList, &it);

                    wchar_t buf[32];
                    swprintf_s(buf, L"%d", n);
                    ListView_SetItemTextW(st->hList, row, 1, buf);
                    swprintf_s(buf, L"%d/%d", s, d);
                    ListView_SetItemTextW(st->hList, row, 2, buf);
                    swprintf_s(buf, L"%d", p);
                    ListView_SetItemTextW(st->hList, row, 3, buf);
                    };

                int r = 0;
                addRow(r++, L"Morning (00:00–11:59)", cntM, avgSysM, avgDiaM, avgPulM);
                addRow(r++, L"Evening (12:00–23:59)", cntE, avgSysE, avgDiaE, avgPulE);
                addRow(r++, L"Overall", cntO, avgSysO, avgDiaO, avgPulO);
                ListView_AutoSizeToHeaderAndContent(st->hList);
            }
            else if (st->hList) {
                LVITEMW it{};
                it.mask = LVIF_TEXT;
                it.iItem = 0;
                it.iSubItem = 0;
                it.pszText = const_cast<wchar_t*>(L"No data");
                ListView_InsertItemW(st->hList, &it);
                ListView_AutoSizeToHeaderAndContent(st->hList);
            }
        }

        // Auto-size columns, then lay out controls and window like the Report All window
        if (st->hList) {
            for (int i = 0; i < 4; ++i) {
                ListView_SetColumnWidth(st->hList, i, LVSCW_AUTOSIZE_USEHEADER);
            }
        }

        // Size the window and position controls
        RECT rc{ 0,0, 760, 420 };
        AdjustWindowRectEx(&rc, WS_CAPTION | WS_SYSMENU, FALSE, WS_EX_DLGMODALFRAME);
        SetWindowPos(hWnd, nullptr, 0, 0, rc.right - rc.left, rc.bottom - rc.top, SWP_NOZORDER);

        RECT rcClient;
        GetClientRect(hWnd, &rcClient);
        const int btnW = 80, btnH = 26;

        if (st->hClose) {
            SetWindowPos(st->hClose, nullptr, rcClient.right - (btnW + margin), rcClient.bottom - (btnH + margin),
                btnW, btnH, SWP_NOZORDER);
        }
        if (st->hSave) {
            int xClose = rcClient.right - (btnW + margin);
            SetWindowPos(st->hSave, nullptr, xClose - (btnW + margin), rcClient.bottom - (btnH + margin),
                btnW, btnH, SWP_NOZORDER);
        }
        if (st->hList) {
            int listRight = rcClient.right - margin;
            int listBottom = (st->hClose ? (rcClient.bottom - (btnH + 2 * margin)) : (rcClient.bottom - margin));
            SetWindowPos(st->hList, nullptr, margin, margin,
                listRight - margin, listBottom - margin, SWP_NOZORDER);
        }

        CenterToOwner(hWnd, GetWindow(hWnd, GW_OWNER));
        EnsureWindowOnScreen(hWnd);    }

        return 0;

    case WM_NOTIFY:
    {
        if (!st) break;
        auto hdr = reinterpret_cast<LPNMHDR>(lParam);
        if (hdr->hwndFrom == st->hList && hdr->code == LVN_KEYDOWN)
        {
            auto p = reinterpret_cast<LPNMLVKEYDOWN>(lParam);
            if ((GetKeyState(VK_CONTROL) & 0x8000) && (p->wVKey == 'C' || p->wVKey == VK_INSERT)) {
                CopyListViewAsTsvToClipboard(hWnd, st->hList);
                return 0;
            }
        }
    }
    break;

    case WM_SIZE:
    {
        if (!st) break;
        RECT rc{};
        GetClientRect(hWnd, &rc);
        const int margin = 10;
        const int btnW = 80, btnH = 26;
        if (st->hClose) {
            SetWindowPos(st->hClose, nullptr, rc.right - (btnW + margin), rc.bottom - (btnH + margin),
                btnW, btnH, SWP_NOZORDER);
        }
        if (st->hSave) {
            int xClose = rc.right - (btnW + margin);
            SetWindowPos(st->hSave, nullptr, xClose - (btnW + margin), rc.bottom - (btnH + margin),
                btnW, btnH, SWP_NOZORDER);
        }
        if (st->hPrint) {
            int xSave = rc.right - (2 * (btnW + margin));
            SetWindowPos(st->hPrint, nullptr, xSave - (btnW + margin), rc.bottom - (btnH + margin),
                btnW, btnH, SWP_NOZORDER);
        }
        if (st->hList) {
            int top = margin * 2;
            top += 24; // simple top offset; no DTPs in this window
            int listRight = rc.right - margin;
            int listBottom = (st->hClose ? (rc.bottom - (btnH + 2 * margin)) : (rc.bottom - margin));
            SetWindowPos(st->hList, nullptr, margin, top, listRight - margin, max(0, listBottom - top), SWP_NOZORDER);

            ListView_AutoSizeToHeaderAndContent(st->hList);
        }
    }
    return 0;

    // Inside ReportAllWndProc(...), add:
    case WM_KEYDOWN:
        if ((GetKeyState(VK_CONTROL) & 0x8000) && (wParam == 'C')) {
            if (st && st->hList) CopyListViewAsTsvToClipboard(hWnd, st->hList);
            return 0;
        }
        if (wParam == VK_RETURN) { SendMessageW(hWnd, WM_COMMAND, IDOK, 0); return 0; }
        if (wParam == VK_ESCAPE) { DestroyWindow(hWnd); return 0; }
        break;

    case WM_COMMAND:
        if (LOWORD(wParam) == IDOK)
        {
            DestroyWindow(hWnd);
            return 0;
        }
        if (LOWORD(wParam) == IDC_REPORTALL_SAVE)
        {
            if (st) SaveListViewAsCsv(hWnd, st->hList, L"Report_Averages.csv", L"Save Averages Report As");
            return 0;
        }
        if (LOWORD(wParam) == IDC_REPORTALL_PRINT)
        {
            if (st) {
                std::wstring title = L"Averages";
                // If possible, include first/last dates
                if (g_db) {
                    std::vector<Reading> readings;
                    if (g_db->GetAllReadings(readings) && !readings.empty()) {
                        // GetAllReadings is ordered by timestamp desc -> front=newest, back=oldest
                        const std::wstring newest = LocalDateYmdFromUtcIso(readings.front().tsUtc);
                        const std::wstring oldest = LocalDateYmdFromUtcIso(readings.back().tsUtc);
                        if (!newest.empty() && !oldest.empty()) {
                            title = L"Averages: " + oldest + L" — " + newest;
                        }
                    }
                }
                PrintListView(hWnd, st->hList, title.c_str());
            }
            return 0;
        }
        break;

    case WM_PAINT:
    {
        // If table exists and is visible, no need to paint fallback text
        if (st && st->hList && IsWindowVisible(st->hList)) {
            ValidateRect(hWnd, nullptr);
            return 0;
        }

        // Fallback text (e.g., no DB/no data)
        PAINTSTRUCT ps;
        HDC hdc = BeginPaint(hWnd, &ps);
        HGDIOBJ old = SelectObject(hdc, st ? st->hFont : GetStockObject(DEFAULT_GUI_FONT));

        int y = 10;
        const int x = 10;
        if (st) {
            for (const auto& line : st->lines) {
                TextOutW(hdc, x, y, line.c_str(), (int)line.size());
                y += 20;
            }
        }

        SelectObject(hdc, old);
        EndPaint(hWnd, &ps);
    }
    return 0;

    case WM_DISPLAYCHANGE:
        EnsureWindowOnScreen(hWnd);
        return 0;

    case WM_DESTROY:
    {
        HWND owner = GetWindow(hWnd, GW_OWNER);
        if (owner && IsWindow(owner)) {
            EnableWindow(owner, TRUE);
            ShowWindow(owner, SW_RESTORE);
            SetForegroundWindow(owner);
        }
        return 0;
    }

    case WM_NCDESTROY:
        if (st) {
            if (st->ownsFont && st->hFont) {
                DeleteObject(st->hFont);
                st->hFont = nullptr;
            }
            SetWindowLongPtrW(hWnd, GWLP_USERDATA, 0);
            delete st;
        }
        return 0;
    }
    return DefWindowProcW(hWnd, msg, wParam, lParam);
}

// ------------------------------
// Report Dates window
// ------------------------------
struct ReportDatesState
{
    HWND hwnd{};
    HFONT hFont{};
    bool ownsFont{};
    HWND hStart{};
    HWND hEnd{};
    HWND hOk{};
    HWND hCancel{};
    SYSTEMTIME stStart{};
    SYSTEMTIME stEnd{};
    int dtpH{24}; // <-- ideal height for DateTime pickers

    // New: combined view pieces
    HWND hList{};           // ListView with results
    bool listInit{ false };   // columns inserted once

    // FIX: Add missing hClose member to match ReportAllState and resolve C2039
    HWND hClose{}; // Add this line
    HWND hSave{}; // <-- Save button
    HWND hPrint{}; // <-- Print button

    bool suppressNotify{}; // <-- guard re-entrant DTP updates
};

// Fills the ListView with Morning/Evening/Overall averages in the given date range
static void FillDatesAveragesList(HWND hList, const SYSTEMTIME& stStart, const SYSTEMTIME& stEnd)
{
    if (!hList) return;

    // Clear current items
    ListView_DeleteAllItems(hList);

    if (!g_db) {
        LVITEMW it{}; it.mask = LVIF_TEXT; it.iItem = 0;
        it.pszText = const_cast<wchar_t*>(L"No database is open.");
        ListView_InsertItemW(hList, &it);
        return;
    }

    std::vector<Reading> all;
    if (!g_db->GetAllReadings(all)) {
        LVITEMW it{}; it.mask = LVIF_TEXT; it.iItem = 0;
        it.pszText = const_cast<wchar_t*>(L"Failed to load readings.");
        ListView_InsertItemW(hList, &it);
        return;
    }

    const int startKey = DateKeyFromSystemTime(stStart);
    const int endKeyRaw = DateKeyFromSystemTime(stEnd);
    const int startK = (startKey <= endKeyRaw) ? startKey : endKeyRaw;
    const int endK = (startKey <= endKeyRaw) ? endKeyRaw : startKey;

    int cntM = 0, cntE = 0, cntO = 0;
    long sumSysM = 0, sumDiaM = 0, sumPulM = 0;
    long sumSysE = 0, sumDiaE = 0, sumPulE = 0;
    long sumSysO = 0, sumDiaO = 0, sumPulO = 0;

    for (const auto& r : all) {
        std::tm local{};
        if (!TryParseUtcIsoToLocalTm(r.tsUtc, local)) continue;

        const int key = DateKeyFromTm(local);
        if (key < startK || key > endK) continue;

        sumSysO += r.systolic; sumDiaO += r.diastolic; sumPulO += r.pulse; ++cntO;

        if (IsMorningHour(local.tm_hour)) {
            sumSysM += r.systolic; sumDiaM += r.diastolic; sumPulM += r.pulse; ++cntM;
        }
        else if (IsEveningHour(local.tm_hour)) {
            sumSysE += r.systolic; sumDiaE += r.diastolic; sumPulE += r.pulse; ++cntE;
        }
    }

    if (cntO == 0) {
        LVITEMW it{}; it.mask = LVIF_TEXT; it.iItem = 0;
        it.pszText = const_cast<wchar_t*>(L"No data in selected range.");
        ListView_InsertItemW(hList, &it);
    }
    else {
        const int avgSysM = RoundAvg((int)sumSysM, cntM);
        const int avgDiaM = RoundAvg((int)sumDiaM, cntM);
        const int avgPulM = RoundAvg((int)sumPulM, cntM);

        const int avgSysE = RoundAvg((int)sumSysE, cntE);
        const int avgDiaE = RoundAvg((int)sumDiaE, cntE);
        const int avgPulE = RoundAvg((int)sumPulE, cntE);

        const int avgSysO = RoundAvg((int)sumSysO, cntO);
        const int avgDiaO = RoundAvg((int)sumDiaO, cntO);
        const int avgPulO = RoundAvg((int)sumPulO, cntO);

        auto addRow = [&](int row, const wchar_t* bucket, int n, int s, int d, int p) {
            LVITEMW it{}; it.mask = LVIF_TEXT; it.iItem = row; it.iSubItem = 0;
            it.pszText = const_cast<wchar_t*>(bucket);
            ListView_InsertItemW(hList, &it);

            wchar_t buf[32];
            swprintf_s(buf, L"%d", n); ListView_SetItemTextW(hList, row, 1, buf);
            swprintf_s(buf, L"%d/%d", s, d); ListView_SetItemTextW(hList, row, 2, buf);
            swprintf_s(buf, L"%d", p); ListView_SetItemTextW(hList, row, 3, buf);
            };

        int r = 0;
        addRow(r++, L"Morning (00:00–11:59)", cntM, avgSysM, avgDiaM, avgPulM);
        addRow(r++, L"Evening (12:00–23:59)", cntE, avgSysE, avgDiaE, avgPulE);
        addRow(r++, L"Overall", cntO, avgSysO, avgDiaO, avgPulO);
    }

    for (int i = 0; i < 4; ++i) {
        ListView_SetColumnWidth(hList, i, LVSCW_AUTOSIZE_USEHEADER);
    }
}

static void ShowReportAllWindow(HWND owner)
{
    static ATOM s_atom = 0;
    if (!s_atom)
    {
        WNDCLASSEXW wc{};
        wc.cbSize = sizeof(wc);
        wc.style = CS_DBLCLKS;
        wc.lpfnWndProc = ReportAllWndProc;
        wc.hInstance = hInst;
        wc.hCursor = LoadCursor(nullptr, IDC_ARROW);
        wc.hIcon = LoadIcon(nullptr, IDI_APPLICATION);
        wc.hbrBackground = (HBRUSH)(COLOR_BTNFACE + 1);
        wc.lpszClassName = L"BP_ReportAllWindow";
        s_atom = RegisterClassExW(&wc);
    }

    const DWORD style = WS_OVERLAPPED | WS_CAPTION | WS_SYSMENU | WS_CLIPCHILDREN | WS_CLIPSIBLINGS;

    HWND hWnd = CreateWindowExW(WS_EX_DLGMODALFRAME,
        L"BP_ReportAllWindow", L"Report - Averages",
        style,
        CW_USEDEFAULT, CW_USEDEFAULT, 620, 280,
        owner, nullptr, hInst, nullptr);

    if (!hWnd) {
        DWORD err = GetLastError();
        wchar_t msg[128];
        swprintf_s(msg, L"Report All window failed (err=%lu).", err);
        MessageBoxW(owner, msg, szTitle, MB_OK | MB_ICONERROR);
        return;
    }

    EnableWindow(owner, FALSE);
    ShowWindow(hWnd, SW_SHOW);
    UpdateWindow(hWnd);
}

// Replace the whole ReportDatesWndProc with this fixed version (removes duplicate case IDOK and handles OK in WM_COMMAND)
static LRESULT CALLBACK ReportDatesWndProc(HWND hWnd, UINT msg, WPARAM wParam, LPARAM lParam)
{
    ReportDatesState* st = reinterpret_cast<ReportDatesState*>(GetWindowLongPtrW(hWnd, GWLP_USERDATA));
    switch (msg)
    {
    case WM_CREATE:
    {
        st = new ReportDatesState();
        st->hwnd = hWnd;
        SetWindowLongPtrW(hWnd, GWLP_USERDATA, (LONG_PTR)st);
        st->suppressNotify = false;

        // Ideal height for DTPs (compute from current UI font to avoid clipping)
        st->dtpH = IdealCtlHeightFromFont(hWnd, st->hFont, 28, 10);

        if (st->hStart) SetWindowPos(st->hStart, nullptr, 0, 0, 160, st->dtpH, SWP_NOMOVE | SWP_NOZORDER);
        if (st->hEnd)   SetWindowPos(st->hEnd, nullptr, 0, 0, 160, st->dtpH, SWP_NOMOVE | SWP_NOZORDER);

        st->hFont = CreateUiFontForWindow(hWnd, 14, FW_SEMIBOLD, L"Segoe UI");
        st->ownsFont = (st->hFont != nullptr);
        if (!st->hFont) st->hFont = (HFONT)GetStockObject(DEFAULT_GUI_FONT);

        const int margin = 10;
        const int ctlH = 24;

        CreateWindowExW(0, L"STATIC", L"Start:",
            WS_CHILD | WS_VISIBLE, margin, margin + 4, 45, ctlH, hWnd, nullptr, hInst, nullptr);

        st->hStart = CreateWindowExW(0, DATETIMEPICK_CLASSW, L"",
            WS_CHILD | WS_VISIBLE | WS_TABSTOP | DTS_SHORTDATECENTURYFORMAT,
            margin + 50, margin, 160, 24, hWnd, (HMENU)IDC_DATES_START, hInst, nullptr);

        CreateWindowExW(0, L"STATIC", L"End:",
            WS_CHILD | WS_VISIBLE, margin + 50 + 160 + 10, margin + 4, 35, ctlH, hWnd, nullptr, hInst, nullptr);

        st->hEnd = CreateWindowExW(0, DATETIMEPICK_CLASSW, L"",
            WS_CHILD | WS_VISIBLE | WS_TABSTOP | DTS_SHORTDATECENTURYFORMAT,
            margin + 50 + 160 + 10 + 40, margin, 160, 24, hWnd, (HMENU)IDC_DATES_END, hInst, nullptr);

        DateTime_SetFormatW(st->hStart, L"yyyy-MM-dd");
        DateTime_SetFormatW(st->hEnd, L"yyyy-MM-dd");

        st->hClose = CreateWindowExW(0, L"BUTTON", L"Close",
            WS_CHILD | WS_VISIBLE | WS_TABSTOP,
            margin, 0, 80, 26, hWnd, (HMENU)IDCANCEL, hInst, nullptr);

        // Save button
        st->hSave = CreateWindowExW(0, L"BUTTON", L"Save...",
            WS_CHILD | WS_VISIBLE | WS_TABSTOP,
            margin, 0, 80, 26, hWnd, (HMENU)IDC_REPORTDATES_SAVE, hInst, nullptr);
        if (st->hSave) {
            SendMessageW(st->hSave, WM_SETFONT, (WPARAM)st->hFont, TRUE); // make font match other buttons
            ShowWindow(st->hSave, SW_SHOW);
        }

        // Print button
        st->hPrint = CreateWindowExW(0, L"BUTTON", L"Print...",
            WS_CHILD | WS_VISIBLE | WS_TABSTOP,
            margin, 0, 80, 26, hWnd, (HMENU)IDC_REPORTDATES_PRINT, hInst, nullptr);
        if (st->hPrint) {
            SendMessageW(st->hPrint, WM_SETFONT, (WPARAM)st->hFont, TRUE);
            ShowWindow(st->hPrint, SW_SHOW);
        }
        
        // Apply fonts
        HWND cts[] = { st->hStart, st->hEnd, st->hClose, st->hSave, st->hPrint };
        for (HWND c : cts) { if (c) SendMessageW(c, WM_SETFONT, (WPARAM)st->hFont, TRUE); }

        // Initialize dates to today
        GetLocalTime(&st->stStart);
        st->stEnd = st->stStart;
        DateTime_SetSystemtime(st->hStart, GDT_VALID, &st->stStart);
        DateTime_SetSystemtime(st->hEnd, GDT_VALID, &st->stEnd);

        // Ideal height for DTPs
        SIZE ideal{};
        if (st->hStart && SendMessageW(st->hStart, DTM_GETIDEALSIZE, 0, (LPARAM)&ideal)) {
            st->dtpH = max(ideal.cy, 24);
        }
        if (st->hStart) SetWindowPos(st->hStart, nullptr, 0, 0, 160, st->dtpH, SWP_NOMOVE | SWP_NOZORDER);
        if (st->hEnd)   SetWindowPos(st->hEnd, nullptr, 0, 0, 160, st->dtpH, SWP_NOMOVE | SWP_NOZORDER);

        // Create ListView for results
        st->hList = CreateWindowExW(WS_EX_CLIENTEDGE, WC_LISTVIEWW, L"",
            WS_CHILD | WS_VISIBLE | WS_TABSTOP | LVS_REPORT | LVS_SINGLESEL | LVS_SHOWSELALWAYS,
            margin, margin * 3 + st->dtpH, 100, 100, hWnd, (HMENU)43060, hInst, nullptr);
        if (st->hList) {
            SendMessageW(st->hList, WM_SETFONT, (WPARAM)st->hFont, TRUE);
            ListView_SetExtendedListViewStyle(st->hList, LVS_EX_FULLROWSELECT | LVS_EX_GRIDLINES | LVS_EX_DOUBLEBUFFER);

            // Columns (insert once)
            LVCOLUMNW col{};
            col.mask = LVCF_TEXT | LVCF_WIDTH | LVCF_SUBITEM | LVCF_FMT;
            struct ColDef { const wchar_t* text; int width; int fmt; } cols[] = {
                { L"Bucket",              160, LVCFMT_LEFT   },
                { L"N",                    80, LVCFMT_RIGHT  },
                { L"Avg Sys/Avg Dia",     120, LVCFMT_CENTER},
                { L"Avg Pulse",           100, LVCFMT_RIGHT  },
            };
            for (int i = 0; i < (int)(sizeof(cols) / sizeof(cols[0])); ++i) {
                col.pszText = const_cast<wchar_t*>(cols[i].text);
                col.cx = cols[i].width;
                col.iSubItem = i;
                col.fmt = cols[i].fmt;
                ListView_InsertColumn(st->hList, i, &col);
            }
            st->listInit = true;
        }

        // Grow window to comfortable size
        RECT rc{ 0,0, 760, 420 };
        AdjustWindowRectEx(&rc, WS_CAPTION | WS_SYSMENU, FALSE, WS_EX_DLGMODALFRAME);
        SetWindowPos(hWnd, nullptr, 0, 0, rc.right - rc.left, rc.bottom - rc.top, SWP_NOZORDER);

        // Place buttons bottom-right
        RECT rcClient;
        GetClientRect(hWnd, &rcClient);
        const int btnW = 80, btnH = 26;
        if (st->hClose) {
            SetWindowPos(st->hClose, nullptr, rcClient.right - (btnW + margin), rcClient.bottom - (btnH + margin),
                btnW, btnH, SWP_NOZORDER);
        }
        if (st->hSave) {
            int xClose = rcClient.right - (btnW + margin);
            SetWindowPos(st->hSave, nullptr, xClose - (btnW + margin), rcClient.bottom - (btnH + margin),
                btnW, btnH, SWP_NOZORDER);
        }
        if (st->hList) {
            ListView_AutoSizeToHeaderAndContent(st->hList);
        }

        CenterToOwner(hWnd, GetWindow(hWnd, GW_OWNER));
        EnsureWindowOnScreen(hWnd);
    }
    return 0;

    case WM_SIZE:
    {
        if (!st) break;
        RECT rc{}; GetClientRect(hWnd, &rc);
        const int margin = 10;
        const int btnW = 80, btnH = 26;

        // Reposition pickers
        if (st->hStart) SetWindowPos(st->hStart, nullptr, margin + 50, margin, 160, st->dtpH, SWP_NOZORDER);
        if (st->hEnd)   SetWindowPos(st->hEnd, nullptr, margin + 50 + 160 + 10 + 40, margin, 160, st->dtpH, SWP_NOZORDER);

        // Buttons bottom-right
        if (st->hClose) {
            SetWindowPos(st->hClose, nullptr, rc.right - (btnW + margin), rc.bottom - (btnH + margin),
                btnW, btnH, SWP_NOZORDER);
        }
        if (st->hSave) {
            int xClose = rc.right - (btnW + margin);
            SetWindowPos(st->hSave, nullptr, xClose - (btnW + margin), rc.bottom - (btnH + margin),
                btnW, btnH, SWP_NOZORDER);
        }
        if (st->hPrint) {
            int xSave = rc.right - (2 * (btnW + margin));
            SetWindowPos(st->hPrint, nullptr, xSave - (btnW + margin), rc.bottom - (btnH + margin),
                btnW, btnH, SWP_NOZORDER);
        }
        if (st->hList) {
            int top = margin * 2;

            // st->dtpH is an int height, not a HWND. Use it directly.
            top += (st->dtpH > 0 ? st->dtpH : 24);

            int listRight = rc.right - margin;
            int listBottom = (st->hClose ? (rc.bottom - (btnH + 2 * margin)) : (rc.bottom - margin));
            SetWindowPos(st->hList, nullptr, margin, top, listRight - margin, max(0, listBottom - top), SWP_NOZORDER);

            for (int i = 0; i < 4; ++i) {
                ListView_SetColumnWidth(st->hList, i, LVSCW_AUTOSIZE_USEHEADER);
            }
        }
    }
    return 0;

    // In ReportDatesWndProc(...) extend the existing WM_NOTIFY to also handle LVN_KEYDOWN from the ListView.
    case WM_NOTIFY:
    {
        if (!st) break;
        auto nm = reinterpret_cast<LPNMHDR>(lParam);

        // New: enable Ctrl+C copy when focus is in the ListView
        if (nm->hwndFrom == st->hList && nm->code == LVN_KEYDOWN) {
            auto p = reinterpret_cast<LPNMLVKEYDOWN>(lParam);
            if ((GetKeyState(VK_CONTROL) & 0x8000) && (p->wVKey == 'C' || p->wVKey == VK_INSERT)) {
                CopyListViewAsTsvToClipboard(hWnd, st->hList);
                return 0;
            }
        }

        const bool isStart = (nm->idFrom == IDC_DATES_START);
        const bool isEnd = (nm->idFrom == IDC_DATES_END);
        if ((isStart || isEnd) && (nm->code == DTN_DATETIMECHANGE || nm->code == DTN_CLOSEUP))
        {
            if (nm->code == DTN_DATETIMECHANGE)
            {
                auto pdt = reinterpret_cast<LPNMDATETIMECHANGE>(lParam);
                if (pdt->dwFlags == GDT_VALID) {
                    if (isStart) st->stStart = pdt->st;
                    else         st->stEnd = pdt->st;
                }
            }
            else
            {
                DateTime_GetSystemtime(st->hStart, &st->stStart);
                DateTime_GetSystemtime(st->hEnd, &st->stEnd);
            }

            // Enforce chronological order: start <= end
            if (!st->suppressNotify && CompareSystemTimes(st->stStart, st->stEnd) > 0) {
                st->suppressNotify = true;
                if (isStart) {
                    // Start moved after End -> move End to Start
                    st->stEnd = st->stStart;
                    DateTime_SetSystemtime(st->hEnd, GDT_VALID, &st->stEnd);
                }
                else {
                    // End moved before Start -> move Start to End
                    st->stStart = st->stEnd;
                    DateTime_SetSystemtime(st->hStart, GDT_VALID, &st->stStart);
                }
                st->suppressNotify = false;
            }
            
            FillDatesAveragesList(st->hList, st->stStart, st->stEnd);
            ListView_AutoSizeToHeaderAndContent(st->hList);
            InvalidateRect(st->hList, nullptr, FALSE);
            return 0;
        }
    }
    break;

    // Inside ReportDatesWndProc(...), add:
    case WM_KEYDOWN:
        if ((GetKeyState(VK_CONTROL) & 0x8000) && (wParam == 'C')) {
            if (st && st->hList) CopyListViewAsTsvToClipboard(hWnd, st->hList);
            return 0;
        }
        if (wParam == VK_RETURN || wParam == VK_ESCAPE) {
            // Simulate pressing the Close button
            SendMessageW(hWnd, WM_COMMAND, IDCANCEL, 0);
            return 0;
        }
        break;

    case WM_COMMAND:
        switch (LOWORD(wParam))
        {
        case IDOK: // Refresh
            if (st) {
                DateTime_GetSystemtime(st->hStart, &st->stStart);
                DateTime_GetSystemtime(st->hEnd, &st->stEnd);

                if (CompareSystemTimes(st->stStart, st->stEnd) > 0) {
                    // Normalize UI and state so start <= end
                    st->suppressNotify = true;
                    st->stEnd = st->stStart;
                    DateTime_SetSystemtime(st->hEnd, GDT_VALID, &st->stEnd);
                    st->suppressNotify = false;
                }

                FillDatesAveragesList(st->hList, st->stStart, st->stEnd);
                ListView_AutoSizeToHeaderAndContent(st->hList);
            }
            return 0;

        case IDCANCEL: // Close
            DestroyWindow(hWnd);
            return 0;

        case IDC_REPORTDATES_SAVE: // Save
            if (st) SaveListViewAsCsv(hWnd, st->hList, L"Report_ByDates.csv", L"Save By Dates Report As");
            return 0;
        
        case IDC_REPORTDATES_PRINT: // Print
            if (st) {
                // Ensure current values are synced and ordered
                DateTime_GetSystemtime(st->hStart, &st->stStart);
                DateTime_GetSystemtime(st->hEnd, &st->stEnd);
                if (CompareSystemTimes(st->stStart, st->stEnd) > 0) {
                    st->stEnd = st->stStart;
                }
                const std::wstring title = L"By Dates: " + FormatYmd(st->stStart) + L" — " + FormatYmd(st->stEnd);
                PrintListView(hWnd, st->hList, title.c_str());
            }
            return 0;
        }
        break;

    case WM_DESTROY:
    {
        HWND owner = GetWindow(hWnd, GW_OWNER);
        if (owner && IsWindow(owner)) {
            EnableWindow(owner, TRUE);
            ShowWindow(owner, SW_RESTORE);
            SetForegroundWindow(owner);
        }
        return 0;
    }
    case WM_NCDESTROY:
        if (st) {
            if (st->ownsFont && st->hFont) { DeleteObject(st->hFont); st->hFont = nullptr; }
            SetWindowLongPtrW(hWnd, GWLP_USERDATA, 0);
            delete st;
        }
        return 0;
    }
    return DefWindowProcW(hWnd, msg, wParam, lParam);
}
        
// Add this function definition near the other Show*Window functions

static void ShowReportDatesWindow(HWND owner)
{
    static ATOM s_atom = 0;
    if (!s_atom)
    {
        WNDCLASSEXW wc{};
        wc.cbSize = sizeof(wc);
        wc.style = CS_DBLCLKS;
        wc.lpfnWndProc = ReportDatesWndProc;
        wc.hInstance = hInst;
        wc.hCursor = LoadCursor(nullptr, IDC_ARROW);
        wc.hIcon = LoadIcon(nullptr, IDI_APPLICATION);
        wc.hbrBackground = (HBRUSH)(COLOR_BTNFACE + 1);
        wc.lpszClassName = L"BP_ReportDatesWindow";
        s_atom = RegisterClassExW(&wc);
    }

    const DWORD style = WS_OVERLAPPED | WS_CAPTION | WS_SYSMENU | WS_CLIPCHILDREN | WS_CLIPSIBLINGS;

    HWND hWnd = CreateWindowExW(WS_EX_DLGMODALFRAME,
        L"BP_ReportDatesWindow", L"Report - By Dates",
        style,
        CW_USEDEFAULT, CW_USEDEFAULT, 620, 200,
        owner, nullptr, hInst, nullptr);

    if (!hWnd) {
        DWORD err = GetLastError();
        wchar_t msg[128];
        swprintf_s(msg, L"Report Dates window failed (err=%lu).", err);
        MessageBoxW(owner, msg, szTitle, MB_OK | MB_ICONERROR);
        return;
    }

    EnableWindow(owner, FALSE);
    ShowWindow(hWnd, SW_SHOW);
    UpdateWindow(hWnd);

    // Modal-like pump so Enter/Esc work regardless of focused child
    MSG msg;
    while (IsWindow(hWnd) && GetMessageW(&msg, nullptr, 0, 0))
    {
        if (msg.message == WM_KEYDOWN &&
            (msg.hwnd == hWnd || IsChild(hWnd, msg.hwnd)))
        {
            if (msg.wParam == VK_RETURN || msg.wParam == VK_ESCAPE) {
                SendMessageW(hWnd, WM_COMMAND, IDCANCEL, 0);
                continue;
            }
        }
        if (!IsDialogMessageW(hWnd, &msg)) {
            TranslateMessage(&msg);
            DispatchMessageW(&msg);
        }
    }

    // Restore and activate owner
    if (owner && IsWindow(owner)) {
        EnableWindow(owner, TRUE);
        ShowWindow(owner, SW_RESTORE);
        SetForegroundWindow(owner);
    }
}