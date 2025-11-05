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
#include <shlobj.h>
#include <KnownFolders.h>
#include <windowsx.h>
#include <windows.h>
#pragma comment(lib, "shell32.lib")
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

// Global Variables:
HINSTANCE hInst;                                // current instance
WCHAR szTitle[MAX_LOADSTRING];                  // The title bar text
WCHAR szWindowClass[MAX_LOADSTRING];            // the main window class name

static std::unique_ptr<Database> g_db;
static HWND g_mainWnd = nullptr;                // new: remember main window
static int g_rowPromptResult = -1;

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
static bool PromptRowNumber(HWND owner, int& outRow);
static void ShowReportDatesResultsWindow(HWND owner, const SYSTEMTIME& stStart, const SYSTEMTIME& stEnd);
static LRESULT CALLBACK ReportDatesResultsWndProc(HWND, UINT, WPARAM, LPARAM);

// Helpers
#ifndef WIDEN
#define WIDEN2(x) L##x
#define WIDEN(x) WIDEN2(x)
#endif

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

static std::wstring GetDatabasePath()
{
    // Use the directory of this source file (i.e., the project folder)
    const wchar_t* srcPath = WIDEN(__FILE__);
    std::wstring dir(srcPath);
    size_t pos = dir.find_last_of(L"\\/");
    if (pos != std::wstring::npos)
        dir.erase(pos + 1);
    return dir + L"BloodPressure.db";
}

// Helpers to compute local time from UTC ISO and averages
static bool TryParseUtcIsoToLocalTm(const std::wstring& isoUtc, std::tm& outLocal)
{
    int Y=0,M=0,D=0,h=0,m=0,s=0;
    if (swscanf_s(isoUtc.c_str(), L"%d-%d-%dT%d:%d:%d", &Y, &M, &D, &h, &m, &s) != 6)
        return false;

    std::tm tmUtc{};
    tmUtc.tm_year = Y - 1900;
    tmUtc.tm_mon  = M - 1;
    tmUtc.tm_mday = D;
    tmUtc.tm_hour = h;
    tmUtc.tm_min  = m;
    tmUtc.tm_sec  = s;

    time_t t = _mkgmtime(&tmUtc);
    if (t == (time_t)-1) return false;

    std::tm tmLocal{};
    if (localtime_s(&tmLocal, &t) != 0) return false;
    outLocal = tmLocal;
    return true;
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
                wchar_t dateHeader[] = L"Date (Local)";
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
                TextOutW(hdc, 10, y, msg1, lstrlenW(msg1));
                y += 17;
                TextOutW(hdc, 10, y, msg2, lstrlenW(msg2));
                y += 17;
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
            // g_rowPromptResult = st->result; // remove: not related to this dialog
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
        if (msg.message == WM_KEYDOWN && msg.wParam == VK_ESCAPE &&
            (msg.hwnd == hDlg || IsChild(hDlg, msg.hwnd)))
        {
            DestroyWindow(hDlg);
            continue;
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
        if (msg.message == WM_KEYDOWN && msg.wParam == VK_ESCAPE &&
            (msg.hwnd == hDlg || IsChild(hDlg, msg.hwnd)))
        {
            DestroyWindow(hDlg);
            continue;
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

// Simple prompt window to get a row number from the user
static LRESULT CALLBACK RowPromptWndProc(HWND hWnd, UINT msg, WPARAM wParam, LPARAM lParam)
{
    struct RowPromptState { HWND owner{}, hEdit{}; int result{ -1 }; };
    RowPromptState* st = reinterpret_cast<RowPromptState*>(GetWindowLongPtrW(hWnd, GWLP_USERDATA));
    switch (msg)
    {
    case WM_CREATE:
    {
        auto cs = reinterpret_cast<LPCREATESTRUCT>(lParam);
        auto owner = reinterpret_cast<HWND>(cs->lpCreateParams);
        auto* s = new RowPromptState();
        s->owner = owner;
        SetWindowLongPtrW(hWnd, GWLP_USERDATA, (LONG_PTR)s);

        HFONT hFont = (HFONT)GetStockObject(DEFAULT_GUI_FONT);

        CreateWindowExW(0, L"STATIC", L"Row No:",
            WS_CHILD | WS_VISIBLE, 12, 12, 60, 20, hWnd, nullptr, hInst, nullptr);
        s->hEdit = CreateWindowExW(WS_EX_CLIENTEDGE, L"EDIT", L"",
            WS_CHILD | WS_VISIBLE | ES_NUMBER | WS_TABSTOP,
            80, 10, 140, 24, hWnd, (HMENU)1001, hInst, nullptr);
        CreateWindowExW(0, L"BUTTON", L"OK",
            WS_CHILD | WS_VISIBLE | WS_TABSTOP | BS_DEFPUSHBUTTON,
            80, 44, 60, 24, hWnd, (HMENU)IDOK, hInst, nullptr);
        CreateWindowExW(0, L"BUTTON", L"Cancel",
            WS_CHILD | WS_VISIBLE | WS_TABSTOP,
            160, 44, 60, 24, hWnd, (HMENU)IDCANCEL, hInst, nullptr);

        SendMessageW(s->hEdit, WM_SETFONT, (WPARAM)hFont, TRUE);
        SetFocus(s->hEdit);

        RECT rc{ 0,0, 260, 100 };
        AdjustWindowRectEx(&rc, WS_CAPTION | WS_SYSMENU, FALSE, WS_EX_DLGMODALFRAME);
        SetWindowPos(hWnd, nullptr, 0, 0, rc.right - rc.left, rc.bottom - rc.top, SWP_NOMOVE | SWP_NOZORDER);
        CenterToOwner(hWnd, owner);
    }
    return 0;

    case WM_COMMAND:
        if (LOWORD(wParam) == IDOK)
        {
            if (!st) break;
            wchar_t buf[32]{};
            GetWindowTextW(st->hEdit, buf, 32);
            int val = (int)wcstol(buf, nullptr, 10);
            if (val <= 0) {
                MessageBoxW(hWnd, L"Enter a positive number.", L"Validation", MB_OK | MB_ICONWARNING);
                return 0;
            }
            st->result = val;
            DestroyWindow(hWnd);
            return 0;
        }
        else if (LOWORD(wParam) == IDCANCEL)
        {
            if (st) st->result = -1;
            DestroyWindow(hWnd);
            return 0;
        }
        break;

    case WM_CLOSE:
        if (st) st->result = -1;
        DestroyWindow(hWnd);
        return 0;

    case WM_NCDESTROY:
        if (st)
        {
            if (st->owner) {
                EnableWindow(st->owner, TRUE);
                ShowWindow(st->owner, SW_RESTORE);
                SetForegroundWindow(st->owner);
            }
            // propagate the result to PromptRowNumber
            g_rowPromptResult = st->result;
            SetWindowLongPtrW(hWnd, GWLP_USERDATA, 0);
            delete st;
        }
        return 0;
    }
    return DefWindowProcW(hWnd, msg, wParam, lParam);
}

static bool PromptRowNumber(HWND owner, int& outRow)
{
    // Register class once
    static ATOM s_atom = 0;
    if (!s_atom)
    {
        WNDCLASSEXW wc{};
        wc.cbSize = sizeof(wc);
        wc.style = CS_DBLCLKS;
        wc.lpfnWndProc = RowPromptWndProc;
        wc.hInstance = hInst;
        wc.hCursor = LoadCursor(nullptr, IDC_ARROW);
        wc.hIcon = LoadIcon(nullptr, IDI_APPLICATION);
        wc.hbrBackground = (HBRUSH)(COLOR_BTNFACE + 1);
        wc.lpszClassName = L"BP_RowPrompt";
        s_atom = RegisterClassExW(&wc);
    }

    HWND hDlg = CreateWindowExW(WS_EX_DLGMODALFRAME,
        L"BP_RowPrompt", L"Edit Row",
        WS_CAPTION | WS_SYSMENU,
        CW_USEDEFAULT, CW_USEDEFAULT, 260, 100,
        owner, nullptr, hInst, owner);
    if (!hDlg) return false;

    EnableWindow(owner, FALSE);
    ShowWindow(hDlg, SW_SHOW);
    UpdateWindow(hDlg);

    // Modal-like loop
    MSG msg;
    while (IsWindow(hDlg) && GetMessageW(&msg, nullptr, 0, 0))
    {
        if (msg.message == WM_KEYDOWN && msg.wParam == VK_ESCAPE &&
            (msg.hwnd == hDlg || IsChild(hDlg, msg.hwnd)))
        {
            DestroyWindow(hDlg);
            continue;
        }

        if (!IsDialogMessageW(hDlg, &msg))
        {
            TranslateMessage(&msg);
            DispatchMessageW(&msg);
        }
    }

    // Use captured result
    if (g_rowPromptResult > 0)
    {
        outRow = g_rowPromptResult;
        g_rowPromptResult = -1;
        return true;
    }
    g_rowPromptResult = -1;
    return false;
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

        // OK button
        const int marginRA = 10;
        st->hClose = CreateWindowExW(0, L"BUTTON", L"OK",
            WS_CHILD | WS_VISIBLE | WS_TABSTOP | BS_DEFPUSHBUTTON,
            marginRA, 0, 80, 26, hWnd, (HMENU)IDOK, hInst, nullptr);
        if (!st->hClose) {
            DWORD err = GetLastError();
            wchar_t msg[128];
            swprintf_s(msg, L"OK button creation failed (err=%lu).", err);
            MessageBoxW(hWnd, msg, szTitle, MB_OK | MB_ICONERROR);
        } else {
            SendMessageW(st->hClose, WM_SETFONT, (WPARAM)st->hFont, TRUE);
            ShowWindow(st->hClose, SW_SHOW);
        }

        // ListView (table) for the averages
        st->hList = CreateWindowExW(WS_EX_CLIENTEDGE, WC_LISTVIEWW, L"",
            WS_CHILD | WS_VISIBLE | WS_TABSTOP | LVS_REPORT | LVS_SINGLESEL | LVS_SHOWSELALWAYS,
            marginRA, marginRA, 100, 100, hWnd, (HMENU)42001, hInst, nullptr);
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
            col.mask = LVCF_TEXT | LVCF_WIDTH | LVCF_SUBITEM;

            struct ColDef { const wchar_t* text; int width; } cols[] = {
                { L"Bucket",   160 },
                { L"N",         80 },
                { L"Avg Sys/Avg Dia",  100 },
                { L"Avg Pulse", 100 },
            };
            for (int i = 0; i < (int)(sizeof(cols) / sizeof(cols[0])); ++i) {
                col.pszText = const_cast<wchar_t*>(cols[i].text);
                col.cx = cols[i].width;
                col.iSubItem = i;
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
            }
            else if (st->hList) {
                LVITEMW it{};
                it.mask = LVIF_TEXT;
                it.iItem = 0;
                it.iSubItem = 0;
                it.pszText = const_cast<wchar_t*>(L"No data");
                ListView_InsertItemW(st->hList, &it);
            }
        }

        // In ReportDatesResultsWndProc -> case WM_CREATE, after the ListView is created and populated, add:
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

        RECT rcClient{};
        GetClientRect(hWnd, &rcClient);
        const int btnW = 80, btnH = 26;

        if (st->hClose) {
            SetWindowPos(st->hClose, nullptr, rcClient.right - (btnW + marginRA), rcClient.bottom - (btnH + marginRA),
                btnW, btnH, SWP_NOZORDER);
        }
        if (st->hList) {
            int listRight = rcClient.right - marginRA;
            int listBottom = (st->hClose ? (rcClient.bottom - (btnH + 2 * marginRA)) : (rcClient.bottom - marginRA));
            SetWindowPos(st->hList, nullptr, marginRA, marginRA,
                listRight - marginRA, listBottom - marginRA, SWP_NOZORDER);
        }

        CenterToOwner(hWnd, GetWindow(hWnd, GW_OWNER));
        EnsureWindowOnScreen(hWnd);    }

        return 0;

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
        if (st->hList) {
            int listRight = rc.right - margin;
            int listBottom = (st->hClose ? (rc.bottom - (btnH + 2 * margin)) : (rc.bottom - margin));
            SetWindowPos(st->hList, nullptr, margin, margin,
                listRight - margin, listBottom - margin, SWP_NOZORDER);

            // Optional: auto-size columns to header/content
            for (int i = 0; i < 5; ++i) {
                ListView_SetColumnWidth(st->hList, i, (i == 0) ? LVSCW_AUTOSIZE_USEHEADER : LVSCW_AUTOSIZE_USEHEADER);
            }
        }
    }
    return 0;

    case WM_COMMAND:
        if (LOWORD(wParam) == IDOK)
        {
            DestroyWindow(hWnd);
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

struct ReportDatesResultsInit
{
    SYSTEMTIME stStart{};
    SYSTEMTIME stEnd{};
};

struct ReportDatesResultsState
{
    HWND hwnd{};
    HFONT hFont{};
    bool ownsFont{};
    HWND hList{};
    HWND hClose{};
    std::vector<Reading> filtered;
};

static LRESULT CALLBACK ReportDatesResultsWndProc(HWND hWnd, UINT msg, WPARAM wParam, LPARAM lParam)
{
    ReportDatesResultsState* st = reinterpret_cast<ReportDatesResultsState*>(GetWindowLongPtrW(hWnd, GWLP_USERDATA));
    switch (msg)
    {
    case WM_CREATE:
    {
        auto cs = reinterpret_cast<LPCREATESTRUCT>(lParam);
        auto init = reinterpret_cast<const ReportDatesResultsInit*>(cs->lpCreateParams);

        st = new ReportDatesResultsState();
        st->hwnd = hWnd;
        SetWindowLongPtrW(hWnd, GWLP_USERDATA, (LONG_PTR)st);

        st->hFont = CreateUiFontForWindow(hWnd, 14, FW_SEMIBOLD, L"Segoe UI");
        st->ownsFont = (st->hFont != nullptr);
        if (!st->hFont) st->hFont = (HFONT)GetStockObject(DEFAULT_GUI_FONT);

        // Filter readings
        st->filtered.clear();
        if (g_db)
        {
            std::vector<Reading> all;
            if (g_db->GetAllReadings(all) && init)
            {
                const int startKey = DateKeyFromSystemTime(init->stStart);
                const int endKeyRaw = DateKeyFromSystemTime(init->stEnd);
                const int endKey = (startKey <= endKeyRaw) ? endKeyRaw : startKey;
                const int startK = (startKey <= endKeyRaw) ? startKey : endKeyRaw;
                for (const auto& r : all)
                {
                    std::tm local{};
                    if (!TryParseUtcIsoToLocalTm(r.tsUtc, local)) continue;
                    const int key = DateKeyFromTm(local);
                    if (key >= startK && key <= endKey)
                        st->filtered.push_back(r);
                }
            }
        }

        const int margin = 10;

        // OK button
        st->hClose = CreateWindowExW(0, L"BUTTON", L"OK",
            WS_CHILD | WS_VISIBLE | WS_TABSTOP | BS_DEFPUSHBUTTON,
            margin, 0, 80, 26, hWnd, (HMENU)IDOK, hInst, nullptr);
        if (st->hClose) SendMessageW(st->hClose, WM_SETFONT, (WPARAM)st->hFont, TRUE);

        // ListView: Morning/Evening/Overall averages for the filtered date range
        st->hList = CreateWindowExW(WS_EX_CLIENTEDGE, WC_LISTVIEWW, L"",
            WS_CHILD | WS_VISIBLE | WS_TABSTOP | LVS_REPORT | LVS_SINGLESEL | LVS_SHOWSELALWAYS,
            margin, margin, 100, 100, hWnd, (HMENU)43050, hInst, nullptr);
        if (st->hList)
        {
            SendMessageW(st->hList, WM_SETFONT, (WPARAM)st->hFont, TRUE);
            ListView_SetExtendedListViewStyle(st->hList,
                LVS_EX_FULLROWSELECT | LVS_EX_GRIDLINES | LVS_EX_DOUBLEBUFFER);

            // Compute averages from filtered readings (same logic as Report All)
            int cntM = 0, cntE = 0, cntO = 0;
            long sumSysM = 0, sumDiaM = 0, sumPulM = 0;
            long sumSysE = 0, sumDiaE = 0, sumPulE = 0;
            long sumSysO = 0, sumDiaO = 0, sumPulO = 0;

            for (const auto& r : st->filtered) {
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

            const int avgSysM = RoundAvg((int)sumSysM, cntM);
            const int avgDiaM = RoundAvg((int)sumDiaM, cntM);
            const int avgPulM = RoundAvg((int)sumPulM, cntM);

            const int avgSysE = RoundAvg((int)sumSysE, cntE);
            const int avgDiaE = RoundAvg((int)sumDiaE, cntE);
            const int avgPulE = RoundAvg((int)sumPulE, cntE);

            const int avgSysO = RoundAvg((int)sumSysO, cntO);
            const int avgDiaO = RoundAvg((int)sumDiaO, cntO);
            const int avgPulO = RoundAvg((int)sumPulO, cntO);

            // Columns: Bucket | N | Avg Sys/Avg Dia | Avg Pulse
            LVCOLUMNW col{};
            col.mask = LVCF_TEXT | LVCF_WIDTH | LVCF_SUBITEM;
            struct ColDef { const wchar_t* text; int width; } cols[] = {
                { L"Bucket",              160 },
                { L"N",                    80 },
                { L"Avg Sys/Avg Dia",     120 },
                { L"Avg Pulse",           100 },
            };
            for (int i = 0; i < (int)(sizeof(cols) / sizeof(cols[0])); ++i) {
                col.pszText = const_cast<wchar_t*>(cols[i].text);
                col.cx = cols[i].width;
                col.iSubItem = i;
                ListView_InsertColumn(st->hList, i, &col);
            }

            if (cntO > 0) {
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
            }
            else {
                LVITEMW it{};
                it.mask = LVIF_TEXT;
                it.iItem = 0;
                it.iSubItem = 0;
                it.pszText = const_cast<wchar_t*>(L"No data");
                ListView_InsertItemW(st->hList, &it);
            }

            for (int i = 0; i < 4; ++i) {
                ListView_SetColumnWidth(st->hList, i, LVSCW_AUTOSIZE_USEHEADER);
            }
        }

        // Layout: size window (keeps requested size) and place controls
        RECT rc{ 0,0, 760, 420 };
        AdjustWindowRectEx(&rc, WS_CAPTION | WS_SYSMENU, FALSE, WS_EX_DLGMODALFRAME);
        SetWindowPos(hWnd, nullptr, 0, 0, rc.right - rc.left, rc.bottom - rc.top, SWP_NOZORDER | SWP_NOMOVE);

        RECT rcClient{}; GetClientRect(hWnd, &rcClient);
        const int btnW = 80, btnH = 26;

        if (st->hClose) {
            SetWindowPos(st->hClose, nullptr,
                rcClient.right - (btnW + margin),
                rcClient.bottom - (btnH + margin),
                btnW, btnH, SWP_NOZORDER);
        }
        if (st->hList) {
            const int listRight = rcClient.right - margin;
            const int listBottom = (st->hClose ? (rcClient.bottom - (btnH + 2 * margin)) : (rcClient.bottom - margin));
            SetWindowPos(st->hList, nullptr,
                margin, margin,
                listRight - margin,
                listBottom - margin,
                SWP_NOZORDER);
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

        if (st->hClose) {
            SetWindowPos(st->hClose, nullptr,
                rc.right - (btnW + margin),
                rc.bottom - (btnH + margin),
                btnW, btnH, SWP_NOZORDER);
        }
        if (st->hList) {
            const int listRight = rc.right - margin;
            const int listBottom = (st->hClose ? (rc.bottom - (btnH + 2 * margin)) : (rc.bottom - margin));
            SetWindowPos(st->hList, nullptr,
                margin, margin,
                listRight - margin,
                listBottom - margin,
                SWP_NOZORDER);

            // Keep columns sensible after resize
            for (int i = 0; i < 4; ++i) {
                ListView_SetColumnWidth(st->hList, i, LVSCW_AUTOSIZE_USEHEADER);
            }
        }
    }
    return 0;

    case WM_COMMAND:
        if (LOWORD(wParam) == IDOK) {
            DestroyWindow(hWnd);
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

        st->hOk = CreateWindowExW(0, L"BUTTON", L"Refresh",
            WS_CHILD | WS_VISIBLE | WS_TABSTOP | BS_DEFPUSHBUTTON,
            margin, 0, 80, 26, hWnd, (HMENU)IDOK, hInst, nullptr);

        st->hCancel = CreateWindowExW(0, L"BUTTON", L"Close",
            WS_CHILD | WS_VISIBLE | WS_TABSTOP,
            margin, 0, 80, 26, hWnd, (HMENU)IDCANCEL, hInst, nullptr);

        // Apply fonts
        HWND cts[] = { st->hStart, st->hEnd, st->hOk, st->hCancel };
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
            LVCOLUMNW col{}; col.mask = LVCF_TEXT | LVCF_WIDTH | LVCF_SUBITEM;
            struct ColDef { const wchar_t* text; int width; } cols[] = {
                { L"Bucket",              160 },
                { L"N",                    80 },
                { L"Avg Sys/Avg Dia",     120 },
                { L"Avg Pulse",           100 },
            };
            for (int i = 0; i < (int)(sizeof(cols) / sizeof(cols[0])); ++i) {
                col.pszText = const_cast<wchar_t*>(cols[i].text);
                col.cx = cols[i].width;
                col.iSubItem = i;
                ListView_InsertColumn(st->hList, i, &col);
            }
            st->listInit = true;
        }

        // Grow window to comfortable size
        RECT rc{ 0,0, 760, 420 };
        AdjustWindowRectEx(&rc, WS_CAPTION | WS_SYSMENU, FALSE, WS_EX_DLGMODALFRAME);
        SetWindowPos(hWnd, nullptr, 0, 0, rc.right - rc.left, rc.bottom - rc.top, SWP_NOZORDER);

        // Place buttons bottom-right
        RECT rcClient{}; GetClientRect(hWnd, &rcClient);
        const int btnW = 80, btnH = 26;
        if (st->hOk)     SetWindowPos(st->hOk, nullptr, rcClient.right - (btnW * 2 + margin * 2), rcClient.bottom - (btnH + margin), btnW, btnH, SWP_NOZORDER);
        if (st->hCancel) SetWindowPos(st->hCancel, nullptr, rcClient.right - (btnW + margin), rcClient.bottom - (btnH + margin), btnW, btnH, SWP_NOZORDER);

        // Initial data
        FillDatesAveragesList(st->hList, st->stStart, st->stEnd);

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
        if (st->hOk)     SetWindowPos(st->hOk, nullptr, rc.right - (btnW * 2 + margin * 2), rc.bottom - (btnH + margin), btnW, btnH, SWP_NOZORDER);
        if (st->hCancel) SetWindowPos(st->hCancel, nullptr, rc.right - (btnW + margin), rc.bottom - (btnH + margin), btnW, btnH, SWP_NOZORDER);

        // List fills remaining space
        if (st->hList) {
            int top = margin * 2 + st->dtpH;
            int listRight = rc.right - margin;
            int listBottom = rc.bottom - (btnH + 2 * margin);
            SetWindowPos(st->hList, nullptr, margin, top, listRight - margin, max(0, listBottom - top), SWP_NOZORDER);

            for (int i = 0; i < 4; ++i) {
                ListView_SetColumnWidth(st->hList, i, LVSCW_AUTOSIZE_USEHEADER);
            }
        }
    }
    return 0;

    case WM_NOTIFY:
    {
        if (!st) break;
        auto nm = reinterpret_cast<LPNMHDR>(lParam);

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
            else // DTN_CLOSEUP: ensure we have the committed value
            {
                DateTime_GetSystemtime(st->hStart, &st->stStart);
                DateTime_GetSystemtime(st->hEnd, &st->stEnd);
            }

            FillDatesAveragesList(st->hList, st->stStart, st->stEnd);
            InvalidateRect(st->hList, nullptr, FALSE);
            return 0;
        }
    }
    break;

    case WM_COMMAND:
        switch (LOWORD(wParam))
        {
        case IDOK: // Refresh
            if (st) {
                DateTime_GetSystemtime(st->hStart, &st->stStart);
                DateTime_GetSystemtime(st->hEnd, &st->stEnd);
                FillDatesAveragesList(st->hList, st->stStart, st->stEnd);
            }
            return 0;

        case IDCANCEL: // Close
            DestroyWindow(hWnd);
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

static void ShowReportDatesResultsWindow(HWND owner, const SYSTEMTIME& stStart, const SYSTEMTIME& stEnd)
{
    static ATOM s_atom = 0;
    if (!s_atom)
    {
        WNDCLASSEXW wc{};
        wc.cbSize = sizeof(wc);
        wc.style = CS_DBLCLKS;
        wc.lpfnWndProc = ReportDatesResultsWndProc;
        wc.hInstance = hInst;
        wc.hCursor = LoadCursor(nullptr, IDC_ARROW);
        wc.hIcon = LoadIcon(nullptr, IDI_APPLICATION);
        wc.hbrBackground = (HBRUSH)(COLOR_BTNFACE + 1);
        wc.lpszClassName = L"BP_ReportDatesResultsWindow";
        s_atom = RegisterClassExW(&wc);
    }

    ReportDatesResultsInit init{ stStart, stEnd };

    const DWORD style = WS_OVERLAPPED | WS_CAPTION | WS_SYSMENU | WS_CLIPCHILDREN | WS_CLIPSIBLINGS;
    HWND hWnd = CreateWindowExW(WS_EX_DLGMODALFRAME,
        L"BP_ReportDatesResultsWindow", L"Report - Averages by Dates",
        style, CW_USEDEFAULT, CW_USEDEFAULT, 760, 420,
        owner, nullptr, hInst, &init);

    if (!hWnd) {
        DWORD err = GetLastError();
        wchar_t msg[128];
        swprintf_s(msg, L"Report Dates Results window failed (err=%lu).", err);
        MessageBoxW(owner, msg, szTitle, MB_OK | MB_ICONERROR);
        return;
    }

    EnableWindow(owner, FALSE);
    ShowWindow(hWnd, SW_SHOW);
    UpdateWindow(hWnd);
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
}
