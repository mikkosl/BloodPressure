// BloodPressure.cpp : Defines the entry point for the application.
//
#include "framework.h"
#include "BloodPressure.h"
#include "Database.h"

#include <commdlg.h>
#pragma comment(lib, "Comdlg32.lib")

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
static void         CloseDatabaseDialog(HWND owner); // <-- add

static LRESULT CALLBACK AddReadingWndProc(HWND hWnd, UINT msg, WPARAM wParam, LPARAM lParam);
static void ShowAddReadingDialog(HWND owner);
static void ShowEditReadingDialog(HWND owner, const Reading& r);
static bool PromptRowNumber(HWND owner, int& outRow);

// Helpers
#ifndef WIDEN
#define WIDEN2(x) L##x
#define WIDEN(x) WIDEN2(x)
#endif

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

int APIENTRY wWinMain(_In_ HINSTANCE hInstance,
                     _In_opt_ HINSTANCE hPrevInstance,
                     _In_ LPWSTR    lpCmdLine,
                     _In_ int       nCmdShow)
{
    UNREFERENCED_PARAMETER(hPrevInstance);
    UNREFERENCED_PARAMETER(lpCmdLine);

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
        PostQuitMessage(0);
        break;
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
