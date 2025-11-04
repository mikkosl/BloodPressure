// BloodPressure.cpp : Defines the entry point for the application.
//
#include "framework.h"
#include "BloodPressure.h"
#include "Database.h"

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

/* Menu command IDs
#define IDM_ABOUT          40001
#define IDM_ADD            40002
#define IDM_EDITROW        40003
#define IDM_EXIT           40004
*/
// Global Variables:
HINSTANCE hInst;                                // current instance
WCHAR szTitle[MAX_LOADSTRING];                  // The title bar text
WCHAR szWindowClass[MAX_LOADSTRING];            // the main window class name

static std::unique_ptr<Database> g_db;
static HWND g_mainWnd = nullptr;                // new: remember main window
static int g_rowPromptResult = -1;

// Forward declarations:
ATOM                MyRegisterClass(HINSTANCE hInstance);
BOOL                InitInstance(HINSTANCE, int);
LRESULT CALLBACK    WndProc(HWND, UINT, WPARAM, LPARAM);
INT_PTR CALLBACK    About(HWND, UINT, WPARAM, LPARAM);

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

// Parse "YYYY-MM-DDTHH:MM:SSZ" (UTC) and format as local time "YYYY-MM-DD HH:MM:SS"
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
    if (wcsftime(buf, 32, L"%Y-%m-%d %H:%M:%S", &tmLocal) == 0)
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
    CoInitializeEx(nullptr, COINIT_APARTMENTTHREADED);

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

    // Initialize SQLite database
    {
        const std::wstring dbPath = GetDatabasePath();
        g_db = std::make_unique<Database>(dbPath.c_str());
        if (!g_db->Initialize())
        {
            MessageBoxW(nullptr, L"Failed to initialize database.", szTitle, MB_ICONERROR | MB_OK);
        }
        // Ensure the window repaints after DB becomes available so readings show immediately
        if (g_mainWnd) InvalidateRect(g_mainWnd, nullptr, TRUE);
    }

    HACCEL hAccelTable = LoadAccelerators(hInstance, MAKEINTRESOURCE(IDC_BLOODPRESSURE));

    MSG msg;

    // Main message loop:
    while (GetMessage(&msg, nullptr, 0, 0))
    {
        if (!TranslateAccelerator(g_mainWnd, hAccelTable, &msg))
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
            case IDM_ADD: // Reading -> Add
                ShowAddReadingDialog(hWnd);
                InvalidateRect(hWnd, nullptr, TRUE);
                break;
            case IDM_EDITROW: // Reading -> Edit by row number
                {
                    int row = 0;
                    if (PromptRowNumber(hWnd, row) && row > 0)
                    {
                        // Fetch the same recent list used for display
                        std::vector<Reading> rows;
                        if (g_db && g_db->GetRecentReadings(50, rows) && row <= (int)rows.size())
                        {
                            ShowEditReadingDialog(hWnd, rows[row - 1]);
                            InvalidateRect(hWnd, nullptr, TRUE);
                        }
                        else
                        {
                            MessageBoxW(hWnd, L"Invalid row number.", szTitle, MB_OK | MB_ICONWARNING);
                        }
                    }
                }
                break;
            case IDM_EXIT:
                DestroyWindow(hWnd);
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
                AppendMenuW(hMenu, MF_STRING, IDM_EDITROW, L"Edit Row...");
                POINT pt{ GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam) };
                ClientToScreen(hWnd, &pt);
                TrackPopupMenu(hMenu, TPM_RIGHTBUTTON | TPM_LEFTALIGN, pt.x, pt.y, 0, hWnd, nullptr);
                DestroyMenu(hMenu);
            }
        }
        break;
    case WM_PAINT:
        {
            PAINTSTRUCT ps;
            HDC hdc = BeginPaint(hWnd, &ps);

            int y = 10;

            if (g_db)
            {
                // Count
                int count = 0;
                if (g_db->GetReadingCount(count))
                {
                    std::wstring msg = L"Readings in database: " + std::to_wstring(count);
                    TextOutW(hdc, 10, y, msg.c_str(), static_cast<int>(msg.length()));
                }
                else
                {
                    const wchar_t* err = L"Database not available.";
                    TextOutW(hdc, 10, y, err, lstrlenW(err));
                }
                y += 20;

                // Header
                HFONT hMono = (HFONT)GetStockObject(SYSTEM_FIXED_FONT);
                HGDIOBJ oldFont = SelectObject(hdc, hMono);
                const wchar_t* header = L"No  Date (Local)         Sys/Dia Pul  Note";
                TextOutW(hdc, 10, y, header, lstrlenW(header));
                y += 18;

                // Divider
                const wchar_t* div = L"----------------------------------------------";
                TextOutW(hdc, 10, y, div, lstrlenW(div));
                y += 18;

                // Rows: last 50 readings
                std::vector<Reading> rows;
                if (g_db->GetRecentReadings(50, rows))
                {
                   int idx = 1;
                    for (const auto& r : rows)
                    {
                        std::wstring note = TruncateForDisplay(r.note, 60);
                        std::wstring tsLocal = UtcIsoToLocalDisplay(r.tsUtc);

                        wchar_t line[512];
                        
                        if (r.diastolic >= 100) {
                           swprintf_s(line, L"%3d %-20s %3d/%3d%3d  %s",
                               idx, tsLocal.c_str(),
                               r.systolic, r.diastolic, r.pulse,
                               note.c_str());
						}
                        else {
                           swprintf_s(line, L"%3d %-20s %3d/%2d %3d  %s",
                               idx, tsLocal.c_str(),
                               r.systolic, r.diastolic, r.pulse,
                               note.c_str());
                        }

                        TextOutW(hdc, 10, y, line, (int)wcslen(line));
                        y += 18;
                        ++idx;
                        if (y > ps.rcPaint.bottom - 20) break; // avoid drawing off-screen
                    }
                    if (rows.empty())
                    {
                        const wchar_t* none = L"(No readings yet. Use Reading -> Add to create one.)";
                        TextOutW(hdc, 10, y, none, lstrlenW(none));
                        y += 18;
                    }
                }
                else
                {
                    const wchar_t* errRows = L"Failed to load readings.";
                    TextOutW(hdc, 10, y, errRows, lstrlenW(errRows));
                    y += 18;
                }

                // Restore font
                SelectObject(hdc, oldFont);
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
    int result{-1}; // <-- Add this member to fix C2039
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

            HWND edits[] = { st->hEditSys, st->hEditDia, st->hEditPulse, st->hEditNote, hOk };
            for (HWND e : edits) { if (e) SendMessageW(e, WM_SETFONT, (WPARAM)st->hFont, TRUE); }

            // Prefill in edit mode
            if (st->editMode && init)
            {
                wchar_t buf[32];
                swprintf_s(buf, L"%d", init->reading.systolic);
                SetWindowTextW(st->hEditSys, buf);
                swprintf_s(buf, L"%d", init->reading.diastolic);
                SetWindowTextW(st->hEditDia, buf);
                swprintf_s(buf, L"%d", init->reading.pulse);
                SetWindowTextW(st->hEditPulse, buf);
                SetWindowTextW(st->hEditNote, init->reading.note.c_str());
            }

            SetFocus(st->hEditSys);

            RECT rc{ 0,0, 360, y + 50 };
            AdjustWindowRectEx(&rc, WS_CAPTION | WS_SYSMENU, FALSE, WS_EX_DLGMODALFRAME);
            SetWindowPos(hWnd, nullptr, 0, 0, rc.right - rc.left, rc.bottom - rc.top, SWP_NOMOVE | SWP_NOZORDER);
            if (st->owner) CenterToOwner(hWnd, st->owner);
        }
        return 0;

    case WM_COMMAND:
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
            if (st->owner) EnableWindow(st->owner, TRUE);
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
