// RoboCopyGUI.cpp
// Win32 drag-and-drop front-end for robocopy.exe
//
// Compile (Developer Command Prompt):
//   cl.exe RoboCopyGUI.cpp /W4 /std:c++17 /EHsc ^
//          /link user32.lib ole32.lib shell32.lib comctl32.lib ^
//               gdi32.lib comdlg32.lib
//
// Features
//   - Drag a folder from Explorer onto the SOURCE or DESTINATION drop zone
//   - Browse buttons as fallback
//   - Copy options panel (mirrors robocopy's most useful switches)
//   - Live command-line preview
//   - Run button launches robocopy in a console window (CREATE_NEW_CONSOLE)
//   - Output log captured and shown in a read-only edit control
//   - Resizable window; all controls scale with it

#define WIN32_LEAN_AND_MEAN
#define UNICODE
#define _UNICODE

#include <windows.h>
#include <windowsx.h>   // GET_X_LPARAM / GET_Y_LPARAM
#include <shlobj.h>     // IDropTarget, SHBrowseForFolder
#include <shellapi.h>
#include <ole2.h>
#include <commctrl.h>
#include <commdlg.h>
#include <wchar.h>
#include <string>
#include <vector>
#include <sstream>
#include <thread>
#include <atomic>

#pragma comment(lib, "user32.lib")
#pragma comment(lib, "ole32.lib")
#pragma comment(lib, "shell32.lib")
#pragma comment(lib, "comctl32.lib")
#pragma comment(lib, "gdi32.lib")
#pragma comment(lib, "comdlg32.lib")

// ─── control IDs ─────────────────────────────────────────────────────────────

enum : UINT {
    IDC_SRC_DROP    = 201,
    IDC_DST_DROP    = 202,
    IDC_SRC_EDIT    = 203,
    IDC_DST_EDIT    = 204,
    IDC_SRC_BROWSE  = 205,
    IDC_DST_BROWSE  = 206,

    // Option checkboxes
    IDC_OPT_S       = 210,  // /S  subdirs
    IDC_OPT_E       = 211,  // /E  all subdirs incl. empty
    IDC_OPT_MIR     = 212,  // /MIR
    IDC_OPT_MOV     = 213,  // /MOV
    IDC_OPT_MOVE    = 214,  // /MOVE
    IDC_OPT_Z       = 215,  // /Z  restartable
    IDC_OPT_B       = 216,  // /B  backup mode
    IDC_OPT_COPYALL = 217,  // /COPYALL
    IDC_OPT_SEC     = 218,  // /SEC
    IDC_OPT_PURGE   = 219,  // /PURGE
    IDC_OPT_NJH     = 220,  // /NJH no job header
    IDC_OPT_NJS     = 221,  // /NJS no job summary
    IDC_OPT_NFL     = 222,  // /NFL no file log
    IDC_OPT_NDL     = 223,  // /NDL no dir log
    IDC_OPT_NP      = 224,  // /NP  no progress
    IDC_OPT_XJ      = 225,  // /XJ  exclude junctions
    IDC_OPT_FFT     = 226,  // /FFT FAT file times

    // Spin-edit pairs
    IDC_RETRY_EDIT  = 230,
    IDC_RETRY_SPIN  = 231,
    IDC_WAIT_EDIT   = 232,
    IDC_WAIT_SPIN   = 233,

    IDC_EXTRA_EDIT  = 235,  // extra flags
    IDC_CMD_EDIT    = 236,  // command preview (read-only)
    IDC_LOG_EDIT    = 237,  // output log
    IDC_RUN_BTN     = 238,
    IDC_COPY_CMD    = 239,
    IDC_STATUS      = 240,
};

// ─── globals ─────────────────────────────────────────────────────────────────

static HWND  g_hWnd        = nullptr;
static HFONT g_hFont       = nullptr;
static HFONT g_hFontBold   = nullptr;
static HFONT g_hFontMono   = nullptr;

// Drop-zone window handles (owner-drawn static labels that accept drops)
static HWND  g_hSrcZone    = nullptr;
static HWND  g_hDstZone    = nullptr;

// Highlight state for drop zones
static bool  g_srcHover    = false;
static bool  g_dstHover    = false;

// Running state
static std::atomic<bool> g_running{ false };

// ─── colors ──────────────────────────────────────────────────────────────────

constexpr COLORREF CLR_BG         = RGB(245, 246, 250);
constexpr COLORREF CLR_ZONE_IDLE  = RGB(228, 232, 240);
constexpr COLORREF CLR_ZONE_HOV   = RGB(198, 218, 255);
constexpr COLORREF CLR_ZONE_BORD  = RGB(160, 174, 200);
constexpr COLORREF CLR_ZONE_BORD2 = RGB( 66, 133, 244);
constexpr COLORREF CLR_ACCENT     = RGB( 30, 100, 210);
constexpr COLORREF CLR_RUN        = RGB( 34, 139,  34);
constexpr COLORREF CLR_RUN_HOV    = RGB( 22, 110,  22);
constexpr COLORREF CLR_TEXT_DIM   = RGB(120, 130, 150);

// ─── forward declarations ─────────────────────────────────────────────────────

LRESULT CALLBACK WndProc(HWND, UINT, WPARAM, LPARAM);
LRESULT CALLBACK DropZoneProc(HWND, UINT, WPARAM, LPARAM);
LRESULT CALLBACK RunBtnProc(HWND, UINT, WPARAM, LPARAM);
void BuildCommandLine(std::wstring& out);
void UpdateCmdPreview();
void RunRobocopy();
void AppendLog(const wchar_t* text);
void BrowseForFolder(HWND hEdit);
void SetStatus(const wchar_t* msg);
HWND MakeCheck(HWND parent, const wchar_t* label, UINT id, int x, int y, int w);
HWND MakeSpin(HWND parent, UINT editId, UINT spinId, int x, int y, int val);

// ─── IDropTarget for the two drop zones ──────────────────────────────────────

class FolderDropTarget : public IDropTarget
{
public:
    explicit FolderDropTarget(HWND hEdit, bool& hover)
        : m_hEdit(hEdit), m_hover(hover) {}

    ULONG STDMETHODCALLTYPE AddRef()  override { return ++m_ref; }
    ULONG STDMETHODCALLTYPE Release() override
    {
        if (--m_ref == 0) { delete this; return 0; }
        return m_ref;
    }
    HRESULT STDMETHODCALLTYPE QueryInterface(REFIID riid, void** ppv) override
    {
        if (riid == IID_IUnknown || riid == IID_IDropTarget)
        {
            *ppv = static_cast<IDropTarget*>(this);
            AddRef(); return S_OK;
        }
        *ppv = nullptr; return E_NOINTERFACE;
    }

    HRESULT STDMETHODCALLTYPE DragEnter(IDataObject* pDO, DWORD,
                                        POINTL, DWORD* pdwEffect) override
    {
        *pdwEffect = HasFolder(pDO) ? DROPEFFECT_COPY : DROPEFFECT_NONE;
        m_hover = true;
        InvalidateRect(GetParent(m_hEdit), nullptr, FALSE);
        return S_OK;
    }
    HRESULT STDMETHODCALLTYPE DragOver(DWORD, POINTL, DWORD* pdwEffect) override
    {
        *pdwEffect = DROPEFFECT_COPY; return S_OK;
    }
    HRESULT STDMETHODCALLTYPE DragLeave() override
    {
        m_hover = false;
        InvalidateRect(GetParent(m_hEdit), nullptr, FALSE);
        return S_OK;
    }
    HRESULT STDMETHODCALLTYPE Drop(IDataObject* pDO, DWORD,
                                   POINTL, DWORD* pdwEffect) override
    {
        *pdwEffect = DROPEFFECT_NONE;
        m_hover = false;

        FORMATETC fe{ CF_HDROP, nullptr, DVASPECT_CONTENT, -1, TYMED_HGLOBAL };
        STGMEDIUM sm{};
        if (FAILED(pDO->GetData(&fe, &sm))) return S_OK;

        HDROP hDrop = static_cast<HDROP>(GlobalLock(sm.hGlobal));
        if (hDrop)
        {
            wchar_t path[MAX_PATH]{};
            // Use the first item dropped; if it's a file, take its directory
            if (DragQueryFileW(hDrop, 0, path, MAX_PATH))
            {
                DWORD attr = GetFileAttributesW(path);
                if (attr != INVALID_FILE_ATTRIBUTES &&
                    !(attr & FILE_ATTRIBUTE_DIRECTORY))
                {
                    // Strip filename → keep folder
                    wchar_t* last = wcsrchr(path, L'\\');
                    if (last) *last = L'\0';
                }
                SetWindowTextW(m_hEdit, path);
                UpdateCmdPreview();
                *pdwEffect = DROPEFFECT_COPY;
            }
            GlobalUnlock(sm.hGlobal);
        }
        ReleaseStgMedium(&sm);
        InvalidateRect(GetParent(m_hEdit), nullptr, FALSE);
        return S_OK;
    }

private:
    HWND  m_hEdit;
    bool& m_hover;
    LONG  m_ref = 1;

    static bool HasFolder(IDataObject* pDO)
    {
        FORMATETC fe{ CF_HDROP, nullptr, DVASPECT_CONTENT, -1, TYMED_HGLOBAL };
        return pDO->QueryGetData(&fe) == S_OK;
    }
};

// ─── Drop zone owner-draw ─────────────────────────────────────────────────────
// Each drop zone is a child STATIC we owner-draw, containing:
//   - coloured rounded-rect border
//   - icon + prompt text
//   - the edit box sits inside it

static WNDPROC g_origZoneProc = nullptr;

void PaintDropZone(HWND hwnd, HDC hdc, bool hover, const wchar_t* label)
{
    RECT rc; GetClientRect(hwnd, &rc);

    // Background
    HBRUSH hbr = CreateSolidBrush(hover ? CLR_ZONE_HOV : CLR_ZONE_IDLE);
    FillRect(hdc, &rc, hbr);
    DeleteObject(hbr);

    // Border (simulate rounded via rectangle + pen)
    HPEN hpen = CreatePen(PS_SOLID, 2,
                          hover ? CLR_ZONE_BORD2 : CLR_ZONE_BORD);
    HPEN hOld = (HPEN)SelectObject(hdc, hpen);
    HBRUSH hNull = (HBRUSH)GetStockObject(NULL_BRUSH);
    HBRUSH hOldBr = (HBRUSH)SelectObject(hdc, hNull);
    RoundRect(hdc, rc.left+1, rc.top+1, rc.right-1, rc.bottom-1, 12, 12);
    SelectObject(hdc, hOld); SelectObject(hdc, hOldBr);
    DeleteObject(hpen);

    // Label text at top-left of zone
    SetBkMode(hdc, TRANSPARENT);
    SetTextColor(hdc, hover ? CLR_ACCENT : CLR_TEXT_DIM);
    SelectObject(hdc, g_hFontBold);
    RECT textRc{ rc.left + 10, rc.top + 6, rc.right - 10, rc.top + 24 };
    DrawTextW(hdc, label, -1, &textRc, DT_LEFT | DT_SINGLELINE | DT_VCENTER);
}

LRESULT CALLBACK DropZoneProc(HWND hwnd, UINT msg, WPARAM wParam, LPARAM lParam)
{
    if (msg == WM_PAINT)
    {
        PAINTSTRUCT ps;
        HDC hdc = BeginPaint(hwnd, &ps);
        bool hover = (hwnd == g_hSrcZone) ? g_srcHover : g_dstHover;
        const wchar_t* lbl = (hwnd == g_hSrcZone)
            ? L"SOURCE  — drop folder here"
            : L"DESTINATION  — drop folder here";
        PaintDropZone(hwnd, hdc, hover, lbl);
        EndPaint(hwnd, &ps);
        return 0;
    }
    // The Browse button and edit control are children of this drop-zone
    // STATIC, not of the main window. Windows sends their notifications
    // (BN_CLICKED, EN_CHANGE, etc.) here first -- forward them up to
    // WndProc so IDC_SRC_BROWSE / IDC_DST_BROWSE actually get handled.
    if (msg == WM_COMMAND || msg == WM_CTLCOLOREDIT || msg == WM_CTLCOLORBTN)
    {
        return SendMessageW(GetParent(hwnd), msg, wParam, lParam);
    }
    return CallWindowProcW(g_origZoneProc, hwnd, msg, wParam, lParam);
}

// ─── Run button (custom draw for green colour) ────────────────────────────────

static WNDPROC g_origRunProc = nullptr;
static bool    g_runHover    = false;

LRESULT CALLBACK RunBtnProc(HWND hwnd, UINT msg, WPARAM wParam, LPARAM lParam)
{
    switch (msg)
    {
    case WM_MOUSEMOVE:
        if (!g_runHover) { g_runHover = true; InvalidateRect(hwnd, nullptr, FALSE); }
        {
            TRACKMOUSEEVENT tme{ sizeof(tme), TME_LEAVE, hwnd, 0 };
            TrackMouseEvent(&tme);
        }
        break;
    case WM_MOUSELEAVE:
        g_runHover = false;
        InvalidateRect(hwnd, nullptr, FALSE);
        break;
    case WM_PAINT:
    {
        PAINTSTRUCT ps;
        HDC hdc = BeginPaint(hwnd, &ps);
        RECT rc; GetClientRect(hwnd, &rc);

        bool disabled = !IsWindowEnabled(hwnd);
        COLORREF bg = disabled   ? RGB(160,160,160)
                    : g_runHover ? CLR_RUN_HOV
                                 : CLR_RUN;

        HBRUSH hbr = CreateSolidBrush(bg);
        HPEN hpen  = CreatePen(PS_SOLID, 1, disabled ? RGB(130,130,130) : RGB(20,100,20));
        SelectObject(hdc, hbr);
        SelectObject(hdc, hpen);
        RoundRect(hdc, rc.left, rc.top, rc.right, rc.bottom, 8, 8);
        DeleteObject(hbr); DeleteObject(hpen);

        SetBkMode(hdc, TRANSPARENT);
        SetTextColor(hdc, RGB(255,255,255));
        SelectObject(hdc, g_hFontBold);
        DrawTextW(hdc, g_running ? L"⏹  Stop" : L"▶  Run Robocopy",
                  -1, &rc, DT_CENTER | DT_VCENTER | DT_SINGLELINE);
        EndPaint(hwnd, &ps);
        return 0;
    }
    }
    return CallWindowProcW(g_origRunProc, hwnd, msg, wParam, lParam);
}

// ─── command-line builder ─────────────────────────────────────────────────────

void BuildCommandLine(std::wstring& out)
{
    auto getText = [](UINT id) -> std::wstring {
        wchar_t buf[MAX_PATH]{};
        // IDC_SRC_EDIT / IDC_DST_EDIT are children of the drop-zone STATIC
        // controls, not of g_hWnd -- GetDlgItemTextW only searches direct
        // children, so those two need the zone as the owner.
        HWND owner = (id == IDC_SRC_EDIT) ? g_hSrcZone
                   : (id == IDC_DST_EDIT) ? g_hDstZone
                   : g_hWnd;
        GetDlgItemTextW(owner, id, buf, MAX_PATH);
        return buf;
    };
    auto checked = [](UINT id) { return IsDlgButtonChecked(g_hWnd, id) == BST_CHECKED; };

    std::wostringstream cmd;
    cmd << L"robocopy.exe";

    std::wstring src = getText(IDC_SRC_EDIT);
    std::wstring dst = getText(IDC_DST_EDIT);

    auto quote = [](const std::wstring& s) -> std::wstring {
        if (s.empty()) return L"\"\"";
        if (s.find(L' ') == std::wstring::npos) return s;
        return L'"' + s + L'"';
    };

    cmd << L" " << quote(src);
    cmd << L" " << quote(dst);

    if (checked(IDC_OPT_S))       cmd << L" /S";
    if (checked(IDC_OPT_E))       cmd << L" /E";
    if (checked(IDC_OPT_MIR))     cmd << L" /MIR";
    if (checked(IDC_OPT_MOV))     cmd << L" /MOV";
    if (checked(IDC_OPT_MOVE))    cmd << L" /MOVE";
    if (checked(IDC_OPT_Z))       cmd << L" /Z";
    if (checked(IDC_OPT_B))       cmd << L" /B";
    if (checked(IDC_OPT_COPYALL)) cmd << L" /COPYALL";
    if (checked(IDC_OPT_SEC))     cmd << L" /SEC";
    if (checked(IDC_OPT_PURGE))   cmd << L" /PURGE";
    if (checked(IDC_OPT_NJH))     cmd << L" /NJH";
    if (checked(IDC_OPT_NJS))     cmd << L" /NJS";
    if (checked(IDC_OPT_NFL))     cmd << L" /NFL";
    if (checked(IDC_OPT_NDL))     cmd << L" /NDL";
    if (checked(IDC_OPT_NP))      cmd << L" /NP";
    if (checked(IDC_OPT_XJ))      cmd << L" /XJ";
    if (checked(IDC_OPT_FFT))     cmd << L" /FFT";

    // Retry / wait
    wchar_t rbuf[8]{}, wbuf[8]{};
    GetDlgItemTextW(g_hWnd, IDC_RETRY_EDIT, rbuf, 8);
    GetDlgItemTextW(g_hWnd, IDC_WAIT_EDIT,  wbuf, 8);
    int rv = _wtoi(rbuf), wv = _wtoi(wbuf);
    if (rv != 1) cmd << L" /R:" << rv;
    if (wv != 1) cmd << L" /W:" << wv;

    // Extra flags
    std::wstring extra = getText(IDC_EXTRA_EDIT);
    if (!extra.empty()) cmd << L" " << extra;

    out = cmd.str();
}

void UpdateCmdPreview()
{
    std::wstring cmd;
    BuildCommandLine(cmd);
    SetDlgItemTextW(g_hWnd, IDC_CMD_EDIT, cmd.c_str());
}

// ─── log helpers ─────────────────────────────────────────────────────────────

void AppendLog(const wchar_t* text)
{
    HWND hLog = GetDlgItem(g_hWnd, IDC_LOG_EDIT);
    int len = GetWindowTextLengthW(hLog);
    SendMessageW(hLog, EM_SETSEL, (WPARAM)len, (LPARAM)len);
    SendMessageW(hLog, EM_REPLACESEL, FALSE, (LPARAM)text);
}

void SetStatus(const wchar_t* msg)
{
    HWND hSb = GetDlgItem(g_hWnd, IDC_STATUS);
    if (hSb) SendMessageW(hSb, SB_SETTEXTW, 0, (LPARAM)msg);
}

// ─── run robocopy in background thread ───────────────────────────────────────

void RunRobocopy()
{
    std::wstring cmdline;
    BuildCommandLine(cmdline);

    // Create pipe for stdout/stderr capture
    HANDLE hReadPipe = nullptr, hWritePipe = nullptr;
    SECURITY_ATTRIBUTES sa{ sizeof(sa), nullptr, TRUE };
    if (!CreatePipe(&hReadPipe, &hWritePipe, &sa, 0))
    {
        AppendLog(L"[ERROR] Failed to create pipe.\r\n");
        g_running = false;
        EnableWindow(GetDlgItem(g_hWnd, IDC_RUN_BTN), TRUE);
        return;
    }
    SetHandleInformation(hReadPipe, HANDLE_FLAG_INHERIT, 0);

    STARTUPINFOW si{};
    si.cb          = sizeof(si);
    si.dwFlags     = STARTF_USESTDHANDLES | STARTF_USESHOWWINDOW;
    si.hStdOutput  = hWritePipe;
    si.hStdError   = hWritePipe;
    si.wShowWindow = SW_HIDE;

    PROCESS_INFORMATION pi{};
    std::wstring buf = cmdline; // CreateProcess may modify the string
    if (!CreateProcessW(nullptr, buf.data(), nullptr, nullptr, TRUE,
                        CREATE_NO_WINDOW, nullptr, nullptr, &si, &pi))
    {
        DWORD err = GetLastError();
        wchar_t msg[256];
        swprintf_s(msg, L"[ERROR] CreateProcess failed (%lu). "
                        L"Is robocopy.exe in PATH?\r\n", err);
        AppendLog(msg);
        CloseHandle(hReadPipe); CloseHandle(hWritePipe);
        g_running = false;
        InvalidateRect(GetDlgItem(g_hWnd, IDC_RUN_BTN), nullptr, FALSE);
        return;
    }

    CloseHandle(hWritePipe); // close our copy of the write end

    // Read pipe on this thread (we ARE the background thread)
    char ansi[4096];
    DWORD bytesRead = 0;
    while (ReadFile(hReadPipe, ansi, sizeof(ansi) - 1, &bytesRead, nullptr)
           && bytesRead > 0)
    {
        ansi[bytesRead] = '\0';
        // Convert to wide and post to main thread via SendMessage
        int wlen = MultiByteToWideChar(CP_OEMCP, 0, ansi, -1, nullptr, 0);
        std::wstring wline(wlen, L'\0');
        MultiByteToWideChar(CP_OEMCP, 0, ansi, -1, wline.data(), wlen);
        // Replace \n with \r\n for the edit control
        std::wstring fixed;
        for (wchar_t c : wline)
        {
            if (c == L'\n') fixed += L'\r';
            fixed += c;
        }
        // Marshal to UI thread
        SendMessageW(g_hWnd, WM_APP + 1, 0,
                     reinterpret_cast<LPARAM>(fixed.c_str()));
    }

    DWORD exitCode = 0;
    GetExitCodeProcess(pi.hProcess, &exitCode);
    CloseHandle(pi.hProcess); CloseHandle(pi.hThread);
    CloseHandle(hReadPipe);

    // Interpret robocopy exit code bitmask
    const wchar_t* verdict =
        (exitCode == 0) ? L"No files copied (source == destination)." :
        (exitCode  < 8) ? L"Success." :
                          L"One or more errors occurred.";

    wchar_t summary[128];
    swprintf_s(summary,
               L"\r\n── Exit code %lu  (%ls) ──\r\n", exitCode, verdict);
    SendMessageW(g_hWnd, WM_APP + 1, 0,
                 reinterpret_cast<LPARAM>(summary));

    wchar_t statusMsg[64];
    swprintf_s(statusMsg, L"Finished — exit code %lu.", exitCode);
    SendMessageW(g_hWnd, WM_APP + 2, 0,
                 reinterpret_cast<LPARAM>(statusMsg));

    g_running = false;
    PostMessageW(g_hWnd, WM_APP + 3, 0, 0); // re-enable Run button
}

// ─── browse helper ────────────────────────────────────────────────────────────
//
// NOTE: We deliberately do NOT use the legacy SHBrowseForFolderW/BROWSEINFO
// API here. With BIF_NEWDIALOGSTYLE it drives Explorer's old shell-namespace
// tree control, which on many machines (usually because of a cloud-sync or
// other shell-namespace extension) can recurse infinitely inside its own
// property-bag advise sink and blow the stack (visible in a debugger as
// repeated "propbagadv.cpp ... 80004005 Unspecified error" spam immediately
// before a c00000fd stack overflow). That is an OS/shell-extension bug, not
// something application code can guard against. IFileOpenDialog (the modern
// Common Item Dialog, available since Vista) uses a completely different,
// non-legacy code path and does not hit this.
void BrowseForFolder(HWND hEdit)
{
    IFileOpenDialog* pDlg = nullptr;
    HRESULT hr = CoCreateInstance(CLSID_FileOpenDialog, nullptr, CLSCTX_INPROC_SERVER,
                                   IID_PPV_ARGS(&pDlg));
    if (FAILED(hr)) return;

    DWORD opts = 0;
    pDlg->GetOptions(&opts);
    pDlg->SetOptions(opts | FOS_PICKFOLDERS | FOS_FORCEFILESYSTEM | FOS_PATHMUSTEXIST);
    pDlg->SetTitle(L"Select folder");

    hr = pDlg->Show(g_hWnd);
    if (SUCCEEDED(hr))
    {
        IShellItem* pItem = nullptr;
        if (SUCCEEDED(pDlg->GetResult(&pItem)) && pItem)
        {
            PWSTR pszPath = nullptr;
            if (SUCCEEDED(pItem->GetDisplayName(SIGDN_FILESYSPATH, &pszPath)) && pszPath)
            {
                SetWindowTextW(hEdit, pszPath);
                CoTaskMemFree(pszPath);
                UpdateCmdPreview();
            }
            pItem->Release();
        }
    }
    // IDCANCEL (user closed the dialog) returns HRESULT_FROM_WIN32(ERROR_CANCELLED);
    // that is an expected, non-error outcome and is silently ignored here.

    pDlg->Release();
}

// ─── helper: create a checkbox ───────────────────────────────────────────────

HWND MakeCheck(HWND parent, const wchar_t* label, UINT id, int x, int y, int w)
{
    HWND h = CreateWindowExW(0, L"BUTTON", label,
        WS_CHILD | WS_VISIBLE | BS_AUTOCHECKBOX,
        x, y, w, 20, parent, (HMENU)(uintptr_t)id,
        GetModuleHandleW(nullptr), nullptr);
    SendMessageW(h, WM_SETFONT, (WPARAM)g_hFont, TRUE);
    return h;
}

// ─── layout constants (adjusted in WM_SIZE) ───────────────────────────────────

constexpr int PAD    = 10;
constexpr int ZONEHT = 72;  // drop zone height
constexpr int EDITHT = 24;
constexpr int BTNHT  = 28;

// ─── WndProc ──────────────────────────────────────────────────────────────────

LRESULT CALLBACK WndProc(HWND hWnd, UINT msg, WPARAM wParam, LPARAM lParam)
{
    switch (msg)
    {
    // ── WM_APP messages from background thread ──────────────────────────────
    case WM_APP + 1: // append log text
        AppendLog(reinterpret_cast<const wchar_t*>(lParam));
        return 0;
    case WM_APP + 2: // set status
        SetStatus(reinterpret_cast<const wchar_t*>(lParam));
        return 0;
    case WM_APP + 3: // re-enable run button
        EnableWindow(GetDlgItem(hWnd, IDC_RUN_BTN), TRUE);
        InvalidateRect(GetDlgItem(hWnd, IDC_RUN_BTN), nullptr, FALSE);
        return 0;

    case WM_CREATE:
    {
        g_hFont = CreateFontW(-13, 0, 0, 0, FW_NORMAL, FALSE, FALSE, FALSE,
                              DEFAULT_CHARSET, OUT_DEFAULT_PRECIS,
                              CLIP_DEFAULT_PRECIS, CLEARTYPE_QUALITY,
                              DEFAULT_PITCH | FF_DONTCARE, L"Segoe UI");
        g_hFontBold = CreateFontW(-13, 0, 0, 0, FW_SEMIBOLD, FALSE, FALSE, FALSE,
                              DEFAULT_CHARSET, OUT_DEFAULT_PRECIS,
                              CLIP_DEFAULT_PRECIS, CLEARTYPE_QUALITY,
                              DEFAULT_PITCH | FF_DONTCARE, L"Segoe UI");
        g_hFontMono = CreateFontW(-12, 0, 0, 0, FW_NORMAL, FALSE, FALSE, FALSE,
                              DEFAULT_CHARSET, OUT_DEFAULT_PRECIS,
                              CLIP_DEFAULT_PRECIS, CLEARTYPE_QUALITY,
                              FIXED_PITCH | FF_DONTCARE, L"Consolas");

        HINSTANCE hInst = GetModuleHandleW(nullptr);

        // ── Source drop zone ────────────────────────────────────────────────
        g_hSrcZone = CreateWindowExW(0, L"STATIC", L"",
            WS_CHILD | WS_VISIBLE | SS_OWNERDRAW,
            PAD, PAD, 400, ZONEHT,
            hWnd, (HMENU)IDC_SRC_DROP, hInst, nullptr);

        HWND hSrcEdit = CreateWindowExW(WS_EX_CLIENTEDGE, L"EDIT", L"",
            WS_CHILD | WS_VISIBLE | ES_AUTOHSCROLL,
            PAD + 8, PAD + 30, 300, EDITHT,
            g_hSrcZone, (HMENU)IDC_SRC_EDIT, hInst, nullptr);
        SendMessageW(hSrcEdit, WM_SETFONT, (WPARAM)g_hFont, TRUE);

        HWND hSrcBrowse = CreateWindowExW(0, L"BUTTON", L"Browse…",
            WS_CHILD | WS_VISIBLE | BS_PUSHBUTTON,
            PAD + 316, PAD + 30, 70, EDITHT,
            g_hSrcZone, (HMENU)IDC_SRC_BROWSE, hInst, nullptr);
        SendMessageW(hSrcBrowse, WM_SETFONT, (WPARAM)g_hFont, TRUE);

        // Register drop target on the zone
        {
            FolderDropTarget* pDT = new FolderDropTarget(hSrcEdit, g_srcHover);
            RegisterDragDrop(g_hSrcZone, pDT);
            pDT->Release();
        }

        g_origZoneProc = (WNDPROC)SetWindowLongPtrW(
            g_hSrcZone, GWLP_WNDPROC, (LONG_PTR)DropZoneProc);

        // ── Destination drop zone ───────────────────────────────────────────
        g_hDstZone = CreateWindowExW(0, L"STATIC", L"",
            WS_CHILD | WS_VISIBLE | SS_OWNERDRAW,
            PAD, PAD * 2 + ZONEHT, 400, ZONEHT,
            hWnd, (HMENU)IDC_DST_DROP, hInst, nullptr);

        HWND hDstEdit = CreateWindowExW(WS_EX_CLIENTEDGE, L"EDIT", L"",
            WS_CHILD | WS_VISIBLE | ES_AUTOHSCROLL,
            PAD + 8, PAD + 30, 300, EDITHT,
            g_hDstZone, (HMENU)IDC_DST_EDIT, hInst, nullptr);
        SendMessageW(hDstEdit, WM_SETFONT, (WPARAM)g_hFont, TRUE);

        HWND hDstBrowse = CreateWindowExW(0, L"BUTTON", L"Browse…",
            WS_CHILD | WS_VISIBLE | BS_PUSHBUTTON,
            PAD + 316, PAD + 30, 70, EDITHT,
            g_hDstZone, (HMENU)IDC_DST_BROWSE, hInst, nullptr);
        SendMessageW(hDstBrowse, WM_SETFONT, (WPARAM)g_hFont, TRUE);

        {
            FolderDropTarget* pDT = new FolderDropTarget(hDstEdit, g_dstHover);
            RegisterDragDrop(g_hDstZone, pDT);
            pDT->Release();
        }

        SetWindowLongPtrW(g_hDstZone, GWLP_WNDPROC, (LONG_PTR)DropZoneProc);

        // ── Options label ───────────────────────────────────────────────────
        int optY = PAD * 3 + ZONEHT * 2;

        HWND hOptLabel = CreateWindowExW(0, L"STATIC", L"Copy Options",
            WS_CHILD | WS_VISIBLE,
            PAD, optY, 200, 18,
            hWnd, nullptr, hInst, nullptr);
        SendMessageW(hOptLabel, WM_SETFONT, (WPARAM)g_hFontBold, TRUE);
        optY += 22;

        // Checkboxes — two columns
        int col1 = PAD, col2 = PAD + 200;
        int cy = optY;
        int cw = 190;

        MakeCheck(hWnd, L"/S  Copy subdirectories",         IDC_OPT_S,    col1, cy, cw);
        MakeCheck(hWnd, L"/MOV  Move files",                IDC_OPT_MOV,  col2, cy, cw); cy += 22;
        MakeCheck(hWnd, L"/E  Include empty subdirs",       IDC_OPT_E,    col1, cy, cw);
        MakeCheck(hWnd, L"/MOVE  Move files+dirs",          IDC_OPT_MOVE, col2, cy, cw); cy += 22;
        MakeCheck(hWnd, L"/MIR  Mirror (= /E /PURGE)",      IDC_OPT_MIR,  col1, cy, cw);
        MakeCheck(hWnd, L"/PURGE  Delete extra dest files", IDC_OPT_PURGE,col2, cy, cw); cy += 22;
        MakeCheck(hWnd, L"/Z  Restartable mode",            IDC_OPT_Z,    col1, cy, cw);
        MakeCheck(hWnd, L"/B  Backup mode",                 IDC_OPT_B,    col2, cy, cw); cy += 22;
        MakeCheck(hWnd, L"/COPYALL  Copy all attributes",   IDC_OPT_COPYALL, col1, cy, cw);
        MakeCheck(hWnd, L"/SEC  Copy with security",        IDC_OPT_SEC,  col2, cy, cw); cy += 22;
        MakeCheck(hWnd, L"/XJ  Exclude junctions",          IDC_OPT_XJ,   col1, cy, cw);
        MakeCheck(hWnd, L"/FFT  FAT file times",            IDC_OPT_FFT,  col2, cy, cw); cy += 22;
        MakeCheck(hWnd, L"/NJH  No job header",             IDC_OPT_NJH,  col1, cy, cw);
        MakeCheck(hWnd, L"/NJS  No job summary",            IDC_OPT_NJS,  col2, cy, cw); cy += 22;
        MakeCheck(hWnd, L"/NFL  No file list",              IDC_OPT_NFL,  col1, cy, cw);
        MakeCheck(hWnd, L"/NDL  No dir list",               IDC_OPT_NDL,  col2, cy, cw); cy += 22;
        MakeCheck(hWnd, L"/NP  No progress %",              IDC_OPT_NP,   col1, cy, cw);
        cy += 28;

        // Retry / Wait row
        auto makeLabel = [&](const wchar_t* t, int x, int y2, int w2) {
            HWND h = CreateWindowExW(0, L"STATIC", t,
                WS_CHILD | WS_VISIBLE, x, y2, w2, 18,
                hWnd, nullptr, hInst, nullptr);
            SendMessageW(h, WM_SETFONT, (WPARAM)g_hFont, TRUE);
        };

        makeLabel(L"Retries /R:", col1, cy + 3, 70);
        HWND hRetry = CreateWindowExW(WS_EX_CLIENTEDGE, L"EDIT", L"1",
            WS_CHILD | WS_VISIBLE | ES_NUMBER,
            col1 + 76, cy, 50, EDITHT,
            hWnd, (HMENU)IDC_RETRY_EDIT, hInst, nullptr);
        SendMessageW(hRetry, WM_SETFONT, (WPARAM)g_hFont, TRUE);
        HWND hRetrySpin = CreateWindowExW(0, UPDOWN_CLASSW, nullptr,
            WS_CHILD | WS_VISIBLE | UDS_ALIGNRIGHT | UDS_SETBUDDYINT | UDS_ARROWKEYS,
            0, 0, 0, 0,
            hWnd, (HMENU)IDC_RETRY_SPIN, hInst, nullptr);
        SendMessageW(hRetrySpin, UDM_SETBUDDY,  (WPARAM)hRetry, 0);
        SendMessageW(hRetrySpin, UDM_SETRANGE,  0, MAKELONG(999, 0));
        SendMessageW(hRetrySpin, UDM_SETPOS,    0, 1);

        makeLabel(L"Wait (s) /W:", col1 + 140, cy + 3, 76);
        HWND hWait = CreateWindowExW(WS_EX_CLIENTEDGE, L"EDIT", L"1",
            WS_CHILD | WS_VISIBLE | ES_NUMBER,
            col1 + 222, cy, 50, EDITHT,
            hWnd, (HMENU)IDC_WAIT_EDIT, hInst, nullptr);
        SendMessageW(hWait, WM_SETFONT, (WPARAM)g_hFont, TRUE);
        HWND hWaitSpin = CreateWindowExW(0, UPDOWN_CLASSW, nullptr,
            WS_CHILD | WS_VISIBLE | UDS_ALIGNRIGHT | UDS_SETBUDDYINT | UDS_ARROWKEYS,
            0, 0, 0, 0,
            hWnd, (HMENU)IDC_WAIT_SPIN, hInst, nullptr);
        SendMessageW(hWaitSpin, UDM_SETBUDDY,  (WPARAM)hWait, 0);
        SendMessageW(hWaitSpin, UDM_SETRANGE,  0, MAKELONG(999, 0));
        SendMessageW(hWaitSpin, UDM_SETPOS,    0, 1);
        cy += 32;

        // Extra flags
        makeLabel(L"Extra flags:", col1, cy + 3, 70);
        HWND hExtra = CreateWindowExW(WS_EX_CLIENTEDGE, L"EDIT", L"",
            WS_CHILD | WS_VISIBLE | ES_AUTOHSCROLL,
            col1 + 76, cy, 310, EDITHT,
            hWnd, (HMENU)IDC_EXTRA_EDIT, hInst, nullptr);
        SendMessageW(hExtra, WM_SETFONT, (WPARAM)g_hFontMono, TRUE);
        cy += 32;

        // Command preview
        makeLabel(L"Command:", col1, cy + 3, 65);
        HWND hCmd = CreateWindowExW(WS_EX_CLIENTEDGE, L"EDIT", L"",
            WS_CHILD | WS_VISIBLE | ES_READONLY | ES_AUTOHSCROLL,
            col1 + 76, cy, 310, EDITHT,
            hWnd, (HMENU)IDC_CMD_EDIT, hInst, nullptr);
        SendMessageW(hCmd, WM_SETFONT, (WPARAM)g_hFontMono, TRUE);

        // Copy button next to command
        HWND hCopyCmd = CreateWindowExW(0, L"BUTTON", L"Copy",
            WS_CHILD | WS_VISIBLE | BS_PUSHBUTTON,
            col1 + 393, cy, 50, EDITHT,
            hWnd, (HMENU)IDC_COPY_CMD, hInst, nullptr);
        SendMessageW(hCopyCmd, WM_SETFONT, (WPARAM)g_hFont, TRUE);
        cy += 36;

        // Run button
        HWND hRun = CreateWindowExW(0, L"BUTTON", L"▶  Run Robocopy",
            WS_CHILD | WS_VISIBLE | BS_PUSHBUTTON | BS_OWNERDRAW,
            col1, cy, 180, 34,
            hWnd, (HMENU)IDC_RUN_BTN, hInst, nullptr);
        SendMessageW(hRun, WM_SETFONT, (WPARAM)g_hFontBold, TRUE);
        g_origRunProc = (WNDPROC)SetWindowLongPtrW(
            hRun, GWLP_WNDPROC, (LONG_PTR)RunBtnProc);
        cy += 42;

        // Log
        makeLabel(L"Output", col1, cy, 50);
        HWND hClearLog = CreateWindowExW(0, L"BUTTON", L"Clear log",
            WS_CHILD | WS_VISIBLE | BS_PUSHBUTTON,
            col1 + 56, cy - 2, 80, 22,
            hWnd, (HMENU)241, hInst, nullptr);
        SendMessageW(hClearLog, WM_SETFONT, (WPARAM)g_hFont, TRUE);
        cy += 20;

        HWND hLog = CreateWindowExW(WS_EX_CLIENTEDGE, L"EDIT", L"",
            WS_CHILD | WS_VISIBLE | WS_VSCROLL |
            ES_MULTILINE | ES_READONLY | ES_AUTOVSCROLL,
            col1, cy, 430, 120,
            hWnd, (HMENU)IDC_LOG_EDIT, hInst, nullptr);
        SendMessageW(hLog, WM_SETFONT, (WPARAM)g_hFontMono, TRUE);

        // Status bar
        INITCOMMONCONTROLSEX icex{ sizeof(icex), ICC_BAR_CLASSES };
        InitCommonControlsEx(&icex);
        HWND hSb = CreateWindowExW(0, STATUSCLASSNAMEW, nullptr,
            WS_CHILD | WS_VISIBLE | SBARS_SIZEGRIP,
            0, 0, 0, 0,
            hWnd, (HMENU)IDC_STATUS, hInst, nullptr);
        (void)hSb;

        UpdateCmdPreview();
        SetStatus(L"Drop a folder onto a zone, or use Browse. Then click Run.");
        return 0;
    }

    case WM_COMMAND:
    {
        UINT id  = LOWORD(wParam);
        UINT ntf = HIWORD(wParam);

        if (id == IDC_SRC_BROWSE)
        {
            BrowseForFolder(GetDlgItem(g_hSrcZone, IDC_SRC_EDIT));
        }
        else if (id == IDC_DST_BROWSE)
        {
            BrowseForFolder(GetDlgItem(g_hDstZone, IDC_DST_EDIT));
        }
        else if (id == IDC_RUN_BTN)
        {
            if (g_running)
            {
                // Stop not directly supported via pipe; just note it
                SetStatus(L"Robocopy is still running in background.");
            }
            else
            {
                // Validate
                wchar_t src[MAX_PATH]{}, dst[MAX_PATH]{};
                GetDlgItemTextW(g_hSrcZone, IDC_SRC_EDIT, src, MAX_PATH);
                GetDlgItemTextW(g_hDstZone, IDC_DST_EDIT, dst, MAX_PATH);
                if (wcslen(src) == 0 || wcslen(dst) == 0)
                {
                    MessageBoxW(hWnd, L"Please set both source and destination folders.",
                                L"RoboCopy GUI", MB_ICONWARNING);
                    break;
                }
                SetDlgItemTextW(hWnd, IDC_LOG_EDIT, L"");
                AppendLog(L"── Starting robocopy ──\r\n");
                g_running = true;
                EnableWindow(GetDlgItem(hWnd, IDC_RUN_BTN), FALSE);
                InvalidateRect(GetDlgItem(hWnd, IDC_RUN_BTN), nullptr, FALSE);
                SetStatus(L"Running…");
                std::thread(RunRobocopy).detach();
            }
        }
        else if (id == IDC_COPY_CMD)
        {
            std::wstring cmd;
            BuildCommandLine(cmd);
            if (OpenClipboard(hWnd))
            {
                EmptyClipboard();
                size_t bytes = (cmd.size() + 1) * sizeof(wchar_t);
                HGLOBAL hg = GlobalAlloc(GMEM_MOVEABLE, bytes);
                if (hg)
                {
                    memcpy(GlobalLock(hg), cmd.c_str(), bytes);
                    GlobalUnlock(hg);
                    SetClipboardData(CF_UNICODETEXT, hg);
                }
                CloseClipboard();
                SetStatus(L"Command copied to clipboard.");
            }
        }
        else if (id == 241) // clear log
        {
            SetDlgItemTextW(hWnd, IDC_LOG_EDIT, L"");
        }
        // Any checkbox / spin change → refresh preview
        else if (ntf == BN_CLICKED || ntf == EN_CHANGE)
        {
            UpdateCmdPreview();
        }
        return 0;
    }

    case WM_DRAWITEM:
    {
        // The run button is BS_OWNERDRAW — handled in its subclass proc
        return DefWindowProcW(hWnd, msg, wParam, lParam);
    }

    case WM_CTLCOLORSTATIC:
    {
        HDC hdc = (HDC)wParam;
        SetBkColor(hdc, CLR_BG);
        SetBkMode(hdc, TRANSPARENT);
        static HBRUSH hBgBrush = CreateSolidBrush(CLR_BG);
        return (LRESULT)hBgBrush;
    }

    case WM_CTLCOLOREDIT:
    case WM_CTLCOLORBTN:
        return (LRESULT)GetSysColorBrush(COLOR_WINDOW);

    case WM_ERASEBKGND:
    {
        HDC hdc = (HDC)wParam;
        RECT rc; GetClientRect(hWnd, &rc);
        HBRUSH hbr = CreateSolidBrush(CLR_BG);
        FillRect(hdc, &rc, hbr);
        DeleteObject(hbr);
        return 1;
    }

    case WM_SIZE:
    {
        HWND hSb = GetDlgItem(hWnd, IDC_STATUS);
        if (hSb) SendMessageW(hSb, WM_SIZE, 0, 0);

        int w = LOWORD(lParam);
        int zoneW = w - PAD * 2 - 2;

        // Resize source zone
        if (g_hSrcZone)
        {
            SetWindowPos(g_hSrcZone, nullptr,
                         PAD, PAD, zoneW, ZONEHT, SWP_NOZORDER);
            // Resize edit and browse inside it
            SetWindowPos(GetDlgItem(g_hSrcZone, IDC_SRC_EDIT), nullptr,
                         PAD + 8, PAD + 30, zoneW - 100, EDITHT, SWP_NOZORDER);
            SetWindowPos(GetDlgItem(g_hSrcZone, IDC_SRC_BROWSE), nullptr,
                         zoneW - 80, PAD + 30, 70, EDITHT, SWP_NOZORDER);
        }
        if (g_hDstZone)
        {
            SetWindowPos(g_hDstZone, nullptr,
                         PAD, PAD * 2 + ZONEHT, zoneW, ZONEHT, SWP_NOZORDER);
            SetWindowPos(GetDlgItem(g_hDstZone, IDC_DST_EDIT), nullptr,
                         PAD + 8, PAD + 30, zoneW - 100, EDITHT, SWP_NOZORDER);
            SetWindowPos(GetDlgItem(g_hDstZone, IDC_DST_BROWSE), nullptr,
                         zoneW - 80, PAD + 30, 70, EDITHT, SWP_NOZORDER);
        }
        return 0;
    }

    case WM_GETMINMAXINFO:
    {
        auto* mm = reinterpret_cast<MINMAXINFO*>(lParam);
        mm->ptMinTrackSize = { 500, 700 };
        return 0;
    }

    case WM_DESTROY:
        RevokeDragDrop(g_hSrcZone);
        RevokeDragDrop(g_hDstZone);
        DeleteObject(g_hFont);
        DeleteObject(g_hFontBold);
        DeleteObject(g_hFontMono);
        PostQuitMessage(0);
        return 0;
    }

    return DefWindowProcW(hWnd, msg, wParam, lParam);
}

// ─── WinMain ─────────────────────────────────────────────────────────────────

int WINAPI wWinMain(HINSTANCE hInst, HINSTANCE, PWSTR, int nCmdShow)
{
    if (FAILED(OleInitialize(nullptr)))
    {
        MessageBoxW(nullptr, L"OleInitialize failed.", L"Error", MB_ICONERROR);
        return 1;
    }

    INITCOMMONCONTROLSEX icex{ sizeof(icex), ICC_UPDOWN_CLASS | ICC_BAR_CLASSES };
    InitCommonControlsEx(&icex);

    WNDCLASSEXW wc{};
    wc.cbSize        = sizeof(wc);
    wc.style         = CS_HREDRAW | CS_VREDRAW;
    wc.lpfnWndProc   = WndProc;
    wc.hInstance     = hInst;
    wc.hCursor       = LoadCursorW(nullptr, IDC_ARROW);
    wc.hbrBackground = CreateSolidBrush(CLR_BG);
    wc.lpszClassName = L"RoboCopyGUI";
    wc.hIcon         = LoadIconW(nullptr, IDI_APPLICATION);
    RegisterClassExW(&wc);

    g_hWnd = CreateWindowExW(
        0,
        L"RoboCopyGUI",
        L"RoboCopy GUI",
        WS_OVERLAPPEDWINDOW,
        CW_USEDEFAULT, CW_USEDEFAULT,
        500, 780,
        nullptr, nullptr, hInst, nullptr);

    if (!g_hWnd) { OleUninitialize(); return 1; }

    ShowWindow(g_hWnd, nCmdShow);
    UpdateWindow(g_hWnd);

    MSG msg{};
    while (GetMessageW(&msg, nullptr, 0, 0))
    {
        TranslateMessage(&msg);
        DispatchMessageW(&msg);
    }

    OleUninitialize();
    return (int)msg.wParam;
}
