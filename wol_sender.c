/*
 * WOL Sender 1.0.1 - Wake-on-LAN client with reachability check.
 *
 * MIT License
 *
 * Copyright (c) 2026 Treblewolf
 *
 * Built with mingw32 gcc; targets Windows 9x and later via runtime
 * compatibility shims (OPENFILENAME size, common-controls init, ICMP).
 *
 *   gcc wolsender.v1.0.1.c resources.o -o wolsender.v1.0.1.exe \
 *       -mwindows -Os -s -march=i386 -mno-sse -mno-sse2 \
 *       -ffunction-sections -fdata-sections -Wl,--gc-sections \
 *       -lwsock32 -lcomdlg32 -lcomctl32 -lshell32
 *   upx --best --lzma wolsender.v1.0.1.exe
 */

#include <windows.h>
#include <windowsx.h>
#include <winsock.h>
#include <stdio.h>
#include <stdlib.h>
#include <string.h>
#include <commctrl.h>
#include <commdlg.h>
#include <shellapi.h>

#pragma comment(lib, "wsock32.lib")
#pragma comment(lib, "comdlg32.lib")
#pragma comment(lib, "comctl32.lib")
#pragma comment(lib, "shell32.lib")

/* ====================================================================
 *  Constants
 * ==================================================================== */

/* Control / menu / context / timer / message IDs */
#define ID_LISTVIEW             101
#define ID_BTN_WAKE             102
#define ID_EDIT_NAME            103
#define ID_EDIT_MAC             104
#define ID_BTN_ADD              105
#define ID_STATUSBAR            106
#define ID_GROUPBOX             107
#define ID_EDIT_COMMENT         108
#define ID_BTN_CHECK            109
#define ID_BTN_EDIT             110

#define ID_FILE_IMPORT          201
#define ID_FILE_EXPORT          202
#define ID_FILE_EXIT            203
#define ID_HELP_ABOUT           204
#define ID_HELP_HELP            205

#define ID_AUTO_NONE            211
#define ID_AUTO_1SEC            212
#define ID_AUTO_2SEC            213

#define ID_CTX_EDIT             301
#define ID_CTX_REMOVE           302

#define IDT_STATUS_TIMER        401
#define IDT_AUTO_CHECK_TIMER    402

#define IDI_APPICON             100

#define WM_CHECK_COMPLETE       (WM_APP + 1)

/* List view columns (insertion order) */
#define COL_NAME                0
#define COL_MAC                 1
#define COL_STATUS              2
#define COL_COMMENT             3
#define COL_COUNT               4

/* Row-color flag stored in the high 16 bits of LVITEM.lParam */
#define ROW_HI_SHIFT            16
#define ROW_HI_MASK             0xFFFFu
#define ROW_COLOR_GREEN         1u
#define ROW_COLOR_RED           2u

/* Status bar */
#define MAX_STATUS_PARTS        24
#define STATUS_CLEAR_MS_SINGLE  2000
#define STATUS_CLEAR_MS_MULTI   10000

/* Default window/column sizes (used until INI overrides) */
#define WIN_DEFAULT_W           420
#define WIN_DEFAULT_H           520
#define WIN_MIN_W               360
#define WIN_MIN_H               440

/* Persistence */
#define INI_FILE_NAME           "wol_contacts.ini"
#define INI_SECTION_CONTACTS    "Contacts"
#define INI_SECTION_COMMENTS    "Comments"
#define INI_SECTION_WINDOW      "Window"

#define PROJECT_URL             "https://github.com/Treblewolf/wol-sender"

/* Common colors */
#define COLOR_OK                RGB(0, 150, 0)
#define COLOR_BAD               RGB(200, 0, 0)
#define COLOR_INFO              RGB(0, 0, 200)
#define COLOR_NEUTRAL           RGB(0, 0, 0)

/* ====================================================================
 *  Types
 * ==================================================================== */

typedef struct {
    char     text[384];
    COLORREF color;
} StatusPartCell;

typedef struct {
    char col0[256];
    char col1[256];
    char col2[256];
    char col3[256];
} SortSnapRow;

typedef struct { char name[256]; } CheckJobItem;

typedef struct {
    HWND         hwnd;
    BOOL         manual;
    int          count;
    CheckJobItem items[1];
} CheckJob;

typedef struct {
    char name[256];
    char detail[256];
    BOOL ok;
} CheckResultItem;

typedef struct {
    BOOL            manual;
    int             count;
    CheckResultItem items[1];
} CheckResult;

typedef HANDLE (WINAPI *PFN_IcmpCreateFile)(void);
typedef BOOL   (WINAPI *PFN_IcmpCloseHandle)(HANDLE);
typedef DWORD  (WINAPI *PFN_IcmpSendEcho)(HANDLE, ULONG, LPVOID, WORD, LPVOID, LPVOID, DWORD, DWORD);

typedef struct {
    ULONG Address;
    ULONG Status;
} ICMP_ECHO_REPLY_HDR;

/* ====================================================================
 *  Static state (single main window; assigned during WM_CREATE)
 * ==================================================================== */

static char g_iniPath[MAX_PATH];

static StatusPartCell g_statusCells[MAX_STATUS_PARTS];
static int            g_statusCellCount = 0;
static BOOL           g_statusMultiCell = FALSE;
static int            g_statusWidths[MAX_STATUS_PARTS];

static int          g_sortColumn = 0;
static int          g_sortAscending = 1;
static SortSnapRow* g_sortSnap = NULL;
static int          g_sortSnapCap = 0;

static volatile LONG g_checkInProgress = 0;
static UINT          g_autoCheckMs = 0;

/* UI handles */
static HWND  g_hMain;
static HWND  g_hListView;
static HWND  g_hBtnEdit, g_hBtnCheck, g_hBtnWake, g_hBtnAdd;
static HWND  g_hLblName, g_hLblMac, g_hLblComment;
static HWND  g_hEditName, g_hEditMac, g_hEditComment;
static HWND  g_hGroupBox, g_hStatusBar;
static HMENU g_hAutoMenu;
static char  g_editingName[256];

/* ====================================================================
 *  Static text
 * ==================================================================== */

static const char CLASS_NAME[]  = "WOLAppClass";
static const char APP_TITLE[]   = "WOL Sender";

static const char ABOUT_MSG[] =
    "WOL Sender v.1.0.1 by Treblewolf\n\n"
    "Changelog:\n\n"
    "v1.0.1 - Status column, async checks, auto-refresh,\n"
    "         sortable headers, saved window/column sizes\n\n"
    "v1.0   - Initial release\n\n\n"
    "Get latest version on GitHub:\n\n"
    PROJECT_URL "\n\n"
    "Open this page in your browser?";

static const char HELP_MSG[] =
    "Check: uses Name/IP as a IPv4 address, DNS name, or NetBIOS name\n";

/* ====================================================================
 *  Forward declarations
 * ==================================================================== */

static void SetStatus(const char* text, COLORREF color, UINT resetMs);
static void SetStatusParts(const StatusPartCell* cells, int n, UINT resetMs);
static void StatusBarApplyParts(void);
static void SortListView(void);
static void StartCheck(BOOL manual, BOOL allIfNoSelection);
static void ApplyCheckResult(CheckResult* result);
static void UpdateAutoMenu(void);
static void LoadColumnWidths(void);
static void SaveColumnWidths(void);
static void LoadContacts(void);
static void LayoutMainWindow(int width, int height);

/* ====================================================================
 *  Networking - WOL magic packet
 * ==================================================================== */

static BOOL SendWakeOnLan(unsigned char* macAddress) {
    WSADATA wsaData;
    SOCKET  udpSocket;
    struct sockaddr_in targetAddress;
    char    magicPacket[102];
    int     i, bOptVal = 1;
    BOOL    success = FALSE;

    if (WSAStartup(MAKEWORD(1, 1), &wsaData) != 0) return FALSE;

    udpSocket = socket(AF_INET, SOCK_DGRAM, 0);
    if (udpSocket == INVALID_SOCKET) { WSACleanup(); return FALSE; }

    setsockopt(udpSocket, SOL_SOCKET, SO_BROADCAST, (char*)&bOptVal, sizeof(bOptVal));
    for (i = 0; i < 6; i++)  magicPacket[i] = (char)0xFF;
    for (i = 1; i <= 16; i++) memcpy(&magicPacket[i * 6], macAddress, 6);

    targetAddress.sin_family      = AF_INET;
    targetAddress.sin_port        = htons(9);
    targetAddress.sin_addr.s_addr = inet_addr("255.255.255.255");

    if (sendto(udpSocket, magicPacket, sizeof(magicPacket), 0,
               (struct sockaddr*)&targetAddress, sizeof(targetAddress)) != SOCKET_ERROR) {
        success = TRUE;
    }
    closesocket(udpSocket);
    WSACleanup();
    return success;
}

/* ====================================================================
 *  Networking - reachability probes
 *  Computer Name = DNS / NetBIOS name or IPv4 literal
 * ==================================================================== */

static BOOL ResolveHostToAddr(const char* name, struct in_addr* out) {
    unsigned long a = inet_addr(name);
    if (a != INADDR_NONE) {
        out->S_un.S_addr = a;
        return TRUE;
    }
    {
        struct hostent* he = gethostbyname(name);
        if (he == NULL || he->h_addr_list[0] == NULL) return FALSE;
        memcpy(&out->S_un.S_addr, he->h_addr_list[0], sizeof(unsigned long));
        return TRUE;
    }
}

static BOOL TryLoadIcmp(PFN_IcmpCreateFile* pC, PFN_IcmpCloseHandle* pCl,
                        PFN_IcmpSendEcho* pS, HMODULE* pMod) {
    HMODULE h = LoadLibraryA("icmp.dll");
    if (h == NULL) h = LoadLibraryA("IPHLPAPI.DLL");
    if (h == NULL) return FALSE;

    *pC  = (PFN_IcmpCreateFile)GetProcAddress(h, "IcmpCreateFile");
    *pCl = (PFN_IcmpCloseHandle)GetProcAddress(h, "IcmpCloseHandle");
    *pS  = (PFN_IcmpSendEcho)GetProcAddress(h, "IcmpSendEcho");
    if (*pC == NULL || *pCl == NULL || *pS == NULL) {
        FreeLibrary(h);
        return FALSE;
    }
    *pMod = h;
    return TRUE;
}

static BOOL TryIcmpPing(PFN_IcmpCreateFile pfnC, PFN_IcmpCloseHandle pfnCl,
                        PFN_IcmpSendEcho pfnS, struct in_addr addr) {
    HANDLE h;
    char   sendData[4];
    char   replyBuf[256];
    DWORD  n;
    ICMP_ECHO_REPLY_HDR* r;

    h = pfnC();
    if (h == INVALID_HANDLE_VALUE) return FALSE;
    n = pfnS(h, addr.S_un.S_addr, sendData, sizeof(sendData),
             NULL, replyBuf, sizeof(replyBuf), 2000);
    pfnCl(h);
    if (n == 0) return FALSE;
    r = (ICMP_ECHO_REPLY_HDR*)replyBuf;
    return (r->Status == 0);
}

static BOOL TcpPortReachable(struct in_addr addr, unsigned short port) {
    SOCKET             s;
    struct sockaddr_in sin;
    u_long             nb = 1;
    fd_set             wf, ef;
    struct timeval     tv;
    int                err, len, r;

    s = socket(AF_INET, SOCK_STREAM, IPPROTO_TCP);
    if (s == INVALID_SOCKET) return FALSE;

    ZeroMemory(&sin, sizeof(sin));
    sin.sin_family = AF_INET;
    sin.sin_addr   = addr;
    sin.sin_port   = htons(port);

    if (ioctlsocket(s, FIONBIO, &nb) != 0) { closesocket(s); return FALSE; }

    r = connect(s, (struct sockaddr*)&sin, sizeof(sin));
    if (r == 0) { closesocket(s); return TRUE; }

    err = WSAGetLastError();
    if (err != WSAEWOULDBLOCK && err != WSAEINPROGRESS) {
        closesocket(s);
        return FALSE;
    }

    FD_ZERO(&wf);
    FD_ZERO(&ef);
    FD_SET(s, &wf);
    FD_SET(s, &ef);
    tv.tv_sec  = 2;
    tv.tv_usec = 0;
    r = select(0, NULL, &wf, &ef, &tv);
    if (r <= 0) { closesocket(s); return FALSE; }

    err = 0;
    len = sizeof(err);
    if (getsockopt(s, SOL_SOCKET, SO_ERROR, (char*)&err, &len) != 0) {
        closesocket(s);
        return FALSE;
    }
    closesocket(s);
    return (err == 0 || err == WSAECONNREFUSED);
}

/* ====================================================================
 *  Per-target evaluators (Check / Wake)
 * ==================================================================== */

static BOOL CheckEvaluateOne(const char* computerName,
                             BOOL haveIcmp,
                             PFN_IcmpCreateFile pfnC, PFN_IcmpCloseHandle pfnCl, PFN_IcmpSendEcho pfnS,
                             char* detail, int detailLen) {
    struct in_addr addr;

    if (computerName == NULL || computerName[0] == '\0') {
        lstrcpyn(detail, "No computer name", detailLen);
        return FALSE;
    }
    if (!ResolveHostToAddr(computerName, &addr)) {
        lstrcpyn(detail, "Unknown host", detailLen);
        return FALSE;
    }
    if (haveIcmp && TryIcmpPing(pfnC, pfnCl, pfnS, addr)) {
        lstrcpyn(detail, "Responding (ICMP)", detailLen);
        return TRUE;
    }
    if (TcpPortReachable(addr, 445)) {
        lstrcpyn(detail, "Responding (TCP 445)", detailLen);
        return TRUE;
    }
    lstrcpyn(detail, "No reply (Off)", detailLen);
    return FALSE;
}

static BOOL WakeEvaluateOne(const char* macStr, char* detail, int detailLen) {
    int  values[6];
    BOOL parsed = FALSE;
    int  i;

    if (sscanf(macStr, "%x:%x:%x:%x:%x:%x",
               &values[0], &values[1], &values[2], &values[3], &values[4], &values[5]) == 6) parsed = TRUE;
    else if (sscanf(macStr, "%x-%x-%x-%x-%x-%x",
                    &values[0], &values[1], &values[2], &values[3], &values[4], &values[5]) == 6) parsed = TRUE;

    if (!parsed) {
        lstrcpyn(detail, "Error: Invalid MAC format.", detailLen);
        return FALSE;
    }
    {
        unsigned char mac[6];
        for (i = 0; i < 6; i++) mac[i] = (unsigned char)values[i];
        if (SendWakeOnLan(mac)) {
            lstrcpyn(detail, "WOL Magic Packet Sent!", detailLen);
            return TRUE;
        }
    }
    lstrcpyn(detail, "Error: Network failure.", detailLen);
    return FALSE;
}

/* ====================================================================
 *  ListView helpers
 * ==================================================================== */

static void SetRowResultColor(int itemIndex, BOOL ok) {
    LVITEM it = {0};
    DWORD  kind = ok ? ROW_COLOR_GREEN : ROW_COLOR_RED;
    DWORD  lp;

    if (itemIndex < 0) return;
    it.mask  = LVIF_PARAM;
    it.iItem = itemIndex;
    if (!ListView_GetItem(g_hListView, &it)) return;
    lp = (DWORD)it.lParam;
    lp = (kind << ROW_HI_SHIFT) | (lp & ROW_HI_MASK);
    it.lParam = (LONG)lp;
    ListView_SetItem(g_hListView, &it);
}

static int FindListItemByName(const char* name) {
    int  i, count = ListView_GetItemCount(g_hListView);
    char current[256];
    for (i = 0; i < count; i++) {
        ListView_GetItemText(g_hListView, i, COL_NAME, current, sizeof(current));
        if (stricmp(current, name) == 0) return i;
    }
    return -1;
}

static int CALLBACK LVCompareProc(LPARAM lp1, LPARAM lp2, LPARAM lpSort) {
    unsigned    idx1 = (unsigned)lp1 & ROW_HI_MASK;
    unsigned    idx2 = (unsigned)lp2 & ROW_HI_MASK;
    const char *s1, *s2;
    int         cmp;
    (void)lpSort;

    if (g_sortSnap == NULL) return 0;
    if (idx1 >= (unsigned)g_sortSnapCap || idx2 >= (unsigned)g_sortSnapCap) return 0;

    switch (g_sortColumn) {
        case COL_MAC:     s1 = g_sortSnap[idx1].col1; s2 = g_sortSnap[idx2].col1; break;
        case COL_STATUS:  s1 = g_sortSnap[idx1].col2; s2 = g_sortSnap[idx2].col2; break;
        case COL_COMMENT: s1 = g_sortSnap[idx1].col3; s2 = g_sortSnap[idx2].col3; break;
        default:          s1 = g_sortSnap[idx1].col0; s2 = g_sortSnap[idx2].col0; break;
    }
    cmp = stricmp(s1, s2);
    if (cmp == 0) cmp = (int)idx1 - (int)idx2;
    return g_sortAscending ? cmp : -cmp;
}

static void SortListView(void) {
    int    count = ListView_GetItemCount(g_hListView);
    int    i;
    LVITEM it = {0};

    if (count <= 1) return;

    if (g_sortSnapCap < count) {
        SortSnapRow* p = (SortSnapRow*)realloc(g_sortSnap, count * sizeof(SortSnapRow));
        if (p == NULL) return;
        g_sortSnap    = p;
        g_sortSnapCap = count;
    }

    for (i = 0; i < count; i++) {
        DWORD lp;
        ListView_GetItemText(g_hListView, i, COL_NAME,    g_sortSnap[i].col0, sizeof(g_sortSnap[i].col0));
        ListView_GetItemText(g_hListView, i, COL_MAC,     g_sortSnap[i].col1, sizeof(g_sortSnap[i].col1));
        ListView_GetItemText(g_hListView, i, COL_STATUS,  g_sortSnap[i].col2, sizeof(g_sortSnap[i].col2));
        ListView_GetItemText(g_hListView, i, COL_COMMENT, g_sortSnap[i].col3, sizeof(g_sortSnap[i].col3));

        it.mask     = LVIF_PARAM;
        it.iItem    = i;
        it.iSubItem = 0;
        if (!ListView_GetItem(g_hListView, &it)) continue;
        lp = (DWORD)it.lParam;
        lp = (lp & 0xFFFF0000u) | ((DWORD)i & ROW_HI_MASK);
        it.lParam = (LONG)lp;
        ListView_SetItem(g_hListView, &it);
    }

    ListView_SortItems(g_hListView, LVCompareProc, 0);
}

/* ====================================================================
 *  Status bar (owner-drawn, optional segmented)
 * ==================================================================== */

static void StatusBarApplyParts(void) {
    RECT rc;
    int  clientW, i, k, seg;

    GetClientRect(g_hMain, &rc);
    clientW = rc.right;
    if (clientW < 20) clientW = 400;

    k = g_statusCellCount;
    if (k <= 0) k = 1;

    if (!g_statusMultiCell || k == 1) {
        g_statusWidths[0] = -1;
        SendMessage(g_hStatusBar, SB_SETPARTS, 1, (LPARAM)g_statusWidths);
    } else {
        seg = clientW / k;
        if (seg < 40) seg = 40;
        for (i = 0; i < k - 1; i++) g_statusWidths[i] = seg * (i + 1);
        g_statusWidths[k - 1] = -1;
        SendMessage(g_hStatusBar, SB_SETPARTS, k, (LPARAM)g_statusWidths);
    }

    for (i = 0; i < k; i++)
        SendMessage(g_hStatusBar, SB_SETTEXT, SBT_OWNERDRAW | i,
                    (LPARAM)g_statusCells[i].text);
}

static void SetStatus(const char* text, COLORREF color, UINT resetMs) {
    g_statusMultiCell = FALSE;
    g_statusCellCount = 1;
    lstrcpyn(g_statusCells[0].text, text, sizeof(g_statusCells[0].text));
    g_statusCells[0].color = color;
    StatusBarApplyParts();
    if (resetMs > 0) SetTimer(g_hMain, IDT_STATUS_TIMER, resetMs, NULL);
    else             KillTimer(g_hMain, IDT_STATUS_TIMER);
}

static void SetStatusParts(const StatusPartCell* cells, int n, UINT resetMs) {
    int i;
    if (n <= 0) return;
    g_statusMultiCell = (n > 1);
    g_statusCellCount = n;
    for (i = 0; i < n && i < MAX_STATUS_PARTS; i++) g_statusCells[i] = cells[i];
    StatusBarApplyParts();
    if (resetMs > 0) SetTimer(g_hMain, IDT_STATUS_TIMER, resetMs, NULL);
    else             KillTimer(g_hMain, IDT_STATUS_TIMER);
}

/* ====================================================================
 *  Async check (worker thread + main-thread result handler)
 * ==================================================================== */

static DWORD WINAPI CheckThreadProc(LPVOID lpParam) {
    CheckJob*           job = (CheckJob*)lpParam;
    CheckResult*        result;
    HMODULE             hIcmp = NULL;
    PFN_IcmpCreateFile  pfnC  = NULL;
    PFN_IcmpCloseHandle pfnCl = NULL;
    PFN_IcmpSendEcho    pfnS  = NULL;
    BOOL                haveIcmp;
    WSADATA             wsaData;
    int                 i;

    result = (CheckResult*)malloc(sizeof(CheckResult) + (job->count - 1) * sizeof(CheckResultItem));
    if (result == NULL) {
        g_checkInProgress = 0;
        free(job);
        return 0;
    }
    result->manual = job->manual;
    result->count  = job->count;

    haveIcmp = TryLoadIcmp(&pfnC, &pfnCl, &pfnS, &hIcmp);

    if (WSAStartup(MAKEWORD(1, 1), &wsaData) != 0) {
        for (i = 0; i < job->count; i++) {
            lstrcpyn(result->items[i].name,   job->items[i].name, sizeof(result->items[i].name));
            lstrcpyn(result->items[i].detail, "Winsock init failed.", sizeof(result->items[i].detail));
            result->items[i].ok = FALSE;
        }
    } else {
        for (i = 0; i < job->count; i++) {
            lstrcpyn(result->items[i].name, job->items[i].name, sizeof(result->items[i].name));
            result->items[i].ok = CheckEvaluateOne(job->items[i].name,
                                                  haveIcmp, pfnC, pfnCl, pfnS,
                                                  result->items[i].detail,
                                                  sizeof(result->items[i].detail));
        }
        WSACleanup();
    }

    if (hIcmp) FreeLibrary(hIcmp);
    if (!PostMessage(job->hwnd, WM_CHECK_COMPLETE, 0, (LPARAM)result)) {
        free(result);
        g_checkInProgress = 0;
    }
    free(job);
    return 0;
}

static void StartCheck(BOOL manual, BOOL allIfNoSelection) {
    int    selected = 0;
    int    count    = ListView_GetItemCount(g_hListView);
    int    i, out = 0;
    CheckJob* job;
    DWORD  threadId;
    HANDLE hThread;

    if (g_checkInProgress) {
        if (manual)
            SetStatus("Checking already in progress...", COLOR_INFO, STATUS_CLEAR_MS_SINGLE);
        return;
    }
    g_checkInProgress = 1;

    for (i = -1; (i = ListView_GetNextItem(g_hListView, i, LVNI_SELECTED)) != -1; )
        selected++;

    if (!allIfNoSelection && selected == 0) {
        g_checkInProgress = 0;
        return;
    }
    if (selected > 0) count = selected;
    if (count <= 0) {
        g_checkInProgress = 0;
        if (manual) SetStatus("No contacts to check.", COLOR_BAD, STATUS_CLEAR_MS_SINGLE);
        return;
    }

    job = (CheckJob*)malloc(sizeof(CheckJob) + (count - 1) * sizeof(CheckJobItem));
    if (job == NULL) {
        g_checkInProgress = 0;
        if (manual) SetStatus("Check failed: out of memory.", COLOR_BAD, STATUS_CLEAR_MS_SINGLE);
        return;
    }
    job->hwnd   = g_hMain;
    job->manual = manual;
    job->count  = count;

    if (selected > 0) {
        for (i = -1; (i = ListView_GetNextItem(g_hListView, i, LVNI_SELECTED)) != -1; ) {
            ListView_GetItemText(g_hListView, i, COL_NAME,
                                 job->items[out].name, sizeof(job->items[out].name));
            out++;
        }
    } else {
        for (i = 0; i < count; i++)
            ListView_GetItemText(g_hListView, i, COL_NAME,
                                 job->items[i].name, sizeof(job->items[i].name));
    }

    if (manual) SetStatus("Checking...", COLOR_INFO, 0);

    hThread = CreateThread(NULL, 0, CheckThreadProc, job, 0, &threadId);
    if (hThread == NULL) {
        free(job);
        g_checkInProgress = 0;
        if (manual) SetStatus("Check failed: could not start worker.",
                              COLOR_BAD, STATUS_CLEAR_MS_SINGLE);
        return;
    }
    CloseHandle(hThread);
}

static void ApplyCheckResult(CheckResult* result) {
    char line[384];
    int  i;

    for (i = 0; i < result->count; i++) {
        int row = FindListItemByName(result->items[i].name);
        if (row >= 0) {
            SetRowResultColor(row, result->items[i].ok);
            ListView_SetItemText(g_hListView, row, COL_STATUS,
                                 (char*)result->items[i].detail);
        }

        if (result->manual && result->count == 1) {
            lstrcpyn(line, result->items[i].name, sizeof(line));
            lstrcat(line, ": ");
            if (lstrlen(line) + lstrlen(result->items[i].detail) < (int)sizeof(line))
                lstrcat(line, result->items[i].detail);
            SetStatus(line,
                      result->items[i].ok ? COLOR_OK : COLOR_BAD,
                      STATUS_CLEAR_MS_SINGLE);
        }
    }

    InvalidateRect(g_hListView, NULL, TRUE);

    if (result->manual && result->count > 1) {
        SetStatus("", COLOR_NEUTRAL, 0);
    }
    g_checkInProgress = 0;
}

/* ====================================================================
 *  Wake operations
 * ==================================================================== */

static void WakeOneComputer(int itemIndex, const char* computerName, const char* macStr) {
    char detail[256];
    char line[384];
    BOOL ok = WakeEvaluateOne(macStr, detail, sizeof(detail));

    lstrcpyn(line, computerName, sizeof(line));
    lstrcat(line, ": ");
    if (lstrlen(line) + lstrlen(detail) < (int)sizeof(line)) lstrcat(line, detail);

    SetRowResultColor(itemIndex, ok);
    SetStatus(line, ok ? COLOR_OK : COLOR_BAD, STATUS_CLEAR_MS_SINGLE);
}

static void WakeListSelection(void) {
    char           detail[256];
    char           line[384];
    StatusPartCell cells[MAX_STATUS_PARTS];
    int            nSel = 0, iPos = -1, ci = 0;

    while ((iPos = ListView_GetNextItem(g_hListView, iPos, LVNI_SELECTED)) != -1) nSel++;
    if (nSel == 0) return;

    iPos = -1;
    while ((iPos = ListView_GetNextItem(g_hListView, iPos, LVNI_SELECTED)) != -1) {
        char name[256], mac[256];
        BOOL ok;

        ListView_GetItemText(g_hListView, iPos, COL_NAME, name, sizeof(name));
        ListView_GetItemText(g_hListView, iPos, COL_MAC,  mac,  sizeof(mac));
        ok = WakeEvaluateOne(mac, detail, sizeof(detail));
        SetRowResultColor(iPos, ok);

        lstrcpyn(line, name, sizeof(line));
        lstrcat(line, ": ");
        if (lstrlen(line) + lstrlen(detail) < (int)sizeof(line)) lstrcat(line, detail);

        if (nSel == 1) {
            SetStatus(line, ok ? COLOR_OK : COLOR_BAD, STATUS_CLEAR_MS_SINGLE);
            return;
        }
        if (ci < MAX_STATUS_PARTS) {
            lstrcpyn(cells[ci].text, line, sizeof(cells[ci].text));
            cells[ci].color = ok ? COLOR_OK : COLOR_BAD;
            ci++;
        }
    }

    if (nSel > MAX_STATUS_PARTS) {
        SetStatus("Too many selections for status bar.", COLOR_BAD, STATUS_CLEAR_MS_MULTI);
        return;
    }
    SetStatusParts(cells, ci, STATUS_CLEAR_MS_MULTI);
}

/* ====================================================================
 *  INI persistence
 * ==================================================================== */

static void SetupIniPath(void) {
    char* lastSlash;
    GetModuleFileName(NULL, g_iniPath, MAX_PATH);
    lastSlash = strrchr(g_iniPath, '\\');
    if (lastSlash != NULL) strcpy(lastSlash + 1, INI_FILE_NAME);
}

static void LoadContacts(void) {
    char  buffer[32767]; /* Win9x doc limit for GetPrivateProfileSection */
    DWORD charsRead;

    ListView_DeleteAllItems(g_hListView);
    charsRead = GetPrivateProfileSection(INI_SECTION_CONTACTS, buffer, sizeof(buffer), g_iniPath);
    if (charsRead == 0) return;

    {
        char* current = buffer;
        int   index   = 0;
        while (*current != '\0') {
            int   originalLen = strlen(current);
            char* equalsSign  = strchr(current, '=');
            if (equalsSign != NULL) {
                char  mac[256];
                char  comment[256];
                char* name;
                char* value;
                char* tab;
                LVITEM lvi = {0};

                *equalsSign = '\0';
                name  = current;
                value = equalsSign + 1;

                /* Backward compat: legacy v1.0 stored "MAC\tComment" in
                 * the [Contacts] value. New format keeps only the MAC
                 * here and stores comments in [Comments]. */
                tab = strchr(value, '\t');
                if (tab != NULL) {
                    *tab = '\0';
                    lstrcpyn(mac, value, sizeof(mac));
                    lstrcpyn(comment, tab + 1, sizeof(comment));
                } else {
                    lstrcpyn(mac, value, sizeof(mac));
                    GetPrivateProfileString(INI_SECTION_COMMENTS, name, "",
                                            comment, sizeof(comment), g_iniPath);
                }

                lvi.mask     = LVIF_TEXT | LVIF_PARAM;
                lvi.iItem    = index;
                lvi.iSubItem = 0;
                lvi.pszText  = name;
                lvi.lParam   = (LPARAM)(DWORD)index;
                ListView_InsertItem(g_hListView, &lvi);
                ListView_SetItemText(g_hListView, index, COL_MAC,     mac);
                ListView_SetItemText(g_hListView, index, COL_STATUS,  "");
                ListView_SetItemText(g_hListView, index, COL_COMMENT, comment);
                index++;
            }
            current += originalLen + 1;
        }
    }
}

static void SaveContact(const char* name, const char* mac, const char* comment) {
    WritePrivateProfileString(INI_SECTION_CONTACTS, name, mac, g_iniPath);
    WritePrivateProfileString(INI_SECTION_COMMENTS, name,
                              (comment != NULL && comment[0] != '\0') ? comment : NULL,
                              g_iniPath);
}

static void RemoveContact(const char* name) {
    WritePrivateProfileString(INI_SECTION_CONTACTS, name, NULL, g_iniPath);
    WritePrivateProfileString(INI_SECTION_COMMENTS, name, NULL, g_iniPath);
}

static void LoadWindowSize(int* width, int* height) {
    int w = GetPrivateProfileInt(INI_SECTION_WINDOW, "Width",  WIN_DEFAULT_W, g_iniPath);
    int h = GetPrivateProfileInt(INI_SECTION_WINDOW, "Height", WIN_DEFAULT_H, g_iniPath);
    if (w < WIN_MIN_W) w = WIN_DEFAULT_W;
    if (h < WIN_MIN_H) h = WIN_DEFAULT_H;
    *width  = w;
    *height = h;
}

static void SaveWindowSize(void) {
    RECT rc;
    char buf[32];
    if (IsIconic(g_hMain)) return;
    GetWindowRect(g_hMain, &rc);
    wsprintf(buf, "%d", rc.right - rc.left);
    WritePrivateProfileString(INI_SECTION_WINDOW, "Width", buf, g_iniPath);
    wsprintf(buf, "%d", rc.bottom - rc.top);
    WritePrivateProfileString(INI_SECTION_WINDOW, "Height", buf, g_iniPath);
}

static const char* g_colWidthKeys[COL_COUNT] = {
    "Col0Width", "Col1Width", "Col2Width", "Col3Width"
};
static const int g_colWidthDefaults[COL_COUNT] = { 150, 120, 130, 150 };

static void LoadColumnWidths(void) {
    int i, w;
    for (i = 0; i < COL_COUNT; i++) {
        w = GetPrivateProfileInt(INI_SECTION_WINDOW, g_colWidthKeys[i],
                                 g_colWidthDefaults[i], g_iniPath);
        if (w < 30 || w > 2000) w = g_colWidthDefaults[i];
        ListView_SetColumnWidth(g_hListView, i, w);
    }
}

static void SaveColumnWidths(void) {
    int  i, w;
    char buf[32];
    for (i = 0; i < COL_COUNT; i++) {
        w = ListView_GetColumnWidth(g_hListView, i);
        if (w <= 0) continue;
        wsprintf(buf, "%d", w);
        WritePrivateProfileString(INI_SECTION_WINDOW, g_colWidthKeys[i], buf, g_iniPath);
    }
}

/* ====================================================================
 *  Win9x compatibility shims
 * ==================================================================== */

static DWORD GetOpenFileNameStructSize(void) {
    OSVERSIONINFO osvi;
    ZeroMemory(&osvi, sizeof(osvi));
    osvi.dwOSVersionInfoSize = sizeof(osvi);
    if (GetVersionEx(&osvi) && osvi.dwPlatformId == VER_PLATFORM_WIN32_WINDOWS)
        return OPENFILENAME_SIZE_VERSION_400;
    return sizeof(OPENFILENAME);
}

static BOOL InitCommonControlsCompat(void) {
    INITCOMMONCONTROLSEX icex;
    ZeroMemory(&icex, sizeof(icex));
    icex.dwSize = sizeof(INITCOMMONCONTROLSEX);
    icex.dwICC  = ICC_LISTVIEW_CLASSES | ICC_BAR_CLASSES;
    if (InitCommonControlsEx(&icex)) return TRUE;
    InitCommonControls();
    return TRUE;
}

/* ====================================================================
 *  File dialogs
 * ==================================================================== */

static void HandleImport(void) {
    OPENFILENAME ofn;
    char szFile[260] = {0};

    ZeroMemory(&ofn, sizeof(ofn));
    ofn.lStructSize = GetOpenFileNameStructSize();
    ofn.hwndOwner   = g_hMain;
    ofn.lpstrFile   = szFile;
    ofn.nMaxFile    = sizeof(szFile);
    ofn.lpstrFilter = "INI Files\0*.ini\0All Files\0*.*\0";
    ofn.nFilterIndex= 1;
    ofn.Flags       = OFN_PATHMUSTEXIST | OFN_FILEMUSTEXIST;

    if (GetOpenFileName(&ofn) == TRUE) {
        CopyFile(szFile, g_iniPath, FALSE);
        LoadContacts();
        SetStatus("Contacts imported.", COLOR_OK, STATUS_CLEAR_MS_SINGLE);
    }
}

static void HandleExport(void) {
    OPENFILENAME ofn;
    char szFile[260] = INI_FILE_NAME;

    ZeroMemory(&ofn, sizeof(ofn));
    ofn.lStructSize = GetOpenFileNameStructSize();
    ofn.hwndOwner   = g_hMain;
    ofn.lpstrFile   = szFile;
    ofn.nMaxFile    = sizeof(szFile);
    ofn.lpstrFilter = "INI Files\0*.ini\0All Files\0*.*\0";
    ofn.nFilterIndex= 1;
    ofn.Flags       = OFN_OVERWRITEPROMPT;
    ofn.lpstrDefExt = "ini";

    if (GetSaveFileName(&ofn) == TRUE) {
        CopyFile(g_iniPath, szFile, FALSE);
        SetStatus("Contacts exported.", COLOR_OK, STATUS_CLEAR_MS_SINGLE);
    }
}

/* ====================================================================
 *  Auto-refresh menu
 * ==================================================================== */

static void UpdateAutoMenu(void) {
    CheckMenuItem(g_hAutoMenu, ID_AUTO_NONE,
                  MF_BYCOMMAND | (g_autoCheckMs == 0    ? MF_CHECKED : MF_UNCHECKED));
    CheckMenuItem(g_hAutoMenu, ID_AUTO_1SEC,
                  MF_BYCOMMAND | (g_autoCheckMs == 1000 ? MF_CHECKED : MF_UNCHECKED));
    CheckMenuItem(g_hAutoMenu, ID_AUTO_2SEC,
                  MF_BYCOMMAND | (g_autoCheckMs == 2000 ? MF_CHECKED : MF_UNCHECKED));
}

static void SetAutoRefresh(UINT ms) {
    g_autoCheckMs = ms;
    if (ms == 0) KillTimer(g_hMain, IDT_AUTO_CHECK_TIMER);
    else         SetTimer(g_hMain, IDT_AUTO_CHECK_TIMER, ms, NULL);
    UpdateAutoMenu();
}

/* ====================================================================
 *  Layout
 *
 *  The "New Contact" group box is a transparent BS_GROUPBOX whose
 *  interior does not erase itself on move, so moving its child labels
 *  with bRepaint=TRUE leaves ghost pixels on vertical resize.
 *  We move the form children with bRepaint=FALSE and invalidate the
 *  whole bottom strip in one pass to force a clean repaint.
 * ==================================================================== */

static void LayoutMainWindow(int width, int height) {
    int  pad = 10;
    int  gbHeight = 110;
    int  btnHeight = 30;
    int  statusHeight, effectiveHeight;
    int  gbY, btnRowY, lvY, lvHeight, innerW;
    RECT rcStatus;

    SendMessage(g_hStatusBar, WM_SIZE, 0, 0);
    GetWindowRect(g_hStatusBar, &rcStatus);
    statusHeight    = rcStatus.bottom - rcStatus.top;
    effectiveHeight = height - statusHeight;

    gbY      = effectiveHeight - pad - gbHeight;
    btnRowY  = gbY - pad - btnHeight;
    lvY      = pad;
    lvHeight = btnRowY - pad - lvY;
    innerW   = width - (pad * 2);

    MoveWindow(g_hListView, pad, lvY, innerW, lvHeight, TRUE);

    {
        int btnGap = 6;
        int btnW   = (innerW - btnGap * 2) / 3;
        int x0 = pad;
        int x1 = pad + btnW + btnGap;
        int x2 = pad + (btnW + btnGap) * 2;
        int wLast = pad + innerW - x2;
        MoveWindow(g_hBtnEdit,  x0, btnRowY, btnW,  btnHeight, TRUE);
        MoveWindow(g_hBtnCheck, x1, btnRowY, btnW,  btnHeight, TRUE);
        MoveWindow(g_hBtnWake,  x2, btnRowY, wLast, btnHeight, TRUE);
    }

    {
        int  lblW   = 52;
        int  editX  = pad + 10 + lblW;
        int  editW  = innerW - (editX - pad) - 80;
        int  inputY1 = gbY + 22;
        int  inputY2 = gbY + 47;
        int  inputY3 = gbY + 72;
        RECT rcForm;

        MoveWindow(g_hGroupBox,    pad,             gbY,         innerW, gbHeight, FALSE);
        MoveWindow(g_hLblName,     pad + 10,        inputY1 + 3, lblW,   20,       FALSE);
        MoveWindow(g_hEditName,    editX,           inputY1,     editW,  20,       FALSE);
        MoveWindow(g_hLblMac,      pad + 10,        inputY2 + 3, lblW,   20,       FALSE);
        MoveWindow(g_hEditMac,     editX,           inputY2,     editW,  20,       FALSE);
        MoveWindow(g_hLblComment,  pad + 10,        inputY3 + 3, lblW,   20,       FALSE);
        MoveWindow(g_hEditComment, editX,           inputY3,     editW,  20,       FALSE);
        MoveWindow(g_hBtnAdd,      width - pad - 70, inputY3 - 1, 60,    22,       FALSE);

        rcForm.left   = 0;
        rcForm.top    = btnRowY + btnHeight;
        rcForm.right  = width;
        rcForm.bottom = effectiveHeight;
        InvalidateRect(g_hMain, &rcForm, TRUE);
    }

    StatusBarApplyParts();
}

/* ====================================================================
 *  Action handlers (WM_COMMAND helpers)
 * ==================================================================== */

static void ShowAboutDialog(void) {
    if (MessageBox(g_hMain, ABOUT_MSG, "Version Info",
                   MB_YESNO | MB_ICONINFORMATION) == IDYES) {
        ShellExecute(g_hMain, "open", PROJECT_URL, NULL, NULL, SW_SHOW);
    }
}

static void ShowHelpDialog(void) {
    MessageBox(g_hMain, HELP_MSG, "Help", MB_OK | MB_ICONINFORMATION);
}

static void HandleEditSelected(void) {
    int  iPos = ListView_GetNextItem(g_hListView, -1, LVNI_SELECTED);
    char name[256], mac[256], comment[256];
    if (iPos == -1) return;

    ListView_GetItemText(g_hListView, iPos, COL_NAME,    name,    sizeof(name));
    ListView_GetItemText(g_hListView, iPos, COL_MAC,     mac,     sizeof(mac));
    ListView_GetItemText(g_hListView, iPos, COL_COMMENT, comment, sizeof(comment));
    SetWindowText(g_hEditName,    name);
    SetWindowText(g_hEditMac,     mac);
    SetWindowText(g_hEditComment, comment);
    lstrcpy(g_editingName, name);
    SetWindowText(g_hBtnAdd, "Save");
    SetStatus("Editing... Modify fields then Save.", COLOR_INFO, 0);
}

static void HandleRemoveSelected(void) {
    int    selCount = ListView_GetSelectedCount(g_hListView);
    char** namesToDelete;
    int    iPos = -1, idx = 0, i;

    if (selCount <= 0) return;

    namesToDelete = (char**)malloc(selCount * sizeof(char*));
    if (namesToDelete == NULL) return;

    while ((iPos = ListView_GetNextItem(g_hListView, iPos, LVNI_SELECTED)) != -1) {
        namesToDelete[idx] = (char*)malloc(256);
        ListView_GetItemText(g_hListView, iPos, COL_NAME, namesToDelete[idx], 256);
        idx++;
    }
    for (i = 0; i < selCount; i++) {
        RemoveContact(namesToDelete[i]);
        free(namesToDelete[i]);
    }
    free(namesToDelete);

    g_editingName[0] = '\0';
    SetWindowText(g_hBtnAdd, "Add");
    LoadContacts();
    SetStatus("Contact(s) removed.", COLOR_BAD, STATUS_CLEAR_MS_SINGLE);
}

static void HandleAddOrSaveContact(void) {
    char name[100], mac[100], comment[256];

    GetWindowText(g_hEditName,    name,    sizeof(name));
    GetWindowText(g_hEditMac,     mac,     sizeof(mac));
    GetWindowText(g_hEditComment, comment, sizeof(comment));

    if (strlen(name) == 0 || strlen(mac) == 0) {
        SetStatus("Error: Name and MAC required.", COLOR_BAD, STATUS_CLEAR_MS_SINGLE);
        return;
    }

    if (g_editingName[0] != '\0') {
        if (stricmp(g_editingName, name) != 0) RemoveContact(g_editingName);
        SaveContact(name, mac, comment);
        g_editingName[0] = '\0';
        SetWindowText(g_hBtnAdd, "Add");
    } else {
        SaveContact(name, mac, comment);
    }
    LoadContacts();
    SetWindowText(g_hEditName,    "");
    SetWindowText(g_hEditMac,     "");
    SetWindowText(g_hEditComment, "");
    SetStatus("Contact saved.", COLOR_OK, STATUS_CLEAR_MS_SINGLE);
}

static void ShowListContextMenu(void) {
    POINT  pt;
    HMENU  hPopup;

    if (ListView_GetNextItem(g_hListView, -1, LVNI_SELECTED) == -1) return;

    GetCursorPos(&pt);
    hPopup = CreatePopupMenu();
    AppendMenu(hPopup, MF_STRING,    ID_CTX_EDIT,   "Edit\tE");
    AppendMenu(hPopup, MF_STRING,    ID_BTN_CHECK,  "Check\tC");
    AppendMenu(hPopup, MF_STRING,    ID_BTN_WAKE,   "Wake Up\tEnter");
    AppendMenu(hPopup, MF_SEPARATOR, 0, NULL);
    AppendMenu(hPopup, MF_STRING,    ID_CTX_REMOVE, "Remove\tDel");
    TrackPopupMenu(hPopup, TPM_RIGHTBUTTON, pt.x, pt.y, 0, g_hMain, NULL);
    DestroyMenu(hPopup);
}

/* ====================================================================
 *  WM_CREATE - menu and control construction
 * ==================================================================== */

static void CreateMainMenu(void) {
    HMENU hMenu      = CreateMenu();
    HMENU hFileMenu  = CreatePopupMenu();
    HMENU hAboutMenu = CreatePopupMenu();

    AppendMenu(hFileMenu, MF_STRING,    ID_FILE_IMPORT, "Import contacts\tCtrl+O");
    AppendMenu(hFileMenu, MF_STRING,    ID_FILE_EXPORT, "Export contacts\tCtrl+S");
    AppendMenu(hFileMenu, MF_SEPARATOR, 0, NULL);
    AppendMenu(hFileMenu, MF_STRING,    ID_FILE_EXIT,   "Exit\tEsc");
    AppendMenu(hMenu,     MF_POPUP,     (UINT_PTR)hFileMenu,  "&File");

    g_hAutoMenu = CreatePopupMenu();
    AppendMenu(g_hAutoMenu, MF_STRING,    ID_AUTO_NONE, "No refresh");
    AppendMenu(g_hAutoMenu, MF_SEPARATOR, 0, NULL);
    AppendMenu(g_hAutoMenu, MF_STRING,    ID_AUTO_1SEC, "1 sec.");
    AppendMenu(g_hAutoMenu, MF_STRING,    ID_AUTO_2SEC, "2 sec.");
    AppendMenu(hMenu,       MF_POPUP,     (UINT_PTR)g_hAutoMenu, "&Auto-check");

    AppendMenu(hAboutMenu, MF_STRING, ID_HELP_ABOUT, "Version Info");
    AppendMenu(hAboutMenu, MF_STRING, ID_HELP_HELP,  "Help");
    AppendMenu(hMenu,      MF_POPUP,  (UINT_PTR)hAboutMenu, "&About");

    SetMenu(g_hMain, hMenu);
    SetAutoRefresh(0);
}

static void CreateMainControls(void) {
    LVCOLUMN lvc = {0};
    HFONT    hFont;

    g_hListView = CreateWindow(WC_LISTVIEW, NULL,
        WS_CHILD | WS_VISIBLE | WS_BORDER | LVS_REPORT | LVS_SHOWSELALWAYS | WS_TABSTOP,
        0, 0, 0, 0, g_hMain, (HMENU)ID_LISTVIEW, NULL, NULL);
    ListView_SetExtendedListViewStyle(g_hListView, LVS_EX_FULLROWSELECT | LVS_EX_GRIDLINES);

    lvc.mask = LVCF_FMT | LVCF_WIDTH | LVCF_TEXT | LVCF_SUBITEM;
    lvc.fmt  = LVCFMT_LEFT;
    lvc.iSubItem = COL_NAME;    lvc.cx = 150; lvc.pszText = "Name/IP";
    ListView_InsertColumn(g_hListView, COL_NAME,    &lvc);
    lvc.iSubItem = COL_MAC;     lvc.cx = 120; lvc.pszText = "MAC Address";
    ListView_InsertColumn(g_hListView, COL_MAC,     &lvc);
    lvc.iSubItem = COL_STATUS;  lvc.cx = 130; lvc.pszText = "Status";
    ListView_InsertColumn(g_hListView, COL_STATUS,  &lvc);
    lvc.iSubItem = COL_COMMENT; lvc.cx = 150; lvc.pszText = "Comment";
    ListView_InsertColumn(g_hListView, COL_COMMENT, &lvc);

    g_editingName[0] = '\0';

    /* Tab order: list, edit, check, wake, then form fields */
    g_hBtnEdit  = CreateWindow("BUTTON", "Edit",
        WS_CHILD | WS_VISIBLE | BS_PUSHBUTTON | WS_TABSTOP,
        0, 0, 0, 0, g_hMain, (HMENU)ID_BTN_EDIT, NULL, NULL);
    g_hBtnCheck = CreateWindow("BUTTON", "Check",
        WS_CHILD | WS_VISIBLE | BS_PUSHBUTTON | WS_TABSTOP,
        0, 0, 0, 0, g_hMain, (HMENU)ID_BTN_CHECK, NULL, NULL);
    g_hBtnWake  = CreateWindow("BUTTON", "Wake Up",
        WS_CHILD | WS_VISIBLE | BS_PUSHBUTTON | WS_TABSTOP,
        0, 0, 0, 0, g_hMain, (HMENU)ID_BTN_WAKE, NULL, NULL);

    g_hGroupBox = CreateWindow("BUTTON", "New Contact",
        WS_CHILD | WS_VISIBLE | BS_GROUPBOX,
        0, 0, 0, 0, g_hMain, (HMENU)ID_GROUPBOX, NULL, NULL);

    g_hLblName     = CreateWindow("STATIC", "Name/IP:",
        WS_CHILD | WS_VISIBLE, 0, 0, 0, 0, g_hMain, NULL, NULL, NULL);
    g_hEditName    = CreateWindow("EDIT", "",
        WS_CHILD | WS_VISIBLE | WS_BORDER | ES_AUTOHSCROLL | WS_TABSTOP,
        0, 0, 0, 0, g_hMain, (HMENU)ID_EDIT_NAME, NULL, NULL);
    g_hLblMac      = CreateWindow("STATIC", "MAC:",
        WS_CHILD | WS_VISIBLE, 0, 0, 0, 0, g_hMain, NULL, NULL, NULL);
    g_hEditMac     = CreateWindow("EDIT", "",
        WS_CHILD | WS_VISIBLE | WS_BORDER | ES_AUTOHSCROLL | WS_TABSTOP,
        0, 0, 0, 0, g_hMain, (HMENU)ID_EDIT_MAC, NULL, NULL);
    g_hLblComment  = CreateWindow("STATIC", "Comment:",
        WS_CHILD | WS_VISIBLE, 0, 0, 0, 0, g_hMain, NULL, NULL, NULL);
    g_hEditComment = CreateWindow("EDIT", "",
        WS_CHILD | WS_VISIBLE | WS_BORDER | ES_AUTOHSCROLL | WS_TABSTOP,
        0, 0, 0, 0, g_hMain, (HMENU)ID_EDIT_COMMENT, NULL, NULL);

    g_hBtnAdd      = CreateWindow("BUTTON", "Add",
        WS_CHILD | WS_VISIBLE | BS_PUSHBUTTON | WS_TABSTOP,
        0, 0, 0, 0, g_hMain, (HMENU)ID_BTN_ADD, NULL, NULL);
    g_hStatusBar   = CreateWindow(STATUSCLASSNAME, "",
        WS_CHILD | WS_VISIBLE | SBARS_SIZEGRIP,
        0, 0, 0, 0, g_hMain, (HMENU)ID_STATUSBAR, NULL, NULL);

    hFont = (HFONT)GetStockObject(DEFAULT_GUI_FONT);
    SendMessage(g_hListView,    WM_SETFONT, (WPARAM)hFont, TRUE);
    SendMessage(g_hBtnEdit,     WM_SETFONT, (WPARAM)hFont, TRUE);
    SendMessage(g_hBtnCheck,    WM_SETFONT, (WPARAM)hFont, TRUE);
    SendMessage(g_hBtnWake,     WM_SETFONT, (WPARAM)hFont, TRUE);
    SendMessage(g_hGroupBox,    WM_SETFONT, (WPARAM)hFont, TRUE);
    SendMessage(g_hLblName,     WM_SETFONT, (WPARAM)hFont, TRUE);
    SendMessage(g_hEditName,    WM_SETFONT, (WPARAM)hFont, TRUE);
    SendMessage(g_hLblMac,      WM_SETFONT, (WPARAM)hFont, TRUE);
    SendMessage(g_hEditMac,     WM_SETFONT, (WPARAM)hFont, TRUE);
    SendMessage(g_hLblComment,  WM_SETFONT, (WPARAM)hFont, TRUE);
    SendMessage(g_hEditComment, WM_SETFONT, (WPARAM)hFont, TRUE);
    SendMessage(g_hBtnAdd,      WM_SETFONT, (WPARAM)hFont, TRUE);
    SendMessage(g_hStatusBar,   WM_SETFONT, (WPARAM)hFont, TRUE);
}

/* ====================================================================
 *  Notification handlers
 * ==================================================================== */

static LRESULT HandleListViewNotify(LPNMHDR lpnmh, LPARAM lParam) {
    if (lpnmh->code == NM_CUSTOMDRAW) {
        LPNMLVCUSTOMDRAW plvcd = (LPNMLVCUSTOMDRAW)lParam;
        switch (plvcd->nmcd.dwDrawStage) {
            case CDDS_PREPAINT:
                return CDRF_NOTIFYITEMDRAW;
            case CDDS_ITEMPREPAINT: {
                DWORD lp = (DWORD)plvcd->nmcd.lItemlParam;
                DWORD hi = lp >> ROW_HI_SHIFT;
                if      (hi == ROW_COLOR_GREEN) plvcd->clrText = COLOR_OK;
                else if (hi == ROW_COLOR_RED)   plvcd->clrText = COLOR_BAD;
                return CDRF_NEWFONT;
            }
            default:
                break;
        }
        return CDRF_DODEFAULT;
    }

    if (lpnmh->code == LVN_COLUMNCLICK) {
        NMLISTVIEW* nml = (NMLISTVIEW*)lParam;
        if (nml->iSubItem == g_sortColumn) {
            g_sortAscending = !g_sortAscending;
        } else {
            g_sortColumn = nml->iSubItem;
            if (g_sortColumn < 0)             g_sortColumn = 0;
            if (g_sortColumn >= COL_COUNT)    g_sortColumn = COL_COUNT - 1;
            g_sortAscending = 1;
        }
        SortListView();
        return 0;
    }

    if (lpnmh->code == NM_DBLCLK) {
        LPNMITEMACTIVATE nmi = (LPNMITEMACTIVATE)lParam;
        if (nmi->iItem != -1) {
            char name[256], mac[256];
            ListView_GetItemText(g_hListView, nmi->iItem, COL_NAME, name, sizeof(name));
            ListView_GetItemText(g_hListView, nmi->iItem, COL_MAC,  mac,  sizeof(mac));
            WakeOneComputer(nmi->iItem, name, mac);
        }
    }
    return 0;
}

static BOOL HandleStatusBarDrawItem(LPDRAWITEMSTRUCT pdis) {
    int          part = (int)pdis->itemID;
    COLORREF     col  = COLOR_NEUTRAL;
    const char*  txt  = "";
    RECT         rc;

    if (part >= 0 && part < g_statusCellCount) {
        col = g_statusCells[part].color;
        txt = g_statusCells[part].text;
    } else if (g_statusCellCount > 0) {
        col = g_statusCells[0].color;
        txt = g_statusCells[0].text;
    }
    SetTextColor(pdis->hDC, col);
    SetBkMode(pdis->hDC, TRANSPARENT);
    rc = pdis->rcItem;
    rc.left += 5;
    DrawText(pdis->hDC, txt, -1, &rc, DT_SINGLELINE | DT_VCENTER | DT_END_ELLIPSIS);
    return TRUE;
}

/* ====================================================================
 *  WindowProc - main message dispatcher
 * ==================================================================== */

static LRESULT CALLBACK WindowProc(HWND hwnd, UINT uMsg, WPARAM wParam, LPARAM lParam) {
    switch (uMsg) {
        case WM_CREATE:
            g_hMain = hwnd;
            SetupIniPath();
            CreateMainMenu();
            CreateMainControls();
            LoadContacts();
            LoadColumnWidths();
            SetFocus(g_hListView);
            return 0;

        case WM_DRAWITEM: {
            LPDRAWITEMSTRUCT pdis = (LPDRAWITEMSTRUCT)lParam;
            if (pdis->CtlID == ID_STATUSBAR) return HandleStatusBarDrawItem(pdis);
            break;
        }

        case WM_SIZE:
            LayoutMainWindow(LOWORD(lParam), HIWORD(lParam));
            return 0;

        case WM_GETMINMAXINFO: {
            MINMAXINFO* mmi = (MINMAXINFO*)lParam;
            mmi->ptMinTrackSize.x = WIN_MIN_W;
            mmi->ptMinTrackSize.y = WIN_MIN_H;
            return 0;
        }

        case WM_TIMER:
            if (wParam == IDT_STATUS_TIMER)          SetStatus("", COLOR_NEUTRAL, 0);
            else if (wParam == IDT_AUTO_CHECK_TIMER) StartCheck(FALSE, TRUE);
            return 0;

        case WM_CHECK_COMPLETE: {
            CheckResult* result = (CheckResult*)lParam;
            if (result != NULL) {
                ApplyCheckResult(result);
                free(result);
            } else {
                g_checkInProgress = 0;
            }
            return 0;
        }

        case WM_NOTIFY: {
            LPNMHDR lpnmh = (LPNMHDR)lParam;
            if (lpnmh->idFrom == ID_LISTVIEW) return HandleListViewNotify(lpnmh, lParam);
            return 0;
        }

        case WM_CONTEXTMENU:
            if ((HWND)wParam == g_hListView) ShowListContextMenu();
            return 0;

        case WM_COMMAND: {
            WORD cmd = LOWORD(wParam);
            switch (cmd) {
                case ID_HELP_ABOUT: ShowAboutDialog(); return 0;
                case ID_HELP_HELP:  ShowHelpDialog();  return 0;

                case ID_FILE_IMPORT:
                    g_editingName[0] = '\0';
                    SetWindowText(g_hBtnAdd, "Add");
                    HandleImport();
                    return 0;
                case ID_FILE_EXPORT: HandleExport();      return 0;
                case ID_FILE_EXIT:   PostQuitMessage(0);  return 0;

                case ID_AUTO_NONE: SetAutoRefresh(0);                                return 0;
                case ID_AUTO_1SEC: SetAutoRefresh(1000); StartCheck(FALSE, TRUE);    return 0;
                case ID_AUTO_2SEC: SetAutoRefresh(2000); StartCheck(FALSE, TRUE);    return 0;

                case ID_BTN_CHECK:                StartCheck(TRUE, TRUE); return 0;
                case ID_BTN_WAKE:                 WakeListSelection();    return 0;
                case ID_BTN_EDIT:
                case ID_CTX_EDIT:                 HandleEditSelected();   return 0;
                case ID_CTX_REMOVE:               HandleRemoveSelected(); return 0;
                case ID_BTN_ADD:                  HandleAddOrSaveContact(); return 0;
            }
            return 0;
        }

        case WM_DESTROY:
            KillTimer(hwnd, IDT_AUTO_CHECK_TIMER);
            SaveWindowSize();
            SaveColumnWidths();
            if (g_sortSnap) {
                free(g_sortSnap);
                g_sortSnap    = NULL;
                g_sortSnapCap = 0;
            }
            PostQuitMessage(0);
            return 0;
    }
    return DefWindowProc(hwnd, uMsg, wParam, lParam);
}

/* ====================================================================
 *  WinMain
 * ==================================================================== */

int WINAPI WinMain(HINSTANCE hInstance, HINSTANCE hPrevInstance, LPSTR lpCmdLine, int nCmdShow) {
    WNDCLASS wc = {0};
    HWND     hwnd;
    MSG      msg;
    HACCEL   hAccel;
    ACCEL    accelerators[5];
    int      winW, winH;

    (void)hPrevInstance; (void)lpCmdLine;

    InitCommonControlsCompat();
    SetupIniPath();
    LoadWindowSize(&winW, &winH);

    wc.lpfnWndProc   = WindowProc;
    wc.hInstance     = hInstance;
    wc.lpszClassName = CLASS_NAME;
    wc.hCursor       = LoadCursor(NULL, IDC_ARROW);
    wc.hbrBackground = (HBRUSH)(COLOR_BTNFACE + 1);
    wc.hIcon         = LoadIcon(hInstance, MAKEINTRESOURCE(IDI_APPICON));
    RegisterClass(&wc);

    hwnd = CreateWindowEx(WS_EX_CONTROLPARENT, CLASS_NAME, APP_TITLE,
        WS_OVERLAPPEDWINDOW,
        CW_USEDEFAULT, CW_USEDEFAULT, winW, winH,
        NULL, NULL, hInstance, NULL);
    if (hwnd == NULL) return 0;
    ShowWindow(hwnd, nCmdShow);

    accelerators[0].fVirt = FCONTROL | FVIRTKEY; accelerators[0].key = 'O';        accelerators[0].cmd = ID_FILE_IMPORT;
    accelerators[1].fVirt = FCONTROL | FVIRTKEY; accelerators[1].key = 'S';        accelerators[1].cmd = ID_FILE_EXPORT;
    accelerators[2].fVirt = FVIRTKEY;            accelerators[2].key = VK_ESCAPE;  accelerators[2].cmd = ID_FILE_EXIT;
    accelerators[3].fVirt = FVIRTKEY;            accelerators[3].key = VK_DELETE;  accelerators[3].cmd = ID_CTX_REMOVE;
    accelerators[4].fVirt = FVIRTKEY;            accelerators[4].key = VK_RETURN;  accelerators[4].cmd = ID_BTN_WAKE;
    hAccel = CreateAcceleratorTable(accelerators, 5);

    while (GetMessage(&msg, NULL, 0, 0)) {
        if (!TranslateAccelerator(hwnd, hAccel, &msg)) {
            if (!IsDialogMessage(hwnd, &msg)) {
                TranslateMessage(&msg);
                DispatchMessage(&msg);
            }
        }
    }

    DestroyAcceleratorTable(hAccel);
    return 0;
}
