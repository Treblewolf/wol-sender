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

// --- Constants & IDs ---
#define ID_LISTVIEW 101
#define ID_BTN_WAKE 102
#define ID_BTN_CHECK 109
#define ID_BTN_EDIT 110
#define ID_EDIT_NAME 103
#define ID_EDIT_MAC 104
#define ID_BTN_ADD 105
#define ID_STATUSBAR 106
#define ID_GROUPBOX 107
#define ID_EDIT_COMMENT 108

#define ID_FILE_IMPORT 201
#define ID_FILE_EXPORT 202
#define ID_FILE_EXIT 203
#define ID_HELP_ABOUT 204

#define ID_CTX_EDIT 301
#define ID_CTX_REMOVE 302

#define IDT_STATUS_TIMER 401
#define IDI_APPICON 100 

#define ROW_HI_SHIFT 16
#define ROW_HI_MASK  0xFFFFu
#define ROW_COLOR_GREEN 1u
#define ROW_COLOR_RED   2u

#define MAX_STATUS_PARTS 24

typedef struct {
    char text[384];
    COLORREF color;
} StatusPartCell;

typedef struct {
    char col0[256];
    char col1[256];
    char col2[256];
} SortSnapRow;

#define INI_SECTION_CONTACTS "Contacts"
#define INI_SECTION_COMMENTS "Comments"
#define STATUS_CLEAR_MS_SINGLE 2000
#define STATUS_CLEAR_MS_MULTI 10000
#define STATUS_AGG_BUF 2048

#define PROJECT_URL "https://github.com/Treblewolf/wol-sender"

// --- Global Variables ---
char iniFilePath[MAX_PATH];
COLORREF statusColor = RGB(0, 0, 0); 
char statusText[STATUS_AGG_BUF] = ""; 
StatusPartCell g_statusCells[MAX_STATUS_PARTS];
int g_statusCellCount = 0;
BOOL g_statusMultiCell = FALSE;
int g_statusWidths[MAX_STATUS_PARTS];

static int g_sortColumn = 0;
static int g_sortAscending = 1;
static SortSnapRow* g_sortSnap = NULL;
static int g_sortSnapCap = 0;

void SetStatus(HWND hwnd, HWND hStatus, const char* text, COLORREF color, UINT resetMs);
static void ClearRowHighlightBits(HWND hListView);
static void StatusBarApplyParts(HWND hwndMain, HWND hStatus);
static void SetStatusParts(HWND hwnd, HWND hStatus, const StatusPartCell* cells, int n, UINT resetMs);
static void SortListView(HWND hListView);

// --- Networking Logic ---
BOOL SendWakeOnLan(unsigned char* macAddress) {
    WSADATA wsaData; SOCKET udpSocket; struct sockaddr_in targetAddress;
    char magicPacket[102]; int i, bOptVal = 1; BOOL success = FALSE;

    if (WSAStartup(MAKEWORD(1, 1), &wsaData) != 0) return FALSE;
    
    udpSocket = socket(AF_INET, SOCK_DGRAM, 0);
    if (udpSocket == INVALID_SOCKET) { WSACleanup(); return FALSE; }

    setsockopt(udpSocket, SOL_SOCKET, SO_BROADCAST, (char*)&bOptVal, sizeof(bOptVal));
    for (i = 0; i < 6; i++) magicPacket[i] = 0xFF;
    for (i = 1; i <= 16; i++) memcpy(&magicPacket[i * 6], macAddress, 6);

    targetAddress.sin_family = AF_INET;
    targetAddress.sin_port = htons(9);
    targetAddress.sin_addr.s_addr = inet_addr("255.255.255.255");

    if (sendto(udpSocket, magicPacket, sizeof(magicPacket), 0, (struct sockaddr*)&targetAddress, sizeof(targetAddress)) != SOCKET_ERROR) {
        success = TRUE;
    }
    closesocket(udpSocket); WSACleanup();
    return success;
}

/* --- Host reachability (Computer Name = DNS / NetBIOS name or IPv4) --- */
typedef HANDLE (WINAPI *PFN_IcmpCreateFile)(void);
typedef BOOL (WINAPI *PFN_IcmpCloseHandle)(HANDLE);
typedef DWORD (WINAPI *PFN_IcmpSendEcho)(HANDLE, ULONG, LPVOID, WORD, LPVOID, LPVOID, DWORD, DWORD);

typedef struct {
    ULONG Address;
    ULONG Status;
} ICMP_ECHO_REPLY_HDR;

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

static BOOL TryLoadIcmp(PFN_IcmpCreateFile* pC, PFN_IcmpCloseHandle* pCl, PFN_IcmpSendEcho* pS, HMODULE* pMod) {
    HMODULE h = LoadLibraryA("icmp.dll");
    if (h == NULL) h = LoadLibraryA("IPHLPAPI.DLL");
    if (h == NULL) return FALSE;
    *pC = (PFN_IcmpCreateFile)GetProcAddress(h, "IcmpCreateFile");
    *pCl = (PFN_IcmpCloseHandle)GetProcAddress(h, "IcmpCloseHandle");
    *pS = (PFN_IcmpSendEcho)GetProcAddress(h, "IcmpSendEcho");
    if (*pC == NULL || *pCl == NULL || *pS == NULL) {
        FreeLibrary(h);
        return FALSE;
    }
    *pMod = h;
    return TRUE;
}

static BOOL TryIcmpPing(PFN_IcmpCreateFile pfnC, PFN_IcmpCloseHandle pfnCl, PFN_IcmpSendEcho pfnS, struct in_addr addr) {
    HANDLE h;
    char sendData[4];
    char replyBuf[256];
    DWORD n;
    ICMP_ECHO_REPLY_HDR* r;

    h = pfnC();
    if (h == INVALID_HANDLE_VALUE) return FALSE;
    n = pfnS(h, addr.S_un.S_addr, sendData, sizeof(sendData), NULL, replyBuf, sizeof(replyBuf), 2000);
    pfnCl(h);
    if (n == 0) return FALSE;
    r = (ICMP_ECHO_REPLY_HDR*)replyBuf;
    return (r->Status == 0);
}

static BOOL TcpPortReachable(struct in_addr addr, unsigned short port) {
    SOCKET s;
    struct sockaddr_in sin;
    u_long nb = 1;
    fd_set wf, ef;
    struct timeval tv;
    int err, len;
    int r;

    s = socket(AF_INET, SOCK_STREAM, IPPROTO_TCP);
    if (s == INVALID_SOCKET) return FALSE;
    ZeroMemory(&sin, sizeof(sin));
    sin.sin_family = AF_INET;
    sin.sin_addr = addr;
    sin.sin_port = htons(port);

    if (ioctlsocket(s, FIONBIO, &nb) != 0) { closesocket(s); return FALSE; }

    r = connect(s, (struct sockaddr*)&sin, sizeof(sin));
    if (r == 0) { closesocket(s); return TRUE; }

    err = WSAGetLastError();
    if (err != WSAEWOULDBLOCK && err != WSAEINPROGRESS) { closesocket(s); return FALSE; }

    FD_ZERO(&wf);
    FD_ZERO(&ef);
    FD_SET(s, &wf);
    FD_SET(s, &ef);
    tv.tv_sec = 2;
    tv.tv_usec = 0;
    r = select(0, NULL, &wf, &ef, &tv);
    if (r <= 0) { closesocket(s); return FALSE; }

    err = 0;
    len = sizeof(err);
    if (getsockopt(s, SOL_SOCKET, SO_ERROR, (char*)&err, &len) != 0) { closesocket(s); return FALSE; }
    closesocket(s);
    if (err == 0 || err == WSAECONNREFUSED) return TRUE;
    return FALSE;
}

static BOOL CheckEvaluateOne(const char* computerName,
    BOOL haveIcmp, PFN_IcmpCreateFile pfnC, PFN_IcmpCloseHandle pfnCl, PFN_IcmpSendEcho pfnS,
    char* detail, int detailLen) {
    struct in_addr addr;

    if (computerName == NULL || computerName[0] == '\0') {
        lstrcpyn(detail, "No computer name", detailLen);
        return FALSE;
    }
    if (!ResolveHostToAddr(computerName, &addr)) {
        lstrcpyn(detail, "Unknown host (use DNS/NetBIOS name or IP in Computer Name)", detailLen);
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
    lstrcpyn(detail, "No reply (off, wrong subnet, or firewall)", detailLen);
    return FALSE;
}

static BOOL WakeEvaluateOne(const char* macStr, char* detail, int detailLen) {
    int values[6], parsed = 0;
    int i;
    if (sscanf(macStr, "%x:%x:%x:%x:%x:%x", &values[0], &values[1], &values[2], &values[3], &values[4], &values[5]) == 6) parsed = 1;
    else if (sscanf(macStr, "%x-%x-%x-%x-%x-%x", &values[0], &values[1], &values[2], &values[3], &values[4], &values[5]) == 6) parsed = 1;

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

static void SetRowResultColor(HWND hListView, int itemIndex, BOOL ok) {
    LVITEM it = {0};
    DWORD kind = ok ? ROW_COLOR_GREEN : ROW_COLOR_RED;
    DWORD lp;
    if (itemIndex < 0) return;
    it.mask = LVIF_PARAM;
    it.iItem = itemIndex;
    if (!ListView_GetItem(hListView, &it)) return;
    lp = (DWORD)it.lParam;
    lp = (kind << ROW_HI_SHIFT) | (lp & ROW_HI_MASK);
    it.lParam = (LONG)lp;
    ListView_SetItem(hListView, &it);
}

static void CheckListSelection(HWND hwnd, HWND hListView, HWND hStatus) {
    char detail[256];
    char line[384];
    StatusPartCell cells[MAX_STATUS_PARTS];
    int nSel = 0;
    int iPos = -1;
    int ci = 0;
    HMODULE hIcmp = NULL;
    PFN_IcmpCreateFile pfnC = NULL;
    PFN_IcmpCloseHandle pfnCl = NULL;
    PFN_IcmpSendEcho pfnS = NULL;
    BOOL haveIcmp;
    WSADATA wsaData;

    while ((iPos = ListView_GetNextItem(hListView, iPos, LVNI_SELECTED)) != -1) nSel++;
    if (nSel == 0) return;

    haveIcmp = TryLoadIcmp(&pfnC, &pfnCl, &pfnS, &hIcmp);
    if (WSAStartup(MAKEWORD(1, 1), &wsaData) != 0) {
        SetStatus(hwnd, hStatus, "Winsock init failed.", RGB(200, 0, 0), STATUS_CLEAR_MS_SINGLE);
        if (hIcmp) FreeLibrary(hIcmp);
        return;
    }

    iPos = -1;
    while ((iPos = ListView_GetNextItem(hListView, iPos, LVNI_SELECTED)) != -1) {
        char name[256];
        BOOL ok;
        ListView_GetItemText(hListView, iPos, 0, name, sizeof(name));
        ok = CheckEvaluateOne(name, haveIcmp, pfnC, pfnCl, pfnS, detail, sizeof(detail));
        SetRowResultColor(hListView, iPos, ok);
        lstrcpyn(line, name, sizeof(line));
        lstrcat(line, ": ");
        if (lstrlen(line) + lstrlen(detail) < (int)sizeof(line)) lstrcat(line, detail);

        if (nSel == 1) {
            WSACleanup();
            if (hIcmp) FreeLibrary(hIcmp);
            SetStatus(hwnd, hStatus, line, ok ? RGB(0, 150, 0) : RGB(200, 0, 0), STATUS_CLEAR_MS_SINGLE);
            return;
        }
        if (ci < MAX_STATUS_PARTS) {
            lstrcpyn(cells[ci].text, line, sizeof(cells[ci].text));
            cells[ci].color = ok ? RGB(0, 150, 0) : RGB(200, 0, 0);
            ci++;
        }
    }

    WSACleanup();
    if (hIcmp) FreeLibrary(hIcmp);

    if (nSel > MAX_STATUS_PARTS) {
        SetStatus(hwnd, hStatus, "Too many selections for status bar.", RGB(200, 0, 0), STATUS_CLEAR_MS_MULTI);
        return;
    }
    SetStatusParts(hwnd, hStatus, cells, ci, STATUS_CLEAR_MS_MULTI);
}

// --- UI Helpers ---
static void StatusBarApplyParts(HWND hwndMain, HWND hStatus) {
    RECT rc;
    int clientW;
    int i, k, seg;

    GetClientRect(hwndMain, &rc);
    clientW = rc.right;
    if (clientW < 20) clientW = 400;

    k = g_statusCellCount;
    if (k <= 0) k = 1;

    if (!g_statusMultiCell || k == 1) {
        g_statusWidths[0] = -1;
        SendMessage(hStatus, SB_SETPARTS, 1, (LPARAM)g_statusWidths);
    } else {
        seg = clientW / k;
        if (seg < 40) seg = 40;
        for (i = 0; i < k - 1; i++)
            g_statusWidths[i] = seg * (i + 1);
        g_statusWidths[k - 1] = -1;
        SendMessage(hStatus, SB_SETPARTS, k, (LPARAM)g_statusWidths);
    }

    for (i = 0; i < k; i++)
        SendMessage(hStatus, SB_SETTEXT, SBT_OWNERDRAW | i, (LPARAM)g_statusCells[i].text);
}

void SetStatus(HWND hwnd, HWND hStatus, const char* text, COLORREF color, UINT resetMs) {
    g_statusMultiCell = FALSE;
    g_statusCellCount = 1;
    lstrcpyn(g_statusCells[0].text, text, sizeof(g_statusCells[0].text));
    g_statusCells[0].color = color;
    lstrcpyn(statusText, text, sizeof(statusText));
    statusColor = color;
    StatusBarApplyParts(hwnd, hStatus);
    if (resetMs > 0) SetTimer(hwnd, IDT_STATUS_TIMER, resetMs, NULL);
    else KillTimer(hwnd, IDT_STATUS_TIMER);
}

static void SetStatusParts(HWND hwnd, HWND hStatus, const StatusPartCell* cells, int n, UINT resetMs) {
    int i;
    if (n <= 0) return;
    g_statusMultiCell = (n > 1);
    g_statusCellCount = n;
    for (i = 0; i < n && i < MAX_STATUS_PARTS; i++) {
        g_statusCells[i] = cells[i];
    }
    StatusBarApplyParts(hwnd, hStatus);
    if (resetMs > 0) SetTimer(hwnd, IDT_STATUS_TIMER, resetMs, NULL);
    else KillTimer(hwnd, IDT_STATUS_TIMER);
}

static void ClearRowHighlightBits(HWND hListView) {
    int i, n = ListView_GetItemCount(hListView);
    LVITEM it = {0};
    it.mask = LVIF_PARAM;
    for (i = 0; i < n; i++) {
        it.iItem = i;
        if (!ListView_GetItem(hListView, &it)) continue;
        it.lParam = (LONG)((DWORD)it.lParam & ROW_HI_MASK);
        ListView_SetItem(hListView, &it);
    }
}

static int CALLBACK LVCompareProc(LPARAM lp1, LPARAM lp2, LPARAM lpSort) {
    unsigned idx1 = (unsigned)(lp1) & ROW_HI_MASK;
    unsigned idx2 = (unsigned)(lp2) & ROW_HI_MASK;
    const char *s1, *s2;
    int cmp;
    (void)lpSort;

    if (g_sortSnap == NULL) return 0;
    if (idx1 >= (unsigned)g_sortSnapCap || idx2 >= (unsigned)g_sortSnapCap) return 0;

    switch (g_sortColumn) {
        case 1: s1 = g_sortSnap[idx1].col1; s2 = g_sortSnap[idx2].col1; break;
        case 2: s1 = g_sortSnap[idx1].col2; s2 = g_sortSnap[idx2].col2; break;
        default: s1 = g_sortSnap[idx1].col0; s2 = g_sortSnap[idx2].col0; break;
    }
    cmp = stricmp(s1, s2);
    if (cmp == 0) cmp = (int)idx1 - (int)idx2;
    return g_sortAscending ? cmp : -cmp;
}

static void SortListView(HWND hListView) {
    int count = ListView_GetItemCount(hListView);
    int i;
    LVITEM it = {0};

    if (count <= 1) return;

    if (g_sortSnapCap < count) {
        SortSnapRow* p = (SortSnapRow*)realloc(g_sortSnap, count * sizeof(SortSnapRow));
        if (p == NULL) return;
        g_sortSnap = p;
        g_sortSnapCap = count;
    }

    for (i = 0; i < count; i++) {
        LPARAM lp;
        ListView_GetItemText(hListView, i, 0, g_sortSnap[i].col0, sizeof(g_sortSnap[i].col0));
        ListView_GetItemText(hListView, i, 1, g_sortSnap[i].col1, sizeof(g_sortSnap[i].col1));
        ListView_GetItemText(hListView, i, 2, g_sortSnap[i].col2, sizeof(g_sortSnap[i].col2));
        it.mask = LVIF_PARAM;
        it.iItem = i;
        it.iSubItem = 0;
        if (!ListView_GetItem(hListView, &it)) continue;
        lp = (DWORD)it.lParam;
        lp = (lp & 0xFFFF0000u) | ((DWORD)i & ROW_HI_MASK);
        it.lParam = (LONG)lp;
        ListView_SetItem(hListView, &it);
    }

    ListView_SortItems(hListView, LVCompareProc, 0);
}

static void WakeOneComputer(HWND hwnd, HWND hStatus, HWND hListView, int itemIndex, const char* computerName, const char* macStr) {
    char detail[256];
    char line[384];
    BOOL ok = WakeEvaluateOne(macStr, detail, sizeof(detail));
    lstrcpyn(line, computerName, sizeof(line));
    lstrcat(line, ": ");
    if (lstrlen(line) + lstrlen(detail) < (int)sizeof(line)) lstrcat(line, detail);
    SetRowResultColor(hListView, itemIndex, ok);
    SetStatus(hwnd, hStatus, line, ok ? RGB(0, 150, 0) : RGB(200, 0, 0), STATUS_CLEAR_MS_SINGLE);
}

static void WakeListSelection(HWND hwnd, HWND hListView, HWND hStatus) {
    char detail[256];
    char line[384];
    StatusPartCell cells[MAX_STATUS_PARTS];
    int nSel = 0;
    int iPos = -1;
    int ci = 0;

    while ((iPos = ListView_GetNextItem(hListView, iPos, LVNI_SELECTED)) != -1) nSel++;
    if (nSel == 0) return;

    iPos = -1;
    while ((iPos = ListView_GetNextItem(hListView, iPos, LVNI_SELECTED)) != -1) {
        char name[256], mac[256];
        BOOL ok;
        ListView_GetItemText(hListView, iPos, 0, name, sizeof(name));
        ListView_GetItemText(hListView, iPos, 1, mac, sizeof(mac));
        ok = WakeEvaluateOne(mac, detail, sizeof(detail));
        SetRowResultColor(hListView, iPos, ok);
        lstrcpyn(line, name, sizeof(line));
        lstrcat(line, ": ");
        if (lstrlen(line) + lstrlen(detail) < (int)sizeof(line)) lstrcat(line, detail);

        if (nSel == 1) {
            SetStatus(hwnd, hStatus, line, ok ? RGB(0, 150, 0) : RGB(200, 0, 0), STATUS_CLEAR_MS_SINGLE);
            return;
        }
        if (ci < MAX_STATUS_PARTS) {
            lstrcpyn(cells[ci].text, line, sizeof(cells[ci].text));
            cells[ci].color = ok ? RGB(0, 150, 0) : RGB(200, 0, 0);
            ci++;
        }
    }

    if (nSel > MAX_STATUS_PARTS) {
        SetStatus(hwnd, hStatus, "Too many selections for status bar.", RGB(200, 0, 0), STATUS_CLEAR_MS_MULTI);
        return;
    }
    SetStatusParts(hwnd, hStatus, cells, ci, STATUS_CLEAR_MS_MULTI);
}

// --- Data Persistence ---
void SetupIniPath() {
    GetModuleFileName(NULL, iniFilePath, MAX_PATH);
    char *lastSlash = strrchr(iniFilePath, '\\');
    if (lastSlash != NULL) strcpy(lastSlash + 1, "wol_contacts.ini"); 
}

void LoadContacts(HWND hListView) {
    char buffer[32767];
    ListView_DeleteAllItems(hListView); 
    DWORD charsRead = GetPrivateProfileSection(INI_SECTION_CONTACTS, buffer, sizeof(buffer), iniFilePath);
    if (charsRead > 0) {
        char* currentStr = buffer;
        int index = 0;
        while (*currentStr != '\0') {
            int originalLen = strlen(currentStr);
            char* equalsSign = strchr(currentStr, '=');
            if (equalsSign != NULL) {
                *equalsSign = '\0'; 
                char* name = currentStr;
                char* value = equalsSign + 1;
                char mac[256];
                char comment[256];
                char* tab = strchr(value, '\t');
                if (tab != NULL) {
                    *tab = '\0';
                    lstrcpyn(mac, value, sizeof(mac));
                    lstrcpyn(comment, tab + 1, sizeof(comment));
                } else {
                    lstrcpyn(mac, value, sizeof(mac));
                    GetPrivateProfileString(INI_SECTION_COMMENTS, name, "", comment, sizeof(comment), iniFilePath);
                }
                LVITEM lvi = {0};
                lvi.mask = LVIF_TEXT | LVIF_PARAM;
                lvi.iItem = index;
                lvi.iSubItem = 0;
                lvi.pszText = name;
                lvi.lParam = (LPARAM)(DWORD)index;
                ListView_InsertItem(hListView, &lvi);
                ListView_SetItemText(hListView, index, 1, mac);
                ListView_SetItemText(hListView, index, 2, comment);
                index++;
            }
            currentStr += originalLen + 1;
        }
    }
}

void SaveContact(char* name, char* mac, char* comment) {
    WritePrivateProfileString(INI_SECTION_CONTACTS, name, mac, iniFilePath);
    if (comment != NULL && comment[0] != '\0') {
        WritePrivateProfileString(INI_SECTION_COMMENTS, name, comment, iniFilePath);
    } else {
        WritePrivateProfileString(INI_SECTION_COMMENTS, name, NULL, iniFilePath);
    }
}
void RemoveContact(char* name) {
    WritePrivateProfileString(INI_SECTION_CONTACTS, name, NULL, iniFilePath);
    WritePrivateProfileString(INI_SECTION_COMMENTS, name, NULL, iniFilePath);
}

static DWORD GetOpenFileNameStructSize(void) {
    OSVERSIONINFO osvi;
    ZeroMemory(&osvi, sizeof(osvi));
    osvi.dwOSVersionInfoSize = sizeof(osvi);
    if (GetVersionEx(&osvi) && osvi.dwPlatformId == VER_PLATFORM_WIN32_WINDOWS) {
        return OPENFILENAME_SIZE_VERSION_400;
    }
    return sizeof(OPENFILENAME);
}

static BOOL InitCommonControlsCompat(void) {
    INITCOMMONCONTROLSEX icex;
    ZeroMemory(&icex, sizeof(icex));
    icex.dwSize = sizeof(INITCOMMONCONTROLSEX);
    icex.dwICC = ICC_LISTVIEW_CLASSES | ICC_BAR_CLASSES;
    if (InitCommonControlsEx(&icex)) return TRUE;
    InitCommonControls();
    return TRUE;
}

void HandleImport(HWND hwnd, HWND hListView, HWND hStatus) {
    OPENFILENAME ofn; char szFile[260] = {0};
    ZeroMemory(&ofn, sizeof(ofn));
    ofn.lStructSize = GetOpenFileNameStructSize(); ofn.hwndOwner = hwnd; ofn.lpstrFile = szFile; ofn.nMaxFile = sizeof(szFile);
    ofn.lpstrFilter = "INI Files\0*.ini\0All Files\0*.*\0"; ofn.nFilterIndex = 1; ofn.Flags = OFN_PATHMUSTEXIST | OFN_FILEMUSTEXIST;
    if (GetOpenFileName(&ofn) == TRUE) {
        CopyFile(szFile, iniFilePath, FALSE); 
        LoadContacts(hListView);
        SetStatus(hwnd, hStatus, "Contacts imported.", RGB(0, 150, 0), STATUS_CLEAR_MS_SINGLE);
    }
}

void HandleExport(HWND hwnd, HWND hStatus) {
    OPENFILENAME ofn; char szFile[260] = "wol_contacts.ini"; 
    ZeroMemory(&ofn, sizeof(ofn));
    ofn.lStructSize = GetOpenFileNameStructSize(); ofn.hwndOwner = hwnd; ofn.lpstrFile = szFile; ofn.nMaxFile = sizeof(szFile);
    ofn.lpstrFilter = "INI Files\0*.ini\0All Files\0*.*\0"; ofn.nFilterIndex = 1; ofn.Flags = OFN_OVERWRITEPROMPT; ofn.lpstrDefExt = "ini";
    if (GetSaveFileName(&ofn) == TRUE) {
        CopyFile(iniFilePath, szFile, FALSE);
        SetStatus(hwnd, hStatus, "Contacts exported.", RGB(0, 150, 0), STATUS_CLEAR_MS_SINGLE);
    }
}

// --- Main Window Procedure ---
LRESULT CALLBACK WindowProc(HWND hwnd, UINT uMsg, WPARAM wParam, LPARAM lParam) {
    static HWND hListView, hBtnEdit, hBtnCheck, hBtnWake, hEditName, hEditMac, hEditComment, hBtnAdd, hStatusBar, hLblName, hLblMac, hLblComment, hGroupBox;
    static char editingOriginalName[256];

    switch (uMsg) {
        case WM_CREATE: {
            SetupIniPath();

            // Menu Creation
            HMENU hMenu = CreateMenu();
            HMENU hFileMenu = CreatePopupMenu();
            AppendMenu(hFileMenu, MF_STRING, ID_FILE_IMPORT, "Import contacts\tCtrl+O");
            AppendMenu(hFileMenu, MF_STRING, ID_FILE_EXPORT, "Export contacts\tCtrl+S");
            AppendMenu(hFileMenu, MF_SEPARATOR, 0, NULL);
            AppendMenu(hFileMenu, MF_STRING, ID_FILE_EXIT, "Exit\tEsc");
            AppendMenu(hMenu, MF_POPUP, (UINT_PTR)hFileMenu, "&File");

            // Added About Section
            HMENU hAboutMenu = CreatePopupMenu();
            AppendMenu(hAboutMenu, MF_STRING, ID_HELP_ABOUT, "Version Info");
            AppendMenu(hMenu, MF_POPUP, (UINT_PTR)hAboutMenu, "&About");

            SetMenu(hwnd, hMenu);

            // UI Controls
            hListView = CreateWindow(WC_LISTVIEW, NULL, WS_CHILD | WS_VISIBLE | WS_BORDER | LVS_REPORT | LVS_SHOWSELALWAYS | WS_TABSTOP,
                                    0, 0, 0, 0, hwnd, (HMENU)ID_LISTVIEW, NULL, NULL);
            ListView_SetExtendedListViewStyle(hListView, LVS_EX_FULLROWSELECT | LVS_EX_GRIDLINES);

            LVCOLUMN lvc = {0};
            lvc.mask = LVCF_FMT | LVCF_WIDTH | LVCF_TEXT | LVCF_SUBITEM;
            lvc.fmt = LVCFMT_LEFT;
            lvc.iSubItem = 0; lvc.cx = 150; lvc.pszText = "Computer Name";
            ListView_InsertColumn(hListView, 0, &lvc);
            lvc.iSubItem = 1; lvc.cx = 120; lvc.pszText = "MAC Address";
            ListView_InsertColumn(hListView, 1, &lvc);
            lvc.iSubItem = 2; lvc.cx = 150; lvc.pszText = "Comment";
            ListView_InsertColumn(hListView, 2, &lvc);

            /* Tab order: list, edit, check, wake, then form fields */
            editingOriginalName[0] = '\0';
            hBtnEdit = CreateWindow("BUTTON", "Edit", WS_CHILD | WS_VISIBLE | BS_PUSHBUTTON | WS_TABSTOP,
                                    0, 0, 0, 0, hwnd, (HMENU)ID_BTN_EDIT, NULL, NULL);
            hBtnCheck = CreateWindow("BUTTON", "Check", WS_CHILD | WS_VISIBLE | BS_PUSHBUTTON | WS_TABSTOP,
                                     0, 0, 0, 0, hwnd, (HMENU)ID_BTN_CHECK, NULL, NULL);
            hBtnWake = CreateWindow("BUTTON", "Wake Up", WS_CHILD | WS_VISIBLE | BS_PUSHBUTTON | WS_TABSTOP,
                                    0, 0, 0, 0, hwnd, (HMENU)ID_BTN_WAKE, NULL, NULL);
            hGroupBox = CreateWindow("BUTTON", "New Contact", WS_CHILD | WS_VISIBLE | BS_GROUPBOX, 
                                     0, 0, 0, 0, hwnd, (HMENU)ID_GROUPBOX, NULL, NULL);
            hLblName = CreateWindow("STATIC", "Name:", WS_CHILD | WS_VISIBLE, 0, 0, 0, 0, hwnd, NULL, NULL, NULL);
            hEditName = CreateWindow("EDIT", "", WS_CHILD | WS_VISIBLE | WS_BORDER | ES_AUTOHSCROLL | WS_TABSTOP,
                                     0, 0, 0, 0, hwnd, (HMENU)ID_EDIT_NAME, NULL, NULL);
            hLblMac = CreateWindow("STATIC", "MAC:", WS_CHILD | WS_VISIBLE, 0, 0, 0, 0, hwnd, NULL, NULL, NULL);
            hEditMac = CreateWindow("EDIT", "", WS_CHILD | WS_VISIBLE | WS_BORDER | ES_AUTOHSCROLL | WS_TABSTOP,
                                    0, 0, 0, 0, hwnd, (HMENU)ID_EDIT_MAC, NULL, NULL);
            hLblComment = CreateWindow("STATIC", "Comment:", WS_CHILD | WS_VISIBLE, 0, 0, 0, 0, hwnd, NULL, NULL, NULL);
            hEditComment = CreateWindow("EDIT", "", WS_CHILD | WS_VISIBLE | WS_BORDER | ES_AUTOHSCROLL | WS_TABSTOP,
                                        0, 0, 0, 0, hwnd, (HMENU)ID_EDIT_COMMENT, NULL, NULL);
            hBtnAdd = CreateWindow("BUTTON", "Add", WS_CHILD | WS_VISIBLE | BS_PUSHBUTTON | WS_TABSTOP,
                                   0, 0, 0, 0, hwnd, (HMENU)ID_BTN_ADD, NULL, NULL);
            hStatusBar = CreateWindow(STATUSCLASSNAME, "", WS_CHILD | WS_VISIBLE | SBARS_SIZEGRIP,
                                        0, 0, 0, 0, hwnd, (HMENU)ID_STATUSBAR, NULL, NULL);

            HFONT hFont = (HFONT)GetStockObject(DEFAULT_GUI_FONT);
            SendMessage(hListView, WM_SETFONT, (WPARAM)hFont, TRUE);
            SendMessage(hBtnEdit, WM_SETFONT, (WPARAM)hFont, TRUE);
            SendMessage(hBtnCheck, WM_SETFONT, (WPARAM)hFont, TRUE);
            SendMessage(hBtnWake, WM_SETFONT, (WPARAM)hFont, TRUE);
            SendMessage(hGroupBox, WM_SETFONT, (WPARAM)hFont, TRUE);
            SendMessage(hLblName, WM_SETFONT, (WPARAM)hFont, TRUE);
            SendMessage(hEditName, WM_SETFONT, (WPARAM)hFont, TRUE);
            SendMessage(hLblMac, WM_SETFONT, (WPARAM)hFont, TRUE);
            SendMessage(hEditMac, WM_SETFONT, (WPARAM)hFont, TRUE);
            SendMessage(hLblComment, WM_SETFONT, (WPARAM)hFont, TRUE);
            SendMessage(hEditComment, WM_SETFONT, (WPARAM)hFont, TRUE);
            SendMessage(hBtnAdd, WM_SETFONT, (WPARAM)hFont, TRUE);
            SendMessage(hStatusBar, WM_SETFONT, (WPARAM)hFont, TRUE);

            LoadContacts(hListView);
            SetFocus(hListView);
            return 0;
        }

        case WM_DRAWITEM: {
            LPDRAWITEMSTRUCT pdis = (LPDRAWITEMSTRUCT)lParam;
            if (pdis->CtlID == ID_STATUSBAR) {
                int part = (int)pdis->itemID;
                COLORREF col = statusColor;
                const char* txt = statusText;
                if (part >= 0 && part < g_statusCellCount) {
                    col = g_statusCells[part].color;
                    txt = g_statusCells[part].text;
                }
                SetTextColor(pdis->hDC, col);
                SetBkMode(pdis->hDC, TRANSPARENT); 
                RECT rc = pdis->rcItem;
                rc.left += 5;
                DrawText(pdis->hDC, txt, -1, &rc, DT_SINGLELINE | DT_VCENTER | DT_END_ELLIPSIS);
                return TRUE;
            }
            break;
        }

        case WM_SIZE: {
            int width = LOWORD(lParam);
            int height = HIWORD(lParam);
            SendMessage(hStatusBar, WM_SIZE, wParam, lParam);
            RECT rcStatus; GetWindowRect(hStatusBar, &rcStatus);
            int statusHeight = rcStatus.bottom - rcStatus.top;
            int effectiveHeight = height - statusHeight; 
            int pad = 10;
            int gbHeight = 110;
            int gbY = effectiveHeight - pad - gbHeight;
            int btnWakeHeight = 30;
            int btnWakeY = gbY - pad - btnWakeHeight;
            int lvY = pad;
            int lvHeight = btnWakeY - pad - lvY;
            int innerW = width - (pad * 2);
            int col0 = (innerW * 38) / 100;
            int col1 = (innerW * 30) / 100;
            int col2;
            if (col0 < 100) col0 = 100;
            if (col1 < 110) col1 = 110;
            col2 = innerW - col0 - col1;
            if (col2 < 80) col2 = 80;

            MoveWindow(hListView, pad, lvY, innerW, lvHeight, TRUE);
            {
                int btnGap = 6;
                int btnW = (innerW - btnGap * 2) / 3;
                int x0 = pad;
                int x1 = pad + btnW + btnGap;
                int x2 = pad + (btnW + btnGap) * 2;
                int wLast = pad + innerW - x2;
                MoveWindow(hBtnEdit, x0, btnWakeY, btnW, btnWakeHeight, TRUE);
                MoveWindow(hBtnCheck, x1, btnWakeY, btnW, btnWakeHeight, TRUE);
                MoveWindow(hBtnWake, x2, btnWakeY, wLast, btnWakeHeight, TRUE);
            }
            MoveWindow(hGroupBox, pad, gbY, innerW, gbHeight, TRUE);

            ListView_SetColumnWidth(hListView, 0, col0);
            ListView_SetColumnWidth(hListView, 1, col1);
            ListView_SetColumnWidth(hListView, 2, col2);

            int lblW = 52;
            int editX = pad + 10 + lblW;
            int editW = innerW - (editX - pad) - 80;
            int inputY1 = gbY + 22;
            int inputY2 = gbY + 47;
            int inputY3 = gbY + 72;
            MoveWindow(hLblName, pad + 10, inputY1 + 3, lblW, 20, TRUE);
            MoveWindow(hEditName, editX, inputY1, editW, 20, TRUE);
            MoveWindow(hLblMac, pad + 10, inputY2 + 3, lblW, 20, TRUE);
            MoveWindow(hEditMac, editX, inputY2, editW, 20, TRUE);
            MoveWindow(hLblComment, pad + 10, inputY3 + 3, lblW, 20, TRUE);
            MoveWindow(hEditComment, editX, inputY3, editW, 20, TRUE);
            MoveWindow(hBtnAdd, width - pad - 70, inputY3 - 1, 60, 22, TRUE);
            StatusBarApplyParts(hwnd, hStatusBar);
            return 0;
        }

        case WM_GETMINMAXINFO: {
            MINMAXINFO* mmi = (MINMAXINFO*)lParam;
            mmi->ptMinTrackSize.x = 360; 
            mmi->ptMinTrackSize.y = 440; 
            return 0;
        }

        case WM_TIMER: {
            if (wParam == IDT_STATUS_TIMER) {
                ClearRowHighlightBits(hListView);
                InvalidateRect(hListView, NULL, TRUE);
                SetStatus(hwnd, hStatusBar, "", RGB(0,0,0), 0);
            }
            return 0;
        }

        case WM_NOTIFY: {
            LPNMHDR lpnmh = (LPNMHDR)lParam;
            if (lpnmh->idFrom == ID_LISTVIEW) {
                if (lpnmh->code == NM_CUSTOMDRAW) {
                    LPNMLVCUSTOMDRAW plvcd = (LPNMLVCUSTOMDRAW)lParam;
                    switch (plvcd->nmcd.dwDrawStage) {
                        case CDDS_PREPAINT:
                            return CDRF_NOTIFYITEMDRAW;
                        case CDDS_ITEMPREPAINT:
                            {
                                DWORD lp = (DWORD)plvcd->nmcd.lItemlParam;
                                DWORD hi = lp >> ROW_HI_SHIFT;
                                if (hi == ROW_COLOR_GREEN)
                                    plvcd->clrText = RGB(0, 150, 0);
                                else if (hi == ROW_COLOR_RED)
                                    plvcd->clrText = RGB(200, 0, 0);
                            }
                            return CDRF_NEWFONT;
                        default:
                            break;
                    }
                    return CDRF_DODEFAULT;
                }
                if (lpnmh->code == LVN_COLUMNCLICK) {
                    NMLISTVIEW* nml = (NMLISTVIEW*)lParam;
                    if (nml->iSubItem == g_sortColumn)
                        g_sortAscending = !g_sortAscending;
                    else {
                        g_sortColumn = nml->iSubItem;
                        if (g_sortColumn > 2) g_sortColumn = 2;
                        if (g_sortColumn < 0) g_sortColumn = 0;
                        g_sortAscending = 1;
                    }
                    SortListView(hListView);
                    return 0;
                }
                if (lpnmh->code == NM_DBLCLK) {
                    LPNMITEMACTIVATE lpnmitem = (LPNMITEMACTIVATE)lParam;
                    if (lpnmitem->iItem != -1) {
                        char name[256], mac[256];
                        ListView_GetItemText(hListView, lpnmitem->iItem, 0, name, sizeof(name));
                        ListView_GetItemText(hListView, lpnmitem->iItem, 1, mac, sizeof(mac));
                        WakeOneComputer(hwnd, hStatusBar, hListView, lpnmitem->iItem, name, mac);
                    }
                }
            }
            return 0;
        }

        case WM_CONTEXTMENU: {
            if ((HWND)wParam == hListView) {
                int iPos = ListView_GetNextItem(hListView, -1, LVNI_SELECTED);
                if (iPos != -1) { 
                    POINT pt; GetCursorPos(&pt); 
                    HMENU hPopup = CreatePopupMenu();
                    AppendMenu(hPopup, MF_STRING, ID_CTX_EDIT, "Edit\tE");
                    AppendMenu(hPopup, MF_STRING, ID_BTN_CHECK, "Check\tC");
                    AppendMenu(hPopup, MF_STRING, ID_BTN_WAKE, "Wake Up\tEnter");
                    AppendMenu(hPopup, MF_SEPARATOR, 0, NULL);
                    AppendMenu(hPopup, MF_STRING, ID_CTX_REMOVE, "Remove\tDel");
                    TrackPopupMenu(hPopup, TPM_RIGHTBUTTON, pt.x, pt.y, 0, hwnd, NULL);
                    DestroyMenu(hPopup);
                }
            }
            return 0;
        }

        case WM_COMMAND: {
            WORD cmd = LOWORD(wParam);
            
            if (cmd == ID_HELP_ABOUT) {
                const char* aboutMsg =
                    "WOL Sender v.1.0 by Treblewolf\n\n"
                    "Check: uses Computer Name as a DNS name, NetBIOS name, or IPv4 address "
                    "(same field as the Name; add IP there if the name does not resolve).\n\n"
                    PROJECT_URL "\n\n"
                    "Open this page in your browser?";
                if (MessageBox(hwnd, aboutMsg, "About", MB_YESNO | MB_ICONINFORMATION) == IDYES)
                    ShellExecute(hwnd, "open", PROJECT_URL, NULL, NULL, SW_SHOW);
            }

            if (cmd == ID_CTX_REMOVE) {
                int selCount = ListView_GetSelectedCount(hListView);
                if (selCount > 0) {
                    char** namesToDelete = (char**)malloc(selCount * sizeof(char*));
                    int iPos = -1; int idx = 0;
                    while ((iPos = ListView_GetNextItem(hListView, iPos, LVNI_SELECTED)) != -1) {
                        namesToDelete[idx] = (char*)malloc(256);
                        ListView_GetItemText(hListView, iPos, 0, namesToDelete[idx], 256);
                        idx++;
                    }
                    for (int i = 0; i < selCount; i++) {
                        RemoveContact(namesToDelete[i]);
                        free(namesToDelete[i]);
                    }
                    free(namesToDelete);
                    editingOriginalName[0] = '\0';
                    SetWindowText(hBtnAdd, "Add");
                    LoadContacts(hListView);
                    SetStatus(hwnd, hStatusBar, "Contact(s) removed.", RGB(200, 0, 0), STATUS_CLEAR_MS_SINGLE); 
                }
            }

            if (cmd == ID_BTN_EDIT || cmd == ID_CTX_EDIT) {
                int iPos = ListView_GetNextItem(hListView, -1, LVNI_SELECTED);
                if (iPos != -1) {
                    char name[256]; char mac[256]; char comment[256];
                    ListView_GetItemText(hListView, iPos, 0, name, sizeof(name));
                    ListView_GetItemText(hListView, iPos, 1, mac, sizeof(mac));
                    ListView_GetItemText(hListView, iPos, 2, comment, sizeof(comment));
                    SetWindowText(hEditName, name);
                    SetWindowText(hEditMac, mac);
                    SetWindowText(hEditComment, comment);
                    lstrcpy(editingOriginalName, name);
                    SetWindowText(hBtnAdd, "Save");
                    SetStatus(hwnd, hStatusBar, "Editing... modify fields then Save.", RGB(0, 0, 200), 0); 
                }
            }

            if (cmd == ID_FILE_IMPORT) {
                editingOriginalName[0] = '\0';
                SetWindowText(hBtnAdd, "Add");
                HandleImport(hwnd, hListView, hStatusBar);
            }
            if (cmd == ID_FILE_EXPORT) HandleExport(hwnd, hStatusBar);
            if (cmd == ID_FILE_EXIT) PostQuitMessage(0);
            
            if (cmd == ID_BTN_CHECK) CheckListSelection(hwnd, hListView, hStatusBar);
            if (cmd == ID_BTN_WAKE) WakeListSelection(hwnd, hListView, hStatusBar);

            if (cmd == ID_BTN_ADD) {
                char name[100]; char mac[100]; char comment[256];
                GetWindowText(hEditName, name, sizeof(name));
                GetWindowText(hEditMac, mac, sizeof(mac));
                GetWindowText(hEditComment, comment, sizeof(comment));
                if (strlen(name) > 0 && strlen(mac) > 0) {
                    if (editingOriginalName[0] != '\0') {
                        if (stricmp(editingOriginalName, name) != 0)
                            RemoveContact(editingOriginalName);
                        SaveContact(name, mac, comment);
                        editingOriginalName[0] = '\0';
                        SetWindowText(hBtnAdd, "Add");
                    } else {
                        SaveContact(name, mac, comment);
                    }
                    LoadContacts(hListView);
                    SetWindowText(hEditName, "");
                    SetWindowText(hEditMac, "");
                    SetWindowText(hEditComment, "");
                    SetStatus(hwnd, hStatusBar, "Contact saved.", RGB(0, 150, 0), STATUS_CLEAR_MS_SINGLE); 
                } else {
                    SetStatus(hwnd, hStatusBar, "Error: Name and MAC required.", RGB(200, 0, 0), STATUS_CLEAR_MS_SINGLE); 
                }
            }
            return 0;
        }

        case WM_DESTROY:
            if (g_sortSnap) { free(g_sortSnap); g_sortSnap = NULL; g_sortSnapCap = 0; }
            PostQuitMessage(0);
            return 0;
    }
    return DefWindowProc(hwnd, uMsg, wParam, lParam);
}

// --- Application Entry Point ---
int WINAPI WinMain(HINSTANCE hInstance, HINSTANCE hPrevInstance, LPSTR lpCmdLine, int nCmdShow) {
    InitCommonControlsCompat();

    const char CLASS_NAME[] = "WOLAppClass";
    WNDCLASS wc = {0}; HWND hwnd; MSG msg;

    wc.lpfnWndProc = WindowProc;
    wc.hInstance = hInstance;
    wc.lpszClassName = CLASS_NAME;
    wc.hCursor = LoadCursor(NULL, IDC_ARROW);
    wc.hbrBackground = (HBRUSH)(COLOR_BTNFACE + 1); 
    wc.hIcon = LoadIcon(hInstance, MAKEINTRESOURCE(IDI_APPICON));

    RegisterClass(&wc);

    hwnd = CreateWindowEx(
        WS_EX_CONTROLPARENT, CLASS_NAME, "WOL Sender",
        WS_OVERLAPPEDWINDOW, 
        CW_USEDEFAULT, CW_USEDEFAULT, 420, 520, 
        NULL, NULL, hInstance, NULL
    );

    if (hwnd == NULL) return 0;
    ShowWindow(hwnd, nCmdShow);

    // Keyboard Shortcuts
    ACCEL accelerators[5];
    accelerators[0].fVirt = FCONTROL | FVIRTKEY; accelerators[0].key = 'O'; accelerators[0].cmd = ID_FILE_IMPORT;
    accelerators[1].fVirt = FCONTROL | FVIRTKEY; accelerators[1].key = 'S'; accelerators[1].cmd = ID_FILE_EXPORT;
    accelerators[2].fVirt = FVIRTKEY; accelerators[2].key = VK_ESCAPE; accelerators[2].cmd = ID_FILE_EXIT;
    accelerators[3].fVirt = FVIRTKEY; accelerators[3].key = VK_DELETE; accelerators[3].cmd = ID_CTX_REMOVE;
    accelerators[4].fVirt = FVIRTKEY; accelerators[4].key = VK_RETURN; accelerators[4].cmd = ID_BTN_WAKE;
    HACCEL hAccelTable = CreateAcceleratorTable(accelerators, 5);

    while (GetMessage(&msg, NULL, 0, 0)) {
        if (!TranslateAccelerator(hwnd, hAccelTable, &msg)) {
            if (!IsDialogMessage(hwnd, &msg)) {
                TranslateMessage(&msg);
                DispatchMessage(&msg);
            }
        }
    }
    
    DestroyAcceleratorTable(hAccelTable);
    return 0;
}