// ctSpaces.cpp

// Because I didn't have the paitence to port this from the ground up, I used Gemini 2.5 Pro for heavylifting and filled in the gaps.

#pragma comment(linker,"/manifestdependency:\"type='win32' name='Microsoft.Windows.Common-Controls' version='6.0.0.0' processorArchitecture='*' publicKeyToken='6595b64144ccf1df' language='*'\"")
#pragma comment(lib, "dwmapi.lib")
#pragma comment(lib, "uxtheme.lib")
#pragma comment(lib, "Comctl32.lib")
#pragma comment(lib, "Shlwapi.lib")
#pragma comment(lib, "Ole32.lib")
#pragma comment(lib, "Propsys.lib")
#pragma comment(lib, "gdiplus.lib")
#pragma comment(lib, "Version.lib")

#include <windows.h>
#include <commctrl.h>
#include <shlobj.h>
#include <shlwapi.h>
#include <dwmapi.h>
#include <uxtheme.h>
#include <propkey.h>
#include <propvarutil.h>
#include <gdiplus.h>
#include <shobjidl.h>
#include "ProgressUI.h"
#include "InProc7z.h"

#include <memory>
#include <string>    
#include <vector>    
#include <filesystem>
#include <format>    
#include <ranges>    
#include <functional> 
#include <regex>
#include <thread>
#include <iostream>
#include <fstream>
#include <algorithm>
#include <iomanip> // FIX: Added for std::put_time
#include <mutex> 
#include <map> 
#include <limits> // for std::numeric_limits

#include "resource.h" // For IDR_7ZA and IDR_DEFAULT_7Z

namespace fs=std::filesystem;

//struct IconButtonInfo{
//    int id;
//    HWND hWnd;
//    std::wstring symbol;
//    std::wstring tooltip;
//    std::function<void()> handler;
//};

#pragma pack(push, 2)
struct GRPICONDIRENTRY{
    BYTE  bWidth;
    BYTE  bHeight;
    BYTE  bColorCount;
    BYTE  bReserved;
    WORD  wPlanes;
    WORD  wBitCount;
    DWORD dwBytesInRes;
    WORD  nID;            // RT_ICON id
};
struct GRPICONDIR{
    WORD idReserved;
    WORD idType;
    WORD idCount;
    GRPICONDIRENTRY idEntries[1];
};
#pragma pack(pop)

#pragma pack(push, 1)
typedef struct{
    BYTE bWidth;
    BYTE bHeight;
    BYTE bColorCount;
    BYTE bReserved;
    WORD wPlanes;
    WORD wBitCount;
    DWORD dwBytesInRes;
    DWORD dwImageOffset;
} ICONDIRENTRY;

typedef struct{
    WORD idReserved;
    WORD idType;
    WORD idCount;
    ICONDIRENTRY idEntries[1];
} ICONDIR;
#pragma pack(pop)

enum class ProfileType{
    Standard,
    Default,
    Temporary
};

const std::wstring APP_ALIAS=L"ctSpaces";
const std::wstring APP_VERSION=L"2.3";
const std::wstring APP_VARIANT=L"pre3";
const std::wstring APP_TITLE=std::format(L"{} v{}{}",APP_ALIAS,APP_VERSION,APP_VARIANT);
const std::wstring GUI_CLASS_NAME=L"ctSpacesLauncherClass";
HICON g_hIconBtnTemp=nullptr;
HICON g_hIconBtnConfig=nullptr;
ULONG_PTR g_gdiplusToken;
HINSTANCE g_hInst;
HWND g_hGui=NULL;
HWND g_hComboClient=NULL;
HWND g_hValidationTooltip=NULL;
HWND g_hBtnTmpProfTip=NULL;
HWND g_hBtnConfigTip=NULL;
HWND g_hBtnGo=NULL;
HWND g_hBtnTmpProf=NULL;
HWND g_hBtnConfig=NULL;
HMENU g_hConfigMenu=nullptr;
std::vector<HBITMAP> g_menuBitmaps; // owned; freed on WM_DESTROY
HFONT g_hFont=NULL;
static UINT g_uiDpi=USER_DEFAULT_SCREEN_DPI;
fs::path g_sDataDir;
fs::path g_sEdgePath;
std::wstring g_sLastValidComboText=L"";
std::wstring g_sClientSel=L"";
std::jthread g_watcherThread;
std::mutex g_activeProfilesMutex;
std::atomic<bool> g_isWatcherRunning=false;
struct IconCacheSet{
    std::map<int,HICON> byPx; // requestedPx -> icon handle
};
std::map<std::wstring,IconCacheSet> g_iconCache;

std::mutex g_iconCacheMutex;
IShellLink* shellLink=NULL;
#define WM_APP_TASK_COMPLETE (WM_APP + 1)
#define IDC_BTN_TEMP   200
#define IDC_BTN_CONFIG 201

// Config menu command IDs
#define IDM_CTX_SET_PROFILE_ICON     41001
#define IDM_CTX_REFRESH_PROFILE      41002
#define IDM_CTX_RESET_PROFILE        41003
#define IDM_CTX_DELETE_PROFILE       41004
#define IDM_CTX_EDIT_DEFAULT_PROFILE 41005
#define IDC_ABOUT_ICON   5101
#define IDC_ABOUT_TITLE  5102
#define IDC_ABOUT_BY     5103
#define IDC_ABOUT_GAP    5104
#define IDC_ABOUT_HOME   5105
#define IDC_STATIC_PROMPT 101
#define IDC_GRP_ICON      301
#define IDC_ICON_PREVIEW  302
#define IDC_LBL_ICON      303

static HWND g_hLblIcon=nullptr;


// Icon preview controls/state
static HWND  g_hGrpIcon=nullptr;
static HWND  g_hIconPreview=nullptr;
static HICON g_hIconPreviewHandle=nullptr;

// Menu tooltip state
static HWND     g_hMenuTip=nullptr;
static TOOLINFOW g_menuTi{};
static std::map<UINT,std::wstring> g_menuTipText; // menu id -> tooltip text



static HFONT g_hFontAboutSmall=nullptr;
static HFONT g_hAboutFont=nullptr;
static HICON g_hAboutIcon64=nullptr;
static HICON g_hAboutDlgSmall=nullptr;
static HICON g_hAboutDlgBig=nullptr;

static HFONT GetAboutSmallFont(){
    if(g_hFontAboutSmall) return g_hFontAboutSmall;
    LOGFONTW lf{};
    HFONT base=g_hAboutFont?g_hAboutFont:g_hFont;
    if(base&&GetObjectW(base,sizeof(lf),&lf)==sizeof(lf)){
        lf.lfHeight=(lf.lfHeight*85)/100; // smaller
        g_hFontAboutSmall=CreateFontIndirectW(&lf);
    }
    return g_hFontAboutSmall;
}

static void EnableDpiAwareness(){
    HMODULE user32=GetModuleHandleW(L"user32.dll");
    if(!user32) return;

    auto setCtx=reinterpret_cast<BOOL (WINAPI*)(DPI_AWARENESS_CONTEXT)>(
        GetProcAddress(user32,"SetProcessDpiAwarenessContext")
    );
    if(setCtx){
        if(setCtx(DPI_AWARENESS_CONTEXT_PER_MONITOR_AWARE_V2)) return;
        setCtx(DPI_AWARENESS_CONTEXT_PER_MONITOR_AWARE);
        return;
    }

    auto setAware=reinterpret_cast<BOOL (WINAPI*)(void)>(
        GetProcAddress(user32,"SetProcessDPIAware")
    );
    if(setAware) setAware();
}

static int ScaleByDpi(int value,UINT dpi){
    return MulDiv(value,dpi,USER_DEFAULT_SCREEN_DPI);
}

static HFONT CreateUiFont(UINT dpi){
    const int basePx=15;
    return CreateFontW(
        ScaleByDpi(basePx,dpi),0,0,0,FW_NORMAL,FALSE,FALSE,FALSE,
        DEFAULT_CHARSET,OUT_DEFAULT_PRECIS,CLIP_DEFAULT_PRECIS,DEFAULT_QUALITY,
        DEFAULT_PITCH|FF_MODERN,L"Consolas"
    );
}

static void UpdateComboBoxMetrics(UINT dpi){
    if(!g_hComboClient||!g_hFont) return;

    HDC hdc=GetDC(g_hComboClient);
    if(!hdc) return;

    HFONT old=(HFONT)SelectObject(hdc,g_hFont);
    TEXTMETRICW tm{};
    if(GetTextMetricsW(hdc,&tm)){
        const int itemH=tm.tmHeight+tm.tmExternalLeading+ScaleByDpi(6,dpi);
        SendMessageW(g_hComboClient,CB_SETITEMHEIGHT,(WPARAM)-1,itemH);
        SendMessageW(g_hComboClient,CB_SETITEMHEIGHT,0,itemH);
    }
    SelectObject(hdc,old);
    ReleaseDC(g_hComboClient,hdc);
}

static void UpdateUiFont(HWND hWnd,UINT dpi){
    HFONT hNew=CreateUiFont(dpi);
    if(!hNew) return;

    if(g_hFont) DeleteObject(g_hFont);
    g_hFont=hNew;

    if(hWnd) SendMessageW(hWnd,WM_SETFONT,(WPARAM)g_hFont,TRUE);
    if(hWnd){
        EnumChildWindows(hWnd,[](HWND hwnd,LPARAM lParam)->BOOL{
            SendMessageW(hwnd,WM_SETFONT,(WPARAM)lParam,TRUE);
            return TRUE;
        },(LPARAM)g_hFont);
    }

    if(g_hValidationTooltip) SendMessageW(g_hValidationTooltip,WM_SETFONT,(WPARAM)g_hFont,TRUE);
    if(g_hMenuTip) SendMessageW(g_hMenuTip,WM_SETFONT,(WPARAM)g_hFont,TRUE);
    if(g_hBtnTmpProfTip) SendMessageW(g_hBtnTmpProfTip,WM_SETFONT,(WPARAM)g_hFont,TRUE);
    if(g_hBtnConfigTip) SendMessageW(g_hBtnConfigTip,WM_SETFONT,(WPARAM)g_hFont,TRUE);

    UpdateComboBoxMetrics(dpi);
}

static void LayoutMainGui(HWND hWnd);
static void UpdateIconPreviewForSelection(bool preferListSelection=false);
static void EnsureConfigMenu(HWND hWnd);
static void SetButtonIcon(HWND hBtn,int iconResId,HICON& hStore);

static void ApplyDpiScaling(HWND hWnd,UINT dpi,const RECT* suggestedRect){
    if(dpi==0) dpi=USER_DEFAULT_SCREEN_DPI;
    g_uiDpi=dpi;

    UpdateUiFont(hWnd,dpi);

    if(g_hValidationTooltip){
        SendMessageW(g_hValidationTooltip,TTM_SETMAXTIPWIDTH,0,ScaleByDpi(400,dpi));
    }
    if(g_hMenuTip){
        SendMessageW(g_hMenuTip,TTM_SETMAXTIPWIDTH,0,ScaleByDpi(450,dpi));
    }

    if(suggestedRect){
        SetWindowPos(
            hWnd,nullptr,
            suggestedRect->left,
            suggestedRect->top,
            suggestedRect->right-suggestedRect->left,
            suggestedRect->bottom-suggestedRect->top,
            SWP_NOZORDER|SWP_NOACTIVATE
        );
    }

    SetButtonIcon(g_hBtnTmpProf,IDI_TEMPB,g_hIconBtnTemp);
    SetButtonIcon(g_hBtnConfig,IDI_CFGB,g_hIconBtnConfig);

    if(g_hConfigMenu){
        for(auto hb:g_menuBitmaps) if(hb) DeleteObject(hb);
        g_menuBitmaps.clear();
        DestroyMenu(g_hConfigMenu);
        g_hConfigMenu=nullptr;
        EnsureConfigMenu(hWnd);
    }

    LayoutMainGui(hWnd);
    UpdateIconPreviewForSelection(true);
}

static void ResetAboutFonts(UINT dpi){
    if(g_hAboutFont){
        DeleteObject(g_hAboutFont);
        g_hAboutFont=nullptr;
    }
    if(g_hFontAboutSmall){
        DeleteObject(g_hFontAboutSmall);
        g_hFontAboutSmall=nullptr;
    }

    g_hAboutFont=CreateUiFont(dpi);
    GetAboutSmallFont();
}

struct _7zUiCtx{
    HWND mainWnd=nullptr;
    std::wstring status;
    int lastPercent=-999;
    DWORD lastTick=0;
    bool sentMarquee=false;
};

struct AboutDlgState{
    HICON hIco64=nullptr;
    HFONT hSmall=nullptr;
};

static void __stdcall _7zProgress(void* user,_7zOp op,unsigned int percent,const wchar_t* /*currentItem*/){
    auto* ctx=reinterpret_cast<_7zUiCtx*>(user);
    if(!ctx||!ctx->mainWnd||!IsWindow(ctx->mainWnd)) return;

    // Unknown progress -> marquee once.
    if(percent<0){
        if(!ctx->sentMarquee){
            ctx->sentMarquee=true;
            ctx->lastPercent=-1;
            ctx->lastTick=GetTickCount();
            ProgressUI_PostUpdate(ctx->mainWnd,ctx->status,-1);
        }
        return;
    }

    if(percent<0) percent=0;
    if(percent>100) percent=100;

    DWORD now=GetTickCount();

    // Throttle: don't spam the GUI thread.
    if(percent==ctx->lastPercent&&(now-ctx->lastTick)<150) return;
    if((now-ctx->lastTick)<33&&percent<100) return; // ~30fps max

    ctx->lastTick=now;
    ctx->lastPercent=percent;
    ProgressUI_PostUpdate(ctx->mainWnd,ctx->status,percent);
}


const std::vector<std::wstring> aKeepDefault={
    L"Local State",
    L"Last Version",
    L"Last Browser",
    L"FirstLaunchAfterInstallation",
    L"First Run",
    L"ctSpaces",
    L"DevToolsActivePort",
    L"Default\\Shortcuts",
    L"Default\\Shortcuts-journal", L"Default\\Secure Preferences",
    L"Default\\Preferences",
    L"Default\\Favicons-journal",
    L"Default\\Favicons",
    L"Default\\Bookmarks",
    L"Default\\Extension State",
    L"Default\\Extensions",
    L"Default\\Local Extension Settings",
    L"Default\\Asset Store",
    L"Default\\Extension Rules",
    L"Default\\Extension Scripts"
};
const std::vector<std::wstring> aKeepActive={
    L"client.ico",
    L"client.png",
    L"ctSpaces",
    L"Local State",
    L"Last Version",
    L"Last Browser",
    L"FirstLaunchAfterInstallation",
    L"First Run",
    L"DevToolsActivePort",
    L"Default\\History",
    L"Default\\Shortcuts",
    L"Default\\Shortcuts-journal",
    L"Default\\Secure Preferences",
    L"Default\\Preferences",
    L"Default\\Favicons-journal",
    L"Default\\Favicons",
    L"Default\\Bookmarks",
    L"Default\\Extension State",
    L"Default\\Extension Rules",
    L"Default\\Extension Scripts",
    L"Default\\Extensions",
    L"Default\\Asset Store",
    L"Default\\Local Extension Settings",
    L"CertificateRevocation",
    L"AutoLaunchProtocolsComponent",
    L"Default\\History",
    L"Default\\Web Data",
    L"Default\\Web Data-journal",
    L"Default\\Login Data",
    L"Default\\Login Data-journal",
    L"Default\\Favicons",
    L"Default\\Favicons-journal",
    L"Default\\MediaDeviceSalts",
    L"Default\\MediaDeviceSalts-journal",
    L"Default\\CdmStorage.db",
    L"Default\\CdmStorage.db-journal",
    L"Default\\DIPS",
    L"Default\\DIPS-journal",
    L"Default\\Local Storage",
    L"Default\\WebStorage",
    //L"Default\\Service Worker\\Database",
    L"Default\\ClientCertificates",
    L"Default\\blob_storage",
    L"Default\\Session Storage",
    L"Default\\IndexedDB",
    L"Default\\Network",
    L"Default\\Sessions"
};
/*
PKIMetadata\
WebAssistDataBase
Web Data
Web Data-journal
Prefrences
Login Data
Login Data-journal
History
History-journal
Sessions\
Network\
Asset Store\

*/
std::map<std::wstring,DWORD> g_activeProfiles;
//std::vector<IconButtonInfo> g_iconButtons;
std::vector<std::jthread> g_reaperThreads;

void GuiProfOpen();
void GuiSetIcon();
void GuiProfReset();
void GuiProfUpd();
void GuiOpenDef();
void GuiOpenTmp();
bool IsValidFilenameChar(wchar_t c);
std::wstring SanitizeName(const std::wstring& name);
void UpdateClientsComboBox();
void SetUiState(bool enabled);
void LaunchAndManageProfile(const std::wstring& clientName,bool isTemp,bool isDefault);
bool ExtractResourceToFile(UINT resourceID,const fs::path& destPath);
bool extDef(const fs::path& profileDataPath,const wchar_t* statusText);
void CleanupProfile(const fs::path& profilePath,const std::vector<std::wstring>& keepList);
void SetWindowAppId(HWND hWnd,const std::wstring& appId);
LRESULT CALLBACK WndProc(HWND,UINT,WPARAM,LPARAM);
ATOM MyRegisterClass(HINSTANCE hInstance);
BOOL InitInstance(HINSTANCE,int);
std::wstring AnsiToWide(const std::string& str);
std::vector<std::wstring> GetSupportedImageTypes();
bool ConvertImageToIcon(const fs::path& sourceImagePath,const fs::path& destIconPath);
bool SaveIconsToFile(const fs::path& filePath,std::vector<HICON>& icons,bool compressLargeImages=true,size_t startIndex=0);
std::vector<BYTE> CompressBitmapToPng(HBITMAP hBitmap);
CLSID GetEncoderClsid(const WCHAR* format);
HICON Create32BitHICON(HICON hIcon);
bool IsAlphaBitmap(HBITMAP hBitmap);
void EnsureWatcherIsRunning();
void WatcherThread();
void ReaperThread(DWORD pid,std::wstring clientName,ProfileType type);
DWORD LaunchProfile(const std::wstring& clientName,bool isTemp,bool isDefault);
bool FindEdgePath();
void TerminateAllProfiles();
std::wstring GetExeVersion(const fs::path& filePath);
bool chkUpdate();
bool doInstall();
HWND CreateToolTip(HWND toolHWND,HWND hDlg,PTSTR pszText);
static void PostTaskComplete(const std::wstring& name);
static void LaunchProfileAsync(const std::wstring& name,bool isTemp,bool isDefault);
void GuiProfDel();
inline void EnsureMouseVisible();
inline void FocusClientEdit();
static void ShowAboutDialog();
static UINT ShowConfigMenuFromButton(HWND hWnd);
static int IcoDim(BYTE b){return (b==0)?256:(int)b;}
static HICON  LoadIconResBestDownscale(HINSTANCE hInst,int groupIconResId,int cxDesired,int cyDesired);
static HICON  LoadIconFromIcoBestDownscale(const fs::path& icoPath,int pxDesired);
static HBITMAP LoadMenuBitmapFromIconRes(int iconResId,UINT dpi,int cx,int cy);
static HFONT CreateSmallerFontFrom(HFONT baseFont,int pxHeight,UINT dpi);
static HICON ScaleIconDown_HQ(HICON hSrc,int dstCx,int dstCy);
static bool GetIconSizePx(HICON hIcon,int& w,int& h);
static void EnsureMenuTooltips(HWND hWnd);
static void HideMenuTooltip();
static void HandleMenuSelect(HWND hWnd,WPARAM wParam,LPARAM lParam);
static void UpdateConfigMenuEnabledState();
static std::wstring GetSelectedClientNameSanitized(bool preferListSelection=false);
static fs::path GetProfileDirFromName(const std::wstring& name);

int APIENTRY wWinMain(_In_ HINSTANCE hInstance,_In_opt_ HINSTANCE,_In_ LPWSTR lpCmdLine,_In_ int nCmdShow){
    CoInitializeEx(NULL,COINIT_APARTMENTTHREADED|COINIT_DISABLE_OLE1DDE);
    EnableDpiAwareness();
    PWSTR path=NULL;
    if(SUCCEEDED(SHGetKnownFolderPath(FOLDERID_LocalAppData,0,NULL,&path))){
        g_sDataDir=fs::path(path)/"InfinitySys"/"ctSpaces";
        CoTaskMemFree(path);
    }
    std::filesystem::create_directories(g_sDataDir);
    wchar_t currentExePathStr[MAX_PATH];
    GetModuleFileNameW(NULL,currentExePathStr,MAX_PATH);
    fs::path currentExePath(currentExePathStr);
    const wchar_t* mutexName=L"Global\\{E19C159D-62C3-4412-A0A3-1A55A67C8C56}";
    HANDLE hMutex=CreateMutexW(NULL,TRUE,mutexName);
    if(hMutex!=NULL&&GetLastError()==ERROR_ALREADY_EXISTS){
        HWND hExistingWnd=FindWindowW(GUI_CLASS_NAME.c_str(),NULL);
        if(hExistingWnd){
            DWORD existingProcId;
            GetWindowThreadProcessId(hExistingWnd,&existingProcId);
            HANDLE hExistingProcess=OpenProcess(PROCESS_QUERY_LIMITED_INFORMATION,FALSE,existingProcId);
            if(hExistingProcess){
                wchar_t existingExePathStr[MAX_PATH]={0};
                DWORD pathSize=MAX_PATH;
                QueryFullProcessImageNameW(hExistingProcess,0,existingExePathStr,&pathSize);
                CloseHandle(hExistingProcess);
                if(!fs::equivalent(currentExePath,existingExePathStr)){
                    MessageBoxW(
                        NULL,
                        L"A different version of ctSpaces is already running.\n\nPlease close the other instance before installing or running this version.",
                        L"Update Conflict",
                        MB_OK|MB_ICONWARNING
                    );
                    ReleaseMutex(hMutex);
                    CloseHandle(hMutex);
                    CoUninitialize();
                    return 0;
                }
            }
            ShowWindow(hExistingWnd,SW_RESTORE);
            SetForegroundWindow(hExistingWnd);
        }
        ReleaseMutex(hMutex);
        CloseHandle(hMutex);
        CoUninitialize();
        return 0;
    }
    bool shouldRun=false;
    fs::path installedExePath=g_sDataDir/L"ctSpaces.exe";
    if(fs::exists(installedExePath)){
        shouldRun=chkUpdate();
    } else{
        shouldRun=doInstall();
    }
    if(!shouldRun){
        CoUninitialize();
        return 0;
    }
    Gdiplus::GdiplusStartupInput gdiplusStartupInput;
    if(!FindEdgePath()){
        MessageBox(NULL,L"Microsoft Edge could not be found in standard installation locations. Please ensure it is installed.",L"Application Error",MB_OK|MB_ICONERROR);
        CoUninitialize();
        return 1;
    }
    Gdiplus::GdiplusStartup(&g_gdiplusToken,&gdiplusStartupInput,NULL);
    ExtractResourceToFile(IDR_DEFPROF,g_sDataDir/"Default.7z");
    INITCOMMONCONTROLSEX icex={sizeof(INITCOMMONCONTROLSEX), ICC_WIN95_CLASSES|ICC_PROGRESS_CLASS};
    InitCommonControlsEx(&icex);

    //InitCommonControlsEx(&icex);
    ProgressUI_Init(hInstance);
    _7zSetHInstance(hInstance);


    //bool useTrdLayout=(wcsstr(lpCmdLine,L"~!Trd:P")!=nullptr);
    //g_iconButtons={
    //    { 200, NULL, L"Ico", L"Set Profile Icon", GuiSetIcon },
    //    { 201, NULL, L"Rst", L"Reset Selected Profile to Default", GuiProfReset },
    //    { 202, NULL, L"Upd", L"Update Selected Profile with Default (Overwrites Profile /w Default)", GuiProfUpd },
    //    { 203, NULL, L"Def", L"Modify Default Profile", GuiOpenDef },
    //    { 204, NULL, L"Tmp", L"Launch temporary profile", GuiOpenTmp },
    //    { 205, NULL, L"Del", L"Delete Profile", GuiProfDel },
    //};
    //if(useTrdLayout){
    //    std::vector<IconButtonInfo> trdLayoutButtons={
    //        g_iconButtons[0],
    //        g_iconButtons[4],
    //        g_iconButtons[2],
    //        g_iconButtons[1],
    //        g_iconButtons[3],
    //        g_iconButtons[5]
    //    };
    //    g_iconButtons.swap(trdLayoutButtons);
    //}

    MyRegisterClass(hInstance);
    if(!InitInstance(hInstance,nCmdShow)){
        return FALSE;
    }

    MSG msg;
    while(GetMessage(&msg,nullptr,0,0)){
        // intercept Enter when focus is in the combo or its edit
        if(msg.message==WM_KEYDOWN&&msg.wParam==VK_RETURN){
            HWND hFocus=GetFocus();
            if(hFocus==g_hComboClient||
               (hFocus&&GetParent(hFocus)==g_hComboClient)){
                // pretend the Go button was pressed
                SendMessage(g_hGui,
                            WM_COMMAND,
                            MAKELONG(IDOK,BN_CLICKED),
                            (LPARAM)g_hBtnGo);
                // don't let the dialog logic eat this
                continue;
            }
        }

        if(!IsDialogMessage(g_hGui,&msg)){
            TranslateMessage(&msg);
            DispatchMessage(&msg);
        }
    }
    //while(GetMessage(&msg,nullptr,0,0)){
    //    if(!IsDialogMessage(g_hGui,&msg)){
    //        TranslateMessage(&msg);
    //        DispatchMessage(&msg);
    //    }
    //}
    Gdiplus::GdiplusShutdown(g_gdiplusToken);
    CoUninitialize();
    return (int)msg.wParam;
}

ATOM MyRegisterClass(HINSTANCE hInstance){
    WNDCLASSEXW wcex={};
    wcex.cbSize=sizeof(WNDCLASSEXW);
    wcex.style=CS_HREDRAW|CS_VREDRAW;
    wcex.lpfnWndProc=WndProc;
    wcex.hInstance=hInstance;
    wcex.hCursor=LoadCursor(nullptr,IDC_ARROW);
    wcex.hbrBackground=(HBRUSH)(COLOR_BTNFACE+1);
    wcex.lpszClassName=GUI_CLASS_NAME.c_str();
    wcex.lpszMenuName=NULL;
    return RegisterClassExW(&wcex);
}

BOOL InitInstance(HINSTANCE hInstance,int nCmdShow){
    g_hInst=hInstance;
    const int baseGuiW=256+64+16+8;
    const int baseGuiH=64+32+16+8+4;
    const int baseGuiM=4;
    const int baseCtrlW=(baseGuiW-baseGuiM*2);

    UINT dpi=GetDpiForSystem();
    g_uiDpi=dpi;

    const int guiClientW=ScaleByDpi(baseGuiW,dpi);
    const int guiClientH=ScaleByDpi(baseGuiH,dpi);
    RECT wr{0,0,guiClientW,guiClientH};

    const DWORD style=WS_OVERLAPPED|WS_CAPTION|WS_SYSMENU|WS_MINIMIZEBOX;
    const DWORD exStyle=0;
    auto adjustForDpi=reinterpret_cast<BOOL (WINAPI*)(LPRECT,DWORD,BOOL,DWORD,UINT)>(
        GetProcAddress(GetModuleHandleW(L"user32.dll"),"AdjustWindowRectExForDpi")
    );
    if(adjustForDpi){
        adjustForDpi(&wr,style,FALSE,exStyle,dpi);
    } else{
        AdjustWindowRectEx(&wr,style,FALSE,exStyle);
    }

    g_hGui=CreateWindowW(
        GUI_CLASS_NAME.c_str(),
        APP_TITLE.c_str(),
        style,
        CW_USEDEFAULT,
        0,
        wr.right-wr.left,
        wr.bottom-wr.top,
        nullptr,
        nullptr,
        hInstance,
        nullptr
    );
    if(!g_hGui) return FALSE;

    dpi=GetDpiForWindow(g_hGui);
    g_uiDpi=dpi;
    g_hFont=CreateUiFont(dpi);

    const int iGuiM=ScaleByDpi(baseGuiM,dpi);
    const int iGuiCtrlW=ScaleByDpi(baseCtrlW,dpi);
    const int labelH=ScaleByDpi(17,dpi);
    HICON hAppIcon=LoadIcon(hInstance,MAKEINTRESOURCE(IDI_CTSPACES));
    if(hAppIcon){
        SendMessage(g_hGui,WM_SETICON,ICON_BIG,(LPARAM)hAppIcon);
        SendMessage(g_hGui,WM_SETICON,ICON_SMALL,(LPARAM)hAppIcon);
    }
    BOOL isDarkMode=TRUE;
    DwmSetWindowAttribute(g_hGui,DWMWA_USE_IMMERSIVE_DARK_MODE,&isDarkMode,sizeof(isDarkMode));
    CreateWindowW(L"STATIC",L"Select or type the client name:",WS_CHILD|WS_VISIBLE,iGuiM,iGuiM,iGuiCtrlW,labelH,g_hGui,(HMENU)101,hInstance,nullptr);
    g_hValidationTooltip=CreateWindowEx(WS_EX_TOPMOST,TOOLTIPS_CLASS,NULL,TTS_BALLOON|TTS_NOPREFIX|TTS_ALWAYSTIP,CW_USEDEFAULT,CW_USEDEFAULT,CW_USEDEFAULT,CW_USEDEFAULT,g_hGui,NULL,g_hInst,NULL);
    SendMessageW(g_hGui,WM_SETFONT,(WPARAM)g_hFont,TRUE);
    SendMessageW(g_hValidationTooltip,WM_SETFONT,(WPARAM)g_hFont,TRUE);
    SendMessage(g_hValidationTooltip,TTM_SETMAXTIPWIDTH,0,ScaleByDpi(400,dpi));
    g_hComboClient=CreateWindowW(L"COMBOBOX",L"",WS_CHILD|WS_VISIBLE|CBS_DROPDOWN|CBS_AUTOHSCROLL|WS_VSCROLL,iGuiM,labelH+iGuiM*2,iGuiCtrlW-iGuiM*4,ScaleByDpi(150,dpi),g_hGui,(HMENU)102,hInstance,nullptr);
    UpdateComboBoxMetrics(dpi);
    TOOLINFOW tic={sizeof(TOOLINFOW)};
    tic.uFlags=TTF_SUBCLASS|TTF_TRANSPARENT|TTF_TRACK;
    tic.hwnd=g_hGui;
    tic.hinst=g_hInst;
    tic.uId=(UINT_PTR)g_hComboClient;
    tic.lpszText=LPSTR_TEXTCALLBACK;
    SendMessage(g_hValidationTooltip,TTM_ADDTOOL,0,(LPARAM)&tic);
    const int goW=ScaleByDpi(64,dpi);
    const int goH=ScaleByDpi(33,dpi);
    g_hBtnGo=CreateWindowW(L"BUTTON",L"Go",WS_CHILD|WS_VISIBLE|BS_DEFPUSHBUTTON,(iGuiCtrlW/2)-ScaleByDpi(32,dpi)-ScaleByDpi(4,dpi),ScaleByDpi(50,dpi)+iGuiM,goW,goH,g_hGui,(HMENU)IDOK,hInstance,nullptr);
    //HWND hToolTip=CreateWindowEx(0,TOOLTIPS_CLASS,NULL,TTS_ALWAYSTIP|TTS_NOPREFIX,CW_USEDEFAULT,CW_USEDEFAULT,CW_USEDEFAULT,CW_USEDEFAULT,g_hGui,NULL,g_hInst,NULL);
    RECT rc{};
    GetClientRect(g_hGui,&rc);
    const int iBtnS=ScaleByDpi(24,dpi);
    const int iBtnT=rc.bottom-iGuiM-iBtnS;
    int iBtnL=rc.right-iGuiM-iBtnS;
    g_hBtnConfig=CreateWindowW(L"BUTTON",L"",WS_CHILD|WS_VISIBLE|BS_ICON,iBtnL,iBtnT,iBtnS,iBtnS,g_hGui,(HMENU)(INT_PTR)201,g_hInst,nullptr);
    g_hBtnTmpProf=CreateWindowW(L"BUTTON",L"",WS_CHILD|WS_VISIBLE|BS_ICON,iBtnL-iBtnS-iGuiM,iBtnT,iBtnS,iBtnS,g_hGui,(HMENU)(INT_PTR)200,g_hInst,nullptr);
    SetButtonIcon(g_hBtnTmpProf,IDI_TEMPB,g_hIconBtnTemp);
    SetButtonIcon(g_hBtnConfig,IDI_CFGB,g_hIconBtnConfig);
    g_hBtnTmpProfTip=CreateToolTip(g_hBtnTmpProf,g_hGui,(LPWSTR)L"Launch temporary profile");
    g_hBtnConfigTip=CreateToolTip(g_hBtnConfig,g_hGui,(LPWSTR)L"Options / Profile actions");
    EnsureConfigMenu(g_hGui);

    // Create the Icon group + preview (positions set by LayoutMainGui)
    g_hGrpIcon=CreateWindowW(
        L"BUTTON",L"",
        WS_CHILD|WS_VISIBLE|BS_GROUPBOX,
        0,0,10,10,
        g_hGui,(HMENU)IDC_GRP_ICON,g_hInst,nullptr
    );
    g_hLblIcon=CreateWindowW(
        L"STATIC",L"Icon",
        WS_CHILD|WS_VISIBLE|SS_LEFT|SS_NOPREFIX,
        0,0,10,10,
        g_hGui,(HMENU)IDC_LBL_ICON,g_hInst,nullptr
    );
    g_hIconPreview=CreateWindowW(
        L"STATIC",L"",
        WS_CHILD|WS_VISIBLE|SS_ICON,
        0,0,10,10,
        g_hGui,(HMENU)IDC_ICON_PREVIEW,g_hInst,nullptr
    );
    SendMessageW(g_hGrpIcon,WM_SETFONT,(WPARAM)g_hFont,TRUE);
    SendMessageW(g_hLblIcon,WM_SETFONT,(WPARAM)g_hFont,TRUE);
    SendMessageW(g_hIconPreview,WM_SETFONT,(WPARAM)g_hFont,TRUE);
    EnsureMenuTooltips(g_hGui);

    // Apply layout (also makes Go square + centered)
    LayoutMainGui(g_hGui);

    // Initial preview
    UpdateIconPreviewForSelection();
    //for(size_t ix=0; ix<g_iconButtons.size(); ++ix){
    //    int iBtnT=iGuiH-((iBtnS+iBtnM)*(2-((ix/4)%((g_iconButtons.size()/4)*4))))-32-6;
    //    int xPos=iBtnL+((iBtnS+iBtnM+10)*(static_cast<int>(ix%4)+1));
    //    g_iconButtons[ix].hWnd=CreateWindowW(L"BUTTON",g_iconButtons[ix].symbol.c_str(),WS_CHILD|WS_VISIBLE,xPos,iBtnT,iBtnS+10,iBtnS,g_hGui,(HMENU)(INT_PTR)g_iconButtons[ix].id,g_hInst,nullptr);
    //    CreateToolTip(g_iconButtons[ix].hWnd,g_hGui,(LPWSTR)g_iconButtons[ix].tooltip.c_str());
    //}
    EnumChildWindows(g_hGui,[](HWND hwnd,LPARAM lParam)->BOOL{SendMessage(hwnd,WM_SETFONT,(WPARAM)lParam,TRUE);return TRUE;},(LPARAM)g_hFont);
    UpdateClientsComboBox();
    SetFocus(g_hComboClient);
    ShowWindow(g_hGui,nCmdShow);
    UpdateWindow(g_hGui);
    return TRUE;
}

LRESULT CALLBACK WndProc(HWND hWnd,UINT message,WPARAM wParam,LPARAM lParam){
    switch(message){
    case WM_COMMAND:{
        int wmId=LOWORD(wParam);
        int wmEvent=HIWORD(wParam);
        if(wmId==102){
            if(wmEvent==CBN_EDITCHANGE){
                wchar_t buffer[256];
                GetWindowText(g_hComboClient,buffer,256);
                std::wstring currentText=buffer;
                std::wstring sanitizedText;
                bool hasInvalidChar=false;
                for(wchar_t c:currentText){
                    if(IsValidFilenameChar(c)){
                        sanitizedText+=c;
                    } else{
                        hasInvalidChar=true;
                    }
                }
                if(GetAsyncKeyState(VK_BACK)&0x8000||GetAsyncKeyState(VK_DELETE)&0x8000){
                    if(!hasInvalidChar){
                        g_sLastValidComboText=currentText;
                        TOOLINFOW ti={sizeof(TOOLINFOW)};
                        ti.hwnd=g_hGui;
                        ti.uId=(UINT_PTR)g_hComboClient;
                        SendMessage(g_hValidationTooltip,TTM_TRACKACTIVATE,FALSE,(LPARAM)&ti);
                    }
                    UpdateIconPreviewForSelection(false);
                    return 0;
                }
                if(hasInvalidChar){
                    SetWindowText(g_hComboClient,g_sLastValidComboText.c_str());
                    SendMessage(g_hComboClient,CB_SETEDITSEL,0,MAKELPARAM(g_sLastValidComboText.length(),g_sLastValidComboText.length()));
                    TOOLINFOW ti={sizeof(TOOLINFOW)};
                    ti.hwnd=g_hGui;
                    ti.uFlags=TTF_ABSOLUTE;
                    ti.uId=(UINT_PTR)g_hComboClient;
                    ti.lpszText=(LPWSTR)L"A client name can't contain any of the following characters:\n    \\ / : * ? \" < > |";
                    SendMessage(g_hValidationTooltip,TTM_UPDATETIPTEXT,0,(LPARAM)&ti);
                    RECT rect;
                    GetWindowRect(g_hComboClient,&rect);
                    const int tipPad=ScaleByDpi(4,g_uiDpi);
                    SendMessage(g_hValidationTooltip,TTM_TRACKPOSITION,0,MAKELPARAM(rect.left+tipPad,rect.bottom-tipPad));
                    SendMessage(g_hValidationTooltip,TTM_TRACKACTIVATE,TRUE,(LPARAM)&ti);
                    SendMessage(g_hComboClient,CB_SHOWDROPDOWN,FALSE,0);
                } else{
                    g_sLastValidComboText=currentText;
                    TOOLINFOW ti={sizeof(TOOLINFOW)};
                    ti.hwnd=g_hGui;
                    ti.uId=(UINT_PTR)g_hComboClient;
                    SendMessage(g_hValidationTooltip,TTM_TRACKACTIVATE,FALSE,(LPARAM)&ti);
                    if(currentText.length()>0){
                        int ciMatchIdx=-1;
                        int csMatchIdx=-1;
                        wchar_t buf[256];
                        for(int i=0;i<(int)SendMessage(g_hComboClient,CB_GETCOUNT,0,0);++i){
                            SendMessage(g_hComboClient,CB_GETLBTEXT,i,(LPARAM)buf);
                            std::wstring listItem=buf;
                            if(csMatchIdx==-1&&listItem.size()>=currentText.size()&&listItem.compare(0,currentText.size(),currentText)==0){
                                csMatchIdx=i;
                                break;
                            }
                            if(ciMatchIdx==-1&&listItem.size()>=currentText.size()&&_wcsnicmp(listItem.c_str(),currentText.c_str(),currentText.size())==0){
                                ciMatchIdx=i;
                            }
                        }
                        if(csMatchIdx!=-1){
                            SendMessage(g_hComboClient,CB_GETLBTEXT,csMatchIdx,(LPARAM)buf);
                            SendMessage(g_hComboClient,CB_SETCURSEL,csMatchIdx,0);
                            SendMessage(g_hComboClient,CB_SHOWDROPDOWN,TRUE,0);
                            EnsureMouseVisible();
                            SetWindowText(g_hComboClient,buf);
                        } else if(ciMatchIdx!=-1){
                            SendMessage(g_hComboClient,CB_SETCURSEL,ciMatchIdx,0);
                            SendMessage(g_hComboClient,CB_SHOWDROPDOWN,TRUE,0);
                            EnsureMouseVisible();
                            SetWindowText(g_hComboClient,currentText.c_str());
                        } else{
                            SendMessage(g_hComboClient,CB_SHOWDROPDOWN,FALSE,0);
                            SetWindowText(g_hComboClient,currentText.c_str());
                        }
                        SendMessage(g_hComboClient,CB_SETEDITSEL,0,MAKELPARAM((DWORD)currentText.length(),(DWORD)-1));
                    } else{
                        SendMessage(g_hComboClient,CB_SHOWDROPDOWN,FALSE,0);
                    }

                }
                UpdateIconPreviewForSelection(false);
                return 0;
            }
            if(wmEvent==CBN_SELCHANGE){
                UpdateIconPreviewForSelection(true);
                return 0;
            }
            if(wmEvent==CBN_SELENDOK){
                UpdateIconPreviewForSelection(true);  // also fine (commit point)
                return 0;
            }
            if(wmEvent==CBN_CLOSEUP){
                UpdateIconPreviewForSelection(true);
                return 0;
            }
            if(wmEvent==CBN_KILLFOCUS){
                UpdateIconPreviewForSelection(false);
                return 0;
            }
        } else if(wmId==IDOK){
            GuiProfOpen();
        } else if(wmId==IDC_BTN_TEMP&&wmEvent==BN_CLICKED){
            GuiOpenTmp();
        } else if(wmId==IDC_BTN_CONFIG&&wmEvent==BN_CLICKED){
            UINT cmd=ShowConfigMenuFromButton(hWnd);
            if(cmd) SendMessageW(hWnd,WM_COMMAND,cmd,0);
        } else if(wmId==IDM_CTX_SET_PROFILE_ICON){
            GuiSetIcon();
        } else if(wmId==IDM_CTX_REFRESH_PROFILE){
            GuiProfUpd();
        } else if(wmId==IDM_CTX_RESET_PROFILE){
            GuiProfReset();
        } else if(wmId==IDM_CTX_DELETE_PROFILE){
            GuiProfDel();
        } else if(wmId==IDM_CTX_EDIT_DEFAULT_PROFILE){
            GuiOpenDef();
        } else if(wmId==IDM_ABOUT){
            ShowAboutDialog();
        } else{
            //auto it=std::find_if(g_iconButtons.begin(),g_iconButtons.end(),[wmId](const auto& btn){
            //    return btn.id==wmId;
            //});
            //if(it!=g_iconButtons.end()&&it->handler){
            //    it->handler();
            //}
        }
        return 0;
    }
    case WM_CTLCOLORBTN:
    case WM_CTLCOLORSTATIC: {
        HDC hdcControl=(HDC)wParam;
        SetTextColor(hdcControl,GetThemeSysColor(NULL,COLOR_BTNTEXT));
        SetBkColor(hdcControl,GetThemeSysColor(NULL,COLOR_BTNFACE));
        return (INT_PTR)GetThemeSysColorBrush(NULL,COLOR_BTNFACE);
    }
    case WM_APP_PROGRESS_SHOW: {
        auto* p=reinterpret_cast<CtProgressPayload*>(lParam);
        if(p){
            ProgressUI_Show(g_hGui?g_hGui:hWnd,p->text,p->percent);
            delete p;
        }
        return 0;
    }
    case WM_APP_PROGRESS_UPDATE: {
        auto* p=reinterpret_cast<CtProgressPayload*>(lParam);
        if(p){
            if(p->text.empty())
                ProgressUI_Update(p->percent);
            else
                ProgressUI_Update(p->text,p->percent);
            delete p;
        }
        return 0;
    }
    case WM_APP_PROGRESS_HIDE: {
        ProgressUI_Hide();
        return 0;
    }
    case WM_APP_TASK_COMPLETE: {
        std::wstring clientName;
        if(lParam){
            std::wstring* pName=reinterpret_cast<std::wstring*>(lParam);
            clientName=*pName;
            delete pName;
        } else{
            wchar_t clientNameBuffer[256];
            GetWindowText(g_hComboClient,clientNameBuffer,256);
            clientName=SanitizeName(clientNameBuffer);
        }
        SetUiState(true);

        if(!clientName.empty()&&clientName!=L"Temp"&&clientName!=L"Default"){
            g_sClientSel=clientName;
            UpdateClientsComboBox();
            LRESULT idx=SendMessage(g_hComboClient,CB_FINDSTRINGEXACT,(WPARAM)-1,(LPARAM)clientName.c_str());
            if(idx!=CB_ERR){
                SendMessage(g_hComboClient,CB_SETCURSEL,(WPARAM)idx,0);
            } else{
                SetWindowText(g_hComboClient,clientName.c_str());
            }
        }
        UpdateIconPreviewForSelection(true);
        FocusClientEdit();
        return 0;
    }
    case WM_CLOSE: {
        bool hasActiveProfiles=false;
        {
            std::lock_guard<std::mutex> lock(g_activeProfilesMutex);
            if(!g_activeProfiles.empty()){
                hasActiveProfiles=true;
            }
        }
        if(hasActiveProfiles){
            int result=MessageBox(
                hWnd,
                L"There are active profiles running. Would you like to exit and close all open profiles?",
                L"Confirm Exit",
                MB_OKCANCEL|MB_ICONWARNING
            );
            if(result==IDOK){
                TerminateAllProfiles();
                DestroyWindow(hWnd);
            }
        } else{
            DestroyWindow(hWnd);
        }
        return 0;
    }
    case WM_DESTROY: {
        for(auto hbmp:g_menuBitmaps){
            if(hbmp) DeleteObject(hbmp);
        }
        g_menuBitmaps.clear();
        if(g_hConfigMenu){
            DestroyMenu(g_hConfigMenu);
            g_hConfigMenu=nullptr;
            
        }
        if(g_hIconBtnTemp){
            DestroyIcon(g_hIconBtnTemp);
            g_hIconBtnTemp=nullptr;
            
        }
        if(g_hIconBtnConfig){
            DestroyIcon(g_hIconBtnConfig);
            g_hIconBtnConfig=nullptr;
            
        }
        if(g_hIconPreviewHandle){
            DestroyIcon(g_hIconPreviewHandle);
            g_hIconPreviewHandle=nullptr;
        }
        if(g_hMenuTip){
            DestroyWindow(g_hMenuTip);
            g_hMenuTip=nullptr;
        }
        if(g_hBtnTmpProfTip){
            DestroyWindow(g_hBtnTmpProfTip);
            g_hBtnTmpProfTip=nullptr;
        }
        if(g_hBtnConfigTip){
            DestroyWindow(g_hBtnConfigTip);
            g_hBtnConfigTip=nullptr;
        }
        if(g_hFontAboutSmall){ DeleteObject(g_hFontAboutSmall); g_hFontAboutSmall=nullptr; }
        if(g_hFont) DeleteObject(g_hFont);
        PostQuitMessage(0);
        break;
    }
    case WM_ACTIVATE: {
        if(LOWORD(wParam)!=WA_INACTIVE){
            FocusClientEdit();
        }
        return 0;
    }
    case WM_SETFOCUS: {
        FocusClientEdit();
        return 0;
    }
    case WM_DPICHANGED: {
        const UINT dpi=HIWORD(wParam);
        const RECT* rc=reinterpret_cast<RECT*>(lParam);
        ApplyDpiScaling(hWnd,dpi,rc);
        return 0;
    }
    case WM_SIZE:
        LayoutMainGui(hWnd);
        return 0;

    case WM_MENUSELECT:
        EnsureMenuTooltips(hWnd);
        HandleMenuSelect(hWnd,wParam,lParam);
        return 0;

    case WM_EXITMENULOOP:
        HideMenuTooltip();
        return 0;

    default:
        return DefWindowProc(hWnd,message,wParam,lParam);
    }
    return 0;
}

void GuiProfOpen(){
    SetUiState(false);
    wchar_t clientNameBuffer[256];
    GetWindowText(g_hComboClient,clientNameBuffer,256);
    std::wstring clientName=SanitizeName(clientNameBuffer);
    if(clientName.empty()){
        MessageBox(g_hGui,L"Please select or enter a valid client name.",L"Input Error",MB_OK|MB_ICONWARNING);
        SetUiState(true);
        return;
    }

    {
        std::lock_guard<std::mutex> lock(g_activeProfilesMutex);
        if(g_activeProfiles.count(clientName)){
            MessageBox(g_hGui,L"This profile is already open.",L"Already Running",MB_OK|MB_ICONINFORMATION);
            SetUiState(true);
            return;
        }
    }
    LaunchProfileAsync(clientName,false,false);
}

void GuiSetIcon(){
    SetUiState(false);

    wchar_t clientNameBuffer[256];
    GetWindowText(g_hComboClient,clientNameBuffer,256);
    std::wstring clientName=SanitizeName(clientNameBuffer);

    if(clientName.empty()){
        MessageBox(g_hGui,L"Please select a client first.",L"Warning",MB_OK|MB_ICONWARNING);
        SetUiState(true);
        return;
    }

    std::wstring filter;
    std::wstring allSupportedExtensions=L"*.ico";
    std::vector<std::wstring> types=GetSupportedImageTypes();
    for(const auto& type:types){
        allSupportedExtensions+=L";*."+type;
    }

    filter+=L"Supported Image Files ("+allSupportedExtensions+L")";
    filter+=L'\0';
    filter+=allSupportedExtensions;
    filter+=L'\0';
    filter+=L"Icon Files (*.ico)";
    filter+=L'\0';
    filter+=L"*.ico";
    filter+=L'\0';
    filter+=L"All Files (*.*)";
    filter+=L'\0';
    filter+=L"*.*";
    filter+=L'\0';
    filter+=L'\0';

    wchar_t szFile[MAX_PATH]={0};
    OPENFILENAMEW ofn={0};
    ofn.lStructSize=sizeof(ofn);
    ofn.hwndOwner=g_hGui;
    ofn.lpstrFile=szFile;
    ofn.nMaxFile=sizeof(szFile)/sizeof(wchar_t);
    ofn.lpstrFilter=filter.c_str();
    ofn.nFilterIndex=1;
    ofn.Flags=OFN_PATHMUSTEXIST|OFN_FILEMUSTEXIST;

    if(GetOpenFileNameW(&ofn)){
        fs::path sourcePath(ofn.lpstrFile);
        fs::path destIconPath=g_sDataDir/"Sites"/clientName/"client.ico";
        fs::create_directories(destIconPath.parent_path());

        bool success=false;
        std::wstring errorDetails;

        try{
            if(_wcsicmp(sourcePath.extension().c_str(),L".ico")==0){
                fs::copy_file(sourcePath,destIconPath,fs::copy_options::overwrite_existing);
                success=true;
            } else{
                success=ConvertImageToIcon(sourcePath,destIconPath);
                if(!success) errorDetails=L"Could not convert image to icon format.";
            }
        } catch(const fs::filesystem_error& e){
            success=false;
            errorDetails=AnsiToWide(e.what());
        }

        if(success){
            MessageBox(g_hGui,L"Icon has been configured.",L"Success",MB_OK|MB_ICONINFORMATION);

            // IMPORTANT: nuke cached icons for this client (all sizes)
            {
                std::lock_guard<std::mutex> cacheLock(g_iconCacheMutex);
                auto it=g_iconCache.find(clientName);
                if(it!=g_iconCache.end()){
                    for(auto& [px,hIcon]:it->second.byPx){
                        if(hIcon) DestroyIcon(hIcon);
                    }
                    g_iconCache.erase(it);
                }
            }
        } else{
            std::wstring errorMsg=L"Failed to set icon.";
            if(!errorDetails.empty()) errorMsg+=L"\n\nDetails: "+errorDetails;
            MessageBox(g_hGui,errorMsg.c_str(),L"Error",MB_OK|MB_ICONERROR);
        }
    }
    UpdateIconPreviewForSelection(false);
    SetUiState(true);
}


void GuiProfReset(){
    SetUiState(false);
    wchar_t clientNameBuffer[256];
    GetWindowText(g_hComboClient,clientNameBuffer,256);
    std::wstring clientName=SanitizeName(clientNameBuffer);
    if(clientName.empty()){
        MessageBox(g_hGui,L"Please select a client first.",L"Warning",MB_OK|MB_ICONWARNING);
        SetUiState(true);
        return;
    }
    std::lock_guard<std::mutex> lock(g_activeProfilesMutex);
    if(g_activeProfiles.count(clientName)){
        MessageBox(g_hGui,L"Cannot reset a profile that is currently active.",L"Action Denied",MB_OK|MB_ICONWARNING);
        SetUiState(true);
        return;
    }
    if(MessageBox(g_hGui,L"This will completely delete and reset the profile. Are you sure?",L"Confirm Reset",MB_YESNO|MB_ICONQUESTION)==IDYES){
        fs::path sData=g_sDataDir/"Sites"/clientName;
        try{
            if(fs::exists(sData)){
                fs::remove_all(sData);
            }
            extDef(sData,L"Resetting Profile...");
            MessageBox(g_hGui,L"Profile has been reset.",L"Success",MB_OK|MB_ICONINFORMATION);
        } catch(...){}
        UpdateClientsComboBox();
        LRESULT selectionIndex=SendMessage(g_hComboClient,CB_FINDSTRINGEXACT,(WPARAM)-1,(LPARAM)clientName.c_str());
        if(selectionIndex!=CB_ERR){
            SendMessage(g_hComboClient,CB_SETCURSEL,(WPARAM)selectionIndex,0);
        }
    }
    SetUiState(true);
}

void GuiProfUpd(){
    SetUiState(false);
    wchar_t clientNameBuffer[256];
    GetWindowText(g_hComboClient,clientNameBuffer,256);
    std::wstring clientName=SanitizeName(clientNameBuffer);
    if(clientName.empty()){
        MessageBox(g_hGui,L"Please select a client first.",L"Warning",MB_OK|MB_ICONWARNING);
        SetUiState(true);
        return;
    }
    std::lock_guard<std::mutex> lock(g_activeProfilesMutex);
    if(g_activeProfiles.count(clientName)){
        MessageBox(g_hGui,L"Cannot update a profile that is currently active.",L"Action Denied",MB_OK|MB_ICONWARNING);
        SetUiState(true);
        return;
    }
    if(MessageBox(g_hGui,L"This will update/overwrite the profile. Are you sure?",L"Confirm Reset",MB_YESNO|MB_ICONQUESTION)==IDYES){
        std::thread([clientName](){
            fs::path sData=g_sDataDir/"Sites"/clientName;
            try{
                if(fs::exists(sData)){
                    fs::remove_all(sData);
                }
                extDef(sData,L"Updating Profile...");
            } catch(...){}
            PostMessage(g_hGui,WM_APP_TASK_COMPLETE,0,0);
        }).detach();
    }
}

void GuiOpenDef(){
    SetUiState(false);
    std::lock_guard<std::mutex> lock(g_activeProfilesMutex);
    if(g_activeProfiles.count(L"Default")){
        MessageBox(g_hGui,L"The Default profile is already open.",L"Already Running",MB_OK|MB_ICONINFORMATION);
        SetUiState(true);
        return;
    }
    LaunchProfileAsync(L"Default",false,true);
}

void GuiOpenTmp(){
    SetUiState(false);
    std::lock_guard<std::mutex> lock(g_activeProfilesMutex);
    if(g_activeProfiles.count(L"Temp")){
        MessageBox(g_hGui,L"A temporary profile is already open.",L"Already Running",MB_OK|MB_ICONINFORMATION);
        SetUiState(true);
        return;
    }
    LaunchProfileAsync(L"Temp",true,false);
}
struct EnumData{
    DWORD processId;
    std::vector<HWND> windows;
};
BOOL CALLBACK EnumWindowsCallback(HWND hWnd,LPARAM lParam){
    EnumData* pData=(EnumData*)lParam;
    DWORD processId=0;
    GetWindowThreadProcessId(hWnd,&processId);
    if(pData->processId==processId&&IsWindowVisible(hWnd)&&GetWindowTextLength(hWnd)>0){
        pData->windows.push_back(hWnd);
    }
    return TRUE;
}
DWORD LaunchProfile(const std::wstring& clientName,bool isTemp,bool isDefault){
    fs::path profilePath;
    if(isTemp) profilePath=g_sDataDir/"Temp";
    else if(isDefault) profilePath=g_sDataDir/"Default";
    else profilePath=g_sDataDir/"Sites"/clientName;
    if(!fs::exists(profilePath/"ctSpaces")){
        const wchar_t* status=
            isDefault?L"Loading Default Profile...":
            isTemp?L"Loading Temp Profile...":
            L"Updating Profile...";

        if(!extDef(profilePath,status)){

            MessageBox(NULL,L"Error: An error occurred when extracting the profile.",APP_TITLE.c_str(),MB_OK|MB_ICONERROR);
            return 0;
        }
    }
    std::wstring cmdLine=std::format(L"\"{}\" --user-data-dir=\"{}\" --no-first-run --disable-sync --disable-features=SyncPromo  --edge-skip-compat-layer-relaunch --no-service-autorun",g_sEdgePath.c_str(),profilePath.c_str());
    STARTUPINFOW si={sizeof(si)};
    PROCESS_INFORMATION pi={};
    if(!CreateProcessW(NULL,&cmdLine[0],NULL,NULL,FALSE,0,NULL,NULL,&si,&pi)){
        MessageBox(NULL,L"Failed to launch Microsoft Edge.",APP_TITLE.c_str(),MB_OK|MB_ICONERROR);
        return 0;
    }
    CloseHandle(pi.hThread);
    return pi.dwProcessId;
}

void UpdateClientsComboBox(){
    SendMessage(g_hComboClient,CB_RESETCONTENT,0,0);
    fs::path sitesDir=g_sDataDir/"Sites";
    if(fs::exists(sitesDir)&&fs::is_directory(sitesDir)){
        for(const auto& entry:fs::directory_iterator(sitesDir)){
            if(entry.is_directory()){
                SendMessage(g_hComboClient,CB_ADDSTRING,0,(LPARAM)entry.path().filename().c_str());
            }
        }
    }

}
std::wstring SanitizeName(const std::wstring& name){
    std::wstring sanitized=name;
    const std::wstring whitespace=L" \t\n\r\f\v";
    sanitized.erase(0,sanitized.find_first_not_of(whitespace));
    sanitized.erase(sanitized.find_last_not_of(whitespace)+1);
    std::wregex invalidChars(LR"([\\/:*?"<>|])");
    sanitized=std::regex_replace(sanitized,invalidChars,L"");
    while(!sanitized.empty()&&sanitized.back()==L'.'){
        sanitized.pop_back();
    }
    std::wregex reservedNames(L"^(CON|PRN|AUX|NUL|COM[1-9]|LPT[1-9])$",std::regex::icase);
    if(std::regex_match(sanitized,reservedNames)){
        return L"";
    }
    return sanitized;
}

bool IsValidFilenameChar(wchar_t c){
    const std::wstring invalidChars=L"\\/:*?\"<>|";
    return invalidChars.find(c)==std::wstring::npos;
}

bool ExtractResourceToFile(UINT resourceID,const fs::path& destPath){
    if(fs::exists(destPath)) return true;
    HRSRC hRes=FindResource(g_hInst,MAKEINTRESOURCE(resourceID),L"BINARY");
    if(!hRes) return false;
    HGLOBAL hResLoad=LoadResource(g_hInst,hRes);
    if(!hResLoad) return false;
    void* pRes=LockResource(hResLoad);
    if(!pRes) return false;
    DWORD dwSize=SizeofResource(g_hInst,hRes);
    std::ofstream outFile(destPath,std::ios::binary);
    if(!outFile) return false;
    outFile.write(static_cast<char*>(pRes),dwSize);
    return outFile.good();
}




void SetUiState(bool enabled){
    EnableWindow(g_hComboClient,enabled);
    EnableWindow(g_hBtnGo,enabled);
    EnableWindow(g_hBtnTmpProf,enabled);
    EnableWindow(g_hBtnConfig,enabled);
    //for(const auto& btn:g_iconButtons){
    //    EnableWindow(btn.hWnd,enabled);
    //}
    if(enabled)
        FocusClientEdit();
}

bool extDef(const fs::path& profileDataPath,const wchar_t* statusText){
    // Default archive lives in your data dir (still extracted from resource once).
    const fs::path default7zPath=g_sDataDir/L"Default.7z";

    if(!fs::exists(default7zPath)){
        if(!ExtractResourceToFile(IDR_DEFPROF,default7zPath)){
            MessageBoxW(g_hGui,L"Failed to extract Default.7z resource.",L"Error",MB_OK|MB_ICONERROR);
            return false;
        }
    }

    try{
        fs::create_directories(profileDataPath);
    } catch(...){
        // best-effort; _7zExtra_7z will fail if path can't be created
    }

    // Always create our marker file on success (your "profile exists" sentinel).
    const fs::path markerPath=profileDataPath/L"ctSpaces";

    _7zUiCtx ctx;
    ctx.mainWnd=g_hGui;
    ctx.status=(statusText&&*statusText)?statusText:L"Loading Default Profile...";

    ProgressUI_PostShow(ctx.mainWnd,ctx.status,-1);

    HRESULT hr=_7zExtra_7z(
        default7zPath.c_str(),
        profileDataPath.c_str(),
        &_7zProgress,
        &ctx
    );

    ProgressUI_PostHide(ctx.mainWnd);

    if(FAILED(hr)){
        MessageBoxW(g_hGui,L"Failed to extract Default.7z (InProc7z).",L"Error",MB_OK|MB_ICONERROR);
        return false;
    }

    // marker file
    try{
        std::wofstream markerFile(markerPath);
        markerFile.close();
    } catch(...){
        // non-fatal
    }

    return true;
}


void SetWindowAppId(HWND hWnd,const std::wstring& appId){
    IPropertyStore* pps;
    if(SUCCEEDED(SHGetPropertyStoreForWindow(hWnd,IID_PPV_ARGS(&pps)))){
        PROPVARIANT pv;
        if(SUCCEEDED(InitPropVariantFromString(appId.c_str(),&pv))){
            pps->SetValue(PKEY_AppUserModel_ID,pv);
            PropVariantClear(&pv);
        }
        pps->Release();
    }
}

void CleanupProfile(const fs::path& profilePath,const std::vector<std::wstring>& keepList){
    std::vector<fs::path> toDelete;
    try{
        for(const auto& entry:fs::recursive_directory_iterator(profilePath)){
            fs::path relativePath=fs::relative(entry.path(),profilePath);
            bool shouldKeep=false;
            for(const auto& keepItem:keepList){
                if(relativePath.wstring().find(keepItem)!=std::wstring::npos){
                    shouldKeep=true;
                    break;
                }
            }
            if(!shouldKeep){
                toDelete.push_back(entry.path());
            }
        }
    } catch(const fs::filesystem_error&){
    }
    std::sort(toDelete.rbegin(),toDelete.rend());
    for(const auto& path:toDelete){
        try{
            if(fs::is_regular_file(path)||fs::is_symlink(path)){
                fs::remove(path);
            } else if(fs::is_directory(path)&&fs::is_empty(path)){
                fs::remove(path);
            }
        } catch(...){ /* ignore errors */ }
    }
}

std::wstring AnsiToWide(const std::string& str){
    if(str.empty()){
        return std::wstring();
    }
    int size_needed=MultiByteToWideChar(CP_ACP,0,str.c_str(),-1,NULL,0);
    if(size_needed==0){
        return std::wstring();
    }
    std::wstring wstrTo(size_needed,0);
    MultiByteToWideChar(CP_ACP,0,str.c_str(),-1,&wstrTo[0],size_needed);
    if(!wstrTo.empty()&&wstrTo.back()==L'\0'){
        wstrTo.pop_back();
    }
    return wstrTo;
}

std::vector<std::wstring> GetSupportedImageTypes(){
    std::vector<std::wstring> supportedTypes;
    UINT numDecoders=0,sizeBytes=0;
    Gdiplus::GetImageDecodersSize(&numDecoders,&sizeBytes);
    if(sizeBytes==0) return supportedTypes;

    std::vector<BYTE> buf(sizeBytes);
    auto* p=(Gdiplus::ImageCodecInfo*)buf.data();
    Gdiplus::GetImageDecoders(numDecoders,sizeBytes,p);

    for(UINT i=0; i<numDecoders; ++i){
        std::wstring extensions(p[i].FilenameExtension);
        std::wstringstream ss(extensions);
        std::wstring ext;
        while(std::getline(ss,ext,L';')){
            if(ext.rfind(L"*.",0)==0) ext=ext.substr(2);
            std::transform(ext.begin(),ext.end(),ext.begin(),::towlower);
            if(ext!=L"ico"&&std::find(supportedTypes.begin(),supportedTypes.end(),ext)==supportedTypes.end()){
                supportedTypes.push_back(ext);
            }
        }
    }
    return supportedTypes;
}

static std::unique_ptr<Gdiplus::Bitmap> ResizeBitmap_StretchSquare(Gdiplus::Bitmap* src,int dstW,int dstH){
    auto dst=std::make_unique<Gdiplus::Bitmap>(dstW,dstH,PixelFormat32bppARGB);
    Gdiplus::Graphics g(dst.get());
    g.SetInterpolationMode(Gdiplus::InterpolationModeHighQualityBicubic);
    g.SetPixelOffsetMode(Gdiplus::PixelOffsetModeHighQuality);

    // AutoIt _GDIPlus_ImageResize() behavior here is effectively a stretch to requested size.
    g.DrawImage(src,0,0,dstW,dstH);
    return dst;
}

bool ConvertImageToIcon(const fs::path& sourceImagePath,const fs::path& destIconPath){
    std::unique_ptr<Gdiplus::Bitmap> src(Gdiplus::Bitmap::FromFile(sourceImagePath.c_str()));
    if(!src||src->GetLastStatus()!=Gdiplus::Ok) return false;

    const UINT w=src->GetWidth();
    const UINT h=src->GetHeight();
    const int iSize=(int)max(w,h);

    // Stretch to square (matches your AutoIt-style behavior)
    auto scaled=ResizeBitmap_StretchSquare(src.get(),iSize,iSize);
    if(!scaled||scaled->GetLastStatus()!=Gdiplus::Ok) return false;

    // Requested canonical sizes
    const int sizes[]={16,32,48,64,128,256};

    std::vector<HICON> icons;
    icons.reserve(6);

    for(int s:sizes){
        if(s>iSize) continue; // NEVER upscale; only generate sizes <= source square
        auto bmp=ResizeBitmap_StretchSquare(scaled.get(),s,s);
        if(!bmp||bmp->GetLastStatus()!=Gdiplus::Ok) return false;

        HICON hIco=NULL;
        if(bmp->GetHICON(&hIco)==Gdiplus::Ok&&hIco){
            icons.push_back(hIco);
        }
    }

    if(icons.empty()) return false;

    // IMPORTANT: don't skip index 0 (this was dropping your 16px)
    bool ok=SaveIconsToFile(destIconPath,icons, /*compress*/true, /*startIndex*/0);

    for(HICON hIco:icons) DestroyIcon(hIco);
    return ok;
}


HICON Create32BitHICON(HICON hIcon){
    ICONINFOEXW iconInfo={sizeof(ICONINFOEXW)};
    if(!GetIconInfoExW(hIcon,&iconInfo)){
        return NULL;
    }
    int width=0;
    int height=0;
    if(iconInfo.hbmColor){
        BITMAP bmp={0};
        if(GetObject(iconInfo.hbmColor,sizeof(BITMAP),&bmp)){
            width=bmp.bmWidth;
            height=bmp.bmHeight;
        }
    } else if(iconInfo.hbmMask){
        BITMAP bmp={0};
        if(GetObject(iconInfo.hbmMask,sizeof(BITMAP),&bmp)){
            width=bmp.bmWidth;
            height=bmp.bmHeight/2;
        }
    }
    if(width==0||height==0){
        width=GetSystemMetrics(SM_CXICON);
        height=GetSystemMetrics(SM_CYICON);
    }
    if(iconInfo.hbmColor){
        BITMAP bmp={0};
        GetObject(iconInfo.hbmColor,sizeof(BITMAP),&bmp);
        if(bmp.bmBitsPixel==32&&IsAlphaBitmap(iconInfo.hbmColor)){
            HICON hCopy=CopyIcon(hIcon);
            DeleteObject(iconInfo.hbmColor);
            DeleteObject(iconInfo.hbmMask);
            return hCopy;
        }
    }
    HDC hdcScreen=GetDC(NULL);
    HDC hdcMem=CreateCompatibleDC(hdcScreen);
    BITMAPINFO bi={0};
    bi.bmiHeader.biSize=sizeof(BITMAPINFOHEADER);
    bi.bmiHeader.biWidth=width;
    bi.bmiHeader.biHeight=-height;
    bi.bmiHeader.biPlanes=1;
    bi.bmiHeader.biBitCount=32;
    bi.bmiHeader.biCompression=BI_RGB;
    void* pBits;
    HBITMAP hDib=CreateDIBSection(hdcScreen,&bi,DIB_RGB_COLORS,&pBits,NULL,0);
    HBITMAP hOldBmp=(HBITMAP)SelectObject(hdcMem,hDib);
    RECT rc={0, 0, width, height};
    HBRUSH hBrush=CreateSolidBrush(RGB(0,0,0));
    FillRect(hdcMem,&rc,hBrush);
    DeleteObject(hBrush);
    DrawIconEx(hdcMem,0,0,hIcon,width,height,0,NULL,DI_NORMAL);
    SelectObject(hdcMem,hOldBmp);
    DeleteDC(hdcMem);
    ReleaseDC(NULL,hdcScreen);
    HBITMAP hMask=CreateBitmap(width,height,1,1,NULL);
    ICONINFO newIconInfo={0};
    newIconInfo.fIcon=TRUE;
    newIconInfo.hbmColor=hDib;
    newIconInfo.hbmMask=hMask;
    HICON hNewIcon=CreateIconIndirect(&newIconInfo);
    DeleteObject(hDib);
    DeleteObject(hMask);
    DeleteObject(iconInfo.hbmColor);
    DeleteObject(iconInfo.hbmMask);
    return hNewIcon;
}

CLSID GetEncoderClsid(const WCHAR* format){
    UINT  num=0;
    UINT  size=0;
    Gdiplus::GetImageEncodersSize(&num,&size);
    if(size==0) return {0};
    auto pImageCodecInfo=std::make_unique<Gdiplus::ImageCodecInfo[]>(size);
    if(!pImageCodecInfo) return {0};
    GetImageEncoders(num,size,pImageCodecInfo.get());
    for(UINT j=0; j<num; ++j){
        if(wcscmp(pImageCodecInfo[j].MimeType,format)==0){
            return pImageCodecInfo[j].Clsid;
        }
    }
    return {0};
}
bool IsAlphaBitmap(HBITMAP hBitmap){
    BITMAP bmp{};
    if(!GetObject(hBitmap,sizeof(bmp),&bmp)) return false;
    if(bmp.bmBitsPixel!=32) return false;

    const int dataSize=bmp.bmWidthBytes*bmp.bmHeight;
    if(dataSize<=0) return false;

    std::vector<BYTE> pixels((size_t)dataSize);
    if(GetBitmapBits(hBitmap,dataSize,pixels.data())==0) return false;

    // AutoIt __AlphaProc(): considers alpha "present" if any A != 0
    for(int i=3; i<dataSize; i+=4){
        if(pixels[i]!=0) return true;
    }
    return false;
}

std::vector<BYTE> CompressBitmapToPng(HBITMAP hBitmap){
    std::unique_ptr<Gdiplus::Bitmap> bitmap(Gdiplus::Bitmap::FromHBITMAP(hBitmap,NULL));
    if(!bitmap||bitmap->GetLastStatus()!=Gdiplus::Ok){
        return {};
    }
    CLSID pngClsid=GetEncoderClsid(L"image/png");
    IStream* pStream=NULL;
    if(CreateStreamOnHGlobal(NULL,TRUE,&pStream)!=S_OK){
        return {};
    }
    if(bitmap->Save(pStream,&pngClsid,NULL)!=Gdiplus::Ok){
        pStream->Release();
        return {};
    }
    ULARGE_INTEGER streamSize;
    pStream->Seek({},STREAM_SEEK_END,&streamSize);
    std::vector<BYTE> buffer(streamSize.QuadPart);
    pStream->Seek({},STREAM_SEEK_SET,NULL);
    ULONG bytesRead;
    pStream->Read(buffer.data(),buffer.size(),&bytesRead);
    pStream->Release();
    if(bytesRead!=buffer.size()){
        return {};
    }
    return buffer;
}
#pragma pack(push, 1)
struct ICONDIRHDR{
    WORD idReserved;
    WORD idType;
    WORD idCount;
};
#pragma pack(pop)

static bool WriteAll(HANDLE hFile,const void* data,DWORD cb){
    const BYTE* p=(const BYTE*)data;
    DWORD total=0;
    while(total<cb){
        DWORD w=0;
        if(!WriteFile(hFile,p+total,cb-total,&w,nullptr)) return false;
        if(w==0) return false;
        total+=w;
    }
    return true;
}

bool SaveIconsToFile(const fs::path& filePath,std::vector<HICON>& icons,bool compressLargeImages,size_t startIndex){
    if(icons.size()<=startIndex) return false;

    const size_t count=icons.size()-startIndex;

    HANDLE hFile=CreateFileW(
        filePath.c_str(),
        GENERIC_WRITE,
        FILE_SHARE_READ,   // matches AutoIt share=4
        nullptr,
        CREATE_ALWAYS,
        FILE_ATTRIBUTE_NORMAL,
        nullptr
    );
    if(hFile==INVALID_HANDLE_VALUE) return false;

    std::vector<ICONDIRENTRY> entries(count);
    ICONDIRHDR hdr{};
    hdr.idReserved=0;
    hdr.idType=1;
    hdr.idCount=(WORD)count;

    // AutoIt writes placeholder header+entries, streams image data, then seeks back and rewrites.
    if(!WriteAll(hFile,&hdr,sizeof(hdr))){
        CloseHandle(hFile);
        DeleteFileW(filePath.c_str());
        return false;
    }
    {
        std::vector<BYTE> zero(sizeof(ICONDIRENTRY)*count,0);
        if(!WriteAll(hFile,zero.data(),(DWORD)zero.size())){
            CloseHandle(hFile);
            DeleteFileW(filePath.c_str());
            return false;
        }
    }

    DWORD offset=(DWORD)(sizeof(ICONDIRHDR)+sizeof(ICONDIRENTRY)*count);

    std::vector<HICON> tempIcons; // like AutoIt $aTemp[]
    tempIcons.reserve(count);

    auto fail=[&](void){
        for(HICON h:tempIcons) DestroyIcon(h);
        CloseHandle(hFile);
        DeleteFileW(filePath.c_str());
    };

    for(size_t i=0; i<count; ++i){
        HICON hIcon=icons[startIndex+i];
        bool didTemp=false;

        for(int attempt=0; attempt<2; ++attempt){
            ICONINFO ii{};
            if(!GetIconInfo(hIcon,&ii)){
                fail();
                return false;
            }

            // AutoIt: CopyImage(..., 0x2008 = LR_CREATEDIBSECTION | LR_COPYDELETEORG)
            HBITMAP hMask=(HBITMAP)CopyImage(ii.hbmMask,IMAGE_BITMAP,0,0,LR_CREATEDIBSECTION|LR_COPYDELETEORG);
            HBITMAP hColor=(HBITMAP)CopyImage(ii.hbmColor,IMAGE_BITMAP,0,0,LR_CREATEDIBSECTION|LR_COPYDELETEORG);

            // If CopyImage failed, originals weren't deleted; clean them.
            if(!hMask||!hColor){
                if(ii.hbmMask) DeleteObject(ii.hbmMask);
                if(ii.hbmColor) DeleteObject(ii.hbmColor);
                if(hMask) DeleteObject(hMask);
                if(hColor) DeleteObject(hColor);
                fail();
                return false;
            }

            DIBSECTION dsMask{};
            DIBSECTION dsColor{};
            if(GetObject(hMask,sizeof(dsMask),&dsMask)!=sizeof(dsMask)||
               GetObject(hColor,sizeof(dsColor),&dsColor)!=sizeof(dsColor)){
                DeleteObject(hMask);
                DeleteObject(hColor);
                fail();
                return false;
            }

            DWORD maskSize=dsMask.dsBmih.biSizeImage;
            DWORD colorSize=dsColor.dsBmih.biSizeImage;
            if(maskSize==0)  maskSize=(DWORD)(dsMask.dsBm.bmWidthBytes*dsMask.dsBm.bmHeight);
            if(colorSize==0) colorSize=(DWORD)(dsColor.dsBm.bmWidthBytes*dsColor.dsBm.bmHeight);

            const int bpp=dsColor.dsBm.bmBitsPixel;
            if(bpp!=16&&bpp!=24&&bpp!=32){
                DeleteObject(hMask);
                DeleteObject(hColor);
                fail();
                return false;
            }

            // AutoIt: for 32bpp, if not alpha -> try Create32BitHICON once and restart this icon.
            if(bpp==32){
                if(!IsAlphaBitmap(hColor)){
                    if(!didTemp){
                        HICON hNew=Create32BitHICON(hIcon);
                        if(hNew){
                            tempIcons.push_back(hNew);
                            hIcon=hNew;
                            didTemp=true;
                            DeleteObject(hMask);
                            DeleteObject(hColor);
                            continue; // retry
                        }
                    }
                } else{
                    // AutoIt compress condition: (colorSize >= 256*256*4) AND compress
                    if(compressLargeImages&&colorSize>=(256u*256u*4u)){
                        std::vector<BYTE> png=CompressBitmapToPng(hColor);
                        if(!png.empty()){
                            ICONDIRENTRY& e=entries[i];
                            e.bWidth=(dsColor.dsBm.bmWidth<256)?(BYTE)dsColor.dsBm.bmWidth:0;
                            e.bHeight=(dsColor.dsBm.bmHeight<256)?(BYTE)dsColor.dsBm.bmHeight:0;
                            e.bColorCount=0;
                            e.bReserved=0;
                            e.wPlanes=1;
                            e.wBitCount=32;
                            e.dwBytesInRes=(DWORD)png.size();
                            e.dwImageOffset=offset;

                            if(!WriteAll(hFile,png.data(),(DWORD)png.size())){
                                DeleteObject(hMask);
                                DeleteObject(hColor);
                                fail();
                                return false;
                            }
                            offset+=(DWORD)png.size();

                            DeleteObject(hMask);
                            DeleteObject(hColor);
                            break; // done with this icon
                        }
                    }
                }
            }

            // Uncompressed path: BITMAPINFOHEADER + color (XOR) + mask (AND)
            BITMAPINFOHEADER bih{};
            bih.biSize=40;
            bih.biWidth=dsColor.dsBm.bmWidth;
            bih.biHeight=2*dsColor.dsBm.bmHeight;
            bih.biPlanes=1;
            bih.biBitCount=(WORD)bpp;
            bih.biCompression=BI_RGB;
            bih.biSizeImage=colorSize+maskSize;

            ICONDIRENTRY& e=entries[i];
            e.bWidth=(dsColor.dsBm.bmWidth<256)?(BYTE)dsColor.dsBm.bmWidth:0;
            e.bHeight=(dsColor.dsBm.bmHeight<256)?(BYTE)dsColor.dsBm.bmHeight:0;
            e.bColorCount=0;
            e.bReserved=0;
            e.wPlanes=1;
            e.wBitCount=(WORD)bpp;
            e.dwBytesInRes=40+colorSize+maskSize;
            e.dwImageOffset=offset;

            if(!WriteAll(hFile,&bih,40)||
               !WriteAll(hFile,dsColor.dsBm.bmBits,colorSize)||
               !WriteAll(hFile,dsMask.dsBm.bmBits,maskSize)){
                DeleteObject(hMask);
                DeleteObject(hColor);
                fail();
                return false;
            }

            offset+=40+colorSize+maskSize;

            DeleteObject(hMask);
            DeleteObject(hColor);
            break;
        }
    }

    // Seek back and write final header + dir entries (AutoIt does this)
    SetFilePointer(hFile,0,nullptr,FILE_BEGIN);
    if(!WriteAll(hFile,&hdr,sizeof(hdr))||
       !WriteAll(hFile,entries.data(),(DWORD)(entries.size()*sizeof(ICONDIRENTRY)))){
        for(HICON h:tempIcons) DestroyIcon(h);
        CloseHandle(hFile);
        DeleteFileW(filePath.c_str());
        return false;
    }

    CloseHandle(hFile);
    for(HICON h:tempIcons) DestroyIcon(h);
    return true;
}



void ReaperThread(DWORD pid,std::wstring clientName,ProfileType type){
    HANDLE hProcess=OpenProcess(SYNCHRONIZE,FALSE,pid);
    if(hProcess){
        WaitForSingleObject(hProcess,INFINITE);
        CloseHandle(hProcess);
    }
    {
        std::lock_guard<std::mutex> lock(g_activeProfilesMutex);
        g_activeProfiles.erase(clientName);
    }
    fs::path profilePath;
    switch(type){
    case ProfileType::Temporary:
        profilePath=g_sDataDir/"Temp";
        try{ fs::remove_all(profilePath); } catch(...){}
        break;
    case ProfileType::Default:
        profilePath=g_sDataDir/"Default";
        CleanupProfile(profilePath,aKeepActive);
        SetUiState(false);
        if(MessageBox(g_hGui,L"Save changes to default profile?",APP_TITLE.c_str(),MB_YESNO|MB_ICONQUESTION|MB_APPLMODAL)==IDYES){
            fs::path backup7z=g_sDataDir/"Default.7z";
            fs::path bakDir=g_sDataDir/"_DefBak";
            fs::create_directories(bakDir);
            auto const now=std::chrono::system_clock::now();
            auto const in_time_t=std::chrono::system_clock::to_time_t(now);
            std::tm tm_buf;
            localtime_s(&tm_buf,&in_time_t);
            std::wostringstream ss;
            ss<<std::put_time(&tm_buf,L"%Y.%m.%d,%H%M%S");
            std::wstring timestamp=ss.str();
            std::wstring bakPath=(bakDir/(timestamp+L"-Default.7z")).wstring();
            try{ fs::rename(backup7z,bakPath); } catch(...){
                // Keep original behavior (best-effort), but try copy+delete if rename fails (cross-volume, AV locks, etc.).
                try{ fs::copy_file(backup7z,bakPath,fs::copy_options::overwrite_existing); } catch(...){}
                try{ fs::remove(backup7z); } catch(...){}
            }
            // Rebuild Default.7z from the cleaned Default profile folder (InProc7z).
            _7zUiCtx ctx;
            ctx.mainWnd=g_hGui;
            ctx.status=L"Updating Default Profile...";

            ProgressUI_PostShow(ctx.mainWnd,ctx.status,-1);

            HRESULT hr=_7zCompress7z(
                backup7z.c_str(),
                profilePath.c_str(),
                false,
                &_7zProgress,
                &ctx
            );

            ProgressUI_PostHide(ctx.mainWnd);

            if(FAILED(hr)){
                MessageBoxW(g_hGui,L"Failed to create Default.7z (InProc7z).",L"Error",MB_OK|MB_ICONERROR);
            }
        }
        SetUiState(true);
        try{ fs::remove_all(profilePath); } catch(...){}
        break;

    case ProfileType::Standard:
    default:
        profilePath=g_sDataDir/"Sites"/clientName;
        CleanupProfile(profilePath,aKeepActive);
        break;
    }
}

void EnsureWatcherIsRunning(){
    if(!g_isWatcherRunning.exchange(true)){
        g_watcherThread=std::jthread(WatcherThread);
    }
}

void WatcherThread(){
    while(true){
        std::this_thread::sleep_for(std::chrono::milliseconds(200));

        std::map<std::wstring,DWORD> profiles_copy;
        {
            std::lock_guard<std::mutex> lock(g_activeProfilesMutex);
            if(g_activeProfiles.empty()){
                g_isWatcherRunning=false;

                // free cache (all sizes)
                std::lock_guard<std::mutex> cacheLock(g_iconCacheMutex);
                for(auto& [name,set]:g_iconCache){
                    for(auto& [px,hIcon]:set.byPx){
                        if(hIcon) DestroyIcon(hIcon);
                    }
                }
                g_iconCache.clear();
                return;
            }
            profiles_copy=g_activeProfiles;
        }

        std::lock_guard<std::mutex> cacheLock(g_iconCacheMutex);

        // prune cache entries for profiles no longer active
        for(auto it=g_iconCache.begin(); it!=g_iconCache.end(); ){
            if(profiles_copy.find(it->first)==profiles_copy.end()){
                for(auto& [px,hIcon]:it->second.byPx){
                    if(hIcon) DestroyIcon(hIcon);
                }
                it=g_iconCache.erase(it);
            } else{
                ++it;
            }
        }

        for(const auto& [clientName,pid]:profiles_copy){
            fs::path profilePath;
            if(clientName==L"Temp") profilePath=g_sDataDir/"Temp";
            else if(clientName==L"Default") profilePath=g_sDataDir/"Default";
            else profilePath=g_sDataDir/"Sites"/clientName;

            fs::path iconPath=profilePath/"client.ico";

            auto& cacheSet=g_iconCache[clientName];

            auto getIconPx=[&](int px)->HICON{
                auto it=cacheSet.byPx.find(px);
                if(it!=cacheSet.byPx.end()) return it->second;

                HICON h=nullptr;
                if(fs::exists(iconPath)){
                    h=LoadIconFromIcoBestDownscale(iconPath,px); // next-highest then downscale
                }
                cacheSet.byPx[px]=h; // cache nullptr too
                return h;
            };

            EnumData data{pid};
            EnumWindows(EnumWindowsCallback,(LPARAM)&data);
            if(data.windows.empty()) continue;

            std::wstring newTitlePartial=std::format(L" - {} - {}",clientName,APP_ALIAS);

            for(HWND ew:data.windows){
                const UINT dpi=GetDpiForWindow(ew);
                const int pxSmall=GetSystemMetricsForDpi(SM_CXSMICON,dpi);
                const int pxBig=GetSystemMetricsForDpi(SM_CXICON,dpi);

                HICON hSmall=getIconPx(pxSmall);
                HICON hBig=getIconPx(pxBig);

                if(hSmall&&(HICON)SendMessage(ew,WM_GETICON,ICON_SMALL,0)!=hSmall){
                    SendMessage(ew,WM_SETICON,ICON_SMALL,(LPARAM)hSmall);
                }
                if(hBig&&(HICON)SendMessage(ew,WM_GETICON,ICON_BIG,0)!=hBig){
                    SendMessage(ew,WM_SETICON,ICON_BIG,(LPARAM)hBig);
                }

                wchar_t currentTitle[256];
                GetWindowTextW(ew,currentTitle,256);
                std::wstring titleStr(currentTitle);

                std::wregex expr(L"(.*) - Profile 1 - Microsoft.*Edge");
                std::wsmatch match;
                if(std::regex_match(titleStr,match,expr)&&match.size()>1){
                    std::wstring newTitle=match[1].str()+newTitlePartial;
                    SetWindowTextW(ew,newTitle.c_str());
                }

                std::wstring appId=std::format(L"ctSpaces.{}.Default",clientName);
                SetWindowAppId(ew,appId);
            }
        }
    }
}


bool FindEdgePath(){
    std::vector<fs::path> searchPaths;
    PWSTR pPath=NULL;
    if(SUCCEEDED(SHGetKnownFolderPath(FOLDERID_ProgramFilesX86,0,NULL,&pPath))){
        searchPaths.push_back(fs::path(pPath)/"Microsoft"/"Edge"/"Application"/"msedge.exe");
        CoTaskMemFree(pPath);
    }
    if(SUCCEEDED(SHGetKnownFolderPath(FOLDERID_ProgramFiles,0,NULL,&pPath))){
        searchPaths.push_back(fs::path(pPath)/"Microsoft"/"Edge"/"Application"/"msedge.exe");
        CoTaskMemFree(pPath);
    }
    if(SUCCEEDED(SHGetKnownFolderPath(FOLDERID_LocalAppData,0,NULL,&pPath))){
        searchPaths.push_back(fs::path(pPath)/"Microsoft"/"Edge"/"Application"/"msedge.exe");
        CoTaskMemFree(pPath);
    }
    for(const auto& path:searchPaths){
        if(fs::exists(path)){
            g_sEdgePath=path;
            return true;
        }
    }
    return false;
}

void TerminateAllProfiles(){
    std::lock_guard<std::mutex> lock(g_activeProfilesMutex);
    for(const auto& [clientName,pid]:g_activeProfiles){
        HANDLE hProcess=OpenProcess(PROCESS_TERMINATE,FALSE,pid);
        if(hProcess){
            TerminateProcess(hProcess,0);
            CloseHandle(hProcess);
        }
    }
    g_activeProfiles.clear();
}
bool RunUpdateCheck();
void RunInstall();
std::wstring GetExeVersion(const fs::path& filePath);
bool CreateShortcut(const fs::path& targetPath,const fs::path& shortcutPath,const fs::path& workingDir,const fs::path& iconPath);

std::wstring GetExeVersion(const fs::path& filePath){
    DWORD handle=0;
    DWORD versionSize=GetFileVersionInfoSizeW(filePath.c_str(),&handle);
    if(versionSize==0) return L"";
    auto versionData=std::make_unique<BYTE[]>(versionSize);
    if(!GetFileVersionInfoW(filePath.c_str(),0,versionSize,versionData.get())) return L"";
    VS_FIXEDFILEINFO* fileInfo=nullptr;
    UINT fileInfoSize=0;
    if(VerQueryValueW(versionData.get(),L"\\",(LPVOID*)&fileInfo,&fileInfoSize)&&fileInfo){
        return std::format(L"{}.{}.{}.{}",
                           HIWORD(fileInfo->dwFileVersionMS),LOWORD(fileInfo->dwFileVersionMS),
                           HIWORD(fileInfo->dwFileVersionLS),LOWORD(fileInfo->dwFileVersionLS)
        );
    }
    return L"";
}


bool CreateShortcut(const fs::path& targetPath,const fs::path& shortcutPath,const fs::path& workingDir,const fs::path& iconPath){
    IShellLink* pShellLink=NULL;
    HRESULT hr=CoCreateInstance(CLSID_ShellLink,NULL,CLSCTX_INPROC_SERVER,IID_IShellLink,(LPVOID*)&pShellLink);
    if(SUCCEEDED(hr)){
        pShellLink->SetPath(targetPath.c_str());
        pShellLink->SetWorkingDirectory(workingDir.c_str());
        pShellLink->SetIconLocation(iconPath.c_str(),0);

        IPersistFile* pPersistFile;
        hr=pShellLink->QueryInterface(IID_IPersistFile,(LPVOID*)&pPersistFile);
        if(SUCCEEDED(hr)){
            hr=pPersistFile->Save(shortcutPath.c_str(),TRUE);
            pPersistFile->Release();
        }
        pShellLink->Release();
    }
    return SUCCEEDED(hr);
}

bool doInstall(){
    if(MessageBox(NULL,L"ctSpaces is not installed. Would you like to install it now?",L"Install ctSpaces",MB_YESNO|MB_ICONQUESTION)==IDYES){
        fs::path installedExePath=g_sDataDir/L"ctSpaces.exe";
        wchar_t currentExePath[MAX_PATH];
        GetModuleFileNameW(NULL,currentExePath,MAX_PATH);
        try{
            fs::copy_file(currentExePath,installedExePath,fs::copy_options::overwrite_existing);
            if(MessageBox(NULL,L"Add ctSpaces to the Start Menu?",L"Installation",MB_YESNO|MB_ICONQUESTION)==IDYES){
                PWSTR pProgramsPath=NULL;
                if(SUCCEEDED(SHGetKnownFolderPath(FOLDERID_Programs,0,NULL,&pProgramsPath))){
                    fs::path shortcutPath=fs::path(pProgramsPath)/L"ctSpaces.lnk";
                    CreateShortcut(installedExePath,shortcutPath,g_sDataDir,installedExePath);
                    CoTaskMemFree(pProgramsPath);
                }
            }
            if(MessageBox(NULL,L"Add ctSpaces to your Desktop?",L"Installation",MB_YESNO|MB_ICONQUESTION)==IDYES){
                PWSTR pDesktopPath=NULL;
                if(SUCCEEDED(SHGetKnownFolderPath(FOLDERID_Desktop,0,NULL,&pDesktopPath))){
                    fs::path shortcutPath=fs::path(pDesktopPath)/L"ctSpaces.lnk";
                    CreateShortcut(installedExePath,shortcutPath,g_sDataDir,installedExePath);
                    CoTaskMemFree(pDesktopPath);
                }
            }
            if(MessageBox(NULL,L"Run ctSpaces on Windows startup?",L"Installation",MB_YESNO|MB_ICONQUESTION)==IDYES){
                HKEY hKey=NULL;
                if(RegOpenKeyExW(HKEY_CURRENT_USER,L"Software\\Microsoft\\Windows\\CurrentVersion\\Run",0,KEY_SET_VALUE,&hKey)==ERROR_SUCCESS){
                    std::wstring pathStr=installedExePath.wstring();
                    RegSetValueExW(hKey,L"ctSpaces",0,REG_SZ,(const BYTE*)pathStr.c_str(),(DWORD)((pathStr.length()+1)*sizeof(wchar_t)));
                    RegCloseKey(hKey);
                }
            }
            if(MessageBox(NULL,L"Installation complete. Would you like to run the installed version now?",L"Installation",MB_YESNO|MB_ICONQUESTION)==IDYES){
                ShellExecuteW(NULL,L"open",installedExePath.c_str(),NULL,g_sDataDir.c_str(),SW_SHOW);
            }
        } catch(...){
            MessageBox(NULL,L"Installation failed. Could not copy file.",L"Error",MB_OK|MB_ICONERROR);
        }
        return false;
    } else{
        if(MessageBox(NULL,L"Would you like to run this version temporarily without installing?",L"Run Temporarily",MB_YESNO|MB_ICONQUESTION)==IDYES){
            return true;
        } else{
            return false;
        }
    }
}

bool chkUpdate(){
    wchar_t currentExePathStr[MAX_PATH];
    GetModuleFileNameW(NULL,currentExePathStr,MAX_PATH);
    fs::path currentExePath(currentExePathStr);
    fs::path installedExePath=g_sDataDir/L"ctSpaces.exe";
    if(fs::equivalent(currentExePath,installedExePath)){
        return true;
    }
    std::wstring currentVersion=GetExeVersion(currentExePath);
    std::wstring installedVersion=GetExeVersion(installedExePath);
    if(currentVersion>installedVersion){
        std::wstring prompt=std::format(L"A new version of ctSpaces is available ({} -> {}).\n\nWould you like to update now?",installedVersion,currentVersion);
        if(MessageBox(NULL,prompt.c_str(),L"Update Available",MB_YESNO|MB_ICONQUESTION)==IDYES){
            try{
                fs::copy_file(currentExePath,installedExePath,fs::copy_options::overwrite_existing);
                if(MessageBox(NULL,L"Update complete! Would you like to run the updated version now?",L"Update Success",MB_OK|MB_ICONINFORMATION)){
                    return true;
                } else{
                    return false;
                }
            } catch(const fs::filesystem_error&){
                MessageBox(NULL,L"Update failed. The installed version of ctSpaces may be running. Please close it and try again.",L"Update Error",MB_OK|MB_ICONERROR);
            }
            return false;
        } else{
            if(MessageBox(NULL,L"Would you like to run this newer version temporarily without updating?",L"Run Temporarily",MB_YESNO|MB_ICONQUESTION)==IDYES){
                return true;
            } else{
                return false;
            }
        }
    } else{
        if(MessageBox(NULL,L"You are running a version that is not installed. Run this version temporarily?",L"Run Temporarily",MB_YESNO|MB_ICONQUESTION)==IDYES){
            return true;
        } else{
            return false;
        }
    }
}

// Ref: https://learn.microsoft.com/en-us/windows/win32/controls/create-a-tooltip-for-a-control
HWND CreateToolTip(HWND hwndTool,HWND hDlg,PTSTR pszText){
    if(!hwndTool||!hDlg||!pszText)
        return FALSE;
    HWND hwndTip=CreateWindowEx(NULL,TOOLTIPS_CLASS,NULL,WS_POPUP|TTS_ALWAYSTIP|TTS_BALLOON,CW_USEDEFAULT,CW_USEDEFAULT,CW_USEDEFAULT,CW_USEDEFAULT,hDlg,NULL,g_hInst,NULL);
    if(!hwndTool||!hwndTip)
        return (HWND)NULL;
    TOOLINFO toolInfo={0};
    toolInfo.cbSize=sizeof(toolInfo);
    toolInfo.hwnd=hDlg;
    toolInfo.uFlags=TTF_IDISHWND|TTF_SUBCLASS;
    toolInfo.uId=(UINT_PTR)hwndTool;
    toolInfo.lpszText=pszText;
    SendMessageW(hwndTip,WM_SETFONT,(WPARAM)g_hFont,TRUE);
    SendMessage(hwndTip,TTM_ADDTOOL,0,(LPARAM)&toolInfo);
    return hwndTip;
}

static void PostTaskComplete(const std::wstring& name){
    auto* pName=new std::wstring(name);
    PostMessage(g_hGui,WM_APP_TASK_COMPLETE,0,(LPARAM)pName);
}

static void LaunchProfileAsync(const std::wstring& name,bool isTemp,bool isDefault){
    SetUiState(false);
    std::thread([name,isTemp,isDefault](){
        DWORD pid=LaunchProfile(name,isTemp,isDefault);
        if(pid>0){
            {
                std::lock_guard<std::mutex> lock(g_activeProfilesMutex);
                g_activeProfiles[name]=pid;
            }
            g_reaperThreads.emplace_back(ReaperThread,pid,name,
                                         isTemp?ProfileType::Temporary:
                                         (isDefault?ProfileType::Default:ProfileType::Standard));
            EnsureWatcherIsRunning();
        }
        PostTaskComplete(name);
    }).detach();
}

void GuiProfDel(){
    SetUiState(false);
    wchar_t clientNameBuffer[256];
    GetWindowText(g_hComboClient,clientNameBuffer,256);
    std::wstring clientName=SanitizeName(clientNameBuffer);

    if(clientName.empty()){
        MessageBox(g_hGui,L"Please select a client first.",L"Warning",MB_OK|MB_ICONWARNING);
        SetUiState(true);
        return;
    }

    std::lock_guard<std::mutex> lock(g_activeProfilesMutex);
    if(g_activeProfiles.count(clientName)){
        MessageBox(g_hGui,L"Cannot delete a profile that is currently active.",L"Action Denied",MB_OK|MB_ICONWARNING);
        SetUiState(true);
        return;
    }

    if(MessageBox(g_hGui,
                  L"This will permanently delete the selected profile folder.\n\nContinue?",
                  L"Confirm Delete",
                  MB_YESNO|MB_ICONQUESTION)!=IDYES){
        SetUiState(true);
        return;
    }

    bool hadError=false;
    std::wstring errMsg;

    fs::path sData=g_sDataDir/L"Sites"/clientName;
    try{
        if(fs::exists(sData)){
            fs::remove_all(sData);
        }
    } catch(const fs::filesystem_error& e){
        hadError=true;
        errMsg=AnsiToWide(e.what());
    } catch(...){
        hadError=true;
        errMsg=L"Unknown error.";
    }
    UpdateClientsComboBox();
    SetWindowText(g_hComboClient,L"");
    if(hadError){
        std::wstring fullMsg=L"Failed to delete profile.";
        if(!errMsg.empty()){
            fullMsg+=L"\n\nDetails: "+errMsg;
        }
        MessageBox(g_hGui,fullMsg.c_str(),L"Delete Profile",MB_OK|MB_ICONERROR);
    } else{
        MessageBox(g_hGui,L"Profile deleted successfully.",L"Delete Profile",MB_OK|MB_ICONINFORMATION);
    }
    SetUiState(true);
}

inline void EnsureMouseVisible(){
    CURSORINFO ci{sizeof(ci)};
    if(GetCursorInfo(&ci)&&ci.flags==0){
        ShowCursor(TRUE);
    }
}

inline void FocusClientEdit(){
    if(g_hComboClient&&IsWindowEnabled(g_hComboClient)){
        SetFocus(g_hComboClient);
        SendMessage(g_hComboClient,CB_SETEDITSEL,0,MAKELPARAM(0,-1));
    }
}

static INT_PTR CALLBACK AboutDlgProc(HWND hDlg,UINT msg,WPARAM wParam,LPARAM lParam){
    constexpr int kIdIcon=1001;
    constexpr int kIdTitle=1002;
    constexpr int kIdBy=1003;
    constexpr int kIdLink=1004;
    constexpr int kIdThanks=1005;

    auto ensureStatic=[&](int id,const wchar_t* text,DWORD style,int x,int y,int width,int height)->HWND{
        HWND hwnd=GetDlgItem(hDlg,id);
        if(!hwnd){
            hwnd=CreateWindowW(L"STATIC",text,WS_CHILD|WS_VISIBLE|style,
                               x,y,width,height,
                               hDlg,(HMENU)id,g_hInst,nullptr);
        } else{
            SetWindowTextW(hwnd,text);
            MoveWindow(hwnd,x,y,width,height,TRUE);
        }
        return hwnd;
    };

    auto ensureLink=[&](int id,const wchar_t* text,int x,int y,int width,int height)->HWND{
        HWND hwnd=GetDlgItem(hDlg,id);
        if(!hwnd){
            hwnd=CreateWindowW(WC_LINK,text,WS_CHILD|WS_VISIBLE|WS_TABSTOP,
                               x,y,width,height,
                               hDlg,(HMENU)id,g_hInst,nullptr);
        } else{
            SetWindowTextW(hwnd,text);
            MoveWindow(hwnd,x,y,width,height,TRUE);
        }
        return hwnd;
    };

    auto relayout=[&](UINT dpi,const RECT* suggested){
        ResetAboutFonts(dpi);

        const int margin=MulDiv(12,dpi,96);
        const int iconPx=MulDiv(64,dpi,96);

        const int dlgW=MulDiv(480,dpi,96);
        const int dlgH=MulDiv(192,dpi,96);

        RECT wnd{};
        if(suggested){
            wnd=*suggested;
        } else{
            GetWindowRect(hDlg,&wnd);
        }
        SetWindowPos(hDlg,nullptr,wnd.left,wnd.top,dlgW,dlgH,SWP_NOZORDER|SWP_NOACTIVATE);

        RECT rc{};
        GetClientRect(hDlg,&rc);

        const int xIco=margin;
        const int yTop=margin;

        HWND hIco=GetDlgItem(hDlg,kIdIcon);
        if(!hIco){
            hIco=CreateWindowW(L"STATIC",L"",WS_CHILD|WS_VISIBLE|SS_ICON,
                               xIco,yTop,iconPx,iconPx,
                               hDlg,(HMENU)kIdIcon,g_hInst,nullptr);
        } else{
            MoveWindow(hIco,xIco,yTop,iconPx,iconPx,TRUE);
        }

        const int xText=xIco+iconPx+margin;
        const int wText=rc.right-xText-margin;

        HWND hTitle=ensureStatic(kIdTitle,APP_TITLE.c_str(),SS_LEFT|SS_NOPREFIX,
                                 xText,yTop,wText,MulDiv(24,dpi,96));

        HWND hBy=ensureStatic(kIdBy,L"by BiatuAutMiahn",SS_LEFT|SS_NOPREFIX,
                              xText+MulDiv(16,dpi,96),yTop+MulDiv(20,dpi,96),wText,MulDiv(20,dpi,96));

        const wchar_t* linkText=
            L"<a href=\"https://github.com/BiatuAutMiahn/ctSpaces\">https://github.com/BiatuAutMiahn/ctSpaces</a>";

        HWND hLink=ensureLink(kIdLink,linkText,
                              xText,yTop+MulDiv(36,dpi,96),rc.right-margin*2,MulDiv(24,dpi,96));

        const wchar_t* thanksText=
            L"Thanks:\r\n"
            L"  Igor Pavlov (7-Zip)\r\n"
            L"  OpenAI (R&D and rapid prototyping)\r\n"
            L"  Google (Material Icons)";

        HWND hThanks=ensureStatic(kIdThanks,thanksText,SS_LEFT|SS_NOPREFIX,
                                  xText,yTop+MulDiv(68,dpi,96),MulDiv(256,dpi,96),MulDiv(80,dpi,96));

        HWND hOk=GetDlgItem(hDlg,IDOK);
        if(hOk){
            RECT br{};
            GetWindowRect(hOk,&br);
            const int bw=br.right-br.left;
            const int bh=br.bottom-br.top;
            MoveWindow(hOk,rc.right-margin-bw,rc.bottom-margin-bh,bw,bh,TRUE);
        }

        HFONT hFont=g_hAboutFont?g_hAboutFont:g_hFont;
        if(hOk)     SendMessageW(hOk,WM_SETFONT,(WPARAM)hFont,TRUE);
        if(hThanks) SendMessageW(hThanks,WM_SETFONT,(WPARAM)hFont,TRUE);
        if(hTitle)  SendMessageW(hTitle,WM_SETFONT,(WPARAM)hFont,TRUE);
        if(hBy)     SendMessageW(hBy,WM_SETFONT,(WPARAM)GetAboutSmallFont(),TRUE);
        if(hLink)   SendMessageW(hLink,WM_SETFONT,(WPARAM)hFont,TRUE);

        const int sm=GetSystemMetricsForDpi(SM_CXSMICON,dpi);
        const int bg=GetSystemMetricsForDpi(SM_CXICON,dpi);

        if(g_hAboutDlgSmall){ DestroyIcon(g_hAboutDlgSmall); g_hAboutDlgSmall=nullptr; }
        if(g_hAboutDlgBig){ DestroyIcon(g_hAboutDlgBig); g_hAboutDlgBig=nullptr; }
        if(g_hAboutIcon64){ DestroyIcon(g_hAboutIcon64); g_hAboutIcon64=nullptr; }

        g_hAboutDlgSmall=LoadIconResBestDownscale(g_hInst,IDI_CTSPACES,sm,sm);
        g_hAboutDlgBig=LoadIconResBestDownscale(g_hInst,IDI_CTSPACES,bg,bg);
        if(g_hAboutDlgSmall) SendMessageW(hDlg,WM_SETICON,ICON_SMALL,(LPARAM)g_hAboutDlgSmall);
        if(g_hAboutDlgBig)   SendMessageW(hDlg,WM_SETICON,ICON_BIG,(LPARAM)g_hAboutDlgBig);

        g_hAboutIcon64=LoadIconResBestDownscale(g_hInst,IDI_IRND,iconPx,iconPx);
        if(hIco&&g_hAboutIcon64){
            SendMessageW(hIco,STM_SETIMAGE,IMAGE_ICON,(LPARAM)g_hAboutIcon64);
        }
    };

    switch(msg){
    case WM_INITDIALOG: {
        const UINT dpi=GetDpiForWindow(hDlg);

        // Title
        std::wstring title=std::format(L"About");
        SetWindowTextW(hDlg,title.c_str());

        // Hide ALL existing resource children except OK/CANCEL
        EnumChildWindows(hDlg,[](HWND c,LPARAM)->BOOL{
            int id=GetDlgCtrlID(c);
            if(id!=IDOK&&id!=IDCANCEL) ShowWindow(c,SW_HIDE);
            return TRUE;
        },0);

        relayout(dpi,nullptr);

        return (INT_PTR)TRUE;
    }

    case WM_DPICHANGED: {
        const UINT dpi=HIWORD(wParam);
        const RECT* rc=reinterpret_cast<RECT*>(lParam);
        relayout(dpi,rc);
        return (INT_PTR)TRUE;
    }

    case WM_NOTIFY: {
        NMHDR* nm=(NMHDR*)lParam;
        if(!nm) break;

        if((nm->code==NM_CLICK||nm->code==NM_RETURN)&&nm->idFrom==kIdLink){
            auto* link=(NMLINK*)lParam;
            if(link) ShellExecuteW(hDlg,L"open",link->item.szUrl,nullptr,nullptr,SW_SHOWNORMAL);
            return (INT_PTR)TRUE;
        }
        break;
    }

    case WM_COMMAND:
        if(LOWORD(wParam)==IDOK||LOWORD(wParam)==IDCANCEL){
            EndDialog(hDlg,LOWORD(wParam));
            return (INT_PTR)TRUE;
        }
        break;

    case WM_DESTROY:
        if(g_hAboutFont){ DeleteObject(g_hAboutFont); g_hAboutFont=nullptr; }
        if(g_hAboutIcon64){ DestroyIcon(g_hAboutIcon64);    g_hAboutIcon64=nullptr; }
        if(g_hAboutDlgSmall){ DestroyIcon(g_hAboutDlgSmall);  g_hAboutDlgSmall=nullptr; }
        if(g_hAboutDlgBig){ DestroyIcon(g_hAboutDlgBig);    g_hAboutDlgBig=nullptr; }
        if(g_hFontAboutSmall){ DeleteObject(g_hFontAboutSmall); g_hFontAboutSmall=nullptr; }
        break;
    }
    return (INT_PTR)FALSE;
}


static void ShowAboutDialog(){
    // Ensure the dialog is created under the same DPI awareness context as the main window.
    DPI_AWARENESS_CONTEXT parentCtx=GetWindowDpiAwarenessContext(g_hGui);
    DPI_AWARENESS_CONTEXT oldCtx=SetThreadDpiAwarenessContext(parentCtx);

    DialogBoxW(g_hInst,MAKEINTRESOURCEW(IDD_ABOUTBOX),g_hGui,AboutDlgProc);

    SetThreadDpiAwarenessContext(oldCtx);
    //DialogBoxW(g_hInst,MAKEINTRESOURCEW(IDD_ABOUTBOX),g_hGui,AboutDlgProc);
}
static HBITMAP IconToMenuBitmap(HICON hIcon,int cx,int cy){
    if(!hIcon) return nullptr;

    BITMAPINFO bmi{};
    bmi.bmiHeader.biSize=sizeof(BITMAPINFOHEADER);
    bmi.bmiHeader.biWidth=cx;
    bmi.bmiHeader.biHeight=-cy; // top-down DIB
    bmi.bmiHeader.biPlanes=1;
    bmi.bmiHeader.biBitCount=32;
    bmi.bmiHeader.biCompression=BI_RGB;

    void* bits=nullptr;
    HDC hdc=GetDC(nullptr);
    HBITMAP hbmp=CreateDIBSection(hdc,&bmi,DIB_RGB_COLORS,&bits,nullptr,0);
    ReleaseDC(nullptr,hdc);
    if(!hbmp||!bits) return nullptr;

    ZeroMemory(bits,(size_t)cx*(size_t)cy*4);

    HDC memDC=CreateCompatibleDC(nullptr);
    HGDIOBJ old=SelectObject(memDC,hbmp);
    DrawIconEx(memDC,0,0,hIcon,cx,cy,0,nullptr,DI_NORMAL);
    SelectObject(memDC,old);
    DeleteDC(memDC);

    return hbmp;
}

static HICON LoadIconResSized(int iconResId,int cx,int cy){
    return (HICON)LoadImageW(g_hInst,MAKEINTRESOURCEW(iconResId),IMAGE_ICON,cx,cy,LR_DEFAULTCOLOR);
}

static HICON LoadIconResScaled(int iconResId,int cx,int cy){
    HICON h=nullptr;
    if(SUCCEEDED(LoadIconWithScaleDown(g_hInst,MAKEINTRESOURCEW(iconResId),cx,cy,&h))&&h)
        return h;
    return (HICON)LoadImageW(
        g_hInst,
        MAKEINTRESOURCEW(iconResId),
        IMAGE_ICON,
        cx,cy,
        LR_DEFAULTCOLOR
    );
}


static void MenuAddItemBmp(HMENU hMenu,UINT id,const wchar_t* text,HBITMAP hbmp){
    MENUITEMINFOW mi{};
    mi.cbSize=sizeof(mi);
    mi.fMask=MIIM_ID|MIIM_STRING|MIIM_BITMAP;
    mi.wID=id;
    mi.dwTypeData=const_cast<wchar_t*>(text);
    mi.hbmpItem=hbmp;
    InsertMenuItemW(hMenu,(UINT)-1,TRUE,&mi);
}

static void MenuAddSep(HMENU hMenu){
    MENUITEMINFOW mi{};
    mi.cbSize=sizeof(mi);
    mi.fMask=MIIM_FTYPE;
    mi.fType=MFT_SEPARATOR;
    InsertMenuItemW(hMenu,(UINT)-1,TRUE,&mi);
}
static void EnsureConfigMenu(HWND hWnd){
    if(g_hConfigMenu) return;

    g_hConfigMenu=CreatePopupMenu();

    const UINT dpi=GetDpiForWindow(hWnd);

    // pick a logical size you like for menu icons
    const int iconPx=MulDiv(28,dpi,96);      // <- change 18 to 16/20/24 etc.
    //const int iconCx=GetSystemMetricsForDpi(SM_CXMENUCHECK,dpi);
    //const int iconCy=GetSystemMetricsForDpi(SM_CYMENUCHECK,dpi);

    const int rowPad=MulDiv(2,dpi,96);

    MENUINFO mnuInfo{};
    mnuInfo.cbSize=sizeof(mnuInfo);
    mnuInfo.fMask=MIM_STYLE;
    mnuInfo.dwStyle=MNS_CHECKORBMP;
    SetMenuInfo(g_hConfigMenu,&mnuInfo);

    auto mkbmp=[&](int iconResId) -> HBITMAP{
        HBITMAP b=LoadMenuBitmapFromIconRes(iconResId,dpi,iconPx,iconPx);
        g_menuBitmaps.push_back(b);
        return b;
    };

    MenuAddItemBmp(g_hConfigMenu,IDM_CTX_SET_PROFILE_ICON,L"Set Profile Icon",mkbmp(IDI_AICOB));
    MenuAddItemBmp(g_hConfigMenu,IDM_CTX_REFRESH_PROFILE,L"Refresh Profile",mkbmp(IDI_UPRFB));
    MenuAddItemBmp(g_hConfigMenu,IDM_CTX_RESET_PROFILE,L"Reset Profile",mkbmp(IDI_RPRFB));
    MenuAddItemBmp(g_hConfigMenu,IDM_CTX_DELETE_PROFILE,L"Delete Profile",mkbmp(IDI_DPRFB));
    MenuAddSep(g_hConfigMenu);
    MenuAddItemBmp(g_hConfigMenu,IDM_CTX_EDIT_DEFAULT_PROFILE,L"Edit Default profile",mkbmp(IDI_EDPFB));
    MenuAddSep(g_hConfigMenu);
    MenuAddItemBmp(g_hConfigMenu,IDM_ABOUT,L"About",mkbmp(IDI_INFOB));
}


static UINT ShowConfigMenuFromButton(HWND hWnd){
    EnsureConfigMenu(hWnd);
    UpdateConfigMenuEnabledState();
    HideMenuTooltip();            
    RECT rc{};
    GetWindowRect(g_hBtnConfig,&rc);

    SetForegroundWindow(hWnd);

    UINT cmd=TrackPopupMenuEx(
        g_hConfigMenu,
        TPM_LEFTALIGN|TPM_TOPALIGN|TPM_RETURNCMD,
        rc.left,
        rc.bottom,
        hWnd,
        nullptr
    );
    HideMenuTooltip();
    PostMessageW(hWnd,WM_NULL,0,0); // allow menu to dismiss cleanly
    return cmd;
}

static void SetButtonIcon(HWND hBtn,int iconResId,HICON& hStore){
    if(!hBtn) return;

    RECT rc{};
    GetClientRect(hBtn,&rc);

    const int w=rc.right-rc.left;
    const int h=rc.bottom-rc.top;

    const UINT dpi=GetDpiForWindow(hBtn);
    const int pad=MulDiv(2,dpi,96);
    const int s=max(16,min(w,h)-pad); // requested size

    HICON hNew=LoadIconResBestDownscale(g_hInst,iconResId,s,s);
    
    if(!hNew) return;

    if(hStore){
        DestroyIcon(hStore);
        hStore=nullptr;
    }
    hStore=hNew;

    SendMessageW(hBtn,BM_SETIMAGE,IMAGE_ICON,(LPARAM)hStore);
}


static HBITMAP LoadMenuBitmapFromIconRes(int iconResId,UINT /*dpi*/,int cx,int cy){
    HICON hIcon=LoadIconResBestDownscale(g_hInst,iconResId,cx,cy);
    HBITMAP hbmp=IconToMenuBitmap(hIcon,cx,cy);
    if(hIcon) DestroyIcon(hIcon);
    return hbmp;
}

static bool GetIconSizePx(HICON hIcon,int& w,int& h){
    w=h=0;
    ICONINFO ii{};
    if(!GetIconInfo(hIcon,&ii)) return false;

    BITMAP bm{};
    bool ok=false;
    if(ii.hbmColor&&GetObject(ii.hbmColor,sizeof(bm),&bm)==sizeof(bm)){
        w=bm.bmWidth; h=bm.bmHeight; ok=true;
    } else if(ii.hbmMask&&GetObject(ii.hbmMask,sizeof(bm),&bm)==sizeof(bm)){
        w=bm.bmWidth; h=bm.bmHeight/2; ok=true;
    }

    if(ii.hbmColor) DeleteObject(ii.hbmColor);
    if(ii.hbmMask)  DeleteObject(ii.hbmMask);
    return ok;
}

static HICON LoadIconResBestDownscale(HINSTANCE hInst,int groupIconResId,int cxDesired,int cyDesired){
    const int want=max(cxDesired,cyDesired);
    if(want<=0) return nullptr;

    HRSRC hGrpRes=FindResourceW(hInst,MAKEINTRESOURCEW(groupIconResId),RT_GROUP_ICON);
    if(!hGrpRes) return nullptr;

    HGLOBAL hGrp=LoadResource(hInst,hGrpRes);
    if(!hGrp) return nullptr;

    const auto* grp=(const GRPICONDIR*)LockResource(hGrp);
    if(!grp||grp->idReserved!=0||grp->idType!=1||grp->idCount==0) return nullptr;

    const GRPICONDIRENTRY* bestGE=nullptr;
    int bestGES=(std::numeric_limits<int>::max)();

    const GRPICONDIRENTRY* bestBig=nullptr;
    int bestBigS=0;

    for(int i=0;i<(int)grp->idCount;i++){
        const auto& e=grp->idEntries[i];
        const int s=max(IcoDim(e.bWidth),IcoDim(e.bHeight));

        if(s>=want&&s<bestGES){
            bestGES=s;
            bestGE=&e;
        }
        if(s>bestBigS){
            bestBigS=s;
            bestBig=&e;
        }
    }

    const GRPICONDIRENTRY* pick=bestGE?bestGE:bestBig;
    if(!pick) return nullptr;

    const int src=max(IcoDim(pick->bWidth),IcoDim(pick->bHeight));

    HRSRC hIconRes=FindResourceW(hInst,MAKEINTRESOURCEW(pick->nID),RT_ICON);
    if(!hIconRes) return nullptr;

    DWORD cb=SizeofResource(hInst,hIconRes);
    HGLOBAL hBits=LoadResource(hInst,hIconRes);
    if(!hBits||cb==0) return nullptr;

    BYTE* pBits=(BYTE*)LockResource(hBits);
    if(!pBits) return nullptr;

    // IMPORTANT: create at native size (0,0) so we do NOT let this API resample.
    HICON hNative=CreateIconFromResourceEx(
        pBits,cb,TRUE,0x00030000,
        0,0,LR_DEFAULTCOLOR
    );
    if(!hNative) return nullptr;

    // If exact match (or smaller-only), return as-is (no upscale).
    if(src<=want) return hNative;

    // Downscale only.
    HICON hScaled=(HICON)CopyImage(hNative,IMAGE_ICON,want,want,0);
    if(!hScaled){
        hScaled=ScaleIconDown_HQ(hNative,want,want); // your GDI+ fallback
    }
    DestroyIcon(hNative);
    return hScaled;
}

static HICON LoadIconFromIcoBestDownscale(const fs::path& icoPath,int pxDesired){
    if(pxDesired<=0) return nullptr;

    std::ifstream f(icoPath,std::ios::binary|std::ios::ate);
    if(!f) return nullptr;

    const std::streamsize sz=f.tellg();
    if(sz<(std::streamsize)sizeof(ICONDIRHDR)) return nullptr;
    f.seekg(0,std::ios::beg);

    std::vector<BYTE> buf((size_t)sz);
    if(!f.read((char*)buf.data(),sz)) return nullptr;

    const auto* hdr=(const ICONDIRHDR*)buf.data();
    if(hdr->idReserved!=0||hdr->idType!=1||hdr->idCount==0) return nullptr;

    const size_t need=sizeof(ICONDIRHDR)+(size_t)hdr->idCount*sizeof(ICONDIRENTRY);
    if(buf.size()<need) return nullptr;

    const auto* ents=(const ICONDIRENTRY*)(buf.data()+sizeof(ICONDIRHDR));

    const int want=pxDesired;

    const ICONDIRENTRY* bestGE=nullptr;
    int bestGES=(std::numeric_limits<int>::max)();

    const ICONDIRENTRY* bestBig=nullptr;
    int bestBigS=0;

    for(int i=0;i<(int)hdr->idCount;i++){
        const auto& e=ents[i];
        int s=max(IcoDim(e.bWidth),IcoDim(e.bHeight));

        if(s>=want&&s<bestGES){
            bestGES=s;
            bestGE=&e;
        }
        if(s>bestBigS){
            bestBigS=s;
            bestBig=&e;
        }
    }

    const ICONDIRENTRY* pick=bestGE?bestGE:bestBig;
    if(!pick) return nullptr;

    const size_t off=(size_t)pick->dwImageOffset;
    const size_t cb=(size_t)pick->dwBytesInRes;
    if(off>=buf.size()||cb==0||off+cb>buf.size()) return nullptr;

    const int src=max(IcoDim(pick->bWidth),IcoDim(pick->bHeight));
    const int out=min(pxDesired,src); // never upscale

    BYTE* pBits=buf.data()+off;

    return CreateIconFromResourceEx(
        pBits,
        (DWORD)cb,
        TRUE,
        0x00030000,
        out,out,
        LR_DEFAULTCOLOR
    );
}


static HFONT CreateSmallerFontFrom(HFONT baseFont,int pxHeight,UINT dpi){
    LOGFONTW lf{};
    if(baseFont&&GetObjectW(baseFont,sizeof(lf),&lf)==sizeof(lf)){
        lf.lfHeight=-MulDiv(pxHeight,dpi,96); // negative = character height
        return CreateFontIndirectW(&lf);
    }
    // fallback
    return CreateFontW(
        -MulDiv(pxHeight,dpi,96),0,0,0,FW_NORMAL,FALSE,FALSE,FALSE,
        DEFAULT_CHARSET,OUT_DEFAULT_PRECIS,CLIP_DEFAULT_PRECIS,DEFAULT_QUALITY,
        DEFAULT_PITCH|FF_MODERN,L"Consolas"
    );
}


static HICON ScaleIconDown_HQ(HICON hSrc,int dstCx,int dstCy){
    if(!hSrc||dstCx<=0||dstCy<=0) return nullptr;

    Gdiplus::Bitmap src(hSrc);
    if(src.GetLastStatus()!=Gdiplus::Ok) return nullptr;

    const UINT srcW=src.GetWidth(),srcH=src.GetHeight();
    if((UINT)dstCx>=srcW&&(UINT)dstCy>=srcH) return CopyIcon(hSrc);

    Gdiplus::Bitmap dst(dstCx,dstCy,PixelFormat32bppARGB);
    if(dst.GetLastStatus()!=Gdiplus::Ok) return nullptr;

    Gdiplus::Graphics g(&dst);
    g.SetCompositingMode(Gdiplus::CompositingModeSourceCopy);
    g.SetCompositingQuality(Gdiplus::CompositingQualityHighQuality);
    g.SetInterpolationMode(Gdiplus::InterpolationModeHighQualityBicubic);
    g.SetSmoothingMode(Gdiplus::SmoothingModeHighQuality);
    g.SetPixelOffsetMode(Gdiplus::PixelOffsetModeHalf);

    g.Clear(Gdiplus::Color(0,0,0,0));
    g.DrawImage(&src,Gdiplus::Rect(0,0,dstCx,dstCy));

    HICON hOut=nullptr;
    if(dst.GetHICON(&hOut)!=Gdiplus::Ok) return nullptr;
    return hOut;
}

static fs::path GetProfileDirFromName(const std::wstring& name){
    if(_wcsicmp(name.c_str(),L"Temp")==0)    return g_sDataDir/"Temp";
    if(_wcsicmp(name.c_str(),L"Default")==0) return g_sDataDir/"Default";
    return g_sDataDir/"Sites"/name;
}

static void UpdateIconPreviewForSelection(bool preferListSelection){
    if(!g_hIconPreview) return;

    const UINT dpi=GetDpiForWindow(g_hIconPreview);
    const int px=MulDiv(32,dpi,96);

    std::wstring name=GetSelectedClientNameSanitized(preferListSelection);


    HICON hNew=nullptr;
    if(!name.empty()){
        fs::path icoPath=GetProfileDirFromName(name)/"client.ico";
        if(fs::exists(icoPath)){
            hNew=LoadIconFromIcoBestDownscale(icoPath,px);
        }
    }
    if(!hNew){
        hNew=LoadIconResBestDownscale(g_hInst,IDI_NIMGB,px,px);
    }

    HICON hOld=(HICON)SendMessageW(g_hIconPreview,STM_SETIMAGE,IMAGE_ICON,(LPARAM)hNew);
    if(hOld&&hOld!=hNew) DestroyIcon(hOld);

    g_hIconPreviewHandle=hNew;
}


static void EnsureMenuTooltips(HWND hWnd){
    if(g_hMenuTip) return;

    g_hMenuTip=CreateWindowExW(
        WS_EX_TOPMOST,
        TOOLTIPS_CLASSW,
        nullptr,
        TTS_ALWAYSTIP|TTS_NOPREFIX|TTS_BALLOON,
        CW_USEDEFAULT,CW_USEDEFAULT,CW_USEDEFAULT,CW_USEDEFAULT,
        hWnd,nullptr,g_hInst,nullptr
    );

    if(!g_hMenuTip) return;

    SetWindowPos(g_hMenuTip,HWND_TOPMOST,0,0,0,0,SWP_NOMOVE|SWP_NOSIZE|SWP_NOACTIVATE);
    SendMessageW(g_hMenuTip,TTM_SETMAXTIPWIDTH,0,ScaleByDpi(450,g_uiDpi));
    SendMessageW(g_hMenuTip,WM_SETFONT,(WPARAM)g_hFont,TRUE);

    ZeroMemory(&g_menuTi,sizeof(g_menuTi));
    g_menuTi.cbSize=sizeof(g_menuTi);
    g_menuTi.uFlags=TTF_TRACK|TTF_ABSOLUTE|TTF_TRANSPARENT;
    g_menuTi.hwnd=hWnd;
    g_menuTi.uId=1;
    g_menuTi.lpszText=(LPWSTR)L"";
    SendMessageW(g_hMenuTip,TTM_ADDTOOL,0,(LPARAM)&g_menuTi);

    // Tooltip text per menu item
    g_menuTipText.clear();
    g_menuTipText[IDM_CTX_SET_PROFILE_ICON]=L"Choose an icon/image for profile.";
    g_menuTipText[IDM_CTX_REFRESH_PROFILE]=L"Rebuild this profile from Default.7z (Extract/Replace)";
    g_menuTipText[IDM_CTX_RESET_PROFILE]=L"Delete this profile and recreate from Default.7z.";
    g_menuTipText[IDM_CTX_DELETE_PROFILE]=L"Permanently delete this profile folder.";
    g_menuTipText[IDM_CTX_EDIT_DEFAULT_PROFILE]=L"Open the Default profile for editing.";
    g_menuTipText[IDM_ABOUT]=L"About ctSpaces";
}

static void HideMenuTooltip(){
    if(!g_hMenuTip) return;
    SendMessageW(g_hMenuTip,TTM_TRACKACTIVATE,FALSE,(LPARAM)&g_menuTi);
}

static void HandleMenuSelect(HWND hWnd,WPARAM wParam,LPARAM lParam){
    // Menu closing / no selection
    const UINT item=(UINT)LOWORD(wParam);
    const UINT flags=(UINT)HIWORD(wParam);

    if(item==0xFFFF||lParam==0){
        HideMenuTooltip();
        return;
    }

    // Popup highlight gives index, not command id
    if(flags&MF_POPUP){
        HideMenuTooltip();
        return;
    }

    auto it=g_menuTipText.find(item);
    if(it==g_menuTipText.end()){
        HideMenuTooltip();
        return;
    }

    // Don't show tooltip for disabled items
    if((flags&MF_DISABLED)||(flags&MF_GRAYED)){
        HideMenuTooltip();
        return;
    }

    // Show near cursor
    POINT pt{};
    GetCursorPos(&pt);

    g_menuTi.lpszText=(LPWSTR)it->second.c_str();
    SendMessageW(g_hMenuTip,TTM_UPDATETIPTEXTW,0,(LPARAM)&g_menuTi);
    SendMessageW(g_hMenuTip,TTM_TRACKPOSITION,0,MAKELPARAM(pt.x+12,pt.y+18));
    SendMessageW(g_hMenuTip,TTM_TRACKACTIVATE,TRUE,(LPARAM)&g_menuTi);
}

static void UpdateConfigMenuEnabledState(){
    if(!g_hConfigMenu) return;

    std::wstring name=GetSelectedClientNameSanitized(false);
    const bool hasSelection=!name.empty();

    auto en=[&](UINT id,bool enabled){
        EnableMenuItem(
            g_hConfigMenu,
            id,
            MF_BYCOMMAND|(enabled?MF_ENABLED:(MF_DISABLED|MF_GRAYED))
        );
    };

    // Disable everything requiring a selection when empty; ALWAYS allow Default/About
    en(IDM_CTX_SET_PROFILE_ICON,hasSelection);
    en(IDM_CTX_REFRESH_PROFILE,hasSelection);
    en(IDM_CTX_RESET_PROFILE,hasSelection);
    en(IDM_CTX_DELETE_PROFILE,hasSelection);

    en(IDM_CTX_EDIT_DEFAULT_PROFILE,true);
    en(IDM_ABOUT,true);
}
static void LayoutMainGui(HWND hWnd){
    if(!hWnd) return;

    const UINT dpi=GetDpiForWindow(hWnd);



    const int m=MulDiv(8,dpi,96);
    const int labelH=MulDiv(16,dpi,96);
    const int comboVisH=MulDiv(24,dpi,96);
    const int comboDropH=MulDiv(220,dpi,96);

    const int btnSmall=MulDiv(28,dpi,96);

    // Icon target size (32 DIP)
    const int iconPx=MulDiv(32,dpi,96);

    // EXACT: 2px margin around icon inside groupbox
    const int iconPad=MulDiv(6,dpi,96);

    // Label: 4px from top-left, and -2px vertical (overlap border like a caption)
    const int lblOffX=MulDiv(2,dpi,96);
    const int lblOffY=MulDiv(1,dpi,96);
    const int lblH=MulDiv(14,dpi,96);
    const int lblRise=MulDiv(2,dpi,96); // "Icon label should be -2px"

    // Tight groupbox: label + pad + icon + pad
    const int grpW=iconPx+iconPad*2;
    const int grpH=(lblOffY+lblH+iconPad+iconPx+iconPad)-MulDiv(6,dpi,96);

    RECT rc{};
    GetClientRect(hWnd,&rc);
    const int W=rc.right-rc.left;
    const int H=rc.bottom-rc.top;

    // Prompt + combo (top area)
    HWND hPrompt=GetDlgItem(hWnd,IDC_STATIC_PROMPT);

    int x=m;
    int y=m;
    int w=W-2*m;

    if(hPrompt) MoveWindow(hPrompt,x,y,w,labelH,TRUE);
    y+=labelH+m;

    MoveWindow(g_hComboClient,x,y,w,comboDropH,TRUE);

    // --- Bottom row anchored to window bottom (match ORIGINAL small buttons) ---
    const int bm=MulDiv(4,dpi,96);          // original iGuiM = 4
    const int yBottom=H-bm;

    // Groupbox top so its bottom sits at (H - m)
    const int yRow=yBottom-grpH;

    // Right buttons anchored to bottom-right with bm margin + bm spacing (original behavior)
    const int xCfg=W-bm-btnSmall;
    const int xTmp=xCfg-bm-btnSmall+MulDiv(2,dpi,96);
    const int yBtn=H-bm-btnSmall;

    const int grpX=bm;

    MoveWindow(g_hGrpIcon,grpX,yRow,grpW,grpH,TRUE);

    MoveWindow(
        g_hLblIcon,
        grpX+lblOffX,
        yRow+lblOffY-lblRise,
        grpW-lblOffX-iconPad-MulDiv(6,dpi,96),
        lblH,
        TRUE
    );

    const int xIco=grpX+iconPad;
    const int yIco=yRow+lblOffY+lblH+iconPad-MulDiv(6,dpi,96);
    MoveWindow(g_hIconPreview,xIco,yIco,iconPx,iconPx,TRUE);

    // Right buttons
    MoveWindow(g_hBtnTmpProf,xTmp,yBtn,btnSmall,btnSmall,TRUE);
    MoveWindow(g_hBtnConfig,xCfg,yBtn,btnSmall,btnSmall,TRUE);


    // --- Go centered/aligned with the combo (horizontal centerline of combo area) ---
    const int leftEdge=x+grpW+m;
    const int rightEdge=xTmp-m;
    const int avail=max(0,rightEdge-leftEdge);

    const int goW=MulDiv(64,dpi,96);
    const int goH=MulDiv(44,dpi,96);

    // Centered on the same horizontal centerline as the combo/content
    const int goX = x + (w - goW) / 2;

    // Under the visible combo height (like the original UI flow)
    const int goY=y+comboVisH+m;//+MulDiv(8,dpi,96);

    MoveWindow(g_hBtnGo, goX, goY, goW, goH, TRUE);
}

static std::wstring GetSelectedClientNameSanitized(bool preferListSelection){
    // Edit text (what the user typed)
    wchar_t editBuf[256]{};
    GetWindowTextW(g_hComboClient,editBuf,256);
    std::wstring editName=SanitizeName(editBuf);

    // Current list selection (what the dropdown is on)
    wchar_t selBuf[256]{};
    std::wstring selName;
    int sel=(int)SendMessageW(g_hComboClient,CB_GETCURSEL,0,0);
    if(sel!=CB_ERR){
        SendMessageW(g_hComboClient,CB_GETLBTEXT,sel,(LPARAM)selBuf);
        selName=SanitizeName(selBuf);
    }

    // If caller wants list selection, use it when available
    if(preferListSelection&&!selName.empty()) return selName;

    // Otherwise prefer what the user typed, fallback to selection
    if(!editName.empty()) return editName;
    return selName; // may be empty
}
