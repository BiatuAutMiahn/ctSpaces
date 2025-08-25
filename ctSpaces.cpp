// ctSpaces.cpp

// Because I didn't have the paitence to port this from the ground up, I used Gemini 2.5 Pro for heavylifting, and filled in the gaps.

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

#include "resource.h" // For IDR_7ZA and IDR_DEFAULT_7Z

namespace fs=std::filesystem;

struct IconButtonInfo{
    int id;
    HWND hWnd;
    std::wstring symbol;
    std::wstring tooltip;
    std::function<void()> handler;
};

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
const std::wstring APP_VERSION=L"2.1";
const std::wstring APP_TITLE=std::format(L"{} v{}b",APP_ALIAS,APP_VERSION);
const std::wstring GUI_CLASS_NAME=L"ctSpacesLauncherClass";
ULONG_PTR g_gdiplusToken;
HINSTANCE g_hInst;
HWND g_hGui=NULL;
HWND g_hComboClient=NULL;
HWND g_hValidationTooltip=NULL;
HWND g_hBtnGo=NULL;
HFONT g_hFont=NULL;
fs::path g_sDataDir;
fs::path g_sEdgePath;
std::wstring g_sLastValidComboText=L"";
std::wstring g_sClientSel=L"";
std::jthread g_watcherThread;
std::mutex g_activeProfilesMutex;
std::atomic<bool> g_isWatcherRunning=false;
std::map<std::wstring,HICON> g_iconCache;
std::mutex g_iconCacheMutex;
IShellLink* shellLink=NULL;
#define WM_APP_TASK_COMPLETE (WM_APP + 1)

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
std::vector<IconButtonInfo> g_iconButtons;
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
bool RunCommand(const std::wstring& command,const fs::path& workingDir);
bool extDef(const fs::path& profileDataPath);
void CleanupProfile(const fs::path& profilePath,const std::vector<std::wstring>& keepList);
void SetWindowAppId(HWND hWnd,const std::wstring& appId);
LRESULT CALLBACK WndProc(HWND,UINT,WPARAM,LPARAM);
ATOM MyRegisterClass(HINSTANCE hInstance);
BOOL InitInstance(HINSTANCE,int);
std::wstring AnsiToWide(const std::string& str);
std::vector<std::wstring> GetSupportedImageTypes();
bool ConvertImageToIcon(const fs::path& sourceImagePath,const fs::path& destIconPath);
bool SaveIconsToFile(const fs::path& filePath,std::vector<HICON>& icons,bool compressLargeImages=true);
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


int APIENTRY wWinMain(_In_ HINSTANCE hInstance,_In_opt_ HINSTANCE,_In_ LPWSTR lpCmdLine,_In_ int nCmdShow){
    CoInitializeEx(NULL,COINIT_APARTMENTTHREADED|COINIT_DISABLE_OLE1DDE);
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
    ExtractResourceToFile(IDR_7ZAX64,g_sDataDir/"7za.exe");
    ExtractResourceToFile(IDR_DEFPROF,g_sDataDir/"Default.7z");
    INITCOMMONCONTROLSEX icex={sizeof(INITCOMMONCONTROLSEX), ICC_WIN95_CLASSES};
    InitCommonControlsEx(&icex);
    bool useTrdLayout=(wcsstr(lpCmdLine,L"~!Trd:P")!=nullptr);
    g_iconButtons={
        { 200, NULL, L"I", L"Set Profile Icon", GuiSetIcon },
        { 201, NULL, L"R", L"Reset Selected Profile to Default (...)", GuiProfReset },
        { 202, NULL, L"U", L"Update Selected Profile with Default (...)", GuiProfUpd },
        { 203, NULL, L"D", L"Open Default Profile", GuiOpenDef },
        { 204, NULL, L"T", L"Launch temporary profile", GuiOpenTmp },
    };
    if(useTrdLayout){
        std::vector<IconButtonInfo> trdLayoutButtons={
            g_iconButtons[0],
            g_iconButtons[4],
            g_iconButtons[2],
            g_iconButtons[1],
            g_iconButtons[3]
        };
        g_iconButtons.swap(trdLayoutButtons);
    }

    MyRegisterClass(hInstance);
    if(!InitInstance(hInstance,nCmdShow)){
        return FALSE;
    }
    MSG msg;
    while(GetMessage(&msg,nullptr,0,0)){
        if(!IsDialogMessage(g_hGui,&msg)){
            TranslateMessage(&msg);
            DispatchMessage(&msg);
        }
    }
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
    const int iGuiW=256+64;
    const int iGuiH=128+4;
    const int iGuiM=4;
    const int iGuiCtrlW=(iGuiW-iGuiM*2);
    g_hFont=CreateFontW(15,0,0,0,FW_NORMAL,FALSE,FALSE,FALSE,DEFAULT_CHARSET,OUT_DEFAULT_PRECIS,CLIP_DEFAULT_PRECIS,DEFAULT_QUALITY,DEFAULT_PITCH|FF_MODERN,L"Consolas");
    g_hGui=CreateWindowW(GUI_CLASS_NAME.c_str(),APP_TITLE.c_str(),WS_OVERLAPPED|WS_CAPTION|WS_SYSMENU|WS_MINIMIZEBOX,CW_USEDEFAULT,0,iGuiW,iGuiH,nullptr,nullptr,hInstance,nullptr);
    if(!g_hGui) return FALSE;
    HICON hAppIcon=LoadIcon(hInstance,MAKEINTRESOURCE(IDI_CTSPACES));
    if(hAppIcon){
        SendMessage(g_hGui,WM_SETICON,ICON_BIG,(LPARAM)hAppIcon);
        SendMessage(g_hGui,WM_SETICON,ICON_SMALL,(LPARAM)hAppIcon);
    }
    BOOL isDarkMode=TRUE;
    DwmSetWindowAttribute(g_hGui,DWMWA_USE_IMMERSIVE_DARK_MODE,&isDarkMode,sizeof(isDarkMode));
    CreateWindowW(L"STATIC",L"Select or type the client name:",WS_CHILD|WS_VISIBLE,iGuiM,iGuiM,iGuiCtrlW,17,g_hGui,(HMENU)101,hInstance,nullptr);
    g_hValidationTooltip=CreateWindowEx(WS_EX_TOPMOST,TOOLTIPS_CLASS,NULL,TTS_BALLOON|TTS_NOPREFIX|TTS_ALWAYSTIP,CW_USEDEFAULT,CW_USEDEFAULT,CW_USEDEFAULT,CW_USEDEFAULT,g_hGui,NULL,g_hInst,NULL);
    SendMessage(g_hValidationTooltip,TTM_SETMAXTIPWIDTH,0,400);
    g_hComboClient=CreateWindowW(L"COMBOBOX",L"",WS_CHILD|WS_VISIBLE|CBS_DROPDOWN|CBS_AUTOHSCROLL|WS_VSCROLL,iGuiM,17+iGuiM*2,iGuiCtrlW-iGuiM*4,150,g_hGui,(HMENU)102,hInstance,nullptr);
    TOOLINFOW tic={sizeof(TOOLINFOW)};
    tic.uFlags=TTF_SUBCLASS|TTF_TRANSPARENT|TTF_TRACK;
    tic.hwnd=g_hGui;
    tic.hinst=g_hInst;
    tic.uId=(UINT_PTR)g_hComboClient;
    tic.lpszText=LPSTR_TEXTCALLBACK;
    SendMessage(g_hValidationTooltip,TTM_ADDTOOL,0,(LPARAM)&tic);
    g_hBtnGo=CreateWindowW(L"BUTTON",L"Go",WS_CHILD|WS_VISIBLE|BS_DEFPUSHBUTTON,(iGuiCtrlW/2)-32-4,(25*2)+iGuiM,64,33,g_hGui,(HMENU)IDOK,hInstance,nullptr);
    HWND hToolTip=CreateWindowEx(0,TOOLTIPS_CLASS,NULL,TTS_ALWAYSTIP|TTS_NOPREFIX,CW_USEDEFAULT,CW_USEDEFAULT,CW_USEDEFAULT,CW_USEDEFAULT,g_hGui,NULL,g_hInst,NULL);
    const int iBtnS=18,iBtnM=2;
    int iBtnL=iGuiCtrlW-iGuiM*2-((iBtnS+iBtnM)*(static_cast<int>(g_iconButtons.size())+1));
    int iBtnT=iGuiH-iBtnM-iBtnS-32-8;
    for(size_t i=0; i<g_iconButtons.size(); ++i){
        int xPos=iBtnL+((iBtnS+iBtnM)*(static_cast<int>(i)+1));
        g_iconButtons[i].hWnd=CreateWindowW(L"BUTTON",g_iconButtons[i].symbol.c_str(),WS_CHILD|WS_VISIBLE,xPos,iBtnT,iBtnS,iBtnS,g_hGui,(HMENU)(INT_PTR)g_iconButtons[i].id,g_hInst,nullptr);
        TOOLINFOW ti={sizeof(TOOLINFOW)};
        ti.uFlags=TTF_SUBCLASS;
        ti.hwnd=g_iconButtons[i].hWnd;
        ti.hinst=g_hInst;
        ti.lpszText=(LPWSTR)g_iconButtons[i].tooltip.c_str();
        SendMessage(hToolTip,TTM_ADDTOOL,0,(LPARAM)&ti);
    }
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
                    SendMessage(g_hValidationTooltip,TTM_TRACKPOSITION,0,MAKELPARAM(rect.left+4,rect.bottom-4));
                    SendMessage(g_hValidationTooltip,TTM_TRACKACTIVATE,TRUE,(LPARAM)&ti);
                } else{
                    g_sLastValidComboText=currentText;
                    TOOLINFOW ti={sizeof(TOOLINFOW)};
                    ti.hwnd=g_hGui;
                    ti.uId=(UINT_PTR)g_hComboClient;
                    SendMessage(g_hValidationTooltip,TTM_TRACKACTIVATE,FALSE,(LPARAM)&ti);
                    if(currentText.length()>0){
                        LRESULT matchIndex=SendMessage(g_hComboClient,CB_FINDSTRING,(WPARAM)-1,(LPARAM)currentText.c_str());
                        if(matchIndex!=CB_ERR){
                            wchar_t listItem[256];
                            SendMessage(g_hComboClient,CB_GETLBTEXT,matchIndex,(LPARAM)listItem);
                            SetWindowText(g_hComboClient,listItem);
                            SendMessage(g_hComboClient,CB_SETEDITSEL,0,MAKELPARAM(currentText.length(),-1));
                        }
                    }
                }
            }
        } else if(wmId==IDOK){
            GuiProfOpen();
        } else{
            auto it=std::find_if(g_iconButtons.begin(),g_iconButtons.end(),[wmId](const auto& btn){
                return btn.id==wmId;
                                 });
            if(it!=g_iconButtons.end()&&it->handler){
                it->handler();
            }
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
    case WM_APP_TASK_COMPLETE: {
        SetUiState(true);
        wchar_t clientNameBuffer[256];
        GetWindowText(g_hComboClient,clientNameBuffer,256);
        std::wstring clientName=SanitizeName(clientNameBuffer);
        g_sClientSel=clientNameBuffer;
        UpdateClientsComboBox();
        LRESULT selectionIndex=SendMessage(g_hComboClient,CB_FINDSTRINGEXACT,(WPARAM)-1,(LPARAM)g_sClientSel.c_str());
        if(selectionIndex!=CB_ERR){
            SendMessage(g_hComboClient,CB_SETCURSEL,(WPARAM)selectionIndex,0);
        } else{
            SetWindowText(g_hComboClient,g_sClientSel.c_str());
        }
        SetFocus(g_hComboClient);
        SendMessage(g_hComboClient,CB_SETEDITSEL,0,MAKELPARAM(0,-1));

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
        if(g_hFont) DeleteObject(g_hFont);
        PostQuitMessage(0);
        break;
    }
    default:
        return DefWindowProc(hWnd,message,wParam,lParam);
    }
    return 0;
}
void GuiProfOpen(){
    wchar_t clientNameBuffer[256];
    GetWindowText(g_hComboClient,clientNameBuffer,256);
    std::wstring clientName=SanitizeName(clientNameBuffer);
    if(clientName.empty()){
        MessageBox(g_hGui,L"Please select or enter a valid client name.",L"Input Error",MB_OK|MB_ICONWARNING);
        return;
    }
    std::lock_guard<std::mutex> lock(g_activeProfilesMutex);
    if(g_activeProfiles.count(clientName)){
        MessageBox(g_hGui,L"This profile is already open.",L"Already Running",MB_OK|MB_ICONINFORMATION);
        return;
    }
    DWORD pid=LaunchProfile(clientName,false,false);
    if(pid>0){
        g_activeProfiles[clientName]=pid;
        g_reaperThreads.emplace_back(ReaperThread,pid,clientName,ProfileType::Standard);
        EnsureWatcherIsRunning();
        UpdateClientsComboBox();
        LRESULT selectionIndex=SendMessage(g_hComboClient,CB_FINDSTRINGEXACT,(WPARAM)-1,(LPARAM)clientName.c_str());
        if(selectionIndex!=CB_ERR){
            SendMessage(g_hComboClient,CB_SETCURSEL,(WPARAM)selectionIndex,0);
        }
    }
}

void GuiSetIcon(){
    wchar_t clientNameBuffer[256];
    GetWindowText(g_hComboClient,clientNameBuffer,256);
    std::wstring clientName=SanitizeName(clientNameBuffer);
    if(clientName.empty()){
        MessageBox(g_hGui,L"Please select a client first.",L"Warning",MB_OK|MB_ICONWARNING);
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
                if(!success){
                    errorDetails=L"Could not convert image to icon format.";
                }
            }
        } catch(const fs::filesystem_error& e){
            success=false;
            errorDetails=AnsiToWide(e.what());
        }
        if(success){
            MessageBox(g_hGui,L"Icon has been configured.",L"Success",MB_OK|MB_ICONINFORMATION);
            std::lock_guard<std::mutex> cacheLock(g_iconCacheMutex);
            if(g_iconCache.count(clientName)){
                if(g_iconCache[clientName]){
                    DestroyIcon(g_iconCache[clientName]);
                }
                HICON hNewIcon=(HICON)LoadImageW(NULL,destIconPath.c_str(),IMAGE_ICON,0,0,LR_LOADFROMFILE|LR_DEFAULTSIZE|LR_SHARED);
                g_iconCache[clientName]=hNewIcon;
            }
        } else{
            std::wstring errorMsg=L"Failed to set icon.";
            if(!errorDetails.empty()){
                errorMsg+=L"\n\nDetails: "+errorDetails;
            }
            MessageBox(g_hGui,errorMsg.c_str(),L"Error",MB_OK|MB_ICONERROR);
        }
    }
}

void GuiProfReset(){
    wchar_t clientNameBuffer[256];
    GetWindowText(g_hComboClient,clientNameBuffer,256);
    std::wstring clientName=SanitizeName(clientNameBuffer);
    if(clientName.empty()){
        MessageBox(g_hGui,L"Please select a client first.",L"Warning",MB_OK|MB_ICONWARNING);
        return;
    }
    std::lock_guard<std::mutex> lock(g_activeProfilesMutex);
    if(g_activeProfiles.count(clientName)){
        MessageBox(g_hGui,L"Cannot reset a profile that is currently active.",L"Action Denied",MB_OK|MB_ICONWARNING);
        return;
    }
    if(MessageBox(g_hGui,L"This will completely delete and reset the profile. Are you sure?",L"Confirm Reset",MB_YESNO|MB_ICONQUESTION)==IDYES){
        fs::path sData=g_sDataDir/"Sites"/clientName;
        try{
            if(fs::exists(sData)){
                fs::remove_all(sData);
            }
            extDef(sData);
            MessageBox(g_hGui,L"Profile has been reset.",L"Success",MB_OK|MB_ICONINFORMATION);
        } catch(...){}
        UpdateClientsComboBox();
        LRESULT selectionIndex=SendMessage(g_hComboClient,CB_FINDSTRINGEXACT,(WPARAM)-1,(LPARAM)clientName.c_str());
        if(selectionIndex!=CB_ERR){
            SendMessage(g_hComboClient,CB_SETCURSEL,(WPARAM)selectionIndex,0);
        }

    }
}

void GuiProfUpd(){
    wchar_t clientNameBuffer[256];
    GetWindowText(g_hComboClient,clientNameBuffer,256);
    std::wstring clientName=SanitizeName(clientNameBuffer);
    if(clientName.empty()){
        MessageBox(g_hGui,L"Please select a client first.",L"Warning",MB_OK|MB_ICONWARNING);
        return;
    }
    std::lock_guard<std::mutex> lock(g_activeProfilesMutex);
    if(g_activeProfiles.count(clientName)){
        MessageBox(g_hGui,L"Cannot update a profile that is currently active.",L"Action Denied",MB_OK|MB_ICONWARNING);
        return;
    }
    if(MessageBox(g_hGui,L"This will update/overwrite the profile. Are you sure?",L"Confirm Reset",MB_YESNO|MB_ICONQUESTION)==IDYES){
        SetUiState(false);
        std::thread([clientName](){
            fs::path sData=g_sDataDir/"Sites"/clientName;
            try{
                if(fs::exists(sData)){
                    fs::remove_all(sData);
                }
                extDef(sData);
            } catch(...){}
            PostMessage(g_hGui,WM_APP_TASK_COMPLETE,0,0);
                    }).detach();
    }
}
void GuiOpenDef(){
    std::lock_guard<std::mutex> lock(g_activeProfilesMutex);
    if(g_activeProfiles.count(L"Default")){
        MessageBox(g_hGui,L"The Default profile is already open.",L"Already Running",MB_OK|MB_ICONINFORMATION);
        return;
    }
    DWORD pid=LaunchProfile(L"Default",false,true);
    if(pid>0){
        g_activeProfiles[L"Default"]=pid;
        g_reaperThreads.emplace_back(ReaperThread,pid,L"Default",ProfileType::Default);
        EnsureWatcherIsRunning();
    }
}

void GuiOpenTmp(){
    std::lock_guard<std::mutex> lock(g_activeProfilesMutex);
    if(g_activeProfiles.count(L"Temp")){
        MessageBox(g_hGui,L"A temporary profile is already open.",L"Already Running",MB_OK|MB_ICONINFORMATION);
        return;
    }
    DWORD pid=LaunchProfile(L"Temp",true,false);
    if(pid>0){
        g_activeProfiles[L"Temp"]=pid;
        g_reaperThreads.emplace_back(ReaperThread,pid,L"Temp",ProfileType::Temporary);
        EnsureWatcherIsRunning();
    }
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
        if(!extDef(profilePath)){
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

bool RunCommand(const std::wstring& command,const fs::path& workingDir){
    fs::path sevenzip=g_sDataDir/L"7za.exe";
    std::wstring fullCmd=std::format(L"\"{}\" {}",sevenzip.c_str(),command);
    STARTUPINFOW si={sizeof(si)};
    PROCESS_INFORMATION pi={};
    si.dwFlags=STARTF_USESHOWWINDOW;
    si.wShowWindow=SW_SHOW;
    if(CreateProcessW(NULL,&fullCmd[0],NULL,NULL,FALSE,0,NULL,workingDir.c_str(),&si,&pi)){
        WaitForSingleObject(pi.hProcess,INFINITE);
        DWORD exitCode;
        GetExitCodeProcess(pi.hProcess,&exitCode);
        CloseHandle(pi.hProcess);
        CloseHandle(pi.hThread);
        return exitCode==0;
    }
    return false;
}

void SetUiState(bool enabled){
    EnableWindow(g_hComboClient,enabled);
    EnableWindow(g_hBtnGo,enabled);
    for(const auto& btn:g_iconButtons){
        EnableWindow(btn.hWnd,enabled);
    }
}

bool extDef(const fs::path& profileDataPath){
    fs::create_directories(profileDataPath);
    std::wstring cmd=std::format(L"x \"{}\\Default.7z\" -y -o\"{}\"",g_sDataDir.c_str(),profileDataPath.c_str());
    if(!RunCommand(cmd,g_sDataDir)){
        return false;
    }
    std::ofstream marker(profileDataPath/"ctSpaces");
    marker.close();
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
    UINT numDecoders=0,size=0;
    Gdiplus::GetImageDecodersSize(&numDecoders,&size);
    if(size==0) return supportedTypes;
    std::unique_ptr<Gdiplus::ImageCodecInfo[]> pImageCodecInfo(new Gdiplus::ImageCodecInfo[size]);
    if(!pImageCodecInfo) return supportedTypes;
    Gdiplus::GetImageDecoders(numDecoders,size,pImageCodecInfo.get());
    for(UINT i=0; i<numDecoders; ++i){
        std::wstring extensions(pImageCodecInfo[i].FilenameExtension);
        std::wstringstream ss(extensions);
        std::wstring ext;
        while(std::getline(ss,ext,L';')){
            if(ext.rfind(L"*.",0)==0){
                ext=ext.substr(2);
            }
            std::transform(ext.begin(),ext.end(),ext.begin(),::towlower);
            if(ext!=L"ico"&&std::find(supportedTypes.begin(),supportedTypes.end(),ext)==supportedTypes.end()){
                supportedTypes.push_back(ext);
            }
        }
    }
    return supportedTypes;
}

bool ConvertImageToIcon(const fs::path& sourceImagePath,const fs::path& destIconPath){
    std::unique_ptr<Gdiplus::Bitmap> sourceBitmap(Gdiplus::Bitmap::FromFile(sourceImagePath.c_str()));
    if(!sourceBitmap||sourceBitmap->GetLastStatus()!=Gdiplus::Ok){
        return false;
    }
    UINT sourceWidth=sourceBitmap->GetWidth();
    UINT sourceHeight=sourceBitmap->GetHeight();
    int masterSize=max(sourceWidth,sourceHeight);
    auto masterScaledBitmap=std::make_unique<Gdiplus::Bitmap>(masterSize,masterSize,PixelFormat32bppARGB);
    {
        auto graphics=std::make_unique<Gdiplus::Graphics>(masterScaledBitmap.get());
        graphics->SetInterpolationMode(Gdiplus::InterpolationModeHighQualityBicubic);
        graphics->SetPixelOffsetMode(Gdiplus::PixelOffsetModeHighQuality);
        auto imageAttributes=std::make_unique<Gdiplus::ImageAttributes>();
        imageAttributes->SetWrapMode(Gdiplus::WrapModeTileFlipXY);
        graphics->DrawImage(
            sourceBitmap.get(),
            Gdiplus::RectF(0.0f,0.0f,(Gdiplus::REAL)masterSize,(Gdiplus::REAL)masterSize),
            0.0f,0.0f,(Gdiplus::REAL)sourceWidth,(Gdiplus::REAL)sourceHeight,
            Gdiplus::UnitPixel,imageAttributes.get()
        );
    }
    const std::vector<int> sizes={256, 128, 64, 48, 32, 24, 16};
    std::vector<HICON> hIcons;
    for(int targetSize:sizes){
        if(targetSize>=masterSize) continue;
        auto finalSizeBitmap=std::make_unique<Gdiplus::Bitmap>(targetSize,targetSize,PixelFormat32bppARGB);
        {
            auto graphics=std::make_unique<Gdiplus::Graphics>(finalSizeBitmap.get());
            graphics->SetInterpolationMode(Gdiplus::InterpolationModeHighQualityBicubic);
            graphics->SetPixelOffsetMode(Gdiplus::PixelOffsetModeHighQuality);
            auto imageAttributes=std::make_unique<Gdiplus::ImageAttributes>();
            imageAttributes->SetWrapMode(Gdiplus::WrapModeTileFlipXY);
            graphics->DrawImage(
                masterScaledBitmap.get(),
                Gdiplus::RectF(0.0f,0.0f,(Gdiplus::REAL)targetSize,(Gdiplus::REAL)targetSize),
                0.0f,0.0f,(Gdiplus::REAL)masterSize,(Gdiplus::REAL)masterSize,
                Gdiplus::UnitPixel,imageAttributes.get()
            );
        }
        HICON hIcon=NULL;
        if(finalSizeBitmap->GetHICON(&hIcon)==Gdiplus::Ok){
            hIcons.push_back(hIcon);
        }
    }
    if(hIcons.empty()){
        return false;
    }
    bool result=SaveIconsToFile(destIconPath,hIcons);
    for(HICON hIcon:hIcons){
        DestroyIcon(hIcon);
    }
    return result;
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
    BITMAP bmp={0};
    if(!GetObject(hBitmap,sizeof(BITMAP),&bmp)){
        return false;
    }
    if(bmp.bmBitsPixel!=32){
        return false;
    }
    int dataSize=bmp.bmWidthBytes*bmp.bmHeight;
    auto pixelData=std::make_unique<BYTE[]>(dataSize);
    if(GetBitmapBits(hBitmap,dataSize,pixelData.get())==0){
        return false;
    }
    for(int i=0; i<dataSize; i+=4){
        if(pixelData[i+3]<255){
            return true;
        }
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
bool SaveIconsToFile(const fs::path& filePath,std::vector<HICON>& icons,bool compressLargeImages){
    if(icons.empty()){
        return false;
    }
    std::ofstream file(filePath,std::ios::binary);
    if(!file.is_open()){
        return false;
    }
    std::vector<HICON> tempIcons;
    auto cleanup=[&](){
        for(HICON hTemp:tempIcons){
            DestroyIcon(hTemp);
        }
    };
    std::vector<std::vector<BYTE>> allIconImageData;
    std::vector<ICONDIRENTRY> allIconDirEntries;
    for(size_t i=0; i<icons.size(); ++i){
        HICON hCurrentImage=icons[i];
        ICONINFOEXW iconInfo={sizeof(ICONINFOEXW)};
        if(!GetIconInfoExW(hCurrentImage,&iconInfo)){
            cleanup();
            file.close();
            return false;
        }
        BITMAP bmpColorInfo={0};
        GetObject(iconInfo.hbmColor,sizeof(BITMAP),&bmpColorInfo);
        if(bmpColorInfo.bmBitsPixel!=32||!IsAlphaBitmap(iconInfo.hbmColor)){
            HICON hNew32BitIcon=Create32BitHICON(hCurrentImage);
            if(hNew32BitIcon){
                DestroyIcon(hCurrentImage);
                icons[i]=hNew32BitIcon;
                hCurrentImage=hNew32BitIcon;
                tempIcons.push_back(hNew32BitIcon);
                GetIconInfoExW(hCurrentImage,&iconInfo);
                GetObject(iconInfo.hbmColor,sizeof(BITMAP),&bmpColorInfo);
            }
        }
        std::vector<BYTE> imageDataBlock;
        bool isCompressed=false;
        if(compressLargeImages&&bmpColorInfo.bmWidth>=256){
            std::vector<BYTE> pngData=CompressBitmapToPng(iconInfo.hbmColor);
            if(!pngData.empty()){
                imageDataBlock=pngData;
                isCompressed=true;
            }
        }
        if(!isCompressed){
            int colorDataSize=bmpColorInfo.bmWidthBytes*bmpColorInfo.bmHeight;
            std::vector<BYTE> colorData(colorDataSize);
            GetBitmapBits(iconInfo.hbmColor,colorDataSize,colorData.data());
            BITMAP bmpMaskInfo={0};
            GetObject(iconInfo.hbmMask,sizeof(BITMAP),&bmpMaskInfo);
            int maskDataSize=bmpMaskInfo.bmWidthBytes*bmpMaskInfo.bmHeight;
            std::vector<BYTE> maskData(maskDataSize);
            GetBitmapBits(iconInfo.hbmMask,maskDataSize,maskData.data());
            BITMAPINFOHEADER bih={0};
            bih.biSize=sizeof(BITMAPINFOHEADER);
            bih.biWidth=bmpColorInfo.bmWidth;
            bih.biHeight=bmpColorInfo.bmHeight*2;
            bih.biPlanes=1;
            bih.biBitCount=bmpColorInfo.bmBitsPixel;
            bih.biCompression=BI_RGB;
            imageDataBlock.insert(imageDataBlock.end(),reinterpret_cast<BYTE*>(&bih),reinterpret_cast<BYTE*>(&bih)+sizeof(bih));
            imageDataBlock.insert(imageDataBlock.end(),colorData.begin(),colorData.end());
            imageDataBlock.insert(imageDataBlock.end(),maskData.begin(),maskData.end());
        }
        allIconImageData.push_back(imageDataBlock);
        ICONDIRENTRY entry={0};
        entry.bWidth=(bmpColorInfo.bmWidth>=256)?0:(BYTE)bmpColorInfo.bmWidth;
        entry.bHeight=(bmpColorInfo.bmHeight>=256)?0:(BYTE)bmpColorInfo.bmHeight;
        entry.wPlanes=1;
        entry.wBitCount=bmpColorInfo.bmBitsPixel;
        entry.dwBytesInRes=static_cast<DWORD>(imageDataBlock.size());
        allIconDirEntries.push_back(entry);
        DeleteObject(iconInfo.hbmColor);
        DeleteObject(iconInfo.hbmMask);
    }
    ICONDIR iconDirHeader={0};
    iconDirHeader.idType=1;
    iconDirHeader.idCount=static_cast<WORD>(icons.size());
    file.write(reinterpret_cast<char*>(&iconDirHeader),sizeof(WORD)*3);
    DWORD currentOffset=sizeof(iconDirHeader.idReserved)+sizeof(iconDirHeader.idType)+sizeof(iconDirHeader.idCount)+(icons.size()*sizeof(ICONDIRENTRY));
    for(auto& entry:allIconDirEntries){
        entry.dwImageOffset=currentOffset;
        currentOffset+=entry.dwBytesInRes;
    }
    file.write(reinterpret_cast<char*>(allIconDirEntries.data()),allIconDirEntries.size()*sizeof(ICONDIRENTRY));
    for(const auto& data:allIconImageData){
        file.write(reinterpret_cast<const char*>(data.data()),data.size());
    }
    file.close();
    cleanup();
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
            try{ fs::rename(backup7z,bakPath); } catch(...){}
            std::wstring pathToArchive=profilePath.wstring()+L"\\*";
            std::wstring sevenZipCmd=std::format(L"a -mx=9 \"{}\" \"{}\"",backup7z.c_str(),pathToArchive.c_str());
            RunCommand(sevenZipCmd,g_sDataDir);
        }
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
    //static std::map<std::wstring,HICON> iconCache;
    while(true){
        std::this_thread::sleep_for(std::chrono::milliseconds(200));
        std::map<std::wstring,DWORD> profiles_copy;
        {
            std::lock_guard<std::mutex> lock(g_activeProfilesMutex);
            if(g_activeProfiles.empty()){
                g_isWatcherRunning=false;
                for(auto const& [name,hIcon]:g_iconCache){
                    if(hIcon) DestroyIcon(hIcon);
                }
                g_iconCache.clear();
                return;
            }
            profiles_copy=g_activeProfiles;
        }
        std::lock_guard<std::mutex> cacheLock(g_iconCacheMutex);
        for(auto it=g_iconCache.begin(); it!=g_iconCache.end(); ){
            if(profiles_copy.find(it->first)==profiles_copy.end()){
                if(it->second) DestroyIcon(it->second);
                it=g_iconCache.erase(it);
            } else{
                ++it;
            }
        }
        for(const auto& [clientName,pid]:profiles_copy){
            if(g_iconCache.find(clientName)==g_iconCache.end()){
                fs::path profilePath;
                if(clientName==L"Temp") profilePath=g_sDataDir/"Temp";
                else if(clientName==L"Default") profilePath=g_sDataDir/"Default";
                else profilePath=g_sDataDir/"Sites"/clientName;

                fs::path iconPath=profilePath/"client.ico";
                HICON hIcon=NULL;
                if(fs::exists(iconPath)){
                    hIcon=(HICON)LoadImageW(NULL,iconPath.c_str(),IMAGE_ICON,0,0,LR_LOADFROMFILE|LR_DEFAULTSIZE|LR_SHARED);
                }
                g_iconCache[clientName]=hIcon;
            }

            HICON hCurrentIcon=g_iconCache[clientName];

            EnumData data={pid};
            EnumWindows(EnumWindowsCallback,(LPARAM)&data);
            if(data.windows.empty()) continue;

            std::wstring newTitlePartial=std::format(L" - {} - {}",clientName,APP_ALIAS);

            for(HWND hWnd:data.windows){
                if(hCurrentIcon){
                    if((HICON)SendMessage(hWnd,WM_GETICON,ICON_SMALL,0)!=hCurrentIcon){
                        SendMessage(hWnd,WM_SETICON,ICON_SMALL,(LPARAM)hCurrentIcon);
                    }
                    if((HICON)SendMessage(hWnd,WM_GETICON,ICON_BIG,0)!=hCurrentIcon){
                        SendMessage(hWnd,WM_SETICON,ICON_BIG,(LPARAM)hCurrentIcon);
                    }
                }
                wchar_t currentTitle[256];
                GetWindowTextW(hWnd,currentTitle,256);
                std::wstring titleStr(currentTitle);
                std::wregex expr(L"(.*) - Profile 1 - Microsoft.*Edge");
                std::wsmatch match;
                if(std::regex_match(titleStr,match,expr)&&match.size()>1){
                    std::wstring newTitle=match[1].str()+newTitlePartial;
                    SetWindowTextW(hWnd,newTitle.c_str());
                }
                std::wstring appId=std::format(L"ctSpaces.{}.Default",clientName);
                SetWindowAppId(hWnd,appId);
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