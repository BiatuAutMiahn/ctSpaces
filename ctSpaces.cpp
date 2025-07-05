// ctSpaces.cpp

// Because I didn't have the paitence to port this from the ground up, I used Gemini 2.5 Pro
//   for boilerplate, and did the rest.

#pragma comment(linker,"/manifestdependency:\"type='win32' name='Microsoft.Windows.Common-Controls' version='6.0.0.0' processorArchitecture='*' publicKeyToken='6595b64144ccf1df' language='*'\"")
#pragma comment(lib, "dwmapi.lib")
#pragma comment(lib, "uxtheme.lib")
#pragma comment(lib, "Comctl32.lib")
#pragma comment(lib, "Shlwapi.lib")
#pragma comment(lib, "Ole32.lib")
#pragma comment(lib, "Propsys.lib")
#pragma comment(lib, "gdiplus.lib")

#include <windows.h>
#include <commctrl.h>
#include <shlobj.h>
#include <shlwapi.h>
#include <dwmapi.h>
#include <uxtheme.h>
#include <propkey.h>
#include <propvarutil.h>
#include <gdiplus.h>

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

#include "resource.h" // For IDR_7ZA and IDR_DEFAULT_7Z

namespace fs=std::filesystem;
const std::wstring APP_ALIAS=L"ctSpaces";
const std::wstring APP_VERSION=L"2.0";
const std::wstring APP_TITLE=std::format(L"{} v{}",APP_ALIAS,APP_VERSION);
const std::wstring GUI_CLASS_NAME=L"ctSpacesLauncherClass";
#define WM_APP_TASK_COMPLETE (WM_APP + 1)

// --- Constants and Globals ---
const std::vector<std::wstring> aKeepDefault={
    L"Local State", L"Last Version", L"Last Browser", L"FirstLaunchAfterInstallation", L"First Run", L"ctSpaces",
    L"DevToolsActivePort", L"Default\\Shortcuts", L"Default\\Shortcuts-journal", L"Default\\Secure Preferences",
    L"Default\\Preferences", L"Default\\Favicons-journal", L"Default\\Favicons", L"Default\\Bookmarks",
    L"Default\\Extension State", L"Default\\Extensions", L"Default\\Local Extension Settings"
};
const std::vector<std::wstring> aKeepActive={
    L"client.ico", L"client.png", L"ctSpaces", L"Local State", L"Last Version", L"Last Browser",
    L"FirstLaunchAfterInstallation", L"First Run", L"DevToolsActivePort", L"Default\\History", L"Default\\Shortcuts",
    L"Default\\Shortcuts-journal", L"Default\\Secure Preferences", L"Default\\Preferences", L"Default\\Favicons-journal",
    L"Default\\Favicons", L"Default\\Bookmarks", L"Default\\Extension State", L"Default\\Extensions",
    L"Default\\Local Extension Settings", L"CertificateRevocation", L"AutoLaunchProtocolsComponent", L"Default\\History",
    L"Default\\Web Data", L"Default\\Web Data-journal", L"Default\\Login Data", L"Default\\Login Data-journal",
    L"Default\\Favicons", L"Default\\Favicons-journal", L"Default\\MediaDeviceSalts", L"Default\\MediaDeviceSalts-journal",
    L"Default\\CdmStorage.db", L"Default\\CdmStorage.db-journal", L"Default\\DIPS", L"Default\\DIPS-journal",
    L"Default\\Local Storage", L"Default\\WebStorage", L"Default\\Service Worker\\Database",
    L"Default\\ClientCertificates", L"Default\\blob_storage", L"Default\\Session Storage", L"Default\\IndexedDB",
    L"Default\\Network", L"Default\\Sessions"
};
std::wstring g_sLastValidComboText=L"";
fs::path g_sDataDir;

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

std::vector<IconButtonInfo> g_iconButtons;
std::wstring g_sClientSel=L"";

ULONG_PTR g_gdiplusToken;
HINSTANCE g_hInst;
HWND g_hGui=NULL;
HWND g_hComboClient=NULL;
HWND g_hValidationTooltip=NULL;
HWND g_hBtnGo=NULL;
HFONT g_hFont=NULL;

// --- Function Prototypes ---
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

int APIENTRY wWinMain(_In_ HINSTANCE hInstance,_In_opt_ HINSTANCE,_In_ LPWSTR lpCmdLine,_In_ int nCmdShow){
    CoInitializeEx(NULL,COINIT_APARTMENTTHREADED|COINIT_DISABLE_OLE1DDE);
    Gdiplus::GdiplusStartupInput gdiplusStartupInput;
    Gdiplus::GdiplusStartup(&g_gdiplusToken,&gdiplusStartupInput,NULL);
    PWSTR path=NULL;
    if(SUCCEEDED(SHGetKnownFolderPath(FOLDERID_LocalAppData,0,NULL,&path))){
        g_sDataDir=fs::path(path)/"InfinitySys"/"ctSpaces";
        CoTaskMemFree(path);
    }
    std::filesystem::create_directories(g_sDataDir);

    // Extract embedded tools if they don't exist
    ExtractResourceToFile(IDR_7ZAX64,g_sDataDir/"7za.exe");
    ExtractResourceToFile(IDR_DEFPROF,g_sDataDir/"Default.7z");

    INITCOMMONCONTROLSEX icex={sizeof(INITCOMMONCONTROLSEX), ICC_WIN95_CLASSES};
    InitCommonControlsEx(&icex);

    bool useTrdLayout=(wcsstr(lpCmdLine,L"~!Trd:P")!=nullptr);
    if(useTrdLayout){
        g_iconButtons={
            { 200, NULL, L"I", L"Set Profile Icon", GuiSetIcon },
            { 204, NULL, L"T", L"Launch temporary profile", GuiOpenTmp },
            { 202, NULL, L"U", L"Update Selected Profile with Default (...)", GuiProfUpd },
            { 201, NULL, L"R", L"Reset Selected Profile to Default (...)", GuiProfReset },
            { 203, NULL, L"D", L"Open Default Profile", GuiOpenDef },
        };
    } else{
        g_iconButtons={
            { 200, NULL, L"I", L"Set Profile Icon", GuiSetIcon },
            { 201, NULL, L"R", L"Reset Selected Profile to Default (...)", GuiProfReset },
            { 202, NULL, L"U", L"Update Selected Profile with Default (...)", GuiProfUpd },
            { 203, NULL, L"D", L"Open Default Profile", GuiOpenDef },
            { 204, NULL, L"T", L"Launch temporary profile", GuiOpenTmp },
        };
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

    g_hGui=CreateWindowW(GUI_CLASS_NAME.c_str(),APP_TITLE.c_str(),WS_OVERLAPPED|WS_CAPTION|WS_SYSMENU,CW_USEDEFAULT,0,iGuiW,iGuiH,nullptr,nullptr,hInstance,nullptr);
    if(!g_hGui) return FALSE;

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
        // FIX: Use (HMENU)(INT_PTR) to cast ID safely on x64
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
        if(wmId==102){ // g_hComboClient
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

                    // Auto-complete logic
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
            // FIX: Use std::find_if from <algorithm> for compatibility.
            auto it=std::find_if(g_iconButtons.begin(),g_iconButtons.end(),[wmId](const auto& btn){
                return btn.id==wmId;
                                 });
            if(it!=g_iconButtons.end()&&it->handler){
                it->handler();
            }
        }
        return 0;
    }
    case WM_DESTROY: {
        if(g_hFont) DeleteObject(g_hFont);
        PostQuitMessage(0);
        break;
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
        SetFocus(g_hComboClient);
        return;
    }
    SetUiState(false);
    std::thread(LaunchAndManageProfile,clientName,false,false).detach();
}
void GuiSetIcon(){
    wchar_t clientNameBuffer[256];
    GetWindowText(g_hComboClient,clientNameBuffer,256);
    std::wstring clientName=SanitizeName(clientNameBuffer);
    if(clientName.empty()){
        MessageBox(g_hGui,L"Please select a client first.",L"Warning",MB_OK|MB_ICONWARNING);
        return;
    }

    // 1. Build the file dialog filter string from supported types
    std::wstring filter;
    std::wstring allSupportedExtensions=L"*.ico";
    std::vector<std::wstring> types=GetSupportedImageTypes();
    for(const auto& type:types){
        allSupportedExtensions+=L";*."+type;
    }

    filter+=L"Supported Image Files ("+allSupportedExtensions+L")";
    filter+=L'\0'; // Null terminator
    filter+=allSupportedExtensions;
    filter+=L'\0'; // Null terminator
    filter+=L"Icon Files (*.ico)";
    filter+=L'\0';
    filter+=L"*.ico";
    filter+=L'\0';
    filter+=L"All Files (*.*)";
    filter+=L'\0';
    filter+=L"*.*";
    filter+=L'\0'; // Double null terminator at the end
    filter+=L'\0';

    // 2. Show the file open dialog
    wchar_t szFile[MAX_PATH]={0};
    OPENFILENAMEW ofn={0};
    ofn.lStructSize=sizeof(ofn);
    ofn.hwndOwner=g_hGui;
    ofn.lpstrFile=szFile;
    ofn.nMaxFile=sizeof(szFile)/sizeof(wchar_t);
    ofn.lpstrFilter=filter.c_str(); // Use the generated filter
    ofn.nFilterIndex=1;
    ofn.Flags=OFN_PATHMUSTEXIST|OFN_FILEMUSTEXIST;

    if(GetOpenFileNameW(&ofn)){
        fs::path sourcePath(ofn.lpstrFile);
        fs::path destIconPath=g_sDataDir/"Sites"/clientName/"client.ico";
        fs::create_directories(destIconPath.parent_path());

        bool success=false;
        std::wstring errorDetails;

        try{
            // 3. Check extension and either copy or convert
            if(_wcsicmp(sourcePath.extension().c_str(),L".ico")==0){
                // It's already an icon, just copy it
                fs::copy_file(sourcePath,destIconPath,fs::copy_options::overwrite_existing);
                success=true;
            } else{
                // It's another image type, convert it
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
    if(MessageBox(g_hGui,L"This will completely delete and reset the profile. Are you sure?",L"Confirm Reset",MB_YESNO|MB_ICONQUESTION)==IDYES){
        SetUiState(false);
        std::thread([clientName](){
            fs::path sData=g_sDataDir/"Sites"/clientName;
            try{
                if(fs::exists(sData)){
                    fs::remove_all(sData);
                }
                extDef(sData);
            } catch(...){ /* Handle errors if necessary */ }
            PostMessage(g_hGui,WM_APP_TASK_COMPLETE,0,0);
                    }).detach();
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
    if(MessageBox(g_hGui,L"This will update/overwrite the profile. Are you sure?",L"Confirm Reset",MB_YESNO|MB_ICONQUESTION)==IDYES){
        SetUiState(false);
        std::thread([clientName](){
            fs::path sData=g_sDataDir/"Sites"/clientName;
            try{
                if(fs::exists(sData)){
                    fs::remove_all(sData);
                }
                extDef(sData);
            } catch(...){ /* Handle errors if necessary */ }
            PostMessage(g_hGui,WM_APP_TASK_COMPLETE,0,0);
                    }).detach();
    }
}

void GuiOpenDef(){
    SetUiState(false);
    std::thread(LaunchAndManageProfile,L"Default",false,true).detach();
}

void GuiOpenTmp(){
    SetUiState(false);
    std::thread(LaunchAndManageProfile,L"Temp",true,false).detach();
}

// --- Helper Functions ---

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

void LaunchAndManageProfile(const std::wstring& clientName,bool isTemp,bool isDefault){
    fs::path profilePath;
    if(isTemp) profilePath=g_sDataDir/"Temp";
    else if(isDefault) profilePath=g_sDataDir/"Default";
    else profilePath=g_sDataDir/"Sites"/clientName;

    if(!fs::exists(profilePath/"ctSpaces")){
        if(!extDef(profilePath)){
            MessageBox(NULL,L"Error: An error occurred when extracting the profile.",APP_TITLE.c_str(),MB_OK|MB_ICONERROR);
            PostMessage(g_hGui,WM_APP_TASK_COMPLETE,0,0);
            return;
        }
    }

    std::wstring cmdLine=std::format(L"\"C:\\Program Files (x86)\\Microsoft\\Edge\\Application\\msedge.exe\" --user-data-dir=\"{}\" --no-first-run",profilePath.c_str());
    STARTUPINFOW si={sizeof(si)};
    PROCESS_INFORMATION pi={};
    if(!CreateProcessW(NULL,&cmdLine[0],NULL,NULL,FALSE,0,NULL,NULL,&si,&pi)){
        MessageBox(NULL,L"Failed to launch Microsoft Edge. Is it installed?",APP_TITLE.c_str(),MB_OK|MB_ICONERROR);
        PostMessage(g_hGui,WM_APP_TASK_COMPLETE,0,0);
        return;
    }

    HICON hIcon=NULL;
    fs::path iconPath=profilePath/"client.ico";
    if(fs::exists(iconPath)){
        hIcon=(HICON)LoadImageW(NULL,iconPath.c_str(),IMAGE_ICON,0,0,LR_LOADFROMFILE|LR_DEFAULTSIZE);
    }

    std::wstring newTitlePartial=std::format(L" - {} - {}",clientName,APP_ALIAS);
    while(WaitForSingleObject(pi.hProcess,500)==WAIT_TIMEOUT){
        EnumData data={pi.dwProcessId};
        EnumWindows(EnumWindowsCallback,(LPARAM)&data);
        for(HWND hWnd:data.windows){
            wchar_t currentTitle[256];
            GetWindowTextW(hWnd,currentTitle,256);
            std::wregex expr(L"(.*) - Profile 1 - Microsoft.*Edge");
            std::wsmatch match;
            std::wstring titleStr(currentTitle);
            if(std::regex_match(titleStr,match,expr)&&match.size()>1){
                std::wstring newTitle=match[1].str()+newTitlePartial;
                SetWindowTextW(hWnd,newTitle.c_str());
            }
            if(hIcon){
                SendMessage(hWnd,WM_SETICON,ICON_SMALL,(LPARAM)hIcon);
                SendMessage(hWnd,WM_SETICON,ICON_BIG,(LPARAM)hIcon);
            }
            std::wstring appId=std::format(L"ctSpaces.{}.Default",clientName);
            SetWindowAppId(hWnd,appId);
        }
    }

    if(hIcon) DestroyIcon(hIcon);
    CloseHandle(pi.hProcess);
    CloseHandle(pi.hThread);

    if(isTemp||isDefault){
        if(isDefault){
            if(MessageBox(g_hGui,L"Save changes to default profile?",APP_TITLE.c_str(),MB_YESNO|MB_ICONQUESTION)==IDYES){
                CleanupProfile(profilePath,aKeepActive);
                fs::path backup7z=g_sDataDir/"Default.7z";
                fs::path bakDir=g_sDataDir/"_DefBak";
                fs::create_directories(bakDir);
                // FIX: Use std::put_time for robust time formatting
                auto const now=std::chrono::system_clock::now();
                auto const in_time_t=std::chrono::system_clock::to_time_t(now);
                std::tm tm_buf;
                localtime_s(&tm_buf,&in_time_t);
                std::wostringstream ss;
                ss<<std::put_time(&tm_buf,L"%Y.%m.%d,%H%M%S");
                std::wstring timestamp=ss.str();
                std::wstring bakPath=bakDir/(timestamp+L"-Default.7z");
                fs::rename(backup7z,bakPath);
                std::wstring pathToArchive=profilePath.wstring()+L"\\*";
                std::wstring sevenZipCmd=std::format(L"a -mx=9 \"{}\" \"{}\"",backup7z.c_str(),pathToArchive.c_str());
                RunCommand(sevenZipCmd,g_sDataDir);
            }
        }
        try{ fs::remove_all(profilePath); } catch(...){}
    } else{
        CleanupProfile(profilePath,aKeepActive);
    }
    PostMessage(g_hGui,WM_APP_TASK_COMPLETE,0,0);
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
    // Trim leading/trailing whitespace
    const std::wstring whitespace=L" \t\n\r\f\v";
    sanitized.erase(0,sanitized.find_first_not_of(whitespace));
    sanitized.erase(sanitized.find_last_not_of(whitespace)+1);

    // FIX 1: Use a raw string literal to avoid complex backslash escaping.
    // The pattern LR"([\\/:*?"<>|])" safely tells the regex engine to match
    // any character inside the brackets, including a literal backslash `\`.
    std::wregex invalidChars(LR"([\\/:*?"<>|])");
    sanitized=std::regex_replace(sanitized,invalidChars,L"");

    // Remove trailing dots
    while(!sanitized.empty()&&sanitized.back()==L'.'){
        sanitized.pop_back();
    }

    // FIX 2: Use the standard C++ flag for case-insensitivity instead of "(?i:...)".
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
        // Ignore errors from iterating, e.g. if a file is locked
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
    // Determine the required size of the wide-character string
    int size_needed=MultiByteToWideChar(CP_ACP,0,str.c_str(),-1,NULL,0);
    if(size_needed==0){
        // Handle error if needed, for now return an empty string
        return std::wstring();
    }
    std::wstring wstrTo(size_needed,0);
    // Perform the conversion
    MultiByteToWideChar(CP_ACP,0,str.c_str(),-1,&wstrTo[0],size_needed);
    // The size includes the null terminator, which std::wstring doesn't need to store explicitly
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
    // --- Step 1: Load the source image ---
    std::unique_ptr<Gdiplus::Bitmap> sourceBitmap(Gdiplus::Bitmap::FromFile(sourceImagePath.c_str()));
    if(!sourceBitmap||sourceBitmap->GetLastStatus()!=Gdiplus::Ok){
        return false;
    }

    // --- Step 2: Create a 1:1 master bitmap, scaled using the script's exact logic ---
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

    // --- Step 3: Generate each icon size by downscaling from the master bitmap ---
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

    // --- Step 4: Call our faithful port to write the icon file ---
    bool result=SaveIconsToFile(destIconPath,hIcons);

    // --- Step 5: Final Cleanup ---
    // The HICONs must be destroyed after they have been written to the file.
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

    // First, determine the dimensions of the icon
    if(iconInfo.hbmColor){
        BITMAP bmp={0};
        if(GetObject(iconInfo.hbmColor,sizeof(BITMAP),&bmp)){
            width=bmp.bmWidth;
            height=bmp.bmHeight;
        }
    } else if(iconInfo.hbmMask){
        // Fallback for monochrome icons
        BITMAP bmp={0};
        if(GetObject(iconInfo.hbmMask,sizeof(BITMAP),&bmp)){
            width=bmp.bmWidth;
            height=bmp.bmHeight/2; // Mask contains both XOR and AND
        }
    }

    if(width==0||height==0){
        width=GetSystemMetrics(SM_CXICON);
        height=GetSystemMetrics(SM_CYICON);
    }

    // Check if it's already a good 32-bpp icon with alpha
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

    // If we reach here, we must create a new 32-bpp DIB.
    HDC hdcScreen=GetDC(NULL);
    HDC hdcMem=CreateCompatibleDC(hdcScreen);

    // Create the DIB Section
    BITMAPINFO bi={0};
    bi.bmiHeader.biSize=sizeof(BITMAPINFOHEADER);
    bi.bmiHeader.biWidth=width;
    bi.bmiHeader.biHeight=-height; // Top-down DIB
    bi.bmiHeader.biPlanes=1;
    bi.bmiHeader.biBitCount=32;
    bi.bmiHeader.biCompression=BI_RGB;
    void* pBits;
    HBITMAP hDib=CreateDIBSection(hdcScreen,&bi,DIB_RGB_COLORS,&pBits,NULL,0);

    HBITMAP hOldBmp=(HBITMAP)SelectObject(hdcMem,hDib);

    // Clear the new DIB with a transparent background
    RECT rc={0, 0, width, height};
    HBRUSH hBrush=CreateSolidBrush(RGB(0,0,0)); // A dummy brush
    FillRect(hdcMem,&rc,hBrush);
    DeleteObject(hBrush);

    // Draw the original icon onto our new DIB
    DrawIconEx(hdcMem,0,0,hIcon,width,height,0,NULL,DI_NORMAL);

    SelectObject(hdcMem,hOldBmp);
    DeleteDC(hdcMem);
    ReleaseDC(NULL,hdcScreen);

    // Create the final icon from our new 32-bpp DIB
    // For a 32bpp DIB, the mask is not needed, so we create a blank one.
    HBITMAP hMask=CreateBitmap(width,height,1,1,NULL);
    ICONINFO newIconInfo={0};
    newIconInfo.fIcon=TRUE;
    newIconInfo.hbmColor=hDib;
    newIconInfo.hbmMask=hMask;

    HICON hNewIcon=CreateIconIndirect(&newIconInfo);

    // Cleanup
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

    // This function only makes sense for 32-bpp bitmaps.
    if(bmp.bmBitsPixel!=32){
        return false;
    }

    // Get the raw pixel data
    int dataSize=bmp.bmWidthBytes*bmp.bmHeight;
    auto pixelData=std::make_unique<BYTE[]>(dataSize);
    if(GetBitmapBits(hBitmap,dataSize,pixelData.get())==0){
        return false;
    }

    // Iterate through each pixel. A 32-bpp pixel is 4 bytes (BGRA).
    for(int i=0; i<dataSize; i+=4){
        // The 4th byte is the alpha channel.
        if(pixelData[i+3]<255){
            return true; // Found a semi-transparent pixel
        }
    }

    return false; // All pixels are fully opaque
}
std::vector<BYTE> CompressBitmapToPng(HBITMAP hBitmap){
    // Create a GDI+ Bitmap object from the HBITMAP
    std::unique_ptr<Gdiplus::Bitmap> bitmap(Gdiplus::Bitmap::FromHBITMAP(hBitmap,NULL));
    if(!bitmap||bitmap->GetLastStatus()!=Gdiplus::Ok){
        return {};
    }

    // Get the CLSID for the PNG encoder
    CLSID pngClsid=GetEncoderClsid(L"image/png");

    // Create an in-memory stream using the Component Object Model (COM)
    IStream* pStream=NULL;
    if(CreateStreamOnHGlobal(NULL,TRUE,&pStream)!=S_OK){
        return {};
    }

    // Save the GDI+ bitmap to the stream as a PNG
    if(bitmap->Save(pStream,&pngClsid,NULL)!=Gdiplus::Ok){
        pStream->Release();
        return {};
    }

    // Get the size of the stream
    ULARGE_INTEGER streamSize;
    pStream->Seek({},STREAM_SEEK_END,&streamSize);

    // Allocate a buffer to hold the PNG data
    std::vector<BYTE> buffer(streamSize.QuadPart);

    // Rewind the stream and read the data into our buffer
    pStream->Seek({},STREAM_SEEK_SET,NULL);
    ULONG bytesRead;
    pStream->Read(buffer.data(),buffer.size(),&bytesRead);

    // Clean up the COM stream
    pStream->Release();

    if(bytesRead!=buffer.size()){
        return {}; // Read error
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
            // ** THE FIX **: Get the bitmap bits directly. DO NOT reverse the rows.
            // The positive biHeight in the header will tell the reader it's a bottom-up DIB.
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
            // Write the color and mask data in their original, bottom-up order.
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