#include "ProgressUI.h"

#include <windows.h>
#include <CommCtrl.h>
#include <algorithm>
#include <string>
#include <uxtheme.h>

#pragma comment(lib, "UxTheme.lib")
#pragma comment(lib, "Comctl32.lib")

static HINSTANCE g_hInst=nullptr;

static HWND g_hWnd=nullptr;   // owned popup (has caption)
static HWND g_hBar=nullptr;   // progress bar control
static HWND g_hOwner=nullptr;   // main GUI window

static std::wstring g_baseText;
static int  g_lastPercent=0;
static bool g_marquee=false;
static UINT g_dpi=USER_DEFAULT_SCREEN_DPI;

static const wchar_t* kProgClass=L"ctSpaces.ProgressUI";

// Forward decl
static void LayoutControls();
static void RepositionInOwner(HWND owner,int desiredW,int desiredH);

static int ScaleByDpi(int value,UINT dpi){
    return MulDiv(value,dpi,USER_DEFAULT_SCREEN_DPI);
}

static std::wstring ComposeCaption(const std::wstring& baseText,int percent){
    if(baseText.empty())
        return L"ctSpaces";

    if(percent<0)
        return baseText;

    wchar_t buf[1024];
    _snwprintf_s(buf,_TRUNCATE,L"%s (%d%%)",baseText.c_str(),percent);
    return buf;
}

static LRESULT CALLBACK ProgressWndProc(HWND hWnd,UINT msg,WPARAM wParam,LPARAM lParam){
    switch(msg){
    case WM_ERASEBKGND:
    {
        RECT rc{};
        GetClientRect(hWnd,&rc);
        FillRect((HDC)wParam,&rc,GetSysColorBrush(COLOR_BTNFACE));
        return 1;
    }

    case WM_PAINT:
    {
        PAINTSTRUCT ps{};
        HDC hdc=BeginPaint(hWnd,&ps);

        RECT rc{};
        GetClientRect(hWnd,&rc);
        FillRect(hdc,&rc,GetSysColorBrush(COLOR_BTNFACE));

        EndPaint(hWnd,&ps);
        return 0;
    }

    case WM_SIZE:
        LayoutControls();
        return 0;

    case WM_DPICHANGED:
    {
        const UINT dpi=HIWORD(wParam);
        g_dpi=(dpi?dpi:USER_DEFAULT_SCREEN_DPI);
        const RECT* rc=reinterpret_cast<RECT*>(lParam);
        if(rc){
            SetWindowPos(hWnd,nullptr,
                         rc->left,rc->top,
                         rc->right-rc->left,
                         rc->bottom-rc->top,
                         SWP_NOZORDER|SWP_NOACTIVATE);
        }
        LayoutControls();
        return 0;
    }

    case WM_CLOSE:
        return 0; // ignore close; controlled by ProgressUI_Hide()

    case WM_NCDESTROY:
        g_hWnd=nullptr;
        g_hBar=nullptr;
        return 0;
    }

    return DefWindowProcW(hWnd,msg,wParam,lParam);
}

static void EnsureWindow(HWND owner){
    if(g_hWnd&&IsWindow(g_hWnd))
        return;

    g_hOwner=owner;

    // Make sure common controls are initialized (safe if already done)
    INITCOMMONCONTROLSEX icc{sizeof(icc), ICC_PROGRESS_CLASS};
    InitCommonControlsEx(&icc);

    static bool s_registered=false;
    if(!s_registered){
        WNDCLASSEXW wcex{};
        wcex.cbSize=sizeof(wcex);
        wcex.lpfnWndProc=ProgressWndProc;
        wcex.hInstance=g_hInst;
        wcex.hCursor=LoadCursor(nullptr,IDC_ARROW);
        wcex.hbrBackground=GetSysColorBrush(COLOR_BTNFACE);
        wcex.lpszClassName=kProgClass;
        RegisterClassExW(&wcex);
        s_registered=true;
    }

    // Owned popup (captioned). This is NOT a child window.
    // Owner keeps it above the main GUI and out of taskbar.
    DWORD exStyle=WS_EX_TOOLWINDOW;   // no taskbar/alt-tab entry
    DWORD style=WS_POPUP|WS_CAPTION; // caption + border, no system menu by default

    g_hWnd=CreateWindowExW(
        exStyle,
        kProgClass,
        L"",
        style,
        0,0,10,10,
        owner,          // OWNER (because WS_POPUP)
        nullptr,
        g_hInst,
        nullptr
    );
    g_dpi=GetDpiForWindow(g_hWnd);

    // If you want to be extra sure there's no Close:
    // - Ensure no system menu
    LONG_PTR s=GetWindowLongPtrW(g_hWnd,GWL_STYLE);
    s&=~WS_SYSMENU;
    SetWindowLongPtrW(g_hWnd,GWL_STYLE,s);

    // Also gray out SC_CLOSE if anything reintroduces it:
    if(HMENU hSys=GetSystemMenu(g_hWnd,FALSE))
        EnableMenuItem(hSys,SC_CLOSE,MF_BYCOMMAND|MF_GRAYED);

    g_hBar=CreateWindowExW(
        0,
        PROGRESS_CLASSW,
        L"",
        WS_CHILD|WS_VISIBLE,
        0,0,10,10,
        g_hWnd,
        nullptr,
        g_hInst,
        nullptr
    );

    SetWindowTheme(g_hBar,L"Explorer",nullptr);
    SendMessageW(g_hBar,PBM_SETBKCOLOR,0,GetSysColor(COLOR_WINDOW));
    SendMessageW(g_hBar,PBM_SETBARCOLOR,0,GetSysColor(COLOR_HIGHLIGHT));
    SendMessageW(g_hBar,PBM_SETRANGE,0,MAKELPARAM(0,100));
    SendMessageW(g_hBar,PBM_SETPOS,0,0);

    // Do NOT override fonts. Let it inherit/default (or your app sets it elsewhere).

    ShowWindow(g_hWnd,SW_HIDE);

    RepositionInOwner(owner,420,84);
    LayoutControls();
}

static void RepositionInOwner(HWND owner,int desiredW,int desiredH){
    if(!g_hWnd||!IsWindow(g_hWnd)||!owner||!IsWindow(owner))
        return;

    g_dpi=GetDpiForWindow(owner);

    RECT rcClient{};
    GetClientRect(owner,&rcClient);

    // Convert owner's client origin to SCREEN coordinates (critical for WS_POPUP)
    POINT origin{0, 0};
    ClientToScreen(owner,&origin);

    const int pad=ScaleByDpi(10,g_dpi);
    int ownerW=rcClient.right-rcClient.left;
    int ownerH=rcClient.bottom-rcClient.top;

    int availW=(std::max)(0,ownerW-pad*2);
    int availH=(std::max)(0,ownerH-pad*2);

    const int desiredWpx=ScaleByDpi(desiredW,g_dpi);
    const int desiredHpx=ScaleByDpi(desiredH,g_dpi);

    int w=(std::min)(desiredWpx,availW);
    int h=(std::min)(desiredHpx,availH);

    w=(std::max)(w,ScaleByDpi(260,g_dpi));
    h=(std::max)(h,ScaleByDpi(70,g_dpi));

    int x=origin.x+pad+(availW-w)/2;
    int y=origin.y+pad+(availH-h)/2;

    SetWindowPos(g_hWnd,HWND_TOP,x,y,w,h,SWP_NOACTIVATE|SWP_NOOWNERZORDER);
}

static void LayoutControls(){
    if(!g_hWnd||!g_hBar) return;

    RECT rc{};
    GetClientRect(g_hWnd,&rc);

    int w=rc.right-rc.left;
    int h=rc.bottom-rc.top;

    // Only a centered progress bar in the CLIENT area (below caption automatically)
    const int barH=ScaleByDpi(22,g_dpi);
    const int padX=ScaleByDpi(16,g_dpi);

    int barW=(std::max)(0,w-padX*2);
    int x=padX;
    int y=(h-barH)/2;

    MoveWindow(g_hBar,x,y,barW,barH,TRUE);
}

static void SetMarquee(bool enable){
    if(!g_hBar) return;

    LONG_PTR style=GetWindowLongPtrW(g_hBar,GWL_STYLE);
    if(enable) style|=PBS_MARQUEE;
    else        style&=~PBS_MARQUEE;

    SetWindowLongPtrW(g_hBar,GWL_STYLE,style);
    SetWindowPos(g_hBar,nullptr,0,0,0,0,
                 SWP_NOMOVE|SWP_NOSIZE|SWP_NOZORDER|SWP_FRAMECHANGED);

    SendMessageW(g_hBar,PBM_SETMARQUEE,enable?TRUE:FALSE,30);
    g_marquee=enable;
}

void ProgressUI_Init(HINSTANCE hInst){
    g_hInst=hInst;
}

void ProgressUI_Show(HWND owner,const std::wstring& baseText,int percent){
    EnsureWindow(owner);

    g_baseText=baseText;
    g_lastPercent=(percent<0)?0:(std::clamp)(percent,0,100);

    // Set *progress window* caption (this fixes "no text in the title")
    SetWindowTextW(g_hWnd,ComposeCaption(g_baseText,percent).c_str());

    if(percent<0){
        SetMarquee(true);
    } else{
        if(g_marquee) SetMarquee(false);
        SendMessageW(g_hBar,PBM_SETPOS,(WPARAM)g_lastPercent,0);
    }

    RepositionInOwner(owner,420,84);
    LayoutControls();

    ShowWindow(g_hWnd,SW_SHOWNOACTIVATE);
    InvalidateRect(g_hWnd,nullptr,TRUE);
    UpdateWindow(g_hWnd);
}

void ProgressUI_Update(int percent){
    ProgressUI_Update(g_baseText,percent);
}

void ProgressUI_Update(const std::wstring& baseText,int percent){
    if(!g_hWnd||!IsWindow(g_hWnd)) return;

    if(!baseText.empty())
        g_baseText=baseText;

    // Update caption every time
    SetWindowTextW(g_hWnd,ComposeCaption(g_baseText,percent).c_str());

    if(g_hOwner&&IsWindow(g_hOwner)){
        RepositionInOwner(g_hOwner,420,84);
        LayoutControls();
    }

    if(percent<0){
        if(!g_marquee) SetMarquee(true);
        return;
    }

    if(g_marquee) SetMarquee(false);

    g_lastPercent=(std::clamp)(percent,0,100);
    SendMessageW(g_hBar,PBM_SETPOS,(WPARAM)g_lastPercent,0);
}

void ProgressUI_Hide(){
    if(!g_hWnd||!IsWindow(g_hWnd)) return;
    ShowWindow(g_hWnd,SW_HIDE);
}

// ===== Thread-safe PostMessage wrappers =====

void ProgressUI_PostShow(HWND mainWnd,const std::wstring& baseText,int percent){
    auto* p=new CtProgressPayload(percent,baseText);
    PostMessageW(mainWnd,WM_APP_PROGRESS_SHOW,0,(LPARAM)p);
}

void ProgressUI_PostUpdate(HWND mainWnd,int percent){
    auto* p=new CtProgressPayload(percent,L"");
    PostMessageW(mainWnd,WM_APP_PROGRESS_UPDATE,0,(LPARAM)p);
}

void ProgressUI_PostUpdate(HWND mainWnd,const std::wstring& baseText,int percent){
    auto* p=new CtProgressPayload(percent,baseText);
    PostMessageW(mainWnd,WM_APP_PROGRESS_UPDATE,0,(LPARAM)p);
}

void ProgressUI_PostHide(HWND mainWnd){
    PostMessageW(mainWnd,WM_APP_PROGRESS_HIDE,0,0);
}
