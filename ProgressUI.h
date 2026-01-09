#pragma once
#pragma once
#include <windows.h>
#include <string>

#ifndef WM_APP_PROGRESS_SHOW
#define WM_APP_PROGRESS_SHOW   (WM_APP + 200)
#define WM_APP_PROGRESS_UPDATE (WM_APP + 201)
#define WM_APP_PROGRESS_HIDE   (WM_APP + 202)
#endif

// Payload posted across threads via PostMessage (must be heap allocated)
struct CtProgressPayload{
    int percent;            // 0..100, or -1 for marquee
    std::wstring text;      // base message, e.g. "Applying default profile..."
    CtProgressPayload(int p,std::wstring t): percent(p),text(std::move(t)){}
};

void ProgressUI_Init(HINSTANCE hInst);
void ProgressUI_SetTheme(COLORREF crWindow,COLORREF crText,COLORREF crBar,COLORREF crBarBk,bool bDark);
void ProgressUI_Show(HWND owner,const std::wstring& baseText,int percent);
void ProgressUI_Update(int percent);          // uses last baseText
void ProgressUI_Update(const std::wstring& baseText,int percent);
void ProgressUI_Hide();

// Thread-safe helpers
void ProgressUI_PostShow(HWND mainWnd,const std::wstring& baseText,int percent);
void ProgressUI_PostUpdate(HWND mainWnd,int percent); // percent only
void ProgressUI_PostUpdate(HWND mainWnd,const std::wstring& baseText,int percent);
void ProgressUI_PostHide(HWND mainWnd);
