#pragma comment(linker,"/manifestdependency:\"type='win32' name='Microsoft.Windows.Common-Controls' version='6.0.0.0' processorArchitecture='*' publicKeyToken='6595b64144ccf1df' language='*'\"")

#include "InProc7z.h"

#include "3p/7zip/CPP/7zip/UI/Client7z/StdAfx.h"


#include "3p/7zip/CPP/Common/MyWindows.h"
#include "3p/7zip/CPP/Common/MyInitGuid.h"

#include "3p/7zip/CPP/Common/Defs.h"
#include "3p/7zip/CPP/Common/IntToString.h"
#include "3p/7zip/CPP/Common/StringConvert.h"

#include "3p/7zip/CPP/Windows/FileDir.h"
#include "3p/7zip/CPP/Windows/FileFind.h"
#include "3p/7zip/CPP/Windows/FileName.h"
#include "3p/7zip/CPP/Windows/NtCheck.h"
#include "3p/7zip/CPP/Windows/PropVariant.h"
#include "3p/7zip/CPP/Windows/PropVariantConv.h"

#include "3p/7zip/CPP/7zip/Common/FileStreams.h"
#include "3p/7zip/CPP/7zip/Archive/IArchive.h"

#include "3p/7zip/CPP/7zip/Common/CreateCoder.h" // keeps internal codecs available
#include "3p/7zip/CPP/7zip/IPassword.h"

#include "3p/7zip/C/7zVersion.h"

#ifdef _WIN32
extern HINSTANCE g_hInstance;
HINSTANCE g_hInstance=NULL;
bool g_IsNT=true;
extern "C" {
    HRESULT WINAPI CreateArchiver(const GUID* clsid,const GUID* iid,void** outObject);
}
#endif

Z7_DIAGNOSTIC_IGNORE_CAST_FUNCTION
// 7z format GUID base: {23170F69-40C1-278A-1000-000110070000}
#define DEFINE_GUID_ARC(name, id) Z7_DEFINE_GUID(name, \
  0x23170F69, 0x40C1, 0x278A, 0x10, 0x00, 0x00, 0x01, 0x10, id, 0x00, 0x00);

enum{ kId_7z=7 };

extern "C" const GUID CLSID_CArchiveHandler=
{0x23170F69, 0x40C1, 0x278A, { 0x10, 0x00, 0x00, 0x01, 0x10, 0x00, 0x00, 0x00 }};

DEFINE_GUID_ARC(CLSID_Format,kId_7z)

using namespace NWindows;
using namespace NFile;
using namespace NDir;


// ----- ctSpaces progress bridge (thread-local) -----
static thread_local _7zProgressCb g__7zProgressCb = nullptr;
static thread_local void* g__7zProgressUser = nullptr;
static thread_local _7zOp g__7zProgressOp = _7zOp::Extract;

struct _7zScopedProgress {
    _7zProgressCb prevCb{};
    void* prevUser{};
    _7zOp prevOp{_7zOp::Extract};
    _7zScopedProgress(_7zOp op, _7zProgressCb cb, void* user){
        prevCb = g__7zProgressCb;
        prevUser = g__7zProgressUser;
        prevOp = g__7zProgressOp;
        g__7zProgressCb = cb;
        g__7zProgressUser = user;
        g__7zProgressOp = op;
    }
    ~_7zScopedProgress(){
        g__7zProgressCb = prevCb;
        g__7zProgressUser = prevUser;
        g__7zProgressOp = prevOp;
    }
};

static void Convert_UString_to_AString(const UString& s,AString& temp){
    const int codePage=CP_OEMCP;
    UnicodeStringToMultiByte2(temp,s,(UINT)codePage);
}

static void DebugPrintA(const char* s){ if(s) ::OutputDebugStringA(s); }
static void DebugPrintW(const wchar_t* s){ if(s) ::OutputDebugStringW(s); }
static void DebugPrintF(const FString& s){ if(s.Ptr()) ::OutputDebugStringW(s.Ptr()); }

static void PrintError(const char* message){
    DebugPrintA("7z: ");
    DebugPrintA(message ? message : "(null)");
    DebugPrintA("\n");
}
static void PrintError(const char* message,const FString& name){
    DebugPrintA("7z: ");
    DebugPrintA(message ? message : "(null)");
    DebugPrintA(" : ");
    DebugPrintF(name);
    DebugPrintA("\n");
}

static FString ArcPathToRelFsPath(const UString& arcPathIn){
    FString f=us2fs(arcPathIn);

#ifdef _WIN32
    for(unsigned i=0; i<f.Len(); ++i)
        if(f[i]==L'/')
            f.ReplaceOneCharAtPos(i,L'\\');

    // strip Win32 extended prefix: \\?\   <-- NOTE: do NOT end the comment with '\'
    if(f.Len()>=4&&
       f[0]==L'\\'&&f[1]==L'\\'&&f[2]==L'?'&&f[3]==L'\\'){
        f.Delete(0,4);
    }

    // remove ':' anywhere else (invalid in Windows file name)
    for (unsigned i = 0; i < f.Len(); ++i)
        if (f[i] == L':')
            f.ReplaceOneCharAtPos(i, L'_');

    // neuter ".." segments so we can't escape outDir
    for (unsigned i = 0; i + 1 < f.Len(); ++i) {
        if (f[i] == L'.' && f[i + 1] == L'.' &&
            (i == 0 || f[i - 1] == L'\\') &&
            (i + 2 == f.Len() || f[i + 2] == L'\\'))
        {
            f.ReplaceOneCharAtPos(i,   L'_');
            f.ReplaceOneCharAtPos(i+1, L'_');
        }
    }
#endif


    if(f.IsEmpty())
        f=FTEXT("_");

    return f;
}


static inline UString FsToUs(const FString& f){ return fs2us(f); }

// -------------------------
// Directory enumeration helpers (Win32)
// -------------------------
struct CDirItem{
    UString Path_For_Handler; // archive internal path (relative)
    FString FullPath;         // full disk path
    bool IsDir=false;

    UInt64 Size=0;
    DWORD Attrib=0;
    FILETIME CTime{};
    FILETIME ATime{};
    FILETIME MTime{};
};

static FString JoinPath(const FString& a,const FString& b){
    if(a.IsEmpty()) return b;
    if(b.IsEmpty()) return a;
    FString r=a;
    if(!r.IsEmpty()){
        wchar_t c=r.Back();
        if(c!=L'\\'&&c!=L'/')
            r.Add_PathSepar();
    }
    r+=b;
    return r;
}

static FString BasenameOfPath(const FString& p){
    // Use 7-Zip helper: reverse find separator and take tail
    int pos=p.ReverseFind_PathSepar();
    if(pos>=0) return p.Ptr(pos+1);
    return p;
}

static HRESULT EnumDirRecursive(
    const FString& rootDirFull,
    const FString& currentDirFull,
    const UString& rootNameInArchive,
    CObjectVector<CDirItem>& outItems){
#ifdef _WIN32
    // pattern: current\*
    FString pattern=currentDirFull;
    if(!pattern.IsEmpty()){
        wchar_t c=pattern.Back();
        if(c!=L'\\'&&c!=L'/')
            pattern.Add_PathSepar();
    }
    pattern+=FTEXT("*");

    WIN32_FIND_DATAW fd;
    HANDLE h=::FindFirstFileW(pattern.Ptr(),&fd);
    if(h==INVALID_HANDLE_VALUE){
        DWORD e=::GetLastError();
        return HRESULT_FROM_WIN32(e);
    }

    const auto CloseFind=[&](){ ::FindClose(h); };

    do{
        const wchar_t* name=fd.cFileName;
        if(!name||!name[0]) continue;
        if(wcscmp(name,L".")==0||wcscmp(name,L"..")==0) continue;

        const bool isDir=(fd.dwFileAttributes&FILE_ATTRIBUTE_DIRECTORY)!=0;

        FString childFull=currentDirFull;
        if(!childFull.IsEmpty()){
            wchar_t c=childFull.Back();
            if(c!=L'\\'&&c!=L'/')
                childFull.Add_PathSepar();
        }
        childFull+=name;

        // Compute relative path from rootDirFull to childFull
        FString rel=childFull;
        if(rel.Len()>=rootDirFull.Len()){
            // Strip rootDirFull prefix (case-insensitive on Windows is typical, but we assume exact form you passed in)
            rel.DeleteFrontal(rootDirFull.Len());
            // strip leading separators
            while(!rel.IsEmpty()&&(rel[0]==L'\\'||rel[0]==L'/'))
                rel.Delete(0);
        }

        // Archive internal path: rootNameInArchive\rel, or just rel if rootNameInArchive is empty
        UString arcPath;
        if(!rootNameInArchive.IsEmpty()){
            arcPath=rootNameInArchive;
            if(!rel.IsEmpty()){
                arcPath.Add_PathSepar();
                arcPath+=FsToUs(rel);
            }
        } else {
            arcPath=FsToUs(rel);
        }

        CDirItem item;
        item.Path_For_Handler=arcPath;
        item.FullPath=childFull;
        item.IsDir=isDir;
        item.Attrib=fd.dwFileAttributes;
        item.CTime=fd.ftCreationTime;
        item.ATime=fd.ftLastAccessTime;
        item.MTime=fd.ftLastWriteTime;
        item.Size=((UInt64)fd.nFileSizeHigh<<32)|(UInt64)fd.nFileSizeLow;

        outItems.Add(item);

        if(isDir){
            HRESULT hr=EnumDirRecursive(rootDirFull,childFull,rootNameInArchive,outItems);
            if(FAILED(hr)){ CloseFind(); return hr; }
        }

    } while(::FindNextFileW(h,&fd));

    const DWORD last=::GetLastError();
    CloseFind();
    if(last!=ERROR_NO_MORE_FILES){
        return HRESULT_FROM_WIN32(last);
    }
    return S_OK;
#else
    (void)rootDirFull; (void)currentDirFull; (void)rootNameInArchive; (void)outItems;
    return E_NOTIMPL;
#endif
}

// -------------------------
// Extract callback (no password)
// -------------------------
static const wchar_t* const kEmptyFileAlias=L"[Content]";

class CArchiveExtractCallback Z7_final:
    public IArchiveExtractCallback,
    public CMyUnknownImp{
    Z7_IFACES_IMP_UNK_1(IArchiveExtractCallback)
        Z7_IFACE_COM7_IMP(IProgress)

        CMyComPtr<IInArchive> _archiveHandler;
    FString _outDir;        // normalized dir prefix (ends with \)
    UString _filePath;      // path inside archive
    FString _diskFilePath;  // full output path on disk
    bool _extractMode;

    struct CProcessedFileInfo{
        UInt32 Attrib;
        bool isDir;
        bool Attrib_Defined;
    } _processed;

    COutFileStream* _outFileStreamSpec;
    CMyComPtr<ISequentialOutStream> _outFileStream;

public:
    void Init(IInArchive* archiveHandler,const FString& outDir){
        NumErrors=0;
        TotalBytes=0;
        CompletedBytes=0;
        LastPercent=(UInt32)(Int32)-1;
        LastTick=0;

        _archiveHandler=archiveHandler;
        _outDir=outDir;
        NName::NormalizeDirPathPrefix(_outDir);
        if(!_outDir.IsEmpty())
            CreateComplexDir(_outDir);
    }
    UInt64 NumErrors=0;

    void EndProgressLine();

private:
    UInt64 TotalBytes=0;
    UInt64 CompletedBytes=0;
    UInt32 LastPercent=(UInt32)(Int32)-1;
    DWORD LastTick=0;

    void PrintProgress(bool force);
    UString CurrentItem;
    unsigned LastLineLen=0;


};

Z7_COM7F_IMF(CArchiveExtractCallback::SetTotal(UInt64 size)){
    TotalBytes=size;
    CompletedBytes=0;
    LastPercent=(UInt32)(Int32)-1;
    LastTick=0;
    PrintProgress(true);
    return S_OK;
}

Z7_COM7F_IMF(CArchiveExtractCallback::SetCompleted(const UInt64* completeValue)){
    if(completeValue)
        CompletedBytes=*completeValue;
    PrintProgress(false);
    return S_OK;
}

void CArchiveExtractCallback::PrintProgress(bool force){
    if(!g__7zProgressCb) return;
    if(TotalBytes==0){
        if(force){
            g__7zProgressCb(g__7zProgressUser, g__7zProgressOp, 0, CurrentItem.Ptr());
        }
        return;
    }

    const DWORD tick=::GetTickCount();
    UInt32 percent=(UInt32)((CompletedBytes*100)/TotalBytes);
    if(percent>100) percent=100;
    if(!force){
        if(percent==LastPercent&&(tick-LastTick)<150)
            return;
    }
    LastPercent=percent;
    LastTick=tick;

    g__7zProgressCb(g__7zProgressUser, g__7zProgressOp, (unsigned)percent, CurrentItem.Ptr());
}


void CArchiveExtractCallback::EndProgressLine(){
    TotalBytes=0;
    LastLineLen=0;
}


Z7_COM7F_IMF(CArchiveExtractCallback::GetStream(UInt32 index,ISequentialOutStream** outStream,Int32 askExtractMode)){
    *outStream=NULL;
    _outFileStream.Release();

    // Path
    {
        NCOM::CPropVariant prop;
        RINOK(_archiveHandler->GetProperty(index,kpidPath,&prop))
            if(prop.vt==VT_EMPTY)
                _filePath=kEmptyFileAlias;
            else{
                if(prop.vt!=VT_BSTR){
                    NumErrors++;
                    PrintError("WARNING: bad kpidPath type (skipping)");
                    return S_FALSE; // skip this item, continue
                }
                _filePath=prop.bstrVal;
            }
    }
    CurrentItem=_filePath;

    if(askExtractMode!=NArchive::NExtract::NAskMode::kExtract)
        return S_OK;

    // Attrib
    {
        NCOM::CPropVariant prop;
        RINOK(_archiveHandler->GetProperty(index,kpidAttrib,&prop))

            if(prop.vt==VT_EMPTY){
                _processed.Attrib=0;
                _processed.Attrib_Defined=false;
            } else if(prop.vt==VT_UI4){
                _processed.Attrib=prop.ulVal;
                _processed.Attrib_Defined=true;
            } else{
                // was: return E_FAIL;
                _processed.Attrib=0;
                _processed.Attrib_Defined=false;
                NumErrors++;
                PrintError("WARNING: unexpected kpidAttrib type (continuing)");
            }
    }


    // IsDir
    {
        NCOM::CPropVariant prop;
        RINOK(_archiveHandler->GetProperty(index,kpidIsDir,&prop))

            if(prop.vt==VT_BOOL){
                _processed.isDir=VARIANT_BOOLToBool(prop.boolVal);
            } else if(prop.vt==VT_EMPTY){
                _processed.isDir=false;
            } else{
                // was: return E_FAIL;
                _processed.isDir=false;
                NumErrors++;
                PrintError("WARNING: unexpected kpidIsDir type (continuing)");
            }
    }


    // Ensure parent dirs exist
    {
        int slashPos=_filePath.ReverseFind_PathSepar();
        if(slashPos>=0){
            const FString parent=_outDir+ArcPathToRelFsPath(_filePath.Left(slashPos));
            if(!CreateComplexDir(parent)){
                NumErrors++;
                PrintError("WARNING: can't create directory (skipping)",parent);
                return S_FALSE;
            }
        }
    }

    _diskFilePath=_outDir+ArcPathToRelFsPath(_filePath);

    if(_processed.isDir){
        if(!CreateComplexDir(_diskFilePath)){
            NumErrors++;
            PrintError("WARNING: can't create directory",_diskFilePath);
        }
        return S_OK;
    }

    // Replace if exists
    {
        NFind::CFileInfo fi;
        if(fi.Find(_diskFilePath)){
            if(!DeleteFileAlways(_diskFilePath)){
                NumErrors++;
                PrintError("WARNING: cannot delete output file (skipping)",_diskFilePath);
                return S_FALSE;
            }
        }
    }

    _outFileStreamSpec=new COutFileStream;
    CMyComPtr<ISequentialOutStream> outLoc(_outFileStreamSpec);
    if(!_outFileStreamSpec->Create_ALWAYS(_diskFilePath)){
        NumErrors++;
        PrintError("WARNING: cannot open output file (skipping)",_diskFilePath);
        return S_FALSE;
    }

    _outFileStream=outLoc;
    *outStream=outLoc.Detach();
    return S_OK;
}

Z7_COM7F_IMF(CArchiveExtractCallback::PrepareOperation(Int32 askExtractMode)){
    _extractMode=(askExtractMode==NArchive::NExtract::NAskMode::kExtract);
    return S_OK;
}

Z7_COM7F_IMF(CArchiveExtractCallback::SetOperationResult(Int32 operationResult)){
    if(operationResult!=NArchive::NExtract::NOperationResult::kOK){
        NumErrors++;
        PrintError("Extract error (operationResult != OK)");
    }

    if(_outFileStream){
        const HRESULT hrClose=_outFileStreamSpec->Close();
        if(hrClose!=S_OK){
            NumErrors++;
            PrintError("WARNING: cannot close output file",_diskFilePath);
        }
    }
    _outFileStream.Release();

    if(_extractMode&&_processed.Attrib_Defined){
        SetFileAttrib_PosixHighDetect(_diskFilePath,_processed.Attrib);
    }
    return S_OK;
}


// -------------------------
// Update callback (no password)
// -------------------------
class CArchiveUpdateCallback Z7_final:
    public IArchiveUpdateCallback2,
    public CMyUnknownImp{
    Z7_IFACES_IMP_UNK_1(IArchiveUpdateCallback2)
        Z7_IFACE_COM7_IMP(IProgress)
        Z7_IFACE_COM7_IMP(IArchiveUpdateCallback)

public:
    const CObjectVector<CDirItem>* Items=NULL;

    void Init(const CObjectVector<CDirItem>* items){
        Items=items;
        FailedFiles.Clear();
        FailedCodes.Clear();
        TotalBytes=0;
        CompletedBytes=0;
        LastPercent=(UInt32)(Int32)-1;
        LastTick=0;
    }
    FStringVector FailedFiles;
    CRecordVector<HRESULT> FailedCodes;

    UInt64 TotalBytes=0;
    UInt64 CompletedBytes=0;
    UInt32 LastPercent=(UInt32)(Int32)-1;
    DWORD LastTick=0;

    void EndProgressLine();

    UInt64 NumErrors=0;
private:
    void PrintProgress(bool force);
    UString CurrentItem;
    unsigned LastLineLen=0;







    // volumes unused
};

Z7_COM7F_IMF(CArchiveUpdateCallback::SetTotal(UInt64 size)){
    TotalBytes=size;
    CompletedBytes=0;
    LastPercent=(UInt32)(Int32)-1;
    LastTick=0;
    PrintProgress(true);
    return S_OK;
}

Z7_COM7F_IMF(CArchiveUpdateCallback::SetCompleted(const UInt64* completeValue)){
    if(completeValue)
        CompletedBytes=*completeValue;
    PrintProgress(false);
    return S_OK;
}

void CArchiveUpdateCallback::PrintProgress(bool force){
    if(!g__7zProgressCb) return;
    if(TotalBytes==0){
        if(force){
            g__7zProgressCb(g__7zProgressUser, g__7zProgressOp, 0, CurrentItem.Ptr());
        }
        return;
    }

    const DWORD tick=::GetTickCount();
    UInt32 percent=(UInt32)((CompletedBytes*100)/TotalBytes);
    if(percent>100) percent=100;
    if(!force){
        if(percent==LastPercent&&(tick-LastTick)<150)
            return;
    }
    LastPercent=percent;
    LastTick=tick;

    g__7zProgressCb(g__7zProgressUser, g__7zProgressOp, (unsigned)percent, CurrentItem.Ptr());
}


void CArchiveUpdateCallback::EndProgressLine(){
    TotalBytes=0;
    LastLineLen=0;
}

Z7_COM7F_IMF(CArchiveUpdateCallback::GetUpdateItemInfo(UInt32 /*index*/,Int32* newData,Int32* newProps,UInt32* indexInArchive)){
    if(newData) *newData=BoolToInt(true);
    if(newProps) *newProps=BoolToInt(true);
    if(indexInArchive) *indexInArchive=(UInt32)(Int32)-1;
    return S_OK;
}

Z7_COM7F_IMF(CArchiveUpdateCallback::GetProperty(UInt32 index,PROPID propID,PROPVARIANT* value)){
    NCOM::CPropVariant prop;
    const CDirItem& it=(*Items)[index];

    switch(propID){
    case kpidPath:   prop=it.Path_For_Handler; break;
    case kpidIsDir:  prop=it.IsDir; break;
    case kpidSize:   prop=it.Size; break;
    case kpidAttrib: prop=(UInt32)it.Attrib; break;
    case kpidCTime:  PropVariant_SetFrom_FiTime(prop,it.CTime); break;
    case kpidATime:  PropVariant_SetFrom_FiTime(prop,it.ATime); break;
    case kpidMTime:  PropVariant_SetFrom_FiTime(prop,it.MTime); break;
    default: break;
    }

    prop.Detach(value);
    return S_OK;
}

Z7_COM7F_IMF(CArchiveUpdateCallback::GetStream(UInt32 index,ISequentialInStream** inStream)){
    *inStream=NULL;
    const CDirItem& it=(*Items)[index];
    CurrentItem=it.Path_For_Handler;

    if(it.IsDir)
        return S_OK;

    CInFileStream* inSpec=new CInFileStream;
    CMyComPtr<ISequentialInStream> inLoc(inSpec);

#ifdef _WIN32
    if(!inSpec->OpenShared(it.FullPath,true))  // allow reading even if someone else has it open
#else
    if(!inSpec->Open(it.FullPath))
#endif
    {
#ifdef _WIN32
        const DWORD e=::GetLastError();
#else
        const DWORD e=::GetLastError();
#endif
        const HRESULT hr=HRESULT_FROM_WIN32(e);
        FailedFiles.Add(it.FullPath);
        FailedCodes.Add(hr);
        PrintError("WARNING: can't open file (skipping)",it.FullPath);
        return S_FALSE; // skip and continue
    }

    *inStream=inLoc.Detach();
    return S_OK;

}

Z7_COM7F_IMF(CArchiveUpdateCallback::SetOperationResult(Int32 operationResult)){
    if(operationResult!=0) // 0 == OK in 7-Zip update results
        NumErrors++;
    return S_OK; // keep going
}

Z7_COM7F_IMF(CArchiveUpdateCallback::GetVolumeSize(UInt32 /*index*/,UInt64* /*size*/)){ return S_FALSE; }

Z7_COM7F_IMF(CArchiveUpdateCallback::GetVolumeStream(UInt32 /*index*/,ISequentialOutStream** /*volumeStream*/)){ return S_FALSE; }


// -------------------------
// Property setup helper (your CLI equivalent)
// -------------------------
static HRESULT Apply7zUltraProps(IOutArchive* outArchive){
    // Mirrors these 7z.exe switches (for 7z format):
    //   -mx=9 -m0=lzma2 -md=256m -mfb=64 -ms=16g -mmt=on -mmemuse=p80
    //
    // Important detail: 7-Zip's console parses "-m<name>=<value>" and passes
    // just "<name>" into ISetProperties. For example "-m0=lzma2" becomes name "0".
    CMyComPtr<ISetProperties> setProps;
    outArchive->QueryInterface(IID_ISetProperties,(void**)&setProps);
    if(!setProps)
        return E_NOINTERFACE;

    const wchar_t* const names[]=
    {
      L"x",       // -mx=9
      L"0",       // -m0=lzma2  (method name for method slot 0)
      L"d",       // -md=256m   (dictionary)
      L"fb",      // -mfb=64
      L"s",       // -ms=16g    (solid block size)
      L"mt",      // -mmt=on
      L"memuse"   // -mmemuse=p80
    };

    const unsigned kNumProps=Z7_ARRAY_SIZE(names);
    NCOM::CPropVariant vals[kNumProps];
    unsigned n=0;

    vals[n++]=(UInt32)9;
    vals[n++]=L"lzma2";
    vals[n++]=L"256m";
    vals[n++]=(UInt32)64;
    vals[n++]=L"16g";
    vals[n++]=L"on";
    vals[n++]=L"p80";

    return setProps->SetProperties(names,vals,kNumProps);
}


// -------------------------
// Public functions you asked for
// -------------------------
static HRESULT CompressDirectory7zImpl(const FString& archivePath,const FString& folderPath,bool includeTopDirectory){
    // Normalize folderPath: must exist and be directory
    NFind::CFileInfo fi;
    if(!fi.Find(folderPath))
        return HRESULT_FROM_WIN32(ERROR_PATH_NOT_FOUND);
    if(!fi.IsDir())
        return HRESULT_FROM_WIN32(ERROR_DIRECTORY);

    CObjectVector<CDirItem> items;

    UString rootName;
    if(includeTopDirectory){
        // Include the folder name as the top-level entry (equivalent to: 7za a Archive.7z "folder")
        FString rootNameFs=BasenameOfPath(folderPath);
        rootName=FsToUs(rootNameFs);

        // Add the root directory entry itself (helps preserve an empty root dir)
        CDirItem root;
        root.Path_For_Handler=rootName;
        root.FullPath=folderPath;
        root.IsDir=true;
        root.Attrib=fi.Attrib;
        root.CTime=fi.CTime;
        root.ATime=fi.ATime;
        root.MTime=fi.MTime;
        root.Size=0;
        items.Add(root);
    } else {
        // Do NOT include the folder name in the archive; add only its contents at the root
        // (equivalent to: 7za a Archive.7z "folder\*")
        rootName.Empty();
    }

    // Enumerate children recursively
    HRESULT hr=EnumDirRecursive(folderPath,folderPath,rootName,items);
    if(FAILED(hr))
        return hr;

    // Create output archive stream
    COutFileStream* outSpec=new COutFileStream;
    CMyComPtr<IOutStream> outStream=outSpec;
    if(!outSpec->Create_NEW(archivePath))
        return HRESULT_FROM_WIN32(::GetLastError());

    // Create 7z out archive handler
    CMyComPtr<IOutArchive> outArchive;
    hr=CreateArchiver(&CLSID_Format,&IID_IOutArchive,(void**)&outArchive);
    if(FAILED(hr)||!outArchive)
        return FAILED(hr)?hr:E_FAIL;

    // Apply your compression settings
    hr=Apply7zUltraProps(outArchive);
    if(FAILED(hr))
        return hr;

    // Update callback
    CArchiveUpdateCallback* cbSpec=new CArchiveUpdateCallback;
    CMyComPtr<IArchiveUpdateCallback2> cb(cbSpec);
    cbSpec->Init(&items);

    hr=outArchive->UpdateItems(outStream,items.Size(),cb);
    if(hr==S_FALSE) hr=S_OK;
    cbSpec->EndProgressLine();
    return hr;
}

HRESULT ExtractArchive7z(const FString& archivePath,const FString& outDir){
    // Open archive file
    CInFileStream* inSpec=new CInFileStream;
    CMyComPtr<IInStream> inStream=inSpec;
    if(!inSpec->Open(archivePath))
        return HRESULT_FROM_WIN32(::GetLastError());

    // Create 7z in archive handler
    CMyComPtr<IInArchive> inArchive;
    HRESULT hr=CreateArchiver(&CLSID_Format,&IID_IInArchive,(void**)&inArchive);
    if(FAILED(hr)||!inArchive)
        return FAILED(hr)?hr:E_FAIL;

    // Open archive (no password callback)
    const UInt64 scanSize=(UInt64)1<<23;
    hr=inArchive->Open(inStream,&scanSize,NULL);
    if(FAILED(hr))
        return hr;


    // Ensure output directory exists
    CreateComplexDir(outDir);

    // Extract callback
    CArchiveExtractCallback* cbSpec=new CArchiveExtractCallback;
    CMyComPtr<IArchiveExtractCallback> cb(cbSpec);
    cbSpec->Init(inArchive,outDir);

    hr=inArchive->Extract(NULL,(UInt32)(Int32)(-1),false,cb);
    if(hr==S_FALSE) hr=S_OK;
    cbSpec->EndProgressLine();
    return hr;
}



// -------------------------
// Public API
// -------------------------

void _7zSetHInstance(HINSTANCE hInst){
#ifdef _WIN32
    g_hInstance = hInst;
#else
    (void)hInst;
#endif
}

HRESULT _7zExtra_7z(
    const wchar_t* archivePath,
    const wchar_t* outDir,
    _7zProgressCb progressCb,
    void* progressUser)
{
    if(!archivePath || !outDir) return E_INVALIDARG;
#ifdef _WIN32
    NT_CHECK
#endif
    _7zScopedProgress sp(_7zOp::Extract, progressCb, progressUser);
    const FString arc = us2fs(UString(archivePath));
    const FString out = us2fs(UString(outDir));
    HRESULT hr = ExtractArchive7z(arc, out);
    if(SUCCEEDED(hr) && progressCb){
        progressCb(progressUser, _7zOp::Extract, 100, nullptr);
    } 
    return hr;
}

HRESULT _7zCompress7z(
    const wchar_t* archivePath,
    const wchar_t* folderPath,
    bool includeTopDirectory,
    _7zProgressCb progressCb,
    void* progressUser)
{
    if(!archivePath || !folderPath) return E_INVALIDARG;
#ifdef _WIN32
    NT_CHECK
#endif
    _7zScopedProgress sp(_7zOp::Compress, progressCb, progressUser);
    const FString arc = us2fs(UString(archivePath));
    const FString dir = us2fs(UString(folderPath));
    HRESULT hr = CompressDirectory7zImpl(arc, dir, includeTopDirectory);
    if(SUCCEEDED(hr) && progressCb){
        progressCb(progressUser, _7zOp::Compress, 100, nullptr);
    }
    return hr;
}
