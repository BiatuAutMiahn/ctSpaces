// Client7z.cpp
// - 7z-only
// - No encryption/password support
// - Adds:
//     HRESULT CompressDirectory7z(const FString& archivePath, const FString& folderPath);
//     HRESULT ExtractArchive7z(const FString& archivePath, const FString& outDir);

#include "../../3p/7zip/CPP/7zip/UI/Client7z/StdAfx.h"

#include <stdio.h>

#include "../../3p/7zip/CPP/Common/MyWindows.h"
#include "../../3p/7zip/CPP/Common/MyInitGuid.h"

#include "../../3p/7zip/CPP/Common/Defs.h"
#include "../../3p/7zip/CPP/Common/IntToString.h"
#include "../../3p/7zip/CPP/Common/StringConvert.h"

#include "../../3p/7zip/CPP/Windows/FileDir.h"
#include "../../3p/7zip/CPP/Windows/FileFind.h"
#include "../../3p/7zip/CPP/Windows/FileName.h"
#include "../../3p/7zip/CPP/Windows/NtCheck.h"
#include "../../3p/7zip/CPP/Windows/PropVariant.h"
#include "../../3p/7zip/CPP/Windows/PropVariantConv.h"

#include "../../3p/7zip/CPP/7zip/Common/FileStreams.h"
#include "../../3p/7zip/CPP/7zip/Archive/IArchive.h"

#include "../../3p/7zip/CPP/7zip/Common/CreateCoder.h" // keeps internal codecs available
#include "../../3p/7zip/CPP/7zip/IPassword.h"

#include "../../3p/7zip/C/7zVersion.h"

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

static void Convert_UString_to_AString(const UString& s,AString& temp){
    const int codePage=CP_OEMCP;
    UnicodeStringToMultiByte2(temp,s,(UINT)codePage);
}

static void Print(const char* s){ fputs(s,stdout); }
static void Print(const AString& s){ Print(s.Ptr()); }
static void Print(const UString& s){ AString as; Convert_UString_to_AString(s,as); Print(as); }
static void PrintNewLine(){ Print("\n"); }

static void PrintError(const char* message){
    Print("Error: ");
    Print(message);
    PrintNewLine();
}
static void PrintError(const char* message,const FString& name){
    PrintError(message);
    Print(name);
    PrintNewLine();
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

        // Archive internal path: rootNameInArchive\rel (or rootNameInArchive alone for root folder itself)
        UString arcPath=rootNameInArchive;
        if(!rel.IsEmpty()){
            arcPath.Add_PathSepar();
            arcPath+=FsToUs(rel);
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
    if(TotalBytes==0)
        return;

    const DWORD tick=::GetTickCount();
    const UInt32 percent=(UInt32)((CompletedBytes*100)/TotalBytes);
    if(!force){
        if(percent==LastPercent&&(tick-LastTick)<150)
            return;
    }
    LastPercent=percent;
    LastTick=tick;

    char head[32];
    sprintf(head,"[Extract %3u%%] ",(unsigned)percent);

    AString name;
    Convert_UString_to_AString(CurrentItem,name);

    AString line(head);
    line+=name;

    fprintf(stderr,"\r%s",line.Ptr());
    const unsigned curLen=(unsigned)line.Len();
    if(curLen<LastLineLen){
        for(unsigned i=0; i<(LastLineLen-curLen); ++i) fputc(' ',stderr);
    }
    LastLineLen=curLen;
    fflush(stderr);
}


void CArchiveExtractCallback::EndProgressLine(){
    if(TotalBytes==0)
        return;
    fprintf(stderr,"\n");
    fflush(stderr);
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
    if(TotalBytes==0)
        return;

    const DWORD tick=::GetTickCount();
    const UInt32 percent=(UInt32)((CompletedBytes*100)/TotalBytes);
    if(!force){
        if(percent==LastPercent&&(tick-LastTick)<150)
            return;
    }
    LastPercent=percent;
    LastTick=tick;

    char head[32];
    sprintf(head,"[Compress %3u%%] ",(unsigned)percent);

    AString name;
    Convert_UString_to_AString(CurrentItem,name);

    AString line(head);
    line+=name;

    fprintf(stderr,"\r%s",line.Ptr());
    const unsigned curLen=(unsigned)line.Len();
    if(curLen<LastLineLen){
        for(unsigned i=0; i<(LastLineLen-curLen); ++i) fputc(' ',stderr);
    }
    LastLineLen=curLen;
    fflush(stderr);
}


void CArchiveUpdateCallback::EndProgressLine(){
    if(TotalBytes==0)
        return;
    fprintf(stderr,"\n");
    fflush(stderr);
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
HRESULT CompressDirectory7z(const FString& archivePath,const FString& folderPath){
    // Normalize folderPath: must exist and be directory
    NFind::CFileInfo fi;
    if(!fi.Find(folderPath))
        return HRESULT_FROM_WIN32(ERROR_PATH_NOT_FOUND);
    if(!fi.IsDir())
        return HRESULT_FROM_WIN32(ERROR_DIRECTORY);

    // Root name in archive should match `7z a Archive.7z .\Folder` behavior:
    // put the folder name as the top-level entry.
    FString rootNameFs=BasenameOfPath(folderPath);
    UString rootName=FsToUs(rootNameFs);

    CObjectVector<CDirItem> items;

    // Add the root directory entry itself (optional but helps preserve empty dirs)
    {
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
    if(cbSpec->FailedFiles.Size()||cbSpec->NumErrors)
        fprintf(stderr,"WARNING: compression completed with skips/errors (files skipped: %u, opErrors: %llu)\n",
                (unsigned)cbSpec->FailedFiles.Size(),
                (unsigned long long)cbSpec->NumErrors);

    if(cbSpec->FailedFiles.Size()!=0){
        PrintError("WARNING: some files were skipped during compression");
    }


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


    {
        UInt32 numItems=0;
        inArchive->GetNumberOfItems(&numItems);
        if(numItems==0)
            fprintf(stderr,"WARNING: archive opened but has 0 items\n");
    }


    // Ensure output directory exists
    CreateComplexDir(outDir);

    // Extract callback
    CArchiveExtractCallback* cbSpec=new CArchiveExtractCallback;
    CMyComPtr<IArchiveExtractCallback> cb(cbSpec);
    cbSpec->Init(inArchive,outDir);

    hr=inArchive->Extract(NULL,(UInt32)(Int32)(-1),false,cb);
    if(hr==S_FALSE) hr=S_OK;
    cbSpec->EndProgressLine();
    if(cbSpec->NumErrors)
        fprintf(stderr,"WARNING: extraction completed with %llu skips/errors\n",
                (unsigned long long)cbSpec->NumErrors);

    return hr;
}

// -------------------------
// Optional test main
// -------------------------
int Z7_CDECL main(int numArgs,const char* args[]){
    NT_CHECK

        if(numArgs<2){
            Print("Usage:\n");
            Print("  7zcl.exe a <Archive.7z> <FolderPath>\n");
            Print("  7zcl.exe x <Archive.7z> <OutDir>\n");
            return 0;
        }

    AString cmd(args[1]);
    cmd.MakeLower_Ascii();

    if(cmd=="a"){
        if(numArgs!=4){ PrintError("Bad args for add"); return 1; }
        const FString arc=us2fs(GetUnicodeString(args[2]));
        const FString dir=us2fs(GetUnicodeString(args[3]));
        HRESULT hr=CompressDirectory7z(arc,dir);
        if(FAILED(hr)){ PrintError("CompressDirectory7z failed"); return 2; }
        return 0;
    }

    if(cmd=="x"){
        if(numArgs!=4){ PrintError("Bad args for extract"); return 1; }
        const FString arc=us2fs(GetUnicodeString(args[2]));
        const FString out=us2fs(GetUnicodeString(args[3]));
        HRESULT hr=ExtractArchive7z(arc,out);
        if(FAILED(hr)){PrintError("ExtractArchive7z failed");return 2; }
        return 0;
    }

    PrintError("Unknown command");
    return 1;
}
