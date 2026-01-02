#pragma once

#include <windows.h>

// In-process 7z pack/unpack wrappers (7z-only) based on 7zip\cmptst_harness.
// - No password / encryption support.

enum class _7zOp : unsigned {
    Extract = 1,
    Compress = 2,
};

// percent is 0..100. currentItem may be nullptr.
using _7zProgressCb = void (*)(void* user, _7zOp op, unsigned percent, const wchar_t* currentItem);

// Provide the application's HINSTANCE to 7-Zip runtime.
void _7zSetHInstance(HINSTANCE hInst);

// Extract a .7z archive to outDir.
HRESULT _7zExtra_7z(
    const wchar_t* archivePath,
    const wchar_t* outDir,
    _7zProgressCb progressCb,
    void* progressUser);

// Create a .7z archive from folderPath.
// If includeTopDirectory == false, the archive contains the folder's *contents* at the root
// (equivalent to: 7za a Archive.7z "folder\*" )
// If includeTopDirectory == true, the archive contains a top-level directory named after folderPath.
HRESULT _7zCompress7z(
    const wchar_t* archivePath,
    const wchar_t* folderPath,
    bool includeTopDirectory,
    _7zProgressCb progressCb,
    void* progressUser);
