cmptst_harness (static 7-Zip SDK smoke test)

What it does
- Compresses ./cmptst -> ./test.7z with mx=9
- Extracts ./test.7z -> ./cmptst_ext overwriting files

Where to put it
- Put this folder at:   .\7zF\cmptst_harness\
- Your 7-Zip sources at: .\3p\7zip\   (headers under .\3p\7zip\CPP\...)
- Your static lib project at: .\7zF\Format7z.vcxproj

Build
- Open cmptst_harness.sln
- Build: Release|x64

Run
- Set Working Directory to repo root (the folder that contains "cmptst")
- Run cmptst_harness.exe

Notes
- This calls the in-process 7-Zip SDK interfaces from the static library (no 7za/7z exe involved).
- If you want to hard-bind to CLSID_CFormat7z instead of enumerating handlers, that’s easy, but enumeration keeps it robust.
