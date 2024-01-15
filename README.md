# PreviewHandler
IPreviewHandler Sample Project

![image](https://github.com/fafalone/PreviewHandler/assets/7834493/394d0a46-8195-41d0-a5fc-d045c532d78b)

**Bringing my IPreviewHandler Sample Project to twinBASIC**

**Project Update - 15 Jan 2024
Version 2 
- Corrects infinite loop when both a local and inproc server handler has the same bug
- Corrects local server being dpi aware when this app is not.
- Better error handling

    
I'm writing this one up as an example of handling a semi-rough import of a project and updating it to x64.\
Files:
PreviewHandlerDemo.twinproj - The immediate VB6 import, no changes besides importing mIID.bas since the relative path was wrong, and re-checking the oleexp.tlb reference.\
PreviewHanderDemo64.twinproj - The final x64 conversion, with a few minor updates.

1) I started from the current public download:

[[VB6] IPreviewHandler: Show non-image file previews from all reg'd preview handlers](http://www.vbforums.com/showthread.php?802107-VB6-IPreviewHandler-Show-non-image-file-previews-from-any-reg-d-preview-handler)

2) I got some errors on load. The path a chose didn't have the same relative path to mIID.bas, so I had to manually add that. The oleexp.tlb reference had the same issue. Both of these can happen in VB6 though, so no big deal so far. Just add the file and fix the reference.

3) At this point it's ready to run. And it works! As long as tB is in 32bit mode, I'm happy to report this project works on initial import without any changes besides repairing references/file paths, both common in VB6 with downloaded projects.

4) It's time to update to x64. WinDevLib is the successor to oleexp.tlb; there's no 64bit oleexp, so WinDevLib is where the interfaces are. We remove the oleexp reference and add the WinDevLib (Windows Development Library for twinBASIC) package.

5) You'll see a lot of red at this point-- but it's pretty simple to clean up. The ReadMe in the WinDevLib repository has [a guide to switching from oleexp](https://github.com/fafalone/WinDevLib#guide-to-switching-from-oleexptlb).

6) Following those steps, we delete the references to oleexp-- `oleexp.SIZE` becomes just `SIZE`, and `oleexp.IUnknown`, as the readme notes, becomes **`IUnknownUnrestricted`**, which is better than the conflicted same name.

7) That leaves us with a single error: 'Too many arguments' in a call to IShellItemImageFactory.GetImage. This was a change made for 64bit compatibility. The first argument for GetImage expects a `ByVal SIZE`, but tB does not currently support this. In oleexp, only thinking of 32bit, the solution was to split the arguments into two, `ByVal cx As Long, ByVal cy As Long`. However, this doesn't work with the 64bit calling convention-- instead of bytes on stack, it uses registers here, so two arguments would occupy two 8-bit registers, instead of a contiguous 8 bytes. So for now, we use `ByVal LongLong`. Now, I could done separate definitions and left 32bit alone, but I thought it was better to have a single approach that worked for both, that way no conditional compilation is needed. So we make the following change (it can be condensed; this version is for maximum clarity):

```vba
Dim cxy As LongLong
Dim pSZ As SIZE
pSZ.cx = CX: pSZ.cy    = CY
CopyMemory cxy, pSZ, LenB(Of SIZE)
isiif.GetImage cxy, SIIGBF_THUMBNAILONLY, GetFileThumbnail
```

8) That gets rid of the last error. So we can now begin the x64 conversion in earnest. This project, like many, has a bunch of API declares. But unlike oleexp, WinDevLib also covers APIs. So while we could go through hundreds of lines of declares and manually adjust them for x64, what I like to do is let WinDevLib handle it instead. So just remove all the API defs: Lines 4-163 in Module1.bas. WinDevLib covers IIDs too, so delete mIID.bas. It's ok to delete the empty folders mIID.bas was in "PARENT FOLDER"/"tl_ole".

9) This produces some errors. WinDevLib documentation covers some of these: For `CreateFile` and `ReadFile`, instead of `ByVal 0` to an `As Any` argument, we pass `vbNullPtr` to a `SECURITY_ATTTRIBUTES` and `OVERLAPPED` argument, respectively. `CoCreateInstance` is just using a slightly different signature; insted of 0, we can pass `Nothing` to the `punkOuter` argument. The `GetObject` one... personally, I would classify this as a tB bug. Packages should be user-code that overrides built in packages without being higher in the priority list. But that's not how it currently works, so we're getting an error here because it's resolved as VB's GetObject, unrelated to the API call. Two options here, explicitly use `GetObjectW`, or a standard 'API' alias is defined, `GetObjectAPI`.

10) A few more thing to change with API calls-- in `BStrFromLPWStr` we change SysReAllocString to SysReAllocStringW because of argument types, then, and this is one is a little insidious: Like VB6, tB does implicitly Long->String conversions, so doesn't notify us of the looming problem here:

    ```vba
    Public Function GetHandlerCLSID(ByVal sExt As String, tID As UUID) As Long
    ...
    Call CLSIDFromString(StrPtr(sRes), tID)
    ```

    WinDevLib has a different signature for that API, especting a `String` instead of `LongPtr`. It won't error-- it will just return GUID_NULL, resulting in 'Class not registered' errors you'd first think were 64bit issues.

    Additionally, `COMDLG_FILTERSPEC` expects a `LongPtr` (`StrPtr()` here) in WinDevLib (my projects have been inconsistent about this, sorry).\
    Finally, `ShowPreviewForFile` has a similar lying-in-wait landmine: `IInitializeWithFile` expects `StrPtr(sFile)`, not a `String`. This changes relates to `Implements` compatibility. 


12) From there, all we have left is the variables. This is where it gets a little complicated and arcane, because whether a variable must be changed from `Long` to `LongPtr` depends on it's C/C++ definition in the Windows API. Information lost in the original VB6 conversion. You need to look at what the APIs and interfaces are expecting. Mainly, we change pointer types (like `PVOID, void*, ULONG_PTR, DWORD_PTR, WCHAR*`, all `PIDL` types, and `LPWSTR` where it's a `Long` instead of `String`), `HANDLE` types (like `HWND, HICON, HIMAGELIST`, or `HBITMAP`), and a few misc. ones (`WPARAM, LPARAM, size_t`). There's a long, long list of less common forms of those groups. In the project here, we have variables named hwnd, lp, lpsz, himl, pidl/pidlFQ, hbmp, hFile, lpGlobal; these are all used with APIs expecting those C/C++ types, so we change them to `LongPtr`.

13) Now, switch the compiler into 64bit mode. You'll likely find a few conversions you missed; the compiler catches a lot of these because it errors on 'Can't convert LongLong to Long'. 


And that's it! It will now run as x64 compatible!

### Project enhancements

There's some areas that really needed improvement. There's a catch-22 with some handlers. The Adobe PDF Previewer seems to work when created with `CLSCTX_INPROC_SERVER` right up until you do `DoPreview()`, where it fails with the error `E_FAIL`. But you can't just use a local server-- the Windows TXT File Previewer fails *silently* at `DoPreview()` unless you use InProc. So I used the dreaded `GoTo` to implement a restart; first it tries with `CLSCTX_LOCAL_SERVER`, and if it fails on `DoPreview()`, tries again with `CLSCTX_INPROC_SERVER`.


