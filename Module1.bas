Attribute VB_Name = "Module1"
Option Explicit
Public Const S_OK = 0&
Public Declare Function SHCreateItemFromParsingName Lib "shell32" (ByVal pszPath As Long, pbc As Any, riid As UUID, ppv As Any) As Long
Public Declare Function SHCreateShellItem Lib "shell32" (ByVal pidlParent As Long, ByVal psfParent As Long, ByVal pidl As Long, ppsi As IShellItem) As Long
Public Declare Function SysReAllocString Lib "oleaut32.dll" (ByVal pBSTR As Long, Optional ByVal pszStrPtr As Long) As Long
Public Declare Function CLSIDFromString Lib "ole32" (ByVal lpszGuid As Long, pGuid As Any) As Long
Public Declare Function CreateFile Lib "kernel32" Alias "CreateFileA" (ByVal lpFileName As String, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, ByVal lpSecurityAttributes As Any, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long
Public Const GENERIC_WRITE As Long = &H40000000
Public Const FILE_READ_DATA As Long = &H1
Public Const FILE_SHARE_READ As Long = &H1&
Public Const OPEN_ALWAYS As Long = 4&
Public Const OPEN_EXISTING As Long = 3&
Public Const CREATE_ALWAYS As Long = 2&
Public Const FILE_ATTRIBUTE_ARCHIVE As Long = &H20&
Public Const GENERIC_READ As Long = &H80000000
Public Const FILE_END As Long = 2&
Public Declare Function ReadFile Lib "kernel32" (ByVal hFile As Long, lpBuffer As Any, ByVal nNumberOfBytesToRead As Long, lpNumberOfBytesRead As Long, ByVal lpOverlapped As Any) As Long
Public Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Public Declare Function AssocGetPerceivedType Lib "shlwapi.dll" (ByVal pszExt As Long, ptype As PERCEIVED, pflag As PERCEIVEDFLAG, ppszType As Long) As Long
Public Declare Function SHCreateItemFromIDList Lib "shell32" (ByVal pidl As Long, riid As UUID, ppv As Any) As Long
Public Declare Function ILCreateFromPathW Lib "shell32" (ByVal pwszPath As Long) As Long
Public Declare Sub ILFree Lib "shell32" (ByVal pidl As Long)
Public Declare Function GetObject Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, lpObject As Any) As Long
Public Enum ImageTypes
  IMAGE_BITMAP = 0
  IMAGE_ICON = 1
  IMAGE_CURSOR = 2
  IMAGE_ENHMETAFILE = 3
End Enum
Public Enum LoadResourceFlags
  LR_DEFAULTCOLOR = &H0
  LR_MONOCHROME = &H1
  LR_COLOR = &H2
  LR_COPYRETURNORG = &H4
  LR_COPYDELETEORG = &H8
  LR_LOADFROMFILE = &H10
  LR_LOADTRANSPARENT = &H20
  LR_DEFAULTSIZE = &H40
  LR_VGACOLOR = &H80
  LR_LOADMAP3DCOLORS = &H1000
  LR_CREATEDIBSECTION = &H2000
  LR_COPYFROMRESOURCE = &H4000
  LR_SHARED = &H8000&
End Enum
Public Declare Function LoadImageW Lib "user32" (ByVal hInst As Long, ByVal lpsz As Long, ByVal dwImageType As ImageTypes, ByVal dwDesiredWidth As Long, ByVal dwDesiredHeight As Long, ByVal dwFlags As LoadResourceFlags) As Long
Public Declare Function DestroyIcon Lib "user32.dll" (ByVal hIcon As Long) As Long

Private Type BITMAP
    BMType As Long
    BMWidth As Long
    BMHeight As Long
    BMWidthBytes As Long
    BMPlanes As Integer
    BMBitsPixel As Integer
    BMBits As Long
End Type
Public Enum IL_CreateFlags
  ILC_MASK = &H1
  ILC_COLOR = &H0
  ILC_COLORDDB = &HFE
  ILC_COLOR4 = &H4
  ILC_COLOR8 = &H8
  ILC_COLOR16 = &H10
  ILC_COLOR24 = &H18
  ILC_COLOR32 = &H20
  ILC_PALETTE = &H800                  ' (no longer supported...never worked anyway)
  '5.0
  ILC_MIRROR = &H2000
  ILC_PERITEMMIRROR = &H8000
  '6.0
  ILC_ORIGINALSIZE = &H10000
  ILC_HIGHQUALITYSCALE = &H20000
End Enum
Public Declare Function ImageList_Create Lib "comctl32.dll" (ByVal CX As Long, ByVal CY As Long, ByVal Flags As IL_CreateFlags, ByVal cInitial As Long, ByVal cGrow As Long) As Long
Public Declare Function ImageList_Add Lib "comctl32.dll" (ByVal himl As Long, ByVal hbmImage As Long, ByVal hBMMask As Long) As Long
Public Declare Function ImageList_Draw Lib "comctl32.dll" (ByVal himl As Long, ByVal i As Long, ByVal hdcDst As Long, ByVal X As Long, ByVal Y As Long, ByVal fStyle As IL_DrawStyle) As Boolean
Public Declare Function ImageList_Destroy Lib "comctl32.dll" (ByVal himl As Long) As Boolean
Public Declare Function ImageList_ReplaceIcon Lib "comctl32.dll" (ByVal himl As Long, ByVal i As Long, ByVal hIcon As Long) As Long

Public Enum IL_DrawStyle
  ILD_NORMAL = &H0
  ILD_TRANSPARENT = &H1
  ILD_MASK = &H10
  ILD_IMAGE = &H20
'#If (WIN32_IE >= &H300) Then
  ILD_ROP = &H40
'#End If
  ILD_BLEND25 = &H2
  ILD_BLEND50 = &H4
  ILD_OVERLAYMASK = &HF00
 
  ILD_SELECTED = ILD_BLEND50
  ILD_FOCUS = ILD_BLEND25
  ILD_BLEND = ILD_BLEND50
End Enum
Public Declare Function GetFileSize Lib "kernel32" (ByVal hFile As Long, lpFileSizeHigh As Long) As Long
Public Declare Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal dwBytes As Long) As Long
Public Declare Function GlobalFree Lib "kernel32" (ByVal hMem As Long) As Long
Public Declare Function GlobalLock Lib "kernel32" (ByVal hMem As Long) As Long
Public Declare Function GlobalUnlock Lib "kernel32" (ByVal hMem As Long) As Long

' Global Memory Flags
Public Const GMEM_FIXED = &H0
Public Const GMEM_MOVEABLE = &H2
Public Const GMEM_NOCOMPACT = &H10
Public Const GMEM_NODISCARD = &H20
Public Const GMEM_ZEROINIT = &H40
Public Const GMEM_MODIFY = &H80
Public Const GMEM_DISCARDABLE = &H100
Public Const GMEM_NOT_BANKED = &H1000
Public Const GMEM_SHARE = &H2000
Public Const GMEM_DDESHARE = &H2000
Public Const GMEM_NOTIFY = &H4000
Public Const GMEM_LOWER = GMEM_NOT_BANKED
Public Const GMEM_VALID_FLAGS = &H7F72
Public Const GMEM_INVALID_HANDLE = &H8000
Public Const GHND = (GMEM_MOVEABLE Or GMEM_ZEROINIT)
Public Const GPTR = (GMEM_FIXED Or GMEM_ZEROINIT)

Public Enum PERCEIVED
  PERCEIVED_TYPE_CUSTOM = -3
  PERCEIVED_TYPE_UNSPECIFIED = -2
  PERCEIVED_TYPE_FOLDER = -1
  PERCEIVED_TYPE_UNKNOWN = 0
  PERCEIVED_TYPE_TEXT = 1
  PERCEIVED_TYPE_IMAGE = 2
  PERCEIVED_TYPE_AUDIO = 3
  PERCEIVED_TYPE_VIDEO = 4
  PERCEIVED_TYPE_COMPRESSED = 5
  PERCEIVED_TYPE_DOCUMENT = 6
  PERCEIVED_TYPE_SYSTEM = 7
  PERCEIVED_TYPE_APPLICATION = 8
  PERCEIVED_TYPE_GAMEMEDIA = 9
  PERCEIVED_TYPE_CONTACTS = 10
End Enum
Public Enum PERCEIVEDFLAG
    PERCEIVEDFLAG_UNDEFINED = &H0 'No perceived type was found (PERCEIVED_TYPE_UNSPECIFIED.
    PERCEIVEDFLAG_SOFTCODED = &H1 'The perceived type was determined through an association in the registry.
    PERCEIVEDFLAG_HARDCODED = &H2 'The perceived type is inherently known to Windows.
    PERCEIVEDFLAG_NATIVESUPPORT = &H4 'The perceived type was determined through a codec provided with Windows.
    PERCEIVEDFLAG_GDIPLUS = &H10 'The perceived type is supported by the GDI+ library.
    PERCEIVEDFLAG_WMSDK = &H20 'The perceived type is supported by the Windows Media SDK.
    PERCEIVEDFLAG_ZIPFOLDER = &H40 'The perceived type is supported by Windows compressed folders.
End Enum

Public Declare Function CreateStreamOnHGlobal Lib "ole32" (ByVal hGlobal As Long, ByVal fDeleteOnRelease As Long, ppstm As IStream) As Long
   
Public Declare Function CoCreateInstance Lib "ole32" _
                                    (rclsid As Any, _
                                     ByVal pUnkOuter As Long, _
                                     ByVal dwClsContext As Long, _
                                     riid As Any, _
                                     pvarResult As Any) As Long
Public Declare Sub CoTaskMemFree Lib "ole32.dll" (ByVal pv As Long) ' Frees memory allocated by the shell
Public Declare Function AssocQueryString Lib "shlwapi.dll" Alias "AssocQueryStringA" ( _
                                                        ByVal Flags As ASSOCF, _
                                                        ByVal str As ASSOCSTR, _
                                                        ByVal pszAssoc As String, _
                                                        ByVal pszExtra As String, _
                                                        ByVal pszOut As String, _
                                                        ByRef pcchOut As Long) As Long

Public Function BStrFromLPWStr(lpWStr As Long, Optional ByVal CleanupLPWStr As Boolean = True) As String
SysReAllocString VarPtr(BStrFromLPWStr), lpWStr
If CleanupLPWStr Then CoTaskMemFree lpWStr
End Function
Public Function ImageList_AddIcon(himl As Long, hIcon As Long) As Long
  ImageList_AddIcon = ImageList_ReplaceIcon(himl, -1, hIcon)
End Function

Public Function GetHandlerCLSID(ByVal sExt As String, tID As UUID) As Long
Dim nBuf As Long
Dim szBuf As String
Dim sIID As String
Dim sRes As String
sIID = "{8895b1c6-b41f-4c1c-a562-0d564250836f}"
szBuf = String(MAX_PATH, 0)
nBuf = Len(szBuf)
Call AssocQueryString(0, ASSOCSTR_SHELLEXTENSION, sExt, sIID, szBuf, nBuf)

If InStr(szBuf, vbNullChar) > 1 Then
     sRes = Left(szBuf, InStr(szBuf, vbNullChar) - 1)
     Call CLSIDFromString(StrPtr(sRes), tID)
     GetHandlerCLSID = 1
Else
    sExt = Mid(sExt, 2)
    sExt = sExt & "file"
    Debug.Print "No handler returned for extension; trying alternate, " & sExt
    Call AssocQueryString(0, ASSOCSTR_SHELLEXTENSION, sExt, sIID, szBuf, nBuf)
    If InStr(szBuf, vbNullChar) > 1 Then
         sRes = Left(szBuf, InStr(szBuf, vbNullChar) - 1)
         Call CLSIDFromString(StrPtr(sRes), tID)
         GetHandlerCLSID = 1
    End If
End If

End Function
Public Function FilePtypeL(sEx As String) As PERCEIVED
Dim lType As PERCEIVED
Dim hr As Long
Dim lFlag As PERCEIVEDFLAG
Dim lpsz As Long
Dim sType As String
Debug.Print "confirm sEx=" & sEx
hr = AssocGetPerceivedType(StrPtr(sEx), lType, lFlag, lpsz)
If hr <> S_OK Then
    'DebugAppend "FilePtypeL->Call to AssocGetPerceivedType failed, hr=" & hr
    FilePtypeL = -1
    Exit Function
End If
Debug.Print "hr=" & hr & ",lType=" & lType & ",lFlag=" & lFlag
FilePtypeL = lType
Call CoTaskMemFree(lpsz)
    
End Function
Public Function GetFileThumbnail(sFile As String, pidlFQ As Long, CX As Long, CY As Long) As Long
Dim isiif As IShellItemImageFactory
Dim pidl As Long
On Error GoTo e0

If pidlFQ Then
    Call SHCreateItemFromIDList(pidlFQ, IID_IShellItemImageFactory, isiif)
Else
    pidl = ILCreateFromPathW(StrPtr(sFile))
    Call SHCreateItemFromIDList(pidl, IID_IShellItemImageFactory, isiif)
    Call ILFree(pidl)
End If
isiif.GetImage CX, CY, SIIGBF_THUMBNAILONLY, GetFileThumbnail
Set isiif = Nothing
On Error GoTo 0
Exit Function

e0:
Debug.Print "GetFileThumbnail.Error->" & Err.Description & " (" & Err.Number & ")"

End Function
Public Function GetFileThumbnail2(sFile As String, pidlFQ As Long, CX As Long, CY As Long) As Long
'alternate method
Dim isi As IShellItem
Dim pidl As Long
Dim iei As IExtractImage
Dim hBmp As Long
Dim uThumbSize As oleexp.Size
    uThumbSize.CX = CX
    uThumbSize.CY = CY
Dim sRet As String
Dim uThumbFlags As IEIFlags
On Error GoTo e0

If pidlFQ Then
    Call SHCreateShellItem(0&, 0&, pidlFQ, isi)
Else
    pidl = ILCreateFromPathW(StrPtr(sFile))
    Call SHCreateShellItem(0&, 0&, pidl, isi)
    Call CoTaskMemFree(pidl) 'also a change that should have been made, had originally used ILFree, which shouldn't be used on Win2k+
End If

isi.BindToHandler ByVal 0&, BHID_ThumbnailHandler, IID_IExtractImage, iei
If (iei Is Nothing) Then
    Debug.Print "GetFileThumbnail2.Failed to create IExtractImage"
    Exit Function
End If

            uThumbFlags = IEIFLAG_ORIGSIZE
            sRet = String$(MAX_PATH, 0)
            iei.GetLocation StrPtr(sRet), MAX_PATH, 0&, uThumbSize, 32, uThumbFlags
hBmp = iei.Extract()
GetFileThumbnail2 = hBmp
Set iei = Nothing

On Error GoTo 0
Exit Function

e0:
Debug.Print "GetFileThumbnail2.Error->" & Err.Description & " (" & Err.Number & ")"
End Function
Public Sub hBitmapToPictureBox(picturebox As Object, hBitmap As Long, Optional X As Long = 0&, Optional Y As Long = 0&)

'This or similar is always given as the example on how to do this
'But it results in transparency being lost
'Dim hdcBitmap As Long
'
'hdcBitmap = CreateCompatibleDC(0)
'SelectObject hdcBitmap, hBitmap
'BitBlt picturebox.hDC, 0, 0, picturebox.ScaleWidth, _
'picturebox.ScaleHeight, hdcBitmap, 0, 0, vbSrcCopy
'So the below method seems a little ackward, but it works. It can
'be done without the ImageList trick, but it's much more code
Dim himlBmp As Long
Dim tBMP As BITMAP
Dim CX As Long, CY As Long
Call GetObject(hBitmap, LenB(tBMP), tBMP)
CX = tBMP.BMWidth
CY = tBMP.BMHeight
If CX = 0 Then
    Debug.Print "no width"
    Exit Sub
End If
himlBmp = ImageList_Create(CX, CY, ILC_COLOR32, 1, 1)

ImageList_Add himlBmp, hBitmap, 0&
ImageList_Draw himlBmp, 0, picturebox.hDC, X, Y, ILD_NORMAL

ImageList_Destroy himlBmp
End Sub
Public Function DoIcoPreview(sIcon As String, hDC As Long, CXY As Long, Optional X As Long = 0&, Optional Y As Long = 0&) As Long

Dim hIcon As Long
Dim himlIcon As Long
himlIcon = ImageList_Create(CXY, CXY, ILC_COLOR32, 1, 1)

hIcon = LoadImageW(App.hInstance, StrPtr(sIcon), IMAGE_ICON, CXY, CXY, LR_LOADFROMFILE)
If hIcon = 0 Then
    Debug.Print "mIShellFolderDefs.DoIcoPreview.Failed to get hIcon", 2
    DoIcoPreview = -1
    Exit Function
End If

ImageList_AddIcon himlIcon, hIcon
ImageList_Draw himlIcon, 0, hDC, X, Y, ILD_NORMAL

Call DestroyIcon(hIcon)
ImageList_Destroy himlIcon
End Function
'Project-associated UUIDs. If your project is using the mIID.bas from oleexp 3.0 or higher, you don't need any of the below
'Public Sub DEFINE_UUID(Name As UUID, L As Long, w1 As Integer, w2 As Integer, B0 As Byte, b1 As Byte, b2 As Byte, B3 As Byte, b4 As Byte, b5 As Byte, b6 As Byte, b7 As Byte)
'  With Name
'    .Data1 = L
'    .Data2 = w1
'    .Data3 = w2
'    .Data4(0) = B0
'    .Data4(1) = b1
'    .Data4(2) = b2
'    .Data4(3) = B3
'    .Data4(4) = b4
'    .Data4(5) = b5
'    .Data4(6) = b6
'    .Data4(7) = b7
'  End With
'End Sub
'
'Public Function IID_IPreviewHandler() As UUID
''{8895b1c6-b41f-4c1c-a562-0d564250836f}
'Static iid As UUID
' If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H8895B1C6, CInt(&HB41F), CInt(&H4C1C), &HA5, &H62, &HD, &H56, &H42, &H50, &H83, &H6F)
' IID_IPreviewHandler = iid
'End Function
'Public Function IID_IPreviewHandlerVisuals() As UUID
''{196bf9a5-b346-4ef0-aa1e-5dcdb76768b1}
'Static iid As UUID
' If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H196BF9A5, CInt(&HB346), CInt(&H4EF0), &HAA, &H1E, &H5D, &HCD, &HB7, &H67, &H68, &HB1)
' IID_IPreviewHandlerVisuals = iid
'End Function
'Public Function IID_IInitializeWithStream() As UUID
''{b824b49d-22ac-4161-ac8a-9916e8fa3f7f}
'Static iid As UUID
' If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HB824B49D, CInt(&H22AC), CInt(&H4161), &HAC, &H8A, &H99, &H16, &HE8, &HFA, &H3F, &H7F)
' IID_IInitializeWithStream = iid
'End Function
'Public Function IID_IInitializeWithFile() As UUID
''{b7d14566-0509-4cce-a71f-0a554233bd9b}
'Static iid As UUID
' If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HB7D14566, CInt(&H509), CInt(&H4CCE), &HA7, &H1F, &HA, &H55, &H42, &H33, &HBD, &H9B)
' IID_IInitializeWithFile = iid
'End Function
'Public Function IID_IInitializeWithItem() As UUID
''{7f73be3f-fb79-493c-a6c7-7ee14e245841}
'Static iid As UUID
' If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H7F73BE3F, CInt(&HFB79), CInt(&H493C), &HA6, &HC7, &H7E, &HE1, &H4E, &H24, &H58, &H41)
' IID_IInitializeWithItem = iid
'End Function
'Public Function IID_IShellItemImageFactory() As UUID
''{BCC18B79-BA16-442F-80C4-8A59C30C463B}
'Static iid As UUID
'If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &HBCC18B79, CInt(&HBA16), CInt(&H442F), &H80, &HC4, &H8A, &H59, &HC3, &HC, &H46, &H3B)
'IID_IShellItemImageFactory = iid
'End Function
'Public Function IID_IShellItem() As UUID
'Static iid As UUID
'If (iid.Data1 = 0) Then Call DEFINE_UUID(iid, &H43826D1E, CInt(&HE718), CInt(&H42EE), &HBC, &H55, &HA1, &HE2, &H61, &HC3, &H7B, &HFE)
'IID_IShellItem = iid
'End Function
'
