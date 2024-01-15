Attribute VB_Name = "Module1"
Option Explicit

Public Function dbg_GUIDToString(tg As UUID, Optional bBrack As Boolean = True) As String
'StringFromGUID2 never works, even "working" code from vbaccelerator AND MSDN
dbg_GUIDToString = Right$("00000000" & Hex$(tg.Data1), 8) & "-" & Right$("0000" & Hex$(tg.Data2), 4) & "-" & Right$("0000" & Hex$(tg.Data3), 4) & _
"-" & Right$("00" & Hex$(CLng(tg.Data4(0))), 2) & Right$("00" & Hex$(CLng(tg.Data4(1))), 2) & "-" & Right$("00" & Hex$(CLng(tg.Data4(2))), 2) & _
Right$("00" & Hex$(CLng(tg.Data4(3))), 2) & Right$("00" & Hex$(CLng(tg.Data4(4))), 2) & Right$("00" & Hex$(CLng(tg.Data4(5))), 2) & _
Right$("00" & Hex$(CLng(tg.Data4(6))), 2) & Right$("00" & Hex$(CLng(tg.Data4(7))), 2)
If bBrack Then dbg_GUIDToString = "{" & dbg_GUIDToString & "}"
End Function
Public Function BStrFromLPWStr(lpWStr As LongPtr, Optional ByVal CleanupLPWStr As Boolean = True) As String
SysReAllocStringW VarPtr(BStrFromLPWStr), lpWStr
If CleanupLPWStr Then CoTaskMemFree lpWStr
End Function
Public Function ImageList_AddIcon(himl As LongPtr, hIcon As LongPtr) As Long
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
     sRes = Left$(szBuf, InStr(szBuf, vbNullChar) - 1)
     Call CLSIDFromString(sRes, tID)
     GetHandlerCLSID = 1
Else
    sExt = Mid(sExt, 2)
    sExt = sExt & "file"
    Debug.Print "No handler returned for extension; trying alternate, " & sExt
    Call AssocQueryString(0, ASSOCSTR_SHELLEXTENSION, sExt, sIID, szBuf, nBuf)
    If InStr(szBuf, vbNullChar) > 1 Then
         sRes = Left$(szBuf, InStr(szBuf, vbNullChar) - 1)
         Call CLSIDFromString(sRes, tID)
         GetHandlerCLSID = 1
    End If
End If

End Function
Public Function FilePtypeL(sEx As String) As PERCEIVED
Dim lType As PERCEIVED
Dim hr As Long
Dim lFlag As PERCEIVEDFLAG
Dim lpsz As LongPtr
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
Public Function GetFileThumbnail(sFile As String, pidlFQ As LongPtr, CX As Long, CY As Long) As LongPtr
Dim isiif As IShellItemImageFactory
Dim pidl As LongPtr
On Error GoTo e0

If pidlFQ Then
    Call SHCreateItemFromIDList(pidlFQ, IID_IShellItemImageFactory, isiif)
Else
    pidl = ILCreateFromPathW(StrPtr(sFile))
    Call SHCreateItemFromIDList(pidl, IID_IShellItemImageFactory, isiif)
    Call ILFree(pidl)
End If
Dim cxy As LongLong
Dim pSZ As SIZE
pSZ.cx = CX: pSZ.cy	= CY
CopyMemory cxy, pSZ, LenB(Of SIZE)
isiif.GetImage cxy, SIIGBF_THUMBNAILONLY, GetFileThumbnail
Set isiif = Nothing
On Error GoTo 0
Exit Function

e0:
Debug.Print "GetFileThumbnail.Error->" & Err.Description & " (" & Err.Number & ")"

End Function
Public Function GetFileThumbnail2(sFile As String, pidlFQ As LongPtr, CX As Long, CY As Long) As LongPtr
'alternate method
Dim isi As IShellItem
Dim pidl As LongPtr
Dim iei As IExtractImage
Dim hBmp As LongPtr
Dim uThumbSize As SIZE
    uThumbSize.cx = CX
    uThumbSize.cy = CY
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
Public Sub hBitmapToPictureBox(picturebox As Object, hBitmap As LongPtr, Optional X As Long = 0&, Optional Y As Long = 0&)

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
Dim himlBmp As LongPtr
Dim tBMP As BITMAP
Dim CX As Long, CY As Long
Call GetObjectW(hBitmap, LenB(tBMP), tBMP)
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

Dim hIcon As LongPtr
Dim himlIcon As LongPtr
himlIcon = ImageList_Create(CXY, CXY, ILC_COLOR32, 1, 1)

hIcon = LoadImageW(App.hInstance, StrPtr(sIcon), IMAGE_ICON, CXY, CXY, LR_LOADFROMFILE)
If hIcon = 0 Then
    Debug.Print "DoIcoPreview.Failed to get hIcon", 2
    DoIcoPreview = -1
    Exit Function
End If

ImageList_AddIcon himlIcon, hIcon
ImageList_Draw himlIcon, 0, hDC, X, Y, ILD_NORMAL

Call DestroyIcon(hIcon)
ImageList_Destroy himlIcon
End Function

