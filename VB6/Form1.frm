VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "IPreviewHandler Example"
   ClientHeight    =   4290
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   5805
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   4290
   ScaleWidth      =   5805
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Preview"
      Height          =   285
      Left            =   4755
      TabIndex        =   5
      Top             =   75
      Width           =   1020
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   1260
      TabIndex        =   4
      Top             =   90
      Width           =   3420
   End
   Begin VB.Frame Frame1 
      Caption         =   "Preview"
      Height          =   3510
      Left            =   60
      TabIndex        =   1
      Top             =   690
      Width           =   5685
      Begin VB.PictureBox Picture1 
         AutoRedraw      =   -1  'True
         Height          =   3180
         Left            =   75
         ScaleHeight     =   208
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   366
         TabIndex        =   3
         Top             =   210
         Width           =   5550
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Choose File..."
      Height          =   330
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   1110
   End
   Begin VB.Label Label1 
      Height          =   225
      Left            =   120
      TabIndex        =   2
      Top             =   405
      Width           =   4800
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Private Type COMDLG_FILTERSPEC
'    pszName As String
'    pszSpec As String
'End Type
Private ipv As IPreviewHandler
Private hGlobal As Long

Private Sub Command1_Click()
Dim isi As IShellItem
Dim FileFilter() As COMDLG_FILTERSPEC
Dim rc As RECT

ReDim FileFilter(0)
FileFilter(0).pszName = "All Files"
FileFilter(0).pszSpec = "*.*"

Dim fod As FileOpenDialog
Set fod = New FileOpenDialog

With fod
    .SetOkButtonLabel "Preview"
    .SetOptions FOS_FILEMUSTEXIST Or FOS_DONTADDTORECENT Or FOS_ALLNONSTORAGEITEMS
    .SetTitle "Select a file that can be previewed"
    .SetFileTypes 1, VarPtr(FileFilter(0))
    
    On Error Resume Next '.show throws an automation error if cancel is clicked
    .Show Me.hWnd
    On Error GoTo 0
    fod.GetResult isi
    If (isi Is Nothing) Then Exit Sub
End With
    rc.Top = 16
    rc.Bottom = (Picture1.Height / Screen.TwipsPerPixelY) - 16
    rc.Left = 16
    rc.Right = (Picture1.Width / Screen.TwipsPerPixelX) - 16

ShowPreviewForFile isi, Picture1.hWnd, rc, Picture1
Set isi = Nothing
Set fod = Nothing
End Sub
Private Sub ShowPreviewForFile(isi As IShellItem, hWnd As Long, rc As RECT, objpic As Object, Optional sFileIn As String = "")
Dim iif As IInitializeWithFile
Dim iis As IInitializeWithStream
Dim iisi As IInitializeWithItem
Dim pVis As IPreviewHandlerVisuals
Dim pUnk As oleexp.IUnknown
Dim hr As Long
Dim sFile As String, sExt As String
Dim lp As Long
Dim tHandler As UUID
On Error GoTo e0

If (isi Is Nothing) Then
    Debug.Print "no isi"
    If sFileIn <> "" Then
        sFile = sFileIn
    End If
Else
    Debug.Print "using isi"
    isi.GetDisplayName SIGDN_FILESYSPATH, lp
    sFile = BStrFromLPWStr(lp)
End If
    Debug.Print "sFile=" & sFile
    sExt = Right$(sFile, (Len(sFile) - InStrRev(sFile, ".")) + 1)
    Debug.Print "sExt=" & sExt

If sExt = "" Then Exit Sub

If (ipv Is Nothing) = False Then
    ipv.Unload
    Set ipv = Nothing
End If


hr = GetHandlerCLSID(sExt, tHandler)
If hr = 1 Then
    Debug.Print "Got handler CLSID; attempting to create IPreviewHandler"
    hr = CoCreateInstance(tHandler, 0, CLSCTX_INPROC_SERVER Or CLSCTX_LOCAL_SERVER, IID_IPreviewHandler, ipv)
    If (ipv Is Nothing) Then
        Debug.Print "Failed to create IPreviewHandler interface, hr=" & hr
        Exit Sub
    End If
    'Set iisi = ipv 'this normally can be used in place of Set pUnk / .QueryInterface, but we need the HRESULT
    Set pUnk = ipv
'    Set iif = ipv
    Set pUnk = ipv
    If pUnk.QueryInterface(IID_IInitializeWithFile, iif) = S_OK Then
        hr = iif.Initialize(sFile, STGM_READ)
        GoTo gpvh
    Else
        Debug.Print "IInitializeWithFile not supported."
    End If
    If pUnk.QueryInterface(IID_IInitializeWithItem, iisi) = S_OK Then
        hr = iisi.Initialize(isi, STGM_READ)
        Debug.Print "iisi.init hr=" & hr
        If hr = S_OK Then GoTo gpvh
    Else
        Debug.Print "IInitializeWithItem not supported."
    End If

        'use IStream
        Dim hFile As Long
        Dim pstrm As IStream
        Dim lpGlobal As Long
        Dim dwSize As Long
        Debug.Print "Attempting to use IStream"
        hFile = CreateFile(sFile, FILE_READ_DATA, FILE_SHARE_READ, ByVal 0&, OPEN_EXISTING, 0, 0)
        If hFile Then
            dwSize = GetFileSize(hFile, ByVal 0&)
            Debug.Print "Got file size=" & dwSize
            If dwSize = 0 Then Exit Sub
            hGlobal = GlobalAlloc(GPTR, dwSize)
            lpGlobal = GlobalLock(hGlobal)
            If lpGlobal Then
                Call ReadFile(hFile, ByVal lpGlobal, dwSize, dwSize, ByVal 0&)
                Call GlobalUnlock(hGlobal)
                Call CreateStreamOnHGlobal(hGlobal, 1, pstrm)
'                Set iis = ipv
                Set pUnk = ipv
                hr = pUnk.QueryInterface(IID_IInitializeWithStream, iis)
                Debug.Print "QI.hr=" & hr
                If (iis Is Nothing) Then
                    Debug.Print "IInitializeWithStream not supported."
                    Call CloseHandle(hFile)
                    GoTo out
                Else
                    hr = iis.Initialize(pstrm, STGM_READ)
                End If
            End If
            
            Call CloseHandle(hFile)

    End If
gpvh:
    hr = ipv.SetWindow(hWnd, rc)
    Debug.Print "SetWindow hr=" & hr
    hr = ipv.DoPreview()
    Debug.Print "DoPreview hr=" & hr
    Dim piunk As oleexp.IUnknown
    Set piunk = ipv
    hr = piunk.QueryInterface(IID_IPreviewHandlerVisuals, pVis)
    If (pVis Is Nothing) = False Then
        Debug.Print "Handler implements IPreviewHandlerVisuals; setting bk color to white"
        pVis.SetBackgroundColor &HFFFFFF
    End If
    isi.GetDisplayName SIGDN_NORMALDISPLAY, lp
    sFile = BStrFromLPWStr(lp)
    Label1.Caption = "DoPreview called for " & sFile
Else
    'images and videos aren't handled that way normally, so we'll do it another way
    Debug.Print "No registered handler; trying alternate method for images..."
    Dim lPcv As Long
    lPcv = FilePtypeL(sExt)
    Debug.Print "Perceived type=" & lPcv
    If lPcv = PERCEIVED_TYPE_IMAGE Then
        If Right$(sFile, 4) = ".ico" Then
            'the below methods don't properly render icons transparent
            'so we'll use a different method that does
            If DoIcoPreview(sFile, objpic.hDC, 32) = -1 Then
                GoTo gfthm
            End If
            objpic.Refresh
            Label1.Caption = "Manually generated preview for icon."
            GoTo out
        Else
gfthm:
            Dim hbm As Long
            hbm = GetFileThumbnail(sFile, 0, objpic.ScaleWidth, objpic.ScaleHeight)
            Debug.Print "hbm=" & hbm
            objpic.Cls
            hBitmapToPictureBox objpic, hbm
            objpic.Refresh
            Label1.Caption = "Manually generated preview for image."
        End If
    Else
        Label1.Caption = "Could not find registered preview handler for file type."
    
    End If
End If
out:

Set iisi = Nothing
Set iif = Nothing
Set iis = Nothing

On Error GoTo 0
Exit Sub

e0:
Debug.Print "ShowPreviewForFile.Error->" & Err.Description & " (" & Err.Number & ")"
End Sub

Private Sub Command2_Click()
Dim psi As IShellItem
Dim rc As RECT
Dim pidl As Long
pidl = ILCreateFromPathW(StrPtr(Text1.Text))
Call SHCreateItemFromIDList(pidl, IID_IShellItem, psi)

'Call SHCreateItemFromParsingName(StrPtr(Text1.Text), ByVal 0&, IID_IShellItem, psi)
    rc.Top = 16
    rc.Bottom = (Picture1.Height / Screen.TwipsPerPixelY) - 16
    rc.Left = 16
    rc.Right = (Picture1.Width / Screen.TwipsPerPixelX) - 16
    ShowPreviewForFile psi, Picture1.hWnd, rc, Picture1, Text1.Text
End Sub

Private Sub Form_Unload(Cancel As Integer)
If (ipv Is Nothing) = False Then
    ipv.Unload
    Set ipv = Nothing
End If


End Sub
