VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmUpdateCheck 
   Caption         =   "Windows Update Checker v1.0"
   ClientHeight    =   6075
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9975
   Icon            =   "frmUpdateCheck.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6075
   ScaleWidth      =   9975
   StartUpPosition =   2  'CenterScreen
   Begin WindowsUpdate.ucSplitBar SplitBarBottom 
      Height          =   30
      Left            =   0
      TabIndex        =   7
      Top             =   4560
      Width           =   4815
      _ExtentX        =   8493
      _ExtentY        =   53
   End
   Begin WindowsUpdate.ucSplitBar SplitBar 
      Height          =   30
      Left            =   0
      TabIndex        =   5
      Top             =   2520
      Width           =   4815
      _ExtentX        =   8493
      _ExtentY        =   53
   End
   Begin VB.PictureBox picBottom 
      Align           =   2  'Align Bottom
      Height          =   1455
      Left            =   0
      ScaleHeight     =   1395
      ScaleWidth      =   9915
      TabIndex        =   1
      Top             =   4620
      Width           =   9975
      Begin VB.CommandButton cmdInstall 
         Caption         =   "Install"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   5160
         TabIndex        =   12
         Top             =   50
         Width           =   1095
      End
      Begin VB.CommandButton cmdGoOnline 
         Caption         =   "MSDN"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   8760
         TabIndex        =   8
         Top             =   50
         Width           =   1095
      End
      Begin VB.CommandButton cmdUpdate 
         Caption         =   "&Check For Updates"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   4
         Top             =   50
         Width           =   1815
      End
      Begin VB.CommandButton cmdOpenKB 
         Caption         =   "&Open KB Article"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2040
         TabIndex        =   3
         Top             =   50
         Width           =   1455
      End
      Begin VB.TextBox txtKB 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3600
         TabIndex        =   2
         Top             =   100
         Width           =   1335
      End
      Begin VB.TextBox txtRemote 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   8520
         TabIndex        =   10
         Top             =   120
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Remote Machine:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   6960
         TabIndex        =   11
         Top             =   195
         Visible         =   0   'False
         Width           =   1500
      End
      Begin VB.Line Line1 
         X1              =   0
         X2              =   9240
         Y1              =   550
         Y2              =   550
      End
      Begin VB.Label lblDescription 
         BackStyle       =   0  'Transparent
         Caption         =   "Description:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   10
         TabIndex        =   6
         Top             =   575
         Width           =   9135
      End
   End
   Begin MSComctlLib.ListView lvUpdateInstalled 
      Height          =   1815
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9135
      _ExtentX        =   16113
      _ExtentY        =   3201
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      _Version        =   393217
      Icons           =   "ImageList1"
      SmallIcons      =   "ImageList1"
      ColHdrIcons     =   "ImageList1"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   8
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Title   (Installed Updates)"
         Object.Width           =   6174
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "URL"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Severity"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Dwld"
         Object.Width           =   1147
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "File Size"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Time"
         Object.Width           =   1940
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "UpdateID"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "Description"
         Object.Width           =   2540
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   480
      Top             =   1800
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUpdateCheck.frx":0442
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUpdateCheck.frx":0894
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUpdateCheck.frx":0CE6
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUpdateCheck.frx":1138
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView lvUpdateNotInstalled 
      Height          =   1815
      Left            =   0
      TabIndex        =   9
      Top             =   2640
      Width           =   9135
      _ExtentX        =   16113
      _ExtentY        =   3201
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      _Version        =   393217
      Icons           =   "ImageList1"
      SmallIcons      =   "ImageList1"
      ColHdrIcons     =   "ImageList1"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   8
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Title   (Updates not Installed)"
         Object.Width           =   6174
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "URL"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Severity"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Dwld"
         Object.Width           =   1147
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "File Size"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Time"
         Object.Width           =   1940
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "UpdateID"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "Description"
         Object.Width           =   2540
      EndProperty
   End
End
Attribute VB_Name = "frmUpdateCheck"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function GetWindowRect Lib "user32.dll" (ByVal hwnd As Long, ByRef lpRect As RECT) As Long

Private Sub cmdGoOnline_Click()
    
    LaunchURLInNewBrowser "http://msdn.microsoft.com/library/default.asp?url=/library/en-us/wua_sdk/wua/portal_client.asp"

End Sub

Private Sub cmdInstall_Click()
If lvUpdateNotInstalled.SelectedItem Is Nothing Then Exit Sub

MousePointer = 11

Dim objCollection As Object
Dim objSearcher As Object
Dim objResults As Object
Dim colUpdates As Object

    'install update
    Set objCollection = CreateObject("Microsoft.Update.UpdateColl")
    Set objSearcher = CreateObject("Microsoft.Update.Searcher")
    Set objResults = objSearcher.Search("UpdateID='" & lvUpdateNotInstalled.SelectedItem.SubItems(6) & "'")

    'Debug.Print "total:" & objResults.Updates.Count
    
    Set colUpdates = objResults.Updates 'ISearchResult.Updates
    objCollection.Add (colUpdates.Item(0))

    Dim updateSession As Object
    Dim downloader As Object
    Dim downloadResult As Object
    
    Set updateSession = CreateObject("Microsoft.Update.Session")
    Set downloader = updateSession.CreateUpdateDownloader()

    downloader.Updates = objCollection
    'Debug.Print "Downloading Update..."
    
    Set downloadResult = downloader.Download()
    
    Dim installer As Object
    Dim installationResult As Object
    Set installer = updateSession.CreateUpdateInstaller()
    
    'Debug.Print "Installing Update..."
    
    installer.Updates = objCollection
    Set installationResult = installer.Install()

    Select Case installationResult.ResultCode
        Case 0
            Debug.Print "Result: not started"
        Case 1
            Debug.Print "Result: in progress"
        Case 2
            Debug.Print "Result: succeeded"
        Case 3
            Debug.Print "Result: succeeded with errors"
        Case 4
            Debug.Print "Result: failed"
        Case 5
            Debug.Print "Result: aborted"
    
    End Select

SkipPast:

    Set installer = Nothing
    Set installationResult = Nothing
    
    Set downloader = Nothing
    Set downloadResult = Nothing
    
    Set updateSession = Nothing

    Set objCollection = Nothing
    Set objSearcher = Nothing
    Set objResults = Nothing
    Set colUpdates = Nothing


    cmdUpdate_Click

    MousePointer = 1
End Sub


Private Sub cmdOpenKB_Click()
    
    If Len(txtKB.Text) > 0 Then
        LaunchURLInNewBrowser "http://support.microsoft.com/default.aspx?kbid=" & txtKB.Text
    End If

End Sub


Private Sub cmdUpdate_Click()

Dim objSession As Object, objSearcher As Object
Dim colUpdates As Object, objResults As Object
Dim objCategories As Object

Dim objIdentity As Object, objInstallationBehavior As Object
Dim strInfo As Object

Dim i As Integer, x As Integer

On Error Resume Next

    lvUpdateInstalled.ListItems.Clear
    lvUpdateNotInstalled.ListItems.Clear
    lblDescription.Caption = "Description:"

    MousePointer = 11
    
    'atl-ws-01 is the computer name
    'Set objSession = CreateObject("Microsoft.Update.Session", "atl-ws-01")
    
    If Len(txtRemote.Text) = 0 Then
        Set objSession = CreateObject("Microsoft.Update.Session")   'IUpdateSession
    Else
        Set objSession = CreateObject("Microsoft.Update.Session", txtRemote.Text)
    End If
    
    Set objSearcher = objSession.CreateUpdateSearcher 'IUpdateSession::CreateUpdateSearcher
            'IUpdateSearcher::Search
    Set objResults = objSearcher.Search("Type='Software'") 'ISearchResult
    Set colUpdates = objResults.Updates 'ISearchResult.Updates
            'IUpdateCollection
    
    
'try and get to IUpdate2 or IUpdateDownloadContentCollection - IUPdateDownloadContent

'http://download.microsoft.com/download/e/1/4/e14c0c02-591b-4696-8552-eb710c26a3cd/NDP1.1sp1-KB886903-X86.exe
Dim lvItem As ListItem

    For i = 0 To colUpdates.Count - 1   'IUpdate
        If colUpdates.Item(i).IsInstalled = "True" Then
            Set lvItem = frmUpdateCheck.lvUpdateInstalled.ListItems.Add(, , colUpdates.Item(i).Title, 0, 0)
            lvItem.Icon = 2
            lvItem.SmallIcon = 2
        Else
            Set lvItem = frmUpdateCheck.lvUpdateNotInstalled.ListItems.Add(, , colUpdates.Item(i).Title, 0, 0)
            lvItem.Icon = 3
            lvItem.SmallIcon = 3
        End If
        
        lvItem.SubItems(2) = colUpdates.Item(i).MsrcSeverity
        If colUpdates.Item(i).MsrcSeverity = "Critical" Then
            lvItem.Icon = 4
            lvItem.SmallIcon = 4
        End If
        
        lvItem.SubItems(7) = colUpdates.Item(i).Description
        lvItem.SubItems(3) = colUpdates.Item(i).IsDownloaded
        lvItem.SubItems(4) = colUpdates.Item(i).MaxDownloadSize
        lvItem.SubItems(5) = colUpdates.Item(i).LastDeploymentChangeTime
        lvItem.SubItems(6) = colUpdates.Item(i).Identity.UpdateID
    
    
        If colUpdates.Item(i).MoreInfoUrls.Count > 0 Then
            For x = 0 To colUpdates.Item(i).MoreInfoUrls.Count - 1
                lvItem.SubItems(1) = colUpdates.Item(i).MoreInfoUrls.Item(x)
            Next x
        End If
        
    Next i

    MousePointer = 1

ErrorHandler:
    If Err.Number > 0 Then
        MsgBox "Error:frmMain:cmdUpdate_Click:" & i & ":Line#:" & Erl & ":" & Err.Number & ":" & Err.Description
    End If
    
    Me.MousePointer = 1

    Set lvItem = Nothing
    
    Set strInfo = Nothing
    Set objSession = Nothing
    Set objSearcher = Nothing
    Set colUpdates = Nothing
    Set objResults = Nothing
    Set objCategories = Nothing
    Set objIdentity = Nothing
    Set objInstallationBehavior = Nothing

End Sub

Private Sub Form_Load()
    
    SplitBar.Orientation = espHorizontal
    SplitBarBottom.Orientation = espHorizontal

End Sub


Private Sub lvUpdateInstalled_Click()
    
    If lvUpdateInstalled.SelectedItem Is Nothing Then Exit Sub
    lblDescription.Caption = lvUpdateInstalled.SelectedItem.SubItems(7)
    
    Dim strName As String
    strName = lvUpdateInstalled.SelectedItem.Text
    If InStr(strName, "(KB") Then
        strName = Mid$(strName, InStr(strName, "(KB") + 3)
        txtKB.Text = Left$(strName, Len(strName) - 1)
    Else
        txtKB.Text = ""
    End If

End Sub

Private Sub lvUpdateInstalled_DblClick()
    
    If lvUpdateInstalled.SelectedItem Is Nothing Then Exit Sub
    
    LaunchURLInNewBrowser lvUpdateInstalled.SelectedItem.SubItems(1)

End Sub

Private Sub lvUpdateInstalled_KeyUp(KeyCode As Integer, Shift As Integer)
    
    lvUpdateInstalled_Click

End Sub

Private Sub lvUpdateNotInstalled_Click()
    If lvUpdateNotInstalled.SelectedItem Is Nothing Then Exit Sub
    lblDescription.Caption = lvUpdateNotInstalled.SelectedItem.SubItems(7)
    
    Dim strName As String
    
    strName = lvUpdateNotInstalled.SelectedItem.Text
    If InStr(strName, "(KB") Then
        strName = Mid$(strName, InStr(strName, "(KB") + 3)
        txtKB.Text = Left$(strName, Len(strName) - 1)
    Else
        txtKB.Text = ""
    End If

End Sub

Private Sub lvUpdateNotInstalled_DblClick()
    
    If lvUpdateNotInstalled.SelectedItem Is Nothing Then Exit Sub
    
    LaunchURLInNewBrowser lvUpdateNotInstalled.SelectedItem.SubItems(1)

End Sub

Private Sub lvUpdateNotInstalled_KeyUp(KeyCode As Integer, Shift As Integer)
    
    lvUpdateNotInstalled_Click

End Sub

Private Sub picBottom_Resize()
    
    lblDescription.Width = picBottom.Width - 20
    lblDescription.Height = picBottom.Height - lblDescription.Top
    Line1.X2 = picBottom.Width

End Sub

Private Sub Form_Resize()
On Error Resume Next
    
    lvUpdateInstalled.Top = 0
    lvUpdateInstalled.Left = 0
    lvUpdateInstalled.Width = Me.ScaleWidth
    lvUpdateNotInstalled.Width = Me.ScaleWidth
    lvUpdateNotInstalled.Left = 0
    SplitBar.Width = Me.ScaleWidth
    SplitBarBottom.Width = Me.ScaleWidth
    lvUpdateInstalled.Height = SplitBar.Top
    lvUpdateNotInstalled.Top = SplitBar.Top + 30
    lvUpdateNotInstalled.Height = Me.ScaleHeight - 90 - SplitBar.Top - picBottom.ScaleHeight
    SplitBarBottom.Top = picBottom.Top - 45

End Sub

Private Sub SplitBar_AfterSize(newSize As Long)

    If newSize + SplitBar.Top + 2 > picBottom.Top Then Exit Sub
    If newSize + SplitBar.Top + 2 < 0 Then Exit Sub
    
    SplitBar.Top = SplitBar.Top + newSize
    Form_Resize

End Sub

Private Sub SplitBar_BeforeSize()
    
    ResizeSplitter

End Sub

Private Sub ResizeSplitter()
    
Dim pRect As RECT
    
    GetWindowRect Me.hwnd, pRect
    SplitBar.RectLeft = pRect.Left
    SplitBar.RectRight = pRect.Right
    SplitBar.RectTop = pRect.Top + 5
    SplitBar.RectBottom = pRect.Bottom - 5

End Sub

Private Sub SplitBarBottom_AfterSize(newSize As Long)
    
    picBottom.Height = picBottom.Height - newSize
    Form_Resize

End Sub

Private Sub SplitBarBottom_BeforeSize()
    
    ResizeSplitterBottom

End Sub

Private Sub ResizeSplitterBottom()

Dim pRect As RECT
    
    GetWindowRect Me.hwnd, pRect
    SplitBarBottom.RectLeft = pRect.Left
    SplitBarBottom.RectRight = pRect.Right
    SplitBarBottom.RectTop = pRect.Top + 5
    SplitBarBottom.RectBottom = pRect.Bottom - 5
End Sub

