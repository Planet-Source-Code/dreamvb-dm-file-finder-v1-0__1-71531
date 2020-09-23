VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmmain 
   Caption         =   "DM File Finder"
   ClientHeight    =   5220
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   8835
   Icon            =   "frmmain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5220
   ScaleWidth      =   8835
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtName 
      Height          =   315
      Left            =   675
      TabIndex        =   1
      Top             =   270
      Width           =   3900
   End
   Begin VB.CommandButton cmdOp 
      Caption         =   ". . . ."
      Height          =   315
      Index           =   3
      Left            =   4440
      TabIndex        =   3
      ToolTipText     =   "Browse Folder"
      Top             =   750
      Width           =   495
   End
   Begin MSComctlLib.StatusBar sBar1 
      Align           =   2  'Align Bottom
      Height          =   300
      Left            =   0
      TabIndex        =   12
      Top             =   4920
      Width           =   8835
      _ExtentX        =   15584
      _ExtentY        =   529
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   12515
            Text            =   "Welcome to DM File Finder V1.0"
            TextSave        =   "Welcome to DM File Finder V1.0"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView LstResult 
      Height          =   2415
      Left            =   30
      TabIndex        =   8
      Top             =   2400
      Width           =   8355
      _ExtentX        =   14737
      _ExtentY        =   4260
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   5
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Filename"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Path"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Filesize"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Date"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Ext"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Timer tmrBusy 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   7110
      Top             =   240
   End
   Begin VB.PictureBox pSrc 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   480
      Left            =   5130
      Picture         =   "frmmain.frx":27A2
      ScaleHeight     =   480
      ScaleWidth      =   1920
      TabIndex        =   11
      Top             =   195
      Visible         =   0   'False
      Width           =   1920
   End
   Begin VB.PictureBox pBusy 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   480
      Left            =   7920
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   1740
      Width           =   480
   End
   Begin VB.ComboBox cboFolder 
      Height          =   315
      Left            =   675
      TabIndex        =   2
      Top             =   750
      Width           =   3675
   End
   Begin VB.CheckBox chkSubFolder 
      Caption         =   "Include subfolders"
      Height          =   240
      Left            =   675
      TabIndex        =   4
      Top             =   1170
      Value           =   1  'Checked
      Width           =   3900
   End
   Begin VB.CommandButton cmdOp 
      Caption         =   "Reset"
      Height          =   350
      Index           =   2
      Left            =   7650
      TabIndex        =   7
      Top             =   1230
      Width           =   1024
   End
   Begin VB.CommandButton cmdOp 
      Caption         =   "&Stop"
      Height          =   350
      Index           =   1
      Left            =   7665
      TabIndex        =   6
      Top             =   705
      Width           =   1024
   End
   Begin VB.CommandButton cmdOp 
      Caption         =   "Find"
      Height          =   350
      Index           =   0
      Left            =   7650
      TabIndex        =   5
      Top             =   210
      Width           =   1024
   End
   Begin VB.Line lnTop 
      BorderColor     =   &H00FFFFFF&
      Index           =   1
      X1              =   0
      X2              =   540
      Y1              =   15
      Y2              =   15
   End
   Begin VB.Line lnTop 
      BorderColor     =   &H00808080&
      Index           =   0
      X1              =   0
      X2              =   540
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Line lb3d 
      BorderColor     =   &H00808080&
      Index           =   1
      X1              =   45
      X2              =   8385
      Y1              =   2325
      Y2              =   2325
   End
   Begin VB.Line lb3d 
      BorderColor     =   &H00FFFFFF&
      Index           =   0
      X1              =   45
      X2              =   8385
      Y1              =   2340
      Y2              =   2340
   End
   Begin VB.Label lblFolder 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Folder:"
      Height          =   195
      Left            =   165
      TabIndex        =   9
      Top             =   825
      Width           =   480
   End
   Begin VB.Label lblName 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Name:"
      Height          =   195
      Left            =   165
      TabIndex        =   0
      Top             =   285
      Width           =   465
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
         Shortcut        =   ^Q
      End
   End
   Begin VB.Menu mnuAbout 
      Caption         =   "About"
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "#"
      Visible         =   0   'False
      Begin VB.Menu mnuOpen 
         Caption         =   "Open File"
      End
      Begin VB.Menu mnuOpenPath 
         Caption         =   "Open Path"
      End
   End
End
Attribute VB_Name = "frmmain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mStop As Boolean
Private mMouseButton As MouseButtonConstants

Private Sub GetDriveNames()
Dim Ld As Long
Dim Count As Integer

    Ld = GetLogicalDrives()
    
    For Count = 0 To 25
        If (Ld And 2 ^ Count) <> 0 Then
            'Load new menu array
            'Add the menus caption to the drive letter
            cboFolder.AddItem Chr(65 + Count) & ":\"
        End If
    Next Count
End Sub

Private Sub ResetBusyDisplay()
    'Update busy display
    With pBusy
        .Cls
        TransparentBlt .hdc, 0, 0, 32, 32, pSrc.hdc, 0, 0, 32, 32, RGB(255, 0, 255)
        .Refresh
    End With
End Sub

Private Sub cboFolder_KeyPress(KeyAscii As Integer)
    If (KeyAscii = 13) Then
        'Preform the serach.
        Call cmdOp_Click(0)
    End If
End Sub

Private Sub cmdOp_Click(Index As Integer)
Dim Fol As String

    Select Case Index
        Case 0
            Call ResetBusyDisplay
            tmrBusy.Enabled = True
            mStop = False
            'Do Serach
            LstResult.ListItems.Clear
            '
            Fol = cboFolder.Text
            If Len(Fol) = 0 Then
                Fol = FixPath(App.Path)
            End If
            
            Call Find(Fol, txtName.Text, chkSubFolder)
            '
            sBar1.Panels(1).Text = "Finished " & LstResult.ListItems.Count & " files found"
            tmrBusy.Enabled = False
        Case 1
            'Stop Serach
            tmrBusy.Enabled = False
            mStop = True
        Case 2
            'Reset serach
            chkSubFolder.Value = 1
            txtName.Text = ""
            sBar1.Panels(1).Text = ""
            sBar1.Panels(2).Text = ""
            LstResult.ListItems.Clear
            Call ResetBusyDisplay
        Case 3
            'Browse Folder
            Fol = FixPath(GetFolder(frmmain.hwnd, ""))
            If Len(Fol) > 1 Then
                cboFolder.Text = Fol
            End If
    End Select
    
    Call ResetBusyDisplay
End Sub

Private Sub Form_Load()
Dim sCmd As String

    Call GetDriveNames
    Call ResetBusyDisplay

    'Get Comamnd
    sCmd = Command$
    'Replace ""
    sCmd = Replace(sCmd, Chr(34), "")
    'Check if we have a string
    If Len(sCmd) Then
        'Check if it's a folder path
        If (GetAttr(sCmd) And vbDirectory) = vbDirectory Then
            cboFolder.Text = FixPath(sCmd)
        End If
    End If
End Sub

Private Sub Form_Resize()
On Error Resume Next
    'Resizeing codes
    cmdOp(0).Left = (frmmain.ScaleWidth - cmdOp(0).Width) - 128
    cmdOp(1).Left = cmdOp(0).Left
    cmdOp(2).Left = cmdOp(0).Left
    txtName.Width = (cmdOp(0).Left - cmdOp(0).Width) + 128
    cboFolder.Width = txtName.Width - cmdOp(3).Width
    cmdOp(3).Left = (cboFolder.Width + cmdOp(3).Width + 200)
    LstResult.Width = (frmmain.ScaleWidth - LstResult.Left)
    LstResult.Height = (frmmain.ScaleHeight - LstResult.Top - sBar1.Height)
    pBusy.Left = (cmdOp(0).Left + pBusy.ScaleWidth \ 2)
    lb3d(0).X2 = (frmmain.ScaleWidth - lb3d(0).X1)
    lb3d(1).X2 = (frmmain.ScaleWidth - lb3d(1).X1)
    lnTop(0).X2 = frmmain.ScaleWidth
    lnTop(1).X2 = frmmain.ScaleWidth
End Sub

Private Sub Form_Unload(Cancel As Integer)
    mStop = True
    Set frmmain = Nothing
    End
End Sub

Private Sub Find(ByVal lPath As String, ByVal SerachFor As String, Optional IncludeSubFolders As Boolean = True)
Dim Dirs As New Collection
Dim lFile As String
Dim mCurDir As String
Dim FileExt As String
On Error Resume Next

    'This is the main sub that does all the file searching
    lPath = FixPath(lPath)
    
    'Check if we need to serach in subfolders
    If (Not IncludeSubFolders) Then
        mCurDir = lPath
        lFile = Dir(mCurDir)
        Do Until Len(lFile) = 0
            If Not (GetAttr(mCurDir & lFile) And vbDirectory) = vbDirectory Then
                With LstResult.ListItems
                    'Add File Info
                    If InStr(1, lFile, SerachFor, vbTextCompare) Then
                        .Add , , lFile
                        .Item(.Count).SubItems(1) = mCurDir
                        .Item(.Count).SubItems(2) = FileLen(mCurDir & lFile)
                        .Item(.Count).SubItems(3) = FileDateTime(mCurDir & lFile)
                        .Item(.Count).SubItems(4) = GetFileExt(lFile)
                    End If
                    'Check for file patten eg *.exe
                    If Left(SerachFor, 2) = "*." Then
                        'Get File Ext
                        FileExt = GetFileExt(lFile)
                        If GetFileExt(SerachFor) = "*" Then SerachFor = vbNullString
                        'Compare FileExt and SerachFor
                        If StrComp(FileExt, GetFileExt(SerachFor), vbTextCompare) = 0 Then
                            'Add File info
                            .Add , , lFile
                            .Item(.Count).SubItems(1) = mCurDir
                            .Item(.Count).SubItems(2) = FileLen(mCurDir & lFile)
                            .Item(.Count).SubItems(3) = FileDateTime(mCurDir & lFile)
                            .Item(.Count).SubItems(4) = FileExt
                        End If
                    End If
                    'Update statusbar text
                    sBar1.Panels(1).Text = mCurDir & lFile
                    sBar1.Panels(2).Text = "Files: " & .Count
                End With
            End If
            lFile = Dir$
            DoEvents
        Loop
        Exit Sub
    End If
    
    'Serach subfolders
    Call Dirs.Add(lPath)
    
    While Dirs.Count
        
        mCurDir = Dirs(1)
        Call Dirs.Remove(1)
        
        lFile = Dir$(mCurDir, vbDirectory)
        
        Do Until (Len(lFile) = 0)
            If mStop Then Exit Do
            If (lFile <> ".") And (lFile <> "..") Then
                If (GetAttr(mCurDir & lFile) And vbDirectory) = vbDirectory Then
                    Dirs.Add mCurDir & lFile & "\"
                Else
                    With LstResult.ListItems
                        If InStr(1, lFile, SerachFor, vbTextCompare) Then
                            'Add File Info
                            .Add , , lFile
                            .Item(.Count).SubItems(1) = mCurDir
                            .Item(.Count).SubItems(2) = FileLen(mCurDir & lFile)
                            .Item(.Count).SubItems(3) = FileDateTime(mCurDir & lFile)
                            .Item(.Count).SubItems(4) = GetFileExt(lFile)
                        End If
                        'Check for file patten eg *.exe
                        If Left(SerachFor, 2) = "*." Then
                            'Get File Ext
                            FileExt = GetFileExt(lFile)
                            If GetFileExt(SerachFor) = "*" Then SerachFor = vbNullString
                            'Compare FileExt and SerachFor
                            If StrComp(FileExt, GetFileExt(SerachFor), vbTextCompare) = 0 Then
                                'Add File info
                                .Add , , lFile
                                .Item(.Count).SubItems(1) = mCurDir
                                .Item(.Count).SubItems(2) = FileLen(mCurDir & lFile)
                                .Item(.Count).SubItems(3) = FileDateTime(mCurDir & lFile)
                                .Item(.Count).SubItems(4) = FileExt
                            End If
                        End If
                        'Update statusbar text
                        sBar1.Panels(1).Text = mCurDir & lFile
                        sBar1.Panels(2).Text = "Files: " & .Count
                    End With
                End If
            End If
            'Get Next File
            lFile = Dir$
            'Let other OS things process
            DoEvents
        Loop
        DoEvents
    Wend
    
    If (mStop) Then
        sBar1.Panels(1).Text = "Finished " & LstResult.ListItems.Count & " files found"
    End If
    
End Sub

Private Sub LstResult_DblClick()
Dim lzFile As String
Dim Ret As Long
    'Check that we have items
    If (LstResult.ListItems.Count) Then
        If (mMouseButton = vbLeftButton) Then
            'Get Filename
            lzFile = LstResult.SelectedItem.SubItems(1) & LstResult.SelectedItem.Text
            Ret = RunApp(frmmain.hwnd, "open", lzFile)
            'Check if the file was opened.
            If (Ret = 2) Then
                MsgBox "There was an error opening the file:" & vbCrLf & vbCrLf & lzFile, vbExclamation, "File Not Found"
            End If
        End If
    End If
    
    lzFile = vbNullString
End Sub

Private Sub LstResult_ItemClick(ByVal Item As MSComctlLib.ListItem)
    If (mMouseButton = vbRightButton) Then
        PopupMenu mnuEdit
    End If
    
End Sub

Private Sub LstResult_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    mMouseButton = Button
End Sub

Private Sub mnuAbout_Click()
    MsgBox frmmain.Caption & " Ver 1.0" & vbCrLf & vbTab & "By DreamVB" _
    & vbCrLf & vbTab & vbTab & "Please vote if youu like this code.", vbInformation, "About"
    
End Sub

Private Sub mnuExit_Click()
    'Exit the program
    Call cmdOp_Click(1)
    Unload frmmain
End Sub

Private Sub mnuOpen_Click()
Dim Ret As Long
Dim lzFile As String

    'Get Filename
    lzFile = LstResult.SelectedItem.SubItems(1) & LstResult.SelectedItem.Text
    Ret = RunApp(frmmain.hwnd, "open", lzFile)
    'Check if the file was opened.
    If (Ret = 2) Then
        MsgBox "There was an error opening the file:" & vbCrLf & vbCrLf & lzFile, vbExclamation, "File Not Found"
    End If
End Sub

Private Sub mnuOpenPath_Click()
Dim Ret As Long
Dim lzFile As String
    'Get Folder
    lzFile = LstResult.SelectedItem.SubItems(1)
    Ret = RunApp(frmmain.hwnd, "open", lzFile)
    'Check if the file was opened.
    If (Ret = 2) Then
        MsgBox "There was an error opening the folder:" & vbCrLf & vbCrLf & lzFile, vbExclamation, "Folder Not Found"
    End If
End Sub

Private Sub tmrBusy_Timer()
Dim iFrames As Integer
Static iCount As Integer

    'Get the Number of frames
    iFrames = 4

    If (iCount >= iFrames) Then
        iCount = 0
    End If
    
    'DRaw the Busy Picture
    With pBusy
        .Cls
        TransparentBlt .hdc, 0, 0, 32, 32, pSrc.hdc, (32 * iCount), 0, 32, 32, vbMagenta
        .Refresh
    End With
    'INC Counter
    iCount = (iCount + 1)
End Sub

Private Sub txtName_KeyPress(KeyAscii As Integer)
    If (KeyAscii = 13) Then
        'Preform the serach.
        Call cmdOp_Click(0)
        'Remove the beep
        KeyAscii = 0
    End If
End Sub
