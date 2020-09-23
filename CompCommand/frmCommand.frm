VERSION 5.00
Object = "{4E3D9D11-0C63-11D1-8BFB-0060081841DE}#1.0#0"; "Xlisten.dll"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmCommand 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "CompCommand v1.0"
   ClientHeight    =   5940
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5535
   Icon            =   "frmCommand.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5940
   ScaleWidth      =   5535
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkIntro 
      Caption         =   "Show introduction form"
      Height          =   240
      Left            =   150
      TabIndex        =   13
      Top             =   5625
      Value           =   1  'Checked
      Width           =   5265
   End
   Begin VB.CommandButton cmdStartup 
      Caption         =   "Startup..."
      Height          =   390
      Left            =   4350
      TabIndex        =   11
      Top             =   5175
      Width           =   1065
   End
   Begin MSComDlg.CommonDialog dlg 
      Left            =   2175
      Top             =   3825
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
      Filter          =   "EXE Files|*.exe|All Files|*.*"
   End
   Begin ACTIVELISTENPROJECTLibCtl.DirectSR Speech 
      Height          =   465
      Left            =   2625
      OleObjectBlob   =   "frmCommand.frx":044A
      TabIndex        =   9
      Top             =   3825
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.ListBox lstPaths 
      Height          =   450
      Left            =   975
      TabIndex        =   8
      Top             =   1350
      Visible         =   0   'False
      Width           =   765
   End
   Begin VB.CommandButton cmdBrowse 
      Caption         =   "Browse..."
      Height          =   390
      Left            =   2850
      TabIndex        =   5
      Top             =   5175
      Width           =   990
   End
   Begin VB.TextBox txtPath 
      Height          =   390
      Left            =   900
      TabIndex        =   4
      Top             =   4650
      Width           =   4515
   End
   Begin VB.TextBox txtName 
      Height          =   390
      Left            =   900
      TabIndex        =   3
      Top             =   4200
      Width           =   4515
   End
   Begin VB.CommandButton cmdDel 
      Caption         =   "Delete..."
      Height          =   390
      Left            =   1800
      TabIndex        =   2
      Top             =   5175
      Width           =   990
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "Add/Modify..."
      Height          =   390
      Left            =   150
      TabIndex        =   1
      Top             =   5175
      Width           =   1590
   End
   Begin VB.ListBox lstNames 
      Columns         =   3
      Height          =   2595
      Left            =   150
      TabIndex        =   0
      Top             =   1425
      Width           =   5265
   End
   Begin VB.Label lblCmd 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "<Current command>"
      Height          =   315
      Left            =   150
      TabIndex        =   12
      Top             =   150
      Width           =   5265
   End
   Begin VB.Label lblInfo 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   $"frmCommand.frx":046E
      Height          =   840
      Left            =   150
      TabIndex        =   10
      Top             =   525
      Width           =   5265
   End
   Begin VB.Label lblPath 
      Caption         =   "Path :"
      Height          =   390
      Left            =   150
      TabIndex        =   7
      Top             =   4650
      Width           =   690
   End
   Begin VB.Label lblName 
      Caption         =   "Name :"
      Height          =   390
      Left            =   150
      TabIndex        =   6
      Top             =   4200
      Width           =   690
   End
   Begin VB.Menu mnu_1 
      Caption         =   "mnu_1"
      Visible         =   0   'False
      Begin VB.Menu mnuShow 
         Caption         =   "Show"
      End
      Begin VB.Menu mnuEnable 
         Caption         =   "Enable"
      End
      Begin VB.Menu mnuDisable 
         Caption         =   "Disable"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
      End
   End
End
Attribute VB_Name = "frmCommand"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim EngineEnable As Boolean

Private Sub chkIntro_Click()
    If chkIntro.Value = vbUnchecked Then
        Open App.Path & "\nointro" For Output As #2
            Write #2, "NoIntro"
        Close #2
    Else
        Kill App.Path & "\nointro"
    End If
End Sub

Private Sub cmdAdd_Click()
    Dim I As Long
    Dim Updated As Boolean
    'check to see if TEXT item already exists...
    
    If Trim(txtName) = "" Then
        Exit Sub
    End If
    
    For I = 0 To lstNames.ListCount - 1
        If LCase(lstNames.List(I)) = LCase(txtName) Then
            'modify...
            lstNames.List(I) = txtName
            lstPaths.List(I) = txtPath
            Updated = True
        End If
    Next I
    
    If Not Updated Then
        'create new
        lstNames.AddItem txtName
        lstPaths.AddItem txtPath
    End If
    
    txtName = "": txtPath = ""
    SaveALL
    InitEngine
End Sub

Private Sub cmdBrowse_Click()
    On Error Resume Next
    
    dlg.ShowOpen
    
    If Err.Number <> 0 Then
        Err.Clear
        Exit Sub
    End If
    
    txtPath = dlg.filename
End Sub

Private Sub cmdDel_Click()
    If lstNames.ListIndex <> -1 Then
        txtName = "": txtPath = ""
        lstPaths.RemoveItem lstNames.ListIndex
        lstNames.RemoveItem lstNames.ListIndex
    End If
    
    SaveALL
    InitEngine
End Sub

Private Sub cmdStartup_Click()
    Dim MBResults As VbMsgBoxResult
    Dim AppExeName As String
    
    MBResults = MsgBox("If you want, you can set CompCommand to auto-start when windows starts, allowing you to tell your commands whitout having to open manually the program! It requires very low ressources. Would you like to have CompRequest started at Windows startup?", vbQuestion + vbYesNo)

    If MBResults = vbYes Then
        AppExeName = App.Path & IIf(Right(App.Path, 1) = "\", "", "\") & App.EXEName
        savestring HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows\CurrentVersion\Run", "CompRequest", AppExeName
    Else
        DeleteValue HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows\CurrentVersion\Run", "CompRequest"
    End If
End Sub

Private Sub Form_Load()
    If App.PrevInstance Then
        MsgBox "There is already an instance of CompCommand running on your system! Just click on the MIC icon near the clock (in the system tray).", vbExclamation
        End
    End If
    
    With nid
        .cbSize = Len(nid)
        .hwnd = Me.hwnd
        .uId = vbNull
        .uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
        .uCallBackMessage = WM_MOUSEMOVE
        .hIcon = Me.Icon
        .szTip = Caption & vbNullChar
    End With
    Shell_NotifyIcon NIM_ADD, nid
    
    EngineEnable = True
    
    LoadALL
    InitEngine
    
    On Error Resume Next
    Open App.Path & "\nointro" For Input As #2
    Close #2
    
    If Err.Number = 53 Then
        chkIntro.Value = vbChecked
        Err.Clear
    Else
        chkIntro.Value = vbUnchecked
    End If
      
    'hide the APP
    Hide
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    'This procedure receives the callbacks f
    '     rom the System Tray icon.
    Dim Result As Long
    Dim MSG As Long
    'The value of X will vary depending upon
    '     the scalemode setting


    If Me.ScaleMode = vbPixels Then
        MSG = x
    Else
        MSG = x / Screen.TwipsPerPixelX
    End If


    Select Case MSG
        Case WM_LBUTTONUP '514 restore form window
        Me.WindowState = vbNormal
        Result = SetForegroundWindow(Me.hwnd)
        Me.Show
        Case WM_LBUTTONDBLCLK '515 restore form window
        Me.WindowState = vbNormal
        Result = SetForegroundWindow(Me.hwnd)
        Me.Show
        Case WM_RBUTTONUP '517 display popup menu
        Result = SetForegroundWindow(Me.hwnd)
        '***** STOP! and make sure that your fir
        '     st menu item
        ' is named "mnu_1", otherwise you will g
        '     et an erro below!!! *******
        
        If EngineEnable Then
            mnuDisable.Visible = True
            mnuEnable.Visible = False
        Else
            mnuDisable.Visible = False
            mnuEnable.Visible = True
        End If
        
        Me.PopupMenu Me.mnu_1
    End Select
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode = 0 Then
        If MsgBox("Instead of closing the program, it is recommended that you minimize it (it will send it to the systray). Would you like to close (and stop listening commands)?", vbQuestion + vbYesNo) = vbNo Then
            Cancel = 1
        End If
    End If
End Sub

Private Sub Form_Resize()
    If WindowState = vbMinimized Then
        WindowState = vbNormal
        Hide
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Shell_NotifyIcon NIM_DELETE, nid
End Sub

Private Sub lstNames_Click()
    txtName = lstNames.List(lstNames.ListIndex)
    txtPath = lstPaths.List(lstNames.ListIndex)
End Sub

Private Sub mnuDisable_Click()
    EngineEnable = False
    Speech.Deactivate
End Sub

Private Sub mnuEnable_Click()
    EngineEnable = True
    Speech.GrammarFromFile App.Path & "\text_cmd.ini"
    Speech.Activate
End Sub

Private Sub mnuExit_Click()
    Unload Me
End Sub

Sub SaveALL()
    Dim I As Long

    Open App.Path & "\prgm_cmd.ini" For Output As #1
        For I = 0 To lstNames.ListCount - 1
            Write #1, lstNames.List(I), lstPaths.List(I)
        Next I
    Close #1
    
    'for MICROSOFT compatibility...
    Open App.Path & "\text_cmd.ini" For Output As #1
        Print #1, "[Grammer]"
        Print #1, ""
        Print #1, "Type=Cfg"
        Print #1, ""
        Print #1, "[<Start>]"
        For I = 0 To lstNames.ListCount - 1
            Print #1, "<Start>=" & lstNames.List(I)
        Next I
    Close #1
End Sub

Sub LoadALL()
    Dim I As Long
    Dim strdata As String, strData2 As String

    On Error Resume Next
    
    Open App.Path & "\prgm_cmd.ini" For Input As #1
        If Err.Number <> 0 Then
            MsgBox "Error while loading from config file -- aborting!", vbCritical
            Err.Clear
            Exit Sub
        End If
    
        Do Until EOF(1)
            Input #1, strdata, strData2
            lstNames.AddItem strdata
            lstPaths.AddItem strData2
        Loop
    Close #1
End Sub

Sub InitEngine()
    If EngineEnable Then
        Speech.Deactivate
        Speech.GrammarFromFile App.Path & "\text_cmd.ini"
        Speech.Activate
    End If
End Sub

Private Sub mnuShow_Click()
    Dim Result

    Me.WindowState = vbNormal
    Result = SetForegroundWindow(Me.hwnd)
    Me.Show
End Sub

Private Sub Speech_PhraseFinish(ByVal flags As Long, ByVal beginhi As Long, ByVal beginlo As Long, ByVal endhi As Long, ByVal endlo As Long, ByVal Phrase As String, ByVal parsed As String, ByVal results As Long)
    Dim I As Integer
    Dim r

    If Trim(Phrase) <> "" Then
        lblCmd = Phrase
        For I = 0 To lstNames.ListCount - 1
            If Phrase = lstNames.List(I) Then
                'attempt to SHELL the app...
                r = ShellExecute(hwnd, "", lstPaths.List(I), "", "", 1)
                
                If r <= 32 Then
                    MsgBox "There was an error while attempting to run the application associated with " & Phrase & "!", vbCritical
                    Err.Clear
                    Exit Sub
                End If
            End If
        Next I
    Else
        lblCmd = "The last command you entered was not understood!"
    End If
End Sub
