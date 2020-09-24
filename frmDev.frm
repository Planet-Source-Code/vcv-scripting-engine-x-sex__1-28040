VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmSexIDE 
   Caption         =   "SEX Editor 1.0"
   ClientHeight    =   5670
   ClientLeft      =   3750
   ClientTop       =   4290
   ClientWidth     =   8385
   Icon            =   "frmDev.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   378
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   559
   Begin RichTextLib.RichTextBox txtCode 
      Height          =   5685
      Left            =   -15
      TabIndex        =   0
      Top             =   -15
      Width           =   8400
      _ExtentX        =   14817
      _ExtentY        =   10028
      _Version        =   393217
      HideSelection   =   0   'False
      ScrollBars      =   2
      TextRTF         =   $"frmDev.frx":1042
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "IBMPC"
         Size            =   9
         Charset         =   255
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComDlg.CommonDialog cmDialog 
      Left            =   7740
      Top             =   105
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
   End
   Begin VB.Menu mnu_File 
      Caption         =   "&File"
      Begin VB.Menu mnu_File_New 
         Caption         =   "&New"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnu_File_Open 
         Caption         =   "&Open..."
         Shortcut        =   ^O
      End
      Begin VB.Menu mnu_File_Save 
         Caption         =   "&Save"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnu_File_SaveAs 
         Caption         =   "Save &As..."
      End
      Begin VB.Menu mnu_File_LB01 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_File_Print 
         Caption         =   "&Print"
         Shortcut        =   ^P
      End
      Begin VB.Menu mnu_File_LB02 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_File_Exit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu frm_Edit 
      Caption         =   "&Edit"
      Begin VB.Menu mnu_Edit_Undo 
         Caption         =   "&Undo"
         Shortcut        =   ^Z
      End
      Begin VB.Menu mnu_Edit_LB01 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_Edit_Cut 
         Caption         =   "Cu&t"
         Shortcut        =   ^X
      End
      Begin VB.Menu mnu_Edit_Copy 
         Caption         =   "&Copy"
         Shortcut        =   ^C
      End
      Begin VB.Menu mnu_Edit_Paste 
         Caption         =   "&Paste"
         Shortcut        =   ^V
      End
      Begin VB.Menu mnu_Edit_Delete 
         Caption         =   "&Delete"
         Shortcut        =   {DEL}
      End
      Begin VB.Menu mnu_Edit_LB02 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_Edit_Find 
         Caption         =   "&Find"
         Shortcut        =   ^F
      End
      Begin VB.Menu mnu_Edit_FindNext 
         Caption         =   "Find &Next"
         Shortcut        =   {F3}
      End
      Begin VB.Menu mnu_Edit_LB03 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_Edit_SelectAll 
         Caption         =   "Select &All"
         Shortcut        =   ^A
      End
   End
   Begin VB.Menu mnu_Run 
      Caption         =   "&Run"
      Begin VB.Menu mnu_Run_Start 
         Caption         =   "&Start"
         Enabled         =   0   'False
         Shortcut        =   {F5}
      End
      Begin VB.Menu mnu_Run_Break 
         Caption         =   "&Break"
         Enabled         =   0   'False
         Shortcut        =   +^{F5}
      End
   End
End
Attribute VB_Name = "frmSexIDE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private bHasChanged     As Boolean
Private strFileName     As String

Private Sub Command2_Click()
MsgBox vbInformation
End Sub


Sub OpenFile(strFileName As String)
    On Error GoTo noLoad
    Dim strFileData As String
    
    Open strFileName For Binary As #1
        strFileData = String(LOF(1), 0)
        Get #1, 1, strFileData
    Close #1
    
    mnu_Run_Start.Enabled = True
    txtCode.Text = strFileData
    strFileData = ""
'    Delete strFileData
    
    Exit Sub
noLoad:

End Sub

Sub SaveFile(strFileName As String)
    
    
    On Error GoTo noSaveAs
    Me.MousePointer = 11
    Open strFileName For Output As #1
        Print #1, txtCode.Text
    Close #1
    Me.MousePointer = 0
    mnu_Run_Start.Enabled = True
    Exit Sub
    
noSaveAs:
    Me.MousePointer = 0
    MsgBox "An error has occured while trying to save the file, [" & Err & "]:" & vbCrLf & vbCrLf & Error, vbExclamation

End Sub

Private Sub Form_Load()
    txtCode.Text = _
        "alias main" & vbCrLf & _
        "   ; main body code goes here" & vbCrLf & _
        "   " & vbCrLf & _
        "end alias"
        
    txtCode.SelStart = 0
    txtCode.SelLength = 5
    txtCode.SelBold = True
    txtCode.SelStart = 48
    txtCode.SelLength = 9
    txtCode.SelBold = True
    txtCode.SelStart = 46
End Sub

Private Sub Form_Resize()
    txtCode.Move 0, 0, Me.ScaleWidth, Me.ScaleHeight
End Sub


Private Sub mnu_Edit_Copy_Click()
    Clipboard.SetText txtCode.SelText
End Sub

Private Sub mnu_Edit_Cut_Click()
    Clipboard.SetText txtCode.SelText
    txtCode.SelText = ""
End Sub

Private Sub mnu_Edit_Delete_Click()
    If txtCode.SelLength > 0 Then
        txtCode.SelText = ""
    Else
        txtCode.SelLength = 1
        txtCode.SelText = ""
    End If
End Sub

Private Sub mnu_Edit_Paste_Click()
    txtCode.SelText = Clipboard.GetText()
End Sub

Private Sub mnu_Edit_SelectAll_Click()
    txtCode.SelStart = 0
    txtCode.SelLength = Len(txtCode.Text)
End Sub

Private Sub mnu_File_New_Click()
    If bHasChanged = False Then
        txtCode.Text = ""
        strFileName = ""
        mnu_Run_Start.Enabled = False
    End If
End Sub

Private Sub mnu_File_Open_Click()
    If bHasChanged Then
    
    Else
        On Error GoTo noopen
        cmDialog.DialogTitle = "Open Script"
        cmDialog.Filter = "SEX Script (*.sex)|*.sex|"
        cmDialog.FilterIndex = 0
        cmDialog.ShowOpen
        strFileName = cmDialog.FileName
        Me.MousePointer = 11
        OpenFile strFileName
        Me.MousePointer = 0
        Exit Sub
noopen:
        If Err = 32755 Then Exit Sub
        MsgBox "An error has occured while trying to access the file [" & Err & "]:" & vbCrLf & vbCrLf & Error, vbExclamation
    End If
End Sub

Private Sub mnu_File_Save_Click()
    If strFileName = "" Then
        Call mnu_File_SaveAs_Click
    Else
        SaveFile strFileName
    End If
    mnu_File_Save.Enabled = False
End Sub

Private Sub mnu_File_SaveAs_Click()
    On Error GoTo noSaveAs
    cmDialog.DialogTitle = "Save Script"
    cmDialog.Filter = "SEX Script (*.sex)|*.sex|"
    cmDialog.FilterIndex = 0
    cmDialog.FileName = strFileName
    cmDialog.ShowSave
    strFileName = cmDialog.FileName
    SaveFile strFileName
    mnu_File_Save.Enabled = False
    
    Exit Sub
noSaveAs:
    If Err = 32755 Then Exit Sub
    MsgBox "An error has occured while trying to access the file [" & Err & "]:" & vbCrLf & vbCrLf & Error, vbExclamation

End Sub


Private Sub mnu_Run_Break_Click()
    Stop
End Sub


Private Sub mnu_Run_Start_Click()
    On Error GoTo errorHandler
    
    Me.MousePointer = 11
    mnu_Run_Break.Enabled = True
    Dim engine  As New clsSSE_Main
    engine.NewScript
    engine.LoadScript engine.ScriptCount, strFileName
    Dim paramlist(0)
    engine.ExecuteAlias "main", paramlist()
    Me.MousePointer = 0
    Exit Sub
errorHandler:
    MsgBox "An error has occured while trying to run the script, [" & Err & "]:" & vbCrLf & vbCrLf & Error
    
End Sub


Private Sub txtCode_Change()
    bHasChanged = True
    mnu_File_Save.Enabled = True
End Sub


