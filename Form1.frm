VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.1#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   Caption         =   "Nilesh Ashra's Text Encryptor"
   ClientHeight    =   7395
   ClientLeft      =   270
   ClientTop       =   735
   ClientWidth     =   9345
   LinkTopic       =   "Form1"
   ScaleHeight     =   7395
   ScaleWidth      =   9345
   Begin VB.TextBox Text3 
      Alignment       =   2  'Center
      BackColor       =   &H80000008&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   285
      Left            =   4770
      Locked          =   -1  'True
      TabIndex        =   23
      Text            =   "Output Window"
      Top             =   330
      Width           =   4545
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      BackColor       =   &H80000008&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   285
      Left            =   60
      Locked          =   -1  'True
      TabIndex        =   22
      Text            =   "Input Window"
      Top             =   330
      Width           =   4680
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H80000008&
      ForeColor       =   &H0000FF00&
      Height          =   300
      Left            =   8820
      Locked          =   -1  'True
      TabIndex        =   20
      Text            =   "OK"
      Top             =   6105
      Width           =   480
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Encrypt"
      Height          =   300
      Left            =   3330
      TabIndex        =   17
      ToolTipText     =   "Encrypt Document"
      Top             =   6150
      Width           =   720
   End
   Begin VB.TextBox KeyText 
      BackColor       =   &H00C00000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   300
      IMEMode         =   3  'DISABLE
      Left            =   3660
      PasswordChar    =   "*"
      TabIndex        =   16
      ToolTipText     =   "Input your personal encryption key for this document here"
      Top             =   5805
      Width           =   2505
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Decrypt"
      Height          =   300
      Left            =   4050
      TabIndex        =   15
      ToolTipText     =   "Decrypt Document"
      Top             =   6150
      Width           =   720
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Save OUTPUT As."
      Height          =   300
      Left            =   4770
      TabIndex        =   14
      ToolTipText     =   "Save output as"
      Top             =   6150
      Width           =   1725
   End
   Begin VB.CommandButton Command11 
      Caption         =   "Show Message Window"
      Height          =   210
      Left            =   45
      TabIndex        =   13
      Top             =   6240
      Visible         =   0   'False
      Width           =   2340
   End
   Begin VB.CommandButton Command10 
      Caption         =   "Hide Message Window"
      Height          =   210
      Left            =   45
      TabIndex        =   12
      Top             =   6240
      Width           =   2340
   End
   Begin VB.CommandButton Command6 
      Caption         =   "PASTE into OUTPUT"
      Height          =   315
      Left            =   7410
      TabIndex        =   10
      Top             =   5100
      Width           =   1875
   End
   Begin VB.CommandButton Command8 
      Caption         =   "Clear"
      Height          =   315
      Left            =   1920
      TabIndex        =   9
      Top             =   5100
      Width           =   870
   End
   Begin VB.TextBox MsgDumpTitle 
      Enabled         =   0   'False
      Height          =   270
      Left            =   60
      Locked          =   -1  'True
      TabIndex        =   8
      Text            =   "--- Messages ---"
      Top             =   6540
      Width           =   9240
   End
   Begin VB.TextBox MsgDump 
      Height          =   555
      Left            =   60
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   7
      Top             =   6810
      Width           =   9240
   End
   Begin VB.CommandButton Command9 
      Caption         =   "Clear"
      Height          =   315
      Left            =   6630
      TabIndex        =   6
      Top             =   5100
      Width           =   780
   End
   Begin ComctlLib.ProgressBar ProgressBar1 
      Height          =   270
      Left            =   840
      TabIndex        =   4
      Top             =   45
      Width           =   8460
      _ExtentX        =   14923
      _ExtentY        =   476
      _Version        =   327682
      Appearance      =   1
   End
   Begin MSComDlg.CommonDialog CommonDialog2 
      Left            =   8790
      Top             =   -30
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   327680
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   8265
      Top             =   -30
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   327680
   End
   Begin VB.CommandButton Command5 
      Caption         =   "PASTE into INPUT box"
      Height          =   315
      Left            =   2805
      TabIndex        =   3
      Top             =   5100
      Width           =   1875
   End
   Begin VB.CommandButton Command4 
      Caption         =   "COPY Input Text"
      Height          =   315
      Left            =   30
      TabIndex        =   2
      Top             =   5100
      Width           =   1875
   End
   Begin VB.CommandButton Command3 
      Caption         =   "COPY Output Text"
      Height          =   315
      Left            =   4755
      TabIndex        =   1
      Top             =   5100
      Width           =   1875
   End
   Begin VB.TextBox InText 
      BackColor       =   &H80000008&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   4470
      Left            =   45
      MultiLine       =   -1  'True
      OLEDropMode     =   2  'Automatic
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      ToolTipText     =   "Input Window"
      Top             =   615
      Width           =   4695
   End
   Begin VB.TextBox OutText 
      BackColor       =   &H80000006&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   4455
      Left            =   4770
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   11
      ToolTipText     =   "Output Window"
      Top             =   615
      Width           =   4530
   End
   Begin VB.Label ShowLabel 
      Caption         =   "Scroll down for further messages"
      Height          =   210
      Left            =   90
      TabIndex        =   24
      Top             =   6015
      Width           =   2340
   End
   Begin VB.Label Label1 
      Caption         =   "Status:"
      Height          =   225
      Left            =   8265
      TabIndex        =   21
      Top             =   6120
      Width           =   525
   End
   Begin VB.Label Label8 
      Caption         =   "Please input key between 4 and 8 inclusive:"
      Height          =   210
      Left            =   3345
      TabIndex        =   19
      Top             =   5565
      Width           =   3135
   End
   Begin VB.Label Label9 
      Caption         =   "Key:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   3135
      TabIndex        =   18
      Top             =   5775
      Width           =   495
   End
   Begin VB.Line Line3 
      BorderColor     =   &H80000003&
      X1              =   15
      X2              =   9375
      Y1              =   6495
      Y2              =   6495
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000003&
      X1              =   15
      X2              =   9375
      Y1              =   5475
      Y2              =   5475
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000003&
      Index           =   0
      X1              =   0
      X2              =   9360
      Y1              =   15
      Y2              =   15
   End
   Begin VB.Label Label3 
      Caption         =   "Progress:"
      Height          =   210
      Left            =   135
      TabIndex        =   5
      Top             =   60
      Width           =   675
   End
   Begin VB.Menu FileM 
      Caption         =   "File"
      Index           =   1
      Begin VB.Menu OpenM 
         Caption         =   "Open"
         Shortcut        =   ^O
      End
      Begin VB.Menu SaveM 
         Caption         =   "Save Output As"
         Shortcut        =   ^S
      End
      Begin VB.Menu ExitM 
         Caption         =   "Quit"
         Shortcut        =   ^Q
      End
   End
   Begin VB.Menu EditM 
      Caption         =   "Edit"
      Begin VB.Menu ClearM 
         Caption         =   "Clear"
         Begin VB.Menu ClsInpM 
            Caption         =   "Clear Input Window"
         End
         Begin VB.Menu ClsOutM 
            Caption         =   "Clear Output Window"
         End
      End
      Begin VB.Menu CutM 
         Caption         =   "Cut"
         Shortcut        =   ^X
      End
      Begin VB.Menu CopyM 
         Caption         =   "Copy"
         Shortcut        =   ^C
      End
      Begin VB.Menu PasteM 
         Caption         =   "Paste"
         Begin VB.Menu InpM 
            Caption         =   "Into Input Window"
         End
         Begin VB.Menu OutM 
            Caption         =   "Into Output Window"
         End
      End
   End
   Begin VB.Menu TextM 
      Caption         =   "Text"
      Begin VB.Menu EncM 
         Caption         =   "Encrypt"
         Shortcut        =   ^E
      End
      Begin VB.Menu DecryptM 
         Caption         =   "Decrypt"
         Shortcut        =   ^D
      End
   End
   Begin VB.Menu HelpM 
      Caption         =   "Help"
      Begin VB.Menu HelpM2 
         Caption         =   "Help"
         Shortcut        =   {F1}
      End
      Begin VB.Menu BugM 
         Caption         =   "Bug Report"
      End
      Begin VB.Menu AboutM 
         Caption         =   "About"
         Shortcut        =   ^A
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub AboutM_Click()
frmAbout.Show
End Sub

Private Sub BugM_Click()
MsgBox ("If you have find a bug in my software, i will happily see to it that it is fixed and that you are sent a fix/update. Any reports send to nileshashra@btinternet.com with as many details of the problem as possible. Thank you :o)")
End Sub

Private Sub ClsInpM_Click()
InText = ""
End Sub

Private Sub ClsOutM_Click()
OutText = ""
End Sub

'################### Nilesh's Text Encryptor ################
'### This was written just to see if i could do it ##########
'### i think it does it's job pretty well, hope u like :o) ##
'############################################################

Private Sub Command1_Click()

showsavebox = True
ProgressBar1.Value = 0
'#############################
'##### CREATION OF KEYS ######
'#############################

Dim keyvar As String
Dim OkayFlag As Boolean
Dim ShowMsgLess As Boolean
Dim ShowMsgMore As Boolean
Dim LenOfKey As Integer
Dim KeyFinal As Integer

If Len(KeyText.Text) = 0 Then
    MsgBox ("Please enter a key")
    MsgDump.Text = MsgDump.Text & "Invalid key(1): No KEY given, unable to encrypt."
    MsgDump.Text = MsgDump.Text & Chr(13)
    MsgDump.Text = MsgDump.Text & Chr(10)
    Text1 = "Error"
    OkFlag = False
End If

keyvar = KeyText.Text
LenOfKey = Len(keyvar)

Select Case LenOfKey
  Case LenOfKey = 0, 1, 2
        ShowMsgLess = True: OkayFlag = False
  
  Case LenOfKey = 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21, 22, 23, 24, 25
        ShowMsgLess = True: OkayFlag = False
        
  Case LenOfKey = 3, 4, 5, 6, 7, 8
        OkayFlag = True
End Select

If ShowMsgLess = True Then
    MsgBox ("Invlaid Key")
    MsgDump.Text = MsgDump.Text & "Invalid Key(2): Key too short or too long, unable to encrypt."
    MsgDump.Text = MsgDump.Text & Chr(13)
    MsgDump.Text = MsgDump.Text & Chr(10)
    Text1 = "Error"
    OkFlag = False
End If

If OkayFlag = True Then

Text1 = "OK"

'----key gen TOTAL-----
    For x = 1 To Len(keyvar)
        currletter = Mid$(keyvar, x, 1)
        currletnum = Asc(currletter)
        total = currletnum + total
    Next x
'----key gen TOTAL-----
End If

If OkayFlag = True Then
    
    'GET CURRENT TIME
    StartTime = Timer
        
        KeyWorking = Int(total / (Len(keyvar)))
            
        Select Case KeyWorking
            Case 1 To 75
                KeyFinal = 5
            Case 75 To 85
                KeyFinal = 8
            Case 85 To 98
                KeyFinal = 9
            Case 98 To 101
                KeyFinal = 3
            Case 102
                KeyFinal = 11
            Case 103 To 104
                KeyFinal = 12
            Case 105
                KeyFinal = 6
            Case 106 To 107
                KeyFinal = 7
            Case 108 To 109
                KeyFinal = 10
            Case 110 To 115
                KeyFinal = 4
            Case Else
                KeyFinal = 15
        End Select
        
        '################################
        '##### CREATION OF KEYS END #####
        '################################
        
        
        Dim txt As String
        Dim CRPresent As Boolean
        
        
        'convert input text box to string: (now stored in memory)
            txt = InText.Text
        

        For Encloop = 1 To Len(txt)
        
          On Error GoTo errorLabel
          
          If ProgressBar1.Value < 99 Then ProgressBar1.Value = ProgressBar1.Value + (100 / Len(txt))
            
            
            CRPresent = True
            
            CurrentLetter = Mid$(txt, Encloop, 1)
            CurrentLetterVal = Asc(CurrentLetter)
            
            'Print CurrentLetterVal
            
            If CurrentLetterVal = 13 Then
                OutText.Text = OutText.Text & Chr$(174)
                CRPresent = False
            End If
            
            If CurrentLetterVal = 10 Then
                OutText.Text = OutText.Text & Chr$(165)
                CRPresent = False
            End If
            
            If CRPresent = True Then
                NewLetterVal = CurrentLetterVal + KeyFinal
                NewLetter$ = Chr$(NewLetterVal)
                OutText.Text = OutText.Text & NewLetter
            End If
        Next Encloop
        
        FinishTime = Timer
    
        totaltime = FinishTime - StartTime
        
        If totaltime > 1 Then
            TotalTimeDec = Mid(Str(totaltime), 1, 4)
            MsgDump.Text = MsgDump.Text & "Time taken to encrypt: " & TotalTimeDec & "s" & Chr$(13) & Chr$(10)
        End If
        
        If totaltime < 1 Then
            totalstring = Str(totaltime)
            Mid(totalstring, 1, 1) = ""
            TotalTimeDec = Mid(totalstring, 1, 4)
            MsgDump.Text = MsgDump.Text & "Time taken to encrypt: 0" & TotalTimeDec & "s" & Chr$(13) & Chr$(10)
        End If

End If



Exit Sub
errorLabel:   If err.Number = 5 Then
                MsgBox ("The text you typed is unencryptable, sorry")
                MsgDump.Text = MsgDump.Text & "Text encryption engine failed to activate; text uncodeable." & Chr$(13) & Chr$(10)
                Text1 = "Error"
            End If
            Text1 = "Error"
            ProgressBar1.Value = 100

End Sub


Private Sub Command10_Click()
Form1.Height = 7190
Command10.Visible = False
Command11.Visible = True
ShowLabel.Visible = False
End Sub

Private Sub Command11_Click()
Form1.Height = 8085
Command11.Visible = False
Command10.Visible = True
ShowLabel.Visible = True

End Sub

Private Sub Command2_Click()
'DECRYPTION
'#############################
'##### CREATION OF KEYS ######
'#############################

ProgressBar1.Value = 0

Dim keyvar As String
Dim OkayFlag As Boolean
Dim ShowMsgLess As Boolean
Dim ShowMsgMore As Boolean
Dim LenOfKey As Integer
Dim KeyFinal As Integer

keyvar = KeyText.Text
LenOfKey = Len(keyvar)

Select Case LenOfKey
  Case LenOfKey = 0, 1, 2
        ShowMsgLess = True: OkayFlag = False
  
  Case LenOfKey = 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21, 22, 23, 24, 25
        ShowMsgMore = True: OkayFlag = False
        
  Case LenOfKey = 3, 4, 5, 6, 7, 8
        OkayFlag = True
End Select

If Len(KeyText.Text) = 0 Then
    MsgBox ("Please enter a key")
    MsgDump.Text = MsgDump.Text & "Invalid key(1): No KEY given, unable to decrypt."
    MsgDump.Text = MsgDump.Text & Chr(13)
    MsgDump.Text = MsgDump.Text & Chr(10)
    Text1 = "Error"
    OkFlag = False
End If

If ShowMsgLess = True Then
    MsgBox ("Invlaid Key")
    MsgDump.Text = MsgDump.Text & "Invalid Key(2): Key too short or too long, unable to encrypt."
    MsgDump.Text = MsgDump.Text & Chr(13)
    MsgDump.Text = MsgDump.Text & Chr(10)
    Text1 = "Error"
    OkFlag = False
End If



If OkayFlag = True Then
Text1 = "OK"
    StartTime = Timer
'----key gen TOTAL-----
    For x = 1 To Len(keyvar)
        currletter = Mid$(keyvar, x, 1)
        currletnum = Asc(currletter)
        total = currletnum + total
    Next x
'----key gen TOTAL-----
End If

If OkayFlag = True Then

KeyWorking = Int(total / (Len(keyvar)))
    
Select Case KeyWorking
    Case 1 To 75
        KeyFinal = 5
    Case 75 To 85
        KeyFinal = 8
    Case 85 To 98
        KeyFinal = 9
    Case 98 To 101
        KeyFinal = 3
    Case 102
        KeyFinal = 11
    Case 103 To 104
        KeyFinal = 12
    Case 105
        KeyFinal = 6
    Case 106 To 107
        KeyFinal = 7
    Case 108 To 109
        KeyFinal = 10
    Case 110 To 115
        KeyFinal = 4
    Case Else
        KeyFinal = 15
End Select



'#############################
'##### CREATION OF KEYS ######
'#############################

Dim DeText As String
Dim CRPresent As Boolean
Dim DeOldCode As Integer
Dim Decode As Integer
'DEFINE DeText$

DeText$ = InText.Text

For m = 1 To Len(DeText$)

        On Error GoTo errorLabel
          
        If ProgressBar1.Value < 99 Then ProgressBar1.Value = ProgressBar1.Value + (100 / Len(DeText$))

    CRPresent = True
    
    Curr$ = Mid$(DeText$, m, 1)
    Decode = Asc(Curr$)
    
    If Decode = 174 Then
        OutText.Text = OutText.Text & Chr(13)
        CRPresent = False
    End If
    
    If Decode = 165 Then
        OutText.Text = OutText.Text & Chr(10)
        CRPresent = False
    End If
    
    If CRPresent = True Then
        DeOldCode = Decode - KeyFinal
        OutText.Text = OutText.Text & Chr(DeOldCode)
    End If


Next m
        FinishTime = Timer

        totaltime = FinishTime - StartTime
        
    If totaltime > 1 Then
        TotalTimeDec = Mid(Str(totaltime), 1, 4)
        MsgDump.Text = MsgDump.Text & "Time taken to decrypt: " & TotalTimeDec & "s" & Chr$(13) & Chr$(10)
    End If

    If totaltime < 1 Then
        totalstring = Str(totaltime)
        Mid(totalstring, 1, 1) = ""
        TotalTimeDec = Mid(totalstring, 1, 4)
        MsgDump.Text = MsgDump.Text & "Time taken to decrypt: 0" & TotalTimeDec & "s" & Chr$(13) & Chr$(10)
    End If
End If
Exit Sub
errorLabel:  If err.Number = 5 Then
                MsgBox ("The text you typed is undecryptable, sorry")
                MsgDump.Text = MsgDump.Text & "Text dencryption engine failed to activate; text undecodeable." & Chr$(13) & Chr$(10)
                Text1 = "Error"
            End If
            
            ProgressBar1.Value = 100
End Sub

Private Sub Command3_Click()
Clipboard.SetText OutText.Text
End Sub


Private Sub Command5_Click()
ClipText = Clipboard.GetText
InText = ClipText
End Sub

Private Sub Command6_Click()
ClipText = Clipboard.GetText
OutText = ClipText
End Sub

Private Sub Command7_Click()
CommonDialog1.Filter = "All Files (*.*)|*.*|Text File(*.txt)|*.txt|"
CommonDialog1.ShowSave
fname = CommonDialog1.filename
On Error GoTo err
Dim txt As String
txt = OutText.Text
Open fname For Output As #1
Print #1, txt
Close #1

err:
End Sub

Private Sub Command8_Click()
InText.Text = ""
End Sub

Private Sub Command9_Click()
OutText.Text = ""
End Sub

Private Sub CopyM_Click()
If InText.SelLength > 1 Then Clipboard.SetText InText.SelText
If OutText.SelLength > 1 Then Clipboard.SetText OutText.SelText
End Sub

Private Sub CutM_Click()
If InText.SelLength > 1 Then Clipboard.SetText InText.SelText: InText.Text = ""
If OutText.SelLength > 1 Then Clipboard.SetText OutText.SelText: OutText.Text = ""
End Sub

Private Sub DecryptM_Click()
Command2_Click
End Sub

Private Sub EncM_Click()
Command1_Click
End Sub

Private Sub ExitM_Click()
Unload frmAbout
Unload Form2
Unload frmSplash
End
End Sub

Private Sub HelpM2_Click()
Form2.Show
End Sub

Private Sub InpM_Click()
ClipText = Clipboard.GetText
InText.Text = InText.Text & ClipText
End Sub

Private Sub OpenM_Click()
On Error GoTo err
CommonDialog2.Filter = "Text Files (*.txt)|*.txt| HTML Documents (*.htm, *.html) |*.htm,*.html|All Files (*.*)|*.*|"
CommonDialog2.ShowOpen

Open CommonDialog2.filename For Input As #1

Do While Not EOF(1)
    Input #1, TxtStream
    InText.Text = InText.Text & TxtStream
Loop
Close #1
Exit Sub
err: MsgDump.Text = MsgDump.Text & "File access error, please try again, or another file. Also ensure that only text or html files are opened." & Chr$(13) & Chr$(10)
        MsgBox ("Path/File Acces Error. Only TEXT or HTM(L) files supported currently by NACrypt")
End Sub

Private Sub OutM_Click()
ClipText = Clipboard.GetText
OutText.Text = OutText.Text & ClipText
End Sub


Private Sub SaveM_Click()
CommonDialog1.Filter = "All Files (*.*)|*.*|Text File(*.txt)|*.txt|"
CommonDialog1.ShowSave
fname = CommonDialog1.filename
On Error GoTo err
Dim txt As String
txt = OutText.Text
Open fname For Output As #1
Print #1, txt
Close #1
Exit Sub
err: MsgBox ("An error occured when you tried to save. Please ensure drive or file is not write protected.")
    MsgDump.Text = MsgDump.Text & "Error occured when saving document." & Chr(13) & Chr(10)
End Sub


