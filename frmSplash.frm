VERSION 5.00
Begin VB.Form frmSplash 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   3240
   ClientLeft      =   2415
   ClientTop       =   1020
   ClientWidth     =   5865
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmSplash.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3240
   ScaleWidth      =   5865
   ShowInTaskbar   =   0   'False
   Begin VB.Line Line2 
      X1              =   5805
      X2              =   0
      Y1              =   2715
      Y2              =   2715
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   5865
      Y1              =   765
      Y2              =   765
   End
   Begin VB.Label Label1 
      Caption         =   "click anywhere"
      Height          =   225
      Left            =   105
      TabIndex        =   6
      Top             =   2985
      Width           =   1080
   End
   Begin VB.Label lblVersion 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   5040
      TabIndex        =   5
      Top             =   2850
      Width           =   60
   End
   Begin VB.Label lblProductName 
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   21.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   90
      TabIndex        =   4
      Top             =   105
      Width           =   120
   End
   Begin VB.Label lblCompanyProduct 
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   675
      TabIndex        =   3
      Top             =   210
      Width           =   90
   End
   Begin VB.Label lblLicenseTo 
      Caption         =   "LicenseTo: MANKIND"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3540
      TabIndex        =   2
      Top             =   2490
      Width           =   1560
   End
   Begin VB.Image Image1 
      Height          =   885
      Left            =   630
      Picture         =   "frmSplash.frx":000C
      Top             =   915
      Width           =   4470
   End
   Begin VB.Label lblCopyright 
      Caption         =   "Copyright 2001"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3990
      TabIndex        =   1
      Top             =   2250
      Width           =   1110
   End
   Begin VB.Label lblPlatform 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Windows 95/98/ME"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   2265
      TabIndex        =   0
      Top             =   1890
      Width           =   2835
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Click()
Unload frmSplash
Form1.Show
End Sub

Private Sub Form_Load()
    lblVersion.Caption = "Version :" & App.Major & "." & App.Minor & "." & App.Revision
    lblProductName.Caption = App.Title
End Sub
