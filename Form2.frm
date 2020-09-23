VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "Instructions"
   ClientHeight    =   5115
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9135
   LinkTopic       =   "Form2"
   ScaleHeight     =   5115
   ScaleWidth      =   9135
   StartUpPosition =   3  'Windows Default
   Begin VB.Label Label11 
      Caption         =   "For Decyrption:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   345
      Left            =   180
      TabIndex        =   10
      Top             =   3120
      Width           =   1995
   End
   Begin VB.Label Label10 
      Caption         =   $"Form2.frx":0000
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   345
      TabIndex        =   9
      Top             =   3780
      Width           =   8505
   End
   Begin VB.Label Label9 
      Caption         =   "1. Input text (by copy/pasting from a text viewer) to be encrypted to the box marked ""INPUT text:"""
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   345
      TabIndex        =   8
      Top             =   3465
      Width           =   8490
   End
   Begin VB.Label Label8 
      Caption         =   "3. Input the key that you encrypted with into the box marked ""Key""."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   345
      TabIndex        =   7
      Top             =   4305
      Width           =   8625
   End
   Begin VB.Label Label7 
      Caption         =   "4. To decrypt the text hit the ""Decrypt"" button."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   345
      TabIndex        =   6
      Top             =   4590
      Width           =   8475
   End
   Begin VB.Label Label6 
      Caption         =   "For Encyrption:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   345
      Left            =   165
      TabIndex        =   5
      Top             =   765
      Width           =   1995
   End
   Begin VB.Label Label5 
      Caption         =   $"Form2.frx":0087
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   330
      TabIndex        =   4
      Top             =   1605
      Width           =   8505
   End
   Begin VB.Label Label2 
      Caption         =   "INSTRUCTIONS"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   495
      Left            =   2610
      TabIndex        =   3
      Top             =   45
      Width           =   3660
   End
   Begin VB.Label Label1 
      Caption         =   "1. Input text (by typing or copy/pasting from a text viewer) to be encrypted to the box marked ""INPUT text:"""
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   330
      TabIndex        =   2
      Top             =   1110
      Width           =   8490
   End
   Begin VB.Label Label3 
      Caption         =   "3. Input a TEXT key into the box marked ""Key"", REMEMBER this as you will need it to decrypt the text when you need it."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Left            =   330
      TabIndex        =   1
      Top             =   2115
      Width           =   8625
   End
   Begin VB.Label Label4 
      Caption         =   "4. To encrypt the text hit the ""Encrypt"" button."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   330
      TabIndex        =   0
      Top             =   2640
      Width           =   8475
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
