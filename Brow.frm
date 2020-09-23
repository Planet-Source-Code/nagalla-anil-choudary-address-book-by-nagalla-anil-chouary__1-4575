VERSION 5.00
Begin VB.Form brow 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Edit Browse"
   ClientHeight    =   3840
   ClientLeft      =   1140
   ClientTop       =   1515
   ClientWidth     =   6000
   Icon            =   "Brow.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3840
   ScaleWidth      =   6000
   Begin VB.CommandButton Command1 
      BackColor       =   &H00000000&
      Caption         =   "&EXIT"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4590
      TabIndex        =   5
      Top             =   3180
      Width           =   1125
   End
   Begin VB.DriveListBox Drive1 
      BackColor       =   &H00000000&
      ForeColor       =   &H00FF00FF&
      Height          =   315
      Left            =   450
      TabIndex        =   4
      Top             =   2970
      Width           =   1935
   End
   Begin VB.DirListBox Dir1 
      BackColor       =   &H00000000&
      ForeColor       =   &H000000FF&
      Height          =   2115
      Left            =   450
      MousePointer    =   1  'Arrow
      TabIndex        =   3
      Top             =   840
      Width           =   1875
   End
   Begin VB.FileListBox File1 
      BackColor       =   &H00000000&
      ForeColor       =   &H0000FF00&
      Height          =   2430
      Left            =   2520
      MousePointer    =   1  'Arrow
      Pattern         =   "*.bmp;*.wmf;*.dib;*.cur;*.ico"
      TabIndex        =   2
      Top             =   780
      Width           =   2025
   End
   Begin VB.CommandButton naccept 
      Caption         =   "&ACCEPT"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4590
      TabIndex        =   1
      Top             =   2700
      Width           =   1125
   End
   Begin VB.TextBox btext 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   1110
      TabIndex        =   0
      Top             =   300
      Width           =   4035
   End
   Begin VB.Label Label1 
      Caption         =   "PATH"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   345
      Left            =   360
      TabIndex        =   6
      Top             =   330
      Width           =   675
   End
End
Attribute VB_Name = "brow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Programmed and designed by
'  Nagalla Anil Choudary
'  D.K.Pallem
'  Bapatla-522 101
'  A.P, India
'  you can redistribute reproduce the source code as u like
'  but mail your comments/updations to anilfriend@hotmail.com
'
Private Sub Command1_Click()
Unload brow
End Sub

Private Sub Dir1_Change()
File1.Path = Dir1.Path
End Sub

Private Sub Drive1_Change()
On Error GoTo anil:
            Dir (Drive1.Drive)
            Dir1.Path = Drive1.Drive
anil:
    Exit Sub
End Sub


Private Sub File1_Click()
If Right(Dir1.Path, 1) Like "\" Then
            brow.btext.Text = Dir1.Path + File1.FileName
        Else
            brow.btext.Text = Dir1.Path + "\" + File1.FileName
        End If
End Sub

Private Sub naccept_Click()
           
       ppath = brow.btext.Text
       brow.Hide
       On Error GoTo anil
       address.Image3.Picture = LoadPicture(ppath)
anil:

End Sub


