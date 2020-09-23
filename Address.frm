VERSION 5.00
Begin VB.Form address 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Anil's    Address Book   ------- Nagalla Anil Choudary"
   ClientHeight    =   5535
   ClientLeft      =   1050
   ClientTop       =   1125
   ClientWidth     =   7365
   Icon            =   "Address.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   5535
   ScaleWidth      =   7365
   Begin VB.CommandButton Command11 
      Caption         =   "new/clear"
      Height          =   285
      Left            =   5310
      TabIndex        =   40
      Top             =   4860
      Width           =   975
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   750
      TabIndex        =   1
      Top             =   810
      Width           =   4005
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   3000
      TabIndex        =   2
      Top             =   1200
      Width           =   2565
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   930
      Sorted          =   -1  'True
      TabIndex        =   3
      Text            =   "Friend"
      Top             =   1590
      Width           =   1695
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   3225
      TabIndex        =   4
      Top             =   1620
      Width           =   3750
   End
   Begin VB.TextBox Text4 
      Height          =   285
      Left            =   705
      TabIndex        =   5
      Top             =   2040
      Width           =   1860
   End
   Begin VB.TextBox Text5 
      Height          =   285
      Left            =   3000
      TabIndex        =   6
      Top             =   2040
      Width           =   1680
   End
   Begin VB.TextBox Text6 
      Height          =   285
      Left            =   5430
      TabIndex        =   7
      Top             =   2040
      Width           =   1545
   End
   Begin VB.TextBox Text7 
      Height          =   285
      Left            =   705
      ScrollBars      =   1  'Horizontal
      TabIndex        =   8
      Top             =   2460
      Width           =   2445
   End
   Begin VB.TextBox Text8 
      Height          =   285
      Left            =   3900
      TabIndex        =   9
      Top             =   2460
      Width           =   3075
   End
   Begin VB.TextBox Text9 
      Height          =   285
      Left            =   840
      TabIndex        =   10
      Top             =   2880
      Width           =   2310
   End
   Begin VB.TextBox Text10 
      Height          =   285
      Left            =   3630
      TabIndex        =   11
      Top             =   2880
      Width           =   3345
   End
   Begin VB.TextBox Text11 
      Height          =   285
      Left            =   705
      TabIndex        =   12
      Top             =   3300
      Width           =   2445
   End
   Begin VB.TextBox Text12 
      Height          =   285
      Left            =   3675
      TabIndex        =   13
      Top             =   3300
      Width           =   3300
   End
   Begin VB.TextBox Text13 
      Height          =   1125
      Left            =   660
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   14
      Top             =   3690
      Width           =   6225
   End
   Begin VB.CommandButton Command2 
      Caption         =   "l <"
      Height          =   285
      Left            =   1050
      TabIndex        =   21
      Top             =   4860
      Width           =   1050
   End
   Begin VB.CommandButton Command3 
      Caption         =   "< <"
      Height          =   285
      Left            =   2085
      TabIndex        =   22
      Top             =   4860
      Width           =   1095
   End
   Begin VB.CommandButton Command4 
      Caption         =   " > >"
      Height          =   285
      Left            =   3180
      TabIndex        =   23
      Top             =   4860
      Width           =   1095
   End
   Begin VB.CommandButton Command5 
      Caption         =   "> l"
      Height          =   285
      Left            =   4245
      TabIndex        =   24
      Top             =   4860
      Width           =   1050
   End
   Begin VB.CommandButton Command10 
      Caption         =   "&Update"
      Height          =   285
      Left            =   5985
      TabIndex        =   20
      Top             =   5160
      Width           =   1005
   End
   Begin VB.CommandButton Command9 
      Caption         =   "&Exit"
      Height          =   285
      Left            =   4860
      TabIndex        =   19
      Top             =   5160
      Width           =   1050
   End
   Begin VB.CommandButton Command8 
      Caption         =   "&Sort"
      Height          =   285
      Left            =   3720
      TabIndex        =   18
      Top             =   5160
      Width           =   1095
   End
   Begin VB.CommandButton Command7 
      Caption         =   "&Find"
      Height          =   285
      Left            =   2520
      TabIndex        =   17
      Top             =   5160
      Width           =   1095
   End
   Begin VB.CommandButton Command6 
      Caption         =   "&Delete"
      Height          =   285
      Left            =   1395
      TabIndex        =   16
      Top             =   5160
      Width           =   1050
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Add"
      Height          =   285
      Left            =   300
      TabIndex        =   15
      Top             =   5160
      Width           =   1050
   End
   Begin VB.Frame Frame1 
      Caption         =   "Image"
      Height          =   1515
      Left            =   5700
      TabIndex        =   0
      Top             =   30
      Width           =   1575
      Begin VB.Image Image3 
         Height          =   1065
         Left            =   180
         Picture         =   "Address.frx":030A
         Stretch         =   -1  'True
         Top             =   300
         Width           =   1245
      End
   End
   Begin VB.Image Image1 
      Height          =   585
      Left            =   1710
      Picture         =   "Address.frx":27C88
      Top             =   0
      Width           =   2805
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "NAME"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   0
      TabIndex        =   39
      Top             =   540
      Width           =   660
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "First"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   270
      TabIndex        =   38
      Top             =   870
      Width           =   375
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Second"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   2250
      TabIndex        =   37
      Top             =   1260
      Width           =   660
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Relation"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   165
      TabIndex        =   36
      Top             =   1680
      Width           =   720
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Place"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   2685
      TabIndex        =   35
      Top             =   1740
      Width           =   495
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "Town"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   165
      TabIndex        =   34
      Top             =   2070
      Width           =   480
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "Pin"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   2640
      TabIndex        =   33
      Top             =   2100
      Width           =   285
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "District"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   4755
      TabIndex        =   32
      Top             =   2070
      Width           =   615
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      Caption         =   "State"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   165
      TabIndex        =   31
      Top             =   2520
      Width           =   465
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      Caption         =   "Country"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   3180
      TabIndex        =   30
      Top             =   2550
      Width           =   660
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      Caption         =   "Tel.No"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   165
      TabIndex        =   29
      Top             =   2910
      Width           =   585
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      Caption         =   "Fax"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   3225
      TabIndex        =   28
      Top             =   2910
      Width           =   315
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      Caption         =   "Email"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   165
      TabIndex        =   27
      Top             =   3330
      Width           =   465
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      Caption         =   "Web"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   3225
      TabIndex        =   26
      Top             =   3360
      Width           =   405
   End
   Begin VB.Label Label15 
      AutoSize        =   -1  'True
      Caption         =   "NOTES"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   0
      TabIndex        =   25
      Top             =   3720
      Width           =   645
   End
   Begin VB.Image Image4 
      Height          =   1080
      Left            =   7110
      Stretch         =   -1  'True
      Top             =   2040
      Visible         =   0   'False
      Width           =   1170
   End
End
Attribute VB_Name = "address"
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

Dim i As Integer
Private Sub img_Click()
browse.Show
img.Visible = False
Image3.Visible = True
End Sub

Private Sub Command1_Click()
   ReDim Preserve sair(smax + 1)
        sair(smax).first = Text1.Text
        sair(smax).second = Text2.Text
        sair(smax).relation = Combo1.Text
        sair(smax).place = Text3.Text
        sair(smax).town = Text4.Text
        sair(smax).pin = Text5.Text
        sair(smax).district = Text6.Text
        sair(smax).state = Text7.Text
        sair(smax).country = Text8.Text
        sair(smax).telno = Text9.Text
        sair(smax).fax = Text10.Text
        sair(smax).email = Text11.Text
        sair(smax).web = Text12.Text
        sair(smax).notes = Text13.Text
        sair(smax).photo = ppath
        smax = smax + 1
        i = smax - 1
End Sub

Private Sub Command10_Click()
 sair(i).first = Text1.Text
        sair(i).second = Text2.Text
        sair(i).relation = Combo1.Text
        sair(i).place = Text3.Text
        sair(i).town = Text4.Text
        sair(i).pin = Text5.Text
        sair(i).district = Text6.Text
        sair(i).state = Text7.Text
        sair(i).country = Text8.Text
        sair(i).telno = Text9.Text
        sair(i).fax = Text10.Text
        sair(i).email = Text11.Text
        sair(i).web = Text12.Text
        sair(i).notes = Text13.Text
        sair(i).photo = ppath
        ppath = sair(i).photo
'        If (ppath <> "") Then
'            Image3.Picture = LoadPicture(ppath)
'        Else
'            Image3.Picture = image4.Picture
'        End If

End Sub

Private Sub Command11_Click()
Text1.Text = sair(i).first
        Text1.Text = ""
        Text2.Text = ""
        Combo1.Text = ""
        Text3.Text = ""
        Text4.Text = ""
        Text5.Text = ""
        Text6.Text = ""
        Text7.Text = ""
        Text8.Text = ""
        Text9.Text = ""
        Text10.Text = ""
        Text11.Text = ""
        Text12.Text = ""
        Text13.Text = ""
End Sub

Private Sub Command2_Click()
  If smax = 0 Then Exit Sub
  On Error Resume Next
      i = 0
      
        Text1.Text = sair(i).first
        Text2.Text = sair(i).second
        Combo1.Text = sair(i).relation
        Text3.Text = sair(i).place
        Text4.Text = sair(i).town
        Text5.Text = sair(i).pin
        Text6.Text = sair(i).district
        Text7.Text = sair(i).state
        Text8.Text = sair(i).country
        Text9.Text = sair(i).telno
        Text10.Text = sair(i).fax
        Text11.Text = sair(i).email
        Text12.Text = sair(i).web
        Text13.Text = sair(i).notes
        ppath = sair(i).photo
        If (ppath <> "") Then
            On Error GoTo anil
            Image3.Picture = LoadPicture(ppath)
        Else
            Image3.Picture = Image4.Picture
        End If
anil:
End Sub

Private Sub Command3_Click()

If smax = 0 Then Exit Sub
On Error Resume Next
 i = (i - 1 + smax) Mod smax
'        i = i - 1
'        If i < 0 Then
'            i = smax - 1
'        End If
        Text1.Text = sair(i).first
        Text2.Text = sair(i).second
        Combo1.Text = sair(i).relation
        Text3.Text = sair(i).place
        Text4.Text = sair(i).town
        Text5.Text = sair(i).pin
        Text6.Text = sair(i).district
        Text7.Text = sair(i).state
        Text8.Text = sair(i).country
        Text9.Text = sair(i).telno
        Text10.Text = sair(i).fax
        Text11.Text = sair(i).email
        Text12.Text = sair(i).web
        Text13.Text = sair(i).notes
        ppath = sair(i).photo
        If (ppath <> "") Then
            On Error GoTo anil
            Image3.Picture = LoadPicture(ppath)
        Else
            Image3.Picture = Image4.Picture
        End If
anil:
End Sub


Private Sub Command4_Click()
If smax = 0 Then Exit Sub
On Error Resume Next
       i = (i + 1) Mod smax
        Text1.Text = sair(i).first
        Text2.Text = sair(i).second
        Combo1.Text = sair(i).relation
        Text3.Text = sair(i).place
        Text4.Text = sair(i).town
        Text5.Text = sair(i).pin
        Text6.Text = sair(i).district
        Text7.Text = sair(i).state
        Text8.Text = sair(i).country
        Text9.Text = sair(i).telno
        Text10.Text = sair(i).fax
        Text11.Text = sair(i).email
        Text12.Text = sair(i).web
        Text13.Text = sair(i).notes
        ppath = sair(i).photo
        If (ppath <> "") Then
            On Error GoTo anil
            Image3.Picture = LoadPicture(ppath)
        Else
            Image3.Picture = Image4.Picture
        End If

anil:
               'i = (i + 1) Mod smax
End Sub


Private Sub Command5_Click()
 
If smax = 0 Then Exit Sub
On Error Resume Next
   i = smax - 1
        Text1.Text = sair(i).first
        Text2.Text = sair(i).second
        Combo1.Text = sair(i).relation
        Text3.Text = sair(i).place
        Text4.Text = sair(i).town
        Text5.Text = sair(i).pin
        Text6.Text = sair(i).district
        Text7.Text = sair(i).state
        Text8.Text = sair(i).country
        Text9.Text = sair(i).telno
        Text10.Text = sair(i).fax
        Text11.Text = sair(i).email
        Text12.Text = sair(i).web
        Text13.Text = sair(i).notes
        ppath = sair(i).photo
        If (ppath <> "") Then
            On Error GoTo anil
            Image3.Picture = LoadPicture(ppath)
        Else
            Image3.Picture = Image4.Picture
        End If
        
        
anil:

End Sub


Private Sub Command6_Click()
If smax < 0 Then
        MsgBox ("NOTHING  TO DELETE ")
        Exit Sub
    
    ElseIf smax = 1 Or smax < 1 Then
        i = 0
        Text1.Text = ""
        Text2.Text = ""
        Combo1.Text = ""
        Text3.Text = ""
        Text4.Text = ""
        Text5.Text = ""
        Text6.Text = ""
        Text7.Text = ""
        Text8.Text = ""
        Text9.Text = ""
        Text10.Text = ""
        Text11.Text = ""
        Text12.Text = ""
        Text13.Text = ""
        ppath = ""
       
'        Exit Sub
        smax = 0
        Exit Sub
    End If
    Dim z
    z = i
    While z <> smax - 1
        sair(z).first = sair((z + 1)).first  '*****
        sair(z).second = sair((z + 1)).second
        sair(z).relation = sair((z + 1)).relation
        sair(z).place = sair((z + 1)).place
        sair(z).town = sair((z + 1)).town
        sair(z).pin = sair((z + 1)).pin
        sair(z).district = sair((z + 1)).district
        sair(z).state = sair((z + 1)).state
        sair(z).country = sair((z + 1)).country
        sair(z).telno = sair((z + 1)).telno
        sair(z).fax = sair((z + 1)).fax
        sair(z).email = sair((z + 1)).email
        sair(z).web = sair((z + 1)).web
        sair(z).notes = sair((z + 1)).notes
        sair(z).photo = sair((z + 1)).photo
        
        z = z + 1
    Wend
        smax = smax - 1
    i = (i) Mod smax
    
    Text1.Text = sair(i).first
    Text2.Text = sair(i).second
    Combo1.Text = sair(i).relation
    Text3.Text = sair(i).place
    Text4.Text = sair(i).town
    Text5.Text = sair(i).pin
    Text6.Text = sair(i).district
    Text7.Text = sair(i).state
    Text8.Text = sair(i).country
    Text9.Text = sair(i).telno
    Text10.Text = sair(i).fax
    Text11.Text = sair(i).email
    Text12.Text = sair(i).web
    Text13.Text = sair(i).notes
    ppath = sair(i).photo
    If (ppath <> "") Then
    On Error GoTo anil
       Image3.Picture = LoadPicture(ppath)
    Else
       Image3.Picture = Image4.Picture
    End If


    

anil:
End Sub

Private Sub Command7_Click()
 
    find.Show
End Sub

Private Sub Command8_Click()
   sort.Show
End Sub

Private Sub Command9_Click()
  Dim j As Integer
    j = 0
    s = App.Path + "\address1.dat"
    Open s For Output As #2
       Do Until j = smax
           Write #2, sair(j).first
           Write #2, sair(j).second
           Write #2, sair(j).relation
           Write #2, sair(j).place
           Write #2, sair(j).town
           Write #2, sair(j).pin
           Write #2, sair(j).district
           Write #2, sair(j).state
           Write #2, sair(j).country
           Write #2, sair(j).telno
           Write #2, sair(j).fax
           Write #2, sair(j).email
           Write #2, sair(j).web
           Write #2, sair(j).notes
           Write #2, sair(j).photo
           j = j + 1
          
           Loop
        Close #2
        Unload Me
      End Sub


Private Sub Form_Load()
    i = 0
    smax = 0
    ppath = ""
'    GoTo anil
    On Error GoTo anil
    s = App.Path + "\address1.dat"
    Open s For Input As #1
       Do Until EOF(1)
           ReDim Preserve sair(smax + 1)
           Input #1, sair(smax).first
           Input #1, sair(smax).second
           Input #1, sair(smax).relation
           Input #1, sair(smax).place
           Input #1, sair(smax).town
           Input #1, sair(smax).pin
           Input #1, sair(smax).district
           Input #1, sair(smax).state
           Input #1, sair(smax).country
           Input #1, sair(smax).telno
           Input #1, sair(smax).fax
           Input #1, sair(smax).email
           Input #1, sair(smax).web
           Input #1, sair(smax).notes
           Input #1, sair(smax).photo
           smax = smax + 1
        Loop

        Close #1
        i = smax - 1
    address.Combo1.AddItem "Friend"
    address.Combo1.AddItem "Enemy"
    address.Combo1.AddItem "Brother"
    address.Combo1.AddItem "Sister"
    address.Combo1.AddItem "Others"
    address.Combo1.AddItem "Company"
    address.Combo1.AddItem "Organisation"
    If i > 0 Or i = 0 Then
        Text1.Text = sair(i).first
        Text2.Text = sair(i).second
        Combo1.Text = sair(i).relation
        Text3.Text = sair(i).place
        Text4.Text = sair(i).town
        Text5.Text = sair(i).pin
        Text6.Text = sair(i).district
        Text7.Text = sair(i).state
        Text8.Text = sair(i).country
        Text9.Text = sair(i).telno
        Text10.Text = sair(i).fax
        Text11.Text = sair(i).email
        Text12.Text = sair(i).web
        Text13.Text = sair(i).notes
        ppath = sair(i).photo
        If StrComp(ppath, "") <> 0 Then
            On Error GoTo anil
            Image3.Picture = LoadPicture(ppath)
        Else
            Image3.Picture = Image4.Picture
        End If
    End If
anil:
        
End Sub


Private Sub Image3_Click()
brow.Show
If (ppath <> "") Then
     On Error GoTo anil
    Image3.Picture = LoadPicture(ppath)
Else
  '  MsgBox ("anil")
    Image3.Picture = Image4.Picture
End If
anil:
End Sub


