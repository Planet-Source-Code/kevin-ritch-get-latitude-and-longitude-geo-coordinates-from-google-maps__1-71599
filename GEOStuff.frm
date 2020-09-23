VERSION 5.00
Begin VB.Form GEOStuffForm 
   BackColor       =   &H007E6436&
   Caption         =   "GEO STUFF (Fun with Map Coordinates in the USA) © Kevin Ritch - Thanks to Google Maps"
   ClientHeight    =   4035
   ClientLeft      =   120
   ClientTop       =   420
   ClientWidth     =   9750
   BeginProperty Font 
      Name            =   "Courier New"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "GEOStuff.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4035
   ScaleWidth      =   9750
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture2 
      BackColor       =   &H007E6436&
      Height          =   2535
      Left            =   4920
      ScaleHeight     =   2475
      ScaleWidth      =   4635
      TabIndex        =   10
      Top             =   240
      Width           =   4695
      Begin VB.TextBox Latitude 
         Alignment       =   2  'Center
         BackColor       =   &H00FFB404&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   2
         Left            =   1200
         TabIndex        =   25
         Top             =   1920
         Width           =   1575
      End
      Begin VB.TextBox Longitude 
         Alignment       =   2  'Center
         BackColor       =   &H00FFB404&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   2
         Left            =   2880
         TabIndex        =   24
         Top             =   1920
         Width           =   1575
      End
      Begin VB.TextBox Zip 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   2
         Left            =   3600
         TabIndex        =   14
         Text            =   "11788"
         Top             =   1320
         Width           =   855
      End
      Begin VB.TextBox State 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   2
         Left            =   2880
         TabIndex        =   13
         Text            =   "NY"
         Top             =   1320
         Width           =   615
      End
      Begin VB.TextBox City 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   2
         Left            =   120
         TabIndex        =   12
         Text            =   "Hauppauge"
         Top             =   1320
         Width           =   2655
      End
      Begin VB.TextBox Street 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   2
         Left            =   120
         TabIndex        =   11
         Text            =   "235 Brooksite Dr"
         Top             =   480
         Width           =   4335
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Lat/Long"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Index           =   10
         Left            =   120
         TabIndex        =   23
         Top             =   1920
         Width           =   900
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Zip"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Index           =   7
         Left            =   3600
         TabIndex        =   18
         Top             =   960
         Width           =   270
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "State"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Index           =   6
         Left            =   2880
         TabIndex        =   17
         Top             =   960
         Width           =   615
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "City"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Index           =   5
         Left            =   120
         TabIndex        =   16
         Top             =   960
         Width           =   615
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Address 2"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Index           =   4
         Left            =   120
         TabIndex        =   15
         Top             =   120
         Width           =   1185
      End
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H007E6436&
      Height          =   2535
      Left            =   120
      ScaleHeight     =   2475
      ScaleWidth      =   4635
      TabIndex        =   1
      Top             =   240
      Width           =   4695
      Begin VB.TextBox Latitude 
         Alignment       =   2  'Center
         BackColor       =   &H00FFB404&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   1
         Left            =   1200
         TabIndex        =   22
         Top             =   1920
         Width           =   1575
      End
      Begin VB.TextBox Longitude 
         Alignment       =   2  'Center
         BackColor       =   &H00FFB404&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   1
         Left            =   2880
         TabIndex        =   21
         Top             =   1920
         Width           =   1575
      End
      Begin VB.TextBox Street 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   1
         Left            =   120
         TabIndex        =   5
         Text            =   "166 Riviera Parkway"
         Top             =   480
         Width           =   4335
      End
      Begin VB.TextBox City 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   1
         Left            =   120
         TabIndex        =   4
         Text            =   "Lindenhurst"
         Top             =   1320
         Width           =   2655
      End
      Begin VB.TextBox State 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   1
         Left            =   2880
         TabIndex        =   3
         Text            =   "NY"
         Top             =   1320
         Width           =   615
      End
      Begin VB.TextBox Zip 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   1
         Left            =   3600
         TabIndex        =   2
         Text            =   "11757"
         Top             =   1320
         Width           =   855
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Lat/Long"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Index           =   9
         Left            =   120
         TabIndex        =   20
         Top             =   1920
         Width           =   900
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Address 1"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Index           =   0
         Left            =   120
         TabIndex        =   9
         Top             =   120
         Width           =   1185
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "City"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Index           =   1
         Left            =   120
         TabIndex        =   8
         Top             =   960
         Width           =   615
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "State"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Index           =   2
         Left            =   2880
         TabIndex        =   7
         Top             =   960
         Width           =   615
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Zip"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Index           =   3
         Left            =   3600
         TabIndex        =   6
         Top             =   960
         Width           =   270
      End
   End
   Begin VB.CommandButton AskGoogleButton 
      Caption         =   "ASK GOOGLE"
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
      TabIndex        =   0
      Top             =   3360
      Width           =   1575
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   $"GEOStuff.frx":0442
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C0C0&
      Height          =   480
      Index           =   11
      Left            =   2760
      TabIndex        =   26
      Top             =   3480
      Width           =   6930
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   " This button will request the Latitudes and Longitudes for these 2 addresses from Google Maps."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   8
      Left            =   120
      TabIndex        =   19
      Top             =   3000
      Width           =   9300
   End
End
Attribute VB_Name = "GEOStuffForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub AskGoogleButton_Click()
 On Error GoTo BolloxedData:
'===========
'CLEAR STUFF
'===========
 Latitude(1) = "": Latitude(2) = ""
 Longitude(1) = "": Longitude(2) = ""
 DoEvents
'==========================================
'CHECK IF USER FILLED IN THE ADDRESS FIELDS
'==========================================
 If MissingData Then
  MsgBox "Please fill in the form correctly.", vbApplicationModal + vbExclamation, "Whoops!"
  Exit Sub
 End If
 AskGoogleButton.Enabled = False
 Screen.MousePointer = 11
 Dim NotFound As Boolean
 For i = 1 To 2
  NotFound = True
  WebURL$ = "http://maps.google.com/maps?q="
  WebURL$ = WebURL$ & Trim$(Street(i)) & ",+"
  WebURL$ = WebURL$ & Trim$(City(i)) & ",+"
  WebURL$ = WebURL$ & Trim$(State(i)) & ",+"
  WebURL$ = WebURL$ & Trim$(Zip(i))
  WebURL$ = Replace$(WebURL$, " ", "+")
  GoogleWebPage$ = GetUrlSource(WebURL$)
  a$ = LCase$(GoogleWebPage$)
 '=======================================================
 'EXTRACT THE LATITUDE AND LONGITUDE FROM GOOGLE WEB PAGE
 '=======================================================
  GC = InStr(a$, ":{lat:") + 2
  If GC > 2 Then
   GC2 = InStr(GC, a$, "},")
   LC = GC2 - GC
   If LC > 20 And LC < 200 Then
    a$ = Mid$(a$, GC + 4, LC - 6)
    Comma = InStr(a$, ",")
    If Comma Then
     Lat$ = Left$(a$, Comma - 1) & String$(12, 48)
     Lat$ = Left$(Lat$, 12)
     Lon$ = Right$(a$, Len(a$) - (Comma + 4)) & String$(12, 48)
     Lon$ = Left$(Lon$, 12)
     If Abs(Val(Lat$)) > 0 And Abs(Val(Lon$)) > 0 Then
      Latitude(i) = Lat$
      Longitude(i) = Lon$
      NotFound = False
     End If
    End If
   End If
  End If
  DoEvents
 Next i
ReturnResult:
 Screen.MousePointer = Default
 If NotFound Then
  MsgBox "NO MAP COORDINATES HAVE BEEN FOUND!", vbApplicationModal + vbExclamation, "ADDRESS???"
  Beep
 Else ' Watch for my next upgrade - ok? (Dinner's waiting)
 '============================================================================
 'Why not put some MORE FUN into this and calculate the distance between them?
 '============================================================================
 End If
 AskGoogleButton.Enabled = True
 Exit Sub
BolloxedData:
 NotFound = True
 Resume ReturnResult:
End Sub
Function MissingData() As Boolean
 For i = 1 To 2
  If Trim(State(i)) = "" Or Trim(Zip(i)) = "" Then
   MissingData = True
   Exit For
  End If
 Next i
End Function
