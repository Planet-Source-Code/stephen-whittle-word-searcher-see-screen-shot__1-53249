VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Word searcher"
   ClientHeight    =   5550
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5460
   LinkTopic       =   "Form1"
   ScaleHeight     =   5550
   ScaleWidth      =   5460
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame3 
      Caption         =   "Crossword finder"
      ForeColor       =   &H00FF0000&
      Height          =   975
      Left            =   120
      TabIndex        =   15
      Top             =   2520
      Width           =   5175
      Begin VB.CommandButton Command7 
         Caption         =   "New word"
         Height          =   255
         Left            =   3840
         TabIndex        =   35
         Top             =   600
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.CommandButton Command6 
         Caption         =   "Search"
         Height          =   285
         Left            =   3840
         TabIndex        =   34
         Top             =   240
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.TextBox Text6 
         Height          =   285
         Index           =   14
         Left            =   3480
         MaxLength       =   1
         TabIndex        =   33
         Top             =   360
         Width           =   210
      End
      Begin VB.TextBox Text6 
         Height          =   285
         Index           =   13
         Left            =   3240
         MaxLength       =   1
         TabIndex        =   32
         Top             =   360
         Width           =   210
      End
      Begin VB.TextBox Text6 
         Height          =   285
         Index           =   12
         Left            =   3000
         MaxLength       =   1
         TabIndex        =   31
         Top             =   360
         Width           =   210
      End
      Begin VB.TextBox Text6 
         Height          =   285
         Index           =   11
         Left            =   2760
         MaxLength       =   1
         TabIndex        =   30
         Top             =   360
         Width           =   210
      End
      Begin VB.TextBox Text6 
         Height          =   285
         Index           =   10
         Left            =   2520
         MaxLength       =   1
         TabIndex        =   29
         Top             =   360
         Width           =   210
      End
      Begin VB.TextBox Text6 
         Height          =   285
         Index           =   9
         Left            =   2280
         MaxLength       =   1
         TabIndex        =   28
         Top             =   360
         Width           =   210
      End
      Begin VB.TextBox Text6 
         Height          =   285
         Index           =   8
         Left            =   2040
         MaxLength       =   1
         TabIndex        =   27
         Top             =   360
         Width           =   210
      End
      Begin VB.TextBox Text6 
         Height          =   285
         Index           =   7
         Left            =   1800
         MaxLength       =   1
         TabIndex        =   26
         Top             =   360
         Width           =   210
      End
      Begin VB.TextBox Text6 
         Height          =   285
         Index           =   6
         Left            =   1560
         MaxLength       =   1
         TabIndex        =   25
         Top             =   360
         Width           =   210
      End
      Begin VB.TextBox Text6 
         Height          =   285
         Index           =   5
         Left            =   1320
         MaxLength       =   1
         TabIndex        =   24
         Top             =   360
         Width           =   210
      End
      Begin VB.TextBox Text6 
         Height          =   285
         Index           =   4
         Left            =   1080
         MaxLength       =   1
         TabIndex        =   23
         Top             =   360
         Width           =   210
      End
      Begin VB.TextBox Text6 
         Height          =   285
         Index           =   3
         Left            =   840
         MaxLength       =   1
         TabIndex        =   22
         Top             =   360
         Width           =   210
      End
      Begin VB.TextBox Text6 
         Height          =   285
         Index           =   2
         Left            =   600
         MaxLength       =   1
         TabIndex        =   21
         Top             =   360
         Width           =   210
      End
      Begin VB.TextBox Text6 
         Height          =   285
         Index           =   1
         Left            =   360
         MaxLength       =   1
         TabIndex        =   20
         Top             =   360
         Width           =   210
      End
      Begin VB.TextBox Text6 
         Height          =   285
         Index           =   0
         Left            =   120
         MaxLength       =   1
         TabIndex        =   19
         Top             =   360
         Width           =   210
      End
      Begin VB.TextBox Text5 
         Height          =   285
         Left            =   2280
         TabIndex        =   18
         Top             =   240
         Width           =   735
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Setup line"
         Height          =   285
         Left            =   3720
         TabIndex        =   16
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label5 
         Caption         =   "How many letters"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   17
         Top             =   270
         Width           =   1935
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Unscramble"
      ForeColor       =   &H00FF0000&
      Height          =   735
      Left            =   120
      TabIndex        =   11
      Top             =   1680
      Width           =   5175
      Begin VB.TextBox Text4 
         Height          =   285
         Left            =   1920
         TabIndex        =   13
         Top             =   240
         Width           =   1695
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Unscramble"
         Height          =   285
         Left            =   3720
         TabIndex        =   12
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label4 
         Caption         =   "Make words from"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   14
         Top             =   270
         Width           =   1575
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Search parts"
      ForeColor       =   &H00FF0000&
      Height          =   1455
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   5175
      Begin VB.CommandButton Command3 
         Caption         =   "Search"
         Height          =   285
         Left            =   3840
         TabIndex        =   8
         Top             =   960
         Width           =   1095
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Search"
         Height          =   285
         Left            =   3840
         TabIndex        =   7
         Top             =   600
         Width           =   1095
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Left            =   2880
         TabIndex        =   6
         Top             =   960
         Width           =   735
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   2880
         TabIndex        =   5
         Top             =   600
         Width           =   735
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Search"
         Height          =   285
         Left            =   3840
         TabIndex        =   3
         Top             =   240
         Width           =   1095
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   2880
         TabIndex        =   0
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label3 
         Caption         =   "Search words with part "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   10
         Top             =   1080
         Width           =   2415
      End
      Begin VB.Label Label2 
         Caption         =   "Search words ending with"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   9
         Top             =   720
         Width           =   2535
      End
      Begin VB.Label Label1 
         Caption         =   "Search words beginning with"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   4
         Top             =   360
         Width           =   2535
      End
   End
   Begin VB.ListBox List1 
      Height          =   1815
      Left            =   120
      TabIndex        =   1
      Top             =   3600
      Width           =   5175
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim words() As String
Dim crossl As Integer

Private Sub Command1_Click()
Dim i As Long, l As Integer, str1 As String, lb As Long, ub As Long

List1.Clear

If Text1.Text <> "" Then
  str1 = Text1.Text
  l = Len(Text1.Text)

   For i = LBound(words) To UBound(words)
    If Left$(words(i), l) = str1 Then
     List1.AddItem words(i)
    End If
   Next i
End If

If List1.ListCount = 0 Then
 List1.AddItem "<< No Results >>"
End If
   

End Sub

Private Sub Command2_Click()
Dim i As Long, l As Integer, str1 As String, lb As Long, ub As Long

List1.Clear

If Text2.Text <> "" Then
 str1 = Text2.Text
 l = Len(Text2.Text)
  For i = LBound(words) To UBound(words)
   If Right$(words(i), l) = str1 Then
    List1.AddItem words(i)
   End If
  Next i
End If

If List1.ListCount = 0 Then
 List1.AddItem "<< No Results >>"
End If


End Sub


Private Sub Command3_Click()
Dim i As Long, l As Integer, str1 As String, lb As Long, ub As Long

List1.Clear

If Text3.Text <> "" Then
 str1 = Text3.Text

  For i = LBound(words) To UBound(words)
   If InStr(1, words(i), str1) Then
    List1.AddItem words(i)
   End If
  Next i
End If

If List1.ListCount = 0 Then
 List1.AddItem "<< No Results >>"
End If


End Sub


Private Sub Command4_Click()
Dim i As Long, l As Long, str1 As String, x As Long, found As Boolean
Dim ins As Long

List1.Clear

l = Len(Text4.Text)

For i = LBound(words) To UBound(words)
 If Len(words(i)) = l Then
     str1 = Text4.Text
     found = True
   
       For x = 1 To l
         ins = InStr(1, str1, Mid$(words(i), x, 1))
        If ins = 0 Then
         found = False
         Exit For
        Else
         str1 = Left$(str1, ins - 1) & Mid$(str1, ins + 1)
        End If
       Next x
    
    If found = True Then
     List1.AddItem words(i)
    End If
    
  End If
Next i

If List1.ListCount = 0 Then
 List1.AddItem "<< No Results >>"
End If

End Sub

Private Sub Command5_Click()
Dim i As Integer, x As Integer


If IsNumeric(Text5.Text) Then
  x = Val(Text5.Text)
 If x < 16 Then
  Command5.Visible = False
  Text5.Visible = False
  Label5.Visible = False
  Command6.Visible = True
  Command7.Visible = True
  crossl = x
For i = 0 To x - 1
 Text6(i).Visible = True
Next i
 Else
  MsgBox "15 letter word is the maximum", vbInformation, "Maximum"
 End If
Else
 MsgBox "Must be numeric", vbInformation, "How many letters"
End If

End Sub

Private Sub Command6_Click()
Dim i As Long, x As Integer, found As Boolean

List1.Clear
List1.AddItem "<< Possible answers >>"

For i = LBound(words) To UBound(words)
 If Len(words(i)) = crossl Then
    found = True
     For x = 0 To crossl - 1
      If Text6(x).Text <> "" Then
        If Mid$(words(i), x + 1, 1) = Text6(x).Text Then
          found = True
        Else
          found = False
          Exit For
        End If
      End If
     Next x
  If found = True Then
    List1.AddItem words(i)
  End If
 End If
Next i

If List1.ListCount = 1 Then
 List1.Clear
 List1.AddItem "<< No Results >>"
End If


End Sub


Private Sub Command7_Click()
Dim i As Integer

Command6.Visible = False
Command5.Visible = True
Text5.Visible = True
Text5.Text = ""
Label5.Visible = True

For i = Text6.LBound To Text6.UBound
 Text6(i).Text = ""
 Text6(i).Visible = False
Next i

List1.Clear

Command7.Visible = False
End Sub

Private Sub Form_Load()
Dim i As Long

Open App.Path & "\wordsdic.DIC" For Input As #1
 Do While Not EOF(1)
  ReDim Preserve words(i)
  Line Input #1, words(i)
  i = i + 1
 Loop
Close #1

For i = Text6.LBound To Text6.UBound
 Text6(i).Visible = False
Next i


End Sub


