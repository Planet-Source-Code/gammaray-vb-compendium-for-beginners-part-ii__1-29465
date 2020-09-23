VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   3  'Fester Dialog
   Caption         =   "VBC 2: String Functions"
   ClientHeight    =   7110
   ClientLeft      =   150
   ClientTop       =   150
   ClientWidth     =   7305
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7110
   ScaleWidth      =   7305
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'Bildschirmmitte
   Begin VB.PictureBox picText 
      BackColor       =   &H00C0E0FF&
      Height          =   2055
      Left            =   7200
      ScaleHeight     =   1995
      ScaleWidth      =   4395
      TabIndex        =   18
      Top             =   2880
      Visible         =   0   'False
      Width           =   4455
   End
   Begin VB.CommandButton cmdex2 
      Caption         =   "Example &2"
      Height          =   1335
      Left            =   6360
      TabIndex        =   17
      Top             =   2880
      Width           =   855
   End
   Begin VB.CommandButton cmdex1 
      Caption         =   "Example &1"
      Height          =   1335
      Left            =   5400
      TabIndex        =   16
      Top             =   2880
      Width           =   855
   End
   Begin VB.Frame Frame1 
      Caption         =   "Click on one of the Labels above to see here the explanation for its function:"
      Height          =   2055
      Left            =   0
      TabIndex        =   14
      Top             =   5040
      Width           =   7215
      Begin VB.Label lblHelp 
         BackColor       =   &H00C0FFC0&
         BorderStyle     =   1  'Fest Einfach
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1695
         Left            =   120
         TabIndex        =   15
         Top             =   240
         Width           =   6975
      End
   End
   Begin VB.Frame fraResults 
      Caption         =   "&Results:"
      Height          =   2895
      Left            =   0
      TabIndex        =   6
      Top             =   2040
      Width           =   5295
      Begin VB.Label lblEx2 
         Caption         =   "Type a text to the first TextBox. Use Hypertext-tags: <B>, </B>, <I> and </I>"
         Height          =   495
         Left            =   4200
         TabIndex        =   19
         Top             =   2640
         Visible         =   0   'False
         Width           =   4455
      End
      Begin VB.Label lblInStrRev 
         AutoSize        =   -1  'True
         Caption         =   "InStrRev(text,parameter)="
         Height          =   195
         Left            =   120
         TabIndex        =   13
         Top             =   720
         Width           =   1815
      End
      Begin VB.Label lblLen 
         AutoSize        =   -1  'True
         Caption         =   "Len(text)="
         Height          =   195
         Left            =   120
         TabIndex        =   12
         Top             =   2160
         Width           =   705
      End
      Begin VB.Label lblStrReverse 
         AutoSize        =   -1  'True
         Caption         =   "StrReverse(text)="
         Height          =   195
         Left            =   120
         TabIndex        =   11
         Top             =   1800
         Width           =   1230
      End
      Begin VB.Label lblLCase 
         AutoSize        =   -1  'True
         Caption         =   "LCase(text)="
         Height          =   195
         Left            =   120
         TabIndex        =   10
         Top             =   1440
         Width           =   885
      End
      Begin VB.Label lblUCase 
         AutoSize        =   -1  'True
         Caption         =   "UCase (text)="
         Height          =   195
         Left            =   120
         TabIndex        =   9
         Top             =   1080
         Width           =   960
      End
      Begin VB.Label lblMid 
         AutoSize        =   -1  'True
         Caption         =   "Mid (text,Len(text)/2,1)="
         Height          =   195
         Left            =   120
         TabIndex        =   8
         Top             =   2520
         Width           =   1695
      End
      Begin VB.Label lblInStr 
         AutoSize        =   -1  'True
         Caption         =   "InStr"
         Height          =   195
         Left            =   120
         TabIndex        =   7
         Top             =   360
         Width           =   330
      End
   End
   Begin VB.Frame fraParameter 
      Caption         =   "&Parameter:"
      Height          =   735
      Left            =   0
      TabIndex        =   4
      Top             =   1200
      Width           =   7215
      Begin VB.TextBox txtParameter 
         Height          =   285
         Left            =   120
         TabIndex        =   5
         Top             =   360
         Width           =   6975
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Quit"
      Height          =   615
      Left            =   5400
      TabIndex        =   3
      Top             =   4320
      Width           =   1815
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&DO IT!!!"
      Height          =   615
      Left            =   5400
      TabIndex        =   2
      Top             =   2160
      Width           =   1815
   End
   Begin VB.Frame fraText 
      Caption         =   "&Enter text to manipulate:"
      Height          =   1095
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7215
      Begin VB.TextBox txtText 
         Height          =   735
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   1  'Horizontal
         TabIndex        =   1
         Top             =   240
         Width           =   6975
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim ExampleNo As Integer

Dim InStrHelp As String
Dim InStrRevHelp As String
Dim LenHelp As String
Dim UCaseHelp As String
Dim LCaseHelp As String
Dim StrReverseHelp As String
Dim MidHelp As String

Private Sub cmdex1_Click()
ExampleNo = 1
lblEx2.Visible = False
picText.Visible = False
End Sub

Private Sub cmdex2_Click()
ExampleNo = 2
lblEx2.Left = 120
lblEx2.Top = 240
lblEx2.Visible = True
picText.Top = 2760
picText.Left = 120
picText.Visible = True
End Sub

Private Sub Command1_Click()
If txtText = "" Then
 MsgBox "the text must have at least one character", vbCritical
 Exit Sub
End If
If ExampleNo = 1 Then
 lblInStr = "InStr(1, text, parameter)  =  " & InStr(1, txtText, txtParameter)
 lblInStrRev = "InStrRev(text, parameter)  =  " & InStrRev(txtText, txtParameter)
 lblMid = "Mid (text, Len(text) / 2, 1)  =  " & Mid(txtText, Len(txtText) / 2, 1)
 lblUCase = "UCase(text)  =  " & UCase(txtText)
 lblLCase = "LCase(text)  =  " & LCase(txtText)
 lblLen = "Len(text)  =  " & Len(txtText)
 lblStrReverse = "StrReverse(text) = " & StrReverse(txtText)
Else
 'Format the text:
 FormatHyperText
End If
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Form_Load()
ExampleNo = 1
InStrHelp = "SYNTAX: pos = InStr([start], text1, text2)" & Chr(13) & "start = the starting position (optional)" & Chr(13) & "text1 = the string which will be checked" & Chr(13) & "text2 = the string we want to find in text1" & Chr(13) & "RESULT: the function returns the position of the first occurance of text2 within text1 or NULL if text2 doesn't appear in text1"
InStrRevHelp = "SYNTAX: pos = InStrRev(text1, text2, [start])" & Chr(13) & "start = the starting position (optional)" & Chr(13) & "text1 = the string which will be checked" & Chr(13) & "text2 = the string we want to find in text1" & Chr(13) & "RESULT: the function returns the position of the last occurance of text2 within text1 or NULL if text2 doesn't appear in text1"
LenHelp = "SYNTAX: length = Len(text)" & Chr(13) & "This function returns the length of the string: 'text'"
StrReverseHelp = "SYNTAX: text2 = StrReverse(text1)" & Chr(13) & "Returns the reversed string of 'text1'"
MidHelp = "SYNTAX: text2 = Mid(text1, start, length)" & Chr(13) & "text1 = the string from which you want to take a part" & Chr(13) & "start = the staring position" & Chr(13) & "length = the length of the string you wanna have" & Chr(13) & "RESULT: Mid returns a string from 'text1' which starts at the position 'start' and has a length of 'length'"
UCaseHelp = "SYNTAX: text2 = UCase(text)" & Chr(13) & "Returns the string 'text' in upper cases"
LCaseHelp = "SYNTAX: text2 = LCase(text)" & Chr(13) & "Returns the string 'text' in lower cases"
lblLen = "Len(text)="
lblInStr = "InStr(1, text, parameter)="
lblMid = "Mid (text, Len(text) / 2, 1)="
lblUCase = "UCase(text)="
lblLCase = "LCase(text)="
lblStrReverse = "StrReverse(text)="
End Sub


Private Sub lblInStr_Click()
lblHelp = InStrHelp
End Sub

Private Sub lblInStrRev_Click()
lblHelp = InStrRevHelp
End Sub

Private Sub lblLCase_Click()
lblHelp = LCaseHelp
End Sub

Private Sub lblLen_Click()
lblHelp = LenHelp
End Sub

Private Sub lblMid_Click()
lblHelp = MidHelp
End Sub

Private Sub lblStrReverse_Click()
lblHelp = StrReverseHelp
End Sub

Private Sub lblUCase_Click()
lblHelp = UCaseHelp
End Sub

Private Sub FormatHyperText()
Dim GetWord As String, CheckString As String, WordToWrite As String

Dim BFlag As Boolean, IFlag As Boolean
'this will hold our currently selected Bold & Italic mode

Dim exitloop As Boolean
'if true then we'll end

CheckString = txtText
'CheckString is the whole text we have to check

picText.Font.Bold = False
picText.Font.Italic = False
picText.Cls

Do
  'check if after there is a space char somewhere in the
  'string
  If InStr(CheckString, " ") = 0 Then 'if not then this is the
   GetWord = CheckString              'the last word
   exitloop = True
  Else
   'If yes, get all chars to the the next space-character
   '(so we'll get a full word):
   GetWord = Left(CheckString, InStr(CheckString, " ") - 1)
   'The Left function works like Mid
   'SYNTAX: Left(Text, L) it returns a string from Text
   'from the first position(that's why it's called Left) with
   'the length L
   'Or: It returns the first L chars
   'Example: Left("1234567890",4) = "1234"
   'Same for Right:
   'Right(Text, L)
   'Retruns the last L chars
   'Right("1234567890",4)= "7890"
   
   CheckString = Right(CheckString, Len(CheckString) - Len(GetWord) - 1)
   'Now we reduce CkeckString because we don't need the
   'current word anymore
  End If
  
  WordToWrite = GetWord
  'Word, which will be displayed
  
  'Test the first 3 characters of out Word:
  Select Case UCase(Left(GetWord, 3))
   Case "<B>" 'If <b> then set Bold=True
    BFlag = True
    picText.Font.Bold = BFlag
    WordToWrite = Right(GetWord, Len(GetWord) - 3)
    'Now we have to eliminate the tag from the word
    GetWord = WordToWrite
   Case "<I>" 'If <i> then set Italic=True
    IFlag = True
    picText.Font.Italic = IFlag
    WordToWrite = Right(GetWord, Len(GetWord) - 3)
    'Now we have to eliminate the tag from the word
    GetWord = WordToWrite
  End Select
  'Same as for the <B> & <I> tags
  'get the last 4 characters of the word
  Select Case UCase(Right(GetWord, 4))
   Case "</B>"
    BFlag = False
    WordToWrite = Left(GetWord, Len(GetWord) - 4)
   Case "</I>"
    IFlag = False
    WordToWrite = Left(GetWord, Len(GetWord) - 4)
  End Select
  
  'If the word is too large -> new line
  If TextWidth(WordToWrite) > picText.Width - picText.CurrentX - 100 Then
   picText.Print
   picText.Print " ";
  End If
  
  'Write the word
  picText.Print WordToWrite & " ";
  picText.Font.Bold = BFlag
  picText.Font.Italic = IFlag
Loop Until exitloop = True

End Sub
