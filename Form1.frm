VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "SPEAK - Speak is a Personlised English Accent Kernel"
   ClientHeight    =   4665
   ClientLeft      =   2895
   ClientTop       =   2160
   ClientWidth     =   5925
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4665
   ScaleWidth      =   5925
   Begin VB.CommandButton Command1 
      Caption         =   "Talk"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2280
      TabIndex        =   2
      Top             =   4200
      Width           =   1695
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2895
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   1
      Top             =   1080
      Width           =   5655
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   5655
   End
   Begin VB.Label Label2 
      Caption         =   "User:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   375
      Left            =   240
      TabIndex        =   3
      Top             =   4200
      Width           =   855
   End
   Begin VB.Menu about 
      Caption         =   "About"
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim InputRemaining As Boolean
Dim BufferInput As String

'Made By:
'Paras Chopra
'CEO , NaramCheez
'http://naramcheez.netfirms.com
'paraschopra@ lycos.com

Private Sub about_Click()
frmAbout.AboutForm Me.icon
End Sub

Private Sub Command1_Click()
    Dim FIRSTLINE As String
    Dim SECONDLINE As String
    Dim QUES() As String
    Dim ANS() As String
    Dim ANSWER(100) As String
    Dim TempStr As String
    Dim m As Integer
    m = 0
    If InputRemaining = True Then
        Open App.Path & "\mem.dat" For Append As #1
        Print #1, BufferInput
        TempStr = Rep(Text1, False)
        TempStr = Rep(RepeatAnsWer(TempStr), True)
        Print #1, TempStr
        Close #1
        InputRemaining = False
        texttemp = TempStr
        GoTo endsub
    End If
    Dim questionasked As String
    questionasked = LCase(Text1.Text)
    questionasked = Replace(questionasked, ",", "")
    questionasked = Replace(questionasked, ">", "")
    questionasked = Replace(questionasked, "?", "")
    questionasked = Replace(questionasked, "!", "")
    questionasked = Replace(questionasked, "£", "")
    questionasked = Replace(questionasked, "$", "")
    questionasked = Replace(questionasked, ")", "")
    questionasked = Replace(questionasked, "^", "")
    questionasked = Replace(questionasked, "&", "")
    questionasked = Replace(questionasked, "*", "")
    questionasked = Replace(questionasked, "(", "")
    questionasked = Replace(questionasked, "<", "")
    questionasked = Replace(questionasked, "_", "")
    questionasked = Replace(questionasked, "-", "")
    questionasked = Replace(questionasked, "=", "")
    questionasked = Replace(questionasked, ";", "")
    questionasked = Replace(questionasked, ":", "")
    questionasked = Replace(questionasked, "/", "")
    questionasked = Replace(questionasked, "\", "")
    questionasked = Replace(questionasked, ".", "")
    questionasked = Replace(questionasked, "@", "")
    questionasked = Replace(questionasked, "'", "")
    
    questionasked = Trim$(questionasked)
    questionasked = Trim$(StripSentence(" " & CStr(questionasked) & " "))
    
    
    QUES = GetTokens(questionasked, " ")
    totspaces = UBound(QUES) + 1
    percentWord = Int(100 / totspaces)
    Open App.Path & "\mem.dat" For Input As #1
    highest = 0
    Do While EOF(1) = False
        thisword = 0
        Line Input #1, FIRSTLINE
        Line Input #1, SECONDLINE
        If FIRSTLINE <> "" And SECONDLINE <> "" Then
            ANS = GetTokens(FIRSTLINE, " ")
            For i = 0 To UBound(ANS)
                For j = 0 To totspaces - 1
                    If QUES(j) = ANS(i) Then
                        thisword = thisword + percentWord
                    End If
                Next j
            Next i
            If thisword > highest Then
             'If thisword > highest Then 'del
                m = 0
                'm = m + 1 ' del
                ANSWER(m) = SECONDLINE
                
                highest = thisword
            ElseIf thisword = highest And highest <> 0 Then
                m = m + 1
                ANSWER(m) = SECONDLINE
            End If
        End If
    Loop
    Close #1
    
    If highest >= 95 Then
        texttemp = ANSWER(RndNum(0, CLng(m)))
    Else
        texttemp = Rep(Text1, False)
        texttemp = Rep(RepeatAnsWer(CStr(texttemp)), True)
        InputRemaining = True
        BufferInput = questionasked
    End If
endsub:
Human = Text1
Label1 = highest
computer = Text2 & "You: " & Human & vbCrLf
Text2.SelStart = Len(computer)
Me.Refresh
computer = computer & "SPEAK: " & texttemp & vbCrLf & vbCrLf
Text2 = computer
Text2.SelStart = Len(computer)
Text1 = ""
Text1.SetFocus
End Sub

Private Sub Form_Load()
    Text2 = ""
    Text1 = ""
End Sub

Public Function StripSentence(strng As String) As String
    
    strng = Replace(strng, " is ", " ")
    strng = Replace(strng, " as ", " ")
    strng = Replace(strng, " on ", " ")
    strng = Replace(strng, " and ", " ")
    strng = Replace(strng, " it ", " ")
    strng = Replace(strng, " at ", " ")
    strng = Replace(strng, " or ", " ")
    strng = Replace(strng, " if ", " ")
    strng = Replace(strng, " an ", " ")
    strng = Replace(strng, " a ", " ")
    strng = Replace(strng, " too ", " ")
    strng = Replace(strng, " but ", " ")
    strng = Replace(strng, " the ", " ")
    strng = Replace(strng, " this ", " ")
    strng = Replace(strng, " that ", " ")
    strng = Replace(strng, " are ", " ")
    strng = Replace(strng, " also ", " ")
    strng = Replace(strng, " where ", " ")
    strng = Replace(strng, " yet ", " ")
    strng = Replace(strng, " though ", " ")
    strng = Replace(strng, " whether ", " ")
    strng = Replace(strng, " either ", " ")
    strng = Replace(strng, " neither ", " ")
    strng = Replace(strng, " not only ", " ")
    strng = Replace(strng, " since ", " ")
    StripSentence = strng
    
End Function


Function RepeatAnsWer(questionasked As String) As String
    Dim temp() As String
    temp = GetTokens(questionasked, " ")
    questionaskedtemp = questionasked
    questionasked = ""
    For i = 0 To UBound(temp)
        find = temp(i)
        If find = "you" Then
        If (i > 0) Then
        If temp(i - 1) <> "are" Then
        If InStr(questionaskedtemp, "who") <> 0 Then
        replace1 = "me"
        GoTo nexti
        End If
        End If
        ElseIf (i < UBound(temp)) Then
        If temp(i + 1) <> "are" Then
        replace1 = "me"
        GoTo nexti
        End If
        End If: End If
        Select Case find
        Case "am": replace1 = "are"
        Case "was": replace1 = "were"
        Case "i": replace1 = "you"
        Case "i'd": replace1 = "you would"
        Case "i've": replace1 = "you have"
        Case "i'll": replace1 = "you will"
        Case "my": replace1 = "your"
        Case "are": replace1 = "am"
        Case "you've": replace1 = "I have"
        Case "you'll": replace1 = "I will"
        Case "your": replace1 = "my"
        Case "yours": replace1 = "mine"
        Case "you": replace1 = "i"
        Case "me": replace1 = "you"
        Case Else:
            replace1 = find
        End Select
nexti:
        questionasked = questionasked & replace1 & " "
    Next i
    
    questionasked = Trim(CStr(questionasked))
    RepeatAnsWer = questionasked
End Function

Public Function RndNum(Min As Long, Max As Long) As Long
    'generates a random integer between the supplied values of Min and Max
    Randomize
    RndNum = CLng(Round((Max - Min) * Rnd + Min))
End Function

Function Rep(questionasked As String, Reverse As Boolean)
    If Reverse = False Then
        questionasked = LCase(questionasked)
        questionasked = Replace(questionasked, ",", " ,")
        questionasked = Replace(questionasked, ">", " >")
        questionasked = Replace(questionasked, "?", " ?")
        questionasked = Replace(questionasked, "!", " !")
        questionasked = Replace(questionasked, "£", " £")
        questionasked = Replace(questionasked, "$", " $")
        questionasked = Replace(questionasked, ")", " )")
        questionasked = Replace(questionasked, "^", " ^")
        questionasked = Replace(questionasked, "&", " &")
        questionasked = Replace(questionasked, "*", " *")
        questionasked = Replace(questionasked, "(", " (")
        questionasked = Replace(questionasked, "<", " <")
        questionasked = Replace(questionasked, "_", " _")
        questionasked = Replace(questionasked, "-", " -")
        questionasked = Replace(questionasked, "=", " =")
        questionasked = Replace(questionasked, ";", " ;")
        questionasked = Replace(questionasked, ":", " :")
        questionasked = Replace(questionasked, "/", " /")
        questionasked = Replace(questionasked, "\", " \")
        questionasked = Replace(questionasked, ".", " .")
        questionasked = Replace(questionasked, "@", " @")
        questionasked = Replace(questionasked, "'", " '")
    Else
        questionasked = LCase(questionasked)
        questionasked = Replace(questionasked, " ,", ",")
        questionasked = Replace(questionasked, " >", ">")
        questionasked = Replace(questionasked, " ?", "?")
        questionasked = Replace(questionasked, " !", "!")
        questionasked = Replace(questionasked, " £", "£")
        questionasked = Replace(questionasked, " $", "$")
        questionasked = Replace(questionasked, " )", ")")
        questionasked = Replace(questionasked, " ^", "^")
        questionasked = Replace(questionasked, " &", "&")
        questionasked = Replace(questionasked, " *", "*")
        questionasked = Replace(questionasked, " (", "(")
        questionasked = Replace(questionasked, " <", "<")
        questionasked = Replace(questionasked, " _", "_")
        questionasked = Replace(questionasked, " -", "-")
        questionasked = Replace(questionasked, " =", "=")
        questionasked = Replace(questionasked, " ;", ";")
        questionasked = Replace(questionasked, " :", ":")
        questionasked = Replace(questionasked, " /", "/")
        questionasked = Replace(questionasked, " \", "\")
        questionasked = Replace(questionasked, " .", ".")
        questionasked = Replace(questionasked, " @", "@")
        questionasked = Replace(questionasked, " '", "'")
    End If
    Rep = questionasked
End Function


Private Sub Text1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
    KeyAscii = 0
        Command1_Click
    End If
End Sub

