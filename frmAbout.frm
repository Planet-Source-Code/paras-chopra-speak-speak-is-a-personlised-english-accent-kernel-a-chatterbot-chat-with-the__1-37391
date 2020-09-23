VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form2"
   ClientHeight    =   3825
   ClientLeft      =   3150
   ClientTop       =   2970
   ClientWidth     =   4785
   ForeColor       =   &H00000000&
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3825
   ScaleWidth      =   4785
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command2 
      Cancel          =   -1  'True
      Caption         =   "&Readme"
      Height          =   375
      Left            =   1800
      TabIndex        =   9
      Top             =   2400
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Ok"
      Height          =   375
      Left            =   3120
      TabIndex        =   2
      Top             =   2400
      Width           =   1335
   End
   Begin VB.Label Label7 
      Caption         =   "74, Vidya Vihar, Sector-9, Rohini, Delhi-110085, India. paraschopra@lycos.com"
      Height          =   495
      Left            =   240
      TabIndex        =   8
      ToolTipText     =   "Hury send me a postcard!"
      Top             =   3120
      Width           =   4335
   End
   Begin VB.Label Label6 
      Height          =   255
      Left            =   3750
      TabIndex        =   7
      Top             =   2040
      Width           =   855
   End
   Begin VB.Label Label5 
      Height          =   255
      Left            =   3750
      TabIndex        =   6
      Top             =   1680
      Width           =   855
   End
   Begin VB.Label Label4 
      Caption         =   "Page File Size:"
      Height          =   255
      Index           =   1
      Left            =   840
      TabIndex        =   5
      Top             =   1920
      Width           =   2895
   End
   Begin VB.Label Label4 
      Caption         =   "Physical Memory Available to Windows:"
      Height          =   255
      Index           =   0
      Left            =   840
      TabIndex        =   4
      Top             =   1680
      Width           =   2895
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      Index           =   1
      X1              =   840
      X2              =   4550
      Y1              =   1575
      Y2              =   1575
   End
   Begin VB.Line Line1 
      Index           =   0
      X1              =   840
      X2              =   4550
      Y1              =   1560
      Y2              =   1560
   End
   Begin VB.Image Image1 
      Height          =   735
      Left            =   120
      Top             =   120
      Width           =   615
   End
   Begin VB.Label Label3 
      Caption         =   "© Paras Chopra, CEO of NaramCheez, http://naramcheez.netfirms.com"
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
      Left            =   840
      TabIndex        =   3
      ToolTipText     =   "Double click to go to the website."
      Top             =   960
      Width           =   3735
   End
   Begin VB.Label Label2 
      Caption         =   "PostcardWare"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   840
      TabIndex        =   1
      ToolTipText     =   "If you want to use this program, you are requested to send me a postcard with your comments on it."
      Top             =   720
      Width           =   3735
   End
   Begin VB.Label Label1 
      Caption         =   "SPEAK beta"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   840
      TabIndex        =   0
      Top             =   120
      Width           =   3735
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'


'

'
'See the AboutForm function
'


Private Declare Sub GlobalMemoryStatus Lib "kernel32" (lpBuffer As MEMORYSTATUS)

Private Type MEMORYSTATUS
        dwLength As Long
        dwMemoryLoad As Long
        dwTotalPhys As Long
        dwAvailPhys As Long
        dwTotalPageFile As Long
        dwAvailPageFile As Long
        dwTotalVirtual As Long
        dwAvailVirtual As Long
 End Type

'
'  Opens a modal About box.
'
Public Sub AboutForm(Optional icon As Variant, Optional note As Variant)
  
   Dim ms As MEMORYSTATUS
   Dim i As Integer
   
   On Error Resume Next
   
   
   Label4(0).Top = Label4(0).Top + i
   Label4(1).Top = Label4(1).Top + i
   Label5.Top = Label5.Top + i
   Label6.Top = Label6.Top + i
   Command1.Top = Command1.Top + i
   Line1(0).Y1 = Line1(0).Y1 + i
   Line1(0).Y2 = Line1(0).Y2 + i
   Line1(1).Y1 = Line1(1).Y1 + i
   Line1(1).Y2 = Line1(1).Y2 + i
   Me.Height = Me.Height + i

      
   Me.Top = (Screen.Height - Me.Height) / 2
   Me.Left = (Screen.Width - Me.Width) / 2
   
   Me.Caption = "About " & Label1.Caption
   
   If Not IsMissing(icon) Then
      Me.icon = icon
      Image1 = icon
   End If
   
   
   Call GlobalMemoryStatus(ms)
   
   Label5.Caption = Format$((ms.dwTotalPhys / 1024), "###,###") & " KB"
   Label6.Caption = Format$(((ms.dwTotalPageFile - ms.dwAvailPageFile) / 1024), "###,###") & " KB"
   
   
   Me.Show vbModal
End Sub


Private Sub Command1_Click()
   Unload Me
End Sub


Private Sub Command2_Click()
Shell ("start " & App.Path & "\readme.txt")
End Sub

Private Sub Label3_DblClick()
Shell ("start http://naramcheez.netfirms.com")

End Sub
