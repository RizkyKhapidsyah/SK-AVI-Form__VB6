VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "AVI Sample Code"
   ClientHeight    =   6165
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6255
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   411
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   417
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command12 
      Caption         =   "Close"
      Height          =   375
      Left            =   5160
      TabIndex        =   25
      Top             =   2040
      Width           =   975
   End
   Begin VB.CommandButton Command11 
      Caption         =   "Click for Windowed AVI"
      Height          =   375
      Left            =   3120
      TabIndex        =   24
      Top             =   2040
      Width           =   2055
   End
   Begin VB.Frame Frame2 
      Caption         =   "Normal AVI"
      Height          =   1215
      Left            =   3120
      TabIndex        =   15
      Top             =   2520
      Width           =   3015
      Begin VB.HScrollBar HScroll1 
         Height          =   135
         Left            =   120
         TabIndex        =   22
         Top             =   960
         Width           =   2775
      End
      Begin VB.Label Label9 
         Height          =   615
         Left            =   1560
         TabIndex        =   20
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   720
         TabIndex        =   19
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label7 
         Caption         =   "Frames"
         Height          =   255
         Left            =   120
         TabIndex        =   18
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   720
         TabIndex        =   17
         Top             =   600
         Width           =   735
      End
      Begin VB.Label Label5 
         Caption         =   "MS"
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   600
         Width           =   375
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Stretched AVI"
      Height          =   1215
      Left            =   3120
      TabIndex        =   10
      Top             =   3720
      Width           =   3015
      Begin VB.HScrollBar HScroll2 
         Height          =   135
         Left            =   120
         TabIndex        =   23
         Top             =   960
         Width           =   2775
      End
      Begin VB.Label Label10 
         Height          =   615
         Left            =   1560
         TabIndex        =   21
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label Label4 
         Caption         =   "MS"
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   600
         Width           =   375
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   720
         TabIndex        =   13
         Top             =   600
         Width           =   735
      End
      Begin VB.Label Label2 
         Caption         =   "Frames"
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   720
         TabIndex        =   11
         Top             =   240
         Width           =   735
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   5160
      Top             =   1440
   End
   Begin VB.CommandButton Command10 
      Caption         =   "Close"
      Height          =   495
      Left            =   5520
      TabIndex        =   9
      Top             =   5640
      Width           =   615
   End
   Begin VB.CommandButton Command9 
      Caption         =   "Close"
      Height          =   495
      Left            =   2400
      TabIndex        =   8
      Top             =   5640
      Width           =   615
   End
   Begin VB.CommandButton Command8 
      Caption         =   "Pause"
      Height          =   495
      Left            =   3960
      TabIndex        =   7
      Top             =   5640
      Width           =   615
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Resume"
      Height          =   495
      Left            =   4680
      TabIndex        =   6
      Top             =   5640
      Width           =   735
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Stop"
      Height          =   495
      Left            =   3240
      TabIndex        =   5
      Top             =   5640
      Width           =   615
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Play AVI (Original Dimensions)"
      Height          =   495
      Left            =   3240
      TabIndex        =   4
      Top             =   5040
      Width           =   2895
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Stop"
      Height          =   495
      Left            =   120
      TabIndex        =   3
      Top             =   5640
      Width           =   615
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Resume"
      Height          =   495
      Left            =   1560
      TabIndex        =   2
      Top             =   5640
      Width           =   735
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Pause"
      Height          =   495
      Left            =   840
      TabIndex        =   1
      Top             =   5640
      Width           =   615
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Play AVI Stretched"
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   5040
      Width           =   2895
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function mciSendString Lib "winmm.dll" Alias _
    "mciSendStringA" (ByVal lpstrCommand As String, ByVal _
    lpstrReturnString As Any, ByVal uReturnLength As Long, ByVal _
    hwndCallback As Long) As Long
    
    
Dim APATH As String
Private Sub Command1_Click()
Dim mssg As String * 255
Dim Tmp

Last$ = Form1.hWnd & " Style " & &H40000000
ToDo$ = "open " & APATH & "test.avi Type avivideo Alias video1 parent " & Last$
i = mciSendString(ToDo$, 0&, 0, 0)
i = mciSendString("put video1 window at 16 10 124 310", 0&, 0, 0)
i = mciSendString("play video1 from 0", 0&, 0, 0)


  i = mciSendString("set video1 time format frames", 0&, 0, 0)
  i = mciSendString("status video1 length", mssg, 255, 0)

HScroll2.Max = mssg
End Sub

Private Sub Command10_Click()
i = mciSendString("close video2", 0&, 0, 0)
End Sub

Private Sub Command11_Click()

i = mciSendString("open " & APATH & "test.avi type avivideo alias video3", 0&, 0, 0)
i = mciSendString("play video3 from 0", 0&, 0, 0)

End Sub

Private Sub Command12_Click()
i = mciSendString("close video3", 0&, 0, 0)
End Sub


Private Sub Command2_Click()
i = mciSendString("pause video1", 0&, 0, 0)
End Sub


Private Sub Command3_Click()
i = mciSendString("resume video1", 0&, 0, 0)
End Sub


Private Sub Command4_Click()
i = mciSendString("stop video1", 0&, 0, 0)
i = mciSendString("seek video1 to start", 0&, 0, 0)

End Sub

Private Sub Command5_Click()
Dim mssg As String * 255
Dim sReturn As String * 128
Dim lPos As Long
Dim lStart As Long
    
Last$ = Form1.hWnd & " Style " & &H40000000
ToDo$ = "open " & APATH & "test.avi Type avivideo Alias video2 parent " & Last$
i = mciSendString(ToDo$, 0&, 0, 0)

i = mciSendString("Where video2 destination", ByVal sReturn, Len(sReturn) - 1, 0)
    
lStart = InStr(1, sReturn, " ") 'pos of top
lPos = InStr(lStart + 1, sReturn, " ") 'pos of left
lStart = InStr(lPos + 1, sReturn, " ") 'pos width
lWidth = Mid(sReturn, lPos, lStart - lPos)
lHeight = Mid(sReturn, lStart + 1)
    
    
i = mciSendString("put video2 window at 176 10 " & lWidth & " " & lHeight, 0&, 0, 0)
i = mciSendString("play video2 from 0", 0&, 0, 0)

  i = mciSendString("set video2 time format frames", 0&, 0, 0)
  i = mciSendString("status video2 length", mssg, 255, 0)

HScroll1.Max = mssg
End Sub


Private Sub Command6_Click()
i = mciSendString("stop video2", 0&, 0, 0)
i = mciSendString("seek video2 to start", 0&, 0, 0)

End Sub

Private Sub Command7_Click()
i = mciSendString("resume video2", 0&, 0, 0)
End Sub

Private Sub Command8_Click()
i = mciSendString("pause video2", 0&, 0, 0)
End Sub

Private Sub Command9_Click()
i = mciSendString("close video1", 0&, 0, 0)
End Sub

Private Sub Form_Load()
  
  
If Right$(App.Path, 1) = "\" Then
  APATH = App.Path
Else
  APATH = App.Path & "\"
End If



End Sub

Private Sub Form_Unload(Cancel As Integer)

Screen.MousePointer = 11 'pointer may not activate on FAST PCs
  i = mciSendString("close video1", 0&, 0, 0)
  i = mciSendString("close video2", 0&, 0, 0)
  i = mciSendString("close video3", 0&, 0, 0)
Screen.MousePointer = 0

End Sub


Private Sub Timer1_Timer()
Dim Msg As String
Dim mssg As String * 255
On Error Resume Next


   i = mciSendString("set video1 time format frames", 0&, 0, 0)
   i = mciSendString("status video1 position", mssg, 255, 0)
Label1.Caption = mssg
HScroll2.Value = Label1.Caption
   i = mciSendString("set video1 time format ms", 0&, 0, 0)
   i = mciSendString("status video1 position", mssg, 255, 0)
Label3.Caption = mssg


   i = mciSendString("set video2 time format frames", 0&, 0, 0)
   i = mciSendString("status video2 position", mssg, 255, 0)
Label8.Caption = mssg
HScroll1.Value = Label8.Caption
   i = mciSendString("set video2 time format ms", 0&, 0, 0)
   i = mciSendString("status video2 position", mssg, 255, 0)
Label6.Caption = mssg



i = mciSendString("status video1 mode", mssg, 255, 0)
Label10.Caption = mssg
i = mciSendString("status video2 mode", mssg, 255, 0)
Label9.Caption = mssg

End Sub
