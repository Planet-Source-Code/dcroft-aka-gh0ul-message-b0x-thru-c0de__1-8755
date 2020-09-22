VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Dynamic Message Box"
   ClientHeight    =   3744
   ClientLeft      =   36
   ClientTop       =   264
   ClientWidth     =   4092
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3744
   ScaleWidth      =   4092
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   252
      Left            =   1920
      ScaleHeight     =   17
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   37
      TabIndex        =   16
      Top             =   3720
      Visible         =   0   'False
      Width           =   492
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Client Edge"
      Height          =   192
      Left            =   2280
      TabIndex        =   15
      Top             =   2280
      Value           =   1  'Checked
      Width           =   1692
   End
   Begin VB.TextBox Text5 
      Height          =   288
      Left            =   480
      TabIndex        =   14
      Text            =   "300"
      Top             =   3360
      Width           =   492
   End
   Begin VB.TextBox Text4 
      Height          =   288
      Left            =   480
      TabIndex        =   11
      Text            =   "400"
      Top             =   3000
      Width           =   492
   End
   Begin VB.CommandButton Command3 
      Caption         =   "..."
      Enabled         =   0   'False
      Height          =   288
      Left            =   3720
      TabIndex        =   9
      Top             =   2520
      Width           =   372
   End
   Begin VB.TextBox Text3 
      Enabled         =   0   'False
      Height          =   288
      Left            =   0
      TabIndex        =   8
      Top             =   2520
      Width           =   3612
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Exit"
      Height          =   252
      Left            =   1320
      TabIndex        =   7
      Top             =   3480
      Width           =   1332
   End
   Begin VB.ListBox List1 
      Height          =   240
      Left            =   2160
      TabIndex        =   6
      Top             =   360
      Width           =   1932
   End
   Begin VB.TextBox Text2 
      Height          =   1212
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   4
      Top             =   960
      Width           =   4092
   End
   Begin VB.TextBox Text1 
      Height          =   288
      Left            =   0
      TabIndex        =   2
      Top             =   360
      Width           =   2052
   End
   Begin VB.CommandButton cmdTest 
      Caption         =   "Test "
      Height          =   252
      Left            =   2760
      TabIndex        =   0
      Top             =   3480
      Width           =   1332
   End
   Begin VB.Label Label4 
      Caption         =   "gh0ul 2000        (ICQ#31047555)"
      Height          =   252
      Left            =   1320
      TabIndex        =   17
      Top             =   3120
      Width           =   2652
   End
   Begin VB.Label Label1 
      Caption         =   "Y"
      Height          =   252
      Index           =   3
      Left            =   240
      TabIndex        =   13
      Top             =   3360
      Width           =   252
   End
   Begin VB.Label Label1 
      Caption         =   "X"
      Height          =   252
      Index           =   2
      Left            =   240
      TabIndex        =   12
      Top             =   3000
      Width           =   252
   End
   Begin VB.Label Label1 
      Caption         =   "Icon Location:"
      Enabled         =   0   'False
      Height          =   252
      Index           =   1
      Left            =   0
      TabIndex        =   10
      Top             =   2280
      Width           =   2052
   End
   Begin VB.Label Label3 
      Caption         =   "Type"
      Height          =   252
      Left            =   2160
      TabIndex        =   5
      Top             =   120
      Width           =   2052
   End
   Begin VB.Label Label2 
      Caption         =   "Message"
      Height          =   252
      Left            =   0
      TabIndex        =   3
      Top             =   720
      Width           =   2052
   End
   Begin VB.Label Label1 
      Caption         =   "Caption"
      Height          =   252
      Index           =   0
      Left            =   0
      TabIndex        =   1
      Top             =   120
      Width           =   2052
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim sCaption As String
Dim sMsg As String
Dim lType As Long
Dim bEdge As Boolean

Private Sub Check1_Click()
   If Check1.Value = vbChecked Then
      bEdge = True
   Else
      bEdge = False
   End If
End Sub

Private Sub cmdTest_Click()
    Dim rv As Integer
    Dim x As Long, y As Long
    
    x = CLng(Text4.Text)
    y = CLng(Text5.Text)
    
    ' call with one simple line of code.
    rv% = MBox(sMsg, _
               sCaption, _
               lType, lType, Text3, _
               x, _
               y, _
               bEdge _
              )
m_BoX:
    rv% = MBox("Test the buttons...", "Simple Question", BX_YESNOCANCEL, 0&, "", x, y, bEdge)
    
    If rv% = mbYes Then
       MBox "YEs was CHosen", "Yes", BX_OK, 0&, "", x, y, bEdge
    ElseIf rv% = mbNo Then
       MBox "No Was Chosen", "No", BX_OK, 0&, "", x, y, bEdge
    Else
       MBox "Cancel Was Chosen.", "Cancel", BX_OK, 0&, "", x, y, bEdge
    End If
    
    rv = MBox("Go Again??", "Once More?", BX_YESNO, 0&, "", x, y, bEdge)
          
    If rv = mbYes Then
       GoTo m_BoX
    End If
    
    
End Sub

Private Sub Command2_Click()
    Unload Me
End Sub

Private Sub Command3_Click()
    Dim Fname As String
    
    Fname = Open_File(Me.hwnd, "Open", _
            "All Files (*.*)" + Chr$(0) + "*.*" + Chr$(0))
    
    If Fname = "" Then Exit Sub
    
    Text3 = Fname

End Sub


Private Sub Form_Load()
    
    With Form1
      .ScaleMode = 3
      .AutoRedraw = True
    End With
    
    With List1
      .AddItem "None", 0
      .AddItem "OK", 1
      .AddItem "OK_CANCEL", 2
      .AddItem "YES_NO", 3
      .AddItem "YES_NO_CANCEL", 4
      .AddItem "Exclamation", 5
      .AddItem "Critical", 6
      .AddItem "Information", 7
      .AddItem "Confirm", 8
      .AddItem "Custom", 9
    End With
    
    ' default
    lType = 1&
    
End Sub

Private Sub List1_Click()
     If List1.Text = "Custom" Then
        Text3.Enabled = True
        Label1(1).Enabled = True
        Command3.Enabled = True
        lType = IC_CUSTOM
     Else
        Text3.Enabled = False
        Label1(1).Enabled = False
        Command3.Enabled = False
        
        lType = CLng(List1.ListIndex)
     End If
     
End Sub

Private Sub Text1_Change()
    sCaption = Text1
End Sub

Private Sub Text2_Change()
   sMsg = Text2
End Sub

