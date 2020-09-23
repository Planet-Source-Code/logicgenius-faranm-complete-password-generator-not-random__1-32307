VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmGenerate 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Password Generator"
   ClientHeight    =   4095
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7875
   Icon            =   "frmGenerate.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4095
   ScaleWidth      =   7875
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "Generating Options"
      Height          =   2535
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   4335
      Begin VB.CheckBox lblAlpha 
         Caption         =   "Generate Alphabets (A to Z)"
         Height          =   255
         Left            =   600
         TabIndex        =   7
         Top             =   2040
         Value           =   1  'Checked
         Width           =   2415
      End
      Begin VB.CheckBox lblSpace 
         Caption         =   "Allow Spaces"
         Height          =   255
         Left            =   600
         TabIndex        =   6
         Top             =   840
         Width           =   1335
      End
      Begin VB.CheckBox lblNum 
         Caption         =   "Mingle Numbers (0 to 9)"
         Height          =   255
         Left            =   600
         TabIndex        =   5
         Top             =   1680
         Value           =   1  'Checked
         Width           =   2055
      End
      Begin VB.CheckBox lblChar 
         Caption         =   "Manipulate Specail Chars"
         Height          =   255
         Left            =   600
         TabIndex        =   4
         Top             =   1440
         Width           =   2175
      End
      Begin VB.CheckBox lblCase 
         Caption         =   "Allow Lower Case"
         Height          =   255
         Left            =   600
         TabIndex        =   3
         Top             =   1080
         Width           =   1575
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H80000018&
         Height          =   360
         Left            =   1800
         TabIndex        =   2
         Top             =   240
         Width           =   2055
      End
      Begin VB.Label lblDigits 
         Caption         =   "Digit Manipulation :"
         Height          =   255
         Left            =   360
         TabIndex        =   1
         Top             =   315
         Width           =   1455
      End
   End
   Begin VB.Frame Frame2 
      BorderStyle     =   0  'None
      Caption         =   "Frame2"
      Height          =   2415
      Left            =   4920
      TabIndex        =   10
      Top             =   360
      Width           =   2655
      Begin VB.CommandButton Command2 
         Caption         =   "Exit"
         Height          =   375
         Left            =   600
         TabIndex        =   12
         Top             =   720
         Width           =   1335
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Start Now"
         Default         =   -1  'True
         Height          =   375
         Left            =   600
         TabIndex        =   11
         Top             =   360
         Width           =   1335
      End
      Begin VB.Label Label2 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "XXX Possibilities"
         Height          =   375
         Left            =   180
         TabIndex        =   13
         Top             =   1680
         Width           =   2415
      End
   End
   Begin MSComctlLib.StatusBar Bar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   9
      Top             =   3720
      Width           =   7875
      _ExtentX        =   13891
      _ExtentY        =   661
      Style           =   1
      SimpleText      =   " Status: Waiting for a Command"
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   375
      Left            =   240
      TabIndex        =   8
      Top             =   3000
      Width           =   7335
      _ExtentX        =   12938
      _ExtentY        =   661
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
   Begin VB.Timer tmrCheck 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   3240
      Top             =   1680
   End
   Begin MSComDlg.CommonDialog Dlg 
      Left            =   3240
      Top             =   1680
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "frmGenerate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private zLen As Long
Private yLen As Integer
Private xLen As Integer

Private CResult As Currency
Private booTask As Boolean

Private str As String
Private strList As String

Option Explicit

Sub Em(xy As Boolean)
 On Error GoTo STOPNOW
  
 Dim x As Long
   
  For x = 1 To Len(strList)
       str = str & Mid(strList, x, 1)

       If xy = False And Len(str) = xLen - 1 Then
            yLen = yLen + 1
            Em True
            yLen = yLen - 1
       Else
            If xy = False And Len(str) <> xLen Then
                yLen = yLen + 1
                Em False
                yLen = yLen - 1
            End If
       End If
        
       DoEvents
        
       If Len(str) = xLen Then
         If booTask Then Exit Sub
                    
            Print #2, (str)
            zLen = zLen + 1
            
            DoEvents
            DoEvents
       End If
                    
       str = Left(str, yLen - 1)
  Next
  
STOPNOW:
End Sub

Private Sub Command1_Click()
     Start
End Sub

Sub Start()
  
  On Error GoTo NOTEXT
    
    xLen = Text1.Text
    
    Text1.Enabled = False
    Command1.Enabled = False
    
    lblAlpha.Enabled = False
    lblCase.Enabled = False
    lblChar.Enabled = False
    lblNum.Enabled = False
    lblSpace.Enabled = False
    
    ProgressBar1.Value = 0
    
    Let yLen = 1
    
    Query
    
    CResult = Len(strList) ^ (Text1.Text)
    Label2.Caption = CResult & " Possibilites"
    
    tmrCheck.Enabled = True
    Command2.Caption = "Stop"
        
    Dlg.CancelError = True
    Dlg.FileName = "Pwds.dat"
    Dlg.Filter = "Data Files (*.dat)|*.dat|All Files (*.*)|*.*"
    Dlg.ShowSave
    
    Open Dlg.FileName For Output As #2
        Em False
    Close #2

    Text1.Enabled = True
    Command1.Enabled = True
    Command2.Caption = "Exit"
    tmrCheck.Enabled = False
    
    lblAlpha.Enabled = True
    lblCase.Enabled = True
    lblChar.Enabled = True
    lblNum.Enabled = True
    lblSpace.Enabled = True

    Bar1.SimpleText = " Status : Process Terminated"
    ProgressBar1.Value = 100
    
    If booTask Then
        MsgBox "Task canceled by user !", vbInformation, "Password Generator"
    Else
        MsgBox "The file has been generated !", vbInformation, "Password Generator"
    End If
    
    zLen = 0: yLen = 0: xLen = 0
    booTask = False
    
    str = vbNullString
    
NOTEXT:
    Text1.Enabled = True
    Command1.Enabled = True
    tmrCheck.Enabled = False
    
    lblAlpha.Enabled = True
    lblCase.Enabled = True
    lblChar.Enabled = True
    lblNum.Enabled = True
    lblSpace.Enabled = True
End Sub

Private Sub Command2_Click()
 If tmrCheck.Enabled = True Then
    booTask = True
    Close #2
 Else
    Unload Me
    Set frmGenerate = Nothing
 End If
End Sub

Private Sub Text1_Change()
  On Error GoTo TooLong
   
    If IsNumeric(Text1.Text) = False Then Text1.Text = "": Exit Sub
    
    Query
    
    If Text1.Text = "0" Then CResult = 0: GoTo DISPLAY
    If Len(strList) = 1 Then CResult = Text1.Text: GoTo DISPLAY
    
    CResult = Len(strList) ^ (Text1.Text)

DISPLAY:
    Label2.Caption = CResult & " Possibilites"
    Exit Sub
    
TooLong:
    Label2.Caption = "Too Long to Fit"
End Sub

Private Sub tmrCheck_Timer()
 On Error Resume Next
   ProgressBar1.Value = 100 * (zLen / CResult)
   Bar1.SimpleText = " Status: Done " & zLen & " out of " & CResult
End Sub

Sub Query()
    strList = ""
    
    If lblAlpha.Value = 1 Then strList = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
    If lblSpace.Value = 1 Then strList = strList & Chr$(32)
    If lblNum.Value = 1 Then strList = strList & "0123456789"
    If lblCase.Value = 1 Then strList = strList & "abcdefghijklmnopqrstuvwxyz"
    If lblChar.Value = 1 Then strList = strList & "`~!@#$%^&*()-_=+[{]};:',<.>/?\|" & Chr$(34)
End Sub
