VERSION 5.00
Begin VB.Form frm 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00800000&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "LCM - GCF Calculator"
   ClientHeight    =   1920
   ClientLeft      =   11565
   ClientTop       =   6555
   ClientWidth     =   5550
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   161
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H000000FF&
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1908.901
   ScaleMode       =   0  'User
   ScaleWidth      =   5550
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtSimplify 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   12
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF00FF&
      Height          =   390
      Left            =   0
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   1080
      Width           =   5595
   End
   Begin VB.TextBox txtInCrease 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   12
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   390
      Left            =   0
      Locked          =   -1  'True
      TabIndex        =   5
      Top             =   360
      Width           =   5595
   End
   Begin VB.TextBox txtEkp 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   390
      Left            =   0
      Locked          =   -1  'True
      TabIndex        =   4
      Top             =   0
      Width           =   5595
   End
   Begin VB.TextBox txtMkd 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   12
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF00FF&
      Height          =   390
      Left            =   0
      Locked          =   -1  'True
      TabIndex        =   2
      Top             =   720
      Width           =   5595
   End
   Begin VB.TextBox txt 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   12
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF8080&
      Height          =   495
      Left            =   0
      TabIndex        =   1
      Top             =   1440
      Width           =   5595
   End
   Begin VB.CommandButton CmdCheck 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "Calc"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   1920
      Width           =   5625
   End
End
Attribute VB_Name = "frm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim iNumber(1 To 50), Max, Min, iEkp, iMkd, iRest, Multiple, Divisor As Double
Dim n, ValuePower As Integer
Dim strNumbers, strEkp, strMkd As String

Sub FindEkp()
On Error Resume Next
Multiple = 0
Do
 iRest = 0
 Multiple = Multiple + 1
 iEkp = Max * Multiple
  For i = 1 To n
  If Max <> iNumber(i) Then
    iRest = iRest + Int(iEkp) Mod Int(iNumber(i))
  End If
  Next i
  
  If iEkp > 99999999 Then
  iEkp = 0
  iRest = 0
  End If
Loop Until iRest = 0

End Sub

Sub InCrease()
Dim iInCrease As String
iInCrease = ""

For i = 1 To n
 If iInCrease = "" Then
  strkomma = ""
 Else
  strkomma = ","
 End If
iInCrease = iInCrease & strkomma & (iEkp / iNumber(i))
Next i

txtInCrease.Text = iInCrease
End Sub

Sub FindMkd()
On Error Resume Next
Divisor = 0
Do
  iRest = 0
  Divisor = Divisor + 1
 If Int(Min) Mod Divisor = 0 Then
  iMkd = Min / Divisor
 End If
  For i = 1 To n
  If Int(Min) <> Int(iNumber(i)) Then
   iRest = iRest + (Int(iNumber(i)) Mod Int(iMkd))
  End If
  Next
  
  If Divisor > 99999999 Then
  Debug.Print "INF"
  iRest = 0
  End If
Loop Until iRest = 0
End Sub

Sub Simplify()
On Error Resume Next
txtSimplify.Text = ""
For i = 1 To n
 If txtSimplify.Text = "" Then
  strkomma = ""
 Else
  strkomma = ","
 End If
 
If iMkd > 0 Then
txtSimplify.Text = txtSimplify & strkomma & (iNumber(i) / iMkd)
Else
txtSimplify.Text = "Inf"
End If
Next i
txtSimplify.Text = txtSimplify.Text

End Sub

Sub InsertNumbers()
On Error Resume Next
Dim i, m As Integer
'Eisagogi arithwn apo textbox
i = 1
n = 0
Max = 0
Min = 0
strNumbers = ""
While Mid(txt, i, 1) <> ""
  If Mid(txt, i, 1) <> "," Then
    m = i
     Do Until Mid(txt, m, 1) = ""
       If Mid(txt, m, 1) = "," Then
        Exit Do
       End If
      m = m + 1
     Loop
      m = m - i
  End If
  n = n + 1
  iNumber(n) = Mid(txt, i, m)
  If Int(iNumber(n)) > Int(Max) Then 'Eyresi Megaliterou arithmou
  Max = Int(iNumber(n))
  End If
 If strNumbers = "" Then
  strkomma = ""
 Else
  strkomma = ","
 End If
 strNumbers = strNumbers & strkomma & iNumber(n)
 strEkp = "LCM(" & strNumbers & ") = "
 strMkd = "GCF(" & strNumbers & ") = "
  i = i + m
i = i + 1
Wend
Min = Max
For i = 1 To n
 If Int(iNumber(i)) < Int(Min) Then 'Eyresi Mikroterou arithmou
 Min = Int(iNumber(i))
 End If
Next i
End Sub

Private Sub CmdCheck_Click()
frm.Cls
Call InsertNumbers

Call FindEkp
If iEkp = 0 Then
txtEkp.Text = strEkp & "Inf"
Else
txtEkp.Text = strEkp & "" & iEkp
End If

Call InCrease


Call FindMkd
If iMkd = 0 Then
txtMkd.Text = strMkd & "Inf"
Else
txtMkd.Text = strMkd & "" & iMkd
End If

Call Simplify

If txt.Text = "" Then
txtEkp.Text = ""
txtMkd.Text = ""
txtInCrease.Text = ""
End If
End Sub

Private Sub Form_Activate()
txt.SetFocus
ValuePower = 0
End Sub

Private Sub Form_Load()
Call CmdCheck_Click
End Sub

Private Sub txt_Change()

If Mid(txt.Text, 1, 1) = "0" Or Mid(txt.Text, 1, 1) = "," Then
txt.Text = Mid(txt.Text, 2, Len(txt.Text))
Beep
End If

If txt.Text = "," Or txt.Text = "0" Then
txt.Text = ""
End If
CmdCheck_Click

End Sub

Private Sub txt_KeyPress(KeyAscii As Integer)
If (KeyAscii = 32 Or KeyAscii = 46 Or KeyAscii = 47 Or KeyAscii = 13) Then KeyAscii = 44
 If txt.Text = "" And (KeyAscii = 32 Or KeyAscii = 44) Then
 KeyAscii = 0
End If
 If txt.Text <> "" Then
  If Mid(txt.Text, Len(txt.Text), 1) = "," And (KeyAscii = 32 Or KeyAscii = 44) Then
  KeyAscii = 0
  End If
   If Mid(txt.Text, Len(txt.Text), 1) = "," And KeyAscii = 48 Then
   KeyAscii = 0
   Debug.Print "dd"
   End If
 Else
 If KeyAscii = 48 Then KeyAscii = 0
 End If
 If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 And KeyAscii <> 46 And KeyAscii <> 44 Then
  KeyAscii = 0
 End If
  If KeyAscii = 0 Then Beep
End Sub

