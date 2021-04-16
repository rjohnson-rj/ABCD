VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "SALARY SLIP"
   ClientHeight    =   6330
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4335
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6330
   ScaleWidth      =   4335
   Begin VB.CommandButton Command1 
      Caption         =   "Print"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   840
      TabIndex        =   1
      Top             =   5640
      Width           =   1575
   End
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   1560
      TabIndex        =   32
      Text            =   "----Select Staff----"
      Top             =   720
      Width           =   2220
   End
   Begin VB.Label Label33 
      BackStyle       =   0  'Transparent
      Caption         =   " A/C No"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   0
      TabIndex        =   30
      Top             =   5160
      Width           =   855
   End
   Begin VB.Label Label22 
      BackStyle       =   0  'Transparent
      Caption         =   "Credit to"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   0
      TabIndex        =   29
      Top             =   4920
      Width           =   1815
   End
   Begin VB.Label Label21 
      BackStyle       =   0  'Transparent
      Caption         =   "Net Amount"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   0
      TabIndex        =   28
      Top             =   4680
      Width           =   1455
   End
   Begin VB.Label Label13 
      BackStyle       =   0  'Transparent
      Caption         =   "HARI SRI VIDYA NIDHI SCHOOL"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   0
      TabIndex        =   27
      Top             =   0
      Width           =   6855
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      Caption         =   "I/R"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   0
      TabIndex        =   26
      Top             =   4320
      Width           =   1215
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "PROF-TAX"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   0
      TabIndex        =   25
      Top             =   3960
      Width           =   1200
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "DA"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   0
      TabIndex        =   24
      Top             =   2160
      Width           =   1200
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Allowance"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   0
      TabIndex        =   23
      Top             =   2520
      Width           =   1200
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Basic Pay"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   0
      TabIndex        =   22
      Top             =   1800
      Width           =   1200
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Name"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   0
      TabIndex        =   21
      Top             =   720
      Width           =   1200
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "EPF"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   0
      TabIndex        =   20
      Top             =   3240
      Width           =   1200
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "ESI"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   0
      TabIndex        =   19
      Top             =   3600
      Width           =   1200
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Designation"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   0
      TabIndex        =   18
      Top             =   1080
      Width           =   1320
   End
   Begin VB.Label Label34 
      BackStyle       =   0  'Transparent
      Caption         =   "PUNKUNNAM, THRISSUR"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   17
      Top             =   240
      Width           =   7335
   End
   Begin VB.Label Label35 
      BackStyle       =   0  'Transparent
      Caption         =   "PAYSLIP - "
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   720
      TabIndex        =   16
      Top             =   480
      Width           =   975
   End
   Begin VB.Label Label19 
      BackStyle       =   0  'Transparent
      Caption         =   "Label19"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1560
      TabIndex        =   15
      Top             =   720
      Visible         =   0   'False
      Width           =   4815
   End
   Begin VB.Label Label36 
      BackStyle       =   0  'Transparent
      Caption         =   "Particulars"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   0
      TabIndex        =   14
      Top             =   1440
      Width           =   1335
   End
   Begin VB.Line Line4 
      X1              =   1440
      X2              =   1440
      Y1              =   720
      Y2              =   5400
   End
   Begin VB.Label Label20 
      BackStyle       =   0  'Transparent
      Caption         =   "Label20"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1560
      TabIndex        =   13
      Top             =   1080
      Visible         =   0   'False
      Width           =   4935
   End
   Begin VB.Label Label15 
      BackStyle       =   0  'Transparent
      Caption         =   "Label15"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1560
      TabIndex        =   12
      Top             =   1800
      Visible         =   0   'False
      Width           =   4935
   End
   Begin VB.Label Label16 
      BackStyle       =   0  'Transparent
      Caption         =   "Label16"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1560
      TabIndex        =   11
      Top             =   2160
      Visible         =   0   'False
      Width           =   5055
   End
   Begin VB.Label Label18 
      BackStyle       =   0  'Transparent
      Caption         =   "Label18"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1560
      TabIndex        =   10
      Top             =   2520
      Visible         =   0   'False
      Width           =   4935
   End
   Begin VB.Label Label17 
      BackStyle       =   0  'Transparent
      Caption         =   "Label17"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1560
      TabIndex        =   9
      Top             =   3120
      Visible         =   0   'False
      Width           =   4095
   End
   Begin VB.Label Label24 
      BackStyle       =   0  'Transparent
      Caption         =   "Label23"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1560
      TabIndex        =   8
      Top             =   3480
      Visible         =   0   'False
      Width           =   3855
   End
   Begin VB.Label Label26 
      BackStyle       =   0  'Transparent
      Caption         =   "Label23"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1560
      TabIndex        =   7
      Top             =   3840
      Visible         =   0   'False
      Width           =   4335
   End
   Begin VB.Label Label27 
      BackStyle       =   0  'Transparent
      Caption         =   "Label23"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1560
      TabIndex        =   6
      Top             =   4200
      Visible         =   0   'False
      Width           =   5175
   End
   Begin VB.Label Label25 
      BackStyle       =   0  'Transparent
      Caption         =   "Label25"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1560
      TabIndex        =   5
      Top             =   4680
      Visible         =   0   'False
      Width           =   3855
   End
   Begin VB.Label Label28 
      BackStyle       =   0  'Transparent
      Caption         =   "Label28"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1560
      TabIndex        =   4
      Top             =   5040
      Visible         =   0   'False
      Width           =   3255
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "================================================="
      Height          =   135
      Left            =   0
      TabIndex        =   3
      Top             =   5400
      Width           =   5655
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Amount"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2040
      TabIndex        =   2
      Top             =   1440
      Width           =   1815
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "Deductions"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   2880
      Width           =   1215
   End
   Begin VB.Label Label14 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1680
      TabIndex        =   31
      Top             =   480
      Width           =   1815
   End
   Begin VB.Label Label23 
      Caption         =   "------------------------------------------------------------------------------------------------"
      Height          =   135
      Index           =   0
      Left            =   0
      TabIndex        =   33
      Top             =   600
      Width           =   4335
   End
   Begin VB.Label Label23 
      Caption         =   "------------------------------------------------------------------------------------------------"
      Height          =   135
      Index           =   1
      Left            =   0
      TabIndex        =   34
      Top             =   1560
      Width           =   4335
   End
   Begin VB.Label Label23 
      Caption         =   "------------------------------------------------------------------------------------------------"
      Height          =   135
      Index           =   2
      Left            =   0
      TabIndex        =   35
      Top             =   3000
      Width           =   4335
   End
   Begin VB.Label Label23 
      Caption         =   "------------------------------------------------------------------------------------------------"
      Height          =   135
      Index           =   3
      Left            =   0
      TabIndex        =   36
      Top             =   2760
      Width           =   4335
   End
   Begin VB.Label Label23 
      Caption         =   "------------------------------------------------------------------------------------------------"
      Height          =   255
      Index           =   4
      Left            =   0
      TabIndex        =   37
      Top             =   4440
      Width           =   4335
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs As New ADODB.Recordset
Dim con As New ADODB.Connection
Public counted As Integer


Private Sub Form_Load()
Me.Width = 4575
Me.Height = 6900
If con.State = 1 Then con.Close
con.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Program Files (x86)\Project1\HariSri.mdb;Persist Security Info=False"

 Label14.Caption = Format(Now, "dd/MM/yyyy")
 Combo1.Visible = True
 
Call combofill

If counted = 1 Then
MsgBox "No Staff details found. Please add details", vbInformation, "Alert...!!!"
End If
End Sub

Private Sub combofill()
If rs.State = 1 Then rs.Close
rs.Open "select * from Salary_Table", con, adOpenDynamic

Dim i As Integer

With MSFlexGrid1

If rs.EOF <> True Then
     Do
     Combo1.AddItem (rs(0))
        i = i + 1
        rs.MoveNext
        
    Loop Until rs.EOF = True
    Combo1.Visible = True
'    Text1.Visible = False
'    Command3.Visible = True
End If

If rs.State = 1 Then rs.Close
rs.Open "select * from Salary_Table", con, adOpenDynamic
If rs.EOF = True Then
'Text1.Visible = True
'Text1.SetFocus
counted = 1
End If

End With
End Sub
Private Sub Combo1_Click()
If con.State = 1 Then con.Close
con.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Program Files (x86)\Project1\HariSri.mdb;Persist Security Info=False"

If rs.State = 1 Then rs.Close
rs.Open " select * from Salary_Table", con, adOpenDynamic
Dim i As Integer
Dim J As Integer
Dim k As Integer
Dim Str As String

J = 0
k = 0
i = 1
If rs.EOF <> True Then
     Do
     Dim a As String
     a = rs(0)
     If a = Combo1.Text Then
     Dim b As Integer
     If rs.State = 1 Then rs.Close
     rs.Open " select ID from Salary_Table where (Sname='" & Combo1.Text & "')"
     b = rs(0)
     End If
     rs.MoveNext
     Loop Until rs.EOF = True
     If b <> 0 Then
     If rs.State = 1 Then rs.Close
     rs.Open " select * from Salary_Table where (ID=" & Val(b) & ")"
     
  Combo1.Visible = False
Label19.Caption = rs(0)
Label20.Caption = rs(1)
Label15.Caption = rs(2)
Label16.Caption = rs(3)
Label18.Caption = rs(4)
Label17.Caption = rs(5)
Label24.Caption = rs(6)
Label26.Caption = rs(7)
Label27.Caption = rs(8)
Label25.Caption = rs(9)
Label28.Caption = rs(10)

Label19.Visible = True
Label20.Visible = True
Label15.Visible = True
Label16.Visible = True
Label18.Visible = True
Label17.Visible = True
Label24.Visible = True
Label26.Visible = True
Label25.Visible = True
Label27.Visible = True
Label28.Visible = True
End If
End If
End Sub



Private Sub Command1_Click()
If con.State = 1 Then con.Close
con.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Program Files (x86)\Project1\HariSri.mdb;Persist Security Info=False"

Call saved


End Sub

Private Sub saved()


Command1.Visible = False
Combo1.Visible = False
Label19.Visible = True
Label20.Visible = True
Label15.Visible = True
Label16.Visible = True
Label18.Visible = True
Label17.Visible = True
Label24.Visible = True
Label26.Visible = True
Label27.Visible = True
Label25.Visible = True
Label28.Visible = True
'On Error GoTo ErrorHandler
PrintForm
'ErrorHandler:
'    MsgBox Err.Description, , "Printing Cancelled"
''Printer.EndDoc
Command1.Visible = True
Label19.Visible = False
Label20.Visible = False
Label15.Visible = False
Label16.Visible = False
Label18.Visible = False
Label17.Visible = False
Label24.Visible = False
Label26.Visible = False
Label27.Visible = False
Label25.Visible = False
Label28.Visible = False

Combo1.Visible = True
Combo1.Clear
Combo1.AddItem ("----Select Staff----")
Call combofill
Combo1.SetFocus
Combo1.SelText = "----Select Staff----"

End Sub

