VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form View_details 
   Caption         =   "SALARY DETAILS"
   ClientHeight    =   6105
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   14610
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   6105
   ScaleWidth      =   14610
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   8160
      Top             =   5760
      Visible         =   0   'False
      Width           =   5295
      _ExtentX        =   9340
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Program Files (x86)\Project1\HariSri.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Program Files (x86)\Project1\HariSri.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   3015
      Left            =   120
      TabIndex        =   6
      Top             =   600
      Width           =   14415
      _ExtentX        =   25426
      _ExtentY        =   5318
      _Version        =   393216
      Cols            =   11
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Delete Staff Details"
      Height          =   495
      Left            =   9000
      TabIndex        =   4
      Top             =   3720
      Width           =   2895
   End
   Begin VB.Frame Frame1 
      Caption         =   "Delete"
      Height          =   1455
      Left            =   120
      TabIndex        =   0
      Top             =   3600
      Visible         =   0   'False
      Width           =   5535
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   1680
         TabIndex        =   2
         Text            =   "---------Select Name--------"
         Top             =   240
         Width           =   3375
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Delete"
         Height          =   375
         Left            =   1800
         TabIndex        =   1
         Top             =   840
         Width           =   1695
      End
      Begin VB.Label Label2 
         Caption         =   "Staff Name       :"
         Height          =   375
         Left            =   240
         TabIndex        =   3
         Top             =   360
         Width           =   1695
      End
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "SALARY DETAILS"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   4455
   End
End
Attribute VB_Name = "View_details"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs As New ADODB.Recordset
Dim con As New ADODB.Connection

Private Sub Command1_Click()
Frame1.Visible = True
If rs.State = 1 Then rs.Close
rs.Open " select * from Salary_Table", con, adOpenDynamic

Dim i As Integer
i = 1
With MSFlexGrid1
If rs.EOF <> True Then
     Do
     Combo1.AddItem (rs(0))
        i = i + 1
        rs.MoveNext
        i = 1
    Loop Until rs.EOF = True
End If
End With
End Sub

Private Sub Command2_Click()
Call connect
If rs.State = 1 Then rs.Close
rs.Open " select ID from Salary_Table where (sname='" & Combo1.Text & "')", con, adOpenDynamic
Dim f As Integer
f = rs(0)
con.Execute "delete from Salary_Table where ID=" & Val(f) & ""
'con.Execute "delete from Salary_Table where Sname='" & Combo1.Text & "'"
MsgBox ("One Record deleted")
Call fillflexgrid

If rs.State = 1 Then rs.Close
rs.Open " select * from Salary_Table", con, adOpenDynamic

Dim i As Integer
i = 1
Combo1.Clear
With MSFlexGrid1
If rs.EOF <> True Then
     Do
     Combo1.AddItem (rs(0))
        i = i + 1
        rs.MoveNext
        i = 1
    Loop Until rs.EOF = True
End If
End With

End Sub

Private Sub Form_Load()
Call connect
Call fillflexgrid

End Sub

Public Sub connect()
If con.State = 1 Then con.Close
'con.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Users\extremaa\Desktop\HariSri Bill Format\HariSri.mdb;Persist Security Info=False"
con.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Program Files (x86)\Project1\HariSri.mdb;Persist Security Info=False"
End Sub

Public Sub fillflexgrid()
MSFlexGrid1.ColWidth(0) = 3000
MSFlexGrid1.ColWidth(1) = 1500
MSFlexGrid1.ColWidth(10) = 2000
If rs.State = 1 Then rs.Close
rs.Open " select * from Salary_Table", con, adOpenDynamic
Dim i As Integer
i = 1
With MSFlexGrid1
.Clear
.Rows = 2
'ALL for heading if you need
.TextMatrix(0, 0) = "STAFF NAME "
.TextMatrix(0, 1) = "DESIGNATION"
.TextMatrix(0, 2) = "BASIC PAY"
.TextMatrix(0, 3) = "DA"
.TextMatrix(0, 4) = "ALLOWANCE"


.TextMatrix(0, 5) = "EPF"
.TextMatrix(0, 6) = "ESI"
.TextMatrix(0, 7) = "PROF-TAX"
.TextMatrix(0, 8) = "I/R"
.TextMatrix(0, 9) = "TOTAL"
.TextMatrix(0, 10) = "A/C NO"
  
If rs.EOF <> True Then
     Do
       If sxx <> rs(0) Then
        .TextMatrix(i, 0) = rs(0)
        .TextMatrix(i, 1) = rs(1)
        .TextMatrix(i, 2) = rs(2)
        .TextMatrix(i, 3) = rs(3)
        .TextMatrix(i, 4) = rs(4)
        .TextMatrix(i, 5) = rs(5)
        .TextMatrix(i, 6) = rs(6)
        .TextMatrix(i, 7) = rs(7)
        .TextMatrix(i, 8) = rs(8)
        .TextMatrix(i, 9) = rs(9)
        .TextMatrix(i, 10) = rs(10)
        i = i + 1
        .Rows = .Rows + 1
        End If
        rs.MoveNext
       Loop Until rs.EOF = True
 End If
 End With
End Sub


