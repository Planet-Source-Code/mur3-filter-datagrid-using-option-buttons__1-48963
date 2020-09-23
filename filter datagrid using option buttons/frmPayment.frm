VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmPayment 
   Caption         =   "Student Payments"
   ClientHeight    =   6570
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9300
   LinkTopic       =   "Form1"
   ScaleHeight     =   6570
   ScaleWidth      =   9300
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame3 
      Caption         =   "Filter Payments By"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004000&
      Height          =   5655
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   9015
      Begin MSDataGridLib.DataGrid DGfilter 
         Height          =   4575
         Left            =   120
         TabIndex        =   6
         Top             =   840
         Width           =   8775
         _ExtentX        =   15478
         _ExtentY        =   8070
         _Version        =   393216
         AllowUpdate     =   0   'False
         DefColWidth     =   93
         HeadLines       =   1
         RowHeight       =   15
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnCount     =   2
         BeginProperty Column00 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column01 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
         EndProperty
         SplitCount      =   1
         BeginProperty Split0 
            BeginProperty Column00 
            EndProperty
            BeginProperty Column01 
            EndProperty
         EndProperty
      End
      Begin VB.OptionButton OptDue 
         Caption         =   "Due"
         Height          =   375
         Left            =   960
         TabIndex        =   2
         Top             =   360
         Width           =   735
      End
      Begin VB.OptionButton OptPaid 
         Caption         =   "Paid"
         Height          =   375
         Left            =   1800
         TabIndex        =   3
         Top             =   360
         Width           =   735
      End
      Begin VB.OptionButton OptAll 
         Caption         =   "All"
         Height          =   375
         Left            =   240
         TabIndex        =   1
         Top             =   360
         Width           =   495
      End
      Begin VB.OptionButton OptName 
         Caption         =   "Name"
         Height          =   375
         Left            =   2760
         TabIndex        =   5
         Top             =   360
         Width           =   855
      End
      Begin VB.TextBox txtFilterName 
         Height          =   375
         Left            =   3600
         TabIndex        =   4
         Top             =   360
         Width           =   1215
      End
   End
   Begin VB.Label Label1 
      Caption         =   "Code by Murshid (MUR3)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   7
      Top             =   6000
      Width           =   3615
   End
End
Attribute VB_Name = "frmPayment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'---------------------------------------------------------------------------------------
' Author    : Murshid
'---------------------------------------------------------------------------------------

Dim cn As Connection
Dim rs1 As Recordset

Private Sub Form_Load()

Set cn = New Connection
Set rs1 = New Recordset

Dim path As String

path = App.path & "\db1.mdb"
cn.Open "provider=microsoft.jet.oledb.4.0;data source=" & path & ";jet oledb:database password=NIM"
cn.CursorLocation = adUseClient

OptAll.Value = True 'executes optall_click() sub when form loads

End Sub

Private Sub optAll_Click()
'all student payments
Call fillGridOptAll

End Sub

Sub fillGridOptAll()
'sub to fillgrid with all students
If rs1.State = adStateOpen Then
    rs1.Close
End If

sqlAll = "SELECT S.RegNo as [Reg No],S.Name,s.TotalAmt as [Total Amount], s.PaidAmt as [Paid Amount],s.totalamt-s.paidamt as Due,S.Address,s.homephone"
sqlAll = sqlAll + " FROM  Students s"

'rs1 to DGfilter
rs1.Open sqlAll, cn, adOpenDynamic, adLockOptimistic
Set DGfilter.DataSource = rs1

gridSet 'set grid details , i.e. hide columns, set width etc...

End Sub

Private Sub OptDue_Click()
' due student payments
Call fillGridOptDue
End Sub

Sub fillGridOptDue()
'sub for due student payments

If rs1.State = adStateOpen Then
    rs1.Close
End If

Dim sqlDue As String

sqlDue = "SELECT S.RegNo as [Reg No],S.Name,s.TotalAmt as [Total Amount], s.PaidAmt as [Paid Amount],s.totalamt-s.paidamt as Due, S.Address,s.homephone"
sqlDue = sqlDue + " FROM  students s"
sqlDue = sqlDue + " Where s.paidamt<s.totalamt"

rs1.Open sqlDue
rs1.Requery
gridSet

End Sub


Private Sub OptPaid_Click()
' filter paid student payments
Call fillGridOptPaid
End Sub

Sub fillGridOptPaid()
' sub to filter paid student payments
If rs1.State = adStateOpen Then
    rs1.Close
End If

Dim sqlPaid As String

sqlPaid = "SELECT S.RegNo as [Reg No],S.Name,s.TotalAmt as [Total Amount], s.PaidAmt as [Paid Amount],s.totalamt-s.paidamt as Due,S.Address,s.homephone"
sqlPaid = sqlPaid + " FROM  Students s"
sqlPaid = sqlPaid + " Where (s.paidamt=s.totalamt or s.paidamt>s.totalamt)"

rs1.Open sqlPaid
rs1.Requery
gridSet

End Sub


Private Sub optName_Click()
' filter student payments by name
txtFilterName.Text = ""
Call fillGridOptAll
txtFilterName.SetFocus
End Sub

Private Sub txtFilterName_Change()
'filters items according to textbox content
If OptName.Value = False Then
    MsgBox "Please select Option - Filter by Name", vbCritical
    Exit Sub
Else
    Call fillgridOptName
End If
End Sub

Sub fillgridOptName()
' sub to filter student payments by name
If rs1.State = adStateOpen Then
    rs1.Close
End If

Dim sqlName As String

sqlName = "SELECT S.RegNo as [Reg No],S.Name,s.TotalAmt as [Total Amount], s.PaidAmt as [Paid Amount],s.totalamt-s.paidamt as Due, S.Address,s.homephone"
sqlName = sqlName + " FROM  Students s"
sqlName = sqlName + " Where s.name like '" & txtFilterName.Text & "%' "

rs1.Open sqlName
rs1.Requery
gridSet

End Sub

Sub gridSet()
DGfilter.Columns(0).Width = 750
DGfilter.Columns(1).Width = 1500
DGfilter.Columns(2).Width = 1500
DGfilter.Columns(3).Width = 1500
DGfilter.Columns(4).Width = 1500
DGfilter.Columns(5).Width = 1500

DGfilter.Columns(5).Visible = False
DGfilter.Columns(6).Visible = False

End Sub



