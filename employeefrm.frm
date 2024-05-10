VERSION 5.00
Begin VB.Form employeefrm 
   Caption         =   "Employee Form"
   ClientHeight    =   7155
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   15375
   LinkTopic       =   "Form1"
   ScaleHeight     =   7155
   ScaleWidth      =   15375
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdClose 
      Caption         =   "CLOSE"
      Height          =   1095
      Left            =   11520
      Picture         =   "employeefrm.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   3120
      Width           =   1335
   End
   Begin VB.CommandButton cmdFind 
      Caption         =   "FIND"
      Height          =   1095
      Left            =   10080
      Picture         =   "employeefrm.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   3120
      Width           =   1335
   End
   Begin VB.CommandButton cmddelete 
      Caption         =   "DELETE"
      Height          =   1095
      Left            =   12960
      Picture         =   "employeefrm.frx":0884
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   1680
      Width           =   1335
   End
   Begin VB.CommandButton cmdsave 
      Caption         =   "SAVE"
      Height          =   1095
      Left            =   11520
      Picture         =   "employeefrm.frx":0CC6
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   1680
      Width           =   1335
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "ADD"
      DisabledPicture =   "employeefrm.frx":1108
      DownPicture     =   "employeefrm.frx":154A
      DragIcon        =   "employeefrm.frx":198C
      Height          =   1095
      Left            =   10080
      Picture         =   "employeefrm.frx":1DCE
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   1680
      Width           =   1335
   End
   Begin VB.TextBox Text6 
      Height          =   735
      Left            =   2400
      TabIndex        =   11
      Top             =   4680
      Width           =   6495
   End
   Begin VB.TextBox txtsalary 
      Height          =   735
      Left            =   2400
      TabIndex        =   9
      Top             =   3840
      Width           =   6495
   End
   Begin VB.TextBox txtposition 
      Height          =   735
      Left            =   2400
      TabIndex        =   7
      Top             =   3000
      Width           =   6495
   End
   Begin VB.TextBox txtaddress 
      Height          =   735
      Left            =   2400
      TabIndex        =   6
      Top             =   2160
      Width           =   6495
   End
   Begin VB.TextBox txtname 
      Height          =   735
      Left            =   2400
      TabIndex        =   5
      Top             =   1320
      Width           =   6495
   End
   Begin VB.TextBox txtEmployeeID 
      Height          =   735
      Left            =   2400
      TabIndex        =   4
      Top             =   480
      Width           =   6495
   End
   Begin VB.Label Label6 
      Caption         =   "Date Hired"
      Height          =   495
      Left            =   480
      TabIndex        =   10
      Top             =   4800
      Width           =   1695
   End
   Begin VB.Label Label5 
      Caption         =   "Salary"
      Height          =   495
      Left            =   480
      TabIndex        =   8
      Top             =   3960
      Width           =   1695
   End
   Begin VB.Label Label4 
      Caption         =   "Position"
      Height          =   495
      Left            =   480
      TabIndex        =   3
      Top             =   3120
      Width           =   1695
   End
   Begin VB.Label Label3 
      Caption         =   "Address"
      Height          =   495
      Left            =   480
      TabIndex        =   2
      Top             =   2280
      Width           =   1695
   End
   Begin VB.Label Label2 
      Caption         =   "Name"
      Height          =   495
      Left            =   480
      TabIndex        =   1
      Top             =   1440
      Width           =   1695
   End
   Begin VB.Label Label1 
      Caption         =   "Employee ID"
      Height          =   495
      Left            =   480
      TabIndex        =   0
      Top             =   600
      Width           =   1695
   End
End
Attribute VB_Name = "employeefrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAdd_Click()
txtEmployeeID.Text = ""
txtname.Text = ""
txtaddress.Text = ""
txtposition.Text = ""
txtsalary.Text = ""
txtdatehired.Text = ""
txtEmployeeID.SetFocus

End Sub

Private Sub cmdClose_Click()
Unload Me

End Sub

Private Sub cmddelete_Click()
conPayroll.Execute "Delete * from employee where employeeid='" & Trim(txtEmployeeID.Text) & "'"
MsgBox "Record has been deleted.."
End Sub

Private Sub cmdFind_Click()
txtEmployeeID.SelStart = 0
txtEmployeeID.SelLength = Len(txtEmployeeID.Text)
txtEmployeeID.SetFocus

End Sub

Private Sub cmdsave_Click()
openrstEmployee "Select * from employee where employeeid='" & Trim(txtEmployeeID.Text) & "'"
If Not rstEmployee.EOF Then
    'not found
    With rstEmployee
        .Edit
            .Fields("employeeid").Value = txtEmployeeID.Text
            .Fields("employeename").Value = txtname.Text
            .Fields("address").Value = txtaddress.Text
            .Fields("position").Value = txtposition.Text
            .Fields("salary").Value = txtsalary.Text
            .Fields("datehired").Value = txtdatehired.Text
            
        .Update
        
        
    End With
Else
    'found
        With rstEmployee
        .AddNew
             .Fields("employeeid").Value = txtEmployeeID.Text
            .Fields("employeename").Value = txtname.Text
            .Fields("address").Value = txtaddress.Text
            .Fields("position").Value = txtposition.Text
            .Fields("salary").Value = txtsalary.Text
            .Fields("datehired").Value = txtdatehired.Text
        .Update
        
        End With
End If

    
End Sub

Private Sub Form_Load()
openWORKSPACEODBC
openconPayroll
End Sub

Private Sub txtaddress_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txtposition.SetFocus
End If

End Sub

Private Sub txtEmployeeID_KeyPress(KeyAscii As Integer)
On Error Resume Next
If KeyAscii = 13 Then
    openrstEmployee "Select * from employee where employeeid ='" & Trim(txtEmployeeID.Text) & "'"
     If Not rstEmployee.EOF Then
        With rstEmployee
            txtEmployeeID.Text = .Fields("employeeid").Value
            txtname.Text = .Fields("employeename").Value
            txtaddress.Text = .Fields("address").Value
            txtposition.Text = .Fields("position").Value
            txtsalary.Text = .Fields("salary").Value
            txtdatehired.Text = .Fields("datehired").Value
        End With
    End If
    
End If

End Sub

Private Sub txtname_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txtaddress.SetFocus
End If

End Sub

Private Sub txtposition_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txtsalary.SetFocus
End If

End Sub

Private Sub txtsalary_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txtdatehired.SetFocus
End If
End Sub
