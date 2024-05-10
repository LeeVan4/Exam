VERSION 5.00
Begin VB.Form payrollform 
   Caption         =   "Payroll Form"
   ClientHeight    =   11850
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   18285
   LinkTopic       =   "Form2"
   ScaleHeight     =   11850
   ScaleWidth      =   18285
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdGross 
      Caption         =   "Compute Gross"
      Height          =   735
      Left            =   360
      TabIndex        =   52
      Top             =   10200
      Width           =   3975
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Height          =   735
      Left            =   13680
      TabIndex        =   51
      Top             =   6360
      Width           =   1815
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Print"
      Height          =   735
      Left            =   13680
      TabIndex        =   50
      Top             =   5280
      Width           =   1815
   End
   Begin VB.CommandButton cmdFind 
      Caption         =   "Find"
      Height          =   735
      Left            =   13680
      TabIndex        =   49
      Top             =   4200
      Width           =   1815
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "Delete"
      Height          =   735
      Left            =   13680
      TabIndex        =   48
      Top             =   3000
      Width           =   1815
   End
   Begin VB.CommandButton cmdsave 
      Caption         =   "Save"
      Height          =   735
      Left            =   13680
      TabIndex        =   47
      Top             =   1920
      Width           =   1815
   End
   Begin VB.CommandButton cmdadd 
      Caption         =   "Add"
      Height          =   735
      Left            =   13680
      TabIndex        =   46
      Top             =   840
      Width           =   1815
   End
   Begin VB.TextBox txtnetincome 
      Height          =   735
      Left            =   10440
      TabIndex        =   45
      Top             =   9960
      Width           =   2295
   End
   Begin VB.TextBox txttotdeduction 
      Height          =   735
      Left            =   10440
      TabIndex        =   44
      Top             =   9120
      Width           =   2295
   End
   Begin VB.CommandButton cmdnet 
      Caption         =   "Net Income"
      Height          =   735
      Left            =   8520
      TabIndex        =   43
      Top             =   9960
      Width           =   1815
   End
   Begin VB.CommandButton cmddeduct 
      Caption         =   "Compute Deduction"
      Height          =   735
      Left            =   8520
      TabIndex        =   42
      Top             =   9120
      Width           =   1815
   End
   Begin VB.TextBox txtpag 
      Height          =   735
      Left            =   10440
      TabIndex        =   41
      Top             =   8160
      Width           =   2295
   End
   Begin VB.TextBox txtphil 
      Height          =   735
      Left            =   10440
      TabIndex        =   39
      Top             =   7320
      Width           =   2295
   End
   Begin VB.TextBox txttax 
      Height          =   735
      Left            =   10440
      TabIndex        =   37
      Top             =   6480
      Width           =   2295
   End
   Begin VB.TextBox txtsss1 
      Height          =   735
      Left            =   10440
      TabIndex        =   35
      Top             =   5640
      Width           =   2295
   End
   Begin VB.TextBox txtGrossPay 
      Height          =   735
      Left            =   2040
      TabIndex        =   32
      Top             =   9120
      Width           =   2295
   End
   Begin VB.TextBox txtmeal 
      Height          =   735
      Left            =   2040
      TabIndex        =   30
      Top             =   8160
      Width           =   2295
   End
   Begin VB.TextBox txtperhour 
      Height          =   735
      Left            =   2040
      TabIndex        =   28
      Top             =   7320
      Width           =   2295
   End
   Begin VB.TextBox txtperday 
      Alignment       =   1  'Right Justify
      Height          =   735
      Left            =   2040
      TabIndex        =   26
      Top             =   6480
      Width           =   2295
   End
   Begin VB.TextBox txt15th 
      Height          =   735
      Left            =   2040
      TabIndex        =   24
      Top             =   5640
      Width           =   2295
   End
   Begin VB.TextBox txtpagibig 
      Height          =   735
      Left            =   10440
      TabIndex        =   21
      Top             =   3600
      Width           =   2295
   End
   Begin VB.TextBox txtphilhealth 
      Height          =   735
      Left            =   10440
      TabIndex        =   19
      Top             =   2760
      Width           =   2295
   End
   Begin VB.TextBox txttin 
      Height          =   735
      Left            =   10440
      TabIndex        =   17
      Top             =   1920
      Width           =   2295
   End
   Begin VB.TextBox txtSSS 
      Height          =   735
      Left            =   10440
      TabIndex        =   15
      Top             =   1080
      Width           =   2295
   End
   Begin VB.TextBox txtdatehired 
      Height          =   735
      Left            =   10440
      TabIndex        =   13
      Top             =   240
      Width           =   2295
   End
   Begin VB.TextBox txtdateTo 
      Height          =   735
      Left            =   5880
      TabIndex        =   11
      Top             =   3600
      Width           =   2295
   End
   Begin VB.TextBox txtdatefrom 
      Height          =   735
      Left            =   1920
      TabIndex        =   9
      Top             =   3600
      Width           =   2295
   End
   Begin VB.TextBox txtmonthlysalaryy 
      Height          =   645
      Left            =   1920
      TabIndex        =   7
      Top             =   2760
      Width           =   2295
   End
   Begin VB.TextBox txtname 
      Height          =   735
      Left            =   1920
      TabIndex        =   5
      Top             =   1920
      Width           =   2295
   End
   Begin VB.TextBox txtempID 
      Height          =   735
      Left            =   1920
      TabIndex        =   3
      Top             =   1080
      Width           =   2295
   End
   Begin VB.TextBox txttranno 
      Height          =   735
      Left            =   1920
      TabIndex        =   1
      Top             =   240
      Width           =   2295
   End
   Begin VB.Label Label22 
      Caption         =   "Pagibig"
      Height          =   615
      Left            =   8880
      TabIndex        =   40
      Top             =   8160
      Width           =   1215
   End
   Begin VB.Label Label21 
      Caption         =   "Philhealth"
      Height          =   615
      Left            =   8880
      TabIndex        =   38
      Top             =   7320
      Width           =   1215
   End
   Begin VB.Label Label20 
      Caption         =   "Tax With Held"
      Height          =   615
      Left            =   8880
      TabIndex        =   36
      Top             =   6480
      Width           =   1215
   End
   Begin VB.Label Label19 
      Caption         =   "SSS"
      Height          =   615
      Left            =   8880
      TabIndex        =   34
      Top             =   5640
      Width           =   1215
   End
   Begin VB.Label Label18 
      Caption         =   "List Of Deductions"
      Height          =   615
      Left            =   8880
      TabIndex        =   33
      Top             =   4800
      Width           =   3855
   End
   Begin VB.Label Label17 
      Caption         =   "Gross Pay"
      Height          =   615
      Left            =   480
      TabIndex        =   31
      Top             =   9120
      Width           =   1215
   End
   Begin VB.Label Label16 
      Caption         =   "Meal/Travel Allowance"
      Height          =   615
      Left            =   480
      TabIndex        =   29
      Top             =   8160
      Width           =   1215
   End
   Begin VB.Label Label15 
      Caption         =   "Rare Per Hour"
      Height          =   615
      Left            =   480
      TabIndex        =   27
      Top             =   7320
      Width           =   1215
   End
   Begin VB.Label Label14 
      Caption         =   "Rate Per Day"
      Height          =   615
      Left            =   480
      TabIndex        =   25
      Top             =   6480
      Width           =   1215
   End
   Begin VB.Label Label13 
      Caption         =   "Rate Per 15th Day Of The Month"
      Height          =   615
      Left            =   480
      TabIndex        =   23
      Top             =   5640
      Width           =   1215
   End
   Begin VB.Label Label12 
      Caption         =   "Breakdown Of Wages"
      Height          =   615
      Left            =   480
      TabIndex        =   22
      Top             =   4800
      Width           =   3855
   End
   Begin VB.Label Label11 
      Caption         =   "PAGIBIG#"
      Height          =   615
      Left            =   8880
      TabIndex        =   20
      Top             =   3600
      Width           =   1215
   End
   Begin VB.Label Label10 
      Caption         =   "PHILHEALTH#"
      Height          =   615
      Left            =   8880
      TabIndex        =   18
      Top             =   2760
      Width           =   1215
   End
   Begin VB.Label Label9 
      Caption         =   "TIN#"
      Height          =   615
      Left            =   8880
      TabIndex        =   16
      Top             =   1920
      Width           =   1215
   End
   Begin VB.Label Label8 
      Caption         =   "SSS#"
      Height          =   615
      Left            =   8880
      TabIndex        =   14
      Top             =   1080
      Width           =   1215
   End
   Begin VB.Label Label7 
      Caption         =   "Date Hired"
      Height          =   615
      Left            =   8880
      TabIndex        =   12
      Top             =   240
      Width           =   1215
   End
   Begin VB.Label Label6 
      Caption         =   "Date Covered To"
      Height          =   615
      Left            =   4320
      TabIndex        =   10
      Top             =   3600
      Width           =   1215
   End
   Begin VB.Label Label5 
      Caption         =   "Date Covered From"
      Height          =   615
      Left            =   360
      TabIndex        =   8
      Top             =   3600
      Width           =   1215
   End
   Begin VB.Label txtmonthlysalary 
      Caption         =   "Monthly Salary"
      Height          =   615
      Left            =   360
      TabIndex        =   6
      Top             =   2760
      Width           =   1215
   End
   Begin VB.Label Label3 
      Caption         =   "Complete Name"
      Height          =   615
      Left            =   360
      TabIndex        =   4
      Top             =   1920
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "Employee ID"
      Height          =   615
      Left            =   360
      TabIndex        =   2
      Top             =   1080
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Transaction No"
      Height          =   615
      Left            =   360
      TabIndex        =   0
      Top             =   240
      Width           =   1215
   End
End
Attribute VB_Name = "payrollform"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdadd_Click()
On Error Resume Next

txttranno.SelStart = 0
txttranno.SelLength = Len(txttranno.Text)
txttranno.SetFocus

txttranno.Text = ""
txtempID.Text = ""
txtname.Text = ""
txtmonthlysalaryy.Text = ""
txtdatefrom.Value = ""
txtdateTo.Value = ""
txtdatehired.Value = ""
txtSSS.Text = ""
txttin.Text = ""
txtphilhealth.Text = ""
txtpagibig.Text = ""
txt15th.Text = ""
txtperday.Text = ""
txtperhour.Text = ""
txtmeal.Text = ""
txtGrossPay.Text = ""
txtsss1.Text = ""
txttax.Text = ""
txtphil.Text = ""
txtpag.Text = ""
txttotdeduction.Text = ""
txtnetincome.Text = ""

End Sub

Private Sub cmdClose_Click()
Unload Me

End Sub

Private Sub cmddeduct_Click()
Dim xsss As Single
Dim xtax As Single
Dim xpag As Single
Dim xphil As Single
Dim xtOTD As Double
xsss = 300
txtsss1.Text = xsss

xtax = 500
txttax.Text = xtax

xpag = 100
txtpag.Text = xpag

xphil = 100
txtphil.Text = xphil

xtOTD = xsss + xtax + xphil + xpag

txttotdeduction.Text = xtOTD


End Sub

Private Sub cmdDelete_Click()
conPayroll.Execute "Delete * from payroll where tranno='" & Trim(txttranno.Text) & "'"
MsgBox "Record has been deleted.."
End Sub

Private Sub cmdFind_Click()
txttranno.SelStart = 0
txttranno.SelLength = Len(txttranno.Text)
txttranno.SetFocus

End Sub

Private Sub cmdGross_Click()
Dim xrate15 As Double
Dim xSalary As Double
Dim xrateperday As Double
Dim xrateperhour As Double
Dim xmeal As Double
Dim xGross As Double



xSalary = txtmonthlysalaryy.Text
xrate15 = xSalary / 2
txt15th.Text = xrate15

xrateperday = txtmonthlysalaryy.Text / 26
txtperday.Text = xrateperday

xrateperhour = txtperday.Text / 8
txtperhour.Text = xrateperhour

xmeal = 500

txtmeal.Text = xmeal

xGross = xmeal + xrate15

txtGrossPay.Text = xGross







End Sub

Private Sub cmdnet_Click()
Dim xNet As Double
Dim xG As Double
Dim xD As Double


xG = txtGrossPay.Text
xD = txttotdeduction.Text

xNet = xG - xD
txtnetincome.Text = xNet

End Sub

Private Sub cmdsave_Click()
OPENRSTPAYROLL "SELECT * FROM payroll where tranno='" & Trim(txttranno.Text) & "'"
If Not rstPayroll.EOF Then
'if not found
    With rstPayroll
        .Edit
            .Fields("tranno").Value = Trim(txttranno.Text)
            .Fields("employeeid").Value = Trim(txtempID.Text)
            .Fields("datefrom").Value = Trim(txtdatefrom.Value)
            .Fields("dateto").Value = Trim(txtdateTo.Value)
            .Fields("rate15").Value = Trim(txt15th.Text)
            .Fields("rateperday").Value = Trim(txtperday.Text)
            .Fields("rateperhour").Value = Trim(txtperhour.Text)
            .Fields("meal").Value = Trim(txtmeal.Text)
            .Fields("grosspay").Value = Trim(txtGrossPay.Text)
            .Fields("datehired").Value = Trim(txtdatehired.Value)
            .Fields("sssno").Value = Trim(txtSSS.Text)
            .Fields("tinno").Value = Trim(txttin.Text)
            .Fields("philhealthno").Value = Trim(txtphilhealth.Text)
            .Fields("pagibigno").Value = Trim(txtpagibig.Text)
            .Fields("sss").Value = Trim(txtsss1.Text)
            .Fields("tax").Value = Trim(txttax.Text)
            .Fields("pagibig").Value = Trim(txtpag.Text)
            .Fields("philhealth").Value = Trim(txtphil.Text)
            .Fields("totaldeduction").Value = Trim(txttotdeduction.Text)
            .Fields("netincome").Value = Trim(txtnetincome.Text)
            
        .Update
        
    End With
Else
    'not found
        With rstPayroll
            .AddNew
                .Fields("tranno").Value = Trim(txttranno.Text)
                .Fields("employeeid").Value = Trim(txtempID.Text)
                .Fields("datefrom").Value = Trim(txtdatefrom.Value)
                .Fields("dateto").Value = Trim(txtdateTo.Value)
                .Fields("rate15").Value = Trim(txt15th.Text)
                .Fields("rateperday").Value = Trim(txtperday.Text)
                .Fields("rateperhour").Value = Trim(txtperhour.Text)
                .Fields("meal").Value = Trim(txtmeal.Text)
                .Fields("grosspay").Value = Trim(txtGrossPay.Text)
                .Fields("datehired").Value = Trim(txtdatehired.Value)
                .Fields("sssno").Value = Trim(txtSSS.Text)
                .Fields("tinno").Value = Trim(txttin.Text)
                .Fields("philhealthno").Value = Trim(txtphilhealth.Text)
                .Fields("pagibigno").Value = Trim(txtpagibig.Text)
                .Fields("sss").Value = Trim(txtsss1.Text)
                .Fields("tax").Value = Trim(txttax.Text)
                .Fields("pagibig").Value = Trim(txtpagibig.Text)
                .Fields("philhealth").Value = Trim(txtphil.Text)
                .Fields("totaldeduction").Value = Trim(txttotdeduction.Text)
                .Fields("netincome").Value = Trim(txtnetincome.Text)
            .Update
            
        End With
End If

    
End Sub


Private Sub Form_Load()
openWORKSPACEODBC
openconPayroll

End Sub

Private Sub txtempID_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    openrstEmployee "SELECT * FROM EMPLOYEE WHERE EMPLOYEEID='" & Trim(txtempID.Text) & "'"
    If Not rstEmployee.EOF Then
        With rstEmployee
            txtempID.Text = .Fields("employeeid").Value
            txtname.Text = .Fields("empname").Value
            txtmonthlysalaryy.Text = .Fields("salary").Value
            txtdatehired.Value = .Fields("datehired").Value
        End With
    End If
End If

End Sub

Private Sub txttranno_KeyPress(KeyAscii As Integer)
On Error Resume Next
If KeyAscii = 13 Then
    OPENRSTPAYROLL "Select * from payroll where tranno='" & Trim(txttranno.Text) & "'"
     If Not rstPayroll.EOF Then
        With rstPayroll
            txttranno.Text = .Fields("tranno").Value
            txtempID.Text = .Fields("employeeid").Value
            txtdatefrom.Value = .Fields("datefrom").Value
            txtdateTo.Value = .Fields("dateto").Value
            txt15th.Text = .Fields("rate15").Value
            txtperday.Text = .Fields("rateperday").Value
            txtperhour.Text = .Fields("rateperhour").Value
            txtmeal.Text = .Fields("meal").Value
            txtGrossPay.Text = .Fields("grosspay").Value
            txtdatehired.Value = .Fields("datehired").Value
            txtSSS.Text = .Fields("sssno").Value
            txttin.Text = .Fields("tinno").Value
            txtphilhealth.Text = .Fields("philhealthno").Value
            txtpagibig.Text = .Fields("pagibigno").Value
            txtsss1.Text = .Fields("sss").Value
            txttax.Text = .Fields("tax").Value
            txtpag.Text = .Fields("pagibig").Value
            txtphil.Text = .Fields("philhealth").Value
            txttotdeduction.Text = .Fields("totaldeduction").Value
            txtnetincome.Text = .Fields("netincome").Value
            
        End With
    End If
    
End If
End Sub
