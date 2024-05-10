Attribute VB_Name = "Module1"
Public WorkspaceODBC As Workspace
Public conPayroll As Connection
Public rstEmployee As Recordset
Public rstPayroll As Recordset

Public SetupReport
'DAO WORKSPACE FUNCTION
Public Sub openWORKSPACEODBC()
    Set WorkspaceODBC = CreateWorkspace("ODBCWorkpace", "", "Admin", dbUseODBC)
End Sub
'DAO CONNECTION FUNCTION
Public Sub openconPayroll()
    Set conPayroll = WorkspaceODBC.OpenConnection("", dbDriverNoPrompt, False, "ODBC;Database=Payroll2d;UID=sa;PWD=pentium;DSN=Payroll2d")
End Sub
'--- DAO RECORDSET COSTTRANHEADER FUNCTION
Public Sub openrstEmployee(SelectString As String)
    Set rstEmployee = conPayroll.OpenRecordset(SelectString, dbOpenDynamic, 0, dbOptimistic)
End Sub
