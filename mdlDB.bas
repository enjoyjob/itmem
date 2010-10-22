Attribute VB_Name = "mdlDB"
Option Explicit
Public OraDB As New ADODB.Connection
Public OraDB_Open As Boolean
Public OraConstr As String

'
''VB连接ORACLE数据库  '打开数据库
Public Sub OpenOraDB()
    On Error GoTo ToExit
    OraDB_Open = False
    'Set OraDB = New ADODB.Connection
    'OraConstr = "Provider=OraOLEDB.Oracle.1;Password=" & strOraPWD & ";User ID=" & strOraUser & ";Data Source=" & OraDBNetName & ";Persist Security Info=False"
    OraConstr = "Provider=OraOLEDB.Oracle.1;Password=prones;User ID=prones-test;Data Source=prones2;Persist Security Info=False"
    OraDB.CursorLocation = adUseServer

    OraDB.Open OraConstr
    OraDB_Open = True

    Exit Sub
ToExit:
    MsgBox "连接数据库服务器错误，您可以在网络正常后继续使用。", vbInformation, "错误信息"
    OraDB_Open = False
End Sub

'关闭数据库
Public Sub CloseOraDB()
    If OraDB_Open = True Then
        If (OraDB.State = adStateOpen) Then
            OraDB.Close
            Set OraDB = Nothing
            OraDB_Open = False
        End If
    End If
    OraDB_Open = False
End Sub




