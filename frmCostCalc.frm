VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmCostCalc 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "成本计算"
   ClientHeight    =   3525
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5115
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3525
   ScaleWidth      =   5115
   StartUpPosition =   3  '窗口缺省
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   2880
      Top             =   720
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton btnOK 
      Caption         =   "确 认"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   360
      TabIndex        =   3
      Top             =   2040
      Width           =   1575
   End
   Begin VB.TextBox txtTo 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1200
      MaxLength       =   8
      TabIndex        =   2
      Text            =   "20101112"
      Top             =   1320
      Width           =   1215
   End
   Begin VB.TextBox txtFrom 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1200
      MaxLength       =   8
      TabIndex        =   1
      Text            =   "20101020"
      Top             =   720
      Width           =   1215
   End
   Begin VB.TextBox txtNo 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1200
      MaxLength       =   20
      TabIndex        =   0
      Text            =   "1234567890ABCDEFGHIJ"
      Top             =   120
      Width           =   2655
   End
   Begin VB.CommandButton btnClear 
      Caption         =   "清 除"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   2400
      TabIndex        =   4
      Top             =   2040
      Width           =   1575
   End
   Begin VB.Label lblCnt 
      Caption         =   "20"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3960
      TabIndex        =   8
      Top             =   240
      Width           =   375
   End
   Begin VB.Label Label3 
      Caption         =   "终止日期"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   7
      Top             =   1440
      Width           =   855
   End
   Begin VB.Label Label2 
      Caption         =   "起始日期"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   6
      Top             =   840
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "订单号"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   5
      Top             =   240
      Width           =   855
   End
End
Attribute VB_Name = "frmCostCalc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub ClearAll()
    lblCnt.Caption = ""
    txtNo.Text = ""
    txtFrom = ""
    txtTo = ""
End Sub

Private Sub btnClear_Click()
    ClearAll
End Sub

Private Function chkDate(tfrom As TextBox, tTo As TextBox) As Boolean
    chkDate = False
    
    'check 起始日期
    With tfrom
        If Not Ymd_chek_proc(.Text) Then
            MsgBox C_MSG_001, vbCritical, C_MSG_TITLE_ERR
            .SetFocus
            Exit Function
        End If
    End With

    'check 终止日期
    With tTo
        If Not Ymd_chek_proc(.Text) Then
            MsgBox C_MSG_001, vbCritical, C_MSG_TITLE_ERR
            .SetFocus
            Exit Function
        End If
    End With
    
    '起始日期 < = 终止日期
        If CDate(Format(tfrom.Text, "####/##/##")) > CDate(Format(tTo.Text, "####/##/##")) Then
            MsgBox C_MSG_002, vbCritical, C_MSG_TITLE_ERR
            tTo.SetFocus
            Exit Function
        End If


    chkDate = True
End Function
Private Sub btnOK_Click()
    'check 日期
    If Not chkDate(txtFrom, txtTo) Then
        Exit Sub
    End If
    
    'check OK
    Screen.MousePointer = vbHourglass
    Me.Enabled = False
    
    Call doExport(txtNo.Text, txtFrom.Text, txtTo.Text)
    
    Me.Enabled = True
    Screen.MousePointer = vbDefault
    Me.SetFocus
End Sub


Private Sub Form_Load()
    ClearAll
End Sub

Private Sub Form_Unload(Cancel As Integer)
    CleanExcelApp
    CloseOraDB
End Sub

Private Sub txtNo_Change()
    If Len(txtNo.Text) = 0 Then
        lblCnt.Caption = ""
    Else
        lblCnt.Caption = Len(txtNo.Text)
    End If
End Sub


Private Sub doExport(sNo As String, sFrom As String, sTo As String)
On Error GoTo ERR_GO
    Dim bRtn As Boolean
    Dim sErr As String, sText As String
    Dim iLine As Long, jCol As Long
    
    Dim dataAry() As Variant
   
    ''''''''''''''''''''''''''''''''''''''''
    ''test for write to excel file
    Dim cnt As Long
    
    If CreateExcelFileFromTemplate(SetExcelFile(Me.CommonDialog1), App.Path & "\" & "moban.xls", sErr) = False Then
        If Trim(sErr) <> "" Then
            MsgBox "系统错误：" & sErr
        End If
        
        Exit Sub
    End If
    
   
    'write to range ,speed up than write to cell
    iLine = 2
    jCol = 1
    
    ReDim dataAry(0 To 5)
    dataAry(0) = 0
    dataAry(1) = "22"
    dataAry(2) = "hello"
    dataAry(3) = Date
    dataAry(4) = Null
    dataAry(5) = Empty
    
    For cnt = 0 To 9
        If cnt > 0 Then
            Call Excel_InsertRow(gxlSheet1, iLine + cnt)  'add row
        End If
        dataAry(0) = cnt
        bRtn = WriteExcelCellEx(gxlSheet1, iLine + cnt, jCol, dataAry)
    Next
    Call Excel_DeleteRow(gxlSheet1, iLine + cnt) 'delete row

    gxlBook.Save    'save  excel file
    CleanExcelApp    'close excel app
    
    MsgBox "write excel OK"
    Exit Sub
   ''test for write to excel file
   ''''''''''''''''''''''''''''''''''''''''
   
    Dim quantity As String
    Dim w_year_s As String
    Dim openff As Boolean
    Dim strOraRS As String
    Dim OraRS As ADODB.Recordset
    
    
    'connect db
    OpenOraDB
    If OraDB_Open = False Then
        Exit Sub
    End If
    
    '打开数据集，写入数据
    
          Set OraRS = New ADODB.Recordset
          OraRS.ActiveConnection = OraDB
          OraRS.CursorLocation = adUseServer
          OraRS.LockType = adLockBatchOptimistic
         ' strOraRS = "select   *   from   " & OraDBtablename
          strOraRS = "select count(*) AS count from T_CLS_MS "
          OraRS.Open strOraRS, OraDB, adOpenStatic, adLockBatchOptimistic
'          OraRS.AddNew
'          OraRS.Fields("PID") = strOraPID
'          OraRS.Fields("pname")   =   strName").Value
'          OraRS.Fields("psex") = strPsex
'          OraRS.Update
    
    If Not OraRS.EOF Then
        quantity = OraRS("count")
    End If
      
          
    '关闭数据集
    OraRS.Close
    Set OraRS = Nothing
    
    'disconnect DB
    CloseOraDB
    Exit Sub
    
ERR_GO:
    MsgBox "系统错误：" & Err.Description
    CleanExcelApp
    CloseOraDB
End Sub
