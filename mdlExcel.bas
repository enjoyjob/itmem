Attribute VB_Name = "mdlExcel"
Option Explicit


'excel file name
Public gReportFile As String '�����ļ���

Public gxlApp As Excel.Application
Public gxlBook As Excel.Workbook
Public gxlSheet1 As Excel.Worksheet
Public gxlRange As Excel.Range


Public Const SHEET_1 = 1

Private Const MSG_REPORT_MOBAN_FILE As String = "���ܸ���ģ���ļ�"
Private Const MSG_NO_MOBAN_FILE As String = "ģ���ļ�������"
Private Const MSG_REPORT_FILE As String = "�������"

'*************************************************************
' �趨һ��EXCEL�ļ���
'*************************************************************
Public Function SetExcelFile(cmmnDlg As CommonDialog) As String

On Error GoTo MyCancelError
With cmmnDlg
    .DialogTitle = MSG_REPORT_FILE
    .FileName = "Report" & Format(Now, "yyyymmdd") & ".xls"
    .Filter = "(Excel)*.xls|*.xls"
    .CancelError = True    '��ȡ���趨Ϊһ������,��������в�׽.
    .Flags = cdlOFNOverwritePrompt '���ļ�������ʱ,������������,���ǽ��и��ǵ�ѡ��.
    '.InitDir = "c:\"
    .ShowSave
    SetExcelFile = .FileName
End With

Exit Function

MyCancelError:
    SetExcelFile = ""
    If Err.Number = 32755 Then
        'cancel open
    End If
End Function


'*************************************************************
' cleanup EXCEL from memory
'*************************************************************
Public Sub CleanExcelApp()
On Error Resume Next
  'gxlApp.Visible = True
  If Not (gxlBook Is Nothing) Then
    gxlBook.Close (False)
  End If
  If Not (gxlApp Is Nothing) Then
    gxlApp.Workbooks.Close
    gxlApp.Quit
  End If

  Set gxlRange = Nothing
  Set gxlSheet1 = Nothing
  
  Set gxlBook = Nothing
  Set gxlApp = Nothing
End Sub


'*************************************************************
' write to a EXCEL cell or cells
'*************************************************************
Public Function WriteExcelCellEx(oxlSheet As Excel.Worksheet, i As Long, j As Long, vData As Variant) As Boolean
On Error GoTo ERR_WRITE_CELL
    Dim dataAry() As Variant
    Dim ArySize As Long
    
    WriteExcelCellEx = True
    If IsArray(vData) Then
        ArySize = UBound(vData) - LBound(vData) + 1
        If ArySize > 0 Then
            dataAry = vData
            With oxlSheet
                .Range(.Cells(i, j), .Cells(i, j + ArySize - 1)).Value = dataAry
            End With
        End If
    
    Else
        If ((vData & "") = "") Then ''IsNull(value) or IsEmpty(value) or (value = "")
            oxlSheet.Cells(i, j).Value = ""
        Else
            oxlSheet.Cells(i, j).Value = vData
        End If
    
    End If
    Exit Function
ERR_WRITE_CELL:
    WriteExcelCellEx = False
End Function

'*************************************************************
' �½�EXCEL�ļ�
'*************************************************************
Public Function CreateExcelFile(sDest As String, sErrMsg As String) As Boolean
On Error GoTo ERR_OPEN_FILE
        
    CreateExcelFile = False
    sErrMsg = ""
    
    gReportFile = Trim(sDest)

    If gReportFile = "" Then
        sErrMsg = ""
        Exit Function
    End If
    
    If Trim(Dir(gReportFile)) <> "" Then
        Kill gReportFile
    End If
    
    Set gxlApp = CreateObject("Excel.Application")
    gxlApp.Visible = False

    Set gxlBook = gxlApp.Workbooks.Add()
    
    gxlBook.Application.DisplayAlerts = False
    gxlBook.SaveAs (gReportFile)
    
    Set gxlSheet1 = Nothing
    Set gxlSheet1 = gxlBook.Worksheets(SHEET_1)
    
    CreateExcelFile = True
    Exit Function
ERR_OPEN_FILE:
    sErrMsg = Err.Description
    CreateExcelFile = False
    CleanExcelApp
End Function

'*************************************************************
' ��ģ���ļ��½�EXCEL�ļ�
'*************************************************************
Public Function CreateExcelFileFromTemplate(sDest As String, sTemplate As String, sErrMsg As String) As Boolean
On Error GoTo ERR_OPEN_FILE
    Dim sTmp As String
    
    CreateExcelFileFromTemplate = False
    sErrMsg = ""
    
    gReportFile = Trim(sDest)

    If gReportFile = "" Then
        sErrMsg = ""
        Exit Function
    End If
    
    '������ģ���ļ����汨��
    If UCase(gReportFile) = UCase(sTemplate) Then
        sErrMsg = ""
        MsgBox MSG_REPORT_MOBAN_FILE, vbOKOnly + vbExclamation, MSG_REPORT_FILE
        Exit Function
    End If

    'ģ���ļ�������
    If Trim(Dir(sTemplate)) = "" Then
        sErrMsg = ""
        sTmp = MSG_NO_MOBAN_FILE & vbCrLf & sTemplate
        MsgBox sTmp, vbOKOnly + vbExclamation, MSG_REPORT_FILE
        Exit Function
    End If

    'Ŀ���ļ����ڣ����֮
    If Trim(Dir(gReportFile)) <> "" Then
        SetAttr gReportFile, vbNormal
        Kill gReportFile
    End If
    
    '����ģ���ļ���Ŀ���ļ�
    FileCopy sTemplate, gReportFile
    SetAttr gReportFile, vbNormal

    Set gxlApp = CreateObject("Excel.Application")
    gxlApp.Visible = False

    Set gxlBook = gxlApp.Workbooks.Open(gReportFile, 0) '0-�������κ����� / 3-��������Զ�����ú��ⲿ����
    gxlBook.Application.DisplayAlerts = False
    
    Set gxlSheet1 = Nothing
    Set gxlSheet1 = gxlBook.Worksheets(SHEET_1)
    
    CreateExcelFileFromTemplate = True
    Exit Function
ERR_OPEN_FILE:
    sErrMsg = Err.Description
    CreateExcelFileFromTemplate = False
    CleanExcelApp
End Function

Public Sub Excel_InsertRow(ByVal oxlSheet As Excel.Worksheet, i As Long, Optional ByVal Direction As Excel.XlDirection = Excel.XlDirection.xlDown)
On Error GoTo MY_ERR
    Dim sErr As String
    With oxlSheet
        .Range(.Cells(i, 1), .Cells(i, 1)).EntireRow.Insert
    End With
    
    Exit Sub
MY_ERR:
    sErr = Err.Description
End Sub

Public Sub Excel_DeleteRow(oxlSheet As Excel.Worksheet, i As Long, Optional ByVal Direction As Excel.XlDirection = Excel.XlDirection.xlUp)
On Error GoTo MY_ERR
    Dim sErr As String
    
    With oxlSheet
        .Range(.Cells(i, 1), .Cells(i, 1)).EntireRow.Delete
    End With
    
    Exit Sub
MY_ERR:
    sErr = Err.Description
End Sub
