VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsExcel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Private WithEvents ThisWorkbook As Workbook
Attribute ThisWorkbook.VB_VarHelpID = -1
Private mExcelApp As Excel.Application
Private mwsParameters As Worksheet
Private mbBlocking As Boolean
Private mstrWBFilename As String
Private mstrProcedureName As String
Private mbNew As Boolean
Private mbIdle As Boolean
Private mstrKey As String
Private mclsParent As clsExcels

Public Event Finsihed(strKey As String)

Private mintTimeOuts As Integer
Private mlLogLastRow As Long

Public Property Set wsLog(ws As Worksheet)
    Set mwsLog = ws
End Property
Public Property Set wsScheduler(ws As Worksheet)
    Set mwsScheduler = ws
End Property


Public Property Set Parent(Excels As clsExcels)
    Set mclsParent = Excels
End Property

Public Property Get ExeclApp() As Excel.Application
    Set ExeclApp = mExcelApp
End Property
Public Property Get IsIdle() As Boolean
    IsIdle = mbIdle
End Property
Public Property Let IsIdle(bIdle As Boolean)
    mbIdle = bIdle
End Property
Public Property Let WorkbookFile(strFile As String)
    mstrWBFilename = strFile
End Property
Public Property Let Key(strKey As String)
    mstrKey = strKey
End Property
Public Property Get ProcedureName() As String
    ProcedureName = mbBlocking
End Property
Public Property Let NewInstance(bNew As Boolean)
    mbNew = bNew
End Property

Public Sub CreateOpen()
    Set mExcelApp = New Excel.Application
    Set ThisWorkbook = mExcelApp.Workbooks.Open(mstrWBFilename, 0)
    Set mwsParameters = ThisWorkbook.Sheets("Parameters")
    mbIdle = True
End Sub
Public Sub SetupParameters(strProcedureName As String, strERP As String, strCommon As String, _
    bSendMail As Boolean, bTrial As Boolean, bPrintDestinations As Boolean, bDPL As Boolean, _
    Optional bBlock As Boolean, Optional strWorkingFileName As String, Optional bReadwrite As Boolean)
    mwsParameters.Cells(3, 1) = strProcedureName ' called procedure number
    mwsParameters.Cells(2, 1) = 1 ' set busy
    mwsParameters.Cells(1, 1) = strERP
    mwsParameters.Cells(1, 2) = strCommon
    mwsParameters.Cells(1, 3) = strWorkingFileName
    mwsParameters.Cells(2, 3) = bReadwrite
    mwsParameters.Cells(1, 4) = bSendMail
    mwsParameters.Cells(1, 5) = bTrial
    mwsParameters.Cells(1, 6) = bPrintDestinations
'    mwsParameters.Cells(1, 7) = frmDashboard.chkDPL
'    mExcelApp.EnableEvents = False
'    ThisWorkbook.Save
'    mExcelApp.EnableEvents = True
'    'report status and schedule if necessary
    mwsParameters.Cells(5, 1) = ""
    mwsParameters.Cells(5, 2) = ""
    With mwsParameters
        If mlLogLastRow = 0 Then
            mlLogLastRow = .UsedRange.Rows.Count + 1
        End If
        .Cells(mlLogLastRow, 1) = mwsParameters.Cells(3, 1)
        .Cells(mlLogLastRow, 2) = Now()
        .Cells(mlLogLastRow, 5) = strWorkingFileName
        .Cells(mlLogLastRow, 6) = "Processing"
    End With

End Sub
Public Sub SetupMail(ByVal strFrom As String, ByVal sPass As String, ByVal sToMail As String, ByVal sCC As String, _
    ByVal sBCC As String, ByVal strMsg As String, ByVal strSubject As String, Optional HTML As Variant = Empty, Optional Attachments As Variant = Empty)

    Dim n As Integer, m As Integer
    
    On Error Resume Next
    
    'put parameters in new created sendmail wb
    If IsArray(Attachments) = True Then
        ' attach all the files in the array.
        mwsParameters.Cells(1, 11) = UBound(Attachments) - LBound(Attachments) + 1
        For n = LBound(Attachments) To UBound(Attachments)
            ' ensure the attachment file exists and attach it.
            If Attachments(n) <> vbNullString Then
                If Dir(Attachments(n), vbNormal) <> vbNullString Then
                    mwsParameters.Cells(1, 13 + n) = Attachments(n)
                End If
            End If
        Next n
    Else
        ' ensure the file exists and if so, attach it to the message.
        If Attachments <> vbNullString Then
            If Dir(CStr(Attachments), vbNormal) <> vbNullString Then
                mwsParameters.Cells(1, 11) = Attachments
            End If
        End If
    End If
    'put parameters in new created sendmail wb
    If IsArray(HTML) = True Then
        mwsParameters.Cells(1, 12) = UBound(HTML) - LBound(HTML) + 1
        For m = LBound(HTML) To UBound(HTML)
            mwsParameters.Cells(1, 12 + n + m) = HTML(m)
        Next m
    Else
        ' ensure the file exists and if so, attach it to the message.
        mwsParameters.Cells(1, 12) = HTML
    End If

    mwsParameters.Cells(3, 1) = "SendMail" ' called procedure number
    mwsParameters.Cells(2, 1) = 1 ' set busy
    mwsParameters.Cells(1, 1) = strFrom
    mwsParameters.Cells(1, 2) = sPass
    mwsParameters.Cells(1, 3) = sToMail
    mwsParameters.Cells(1, 4) = sCC
    mwsParameters.Cells(1, 5) = sBCC
    mwsParameters.Cells(1, 6) = strMsg
    mwsParameters.Cells(1, 7) = strSubject
    mwsParameters.Cells(1, 8) = strFrom
'    mwsParameters.Cells(1, 9) = strHTML(1)
'    mwsParameters.Cells(1, 10) = strHTML2
'    mExcelApp.EnableEvents = False
'    ThisWorkbook.Save
'    mExcelApp.EnableEvents = True

End Sub

Public Function RangeToHTML(ByRef rRng As Range) As String
    
    Dim rRow As Range, rCell As Range
    Dim sTable As String, sTd As String, sHead As String
    Dim aCells() As String, aRows() As String, aAttr() As String, aHead(1 To 2) As String
    Dim lCellCnt As Long, lRowCnt As Long
    Dim lFontSize As Long
    
    '1. Get the font size of the last cell
    lFontSize = rRng.Cells(rRng.Cells.Count).Font.Size
    ReDim aRows(1 To rRng.Rows.Count)
    
    '2 create the style in the header
    aHead(1) = "td {font-family:" & rRng.Cells(1).Font.Name & "; font-size: " & lFontSize & "pt}"
    aHead(2) = ".bb {border-bottom: 1px solid black}"
    sHead = Tag(Tag(Join(aHead, vbNewLine), "style", , True), "head", , True)
    
    '3. Load up a 'cells' array and a 'rows' array FOR joining.
    For Each rRow In rRng.Rows
        lRowCnt = lRowCnt + 1: lCellCnt = 0
        ReDim aCells(1 To rRng.Columns.Count)
        For Each rCell In rRow.Cells
            lCellCnt = lCellCnt + 1
            If Not rCell.EntireColumn.Hidden Then
            
                '4. Deal with empty cells and multi-line cells
                If IsEmpty(rCell.Value) Then
                    sTd = "&nbsp;"
                Else
                    sTd = Replace(rCell.Text, Chr$(10), "<br />")
                End If
                
                '5. Bold and italic
                If rCell.Font.Bold Then sTd = Tag(sTd, "strong")
                If rCell.Font.Italic Then sTd = Tag(sTd, "em")
                
                '6. Font size
                If rCell.Font.Size <> lFontSize Then
                    sTd = Tag(sTd, "div", "style=font-size:" & rCell.Font.Size & "pt")
                End If
                
                '7. Setting the cell alignment
                ReDim aAttr(1 To 4)
                aAttr(1) = AlignmentAttr(rCell)
                
                '8. Setting the cell color
                aAttr(2) = ColorAttr(rCell)
                
                '8. Span rows and columns for merged  cells
                If rCell.MergeArea.Address <> rCell.Address Then
                    aAttr(3) = "COLSPAN=""" & rCell.MergeArea.Columns.Count & """ ROWSPAN=""" & rCell.MergeArea.Rows.Count & """"
                End If
                
                '9. Bottom border
                If rCell.Borders(xlEdgeBottom).LineStyle <> xlLineStyleNone Then
                    aAttr(4) = "class=""bb"""
                End If
                
                '10. Make string
                If rCell.MergeArea.Cells(1).Address = rCell.Address Then
                    aCells(lCellCnt) = Tag(sTd, "td", Join(aAttr, Space(1)))
                End If
            End If
        Next rCell
        aRows(lRowCnt) = Tag(Join(aCells, vbNewLine), "tr", , True)
    Next rRow
    
    sTable = Tag(Join(aRows, vbNewLine), "table", "cellpadding=""2px""", True)
    
    RangeToHTML = Tag(sHead & vbNewLine & sTable, "html", , True)
    
End Function
Function Tag(sValue As String, sTag As String, Optional sAttr As String = "", Optional bIndent As Boolean = False) As String
   
    Dim sReturn As String
   
    If Len(sAttr) > 0 Then
        sAttr = Space(1) & sAttr
    End If
   
    If bIndent Then
        sValue = vbTab & Replace(sValue, vbNewLine, vbNewLine & vbTab)
        sReturn = "<" & sTag & sAttr & ">" & vbNewLine & sValue & vbNewLine & "</" & sTag & ">"
    Else
        sReturn = "<" & sTag & sAttr & ">" & sValue & "</" & sTag & ">"
    End If
   
    Tag = sReturn
   
End Function

Public Function AlignmentAttr(ByRef rCell As Range) As String
    
    Dim sReturn As String
    
    Select Case True
        Case rCell.HorizontalAlignment = xlLeft, (rCell.HorizontalAlignment = 1 And Not IsNumeric(rCell.Value))
            sReturn = "align=""left"""
        Case rCell.HorizontalAlignment = xlRight, (rCell.HorizontalAlignment = 1 And IsNumeric(rCell.Value))
            sReturn = "align=""right"""
        Case rCell.HorizontalAlignment = xlCenter
            sReturn = "align=""center"""
    End Select
    
    AlignmentAttr = sReturn
    
End Function

Public Function ColorAttr(ByRef rCell As Range) As String
    
    Dim sReturn As String
    Dim strBColor As String
    Dim strColor As String
    
    On Error Resume Next
    
    strBColor = Color(rCell, 2)
    strColor = Color(rCell, 1)
    If strBColor = "" Then
        If strColor = "" Then
            sReturn = ""
        Else
            sReturn = "style=" & """color: #" & strColor & ";"""
        End If
    Else
        If strColor = "" Then
            sReturn = "style=" & """background-color: #" & strBColor & ";"""
        Else
            sReturn = "style=" & """background-color: #" & strBColor & "; color: #" & Color(rCell, 1) & ";"""
        End If
    End If
    
    ColorAttr = sReturn
    
End Function
Function Color(rCell As Range, Optional FormatType As Integer = 0) As Variant
'   Purpose             Determine the Background Color Of a Cell
'   @Param rng          Range to Determine Background Color of
'   @Param formatType   Default Value = 0
'                       0   Integer
'                       1   Hex
'                       2   RGB
'                       3   Excel Color Index
'   Usage               Color(A1)      -->   9507341
'                       Color(A1, 0)   -->   9507341
'                       Color(A1, 1)   -->   91120D
'                       Color(A1, 2)   -->   13, 18, 145
'                       Color(A1, 3)   -->   6
    Dim colorVal As Variant
    Select Case FormatType
        Case 1
            colorVal = rCell.Font.Color
            Color = FmtHex(colorVal Mod 256, 2) & FmtHex((colorVal \ 256) Mod 256, 2) & FmtHex(colorVal \ 65536, 2)
        Case 2
            colorVal = rCell.Interior.Color
            Color = FmtHex(colorVal Mod 256, 2) & FmtHex((colorVal \ 256) Mod 256, 2) & FmtHex(colorVal \ 65536, 2)
        Case 3
            colorVal = rCell.Font.Background
            Color = FmtHex(colorVal Mod 256, 2) & FmtHex((colorVal \ 256) Mod 256, 2) & FmtHex(colorVal \ 65536, 2)
        Case Else
            colorVal = rCell.Interior.Color
            Color = FmtHex(colorVal Mod 256, 2) & FmtHex((colorVal \ 256) Mod 256, 2) & FmtHex(colorVal \ 65536, 2)
    End Select
    If Color = "FFFFFF" Then Color = ""
    If Color = "000000" Then Color = ""
End Function


Public Function Launch() As Boolean
    On Error Resume Next
    
'    DoEvents
    
    mExcelApp.WindowState = xlMinimized
    mExcelApp.Visible = False
'        mExcelApp.Visible = True
    Launch = True
    mbIdle = False
    mbNew = False
'    ThisWorkbook.Activate
    mExcelApp.Run "Launcher"

End Function

Private Sub ThisWorkbook_AfterSave(ByVal Success As Boolean)
    
    
    On Error GoTo ErrorHandler
    mwsParameters.Cells(6, 1) = "Mail Sent"
    
    DoEvents
'    RaiseEvent Finsihed(mstrKey)

    ThisWorkbook.Close False
'    DoEvents
    mExcelApp.Quit
'    DoEvents
    Kill mstrWBFilename
    mclsParent.Finsihed mstrKey
    Exit Sub
ErrorHandler:
    mwsParameters.Cells(5, 1) = Err.Description
    Resume Next
End Sub
Public Function Terminate() As Boolean
            Terminate = True

End Function

Private Sub ThisWorkbook_NewSheet(ByVal Sh As Object)
    On Error Resume Next
    
    Select Case mwsParameters.Cells(5, 1)
            Case -2147220977  'Likely cause, Incorrectly Formatted Email Address, server rejected the Email Format
                Sh.Cells(1, 1) = 2
                With Sh
                    mlLogLastRow = .UsedRange.Rows.Count + 1
                    .Cells(mlLogLastRow, 1) = mwsParameters.Cells(3, 1)
                    .Cells(mlLogLastRow, 3) = Now()
                    .Cells(mlLogLastRow, 4) = "Failed"
                End With

            Case -2147220980  'Likely cause, No Recipient Provided (No Email Address)
                Sh.Cells(1, 1) = 2
                With Sh
                    mlLogLastRow = .UsedRange.Rows.Count + 1
                    .Cells(mlLogLastRow, 1) = mwsParameters.Cells(3, 1)
                    .Cells(mlLogLastRow, 3) = Now()
                    .Cells(mlLogLastRow, 4) = "Failed"
                End With

            Case -2147220960 'Likely cause, SendUsing Configuration Error
                Sh.Cells(1, 1) = 2
            
            Case -2147220973, -2147220974  'Likely cause, No Internet Connection
                If mintTimeOuts < 10 Then
                    Sh.Cells(1, 1) = 1
                    Sh.Cells(1, 2) = 60
                    mintTimeOuts = mintTimeOuts + 1
                    With Sh
                        If mlLogLastRow = 0 Then
                            mlLogLastRow = .UsedRange.Rows.Count + 1
                        End If
                        .Cells(mlLogLastRow, 1) = mwsParameters.Cells(3, 1)
                        .Cells(mlLogLastRow, 3) = Now()
                        .Cells(mlLogLastRow, 4) = "Retry"
                    End With
                Else
                    'schedule this for later pickup
                    mintTimeOuts = 1
                    With Sh
                        If mlLogLastRow = 0 Then
                            mlLogLastRow = .UsedRange.Rows.Count + 1
                        End If
                        .Cells(mlLogLastRow, 1) = mwsParameters.Cells(3, 1)
                        .Cells(mlLogLastRow, 3) = Now()
                        .Cells(mlLogLastRow, 4) = "Rescheduled after 60 minutes"
                    End With
                End If
            
            Case -2147220975  'Likely cause, Incorrect Password
                Sh.Cells(1, 1) = 2
                With Sh
                    mlLogLastRow = .UsedRange.Rows.Count + 1
                    .Cells(mlLogLastRow, 1) = mwsParameters.Cells(3, 1)
                    .Cells(mlLogLastRow, 3) = Now()
                    .Cells(mlLogLastRow, 4) = "Failed"
                End With
            
            Case Else   'Report Other Errors
                If mintTimeOuts < 10 Then
                    Sh.Cells(1, 1) = 1
                    Sh.Cells(1, 2) = 60
                    mintTimeOuts = mintTimeOuts + 1
                    With Sh
                        If mlLogLastRow = 0 Then
                            mlLogLastRow = .UsedRange.Rows.Count + 1
                        End If
                        .Cells(mlLogLastRow, 1) = mwsParameters.Cells(3, 1)
                        .Cells(mlLogLastRow, 3) = Now()
                        .Cells(mlLogLastRow, 4) = "Retry"
                    End With
                Else
                    'schedule this for later pickup
                    mintTimeOuts = 1
                    With Sh
                        If mlLogLastRow = 0 Then
                            mlLogLastRow = .UsedRange.Rows.Count + 1
                        End If
                        .Cells(mlLogLastRow, 1) = mwsParameters.Cells(3, 1)
                        .Cells(mlLogLastRow, 3) = Now()
                        .Cells(mlLogLastRow, 4) = "Rescheduled after 60 minutes"
                    End With
                End If
    End Select

End Sub
