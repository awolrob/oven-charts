Attribute VB_Name = "StartUp"
Option Explicit

Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Global sFolderPath As String
Dim sFileName, sTitle, sMessage, sOvenFileName, sYear, sYear4, sMonth, sDay As String
Dim wbResult As Workbook
Dim WB As Workbook
Dim dDate As Date
Dim sLastRow As String
Dim sFindMarkerRange As Range
Dim iResponse As Integer
Dim bFilesFound As Boolean

Sub Main()
    
    frmSplash.lblStatus.Visible = False
    frmSplash.lblStatus2.Visible = False
    frmSplash.Show
    frmSplash.Refresh
    Sleep (1000)

    sFolderPath = "X:\OVENS\"
    
    Do
        sTitle = "Enter Date To Run Daily Oven Chart"
        sMessage = "Enter Two Digit Year in YY format"
        sYear = Format(Date, "yy")
        sYear = InputBox(sMessage, sTitle, sYear)
        sYear4 = Format(Date, "yyyy")
        
        sTitle = "Enter Date To Run Daily Oven Chart"
        sMessage = "Enter Two Digit Month in MM format"
        sMonth = Format(Date, "mm")
        sMonth = InputBox(sMessage, sTitle, sMonth)
        If Len(sMonth) <> 2 Then
          sMonth = "0" & sMonth
        End If
        
        sTitle = "Enter Date To Run Daily Oven Chart"
        sMessage = "Enter Two Digit Day in DD format"
        sDay = Format(Date, "dd")
        sDay = InputBox(sMessage, sTitle, sDay)
        If Len(sDay) <> 2 Then
          sDay = "0" & sDay
        End If
        
        sOvenFileName = sYear4 & "/" & sMonth & "/" & sDay
    
        If Not IsDate(sOvenFileName) Then
            iResponse = MsgBox("Invaild Date.  Please Reenter Or Abort (YYMMDD) ", vbRetryCancel, "Please enter a valid date")
        End If
    
        If iResponse = vbRetry Then
            ' stay in look
        End If
    Loop While Not IsDate(sOvenFileName) And Not iResponse = vbCancel
            
    If IsDate(sOvenFileName) And (Not iResponse = vbCancel) Then
        sOvenFileName = "*" & sYear & sMonth & sDay & "*.csv"
        CombineCsvs
        If bFilesFound Then
            Ovens
        Else
            frmSplash.lblStatus.Caption = "Search FAILED: " & sOvenFileName
            frmSplash.lblStatus.ForeColor = vbRed
            frmSplash.lblStatus.Visible = True
            frmSplash.lblStatus2.Caption = "*** No Oven Data Files Found - Canceling *** "
            frmSplash.lblStatus2.ForeColor = vbRed
            frmSplash.lblStatus2.Visible = True
            DoEvents
            Sleep (1000)
       End If
    End If
    Unload frmSplash
End Sub

Sub CombineCsvs()

  bFilesFound = False
  
  frmSplash.lblStatus.Caption = frmSplash.lblStatus.Caption & "... Search String: " & sOvenFileName
  frmSplash.lblStatus.Visible = True
    
  frmSplash.lblStatus2.Caption = " "
  frmSplash.lblStatus2.Visible = True
  DoEvents
  
  sFileName = Dir(sFolderPath & sOvenFileName)
  
  If sFileName = vbNullString Then
    '* do nothing
  Else
      Set wbResult = Workbooks.Add

      Application.DisplayAlerts = False
      Application.ScreenUpdating = False
      
      Do While sFileName <> vbNullString
        bFilesFound = True
        frmSplash.lblStatus2.Caption = sFolderPath & sFileName
        DoEvents
        Set WB = Workbooks.Open(sFolderPath & sFileName)
        WB.ActiveSheet.UsedRange.Copy wbResult.ActiveSheet.UsedRange.Rows(wbResult.ActiveSheet.UsedRange.Rows.Count).Offset(1).Resize(1)
        WB.Close False
        sFileName = Dir()
      Loop
      
    '*       Debug code to watch excel process
    '        Application.Visible = True
    '        Application.ScreenUpdating = True
    '        Application.DisplayAlerts = True
          
      wbResult.ActiveSheet.Rows(1).EntireRow.Delete
  End If
  
End Sub

Sub Ovens()
'
' Ovens Macro
'
' Keyboard Shortcut: Ctrl+Shift+O
'
    
    frmSplash.lblStatus.Caption = "Formatting Data"
    frmSplash.lblStatus2.Visible = False
    DoEvents
    
    Cells.Select
    Selection.Columns.AutoFit
    
    frmSplash.lblStatus.Caption = "Sorting Data"
    frmSplash.lblStatus2.Visible = False
    DoEvents
    
    With ActiveSheet
        sLastRow = .Cells(.Rows.Count, "C").End(xlUp).Row
        .Sort.SortFields.Clear
    
        .Sort.SortFields.Add2 Key:=Range( _
            "A2:A" & sLastRow), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
            xlSortNormal
        .Sort.SortFields.Add2 Key:=Range( _
            "B2:B" & sLastRow), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
            xlSortNormal
    End With
    
    With ActiveSheet.Sort
        .SetRange Range("A1:N" & sLastRow)
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With

    frmSplash.lblStatus.Caption = "Formatting Data"
    DoEvents

    Columns("B:B").Select
    Selection.NumberFormat = "hh:mm;@"

    Range("A1").Select
    Set sFindMarkerRange = Cells.Find(What:="Marker", After:=ActiveCell, LookIn:=xlFormulas2, _
        LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
        MatchCase:=False, SearchFormat:=False)
    Do While Not sFindMarkerRange Is Nothing
        Range(sFindMarkerRange.Address).Activate
        Selection.EntireRow.Delete
        Selection.EntireRow.Delete
        
        Range("A1").Select
        Set sFindMarkerRange = Cells.Find(What:="Marker", After:=ActiveCell, LookIn:=xlFormulas2, _
            LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
            MatchCase:=False, SearchFormat:=False)
    Loop
    
    dDate = Range("A2").Value

    frmSplash.lblStatus.Caption = "Generating Chart"
    DoEvents

    Columns("B:H").Select
    ActiveSheet.Shapes.AddChart2(227, xlLineMarkers).Select
    ActiveChart.SetSourceData Source:=Range("$B:$H")
    
    ActiveChart.FullSeriesCollection(1).Delete
    ActiveChart.FullSeriesCollection(1).XValues = Sheets(1).Range("$B:$B")
    ActiveChart.Location Where:=xlLocationAsNewSheet
    ActiveChart.ChartTitle.Text = dDate
    
    frmSplash.lblStatus.Caption = "Saving File"
    DoEvents
    
    ChDir "X:\Reports"
    
    Application.DisplayAlerts = False
    ActiveWorkbook.SaveAs filename:="X:\Reports\" & sYear4 & "-" & sMonth & "-" & sDay & ".xlsx", FileFormat _
        :=xlOpenXMLWorkbook, CreateBackup:=False, AccessMode:=xlExclusive, _
        ConflictResolution:=Excel.XlSaveConflictResolution.xlLocalSessionChanges
    ActiveWorkbook.Close True
    
    Set WB = Workbooks.Open("X:\Reports\" & sYear4 & "-" & sMonth & "-" & sDay & ".xlsx")
    Application.Visible = True
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True

    AppActivate Application.Caption
    
    frmSplash.lblStatus.Caption = "Complete"
    DoEvents
    
End Sub
