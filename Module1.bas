Attribute VB_Name = "Module1"
Option Explicit
Public strSQLDriver As String
Public Type DBAttenStats
    ExTypeName As String
    ExTypeCount As Long
    ExTypePct As Double
End Type
Public AttenStats() As DBAttenStats
Public strChartData()
Public strUsername                      As String, strPassword As String, strServerAddress As String
Public ControlFocus(8)                  As Integer
Public bolAlarmOKed                     As Boolean
Public NewEmp                           As Boolean
Public SelGUID                          As String, SelDate As String, SelExcuse As String, SelType As String, SelHours As String
Public intEmpNum()                      As String
Public strListLine()                    As String
Public intGridRow                       As Integer
Public strReportType                    As String
Public intRowsAdded                     As Integer, intExcusedRowsAdded As Integer, intUnExcusedRowsAdded As Integer, intOtherRowsAdded As Integer
Public strReportName                    As String, strReportNum As String
Public strReportTitle                   As String, strReportSubTitle As String
Public strReportMsg                     As String
Public strReportInfo                    As String, strReportEntryCount As String
Public bolStop                          As Boolean
Public bolOpenEmp                       As Boolean
Public bolIsDateRange                   As Boolean
Public DTStartDate                      As Date
Public DTEndDate                        As Date
Public bolCancelPrint                   As Boolean
Public strCurrentEmpInfo                As EmpInfo
Public strEmpInfo()                     As EmpInfo
Public moApp                            As Word.Application
Private mbKillMe                        As Boolean
Public Const strDBDateFormat            As String = "YYYY-MM-DD"
Public Const strUserDateFormat          As String = "MM/DD/YYYY"
Public Const intPartialUnExcusedAllowed As Integer = 12
Public Const intPartialExcusedAllowed   As Integer = 12
Public Const intFullUnExcusedAllowed    As Integer = 6
Public Const intFullExcusedAllowed      As Integer = 6
Public Flashes                          As Integer
Public NoEntries                        As Boolean
Public strConfirmedAlarms()             As String
Public SelectedFilters                  As Integer
Public bolPartialUnExcusedExceeded      As Boolean, bolFullUnExcusedExceeded As Boolean, bolFullExcusedExceeded As Boolean, bolPartialExcusedExceeded As Boolean
Public bolAlarmsCalled                  As Boolean
Public strAlarmTitleString              As String
Public Const intSearchWait              As Integer = 2
Public intQryIndex                      As Integer
Public lngQryTimes()                    As Long
Public strTimeRemaining                 As String
Public lngTimeRemainingArray()          As Long
Type AnniDates
    CurrentYear As Date
    CurrentYearPlus1Week As Date
    PreviousYear As Date
    PreviousYearSub1Week As Date
End Type
Type MonthDates
    StartDate As Date
    EndDate As Date
End Type
Public bolVacationOpen As Boolean
Public dtFiscalYearEnd As Date
Public dtVacaPeriod    As MonthDates
Public Type EmpInfo
    Name As String
    Location1 As String
    Location2 As String
    Number As String
    HireDate As Date
    VacaWeeks As Integer
    VacaHours As Single
    IsActive As Boolean
End Type
Public Type Tenure
    YearsWorked As Integer
    VacaWeeksAvail As Integer
    VacaHoursAvail As Single
End Type
Public Type DBStats
    TotalAttenEntries As Integer
    TotalVacaEntries As Integer
    TotalEmployees As Integer
End Type
Public DataBaseStats As DBStats
Public Type AttenExcuses
    PartialExcused As Integer
    FullExcused As Integer
    FullUnExcused As Integer
    PartialUnExcused As Integer
End Type
Public AttenVals As AttenExcuses
Declare Function QueryPerformanceCounter Lib "kernel32" (X As Currency) As Boolean
Declare Function QueryPerformanceFrequency Lib "kernel32" (X As Currency) As Boolean
Public total As Currency
Public Ctr1  As Currency, Ctr2 As Currency, Freq As Currency
Public Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer
Public Const VK_TAB = &H9
Public intPartialUnExcused As Integer, intFullUnExcused As Integer, intFullExcused As Integer, intPartialExcused As Integer
Public lngAddQry           As Double
Public Const colGridBusy   As Long = &HE0E0E0        '&H80FF80
Const HKEY_LOCAL_MACHINE = &H80000002
Private Declare Function RegOpenKeyEx _
                Lib "advapi32.dll" _
                Alias "RegOpenKeyExA" (ByVal hKey As Long, _
                                       ByVal lpSubKey As String, _
                                       ByVal ulOptions As Long, _
                                       ByVal samDesired As Long, _
                                       phkResult As Long) As Long
Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Private Declare Function RegEnumValue _
                Lib "advapi32.dll" _
                Alias "RegEnumValueA" (ByVal hKey As Long, _
                                       ByVal dwIndex As Long, _
                                       ByVal lpValueName As String, _
                                       lpcbValueName As Long, _
                                       ByVal lpReserved As Long, _
                                       lpType As Long, _
                                       lpData As Any, _
                                       lpcbData As Long) As Long
Private Declare Sub CopyMemory _
                Lib "kernel32" _
                Alias "RtlMoveMemory" (dest As Any, _
                                       Source As Any, _
                                       ByVal numBytes As Long)
Const REG_SZ = 1
Const REG_EXPAND_SZ = 2
Const REG_BINARY = 3
Const REG_DWORD = 4
Const REG_MULTI_SZ = 7
Const ERROR_MORE_DATA = 234
Const KEY_READ = &H20019 ' ((READ_CONTROL Or KEY_QUERY_VALUE Or
Public Type EmpExist
    Exist As Boolean
    Number As String
End Type
Global cn_Global As New ADODB.Connection

Public Function Addres_Excel(ByVal lng_row As Long, ByVal lng_col As Long) As String
    'this function is used to send the columns from grid to excel'
    'make column header to look like the letters used in excel'
    'for example for col 1 the first column we will send "1" and will return "A"'
    Dim modval As Long  'used to get the reminder'
    Dim strval As String   'get the transferd letter'
    modval = (lng_col - 1) Mod 26   'using mode we get the reminder. 26 is for the letters in engl.'
    strval = Chr$(Asc("A") + modval) 'using the reminder we get the letter'
    modval = ((lng_col - 1) \ 26) - 1 'check to see if it is not addres like "AA"'
    If modval >= 0 Then strval = Chr$(Asc("A") + modval) & strval 'if we have more then we add the letter'
    Addres_Excel = strval & lng_row 'return the value to the function'
End Function
' KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY) And (Not
' SYNCHRONIZE))
' Enumerate values under a given registry key
'
' returns a collection, where each element of the collection
' is a 2-element array of Variants:
' element(0) is the value name, element(1) is the value's value
Function EnumRegistryValues(ByVal hKey As Long, ByVal KeyName As String) As Collection
    Dim handle            As Long
    Dim index             As Long
    Dim valueType         As Long
    Dim Name              As String
    Dim nameLen           As Long
    Dim resLong           As Long
    Dim resString         As String
    Dim dataLen           As Long
    Dim valueInfo(0 To 1) As Variant
    Dim retVal            As Long
    ' initialize the result
    Set EnumRegistryValues = New Collection
    ' Open the key, exit if not found.
    If Len(KeyName) Then
        If RegOpenKeyEx(hKey, KeyName, 0, KEY_READ, handle) Then Exit Function
        ' in all cases, subsequent functions use hKey
        hKey = handle
    End If
    Do
        ' this is the max length for a key name
        nameLen = 260
        Name = Space$(nameLen)
        ' prepare the receiving buffer for the value
        dataLen = 4096
        ReDim resBinary(0 To dataLen - 1) As Byte
        ' read the value's name and data
        ' exit the loop if not found
        retVal = RegEnumValue(hKey, index, Name, nameLen, ByVal 0&, valueType, resBinary(0), dataLen)
        ' enlarge the buffer if you need more space
        If retVal = ERROR_MORE_DATA Then
            ReDim resBinary(0 To dataLen - 1) As Byte
            retVal = RegEnumValue(hKey, index, Name, nameLen, ByVal 0&, valueType, resBinary(0), dataLen)
        End If
        ' exit the loop if any other error (typically, no more values)
        If retVal Then Exit Do
        ' retrieve the value's name
        valueInfo(0) = Left$(Name, nameLen)
        ' return a value corresponding to the value type
        Select Case valueType
            Case REG_DWORD
                CopyMemory resLong, resBinary(0), 4
                valueInfo(1) = resLong
            Case REG_SZ, REG_EXPAND_SZ
                ' copy everything but the trailing null char
                resString = Space$(dataLen - 1)
                CopyMemory ByVal resString, resBinary(0), dataLen - 1
                valueInfo(1) = resString
            Case REG_BINARY
                ' shrink the buffer if necessary
                If dataLen < UBound(resBinary) + 1 Then
                    ReDim Preserve resBinary(0 To dataLen - 1) As Byte
                End If
                valueInfo(1) = resBinary()
            Case REG_MULTI_SZ
                ' copy everything but the 2 trailing null chars
                resString = Space$(dataLen - 2)
                CopyMemory ByVal resString, resBinary(0), dataLen - 2
                valueInfo(1) = resString
            Case Else
                ' Unsupported value type - do nothing
        End Select
        ' add the array to the result collection
        ' the element's key is the value's name
        EnumRegistryValues.Add valueInfo, valueInfo(0)
        index = index + 1
    Loop
    ' Close the key, if it was actually opened
    If handle Then RegCloseKey handle
End Function
' get the list of ODBC drivers through the registry
'
' returns a collection of strings, each one holding the
' name of a driver, e.g. "Microsoft Access Driver (*.mdb)"
'
' requires the EnumRegistryValues function
Function GetODBCDrivers() As Collection
    Dim res    As Collection
    Dim values As Variant
    ' initialize the result
    Set GetODBCDrivers = New Collection
    ' the names of all the ODBC drivers are kept as values
    ' under a registry key
    ' the EnumRegistryValue returns a collection
    For Each values In EnumRegistryValues(HKEY_LOCAL_MACHINE, "Software\ODBC\ODBCINST.INI\ODBC Drivers")
        ' each element is a two-item array:
        ' values(0) is the name, values(1) is the data
        If StrComp(values(1), "Installed", 1) = 0 Then
            ' if installed, add to the result collection
            GetODBCDrivers.Add values(0), values(0)
        End If
    Next
End Function
Public Sub FindMySQLDriver()
    GetODBCDrivers
    Dim i           As Integer
    Dim strPossis() As String
    Dim blah
    ReDim strPossis(0)
    Debug.Print GetODBCDrivers.Count
    For i = 1 To GetODBCDrivers.Count
        If InStr(1, GetODBCDrivers.Item(i), "MySQL") Then
            strPossis(UBound(strPossis)) = GetODBCDrivers.Item(i)
            ReDim Preserve strPossis(UBound(strPossis) + 1)
        End If
    Next i
    For i = 0 To UBound(strPossis)
        Debug.Print i & " " & strPossis(i)
    Next i
    If UBound(strPossis) > 1 Then
        blah = MsgBox("Multiple MySQL Drivers detected!", vbExclamation + vbOKOnly, "Gasp!")
        strSQLDriver = strPossis(0)
    Else
        strSQLDriver = strPossis(0)
    End If
End Sub
Public Sub CountExcusesInYear(Excused As String, TimeOffType As String, EntryDate As Date)
    Dim LastYearDate As Date
    Dim CurrentDate  As String
    CurrentDate = Format$(Date, "YYYY-MM-DD")
    LastYearDate = Format$(DateAdd("yyyy", -1, CurrentDate), strDBDateFormat)
    If TimeOffType = "Left Early" And Excused = "UNEXCUSED" And DateDiff("d", Format$(LastYearDate, strDBDateFormat), EntryDate) >= 0 Or TimeOffType = "Left & Came Back" And Excused = "UNEXCUSED" And DateDiff("d", Format$(LastYearDate, strDBDateFormat), EntryDate) >= 0 Or TimeOffType = "Late for DAY" And Excused = "UNEXCUSED" And DateDiff("d", Format$(LastYearDate, strDBDateFormat), EntryDate) >= 0 Or TimeOffType = "Late from LUNCH" And Excused = "UNEXCUSED" And DateDiff("d", Format$(LastYearDate, strDBDateFormat), EntryDate) >= 0 Then AttenVals.PartialUnExcused = AttenVals.PartialUnExcused + 1
    If TimeOffType = "Called Off" And Excused = "UNEXCUSED" And DateDiff("d", Format$(LastYearDate, strDBDateFormat), EntryDate) >= 0 Or TimeOffType = "No Call, No Show" And Excused = "UNEXCUSED" And DateDiff("d", Format$(LastYearDate, strDBDateFormat), EntryDate) >= 0 Or TimeOffType = "Requested Day Off" And Excused = "UNEXCUSED" And DateDiff("d", Format$(LastYearDate, strDBDateFormat), EntryDate) >= 0 Then AttenVals.FullUnExcused = AttenVals.FullUnExcused + 1
    If TimeOffType = "Called Off" And Excused = "EXCUSED" And DateDiff("d", Format$(LastYearDate, strDBDateFormat), EntryDate) >= 0 Or TimeOffType = "No Call, No Show" And Excused = "EXCUSED" And DateDiff("d", Format$(LastYearDate, strDBDateFormat), EntryDate) >= 0 Or TimeOffType = "Requested Day Off" And Excused = "EXCUSED" And DateDiff("d", Format$(LastYearDate, strDBDateFormat), EntryDate) >= 0 Then AttenVals.FullExcused = AttenVals.FullExcused + 1
    If TimeOffType = "Left Early" And Excused = "EXCUSED" And DateDiff("d", Format$(LastYearDate, strDBDateFormat), EntryDate) >= 0 Or TimeOffType = "Left & Came Back" And Excused = "EXCUSED" And DateDiff("d", Format$(LastYearDate, strDBDateFormat), EntryDate) >= 0 Or TimeOffType = "Late for DAY" And Excused = "EXCUSED" And DateDiff("d", Format$(LastYearDate, strDBDateFormat), EntryDate) >= 0 Or TimeOffType = "Late from LUNCH" And Excused = "EXCUSED" And DateDiff("d", Format$(LastYearDate, strDBDateFormat), EntryDate) >= 0 Then AttenVals.PartialExcused = AttenVals.PartialExcused + 1
End Sub
Public Sub CountExcusesAll(Excused As String, TimeOffType As String)
    If TimeOffType = "Left Early" And Excused = "UNEXCUSED" Or TimeOffType = "Left & Came Back" And Excused = "UNEXCUSED" Or TimeOffType = "Late for DAY" And Excused = "UNEXCUSED" Or TimeOffType = "Late from LUNCH" And Excused = "UNEXCUSED" Then AttenVals.PartialUnExcused = AttenVals.PartialUnExcused + 1
    If TimeOffType = "Called Off" And Excused = "UNEXCUSED" Or TimeOffType = "No Call, No Show" And Excused = "UNEXCUSED" Or TimeOffType = "Requested Day Off" And Excused = "UNEXCUSED" Then AttenVals.FullUnExcused = AttenVals.FullUnExcused + 1
    If TimeOffType = "Called Off" And Excused = "EXCUSED" Or TimeOffType = "No Call, No Show" And Excused = "EXCUSED" Or TimeOffType = "Requested Day Off" And Excused = "EXCUSED" Then AttenVals.FullExcused = AttenVals.FullExcused + 1
    If TimeOffType = "Left Early" And Excused = "EXCUSED" Or TimeOffType = "Left & Came Back" And Excused = "EXCUSED" Or TimeOffType = "Late for DAY" And Excused = "EXCUSED" Or TimeOffType = "Late from LUNCH" And Excused = "EXCUSED" Then AttenVals.PartialExcused = AttenVals.PartialExcused + 1
End Sub
Public Sub GetDataBaseStats()
    DataBaseStats.TotalAttenEntries = 0
    DataBaseStats.TotalEmployees = 0
    DataBaseStats.TotalVacaEntries = 0
    Dim rs      As New ADODB.Recordset
    Dim strSQL1 As String
    Dim strSQL2 As String
    Dim strSQL3 As String
    strSQL1 = "SELECT COUNT(*) FROM attendb.attenentries attenentries_0"
    strSQL2 = "SELECT COUNT(*) FROM attendb.emplist emplist_0"
    strSQL3 = "SELECT COUNT(*) FROM attendb.vacations vacations_0"
    cn_Global.CursorLocation = adUseClient
    Set rs = cn_Global.Execute(strSQL1)
    DataBaseStats.TotalAttenEntries = rs.Fields(0)
    rs.Close
    Set rs = cn_Global.Execute(strSQL2)
    DataBaseStats.TotalEmployees = rs.Fields(0)
    rs.Close
    Set rs = cn_Global.Execute(strSQL3)
    DataBaseStats.TotalVacaEntries = rs.Fields(0)
    rs.Close
End Sub
Public Function GetTabState() As Boolean
    GetTabState = False
    If GetKeyState(VK_TAB) And -256 Then
        GetTabState = True
    End If
End Function
Public Function DateDiffW(BegDate As Date, EndDate As Date)
    Const SUNDAY = 1
    Const SATURDAY = 7
    Dim NumWeeks As Integer
    If BegDate > EndDate Then
        DateDiffW = 0
    Else
        Select Case Weekday(BegDate)
            Case SUNDAY: BegDate = BegDate + 1
            Case SATURDAY: BegDate = BegDate + 2
        End Select
        Select Case Weekday(EndDate)
            Case SUNDAY: EndDate = EndDate - 2
            Case SATURDAY: EndDate = EndDate - 1
        End Select
        NumWeeks = DateDiff("ww", BegDate, EndDate)
        DateDiffW = NumWeeks * 5 + Weekday(EndDate) - Weekday(BegDate) + 1
    End If
End Function
Public Function CalcYearsWorked(EmpNum As Variant) As Tenure
    Dim lngYearsWorked As Integer
    CalcYearsWorked.YearsWorked = 0
    CalcYearsWorked.VacaWeeksAvail = 0
    lngYearsWorked = DateDiff("yyyy", Format$(ReturnEmpInfo(EmpNum).HireDate, strUserDateFormat), Format$(Date, strUserDateFormat)) ' - 1
    CalcYearsWorked.YearsWorked = lngYearsWorked
    If ReturnEmpInfo(EmpNum).VacaWeeks <> 0 Then
        CalcYearsWorked.VacaWeeksAvail = ReturnEmpInfo(EmpNum).VacaWeeks
        Exit Function
    End If
    Select Case CalcYearsWorked.YearsWorked
        Case 1 To 4
            CalcYearsWorked.VacaHoursAvail = 80
        Case 5 To 11
            CalcYearsWorked.VacaHoursAvail = 120
        Case 12 To 20
            CalcYearsWorked.VacaHoursAvail = 160
        Case 21
            CalcYearsWorked.VacaHoursAvail = 168
        Case 22
            CalcYearsWorked.VacaHoursAvail = 176
        Case 23
            CalcYearsWorked.VacaHoursAvail = 184
        Case 24
            CalcYearsWorked.VacaHoursAvail = 192
        Case Is >= 25
            CalcYearsWorked.VacaHoursAvail = 200
    End Select
    CalcYearsWorked.VacaHoursAvail = CalcYearsWorked.VacaHoursAvail + 8 'Add 8 hours to all for floating holiday.
End Function
Public Sub GetEmpInfo()
    Dim rs      As New ADODB.Recordset
    Dim strSQL1 As String
    cn_Global.CursorLocation = adUseClient
    strSQL1 = "SELECT * FROM emplist order by idName"
    Set rs = cn_Global.Execute(strSQL1)
    With rs
        ReDim strEmpInfo(.RecordCount)
        Do Until .EOF
            strEmpInfo(.AbsolutePosition - 1).HireDate = !idHireDate
            strEmpInfo(.AbsolutePosition - 1).Location1 = !idLocation1
            strEmpInfo(.AbsolutePosition - 1).Location2 = !idLocation2
            strEmpInfo(.AbsolutePosition - 1).Name = !idName
            strEmpInfo(.AbsolutePosition - 1).Number = !idNumber
            strEmpInfo(.AbsolutePosition - 1).VacaHours = !idVacaHours
            strEmpInfo(.AbsolutePosition - 1).IsActive = !idIsActive
            .MoveNext
        Loop
        .Close
    End With
End Sub
Public Sub GetCurrentEmp(EmpNum As String)
    strCurrentEmpInfo.Number = EmpNum
    strCurrentEmpInfo.Name = ReturnEmpInfo(EmpNum).Name
    strCurrentEmpInfo.HireDate = ReturnEmpInfo(EmpNum).HireDate
    strCurrentEmpInfo.VacaWeeks = ReturnEmpInfo(EmpNum).VacaWeeks
    strCurrentEmpInfo.Location2 = ReturnEmpInfo(EmpNum).Location2
    strCurrentEmpInfo.Location1 = ReturnEmpInfo(EmpNum).Location1
End Sub
Public Function ReturnEmpInfo(strEmpNum As Variant) As EmpInfo
    ReturnEmpInfo.Name = vbNull
    ReturnEmpInfo.HireDate = vbNull
    ReturnEmpInfo.IsActive = vbNull
    ReturnEmpInfo.Location1 = vbNull
    ReturnEmpInfo.Location2 = vbNull
    ReturnEmpInfo.Number = vbNull
    ReturnEmpInfo.VacaHours = vbNull
    Dim i As Integer
    For i = 0 To UBound(strEmpInfo)
        If strEmpInfo(i).Number = strEmpNum Then
            ReturnEmpInfo.Name = strEmpInfo(i).Name
            ReturnEmpInfo.HireDate = strEmpInfo(i).HireDate
            ReturnEmpInfo.IsActive = strEmpInfo(i).IsActive
            ReturnEmpInfo.Location1 = strEmpInfo(i).Location1
            ReturnEmpInfo.Location2 = strEmpInfo(i).Location2
            ReturnEmpInfo.Number = strEmpInfo(i).Number
            ReturnEmpInfo.VacaHours = strEmpInfo(i).VacaHours
            Exit Function
        End If
    Next i
    MsgBox (strEmpNum & " not found")
End Function
Public Function CheckForName(FirstLast As String) As EmpExist
    Dim i As Integer
    CheckForName.Exist = False
    For i = 0 To UBound(strEmpInfo)
        If strEmpInfo(i).Name = UCase$(FirstLast) Then
            CheckForName.Exist = True
            CheckForName.Number = strEmpInfo(i).Number
        End If
    Next i
End Function
Public Sub ClearEmpInfo()
    strCurrentEmpInfo.HireDate = vbNull
    strCurrentEmpInfo.IsActive = vbNull
    strCurrentEmpInfo.Location1 = vbNullString
    strCurrentEmpInfo.Location2 = vbNullString
    strCurrentEmpInfo.Name = vbNullString
    strCurrentEmpInfo.Number = vbNullString
    strCurrentEmpInfo.VacaWeeks = vbNull
End Sub
Public Function dtAnniDate(HireDate As Date, Optional Period As Integer) As AnniDates
    dtAnniDate.CurrentYear = Date
    dtAnniDate.CurrentYearPlus1Week = Date
    dtAnniDate.PreviousYear = Date
    dtAnniDate.PreviousYearSub1Week = Date
    Dim strHireDate        As String, strCurrentDate As String
    Dim strSplitHireDate() As String, strSplitCurDate() As String
    strHireDate = Format$(HireDate, "YYYY-MM-DD")
    strCurrentDate = Format$(Date, "YYYY-MM-DD")
    strSplitHireDate = Split(strHireDate, "-")
    strSplitCurDate = Split(strCurrentDate, "-")
    If Period > 0 Then
        If DateDiff("d", strSplitCurDate(0) & "-" & strSplitHireDate(1) & "-" & strSplitHireDate(2), DateTime.Date) >= 1 Then  'if current date is past current anni date, then calc for current anni to next year anni
            dtAnniDate.PreviousYear = strSplitCurDate(0) - Period & "-" & strSplitHireDate(1) & "-" & strSplitHireDate(2)
            dtAnniDate.CurrentYear = strSplitCurDate(0) + 1 - Period & "-" & strSplitHireDate(1) & "-" & strSplitHireDate(2)
        Else
            dtAnniDate.CurrentYear = strSplitCurDate(0) - Period & "-" & strSplitHireDate(1) & "-" & strSplitHireDate(2)
            dtAnniDate.PreviousYear = strSplitCurDate(0) - 1 - Period & "-" & strSplitHireDate(1) & "-" & strSplitHireDate(2)
        End If
    Else
        If DateDiff("d", strSplitCurDate(0) & "-" & strSplitHireDate(1) & "-" & strSplitHireDate(2), DateTime.Date) >= 1 Then  'if current date is past current anni date, then calc for current anni to next year anni
            dtAnniDate.PreviousYear = strSplitCurDate(0) & "-" & strSplitHireDate(1) & "-" & strSplitHireDate(2)
            dtAnniDate.CurrentYear = strSplitCurDate(0) + 1 & "-" & strSplitHireDate(1) & "-" & strSplitHireDate(2)
        Else
            dtAnniDate.CurrentYear = strSplitCurDate(0) & "-" & strSplitHireDate(1) & "-" & strSplitHireDate(2)
            dtAnniDate.PreviousYear = strSplitCurDate(0) - 1 & "-" & strSplitHireDate(1) & "-" & strSplitHireDate(2)
        End If
    End If
    dtAnniDate.PreviousYearSub1Week = DateAdd("ww", -1, dtAnniDate.PreviousYear) ' Prev anni minus one week (to show vacas on fringe)
    dtAnniDate.CurrentYearPlus1Week = DateAdd("ww", 1, dtAnniDate.CurrentYear)
End Function
Public Function GetMonthDates(Month As Integer, Year As Integer) As MonthDates
    On Error Resume Next
    With GetMonthDates
        .StartDate = Month & "/1/" & Year '"1/1/2012"
        .EndDate = DateAdd("m", 1, .StartDate)
        .EndDate = DateAdd("d", -1, .EndDate)
    End With
End Function
Public Function BoundedText(ByVal ptr As Object, _
                            ByVal txt As String, _
                            ByVal max_wid As Single) As String
    Dim intHalfWid As Integer
    If Printer.TextWidth(txt) > max_wid Then
        intHalfWid = Len(txt) / 2
        Mid$(txt, intHalfWid, 0) = " " & vbCrLf & " "
    End If
    BoundedText = txt
End Function
Public Sub PrintFlexGridSGrid(FlexGrid As vbalGrid, _
                              Optional sTitle As String, _
                              Optional sSubTitle As String)
    FlexGrid.Redraw = False
    Dim HWidth          As Integer
    Dim intPadding      As Integer
    Dim PrevX           As Integer, PrevY As Integer, intMidStart As Integer, intMidLen As Integer, intTotLen As Integer, intPossibleLen As Integer
    Dim strSizedTxt     As String, strOrigTxt As String
    Dim bolLongLine     As Boolean, bolFirstLoop As Boolean
    Dim TwipPix         As Long
    Dim intCenterOffset As Long
    bolLongLine = False
    On Error Resume Next
    Printer.Font.Underline = False
    Printer.ScaleMode = 1
    Printer.Orientation = vbPRORLandscape
    Printer.DrawWidth = 1
    Printer.DrawStyle = vbSolid
    With Printer
        .ScaleMode = 1
        Printer.Print
        .FontSize = 20
        HWidth = Printer.TextWidth(sTitle) / 2
        Printer.CurrentX = Printer.ScaleWidth / 2 - HWidth
        Printer.Print sTitle
        Printer.FontSize = 8
    End With
    Printer.FontSize = 7
    Printer.Print "    " & sSubTitle
    Printer.Print ""
    Printer.Print "    Report date: " & Date & " " & Time & "      Printed by: " & UCase$(Environ$("USERNAME"))
    Const GAP = 40
    Dim xmax As Single, xmin As Single
    xmin = 300
    xmax = 14800
    Dim ymax As Single, ymin As Single
    ymin = 1500
    ymax = 10800
    Dim X As Single, XFirstColumn As Single
    Dim c As Integer, cc As Integer
    Dim R As Integer
    intMidStart = 1
    With Printer.Font
        .Name = FlexGrid.Font.Name
        .Size = 9
    End With
    With FlexGrid
        PrevX = Printer.CurrentX
        PrevY = Printer.CurrentY
        Printer.CurrentX = xmax - 600
        Printer.CurrentY = ymax + 300
        Printer.ForeColor = vbBlack
        Printer.Font.Underline = False
        Printer.Print "Page " & Printer.Page
        Printer.CurrentX = PrevX
        Printer.CurrentY = PrevY
        Printer.CurrentY = ymin
        intPadding = 150
        frmPBar.PBar1.Max = .Rows
        frmPBar.PBar1.Value = 0
        Form1.tmrUpdateTimeRemaining.Enabled = False
        frmPBar.lblInfo.Caption = "Spooling..."
        TwipPix = .ColumnWidth(c) * Screen.TwipsPerPixelX
        XFirstColumn = xmin + TwipPix * GAP
        X = xmin + GAP
        If FlexGrid.Header = True Then
            For c = 1 To .Columns
                Printer.CurrentX = X
                TwipPix = .ColumnWidth(c) * Screen.TwipsPerPixelX
                Printer.Print BoundedText(Printer, .ColumnHeader(c), TwipPix);
                TwipPix = .ColumnWidth(c) * Screen.TwipsPerPixelX
                X = X + TwipPix + 2 * GAP
            Next c
        End If
        Printer.Print
        For R = 1 To .Rows ' - 1
            If bolStop = True Then
                Printer.EndDoc
                bolStop = False
                frmPBar.Visible = False
                Exit Sub
            End If
            frmPBar.PBar1.Value = R ' * .Cols
            frmPBar.lblQryTime = "Row " & R & " of " & .Rows
            DoEvents
            ' Draw a line above this row.
            If R > 0 Then
                Printer.Line (XFirstColumn, Printer.CurrentY)-(xmax, Printer.CurrentY), vbBlack
            End If
            Printer.CurrentY = Printer.CurrentY + GAP
            ' Print the entries on this row.
            X = xmin + GAP
            For c = 1 To .Columns
                If frmPrinters.optCenterJust And c < .Columns Then
                    intCenterOffset = ((.ColumnWidth(c) * Screen.TwipsPerPixelX) / 2) - (Printer.TextWidth(.CellText(R, c)) / 2)
                Else
                    intCenterOffset = 0
                End If
                Printer.CurrentX = X
                If .CellText(R, c) <> "" And Printer.TextWidth(.CellText(R, c)) + intPadding >= xmax - Printer.CurrentX Then
                    strOrigTxt = .CellText(R, c)
                    intTotLen = 1
                    bolFirstLoop = True
                    Do Until intTotLen >= Len(strOrigTxt)
                        intMidLen = Len(strOrigTxt) - intMidStart + 1
                        If Not bolFirstLoop Then
                            strSizedTxt = Mid$(strOrigTxt, intMidStart, intPossibleLen)
                        Else
                            strSizedTxt = strOrigTxt
                            'intMidLen = (Len(strOrigTxt) - Round(((Printer.TextWidth(.CellText(R, c)) + intPadding) - (xmax - Printer.CurrentX)) / (Len(strOrigTxt) / 2), 0)) '80
                        End If
                        Dim lngColEnd  As Long
                        Dim intOrigLen As Integer
                        intOrigLen = Len(strOrigTxt)
                        lngColEnd = xmax - Printer.CurrentX
                        Do Until Printer.TextWidth(strSizedTxt) + intPadding <= lngColEnd Or intTotLen >= intOrigLen
                            strSizedTxt = Mid$(strOrigTxt, intMidStart, intMidLen)
                            intMidLen = intMidLen - 1
                        Loop
                        If Not bolFirstLoop Then
                            intMidStart = intMidStart + intPossibleLen ' + 1
                        Else
                            intPossibleLen = Len(strSizedTxt)
                            intMidStart = intMidStart + intMidLen + 1
                        End If
                        bolFirstLoop = False
                        intTotLen = intTotLen + Len(strSizedTxt) + 1
                        Printer.Font.Underline = .CellFont(R, c).Underline
                        If .CellFont(R, c).Underline = True Then
                            Printer.ForeColor = vbBlack
                        Else
                            Printer.ForeColor = &H808080
                        End If
                        Printer.Print strSizedTxt
                        If Printer.CurrentY >= ymax Then ' new page
                            Printer.Line (XFirstColumn, ymin)-(xmax, Printer.CurrentY), vbBlack, B
                            X = xmin
                            For cc = 1 To .Columns - 1
                                TwipPix = .ColumnWidth(cc) * Screen.TwipsPerPixelX
                                X = X + TwipPix + 2 * GAP
                                Printer.Line (X, ymin)-(X, Printer.CurrentY), vbBlack
                            Next cc
                            Printer.NewPage
                            Printer.CurrentX = xmax - 600
                            Printer.CurrentY = ymax + 300
                            Printer.ForeColor = vbBlack
                            Printer.Font.Underline = False
                            Printer.Print "Page " & Printer.Page
                            Printer.CurrentX = xmin
                            ymin = 400
                            Printer.CurrentY = ymin
                        End If
                        Printer.CurrentX = X + GAP
                        strSizedTxt = ""
                    Loop
                    intMidStart = 1
                    intMidLen = 0
                    intTotLen = 0
                    strSizedTxt = ""
                    bolLongLine = True
                Else
                    'bolLongLine = False
                    Printer.CurrentX = X + intCenterOffset
                    TwipPix = .ColumnWidth(c) * Screen.TwipsPerPixelX
                    Printer.Font.Underline = .CellFont(R, c).Underline
                    If .CellFont(R, c).Underline = True Then
                        Printer.ForeColor = vbBlack
                    Else
                        Printer.ForeColor = &H808080
                    End If
                    Printer.Print BoundedText(Printer, .CellText(R, c), TwipPix);
                End If
                TwipPix = .ColumnWidth(c) * Screen.TwipsPerPixelX
                X = X + TwipPix + 2 * GAP
            Next c
            Printer.CurrentY = Printer.CurrentY + GAP
            ' Move to the next line.
            If bolLongLine = True Then
                bolLongLine = False
            Else
                Printer.Print
                bolLongLine = False
            End If
            ' if near end of page, start a new one
            If Printer.CurrentY >= ymax Then
                Printer.Line (XFirstColumn, ymin)-(xmax, Printer.CurrentY), vbBlack, B
                X = xmin
                For c = 1 To .Columns - 1 '3
                    TwipPix = .ColumnWidth(c) * Screen.TwipsPerPixelX '+ GAP
                    X = X + TwipPix + 2 * GAP
                    Printer.Line (X, ymin)-(X, Printer.CurrentY), vbBlack 'ymax
                Next c
                Printer.NewPage
                Printer.CurrentX = xmax - 600
                Printer.CurrentY = ymax + 300
                Printer.ForeColor = vbBlack
                Printer.Font.Underline = False
                Printer.Print "Page " & Printer.Page
                Printer.CurrentX = xmin
                ymin = 400
                Printer.CurrentY = ymin
            End If
        Next R
        ymax = Printer.CurrentY
        'Draw a box around everything.
        Printer.Line (XFirstColumn, ymin)-(xmax, ymax), vbBlack, B
        X = xmin
        ' Draw lines between the columns.
        For c = 1 To .Columns - 1 '3
            TwipPix = .ColumnWidth(c) * Screen.TwipsPerPixelX
            X = X + TwipPix + 2 * GAP
            'vbBlack
            Printer.Line (X, ymin)-(X, Printer.CurrentY), vbBlack 'Printer.CurrentY
        Next c
    End With
    frmPBar.Hide
    Form1.tmrUpdateTimeRemaining.Enabled = True
    Printer.EndDoc
    FlexGrid.Redraw = True
    ' blah = MsgBox("Report has been sent to " & Printer.DeviceName, vbOKOnly, "Print Report")
End Sub
Public Function CheckForEmp(EmpNum As String) As Boolean
    Dim rs As New ADODB.Recordset
    Dim cn As New ADODB.Connection
    CheckForEmp = False
    Dim i As Integer
    For i = 0 To UBound(strEmpInfo)
        If strEmpInfo(i).Number = EmpNum Then
            CheckForEmp = True
            Exit Function
        End If
    Next i
    CheckForEmp = False
End Function
Public Sub ShowData()
    StartTimer
End Sub
Public Sub StartTimer()
    total = 0
    QueryPerformanceFrequency Freq
    QueryPerformanceCounter Ctr1
End Sub
Public Function StopTimer() As Double
    StopTimer = 0
    QueryPerformanceCounter Ctr2
    total = total + (Ctr2 - Ctr1)
    StopTimer = Round(CDbl(total / Freq) * 1000, 3)
End Function
Public Sub HideData()
    Dim lngCurQry As Double, lngAvgQry As Double
    lngCurQry = StopTimer
    intQryIndex = intQryIndex + 1
    ReDim Preserve lngQryTimes(intQryIndex)
    lngQryTimes(intQryIndex) = lngCurQry
    lngAddQry = lngAddQry + lngQryTimes(intQryIndex)
    lngAvgQry = lngAddQry / UBound(lngQryTimes)
    strTimeRemaining = Round(((lngAvgQry * (frmPBar.PBar1.Max - frmPBar.PBar1.Value)) / 1000), 2) & " seconds remaining"
End Sub
Public Sub ClearAvgQryTimes()
    intQryIndex = 0
    lngAddQry = 0
    Erase lngQryTimes
    Erase lngTimeRemainingArray
    ReDim lngTimeRemainingArray(1)
End Sub
Public Property Get KillMe() As Boolean
    InitializeMe
    KillMe = mbKillMe
End Property
Public Property Let KillMe(Value As Boolean)
    mbKillMe = Value
End Property
Public Sub InitializeMe()
    On Error Resume Next
    '<INITIALIZE WORD>
    Set moApp = GetObject(, "Word.Application")
    If TypeName(moApp) <> "Nothing" Then
        Set moApp = GetObject(, "Word.Application")
    Else
        Set moApp = CreateObject("Word.Application")
        mbKillMe = True
    End If
End Sub
Public Function SpellMe(ByVal msSpell As String) As String
    On Error GoTo No_Bugs
    Dim oDoc     As Word.Document
    Dim iWSE     As Integer
    Dim iWGE     As Integer
    Dim sReplace As String
    Dim lResp    As Long
    If msSpell = vbNullString Then Exit Function
    InitializeMe
    Select Case moApp.Version
        Case "9.0", "10.0", "11.0", "12.0", "14.0"
            Set oDoc = moApp.Documents.Add(, , 1, True)
        Case "8.0"
            Set oDoc = moApp.Documents.Add
        Case Else
            MsgBox "Unsupported Version of Word.", vbOKOnly + vbExclamation, "VB/Office Guru™ SpellChecker™"
            Exit Function
    End Select
    Screen.MousePointer = vbHourglass
    App.OleRequestPendingTimeout = 999999
    oDoc.Words.First.InsertBefore msSpell
    iWSE = oDoc.SpellingErrors.Count
    iWGE = oDoc.GrammaticalErrors.Count
    '<CHECK SPELLING AND GRAMMER DIALOG BOX>
    If iWSE > 0 Or iWGE > 0 Then
        '<HIDE MAIN WORD WINDOW>
        moApp.Visible = False
        If (moApp.WindowState = wdWindowStateNormal) Or (moApp.WindowState = wdWindowStateMaximize) Then
            moApp.WindowState = wdWindowStateMinimize
        Else
            moApp.WindowState = wdWindowStateMinimize
        End If
        '</HIDE MAIN WORD WINDOW>
        '<PREP CHECK SPELLING OPTIONS DIALOG BOX (MODIFY TO YOUR PREFERENCES)>
        moApp.Dialogs(wdDialogToolsSpellingAndGrammar).Application.Options.CheckGrammarWithSpelling = True
        moApp.Dialogs(wdDialogToolsSpellingAndGrammar).Application.Options.SuggestSpellingCorrections = True
        moApp.Dialogs(wdDialogToolsSpellingAndGrammar).Application.Options.IgnoreUppercase = True
        moApp.Dialogs(wdDialogToolsSpellingAndGrammar).Application.Options.IgnoreInternetAndFileAddresses = True
        moApp.Dialogs(wdDialogToolsSpellingAndGrammar).Application.Options.IgnoreMixedDigits = False
        moApp.Dialogs(wdDialogToolsSpellingAndGrammar).Application.Options.ShowReadabilityStatistics = False
        '</PREP CHECK SPELLING OPTIONS DIALOG BOX (MODIFY TO YOUR PREFERENCES)>
        '<DO ACTUAL SPELL CHECKING>
        moApp.Visible = True
        moApp.Activate
        lResp = moApp.Dialogs(wdDialogToolsSpellingAndGrammar).Display
        '</DO ACTUAL SPELL CHECKING>
        If lResp < 0 Then
            moApp.Visible = True
            ' MsgBox "Corrections Being Updated!", vbOKOnly + vbInformation, App.ProductName
            Clipboard.Clear
            oDoc.Select
            oDoc.Range.Copy
            sReplace = Clipboard.GetText(1)
            '<FIX FOR POSSIBLE EXTRA LINE BREAK AT END OF TEXT>
            If (InStrRev(sReplace, Chr$(13) & Chr$(10))) = (Len(sReplace) - 1) Then
                sReplace = Mid$(sReplace, 1, Len(sReplace) - 2)
            End If
            '</FIX FOR POSSIBLE EXTRA LINE BREAK AT END OF TEXT>
            SpellMe = sReplace
        ElseIf lResp = 0 Then
            MsgBox "Spelling Corrections Have Been Canceled!", vbOKOnly + vbCritical, "VB/Office Guru™ SpellChecker"
            SpellMe = msSpell
        End If
    Else
        MsgBox "No Spelling Errors Found" & vbNewLine & "Or No Suggestions Available!", vbOKOnly + vbInformation, "VB/Office Guru™ SpellChecker"
        SpellMe = msSpell
    End If
    '</CHECK SPELLING AND GRAMMER DIALOG BOX>
    oDoc.Close False
    Set oDoc = Nothing
    '<HIDE WORD IF THERE ARE NO OTHER INSTANCES>
    If KillMe = True Then
        moApp.Visible = False
    End If
    '</HIDE WORD IF THERE ARE NO OTHER INSTANCES>
    Screen.MousePointer = vbNormal
    Exit Function
No_Bugs:
    If Err.Number = "91" Then
        Resume Next
    ElseIf Err.Number = "462" Then
        MsgBox "Spell Checking Is Temporary Un-Available!" & vbNewLine & "Try Again After Program Re-Start.", vbOKOnly + vbInformation, "ActiveX Server Not Responding"
        Screen.MousePointer = vbNormal
    ElseIf Err.Number = 429 Then
        Set moApp = Nothing
        Resume Next
    Else
        MsgBox Err.Number & " " & Err.Description, vbOKOnly + vbInformation, App.ProductName
        Screen.MousePointer = vbNormal
    End If
End Function
