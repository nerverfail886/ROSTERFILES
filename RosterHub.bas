Attribute VB_Name = "RosterHub"
Option Explicit

' =====================================================================
' RosterHub daily allocation
' One module for all seven day sheets (Monday ... Sunday).
'   Generate button -> GenerateRoster
'   Reset button    -> ResetRoster
' The macro works on the sheet the button was clicked on.
' =====================================================================

' ---------------------------------------------------------------------
' SETTINGS (change here only)
' ---------------------------------------------------------------------
Private Const DB_PATH As String = "T:\Riverside\Production\Cold Kitchen\roster\RosterDatabase.accdb"

Private Const ROSTER_SHEET As String = "roster"
Private Const BCR_SHEET As String = "BCR"
Private Const ROSTER_START_ROW As Long = 4    ' row 3 is the boss, not rostered
Private Const BCR_START_ROW As Long = 3

Private Const END_COLUMN As String = "B"      ' END anchor on the day sheet

' Unassigned areas
Private Const AM_UN_TITLE As String = "H27"
Private Const AM_UN_FIRST_ROW As Long = 28
Private Const AM_UN_COL As String = "G"
Private Const PM_UN_TITLE As String = "M33"
Private Const PM_UN_FIRST_ROW As Long = 34
Private Const PM_UN_COL As String = "L"

' Leave areas, starting at END row + 1
Private Const LEAVE_MAX_ROWS As Long = 14
Private Const AL_COL As String = "G"          ' AL and RDO
Private Const WC_COL As String = "L"          ' WCOMP and SL

' Colours
Private Const COLOR_ROSTER As Long = 14277081 ' RGB(217, 217, 217) grey
Private Const COLOR_BCR As Long = 13431551    ' RGB(255, 242, 204) light yellow
Private Const COLOR_WHITE As Long = 16777215  ' RGB(255, 255, 255)
Private Const COLOR_RED As Long = 255         ' RGB(255, 0, 0)

Private gSlots As Object   ' TaskID -> Collection of ID cell addresses
Private gToday As String   ' "MON" ... "SUN", set by BuildSlotList


' ---------------------------------------------------------------------
' SLOT LIST: which cells each task fills (ID cell; name goes one cell
' to the right, time two cells to the right). Order of cells = order
' of skill score (highest first).
' Optional third argument = days this line applies. Empty = every day.
' A task with no line for today is not assigned that day.
' The same task can have different cells on different days.
' ---------------------------------------------------------------------
Private Sub BuildSlotList(ByVal dayCode As String)
    Set gSlots = CreateObject("Scripting.Dictionary")
    gToday = dayCode

    ' ----- AM area (ID in column B) -----
    AddSlots 1, "B3, G3"                        ' Team Leader (2nd one = CRA Shift Leader G3)
    AddSlots 3, "B5"                            ' Garlic Bread (FEDEX/Garlic bread)
    AddSlots 9, "B6"                            ' Ice Cream (Ice cream/ICRM Garnish)
    AddSlots 2, "B7"                            ' Fedex/ICRM Garnish (ICRM GNH/FEDEX)
    AddSlots 4, "B10, B11, B12, B13"            ' SPML-AM
    AddSlots 5, "B15, B16, B17"                 ' Sandwich-AM
    AddSlots 6, "B20, B21"                      ' VA
    AddSlots 7, "B23"                           ' Cheese
    AddSlots 8, "B24"                           ' Garnish
    AddSlots 26, "B26"                          ' K1
    AddSlots 27, "B27"                          ' K2
    AddSlots 28, "B28"                          ' F1
    AddSlots 13, "B30, B31"                     ' Lettuce Check
    AddSlots 10, "B33, B34"                     ' Admin-AM
    AddSlots 33, "B36, B37", "TUE,THU,FRI,SAT,SUN"  ' HA SW

    ' ----- CRA area (ID in column G) -----
    AddSlots 14, "G6, G5"                       ' CRA/Protein, then CRA/SW
    AddSlots 15, "G7, G4"                       ' CRA/VEGE, then CRA/EY
    AddSlots 16, "G9, G10, G11, G12, G13"       ' CRA/Assembly
    AddSlots 17, "G16"                          ' CRA/Dressing
    AddSlots 18, "G17"                          ' CRA/Cooking
    AddSlots 19, "G18, G19"                     ' CRA/Runner
    AddSlots 20, "G21, G22, G23, G24, G25"      ' Fruit

    ' ----- PM area (ID in column L) -----
    AddSlots 21, "L3, L7"                       ' Shift Leader
    AddSlots 22, "L4, L5"                       ' Prep UA AA SB HA
    AddSlots 23, "L8, L9, L10, L11, L12"        ' Sandwich-PM
    AddSlots 24, "L14, L15, L16, L17"           ' SPML-PM
    AddSlots 25, "L19, L20"                     ' HALAL
    AddSlots 31, "L22, L23", "MON,TUE,WED,THU,SUN"   ' TIKTOK prep
    AddSlots 11, "L25, L26", "MON,TUE,WED,THU,SUN"   ' Admin-PM
    AddSlots 11, "L22, L23", "FRI,SAT"               ' Admin-PM moves up (no TIKTOK prep)

    ' Not used: 12 IDS, 29 F2, 30 FA (0 people), 35 CRA EY (uses VEGE for now)
    ' L29:L32 (TIKTOK / TT Wash) are filled by hand
End Sub

' Task order: tasks earlier in the list pick staff first
Private Function AmSequence() As Variant
    AmSequence = Array(26, 27, 28, 19, 18, 17, 14, 15, 16, 4, 6, 7, 8, 3, 5, 2, 1, 9, 20, 10, 13, 33)
End Function

Private Function PmSequence() As Variant
    PmSequence = Array(11, 21, 22, 25, 24, 23, 31)
End Function


' =====================================================================
' BUTTONS
' =====================================================================
Public Sub GenerateRoster()
    Dim ws As Worksheet
    Dim dayIdx As Long, dayCode As String, dayCol As Long

    Set ws = ActiveSheet
    dayIdx = DayIndex(ws.Name)
    If dayIdx < 0 Then Exit Sub
    dayCode = Choose(dayIdx + 1, "MON", "TUE", "WED", "THU", "FRI", "SAT", "SUN")
    dayCol = 3 + dayIdx                          ' MON = column C ... SUN = column I

    BuildSlotList dayCode
    DropSlotsBelowEnd ws

    ' 1. Read who works today
    Dim amPool As Object, pmPool As Object
    Dim alList As Collection, rdoList As Collection, wcList As Collection
    Set amPool = CreateObject("Scripting.Dictionary")
    Set pmPool = CreateObject("Scripting.Dictionary")
    Set alList = New Collection
    Set rdoList = New Collection
    Set wcList = New Collection

    Dim wsRoster As Worksheet, wsBCR As Worksheet
    Set wsRoster = ThisWorkbook.Sheets(ROSTER_SHEET)
    Set wsBCR = ThisWorkbook.Sheets(BCR_SHEET)
    ReadStaff wsRoster, ROSTER_START_ROW, dayCol, COLOR_ROSTER, True, _
              amPool, pmPool, alList, rdoList, wcList
    ReadStaff wsBCR, BCR_START_ROW, dayCol, COLOR_BCR, False, _
              amPool, pmPool, alList, rdoList, wcList

    ' 2. Read tasks and skills from the database
    Dim skills As Object, required As Object
    If Not LoadTaskData(skills, required) Then Exit Sub

    ' 3. Assign (nothing is written yet)
    Dim amResult As Object, pmResult As Object
    Dim amUn As Collection, pmUn As Collection
    Set amResult = AssignShift(AmSequence(), amPool, skills, required, dayCode, amUn)
    Set pmResult = AssignShift(PmSequence(), pmPool, skills, required, dayCode, pmUn)

    Application.ScreenUpdating = False

    ' 4. Clear the old allocation on this sheet
    ClearAllocation ws

    ' 5. Make room for Unassigned: insert rows above END if needed
    '    (END not found -> endRow = 0: no inserting, no leave writing)
    Dim endRow As Long, extra As Long
    endRow = FindEndRow(ws)
    extra = 0
    If endRow > 0 Then
        extra = Application.WorksheetFunction.Max( _
                    amUn.Count - (endRow - AM_UN_FIRST_ROW), _
                    pmUn.Count - (endRow - PM_UN_FIRST_ROW), 0)
    End If
    If extra > 0 Then
        ' Insert above the last row before END (the row with the thick
        ' bottom border), so new rows copy the plain row above it and the
        ' border row stays at the bottom of the block.
        ws.Rows(endRow - 1).Resize(extra).Insert Shift:=xlDown, _
                                                CopyOrigin:=xlFormatFromLeftOrAbove
        endRow = endRow + extra                  ' END has moved down
    End If

    ' 6. Write tasks and Unassigned
    WriteTasks ws, amResult, amPool
    WriteTasks ws, pmResult, pmPool
    WriteUnassigned ws, AM_UN_TITLE, AM_UN_FIRST_ROW, AM_UN_COL, amUn, amPool, (endRow > 0)
    WriteUnassigned ws, PM_UN_TITLE, PM_UN_FIRST_ROW, PM_UN_COL, pmUn, pmPool, (endRow > 0)

    ' 7. Write leave, below the (new) END row. Skipped if END is missing,
    '    so nothing gets written on top of the hours summary.
    If endRow > 0 Then
        Dim leaveList As Collection, item As Variant
        Set leaveList = New Collection
        For Each item In alList: leaveList.Add item: Next item
        For Each item In rdoList: leaveList.Add item: Next item
        WriteLeave ws, endRow + 1, AL_COL, leaveList
        WriteLeave ws, endRow + 1, WC_COL, wcList
    End If

    Application.ScreenUpdating = True
End Sub

Public Sub ResetRoster()
    Dim ws As Worksheet
    Set ws = ActiveSheet
    If DayIndex(ws.Name) < 0 Then Exit Sub
    If MsgBox("Are you sure you want to reset the sheet?", _
              vbQuestion + vbYesNo, "Reset Confirmation") = vbNo Then Exit Sub

    BuildSlotList Choose(DayIndex(ws.Name) + 1, "MON", "TUE", "WED", "THU", "FRI", "SAT", "SUN")
    DropSlotsBelowEnd ws
    Application.ScreenUpdating = False
    ClearAllocation ws
    Application.ScreenUpdating = True
End Sub


' =====================================================================
' READING
' =====================================================================

' Reads one staff sheet from startRow down to the END row in column A.
' readLeave = True for roster (AL / RDO / WCOMP / SL), False for BCR.
Private Sub ReadStaff(ws As Worksheet, ByVal startRow As Long, ByVal dayCol As Long, _
                      ByVal srcColor As Long, ByVal readLeave As Boolean, _
                      amPool As Object, pmPool As Object, _
                      alList As Collection, rdoList As Collection, wcList As Collection)
    Dim r As Long, lastRow As Long
    Dim id As String, nm As String, t As String
    Dim h As Long

    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row

    For r = startRow To lastRow
        If UCase$(CellText(ws.Cells(r, "A"))) = "END" Then Exit For

        id = CleanID(ws.Cells(r, "A").Value)
        If Len(id) > 0 Then
            If Not amPool.Exists(id) And Not pmPool.Exists(id) Then
                nm = CellText(ws.Cells(r, "B"))
                t = UCase$(CellText(ws.Cells(r, dayCol)))

                Select Case t
                    Case "", "NA"
                        ' not working today
                    Case "AL"
                        If readLeave Then alList.Add Array(id, nm, "AL")
                    Case "RDO"
                        If readLeave Then rdoList.Add Array(id, nm, "RDO")
                    Case "WCOMP", "SL"
                        If readLeave Then wcList.Add Array(id, nm, t)
                    Case Else
                        h = StartHour(t)
                        If h >= 0 Then
                            If h < 12 Then
                                amPool.Add id, Array(nm, t, srcColor)
                            Else
                                pmPool.Add id, Array(nm, t, srcColor)
                            End If
                        End If
                End Select
            End If
        End If
    Next r
End Sub

' Reads RequiredPeople per task and skilled staff per task (highest score first)
Private Function LoadTaskData(skills As Object, required As Object) As Boolean
    Dim conn As Object, rs As Object
    Dim tid As Long, sid As String

    Set skills = CreateObject("Scripting.Dictionary")
    Set required = CreateObject("Scripting.Dictionary")
    Set conn = CreateObject("ADODB.Connection")

    On Error Resume Next
    conn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & DB_PATH
    If Err.Number <> 0 Then
        On Error GoTo 0
        MsgBox "Cannot open the roster database:" & vbCrLf & DB_PATH, vbExclamation
        Exit Function
    End If
    On Error GoTo 0

    Set rs = conn.Execute("SELECT TaskID, RequiredPeople FROM Tasks")
    Do While Not rs.EOF
        If Not IsNull(rs.Fields("TaskID").Value) Then
            required(CLng(rs.Fields("TaskID").Value)) = NzLong(rs.Fields("RequiredPeople").Value)
        End If
        rs.MoveNext
    Loop
    rs.Close

    Set rs = conn.Execute("SELECT TaskID, StaffID FROM peopleSkills ORDER BY TaskID, SkillScore DESC")
    Do While Not rs.EOF
        If Not IsNull(rs.Fields("TaskID").Value) Then
            tid = CLng(rs.Fields("TaskID").Value)
            sid = CleanID(rs.Fields("StaffID").Value)
            If Len(sid) > 0 Then
                If Not skills.Exists(tid) Then skills.Add tid, New Collection
                skills(tid).Add sid
            End If
        End If
        rs.MoveNext
    Loop
    rs.Close
    conn.Close

    LoadTaskData = True
End Function


' =====================================================================
' ASSIGNING
' =====================================================================

' For each task in order: pick up to min(RequiredPeople, number of cells)
' people from the pool, best skill first. A person is used once only.
' People not picked by any task end up in "unassigned".
Private Function AssignShift(seq As Variant, pool As Object, skills As Object, _
                             required As Object, ByVal dayCode As String, _
                             unassigned As Collection) As Object
    Dim result As Object, used As Object
    Dim t As Variant, taskID As Long, need As Long
    Dim picked As Collection, sid As Variant

    Set result = CreateObject("Scripting.Dictionary")
    Set used = CreateObject("Scripting.Dictionary")

    For Each t In seq
        taskID = CLng(t)
        need = 0
        If gSlots.Exists(taskID) And required.Exists(taskID) Then
            need = required(taskID)
            If need > gSlots(taskID).Count Then need = gSlots(taskID).Count
        End If

        If need > 0 And skills.Exists(taskID) Then
            Set picked = New Collection
            For Each sid In skills(taskID)
                If pool.Exists(sid) And Not used.Exists(sid) Then
                    picked.Add sid
                    used(sid) = True
                    If picked.Count >= need Then Exit For
                End If
            Next sid
            If picked.Count > 0 Then Set result(taskID) = picked
        End If
    Next t

    Set unassigned = New Collection
    For Each sid In pool.Keys
        If Not used.Exists(sid) Then unassigned.Add sid
    Next sid

    Set AssignShift = result
End Function


' =====================================================================
' WRITING
' =====================================================================
Private Sub WriteTasks(ws As Worksheet, result As Object, pool As Object)
    Dim t As Variant, i As Long, slots As Collection, picked As Collection

    For Each t In result.Keys
        Set slots = gSlots(t)
        Set picked = result(t)
        For i = 1 To picked.Count
            WritePerson ws.Range(slots(i)), picked(i), pool(picked(i))
        Next i
    Next t
End Sub

Private Sub WriteUnassigned(ws As Worksheet, ByVal titleAddr As String, ByVal firstRow As Long, _
                           ByVal idCol As String, list As Collection, pool As Object, _
                           ByVal endFound As Boolean)
    Dim i As Long
    If list.Count = 0 Then Exit Sub

    With ws.Range(titleAddr)
        If endFound Then
            .Value = "Unassigned Staff"
        Else
            .Value = "Unassigned Staff (END not found)"
        End If
        .Interior.Color = COLOR_RED
    End With

    For i = 1 To list.Count
        With ws.Range(idCol & (firstRow + i - 1))
            WritePerson .Cells(1, 1), list(i), pool(list(i))
            .Resize(1, 3).Font.Color = vbBlack   ' last row above END has white font
        End With
    Next i
End Sub

' data = Array(name, time, colour)
Private Sub WritePerson(idCell As Range, ByVal staffID As String, data As Variant)
    idCell.NumberFormat = "0"                    ' long IDs, no 1.21E+08
    idCell.Value = staffID
    idCell.Interior.Color = data(2)
    idCell.Offset(0, 1).Value = data(0)
    idCell.Offset(0, 2).Value = data(1)
End Sub

' Writes leave rows from startRow down. Stops after LEAVE_MAX_ROWS rows,
' or at the first row that already has something in it (summary labels).
' Colours are not changed in the leave area.
Private Sub WriteLeave(ws As Worksheet, ByVal startRow As Long, ByVal idCol As String, list As Collection)
    Dim i As Long, r As Long, idCell As Range, rec As Variant

    r = startRow
    For i = 1 To list.Count
        If r >= startRow + LEAVE_MAX_ROWS Then Exit For
        Set idCell = ws.Range(idCol & r)
        If Application.WorksheetFunction.CountA(idCell.Resize(1, 3)) > 0 Then Exit For

        rec = list(i)                            ' Array(id, name, status)
        idCell.Value = rec(0)
        idCell.Offset(0, 1).Value = rec(1)
        idCell.Offset(0, 2).Value = rec(2)
        r = r + 1
    Next i
End Sub


' =====================================================================
' CLEARING
' =====================================================================
Private Sub ClearAllocation(ws As Worksheet)
    Dim t As Variant, addr As Variant, endRow As Long

    ' Task cells
    For Each t In gSlots.Keys
        For Each addr In gSlots(t)
            ClearPersonCell ws.Range(addr)
        Next addr
    Next t

    endRow = FindEndRow(ws)

    ' Unassigned areas (rows between first row and END)
    ClearUnassigned ws, AM_UN_TITLE, AM_UN_FIRST_ROW, AM_UN_COL, endRow
    ClearUnassigned ws, PM_UN_TITLE, PM_UN_FIRST_ROW, PM_UN_COL, endRow

    ' Leave areas: clear only rows we wrote (ID cell holds a staff ID)
    If endRow > 0 Then
        ClearLeave ws, endRow + 1, AL_COL
        ClearLeave ws, endRow + 1, WC_COL
    End If
End Sub

Private Sub ClearPersonCell(idCell As Range)
    idCell.ClearContents
    idCell.Interior.Color = COLOR_WHITE
    idCell.Offset(0, 1).Resize(1, 2).ClearContents
End Sub

Private Sub ClearUnassigned(ws As Worksheet, ByVal titleAddr As String, ByVal firstRow As Long, _
                           ByVal idCol As String, ByVal endRow As Long)
    Dim r As Long

    With ws.Range(titleAddr)
        .ClearContents
        .Interior.Color = COLOR_WHITE
    End With

    If endRow > 0 Then
        For r = firstRow To endRow - 1
            ClearPersonCell ws.Range(idCol & r)
        Next r
    Else
        ' END missing: clear only the unbroken list of IDs we wrote
        r = firstRow
        Do While Len(CleanID(ws.Range(idCol & r).Value)) > 0
            ClearPersonCell ws.Range(idCol & r)
            r = r + 1
        Loop
    End If
End Sub

Private Sub ClearLeave(ws As Worksheet, ByVal startRow As Long, ByVal idCol As String)
    Dim r As Long
    For r = startRow To startRow + LEAVE_MAX_ROWS - 1
        If Len(CleanID(ws.Range(idCol & r).Value)) = 0 Then Exit For
        ws.Range(idCol & r).Resize(1, 3).ClearContents
    Next r
End Sub


' =====================================================================
' HELPERS
' =====================================================================

' Row of the cell in column B whose text is "END" (any case, extra spaces
' ignored). Returns 0 if there is no END. Never guesses a row.
Private Function FindEndRow(ws As Worksheet) As Long
    Dim r As Long, lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, END_COLUMN).End(xlUp).Row
    For r = 1 To lastRow
        If UCase$(CellText(ws.Cells(r, END_COLUMN))) = "END" Then
            FindEndRow = r
            Exit Function
        End If
    Next r
    FindEndRow = 0
End Function

Private Function DayIndex(ByVal sheetName As String) As Long
    Select Case LCase$(Trim$(sheetName))
        Case "monday":    DayIndex = 0
        Case "tuesday":   DayIndex = 1
        Case "wednesday": DayIndex = 2
        Case "thursday":  DayIndex = 3
        Case "friday":    DayIndex = 4
        Case "saturday":  DayIndex = 5
        Case "sunday":    DayIndex = 6
        Case Else:        DayIndex = -1
    End Select
End Function

' Safety: a slot cell on or below the END row is never cleared or written.
' (If rows above END are deleted, a slot like B36 could land on END itself.)
' Tasks that lose cells this way simply get fewer people; the rest go to Unassigned.
Private Sub DropSlotsBelowEnd(ws As Worksheet)
    Dim endRow As Long, t As Variant, addr As Variant, kept As Collection
    endRow = FindEndRow(ws)
    If endRow = 0 Then Exit Sub

    For Each t In gSlots.Keys
        Set kept = New Collection
        For Each addr In gSlots(t)
            If ws.Range(addr).Row < endRow Then kept.Add addr
        Next addr
        Set gSlots(t) = kept
    Next t
End Sub

Private Sub AddSlots(ByVal taskID As Long, ByVal cellList As String, Optional ByVal days As String = "")
    Dim col As Collection, part As Variant

    ' Line only applies on the listed days
    If Len(days) > 0 Then
        If InStr(1, "," & UCase$(Replace(days, " ", "")) & ",", "," & gToday & ",") = 0 Then Exit Sub
    End If

    Set col = New Collection
    For Each part In Split(cellList, ",")
        If Len(Trim$(part)) > 0 Then col.Add UCase$(Trim$(part))
    Next part
    Set gSlots(taskID) = col
End Sub

' Staff ID as text: digits only, hidden spaces removed. "" if not a valid ID.
Private Function CleanID(ByVal v As Variant) As String
    Dim s As String
    If IsError(v) Or IsNull(v) Or IsEmpty(v) Then Exit Function
    s = CStr(v)
    s = Replace(s, Chr(160), "")
    s = Trim$(s)
    If Len(s) = 0 Then Exit Function
    If s Like "*[!0-9]*" Then Exit Function
    If Val(s) = 0 Then Exit Function
    CleanID = s
End Function

' Start hour of "06:00-14:00" (also "6:00-14:00" and "0600-1400"); -1 if not a shift
Private Function StartHour(ByVal t As String) As Long
    Dim p As Long, s As String
    StartHour = -1
    p = InStr(t, "-")
    If p = 0 Then Exit Function
    s = Trim$(Left$(t, p - 1))
    If InStr(s, ":") > 0 Then
        s = Left$(s, InStr(s, ":") - 1)
    ElseIf Len(s) >= 3 Then
        s = Left$(s, Len(s) - 2)
    End If
    If Len(s) = 0 Or s Like "*[!0-9]*" Then Exit Function
    StartHour = CLng(s)
End Function

Private Function CellText(c As Range) As String
    If IsError(c.Value) Then Exit Function
    CellText = Trim$(Replace(CStr(c.Value), Chr(160), " "))
End Function

Private Function NzLong(ByVal v As Variant) As Long
    If IsNull(v) Or IsEmpty(v) Then
        NzLong = 0
    ElseIf IsNumeric(v) Then
        NzLong = CLng(v)
    End If
End Function
