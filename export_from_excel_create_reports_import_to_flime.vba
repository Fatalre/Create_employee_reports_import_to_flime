Private Sub CommandButton1_Click()
Workbooks("Create_employee_reports.xlsm").Activate
'Âûáîð ôàéëîâ è çàïèñü èõ â ÿ÷åéêó
Filename = Application _
    .GetOpenFilename("Excel Files (*.xlsx), *.xlsx", 1, "Âûáåðèòå ôàéë", "Âûáðàòü", True)
Dim element As Variant
Dim i As Integer, y As Integer
i = 1
Dim l As Integer
l = 1
For l = 1 To 1000
Cells(l, 1) = ""
Next l
For Each element In Filename
Cells(i, 1).Value = element
i = i + 1
Next

'Çàïèñü ïóòåé â ñòðîêîâûé ìàññèâ
Dim path(100) As String
i = 1
For y = 0 To 100
path(y) = Cells(i, 1).Value
i = i + 1
Next y

Dim m As Integer
Dim ch As Integer

'Îòêðûòèå ôàéëîâ è äåéñòâèé íàä íèìè, ïðîâåðêà âñåõ ëèñòîâ
i = 0
ch = 0
Dim Data(1000) As String, Dlitelnost(1000) As String, Stoimost(1000) As String, Zadacha(1000) As String, proekt(1000) As String, OpisZadach(1000) As String, Ssulka(1000) As String, NameUsser(1000) As String
Do While path(i) <> ""
Workbooks.Open Filename:=path(i)
For m = 1 To ActiveWorkbook.Worksheets.Count
ActiveWorkbook.Worksheets(m).Activate
If ActiveWorkbook.Worksheets(m).Name <> "info" Then
'Çàïèñü çíà÷åíèé â ïåðåìåííûå
y = 2
Do While ActiveWorkbook.ActiveSheet.Cells(y, 1).Text <> ""
Data(ch) = ActiveWorkbook.ActiveSheet.Cells(y, 1).Value
Dlitelnost(ch) = ActiveWorkbook.ActiveSheet.Cells(y, 2).Text
Stoimost(ch) = ActiveWorkbook.ActiveSheet.Cells(y, 3).Text
Zadacha(ch) = ActiveWorkbook.ActiveSheet.Cells(y, 4).Text
proekt(ch) = ActiveWorkbook.ActiveSheet.Cells(y, 5).Text
OpisZadach(ch) = ActiveWorkbook.ActiveSheet.Cells(y, 6).Text
Ssulka(ch) = ActiveWorkbook.ActiveSheet.Cells(y, 7).Text
NameUsser(ch) = ActiveWorkbook.ActiveSheet.Cells(1, 8).Text
ch = ch + 1
y = y + 1
Loop
End If
Next m


'Çàìåíà òî÷åê è äâîåòî÷èå â íàçâàíèå ïðîåêòà
Dim n As Integer
Dim findstr As String, newstr As String, findstr1 As String, newstr1 As String
For n = 0 To 1000
findstr = ":"
newstr = " "
proekt(n) = Replace(proekt(n), findstr, newstr)
findstr = "."
newstr = " "
proekt(n) = Replace(proekt(n), findstr1, newstr1)
Next n
ActiveWorkbook.Close
i = i + 1
Loop
    Columns("B:B").Select
    Selection.ClearContents
'Ïåðåèìåíîâàíèå ïóñòûõ íàçâàíèé ïðîåêòà â òî÷êè
Dim namefile(1000) As String
Workbooks("Create_employee_reports.xlsm").Activate
x = 0
For y = 1 To 1000
If proekt(x) = "" Then
proekt(x) = "."
End If
Cells(y, 2) = proekt(x)
x = x + 1
Next y
'Ñîðòèðîâêà ïî àëôàâèòó äëÿ âû÷èñëåíèé äóáëèêàòîâ
    Columns("B:B").Select
    ActiveWorkbook.Worksheets("Ëèñò1").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("Ëèñò1").Sort.SortFields.Add Key:=Range("B1"), _
        SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("Ëèñò1").Sort
        .SetRange Range("B1:B1000")
        .Header = xlNo
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
'Óäàëåíèå äóáëèêàòîâ
    For x = 1 To 1000
    If Cells(x + 1, 2).Text = Cells(x, 2).Text And Cells(x + 1, 2) <> "" Then
    Rows(x + 1 & ":" & x + 1).Select
    Selection.Delete Shift:=xlUp
    x = x - 1
    End If
    Next x
'Ñîçäàíèå íîâîãî ìàññèâà ñ èìåíàìè ïðîåêòà

Dim OneName(1000) As String
Dim z As Integer
For z = 0 To 1000
OneName(z) = ""
Next z

x = 1
y = 0
Workbooks("Create_employee_reports.xlsm").Activate
Do While (Cells(x, 2)) <> ""
Cells(x, 2).Select
OneName(y) = ActiveWorkbook.ActiveSheet.Cells(x, 2).Text
'MsgBox (OneName(y))
y = y + 1
x = x + 1
Loop






'Ñîçäàíèå êíèã ïî ïðîåêòàì, îòêðûòèå êíèã, çàïèñü çíà÷åíèé
Dim temp As String
Dim p As Integer
Dim lkl As Integer

For y = 0 To 1000
If proekt(y) = "." Then
temp = Application.ActiveWorkbook.path & "\No project.xlsx"
    If Dir(temp) <> "" Then
        If bBookOpen("No project.xlsx") Then
'            MsgBox ("Êíèãà îòêðûòà")
            Workbooks("No project.xlsx").Activate
            For i = 1 To ActiveWorkbook.Worksheets.Count
            Dim mesac As String
            mesac = Right(Left(Data(y), 5), 2)
'            MsgBox (mesac)
            If ActiveWorkbook.Worksheets(i).Name = mesac Then
            ActiveWorkbook.Sheets(ActiveWorkbook.Worksheets(i).Name).Select
            p = 1
            Do While (ActiveWorkbook.ActiveSheet.Cells(p, 1)) <> ""
            For lkl = 0 To 1000
            If (Data(lkl) = ActiveWorkbook.ActiveSheet.Cells(p, 1).Value And Dlitelnost(lkl) = ActiveWorkbook.ActiveSheet.Cells(p, 3).Text And OpisZadach(lkl) = ActiveWorkbook.ActiveSheet.Cells(p, 6).Text And Ssulka(lkl) = ActiveWorkbook.ActiveSheet.Cells(p, 7).Text) Then
            Data(lkl) = ""
            Dlitelnost(lkl) = ""
            Stoimost(lkl) = ""
            Zadacha(lkl) = ""
            proekt(lkl) = ""
            OpisZadach(lkl) = ""
            Ssulka(lkl) = ""
            NameUsser(lkl) = ""
            End If
            Next lkl
            p = p + 1
'            MsgBox (p)
            Loop
            If (Data(y) <> "") Then
            ActiveWorkbook.ActiveSheet.Cells(p, 1).Value = Data(y)
            ActiveWorkbook.ActiveSheet.Cells(p, 2) = NameUsser(y)
            ActiveWorkbook.ActiveSheet.Cells(p, 3) = Dlitelnost(y)
            ActiveWorkbook.ActiveSheet.Cells(p, 4) = Stoimost(y)
            ActiveWorkbook.ActiveSheet.Cells(p, 5) = Zadacha(y)
            ActiveWorkbook.ActiveSheet.Cells(p, 6) = OpisZadach(y)
            ActiveWorkbook.ActiveSheet.Cells(p, 7) = Ssulka(y)
            End If
            Else
            End If
            Next i
        Else
'            MsgBox ("Êíèãà íå îòêðûòà")
            Workbooks.Open (temp)
            Workbooks("No project.xlsx").Activate
            For i = 1 To ActiveWorkbook.Worksheets.Count
            mesac = Right(Left(Data(y), 5), 2)
'            MsgBox (mesac)
            If ActiveWorkbook.Worksheets(i).Name = mesac Then
            ActiveWorkbook.Sheets(ActiveWorkbook.Worksheets(i).Name).Select
            p = 1
            Do While (ActiveWorkbook.ActiveSheet.Cells(p, 1)) <> ""
            For lkl = 0 To 1000
            If (Data(lkl) = ActiveWorkbook.ActiveSheet.Cells(p, 1).Value And Dlitelnost(lkl) = ActiveWorkbook.ActiveSheet.Cells(p, 3).Text And OpisZadach(lkl) = ActiveWorkbook.ActiveSheet.Cells(p, 6).Text And Ssulka(lkl) = ActiveWorkbook.ActiveSheet.Cells(p, 7).Text) Then
            Data(lkl) = ""
            Dlitelnost(lkl) = ""
            Stoimost(lkl) = ""
            Zadacha(lkl) = ""
            proekt(lkl) = ""
            OpisZadach(lkl) = ""
            Ssulka(lkl) = ""
            NameUsser(lkl) = ""
            End If
            Next lkl
            p = p + 1
'            MsgBox (p)
            Loop
            If (Data(y) <> "") Then
            ActiveWorkbook.ActiveSheet.Cells(p, 1).Value = Data(y)
            ActiveWorkbook.ActiveSheet.Cells(p, 2) = NameUsser(y)
            ActiveWorkbook.ActiveSheet.Cells(p, 3) = Dlitelnost(y)
            ActiveWorkbook.ActiveSheet.Cells(p, 4) = Stoimost(y)
            ActiveWorkbook.ActiveSheet.Cells(p, 5) = Zadacha(y)
            ActiveWorkbook.ActiveSheet.Cells(p, 6) = OpisZadach(y)
            ActiveWorkbook.ActiveSheet.Cells(p, 7) = Ssulka(y)
            End If
            End If
            Next i
        End If
        
    Else
    
        Workbooks.Add
        ActiveWorkbook.SaveAs Filename:=temp
        Workbooks("No project.xlsx").Activate
        
        
        Sheets("Ëèñò1").Select
        Sheets("Ëèñò1").Name = "01"
                ActiveWorkbook.ActiveSheet.Cells(1, 1) = "Äàòà"
        ActiveWorkbook.ActiveSheet.Cells(1, 2) = "ÔÈÎ"
        ActiveWorkbook.ActiveSheet.Cells(1, 3) = "×àñû"
        ActiveWorkbook.ActiveSheet.Cells(1, 4) = "Ñòîèìîñòü"
        ActiveWorkbook.ActiveSheet.Cells(1, 5) = "Ãðóïïà çàäà÷"
        ActiveWorkbook.ActiveSheet.Cells(1, 6) = "Îïèñàíèå"
        ActiveWorkbook.ActiveSheet.Cells(1, 7) = "Ññûëêà"
                Sheets.Add After:=ActiveSheet
        Sheets("Ëèñò2").Select
        Sheets("Ëèñò2").Name = "02"
                ActiveWorkbook.ActiveSheet.Cells(1, 1) = "Äàòà"
        ActiveWorkbook.ActiveSheet.Cells(1, 2) = "ÔÈÎ"
        ActiveWorkbook.ActiveSheet.Cells(1, 3) = "×àñû"
        ActiveWorkbook.ActiveSheet.Cells(1, 4) = "Ñòîèìîñòü"
        ActiveWorkbook.ActiveSheet.Cells(1, 5) = "Ãðóïïà çàäà÷"
        ActiveWorkbook.ActiveSheet.Cells(1, 6) = "Îïèñàíèå"
        ActiveWorkbook.ActiveSheet.Cells(1, 7) = "Ññûëêà"
                Sheets.Add After:=ActiveSheet
        Sheets("Ëèñò3").Select
        Sheets("Ëèñò3").Name = "03"
                ActiveWorkbook.ActiveSheet.Cells(1, 1) = "Äàòà"
        ActiveWorkbook.ActiveSheet.Cells(1, 2) = "ÔÈÎ"
        ActiveWorkbook.ActiveSheet.Cells(1, 3) = "×àñû"
        ActiveWorkbook.ActiveSheet.Cells(1, 4) = "Ñòîèìîñòü"
        ActiveWorkbook.ActiveSheet.Cells(1, 5) = "Ãðóïïà çàäà÷"
        ActiveWorkbook.ActiveSheet.Cells(1, 6) = "Îïèñàíèå"
        ActiveWorkbook.ActiveSheet.Cells(1, 7) = "Ññûëêà"
                Sheets.Add After:=ActiveSheet
        Sheets("Ëèñò4").Select
        Sheets("Ëèñò4").Name = "04"
                ActiveWorkbook.ActiveSheet.Cells(1, 1) = "Äàòà"
        ActiveWorkbook.ActiveSheet.Cells(1, 2) = "ÔÈÎ"
        ActiveWorkbook.ActiveSheet.Cells(1, 3) = "×àñû"
        ActiveWorkbook.ActiveSheet.Cells(1, 4) = "Ñòîèìîñòü"
        ActiveWorkbook.ActiveSheet.Cells(1, 5) = "Ãðóïïà çàäà÷"
        ActiveWorkbook.ActiveSheet.Cells(1, 6) = "Îïèñàíèå"
        ActiveWorkbook.ActiveSheet.Cells(1, 7) = "Ññûëêà"
                Sheets.Add After:=ActiveSheet
        Sheets("Ëèñò5").Select
        Sheets("Ëèñò5").Name = "05"
                ActiveWorkbook.ActiveSheet.Cells(1, 1) = "Äàòà"
        ActiveWorkbook.ActiveSheet.Cells(1, 2) = "ÔÈÎ"
        ActiveWorkbook.ActiveSheet.Cells(1, 3) = "×àñû"
        ActiveWorkbook.ActiveSheet.Cells(1, 4) = "Ñòîèìîñòü"
        ActiveWorkbook.ActiveSheet.Cells(1, 5) = "Ãðóïïà çàäà÷"
        ActiveWorkbook.ActiveSheet.Cells(1, 6) = "Îïèñàíèå"
        ActiveWorkbook.ActiveSheet.Cells(1, 7) = "Ññûëêà"
                Sheets.Add After:=ActiveSheet
        Sheets("Ëèñò6").Select
        Sheets("Ëèñò6").Name = "06"
                ActiveWorkbook.ActiveSheet.Cells(1, 1) = "Äàòà"
        ActiveWorkbook.ActiveSheet.Cells(1, 2) = "ÔÈÎ"
        ActiveWorkbook.ActiveSheet.Cells(1, 3) = "×àñû"
        ActiveWorkbook.ActiveSheet.Cells(1, 4) = "Ñòîèìîñòü"
        ActiveWorkbook.ActiveSheet.Cells(1, 5) = "Ãðóïïà çàäà÷"
        ActiveWorkbook.ActiveSheet.Cells(1, 6) = "Îïèñàíèå"
        ActiveWorkbook.ActiveSheet.Cells(1, 7) = "Ññûëêà"
                Sheets.Add After:=ActiveSheet
        Sheets("Ëèñò7").Select
        Sheets("Ëèñò7").Name = "07"
                ActiveWorkbook.ActiveSheet.Cells(1, 1) = "Äàòà"
        ActiveWorkbook.ActiveSheet.Cells(1, 2) = "ÔÈÎ"
        ActiveWorkbook.ActiveSheet.Cells(1, 3) = "×àñû"
        ActiveWorkbook.ActiveSheet.Cells(1, 4) = "Ñòîèìîñòü"
        ActiveWorkbook.ActiveSheet.Cells(1, 5) = "Ãðóïïà çàäà÷"
        ActiveWorkbook.ActiveSheet.Cells(1, 6) = "Îïèñàíèå"
        ActiveWorkbook.ActiveSheet.Cells(1, 7) = "Ññûëêà"
                Sheets.Add After:=ActiveSheet
        Sheets("Ëèñò8").Select
        Sheets("Ëèñò8").Name = "08"
                ActiveWorkbook.ActiveSheet.Cells(1, 1) = "Äàòà"
        ActiveWorkbook.ActiveSheet.Cells(1, 2) = "ÔÈÎ"
        ActiveWorkbook.ActiveSheet.Cells(1, 3) = "×àñû"
        ActiveWorkbook.ActiveSheet.Cells(1, 4) = "Ñòîèìîñòü"
        ActiveWorkbook.ActiveSheet.Cells(1, 5) = "Ãðóïïà çàäà÷"
        ActiveWorkbook.ActiveSheet.Cells(1, 6) = "Îïèñàíèå"
        ActiveWorkbook.ActiveSheet.Cells(1, 7) = "Ññûëêà"
                Sheets.Add After:=ActiveSheet
        Sheets("Ëèñò9").Select
        Sheets("Ëèñò9").Name = "09"
                ActiveWorkbook.ActiveSheet.Cells(1, 1) = "Äàòà"
        ActiveWorkbook.ActiveSheet.Cells(1, 2) = "ÔÈÎ"
        ActiveWorkbook.ActiveSheet.Cells(1, 3) = "×àñû"
        ActiveWorkbook.ActiveSheet.Cells(1, 4) = "Ñòîèìîñòü"
        ActiveWorkbook.ActiveSheet.Cells(1, 5) = "Ãðóïïà çàäà÷"
        ActiveWorkbook.ActiveSheet.Cells(1, 6) = "Îïèñàíèå"
        ActiveWorkbook.ActiveSheet.Cells(1, 7) = "Ññûëêà"
                Sheets.Add After:=ActiveSheet
        Sheets("Ëèñò10").Select
        Sheets("Ëèñò10").Name = "10"
                ActiveWorkbook.ActiveSheet.Cells(1, 1) = "Äàòà"
        ActiveWorkbook.ActiveSheet.Cells(1, 2) = "ÔÈÎ"
        ActiveWorkbook.ActiveSheet.Cells(1, 3) = "×àñû"
        ActiveWorkbook.ActiveSheet.Cells(1, 4) = "Ñòîèìîñòü"
        ActiveWorkbook.ActiveSheet.Cells(1, 5) = "Ãðóïïà çàäà÷"
        ActiveWorkbook.ActiveSheet.Cells(1, 6) = "Îïèñàíèå"
        ActiveWorkbook.ActiveSheet.Cells(1, 7) = "Ññûëêà"
                Sheets.Add After:=ActiveSheet
        Sheets("Ëèñò11").Select
        Sheets("Ëèñò11").Name = "11"
                ActiveWorkbook.ActiveSheet.Cells(1, 1) = "Äàòà"
        ActiveWorkbook.ActiveSheet.Cells(1, 2) = "ÔÈÎ"
        ActiveWorkbook.ActiveSheet.Cells(1, 3) = "×àñû"
        ActiveWorkbook.ActiveSheet.Cells(1, 4) = "Ñòîèìîñòü"
        ActiveWorkbook.ActiveSheet.Cells(1, 5) = "Ãðóïïà çàäà÷"
        ActiveWorkbook.ActiveSheet.Cells(1, 6) = "Îïèñàíèå"
        ActiveWorkbook.ActiveSheet.Cells(1, 7) = "Ññûëêà"
                Sheets.Add After:=ActiveSheet
        Sheets("Ëèñò12").Select
        Sheets("Ëèñò12").Name = "12"
    
        ActiveWorkbook.ActiveSheet.Cells(1, 1) = "Äàòà"
        ActiveWorkbook.ActiveSheet.Cells(1, 2) = "ÔÈÎ"
        ActiveWorkbook.ActiveSheet.Cells(1, 3) = "×àñû"
        ActiveWorkbook.ActiveSheet.Cells(1, 4) = "Ñòîèìîñòü"
        ActiveWorkbook.ActiveSheet.Cells(1, 5) = "Ãðóïïà çàäà÷"
        ActiveWorkbook.ActiveSheet.Cells(1, 6) = "Îïèñàíèå"
        ActiveWorkbook.ActiveSheet.Cells(1, 7) = "Ññûëêà"
        
                    For i = 1 To ActiveWorkbook.Worksheets.Count
            mesac = Right(Left(Data(y), 5), 2)
'            MsgBox (mesac)
            If ActiveWorkbook.Worksheets(i).Name = mesac Then
            ActiveWorkbook.Sheets(ActiveWorkbook.Worksheets(i).Name).Select
        p = 1
        Do While (ActiveWorkbook.ActiveSheet.Cells(p, 1)) <> ""
        p = p + 1
'            MsgBox (p)
        Loop
        ActiveWorkbook.ActiveSheet.Cells(p, 1).Value = Data(y)
        ActiveWorkbook.ActiveSheet.Cells(p, 2) = NameUsser(y)
        ActiveWorkbook.ActiveSheet.Cells(p, 3) = Dlitelnost(y)
        ActiveWorkbook.ActiveSheet.Cells(p, 4) = Stoimost(y)
        ActiveWorkbook.ActiveSheet.Cells(p, 5) = Zadacha(y)
        ActiveWorkbook.ActiveSheet.Cells(p, 6) = OpisZadach(y)
        ActiveWorkbook.ActiveSheet.Cells(p, 7) = Ssulka(y)
        End If
        Next i
    End If
    
ElseIf proekt(y) = "" Then

Else

For x = 0 To 999
If proekt(y) = OneName(x) Then
temp = Application.ActiveWorkbook.path & "\" & OneName(x) & ".xlsx"
    If Dir(temp) <> "" Then
        If bBookOpen(OneName(x) & ".xlsx") Then
'            MsgBox ("Êíèãà îòêðûòà")
            Workbooks(OneName(x) & ".xlsx").Activate
            For i = 1 To ActiveWorkbook.Worksheets.Count
            mesac = Right(Left(Data(y), 5), 2)
'            MsgBox (mesac)
            If ActiveWorkbook.Worksheets(i).Name = mesac Then
            ActiveWorkbook.Sheets(ActiveWorkbook.Worksheets(i).Name).Select
            p = 1
            Do While (ActiveWorkbook.ActiveSheet.Cells(p, 1)) <> ""
            For lkl = 0 To 1000
            If (Data(lkl) = ActiveWorkbook.ActiveSheet.Cells(p, 1).Value And Dlitelnost(lkl) = ActiveWorkbook.ActiveSheet.Cells(p, 3).Text And OpisZadach(lkl) = ActiveWorkbook.ActiveSheet.Cells(p, 6).Text And Ssulka(lkl) = ActiveWorkbook.ActiveSheet.Cells(p, 7).Text) Then
            Data(lkl) = ""
            Dlitelnost(lkl) = ""
            Stoimost(lkl) = ""
            Zadacha(lkl) = ""
            proekt(lkl) = ""
            OpisZadach(lkl) = ""
            Ssulka(lkl) = ""
            NameUsser(lkl) = ""
            End If
            Next lkl
            p = p + 1
'            MsgBox (p)
            Loop
            If (Data(y) <> "") Then
            ActiveWorkbook.ActiveSheet.Cells(p, 1).Value = Data(y)
            ActiveWorkbook.ActiveSheet.Cells(p, 2) = NameUsser(y)
            ActiveWorkbook.ActiveSheet.Cells(p, 3) = Dlitelnost(y)
            ActiveWorkbook.ActiveSheet.Cells(p, 4) = Stoimost(y)
            ActiveWorkbook.ActiveSheet.Cells(p, 5) = Zadacha(y)
            ActiveWorkbook.ActiveSheet.Cells(p, 6) = OpisZadach(y)
            ActiveWorkbook.ActiveSheet.Cells(p, 7) = Ssulka(y)
            End If
            End If
            Next i
        Else
'            MsgBox ("Êíèãà íå îòêðûòà")
            Workbooks.Open (temp)
            Workbooks(OneName(x) & ".xlsx").Activate
            For i = 1 To ActiveWorkbook.Worksheets.Count
            mesac = Right(Left(Data(y), 5), 2)
'            MsgBox (mesac)
            If ActiveWorkbook.Worksheets(i).Name = mesac Then
            ActiveWorkbook.Sheets(ActiveWorkbook.Worksheets(i).Name).Select
            p = 1
            Do While (ActiveWorkbook.ActiveSheet.Cells(p, 1)) <> ""
            For lkl = 0 To 1000
            If (Data(lkl) = ActiveWorkbook.ActiveSheet.Cells(p, 1).Value And Dlitelnost(lkl) = ActiveWorkbook.ActiveSheet.Cells(p, 3).Text And OpisZadach(lkl) = ActiveWorkbook.ActiveSheet.Cells(p, 6).Text And Ssulka(lkl) = ActiveWorkbook.ActiveSheet.Cells(p, 7).Text) Then
            Data(lkl) = ""
            Dlitelnost(lkl) = ""
            Stoimost(lkl) = ""
            Zadacha(lkl) = ""
            proekt(lkl) = ""
            OpisZadach(lkl) = ""
            Ssulka(lkl) = ""
            NameUsser(lkl) = ""
            End If
            Next lkl
            p = p + 1
'            MsgBox (p)
            Loop
            If (Data(y) <> "") Then
            ActiveWorkbook.ActiveSheet.Cells(p, 1).Value = Data(y)
            ActiveWorkbook.ActiveSheet.Cells(p, 2) = NameUsser(y)
            ActiveWorkbook.ActiveSheet.Cells(p, 3) = Dlitelnost(y)
            ActiveWorkbook.ActiveSheet.Cells(p, 4) = Stoimost(y)
            ActiveWorkbook.ActiveSheet.Cells(p, 5) = Zadacha(y)
            ActiveWorkbook.ActiveSheet.Cells(p, 6) = OpisZadach(y)
            ActiveWorkbook.ActiveSheet.Cells(p, 7) = Ssulka(y)
            End If
            End If
            Next i
        End If
        
    Else
    
        Workbooks.Add
        ActiveWorkbook.SaveAs Filename:=temp
        Workbooks(OneName(x) & ".xlsx").Activate
        
        Sheets("Ëèñò1").Select
        Sheets("Ëèñò1").Name = "01"
                ActiveWorkbook.ActiveSheet.Cells(1, 1) = "Äàòà"
        ActiveWorkbook.ActiveSheet.Cells(1, 2) = "ÔÈÎ"
        ActiveWorkbook.ActiveSheet.Cells(1, 3) = "×àñû"
        ActiveWorkbook.ActiveSheet.Cells(1, 4) = "Ñòîèìîñòü"
        ActiveWorkbook.ActiveSheet.Cells(1, 5) = "Ãðóïïà çàäà÷"
        ActiveWorkbook.ActiveSheet.Cells(1, 6) = "Îïèñàíèå"
        ActiveWorkbook.ActiveSheet.Cells(1, 7) = "Ññûëêà"
                Sheets.Add After:=ActiveSheet
        Sheets("Ëèñò2").Select
        Sheets("Ëèñò2").Name = "02"
                ActiveWorkbook.ActiveSheet.Cells(1, 1) = "Äàòà"
        ActiveWorkbook.ActiveSheet.Cells(1, 2) = "ÔÈÎ"
        ActiveWorkbook.ActiveSheet.Cells(1, 3) = "×àñû"
        ActiveWorkbook.ActiveSheet.Cells(1, 4) = "Ñòîèìîñòü"
        ActiveWorkbook.ActiveSheet.Cells(1, 5) = "Ãðóïïà çàäà÷"
        ActiveWorkbook.ActiveSheet.Cells(1, 6) = "Îïèñàíèå"
        ActiveWorkbook.ActiveSheet.Cells(1, 7) = "Ññûëêà"
                Sheets.Add After:=ActiveSheet
        Sheets("Ëèñò3").Select
        Sheets("Ëèñò3").Name = "03"
                ActiveWorkbook.ActiveSheet.Cells(1, 1) = "Äàòà"
        ActiveWorkbook.ActiveSheet.Cells(1, 2) = "ÔÈÎ"
        ActiveWorkbook.ActiveSheet.Cells(1, 3) = "×àñû"
        ActiveWorkbook.ActiveSheet.Cells(1, 4) = "Ñòîèìîñòü"
        ActiveWorkbook.ActiveSheet.Cells(1, 5) = "Ãðóïïà çàäà÷"
        ActiveWorkbook.ActiveSheet.Cells(1, 6) = "Îïèñàíèå"
        ActiveWorkbook.ActiveSheet.Cells(1, 7) = "Ññûëêà"
                Sheets.Add After:=ActiveSheet
        Sheets("Ëèñò4").Select
        Sheets("Ëèñò4").Name = "04"
                ActiveWorkbook.ActiveSheet.Cells(1, 1) = "Äàòà"
        ActiveWorkbook.ActiveSheet.Cells(1, 2) = "ÔÈÎ"
        ActiveWorkbook.ActiveSheet.Cells(1, 3) = "×àñû"
        ActiveWorkbook.ActiveSheet.Cells(1, 4) = "Ñòîèìîñòü"
        ActiveWorkbook.ActiveSheet.Cells(1, 5) = "Ãðóïïà çàäà÷"
        ActiveWorkbook.ActiveSheet.Cells(1, 6) = "Îïèñàíèå"
        ActiveWorkbook.ActiveSheet.Cells(1, 7) = "Ññûëêà"
                Sheets.Add After:=ActiveSheet
        Sheets("Ëèñò5").Select
        Sheets("Ëèñò5").Name = "05"
                ActiveWorkbook.ActiveSheet.Cells(1, 1) = "Äàòà"
        ActiveWorkbook.ActiveSheet.Cells(1, 2) = "ÔÈÎ"
        ActiveWorkbook.ActiveSheet.Cells(1, 3) = "×àñû"
        ActiveWorkbook.ActiveSheet.Cells(1, 4) = "Ñòîèìîñòü"
        ActiveWorkbook.ActiveSheet.Cells(1, 5) = "Ãðóïïà çàäà÷"
        ActiveWorkbook.ActiveSheet.Cells(1, 6) = "Îïèñàíèå"
        ActiveWorkbook.ActiveSheet.Cells(1, 7) = "Ññûëêà"
                Sheets.Add After:=ActiveSheet
        Sheets("Ëèñò6").Select
        Sheets("Ëèñò6").Name = "06"
                ActiveWorkbook.ActiveSheet.Cells(1, 1) = "Äàòà"
        ActiveWorkbook.ActiveSheet.Cells(1, 2) = "ÔÈÎ"
        ActiveWorkbook.ActiveSheet.Cells(1, 3) = "×àñû"
        ActiveWorkbook.ActiveSheet.Cells(1, 4) = "Ñòîèìîñòü"
        ActiveWorkbook.ActiveSheet.Cells(1, 5) = "Ãðóïïà çàäà÷"
        ActiveWorkbook.ActiveSheet.Cells(1, 6) = "Îïèñàíèå"
        ActiveWorkbook.ActiveSheet.Cells(1, 7) = "Ññûëêà"
                Sheets.Add After:=ActiveSheet
        Sheets("Ëèñò7").Select
        Sheets("Ëèñò7").Name = "07"
                ActiveWorkbook.ActiveSheet.Cells(1, 1) = "Äàòà"
        ActiveWorkbook.ActiveSheet.Cells(1, 2) = "ÔÈÎ"
        ActiveWorkbook.ActiveSheet.Cells(1, 3) = "×àñû"
        ActiveWorkbook.ActiveSheet.Cells(1, 4) = "Ñòîèìîñòü"
        ActiveWorkbook.ActiveSheet.Cells(1, 5) = "Ãðóïïà çàäà÷"
        ActiveWorkbook.ActiveSheet.Cells(1, 6) = "Îïèñàíèå"
        ActiveWorkbook.ActiveSheet.Cells(1, 7) = "Ññûëêà"
                Sheets.Add After:=ActiveSheet
        Sheets("Ëèñò8").Select
        Sheets("Ëèñò8").Name = "08"
                ActiveWorkbook.ActiveSheet.Cells(1, 1) = "Äàòà"
        ActiveWorkbook.ActiveSheet.Cells(1, 2) = "ÔÈÎ"
        ActiveWorkbook.ActiveSheet.Cells(1, 3) = "×àñû"
        ActiveWorkbook.ActiveSheet.Cells(1, 4) = "Ñòîèìîñòü"
        ActiveWorkbook.ActiveSheet.Cells(1, 5) = "Ãðóïïà çàäà÷"
        ActiveWorkbook.ActiveSheet.Cells(1, 6) = "Îïèñàíèå"
        ActiveWorkbook.ActiveSheet.Cells(1, 7) = "Ññûëêà"
                Sheets.Add After:=ActiveSheet
        Sheets("Ëèñò9").Select
        Sheets("Ëèñò9").Name = "09"
                ActiveWorkbook.ActiveSheet.Cells(1, 1) = "Äàòà"
        ActiveWorkbook.ActiveSheet.Cells(1, 2) = "ÔÈÎ"
        ActiveWorkbook.ActiveSheet.Cells(1, 3) = "×àñû"
        ActiveWorkbook.ActiveSheet.Cells(1, 4) = "Ñòîèìîñòü"
        ActiveWorkbook.ActiveSheet.Cells(1, 5) = "Ãðóïïà çàäà÷"
        ActiveWorkbook.ActiveSheet.Cells(1, 6) = "Îïèñàíèå"
        ActiveWorkbook.ActiveSheet.Cells(1, 7) = "Ññûëêà"
                Sheets.Add After:=ActiveSheet
        Sheets("Ëèñò10").Select
        Sheets("Ëèñò10").Name = "10"
                ActiveWorkbook.ActiveSheet.Cells(1, 1) = "Äàòà"
        ActiveWorkbook.ActiveSheet.Cells(1, 2) = "ÔÈÎ"
        ActiveWorkbook.ActiveSheet.Cells(1, 3) = "×àñû"
        ActiveWorkbook.ActiveSheet.Cells(1, 4) = "Ñòîèìîñòü"
        ActiveWorkbook.ActiveSheet.Cells(1, 5) = "Ãðóïïà çàäà÷"
        ActiveWorkbook.ActiveSheet.Cells(1, 6) = "Îïèñàíèå"
        ActiveWorkbook.ActiveSheet.Cells(1, 7) = "Ññûëêà"
                Sheets.Add After:=ActiveSheet
        Sheets("Ëèñò11").Select
        Sheets("Ëèñò11").Name = "11"
                ActiveWorkbook.ActiveSheet.Cells(1, 1) = "Äàòà"
        ActiveWorkbook.ActiveSheet.Cells(1, 2) = "ÔÈÎ"
        ActiveWorkbook.ActiveSheet.Cells(1, 3) = "×àñû"
        ActiveWorkbook.ActiveSheet.Cells(1, 4) = "Ñòîèìîñòü"
        ActiveWorkbook.ActiveSheet.Cells(1, 5) = "Ãðóïïà çàäà÷"
        ActiveWorkbook.ActiveSheet.Cells(1, 6) = "Îïèñàíèå"
        ActiveWorkbook.ActiveSheet.Cells(1, 7) = "Ññûëêà"
                Sheets.Add After:=ActiveSheet
        Sheets("Ëèñò12").Select
        Sheets("Ëèñò12").Name = "12"
    
        ActiveWorkbook.ActiveSheet.Cells(1, 1) = "Äàòà"
        ActiveWorkbook.ActiveSheet.Cells(1, 2) = "ÔÈÎ"
        ActiveWorkbook.ActiveSheet.Cells(1, 3) = "×àñû"
        ActiveWorkbook.ActiveSheet.Cells(1, 4) = "Ñòîèìîñòü"
        ActiveWorkbook.ActiveSheet.Cells(1, 5) = "Ãðóïïà çàäà÷"
        ActiveWorkbook.ActiveSheet.Cells(1, 6) = "Îïèñàíèå"
        ActiveWorkbook.ActiveSheet.Cells(1, 7) = "Ññûëêà"
                    For i = 1 To ActiveWorkbook.Worksheets.Count
            mesac = Right(Left(Data(y), 5), 2)
'            MsgBox (mesac)
            If ActiveWorkbook.Worksheets(i).Name = mesac Then
            ActiveWorkbook.Sheets(ActiveWorkbook.Worksheets(i).Name).Select
        p = 1
        Do While (ActiveWorkbook.ActiveSheet.Cells(p, 1)) <> ""
        p = p + 1
'            MsgBox (p)
        Loop
        ActiveWorkbook.ActiveSheet.Cells(p, 1).Value = Data(y)
        ActiveWorkbook.ActiveSheet.Cells(p, 2) = NameUsser(y)
        ActiveWorkbook.ActiveSheet.Cells(p, 3) = Dlitelnost(y)
        ActiveWorkbook.ActiveSheet.Cells(p, 4) = Stoimost(y)
        ActiveWorkbook.ActiveSheet.Cells(p, 5) = Zadacha(y)
        ActiveWorkbook.ActiveSheet.Cells(p, 6) = OpisZadach(y)
        ActiveWorkbook.ActiveSheet.Cells(p, 7) = Ssulka(y)
        End If
        Next i
    
    End If
End If
Next x

End If

Next y

Workbooks("No project.xlsx").Close SaveChanges:=True
For x = 0 To 1000
If OneName(x) <> "" And OneName(x) <> "." Then
Workbooks(OneName(x) & ".xlsx").Close SaveChanges:=True
'MsgBox (x)
End If
Next x

MsgBox ("Ãîòîâî")
End Sub


Function bBookOpen(wbName As String) As Boolean
    Dim wbBook As Workbook: On Error Resume Next
    Set wbBook = Workbooks(wbName)
    bBookOpen = Not wbBook Is Nothing
End Function

Private Sub CommandButton2_Click()

Workbooks("Create_employee_reports.xlsm").Activate
'Âûáîð ôàéëîâ è çàïèñü èõ â ÿ÷åéêó
Filename = Application _
    .GetOpenFilename("Excel Files (*.xlsx), *.xlsx", 1, "Âûáåðèòå ôàéë", "Âûáðàòü", True)
Dim element As Variant
Dim i As Integer, y As Integer
i = 1
Dim l As Integer
l = 1
For l = 1 To 1000
Cells(l, 1) = ""
Next l
For Each element In Filename
Cells(i, 1).Value = element
i = i + 1
Next

'Çàïèñü ïóòåé â ñòðîêîâûé ìàññèâ
Dim path(100) As String
i = 1
For y = 0 To 100
path(y) = Cells(i, 1).Value
i = i + 1
Next y


'Îòêðûòèå ôàéëîâ è äåéñòâèé íàä íèìè, ïðîâåðêà âñåõ ëèñòîâ
i = 0
ch = 0
Dim m As Integer
Dim u As Integer, f As Integer, f1 As Integer, f2 As Integer, f3 As Integer, f4 As Integer
f3 = 1
f2 = 2
f = 0
f1 = 1
u = 0
l = 1
f4 = 1
Dim Data(1000) As String, Dlitelnost(1000) As String, str1 As String, proekt(1000) As String


Do While path(i) <> ""

Workbooks.Open Filename:=path(i)

For m = 1 To ActiveWorkbook.Worksheets.Count

ActiveWorkbook.Worksheets(m).Activate
Dim shna As String
shna = Left(ActiveWorkbook.Worksheets(m).Name, 5)
If shna = "Count" Then
ActiveWorkbook.ActiveSheet.Cells(1, 12) = "Èòîãî ïî ÷àñàì è äíÿì"
Do While ActiveWorkbook.ActiveSheet.Cells(l, 1) <> ""
Data(u) = Left(ActiveWorkbook.ActiveSheet.Cells(l, 1).Text, 5)
Dlitelnost(u) = ActiveWorkbook.ActiveSheet.Cells(l, 2).Text
u = u + 1
l = l + 1
Loop
Do While f <> 999
Do While (Data(f) = Data(f1) And f1 <> 999)
ActiveWorkbook.ActiveSheet.Cells(f2, 12) = Data(f)
ActiveWorkbook.ActiveSheet.Cells(f4, 14) = Dlitelnost(f1)
f4 = f4 + 1
f1 = f1 + 1
ActiveWorkbook.ActiveSheet.Cells(1, 15).FormulaR1C1 = "=SUM(C[-1])"
ActiveWorkbook.ActiveSheet.Cells(f2, 13) = ActiveWorkbook.ActiveSheet.Cells(1, 15).Text
Loop
f = f + 1
ActiveWorkbook.ActiveSheet.Columns("N:O").Select
Selection.ClearContents
Do While ActiveWorkbook.ActiveSheet.Cells(f3, 12).Text <> ""
f3 = f3 + 1
'MsgBox f3
Loop
f2 = f3
Loop
f1 = 0
f4 = 1



Dim n As Integer, n2 As Integer, n3 As Integer, n4 As Integer

n = 1
n1 = 0


For n = 2 To 1000
proekt(n1) = ActiveWorkbook.ActiveSheet.Cells(n, 5).Text
n1 = n1 + 1
Next n

Dim j As Integer
j = 1
Do While ActiveWorkbook.ActiveSheet.Cells(j, 1).Text <> ""
j = j + 1
Loop

j = j + 2
For n1 = 0 To 999
ActiveWorkbook.ActiveSheet.Cells(j, 11) = proekt(n1)
j = j + 1
Next n1

Dim x As Integer
'Ñîðòèðîâêà ïî àëôàâèòó äëÿ âû÷èñëåíèé äóáëèêàòîâ
    ActiveWorkbook.ActiveSheet.Columns("K:K").Select
    ActiveWorkbook.Worksheets(m).Sort.SortFields.Clear
    ActiveWorkbook.Worksheets(m).Sort.SortFields.Add2 Key:=Range( _
        "K1"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
        xlSortNormal
    With ActiveWorkbook.Worksheets(m).Sort
        .SetRange Range("K1:K840")
        .Header = xlNo
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
    
n1 = 1
n = 0
For n1 = 1 To 1000
proekt(n) = ActiveWorkbook.ActiveSheet.Cells(n1, 11).Text
n = n + 1
Next n1

    ActiveWorkbook.ActiveSheet.Columns("K:K").Select
    Selection.ClearContents

j = 1
Do While ActiveWorkbook.ActiveSheet.Cells(j, 1).Text <> ""
j = j + 1
Loop

Dim v As Integer
j = j + 2
v = j
For n1 = 0 To 999
ActiveWorkbook.ActiveSheet.Cells(j, 11) = proekt(n1)
j = j + 1
Next n1
    
'Óäàëåíèå äóáëèêàòîâ
    For x = v To 2000
    If ActiveWorkbook.ActiveSheet.Cells(x + 1, 11).Text = ActiveWorkbook.ActiveSheet.Cells(x, 11).Text And ActiveWorkbook.ActiveSheet.Cells(x + 1, 11) <> "" Then
    ActiveWorkbook.ActiveSheet.Rows(x + 1 & ":" & x + 1).Select
    Selection.Delete Shift:=xlUp
    x = x - 1
    End If
    Next x


Dim s As String
s = v
Do While ActiveWorkbook.ActiveSheet.Cells(s, 11).Text <> ""
s = s + 1
Loop
ActiveWorkbook.ActiveSheet.Cells(s, 11) = "ÎÏÐ"


    Application.CutCopyMode = False
    ActiveWorkbook.ActiveSheet.Cells(v - 1, 8) = "Total time:"
    ActiveWorkbook.ActiveSheet.Cells(v - 1, 9).FormulaR1C1 = "=SUM(C[-7])"
    ActiveWorkbook.ActiveSheet.Columns("I:I").Select
    Selection.NumberFormat = "[h]:mm:ss"


s = v
x = 0
Do While ActiveWorkbook.ActiveSheet.Cells(s, 11).Text <> ""
proekt(x) = ActiveWorkbook.ActiveSheet.Cells(s, 11).Text
s = s + 1
x = x + 1
Loop
proekt(x + 1) = ""

Dim kk As Integer, uv As Integer
uv = v
x = 0
f1 = v
Do While proekt(x) <> ""
'MsgBox Proekt(x)
For kk = 1 To 1000
If (proekt(x) = "ÎÏÐ") Then
GoTo ex
End If
If (proekt(x) = ActiveWorkbook.ActiveSheet.Cells(kk, 5).Text) Then
ActiveWorkbook.ActiveSheet.Cells(f1, 15) = ActiveWorkbook.ActiveSheet.Cells(kk, 2).Text
f1 = f1 + 1
End If
Next kk

    Application.CutCopyMode = False
    ActiveWorkbook.ActiveSheet.Cells(106, 16).FormulaR1C1 = "=SUM(C[-1])"
    ActiveWorkbook.ActiveSheet.Cells(106, 16).Range("P107").Select
    ActiveWorkbook.ActiveSheet.Columns("J:J").Select
    ActiveWorkbook.ActiveSheet.Range("J100").Activate
    Selection.NumberFormat = "[h]:mm:ss"
    ActiveWorkbook.ActiveSheet.Cells(uv, 10) = ActiveWorkbook.ActiveSheet.Cells(106, 16).Text
    ActiveWorkbook.ActiveSheet.Cells(106, 16) = ""

    ActiveWorkbook.ActiveSheet.Columns("O:O").Select
    ActiveWorkbook.ActiveSheet.Range("O100").Activate
    Selection.ClearContents
uv = uv + 1
x = x + 1
Loop
ex:

Dim a As String, cenana As Integer, banana As Integer
banana = 1
'a = InputBox("Ââåäèòå ñòàâêó äëÿ " & ActiveWorkbook.Name)
Do While (Workbooks("Create_employee_reports.xlsm").Sheets("Stavki_email").Cells(banana, 1) <> ActiveWorkbook.ActiveSheet.Cells(1, 8).Text)
banana = banana + 1
Loop

a = Workbooks("Create_employee_reports.xlsm").Sheets("Stavki_email").Cells(banana, 3)

ActiveWorkbook.ActiveSheet.Cells(v - 5, 13) = "Ñòàâêà: "
ActiveWorkbook.ActiveSheet.Cells(v - 5, 14) = a
ActiveWorkbook.ActiveSheet.Cells(v - 5, 14).Select
    Selection.NumberFormat = "#,##0.00 $"
'    a = InputBox("Ââåäèòå âòîðîå çíà÷åíèå äëÿ " & ActiveWorkbook.Name)
a = Workbooks("Create_employee_reports.xlsm").Sheets("Stavki_email").Cells(banana, 4)
ActiveWorkbook.ActiveSheet.Cells(v - 4, 14) = a
ActiveWorkbook.ActiveSheet.Cells(v - 4, 14).Select
    Selection.NumberFormat = "#,##0.00 $"
    
    ActiveWorkbook.ActiveSheet.Cells(v - 3, 14).Select
    Application.CutCopyMode = False
    ActiveCell.FormulaR1C1 = "=R[-1]C"
ActiveWorkbook.ActiveSheet.Cells(v - 3, 14).Select
    Selection.NumberFormat = "#,##0.00 $"
cenana = v - 3

Dim qw As Integer, cifr As Integer, cifr1 As Integer
qw = v
cifr = qw - 3
Do While (ActiveWorkbook.ActiveSheet.Cells(qw, 11).Text <> "" And ActiveWorkbook.ActiveSheet.Cells(qw, 11).Text <> "ÎÏÐ")
    ActiveWorkbook.ActiveSheet.Cells(qw, 12).Select
    ActiveWorkbook.ActiveSheet.Cells(qw, 12).FormulaR1C1 = "=R" & cifr & "C14*RC[-2]*24"
    ActiveWorkbook.ActiveSheet.Cells(qw, 12).Select
    Selection.NumberFormat = "#,##0.00 $"
    qw = qw + 1
Loop

ActiveWorkbook.ActiveSheet.Cells(v + 3, 13) = "Îñòàòîê:"
    Application.CutCopyMode = False
    ActiveWorkbook.ActiveSheet.Cells(v + 3, 14).FormulaR1C1 = "=R101C14-R[-3]C[-2]:R[82]C[-2]"
    ActiveWorkbook.ActiveSheet.Cells(v + 3, 14).Select
    Selection.NumberFormat = "#,##0.00 $"


qw = v
Do While ActiveWorkbook.ActiveSheet.Cells(v, 12) <> ""
v = v + 1
Loop

ActiveWorkbook.ActiveSheet.Cells(v, 12) = ActiveWorkbook.ActiveSheet.Cells(qw + 3, 14).Text


'Ïðåîáðàçîâàíèå ìàññèâà ñ íàçâàíèåì ïðîåêòà
Dim klk As Integer
klk = 0
Do While proekt(klk) <> ""
klk = klk + 1
Loop
proekt(klk - 1) = "."

qw = v + 2
ActiveWorkbook.ActiveSheet.Cells(qw, 12) = "Çàäà÷è ïî ïðîåêòàì:"
Dim novotch As Integer
novotch = qw + 1
f1 = 0

f3 = 0
f4 = 0
qw = qw + 1
Do While proekt(f1) <> ""
f2 = 1
qw = qw + 1
ActiveWorkbook.ActiveSheet.Cells(qw, 12) = proekt(f1)

Dim OPR As String
Do While ActiveWorkbook.ActiveSheet.Cells(f2, 4) <> "" And proekt(f1) <> "."

If (ActiveWorkbook.ActiveSheet.Cells(f2, 5) = proekt(f1)) Then
ActiveWorkbook.ActiveSheet.Cells(qw, 13) = ActiveWorkbook.ActiveSheet.Cells(f2, 4).Text
ActiveWorkbook.ActiveSheet.Cells(qw, 14) = ActiveWorkbook.ActiveSheet.Cells(f2, 2).Text
qw = qw + 1
End If
If proekt(f1) = "ÎÏÐ" Then
If (ActiveWorkbook.ActiveSheet.Cells(f2, 5) = "") Then
ActiveWorkbook.ActiveSheet.Cells(qw, 13) = ActiveWorkbook.ActiveSheet.Cells(f2, 4).Text
ActiveWorkbook.ActiveSheet.Cells(qw, 14) = ActiveWorkbook.ActiveSheet.Cells(f2, 2).Text
qw = qw + 1
End If
End If
f2 = f2 + 1
Loop
f1 = f1 + 1
Loop

Dim pust As Integer, pust1 As Integer, nach As Integer, konec As Integer, poisk As Integer, fir As Integer, las As Integer, las1 As Integer, las2 As Integer
pust = novotch
pust1 = 0
nach = 0
konec = 0

'Ðàññòàíîâêà ïî ôèëüòðó ãðóïï çàäà÷à ïî ïðîåêòàì è ðàñ÷åò ÷àñîâ, äåíåã
Do While ((ActiveWorkbook.ActiveSheet.Cells(pust, 13) <> "" And ActiveWorkbook.ActiveSheet.Cells(pust + 1, 13) = "") Or (ActiveWorkbook.ActiveSheet.Cells(pust, 13) <> "" And ActiveWorkbook.ActiveSheet.Cells(pust + 1, 13) <> "") Or (ActiveWorkbook.ActiveSheet.Cells(pust, 13) = "" And ActiveWorkbook.ActiveSheet.Cells(pust + 1, 13) <> ""))

If (ActiveWorkbook.ActiveSheet.Cells(pust, 12) <> "" And pust1 = 0) Then
pust1 = 1
nach = pust

ElseIf (ActiveWorkbook.ActiveSheet.Cells(pust + 1, 12) <> "" And pust1 = 1) Then
pust1 = 0
konec = pust

 ActiveWorkbook.ActiveSheet.Range(ActiveWorkbook.ActiveSheet.Cells(nach, 13), ActiveWorkbook.ActiveSheet.Cells(konec - 1, 14)).Select
ActiveWorkbook.Worksheets(m).Sort.SortFields.Clear
    ActiveWorkbook.Worksheets(m).Sort.SortFields.Add2 Key:=Range("M1"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
        xlSortNormal
    With ActiveWorkbook.Worksheets(m).Sort
        .SetRange Range(Cells(nach, 13), Cells(konec - 1, 14))
        .Header = xlNo
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
    If (nach <> 0) Then
    poisk = nach
    Do While poisk <> konec + 1
    
        If (ActiveWorkbook.ActiveSheet.Cells(nach, 13) = ActiveWorkbook.ActiveSheet.Cells(poisk, 13)) Then
        ActiveWorkbook.ActiveSheet.Cells(poisk, 15) = ActiveWorkbook.ActiveSheet.Cells(poisk, 14).Text
        ElseIf (ActiveWorkbook.ActiveSheet.Cells(nach, 13) <> ActiveWorkbook.ActiveSheet.Cells(poisk, 13)) Then
        Application.CutCopyMode = False
        ActiveWorkbook.ActiveSheet.Cells(nach, 16).FormulaR1C1 = "=SUM(RC[-1]:R" & poisk - 1 & "C[-1])"
        ActiveWorkbook.ActiveSheet.Cells(nach, 17).FormulaR1C1 = "=RC[-3]*R" & cenana & "C[-3]*24"
        nach = poisk
        poisk = poisk - 1
        End If
        
    poisk = poisk + 1
    Loop
    
    End If
    

'ElseIf (ActiveWorkbook.ActiveSheet.Cells(pust - 1, 12).Text = "ÎÏÐ" And pust1 = 1) Then
'pust1 = 0
'konec = pust
'
' ActiveWorkbook.ActiveSheet.Range(ActiveWorkbook.ActiveSheet.Cells(nach, 13), ActiveWorkbook.ActiveSheet.Cells(konec - 1, 14)).Select
'ActiveWorkbook.Worksheets(m).Sort.SortFields.Clear
'    ActiveWorkbook.Worksheets(m).Sort.SortFields.Add2 Key:=Range("M1"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
'        xlSortNormal
'    With ActiveWorkbook.Worksheets(m).Sort
'        .SetRange Range(Cells(nach, 13), Cells(konec - 1, 14))
'        .Header = xlNo
'        .MatchCase = False
'        .Orientation = xlTopToBottom
'        .SortMethod = xlPinYin
'        .Apply
'    End With
'
'    If (nach <> 0) Then
'    poisk = nach
'    Do While poisk <> konec + 1
'
'        If (ActiveWorkbook.ActiveSheet.Cells(nach, 13) = ActiveWorkbook.ActiveSheet.Cells(poisk, 13)) Then
'        ActiveWorkbook.ActiveSheet.Cells(poisk, 15) = ActiveWorkbook.ActiveSheet.Cells(poisk, 14).Text
'        ElseIf (ActiveWorkbook.ActiveSheet.Cells(nach, 13) <> ActiveWorkbook.ActiveSheet.Cells(poisk, 13)) Then
'        Application.CutCopyMode = False
'        ActiveWorkbook.ActiveSheet.Cells(nach, 16).FormulaR1C1 = "=SUM(RC[-1]:R" & poisk - 1 & "C[-1])"
'        ActiveWorkbook.ActiveSheet.Cells(nach, 17).FormulaR1C1 = "=RC[-3]*R" & cenana & "C[-3]*24"
'        nach = poisk
'        poisk = poisk - 1
'        End If
'
'    poisk = poisk + 1
'    Loop
'
'    End If


End If

pust = pust + 1

'MsgBox nach
'MsgBox konec
'MsgBox pust
Loop

las = 1
Do While ActiveWorkbook.ActiveSheet.Cells(las, 12) <> "ÎÏÐ"
las = las + 1
Loop


las1 = las
Do While ActiveWorkbook.ActiveSheet.Cells(las1, 14) <> ""
las1 = las1 + 1
Loop

las1 = las1 - 1



konec = las1 + 1
nach = las

 ActiveWorkbook.ActiveSheet.Range(ActiveWorkbook.ActiveSheet.Cells(nach, 13), ActiveWorkbook.ActiveSheet.Cells(konec - 1, 14)).Select
ActiveWorkbook.Worksheets(m).Sort.SortFields.Clear
    ActiveWorkbook.Worksheets(m).Sort.SortFields.Add2 Key:=Range("M1"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
        xlSortNormal
    With ActiveWorkbook.Worksheets(m).Sort
        .SetRange Range(Cells(nach, 13), Cells(konec - 1, 14))
        .Header = xlNo
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With

ActiveWorkbook.ActiveSheet.Cells(las, 16).FormulaR1C1 = "=SUM(RC[-2]:R" & las1 & "C[-2])"


las2 = 1
Do While ActiveWorkbook.ActiveSheet.Cells(las2, 11) <> "ÎÏÐ"
las2 = las2 + 1
Loop

        OPR = ActiveWorkbook.ActiveSheet.Cells(las2, 12).Text
ActiveWorkbook.ActiveSheet.Cells(las, 17) = OPR
    ActiveWorkbook.ActiveSheet.Columns("O:P").Select
    Selection.NumberFormat = "h:mm"

las = 1
Do While las <> 2000
If ActiveWorkbook.ActiveSheet.Cells(las, 16) <> "" Then
ActiveWorkbook.ActiveSheet.Cells(las, 14) = ActiveWorkbook.ActiveSheet.Cells(las, 16).Text
End If
las = las + 1
Loop


las = 1
Do While ActiveWorkbook.ActiveSheet.Cells(las, 17) = ""
las = las + 1
Loop

Do While las <> 2000

If (ActiveWorkbook.ActiveSheet.Cells(las, 17) <> "") Then
ActiveWorkbook.ActiveSheet.Cells(las, 15) = ActiveWorkbook.ActiveSheet.Cells(las, 17).Text
End If

las = las + 1
Loop


las = 1
Do While ActiveWorkbook.ActiveSheet.Cells(las, 17) = ""
las = las + 1
Loop

Do While las <> 2000
If (ActiveWorkbook.ActiveSheet.Cells(las, 17) = "" And ActiveWorkbook.ActiveSheet.Cells(las, 15) <> "") Then
    ActiveWorkbook.ActiveSheet.Rows(las).Select
    Selection.Delete Shift:=xlUp
    las = las - 1
End If

las = las + 1
Loop

las2 = 1
Do While ActiveWorkbook.ActiveSheet.Cells(las2, 12) <> "ÎÏÐ"
las2 = las2 + 1
Loop
las = las2

Dim pop As Integer

For pop = 0 To 500
If (ActiveWorkbook.ActiveSheet.Cells(las, 17) = "" And ActiveWorkbook.ActiveSheet.Cells(las, 15) = "") Then
    ActiveWorkbook.ActiveSheet.Rows(las).Select
    Selection.Delete Shift:=xlUp
    las = las - 1
End If

las = las + 1
Next pop


    ActiveWorkbook.ActiveSheet.Columns("P:Q").Select
    ActiveWorkbook.ActiveSheet.Range("P190").Activate
    Selection.ClearContents
    ActiveWorkbook.ActiveSheet.Columns("O:O").Select
    ActiveWorkbook.ActiveSheet.Range("O190").Activate
    Selection.NumberFormat = "#,##0.00 $"


Dim fgh As String
las = 1
Do While ActiveWorkbook.ActiveSheet.Cells(las, 11) <> "ÎÏÐ"
las = las + 1
Loop

fgh = ActiveWorkbook.ActiveSheet.Cells(las, 12).Text
ActiveWorkbook.ActiveSheet.Cells(las - 2, 14) = fgh

End If


Next m
ActiveWorkbook.Save
i = i + 1
Loop




'Âûçîâ îêîí
''Dim a As String
''a = InputBox("Íàïèøèòå ÷òî-íèáóäü ...")
''MsgBox a

End Sub

Private Sub CommandButton3_Click()
Workbooks("Create_employee_reports.xlsm").Activate
'Âûáîð ôàéëîâ è çàïèñü èõ â ÿ÷åéêó
Filename = Application _
    .GetOpenFilename("Excel Files (*.xlsx), *.xlsx", 1, "Âûáåðèòå ôàéë", "Âûáðàòü", True)
Dim element As Variant
Dim i As Integer, y As Integer
i = 1
Dim l As Integer
l = 1
For l = 1 To 1000
Cells(l, 1) = ""
Next l
For Each element In Filename
Cells(i, 1).Value = element
i = i + 1
Next

'Çàïèñü ïóòåé â ñòðîêîâûé ìàññèâ
Dim path(100) As String
i = 1
For y = 0 To 100
path(y) = Cells(i, 1).Value
i = i + 1
Next y


'Îòêðûòèå ôàéëîâ è äåéñòâèé íàä íèìè, ïðîâåðêà âñåõ ëèñòîâ
i = 0
ch = 0
Dim m As Integer
Dim u As Integer, f As Integer, f1 As Integer, f2 As Integer, f3 As Integer, f4 As Integer
f3 = 1
f2 = 2
f = 0
f1 = 1
u = 0
l = 1
f4 = 1
Dim Data(1000) As String, Dlitelnost(1000) As String, str1 As String, proekt(1000) As String


Do While path(i) <> ""

Workbooks.Open Filename:=path(i)

Dim nazkn As String
nazkn = ActiveWorkbook.Name

For m = 1 To ActiveWorkbook.Worksheets.Count

ActiveWorkbook.Worksheets(m).Activate
Dim shna As String
shna = Left(ActiveWorkbook.Worksheets(m).Name, 5)
If shna = "Count" Then

Dim a As String
a = ActiveWorkbook.ActiveSheet.Cells(1, 8).Text
Workbooks.Add
ActiveWorkbook.SaveAs Filename:="import_flim_" & a & ".xls"

ActiveWorkbook.ActiveSheet.Cells(1, 1) = "worker_email"
ActiveWorkbook.ActiveSheet.Cells(1, 2) = "legal_type_id"
ActiveWorkbook.ActiveSheet.Cells(1, 3) = "title"
ActiveWorkbook.ActiveSheet.Cells(1, 4) = "group_name"
ActiveWorkbook.ActiveSheet.Cells(1, 5) = "attribute_1_value"
ActiveWorkbook.ActiveSheet.Cells(1, 6) = "attribute_2_value"
ActiveWorkbook.ActiveSheet.Cells(1, 7) = "attribute_3_value"
ActiveWorkbook.ActiveSheet.Cells(1, 8) = "attribute_4_value"
ActiveWorkbook.ActiveSheet.Cells(1, 9) = "end_date"
ActiveWorkbook.ActiveSheet.Cells(1, 10) = "copyright"
ActiveWorkbook.ActiveSheet.Cells(1, 11) = "amount"
ActiveWorkbook.ActiveSheet.Cells(1, 12) = "currency_id"
ActiveWorkbook.ActiveSheet.Cells(1, 13) = "email_language"
ActiveWorkbook.ActiveSheet.Cells(1, 14) = "file_links"

Dim b As Integer
b = 2

Workbooks(nazkn).Activate
Dim banana As Integer, email As String

banana = 1
Do While (Workbooks("Create_employee_reports.xlsm").Sheets("Stavki_email").Cells(banana, 1) <> ActiveWorkbook.ActiveSheet.Cells(1, 8))
banana = banana + 1
Loop
email = Workbooks("Create_employee_reports.xlsm").Sheets("Stavki_email").Cells(banana, 2).Text

Dim ff As Integer, title As String, titleeng As String, title_id As String, att1 As String, att2 As String, att3 As String, legalType As String, time As String, amount As String
Dim proekt1 As String, dd As String

titleeng = "ne nashel"
ff = 1
Do While ActiveWorkbook.ActiveSheet.Cells(ff, 12) <> "Çàäà÷è ïî ïðîåêòàì:"
ff = ff + 1
Loop

Dim uh As Integer, ah As Integer

ff = ff + 2
Do While ActiveWorkbook.ActiveSheet.Cells(ff, 12) <> "ÎÏÐ"
If (ActiveWorkbook.ActiveSheet.Cells(ff, 12) <> "") Then
proekt1 = ActiveWorkbook.ActiveSheet.Cells(ff, 12).Text
uh = ff
Do While ActiveWorkbook.ActiveSheet.Cells(uh, 13) <> ""
title = ActiveWorkbook.ActiveSheet.Cells(uh, 13).Text
time = ActiveWorkbook.ActiveSheet.Cells(uh, 14).Text
amount = ActiveWorkbook.ActiveSheet.Cells(uh, 15).Text
banana = 1
For ah = 1 To 300
'Do While (Workbooks("Create_employee_reports.xlsm").Sheets("title_list_task").Cells(banana, 3) <> title Or Workbooks("Create_employee_reports.xlsm").Sheets("title_list_task").Cells(banana, 3) <> "")
If (Workbooks("Create_employee_reports.xlsm").Sheets("title_list_task").Cells(banana, 3) = title Or Workbooks("Create_employee_reports.xlsm").Sheets("title_list_task").Cells(banana, 3) = "") Then
GoTo tut
End If
banana = banana + 1
'Loop
Next ah
tut:
titleeng = Workbooks("Create_employee_reports.xlsm").Sheets("title_list_task").Cells(banana, 1)
title_id = Workbooks("Create_employee_reports.xlsm").Sheets("title_list_task").Cells(banana, 2)

If (title_id = "1129" Or title_id = "1195" Or title_id = "1219" Or title_id = "1202" Or title_id = "1092" Or title_id = "1072" Or title_id = "1198" Or title_id = "1076") Then
att1 = proekt1
ElseIf (title_id = "1215" Or title_id = "1107" Or title_id = "1218" Or title_id = "1405" Or title_id = "1195") Then
att1 = proekt1
Else
att1 = Workbooks("Create_employee_reports.xlsm").Sheets("title_list_task").Cells(banana, 5)
End If
att2 = Workbooks("Create_employee_reports.xlsm").Sheets("title_list_task").Cells(banana, 6)
att3 = Workbooks("Create_employee_reports.xlsm").Sheets("title_list_task").Cells(banana, 7)
legalType = Workbooks("Create_employee_reports.xlsm").Sheets("title_list_task").Cells(banana, 4)

Workbooks("import_flim_" & a & ".xls").Activate
ActiveWorkbook.ActiveSheet.Cells(b, 1) = email
ActiveWorkbook.ActiveSheet.Cells(b, 2) = title_id
ActiveWorkbook.ActiveSheet.Cells(b, 3) = titleeng
ActiveWorkbook.ActiveSheet.Cells(b, 4) = proekt1
ActiveWorkbook.ActiveSheet.Cells(b, 5) = " " & att1 & " "
ActiveWorkbook.ActiveSheet.Cells(b, 6) = " " & att2 & " "
ActiveWorkbook.ActiveSheet.Cells(b, 7) = " " & att3 & " "
ActiveWorkbook.ActiveSheet.Cells(b, 8) = ""
ActiveWorkbook.ActiveSheet.Cells(b, 9).Value = Format(Now, "dd.mm.yyyy")
ActiveWorkbook.ActiveSheet.Cells(b, 10) = "0"
ActiveWorkbook.ActiveSheet.Cells(b, 11) = amount
ActiveWorkbook.ActiveSheet.Cells(b, 12) = "1"
ActiveWorkbook.ActiveSheet.Cells(b, 13) = "ru"
ActiveWorkbook.ActiveSheet.Cells(b, 14) = ""
Workbooks(nazkn).Activate
b = b + 1
uh = uh + 1
Loop
End If
ff = ff + 1
Loop


amount = ActiveWorkbook.ActiveSheet.Cells(ff, 15).Text
titleeng = "Services of processing of incoming and outgoing documents and phone calls"
title_id = "1177"
att1 = "Äàòû"
att2 = ""
att3 = ""
legalType = "Incoming/outgoing documents and phone calls processing"

Workbooks("import_flim_" & a & ".xls").Activate
ActiveWorkbook.ActiveSheet.Cells(b, 1) = email
ActiveWorkbook.ActiveSheet.Cells(b, 2) = title_id
ActiveWorkbook.ActiveSheet.Cells(b, 3) = titleeng
ActiveWorkbook.ActiveSheet.Cells(b, 4) = proekt1
ActiveWorkbook.ActiveSheet.Cells(b, 5) = "" & att1 & ""
ActiveWorkbook.ActiveSheet.Cells(b, 6) = "" & att2 & ""
ActiveWorkbook.ActiveSheet.Cells(b, 7) = "" & att3 & ""
ActiveWorkbook.ActiveSheet.Cells(b, 8) = ""
ActiveWorkbook.ActiveSheet.Cells(b, 9).Value = Format(Now, "dd.mm.yyyy")
ActiveWorkbook.ActiveSheet.Cells(b, 10) = "0"
ActiveWorkbook.ActiveSheet.Cells(b, 11) = amount
ActiveWorkbook.ActiveSheet.Cells(b, 12) = "1"
ActiveWorkbook.ActiveSheet.Cells(b, 13) = "ru"
ActiveWorkbook.ActiveSheet.Cells(b, 14) = ""
ActiveWorkbook.Save
Workbooks(nazkn).Activate


End If
Next m
i = i + 1
Loop
End Sub
