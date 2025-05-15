Attribute VB_Name = "ModAMR"
' A module for working with the data of uploads (books) of arithmetic mean balances (CFT)
Option Explicit
Option Compare Text

' A data set for unloading arithmetic mean balances from the CFT
Private Type DataAMR
    Account As String
    CurrencyCode As String
    ClientName As String
    IncomingBalance As Currency
    DebitTurnover As Currency
    CreditTurnover As Currency
    OutgoingBalance As Currency
    AverageBalance As Currency
    AverageBalanceOfNP As Currency
    Division As String
    RegistrationNumber As Long
End Type

' Types of accounts
Private Enum ViewAccount
    LEGAL
    DEPOSIT_LEGAL
    PHYSICAL
    DEPOSIT_PHYSICAL
    LETTERS_OF_CREDIT
End Enum

' File Dialog Filter
Private Enum FileFilter
    MS_WORD
    MS_EXCEL
    MS_ACCESS
    Txt
    XML
    IMAGE
End Enum

' Splits the book into parts (by account number)
Public Sub BreakBookIntoParts()
    On Error GoTo ErrorHandler
    Dim ErrorMsg As String
    
    Dim WshShell As Object
    Dim PathToDocuments As String
    Set WshShell = CreateObject("WScript.Shell")
    PathToDocuments = WshShell.SpecialFolders("MyDocuments")

    Dim PathToUploadFile As String
    PathToUploadFile = _
        GetPathToFile("Выбирите файл выгрузки ср. арифм. ост.", MS_EXCEL, PathToDocuments)
    
    If PathToUploadFile = "NotFound" Then
        ErrorMsg = "Не выбран файл выгрузки ср. арифм. ост."
        GoTo ErrorHandler
    End If
    
    Dim TableName As String
    
    Dim FullDataAMRs() As DataAMR
    FullDataAMRs = GetFullDataAMR(PathToUploadFile, TableName, ErrorMsg)
    If ErrorMsg <> "" Then GoTo ErrorHandler

    Dim FSO As Object
    Set FSO = CreateObject("Scripting.FileSystemObject")
    Dim PathToFolder As String
    PathToFolder = FSO.GetParentFolderName(PathToUploadFile)
    
    Dim PathToResultFile As String
    Dim FiltDataAMRs() As DataAMR
    
    FiltDataAMRs = FiltArrayAMR(FullDataAMRs, ConstArrayMasks(LEGAL))
    If FiltDataAMRs(0).Account <> "NotFound" Then
        PathToResultFile = GetUniqueFullPathToFile(FSO, PathToFolder, "Юр_Лица_Ср_Ост_" & StringDateToFormatDate(TableName) & ".xlsx")
        If PathToResultFile = "Error" Then
            ErrorMsg = "Объект Scripting.FileSystemObject не инициализирован."
            GoTo ErrorHandler
        End If
        ErrorMsg = CreateWorkbookAMR(FiltDataAMRs, TableName, PathToResultFile, ErrorMsg, False)
        If ErrorMsg <> "" Then GoTo ErrorHandler
    End If
    Erase FiltDataAMRs
    
    FiltDataAMRs = FiltArrayAMR(FullDataAMRs, ConstArrayMasks(DEPOSIT_LEGAL))
    If FiltDataAMRs(0).Account <> "NotFound" Then
        PathToResultFile = GetUniqueFullPathToFile(FSO, PathToFolder, "Юр_Лица_Депозиты_СрОст_" & StringDateToFormatDate(TableName) & ".xlsx")
        If PathToResultFile = "Error" Then
            ErrorMsg = "Объект Scripting.FileSystemObject не инициализирован."
            GoTo ErrorHandler
        End If
        ErrorMsg = CreateWorkbookAMR(FiltDataAMRs, TableName, PathToResultFile, ErrorMsg, False)
        If ErrorMsg <> "" Then GoTo ErrorHandler
    End If
    Erase FiltDataAMRs
    
    FiltDataAMRs = FiltArrayAMR(FullDataAMRs, ConstArrayMasks(PHYSICAL))
    If FiltDataAMRs(0).Account <> "NotFound" Then
        PathToResultFile = GetUniqueFullPathToFile(FSO, PathToFolder, "Физ_Лица_Ср_Ост_" & StringDateToFormatDate(TableName) & ".xlsx")
        If PathToResultFile = "Error" Then
            ErrorMsg = "Объект Scripting.FileSystemObject не инициализирован."
            GoTo ErrorHandler
        End If
        ErrorMsg = CreateWorkbookAMR(FiltDataAMRs, TableName, PathToResultFile, ErrorMsg, False)
        If ErrorMsg <> "" Then GoTo ErrorHandler
    End If
    Erase FiltDataAMRs
    
    FiltDataAMRs = FiltArrayAMR(FullDataAMRs, ConstArrayMasks(DEPOSIT_PHYSICAL))
    If FiltDataAMRs(0).Account <> "NotFound" Then
        PathToResultFile = GetUniqueFullPathToFile(FSO, PathToFolder, "Физ_Лица_Депозиты_СрОст_" & StringDateToFormatDate(TableName) & ".xlsx")
        If PathToResultFile = "Error" Then
            ErrorMsg = "Объект Scripting.FileSystemObject не инициализирован."
            GoTo ErrorHandler
        End If
        ErrorMsg = CreateWorkbookAMR(FiltDataAMRs, TableName, PathToResultFile, ErrorMsg, False)
        If ErrorMsg <> "" Then GoTo ErrorHandler
    End If
    Erase FiltDataAMRs
    
    FiltDataAMRs = FiltArrayAMR(FullDataAMRs, ConstArrayMasks(LETTERS_OF_CREDIT))
    If FiltDataAMRs(0).Account <> "NotFound" Then
        PathToResultFile = GetUniqueFullPathToFile(FSO, PathToFolder, "Аккредитивы_СрОст_" & StringDateToFormatDate(TableName) & ".xlsx")
        If PathToResultFile = "Error" Then
            ErrorMsg = "Объект Scripting.FileSystemObject не инициализирован."
            GoTo ErrorHandler
        End If
        ErrorMsg = CreateWorkbookAMR(FiltDataAMRs, TableName, PathToResultFile, ErrorMsg, False)
        If ErrorMsg <> "" Then GoTo ErrorHandler
    End If
    Erase FiltDataAMRs
    
Complete:
    Set WshShell = Nothing
    Set FSO = Nothing
    
    MsgBox "Успешное преобразование файла выгрузки.", vbOKOnly, "Выполнено"
    Exit Sub
    
ErrorHandler:
    If ErrorMsg = "" Then ErrorMsg = "Ошибка №" + Err.Number + ", описание:" + Err.Description
    MsgBox ErrorMsg, vbCritical, "Ошибка"
    Resume Complete
End Sub

' Extracts the desired date in the MMMMYYYY format from the row (table name).
' Str   - a string containing a date in the dd/mm/yyyy format at the end.
Private Function StringDateToFormatDate(Str As String) As String
    Dim Dt As Date
    Dt = CDate(Mid(Str, Len(Str) - 10, 11))
    StringDateToFormatDate = MonthName(Month(Dt)) & Year(Dt)
End Function

' Returns the unique path to the file (if the file is in the specified directory, appends a digit to the file name)
' FSO           - "Scripting.FileSystemObject"
' PathToFolder  - The path to the directory where you want to save the file
' NameFile      - File name (Example: Report.xlsx )
Private Function GetUniqueFullPathToFile(FSO As Object, PathToFolder As String, NameFile As String) As String
    If FSO Is Nothing Then
        GetUniqueFullPathToFile = "Error"
        Exit Function
    End If
    
    Dim FullPathToFile As String
    Dim i As Long
    
    FullPathToFile = PathToFolder & "\" & NameFile
    
    Do While FSO.FileExists(FullPathToFile)
        FullPathToFile = PathToFolder & "\" & "(" & i & ")" & NameFile
        i = i + 1
    Loop
    
    GetUniqueFullPathToFile = FullPathToFile
End Function


' Filters the upload file data by masks
' DataAMRs      - Upload file Data
' Masks         - Masks of required accounts
Private Function FiltArrayAMR(DataAMRs() As DataAMR, Masks() As String) As DataAMR()
    Dim Results() As DataAMR
    Dim IndexAMR As Long, IndexMask As Long, IndexResult As Long
    
    For IndexAMR = LBound(DataAMRs) To UBound(DataAMRs)
        For IndexMask = LBound(Masks) To UBound(Masks)
            If DataAMRs(IndexAMR).Account Like Masks(IndexMask) & "*" Then
                ReDim Preserve Results(IndexResult)
                Results(IndexResult) = DataAMRs(IndexAMR)
                IndexResult = IndexResult + 1
            End If
        Next IndexMask
    Next IndexAMR
    
    If IndexResult = 0 Then
        ReDim Preserve Results(IndexResult)
        Results(IndexResult).Account = "NotFound"
    End If
    
    FiltArrayAMR = Results
End Function
Sub TestFiltArrayAMR()
    Dim DataList(20) As DataAMR
    DataList(0).Account = "40701810000000000001"
    DataList(1).Account = "40702810000000000002"
    DataList(2).Account = "40703810000000000003"
    DataList(3).Account = "40705810000000000004"
    DataList(4).Account = "40802810000000000005"
    DataList(5).Account = "40807810000000000006"
    DataList(6).Account = "40827810000000000007"
    DataList(7).Account = "40832810000000000008"
    
    DataList(8).Account = "42001810000000000009"
    DataList(9).Account = "42112810000000000010"
    DataList(10).Account = "42503810000000000011"
    DataList(11).Account = "42104810000000000012"
    DataList(12).Account = "42205810000000000013"
    
    DataList(13).Account = "40817810000000000014"
    DataList(14).Account = "40820810000000000015"
    
    DataList(15).Account = "42307810000000000016"
    
    DataList(16).Account = "40901810000000000017"
    DataList(17).Account = "40902810000000000018"
    DataList(18).Account = "47410810000000000019"
    DataList(19).Account = "47431810000000000020"
    
    DataList(20).Account = "99999810000000000020"
    
    Dim Assert As New Assert
    Assert.NameMethod = "Tests function FiltArrayAMR"
    
    Dim Actuals() As DataAMR
    
    Actuals = FiltArrayAMR(DataList, ConstArrayMasks(LEGAL))
    Call Assert.Equal_Long(7, UBound(Actuals))
    Call Assert.Equal_String("40701810000000000001", Actuals(0).Account)
    Call Assert.Equal_String("40702810000000000002", Actuals(1).Account)
    Call Assert.Equal_String("40703810000000000003", Actuals(2).Account)
    Call Assert.Equal_String("40705810000000000004", Actuals(3).Account)
    Call Assert.Equal_String("40802810000000000005", Actuals(4).Account)
    Call Assert.Equal_String("40807810000000000006", Actuals(5).Account)
    Call Assert.Equal_String("40827810000000000007", Actuals(6).Account)
    Call Assert.Equal_String("40832810000000000008", Actuals(7).Account)
    Erase Actuals
    
    Actuals = FiltArrayAMR(DataList, ConstArrayMasks(DEPOSIT_LEGAL))
    Call Assert.Equal_Long(4, UBound(Actuals))
    Call Assert.Equal_String("42001810000000000009", Actuals(0).Account)
    Call Assert.Equal_String("42112810000000000010", Actuals(1).Account)
    Call Assert.Equal_String("42503810000000000011", Actuals(2).Account)
    Call Assert.Equal_String("42104810000000000012", Actuals(3).Account)
    Call Assert.Equal_String("42205810000000000013", Actuals(4).Account)
    Erase Actuals
    
    Actuals = FiltArrayAMR(DataList, ConstArrayMasks(PHYSICAL))
    Call Assert.Equal_Long(1, UBound(Actuals))
    Call Assert.Equal_String("40817810000000000014", Actuals(0).Account)
    Call Assert.Equal_String("40820810000000000015", Actuals(1).Account)
    Erase Actuals
    
    Actuals = FiltArrayAMR(DataList, ConstArrayMasks(DEPOSIT_PHYSICAL))
    Call Assert.Equal_Long(0, UBound(Actuals))
    Call Assert.Equal_String("42307810000000000016", Actuals(0).Account)
    Erase Actuals
    
    Actuals = FiltArrayAMR(DataList, ConstArrayMasks(LETTERS_OF_CREDIT))
    Call Assert.Equal_Long(3, UBound(Actuals))
    Call Assert.Equal_String("40901810000000000017", Actuals(0).Account)
    Call Assert.Equal_String("40902810000000000018", Actuals(1).Account)
    Call Assert.Equal_String("47410810000000000019", Actuals(2).Account)
    Call Assert.Equal_String("47431810000000000020", Actuals(3).Account)
    Erase Actuals
    
    Dim EmptyArrMask(0) As String
    EmptyArrMask(0) = "0234"
    
    Actuals = FiltArrayAMR(DataList, EmptyArrMask)
    Call Assert.Equal_Long(0, UBound(Actuals))
    Call Assert.Equal_String("NotFound", Actuals(0).Account)
    Erase Actuals
    
    Call Assert.ResultAssert
     
    Set Assert = Nothing
End Sub

' Creates a book for unloading arithmetic averages
' DataAMRs      - Data for the upload file
' TableName     - Name of the table
' PathToFile    - The full path to the new file
' ErrorMsg      - Possible error message
' IsVisible     - Whether to display the creation of a book
Private Function CreateWorkbookAMR(DataAMRs() As DataAMR, TableName As String, PathToFile As String, ErrorMsg As String, Optional IsVisible As Boolean = False) As String
    On Error GoTo ErrorHandler
    
    Dim ExcelApp As Object
    Set ExcelApp = CreateObject("Excel.Application")

    Dim Workbook As Object
    Set Workbook = ExcelApp.Workbooks.Add
    ExcelApp.Visible = IsVisible
    
    Dim Worksheet As Object
    Set Worksheet = Workbook.Sheets(1)
    
    Worksheet.Range("B2").Value = TableName
    With Worksheet.Range("B2:L2")
        .Merge
        .Font.Bold = True
        .HorizontalAlignment = -4108 'xlCenter
    End With
    Worksheet.Range("B5:D" & (UBound(DataAMRs) + 5)).NumberFormat = "@"
    Worksheet.Range("E5:J" & (UBound(DataAMRs) + 5)).NumberFormat = "#,##0.00"
    Worksheet.Range("K5:K" & (UBound(DataAMRs) + 5)).NumberFormat = "@"
    Worksheet.Range("L5:L" & (UBound(DataAMRs) + 5)).NumberFormat = "0"
    
    Worksheet.Cells(4, 2).Value = "№ счета"
    Worksheet.Cells(4, 3).Value = "Валюта"
    Worksheet.Cells(4, 4).Value = "Клиент"
    Worksheet.Cells(4, 5).Value = "Входящее сальдо"
    Worksheet.Cells(4, 6).Value = "Оборот Дебет"
    Worksheet.Cells(4, 7).Value = "Оборот Кредит"
    Worksheet.Cells(4, 8).Value = "Исходящее сальдо"
    Worksheet.Cells(4, 9).Value = "Средний остаток"
    Worksheet.Cells(4, 10).Value = "Средний остаток НП"
    Worksheet.Cells(4, 11).Value = "Подразделение"
    Worksheet.Cells(4, 12).Value = "Рег. номер"
    
    Dim IndexRow As Long
    For IndexRow = LBound(DataAMRs) To UBound(DataAMRs)
        Worksheet.Cells(IndexRow + 5, 2).Value = DataAMRs(IndexRow).Account
        Worksheet.Cells(IndexRow + 5, 3).Value = DataAMRs(IndexRow).CurrencyCode
        Worksheet.Cells(IndexRow + 5, 4).Value = DataAMRs(IndexRow).ClientName
        Worksheet.Cells(IndexRow + 5, 5).Value = DataAMRs(IndexRow).IncomingBalance
        Worksheet.Cells(IndexRow + 5, 6).Value = DataAMRs(IndexRow).DebitTurnover
        Worksheet.Cells(IndexRow + 5, 7).Value = DataAMRs(IndexRow).CreditTurnover
        Worksheet.Cells(IndexRow + 5, 8).Value = DataAMRs(IndexRow).OutgoingBalance
        Worksheet.Cells(IndexRow + 5, 9).Value = DataAMRs(IndexRow).AverageBalance
        Worksheet.Cells(IndexRow + 5, 10).Value = DataAMRs(IndexRow).AverageBalanceOfNP
        Worksheet.Cells(IndexRow + 5, 11).Value = DataAMRs(IndexRow).Division
        Worksheet.Cells(IndexRow + 5, 12).Value = DataAMRs(IndexRow).RegistrationNumber
    Next IndexRow
    
    With Worksheet.Range("B4:L" & (UBound(DataAMRs) + 5)).Borders
        .LineStyle = 1
        .Weight = 2
    End With
    With Worksheet.Range("B4:L4")
        .WrapText = True
        .Font.Bold = True
        .HorizontalAlignment = -4108 'xlCenter
    End With
    Worksheet.Columns("B:L").AutoFit
    
    Workbook.SaveAs PathToFile, 51
    Workbook.Close False
Complete:
    If Not ExcelApp Is Nothing Then
        ExcelApp.Quit
        Set ExcelApp = Nothing
    End If
    Set Workbook = Nothing
    Set Worksheet = Nothing
    
    CreateWorkbookAMR = ErrorMsg
    Exit Function
ErrorHandler:
    If ErrorMsg = "" Then ErrorMsg = "Ошибка №" + Err.Number + ", описание:" + Err.Description
    Resume Complete
End Function

' Returns an array of account masks
' View - the view of masks required
Private Function ConstArrayMasks(View As ViewAccount) As String()
    Dim ArrayLegal(7) As String
    ArrayLegal(0) = "40701"
    ArrayLegal(1) = "40702"
    ArrayLegal(2) = "40703"
    ArrayLegal(3) = "40705"
    ArrayLegal(4) = "40802"
    ArrayLegal(5) = "40807"
    ArrayLegal(6) = "40827"
    ArrayLegal(7) = "40832"
    
    Dim ArrayDepositLegal(4) As String
    ArrayDepositLegal(0) = "4200"
    ArrayDepositLegal(1) = "4211"
    ArrayDepositLegal(2) = "4250"
    ArrayDepositLegal(3) = "4210"
    ArrayDepositLegal(4) = "4220"
    
    Dim ArrayPhysical(1) As String
    ArrayPhysical(0) = "40817810"
    ArrayPhysical(1) = "40820810"
    
    Dim ArrayDepositPhysical(0) As String
    ArrayDepositPhysical(0) = "4230"
    
    Dim ArrayLettersOfCredit(3) As String
    ArrayLettersOfCredit(0) = "40901"
    ArrayLettersOfCredit(1) = "40902"
    ArrayLettersOfCredit(2) = "47410"
    ArrayLettersOfCredit(3) = "47431"
    
    Dim ArrayDefault(0) As String
    
    Select Case View
        Case ViewAccount.LEGAL:
           ConstArrayMasks = ArrayLegal
        Case ViewAccount.DEPOSIT_LEGAL:
           ConstArrayMasks = ArrayDepositLegal
        Case ViewAccount.PHYSICAL:
           ConstArrayMasks = ArrayPhysical
        Case ViewAccount.DEPOSIT_PHYSICAL:
           ConstArrayMasks = ArrayDepositPhysical
        Case ViewAccount.LETTERS_OF_CREDIT:
           ConstArrayMasks = ArrayLettersOfCredit
        Case Else:
           ConstArrayMasks = ArrayDefault
    End Select
End Function

' Returns all the data from the file unloading arithmetic averages
' PathToUploadFile  - The full path to the upload file
' ErrorMsg          - Possible error message
Private Function GetFullDataAMR(PathToUploadFile As String, TableName As String, ErrorMsg As String) As DataAMR()
    On Error GoTo ErrorHandler
    
    Dim ExcelApp As Object
    Set ExcelApp = CreateObject("Excel.Application")
    
    Dim Workbook As Object
    Set Workbook = ExcelApp.Workbooks.Open(PathToUploadFile)
    
    Dim Worksheet As Object
    Set Worksheet = Workbook.Sheets(1)
    
    Dim LastRow As Long, LastCol As Long
    
    LastRow = Worksheet.Cells(Worksheet.Rows.Count, 2).End(-4162).Row ' xlUp = -4162
    LastCol = Worksheet.Cells(4, Worksheet.Columns.Count).End(-4159).Column ' xlToLeft = -4159
    
    Dim Results() As DataAMR
    Dim IndexRow As Long, IndexColumn As Long, IndexResult As Long
    For IndexRow = 4 To LastRow
        If IndexRow = 4 Then
            If Worksheet.Cells(IndexRow, 2).Value <> "№ счета" Or _
               Worksheet.Cells(IndexRow, 3).Value <> "Валюта" Or _
               Worksheet.Cells(IndexRow, 4).Value <> "Клиент" Or _
               Worksheet.Cells(IndexRow, 5).Value <> "Входящее сальдо" Or _
               Worksheet.Cells(IndexRow, 6).Value <> "Оборот Дебет" Or _
               Worksheet.Cells(IndexRow, 7).Value <> "Оборот Кредит" Or _
               Worksheet.Cells(IndexRow, 8).Value <> "Исходящее сальдо" Or _
               Worksheet.Cells(IndexRow, 9).Value <> "Средний остаток" Or _
               Worksheet.Cells(IndexRow, 10).Value <> "Средний остаток НП" Or _
               Worksheet.Cells(IndexRow, 11).Value <> "Подразделение" Or _
               Worksheet.Cells(IndexRow, 12).Value <> "Рег. номер" Then
               ErrorMsg = "Содержание файла не соответствует шаблону выгрузки ср. арифм. ост."
               GoTo ErrorHandler
            End If
        Else
            ReDim Preserve Results(IndexResult)
            
            Results(IndexResult).Account = CStr(Worksheet.Cells(IndexRow, 2).Value)
            Results(IndexResult).CurrencyCode = CStr(Worksheet.Cells(IndexRow, 3).Value)
            Results(IndexResult).ClientName = CStr(Worksheet.Cells(IndexRow, 4).Value)
            Results(IndexResult).IncomingBalance = CCur(Worksheet.Cells(IndexRow, 5).Value)
            Results(IndexResult).DebitTurnover = CCur(Worksheet.Cells(IndexRow, 6).Value)
            Results(IndexResult).CreditTurnover = CCur(Worksheet.Cells(IndexRow, 7).Value)
            Results(IndexResult).OutgoingBalance = CCur(Worksheet.Cells(IndexRow, 8).Value)
            Results(IndexResult).AverageBalance = CCur(Worksheet.Cells(IndexRow, 9).Value)
            Results(IndexResult).AverageBalanceOfNP = CCur(Worksheet.Cells(IndexRow, 10).Value)
            Results(IndexResult).Division = CStr(Worksheet.Cells(IndexRow, 11).Value)
            Results(IndexResult).RegistrationNumber = CLng(Worksheet.Cells(IndexRow, 12).Value)
            
            IndexResult = IndexResult + 1
        End If
    Next IndexRow
    
    TableName = Worksheet.Cells(2, 2).Value
    
    If IndexResult = 0 Then ReDim Preserve Results(IndexResult) ' если нет данных то пустой массив
    
    GetFullDataAMR = Results
Complete:
    If Not ExcelApp Is Nothing Then
        ExcelApp.Quit
        Set ExcelApp = Nothing
    End If
    Set Workbook = Nothing
    Set Worksheet = Nothing
    Exit Function
    
ErrorHandler:
    If ErrorMsg = "" Then ErrorMsg = "Ошибка №" + Err.Number + ", описание:" + Err.Description
    Resume Complete
End Function

' Opens the file dialog box (file selection)
' Title             - The content of the window title
' Filter            - Filter by required files
' InitialFileName   - The default directory when opening the window (optional)
Private Function GetPathToFile(Title As String, Filter As FileFilter, Optional InitialFileName As String = "C:\") As String
    
    Dim FilterStr As String
    Select Case Filter
        Case 0:
            FilterStr = "*.docx*; *.doc*"
        Case 1:
            FilterStr = "*.xlsx*; *.xls*"
        Case 2:
            FilterStr = "*.accdb*; *.mdb*"
        Case 3:
            FilterStr = "*.txt*"
        Case 4:
            FilterStr = "*.xml*"
        Case 5:
            FilterStr = "*.png*; *.jpeg*; *.bmp*; *.jpg*"
        Case Else:
            FilterStr = "*"
    End Select
    
    With Application.FileDialog(1)
        .Title = Title
        .InitialFileName = InitialFileName
        .AllowMultiSelect = False
        .Filters.Clear
        .Filters.Add "Files", FilterStr, 1
        Dim ResultShow As Integer
        ResultShow = .Show
        If ResultShow = 0 Then
            GetPathToFile = "NotFound"
        Else
            GetPathToFile = Trim(.SelectedItems.Item(1))
        End If
    End With

End Function
