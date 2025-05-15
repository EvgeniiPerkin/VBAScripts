VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CyrName"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'***************Description*******************************
'*********************************************************

'***************Methods***********************************
' Слоняет имя (по раздельности) в указанный падеж
' Возвращает результат склонения в виде CyrNameResult
Public Function Decline_OneCase_MuchWord(surname As String, Name As String, patronymic As String, case_ As CasesEnum, Optional gender As GendersEnum = 0, Optional shorten As Boolean = False) As CyrNameResult
    Dim values() As String
    values = Decline_MuchWord(surname, Name, patronymic, CLng(case_), CLng(gender), shorten)
    
    Dim Result As New CyrNameResult
    Call Result.Constructor2(values)
    
    Set Decline_OneCase_MuchWord = Result
End Function

' Слоняет имя (одним словом) в указанный падеж
' Возвращает результат склонения в виде CyrNameResult
Public Function Decline_OneCase_OneWord(FullName As String, case_ As CasesEnum, Optional gender As GendersEnum = 0, Optional shorten As Boolean = False) As CyrNameResult
    Dim values() As String
    values = Decline_OneWord(FullName, CLng(case_), CLng(gender), shorten)
    
    Dim Result As New CyrNameResult
    Call Result.Constructor2(values)
    
    Set Decline_OneCase_OneWord = Result
End Function

' Слоняет имя (по раздельности) во все возможные падежи
' Возвращает результат склонения в виде CyrResult
Public Function Decline_AllCases_MuchWord(surname As String, Name As String, patronymic As String, Optional gender As GendersEnum = 0, Optional shorten As Boolean = False) As CyrResult
    Dim cases() As String
    ReDim cases(5)
     
    Dim cas As New CyrDeclineCase
    Dim item As Variant
    
    For Each item In cas.GetEnumerable
        Dim caseIndex As Long
        caseIndex = item.ind
        
        Dim values() As String
        values = Decline_MuchWord(surname, Name, patronymic, caseIndex, CLng(gender), shorten)
        
        Dim cyr_name_result As New CyrNameResult
        Call cyr_name_result.Constructor2(values)
    
        cases(caseIndex - 1) = cyr_name_result.ToString
    Next item
    
    Dim Result As New CyrResult
    Call Result.Constructor3(cases)
    
    Set Decline_AllCases_MuchWord = Result
End Function

' Слоняет имя (одним словом) во все возможные падежи
' Возвращает результат склонения в виде CyrResult
Public Function Decline_AllCases_OneWord(FullName As String, Optional gender As GendersEnum = 0, Optional shorten As Boolean = False) As CyrResult
    Dim cases() As String
    ReDim cases(5)
     
    Dim cas As New CyrDeclineCase
    Dim item As Variant
    Dim full_name As String
    
    full_name = FullName
     
    For Each item In cas.GetEnumerable
        Dim caseIndex As Long
        caseIndex = item.ind
        
        Dim values() As String
        values = Decline_OneWord(FullName, caseIndex, CLng(gender), shorten)
        
        Dim cyr_name_result As New CyrNameResult
        Call cyr_name_result.Constructor2(values)
    
        cases(caseIndex - 1) = cyr_name_result.ToString
        FullName = full_name
    Next item
    
    Dim Result As New CyrResult
    Call Result.Constructor3(cases)
    
    Set Decline_AllCases_OneWord = Result
End Function

' Слоняет имя (по раздельности) в указанный падеж
' inputSurname      Фамилия, в именительном падеже
' inputName         Имя, в именительном падеже
' inputPatronymic   Отчество, в именительном падеже
' inputCase         Падеж, в который нужно просклонять, где 1 - именительный, 2 - родительный, 3 - дательный, 4 - винительный, 5 - творительный, 6 - предложный
' inputGender       Пол, указанного имени, где 0 - определить автоматически, 1 - мужской, 2 - женский
' inputShorten      Сократить ли имя и отчество в результате склонения. К примеру, Иванов Иван Иванович будет Иванов И.И.
' Возвращает результат в виде массива из трех элементов [Фамилия, Имя, Отчество]
Public Function Decline_MuchWord(inputSurname As String, inputName As String, inputPatronymic As String, inputCase As Long, Optional inputGender As Long = 0, Optional inputShorten As Boolean = False) As String()
    Dim temp As String, surname As String, Name As String
    Dim patronymic As String, patronymicAfter As String, patronymicBefore As String
    Dim surnameNew As String, surnameOld As String
    Dim caseNumber As Long, gender As Long, index As Long
    Dim declinePatronymic As Boolean, isFeminine As Boolean
    
    caseNumber = inputCase
    gender = inputGender
    surname = ProperCase(inputSurname)
    Name = ProperCase(inputName)
    
    Call SplitPatronymic(inputPatronymic, patronymicBefore, patronymic, patronymicAfter)
    If patronymicAfter = "" Then
        declinePatronymic = True
    End If
    patronymic = ProperCase(patronymic)
    
    If caseNumber < 1 Or caseNumber > 6 Then
        caseNumber = 1
    End If

    If gender < 0 Or gender > 2 Then
        gender = 0
    End If

    If gender = 0 Then
        gender = IIf(Right(patronymic, 2) = "на", 2, 1)
    End If
    
    If gender = 2 Then
        isFeminine = True
    End If
            
    surnameOld = surname
    surnameNew = ""
    index = InStr(surnameOld, "-") - 1
    
    While (index > 0)
        temp = ProperCase(Mid(surnameOld, 1, index))
        surnameNew = surnameNew + DeclineSurname(temp, caseNumber, isFeminine) + "-"
        surnameOld = Right(surnameOld, Len(surnameOld) - (index + 1))
        index = InStr(surnameOld, "-") - 1
    Wend
    
    temp = ProperCase(surnameOld)
    surnameNew = surnameNew + DeclineSurname(temp, caseNumber, isFeminine)
    surname = surnameNew
    
    Select Case caseNumber
        Case 1:
            If inputShorten Then
                Name = shorten(Name)
                patronymic = shorten(patronymic)
            End If
        Case 2:
            Name = DeclineNameGenitive(Name, isFeminine, inputShorten)
            If declinePatronymic Or inputShorten Then
                patronymic = DeclinePatronymicGenitive(patronymic, isFeminine, inputShorten)
            End If
        Case 3:
            Name = DeclineNameDative(Name, isFeminine, inputShorten)
            If declinePatronymic Or inputShorten Then
                patronymic = DeclinePatronymicDative(patronymic, isFeminine, inputShorten)
            End If
        Case 4:
            Name = DeclineNameAccusative(Name, isFeminine, inputShorten)

            If declinePatronymic Or inputShorten Then
                patronymic = DeclinePatronymicAccusative(patronymic, isFeminine, inputShorten)
            End If

        Case 5:
            Name = DeclineNameInstrumental(Name, isFeminine, inputShorten)
            If declinePatronymic Or inputShorten Then
                patronymic = DeclinePatronymicInstrumental(patronymic, isFeminine, inputShorten)
            End If
        Case 6:
            Name = DeclineNamePrepositional(Name, isFeminine, inputShorten)

            If declinePatronymic Or inputShorten Then
                patronymic = DeclinePatronymicPrepositional(patronymic, isFeminine, inputShorten)
            End If
    End Select
    
    If Not inputShorten Then
        patronymic = patronymicBefore + patronymic + patronymicAfter
    End If

    Dim resultArr() As String
    ReDim resultArr(2)
    resultArr(0) = surname
    resultArr(1) = Name
    resultArr(2) = patronymic
    
    Decline_MuchWord = resultArr
End Function

' Слоняет имя (одним словом) в указанный падеж
' fullName      Полное имя, в именительном падеже
' case_         Падеж, в который нужно просклонять, где 1 - именительный, 2 - родительный, 3 - дательный, 4 - винительный, 5 - творительный, 6 - предложный
' gender        Пол, указанного имени, где 0 - определить автоматически, 1 - мужской, 2 - женский
' shorten       Сократить ли имя и отчество в результате склонения. К примеру, Иванов Иван Иванович будет Иванов И.И.
' Возвращает результат в виде массива из трех элементов [Фамилия, Имя, Отчество]
Public Function Decline_OneWord(FullName As String, case_ As Long, Optional gender As Long = 0, Optional shorten As Boolean = False) As String()
    Dim surname As String, Name As String, patronymic As String, str1 As String, str2 As String, str3 As String
    Dim paceIndex As Long, spaceIndex As Long
    
    spaceIndex = InStr(FullName, " ") - 1
    If spaceIndex > 0 Then
        str1 = Trim(Mid(FullName, 1, spaceIndex))
        FullName = Trim(Right(FullName, Len(FullName) - spaceIndex))

        spaceIndex = InStr(FullName, " ") - 1

        If (spaceIndex > 0) Then
            str2 = Trim(Mid(FullName, 1, spaceIndex))
            str3 = Trim(Right(FullName, Len(FullName) - spaceIndex))
        Else
            str2 = Trim(FullName)
        End If
    Else
        str1 = Trim(FullName)
    End If
    
    If Not str3 = "" Then
        If Right(str2, 2) = "ич" Or Right(str2, 3) = "вна" Or Right(str2, 3) = "чна" Then
            surname = str3
            Name = str1
            patronymic = str2
        Else
            surname = str1
            Name = str2
            patronymic = str3
        End If
    ElseIf Not str2 = "" Then
        If Right(str2, 2) = "ич" Or Right(str2, 3) = "вна" Or Right(str2, 3) = "чна" Then
            Name = str1
            patronymic = str2
        Else
            surname = str1
            Name = str2
        End If
    Else
        surname = str1
    End If
    Decline_OneWord = Decline_MuchWord(surname, Name, patronymic, case_, gender, shorten)
End Function
 
' Сокращает указанное значение. Пример: "Петровна" -> "П.", "Николай" -> "Н."
' Возвращает пустую строку если входное значение является пустой строкой
Public Function shorten(Value As String) As String
    If Trim(Value & vbNullString) = vbNullString Then
        shorten = ""
        Exit Function
    End If
    
    Dim Result As String
    Result = Mid(Value, 1, 1) & "."
    shorten = Result
End Function


' Возвращает true, если входная строка является пустой или заканчивается на точку
Public Function IsShorten(Value As String) As Boolean
    If Value = "" Then
        IsShorten = True
        Exit Function
    End If
    Dim point As String
    point = Mid(Value, Len(Value), 1)
    
    If point = "." Then
        IsShorten = True
    Else
        IsShorten = False
    End If
End Function

Public Function ProperCase(Value As String) As String
    'If value <> "" Then
    '    value = Trim(Replace(value, Chr("\uFEFF"), ""))
    'End If
    If Value = "" Then
        ProperCase = ""
        Exit Function
    End If
    ProperCase = UCase(Mid(Value, 1, 1)) + Right(Value, Len(Value) - 1)
End Function

Public Function SetEnd(Value As String, add As String) As String
    SetEnd = SetEndS(Value, Len(add), add)
End Function

Public Function SetEndS(Value As String, cut As Long, add As String) As String
    Dim Result As String
    Result = Left(Value, Len(Value) - cut) & add
    SetEndS = Result
End Function

Public Function SubstringRight(Value As String, cut As Long) As String
    If cut > Len(Value) Then
        cut = Len(Value)
    End If
    SubstringRight = Right(Value, cut)
End Function

Public Function PatronymicPrefixRegex(ByRef patronymic As String, ByRef Prefix As String) As Boolean
    Const param1 As String = "ибн "
    Const param2 As String = "ибн-"
    
    If InStr(1, LCase(patronymic), param1) = 1 Then
        Prefix = Mid(patronymic, 1, Len(param1))
        patronymic = Mid(patronymic, Len(param1) + 1, Len(patronymic))
        PatronymicPrefixRegex = True
        Exit Function
    End If
    If InStr(1, LCase(patronymic), param2) = 1 Then
        Prefix = Mid(patronymic, 1, Len(param2))
        patronymic = Mid(patronymic, Len(param2) + 1, Len(patronymic))
        PatronymicPrefixRegex = True
        Exit Function
    End If
    PatronymicPrefixRegex = False
End Function
Public Function PatronymicSuffixRegex(ByRef patronymic As String, ByRef suffix As String) As Boolean
    Const param1 As String = " оглы"
    Const param2 As String = " улы"
    Const param3 As String = " уулу"
    Const param4 As String = " кызы"
    Const param5 As String = " гызы"
    Const param6 As String = "-оглы"
    Const param7 As String = "-улы"
    Const param8 As String = "-уулу"
    Const param9 As String = "-кызы"
    Const param10 As String = "-гызы"
     
    If LCase(Right(patronymic, Len(param1))) = param1 Then
        suffix = Mid(patronymic, InStr(LCase(patronymic), param1), Len(param1))
        patronymic = Replace(patronymic, suffix, "")
        PatronymicSuffixRegex = True
        Exit Function
    End If
    If LCase(Right(patronymic, Len(param2))) = param2 Then
        suffix = Mid(patronymic, InStr(LCase(patronymic), param2), Len(param2))
        patronymic = Replace(patronymic, suffix, "")
        PatronymicSuffixRegex = True
        Exit Function
    End If
    If LCase(Right(patronymic, Len(param3))) = param3 Then
        suffix = Mid(patronymic, InStr(LCase(patronymic), param3), Len(param3))
        patronymic = Replace(patronymic, suffix, "")
        PatronymicSuffixRegex = True
        Exit Function
    End If
    If LCase(Right(patronymic, Len(param4))) = param4 Then
        suffix = Mid(patronymic, InStr(LCase(patronymic), param4), Len(param4))
        patronymic = Replace(patronymic, suffix, "")
        PatronymicSuffixRegex = True
        Exit Function
    End If
    If LCase(Right(patronymic, Len(param5))) = param5 Then
        suffix = Mid(patronymic, InStr(LCase(patronymic), param5), Len(param5))
        patronymic = Replace(patronymic, suffix, "")
        PatronymicSuffixRegex = True
        Exit Function
    End If
    If LCase(Right(patronymic, Len(param6))) = param6 Then
        suffix = Mid(patronymic, InStr(LCase(patronymic), param6), Len(param6))
        patronymic = Replace(patronymic, suffix, "")
        PatronymicSuffixRegex = True
        Exit Function
    End If
    If LCase(Right(patronymic, Len(param7))) = param7 Then
        suffix = Mid(patronymic, InStr(LCase(patronymic), param7), Len(param7))
        patronymic = Replace(patronymic, suffix, "")
        PatronymicSuffixRegex = True
        Exit Function
    End If
    If LCase(Right(patronymic, Len(param8))) = param8 Then
        suffix = Mid(patronymic, InStr(LCase(patronymic), param8), Len(param8))
        patronymic = Replace(patronymic, suffix, "")
        PatronymicSuffixRegex = True
        Exit Function
    End If
    If LCase(Right(patronymic, Len(param9))) = param9 Then
        suffix = Mid(patronymic, InStr(LCase(patronymic), param9), Len(param9))
        patronymic = Replace(patronymic, suffix, "")
        PatronymicSuffixRegex = True
        Exit Function
    End If
    If LCase(Right(patronymic, Len(param10))) = param10 Then
        suffix = Mid(patronymic, InStr(LCase(patronymic), param10), Len(param10))
        patronymic = Replace(patronymic, suffix, "")
        PatronymicSuffixRegex = True
        Exit Function
    End If
    PatronymicSuffixRegex = False
End Function

' Используется для разбивки составных отчеств, к примеру тюркские варианты Салим-оглы или Салим-кызы
' fullPatronymic -  Полное отчество. Примеры: "Салим-оглы", "Салим Оглы", "Ибн Салим"
' prefix -          Возвращает префиксную часть составного отчества, которая не должна склоняться
' patronymic -      Возвращает основную часть отчества, которая может склоняться.
' suffix -          Возвращает суффиксную часть состаного отчества, которая не должна склоняться
Public Sub SplitPatronymic(fullPatronymic As String, ByRef Prefix As String, ByRef patronymic As String, ByRef suffix As String)
    If fullPatronymic = "" Then
        Prefix = ""
        patronymic = fullPatronymic
        suffix = ""
        Exit Sub
    End If
    patronymic = fullPatronymic
    If Not PatronymicPrefixRegex(patronymic, Prefix) Then
        Prefix = ""
    End If

    If Not PatronymicSuffixRegex(patronymic, suffix) Then
        suffix = ""
    End If
End Sub

' Родительный, Кого? Чего? (нет)
' name -        имя, для склонения
' isFeminine -  истина, для женских имен
' Is_Shorten -  Сократить ли имя, к примеру, Иван будет И.
' Возвращает результат склонения
Public Function DeclineNameGenitive(Name As String, isFeminine As Boolean, is_shorten As Boolean) As String
    Dim temp As String

    If IsShorten(Name) Then
        DeclineNameGenitive = Name
        Exit Function
    End If

    If is_shorten Then
        Name = shorten(Name)
    Else
        temp = Name
        Dim Str As String
        Str = LCase(SubstringRight(Name, 3))
        
        Select Case Str
            Case "лев":
                Name = SetEndS(Name, 2, "ьва")
        End Select

        If Name = temp Then
            Select Case SubstringRight(Name, 2)
                Case "ей", "ий", "ай":
                    Name = SetEnd(Name, "я")
                Case "ел":
                    Name = SetEnd(Name, "ла")
                Case "ец":
                    Name = SetEnd(Name, "ца")
                Case "га", "жа", "ка", "ха", "ча", "ща":
                    Name = SetEnd(Name, "и")
            End Select
        End If

        If Name = temp Then
            Select Case SubstringRight(Name, 1)
                Case "а":
                    Name = SetEnd(Name, "ы")
                Case "е", "ё", "и", "о", "у", "э", "ю":
                     
                Case "я":
                    Name = SetEnd(Name, "и")
                Case "ь":
                    Name = SetEnd(Name, IIf(isFeminine, "и", "я"))
                Case Else
                    If Not isFeminine Then
                        Name = Name & "а"
                    End If
            End Select
        End If
    End If
    DeclineNameGenitive = Name
End Function

' Дательный, Кому? Чему? (дам)
' name          Имя, для склонения
' isFeminine    Истина, для женских имен
' shorten       Сократить ли имя, к примеру, Иван бедет И.
' Возвращает результат склонения
Public Function DeclineNameDative(Name As String, isFeminine As Boolean, is_shorten As Boolean) As String
    Dim temp As String

    If IsShorten(Name) Then
        DeclineNameDative = Name
        Exit Function
    End If

    If is_shorten Then
        Name = shorten(Name)
    Else
        temp = Name
        Dim Str As String
        Str = LCase(SubstringRight(Name, 3))
        
        Select Case Str
            Case "лев":
                Name = SetEndS(Name, 2, "ьву")
        End Select

        If Name = temp Then
            Select Case SubstringRight(Name, 2)
                Case "ей", "ий", "ай":
                    Name = SetEnd(Name, "ю")
                Case "ел":
                    Name = SetEnd(Name, "лу")
                Case "ец":
                    Name = SetEnd(Name, "цу")
                Case "ия":
                    Name = SetEnd(Name, "ии")
            End Select
        End If

        If Name = temp Then
            Select Case SubstringRight(Name, 1)
                Case "а", "я":
                    Name = SetEnd(Name, "е")
                Case "е", "ё", "и", "о", "у", "э", "ю":
                      
                Case "ь":
                    Name = SetEnd(Name, IIf(isFeminine, "и", "ю"))
                Case Else
                    If Not isFeminine Then
                        Name = Name & "у"
                    End If
            End Select
        End If
    End If
    DeclineNameDative = Name
End Function

' Винительный, Кого? Что? (вижу)
' name          Имя, для склонения
' isFeminine    Истина, для женских имен
' shorten       Сократить ли имя, к примеру, Иван бедет И.
' Возвращает результат склонения
Public Function DeclineNameAccusative(Name As String, isFeminine As Boolean, is_shorten As Boolean) As String
    Dim temp As String

    If IsShorten(Name) Then
        DeclineNameAccusative = Name
        Exit Function
    End If

    If is_shorten Then
        Name = shorten(Name)
    Else
        temp = Name
        Dim Str As String
        Str = LCase(SubstringRight(Name, 3))
        
        Select Case Str
            Case "лев":
                Name = SetEndS(Name, 2, "ьва")
        End Select

        If Name = temp Then
            Select Case SubstringRight(Name, 2)
                Case "ей", "ий", "ай":
                    Name = SetEnd(Name, "я")
                Case "ел":
                    Name = SetEnd(Name, "ла")
                Case "ец":
                    Name = SetEnd(Name, "ца")
            End Select
        End If

        If Name = temp Then
            Select Case SubstringRight(Name, 1)
                Case "а":
                    Name = SetEnd(Name, "у")
                Case "е", "ё", "и", "о", "у", "э", "ю":
                      
                Case "я":
                    Name = SetEnd(Name, "ю")
                Case "ь":
                    If Not isFeminine Then
                        Name = SetEnd(Name, "я")
                    End If
                Case Else
                    If Not isFeminine Then
                        Name = Name & "а"
                    End If
            End Select
        End If
    End If
    DeclineNameAccusative = Name
End Function

' Творительный, Кем? Чем? (горжусь)
' name          Имя, для склонения
' isFeminine    Истина, для женских имен
' shorten       Сократить ли имя, к примеру, Иван бедет И.
' Возвращает результат склонения
Public Function DeclineNameInstrumental(Name As String, isFeminine As Boolean, is_shorten As Boolean) As String
    Dim temp As String

    If IsShorten(Name) Then
        DeclineNameInstrumental = Name
        Exit Function
    End If

    If is_shorten Then
        Name = shorten(Name)
    Else
        temp = Name
        Dim Str As String
        Str = LCase(SubstringRight(Name, 3))
        
        Select Case Str
            Case "лев":
                Name = SetEndS(Name, 2, "ьвом")
        End Select

        If Name = temp Then
            Select Case SubstringRight(Name, 2)
                Case "ей", "ий", "ай":
                    Name = SetEndS(Name, 1, "ем")
                Case "ел":
                    Name = SetEndS(Name, 2, "лом")
                Case "ец":
                    Name = SetEndS(Name, 2, "цом")
                Case "жа", "ца", "ча", "ша", "ща":
                    Name = SetEndS(Name, 1, "ей")
            End Select
        End If

        If Name = temp Then
            Select Case SubstringRight(Name, 1)
                Case "а":
                    Name = SetEndS(Name, 1, "ой")
                Case "е", "ё", "и", "о", "у", "э", "ю":
                      
                Case "я":
                    Name = SetEndS(Name, 1, "ей")
                Case "ь":
                    Name = SetEndS(Name, 1, IIf(isFeminine, "ью", "ем"))
                Case Else
                    If Not isFeminine Then
                        Name = Name & "ом"
                    End If
            End Select
        End If
    End If
    DeclineNameInstrumental = Name
End Function

' Предложный, О ком? О чем? (думаю)
' name          Имя, для склонения
' isFeminine    Истина, для женских имен
' shorten       Сократить ли имя, к примеру, Иван бедет И.
' Возвращает результат склонения
Public Function DeclineNamePrepositional(Name As String, isFeminine As Boolean, is_shorten As Boolean) As String
    Dim temp As String

    If IsShorten(Name) Then
        DeclineNamePrepositional = Name
        Exit Function
    End If

    If is_shorten Then
        Name = shorten(Name)
    Else
        temp = Name
        Dim Str As String
        Str = LCase(SubstringRight(Name, 3))
        
        Select Case Str
            Case "лев":
                Name = SetEndS(Name, 2, "ьве")
        End Select

        If Name = temp Then
            Select Case SubstringRight(Name, 2)
                Case "ей", "ай":
                    Name = SetEnd(Name, "е")
                Case "ий":
                    Name = SetEnd(Name, "и")
                Case "ел":
                    Name = SetEnd(Name, "ле")
                Case "ец":
                    Name = SetEnd(Name, "це")
                Case "ия":
                    Name = SetEnd(Name, "ии")
            End Select
        End If

        If Name = temp Then
            Select Case SubstringRight(Name, 1)
                Case "а", "я":
                    Name = SetEnd(Name, "е")
                Case "е", "ё", "и", "о", "у", "э", "ю":
                       
                Case "ь":
                    Name = SetEnd(Name, IIf(isFeminine, "и", "е"))
                Case Else
                    If Not isFeminine Then
                        Name = Name & "е"
                    End If
            End Select
        End If
    End If
    DeclineNamePrepositional = Name
End Function

' Родительный, Кого? Чего? (нет)
' patronymic    Отчество, для склонения
' isFeminine    Истина, для женских имен
' shorten       Сократить ли отчество, к примеру, Иванович бедет И.
' Возвращает результат склонения
Public Function DeclinePatronymicGenitive(patronymic As String, isFeminine As Boolean, is_shorten As Boolean) As String
    Dim temp As String

    If IsShorten(patronymic) Then
        DeclinePatronymicGenitive = patronymic
        Exit Function
    End If

    If is_shorten Then
        patronymic = shorten(patronymic)
    Else
        Select Case SubstringRight(patronymic, 1)
            Case "а":
                patronymic = SetEnd(patronymic, "ы")
            Case "е", "ё", "и", "о", "у", "э", "ю":
                  
            Case "я":
                patronymic = SetEnd(patronymic, "и")
            Case "ь":
                patronymic = SetEnd(patronymic, IIf(isFeminine, "и", "я"))
            Case Else
                If Not isFeminine Then
                    patronymic = patronymic & "а"
                End If
        End Select
    End If
    DeclinePatronymicGenitive = patronymic
End Function

' Дательный, Кому? Чему? (дам)
' patronymic    Отчество, для склонения
' isFeminine    Истина, для женских имен
' shorten       Сократить ли отчество, к примеру, Иванович бедет И.
' Возвращает результат склонения
Public Function DeclinePatronymicDative(patronymic As String, isFeminine As Boolean, is_shorten As Boolean) As String
    Dim temp As String

    If IsShorten(patronymic) Then
        DeclinePatronymicDative = patronymic
        Exit Function
    End If

    If is_shorten Then
        patronymic = shorten(patronymic)
    Else
        Select Case SubstringRight(patronymic, 1)
            Case "а", "я":
                patronymic = SetEnd(patronymic, "е")
            Case "е", "ё", "и", "о", "у", "э", "ю":
                  
            Case "ь":
                patronymic = SetEnd(patronymic, IIf(isFeminine, "и", "ю"))
            Case Else
                If Not isFeminine Then
                    patronymic = patronymic & "у"
                End If
        End Select
    End If
    DeclinePatronymicDative = patronymic
End Function

' Винительный, Кого? Чего? (вижу)
' patronymic    Отчество, для склонения
' isFeminine    Истина, для женских имен
' shorten       Сократить ли отчество, к примеру, Иванович бедет И.
' Возвращает результат склонения
Public Function DeclinePatronymicAccusative(patronymic As String, isFeminine As Boolean, is_shorten As Boolean) As String
    Dim temp As String

    If IsShorten(patronymic) Then
        DeclinePatronymicAccusative = patronymic
        Exit Function
    End If

    If is_shorten Then
        patronymic = shorten(patronymic)
    Else
        Select Case SubstringRight(patronymic, 1)
            Case "а":
                patronymic = SetEnd(patronymic, "у")
            Case "е", "ё", "и", "о", "у", "э", "ю":
                  
            Case "я":
                patronymic = SetEnd(patronymic, "ю")
            Case "ь":
                If Not isFeminine Then patronymic = SetEnd(patronymic, "я")
            Case Else
                If Not isFeminine Then
                    patronymic = patronymic & "а"
                End If
        End Select
    End If
    DeclinePatronymicAccusative = patronymic
End Function

' Творительный, Кем? Чем? (горжусь)
' patronymic    Отчество, для склонения
' isFeminine    Истина, для женских имен
' shorten       Сократить ли отчество, к примеру, Иванович бедет И.
' Возвращает результат склонения
Public Function DeclinePatronymicInstrumental(patronymic As String, isFeminine As Boolean, is_shorten As Boolean) As String
    Dim temp As String

    If IsShorten(patronymic) Then
        DeclinePatronymicInstrumental = patronymic
        Exit Function
    End If

    If is_shorten Then
        patronymic = shorten(patronymic)
    Else
        temp = patronymic
        
        Select Case SubstringRight(patronymic, 2)
            Case "ич":
                patronymic = patronymic & IIf(LCase(patronymic) = "ильич", "ом", "ем")
            Case "на":
                patronymic = SetEndS(patronymic, 2, "ной")
        End Select
        
        If patronymic = temp Then
            Select Case SubstringRight(patronymic, 1)
                Case "а":
                    patronymic = SetEndS(patronymic, 1, "ой")
                Case "е", "ё", "и", "о", "у", "э", "ю":
                      
                Case "я":
                    patronymic = SetEndS(patronymic, 1, "ей")
                Case "ь":
                    patronymic = SetEndS(patronymic, 1, IIf(isFeminine, "ью", "ем"))
                Case Else
                    If Not isFeminine Then
                        patronymic = patronymic & "ом"
                    End If
            End Select
        End If
    End If
    DeclinePatronymicInstrumental = patronymic
End Function

' Творительный, Кем? Чем? (горжусь)
' patronymic    Отчество, для склонения
' isFeminine    Истина, для женских имен
' shorten       Сократить ли отчество, к примеру, Иванович бедет И.
' Возвращает результат склонения
Public Function DeclinePatronymicPrepositional(patronymic As String, isFeminine As Boolean, is_shorten As Boolean) As String
    Dim temp As String

    If IsShorten(patronymic) Then
        DeclinePatronymicPrepositional = patronymic
        Exit Function
    End If

    If is_shorten Then
        patronymic = shorten(patronymic)
    Else
        Select Case SubstringRight(patronymic, 1)
            Case "а", "я":
                patronymic = SetEnd(patronymic, "е")
            Case "е", "ё", "и", "о", "у", "э", "ю":
                   
            Case "ь":
                patronymic = SetEnd(patronymic, IIf(isFeminine, "и", "е"))
            Case Else
                If Not isFeminine Then
                    patronymic = patronymic & "е"
                End If
        End Select
    End If
    DeclinePatronymicPrepositional = patronymic
End Function

' Родительный, Кого? Чего? (нет)
' surname    Фамилия, для склонения
' isFeminine    Истина, для женских фамилий
' Возвращает результат склонения
Public Function DeclineSurnameGenitive(surname As String, isFeminine As Boolean) As String
    Dim temp As String, end_s As String
    temp = surname
    end_s = SubstringRight(surname, 3)
 
    If Not isFeminine Then
         Select Case end_s
            Case "жий", "ний", "ций", "чий", "ший", "щий":
                surname = SetEndS(surname, 2, "его")
            Case "лец":
                surname = SetEndS(surname, 2, "ьца")
            Case "нок":
                surname = SetEnd(surname, "нка")
        End Select
    Else
        Select Case end_s
            Case "ова", "ева", "ина", "ына"
                surname = SetEndS(surname, 1, "ой")
            Case "жая", "цая", "чая", "шая", "щая":
                surname = SetEndS(surname, 2, "ей")
             Case "ска", "цка":
                surname = SetEndS(surname, 1, "ой")
        End Select
    End If
    
    If surname <> temp Then
        DeclineSurnameGenitive = surname
        Exit Function
    End If
    
    end_s = SubstringRight(surname, 2)
    
    Select Case end_s
        Case "га", "жа", "ка", "ха", "ча", "ша", "ща"
            surname = SetEndS(surname, 1, "и")
    End Select
    
    If surname <> temp Then
        DeclineSurnameGenitive = surname
        Exit Function
    End If
    
    If Not isFeminine Then
         Select Case end_s
            Case "ок":
                surname = SetEndS(surname, 1, "ка")
            Case "ёк", "ек":
                surname = SetEndS(surname, 2, "ька")
            Case "ец":
                surname = SetEndS(surname, 2, "ца")
            Case "ий", "ый", "ой":
                If Len(surname) > 4 Then
                    surname = SetEndS(surname, 2, "ого")
                End If
            Case "ей":
                If LCase(surname) = "соловей" Or LCase(surname) = "воробей" Then
                    surname = SetEndS(surname, 2, "ья")
                Else
                    surname = SetEndS(surname, 2, "ея")
                End If
        End Select
    Else
        Select Case end_s
            Case "ая"
                surname = SetEndS(surname, 2, "ой")
            Case "яя":
                surname = SetEndS(surname, 2, "ей")
        End Select
    End If
    
    If surname <> temp Then
        DeclineSurnameGenitive = surname
        Exit Function
    End If
    
    end_s = SubstringRight(surname, 1)
    
     If Not isFeminine Then
         Select Case end_s
            Case "а":
                Select Case Mid(surname, Len(surname) - 1, 1)               '2 буква с конца
                    Case "а", "е", "ё", "и", "о", "у", "э", "ы", "ю", "я":
                        
                    Case Else:
                        surname = SetEndS(surname, 1, "ы")
                End Select
            Case "я":
                surname = SetEndS(surname, 1, "и")
            Case "б", "в", "г", "д", "ж", "з", "к", "л", "м", "н", "п", "р", "с", "т", "ф", "ц", "ч", "ш", "щ":
                surname = surname & "а"
            Case "х":
                If Right(surname, 2) <> "их" And Right(surname, 2) <> "ых" Then
                    surname = surname & "а"
                End If
            Case "ь", "й":
                surname = SetEndS(surname, 1, "я")
        End Select
    Else
        Select Case end_s
            Case "а":
                Select Case Mid(surname, Len(surname) - 1, 1)               '2 буква с конца
                    Case "а", "е", "ё", "и", "о", "у", "э", "ы", "ю", "я":
                        
                    Case Else:
                        surname = SetEndS(surname, 1, "ы")
                End Select
            Case "я":
                surname = SetEndS(surname, 1, "и")
        End Select
    End If
    
    DeclineSurnameGenitive = surname
End Function

' Дательный, Кому? Чему? (дам)
' surname    Фамилия, для склонения
' isFeminine    Истина, для женских фамилий
' Возвращает результат склонения
Public Function DeclineSurnameDative(surname As String, isFeminine As Boolean) As String
    Dim temp As String, end_s As String
    temp = surname
    end_s = SubstringRight(surname, 3)
 
    If Not isFeminine Then
         Select Case end_s
            Case "жий", "ний", "ций", "чий", "ший", "щий":
                surname = SetEndS(surname, 2, "ему")
            Case "лец":
                surname = SetEndS(surname, 2, "ьцу")
        End Select
    Else
        Select Case end_s
            Case "ова", "ева", "ина", "ына"
                surname = SetEndS(surname, 1, "ой")
            Case "жая", "цая", "чая", "шая", "щая":
                surname = SetEndS(surname, 2, "ей")
             Case "ска", "цка":
                surname = SetEndS(surname, 1, "ой")
        End Select
    End If
    
    If surname <> temp Then
        DeclineSurnameDative = surname
        Exit Function
    End If
    
    end_s = SubstringRight(surname, 2)
    
    Select Case end_s
        Case "ия"
            surname = SetEndS(surname, 1, "и")
    End Select
    
    If surname <> temp Then
        DeclineSurnameDative = surname
        Exit Function
    End If
    
    If Not isFeminine Then
         Select Case end_s
            Case "ок":
                surname = SetEndS(surname, 2, "ку")
            Case "ёк", "ек":
                surname = SetEndS(surname, 2, "ьку")
            Case "ец":
                surname = SetEndS(surname, 2, "цу")
            Case "ий", "ый", "ой":
                If Len(surname) > 4 Then
                    surname = SetEndS(surname, 2, "ому")
                End If
            Case "ей":
                If LCase(surname) = "соловей" Or LCase(surname) = "воробей" Then
                    surname = SetEndS(surname, 2, "ью")
                Else
                    surname = SetEndS(surname, 2, "ею")
                End If
        End Select
    Else
        Select Case end_s
            Case "ая"
                surname = SetEndS(surname, 2, "ой")
            Case "яя":
                surname = SetEndS(surname, 2, "ей")
        End Select
    End If
    
    If surname <> temp Then
        DeclineSurnameDative = surname
        Exit Function
    End If
    
    end_s = SubstringRight(surname, 1)
    
    If Not isFeminine Then
         Select Case end_s
            Case "а":
                Select Case Mid(surname, Len(surname) - 1, 1)               '2 буква с конца
                    Case "а", "е", "ё", "и", "о", "у", "э", "ы", "ю", "я":
                        
                    Case Else:
                        surname = SetEndS(surname, 1, "е")
                End Select
            Case "я":
                surname = SetEndS(surname, 1, "е")
            Case "б", "в", "г", "д", "ж", "з", "к", "л", "м", "н", "п", "р", "с", "т", "ф", "ц", "ч", "ш", "щ":
                surname = surname & "у"
            Case "х":
                If Right(surname, 2) <> "их" And Right(surname, 2) <> "ых" Then
                    surname = surname & "у"
                End If
            Case "ь", "й":
                surname = SetEndS(surname, 1, "ю")
        End Select
    Else
        Select Case end_s
            Case "а":
                Select Case Mid(surname, Len(surname) - 1, 1)               '2 буква с конца
                    Case "а", "е", "ё", "и", "о", "у", "э", "ы", "ю", "я":
                        
                    Case Else:
                        surname = SetEndS(surname, 1, "е")
                End Select
            Case "я":
                surname = SetEndS(surname, 1, "е")
        End Select
    End If
    
    DeclineSurnameDative = surname
End Function

' Винительный, Кого? Чего? (вижу)
' surname    Фамилия, для склонения
' isFeminine    Истина, для женских фамилий
' Возвращает результат склонения
Public Function DeclineSurnameAccusative(surname As String, isFeminine As Boolean) As String
    Dim temp As String, end_s As String
    temp = surname
    end_s = SubstringRight(surname, 3)
 
    If Not isFeminine Then
         Select Case end_s
            Case "жий", "ний", "ций", "чий", "ший", "щий":
                surname = SetEndS(surname, 2, "его")
            Case "лец":
                surname = SetEndS(surname, 2, "ьца")
        End Select
    Else
        Select Case end_s
            Case "ова", "ева", "ина", "ына"
                surname = SetEnd(surname, "у")
             Case "ска", "цка":
                surname = SetEndS(surname, 1, "ую")
        End Select
    End If
    
    If surname <> temp Then
        DeclineSurnameAccusative = surname
        Exit Function
    End If
    
    end_s = SubstringRight(surname, 2)
        
    If Not isFeminine Then
         Select Case end_s
            Case "ок":
                surname = SetEnd(surname, "ка")
            Case "ёк", "ек":
                surname = SetEndS(surname, 2, "ька")
            Case "ец":
                surname = SetEndS(surname, 2, "ца")
            Case "ий", "ый", "ой":
                If Len(surname) > 4 Then
                    surname = SetEndS(surname, 2, "ого")
                End If
            Case "ей":
                If LCase(surname) = "соловей" Or LCase(surname) = "воробей" Then
                    surname = SetEndS(surname, 2, "ья")
                Else
                    surname = SetEndS(surname, 2, "ея")
                End If
        End Select
    Else
        Select Case end_s
            Case "ая"
                surname = SetEnd(surname, "ую")
            Case "яя":
                surname = SetEnd(surname, "юю")
        End Select
    End If
    
    If surname <> temp Then
        DeclineSurnameAccusative = surname
        Exit Function
    End If
    
    end_s = SubstringRight(surname, 1)
    
    If Not isFeminine Then
         Select Case end_s
            Case "а":
                Select Case Mid(surname, Len(surname) - 1, 1)               '2 буква с конца
                    Case "а", "е", "ё", "и", "о", "у", "э", "ы", "ю", "я":
                        
                    Case Else:
                        surname = SetEnd(surname, "у")
                End Select
            Case "я":
                surname = SetEnd(surname, "ю")
            Case "б", "в", "г", "д", "ж", "з", "к", "л", "м", "н", "п", "р", "с", "т", "ф", "ц", "ч", "ш", "щ":
                surname = surname & "а"
            Case "х":
                If Right(surname, 2) <> "их" And Right(surname, 2) <> "ых" Then
                    surname = surname & "а"
                End If
            Case "ь", "й":
                surname = SetEnd(surname, "я")
        End Select
    Else
        Select Case end_s
            Case "а":
                Select Case Mid(surname, Len(surname) - 1, 1)               '2 буква с конца
                    Case "а", "е", "ё", "и", "о", "у", "э", "ы", "ю", "я":
                        
                    Case Else:
                        surname = SetEnd(surname, "у")
                End Select
            Case "я":
                surname = SetEnd(surname, "ю")
        End Select
    End If
    
    DeclineSurnameAccusative = surname
End Function

' Творительный, Кем? Чем? (горжусь)
' surname    Фамилия, для склонения
' isFeminine    Истина, для женских фамилий
' Возвращает результат склонения
Public Function DeclineSurnameInstrumental(surname As String, isFeminine As Boolean) As String
    Dim temp As String, end_s As String
    temp = surname
    end_s = SubstringRight(surname, 3)
 
    If Not isFeminine Then
         Select Case end_s
            Case "лец":
                surname = SetEndS(surname, 2, "ьцом")
            Case "бец":
                surname = SetEndS(surname, 2, "цем")
            Case "кой":
                surname = SetEndS(surname, 2, "им")
        End Select
    Else
        Select Case end_s
            Case "жая", "цая", "чая", "шая", "щая":
                surname = SetEnd(surname, "ей")
            Case "ска", "цка":
                surname = SetEndS(surname, 1, "ой")
            Case "еца", "ица", "аца", "ьца":
                surname = SetEndS(surname, 1, "ей")
        End Select
    End If
    
    If surname <> temp Then
        DeclineSurnameInstrumental = surname
        Exit Function
    End If
    
    end_s = SubstringRight(surname, 2)
        
    If Not isFeminine Then
         Select Case end_s
            Case "ок":
                surname = SetEndS(surname, 2, "ком")
            Case "ёк", "ек":
                surname = SetEndS(surname, 2, "ьком")
            Case "ец":
                surname = SetEndS(surname, 2, "цом")
            Case "ий":
                If Len(surname) > 4 Then
                    surname = SetEndS(surname, 2, "им")
                End If
            Case "ый", "ой":
                If Len(surname) > 4 Then
                    surname = SetEnd(surname, "ым")
                End If
            Case "ей":
                If LCase(surname) = "соловей" Or LCase(surname) = "воробей" Then
                    surname = SetEndS(surname, 2, "ьем")
                Else
                    surname = SetEndS(surname, 2, "еем")
                End If
            Case "оч", "ич", "иц", "ьц", "ьш", "еш", "ыш", "яц":
                surname = surname & "ем"
            Case "ин", "ын", "ен", "эн", "ов", "ев", "ёв", "ун":
                 If LCase(surname) <> "дарвин" And LCase(surname) <> "франклин" And LCase(surname) <> "чаплин" And LCase(surname) <> "грин" Then
                    surname = surname + "ым"
                End If
            Case "жа", "ца", "ча", "ша", "ща":
                surname = SetEndS(surname, 1, "ей")
        End Select
    Else
        Select Case end_s
            Case "ая"
                surname = SetEnd(surname, "ой")
            Case "яя":
                surname = SetEnd(surname, "ей")
        End Select
    End If
    
    If surname <> temp Then
        DeclineSurnameInstrumental = surname
        Exit Function
    End If
    
    end_s = SubstringRight(surname, 1)
    
    If Not isFeminine Then
         Select Case end_s
            Case "а":
                Select Case Mid(surname, Len(surname) - 1, 1)               '2 буква с конца
                    Case "а", "е", "ё", "и", "о", "у", "э", "ы", "ю", "я":
                        
                    Case Else:
                        surname = SetEndS(surname, 1, "ой")
                End Select
            Case "я":
                surname = SetEndS(surname, 1, "ей")
            Case "б", "в", "г", "д", "ж", "з", "к", "л", "м", "н", "п", "р", "с", "т", "ф", "ц", "ч", "ш":
                surname = surname & "ом"
            Case "х":
                If Right(surname, 2) <> "их" And Right(surname, 2) <> "ых" Then
                    surname = surname & "ом"
                End If
            Case "щ":
                surname = surname & "ем"
            Case "ь", "й":
                surname = SetEndS(surname, 1, "ем")
        End Select
    Else
        Select Case end_s
            Case "а":
                Select Case Mid(surname, Len(surname) - 1, 1)               '2 буква с конца
                    Case "а", "е", "ё", "и", "о", "у", "э", "ы", "ю", "я":
                        
                    Case Else:
                        surname = SetEndS(surname, 1, "ой")
                End Select
            Case "я":
                surname = SetEndS(surname, 1, "ей")
        End Select
    End If
    
    DeclineSurnameInstrumental = surname
End Function

' Предложный, О ком? О чем? (думаю)
' surname    Фамилия, для склонения
' isFeminine    Истина, для женских фамилий
' Возвращает результат склонения
Public Function DeclineSurnamePrepositional(surname As String, isFeminine As Boolean) As String
    Dim temp As String, end_s As String
    temp = surname
    end_s = SubstringRight(surname, 3)
 
    If Not isFeminine Then
         Select Case end_s
            Case "жий", "ний", "ций", "чий", "ший", "щий":
                surname = SetEndS(surname, 2, "ем")
            Case "лец":
                surname = SetEndS(surname, 2, "ьце")
        End Select
    Else
        Select Case end_s
            Case "ова", "ева", "ина", "ына":
                surname = SetEndS(surname, 1, "ой")
            Case "жая", "цая", "чая", "шая", "щая":
                surname = SetEnd(surname, "ей")
            Case "ска", "цка":
                surname = SetEndS(surname, 1, "ой")
        End Select
    End If
    
    If surname <> temp Then
        DeclineSurnamePrepositional = surname
        Exit Function
    End If
    
    end_s = SubstringRight(surname, 2)
    
    Select Case end_s
        Case "ия":
            surname = SetEnd(surname, "и")
    End Select
    
    If surname <> temp Then
        DeclineSurnamePrepositional = surname
        Exit Function
    End If
        
    If Not isFeminine Then
         Select Case end_s
            Case "ок":
                surname = SetEnd(surname, "ке")
            Case "ёк", "ек":
                surname = SetEndS(surname, 2, "ьке")
            Case "ец":
                surname = SetEndS(surname, 2, "це")
            Case "ий", "ый", "ой":
                If Len(surname) > 4 Then
                    surname = SetEnd(surname, "ом")
                End If
            Case "ей":
                If LCase(surname) = "соловей" Or LCase(surname) = "воробей" Then
                    surname = SetEndS(surname, 2, "ье")
                Else
                    surname = SetEndS(surname, 2, "ее")
                End If
        End Select
    Else
        Select Case end_s
            Case "ая"
                surname = SetEnd(surname, "ой")
            Case "яя":
                surname = SetEnd(surname, "ей")
        End Select
    End If
    
    If surname <> temp Then
        DeclineSurnamePrepositional = surname
        Exit Function
    End If
    
    end_s = SubstringRight(surname, 1)
    
    If Not isFeminine Then
         Select Case end_s
            Case "а":
                Select Case Mid(surname, Len(surname) - 1, 1)               '2 буква с конца
                    Case "а", "е", "ё", "и", "о", "у", "э", "ы", "ю", "я":
                        
                    Case Else:
                        surname = SetEnd(surname, "е")
                End Select
            Case "я":
                surname = SetEnd(surname, "е")
            Case "б", "в", "г", "д", "ж", "з", "к", "л", "м", "н", "п", "р", "с", "т", "ф", "ц", "ч", "ш", "щ":
                surname = surname & "е"
            Case "х":
                If Right(surname, 2) <> "их" And Right(surname, 2) <> "ых" Then
                    surname = surname & "е"
                End If
            Case "ь", "й":
                surname = SetEnd(surname, "е")
        End Select
    Else
        Select Case end_s
            Case "а":
                Select Case Mid(surname, Len(surname) - 1, 1)               '2 буква с конца
                    Case "а", "е", "ё", "и", "о", "у", "э", "ы", "ю", "я":
                        
                    Case Else:
                        surname = SetEnd(surname, "е")
                End Select
            Case "я":
                surname = SetEnd(surname, "е")
        End Select
    End If
    
    DeclineSurnamePrepositional = surname
End Function

' Склоняет фамилию в указанный падеж
' surname       фамилия в иминительном падеже
' p_case        Падеж в каторый нужно склонять фамилию, где 1 - именительный, 2 - родительный, 3 - дательный, 4 - винительный, 5 - творительный, 6 - предложный
' isFeminine    Иститна для женских фамилий
' Возвращает результат склонения
Public Function DeclineSurname(surname As String, p_case As Long, isFeminine As Boolean) As String
    Dim Result As String
    Result = surname
    
    If Len(surname) <= 1 Or p_case < 2 Or p_case > 6 Then
        Result = surname
        DeclineSurname = Result
        Exit Function
    End If
    
    Select Case p_case
        Case 2:
            Result = DeclineSurnameGenitive(surname, isFeminine)
        Case 3:
            Result = DeclineSurnameDative(surname, isFeminine)
        Case 4:
            Result = DeclineSurnameAccusative(surname, isFeminine)
        Case 5:
            Result = DeclineSurnameInstrumental(surname, isFeminine)
        Case 6:
            Result = DeclineSurnamePrepositional(surname, isFeminine)
    End Select
    
    DeclineSurname = Result
End Function
'*********************************************************


