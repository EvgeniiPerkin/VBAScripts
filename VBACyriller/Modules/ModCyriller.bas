Attribute VB_Name = "ModCyriller"

Option Explicit

Public Type CasesEnumAndString
    Key As CasesEnum
    Value As String
End Type

Public Type StringAndString
    Key As String
    Value As String
End Type
 
Public Enum CasesEnum
    Nominative = 1      ' Именительный, Кто? Что?   (есть)
    Genitive = 2        ' Родительный, Кого? Чего?  (нет)
    Dative = 3          ' Дательный, Кому? Чего?    (дам)
    Accusative = 4      ' Винительный, Кого? Что?   (вижу)
    Instrumental = 5    ' Творительный, Кем? Чем?   (горжусь)
    Prepositional = 6   ' Предложный, О ком? О Чем? (думаю)
End Enum

Public Enum GendersEnum
    Undefined = 0       ' Неопределённый
    Masculine = 1       ' Мужской
    Feminine = 2        ' Женский
    Neuter = 3          ' Средний
End Enum

Public Function DeclineSeparateName(LastName As String, FirstName As String, SecondName As String, case_p As Long, is_shorten As Boolean) As String
    Dim Cyr As New CyrName
    Dim Result As New CyrResult
    Dim resultStr As String
    Set Result = Cyr.Decline_AllCases_MuchWord(LastName, FirstName, SecondName, , is_shorten)
        
    Select Case case_p
        Case 1:
            resultStr = Result.Именительный
        Case 2:
            resultStr = Result.Родительный
        Case 3:
            resultStr = Result.Дательный
        Case 4:
            resultStr = Result.Винительный
        Case 5:
            resultStr = Result.Творительный
        Case 6:
            resultStr = Result.Предложный
    End Select
    
    Set Result = Nothing
    Set Cyr = Nothing
    
    DeclineSeparateName = resultStr
End Function

Public Function DeclineMergeName(FullName As String, case_p As Long, is_shorten As Boolean) As String
    Dim Cyr As New CyrName
    Dim Result As New CyrResult
    Dim resultStr As String
    Set Result = Cyr.Decline_AllCases_OneWord(FullName, , is_shorten)
        
    Select Case case_p
        Case 1:
            resultStr = Result.Именительный
        Case 2:
            resultStr = Result.Родительный
        Case 3:
            resultStr = Result.Дательный
        Case 4:
            resultStr = Result.Винительный
        Case 5:
            resultStr = Result.Творительный
        Case 6:
            resultStr = Result.Предложный
    End Select
    
    Set Result = Nothing
    Set Cyr = Nothing
    
    DeclineMergeName = resultStr
End Function


'-----------------TESTING------------------------------
Public Sub GLOBAL_TEST()
    Dim Assert As New Assert
    
    Call TestIsShorten(Assert)
    Call TestShorten(Assert)
    Call TestSubstringRight(Assert)
    Call TestSetEndS(Assert)
    Call TestSetEnd(Assert)
    Call TestProperCase(Assert)
    Call TestPatronymicPrefixRegex(Assert)
    Call TestPatronymicSuffixRegex(Assert)
    Call TestSplitPatronymic(Assert)
    
    Call FeminineFullNameIsCorrectlyDeclined(Assert)
    Call MasculineFullNameIsCorrectlyDeclined(Assert)
    Call TestNameDeclension(Assert)
    Call TestSurnameDeclension(Assert)
    Call TestPatronymicDeclension(Assert)
    
    Assert.ResultAssert
End Sub

Private Sub TestIsShorten(Assert As Assert)
    Dim test As New CyrName
    Assert.NameMethod = "TestIsShorten"
    
    Call Assert.Equal_Boolean(True, test.IsShorten(""))
    Call Assert.Equal_Boolean(False, test.IsShorten("Петровна"))
    Call Assert.Equal_Boolean(True, test.IsShorten(""))
End Sub
Private Sub TestShorten(Assert As Assert)
    Dim test As New CyrName
    Assert.NameMethod = "TestShorten"
    
    Call Assert.Equal_String("", test.shorten("     "))
    Call Assert.Equal_String("", test.shorten(""))
    Call Assert.Equal_String("П.", test.shorten("Перкин"))
End Sub
Private Sub TestSubstringRight(Assert As Assert)
    Dim test As New CyrName
    Assert.NameMethod = "TestSubstringRight"
    Call Assert.Equal_String("н", test.SubstringRight("Перкин", 1))
    Call Assert.Equal_String("Перкни", test.SubstringRight("Перкни", 7))
End Sub
Private Sub TestSetEndS(Assert As Assert)
    Dim test As New CyrName
    Assert.NameMethod = "TestSetEndS"
    
    Call Assert.Equal_String("ПеркиСон", test.SetEndS("Перкин", 1, "Сон"))
    Call Assert.Equal_String("ПеркСон", test.SetEndS("Перкин", 2, "Сон"))
    Call Assert.Equal_String("ПерСон", test.SetEndS("Перкин", 3, "Сон"))
End Sub
Private Sub TestSetEnd(Assert As Assert)
    Dim test As New CyrName
    Assert.NameMethod = "TestSetEnd"
    
    Call Assert.Equal_String("ПерСон", test.SetEnd("Перкин", "Сон"))
    Call Assert.Equal_String("ПеСонч", test.SetEnd("Перкин", "Сонч"))
    Call Assert.Equal_String("ПСончи", test.SetEnd("Перкин", "Сончи"))
End Sub
Private Sub TestProperCase(Assert As Assert)
    Dim test As New CyrName
    Assert.NameMethod = "TestProperCase"
    
    Call Assert.Equal_String("Перкин", test.ProperCase("перкин"))
    Call Assert.Equal_String("Евгений", test.ProperCase("евгений"))
    Call Assert.Equal_String("Владимирович", test.ProperCase("владимирович"))
End Sub
Private Sub TestPatronymicPrefixRegex(Assert As Assert)
    Dim test As New CyrName
    Dim patronymic As String
    Dim Prefix As String
    Assert.NameMethod = "TestPatronymicPrefixRegex"
    
    patronymic = "Ибн Салим"
    Call Assert.Equal_Boolean(True, test.PatronymicPrefixRegex(patronymic, Prefix))
    Call Assert.Equal_String("Салим", patronymic)
    Call Assert.Equal_String("Ибн ", Prefix)
    
    patronymic = "ибн-Салим"
    Prefix = ""
    Call Assert.Equal_Boolean(True, test.PatronymicPrefixRegex(patronymic, Prefix))
    Call Assert.Equal_String("Салим", patronymic)
    Call Assert.Equal_String("ибн-", Prefix)
    
    patronymic = "о ибн-Салим"
    Prefix = ""
    Call Assert.Equal_Boolean(False, test.PatronymicPrefixRegex(patronymic, Prefix))
    Call Assert.Equal_String("о ибн-Салим", patronymic)
    Call Assert.Equal_String("", Prefix)
    
    patronymic = "ибн-Салим Ука"
    Prefix = ""
    Call Assert.Equal_Boolean(True, test.PatronymicPrefixRegex(patronymic, Prefix))
    Call Assert.Equal_String("Салим Ука", patronymic)
    Call Assert.Equal_String("ибн-", Prefix)
End Sub
Private Sub TestPatronymicSuffixRegex(Assert As Assert)
    Dim test As New CyrName
    Dim patronymic As String
    Dim suffix As String
    Assert.NameMethod = "TestPatronymicSuffixRegex"
    
    patronymic = "Салим Оглы"
    Call Assert.Equal_Boolean(True, test.PatronymicSuffixRegex(patronymic, suffix))
    Call Assert.Equal_String("Салим", patronymic)
    Call Assert.Equal_String(" Оглы", suffix)
    
    patronymic = "Салим-Оглы"
    suffix = ""
    Call Assert.Equal_Boolean(True, test.PatronymicSuffixRegex(patronymic, suffix))
    Call Assert.Equal_String("Салим", patronymic)
    Call Assert.Equal_String("-Оглы", suffix)
    
    patronymic = "Салим-кызы"
    suffix = ""
    Call Assert.Equal_Boolean(True, test.PatronymicSuffixRegex(patronymic, suffix))
    Call Assert.Equal_String("Салим", patronymic)
    Call Assert.Equal_String("-кызы", suffix)
    
    patronymic = "Салим кызы"
    suffix = ""
    Call Assert.Equal_Boolean(True, test.PatronymicSuffixRegex(patronymic, suffix))
    Call Assert.Equal_String("Салим", patronymic)
    Call Assert.Equal_String(" кызы", suffix)
    
    patronymic = "Салим-гызы"
    suffix = ""
    Call Assert.Equal_Boolean(True, test.PatronymicSuffixRegex(patronymic, suffix))
    Call Assert.Equal_String("Салим", patronymic)
    Call Assert.Equal_String("-гызы", suffix)
    
    patronymic = "Салим гызы"
    suffix = ""
    Call Assert.Equal_Boolean(True, test.PatronymicSuffixRegex(patronymic, suffix))
    Call Assert.Equal_String("Салим", patronymic)
    Call Assert.Equal_String(" гызы", suffix)
    
    patronymic = "Абдибаит-Уулу"
    suffix = ""
    Call Assert.Equal_Boolean(True, test.PatronymicSuffixRegex(patronymic, suffix))
    Call Assert.Equal_String("Абдибаит", patronymic)
    Call Assert.Equal_String("-Уулу", suffix)
    
    patronymic = "Абдибаит Уулу"
    suffix = ""
    Call Assert.Equal_Boolean(True, test.PatronymicSuffixRegex(patronymic, suffix))
    Call Assert.Equal_String("Абдибаит", patronymic)
    Call Assert.Equal_String(" Уулу", suffix)
         
    patronymic = "Абдибаит Уулу Мулу"
    suffix = ""
    Call Assert.Equal_Boolean(False, test.PatronymicSuffixRegex(patronymic, suffix))
    Call Assert.Equal_String("Абдибаит Уулу Мулу", patronymic)
    Call Assert.Equal_String("", suffix)
    
    patronymic = "Мулу Абдибаит Уулу"
    suffix = ""
    Call Assert.Equal_Boolean(True, test.PatronymicSuffixRegex(patronymic, suffix))
    Call Assert.Equal_String("Мулу Абдибаит", patronymic)
    Call Assert.Equal_String(" Уулу", suffix)
End Sub
Private Sub TestSplitPatronymic(Assert As Assert)
    Dim test As New CyrName
    Dim fullPatronymic As String
    Dim patronymic As String
    Dim suffix As String
    Dim Prefix As String
    
    Assert.NameMethod = "TestSplitPatronymic"
    
    fullPatronymic = "Ибн Салим"
    patronymic = ""
    Prefix = ""
    suffix = ""
    Call test.SplitPatronymic(fullPatronymic, Prefix, patronymic, suffix)
    Call Assert.Equal_String("Салим", patronymic)
    Call Assert.Equal_String("Ибн ", Prefix)
    Call Assert.Equal_String("", suffix)
    
    fullPatronymic = "ибн-Салим"
    patronymic = ""
    Prefix = ""
    suffix = ""
    Call test.SplitPatronymic(fullPatronymic, Prefix, patronymic, suffix)
    Call Assert.Equal_String("Салим", patronymic)
    Call Assert.Equal_String("ибн-", Prefix)
    Call Assert.Equal_String("", suffix)
    
    fullPatronymic = "о ибн-Салим"
    patronymic = ""
    Prefix = ""
    suffix = ""
    Call test.SplitPatronymic(fullPatronymic, Prefix, patronymic, suffix)
    Call Assert.Equal_String("о ибн-Салим", patronymic)
    Call Assert.Equal_String("", Prefix)
    Call Assert.Equal_String("", suffix)
    
    fullPatronymic = "ибн-Салим Ука"
    patronymic = ""
    Prefix = ""
    suffix = ""
    Call test.SplitPatronymic(fullPatronymic, Prefix, patronymic, suffix)
    Call Assert.Equal_String("Салим Ука", patronymic)
    Call Assert.Equal_String("ибн-", Prefix)
    Call Assert.Equal_String("", suffix)
    
    fullPatronymic = "Салим Оглы"
    patronymic = ""
    Prefix = ""
    suffix = ""
    Call test.SplitPatronymic(fullPatronymic, Prefix, patronymic, suffix)
    Call Assert.Equal_String("Салим", patronymic)
    Call Assert.Equal_String("", Prefix)
    Call Assert.Equal_String(" Оглы", suffix)
    
    fullPatronymic = "Салим-Оглы"
    patronymic = ""
    Prefix = ""
    suffix = ""
    Call test.SplitPatronymic(fullPatronymic, Prefix, patronymic, suffix)
    Call Assert.Equal_String("Салим", patronymic)
    Call Assert.Equal_String("", Prefix)
    Call Assert.Equal_String("-Оглы", suffix)
    
    fullPatronymic = "Салим-кызы"
    patronymic = ""
    Prefix = ""
    suffix = ""
    Call test.SplitPatronymic(fullPatronymic, Prefix, patronymic, suffix)
    Call Assert.Equal_String("Салим", patronymic)
    Call Assert.Equal_String("", Prefix)
    Call Assert.Equal_String("-кызы", suffix)
    
    fullPatronymic = "Салим кызы"
    patronymic = ""
    Prefix = ""
    suffix = ""
    Call test.SplitPatronymic(fullPatronymic, Prefix, patronymic, suffix)
    Call Assert.Equal_String("Салим", patronymic)
    Call Assert.Equal_String("", Prefix)
    Call Assert.Equal_String(" кызы", suffix)
    
    fullPatronymic = "Салим-гызы"
    patronymic = ""
    Prefix = ""
    suffix = ""
    Call test.SplitPatronymic(fullPatronymic, Prefix, patronymic, suffix)
    Call Assert.Equal_String("Салим", patronymic)
    Call Assert.Equal_String("", Prefix)
    Call Assert.Equal_String("-гызы", suffix)
    
    fullPatronymic = "Салим гызы"
    patronymic = ""
    Prefix = ""
    suffix = ""
    Call test.SplitPatronymic(fullPatronymic, Prefix, patronymic, suffix)
    Call Assert.Equal_String("Салим", patronymic)
    Call Assert.Equal_String("", Prefix)
    Call Assert.Equal_String(" гызы", suffix)
    
    fullPatronymic = "Абдибаит Уулу"
    patronymic = ""
    Prefix = ""
    suffix = ""
    Call test.SplitPatronymic(fullPatronymic, Prefix, patronymic, suffix)
    Call Assert.Equal_String("Абдибаит", patronymic)
    Call Assert.Equal_String("", Prefix)
    Call Assert.Equal_String(" Уулу", suffix)
    
    fullPatronymic = "Абдибаит-Уулу"
    patronymic = ""
    Prefix = ""
    suffix = ""
    Call test.SplitPatronymic(fullPatronymic, Prefix, patronymic, suffix)
    Call Assert.Equal_String("Абдибаит", patronymic)
    Call Assert.Equal_String("", Prefix)
    Call Assert.Equal_String("-Уулу", suffix)
    
    fullPatronymic = "Абдибаит Уулу Мулу"
    patronymic = ""
    Prefix = ""
    suffix = ""
    Call test.SplitPatronymic(fullPatronymic, Prefix, patronymic, suffix)
    Call Assert.Equal_String("Абдибаит Уулу Мулу", patronymic)
    Call Assert.Equal_String("", Prefix)
    Call Assert.Equal_String("", suffix)
    
    fullPatronymic = "Мулу Абдибаит Уулу"
    patronymic = ""
    Prefix = ""
    suffix = ""
    Call test.SplitPatronymic(fullPatronymic, Prefix, patronymic, suffix)
    Call Assert.Equal_String("Мулу Абдибаит", patronymic)
    Call Assert.Equal_String("", Prefix)
    Call Assert.Equal_String(" Уулу", suffix)
    
    fullPatronymic = "Ибн-Абдибаит Уулу"
    patronymic = ""
    Prefix = ""
    suffix = ""
    Call test.SplitPatronymic(fullPatronymic, Prefix, patronymic, suffix)
    Call Assert.Equal_String("Абдибаит", patronymic)
    Call Assert.Equal_String("Ибн-", Prefix)
    Call Assert.Equal_String(" Уулу", suffix)
End Sub

Private Sub FeminineFullNameIsCorrectlyDeclined(Assert As Assert)
    Assert.NameMethod = "FeminineFullNameIsCorrectlyDeclined"
    Dim Result As New CyrResult
    Dim CyrName As New CyrName
    
    Set Result = CyrName.Decline_AllCases_OneWord("Иванова Наталья Петровна", Feminine, False)
    
    Call Assert.Equal_String("Ивановой Натальи Петровны", Result.GetCase(Genitive))
    Call Assert.Equal_String("Ивановой Наталье Петровне", Result.GetCase(Dative))
    Call Assert.Equal_String("Иванову Наталью Петровну", Result.GetCase(Accusative))
    Call Assert.Equal_String("Ивановой Натальей Петровной", Result.GetCase(Instrumental))
    Call Assert.Equal_String("Ивановой Наталье Петровне", Result.GetCase(Prepositional))
    
    Set Result = CyrName.Decline_AllCases_OneWord("Сафаралиева Койкеб Кямил Кызы", Feminine, False)
    
    Call Assert.Equal_String("Сафаралиевой Койкеб Кямил Кызы", Result.GetCase(Genitive))
    Call Assert.Equal_String("Сафаралиевой Койкеб Кямил Кызы", Result.GetCase(Dative))
    Call Assert.Equal_String("Сафаралиеву Койкеб Кямил Кызы", Result.GetCase(Accusative))
    Call Assert.Equal_String("Сафаралиевой Койкеб Кямил Кызы", Result.GetCase(Instrumental))
    Call Assert.Equal_String("Сафаралиевой Койкеб Кямил Кызы", Result.GetCase(Prepositional))
    
    Set Result = CyrName.Decline_AllCases_OneWord("Сафаралиева Койкеб Кямил-Кызы", Feminine, False)
    
    Call Assert.Equal_String("Сафаралиевой Койкеб Кямил-Кызы", Result.GetCase(Genitive))
    Call Assert.Equal_String("Сафаралиевой Койкеб Кямил-Кызы", Result.GetCase(Dative))
    Call Assert.Equal_String("Сафаралиеву Койкеб Кямил-Кызы", Result.GetCase(Accusative))
    Call Assert.Equal_String("Сафаралиевой Койкеб Кямил-Кызы", Result.GetCase(Instrumental))
    Call Assert.Equal_String("Сафаралиевой Койкеб Кямил-Кызы", Result.GetCase(Prepositional))
    
    Set Result = CyrName.Decline_AllCases_OneWord("Иванова Наталья Петровна", Feminine, True)
    
    Call Assert.Equal_String("Ивановой Н. П.", Result.GetCase(Genitive))
    Call Assert.Equal_String("Ивановой Н. П.", Result.GetCase(Dative))
    Call Assert.Equal_String("Иванову Н. П.", Result.GetCase(Accusative))
    Call Assert.Equal_String("Ивановой Н. П.", Result.GetCase(Instrumental))
    Call Assert.Equal_String("Ивановой Н. П.", Result.GetCase(Prepositional))
    
    Set Result = CyrName.Decline_AllCases_OneWord("Сафаралиева Койкеб Кямил Кызы", Feminine, True)
    
    Call Assert.Equal_String("Сафаралиевой К. К.", Result.GetCase(Genitive))
    Call Assert.Equal_String("Сафаралиевой К. К.", Result.GetCase(Dative))
    Call Assert.Equal_String("Сафаралиеву К. К.", Result.GetCase(Accusative))
    Call Assert.Equal_String("Сафаралиевой К. К.", Result.GetCase(Instrumental))
    Call Assert.Equal_String("Сафаралиевой К. К.", Result.GetCase(Prepositional))
End Sub
Private Sub MasculineFullNameIsCorrectlyDeclined(Assert As Assert)
    Assert.NameMethod = "MasculineFullNameIsCorrectlyDeclined"
    Dim Result As New CyrResult
    Dim CyrName As New CyrName
    
    Set Result = CyrName.Decline_AllCases_OneWord("Иванов Иван Иванович", Masculine, False)
    
    Call Assert.Equal_String("Иванова Ивана Ивановича", Result.GetCase(Genitive))
    Call Assert.Equal_String("Иванову Ивану Ивановичу", Result.GetCase(Dative))
    Call Assert.Equal_String("Иванова Ивана Ивановича", Result.GetCase(Accusative))
    Call Assert.Equal_String("Ивановым Иваном Ивановичем", Result.GetCase(Instrumental))
    Call Assert.Equal_String("Иванове Иване Ивановиче", Result.GetCase(Prepositional))
    
    Set Result = CyrName.Decline_AllCases_OneWord("Карим Куржов Салим Оглы", Masculine, False)
    
    Call Assert.Equal_String("Карима Куржова Салим Оглы", Result.GetCase(Genitive))
    Call Assert.Equal_String("Кариму Куржову Салим Оглы", Result.GetCase(Dative))
    Call Assert.Equal_String("Карима Куржова Салим Оглы", Result.GetCase(Accusative))
    Call Assert.Equal_String("Каримом Куржовом Салим Оглы", Result.GetCase(Instrumental))
    Call Assert.Equal_String("Кариме Куржове Салим Оглы", Result.GetCase(Prepositional))
    
    Set Result = CyrName.Decline_AllCases_OneWord("Карим Куржов Салим-Оглы", Masculine, False)
    
    Call Assert.Equal_String("Карима Куржова Салим-Оглы", Result.GetCase(Genitive))
    Call Assert.Equal_String("Кариму Куржову Салим-Оглы", Result.GetCase(Dative))
    Call Assert.Equal_String("Карима Куржова Салим-Оглы", Result.GetCase(Accusative))
    Call Assert.Equal_String("Каримом Куржовом Салим-Оглы", Result.GetCase(Instrumental))
    Call Assert.Equal_String("Кариме Куржове Салим-Оглы", Result.GetCase(Prepositional))
    
    Set Result = CyrName.Decline_AllCases_OneWord("Иванов Иван Иванович", Masculine, True)
    
    Call Assert.Equal_String("Иванова И. И.", Result.GetCase(Genitive))
    Call Assert.Equal_String("Иванову И. И.", Result.GetCase(Dative))
    Call Assert.Equal_String("Иванова И. И.", Result.GetCase(Accusative))
    Call Assert.Equal_String("Ивановым И. И.", Result.GetCase(Instrumental))
    Call Assert.Equal_String("Иванове И. И.", Result.GetCase(Prepositional))
    
    Set Result = CyrName.Decline_AllCases_OneWord("Карим Куржов Салим Оглы", Masculine, True)
    
    Call Assert.Equal_String("Карима К. С.", Result.GetCase(Genitive))
    Call Assert.Equal_String("Кариму К. С.", Result.GetCase(Dative))
    Call Assert.Equal_String("Карима К. С.", Result.GetCase(Accusative))
    Call Assert.Equal_String("Каримом К. С.", Result.GetCase(Instrumental))
    Call Assert.Equal_String("Кариме К. С.", Result.GetCase(Prepositional))
    
    Set Result = CyrName.Decline_AllCases_OneWord("Илон МакФерсон", Masculine, False)
    
    Call Assert.Equal_String("Илона МакФерсона", Result.GetCase(Genitive))
    Call Assert.Equal_String("Илону МакФерсону", Result.GetCase(Dative))
    Call Assert.Equal_String("Илона МакФерсона", Result.GetCase(Accusative))
    Call Assert.Equal_String("Илоном МакФерсоном", Result.GetCase(Instrumental))
    Call Assert.Equal_String("Илоне МакФерсоне", Result.GetCase(Prepositional))
    
    Set Result = CyrName.Decline_AllCases_OneWord("Ахмед Гафуров ибн Мухаммад", Masculine, False)
    
    Call Assert.Equal_String("Ахмеда Гафурова ибн Мухаммада", Result.GetCase(Genitive))
    Call Assert.Equal_String("Ахмеду Гафурову ибн Мухаммаду", Result.GetCase(Dative))
    Call Assert.Equal_String("Ахмеда Гафурова ибн Мухаммада", Result.GetCase(Accusative))
    Call Assert.Equal_String("Ахмедом Гафуровом ибн Мухаммадом", Result.GetCase(Instrumental))
    Call Assert.Equal_String("Ахмеде Гафурове ибн Мухаммаде", Result.GetCase(Prepositional))
End Sub
Private Sub TestNameDeclension(Assert As Assert)
    Assert.NameMethod = "TestNameDeclension"
    Dim CyrName As New CyrName
    
    Call Assert.Equal_String("ивана", CyrName.DeclineNameAccusative("иван", False, False))
    Call Assert.Equal_String("ивану", CyrName.DeclineNameDative("иван", False, False))
    Call Assert.Equal_String("ивана", CyrName.DeclineNameGenitive("иван", False, False))
    Call Assert.Equal_String("иваном", CyrName.DeclineNameInstrumental("иван", False, False))
    Call Assert.Equal_String("иване", CyrName.DeclineNamePrepositional("иван", False, False))
    Call Assert.Equal_String("наталью", CyrName.DeclineNameAccusative("наталья", True, False))
    Call Assert.Equal_String("наталье", CyrName.DeclineNameDative("наталья", True, False))
    Call Assert.Equal_String("натальи", CyrName.DeclineNameGenitive("наталья", True, False))
    Call Assert.Equal_String("натальей", CyrName.DeclineNameInstrumental("наталья", True, False))
    Call Assert.Equal_String("наталье", CyrName.DeclineNamePrepositional("наталья", True, False))
End Sub
Private Sub TestSurnameDeclension(Assert As Assert)
    Assert.NameMethod = "TestSurnameDeclension"
    Dim CyrName As New CyrName
    
    Call Assert.Equal_String("иванова", CyrName.DeclineSurnameAccusative("иванов", False))
    Call Assert.Equal_String("иванову", CyrName.DeclineSurnameDative("иванов", False))
    Call Assert.Equal_String("иванова", CyrName.DeclineSurnameGenitive("иванов", False))
    Call Assert.Equal_String("ивановым", CyrName.DeclineSurnameInstrumental("иванов", False))
    Call Assert.Equal_String("иванове", CyrName.DeclineSurnamePrepositional("иванов", False))
    Call Assert.Equal_String("петрову", CyrName.DeclineSurnameAccusative("петрова", True))
    Call Assert.Equal_String("петровой", CyrName.DeclineSurnameDative("петрова", True))
    Call Assert.Equal_String("петровой", CyrName.DeclineSurnameGenitive("петрова", True))
    Call Assert.Equal_String("петровой", CyrName.DeclineSurnameInstrumental("петрова", True))
    Call Assert.Equal_String("петровой", CyrName.DeclineSurnamePrepositional("петрова", True))
End Sub
Private Sub TestPatronymicDeclension(Assert As Assert)
    Assert.NameMethod = "TestPatronymicDeclension"
    Dim CyrName As New CyrName

    Call Assert.Equal_String("ивановича", CyrName.DeclinePatronymicAccusative("иванович", False, False))
    Call Assert.Equal_String("ивановичу", CyrName.DeclinePatronymicDative("иванович", False, False))
    Call Assert.Equal_String("ивановича", CyrName.DeclinePatronymicGenitive("иванович", False, False))
    Call Assert.Equal_String("ивановичем", CyrName.DeclinePatronymicInstrumental("иванович", False, False))
    Call Assert.Equal_String("ивановиче", CyrName.DeclinePatronymicPrepositional("иванович", False, False))
    Call Assert.Equal_String("ивановну", CyrName.DeclinePatronymicAccusative("ивановна", True, False))
    Call Assert.Equal_String("ивановне", CyrName.DeclinePatronymicDative("ивановна", True, False))
    Call Assert.Equal_String("ивановны", CyrName.DeclinePatronymicGenitive("ивановна", True, False))
    Call Assert.Equal_String("ивановной", CyrName.DeclinePatronymicInstrumental("ивановна", True, False))
    Call Assert.Equal_String("ивановне", CyrName.DeclinePatronymicPrepositional("ивановна", True, False))
End Sub
'------------------------------------------------------
