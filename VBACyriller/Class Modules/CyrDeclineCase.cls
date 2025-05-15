VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CyrDeclineCase"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'***************Fields'***********************************
Private NameRu As String
Private NameEn As String
Private description As String
Private index As Long
Private Value As CasesEnum
'*********************************************************

'***************Properties********************************
Property Get Name_Ru() As String
    Name_Ru = NameRu
End Property
Public Property Let Name_Ru(p_NameRu As String)
    NameRu = p_NameRu
End Property
Property Get Name_En() As String
    Name_En = NameEn
End Property
Public Property Let Name_En(p_NameEn As String)
    NameEn = p_NameEn
End Property
Property Get Desc() As String
    Desc = description
End Property
Public Property Let Desc(p_Desc As String)
    description = p_Desc
End Property
Property Get ind() As Long
    ind = index
End Property
Public Property Let ind(p_Index As Long)
    index = p_Index
End Property
Property Get Val() As CasesEnum
    Val = Value
End Property
Public Property Let Val(p_Value As CasesEnum)
    Value = p_Value
End Property
Public Property Let CyrDeclineCase(NameRu As String, NameEn As String, description As String, index As Long, Value As CasesEnum)
    NameRu = NameRu
    NameEn = NameEn
    description = description
    index = index
    Value = Value
End Property
'*********************************************************
'***************Methods***********************************
Public Function GetEnumerable() As CyrDeclineCase()
    Dim Result() As CyrDeclineCase
    
    ReDim Result(5)
    Set Result(0) = Case1
    Set Result(1) = case2
    Set Result(2) = Case3
    Set Result(3) = Case4
    Set Result(4) = Case5
    Set Result(5) = Case6
    
    GetEnumerable = Result
End Function

Public Function Case1() As CyrDeclineCase
    Dim Result As New CyrDeclineCase
         
    Result.Name_Ru = "Именительный"
    Result.Name_En = "Nominative"
    Result.Desc = "Кто? Что? (есть)"
    Result.ind = 1
    Result.Val = Nominative
    
    Set Case1 = Result
End Function
Public Function case2() As CyrDeclineCase
    Dim Result As New CyrDeclineCase
     
    Result.Name_Ru = "Родительный"
    Result.Name_En = "Genitive"
    Result.Desc = "Кого? Чего? (нет)"
    Result.ind = 2
    Result.Val = Genitive
    
    Set case2 = Result
End Function
Public Function Case3() As CyrDeclineCase
    Dim Result As New CyrDeclineCase
     
    Result.Name_Ru = "Дательный"
    Result.Name_En = "Dative"
    Result.Desc = "Кому? Чему? (дам)"
    Result.ind = 3
    Result.Val = Dative
    
    Set Case3 = Result
End Function
Public Function Case4() As CyrDeclineCase
    Dim Result As New CyrDeclineCase
     
    Result.Name_Ru = "Винительный"
    Result.Name_En = "Accusative"
    Result.Desc = "Кого? Что? (вижу)"
    Result.ind = 4
    Result.Val = Accusative
    
    Set Case4 = Result
End Function
Public Function Case5() As CyrDeclineCase
    Dim Result As New CyrDeclineCase
     
    Result.Name_Ru = "Творительный"
    Result.Name_En = "Instrumental"
    Result.Desc = "Кем? Чем? (горжусь)"
    Result.ind = 5
    Result.Val = Instrumental
    
    Set Case5 = Result
End Function
Public Function Case6() As CyrDeclineCase
    Dim Result As New CyrDeclineCase
     
    Result.Name_Ru = "Предложный"
    Result.Name_En = "Prepositional"
    Result.Desc = "О ком? О чем? (думаю)"
    Result.ind = 6
    Result.Val = Prepositional
    
    Set Case6 = Result
End Function

Public Function Nominative_() As CyrDeclineCase
    Set Nominative_ = Case1
End Function
Public Function Genitive_() As CyrDeclineCase
    Set Genitive_ = case2
End Function
Public Function Dative_() As CyrDeclineCase
    Set Dative_ = Case3
End Function
Public Function Accusative_() As CyrDeclineCase
    Set Accusative_ = Case4
End Function
Public Function Instrumental_() As CyrDeclineCase
    Set Instrumental_ = Case5
End Function
Public Function Prepositional_() As CyrDeclineCase
    Set Prepositional_ = Case6
End Function

Public Function Именительный() As CyrDeclineCase
    Set Именительный = Case1
End Function
Public Function Родительный() As CyrDeclineCase
    Set Родительный = case2
End Function
Public Function Дательный() As CyrDeclineCase
    Set Дательный = Case3
End Function
Public Function Винительный() As CyrDeclineCase
    Set Винительный = Case4
End Function
Public Function Творительный() As CyrDeclineCase
    Set Творительный = Case5
End Function
Public Function Предложный() As CyrDeclineCase
    Set Предложный = Case6
End Function
'*********************************************************

