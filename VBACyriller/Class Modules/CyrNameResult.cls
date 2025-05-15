VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CyrNameResult"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'***************Description*******************************
'*********************************************************

'***************Fields'***********************************
Private Name As String
Private surname As String
Private patronymic As String
'*********************************************************

'***************Properties********************************
Property Get Name_() As String
    Name_ = Name
End Property
Public Property Let Name_(p_name As String)
    Name = p_name
End Property
Property Get Surname_() As String
    Surname_ = surname
End Property
Public Property Let Surname_(p_surname As String)
    surname = p_surname
End Property
Property Get patronymic_() As String
    patronymic_ = patronymic
End Property
Public Property Let patronymic_(p_patronymic As String)
    patronymic = p_patronymic
End Property

'*********************************************************

'***************Methods***********************************
Public Sub Constructor(Name As String, surname As String, patronymic As String)
    Name = Name
    surname = surname
    patronymic = patronymic
End Sub
Public Function Constructor2(values() As String)
    Dim i As Long
    i = UBound(values)
    If UBound(values) = 1 Or UBound(values) = 2 Then
        surname = values(0)
        Name = values(1)
        patronymic = IIf(UBound(values) = 2, values(2), "")
    End If
End Function

Public Function FullName() As String
    FullName = Trim(Trim(surname) + " " + Trim(Name) + " " + Trim(patronymic))
End Function

Public Function ToString() As String
    ToString = FullName
End Function

Public Function Equals(cyr_name_result As CyrNameResult) As Boolean
    Equals = cyr_name_result.Name_ = Name And cyr_name_result.Surname_ = surname And cyr_name_result.patronymic_ = patronymic_
End Function

'*********************************************************
