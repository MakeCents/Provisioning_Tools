VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "PLISN"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public plisn As String
Public PCCN As String
Public partNumber As String
Public nsn As String
Public cage As String
Public SMR As String
Public Index As Integer

Public a As New Collection
Public B As New Collection
Public c As New Collection
Public d As New Collection
Public E As New Collection
Public f As New Collection
Public G As New Collection
Public H As New Collection
Public J As New Collection
Public K As New Collection
Public m As New Collection


Property Get Acard(X) As String
    Acard = a(X)
End Property
Property Get Bcard(X) As String
    Bcard = B(X)
End Property


Property Get cards() As String
    cards = ""
    On Error Resume Next
    Dim var As Variant
    Dim Item As Variant
    Dim X As Variant
    
    For Each var In Array("01", "02", "03", "04", "05", "06")
        cards = cards + vbCr + a(var)
    Next
    For Each var In Array("01", "02", "03", "04", "05", "06")
        cards = cards + vbCr + B(var)
    Next
    For Each var In Array("01", "02", "03", "04", "05", "06")
        cards = cards + vbCr + c(var)
    Next
    For Each var In Array("01", "02", "03", "04", "05", "06")
        cards = cards + vbCr + d(var)
    Next
    For Each var In Array("01", "02", "03", "04", "05", "06")
        cards = cards + vbCr + E(var)
    Next
    For Each var In Array("01", "02", "03", "04", "05", "06")
        cards = cards + vbCr + f(var)
    Next
    For Each var In Array("01", "02", "03", "04", "05", "06")
        cards = cards + vbCr + G(var)
    Next
    For Each var In Array("01", "02", "03", "04", "05", "06")
        cards = cards + vbCr + H(var)
    Next
    For Each var In Array("01", "02", "03", "04", "05", "06")
        cards = cards + vbCr + J(var)
    Next
    For Each var In Array("01", "02", "03", "04", "05", "06")
        cards = cards + vbCr + K(var)
    Next
    For Each var In Array("01", "02", "03", "04", "05", "06")
        cards = cards + vbCr + m(var)
    Next
    
    
End Property


