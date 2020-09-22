VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsEncodeDecode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public Function EncodeString(ByVal S As String) As String
 EncoString = "INTRUDUCTORY METHODS OF NUMERICAL ANALYSIS"
 l = Len(S)
 For i = 1 To l
  value1 = Mid(S, i, 1)
  Value2 = Mid(EncoString, i, 1)
  num = Asc(Value2)
  ascii = Asc(value1)
  ascii = ascii + num
  Mid(S, i, 1) = Chr(ascii)
 Next
 EncodeString = S
End Function
Public Function DecodeString(ByVal S As String) As String
 EncoString = "INTRUDUCTORY METHODS OF NUMERICAL ANALYSIS"
 l = Len(S)
 For i = 1 To l
  value1 = Mid(S, i, 1)
  Value2 = Mid(EncoString, i, 1)
  num = Asc(Value2)
  ascii = Asc(value1)
  ascii = ascii - num
  Mid(S, i, 1) = Chr(ascii)
 Next
 DecodeString = S
End Function

