VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Course"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private pClassName As String
Private pDate As String
Private pTime As String
Private pProfessor As String
Private pRoom As String

Public Property Let className(name As String)
    pClassName = name
End Property
Public Property Get className() As String
    className = pClassName
End Property
Public Property Let datum(newDate As String)
    pDate = newDate
End Property
Public Property Get datum() As String
    datum = pDate
End Property
Public Property Let time(newTime As String)
    pTime = newTime
End Property
Public Property Get time() As String
    time = pTime
End Property
Public Property Let professor(prof As String)
    pProfessor = prof
End Property
Public Property Get professor() As String
    professor = pProfessor
End Property
Public Property Let room(newRoom As String)
    pRoom = newRoom
End Property
Public Property Get room() As String
    room = pRoom
End Property


