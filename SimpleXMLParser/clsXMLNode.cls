VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsXMLNode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_name          As String
Private m_child()       As clsXMLNode
Private m_attrib()      As String
Private m_value()       As String
Private m_text          As String
'


Public Property Let Name(n As String)
    m_name = n
End Property
Public Property Get Name() As String
    Name = m_name
End Property

Public Property Let Text(t As String)
    m_text = t
End Property
Public Property Get Text() As String
    Text = m_text
End Property


Public Function getAttribute(attribName As String, _
                        Optional Default As String = vbNullString) As String
    Dim i   As Long
    
    i = attribLookup(attribName)
    
    If (i > 0) Then
        getAttribute = m_value(i)
    Else
        getAttribute = Default
    End If
    
End Function
'

Public Sub setAttribute(attribName As String, Value As String)
    On Error Resume Next
    
    Dim i   As Long
    
    i = attribLookup(attribName)
    
    If (i > 0) Then
        m_value(i) = Value
    Else
        i = UBound(m_attrib) + 1
        
        ReDim Preserve m_attrib(i)
        ReDim Preserve m_value(i)
        
        m_attrib(i) = attribName
        m_value(i) = Value
        
    End If
    
End Sub
'

Public Property Get AttribCount() As Long
    AttribCount = UBound(m_attrib)
End Property
'

Public Property Get ChildrenCount() As Long
    ChildrenCount = UBound(m_child)
End Property
'

Public Sub addChild(Child As clsXMLNode)
    On Error Resume Next
    
    Dim i   As Long
    
    i = UBound(m_child) + 1
    
    ReDim Preserve m_child(i)
    
    Set m_child(i) = Child
    'Debug.Print "Added Node '" & Child.Name & "' into '" & m_name & "'"
End Sub
'

Public Function enumChild(ByVal Index As Long) As clsXMLNode
    Set enumChild = m_child(Index)
End Function
'

Public Function AttributeName(Index As Long) As String
    AttributeName = m_attrib(Index)
End Function
'

Private Function attribLookup(attribName As String) As Long
    On Error GoTo notFound
    
    Dim i   As Long
    
    For i = 1 To UBound(m_attrib)
        If (LCase$(m_attrib(i)) = LCase$(attribName)) Then
            attribLookup = i
            Exit Function
        End If
    Next
    
notFound:
    Call Err.Clear
    attribLookup = 0
End Function
'

Private Sub Class_Initialize()
    ReDim m_attrib(0)
    ReDim m_value(0)
    ReDim m_child(0)
End Sub
