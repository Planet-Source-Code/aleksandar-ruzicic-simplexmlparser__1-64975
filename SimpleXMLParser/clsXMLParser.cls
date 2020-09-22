VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsXMLParser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Enum xmlParserError
    [xmlExpected] = 42
    [xmlMismatchedTag]
End Enum

Public Enum xmlIdentType
    [xmlSpace] = 0
    [xmlTab]
End Enum

Private m_parent        As clsXMLNode

Private m_look          As String * 1
Private m_pos           As Long

Private m_line          As Long
Private m_col           As Long

Private m_hasxml        As Boolean
Private m_xmlnode       As clsXMLNode

Private m_filenum       As Integer
'

Public Sub Parse(Filename As String)
    
    Call OpenFile(Filename)
    Set m_parent = parseParentNode
    Call CloseFile
    
End Sub
'

Public Sub Save(Filename As String, Optional PrettyPrint As Boolean = False, _
                                Optional IdentType As xmlIdentType = xmlTab)
    Call OpenFile(Filename)
    Call writeXML(PrettyPrint, IdentType)
    Call CloseFile
End Sub
'

Public Property Get ParentNode() As clsXMLNode
    Set ParentNode = m_parent
End Property
'

Public Property Let ParentNode(Node As clsXMLNode)
    Set m_parent = Node
End Property
'

Public Property Get XMLNode() As clsXMLNode
    Set XMLNode = m_xmlnode
End Property
'

Public Property Let XMLNode(Node As clsXMLNode)
    Set m_xmlnode = Node
End Property
'


'printing functions (used by Save)

Private Sub writeXML(pp As Boolean, it As xmlIdentType)
    
    Dim i   As Long
    
    If m_hasxml Then 'print <?xml ... ?>
        Call WriteString("<?xml")
        For i = 1 To m_xmlnode.AttribCount
            Call WriteString(" " & m_xmlnode.AttributeName(i) & "=" & _
              Chr$(34) & m_xmlnode.getAttribute(m_xmlnode.AttributeName(i)) _
              & Chr$(34))
        Next
        Call WriteString("?>" & IIf(pp, vbNewLine, vbNullString))
    End If
    
    Call printNode(m_parent, pp, 0, it)
    
End Sub
'

Private Sub printNode(Node As clsXMLNode, pp As Boolean, ident As Long, it As xmlIdentType)
    Dim i   As Long
    Dim b   As String
    Dim e   As String
    
    b = IIf(pp, String$(ident, IIf(it = xmlSpace, " ", vbTab)), vbNullString)
    e = IIf(pp, vbNewLine, vbNullString)
    
    Call WriteString(b & "<")
    Call WriteString(Node.Name)
    
    If Node.AttribCount > 0 Then
        For i = 1 To Node.AttribCount
            Call WriteString(" " & Node.AttributeName(i) & "=")
            Call WriteString(Chr$(34) & _
                             Node.getAttribute(Node.AttributeName(i)) & _
                             Chr$(34))
        Next
    End If
    
    
    If Node.ChildrenCount > 0 Then
        
        Call WriteString(">" & e)
        
        For i = 1 To Node.ChildrenCount
            Call printNode(Node.enumChild(i), pp, _
                IIf(it = xmlSpace, ident + 4, ident + 1), it)
        Next
        
        Call WriteString(b & "</" & Node.Name & ">" & e)
        
    Else
        If Node.Text = "" Then
            Call WriteString("/>" & e)
        Else
            Call WriteString(">" & Node.Text & "</" & Node.Name & ">" & e)
        End If
    End If
    
End Sub
'


'parse functions (used by Parse)

Private Sub Abort(errType As xmlParserError, Optional info As String = "")
    
    Dim Description As String
    
    Select Case errType
        Case xmlExpected
            Description = "Expected: " & info & " at line " & m_line & " [" & m_col & "]"
        
        Case xmlMismatchedTag
            Description = "Mismatched tag: '" & info & "' at line " & m_line & " [" & m_col - 2 & "]"
            
        Case Else 'never happens :)
            Description = "Unknown"
    End Select
    
    Call Err.Clear
    Call Err.Raise(errType, "XMLParser", Description)

End Sub
'

Private Sub getChar()
    m_look = ReadString
End Sub
'

Private Sub Match(What As String)
    If m_look = What Then
        Call getChar
    Else
        Call Abort(xmlExpected, What)
    End If
End Sub
'

Private Sub skipWhite()
    While InStr(1, " " & vbTab & vbCr & vbLf, m_look) > 0
        If m_look = vbCr Then
            m_line = m_line + 1
            m_col = 0
        End If
        Call getChar
    Wend
End Sub
'

Private Sub skipComment()

    If m_look = "<" Then
        If ReadString(3) = "!--" Then

            Do
                Call getChar

                If m_look = "-" Then
                    If ReadString(2) = "->" Then
                        Call getChar
                        Call skipWhite
                        Call skipComment
                        Exit Do
                    Else
                        Call SeekFile(m_pos - 1)
                    End If
                Else
                    If m_look = vbCr Then
                        m_line = m_line + 1
                        m_col = 0
                    End If
                End If
                
            Loop Until EOF(m_filenum)
            
        Else
            Call SeekFile(m_pos - 2)
        End If
    End If
    
    Call skipWhite
    
End Sub
'

Private Function isAlpha() As Boolean
    isAlpha = InStr(1, "ABCDEFGHIJKLMNOPQRSTUVWXYZ", UCase$(m_look)) > 0
End Function
'

Private Function isDigit() As Boolean
    isDigit = IsNumeric(m_look)
End Function
'

Private Function isValidName() As Boolean
    isValidName = isAlpha Or isDigit Or (InStr(1, "`!@#$^_-:", m_look) > 0)
End Function
'

Private Function getName() As String
    Dim n   As String
    
    n = vbNullString
    While isValidName() And Not EOF(m_filenum)
        n = n & m_look
        Call getChar
    Wend
    
    Call skipWhite
    
    getName = n
    
End Function
'

Private Function getValue(Terminator As String) As String
    Dim v   As String
    
    v = vbNullString
    While m_look <> Terminator And Not EOF(m_filenum)
        v = v & m_look
        Call getChar
    Wend
    
    Call skipWhite
    
    getValue = v
    
End Function
'

Private Function parseParentNode() As clsXMLNode
    
    Dim Name As String
    
    Call getChar
    Call skipWhite
    Call skipComment
    
    Call Match("<")
    
    m_hasxml = False
    If m_look = "?" Then '<?xml ... ?>
        Call Match("?")
        Name = getName
        
        If LCase$(Name) <> "xml" Then
            Call Abort(xmlExpected, "xml")
        End If
        
        m_hasxml = True
        
        Set m_xmlnode = New clsXMLNode
        m_xmlnode.Name = "xml"
        
        Call readAttributes(m_xmlnode)
        Call getChar
        Call skipWhite
        Call skipComment
        Call Match("<")
        Call SeekFile(m_pos - 2)
        Call getChar
    End If
    
    Name = getName

    Set parseParentNode = parseNode(Name)
    
End Function
'

Private Sub readAttributes(ByRef Node As clsXMLNode)
    
    Dim Name    As String
    Dim Value   As String

    While m_look <> "/" And m_look <> ">" And Not EOF(m_filenum)
        Call skipWhite
       
        Name = getName
        
        Call Match("=")
        Call Match("""")
        Value = getValue("""")
        Call Node.setAttribute(Name, Value)
        Call Match("""")

        Call skipWhite

        If m_hasxml And m_look = "?" Then
            Call Match("?")
        End If
    Wend
    
End Sub
'

Private Function parseNode(NodeName As String) As clsXMLNode
    
    Dim ret     As clsXMLNode
    Dim Name    As String
    Dim Child   As clsXMLNode
    Dim closed  As Boolean
    
    Set ret = New clsXMLNode
    ret.Name = NodeName
    closed = False
    
    Call skipWhite
    
    Call readAttributes(ret)
    
    If m_look = ">" Then
        Call Match(">")
        Call skipWhite
        Call skipComment
        
        If m_look = "<" Then
            
            Do
                Call Match("<")
                Call SeekFile(m_pos - 2)
                Call getChar
                Call skipWhite
                
                If m_look = "/" Then
                    Call Match("/")
                    Name = getName
                        
                    If LCase$(Name) = LCase$(NodeName) Then
                        Call Match(">")
                        closed = True
                        Exit Do
                    Else
                        Call Abort(xmlMismatchedTag, NodeName)
                    End If
                End If
                
                Name = getName
                Set Child = parseNode(Name)
                
                Call ret.addChild(Child)

                Call skipWhite
                Call skipComment
                
            Loop Until EOF(m_filenum)
            
        Else
            ret.Text = getValue("<")
        End If
        
        If Not closed Then
            Call Match("<")
            Call Match("/")
            
            Name = getName
            
            If LCase$(Name) = LCase$(NodeName) Then
                Call Match(">")
            Else
                Call Abort(xmlMismatchedTag, NodeName)
            End If
        End If
    Else
        Call Match("/")
        Call Match(">")
    End If
    
    Call skipWhite
    
    Set parseNode = ret
    
End Function
'

'file I/O functions (used by print and parse functions)

Private Sub OpenFile(Filename As String)
    m_filenum = FreeFile()
    m_pos = 0
    m_line = 1
    m_col = 0
    Open Filename For Binary As #m_filenum
End Sub
'

Private Sub CloseFile()
    Close #m_filenum
End Sub
'

Private Sub SeekFile(pos As Long)
    m_pos = pos
    Seek #m_filenum, m_pos
End Sub
'

Private Function ReadString(Optional Length As Long = 1) As String
    Dim sBuff   As String
    sBuff = Space(Length)
    Get #m_filenum, , sBuff
    ReadString = RTrim$(sBuff)
    m_pos = m_pos + Length
    m_col = m_col + Length
End Function
'

Private Sub WriteString(str As String)
    Put #m_filenum, , str
End Sub
