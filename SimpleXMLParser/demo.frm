VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmDemo 
   Caption         =   "SimpleXMLParser Demo"
   ClientHeight    =   5865
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9255
   LinkTopic       =   "Form1"
   ScaleHeight     =   391
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   617
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog CD 
      Left            =   1770
      Top             =   2100
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
      DefaultExt      =   "xml"
      Filter          =   "XML Files (*.xml)|*.xml|All Files (*.*)|*.*"
   End
   Begin VB.TextBox txtInfo 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4935
      Left            =   3870
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      Top             =   300
      Width           =   5325
   End
   Begin VB.CommandButton cmdOpen 
      Caption         =   "Open XML File"
      Height          =   405
      Left            =   75
      TabIndex        =   1
      Top             =   5385
      Width           =   2520
   End
   Begin MSComctlLib.ImageList imlFolders 
      Left            =   975
      Top             =   2040
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "demo.frx":0000
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.TreeView xmlTree 
      Height          =   5190
      Left            =   60
      TabIndex        =   0
      Top             =   75
      Width           =   3690
      _ExtentX        =   6509
      _ExtentY        =   9155
      _Version        =   393217
      Indentation     =   397
      LabelEdit       =   1
      Style           =   7
      ImageList       =   "imlFolders"
      Appearance      =   1
   End
   Begin VB.Label lblInfo 
      Caption         =   "Selected node info:"
      Height          =   225
      Left            =   4380
      TabIndex        =   3
      Top             =   60
      Width           =   2085
   End
End
Attribute VB_Name = "frmDemo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Parser      As clsXMLParser
Dim xmlNode()   As clsXMLNode
'

Private Sub cmdOpen_Click()

    On Error Resume Next
    Call CD.ShowOpen
    
    
    If Err.Number = cdlCancel Then Exit Sub 'Cancel pressed
    If Dir$(CD.Filename) = "" Then Exit Sub 'File not exists
    
    On Error GoTo 0
    
    Dim Node    As clsXMLNode
    Dim i       As Long
    Dim pn      As Node
    
    ReDim xmlNode(1)
    
    'initialize parser
    Set Parser = New clsXMLParser
    
    'parse xml file
    Call Parser.Parse(CD.Filename)
    
    Call xmlTree.Nodes.Clear
    
    'get xml node (if exists)
    Set Node = Parser.xmlNode
    If Not Node Is Nothing Then
        Call xmlTree.Nodes.Add(, , "xml node", "<?xml ... ?>", 1)
    End If
    
    'get main node
    Set Node = Parser.ParentNode
    Set xmlNode(1) = Node
    
    Set pn = xmlTree.Nodes.Add(, , "k1", Node.Name, 1)
    
    For i = 1 To Node.ChildrenCount
        Call makeXMLTree(Node.enumChild(i), "k1")
    Next
    
    pn.Expanded = True 'expand parent node
    
End Sub
'

'mirros structure of xml file into TreeView control
Private Sub makeXMLTree(ByVal Node As clsXMLNode, ParentNode As String)
    
    Dim i   As Long
    Dim k   As Long
    
    k = UBound(xmlNode) + 1
    ReDim Preserve xmlNode(k)
    
    Set xmlNode(k) = Node
    
    Call xmlTree.Nodes.Add(ParentNode, tvwChild, "k" & CStr(k), Node.Name, 1)
    
    For i = 1 To Node.ChildrenCount
        Call makeXMLTree(Node.enumChild(i), "k" & CStr(k))
    Next
    
End Sub
'

'resize/move controls to fit form's area
Private Sub Form_Resize()
    On Error Resume Next
    
    cmdOpen.Top = Me.ScaleHeight - cmdOpen.Height - 5
    
    txtInfo.Left = Me.ScaleWidth - txtInfo.Width - 5
    lblInfo.Left = txtInfo.Left
    
    xmlTree.Height = Me.ScaleHeight - cmdOpen.Height - 20
    txtInfo.Height = xmlTree.Height - 15
    
    xmlTree.Width = Me.ScaleWidth - txtInfo.Width - 15
    
End Sub

Private Sub xmlTree_NodeClick(ByVal Node As MSComctlLib.Node)
    
    Dim i   As Long
    
    If Node.Key = "xml node" Then
        txtInfo.Text = "This is xml language definition node." & vbNewLine & _
                    "It is not required (for this parser) but it defines language version and some other options like file encoding..." & vbNewLine & _
                    vbNewLine & "Attributes:" & vbNewLine
        For i = 1 To Parser.xmlNode.AttribCount
            txtInfo.Text = txtInfo.Text & _
            Parser.xmlNode.AttributeName(i) & ": " & _
            Parser.xmlNode.getAttribute(Parser.xmlNode.AttributeName(i)) & vbNewLine
        Next
    Else
        Dim n   As clsXMLNode
        
        Set n = xmlNode(CLng(Mid(Node.Key, 2)))
        
        txtInfo.Text = "Name: " & n.Name & vbNewLine & vbNewLine & _
                "Text: " & n.Text & vbNewLine & vbNewLine & _
                "Attributes:" & vbNewLine
        
        For i = 1 To n.AttribCount
            txtInfo.Text = txtInfo.Text & n.AttributeName(i) & ": " & _
                            n.getAttribute(n.AttributeName(i)) & vbNewLine
        Next
        
    End If
    
End Sub
