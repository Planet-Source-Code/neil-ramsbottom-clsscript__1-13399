VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Script"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

'=========================================================================

'Author:    Neil Ramsbottom
'Date:      06/12/2000
'Purpose:   To allow fast access to any
'           location in a .script file
'Notes:
'           Use this code as you will but
'           dont forget to give me credit
'           where it is due.

'Website:   http://www.nramsbottom.co.uk

'=========================================================================

Option Explicit

Private m_strFilename As String
Private m_bLoaded As Boolean
Private m_arrScriptBuffer() As String

Public Property Get Filename() As String

    Filename = m_strFilename
    
End Property

Public Property Let Filename(ByVal strFilename As String)

    m_strFilename = strFilename

End Property

Public Function LoadScript(Optional strFilename As String) As Boolean

'On Error GoTo err   'If there is an error, exit function

If strFilename <> "" Then m_strFilename = strFilename 'Overrides module level variable

If m_strFilename = "" Then Exit Function    'Stop because there is not file to work on

Dim intFileNumber As Integer
Dim strBuffer As String
Dim strTmp As String
Dim lngNewElement As Long

ReDim m_arrScriptBuffer(1)



intFileNumber = FreeFile
Open m_strFilename For Input As #intFileNumber

While Not EOF(intFileNumber)

    Line Input #intFileNumber, strTmp
    
    'strBuffer = strBuffer & strTmp & vbCrLf
    
    ReDim Preserve m_arrScriptBuffer(lngNewElement)
    m_arrScriptBuffer(lngNewElement) = strTmp
    
    lngNewElement = UBound(m_arrScriptBuffer) + 1
    
Wend

Close #intFileNumber

m_bLoaded = True
LoadScript = True

Exit Function
err:
    
    LoadScript = False  'Not really needed, BWTF

End Function

Public Function GetLine(lngLineNumber As Long) As String
    
On Error Resume Next

    GetLine = m_arrScriptBuffer(lngLineNumber)

End Function

Public Sub SetLine(lngLineNumber As Long, strText As String)

On Error Resume Next

Dim lngLineCount As Long

lngLineCount = LineCount

    If lngLineNumber > LineCount Then
    
        ReDim Preserve m_arrScriptBuffer(lngLineNumber)
        
    End If
    
    m_arrScriptBuffer(lngLineNumber) = strText



End Sub


Public Function SaveScript() As Boolean

On Error Resume Next

Dim intFileNumber As Integer
Dim i As Integer

intFileNumber = FreeFile
Open m_strFilename For Output As #intFileNumber

For i = 0 To UBound(m_arrScriptBuffer)
    
    Print #intFileNumber, m_arrScriptBuffer(i)
    
Next i

Close #intFileNumber

If err.Number = 0 Then

    SaveScript = True
    
End If

End Function


Public Property Get LineCount() As Long

On Error Resume Next

    LineCount = UBound(m_arrScriptBuffer)
    
End Property

Public Property Get Loaded() As Boolean

    Loaded = m_bLoaded

End Property

Public Function GetWord(lngLineNumber As Long, intWordNumber As Integer) As String

On Error Resume Next

Dim arrTmp() As String

    arrTmp = Split(m_arrScriptBuffer(lngLineNumber), " ")
    GetWord = arrTmp(intWordNumber)

End Function

Public Function GetWordCount(lngLineNumber As Long) As Long

On Error Resume Next

Dim arrTmp() As String

    arrTmp = Split(m_arrScriptBuffer(lngLineNumber), " ")
    GetWordCount = UBound(arrTmp)
    
End Function

Private Sub Class_Initialize()

    ReDim m_arrScriptBuffer(1)

End Sub
