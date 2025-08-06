Attribute VB_Name = "xlwings_utils"
Sub EncodeFile(Sheet As Worksheet, File As String, Optional Row As Integer = 1, Optional Column As Integer = 1)
    Dim S As String
    Dim I As Long
    Dim NumberOfChunks As Long
    Dim ChunkSize As Integer
    
    Dim BinaryStream As Object
    Dim XmlObj As Object
    Dim Element As Object
  
    Set BinaryStream = CreateObject("ADODB.Stream")
    With BinaryStream
        .Type = 1 ' adTypeBinary
        .Open
        .LoadFromFile File
    End With

    Set XmlObj = CreateObject("MSXML2.DOMDocument")
    Set Element = XmlObj.createElement("b64")
    Element.DataType = "bin.base64"
    Element.nodeTypedValue = BinaryStream.Read() ' read all bytes
    BinaryStream.Close

    Encoded = Replace(Element.Text, Chr(10), "") ' remove linefeeds
   
    ChunkSize = 5000
    NumberOfChunks = Application.WorksheetFunction.Ceiling(Len(Encoded) / ChunkSize, 1)
    ReDim Chunks(0 To NumberOfChunks - 1)
    
    Sheet.Cells(Row, Column) = "<file=" + File + ">"

    For I = 0 To NumberOfChunks - 1
        Sheet.Cells(I + Row + 1, Column) = Mid$(Encoded, I * ChunkSize + 1, ChunkSize)
    Next I
    
    Sheet.Cells(I + Row + 1, Column) = "</file>"
    
End Sub

Private Function Base64ToArray(base64 As String) As Variant
    
    Dim xmlDoc As Object
    Dim xmlNode As Object
    
    Set xmlDoc = CreateObject("MSXML2.DOMDocument")
    Set xmlNode = xmlDoc.createElement("b64")
    
    xmlNode.DataType = "bin.base64"
    xmlNode.Text = base64
    
    Base64ToArray = xmlNode.nodeTypedValue
  
End Function

Function DecodeFile(Sheet As Worksheet, Optional Row As Integer = 1, Optional Column As Integer = 1)
    Dim vArr() As Byte
    Dim S As String
    Dim Count As Integer
    Dim FileNames As String
    
    ThisDir = ThisWorkbook.Path
    Count = 0
    While Row < 30000
        Line = Sheet.Cells(Row, Column)
        If InStr(Line, "<file=") = 1 And Right(Line, 1) = ">" Then
            If Count <> 0 Then
                FileNames = FileNames & ", "
            End If
            Count = Count + 1
            FileNameOnly = Mid(Line, 7, Len(Line) - 7)
            Filename = ThisDir & "/" & FileNameOnly
            FileNames = FileNames & FileNameOnly
            
            Row = Row + 1
            S = ""
            While Sheet.Cells(Row, Column) <> "</file>"
                S = S & Sheet.Cells(Row, Column)
                Row = Row + 1
            Wend
                
            vArr = Base64ToArray(S)
        
            Open Filename For Binary Access Write As #1
            WritePos = 1
            Put #1, WritePos, vArr
            Close #1
        End If
        
        Row = Row + 1
    
    Wend
    
    DecodeFile = Count
   

End Function


