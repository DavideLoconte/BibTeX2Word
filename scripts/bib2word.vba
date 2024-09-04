'------------------------------------------------------------------------------
' Title: BibTeX to Word Bibliography Importer
' Description: A VBA script to import BibTeX citations into Word's bibliography.
' Author: Davide Loconte
' Date: 12/07/2024
'
' Copyright (c) 2024 Davide Loconte
'
' This program is free software: you can redistribute it and/or modify
' it under the terms of the GNU General Public License as published by
' the Free Software Foundation, either version 3 of the License, or
' (at your option) any later version.
'
' This program is distributed in the hope that it will be useful,
' but WITHOUT ANY WARRANTY; without even the implied warranty of
' MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE. See the
' GNU General Public License for more details.
'
' You should have received a copy of the GNU General Public License
' along with this program. If not, see <https://www.gnu.org/licenses/>.
'------------------------------------------------------------------------------

Sub TransformIntoCitation()
    Dim selectedText As String
    If Selection.Type = wdSelectionNormal Then
        ParseCitation (Selection.Text)
    End If
End Sub

Sub PasteIntoCitation()
    ' Using MSForms.DataObject require “Microsoft Forms 2.0 Object Library”
    ' If you cannot find the tool in reference list, import FM20.DLL file from the system32
    Dim objData As New MSForms.DataObject
    Dim strText
    objData.GetFromClipboard
    strText = objData.GetText()
    ParseCitation (strText)
End Sub

' ------------------------------------------------------------------------------
' Private
' ------------------------------------------------------------------------------

Sub ParseCitation(bibTeX As String)
    Dim xml As String

    Dim citationTag As String
    Dim citationClass As String
    Dim author As String
    Dim title As String
    Dim year As String
    Dim city As String
    Dim publisher As String

    ' Journal specifics
    Dim pages As String
    Dim journalName As String
    Dim volume As String
    Dim issue As String
    
    ' Book specific
    Dim bookTitle As String
    
    citationClass = GetCitationClass(bibTeX)
    
    If citationClass = "" Then
        MsgBox "Invalid bibTeX reference: " & bibTeX
    Else
    
        citationTag = GetBibKey(bibTeX)
        author = GetAuthorXML(bibTeX)
        title = GetField(bibTeX, "title")
        year = GetField(bibTeX, "year")
        city = GetField(bibTeX, "address")
        publisher = GetField(bibTeX, "publisher")
        
        xml = "<b:Source xmlns:b=""http://schemas.openxmlformats.org/officeDocument/2006/bibliography"">" & vbCrLf
        xml = xml & "  <b:Tag>" & citationTag & "</b:Tag>" & vbCrLf
        xml = xml & "  <b:SourceType>" & citationClass & "</b:SourceType>" & vbCrLf
        xml = xml & "  <b:Author>" & author & "</b:Author>" & vbCrLf
        
        If title <> "" Then xml = xml & "  <b:Title>" & title & "</b:Title>" & vbCrLf
        If year <> "" Then xml = xml & "  <b:Year>" & year & "</b:Year>" & vbCrLf
        If city <> "" Then xml = xml & "  <b:City>" & city & "</b:City>" & vbCrLf
        If publisher <> "" Then xml = xml & "  <b:Publisher>" & publisher & "</b:Publisher>" & vbCrLf
        
        If citationClass = "JournalArticle" Then
            pages = GetField(bibTeX, "pages")
            journalName = GetField(bibTeX, "journal")
            volume = GetField(bibTeX, "volume")
            issue = GetField(bibTeX, "number")
            
            If pages <> "" Then xml = xml & "  <b:Pages>" & pages & "</b:Pages>" & vbCrLf
            If journalName <> "" Then xml = xml & "  <b:JournalName>" & journalName & "</b:JournalName>" & vbCrLf
            If volume <> "" Then xml = xml & "  <b:Volume>" & volume & "</b:Volume>" & vbCrLf
            If issue <> "" Then xml = xml & "  <b:Issue>" & issue & "</b:Issue>" & vbCrLf
        End If
        
        If citationClass = "BookSection" Then
            bookTitle = GetField(bibTeX, "booktitle")
            If bookTitle <> "" Then xml = xml & "  <b:BookTitle>" & bookTitle & "</b:BookTitle>" & vbCrLf
        End If
        
        If citationClass = "ConferenceProceedings" Then
            bookTitle = GetField(bibTeX, "booktitle")
            If bookTitle <> "" Then xml = xml & "  <b:ConferenceName>" & bookTitle & "</b:ConferenceName>" & vbCrLf
        End If
        
        xml = xml & "</b:Source>"
        
        InsertBibCitation citationTag, xml
    End If
End Sub


Function GetCitationClass(bibTeX As String) As String
    Dim startPos As Integer
    Dim endPos As Integer
    
    Dim citationClass As String
    Dim ret As String
    
    startPos = InStr(bibTeX, "@") + 1
    endPos = InStr(bibTeX, "{")
    citationClass = ""
    
    If startPos > 0 And endPos > startPos Then
        citationClass = Mid(bibTeX, startPos, endPos - startPos)
        citationClass = LCase(citationClass)
    End If
    
    Select Case citationClass
        Case "article"
            ret = "JournalArticle"
        Case "book"
            ret = "Book"
        Case "booklet"
            ret = "Book"
        Case "inbook"
            ret = "BookSection"
        Case "incollection"
            ret = "BookSection"
        Case "inproceedings"
            ret = "ConferenceProceedings"
        Case "conference"
            ret = "ConferenceProceedings"
        Case "proceedings"
            ret = "ConferenceProceedings"
        Case "manual"
            ret = "Book"
        Case "mastersthesis"
            ret = "Misc"
        Case "phdthesis"
            ret = "Misc"
        Case "techreport"
            ret = "Report"
        Case "unpublished"
            ret = "Misc"
        Case "misc"
            ret = "Misc"
    End Select
    
    GetCitationClass = ret
End Function

Function GetBibKey(bibTeX As String) As String
    Dim startPos As Long
    Dim endPos As Long
    Dim bibKey As String
    
    startPos = InStr(bibTeX, "{") + 1
    endPos = InStr(bibTeX, ",")
    
    bibKey = Mid(bibTeX, startPos, endPos - startPos)
    GetBibKey = bibKey
End Function

Function GetField(bibTeX As String, fieldName As String) As String
    GetField = EscapeXMLSpecialChars(GetFieldInner(bibTeX, fieldName, 1))
End Function

Function GetFieldInner(bibTeX As String, fieldName As String, startPos As Integer) As String
    Dim firstEqual As Integer
    Dim ret As String
    Dim i As Integer
    Dim ch As String
    
    ret = ""
    startPos = InStr(startPos, bibTeX, fieldName)
    
    If startPos > 0 Then
        firstEqual = InStr(startPos, bibTeX, "=")
        If firstEqual > 0 Then
            For i = (firstEqual + 1) To Len(bibTeX)
                ch = Mid(bibTeX, i, 1)
                
                If ch = "{" Then
                    ret = GetFieldInBrackets(bibTeX, i)
                    Exit For
                ElseIf ch = """" Then
                    ret = GetQuotedField(bibTeX, i)
                    Exit For
                ElseIf Not CharIsWhitespace(ch) Then
                    ret = GetFieldUnbounded(bibTeX, i)
                    Exit For
                End If
            Next i
        End If
    End If
    
    If ret = "" And startPos > 0 Then
        ret = GetFieldInner(bibTeX, fieldName, startPos + Len(fieldName))
    End If
    
    GetFieldInner = ret
End Function

Function GetQuotedField(bibTeX As String, firstQuote As Integer)
    Dim i As Integer
    Dim ch As String
    Dim endPos As Integer
    Dim escaping As Boolean
    Dim ret As String
    Dim found As Boolean
    
    ret = ""
    found = False
    endPos = firstQuote
    
    For i = (firstQuote + 1) To Len(bibTeX)
        ch = Mid(bibTeX, i, 1)
        If ch = "\" And Not escaping Then
            escaping = True
        ElseIf Not escaping And ch = """" Then
            endPos = i
            found = True
            Exit For
        Else
            escaping = False
        End If
    Next i
    
    If found Then
        ret = Mid(bibTeX, firstQuote + 1, endPos - firstQuote - 1)
    End If
    
    
    GetQuotedField = ret
End Function

Function GetFieldInBrackets(bibTeX As String, firstBracket As Integer)
    Dim i As Integer
    Dim ch As String
    Dim endPos As Integer
    Dim levels As Integer
    Dim ret As String
    Dim found As Boolean
    
    ret = ""
    found = False
    endPos = firstBracket
    
    For i = firstBracket To Len(bibTeX)
        ch = Mid(bibTeX, i, 1)
        If ch = "{" Then
            levels = levels + 1
        ElseIf ch = "}" Then
            levels = levels - 1
        End If
        If levels = 0 Then
            endPos = i
            found = True
            Exit For
        End If
    Next i
    
    ' MsgBox "Bracket search end Pos = " & endPos & " and start POs = " & firstBracket
    If found Then
        ret = Mid(bibTeX, firstBracket + 1, endPos - firstBracket - 1)
    End If
    
    GetFieldInBrackets = ret
End Function

Function GetFieldUnbounded(bibTeX As String, firstChar As Integer)
    Dim i As Integer
    Dim ch As String
    Dim endPos As Integer
    Dim ret As String
    Dim found As Boolean
    
    ret = ""
    found = False
    endPos = firstChar
    
    For i = (firstChar + 1) To Len(bibTeX)
        ch = Mid(bibTeX, i, 1)
        If CharIsWhitespace(ch) Then
            endPos = i
            found = True
            Exit For
        End If
    Next i
    
    If found Then
        ret = Mid(bibTeX, firstChar, endPos - firstChar)
    End If
    
    GetFieldUnbounded = CleanString(ret)
End Function

Function GetAuthorXML(bibTeX As String) As String
    Dim authorField As String
    authorField = GetField(bibTeX, "author")
    
    If InStr(authorField, "{") = 1 Then
        GetAuthorXML = GetCorporateAuthor(authorField)
    Else
        GetAuthorXML = GetAuthorNameList(authorField)
    End If
    
End Function

Function GetAuthorNameList(authorField As String) As String
    Dim authors() As String
    Dim i As Integer
    Dim xml As String
    
    xml = "<b:Author><b:NameList>"
    authors = Split(authorField, " and ")

    For i = LBound(authors) To UBound(authors)
        Dim nameParts() As String
        nameParts = Split(authors(i), ",")
        
        If UBound(nameParts) = 1 Then
            xml = xml & "<b:Person><b:Last>" & EscapeXMLSpecialChars(CleanString(nameParts(0))) & "</b:Last><b:First>" & EscapeXMLSpecialChars(CleanString(nameParts(1))) & "</b:First></b:Person>"
        Else
            xml = xml & "<b:Person><b:Last>" & EscapeXMLSpecialChars(CleanString(nameParts(0))) & "</b:Last></b:Person>"
        End If
    Next i

    xml = xml & "</b:NameList></b:Author>"
    GetAuthorNameList = xml
End Function

Function GetCorporateAuthor(authorField As String) As String
    Dim corporateAuthor As String
    Dim startPos As Integer
    Dim endPos As Integer
    
    startPos = InStr(authorField, "{")
    endPos = InStr(authorField, "}")

    If startPos > 0 And endPos > startPos Then
        corporateAuthor = Mid(authorField, startPos + 1, endPos - startPos - 1)
    Else
        corporateAuthor = ""
    End If
    
    GetCorporateAuthor = "<b:Author><b:Corporate>" & EscapeXMLSpecialChars(corporateAuthor) & "</b:Corporate></b:Author>"
End Function

Function CleanString(str As String) As String
    Dim i As Integer
    Dim ch As String
    Dim resultStr As String
    
    resultStr = ""
    
    For i = 1 To Len(str)
        ch = Mid(str, i, 1)
        If (ch >= "A" And ch <= "Z") Or _
           (ch >= "a" And ch <= "z") Or _
           (ch >= "0" And ch <= "9") Or _
           (ch >= ChrW(192) And ch <= ChrW(382)) Then
            resultStr = resultStr & ch
        End If
    Next i
    
    CleanString = resultStr
End Function

Sub InsertBibCitation(citationTag As String, xml As String)
    If Not CitationTagExists(citationTag) Then
        ActiveDocument.Bibliography.Sources.Add xml
    End If
    Selection.Fields.Add Selection.Range, wdFieldCitation, citationTag
End Sub

Function CitationTagExists(citationTag As String) As Boolean
    Exists = False
    For i = 1 To ActiveDocument.Bibliography.Sources.Count
        If ActiveDocument.Bibliography.Sources(i).tag = citationTag Then
            Exists = True
            Exit For
        End If
    Next i
    CitationTagExists = Exists
End Function

Function CharIsWhitespace(ch As String) As Boolean
    If Not (ch = " " Or ch = vbTab Or ch = vbCr Or ch = vbLf) Then
        CharIsWhitespace = False
    Else
        CharIsWhitespace = True
    End If
End Function

Function IsAllWhitespace(inputStr As String) As Boolean
    Dim i As Integer
    Dim ch As String
    Dim isWhitespace As Boolean
    
    isWhitespace = True
    
    For i = 1 To Len(inputStr)
        ch = Mid(inputStr, i, 1)
        If Not (ch = " " Or ch = vbTab Or ch = vbCr Or ch = vbLf) Then
            isWhitespace = False
            Exit For
        End If
    Next i
    
    IsAllWhitespace = isWhitespace
End Function

Function EscapeXMLSpecialChars(inputString As String) As String
    inputString = Replace(inputString, "&", "&amp;")
    inputString = Replace(inputString, "<", "&lt;")
    inputString = Replace(inputString, ">", "&gt;")
    inputString = Replace(inputString, """", "&quot;")
    inputString = Replace(inputString, "'", "&apos;")
    inputString = Replace(inputString, "\", "")
    EscapeXMLSpecialChars = inputString
End Function

