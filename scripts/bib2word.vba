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

' Private

Sub ParseCitation(bibTeX As String)
    Dim citationClass As String
    citationClass = GetCitationClass(bibTeX)
    
    Select Case citationClass
        Case "article"
            ParseArticle (bibTeX)
        Case "book"
            ParseBook (bibTeX)
        Case "booklet"
            ParseBook (bibTeX)
        Case "inbook"
            ParseBookSection (bibTeX)
        Case "incollection"
            ParseBookSection (bibTeX)
        Case "inproceedings"
            ParseConference (bibTeX)
        Case "conference"
            ParseConference (bibTeX)
        Case "proceedings"
            ParseConference (bibTeX)
        Case "manual"
            ParseBook (bibTeX)
        Case "mastersthesis"
            ParseMisc (bibTeX)
        Case "phdthesis"
            ParseMisc (bibTeX)
        Case "techreport"
            ParseReport (bibTeX)
        Case "unpublished"
            ParseMisc (bibTeX)
        Case "misc"
            ParseMisc (bibTeX)
    End Select

End Sub

Function GetCitationClass(bibTeX As String) As String
    Dim startPos As Integer
    Dim endPos As Integer
    Dim citationClass As String
    
    ' Find the start position of the class (after '@' symbol)
    startPos = InStr(bibTeX, "@") + 1
    ' Find the end position of the class (before '{' symbol)
    endPos = InStr(bibTeX, "{")
    
    ' Extract the class of the citation
    If startPos > 0 And endPos > startPos Then
        citationClass = Mid(bibTeX, startPos, endPos - startPos)
        ' Convert the class to lowercase
        citationClass = LCase(citationClass)
    Else
        citationClass = ""
    End If
    
    GetCitationClass = citationClass
End Function

Sub ParseMisc(bibTeX As String)
    Dim citationTag As String
    Dim author As String
    Dim title As String
    Dim year As String
    Dim city As String
    Dim publisher As String
    
    Dim xml As String
    
    Dim authorField As String
    authorField = GetField(bibTeX, "author")
    citationTag = CleanString(Left(Replace(title, " ", ""), 10) & year & Left(CleanString(authorField), 10))
    
    author = GetAuthorXML(bibTeX)
    title = GetField(bibTeX, "title")
    year = GetField(bibTeX, "year")
    city = GetField(bibTeX, "address")
    publisher = GetField(bibTeX, "publisher")
    
    xml = "<b:Source xmlns:b=""http://schemas.openxmlformats.org/officeDocument/2006/bibliography"">" & vbCrLf
    xml = xml & "  <b:Tag>" & citationTag & "</b:Tag>" & vbCrLf
    xml = xml & "  <b:SourceType>Misc</b:SourceType>" & vbCrLf
    xml = xml & "  <b:Author>" & author & "</b:Author>" & vbCrLf
    
    If title <> "" Then xml = xml & "  <b:Title>" & title & "</b:Title>" & vbCrLf
    If year <> "" Then xml = xml & "  <b:Year>" & year & "</b:Year>" & vbCrLf
    If city <> "" Then xml = xml & "  <b:City>" & city & "</b:City>" & vbCrLf
    If publisher <> "" Then xml = xml & "  <b:Publisher>" & publisher & "</b:Publisher>" & vbCrLf
    xml = xml & "</b:Source>"
    
    ActiveDocument.Bibliography.Sources.Add xml
    Selection.Fields.Add Selection.Range, _
        wdFieldCitation, citation
End Sub

Sub ParseReport(bibTeX As String)
    ' Variables to hold the extracted fields
    Dim author As String
    Dim title As String
    Dim year As String
    Dim city As String
    Dim publisher As String
    Dim xml As String
    Dim citationTag As String

    ' Extract the fields from the BibTeX entry
    author = GetAuthorXML(bibTeX)
    title = GetField(bibTeX, "title")
    year = GetField(bibTeX, "year")
    city = GetField(bibTeX, "address")
    publisher = GetField(bibTeX, "institution")

    ' Create a unique citation tag
    Dim authorField As String
    authorField = GetField(bibTeX, "author")
    citationTag = CleanString(Left(Replace(title, " ", ""), 10) & year & Left(Replace(authorField, " ", ""), 10))

    ' Start of the XML
    xml = "<b:Source xmlns:b=""http://schemas.openxmlformats.org/officeDocument/2006/bibliography"">" & vbCrLf
    xml = xml & "  <b:Tag>" & citationTag & "</b:Tag>" & vbCrLf
    xml = xml & "  <b:SourceType>Report</b:SourceType>" & vbCrLf
    
    ' Add author
    xml = xml & "  <b:Author>" & author & "</b:Author>" & vbCrLf
    
    ' Add title
    If title <> "" Then xml = xml & "  <b:Title>" & title & "</b:Title>" & vbCrLf
    
    ' Add year
    If year <> "" Then xml = xml & "  <b:Year>" & year & "</b:Year>" & vbCrLf
    
    ' Add city
    If city <> "" Then xml = xml & "  <b:City>" & city & "</b:City>" & vbCrLf
    
    ' Add publisher
    If publisher <> "" Then xml = xml & "  <b:Publisher>" & publisher & "</b:Publisher>" & vbCrLf
    
    ' End of the XML
    xml = xml & "</b:Source>"

    ' Add the XML to the bibliography
    ActiveDocument.Bibliography.Sources.Add xml

    ' Insert citation at the current selection
    Selection.Fields.Add Selection.Range, wdFieldCitation, citationTag
End Sub

Sub ParseBook(bibTeX As String)
    ' Variables to hold the extracted fields
    Dim author As String
    Dim title As String
    Dim year As String
    Dim city As String
    Dim publisher As String
    Dim xml As String
    Dim citationTag As String

    ' Extract the fields from the BibTeX entry
    author = GetAuthorXML(bibTeX)
    title = GetField(bibTeX, "title")
    year = GetField(bibTeX, "year")
    city = GetField(bibTeX, "address")
    publisher = GetField(bibTeX, "publisher")

    ' Create a unique citation tag
    Dim authorField As String
    authorField = GetField(bibTeX, "author")
    citationTag = CleanString(Left(Replace(title, " ", ""), 10) & year & Left(Replace(authorField, " ", ""), 10))

    ' Start of the XML
    xml = "<b:Source xmlns:b=""http://schemas.openxmlformats.org/officeDocument/2006/bibliography"">" & vbCrLf
    xml = xml & "  <b:Tag>" & citationTag & "</b:Tag>" & vbCrLf
    xml = xml & "  <b:SourceType>Book</b:SourceType>" & vbCrLf
    
    ' Add author
    xml = xml & "  <b:Author>" & author & "</b:Author>" & vbCrLf
    
    ' Add title
    If title <> "" Then xml = xml & "  <b:Title>" & title & "</b:Title>" & vbCrLf
    
    ' Add year
    If year <> "" Then xml = xml & "  <b:Year>" & year & "</b:Year>" & vbCrLf
    
    ' Add city
    If city <> "" Then xml = xml & "  <b:City>" & city & "</b:City>" & vbCrLf
    
    ' Add publisher
    If publisher <> "" Then xml = xml & "  <b:Publisher>" & publisher & "</b:Publisher>" & vbCrLf
    
    ' End of the XML
    xml = xml & "</b:Source>"

    ' Add the XML to the bibliography
    ActiveDocument.Bibliography.Sources.Add xml

    ' Insert citation at the current selection
    Selection.Fields.Add Selection.Range, wdFieldCitation, citationTag
End Sub

Sub ParseArticle(bibTeX As String)
    ' Variables to hold the extracted fields
    Dim author As String
    Dim title As String
    Dim year As String
    Dim pages As String
    Dim journalName As String
    Dim volume As String
    Dim issue As String
    Dim xml As String
    Dim citationTag As String

    ' Extract the fields from the BibTeX entry
    author = GetAuthorXML(bibTeX)
    title = GetField(bibTeX, "title")
    year = GetField(bibTeX, "year")
    pages = GetField(bibTeX, "pages")
    journalName = GetField(bibTeX, "journal")
    volume = GetField(bibTeX, "volume")
    issue = GetField(bibTeX, "number")

    ' Create a unique citation tag
    Dim authorField As String
    authorField = GetField(bibTeX, "author")
    citationTag = CleanString(Left(Replace(title, " ", ""), 10) & year & Left(Replace(authorField, " ", ""), 10))

    ' Start of the XML
    xml = "<b:Source xmlns:b=""http://schemas.openxmlformats.org/officeDocument/2006/bibliography"">" & vbCrLf
    xml = xml & "  <b:Tag>" & citationTag & "</b:Tag>" & vbCrLf
    xml = xml & "  <b:SourceType>JournalArticle</b:SourceType>" & vbCrLf
    
    ' Add author
    xml = xml & "  <b:Author>" & author & "</b:Author>" & vbCrLf
    
    ' Add title
    If title <> "" Then xml = xml & "  <b:Title>" & title & "</b:Title>" & vbCrLf
    
    ' Add year
    If year <> "" Then xml = xml & "  <b:Year>" & year & "</b:Year>" & vbCrLf
    
    ' Add pages
    If pages <> "" Then xml = xml & "  <b:Pages>" & pages & "</b:Pages>" & vbCrLf
    
    ' Add journal name
    If journalName <> "" Then xml = xml & "  <b:JournalName>" & journalName & "</b:JournalName>" & vbCrLf
    
    ' Add volume
    If volume <> "" Then xml = xml & "  <b:Volume>" & volume & "</b:Volume>" & vbCrLf
    
    ' Add issue
    If issue <> "" Then xml = xml & "  <b:Issue>" & issue & "</b:Issue>" & vbCrLf
    
    ' End of the XML
    xml = xml & "</b:Source>"

    ' Add the XML to the bibliography
    ' TODO add to bibliography only if reference is not present in the list
    ActiveDocument.Bibliography.Sources.Add xml

    ' Insert citation at the current selection
    Selection.Fields.Add Selection.Range, wdFieldCitation, citationTag
End Sub


Sub ParseBookSection(bibTeX As String)
    ' Variables to hold the extracted fields
    Dim author As String
    Dim title As String
    Dim year As String
    Dim pages As String
    Dim bookTitle As String
    Dim city As String
    Dim publisher As String
    Dim xml As String
    Dim citationTag As String

    ' Extract the fields from the BibTeX entry
    author = GetAuthorXML(bibTeX)
    title = GetField(bibTeX, "title")
    year = GetField(bibTeX, "year")
    pages = GetField(bibTeX, "pages")
    bookTitle = GetField(bibTeX, "booktitle")
    city = GetField(bibTeX, "address")
    publisher = GetField(bibTeX, "publisher")

    ' Create a unique citation tag
    Dim authorField As String
    authorField = GetField(bibTeX, "author")
    citationTag = CleanString(Left(Replace(title, " ", ""), 10) & year & Left(Replace(authorField, " ", ""), 10))

    ' Start of the XML
    xml = "<b:Source xmlns:b=""http://schemas.openxmlformats.org/officeDocument/2006/bibliography"">" & vbCrLf
    xml = xml & "  <b:Tag>" & citationTag & "</b:Tag>" & vbCrLf
    xml = xml & "  <b:SourceType>BookSection</b:SourceType>" & vbCrLf
    
    ' Add author
    xml = xml & "  <b:Author>" & author & "</b:Author>" & vbCrLf
    
    ' Add title
    If title <> "" Then xml = xml & "  <b:Title>" & title & "</b:Title>" & vbCrLf
    
    ' Add year
    If year <> "" Then xml = xml & "  <b:Year>" & year & "</b:Year>" & vbCrLf
    
    ' Add pages
    If pages <> "" Then xml = xml & "  <b:Pages>" & pages & "</b:Pages>" & vbCrLf
    
    ' Add book title
    If bookTitle <> "" Then xml = xml & "  <b:BookTitle>" & bookTitle & "</b:BookTitle>" & vbCrLf
    
    ' Add city
    If city <> "" Then xml = xml & "  <b:City>" & city & "</b:City>" & vbCrLf
    
    ' Add publisher
    If publisher <> "" Then xml = xml & "  <b:Publisher>" & publisher & "</b:Publisher>" & vbCrLf
    
    ' End of the XML
    xml = xml & "</b:Source>"

    ' Add the XML to the bibliography
    ActiveDocument.Bibliography.Sources.Add xml

    ' Insert citation at the current selection
    Selection.Fields.Add Selection.Range, wdFieldCitation, citationTag
End Sub

Sub ParseConference(bibTeX As String)
    ' Variables to hold the extracted fields
    Dim author As String
    Dim title As String
    Dim year As String
    Dim city As String
    Dim conferenceName As String
    Dim xml As String
    Dim citationTag As String

    ' Extract the fields from the BibTeX entry
    author = GetAuthorXML(bibTeX)
    title = GetField(bibTeX, "title")
    year = GetField(bibTeX, "year")
    city = GetField(bibTeX, "address")
    conferenceName = GetField(bibTeX, "booktitle")

    ' Create a unique citation tag
    Dim authorField As String
    authorField = GetField(bibTeX, "author")
    citationTag = CleanString(Left(Replace(title, " ", ""), 10) & year & Left(Replace(authorField, " ", ""), 10))

    ' Start of the XML
    xml = "<b:Source xmlns:b=""http://schemas.openxmlformats.org/officeDocument/2006/bibliography"">" & vbCrLf
    xml = xml & "  <b:Tag>" & citationTag & "</b:Tag>" & vbCrLf
    xml = xml & "  <b:SourceType>ConferenceProceedings</b:SourceType>" & vbCrLf
    xml = xml & "  <b:Guid>{" & CreateObject("Scriptlet.TypeLib").GUID & "}</b:Guid>" & vbCrLf
    
    ' Add author
    xml = xml & "  <b:Author>" & author & "</b:Author>" & vbCrLf
    
    ' Add title
    If title <> "" Then xml = xml & "  <b:Title>" & title & "</b:Title>" & vbCrLf
    
    ' Add year
    If year <> "" Then xml = xml & "  <b:Year>" & year & "</b:Year>" & vbCrLf
    
    ' Add city
    If city <> "" Then xml = xml & "  <b:City>" & city & "</b:City>" & vbCrLf
    
    ' Add conference name
    If conferenceName <> "" Then xml = xml & "  <b:ConferenceName>" & conferenceName & "</b:ConferenceName>" & vbCrLf
    
    ' End of the XML
    xml = xml & "</b:Source>"

    ' Add the XML to the bibliography
    ActiveDocument.Bibliography.Sources.Add xml

    ' Insert citation at the current selection
    Selection.Fields.Add Selection.Range, wdFieldCitation, citationTag
End Sub

Function GetField(bibTeX As String, fieldName As String) As String
    Set regex = CreateObject("VBScript.RegExp")
    
    ' Create a pattern to find the field value
    fieldPattern = fieldName & "\s*=\s*(""(.*)""|{(.*)})"
    regex.IgnoreCase = True
    regex.Global = False
    regex.Pattern = fieldPattern

    Set matches = regex.Execute(bibTeX)
    If matches.Count > 0 Then
        Set match = matches(0)
        If Len(match.SubMatches(1)) > Len(match.SubMatches(2)) Then
            fieldValue = match.SubMatches(1)
        Else
            fieldValue = match.SubMatches(2)
        End If
    Else
        fieldValue = ""
    End If
    
    GetField = fieldValue
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
            xml = xml & "<b:Person><b:Last>" & CleanString(nameParts(0)) & "</b:Last><b:First>" & CleanString(nameParts(1)) & "</b:First></b:Person>"
        Else
            xml = xml & "<b:Person><b:Last>" & CleanString(nameParts(0)) & "</b:Last></b:Person>"
        End If
    Next i

    xml = xml & "</b:NameList></b:Author>"
    GetAuthorNameList = xml
End Function

Function GetCorporateAuthor(authorField As String) As String
    Dim corporateAuthor As String
    Dim startPos As Integer
    Dim endPos As Integer

    ' Find the start and end positions of the corporate author
    startPos = InStr(authorField, "{")
    endPos = InStr(authorField, "}")

    If startPos > 0 And endPos > startPos Then
        ' Extract the corporate author
        corporateAuthor = Mid(authorField, startPos + 1, endPos - startPos - 1)
    Else
        ' No corporate author found, return empty string
        corporateAuthor = ""
    End If
    
    GetCorporateAuthor = "<b:Author><b:Corporate>" & corporateAuthor & "</b:Corporate></b:Author>"
End Function


Function CleanString(str As String) As String
    Dim regex As Object
    Set regex = CreateObject("VBScript.RegExp")
    
    regex.Pattern = "[^A-Za-z0-9À-ž\s]"
    regex.Global = True
    regex.IgnoreCase = True
    
    CleanString = regex.Replace(str, "")
End Function

Sub InsertBibCitation(citation As String)
    Selection.Fields.Add Selection.Range, _
        wdFieldCitation, citation
End Sub
