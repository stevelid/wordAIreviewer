Attribute VB_Name = "VAHelpers"
Option Explicit

'Utility helper mirrored from the VA Addin so FillBMs is available locally.
Public Sub FillBMs(ByRef targetBookmark As Bookmark, ByVal textValue As String)
    Dim bookmarkRange As Word.Range
    Dim bookmarkName As String

    If targetBookmark Is Nothing Then Exit Sub

    Set bookmarkRange = targetBookmark.Range
    bookmarkName = targetBookmark.Name

    bookmarkRange.Text = textValue
    ActiveDocument.Bookmarks.Add bookmarkName, bookmarkRange
End Sub

'Returns a ready-to-use Scripting.Dictionary without requiring references.
Public Function NewDictionary() As Object
    Set NewDictionary = CreateObject("Scripting.Dictionary")
End Function

'Safely checks whether an object behaves like a dictionary and contains a key.
Public Function HasDictionaryKey(ByVal dict As Object, ByVal keyName As String) As Boolean
    On Error GoTo CleanFail
    HasDictionaryKey = False
    If dict Is Nothing Then Exit Function
    If TypeName(dict) = "Collection" Then Exit Function
    If TypeName(dict) = "Dictionary" Or TypeName(dict) = "Scripting.Dictionary" Then
        If dict.Exists(keyName) Then HasDictionaryKey = True
        Exit Function
    End If
    If dict.Exists(keyName) Then HasDictionaryKey = True
    Exit Function
CleanFail:
    HasDictionaryKey = False
End Function

'Extracts context text from a suggestion object with defensive null handling.
Public Function GetSuggestionContextText(ByVal suggestion As Object) As String
    On Error GoTo CleanFail
    GetSuggestionContextText = "<missing context>"
    If suggestion Is Nothing Then Exit Function
    If TypeName(suggestion) = "Collection" Then Exit Function
    If HasDictionaryKey(suggestion, "context") Then
        Dim ctxValue As Variant
        ctxValue = suggestion("context")
        If IsNull(ctxValue) Then
            GetSuggestionContextText = "<missing context>"
        Else
            GetSuggestionContextText = CStr(ctxValue)
        End If
    End If
    Exit Function
CleanFail:
    GetSuggestionContextText = "<context error>"
End Function

'Retrieves a suggestion field as text with an optional default.
Public Function GetSuggestionText(ByVal suggestion As Object, ByVal key As String, Optional ByVal defaultText As String = "") As String
    On Error GoTo CleanFail
    If HasDictionaryKey(suggestion, key) Then
        Dim val As Variant
        val = suggestion(key)
        If IsNull(val) Then
            GetSuggestionText = defaultText
        Else
            GetSuggestionText = CStr(val)
        End If
    Else
        GetSuggestionText = defaultText
    End If
    Exit Function
CleanFail:
    GetSuggestionText = defaultText
End Function

'Extracts a representative keyword (first >4 letters, otherwise first non-empty word).
Public Function GetFirstSignificantWord(ByVal text As String) As String
    On Error Resume Next
    Dim words() As String
    Dim word As Variant
    Dim cleanText As String

    cleanText = Replace(text, ",", " ")
    cleanText = Replace(cleanText, ".", " ")
    cleanText = Replace(cleanText, ";", " ")
    cleanText = Replace(cleanText, ":", " ")
    cleanText = Replace(cleanText, "'", " ")
    cleanText = Replace(cleanText, """", " ")

    words = Split(Trim(cleanText), " ")

    For Each word In words
        If Len(word) > 4 Then
            GetFirstSignificantWord = word
            Exit Function
        End If
    Next word

    For Each word In words
        If Len(word) > 0 Then
            GetFirstSignificantWord = word
            Exit Function
        End If
    Next word

    GetFirstSignificantWord = ""
End Function

'Normalizes document text for deterministic matching.
Public Function NormalizeForDocument(ByVal value As String) As String
    Dim result As String
    result = value

    If result = "" Then
        NormalizeForDocument = ""
        Exit Function
    End If

    result = Replace(result, vbCrLf, Chr(13))
    result = Replace(result, vbLf, Chr(13))
    result = Replace(result, Chr(160), " ")
    result = Replace(result, vbTab, " ")
    result = Replace(result, Chr(11), " ")
    result = Replace(result, Chr(12), " ")

    result = Replace(result, ChrW(8220), Chr(34))
    result = Replace(result, ChrW(8221), Chr(34))
    result = Replace(result, ChrW(8216), "'")
    result = Replace(result, ChrW(8217), "'")
    result = Replace(result, ChrW(8211), "-")
    result = Replace(result, ChrW(8212), "-")

    Do While InStr(result, "  ") > 0
        result = Replace(result, "  ", " ")
    Loop

    result = NormalizeLinesWhitespace(result)

    NormalizeForDocument = result
End Function

Private Function NormalizeLinesWhitespace(ByVal text As String) As String
    Dim lines() As String
    Dim i As Long
    Dim result As String

    If text = "" Then
        NormalizeLinesWhitespace = ""
        Exit Function
    End If

    lines = Split(text, Chr(13))
    result = ""

    For i = LBound(lines) To UBound(lines)
        If i > LBound(lines) Then result = result & Chr(13)
        result = result & Trim(lines(i))
    Next i

    NormalizeLinesWhitespace = result
End Function
