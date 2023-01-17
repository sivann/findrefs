Sub pasteunformated()
'
' pasteunformated Macro
' Macro recorded 6/29/2004 by w
'
    Selection.PasteAndFormat (wdFormatPlainText)
End Sub
Sub finduncited()
'
' finduncited Macro
' Macro created 19/7/2005 by sivann
'
Dim c As Long
Dim fieldCode As String
Dim msg As String
Dim flds(255) As Long
Dim fldno As Long


fldno = 1
'read fields into a table
For ref = 1 To ActiveDocument.Fields.Count
    If ActiveDocument.Fields(ref).Type = wdFieldRef Then
        reftxt = ActiveDocument.Fields(ref).Result.Text
        l = InStr(reftxt, "[")
        r = InStr(reftxt, "]")
        lr = r - l - 1
        If l > 0 Then
            flds(fldno) = Mid(reftxt, l + 1, lr)
            fldno = fldno + 1
        End If
    End If
Next
   
c = 1

'get numbered items
For Each itm In ActiveDocument.GetCrossReferenceItems(wdRefTypeNumberedItem)
    If InStr(itm, "[") > 0 Then
        match = 0
        l = InStr(itm, "[")
        r = InStr(itm, "]")
        lr = r - l - 1
        itmnum = Mid(itm, l + 1, lr)
        If IsNumeric(itmnum) Then
            itmnum = CLng(itmnum)
        Else
            GoTo nxt
        End If
        
        For i = 1 To fldno
            If flds(i) = itmnum Then
                    match = match + 1 'match
            End If
        Next
        'MsgBox (itm & " Matches = " & match)
        If match = 0 Then
            MsgBox (itm & " : No Matches Found")
            vResponse = MsgBox("Continue Searching? ", vbYesNo)
            If vResponse = vbNo Then Exit For
        End If
    End If
nxt:
Next
End Sub

Sub FindUnrefedTablFigures()
'
' findunrefedfigures Macro
' Macro created 19/7/2005 by sivann
' find figures and talbes that have not been cited
Dim figseqs(384) As String
Dim figseqnames(384) As String
Dim figseqsno As Integer

Dim tabseqs(384) As String
Dim tabseqnames(384) As String
Dim tabseqsno As Integer


Dim refs(384) As String
Dim refsno As Integer
Dim str As String
Dim seqstr As String

figseqsno = 0
tabseqsno = 0
refsno = 0

'prepei:    1) check poia den exoun bookmark
'           2) poia exoun alla den uparxei _Ref s'ayta
ActiveDocument.Bookmarks.ShowHidden = True

'1) find sequences without a bookmark (probably never cross-referenced)
For ref = 1 To ActiveDocument.Fields.Count
    ftype = ActiveDocument.Fields(ref).Type
    If ftype = wdFieldSequence Then
        rtxt = ActiveDocument.Fields(ref).Result.Text
        fld = ActiveDocument.Fields(ref)
        
        'get seq type
        If InStr(fld, "SEQ Fig") > 0 Then
            seqstr = "Figure"
            t = 1
        ElseIf InStr(fld, "SEQ Tab") > 0 Then
            seqstr = "Table"
            t = 2
        Else
            GoTo nxt
        End If
        
        'find seq bookmarks
        s = fld.Start
        e = fld.End
        bk = ActiveDocument.Range(s, e).Bookmarks.Count
        If bk > 0 Then
            For i = 1 To bk
                str = ActiveDocument.Range(s, e).Bookmarks(i).Name
                If InStr(str, "_Ref") And t = 1 Then
                    figseqs(figseqsno) = str
                    figseqnames(figseqsno) = rtxt
                    figseqsno = figseqsno + 1
                ElseIf InStr(str, "_Ref") And t = 2 Then
                    tabseqs(tabeqsno) = str
                    tabseqnames(tabseqsno) = rtxt
                    tabseqsno = tabseqsno + 1
                End If
            Next
        Else 'no bookmark exists. crossreference auto-adds a bookmark
            vResponse = MsgBox(seqstr & " " & rtxt & " has never been referenced. Continue?", vbYesNo)
            If vResponse = vbNo Then Exit For
        End If
    ElseIf ftype = wdFieldRef Then
        rtxt = ActiveDocument.Fields(ref).Result.Text
        fld = ActiveDocument.Fields(ref)
        str = fld
        str = Split(str)(2) 'get bookmark no
        refs(refsno) = str
        refsno = refsno + 1
    End If
nxt:
Next

'poia exoun bookmark alla den uparxei _Ref s'ayta
For sq = 0 To figseqsno
    match = 0
    For rf = 0 To refsno
        If refs(rf) = figseqs(sq) Then match = match + 1
    Next
    If match = 0 Then
        rtxt = figseqnames(sq)
        vResponse = MsgBox("Figure " & rtxt & " isn't referenced. Continue?", vbYesNo)
        If vResponse = vbNo Then Exit For
    End If
Next

For sq = 0 To tabseqsno
    match = 0
    For rf = 0 To refsno
        If refs(rf) = tabseqs(sq) Then match = match + 1
    Next
    If match = 0 Then
        rtxt = tabseqnames(sq)
        vResponse = MsgBox("Table " & rtxt & " isn't referenced. Continue?", vbYesNo)
        If vResponse = vbNo Then Exit For
    End If
Next

End Sub


Sub ShowSelectionBookmarks()
Dim str As String

    ActiveDocument.Bookmarks.ShowHidden = True
    s = Selection.Start
    e = Selection.End
    bk = ActiveDocument.Range(s, e).Bookmarks.Count
    For i = 1 To bk
        str = ActiveDocument.Range(s, e).Bookmarks(i).Name
        vResponse = MsgBox(str & " Continue?", vbYesNo)
        If vResponse = vbNo Then Exit For
     Next
End Sub

Sub FindUncitedRefs()
' FindUncited References Macro
' references are considered items of the last numbered list
' Macro created 20/7/2005 by sivann

Dim str As String
Dim bk As Integer
Dim refs(384) As String
Dim refnames(384) As String
Dim refsno As Integer
Dim xrefs(384) As String
Dim xrefsno As Integer
Dim bl As Boolean


refsno = 0
bl = ActiveDocument.Bookmarks.ShowHidden
ActiveDocument.Bookmarks.ShowHidden = True


totlists = ActiveDocument.Lists.Count
totrefs = ActiveDocument.Lists.Item(totlists).ListParagraphs.Count
'ActiveDocument.Lists.Item(totlists).Range.Select
's = Selection.Start
'e = Selection.End
'bk = ActiveDocument.Range(s, e).Bookmarks.Count


'make a list (refs) with bookmarks on the last numbered list
'references are probably the last numbered list
For rf = 1 To totrefs
    bk = ActiveDocument.Lists.Item(totlists).ListParagraphs.Item(rf).Range.Bookmarks.Count
    If bk = 0 Then
        str = ActiveDocument.Lists.Item(totlists).ListParagraphs.Item(rf).Range.Text
        vResponse = MsgBox("Reference " & rf & " " & str & vbCr & " has never been referenced. Continue?", vbYesNo)
        If vResponse = vbNo Then Exit Sub
    Else
        str = ActiveDocument.Lists.Item(totlists).ListParagraphs.Item(rf).Range.Bookmarks(1).Name
        refs(refsno) = str
        refnames(refsno) = ActiveDocument.Lists.Item(totlists).ListParagraphs.Item(rf).Range.Text
        refsno = refsno + 1
    End If
Next

'make a list (xrefs) with bookmark numbers in reference fields
For ref = 1 To ActiveDocument.Fields.Count
    ftype = ActiveDocument.Fields(ref).Type
    If ftype = wdFieldRef Then
        rtxt = ActiveDocument.Fields(ref).Result.Text
        fld = ActiveDocument.Fields(ref)
        str = fld
        str = Split(str)(2) 'get bookmark no
        xrefs(xrefsno) = str
        xrefsno = xrefsno + 1
    End If
Next

'compare lists
For rf = 0 To refsno
    match = 0
    For xrf = 0 To xrefsno
        If xrefs(xrf) = refs(rf) Then match = match + 1
    Next
    If match = 0 Then
        rtxt = refnames(rf)
        vResponse = MsgBox("Reference " & rf + 1 & " " & rtxt & vbCr & " isn't referenced. Continue?", vbYesNo)
        If vResponse = vbNo Then Exit Sub
    End If
Next

ActiveDocument.Bookmarks.ShowHidden = bl

End Sub

Sub SortRefsByOrderofAppearance()
' sort references in order of citation
' SortRefsByOrderofAppearance Macro
' Macro created 20/7/2005 by sivann
' first working version 29/8/2005

Dim str As String
Dim s1, s2 As String
Dim rtxt As String
Dim bl As Boolean
Dim xrefs(384) As String
Dim xrefspos(384) As Integer
Dim xrefsno As Integer
Dim citerefs(384) As String
Dim citerefspar(384) As Integer
Dim citerefsno As Integer

citerefsno = 0
xrefsno = 0

bl = ActiveDocument.Bookmarks.ShowHidden
ActiveDocument.Bookmarks.ShowHidden = True

totlists = ActiveDocument.Lists.Count
totrefs = ActiveDocument.Lists.Item(totlists).ListParagraphs.Count

'create a list with the bookmark numbers of references(citations)
'citations are the last numbered list.
xxx = ActiveDocument.Range.Bookmarks.Count
For i = 1 To totrefs
    bkcnt = ActiveDocument.Lists.Item(totlists).ListParagraphs.Item(i).Range.Bookmarks.Count
    If bkcnt = 0 Then GoTo nxt1:
    For j = 1 To bkcnt
        str = ActiveDocument.Lists.Item(totlists).ListParagraphs.Item(i).Range.Bookmarks(j).Name
        citerefs(citerefsno) = str
        citerefspar(citerefsno) = i
        citerefsno = citerefsno + 1
        tmp1 = ActiveDocument.Lists(totlists).ListParagraphs(i).Range.FormattedText.Text
        
    Next
nxt1:
Next



'make a list of cross-references (xrefs) which cite references
'xrefs will contain bookmark numbers of references by order of appearance in text
'parse in reverse so as for multiple citations to the same reference, only the
'first one (last in parse order) gets recorded

For ref = ActiveDocument.Fields.Count To 1 Step -1
    ftype = ActiveDocument.Fields(ref).Type
    If ftype = wdFieldRef Then
        rtxt = ActiveDocument.Fields(ref).Result.Text
        fld = ActiveDocument.Fields(ref)
        str = fld
        str = Split(str)(2) 'get bookmark no
        'search for str in citerefs
        match = 0
        For crf = 0 To citerefsno - 1
            If str = citerefs(crf) Then
                match = match + 1
                Exit For
            End If
        Next
        'it is a reference to a citation.
        If match > 0 Then
            xrefs(xrefsno) = str
            x = citerefspar(crf) ' paragraph of bkmrk
            xrefspos(x) = ref
            xrefsno = xrefsno + 1
        End If
    End If 'field type
Next


'prepend a number denoting the place to go
For i = 1 To ActiveDocument.Lists.Item(totlists).ListParagraphs.Count
    wrd = ActiveDocument.Lists(totlists).ListParagraphs(i).Range.FormattedText.Words.First.Text
    If xrefspos(i) = 0 Then xrefspos(i) = 999 'put uncited to the end
    'newstr = xrefspos(i) & "%"
    newstr = Format(xrefspos(i), "000") & "%"
    ActiveDocument.Lists(totlists).ListParagraphs(i).Range.FormattedText.Words.First.Text = newstr & wrd
Next

'ActiveDocument.Bookmarks.ShowHidden = False

ActiveDocument.Lists(totlists).Range.Select

Selection.Sort ExcludeHeader:=False, FieldNumber:="Paragraphs", _
        SortFieldType:=wdSortFieldAlphanumeric, SortOrder:=wdSortOrderAscending, _
        FieldNumber2:="", SortFieldType2:=wdSortFieldAlphanumeric, SortOrder2:= _
        wdSortOrderAscending, FieldNumber3:="", SortFieldType3:= _
        wdSortFieldAlphanumeric, SortOrder3:=wdSortOrderAscending, Separator:= _
        wdSortSeparateByDefaultTableSeparator, SortColumn:=False, CaseSensitive:= _
        False, LanguageID:=wdEnglishUS, SubFieldNumber:="Paragraphs", _
        SubFieldNumber2:="Paragraphs", SubFieldNumber3:="Paragraphs"

'remove prepended string
For i = 1 To ActiveDocument.Lists.Item(totlists).ListParagraphs.Count
    wrd = ActiveDocument.Lists(totlists).ListParagraphs(i).Range.FormattedText.Words.First.Text
    If wrd = "%" Then 'no number, reference was uncited
        ActiveDocument.Lists(totlists).ListParagraphs(i).Range.FormattedText.Words.First.Text = ""
        GoTo nxt2
    End If
    'delete number and %
    ActiveDocument.Lists(totlists).ListParagraphs(i).Range.FormattedText.Words.First.Text = ""
    ActiveDocument.Lists(totlists).ListParagraphs(i).Range.FormattedText.Words.First.Text = ""
    
nxt2:
Next


ActiveDocument.Bookmarks.ShowHidden = bl

End Sub

Sub removeLineBreaks()
'
' remove_line_breaks Macro
' Macro recorded 11/13/2002 by Norm Jones
'
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    
    With Selection.Find
        .Text = "^p^p"
        .Replacement.Text = "^t"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With

    Selection.Find.Execute Replace:=wdReplaceAll
    With Selection.Find
        .Text = "^p"
        .Replacement.Text = " "
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    
    Selection.Find.Execute Replace:=wdReplaceAll
    
    With Selection.Find
        .Text = "^t"
        .Replacement.Text = "^p^p"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    
    Selection.Find.Execute Replace:=wdReplaceAll
    
    
End Sub

