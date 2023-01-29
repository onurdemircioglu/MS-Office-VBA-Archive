Sub AddPageNumbers()
    'https://learn.microsoft.com/en-us/office/vba/api/word.pagenumbers.add
    With ActiveDocument.Sections(1)
        .Footers(wdHeaderFooterPrimary).PageNumbers.add _
         PageNumberAlignment:=wdAlignPageNumberCenter, _
         FirstPage:=True
    End With

End Sub



Sub RemovePageNumbers()
    'Open AI ChatGPT Jan 9 Version
    Dim sec As Section
    For Each sec In ActiveDocument.Sections
        sec.Footers(wdHeaderFooterPrimary).Range.Delete
    Next sec
End Sub
