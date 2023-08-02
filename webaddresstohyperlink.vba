Sub ConvertWebAddressesToHyperlinks()
    Dim doc As Document
    Set doc = ActiveDocument
    
    Dim rng As Range
    Set rng = doc.Content
    
    Dim regEx As Object
    '''
    ''' To identify a web address, this regular expression searches for any string that begins with "http:" "https:" or "ftp:"
    ''' and ends in a whitespace character (such as a line break or space).
    ''' Since this variable is also used as the hyperlink URL, I'm hesitant to include "www." or "doi" as expressions - but we could.
    '''
    '''
    Set regEx = CreateObject("VBScript.RegExp")
    regEx.Global = True
    regEx.Pattern = "((?:https?|ftp)://\S+)" ' Modify the regular expression here.
    
    Dim matches As Object
    Set matches = regEx.Execute(rng.Text)
    
    Dim match As Object
    
    Dim counter As Long
    counter = 0
    
    For Each match In matches
        Dim linkText As String
        linkText = match.Value
        If Not HyperlinkExists(doc, linkText) Then
            Do
                rng.Find.ClearFormatting
                rng.Find.Text = linkText
                rng.Find.Execute
                If rng.Find.Found Then
                    rng.Hyperlinks.Add Anchor:=rng, address:=linkText
                    rng.Collapse wdCollapseEnd ' Move to the end of the hyperlink range
                    counter = counter + 1
                End If
            Loop While rng.Find.Found And counter < 400 ' Terminates after 400 hyperlink attempts to prevent risk of infinite loop
        End If
    Next match
    
    Set regEx = Nothing
    Set matches = Nothing
    Set rng = Nothing
    Set doc = Nothing
End Sub

Function HyperlinkExists(doc As Document, address As String) As Boolean
'''
''' Checks to see if the link already exists in the document.
''' The way this is called in ConvertWebAddressToHyperlinks(), if the hyperlink already exists in the document
''' (even if it exists elsewhere), it will move on to the next address withoput making a new link.
'''
'''
    Dim hlink As Hyperlink
    For Each hlink In doc.Hyperlinks
        If hlink.address = address Then
            HyperlinkExists = True
            Exit Function
        End If
    Next hlink
    HyperlinkExists = False
End Function
