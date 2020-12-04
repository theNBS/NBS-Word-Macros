Attribute VB_Name = "modNBSMacrosWord"
' ========================================================================
'
' Macro code written to help work with MS Word
' either prior to importing into NBS Chorus
' or once exported out of Chorus
'
' Contributors:
' 1. Stephen Hamil
'
' Use at own risk
' Always save and keep copies of Word documents before running the macros
'
' Support Community - https://support.thenbs.com/support/home
'
' modNBSMacrosWord      3rd Dec 2020
' frmNBSSelectStyles    3rd Dec 2020
' frmNBSWarning         3rd Dec 2020
' frmKeyNoteText        3rd Dec 2020
'
' modNBSMacros
' ============
' This is the main module with the primary function cals
'
' 1. SetStyles() - quickly sets styles prior to stylesheet import
' 2. PrepareNBSForImport() - does some pre-processing for NBS word processing files prior to import
' 3. GenerateKeyNoteText() - loops through an exported DOCX file and generates keynotes for CAWS or Uniclass 2015
' ========================================================================

' Ability to debug in real time (if needed)
Dim g_strDebugText

Enum enumNBSParagraphType
    enNBSSectionHeading = 1
    enNBSSectionSubHeading = 2
    enNBSClause = 3
    enNBSClauseHeading = 4
    enNBSClauseRow = 5
    enNBSBlank = 6
End Enum


Enum enumClassificationType
    enClassCAWS = 1
    enClassMasterFormat = 2
    enClassUniclass = 3
    enClassUnknown = 4
End Enum

' Primary function - SetStyles()
' ==============================
'
' Designed to work with the NBS print template for UK and Aus
' NBS-stylesheet-UK-AUS-01-Oct-2020.docx
'
' Will display a dialog in which
' (a) a colour can be specified that
' will change every heading and hyperlink
' (b) a font can be specificed that will
' change all text in the template
' ===============================
Sub SetStyles()
Attribute SetStyles.VB_Description = "123"
Attribute SetStyles.VB_ProcData.VB_Invoke_Func = "Normal.modNBSMacros.SetStyles"
    
    ' Set up the variables
    Dim strFontNormal, intColourR, intColourG, intColourB
 
    ' Fire up a form so the user can specify their requirements
    Dim objNewForm
    Set objNewForm = New frmNBSSelectStyles
    objNewForm.Show
    With objNewForm
        If .ApplyChanges = True Then
            ' User has clicked OK - let's grab the values
            strFontNormal = .FontFamily
            intColourR = .HeadingColourR
            intColourG = .HeadingColourG
            intColourB = .HeadingColourB
        Else
            ' User has clicked cancel - we are out of here
            Exit Sub
        End If
    End With
    
       
    ''''''''''''''''''''''''''''''''''''
    ' Main body text
    ' Set 'Normal' to the choice of font
    ' Note: Maybe code could be improved through looping through some sort of array
    ' - but a big list seems to work OK
    ActiveDocument.Styles("Normal").Font.Name = strFontNormal
    ActiveDocument.Styles("chorus-cite-clause").Font.Name = strFontNormal
    ActiveDocument.Styles("chorus-clause-code").Font.Name = strFontNormal
    ActiveDocument.Styles("chorus-clause-row").Font.Name = strFontNormal
    ActiveDocument.Styles("chorus-clause-row-bullet").Font.Name = strFontNormal
    ActiveDocument.Styles("chorus-clause-row-label").Font.Name = strFontNormal
    ActiveDocument.Styles("chorus-clause-row-title").Font.Name = strFontNormal
    ActiveDocument.Styles("chorus-clause-row-value").Font.Name = strFontNormal
    ActiveDocument.Styles("chorus-clause-row-value-bullet-list-item").Font.Name = strFontNormal
    ActiveDocument.Styles("chorus-clause-row-value-numbered-list-item").Font.Name = strFontNormal
    ActiveDocument.Styles("chorus-clause-title").Font.Name = strFontNormal
    ActiveDocument.Styles("chorus-clause-title-deleted").Font.Name = strFontNormal
    ActiveDocument.Styles("chorus-section-end").Font.Name = strFontNormal
    ActiveDocument.Styles("chorus-section-header-code").Font.Name = strFontNormal
    ActiveDocument.Styles("chorus-shared-by").Font.Name = strFontNormal
   
    
    '''''''''''''''''''''''''
    ' HEADINGS AND HYPERLINKS
    ' Set font and color to the headings
    With ActiveDocument.Styles("chorus-clause-group-title").Font
        .Name = strFontNormal
        .Color = RGB(intColourR, intColourG, intColourB)
    End With
    
    With ActiveDocument.Styles("chorus-section-header").Font
        .Name = strFontNormal
        .Color = RGB(intColourR, intColourG, intColourB)
    End With
    
    With ActiveDocument.Styles("chorus-section-header-code").Font
        .Name = strFontNormal
        .Color = RGB(intColourR, intColourG, intColourB)
    End With
    
    
    With ActiveDocument.Styles("Subtitle").Font
        .Name = strFontNormal
        .Color = RGB(intColourR, intColourG, intColourB)
    End With
    
    With ActiveDocument.Styles("TOC Heading").Font
        .Name = strFontNormal
        .Color = RGB(intColourR, intColourG, intColourB)
    End With
    
    With ActiveDocument.Styles("chorus-clause-link").Font
        .Name = strFontNormal
        .Color = RGB(intColourR, intColourG, intColourB)
    End With
    
    With ActiveDocument.Styles("Hyperlink").Font
        .Name = strFontNormal
        .Color = RGB(intColourR, intColourG, intColourB)
    End With
    
    Set objNewForm = Nothing
End Sub


' Primary function - PrepareNBSForImport()
' ==============================
'
' Designed to work with the NBS word processing disc files
' that were in use in the twenty years+ leading through to 2020
'
' Will display a dialog to confirm the user is happy
' for their document to modified
' then it will:
' (a) Set Heading 1, Heading 2, Heading 3
' (b) remove '-' and '______.' content from rows
' (c) remove the section prefix from clauses
' (d) add 'GENERAL' sub heading if it cannot find a Heading 2
' (e) TO DO - add '||' separator field to seperate code from titles in headings and clauses
' ===============================
Sub PrepareNBSForImport()
Attribute PrepareNBSForImport.VB_ProcData.VB_Invoke_Func = "Normal.NewMacros.Test"

    Dim singleLine As Paragraph 'use when looping through every line of the word doc
    Dim strLineText As String ' use to grab the text value of each line
    Dim strSectionCode ' will try and grab the section code each time - i.e. F10
    strLineText = ""
    strSection = ""
       

    Dim objForm
    Set objForm = New frmNBSWarning
    objForm.Show
    With objForm
        If .ApplyChanges = False Then
            ' User has clicked cancel - we are out of here
            Exit Sub
        End If
    End With
    

    ' Firstly, turn on hidden text else the macro will not work properly
    ActiveWindow.ActivePane.View.ShowAll = True
    
    RemoveAllBreaks

    ' Loop through every paragraph in the word document
    ' looking out for headings, clauses and normal text...
    For Each singleLine In ActiveDocument.Paragraphs
        strLineText = singleLine.Range.Text ' get the text of the paragraph
        ' strLineText = StripUnwantedCharacters(strLineText) ' clean it up (page breaks etc...)
        
        Dim eNBSParagraphType
        eNBSParagraphType = enNBSClauseRow ' cast to the default type
        eNBSParagraphType = DetermineParagraphType(strSectionCode, strLineText, singleLine)
        
        Select Case eNBSParagraphType
            Case enNBSSectionHeading
                ''''''''''''''''''''''''''
                 ' We have a new section (Heading 1)
                 ' Grab the section code now
                 strSectionCode = Mid(strLineText, 1, 3)
                 ' Make this line Heading 1 style
                 singleLine.Style = wdStyleHeading1
                 
                 ' TO DO - this only needs to happen if the next paragraph is *not* Heading 2
                 ' Add a new line that is of style Heading 2...
                singleLine.Range.InsertAfter "GENERAL" & vbCrLf
                
            Case enNBSSectionSubHeading
                ' Not sure if there is any way of determining this programatically ???
                
            Case enNBSClause
                ''''''''''''''''''''''''''''''
                ' We have a clause (Heading 3)
                ' Set as Heading 3 style
                singleLine.Range.Style = wdStyleHeading3
                                
                ' get rid of the section code and the slash
                ' NOTE - Find.Execute means that the style is not overridden (what happens with normal property set)
                singleLine.Range.Find.Execute strSectionCode & "/", , , , , , , , , ""
                
                
            Case enNBSClauseHeading
                ' Not sure if there is any way of determining this programatically ???
                
            Case enNBSClauseRow
                ' we have a row - Set the style
                singleLine.Range.Style = "Normal"
                ''''''''''''''''''''''''''''''
                ' we have a row
                
                ' we get rid of the tab dash tab
                singleLine.Range.Find.Execute vbTab & "-" & vbTab, , , , , , , , , ""
                
                ' and the _______ - but ensure there is still a space there so the import recognises it
                singleLine.Range.Find.Execute " ______ .", , , , , , , , , ""
        
        End Select
            
    Next singleLine
    
    
    ' A bit bodgey - but we now do a second pass setting the Heading 2 styles for where we have inserted 'GENERAL'
    ' Happy to discuss better/more robust logic.
    For Each singleLine In ActiveDocument.Paragraphs
        strLineText = singleLine.Range.Text

        If Mid(strLineText, 1, 7) = "GENERAL" Then
            singleLine.Style = wdStyleHeading2
        End If
    Next singleLine

    
    ' Enable this method if you want a log file
    ' SaveDebugInfoToTextFile

    
    ' Old code
    ' Initially I tried to just change the NBS styles use
    ' by production team in a 'one-er' - problem here is that you need to loop
    ' through the lines anyway - so changed the way of thinking
    ' may want to re-evaluate some day though
    'With ActiveDocument.Styles("NBS heading")
    '    .AutomaticallyUpdate = False
    '    .BaseStyle = "Heading 1"
    '    .NextParagraphStyle = "NBS minor clause"
    'End With
    
    'With ActiveDocument.Styles("NBS minor clause")
    '    .AutomaticallyUpdate = False
    '    .BaseStyle = "Heading 3"
    '    .NextParagraphStyle = "NBS minor clause"
    'End With
End Sub



' Primary function - GenerateKeyNoteText()
' ========================================
'
' Designed to work with the NBS Chorus Word export DOCX
'
' Loops through the document creating trhe tree view structure
' based on the style logic Heading 1, Heading 2 etc...'
'
'===============================
Sub GenerateKeyNoteText()
    ' This will be the key note text
    Dim strKeyNoteText As String
    strKeyNoteText = ""
    
    ' Try our best using VBA(!) to get a wait cursor up
    ' if time presented itself, a nice progress bar would be lovely
    System.Cursor = wdCursorWait
    DoEvents

    ' Logic is different between CAWS, MasterFormat and Uniclass
    ' So we need to work out what classification we're working with from the first
    ' section we find
    
    Dim enClassType As enumClassificationType
    enClassType = ReturnClassificationType()
    
    Select Case enClassType
        Case enClassCAWS
                strKeyNoteText = ReturnCAWSKeyNoteText()
        
        Case enClassUniclass
                strKeyNoteText = ReturnUniclassKeyNoteText()
        
        Case enClassMasterFormat
            ' to do
    
    End Select
     
    ' Debug.Print "Keynote text: " & strKeyNoteText

    ' cursor back to normal
    System.Cursor = wdCursorNormal


    ' bang up the keynote text in a dialog
    ' the user can copy and paste into notepad and save - todo - some nice save dialog
    Dim objForm As New frmKeyNoteText
    objForm.SetText strKeyNoteText
    objForm.Show

    

End Sub



Private Function ReturnClassificationType() As enumClassificationType

    ' Loop through every paragraph until we can work out the classification type - then exit loop
    ' It would be nice if we were working with a proper object model
    ' but we're not - so we'll have to try and find patterns
    
    Dim sLine As Paragraph
    Dim strLineText As String
    Dim intPos As Integer
    Dim strSectionCode As String
    
    For Each sLine In ActiveDocument.Paragraphs
        strLineText = sLine.Range.Text ' get the text of the paragraph
        
        Select Case sLine.Style
            ' First time into a new group - we need to grab the group code...
            Case "chorus-section-header"
                intPos = InStr(1, strLineText, "", vbTextCompare)
                strSectionCode = Mid(strLineText, 1, intPos - 1)
                
                ' For CAWS the section code is three digits - C10 for example
                If Len(strSectionCode) = 3 Then
                    ReturnClassificationType = enClassCAWS
                    Exit For
                End If
                
                ' For Uniclass the section code contains an underbar _ - Ss_25_30_25 for example
                If InStr(1, strSectionCode, "_", vbTextCompare) <> 0 Then
                    ReturnClassificationType = enClassUniclass
                    Exit For
                End If
                
                ' All others, we'll assume Masterformat for now
                ReturnClassificationType = enClassMasterFormat
                Exit For ' no point in sending the code through every line
        End Select
        
    Next sLine

End Function


Private Function ReturnCAWSKeyNoteText() As String


    Dim sLine As Paragraph
    Dim strLineText
    
    Dim strGroupCode As String
    Dim strSectionCode As String
    Dim strSectionTitle As String
    
    Dim strClauseCode As String
    Dim strClauseTitle As String
    Dim intPos As Integer
    
    Dim strGroupLetter As String
    Dim strPreviousGroupLetter As String
    Dim strGroupTitle As String
    
    Dim nCount As Integer
    nCount = 0
    
    System.Cursor = wdCursorWait
    
    

    ' Loop through every paragraph in the word document
    ' looking out for the word styles we need
    For Each sLine In ActiveDocument.Paragraphs
        
                
        Select Case sLine.Style
            ' First time into a new group - we need to grab the group code...
            
            Case "chorus-section-header"
                strLineText = sLine.Range.Text ' get the text of the paragraph
                ' First time into a new group - we need to grab the group code...
                ' Get first letter
                strGroupLetter = Mid(strLineText, 1, 1)
                If strGroupLetter = strPreviousGroupLetter Then
                    ' still in the same group - ignore
                Else
                    strPreviousGroupLetter = strGroupLetter ' set it for next time
                    strKeyNoteText = strKeyNoteText & strGroupLetter & vbTab & GetGroupTitleCAWS(strGroupLetter) & vbCrLf
                
                End If
            
            
                ' We have text of the format E10Concrete
                ' Splitting on the  character
                intPos = InStr(1, strLineText, "", vbTextCompare)
                strSectionCode = Mid(strLineText, 1, intPos - 1)
                strSectionTitle = Mid(strLineText, intPos + 1, Len(strLineText) - intPos - 1)
                strSectionCode = Trim(strSectionCode)
                strSectionTitle = Trim(strSectionTitle)
                strKeyNoteText = strKeyNoteText & strSectionCode & vbTab & strSectionTitle & vbTab & strGroupLetter & vbCrLf
                
                
            Case "chorus-clause-title"
                strLineText = sLine.Range.Text ' get the text of the paragraph
                ' We have text of the format 120 RC Concrete
                ' Splitting on the [space] character
                intPos = InStr(1, strLineText, " ", vbTextCompare)
                strClauseCode = Mid(strLineText, 1, intPos - 1)
                strClauseTitle = Mid(strLineText, intPos + 1, Len(strLineText) - intPos - 1)
                strClauseCode = Trim(strClauseCode)
                strClauseTitle = Trim(strClauseTitle)
                strKeyNoteText = strKeyNoteText & strSectionCode & "/" & strClauseCode & vbTab & strClauseTitle & vbTab & strSectionCode & vbCrLf
                
        End Select
        
    Next sLine
    
    
    System.Cursor = wdCursorNormal
    
    ReturnCAWSKeyNoteText = strKeyNoteText

End Function



Private Function ReturnUniclassKeyNoteText() As String


    Dim sLine As Paragraph
    Dim strLineText
    
    Dim strGroupCode As String
    Dim strSectionCode As String
    Dim strSectionTitle As String
    
    Dim strClauseCode As String
    Dim strClauseTitle As String
    Dim intPos As Integer
    
    Dim strGroupLetter As String
    Dim strPreviousGroupLetter As String
    Dim strGroupTitle As String
    
    
    System.Cursor = wdCursorWait
    
    
    ' Loop through every paragraph in the word document
    ' looking out for headings, clauses and normal text...
    For Each sLine In ActiveDocument.Paragraphs
        
        
        Select Case sLine.Style
            ' First time into a new group - we need to grab the group code...
            
            Case "chorus-section-header"
                strLineText = sLine.Range.Text ' get the text of the paragraph
                ' First time into a new group - we need to grab the group code...
                ' Get first letter
                strGroupLetter = Mid(strLineText, 1, 5)
                If strGroupLetter = strPreviousGroupLetter Then
                    ' still in the same group - ignore
                Else
                    strPreviousGroupLetter = strGroupLetter ' set it for next time
                    strKeyNoteText = strKeyNoteText & strGroupLetter & vbTab & GetGroupTitleUniclass(strGroupLetter) & vbCrLf
                
                End If
            
            
                ' We have text of the format E10Concrete
                ' Splitting on the  character
                intPos = InStr(1, strLineText, "", vbTextCompare)
                strSectionCode = Mid(strLineText, 1, intPos - 1)
                strSectionTitle = Mid(strLineText, intPos + 1, Len(strLineText) - intPos - 1)
                strSectionCode = Trim(strSectionCode)
                strSectionTitle = Trim(strSectionTitle)
                strKeyNoteText = strKeyNoteText & strSectionCode & vbTab & strSectionTitle & vbTab & strGroupLetter & vbCrLf
                
                
            Case "chorus-clause-title"
                strLineText = sLine.Range.Text ' get the text of the paragraph
                ' We have text of the format 120 RC Concrete
                ' Splitting on the [space] character
                intPos = InStr(1, strLineText, " ", vbTextCompare)
                strClauseCode = Mid(strLineText, 1, intPos - 1)
                
                If InStr(1, strClauseCode, "/", vbTextCompare) = 0 Then  ' we only want system and product clauses - not syst perf/exec etc...
                    strClauseTitle = Mid(strLineText, intPos + 1, Len(strLineText) - intPos - 1)
                    strClauseCode = Trim(strClauseCode)
                    strClauseTitle = Trim(strClauseTitle)
                    strKeyNoteText = strKeyNoteText & strClauseCode & vbTab & strClauseTitle & vbTab & strSectionCode & vbCrLf
                End If
                
        End Select
        
    Next sLine
    
    System.Cursor = wdCursorNormal
    
    ReturnUniclassKeyNoteText = strKeyNoteText

End Function



Private Function IsSectionHeader(strText)
    ' We're looking for 'F10->' as a pattern

    ' Are characters 2 and 3 numeric?
    Dim strSectionCode
    strSectionCode = ""
    strSectionCode = Mid(strText, 2, 2)
    
    Dim boolIsNumeric
    boolIsNumeric = False
    boolIsNumeric = IsNumeric(strSectionCode)
    
    
    ' Is character 4 a tab?
    Dim strTabCheck
    strTabCheck = ""
    
    If boolIsNumeric = True Then
        strTabCheck = Mid(strText, 4, 1)
        If strTabCheck = vbTab Then
            ' We have the 'F10->' pattern return true
            IsSectionHeader = True
            Exit Function
        End If
    End If
    
    ' Return false for all other cases
    IsSectionHeader = False

End Function

Private Function IsClauseTitle(strText)
    ' We're looking for 'F10/10' as a pattern

    Dim boolIsNumeric
    boolIsNumeric = False
        

    ' Are characters 2 and 3 numeric?
    Dim strSectionCode
    strSectionCode = ""
    strSectionCode = Mid(strText, 2, 2)
    
    boolIsNumeric = False
    boolIsNumeric = IsNumeric(strSectionCode)
    
    If boolIsNumeric = False Then
        IsClauseTitle = False
        Exit Function
    End If
    
    
    ' Is character 4 a slash?
    Dim strSlash
    If Mid(strText, 4, 1) = "/" Then
            ' We have the 'F10/' pattern return true
            IsClauseTitle = True
            Exit Function
    End If
    
    
    
    
    ' Return false for all other cases
    IsClauseTitle = False

End Function

Private Function StripUnwantedCharacters(sText)
    sText = Replace(sText, Chr(12), "") ' strip out any carriage returns
    sText = Replace(sText, "", "") ' strip out any page breaks (bit bodgey) :)
    
    StripUnwantedCharacters = sText

End Function

' Note objP must be a Word paragraph object
Private Function DetermineParagraphType(sCode, sLine, objP)
    
    If Len(sLine) < 3 Then
        DetermineParagraphType = enNBSBlank ' treat less then 3 characters as a line to ignore
        Exit Function
    End If

    If IsSectionHeader(sLine) = True Then
        DetermineParagraphType = enNBSSectionHeading
        Exit Function
    End If

    If IsClauseTitle(sLine) = True Then
        DetermineParagraphType = enNBSClause
        Exit Function
    End If

    ' If all fails, we assume it's just a row
    DetermineParagraphType = enNBSClauseRow

End Function


Private Sub RemoveAllBreaks()
    Dim arr() As Variant
    Dim i As Byte
    
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    '"^b" - Selection Breaks, "^m" - Page Break, "^n" - Column Break
    arr = Array("^b", "^m", "^n")
    For i = LBound(arr) To UBound(arr)
        With Selection.Find
            .Text = arr(i)
            .Replacement.Text = ""
        End With
    Selection.Find.Execute Replace:=wdReplaceAll
    Next
End Sub

Private Sub SaveDebugInfoToTextFile()

    Dim filePath As String
    filePath = "C:\temp\NBS-Debug.txt"

    ' The advantage of correctly typing fso as FileSystemObject is to make autocompletion
    ' (Intellisense) work, which helps you avoid typos and lets you discover other useful
    ' methods of the FileSystemObject
    Dim fso As FileSystemObject
    Set fso = New FileSystemObject
    Dim fileStream As TextStream

    ' Delete previous file if it exists
    If fso.FileExists(filePath) = True Then
        fso.DeleteFile (filePath)
    End If

    ' Here the actual file is created and opened for write access
    Set fileStream = fso.CreateTextFile(filePath)

    ' Write something to the file
    fileStream.WriteLine g_strDebugText

    ' Close it, so it is not locked anymore
    fileStream.Close

    ' Here is another great method of the FileSystemObject that checks if a file exists
    ' If fso.FileExists(filePath) Then
    '    MsgBox "Yay! The file was created! :D"
    ' End If

    ' Explicitly setting objects to Nothing should not be necessary in most cases, but if
    ' you're writing macros for Microsoft Access, you may want to uncomment the following
    ' two lines (see https://stackoverflow.com/a/517202/2822719 for details):
    'Set fileStream = Nothing
    'Set fso = Nothing

End Sub

Private Function GetGroupTitleCAWS(valLetter) As String
    Dim strReturn As String
    strReturn = "[UNDEFINED]"
    
    Select Case valLetter
        Case "A"
            strReturn = "Preliminaries/General conditions"
        Case "B"
            strReturn = "Complete buildings/structures/units"
        Case "C"
            strReturn = "Existing site/buildings/services"
        Case "D"
            strReturn = "Groundwork"
        Case "E"
            strReturn = "In situ concrete/Large precast concrete"
        Case "F"
            strReturn = "Masonry"
        Case "G"
            strReturn = "Structural/Carcassing metal/timber"
        Case "H"
            strReturn = "Cladding/Covering"
        Case "I"
            'strReturn = ""
        Case "J"
            strReturn = "Waterproofing"
        Case "K"
            strReturn = "Linings/Sheathing/Dry partitioning"
        Case "L"
            strReturn = "Windows/Doors/Stairs"
        Case "M"
            strReturn = "Surface finishes"
        Case "N"
            strReturn = "Furniture/Equipment"
        Case "O"
            'strReturn = ""
        Case "P"
            strReturn = "Building fabric sundries"
        Case "Q"
            strReturn = "Paving/Planting/Fencing/Site furniture"
        Case "R"
            strReturn = "Disposal systems"
        Case "S"
            strReturn = "Piped supply systems"
        Case "T"
            strReturn = "Mechanical heating, cooling and refrigeration systems"
        Case "U"
            strReturn = "Ventilation and air conditioning systems"
        Case "V"
            strReturn = "Electrical systems"
        Case "W"
            strReturn = "Communications, security, safety and protection systems"
        Case "X"
            strReturn = "Transport systems"
        Case "Y"
            strReturn = "General engineering services"
        Case "Z"
            strReturn = "Building fabric reference specification"
    
    End Select
    GetGroupTitleCAWS = strReturn
End Function

Private Function GetGroupTitleUniclass(valLetter) As String
    Dim strReturn As String
    strReturn = "[UNDEFINED]"
    
    Select Case valLetter
        Case "Ss_15"
            strReturn = "Earthworks, remediation and temporary systems"
        Case "Ss_20"
            strReturn = "Structural systems"
        Case "Ss_25"
            strReturn = "Wall and barrier systems"
        Case "Ss_30"
            strReturn = "Roof, floor and paving systems"
        Case "Ss_32"
            strReturn = "Damp-proofing, waterproofing and plaster-finishing systems"
        Case "Ss_35"
            strReturn = "Stair and ramp systems"
        Case "Ss_37"
            strReturn = "Tunnel, shaft, vessel and tower systems"
        Case "Ss_40"
            strReturn = "Signage, fittings, furnishings and equipment (FF&E) and general finishing systems"
        Case "Ss_45"
            strReturn = "Flora and fauna systems"
        Case "Ss_50"
            strReturn = "Disposal systems"
        Case "Ss_55"
            strReturn = "Piped supply systems"
        Case "Ss_60"
            strReturn = "Heating, cooling and refrigeration systems"
        Case "Ss_65"
            strReturn = "Ventilation and air conditioning systems"
        Case "Ss_70"
            strReturn = "Electrical systems"
        Case "Ss_75"
            strReturn = "Communications, security, safety, control and protection systems"
        Case "Ss_80"
            strReturn = "Transport systems"
        Case "Ss_85"
            strReturn = "Process engineering systems"
        Case "Ss_90"
            strReturn = "Soft facility management systems"
    
            ' ToDO
            ' For child products the top level product titles need added - same for the activities
    
    End Select
    GetGroupTitleUniclass = strReturn
End Function
