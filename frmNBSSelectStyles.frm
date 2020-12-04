VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmNBSSelectStyles 
   Caption         =   "Select styles"
   ClientHeight    =   2805
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4245
   OleObjectBlob   =   "frmNBSSelectStyles.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmNBSSelectStyles"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' ========================================================================
'
' Macro code written to help work with MS Word
' manipulation prior to import into NBS Chorus
'
' Contributors:
' 1. Stephen Hamil
'
' Use at own risk
' Always save and keep copies of Word documents before running the macros
'
' Support Community - https://support.thenbs.com/support/home
'
' modNBSMacros          5th Oct 2020
' frmNBSSelectStyles    5th Oct 2020
' frmNBSWarning         5th Oct 2020
'
' frmNBSSelectStyles
' ==================
' Data entry form for font and RGB colours
' Basic level of validation takes place once user clicks OK
' Anyone want to improve this and potential error handling - fill your boots :)
' ========================================================================

Dim bApplyChanges


Public Function ApplyChanges()
    ApplyChanges = bApplyChanges
End Function

Public Function FontFamily()
    FontFamily = txtFontFamily.value
End Function

Public Function HeadingColourR()
    HeadingColourR = txtHeadingsColourR.value
End Function

Public Function HeadingColourG()
    HeadingColourG = txtHeadingsColourG.value
End Function

Public Function HeadingColourB()
    HeadingColourB = txtHeadingsColourB.value
End Function


Private Sub cmdOK_Click()
    ' Check for valid values before hiding.
    
    ' Is there a valid font value (at least check it's not pretty much blank
    If Len(FontFamily) < 3 Then
        MsgBox "Please enter a valid font", vbOKOnly, "Select styles"
        txtFontFamily.SetFocus
        Exit Sub
    End If
    
    ' Is there valid RGB values - we already limit the data entry to 3 digits
    If CheckEntryForInteger(HeadingColourR) = False Then
        MsgBox "Please enter a three digit number for the R value", vbOKOnly, "Select styles"
        txtHeadingsColourR.SetFocus
        Exit Sub
    End If
    
    ' Is there valid RGB values - we already limit the data entry to 3 digits
    If CheckEntryForInteger(HeadingColourG) = False Then
        MsgBox "Please enter a three digit number for the G value", vbOKOnly, "Select styles"
        txtHeadingsColourG.SetFocus
        Exit Sub
    End If
    
    ' Is there valid RGB values - we already limit the data entry to 3 digits
    If CheckEntryForInteger(HeadingColourB) = False Then
        MsgBox "Please enter a three digit number for the B value", vbOKOnly, "Select styles"
        txtHeadingsColourB.SetFocus
        Exit Sub
    End If
    
    bApplyChanges = True
    Hide
End Sub

Private Sub cmdCancel_Click()
    bApplyChanges = False
    Hide
End Sub


Private Sub UserForm_Initialize()
    bApplyChanges = False
End Sub

Private Function CheckEntryForInteger(valTextEntry)
    If IsNumeric(valTextEntry) Then
        If valTextEntry = 0 Then
            CheckEntryForInteger = True
            Exit Function
        End If
        ' Here, it still could be an integer or a floating point number
        If CLng(valTextEntry) Then
           CheckEntryForInteger = True
        Else
           CheckEntryForInteger = False
        End If
    End If
End Function
