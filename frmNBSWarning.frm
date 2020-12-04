VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmNBSWarning 
   Caption         =   "Pre-processing support"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   OleObjectBlob   =   "frmNBSWarning.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmNBSWarning"
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
' frmNBSWarning
' =============
' Not much, just flashes a warning message up - gives the opportunity for
' the user to cancel
' ========================================================================

Dim bApplyChanges

Public Function ApplyChanges()
    ApplyChanges = bApplyChanges
End Function

Private Sub cmdCancel_Click()
    bApplyChanges = False
    Hide
End Sub

Private Sub cmdOK_Click()
    bApplyChanges = True
    Hide
End Sub

Private Sub UserForm_Initialize()
    bApplyChanges = False
End Sub
