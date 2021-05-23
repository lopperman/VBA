VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ProtectionBase"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
''
' RangeMonger v0.1.3
' (c) Paul Brower - https://github.com/lopperman/VBA
'
' Base class for defined Protect/Unprotect functions
'
' @class ProtectionBase
' @author Paul Brower
' @license GNU General Public License v3.0
' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ '
Option Explicit


Public Function Protect(wkst As Worksheet) As Boolean
End Function

Public Function Unprotect(wkst As Worksheet) As Boolean
End Function

Public Sub UnprotectAll(wkbk As Workbook)
End Sub

Public Sub UnprotectAllForced(wkbk As Workbook)
End Sub

Public Property Get PreventProtection() As Boolean
End Property
Public Property Let PreventProtection(preventProt As Boolean)
End Property

