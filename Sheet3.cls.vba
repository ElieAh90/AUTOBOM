Attribute VB_Name = "Sheet3"
Attribute VB_Base = "0{00020820-0000-0000-C000-000000000046}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = True
Attribute VB_Control = "cb_AutoSelectAllReferences, 3, 0, MSForms, CheckBox"
Attribute VB_Control = "cb_AutoSingleStandard, 2, 1, MSForms, CheckBox"
Public Function IsAutoSelectSingleStandard() As Boolean
    IsAutoSelectSingleStandard = cb_AutoSingleStandard.value
End Function

Public Function IsAutoSelectAllReferences() As Boolean
    IsAutoSelectAllReferences = cb_AutoSelectAllReferences.value
End Function

Public Function FixturesUsedIn(index As Integer) As Integer
    Dim Count As Integer
    Count = 0
    Dim i As Integer
    
    For i = 1 To lengthsCount
        Count = Count + usedStandards(index, i)
    Next i
    
    FixturesUsedIn = Count

End Function
