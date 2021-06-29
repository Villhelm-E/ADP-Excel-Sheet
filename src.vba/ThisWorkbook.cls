Option Explicit

'Run on Workbook open
Private Sub Workbook_Open()

''''Connect to Master Database
    ConnectMasterDatabase
    
''''Connect to ADP Find Sets Database
    ConnectFindSetsDatabase
    
''''Connect to Sixbit Database
    ConnectSixbitDatabase
    
    'toggle for debugging
    DebugMode = False 'restart excel sheet to save changes

End Sub

'Run whenever any sheet changes
Private Sub Workbook_SheetChange(ByVal sh As Object, ByVal target As range)
    
    'If code writes anything to worksheet, screenupdating will be turned off bypassing this subroutine
    'will only update ribbon if code isn't making changes to sheet
    If DebugMode = False Or Application.ScreenUpdating = True Then    'DebugMode is for debugging (specifically stepping into the code to skip refreshing the ribbon all the time)
        RibbonCategories    'Ribbon Module
    End If

End Sub

'Run whenever user switches sheet
Private Sub Workbook_SheetActivate(ByVal sh As Object)
    
    RibbonCategories    'Ribbon Module

End Sub
