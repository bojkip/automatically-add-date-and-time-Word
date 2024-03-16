Attribute VB_Name = "NewMacros"
Sub FileSave()
    ' Run as substitute for FileSave to add date to document title

    '    for filename

    ' Put in a document or document template

    '    but probably not in a global
    '
    On Error Resume Next
    '
    Dim strName As String, dlgSave As Dialog
    Set dlgSave = Dialogs(wdDialogFileSaveAs)
    vrijeme = Time

    strName = ActiveDocument.BuiltInDocumentProperties("Title").Value

    '    get name in title
    strName = strName & ActiveDocument.Name & "_" & Format((Day(Now()) Mod 100), "0#") & "-" & _
    Format((Month(Now() + 1) Mod 100), "0#") & "-" & _
    Format((Year(Now() + 1) Mod 100), _
        "20##") & "_" & _
    Format(Now, "hh-mm") & "h"

    With dlgSave
        .Name = strName
        .Show
        
    End With
End Sub

Sub FileSaveAs()
    Call FileSave
End Sub

