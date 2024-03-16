# Automatically add date and time to filename when saving a word document

1.
![Snimka zaslona 2024-03-16 164349](https://github.com/bojkip/nekoime/assets/91488932/7e42c018-2add-4a28-b3f7-73bf21e0f910)

Alt + F8 or Alt + F11


2.
![Snimka zaslona 2024-03-16 164511](https://github.com/bojkip/nekoime/assets/91488932/b31f8b60-315a-4455-81c4-5953aad3a2de)

Copy⬇️ and paste⬆️ code:

```

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



```


4.
![Snimka zaslona 2024-03-16 165609](https://github.com/bojkip/nekoime/assets/91488932/15ff0cdf-2efc-4625-bf42-1f05691643bd)

    4.1
      


6.
![Snimka zaslona 2024-03-16 170200](https://github.com/bojkip/nekoime/assets/91488932/6843c56c-55be-4169-bae0-1852f0b101d0)






