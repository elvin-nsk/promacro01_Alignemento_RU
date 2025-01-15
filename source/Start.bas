Attribute VB_Name = "Start"

Public Sub Alignement_Shapes()
 
    If ActiveSelection.Shapes.Count < 2 Then _
        MsgBox "Выберите как минимум 2 объекта.", vbExclamation: Exit Sub
     
    With New ProMacro01
        .Show
    End With
  
End Sub
