Attribute VB_Name = "Module2"
Sub Macro1Planning()
Attribute Macro1Planning.VB_Description = "Macro de planning pour le totem espace formation\n"
Attribute Macro1Planning.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Macro1Planning Macro
' Macro de planning pour le totem espace formation
'

'
    Range("B:B,C:C").Select
    Range("C1").Activate
    Selection.NumberFormat = "h:mm"
    Range("A1").Select
    ActiveCell.FormulaR1C1 = "Classe"
    Range("B1").Select
    ActiveCell.FormulaR1C1 = "Event heure debut"
    Range("L16").Select
    ActiveCell.FormulaR1C1 = ""
    Range("B1").Select
    ActiveCell.FormulaR1C1 = "Heure_Debut"
    Range("C1").Select
    ActiveCell.FormulaR1C1 = "Heure_de_Fin"
    Range("D1").Select
    ActiveCell.FormulaR1C1 = "Salle"
    Range("D2").Select
End Sub
