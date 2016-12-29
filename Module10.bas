Attribute VB_Name = "Module10"
Sub Reset_Filters()
Attribute Reset_Filters.VB_Description = "Macro recorded 12/14/2007 by Eloise Roche"
Attribute Reset_Filters.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Reset_Filters Macro
' Macro recorded 12/14/2007 by Eloise Roche
'
ActiveSheet.ShowAllData
End Sub
Sub CopiarSensors2AI()
Attribute CopiarSensors2AI.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Macro made by Freddy Gonzalez
' Tica Polyols Cartagena

'
    Dim PosX As Integer
    Dim PosY As Integer
    Dim Celda As Range
    Dim CeldaMaster As Range
        
'   Se selecciona la primera celda de la columna donde se va a copiar los datos
    
    Set Celda = ActiveCell
        
    Windows("V2.200136_LOPADB DowGEP Mod5 Ass WB.xls").Activate
    Set CeldaMaster = ActiveCell 'Selecciona celda en segundo Workbook
    
    Windows("COPY2 Cartagena DowGEP AI Prioritization harmoniization.xls").Activate
  
'--------------------------------------------------------------------
    Celda.Offset(0, 2).Value = "" & CeldaMaster.Offset(1, 0).Value
    
    Celda.Offset(0, 3).Value = "" & CeldaMaster.Offset(1, 0).Value
   
    Celda.Offset(0, 4).Value = "" & CeldaMaster.Offset(2, 0).Value
    
    Celda.Offset(0, 5).Value = "SENSOR"

    Celda.Offset(0, 6).Value = "Especificar y comprar"
    
    Celda.Offset(0, 7).Value = "Mandatory"
    
     
End Sub

Sub CopiarBPCS2AI()
Attribute CopiarBPCS2AI.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Macro made by Freddy Gonzalez
' Tica Polyols Cartagena

'
    Dim PosX As Integer
    Dim PosY As Integer
    Dim Celda As Range
    Dim CeldaMaster As Range
        
'   Se selecciona la primera celda de la columna donde se va a copiar los datos
    
    Set Celda = ActiveCell
        
    Windows("V2.200136_LOPADB DowGEP Mod5 Ass WB.xls").Activate
    Set CeldaMaster = ActiveCell 'Selecciona celda en segundo Workbook
    
    Windows("COPY2 Cartagena DowGEP AI Prioritization harmoniization.xls").Activate
  
'--------------------------------------------------------------------
    Celda.Offset(0, 2).Value = "" & CeldaMaster.Offset(1, 0).Value
    
    Celda.Offset(0, 3).Value = "" & CeldaMaster.Offset(1, 0).Value
   
    Celda.Offset(0, 4).Value = "" & CeldaMaster.Offset(2, 0).Value
    
    Celda.Offset(0, 5).Value = "BPCS/ALM"

    Celda.Offset(0, 6).Value = "Implement Alkoxylation MET2 DowGEP IPL ACM Code"
    
    Celda.Offset(0, 7).Value = "Mandatory"
    
    Celda.Offset(0, 8).Value = "Functional"
     
End Sub

Sub CopiarSIS2AI()
Attribute CopiarSIS2AI.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Macro made by Freddy Gonzalez
' Tica Polyols Cartagena

'
    Dim PosX As Integer
    Dim PosY As Integer
    Dim Celda As Range
    Dim CeldaMaster As Range
        
'   Se selecciona la primera celda de la columna donde se va a copiar los datos
    
    Set Celda = ActiveCell
        
    Windows("V2.200136_LOPADB DowGEP Mod5 Ass WB.xls").Activate
    Set CeldaMaster = ActiveCell 'Selecciona celda en segundo Workbook
    
    Windows("COPY2 Cartagena DowGEP AI Prioritization harmoniization.xls").Activate
  
'--------------------------------------------------------------------
    Celda.Offset(0, 2).Value = "" & CeldaMaster.Offset(1, 0).Value
    
    Celda.Offset(0, 3).Value = "" & CeldaMaster.Offset(1, 0).Value
   
    Celda.Offset(0, 4).Value = "" & CeldaMaster.Offset(3, 0).Value
    
    Celda.Offset(0, 5).Value = "SIL" & CeldaMaster.Offset(4, 0).Value

    Celda.Offset(0, 6).Value = "Implement Alkoxylation MET2 DowGEP IPL ACM Code"
    
    Celda.Offset(0, 7).Value = "Mandatory"
    
    Celda.Offset(0, 8).Value = "Functional"
     
End Sub

Sub RemedialSIS2AI()
Attribute RemedialSIS2AI.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Macro made by Freddy Gonzalez
' Tica Polyols Cartagena

'
    Dim PosX As Integer
    Dim PosY As Integer
    Dim Celda As Range
    Dim CeldaMaster As Range
        
'   Se selecciona la primera celda de la columna donde se va a copiar los datos
    
    Set Celda = ActiveCell
        
    Windows("V2.200136_LOPADB DowGEP Mod5 Ass WB.xls").Activate
    Set CeldaMaster = ActiveCell 'Selecciona celda en segundo Workbook
    
    Windows("COPY2 Cartagena DowGEP AI Prioritization harmoniization.xls").Activate
  
'--------------------------------------------------------------------
    Celda.Offset(0, 3).Value = "" & CeldaMaster.Offset(1, 0).Value
    
'   Celda.Offset(0, 3).Value = "" & CeldaMaster.Offset(1, 0).Value
   
    Celda.Offset(0, 4).Value = "" & CeldaMaster.Offset(3, 0).Value
    
    Celda.Offset(0, 5).Value = "SIL" & CeldaMaster.Offset(4, 0).Value

'   Celda.Offset(0, 6).Value = "Implement Alkoxylation MET2 DowGEP IPL ACM Code"
    
'   Celda.Offset(0, 7).Value = "Mandatory"
    
'   Celda.Offset(0, 8).Value = "Functional"
     
End Sub

Sub RemedialBPCS2AI()
Attribute RemedialBPCS2AI.VB_ProcData.VB_Invoke_Func = "t\n14"
'
' Macro made by Freddy Gonzalez
' Tica Polyols Cartagena

'
    Dim PosX As Integer
    Dim PosY As Integer
    Dim Celda As Range
    Dim CeldaMaster As Range
        
'   Se selecciona la primera celda de la columna donde se va a copiar los datos
    
    Set Celda = ActiveCell
        
    Windows("V2.200136_LOPADB DowGEP Mod5 Ass WB.xls").Activate
    Set CeldaMaster = ActiveCell 'Selecciona celda en segundo Workbook
    
    Windows("COPY2 Cartagena DowGEP AI Prioritization harmoniization.xls").Activate
  
'--------------------------------------------------------------------
    Celda.Offset(0, 3).Value = "" & CeldaMaster.Offset(1, 0).Value
    
'   Celda.Offset(0, 3).Value = "" & CeldaMaster.Offset(1, 0).Value
   
    Celda.Offset(0, 4).Value = "" & CeldaMaster.Offset(2, 0).Value
    
'    Celda.Offset(0, 5).Value = "SIL" & CeldaMaster.Offset(4, 0).Value

'   Celda.Offset(0, 6).Value = "Implement Alkoxylation MET2 DowGEP IPL ACM Code"
    
'   Celda.Offset(0, 7).Value = "Mandatory"
    
'   Celda.Offset(0, 8).Value = "Functional"
     
End Sub
