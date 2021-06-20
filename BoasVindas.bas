Attribute VB_Name = "Módulo2"
Sub BoasVindas()
    Dim MyOlapp     As Object, MeuItem As Object
    Dim Suit        As String
    Dim EmailCopia  As String
    Dim Email       As String
    Dim Linha       As Integer
    Dim PauseTime   As Integer
    Dim Start       As Single
    Dim result      As Byte
    Dim Cliente     As String
    Dim Corpo       As String
    
   
    result = MsgBox("Lembrou de trocar a assinatura?", vbOKCancel)
    
    If result = 1 Then
    
         Linha = Sheets("Bem-Vindo").Cells(Sheets("Bem-Vindo").Rows.Count, 1).End(xlUp).Row
         
         Set MyOlapp = CreateObject("Outlook.Application")
         PauseTime = Range("H2")
         
        
        Do While Linha >= 2
             EmailCopia = Range("D" & Linha)
             Email = Range("B" & Linha)
             Suit = Range("E" & Linha)
             Cliente = Range("A" & Linha)
         
             Set MeuItem = MyOlapp.CreateItem(olMailItem)
             With MeuItem
                 
                 .to = Email
                 .CC = EmailCopia & ";" & "cadastro@fiduc.com.br"
                 .Subject = "BEM-VINDO À FIDUC, SEU PERFIL SUITABILITY É " & Suit
                 .Display
                 Corpo = "<font size=4 color=1F497D face=calibri>Olá, <br >" & Cliente
                 Corpo = Corpo & "<br>"
                 .HTMLBody = Corpo & .HTMLBody
                 .Send
                 
             End With
             Start = Timer    ' Set start time.
             Do While Timer < Start + PauseTime
                 DoEvents    ' Yield to other processes.
             Loop
             
             
             Linha = Linha - 1
             
         Loop
         
         MsgBox "Troxa!"
    
    End If

End Sub




