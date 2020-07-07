Public Declare Function SetCursorPos Lib "user32" (ByVal x As Long, ByVal y As Long) As Long
Public Declare Sub mouse_event Lib "user32" (ByVal dwFlags As Long, ByVal dx As Long, ByVal dy As Long, ByVal cButtons As Long, ByVal dwExtraInfo As Long)
Public WrkB As Workbook
Public WrkS As Worksheet

Public IntervaloRotina As Range
Public Celula          As Range

'Dim Account As String


Declare Sub Sleep Lib "kernel32" _
(ByVal dwMilliseconds As Long)
Public Sub CadastroSMC()

Set WrkB = ThisWorkbook
Set WrkS = WrkB.Sheets("Cadastro_SMC")

Set IntervaloRotina = WrkS.Range("A5:A100000")


With WrkS
    .Select
        For Each Celula In IntervaloRotina
            Call InicioHemera
            'Sleep (8000)
            Next
        
End With

End Sub
Public Sub InicioHemera()
'On Error GoTo tratar_erro
    Dim IE As New ieRV
    Dim login As String
    Dim senha As String
    Dim tela As String
    
    Dim IEobj As Object
    Set IEobj = CreateObject("InternetExplorer.application")
      
    Dim medidor As Double
    Dim Button As HTMLInputElement
    login = InputBox("Inserir Login de Rede ")
    
    senha = InputBox("Inserir Senha de Rede ")
       
    
    With IE
        .iniciaIE
        .NAVEGAR "http://portal", SW_SHOWMAXIMIZED 'Entrar Na Página Web
        .wait (3000) 'Aguardar 5 Segundos
        .getElement(1, "name", "username").innerText = login 'Inserir Login
        .getElement(1, "name", "password").innerText = senha 'Inserir Senha
        .getElement(1, "id", "ext-gen22").Click 'Clicar Na Rolagem para Seleção
        .wait (1000) 'Aguardar 1 Segundo
        SendKeys "{UP}", True 'Enviar Comando do Botão Para Cima
        .wait (1000) 'Aguardar 1 Segundo
        SendKeys "{UP}", True 'Enviar Comando do Botão Para Cima
        .wait (1000) 'Aguardar 1 Segundo
        SendKeys "{UP}", True 'Enviar Comando do Botão Para Cima
        .wait (1000) 'Aguardar 1 Segundo
        SendKeys "~", True 'Pressionar Enter
        .getElement(1, "id", "divCenterButton").Click 'Botão Entrar
        .wait (8000) 'Aguardar 8 segundos
        .getElement(1, "id", "ext-gen119").Click
        .wait (5000) 'Aguardar 5 segundos
        .getElement(1, "id", "ext-gen72").Click
        .wait (5000) 'Aguardar 5 segundos
        .getElement(1, "id", "ext-comp-1022-span-collapse").Click 'Clicar no botão Grupo B para expandir
        .wait (5000) 'Aguardar 5 segundos
        .getElement(1, "name", "txtShuntSerial").Click 'Clicar no campo do nº do medidor
        .getElement(1, "name", "txtShuntSerial").innerText = Cells(Celula.Row, 1).Value 'Inserir nº do medidor
        .wait (1000) 'Aguardar 1 segundo
        SendKeys "{ENTER}" 'Pressionar enter para pesquisar nº do medidor
        .wait (8000) 'Aguardar 8 segundos
        'For Each Button In IEobj.Document.getElementsByTagName("Pesquisar")
        'Button.Click
        'Next
        .getElement(1, "id", "ext-gen660").Click 'Clicar no medidor
        .wait (5000) 'Aguardar 5 segundos
        .getElement(1, "id", "ext-gen29").Click 'Clicar no botão geral
        .wait (3000) 'Aguardar 3 segundos
        .getElement(1, "id", "ext-gen75").Click 'Clicar em alterar medidor
        .wait (3000) 'Aguardar 3 segundos
        .getElement(1, "id", "ext-gen129").Click 'Clicar em selecionar medidor
        '.wait (1000) 'Aguardar 1 segundo
        '.waitElem(1, "id", "ext-gen389").Value = "Seleção de UC" 'Aguardar tela de seleção de uc
        .wait (3000) 'Aguardar 1 segundo
        .getElement(1, "name", "searchName").Click 'Clicar em nome
        .getElement(1, "name", "searchName").innerText = Cells(Celula.Row, 2).Value 'Inserir instalação
        .wait (1000) 'Aguardar 1 segundo
        .getElement(1, "id", "ext-gen448").Click 'Clicar em pesquisar
        .wait (3000) 'Aguardar 3 segundos
        .getElement(1, "id", "ext-gen109").Click 'Clicar no texto, Sem registros para exibir
        'If .getElement(1, "id", "ext-gen109").Value = "1 - 1 | Total 1" Then 'Validar se há instalação em outro medidor
        '.getElement(1, "id", "ext-gen767").Click 'Fechar a aba alterar medidor
        'WrkS.Cells(Celula.Row, 3) = "Instalação Associada a Outro Medidor"
        'End If
        .wait (2000) 'Aguardar 2 segundos
        .getElement(1, "id", "ext-gen439").Click 'Clicar em Nova UC
        .wait (4000) 'Aguardar 2 segundos
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        '.waitElem(1, "id", "ext-gen719") = True
        '.getElement(1,
        
        
        
     
        
          '.closeAllIE
    End With
    
Exit Sub
'tratar_erro:

'Resume Next


'With ie
  
'End With
medidor = Empty

End Sub
