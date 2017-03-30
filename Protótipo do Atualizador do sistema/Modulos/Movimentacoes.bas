Attribute VB_Name = "Movimentacoes"
'Declaração das variaveis da acessibilidade
Dim strID_Acessibilidade As String
Dim rstAplicacao As New ADODB.Recordset
Dim acesso As New DLLSystemManager.Acessibilidade
Dim log As New DLLSystemManager.log

Function Movimenta_Data_Grid(strSql As String, datagrid As Object, tamanho_colunas As String, caption_campos As String, Banco As String, Aplicacao As String, Form As Object)
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Only Tech                                                                                  '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Sistema................: PDV Only Tech                                                     '
' Equipe Responsável.....: Giordano Vilela,Marcos Baião                                   '
' Data da criação........: 24/02/2003                                                     '
' Data última manutenção.:                                                                '
' Observação.............: Public Property Let --> Recebe um valor para a propriedade     '
'                          (em run time)                                                  '
'                          Public Property Get --> Retorna para a aplicação um valor      '
'                                                                                         '
'                         *Esta Classe movimenta o controle datagrid. Cuidados Especiais  '
'                          Apos o uso não esquecer de descarregar a                       '
'                                                                                         '
'                                                                                         '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    Dim rstgrid As New ADODB.Recordset
    Dim vdatagrid As Object
    Dim Contador As Integer
    Dim Colunas As Integer
    Dim matriz_tamanho() As String
    Dim matriz_caption() As String
    Dim conexao_grid As New DLLConexao_Sistema.conexao
    
    'On Error GoTo Erro
    
    Set vdatagrid = datagrid
    
    'Indicando o banco à conectar-se
    conexao_grid.Initial_Catalog = Banco
    
    conexao_grid.Abrir_conexao (Aplicacao)
       
    rstgrid.CursorLocation = adUseClient
    rstgrid.Open strSql, conexao_grid.CNConexao, adOpenKeyset, adLockOptimistic, adCmdText
    
    Set vdatagrid = datagrid
    Set vdatagrid.DataSource = rstgrid
    
    'Montando a matriz
    matriz_tamanho = Split(tamanho_colunas, ",")
    matriz_caption = Split(caption_campos, ",")
    
    Contador = 0
    Colunas = 0
    
    'Montando as características do grid - Caption e Tamanho
    Do While matriz_caption(Contador) <> Empty
        vdatagrid.Columns(Colunas).Caption = matriz_caption(Contador)
        vdatagrid.Columns(Colunas).Width = Val(matriz_tamanho(Contador))
        Contador = Contador + 1
        Colunas = Colunas + 1
        If Contador > UBound(matriz_caption) Then Exit Do
    Loop
   
   ' Exit Function
    
'Erro:
'    Call Erro.Erro(Form, Aplicacao, "Movimenta_Data_Grid")
    
End Function
Function Movimenta_HFlex_Grid(strSql As String, HflexGrid As Object, tamanho_colunas As String, caption_campos As String, Banco As String, Aplicacao As String, Form As Object, Optional controle_focu As String)

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Only Tech                                                                                  '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Sistema................: PDV Only Tech                                                     '
' Equipe Responsável.....: Giordano Vilela,Marcos Baião                                   '
' Data da criação........: 05/08/2003                                                     '
' Data última manutenção.:                                                                '
' Observação.............:                                                                '
'                                                                                         '
'                                                                                         '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    Dim rstHFlexgrid As New ADODB.Recordset
    Dim Contador As Integer
    Dim Colunas As Integer
    Dim matriz_tamanho() As String
    Dim matriz_caption() As String
    Dim contador_colunas As Integer
    Dim I As Integer
    Dim conexao_grid As New DLLConexao_Sistema.conexao
    
 '   On Error GoTo Erro_HFlexgrid
        
    'Indicando o banco à conectar-se
    conexao_grid.Initial_Catalog = Banco
    
    conexao_grid.Abrir_conexao (Aplicacao)
       
    rstHFlexgrid.CursorLocation = adUseClient
    rstHFlexgrid.Open strSql, conexao_grid.CNConexao, adOpenKeyset, adLockOptimistic, adCmdText
    
    'INSERIDO AQUI PARA SOLUCIONAR UM PROBLEMA DE ÚLTIMO ITEM EXCLUIDO E CONTINUA A APARECER NO GRID 16/06/2004
    If rstHFlexgrid.BOF = True And rstHFlexgrid.EOF = True Then HflexGrid.Clear: Exit Function
    
    'Marreta para acertar bug do Hflex grid
    'Não tirar o IF abaixo dessa posição. Se tirar vão ocorrer erros no FlexGrid. (Giordano e Marcos).
    
    If rstHFlexgrid.RecordCount = 0 Then
       Set rstHFlexgrid = Nothing
       Exit Function
    End If
    
    Set HflexGrid.DataSource = rstHFlexgrid
   
    HflexGrid.Clear
         
    'Montando a matriz
    matriz_tamanho = Split(tamanho_colunas, ",")
    matriz_caption = Split(caption_campos, ",")
    
    Contador = 0
    Colunas = 1
    
    HflexGrid.Cols = rstHFlexgrid.Fields.Count + 1
    
    HflexGrid.ColWidth(0) = 480
    
    'Montando as características do cabeçalho do MSHFlexgrid - Caption e Tamanho
    Do While matriz_caption(Contador) <> Empty
        DoEvents
        HflexGrid.Row = 0
        HflexGrid.Col = Colunas
        HflexGrid.FixedAlignment(Colunas) = 2
        HflexGrid.Font.Name = "Tahoma"
        HflexGrid.Text = matriz_caption(Contador)
        HflexGrid.ColWidth(Colunas) = Val(matriz_tamanho(Contador))
        Contador = Contador + 1
        Colunas = Colunas + 1
        If Contador > UBound(matriz_caption) Then Exit Do
    Loop
    
    HflexGrid.Rows = rstHFlexgrid.RecordCount + 1
    
    'Essa Mudança não pode ser retirada, senão o produto nao funcionara
    If rstHFlexgrid.EOF <> True And rstHFlexgrid.BOF <> True Then
       rstHFlexgrid.MoveFirst
    End If
    
    contador_colunas = 1
    Linhas = 1
    I = 0
    
    Do While Linhas <= rstHFlexgrid.RecordCount
       DoEvents
       HflexGrid.Row = Linhas
       HflexGrid.Col = 0
       HflexGrid.CellBackColor = &H80FFFF
       HflexGrid.CellFontBold = False
       HflexGrid.CellFontSize = 7
       HflexGrid.Text = Linhas
       Do While contador_colunas <= rstHFlexgrid.Fields.Count
          HflexGrid.Col = contador_colunas
          'Essse if abaixo foi incluido para identificar os campos "Booleanos" e
          'mover "Sim" ou "Não" para o Grid ao inves de "True" ou "False".(Giordano Vilela)
          If rstHFlexgrid.Fields.Item(I).Type = adBoolean Then
             If rstHFlexgrid.Fields(I).Value = False Then
                HflexGrid.Text = "Não"
             Else
                HflexGrid.Text = "Sim"
             End If
          Else
              'if inserido para atribyir ao campo NULL espaço para não dar problema no HFlexgrid
              If IsNull(rstHFlexgrid.Fields(I).Value) Then
                 HflexGrid.Text = " "
              Else
                If rstHFlexgrid.Fields.Item(I).Type = adCurrency Then
                   HflexGrid.Text = Format(rstHFlexgrid.Fields(I).Value, "#,###0.000")
                Else
                   HflexGrid.Text = rstHFlexgrid.Fields(I).Value
                End If
              End If
          End If
          contador_colunas = contador_colunas + 1
          I = I + 1
       Loop
       rstHFlexgrid.MoveNext
       I = 0
       contador_colunas = 1
       Linhas = Linhas + 1
    Loop
      
    Set rstHFlexgrid = Nothing
    
    HflexGrid.Row = 1
    HflexGrid.Col = 0
    
    If controle_focu = "" Then controle_focu = "S"
    If controle_focu = "S" Then HflexGrid.SetFocus
    
  '  Exit Function
    
'Erro_HFlexgrid:

 '   Call Erro.Erro(Form, "Otica", "Movimenta_HFlex_Grid")
    
End Function
Function Movimenta_HFlex_GridII(strSql As String, HflexGrid As Object, tamanho_colunas As String, caption_campos As String, Banco As String, Aplicacao As String, Form As Object)
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Only Tech                                                                                  '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Sistema................: PDV Only Tech                                                     '
' Equipe Responsável.....: Giordano Vilela,Marcos Baião                                   '
' Data da criação........: 05/08/2003                                                     '
' Data última manutenção.:                                                                '
' Observação.............:                                                                '
'                                                                                         '
'                                                                                         '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    Dim rstHFlexgrid As New ADODB.Recordset
    Dim Contador As Integer
    Dim Colunas As Integer
    Dim matriz_tamanho() As String
    Dim matriz_caption() As String
    Dim contador_colunas As Integer
    Dim I As Integer
    Dim conexao_grid As New DLLConexao_Sistema.conexao
    Dim C As Integer

  '  On Error GoTo Erro_HFlexgrid

    'Indicando o banco à conectar-se
    conexao_grid.Initial_Catalog = Banco

    conexao_grid.Abrir_conexao (Aplicacao)

    rstHFlexgrid.CursorLocation = adUseClient
    rstHFlexgrid.Open strSql, conexao_grid.CNConexao, adOpenKeyset, adLockOptimistic, adCmdText

    Set HflexGrid.DataSource = rstHFlexgrid
    HflexGrid.Clear

    'Montando a matriz
    matriz_tamanho = Split(tamanho_colunas, ",")
    matriz_caption = Split(caption_campos, ",")

    Contador = 0
    Colunas = 1

    HflexGrid.Cols = rstHFlexgrid.Fields.Count + 1

    HflexGrid.ColWidth(0) = 300

    'Montando as características do cabeçalho do MSHFlexgrid - Caption e Tamanho
    Do While matriz_caption(Contador) <> Empty
        HflexGrid.Row = 0
        HflexGrid.Col = Colunas
        HflexGrid.FixedAlignment(Colunas) = 2
        HflexGrid.Font.Name = "Tahoma"
        HflexGrid.Text = matriz_caption(Contador)
        HflexGrid.ColWidth(Colunas) = Val(matriz_tamanho(Contador))
        Contador = Contador + 1
        Colunas = Colunas + 1
        If Contador > UBound(matriz_caption) Then Exit Do
    Loop

    HflexGrid.Rows = rstHFlexgrid.RecordCount + 1

    Linhas = 1
    
    I = 0
    C = 0
    Dim Teste As String
    
    'Essa Mudança não pode ser retirada, senão o produto nao funcionara
    If rstHFlexgrid.EOF <> True And rstHFlexgrid.BOF <> True Then
       rstHFlexgrid.MoveFirst
    End If
    
    Dim strLinha As String
    
    Do While Linhas <= rstHFlexgrid.RecordCount
       DoEvents
       Do While C < rstHFlexgrid.Fields.Count
          'strRecebe = rstHFlexgrid.Fields.Item(C).Value
          'vetNome(C) = strRecebe & "ø"
          If IsNull(rstHFlexgrid.Fields.Item(C).Value) Then
              rstHFlexgrid.Fields.Item(C).Value = " "
          End If
          If C = rstHFlexgrid.Fields.Count - 1 Then
            strLinha = strLinha & rstHFlexgrid.Fields.Item(C).Value
          Else
            strLinha = strLinha & rstHFlexgrid.Fields.Item(C).Value & " + Chr$(9) + "
          End If
          C = C + 1
       Loop
       Teste = Linhas
       strLinha = strLinha
       HflexGrid.AddItem Teste + Chr(9) + strLinha, Teste
       HflexGrid.Refresh
       HflexGrid.Row = Linhas
       HflexGrid.Col = 0
       HflexGrid.CellBackColor = &H80FFFF
       HflexGrid.CellFontBold = False
       HflexGrid.CellFontSize = 7
       HflexGrid.Text = Linhas
       strLinha = Empty
       Linhas = Linhas + 1
       C = 0
       If rstHFlexgrid.EOF <> True Then rstHFlexgrid.MoveNext
    Loop

    Set rstHFlexgrid = Nothing

   ' Exit Function

'Erro_HFlexgrid:

 '   Call Erro.Erro(Form, "Otica", "Movimenta_HFlex_GridII")
    
End Function

Public Function Movimenta_DataCombo(Nome_Campo_Codigo As String, Nome_Campo_Descricao As String, DataCombo As Object, String_Sql As String, Banco As String, Aplicacao As String, Form As Object) As String

    Dim rstCombo As New ADODB.Recordset
    Dim vDatacombo As Object
    Dim conexao_combo As New DLLConexao_Sistema.conexao
    
  '  On Error GoTo Erro
    
    'Indicando o banco à conectar-se
    conexao_combo.Initial_Catalog = Banco
    
    conexao_combo.Abrir_conexao (Aplicacao)
    
    DoEvents
    rstCombo.CursorLocation = adUseClient
    String_Sql = String_Sql & " ORDER BY " & Nome_Campo_Descricao
    rstCombo.Open String_Sql, conexao_combo.CNConexao, adOpenStatic, adLockReadOnly
                
    Set vDatacombo = DataCombo
    Set vDatacombo.DataSource = rstCombo
    
    Set vDatacombo.RowSource = rstCombo
        vDatacombo.ListField = Nome_Campo_Descricao
        'BoundColumn -> sendo usado para retornar no TextBox o valor pedido, neste caso
        '               Nome_Campo_Codigo
        vDatacombo.BoundColumn = Nome_Campo_Codigo
    
    Set rstCombo.ActiveConnection = Nothing
    
    conexao_combo.Fechar_conexao
    
   ' Exit Function
    
'Erro:

 '   Call Erro.Erro(Form, Aplicacao, "Movimenta Combo")
    
End Function

Public Function Verifica_Numero(Nome_Campo As String, Nome_Tabela As String, Nome_textbox As Object, Aplicacao As String, Form As Object) As Boolean

    Dim conexao_verifica As New DLLConexao_Sistema.conexao
    Dim rstVerificacao As New ADODB.Recordset
    
    If Nome_textbox.Text = Empty Then
       Exit Function
    End If
    
    Dim strSql As String
    
    On Error GoTo erro
    
    strSql = Empty
    strSql = "SELECT " & Nome_Campo & " " & _
             "FROM " & Nome_Tabela & " " & _
             "WHERE " & Nome_Campo & " = '" & Nome_textbox.Text & "' "

    conexao_verifica.Abrir_conexao (Aplicacao)
    
    rstVerificacao.CursorLocation = adUseClient
    rstVerificacao.Open strSql, conexao_verifica.CNConexao, adOpenStatic, adLockReadOnly
    If rstVerificacao.EOF <> False And rstVerificacao.BOF <> True Then
        Verifica_Numero = False
    Else
      If Val(rstVerificacao(Nome_Campo)) = Val(Nome_textbox.Text) Then
          MsgBox "Registro já existente.", vbCritical, "Only Tech"
          Nome_textbox.Text = Empty
          Nome_textbox.SetFocus
          Verifica_Numero = True
      Else
          Verifica_Numero = False
      End If
      
    End If
    Exit Function
    
erro:

    If Err.Number = 3021 Then
        Verifica_Numero = False
        Exit Function
    Else
        Call erro.erro(Form, Aplicacao, "Verifica Registro")
    End If
    Nome_textbox.SetFocus
    
End Function

Public Function Select_geral(String_Sql As String, Banco As String, recordset_aplicacao As ADODB.Recordset, Aplicacao As String, Form As Object)

    Dim conexao_select As New DLLConexao_Sistema.conexao
    
  '  On Error GoTo Erro
    
    'Trecho inserido aqui porque é necessario que esta classe tenha a possibilidade de se conectar
    'a mais de um banco no SQL Server
    conexao_select.Initial_Catalog = Banco
    
    'Estabelecendo conexão com o banco
    conexao_select.Abrir_conexao (Aplicacao)
    
    DoEvents
    recordset_aplicacao.CursorLocation = adUseClient
    
    recordset_aplicacao.Open String_Sql, conexao_select.CNConexao, adOpenStatic, adLockBatchOptimistic
     
    'Desconectando Recordset
    recordset_aplicacao.ActiveConnection = Nothing
    
    'Fecha a conexão com o banco
    conexao_select.Fechar_conexao

   ' Exit Function

'Erro:

    'Fecha a conexão com o banco
 '   conexao_select.Fechar_conexao

  '  Call Erro.Erro(Form, Aplicacao, "Movimentações")
    
End Function

Function Monta_HFlex_Grid(HflexGrid As Object, tamanho_colunas As String, caption_campos As String, Quant_Campos As Integer, Aplicacao As String, Form As Object)

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Only Tech                                                                                  '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Sistema................: Otica/PDV Only Tech                                               '
' Equipe Responsável.....: Giordano Vilela,Marcos Baião                                   '
' Data da criação........: 18/08/2003                                                     '
' Data última manutenção.:                                                                '
' Observação.............:                                                                '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    Dim Contador As Integer
    Dim Colunas As Integer
    Dim matriz_tamanho() As String
    Dim matriz_caption() As String
    Dim contador_colunas As Integer
    Dim I As Integer
        
   ' On Error GoTo Erro_grid
    
    HflexGrid.Clear
 
    'Montando a matriz
    matriz_tamanho = Split(tamanho_colunas, ",")
    matriz_caption = Split(caption_campos, ",")
    
    Contador = 0
    Colunas = 1
    
    HflexGrid.Cols = Quant_Campos + 1
    HflexGrid.ColWidth(0) = 200
    
    'Montando as características do cabeçalho do MSHFlexgrid - Caption e Tamanho
    Do While matriz_caption(Contador) <> Empty
        DoEvents
        HflexGrid.Row = 0
        HflexGrid.Col = Colunas
        HflexGrid.Font.Name = "Tahoma"
        HflexGrid.Text = matriz_caption(Contador)
        HflexGrid.ColWidth(Colunas) = Val(matriz_tamanho(Contador))
        Contador = Contador + 1
        Colunas = Colunas + 1
        If Contador > UBound(matriz_caption) Then Exit Do
    Loop
        
    'Exit Function
    
'Erro_grid:

 '   Call Erro.Erro(Form, "Otica", "Monta_HFlex_Grid")
        
End Function
Public Function Acessibilidade_inicio(Caption_form As String, botao_consultar As Object, botao_refresh As Object, SSTab_form As Object, Variavel_Incluir As Boolean, Variavel_Alterar As Boolean, Variavel_Excluir As Boolean, Variavel_Consultar As Boolean, Codigo_Usuario As Long, barra_ferramentas As Object, Form As Object, Aplicacao As String, Banco As String)

        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        ' Only Tech                                                                                  '
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        ' Sistema................: Retaguarda                                                     '
        ' Equipe Responsável.....: Giordano Vilela,Marcos Baião                                   '
        ' Data da criação........: 10/07/2004                                                     '
        ' Data última manutenção.:                                                                '
        ' Observação.............: Função desenvolvida para trabalhar em conjunto com o módulo    '
        '                          movimentacoes.acessibilidade_inicio                            '
        '                          da DLLSystem_manager.Acessibilidade e será usada no Activate do'
        '                          form, será o primeiro teste de acessibilidade do form          '
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
        'Verificando acessibilidade
        'Tenho a descricao(caption do form) do programa e indico pra conseguir o ID
        acesso.Consulta_ID_programa Caption_form, Aplicacao, Banco, strID_Acessibilidade
        
        'Indicar o usuário a verificar sua acessibilidade
        acesso.Codigo_Usuario = Codigo_Usuario
    
        Set rstAplicacao = acesso.Verifica_acessibilidade(strID_Acessibilidade, Aplicacao, Banco)
        
        'Abastecendo a OCX do usuário para ser usada pelo restante da execução em Run Time
        Variavel_Alterar = rstAplicacao!DFAlteracao_TBAcessibilidade
        Variavel_Consultar = rstAplicacao!DFConsulta_TBAcessibilidade
        Variavel_Excluir = rstAplicacao!DFExclusao_TBAcessibilidade
        Variavel_Incluir = rstAplicacao!DFInclusao_TBAcessibilidade
         
        If Not rstAplicacao.EOF And Not rstAplicacao.BOF Then
            SSTab_form = 1
            'Habilita somente as operações que o usuário tem permissão
            'Novo
            barra_ferramentas.Buttons.Item(1).Enabled = Variavel_Incluir
            'Gravar
            barra_ferramentas.Buttons.Item(2).Enabled = False
            'Cancelar
            barra_ferramentas.Buttons.Item(3).Enabled = False
            'Excluir
            barra_ferramentas.Buttons.Item(4).Enabled = False
            'Imprimir
            barra_ferramentas.Buttons.Item(5).Enabled = Variavel_Consultar
            
            If Variavel_Consultar = False Then
               botao_consultar.Enabled = False
               botao_refresh.Enabled = False
            End If
         Else
            botao_consultar.Enabled = False
            botao_refresh.Enabled = False
            MsgBox "Usuário não tem privilégio para esta operação Verifique com o administrador do sistema!", vbInformation, "Only Tech"
            log.Descricao = "Operação cancelada por falta de privilégio do usuário - " & Caption_form & " "
            'Gravando o log
            log.Gravar_log "Otica", Form
            Set rstAplicacao = Nothing
            'Finaliza o executável
           ' End
         End If
         Set rstAplicacao = Nothing
    
End Function
Public Function Acessibilidade_inicio_relatorios(Caption_form As String, OCXUsuario_form As Object, Form As Object, Aplicacao As String, Banco As String)

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Only Tech                                                                                  '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Sistema................: Otica/Only Tech                                                   '
' Equipe Responsável.....: Giordano Vilela,Marcos Baião                                   '
' Data da criação........: 30/08/2003                                                     '
' Data última manutenção.:                                                                '
' Observação.............: Função desenvolvida para trabalhar em conjunto com o método da '
'                          da DLLSystem_manager.Acessibilidade e será usada no Activate do'
'                          form, será o primeiro teste de acessibilidade do form          '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  '
   '     On Error GoTo Erro_acessibilidade
    
        'Verificando acessibilidade
        'Tenho a descricao(caption do form) do programa e indico pra conseguir o ID
        acesso.Consulta_ID_programa Caption_form, Aplicacao, Banco, strID_Acessibilidade
        'Indicar o usuário a verificar sua acessibilidade
        acesso.Codigo_Usuario = OCXUsuario_form.Codigo
    
        Set rstAplicacao = acesso.Verifica_acessibilidade(strID_Acessibilidade, Aplicacao, Banco)
        
        'Abastecendo a OCX do usuário para ser usada pelo restante da execução em Run Time
        OCXUsuario_form.PrivilégioAcessar = rstAplicacao!DFAcesso_TBAcessibilidade
        OCXUsuario_form.PrivilégioAlterar = rstAplicacao!DFAlteracao_TBAcessibilidade
        OCXUsuario_form.PrivilégioConsultar = rstAplicacao!DFConsulta_TBAcessibilidade
        OCXUsuario_form.PrivilégioExcluir = rstAplicacao!DFExclusao_TBAcessibilidade
        OCXUsuario_form.PrivilégioIncluir = rstAplicacao!DFInclusao_TBAcessibilidade
        
        If rstAplicacao.EOF And rstAplicacao.BOF And OCXUsuario_form.PrivilégioAcessar = True Then
            MsgBox "Usuário não tem privilégio para esta operação Verifique com o administrador do sistema!", vbInformation, "Only Tech"
            log.Descricao = "Operação cancelada por falta de privilégio do usuário - " & Caption_form & " "
            'Gravando o log
            log.Gravar_log "Otica", Form
            Set rstAplicacao = Nothing
            'Finaliza o executável
         '   End
         End If
         Set rstAplicacao = Nothing
    
    Exit Function
    
'Erro_acessibilidade:

    'Call Erro.Erro(Form, "Otica", "Acessibilidade_inicio_relatorios")
      
End Function

Public Function Refresh_Grid(FlexGrid_Form As Object, Campos_Query As String, Nome_Tabela As String, FlexGrid_Linha_Atualizar As Integer, FlexGrid_Colunas_Atualizar As String, Indice_Fields As String, Contador As Integer, Aplicacao As String, Banco As String, Form As Object, Optional Clausula_WHERE As String, Optional Clausula_INNER_JOIN As String)
    
    Dim Conexao_Refresh As New DLLConexao_Sistema.conexao
    Dim rstRefresh As New ADODB.Recordset
    Dim Vetor_Colunas() As String
    Dim Vetor_Indice_Fields() As String
    Dim I As Integer
    Dim strSql As String
    
    Vetor_Colunas = Split(FlexGrid_Colunas_Atualizar, ",")
    Vetor_Indice_Fields = Split(Indice_Fields, ",")
              
    Conexao_Refresh.Initial_Catalog = Banco
    Conexao_Refresh.Abrir_conexao (Aplicacao)
    DoEvents
                 
    strSql = ""
      
    strSql = "SELECT " & Campos_Query & " " & _
             "FROM " & Nome_Tabela & " "
             
    If Clausula_INNER_JOIN <> Empty Then
       strSql = strSql & " INNER JOIN " & Clausula_INNER_JOIN & " "
    End If
    
    If Clausula_WHERE <> Empty Then
       strSql = strSql & " WHERE " & Clausula_WHERE & " "
    End If
        
    rstRefresh.CursorLocation = adUseClient
    rstRefresh.Open strSql, Conexao_Refresh.CNConexao, adOpenStatic, adLockBatchOptimistic
    
    If rstRefresh.RecordCount <> 0 Then
       FlexGrid_Form.Row = FlexGrid_Linha_Atualizar
       For I = 0 To Contador - 1
          FlexGrid_Form.Col = Vetor_Colunas(I)
          If rstRefresh.Fields.Item(Val(Vetor_Indice_Fields(I))).Type = adCurrency Then
             FlexGrid_Form.Text = Format(rstRefresh.Fields.Item(Val(Vetor_Indice_Fields(I))).Value, "#,###0.00")
          Else
             FlexGrid_Form.Text = rstRefresh.Fields.Item(Val(Vetor_Indice_Fields(I))).Value
          End If
       Next I
    End If
              
    FlexGrid_Form.Row = FlexGrid_Linha_Atualizar
    FlexGrid_Form.Col = 0
    FlexGrid_Form.SetFocus
    
    FlexGrid_Form.Refresh
        
    rstRefresh.ActiveConnection = Nothing
    Conexao_Refresh.Fechar_conexao
    
    Exit Function
    
End Function

Public Function Verifica_DataCombo(Valor_Data_Combo As String)
        
    If Valor_Data_Combo = Empty Then
       SendKeys "{F4}"
    Else
       SendKeys "{TAB}"
    End If

End Function

Public Function Digito_Nosso_Numero(Codigo_Banco As Long, Agencia As String, CC As String, Carteira As String, Nosso_Numero As String) As String

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Only Tech                                                                                  '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Sistema................: Retaguarda Only Tech                                              '
' Equipe Responsável.....: Giordano Vilela,Marcos Baião,Alex,Sérgio,Rafael                '
' Data da criação........: 24/02/2003                                                     '
' Data última manutenção.:                                                                '
' Observação.............: Função criada para geração automática do gig. do nosso número, '
'                          variando, de banco para banco                                  '
'                                                                                         '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    Dim strNumero_Concatenado As String
    Dim intTamanho_string_concatenado As Integer
    Dim intTamanho_string_resultado_parcial As Integer
    Dim intContador_string As Integer
    Dim dblResultado_parcial As Double
    Dim dblResultado_final As Double
    Dim dblCalculo_Digito  As Double
    Dim dblDigito  As Double
    
    '--------------------------------------------------------------------------------------------------
    'Itaú
    'Regra:
    '1 - Concatenar Agência,CC,Carteira e Nosso Número;
    '2 - Multiplicar os digitos um a um alternados por 1 e 2 da esquerda para a direita;
    '3 - Somar os produtos acima de 9 entre eles ex: 10 ---> 1 + 0;
    '4 - Dividir o resultado por 10;
    '5 - Subtrair Divisor por 10.
    '--------------------------------------------------------------------------------------------------
    
    If Codigo_Banco = 341 Then 'Itau
    
       strNumero_Concatenado = Trim(Agencia) & Trim(CC) & Trim(Carteira) & Trim(Nosso_Numero)
       intTamanho_string_concatenado = Len(strNumero_Concatenado)
       
       intContador_string = 1
       
       Do While intContador_string <= intTamanho_string_concatenado
          dblResultado_parcial = CDbl(Mid(strNumero_Concatenado, intContador_string, 1))
          'Multiplica os impares por 1
          If intContador_string Mod 2 <> 0 Then
             dblResultado_parcial = dblResultado_parcial * 1
          Else
             dblResultado_parcial = dblResultado_parcial * 2
          End If
          
          intTamanho_string_resultado_parcial = Len(CStr(dblResultado_parcial))
          'Soma os digitos dos produtos acima de 2
          If intTamanho_string_resultado_parcial > 1 Then
             dblResultado_parcial = CDbl(Mid(dblResultado_parcial, 1, 1)) + CDbl(Mid(dblResultado_parcial, 2, 1))
          End If
          
          dblResultado_final = dblResultado_final + dblResultado_parcial
                    
          intContador_string = intContador_string + 1
       Loop
       
       dblCalculo_Digito = dblResultado_final Mod 10
    
       dblDigito = 10 - dblCalculo_Digito
       
       If dblDigito = 10 Then dblDigito = 0
       
    End If
    
    Digito_Nosso_Numero = dblDigito
    
End Function
Public Function Acessibilidade_Menu(Caption_Menu As String, Aplicacao As String, Banco As String, Usuario As Long) As Boolean

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Only Tech                                                                                  '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Sistema................: Only Tech                                                         '
' Equipe Responsável.....: Marcos Baião                                                   '
' Data da criação........: 30/08/2003                                                     '
' Data última manutenção.:                                                                '
' Observação.............: Função desenvolvida para trabalhar em conjunto com o método da '
'                          da DLLSystem_manager.Acessibilidade e será usada no Admin      '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    'Verificando acessibilidade
    'Tenho a descricao(caption do form) do menu e indico pra conseguir o ID
    acesso.Consulta_ID_menu Caption_Menu, Aplicacao, Banco, strID_Acessibilidade
    'Indicar o usuário a verificar sua acessibilidade
    acesso.Codigo_Usuario = Usuario

    Set rstAplicacao = acesso.Verifica_acessibilidade_menu(strID_Acessibilidade, Aplicacao, Banco)
    
    If strID_Acessibilidade = 0 Then Acessibilidade_Menu = False: Exit Function
    
    If Not rstAplicacao.EOF And Not rstAplicacao.BOF Then
       If rstAplicacao!DFAcesso_TBAcessibilidade_Menu = True Then
          Acessibilidade_Menu = True
       Else
          Acessibilidade_Menu = False
       End If
    Else
       Acessibilidade_Menu = False
    End If
    
    Set rstAplicacao = Nothing
    
End Function

Public Function Acessibilidade_Item_Menu(Caption_Menu As String, Aplicacao As String, Banco As String, Usuario As Long, Item_Menu_Componente As Object)

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Only Tech                                                                                  '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Sistema................: Only Tech                                                         '
' Equipe Responsável.....: Marcos Baião                                                   '
' Data da criação........: 05/07/2004                                                     '
' Data última manutenção.:                                                                '
' Observação.............: Função desenvolvida para trabalhar em conjunto com o método da '
'                          da DLLSystem_manager.Acessibilidade e será usada no MDI        '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    'Verificando acessibilidade
    'Tenho a descricao(caption do form) do menu e indico pra conseguir o ID
    acesso.Consulta_ID_programa Caption_Menu, Aplicacao, Banco, strID_Acessibilidade
    'Indicar o usuário a verificar sua acessibilidade
    acesso.Codigo_Usuario = Usuario
    
    If strID_Acessibilidade = 0 Then Item_Menu_Componente.Enabled = False: Exit Function
    
    Set rstAplicacao = acesso.Verifica_acessibilidade(strID_Acessibilidade, Aplicacao, Banco)
    
    If strID_Acessibilidade = 0 Then Item_Menu_Componente.Enabled = False: Exit Function
     
    If Not rstAplicacao.EOF And Not rstAplicacao.BOF Then
       If rstAplicacao!DFAcesso_TBAcessibilidade = True Then
          Item_Menu_Componente.Enabled = True
       Else
          Item_Menu_Componente.Enabled = False
       End If
    Else
       Item_Menu_Componente.Enabled = False
    End If
    
    Set rstAplicacao = Nothing
    
End Function

Private Function Montando_Acessibilidade_Programas(Caption_form As String, Variavel_Incluir As Boolean, Variavel_Alterar As Boolean, Variavel_Excluir As Boolean, Variavel_Consultar As Boolean, Aplicacao As String, Banco As String)

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Only Tech                                                                                  '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Sistema................: Retaguarda                                                     '
' Equipe Responsável.....: Giordano Vilela,Marcos Baião                                   '
' Data da criação........: 10/07/2004                                                     '
' Data última manutenção.:                                                                '
' Observação.............: Função desenvolvida para trabalhar em conjunto com o módulo    '
'                          movimentacoes.acessibilidade_inicio                            '
'                          da DLLSystem_manager.Acessibilidade e será usada no Activate do'
'                          form, será o primeiro teste de acessibilidade do form          '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    'Verificando acessibilidade
    'Tenho a descricao(caption do form) do programa e indico pra conseguir o ID
    acesso.Consulta_ID_programa Caption_form, Aplicacao, Banco, strID_Acessibilidade
    
    'Indicar o usuário a verificar sua acessibilidade
    acesso.Codigo_Usuario = OCXUsuario_form.Codigo

    Set rstAplicacao = acesso.Verifica_acessibilidade(strID_Acessibilidade, Aplicacao, Banco)
    
    'Abastecendo a OCX do usuário para ser usada pelo restante da execução em Run Time
    Variavel_Alterar = rstAplicacao!DFAlteracao_TBAcessibilidade
    Variavel_Consultar = rstAplicacao!DFConsulta_TBAcessibilidade
    Variavel_Excluir = rstAplicacao!DFExclusao_TBAcessibilidade
    Variavel_Incluir = rstAplicacao!DFInclusao_TBAcessibilidade
        
End Function
Public Function Consulta_Contingencia_Acessibilidade(Aplicacao As String) As String

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Only Tech                                                                                   '
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Sistema................: Retaguarda                                                      '
' Equipe Responsável.....: Giordano Vilela,Marcos Baião                                    '
' Data da criação........: 20/07/2004                                                      '
' Data última manutenção.:                                                                 '
' Observação.............: Função desenvolvida para trabalhar em conjunto com o todos os   '
'                          módulos nos MDI's ela será útil, quando o programa não conseguir'                            '
'                          acessar o endereço de memória para colher as informações dos us_'
'                          uários logados no admin                                         '
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    Dim Registro As New DLLSystemManager.Registro
    Dim Nome_Usuario As String, Codigo_Usuario As String, Empresa_Usuario As String, Estacao_Usuario As String

    Empresa_Usuario = CStr(Registro.WinRegLerSequência("HKEY_LOCAL_MACHINE\SOFTWARE\" + Aplicacao + "\INF.USER", "EMPRESA"))
    Codigo_Usuario = CStr(Registro.WinRegLerSequência("HKEY_LOCAL_MACHINE\SOFTWARE\" + Aplicacao + "\INF.USER", "CODIGO"))
    Nome_Usuario = CStr(Registro.WinRegLerSequência("HKEY_LOCAL_MACHINE\SOFTWARE\" + Aplicacao + "\INF.USER", "NOME"))
    Estacao_Usuario = CStr(Registro.WinRegLerSequência("HKEY_LOCAL_MACHINE\SOFTWARE\" + Aplicacao + "\INF.USER", "ESTACAO"))
    
    '--------------------------------------------------------------------------------------------------
    'Monta a mensagem na mesma ordem da Intercomunicador
    '--------------------------------------------------------------------------------------------------
    Consulta_Contingencia_Acessibilidade = Estacao_Usuario & "¤" & _
                                           Nome_Usuario & "¤" & _
                                           "********" & "¤" & _
                                           Nome_Usuario & "¤" & _
                                           Codigo_Usuario & "¤" & _
                                           "#" & "¤" & _
                                           "#" & "¤" & _
                                           "#" & "¤" & _
                                           "#" & "¤" & _
                                           "#" & "¤" & _
                                           "#" & "¤" & _
                                           Empresa_Usuario
End Function
Public Function Grava_Contingencia_Acessibilidade(Estacao_Usuario As String, Nome_Usuario As String, Codigo_Usuario As Integer, Empresa_Usuario As String, Aplicacao As String) As String

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Only Tech                                                                                   '
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Sistema................: Retaguarda                                                      '
' Equipe Responsável.....: Giordano Vilela,Marcos Baião                                    '
' Data da criação........: 20/07/2004                                                      '
' Data última manutenção.:                                                                 '
' Observação.............: Função desenvolvida para trabalhar em conjunto com o todos os   '
'                          módulos nos MDI's ela será útil, quando o programa não conseguir'                            '
'                          acessar o endereço de memória para colher as informações dos us_'
'                          uários logados no admin                                         '
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    Dim Registro As New DLLSystemManager.Registro

    'Criando chave no registro
    Registro.WinRegCriarChave "HKEY_LOCAL_MACHINE\SOFTWARE\" + Aplicacao + "\INF.USER"
    
    'Limpando as inf. do usuário
    Call Limpa_Contingencia_Acessibilidade(Aplicacao)
    
    'Gravando as inf. do usuário
    Registro.WinRegAdicionarSequência "HKEY_LOCAL_MACHINE\SOFTWARE\" + Aplicacao + "\INF.USER", "ESTACAO", Estacao_Usuario
    Registro.WinRegAdicionarSequência "HKEY_LOCAL_MACHINE\SOFTWARE\" + Aplicacao + "\INF.USER", "EMPRESA", Empresa_Usuario
    Registro.WinRegAdicionarSequência "HKEY_LOCAL_MACHINE\SOFTWARE\" + Aplicacao + "\INF.USER", "CODIGO", CInt(Codigo_Usuario)
    Registro.WinRegAdicionarSequência "HKEY_LOCAL_MACHINE\SOFTWARE\" + Aplicacao + "\INF.USER", "NOME", Nome_Usuario
    Registro.WinRegAdicionarSequência "HKEY_LOCAL_MACHINE\SOFTWARE\" + Aplicacao + "\INF.USER", "DATA", Format(Now, "DD/MM/YYYY")
    Registro.WinRegAdicionarSequência "HKEY_LOCAL_MACHINE\SOFTWARE\" + Aplicacao + "\INF.USER", "HORA", Format(Now, "HH:MM:SS")
    
End Function
Public Function Limpa_Contingencia_Acessibilidade(Aplicacao As String) As String

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Only Tech                                                                                   '
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Sistema................: Retaguarda                                                      '
' Equipe Responsável.....: Giordano Vilela,Marcos Baião                                    '
' Data da criação........: 20/07/2004                                                      '
' Data última manutenção.:                                                                 '
' Observação.............: Função desenvolvida para trabalhar em conjunto com o todos os   '
'                          módulos nos MDI's ela será útil, quando o programa não conseguir'                            '
'                          acessar o endereço de memória para colher as informações dos us_'
'                          uários logados no admin                                         '
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    Dim Registro As New DLLSystemManager.Registro

    'Limpando as inf. do usuário
    Registro.WinRegAdicionarSequência "HKEY_LOCAL_MACHINE\SOFTWARE\" + Aplicacao + "\INF.USER", "EMPRESA", ""
    Registro.WinRegAdicionarSequência "HKEY_LOCAL_MACHINE\SOFTWARE\" + Aplicacao + "\INF.USER", "ESTACAO", ""
    Registro.WinRegAdicionarSequência "HKEY_LOCAL_MACHINE\SOFTWARE\" + Aplicacao + "\INF.USER", "CODIGO", ""
    Registro.WinRegAdicionarSequência "HKEY_LOCAL_MACHINE\SOFTWARE\" + Aplicacao + "\INF.USER", "NOME", ""
    Registro.WinRegAdicionarSequência "HKEY_LOCAL_MACHINE\SOFTWARE\" + Aplicacao + "\INF.USER", "DATA", ""
    Registro.WinRegAdicionarSequência "HKEY_LOCAL_MACHINE\SOFTWARE\" + Aplicacao + "\INF.USER", "HORA", ""
   
End Function

