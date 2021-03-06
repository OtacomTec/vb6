VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsImpressao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Criado em 28/09/2005 por Fabiano Salles
'Altera��es e manuten��es: Fabiano Salles e Leandro Nolosco

Option Explicit

Dim Orientacao_Papel As Byte                'Retrato, Paisagem
Dim Tamanho_Papel As Integer                'N�mero de registros impressos por folha
Dim Tamanho_PapelVB As String               'Tamanho do papel para o VB

Dim Destino_Relatorio As Boolean            'Tela ou Impressora
Dim Conteudo_Relatorio As New Collection    'Cole��o contendo o conteudo do relat�rio

Dim Fonte_Relatorio As String               'Fonte do relat�rio
Dim Tamanho_Fonte_Rel As Integer            'Tamanho da Fonte do CORPO
Dim Tamanho_Fonte_Cab As Integer            'Tamanho da fonte do CABE�ALHO do REALT�RIO

Dim Titulo_Relatorio As String
Dim Titulo_Empresa_Relatorio As String
Dim Titulo_Cliente_Relatorio As String
Dim Data_Criacao_Relatorio As String
Dim Sub_Titulo_Relatorio As String
Dim Linhas_RCabecalho As Integer            'Esta vari�vel guarda quantas linhas ser�o repetidas
                                            'para a forma��o do cabe�alho das COLUNAS
Dim dbl_N_Pag As Double                     'N�mero de p�gians
Dim int_Conta_Pagina As Integer             'Contador

Dim int_Controle_Graf As Integer            'Controla a p�gina onde ser� impresso o gr�fico

Private Type tDadosGrafico
    bol_Grafico As Boolean                  'Determina se ter� gr�fico
    int_Indice As Integer                   'Indica quantos gr�ficos ter� o relat�rio
    obj_Grafico As PictureBox               'Propriedade que armazenar� o gr�fico
    int_Alinha As Integer                   'Guradar� o alinhamento
    int_GrafY As Integer                    'Posi��o X do gr�fico
    int_GrafX As Integer                    'Posi��o Y do gr�fico
    dbl_Tamanho As Double                   'Tamanho do Gr�fico em % (1 = 100%)
End Type

Dim tDadosGrafico As tDadosGrafico

Public Enum Alinhamento                     'Alinhamentos
    alAmbos = 1
    alCentraliza_H = 2
    alCentraliza_V = 3
    alManual = 4
End Enum

Dim strRodape_Pagina As String

'Essa propriedade ir� receber a orienta��o do papel
Public Property Let Orientacao(ByVal New_Orientacao As Byte)

    '1 = Retrato
    '2 = Paisagem
    
    Orientacao_Papel = New_Orientacao
    
End Property

Public Property Get Orientacao() As Byte

    Orientacao = Orientacao_Papel
    
End Property

'Tamanho do papel
Public Property Let Papel(ByVal New_Papel As String)

    If New_Papel = "A4" And Orientacao_Papel = 1 Then
    
        Tamanho_Papel = 75 '78
        Tamanho_PapelVB = vbPRPSA4
    
    ElseIf New_Papel = "A4" And Orientacao_Papel = 2 Then
        
        Tamanho_Papel = 50 '53
        Tamanho_PapelVB = vbPRPSA4
    
    'como nunca usou, configurar outra orienta��o e verificar espa�o de rodap�
    ElseIf New_Papel = "OFICIO" And Orientacao_Papel = 1 Then
                
        Tamanho_Papel = 66
        Tamanho_PapelVB = vbPRPSLetter
            
    ElseIf New_Papel = "CARTA" And Orientacao_Papel = 1 Then
                
        Tamanho_Papel = 49
        Tamanho_PapelVB = vbPRPSLegal
            
    End If
    
End Property

Public Property Get Atual_Papel()

    Select Case UCase(Tamanho_Papel)
    
        Case 75
            Atual_Papel = "A4"
            
        Case 66
            Atual_Papel = "OFICIO"
            
        Case 49
            Atual_Papel = "CARTA"
            
    End Select
    
End Property

'Sa�da - Destino
Public Property Let Destino(ByVal New_Destino As Boolean)

    ' 0 (False) = Tela
    '-1 (True) = Impressora
    
    Destino_Relatorio = New_Destino

End Property

'Fonte
Public Property Let Fonte(ByVal New_Fonte As FontName)

    Fonte_Relatorio = New_Fonte

End Property

'Tamanho Fonte Relat�rio
Public Property Let Tamanho_Fonte_Corpo(ByVal New_Tamanho As Integer)

    Tamanho_Fonte_Rel = New_Tamanho

End Property

'Tamanho Fonte cabe�alho
Public Property Let Tamanho_Fonte_Cabecalho(ByVal New_Tamanho As Integer)

    Tamanho_Fonte_Cab = New_Tamanho

End Property

'T�tulo do Relat�rio
Public Property Let Titulo(ByVal New_Titulo As String)

    Titulo_Relatorio = New_Titulo

End Property

'T�tulo do Relat�rio
Public Property Let Titulo_Empresa(ByVal New_Titulo_Empresa As String)

    Titulo_Empresa_Relatorio = New_Titulo_Empresa

End Property

'Sub-T�tulo do Relat�rio
Public Property Let Sub_Titulo(ByVal New_Sub_Titulo As String)

    Sub_Titulo_Relatorio = New_Sub_Titulo

End Property

'Cliente do Relat�rio
Public Property Let Titulo_Cliente(ByVal New_Titulo_Cliente As String)

    Titulo_Cliente_Relatorio = New_Titulo_Cliente

End Property

'Data de Cria��o do Relat�rio
Public Property Let Data_Criacao(ByVal New_Data_Criacao As String)

    Data_Criacao_Relatorio = New_Data_Criacao

End Property



'Esta propriedade guardara a quantidade de linhas que ser�o
'repetidas nos cabe�alhos das colunas nas quebras de p�gina
Public Property Let Linhas_Cabecalho(ByVal New_Linhas As Integer)

    Linhas_RCabecalho = New_Linhas

End Property

'Informa��es que ser�o plotadas no relat�rio
Public Property Let Conteudo(ByVal New_Conteudo As Collection)
    
    Set Conteudo_Relatorio = New_Conteudo
    
End Property

'Rotina que executa a impressao
Public Sub Executa_Impressao(ByVal int_Pag As Integer)
        
    If Destino_Relatorio = 0 Then
        Imprimir_Tela (int_Pag)
    Else
        Imprimir_Impressora
    End If

End Sub

Public Property Get Numero_Paginas()

    Numero_Paginas = dbl_N_Pag
    
End Property

Public Sub setDados_Grafico(ByRef New_Grafico As PictureBox, ByVal New_Alinhameno As Alinhamento, Optional ByVal New_X As Integer, Optional ByVal New_Y As Integer, Optional ByVal New_Tamanho As Double = 1, Optional ByVal New_Indice As Integer = 0)

    With tDadosGrafico
    
        .bol_Grafico = True
        .int_Alinha = New_Alinhameno
        .int_Indice = New_Indice
        .dbl_Tamanho = New_Tamanho
        
        'If New_Alinhameno = alNONE Then
            .int_GrafY = New_X
            .int_GrafX = New_Y
        'End If
        
        Set .obj_Grafico = New_Grafico
        
    End With
    
End Sub

Private Sub Imprimir_Impressora()

    Dim int_Tamanho_Folha As Integer
    Dim int_X As Integer
    Dim int_Y As Integer
    
    'Vari�veis de controle de  para quebra de p�gina
    Dim int_Controle As Integer
    Dim bol_Troca_Pag As Boolean
    Dim str_Linha_Ant As String
    Dim str_Negrito_Ant As String
    
    Dim lng_Cont_Linhas_Cab_Grupo As Long               'n�mero de linhas do cabe�alho de grupos
    Dim lng_Cont_Detalhe As Long                        'n�mero de registro dos detalhes do grupo
    Dim lng_IGrupo_Atual As Long                        '�ndice do grupo atual
    
    Dim lng_Area As Long                                'em caso de gr�fico, serve para verificar se
                                                        'o gr�fico cabe na �rea apontada na chamada da fun��o
    
    
    On Error GoTo TrataErro
    
    Printer.Orientation = Orientacao_Papel
    Printer.PaperSize = Tamanho_PapelVB
    Printer.FontName = Fonte_Relatorio
    int_Tamanho_Folha = Printer.ScaleHeight - 1200 '- 1440
    
    int_Conta_Pagina = 1
    
    Imprimir_Cabecalho_Impressora
    Imprimir_Cabecalho_Dados
    
    Printer.Print Spc(2); ""
    
    '    For int_X = Linhas_RCabecalho To Conteudo_Relatorio.Count - 1
    '
    '        If Printer.CurrentY >= int_Tamanho_Folha Then
    '
    '            Printer.NewPage
    '            Imprimir_Cabecalho_Impressora
    '            'Posiciona a primeira linha depois do cabe�alho dos registros...
    '            'Printer.CurrentY = 1490
    '            Printer.CurrentY = 1390 + (Linhas_RCabecalho * 110)
    '
    '        End If
    '
    '        Printer.FontBold = Not (Conteudo_Relatorio(CStr(int_X)).Negrito = "N")
    '
    '        'se a primeira informacao da p�gina for uma linha n�o imprime
    '        If Printer.CurrentY = 1490 And Right(Conteudo_Relatorio(CStr(int_X)).Linha, 1) = "_" Then
    '        Else
    '             Printer.Print Conteudo_Relatorio(CStr(int_X)).Linha
    '        End If
    '
    '    Next
    
    '********************************************************************************************************
    
    For int_X = 0 To Conteudo_Relatorio.Count - 1
    
        'incrementa linha
        int_Controle = int_Controle + 1
            
        'se atingimos o tamanho m�ximo da folha ou temmos que
        'quebrar a p�gina para n�o quebra o grupo...
        If int_Controle > Tamanho_Papel Or bol_Troca_Pag Then
            
            'n�mero de p�ginas
            int_Conta_Pagina = int_Conta_Pagina + 1
            int_Controle = Linhas_RCabecalho + 1
            
            'flag para trocar de p�gina
            bol_Troca_Pag = False
            
            Printer.NewPage
            Imprimir_Cabecalho_Impressora
            
            'Posiciona a primeira linha depois do cabe�alho dos registros...
            'Printer.CurrentY = IIf(Linhas_RCabecalho <= 1, 1490, 1390 + (Linhas_RCabecalho * 120))
            For int_Y = 0 To Linhas_RCabecalho
                Printer.FontBold = Not (Conteudo_Relatorio(CStr(int_X)).Negrito = "N")
                Printer.Print Spc(2); Conteudo_Relatorio(CStr(int_Y)).Linha
            Next
            
        End If
            
        'rotina que plota a linha anterior caso tenha sido necess�rio for�ar a quebra da p�gina
        'If Not Vazio(str_Linha_Ant) Then
        If str_Linha_Ant <> Empty Then
            Printer.FontBold = Not (str_Negrito_Ant = "N")
            Printer.Print Spc(2); str_Linha_Ant,
            str_Linha_Ant = ""
            str_Negrito_Ant = ""
        End If
            
        'controla se a linha ser� em negrito
        Printer.FontBold = Not (Conteudo_Relatorio(CStr(int_X)).Negrito = "N")
            
        'se for uma linha na primeira posi��o do relat�rio, n�o imprime...
        If Printer.CurrentY = 1490 And Right(Conteudo_Relatorio(CStr(int_X)).Linha, 1) = "_" Then
        Else
            'se tem grupo...
            If Conteudo_Relatorio(CStr(int_X)).IGrupo > -1 Then
                
                'capturar o �ndice e inicializando vari�veis de controle
                lng_IGrupo_Atual = Conteudo_Relatorio(CStr(int_X)).IGrupo
                lng_Cont_Detalhe = 0
                lng_Cont_Linhas_Cab_Grupo = 1
                
                'contar detalhes e linhas de cabe�alho do grupo
                For int_Y = int_X + 1 To Conteudo_Relatorio.Count - 1
                    If Conteudo_Relatorio(CStr(int_Y)).IGrupo = -1 Then
                        lng_Cont_Detalhe = lng_Cont_Detalhe + 1
                    ElseIf Conteudo_Relatorio(CStr(int_Y)).IGrupo > lng_IGrupo_Atual Then
                        Exit For
                    Else
                        lng_Cont_Linhas_Cab_Grupo = lng_Cont_Linhas_Cab_Grupo + 1
                    End If
                Next
                
                'verificando se o grupo cabe no papel...
                If lng_Cont_Detalhe + lng_Cont_Linhas_Cab_Grupo <= Tamanho_Papel Then
                    'cabe, logo verifica se o grupo cabe no espa�o restante
                    If lng_Cont_Detalhe + lng_Cont_Linhas_Cab_Grupo > Tamanho_Papel - int_Controle Then
                        bol_Troca_Pag = True
                    Else
                        bol_Troca_Pag = False
                    End If
                Else
                    'se n�o couber, verifica se cabe o cabe�alho do grupo e o primeiro detalhe
                    If lng_Cont_Linhas_Cab_Grupo + IIf(lng_Cont_Detalhe > 0, 1, 0) > Tamanho_Papel - int_Controle Then
                        bol_Troca_Pag = True
                    Else
                        bol_Troca_Pag = False
                    End If
                End If
            End If
            
            'se n�o quebra p�gina, IMPRIME!
            If Not bol_Troca_Pag Then
                Printer.Print Spc(2); Conteudo_Relatorio(CStr(int_X)).Linha
            'se quebra...
            Else
                'grava a linha para plotar na pr�xima p�gina
                str_Linha_Ant = Conteudo_Relatorio(CStr(int_X)).Linha
                str_Negrito_Ant = Conteudo_Relatorio(CStr(int_X)).Negrito
            End If
        End If
    Next
    '********************************************************************************************************
    
    'imprime o(s) Gr�fico(s)
    If tDadosGrafico.bol_Grafico Then
        Printer.CurrentX = 30
    
        For int_X = 0 To tDadosGrafico.int_Indice
        
            lng_Area = tDadosGrafico.int_GrafX + (frmVisualiza_Impressao.imgGrafico(int_X).width - 1000)
            
            Select Case ((Printer.ScaleWidth - tDadosGrafico.int_GrafX) - lng_Area)
                Case Is > 0 'CABE
                    Printer.PaintPicture frmVisualiza_Impressao.imgGrafico(int_X).Picture, IIf(tDadosGrafico.int_GrafX = 0, 30, tDadosGrafico.int_GrafX), tDadosGrafico.int_GrafY, frmVisualiza_Impressao.imgGrafico(int_X).width * tDadosGrafico.dbl_Tamanho, frmVisualiza_Impressao.imgGrafico(int_X).Height * tDadosGrafico.dbl_Tamanho
                Case Else
                    Printer.NewPage
                    
                    int_Conta_Pagina = int_Conta_Pagina + 1
                    Imprimir_Cabecalho_Impressora
                    Printer.CurrentX = 30
                    
                    'auto ajuste...
                    If frmVisualiza_Impressao.imgGrafico(int_X).width <= Printer.ScaleWidth Then
                        Printer.PaintPicture frmVisualiza_Impressao.imgGrafico(int_X).Picture, ((Printer.ScaleWidth - 30) - frmVisualiza_Impressao.imgGrafico(int_X).width) / 2, 2000, frmVisualiza_Impressao.imgGrafico(int_X).width * tDadosGrafico.dbl_Tamanho, frmVisualiza_Impressao.imgGrafico(int_X).Height * tDadosGrafico.dbl_Tamanho
                    Else
                        Printer.PaintPicture frmVisualiza_Impressao.imgGrafico(int_X).Picture, 30, 2000, frmVisualiza_Impressao.imgGrafico(int_X).width * Round(1 - (frmVisualiza_Impressao.imgGrafico(int_X).width / (Printer.ScaleWidth - 150) - 1), 2), frmVisualiza_Impressao.imgGrafico(int_X).Height * tDadosGrafico.dbl_Tamanho
                    End If
            End Select
            'Printer.PaintPicture frmVisualiza_Impressao.imgGrafico(int_X), tDadosGrafico.int_GrafX, tDadosGrafico.int_GrafY, frmVisualiza_Impressao.imgGrafico(int_X).Width * tDadosGrafico.dbl_Tamanho, frmVisualiza_Impressao.imgGrafico(int_X).Height * tDadosGrafico.dbl_Tamanho
            
        Next
       
    End If
    
    Printer.EndDoc
    
    MsgBox "Os dados foram enviados para a impressora! Verifique!", vbInformation, wNomeSistema
        
    Exit Sub

TrataErro:
    If Err.Number <> 0 Then
        MsgBox Err.Number & " - " & Err.Description, vbCritical, wNomeSistema
    End If
End Sub

Private Sub Imprimir_Cabecalho_Impressora()

    Dim int_X As Integer
    Dim int_X_Atual As Integer
    Dim int_Y_Atual As Integer
     
    On Error GoTo TrataErro
    
    Printer.FontBold = True
    Printer.CurrentY = 180
    Printer.FontName = Fonte_Relatorio
    Printer.FontSize = Tamanho_Fonte_Cab
    
    'Cabe�alho do Relat�rio
    'Printer.PaintPicture frmVisualiza_Impressao.imgLogo.Picture, 1, 1
    
    'aki
    'Retrato
    If Orientacao = 1 Then
    
        Printer.Print Titulo_Empresa_Relatorio & Space(50) & "Data de Criacao: " & Data_Criacao_Relatorio
        Printer.Print String(96, "-")
        Printer.Print Titulo_Cliente_Relatorio & Space(96 - (Len(Trim(Titulo_Cliente_Relatorio)) + Len("Data de Geracao: " & Format(Date, "dd/MM/yyyy")))) & "Data de Geracao: " & Format(Date, "dd/MM/yyyy")
        Printer.Print Titulo_Relatorio & Space(96 - (Len(Titulo_Relatorio) + (Len("Hora de Geracao: " & Format(Now, "HH:nn:ss  "))))) & "Hora de Geracao: " & Format(Now, "HH:nn:ss  ")
        Printer.Print Sub_Titulo_Relatorio & Space(92 - (Len(Sub_Titulo_Relatorio) + Len("P�ginas: " & int_Conta_Pagina & " de " & Mid(frmVisualiza_Impressao.lblPagina, InStr(frmVisualiza_Impressao.lblPagina, " de ") + 4)))) & "P�ginas: " & int_Conta_Pagina & " de " & Mid(frmVisualiza_Impressao.lblPagina, InStr(frmVisualiza_Impressao.lblPagina, " de ") + 4)
        Printer.Print String(96, "-")
        Printer.Print
    
    'Paisagem
    ElseIf Orientacao = 2 Then
    
        Printer.Print Titulo_Empresa_Relatorio & Space(68) & "Data de Criacao: " & Data_Criacao_Relatorio
        Printer.Print String(114, "-")
        Printer.Print Titulo_Cliente_Relatorio & Space(114 - (Len(Trim(Titulo_Cliente_Relatorio)) + Len("Data de Geracao: " & Format(Date, "dd/MM/yyyy")))) & "Data de Geracao: " & Format(Date, "dd/MM/yyyy")
        Printer.Print Titulo_Relatorio & Space(114 - Len(Titulo_Relatorio) + (Len("Hora de Geracao: " & Format(Now, "HH:nn:ss  ")))) & "Hora de Geracao: " & Format(Now, "HH:nn:ss  ")
        Printer.Print Sub_Titulo_Relatorio & Space(114 - (Len(Sub_Titulo_Relatorio) + Len("P�ginas: " & int_Conta_Pagina & " de " & Mid(frmVisualiza_Impressao.lblPagina, InStr(frmVisualiza_Impressao.lblPagina, " de ") + 4)))) & "P�ginas: " & int_Conta_Pagina & " de " & Mid(frmVisualiza_Impressao.lblPagina, InStr(frmVisualiza_Impressao.lblPagina, " de ") + 4)
        Printer.Print String(114, "-")
        Printer.Print
        
    End If
    
    Printer.FontBold = False
    Printer.FontSize = Tamanho_Fonte_Rel

    int_X_Atual = Printer.CurrentX
    int_Y_Atual = Printer.CurrentY
    
    'Linha
    Printer.DrawWidth = 2
    Printer.Line (0, 1100)-(Printer.ScaleWidth - 160, Printer.ScaleHeight - 600), vbBlack, B

    Printer.CurrentX = int_X_Atual
    Printer.CurrentY = int_Y_Atual
    
    Imprimir_Rodape_Pagina_Impressora
    
    Printer.Print
    'Ajuste � moldura
    Printer.Print
    
    Exit Sub
    
TrataErro:
    
    If Err.Number <> 0 Then
        MsgBox Err.Number & " - " & Err.Description, vbCritical, wNomeSistema
    End If
    Exit Sub
    Resume
End Sub

Private Sub Imprimir_Cabecalho_Dados()
    
    Dim int_X As Integer
    
    On Error GoTo TrataErro
    
    'Cabe�aho das Colunas
    If Linhas_RCabecalho > 0 And Printer.Page > 1 Then
    
        For int_X = 0 To Linhas_RCabecalho - 1
        
            Printer.FontBold = Not (Conteudo_Relatorio(CStr(int_X)).Negrito = "N")
            Printer.Print Spc(2); Conteudo_Relatorio(CStr(int_X)).Linha
            
        Next

    End If
    
    Exit Sub
    
TrataErro:
    If Err.Number <> 0 Then
        MsgBox Err.Number & " - " & Err.Description, vbCritical, wNomeSistema
    End If
End Sub

Private Sub Imprimir_Tela(ByVal int_Pag As Integer)

    Dim dbl_Pos_PLinha As Double                        'posi��o da primeira linha
    Dim int_Controle As Integer                         'linhas da p�gina
    Dim int_Tamanho_Folha As Integer                    'tamanho da folha em linhas
    Dim int_X As Integer
    Dim int_Y As Integer
    Dim X As String
    
    'Vari�veis de controle de  para quebra de p�gina
    Dim bol_Troca_Pag As Boolean
    Dim str_Linha_Ant As String
    Dim str_Negrito_Ant As String
    
    Dim lng_Cont_Linhas_Cab_Grupo As Long               'n�mero de linhas do cabe�alho de grupos
    Dim lng_Cont_Detalhe As Long                        'n�mero de registro dos detalhes do grupo
    Dim lng_IGrupo_Atual As Long                        '�ndice do grupo atual
    
    On Error GoTo TrataErro
    
    Reseta_Tela
    
    int_Tamanho_Folha = frmVisualiza_Impressao.picVisualiza.ScaleHeight - 1000 '1440
    
    If dbl_N_Pag = 0 Then
    
        'J� estamos incluido as tr�s Linhas do Cabe�alho
        dbl_N_Pag = 0
        
        For int_X = 0 To Conteudo_Relatorio.Count
            dbl_N_Pag = dbl_N_Pag + 1
        Next
        
        'calculando o n�mero de p�ginas
        dbl_N_Pag = dbl_N_Pag / Tamanho_Papel
        
        If dbl_N_Pag <= 1 Then
            dbl_N_Pag = 1
        ElseIf dbl_N_Pag < Int(dbl_N_Pag) + 1 Then
            dbl_N_Pag = Int(dbl_N_Pag) + 1
        End If
        
        If dbl_N_Pag < 1 Then dbl_N_Pag = 1
    
    End If
        
    int_Conta_Pagina = 1
    
    Imprimir_Cabecalho_Tela int_Pag
    
    'zera contador de linhas da folha
    int_Controle = 0
    
    'vamo pegar a posi��o da primeira linha
    dbl_Pos_PLinha = frmVisualiza_Impressao.picVisualiza.CurrentY
    
    For int_X = 0 To Conteudo_Relatorio.Count - 1
    
        'incrementa linha
        int_Controle = int_Controle + 1
            
        'se atingirmos o tamanho m�ximo da folha ou tivermos que
        'quebrar a p�gina para n�o quebra o grupo...
        If int_Controle > Tamanho_Papel Or bol_Troca_Pag Then
            
            'n�mero de p�ginas
            int_Conta_Pagina = int_Conta_Pagina + 1
            If int_Pag = 1 Then
                int_Controle = 0
            Else
                int_Controle = Linhas_RCabecalho + 1
            End If
            
            'flag para trocar de p�gina
            bol_Troca_Pag = False
            
        End If
        
        frmVisualiza_Impressao.imgGrafico(tDadosGrafico.int_Indice).Visible = False
        
        'a rotina ir� varrer todo o array verificando registro por registro...
        'e s� entrar� neste IF se o registro pertencer � esta p�gina.
        If int_Pag = int_Conta_Pagina Then
            
            'Se estivermos na primeira linha E tivermos linhas de cabe�alho E n�o estivermos na primira posi��o do array...
            If frmVisualiza_Impressao.picVisualiza.CurrentY = dbl_Pos_PLinha And Linhas_RCabecalho > 0 And int_Pag <> 1 Then
                
                'plota o cabe�alho...
                For int_Y = 0 To Linhas_RCabecalho - 1
                    
                    frmVisualiza_Impressao.picVisualiza.FontBold = Not (Conteudo_Relatorio(CStr(int_Y)).Negrito = "N") 'False
                    frmVisualiza_Impressao.picVisualiza.CurrentX = 150
                    X = Imprime_Tela(frmVisualiza_Impressao.picVisualiza, Conteudo_Relatorio(CStr(int_Y)).Linha)
                    
                Next
            End If
            
            'rotina que plota a linha anterior caso tenha sido necess�rio for�ar a quebra da p�gina
            If str_Linha_Ant <> Empty Then
                frmVisualiza_Impressao.picVisualiza.FontBold = Not (str_Negrito_Ant = "N")
                frmVisualiza_Impressao.picVisualiza.CurrentX = 150
                X = Imprime_Tela(frmVisualiza_Impressao.picVisualiza, str_Linha_Ant)
                str_Linha_Ant = ""
                str_Negrito_Ant = ""
            End If
            
            'controla se a linha ser� em negrito
            frmVisualiza_Impressao.picVisualiza.FontBold = Not (Conteudo_Relatorio(CStr(int_X)).Negrito = "N")
            
            'se for uma linha na primeira posi��o do relat�rio, n�o imprime...
            If frmVisualiza_Impressao.picVisualiza.CurrentY = dbl_Pos_PLinha And Left(Conteudo_Relatorio(CStr(int_X)).Linha, 1) = "_" Then
            Else
                'se tem grupo...
                If Conteudo_Relatorio(CStr(int_X)).IGrupo > -1 Then
                    
                    'capturar o �ndice e inicializando vari�veis de controle
                    lng_IGrupo_Atual = Conteudo_Relatorio(CStr(int_X)).IGrupo
                    lng_Cont_Detalhe = 0
                    lng_Cont_Linhas_Cab_Grupo = 1
                    
                    'contar detalhes e linhas de cabe�alho do grupo
                    For int_Y = int_X + 1 To Conteudo_Relatorio.Count - 1
                    
                        If Conteudo_Relatorio(CStr(int_Y)).IGrupo = -1 Then
                            lng_Cont_Detalhe = lng_Cont_Detalhe + 1
                        ElseIf Conteudo_Relatorio(CStr(int_Y)).IGrupo > lng_IGrupo_Atual Then
                            Exit For
                        Else
                            lng_Cont_Linhas_Cab_Grupo = lng_Cont_Linhas_Cab_Grupo + 1
                        End If
                    Next
                    
                    'verificando se o grupo cabe no papel...
                    If lng_Cont_Detalhe + lng_Cont_Linhas_Cab_Grupo <= Tamanho_Papel Then
                        'cabe, logo verifica se o grupo cabe no espa�o restante
                        If lng_Cont_Detalhe + lng_Cont_Linhas_Cab_Grupo > Tamanho_Papel - int_Controle Then
                            bol_Troca_Pag = True
                        Else
                            bol_Troca_Pag = False
                        End If
                    Else
                        'se n�o couber, verifica se cabe o cabe�alho do grupo e o primeiro detalhe
                        If lng_Cont_Linhas_Cab_Grupo + IIf(lng_Cont_Detalhe > 0, 1, 0) > Tamanho_Papel - int_Controle Then
                            bol_Troca_Pag = True
                        Else
                            bol_Troca_Pag = False
                        End If
                    End If
                End If
                
                'se n�o quebra p�gina, IMPRIME!
                If Not bol_Troca_Pag Then
                    frmVisualiza_Impressao.picVisualiza.CurrentX = 150
                    X = Imprime_Tela(frmVisualiza_Impressao.picVisualiza, Conteudo_Relatorio(CStr(int_X)).Linha)
                End If
                                
            End If
        
        'Imprime gr�fico na p�gina que foi calculada...
        ElseIf int_Pag = int_Controle_Graf Then
        
            'gr�fico...
            If frmVisualiza_Impressao.imgGrafico(tDadosGrafico.int_Indice).width > frmVisualiza_Impressao.picVisualiza.ScaleWidth Then
                'auto ajuste...
                frmVisualiza_Impressao.imgGrafico(tDadosGrafico.int_Indice).width = frmVisualiza_Impressao.imgGrafico(tDadosGrafico.int_Indice).width * Round(1 - (frmVisualiza_Impressao.imgGrafico(tDadosGrafico.int_Indice).width / frmVisualiza_Impressao.picVisualiza.ScaleWidth - 1), 2)
            End If
            frmVisualiza_Impressao.imgGrafico(tDadosGrafico.int_Indice).Visible = True
            
        
        Else
            'o else � necess�rio para calcular as p�ginas com quebras de grupo e...
            'armazenar as informa��es que por ventura ter�o de ser impressas em uma nova p�gina
            If frmVisualiza_Impressao.picVisualiza.CurrentY = dbl_Pos_PLinha And Left(Conteudo_Relatorio(CStr(int_X)).Linha, 1) = "_" Then
            Else
                If Conteudo_Relatorio(CStr(int_X)).IGrupo > -1 Then
                
                    lng_IGrupo_Atual = Conteudo_Relatorio(CStr(int_X)).IGrupo
                    lng_Cont_Detalhe = 0
                    lng_Cont_Linhas_Cab_Grupo = 1
                    
                    For int_Y = int_X + 1 To Conteudo_Relatorio.Count - 1
                        If Conteudo_Relatorio(CStr(int_Y)).IGrupo = -1 Then
                            lng_Cont_Detalhe = lng_Cont_Detalhe + 1
                        ElseIf Conteudo_Relatorio(CStr(int_Y)).IGrupo > lng_IGrupo_Atual Then
                            Exit For
                        Else
                            lng_Cont_Linhas_Cab_Grupo = lng_Cont_Linhas_Cab_Grupo + 1
                        End If
                    Next
                    
                    If lng_Cont_Detalhe + lng_Cont_Linhas_Cab_Grupo <= Tamanho_Papel Then
                        If lng_Cont_Detalhe + lng_Cont_Linhas_Cab_Grupo > Tamanho_Papel - int_Controle Then
                            bol_Troca_Pag = True
                        Else
                            bol_Troca_Pag = False
                        End If
                    Else
                        If lng_Cont_Linhas_Cab_Grupo + IIf(lng_Cont_Detalhe > 0, 1, 0) > Tamanho_Papel - int_Controle Then
                            bol_Troca_Pag = True
                        Else
                            bol_Troca_Pag = False
                        End If
                    End If
                    
                End If
                'se tem que quebrar a p�gina
                If bol_Troca_Pag And int_Conta_Pagina = int_Pag - 1 Then
                    'grava a linha para plotar na pr�xima p�gina
                    str_Linha_Ant = Conteudo_Relatorio(CStr(int_X)).Linha
                    str_Negrito_Ant = Conteudo_Relatorio(CStr(int_X)).Negrito
                End If
            
            End If
            
        End If
       
    Next
    
    If tDadosGrafico.bol_Grafico Then
       Plota_Grafico_Tela tDadosGrafico.int_Indice
    End If
    
    Imprimir_Rodape_Pagina_Tela

    'Ajusta label do Visualiza
    frmVisualiza_Impressao.lblPagina = "P�gina " & CStr(int_Pag) & " de " & CStr(dbl_N_Pag)
    Exit Sub

TrataErro:
   
   If Err.Number <> 0 Then
        MsgBox Err.Number & " - " & Err.Description, vbCritical, wNomeSistema
   End If
End Sub

Private Sub Imprimir_Cabecalho_Tela(ByVal int_Pag As Integer)

    Dim X As String
    Dim int_Aux As Integer
    
    On Error GoTo TrataErro
    
    'frmVisualiza_Impressao.picVisualiza.FontBold = True
    
    frmVisualiza_Impressao.picVisualiza.FontName = Fonte_Relatorio
    'frmVisualiza_Impressao.picVisualiza.FontSize = Tamanho_Fonte_Cab - 2
    
    'frmVisualiza_Impressao.picVisualiza.CurrentX = 150
    
    'aki - 30/08/06 21h
    frmVisualiza_Impressao.picVisualiza.CurrentY = 100
    Call Imprime_Tela(frmVisualiza_Impressao.picVisualiza, Titulo_Empresa_Relatorio & Space(50) & "Data de Criacao: " & Data_Criacao_Relatorio)
    'frmVisualiza_Impressao.picVisualiza.CurrentY = frmVisualiza_Impressao.picVisualiza.CurrentY '+ 1
    Call Imprime_Tela(frmVisualiza_Impressao.picVisualiza, String(96, "-"))
    
    'frmVisualiza_Impressao.picVisualiza.CurrentY = frmVisualiza_Impressao.picVisualiza.CurrentY '+ 1
    Call Imprime_Tela(frmVisualiza_Impressao.picVisualiza, Titulo_Cliente_Relatorio & Space(96 - (Len(Trim(Titulo_Cliente_Relatorio)) + Len("Data de Geracao: " & Format(Date, "dd/MM/yyyy")))) & "Data de Geracao: " & Format(Date, "dd/MM/yyyy"))
    'frmVisualiza_Impressao.picVisualiza.CurrentY = frmVisualiza_Impressao.picVisualiza.CurrentY '+ 30
    Call Imprime_Tela(frmVisualiza_Impressao.picVisualiza, Titulo_Relatorio & Space(96 - (Len(Titulo_Relatorio) + (Len("Hora de Geracao: " & Format(Now, "HH:nn:ss  "))))) & "Hora de Geracao: " & Format(Now, "HH:nn:ss  "))
    'frmVisualiza_Impressao.picVisualiza.CurrentY = frmVisualiza_Impressao.picVisualiza.CurrentY '+ 30
    Call Imprime_Tela(frmVisualiza_Impressao.picVisualiza, Sub_Titulo_Relatorio & Space(92 - (Len(Sub_Titulo_Relatorio) + Len("P�ginas: " & int_Conta_Pagina & " de " & Mid(frmVisualiza_Impressao.lblPagina, InStr(frmVisualiza_Impressao.lblPagina, " de ") + 4)))) & "P�ginas: " & int_Conta_Pagina & " de " & Mid(frmVisualiza_Impressao.lblPagina, InStr(frmVisualiza_Impressao.lblPagina, " de ") + 4))
    'frmVisualiza_Impressao.picVisualiza.CurrentY = frmVisualiza_Impressao.picVisualiza.CurrentY '+ 1
    Call Imprime_Tela(frmVisualiza_Impressao.picVisualiza, String(96, "-"))
    
    'frmVisualiza_Impressao.picVisualiza.CurrentX = frmVisualiza_Impressao.picVisualiza.CurrentY '+ 1
    Call Imprime_Tela(frmVisualiza_Impressao.picVisualiza, "")
    'call Imprime_Tela(frmVisualiza_Impressao.picVisualiza, "")
    
    Exit Sub
    
TrataErro:

   If Err.Number <> 0 Then
        MsgBox Err.Number & " - " & Err.Description, vbCritical, wNomeSistema
   End If
   
   Exit Sub
   Resume
    
End Sub

Private Function Imprime_Tela(objPrint As Object, ByVal Texto As String) As String

    objPrint.Print Space(2) & Texto

End Function

Public Sub Zera_Colecao()

    Dim int_X As Integer
        
    For int_X = 0 To Conteudo_Relatorio.Count - 1
        Conteudo_Relatorio.Remove (CStr(int_X))
    Next
    
    Set Conteudo_Relatorio = Nothing
    
End Sub

'Centraliza texto na tela (n�o � utilizada ainda para a impress�o na imp.
Private Function Centraliza_Texto(str_Texto) As Integer
    
    Dim int_Len As Integer
    
    int_Len = Len(str_Texto)
    Centraliza_Texto = (frmVisualiza_Impressao.picVisualiza.width - (120 * int_Len)) / 2
    
End Function

'Prepara a tela de acordo com o tamanho e formato do papel
Public Sub Configura_Tela()

    'Configura��es para A4 retrato
    If Tamanho_Papel = 75 And Orientacao_Papel = 1 Then
    
        frmVisualiza_Impressao.picVisualiza.Height = 18000
        frmVisualiza_Impressao.picVisualiza.width = 11400
        frmVisualiza_Impressao.vscVisualiza.Max = 11800
        frmVisualiza_Impressao.fraImpressao.Height = 6300
        frmVisualiza_Impressao.vschVisualiza.Visible = False
        
    'Configura��es para A4 paisagem
    ElseIf Tamanho_Papel = 50 And Orientacao_Papel = 2 Then
    
        frmVisualiza_Impressao.picVisualiza.Height = 13000
        frmVisualiza_Impressao.picVisualiza.width = 18010
        frmVisualiza_Impressao.vscVisualiza.Max = 7000
        frmVisualiza_Impressao.fraImpressao.Height = 6000
        frmVisualiza_Impressao.vschVisualiza.Visible = True
        frmVisualiza_Impressao.vschVisualiza.Max = 6700
        
    End If
    
End Sub

'Reinicia a tela
Private Sub Reseta_Tela()

    frmVisualiza_Impressao.vscVisualiza.Value = 0
    frmVisualiza_Impressao.picVisualiza.Cls
    Set frmVisualiza_Impressao.picVisualiza.Picture = Nothing
    frmVisualiza_Impressao.picVisualiza.Refresh
    
End Sub

Private Sub Plota_Grafico_Tela(ByVal I As Integer)
    
    Dim lng_Area As Long
    
    If int_Controle_Graf = Impressao.Numero_Paginas Then Exit Sub
    
    With tDadosGrafico
    
        'Picture que armazenar� o gr�fico
        If .int_Indice <> 0 Then
            Load frmVisualiza_Impressao.imgGrafico(I)
        End If
                
        frmVisualiza_Impressao.imgGrafico(I).BorderStyle = 0
        frmVisualiza_Impressao.imgGrafico(I).width = .obj_Grafico.ScaleWidth * .dbl_Tamanho
        frmVisualiza_Impressao.imgGrafico(I).Height = .obj_Grafico.ScaleHeight * .dbl_Tamanho
        frmVisualiza_Impressao.imgGrafico(I).Stretch = True
        frmVisualiza_Impressao.imgGrafico(I).Visible = True
        
        lng_Area = .int_GrafX + frmVisualiza_Impressao.imgGrafico(I).width

        'ambos
        If .int_Alinha = alAmbos Then
        
            frmVisualiza_Impressao.imgGrafico(I).Top = (frmVisualiza_Impressao.picVisualiza.ScaleHeight - .obj_Grafico.ScaleHeight) \ 2
            frmVisualiza_Impressao.imgGrafico(I).Left = (frmVisualiza_Impressao.picVisualiza.ScaleWidth - .obj_Grafico.ScaleWidth) \ 2
            .int_GrafY = (frmVisualiza_Impressao.picVisualiza.ScaleHeight - .obj_Grafico.ScaleHeight) \ 2
            .int_GrafX = (frmVisualiza_Impressao.picVisualiza.ScaleWidth - .obj_Grafico.ScaleWidth) \ 2
            
        'horizontal
        ElseIf .int_Alinha = alCentraliza_H Then
        
            frmVisualiza_Impressao.imgGrafico(I).Top = .int_GrafY + 500
            frmVisualiza_Impressao.imgGrafico(I).Left = (frmVisualiza_Impressao.picVisualiza.ScaleWidth - .obj_Grafico.ScaleWidth) \ 2
            .int_GrafX = (frmVisualiza_Impressao.picVisualiza.ScaleWidth - .obj_Grafico.ScaleWidth) \ 2
            
        
        '*** SOMENTE NOS CASOS ABAIXO PODE SER IMPLEMENTADO O GR�FICO ISOLADO EM UMA NOVA P�GINA ***
        
        'vertical
        ElseIf .int_Alinha = alCentraliza_V Then
        
            'se cabe...
            If ((frmVisualiza_Impressao.picVisualiza.ScaleWidth - .int_GrafX) - lng_Area) > 0 Then
            
                frmVisualiza_Impressao.imgGrafico(I).Top = (frmVisualiza_Impressao.picVisualiza.ScaleHeight - .obj_Grafico.ScaleHeight) \ 2
                frmVisualiza_Impressao.imgGrafico(I).Left = .int_GrafX
                .int_GrafY = (frmVisualiza_Impressao.picVisualiza.ScaleHeight - .obj_Grafico.ScaleHeight) \ 2
            
            Else
            
                frmVisualiza_Impressao.imgGrafico(I).Top = (frmVisualiza_Impressao.picVisualiza.ScaleHeight - .obj_Grafico.ScaleHeight) \ 2
                frmVisualiza_Impressao.imgGrafico(I).Left = 0
                .int_GrafY = (frmVisualiza_Impressao.picVisualiza.ScaleHeight - .obj_Grafico.ScaleHeight) \ 2 + 2000
                dbl_N_Pag = dbl_N_Pag + 1
                int_Controle_Graf = dbl_N_Pag
                frmVisualiza_Impressao.imgGrafico(I).Visible = False
            
            End If
            
        'manual
        'ElseIf .int_Alinha = 4 Then
        Else
                    
            'se cabe...
            If ((frmVisualiza_Impressao.picVisualiza.ScaleWidth - .int_GrafX) - lng_Area) > 0 Then
            
                frmVisualiza_Impressao.imgGrafico(I).Top = .int_GrafY + 500
                frmVisualiza_Impressao.imgGrafico(I).Left = .int_GrafX
                
            Else
            
                frmVisualiza_Impressao.imgGrafico(I).Top = 100 + 2000
                
                If frmVisualiza_Impressao.imgGrafico(I).width > frmVisualiza_Impressao.picVisualiza.ScaleWidth Then
                    frmVisualiza_Impressao.imgGrafico(I).Left = 0
                Else
                    'tentando centralizar horizontal...
                    frmVisualiza_Impressao.imgGrafico(I).Left = (frmVisualiza_Impressao.picVisualiza.ScaleWidth - .obj_Grafico.ScaleWidth) \ 2
                End If
                
                dbl_N_Pag = dbl_N_Pag + 1
                int_Controle_Graf = dbl_N_Pag
                frmVisualiza_Impressao.imgGrafico(I).Visible = False
            
            End If
            
        End If
                
        frmVisualiza_Impressao.imgGrafico(I).Picture = .obj_Grafico.Picture
    
    End With
    
End Sub

Public Function Linha_Relatorio() As String

    Dim Linha As String
    
    On Error GoTo TrataErro

    Linha = "_"
    
    If Destino_Relatorio = False Then
        If Orientacao_Papel = 1 Then
            While frmVisualiza_Impressao.CurrentX + frmVisualiza_Impressao.TextWidth(Linha) <= 9110 '9500                Linha = Linha & "_"
                Linha = Linha & "_"
            Wend
        Else
            While frmVisualiza_Impressao.CurrentX + frmVisualiza_Impressao.TextWidth(Linha) <= 14760 '15000
                Linha = Linha & "_"
            Wend
        End If
    Else
        While Printer.CurrentX + Printer.TextWidth(Linha) <= Printer.ScaleWidth
            Linha = Linha & "_"
        Wend
    End If
    
    Linha_Relatorio = Linha
    
    Exit Function
    
TrataErro:

   If Err.Number <> 0 Then
        MsgBox Err.Number & " - " & Err.Description, vbCritical, wNomeSistema
        'DB.Rollback
   End If

End Function

Private Sub Class_Initialize()
    int_Controle_Graf = 0
    tDadosGrafico.bol_Grafico = False
End Sub

Private Sub Imprimir_Rodape_Pagina_Tela()
    Dim booEstado As Boolean
    
    booEstado = frmVisualiza_Impressao.picVisualiza.FontBold
    
    frmVisualiza_Impressao.picVisualiza.FontBold = True
    frmVisualiza_Impressao.picVisualiza.CurrentY = frmVisualiza_Impressao.picVisualiza.ScaleHeight - 475
    Call Imprime_Tela(frmVisualiza_Impressao.picVisualiza, strRodape_Pagina)
    frmVisualiza_Impressao.picVisualiza.FontBold = booEstado
    
End Sub

Private Sub Imprimir_Rodape_Pagina_Impressora()
    
    Dim int_X_Atual As Integer
    Dim int_Y_Atual As Integer

    int_X_Atual = Printer.CurrentX
    int_Y_Atual = Printer.CurrentY

    Printer.FontBold = True
    Printer.CurrentY = Printer.ScaleHeight - 475
    Printer.Print strRodape_Pagina
    Printer.FontBold = False
    
    Printer.CurrentX = int_X_Atual
    Printer.CurrentY = int_Y_Atual

End Sub

Public Property Let Rodape_Pagina(strDado As String)
    strRodape_Pagina = strDado
End Property
