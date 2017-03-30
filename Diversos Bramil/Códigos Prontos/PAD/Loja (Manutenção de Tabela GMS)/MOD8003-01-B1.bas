Attribute VB_Name = "Module1"
Option Explicit
'*************** 1 - Adicionar aqui os bancos de dados a serem utilizados
Public Const pstrBancosUsados = "bdConfus - bdLog - bdGMS002 "

Public pdbGMS002 As Database
Public pwrkGMS002 As Workspace

Public pstrLocacaobdGMS001 As String
Public pstrLocacaobdGMS002 As String
Public pstrLocacaobdGMS005 As String
'Public pstrLocacaobdLog    As String
Public pstrSenhabdLog      As String


Public pstrPathL As String
Public piCodEmpresatPar As Integer

Public pdtInicialAtualizCupomtPar As Date

Public Sub Main()
    On Error GoTo Erro
    If Not pfboRotinasIniciais("GMS8003-01-VB") Then End
    Call ppAbre_BDAcesso(pWrkArea, pstrCodPrograma, pdbConfus, pstrLocacaobdConfus, pstrSenhaBancoDadosBDCONFUS)
    
    'Pego a Locação do arquivo bdGMS002
    pstrSql = "Select * " & _
                "FROM tLocacaoBancoDados " & _
               "where bCodEsquemaBancoDadostLoc = " & pbCodEsquemaBancoDadostLogin & " " & _
               "AND strCodBancoDadostLoc = 'bdGMS002'"
    Set prsSeleção = pfrsSelecao(pdbConfus, pstrSql)
    If prsSeleção.EOF Then
        MsgBox "Locação do Banco de Dados bdGMS002 nao encontrado na Tabela tLocacaoBancoDados", vbCritical + vbOKOnly, "Form_Load"
        End: Exit Sub
    End If
    pstrLocacaobdGMS002 = Empty & prsSeleção.Fields("strLocacaoBancoDadostLoc")
    
    'Pega locação da base de dados bdGMS005
    pstrSql = "Select * " & _
                "FROM tLocacaoBancoDados " & _
               "where bCodEsquemaBancoDadostLoc = " & pbCodEsquemaBancoDadostLogin & " " & _
               "AND strCodBancoDadostLoc = 'bdGMS005'"
    Set prsSeleção = pfrsSelecao(pdbConfus, pstrSql)
    If prsSeleção.EOF Then
        MsgBox "Locação do Banco de Dados bdGMS005 nao encontrado na Tabela tLocacaoBancoDados", vbCritical + vbOKOnly, "Form_Load"
        End: Exit Sub
    End If
    pstrLocacaobdGMS005 = Empty & prsSeleção.Fields("strLocacaoBancoDadostLoc")
    pstrSenhaBancoDadostLoc = Empty & prsSeleção.Fields("strSenhaBancoDadostLoc")
    
    'Pega locação da base de dados bdLog
    pstrSql = "Select * " & _
                "FROM tLocacaoBancoDados " & _
               "where bCodEsquemaBancoDadostLoc = " & pbCodEsquemaBancoDadostLogin & " " & _
               "AND strCodBancoDadostLoc = 'bdLog'"
    Set prsSeleção = pfrsSelecao(pdbConfus, pstrSql)
    If prsSeleção.EOF Then
        MsgBox "Locação do Banco de Dados bdLog nao encontrado na Tabela tLocacaoBancoDados", vbCritical + vbOKOnly, "Form_Load"
        End: Exit Sub
    End If
    pstrLocacaobdLog = Empty & prsSeleção.Fields("strLocacaoBancoDadostLoc")
    pstrSenhabdLog = Empty & prsSeleção.Fields("strSenhaBancoDadostLoc")
    
    
    
    
    'Pega locação da base de dados bdGMS001
    pstrSql = "Select * " & _
                "FROM tLocacaoBancoDados " & _
               "where bCodEsquemaBancoDadostLoc = " & pbCodEsquemaBancoDadostLogin & " " & _
               "AND strCodBancoDadostLoc = 'bdGMS001'"
    Set prsSeleção = pfrsSelecao(pdbConfus, pstrSql)
    If prsSeleção.EOF Then
        MsgBox "Locação do Banco de Dados bdGMS001 nao encontrado na Tabela tLocacaoBancoDados", vbCritical + vbOKOnly, "Form_Load"
        End: Exit Sub
    End If
    pstrLocacaobdGMS001 = Empty & prsSeleção.Fields("strLocacaoBancoDadostLoc")
    
    pdbConfus.Close
    
    'Abrindo arquivo bdGMS002
    'Call ppAbre_BDAcesso(pwrkGMS002, pstrCodPrograma & "dbGMS002", pdbGMS002, pstrLocacaobdGMS002, pstrSenhaBancoDadostLoc)
    FormPrincipal.Show
    
    Exit Sub
Erro:
    MsgBox "Ocorreu o erro número " & Str(Err.Number) & Chr(13) & Chr(13) & Err.Description, vbCritical, "SubMain"
    Err.Clear
    End
End Sub

