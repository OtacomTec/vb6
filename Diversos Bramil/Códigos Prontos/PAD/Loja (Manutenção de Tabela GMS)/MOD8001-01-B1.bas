Attribute VB_Name = "Module1"
Option Explicit
'*************** 1 - Adicionar aqui os bancos de dados a serem utilizados
Public Const pstrBancosUsados = "bdConfus - bdLog - bdGMS002 - txtGMS002"

Public pdbGMS002 As Database
Public pwrkGMS002 As Workspace

Public pstrLocacaobdGMS002 As String
Public pstrLocacaobdLogVenda As String
Public pstrLocacaotxtGMS002 As String

Public pstrPathL As String
Public piCodEmpresatPar As Integer

Public pdtInicialAtualizCupomtPar As Date

'Numeros Inical e Final dos PDV's
Public pbNumInicialPDVtPar As Byte
Public pbNumFinalPDVtPar As Byte
Public pboRepouso As Boolean


Public Sub Main()
    
    On Error GoTo Erro
    
    If Not pfboRotinasIniciais("GMS8003-01-VB") Then End
    
    Call ppAbre_BDAcesso(pWrkArea, pstrCodPrograma, pdbConfus, pstrLocacaobdConfus, pstrSenhaBancoDadosBDCONFUS)
    
    'Pego a Locação do arquivo bdGMS002
    pstrSql = "Select * FROM tLocacaoBancoDados where bCodEsquemaBancoDadostLoc = " & pbCodEsquemaBancoDadostLogin & " AND strCodBancoDadostLoc = 'bdGMS002'"
    Set prsSeleção = pfrsSelecao(pdbConfus, pstrSql)
    
    If prsSeleção.EOF Then
        MsgBox "Locação do Banco de Dados bdGMS002 nao encontrado na Tabela tLocacaoBancoDados", vbCritical + vbOKOnly, "Form_Load"
        End: Exit Sub
    End If
    
    pstrLocacaobdGMS002 = Empty & prsSeleção.Fields("strLocacaoBancoDadostLoc")
    pstrSenhaBancoDadostLoc = Empty & prsSeleção.Fields("strSenhaBancoDadostLoc")

    
    'Abrindo arquivo bdGMS002
    Call ppAbre_BDAcesso(pwrkGMS002, pstrCodPrograma & "dbGMS002", pdbGMS002, pstrLocacaobdGMS002, pstrSenhaBancoDadostLoc)
    
    FormPrincipal.Show
    
    Exit Sub
Erro:
    MsgBox "Ocorreu o erro número " & Str(Err.Number) & Chr(13) & Chr(13) & Err.Description, vbCritical, "SubMain"
    Err.Clear
    End
End Sub




