Attribute VB_Name = "GMMod_bdGMS002"
Public Enum GMen_tProdutosBarra
    'Flags para
    Public Enum GMen_bIndCompratPrdBar
        GMen_CompraLiberada = 1
        GMen_CompraSuspensa = 2
    End Enum
    
    Public Enum GMen_bIndEstoqueSalaotPrdBar
        GMen_EstSal_SemControle = 1
        GMen_EstSal_ComControle = 2
    End Enum
    
    Public Enum GMen_bIndEstoquetPrdBar
        GMen_Est_SemControle = 1
        GMen_Est_ComQuebra = 2
        GMen_Est_B�sico = 3
        GMen_Est_Regular = 4
        GMen_Est_Priorit�rio = 5
    End Enum
    
    Public Enum GMen_bIndEtiquetaPrecoVendatPrdBar
        GMen_EtiqPr�Venda_SemEtiq = 1
        GMen_EtiqPr�Venda_Simples = 2
        GMen_EtiqPr�Venda_Balan�a = 3
    End Enum
    
    Public Enum GMen_IndExposicaoPrecoVendatPrdBar
        GMen_ExpPr�Venda_SemExp = 1
        GMen_ExpPr�Venda_EtiqGond = 2
        GMen_ExpPr�Venda_Cart025 = 3
        GMen_ExpPr�Venda_Cart050 = 4
        GMen_ExpPr�Venda_Cart1Pag = 5
        GMen_ExpPr�Venda_Cart2Pag = 6
        GMen_ExpPr�Venda_TabPre�Cart = 7
        GMen_ExpPr�Venda_TabPre�Mont = 8
    End Enum
        
    Public Enum GMen_bIndPedidotPrdBar
        GMen_PedMatrizMatriz = 1
        GMen_PedMatrizLoja = 2
        GMen_PedLojaLoja = 3
    End Enum
    
    
    
    
    
    Public Enum GMen_
        GMen_StatusAtivoExpe = 1
        GMen_StatusAtivoConf = 2
        GMen_StatusAtivoExcl = 3
        GMen_StatusAtivoSusp = 4
    End Enum
    
    GMen_TabelaC�digo = 1
    GMen_TabelaPre�oListagem = 2
    GMen_TabelaPre�oCartaz = 3
    
    
    
End Enum
