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
        GMen_Est_Básico = 3
        GMen_Est_Regular = 4
        GMen_Est_Prioritário = 5
    End Enum
    
    Public Enum GMen_bIndEtiquetaPrecoVendatPrdBar
        GMen_EtiqPrçVenda_SemEtiq = 1
        GMen_EtiqPrçVenda_Simples = 2
        GMen_EtiqPrçVenda_Balança = 3
    End Enum
    
    Public Enum GMen_IndExposicaoPrecoVendatPrdBar
        GMen_ExpPrçVenda_SemExp = 1
        GMen_ExpPrçVenda_EtiqGond = 2
        GMen_ExpPrçVenda_Cart025 = 3
        GMen_ExpPrçVenda_Cart050 = 4
        GMen_ExpPrçVenda_Cart1Pag = 5
        GMen_ExpPrçVenda_Cart2Pag = 6
        GMen_ExpPrçVenda_TabPreçCart = 7
        GMen_ExpPrçVenda_TabPreçMont = 8
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
    
    GMen_TabelaCódigo = 1
    GMen_TabelaPreçoListagem = 2
    GMen_TabelaPreçoCartaz = 3
    
    
    
End Enum
