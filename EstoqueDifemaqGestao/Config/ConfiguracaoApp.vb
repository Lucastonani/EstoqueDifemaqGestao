Imports System.Drawing

Public Class ConfiguracaoApp

    ' Nomes das tabelas do Power Query
    Public Const TABELA_PRODUTOS As String = "tblProdutos"
    Public Const TABELA_ESTOQUE As String = "tblEstoqueVisao"
    Public Const TABELA_COMPRAS As String = "tblCompras"
    Public Const TABELA_VENDAS As String = "tblVendas"

    ' Configurações de imagens
    Public Const CAMINHO_IMAGENS As String = "C:\ImagesEstoque"
    Public Shared ReadOnly EXTENSOES_IMAGEM As String() = {".jpg", ".jpeg", ".png", ".bmp", ".gif"}

    ' Configurações de interface
    Public Const COR_HEADER_GRID As String = "#2E86AB"
    Public Const COR_LINHA_ALTERNADA As String = "#F5F5F5"
    Public Const COR_SELECAO As String = "#4A90E2"
    Public Const COR_HEADER As String = "#2E86AB"
    Public Const COR_ALTERNADA As String = "#F8F9FA"
    Public Const ALTURA_LINHA_GRID As Integer = 25

    ' Configurações de sistema
    Public Const TIMEOUT_POWERQUERY As Integer = 60 ' segundos
    Public Const DEBOUNCE_DELAY As Integer = 300 ' milissegundos
    Public Const TAMANHO_MAXIMO_IMAGEM As Long = 5242880 ' 5MB

    ' Configurações de log
    Public Const CAMINHO_LOG As String = "C:\Logs\GestaoEstoque"
    Public Const TAMANHO_MAX_LOG As Long = 10485760 ' 10MB

    ' Configurações de performance
    Public Const LIMITE_REGISTROS_GRID As Integer = 5000
    Public Const TIMEOUT_CARREGAMENTO_IMAGEM As Integer = 3000 ' ms

    ' Mensagens padrão
    Public Const MSG_PRODUTO_NAO_ENCONTRADO As String = "Produto não encontrado"
    Public Const MSG_IMAGEM_NAO_ENCONTRADA As String = "Imagem não disponível"
    Public Const MSG_DADOS_CARREGANDO As String = "Carregando dados..."

    ' Métodos auxiliares para cores
    Public Shared Function ObterCorHeader() As Color
        Return ColorTranslator.FromHtml(COR_HEADER_GRID)
    End Function

    Public Shared Function ObterCorSelecao() As Color
        Return ColorTranslator.FromHtml(COR_SELECAO)
    End Function

    Public Shared Function ObterCorAlternada() As Color
        Return ColorTranslator.FromHtml(COR_LINHA_ALTERNADA)
    End Function

End Class