# 📋 Guia de Padrões de Código - EstoqueDifemaqGestao

## 🎯 Objetivo
Este documento estabelece padrões de código consistentes para o projeto EstoqueDifemaqGestao, baseados nas melhores práticas já implementadas no sistema.

---

## 📏 1. Convenções de Nomenclatura

### **1.1 Classes e Interfaces**
**Padrão**: PascalCase
```vb
✅ Correto:
Public Class PowerQueryManager
Public Class UcReposicaoEstoque
Public Interface IDataProcessor

❌ Incorreto:
Public Class powerQueryManager
Public Class ucReposicaoestoque
```

**Convenções Específicas:**
- **UserControls**: Prefixo `Uc` + nome descritivo
- **Forms**: Sufixo `Form` (ex: `MainForm`)
- **Managers**: Sufixo `Manager` (ex: `PowerQueryManager`)
- **Helpers**: Sufixo `Helper` (ex: `DataHelper`)

### **1.2 Métodos e Propriedades**
**Padrão**: PascalCase
```vb
✅ Correto:
Public Function ObterTabela(nomeTabela As String) As ListObject
Public Sub AtualizarTodasConsultas()
Public Property DadosCarregados As Boolean

❌ Incorreto:
Public Function obterTabela()
Public Sub atualizarTodasConsultas()
```

**Convenções por Tipo:**
- **Ações**: Verbo + substantivo (`CarregarDados`, `AtualizarConsultas`)
- **Consultas**: `Obter` + substantivo (`ObterTabela`, `ObterStatus`)
- **Verificações**: `Verificar` + substantivo (`VerificarStatus`)
- **Propriedades**: Substantivo/adjetivo (`DadosCarregados`, `CacheValido`)

### **1.3 Variáveis e Parâmetros**
**Padrão**: camelCase
```vb
✅ Correto:
Dim dadosCarregados As Boolean
Dim nomeTabela As String
Private ultimaAtualizacao As DateTime
Function ProcessarDados(codigoProduto As String)

❌ Incorreto:
Dim DadosCarregados As Boolean
Dim nome_tabela As String
```

**Convenções Específicas:**
- **Variáveis booleanas**: `is`, `tem`, `esta` + descrição (`isCarregando`, `temDados`, `estaValido`)
- **Contadores**: `total`, `numero`, `quantidade` + substantivo (`totalRegistros`, `numeroLinhas`)
- **Índices**: `i`, `j`, `k` para loops simples, nomes descritivos para contextos específicos

### **1.4 Constantes**
**Padrão**: UPPER_CASE com underscore
```vb
✅ Correto:
Public Const CACHE_TIMEOUT_MINUTES As Integer = 5
Public Const CAMINHO_IMAGENS As String = "C:\ImagesEstoque\"
Private Const WM_SETREDRAW As Integer = &HB

❌ Incorreto:
Public Const CacheTimeoutMinutes As Integer = 5
Public Const caminhoImagens As String = "..."
```

### **1.5 Controles de Interface**
**Padrão**: Prefixo de tipo + nome descritivo (PascalCase)
```vb
✅ Correto:
Private dgvProdutos As DataGridView
Private btnAtualizar As Button
Private txtFiltro As TextBox
Private lblStatus As Label
Private grpImagem As GroupBox

❌ Incorreto:
Private dataGridView1 As DataGridView
Private button1 As Button
```

**Prefixos Padrão:**
- `btn` - Button
- `txt` - TextBox  
- `lbl` - Label
- `dgv` - DataGridView
- `cmb` - ComboBox
- `grp` - GroupBox
- `pnl` - Panel
- `pic` - PictureBox

### **1.6 Campos Privados**
**Padrão**: camelCase, começando com letra minúscula
```vb
✅ Correto:
Private workbook As Workbook
Private isDisposed As Boolean
Private tabelasCache As Dictionary(Of String, ListObject)
Private ultimaAtualizacao As DateTime

❌ Incorreto:
Private _workbook As Workbook
Private IsDisposed As Boolean
```

---

## 📝 2. Documentação e Comentários

### **2.1 Comentários XML (Obrigatório para APIs Públicas)**
```vb
✅ Correto:
''' <summary>
''' Obtém uma tabela específica do workbook com cache inteligente
''' </summary>
''' <param name="nomeTabela">Nome da tabela a ser obtida (case-insensitive)</param>
''' <returns>ListObject da tabela encontrada ou Nothing se não encontrada</returns>
''' <exception cref="ArgumentException">Quando nomeTabela é nulo ou vazio</exception>
''' <example>
''' <code>
''' Dim tabela = manager.ObterTabela("tblProdutos")
''' If tabela IsNot Nothing Then
'''     Console.WriteLine($"Tabela tem {tabela.ListRows.Count} linhas")
''' End If
''' </code>
''' </example>
''' <remarks>
''' Implementa cache inteligente com timeout de 5 minutos.
''' Busca case-insensitive em todas as planilhas do workbook.
''' </remarks>
Public Function ObterTabela(nomeTabela As String) As ListObject
```

**Elementos Obrigatórios:**
- `<summary>`: Descrição concisa (1-2 linhas)
- `<param>`: Para cada parâmetro
- `<returns>`: Para funções que retornam valores
- `<exception>`: Para exceções que podem ser lançadas

**Elementos Opcionais:**
- `<example>`: Código de exemplo funcional
- `<remarks>`: Informações adicionais, performance, comportamento
- `<see>`: Referências a outras classes/métodos

### **2.2 Comentários Inline**
```vb
✅ Correto:
' Configurar Excel para melhor performance durante atualização
.DisplayAlerts = False
.ScreenUpdating = False

' Cache é invalidado automaticamente após atualizações
InvalidarCache()

' Buscar apenas em tabelas críticas do sistema
For Each nomeTabela In tabelasCriticas

❌ Incorreto:
' set alerts to false
.DisplayAlerts = False

' invalidate cache
InvalidarCache()

' loop through tables
For Each nomeTabela In tabelasCriticas
```

**Diretrizes:**
- Explique **POR QUE**, não **O QUE**
- Use português claro e objetivo
- Uma linha antes do código relevante
- Agrupe comentários relacionados

### **2.3 Comentários de Seção**
```vb
✅ Correto:
' ===================================================
' MÉTODOS DE CACHE E PERFORMANCE
' ===================================================

' ===================================================
' CONFIGURAÇÕES DO EXCEL
' ===================================================

' ✅ CORREÇÃO: Aplicar diretamente SEM limpar primeiro
' O problema era: pbProduto.Image = Nothing estava causando a limpeza
```

---

## 🏗️ 3. Estrutura de Classes e Métodos

### **3.1 Ordem de Membros na Classe**
```vb
Public Class ExemploClasse
    ' 1. Constantes
    Private Const TIMEOUT_PADRAO As Integer = 30
    
    ' 2. Campos privados
    Private workbook As Workbook
    Private isDisposed As Boolean
    
    ' 3. Propriedades
    Public Property DadosCarregados As Boolean
    
    ' 4. Construtor
    Public Sub New()
    
    ' 5. Métodos públicos (por ordem de importância)
    Public Function MetodoPrincipal() As Boolean
    Public Function MetodoSecundario() As String
    
    ' 6. Métodos privados
    Private Sub MetodoHelper()
    
    ' 7. Event handlers
    Private Sub Button_Click(sender As Object, e As EventArgs)
    
    ' 8. Dispose/Cleanup
    Public Sub Dispose()
End Class
```

### **3.2 Estrutura de Métodos**
```vb
✅ Padrão Recomendado:
Public Function MetodoExemplo(parametro As String) As Boolean
    Try
        ' 1. Validação de parâmetros
        If String.IsNullOrEmpty(parametro) Then
            Throw New ArgumentException("Parâmetro não pode ser vazio")
        End If
        
        ' 2. Variáveis locais
        Dim resultado As Boolean = False
        Dim dadosProcessados As List(Of String)
        
        ' 3. Log de início (para métodos importantes)
        LogErros.RegistrarInfo($"Iniciando processamento: {parametro}", "MetodoExemplo")
        
        ' 4. Lógica principal
        dadosProcessados = ProcessarDados(parametro)
        resultado = ValidarResultado(dadosProcessados)
        
        ' 5. Log de sucesso
        LogErros.RegistrarInfo("Processamento concluído com sucesso", "MetodoExemplo")
        
        Return resultado
        
    Catch ex As Exception
        ' 6. Tratamento de erro
        LogErros.RegistrarErro(ex, "MetodoExemplo")
        Return False
    End Try
End Function
```

---

## 🔧 4. Tratamento de Erros

### **4.1 Padrão Try-Catch**
```vb
✅ Correto:
Public Function ProcessarDados() As Boolean
    Try
        ' Lógica principal
        Return True
        
    Catch ex As Exception
        LogErros.RegistrarErro(ex, "ProcessarDados")
        Return False
    End Try
End Function

' Para métodos que devem propagar erros:
Public Sub AtualizarDados()
    Try
        ' Lógica principal
        
    Catch ex As Exception
        LogErros.RegistrarErro(ex, "AtualizarDados")
        Throw New Exception($"Erro ao atualizar dados: {ex.Message}", ex)
    End Try
End Sub
```

### **4.2 Validação de Parâmetros**
```vb
✅ Correto:
Public Function ObterTabela(nomeTabela As String) As ListObject
    ' Validação no início do método
    If String.IsNullOrWhiteSpace(nomeTabela) Then
        Throw New ArgumentException("Nome da tabela não pode ser nulo ou vazio", NameOf(nomeTabela))
    End If
    
    If workbook Is Nothing Then
        Throw New InvalidOperationException("Workbook não está disponível")
    End If
    
    ' Resto da lógica...
End Function
```

### **4.3 Logging Padronizado**
```vb
✅ Padrão de Logging:
' Para início de operações importantes
LogErros.RegistrarInfo("Iniciando atualização Power Query", "PowerQueryManager.AtualizarConsultas")

' Para sucesso
LogErros.RegistrarInfo("Consultas atualizadas com sucesso", "PowerQueryManager.AtualizarConsultas")

' Para erros (sempre incluir contexto)
LogErros.RegistrarErro(ex, "PowerQueryManager.AtualizarConsultas")

' Para debug/diagnóstico
LogErros.RegistrarInfo($"Tabela encontrada: {nomeTabela} na planilha {worksheet.Name}", "PowerQueryManager.ObterTabela")
```

---

## ⚡ 5. Performance e Otimização

### **5.1 Cache Pattern**
```vb
✅ Padrão de Cache:
Private tabelasCache As New Dictionary(Of String, ListObject)
Private cacheValido As DateTime = DateTime.MinValue
Private Const CACHE_TIMEOUT_MINUTES As Integer = 5

Public Function ObterTabela(nomeTabela As String) As ListObject
    ' 1. Verificar cache primeiro
    If CacheEstaValido() AndAlso tabelasCache.ContainsKey(nomeTabela) Then
        Return tabelasCache(nomeTabela)
    End If
    
    ' 2. Buscar dados
    Dim tabela = BuscarTabela(nomeTabela)
    
    ' 3. Armazenar no cache
    If tabela IsNot Nothing Then
        tabelasCache(nomeTabela) = tabela
        cacheValido = DateTime.Now
    End If
    
    Return tabela
End Function

Private Function CacheEstaValido() As Boolean
    Return DateTime.Now.Subtract(cacheValido).TotalMinutes < CACHE_TIMEOUT_MINUTES
End Function
```

### **5.2 Otimização de UI**
```vb
✅ Padrão para Updates de UI:
Private Sub AtualizarInterfaceOtimizada()
    ' Suspender redesenho durante updates
    PararRedesenhoCompleto()
    
    Try
        ' Fazer todas as alterações de UI
        dgvProdutos.DataSource = novosDados
        AtualizarLabels()
        
    Finally
        ' Sempre reabilitar redesenho
        ReabilitarRedesenhoCompleto()
    End Try
End Sub
```

---

## 🧪 6. Testes e Validação

### **6.1 Padrão de Testes**
```vb
✅ Estrutura de Teste:
TestFramework.RunTest("Nome Descritivo do Teste", Sub()
    ' Arrange - Preparar dados
    Dim manager As New PowerQueryManager(workbook)
    Dim tabelaTeste = "tblProdutos"
    
    ' Act - Executar ação
    Dim resultado = manager.ObterTabela(tabelaTeste)
    
    ' Assert - Verificar resultado
    TestFramework.AssertNotNull(resultado, "Tabela deve ser encontrada")
    TestFramework.AssertEquals(tabelaTeste, resultado.Name, "Nome da tabela deve coincidir")
End Sub)
```

### **6.2 Validação de Estado**
```vb
✅ Verificações de Estado:
Public Function ObterStatus() As Dictionary(Of String, Object)
    Dim status As New Dictionary(Of String, Object)
    
    status("WorkbookDisponivel") = (workbook IsNot Nothing)
    status("CacheValido") = CacheEstaValido()
    status("TabelasEmCache") = tabelasCache.Count
    status("UltimaAtualizacao") = ultimaAtualizacao
    
    Return status
End Function
```

---

## 📐 7. Formatação e Estilo

### **7.1 Indentação e Espaçamento**
```vb
✅ Correto:
Public Function ExemploFormatacao() As Boolean
    Try
        If condicao1 Then
            For Each item In lista
                If item.IsValid Then
                    ProcessarItem(item)
                End If
            Next
        End If
        
        Return True
        
    Catch ex As Exception
        LogErros.RegistrarErro(ex, "ExemploFormatacao")
        Return False
    End Try
End Function
```

**Regras:**
- **4 espaços** para indentação (não tabs)
- **Linha em branco** entre seções lógicas
- **Alinhamento** de declarações relacionadas
- **Quebra de linha** antes de `Catch`, `Finally`, `End If`, etc.

### **7.2 Declarações de Variáveis**
```vb
✅ Correto:
' Agrupar por tipo e propósito
Dim resultado As Boolean = False
Dim dadosProcessados As List(Of String)
Dim contadorItens As Integer = 0

' Declarar perto do primeiro uso
For Each item In lista
    Dim itemProcessado = ProcessarItem(item)
    If itemProcessado IsNot Nothing Then
        ' usar itemProcessado...
    End If
Next

❌ Incorreto:
' Todas as variáveis no topo sem agrupamento
Dim a As String
Dim b As Integer  
Dim c As Boolean
Dim d As List(Of String)
```

---

## 🎯 8. Padrões Específicos do Projeto

### **8.1 Integração com Excel**
```vb
✅ Padrão para Excel:
' Sempre configurar Excel para operações pesadas
Private Function ConfigurarExcelParaOperacao() As Dictionary(Of String, Object)
    Dim estado As New Dictionary(Of String, Object)
    
    With workbook.Application
        estado("DisplayAlerts") = .DisplayAlerts
        estado("ScreenUpdating") = .ScreenUpdating
        estado("EnableEvents") = .EnableEvents
        
        .DisplayAlerts = False
        .ScreenUpdating = False
        .EnableEvents = False
    End With
    
    Return estado
End Function

' Sempre restaurar configurações
Private Sub RestaurarExcel(estado As Dictionary(Of String, Object))
    ' Implementação de restauração...
End Sub
```

### **8.2 Padrão para UserControls**
```vb
✅ Estrutura de UserControl:
Public Class UcExemplo
    ' Cache e estado
    Private dadosCarregados As Boolean = False
    Private ultimaAtualizacao As DateTime
    
    ' Configuração inicial
    Public Sub New()
        InitializeComponent()
        ConfigurarComponentes()
    End Sub
    
    ' Método principal de carregamento
    Public Function CarregarDados() As Boolean
        ' Implementação...
    End Function
    
    ' Método de limpeza
    Public Sub LimparDados()
        ' Implementação...
    End Sub
End Class
```

### **8.3 Padrão de Configuração**
```vb
✅ Classe de Configuração:
Public Class ConfiguracaoApp
    ' Constantes agrupadas por categoria
    
    ' Caminhos
    Public Const CAMINHO_IMAGENS As String = "C:\ImagesEstoque\"
    Public Const CAMINHO_LOG As String = "C:\Logs\GestaoEstoque\"
    
    ' Timeouts
    Public Const CACHE_TIMEOUT_MINUTES As Integer = 5
    Public Const TIMEOUT_POWERQUERY As Integer = 60
    
    ' Tabelas
    Public Const TABELA_PRODUTOS As String = "tblProdutos"
    Public Const TABELA_ESTOQUE As String = "tblEstoque"
End Class
```

---

## ✅ 9. Checklist de Qualidade

### **Antes de Fazer Commit:**
- [ ] **Nomenclatura**: Segue padrões estabelecidos
- [ ] **Documentação**: Métodos públicos têm XML documentation
- [ ] **Tratamento de erro**: Try-catch apropriado com logging
- [ ] **Performance**: Cache implementado onde necessário
- [ ] **Testes**: Funcionalidade testada
- [ ] **Formatação**: Código bem formatado e legível
- [ ] **Build**: Compila sem warnings
- [ ] **Funcionalidade**: Feature funciona como esperado

### **Code Review Checklist:**
- [ ] **Lógica**: Implementação faz sentido
- [ ] **Segurança**: Validação de parâmetros adequada
- [ ] **Manutenibilidade**: Código fácil de entender e modificar
- [ ] **Padrões**: Segue convenções estabelecidas
- [ ] **Documentação**: Comentários úteis e atualizados

---

## 🎓 10. Exemplos Completos

### **10.1 Classe Completa Exemplo**
```vb
''' <summary>
''' Gerenciador de exemplo seguindo todos os padrões estabelecidos
''' </summary>
Public Class ExemploManager
    ' Constantes
    Private Const TIMEOUT_PADRAO As Integer = 30
    
    ' Campos privados
    Private isInicializado As Boolean = False
    Private dadosCache As Dictionary(Of String, Object)
    
    ''' <summary>
    ''' Inicializa nova instância do ExemploManager
    ''' </summary>
    Public Sub New()
        dadosCache = New Dictionary(Of String, Object)
        isInicializado = True
        LogErros.RegistrarInfo("ExemploManager inicializado", "ExemploManager.New")
    End Sub
    
    ''' <summary>
    ''' Processa dados com validação e cache
    ''' </summary>
    ''' <param name="entrada">Dados a serem processados</param>
    ''' <returns>True se processamento foi bem-sucedido</returns>
    Public Function ProcessarDados(entrada As String) As Boolean
        Try
            ' Validação
            If String.IsNullOrWhiteSpace(entrada) Then
                Throw New ArgumentException("Entrada não pode ser vazia", NameOf(entrada))
            End If
            
            If Not isInicializado Then
                Throw New InvalidOperationException("Manager não foi inicializado")
            End If
            
            ' Log início
            LogErros.RegistrarInfo($"Processando entrada: {entrada}", "ExemploManager.ProcessarDados")
            
            ' Lógica principal
            Dim resultado = ExecutarProcessamento(entrada)
            
            ' Cache resultado
            dadosCache(entrada) = resultado
            
            ' Log sucesso
            LogErros.RegistrarInfo("Processamento concluído com sucesso", "ExemploManager.ProcessarDados")
            
            Return True
            
        Catch ex As Exception
            LogErros.RegistrarErro(ex, "ExemploManager.ProcessarDados")
            Return False
        End Try
    End Function
    
    ''' <summary>
    ''' Executa o processamento principal dos dados
    ''' </summary>
    ''' <param name="entrada">Dados de entrada</param>
    ''' <returns>Resultado do processamento</returns>
    Private Function ExecutarProcessamento(entrada As String) As Object
        ' Implementação específica...
        Return New Object()
    End Function
End Class
```

---

Este guia deve ser seguido por todos os desenvolvedores do projeto EstoqueDifemaqGestao para manter consistência e qualidade do código.
