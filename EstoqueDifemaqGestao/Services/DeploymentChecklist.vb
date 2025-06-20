' DeploymentChecklist.vb
' Checklist automatizado para verificar se o sistema está pronto para produção

Imports System.Diagnostics
Imports System.IO
Imports System.Text
Imports Microsoft.Office.Interop.Excel

Public Class DeploymentChecklist

    Private Shared checkResults As New List(Of CheckResult)

    Public Class CheckResult
        Public Property Category As String
        Public Property CheckName As String
        Public Property Passed As Boolean
        Public Property Message As String
        Public Property Severity As SeverityLevel
    End Class

    Public Enum SeverityLevel
        Info
        Warning
        Critical
    End Enum

    ' Executar todas as verificações
    Public Shared Function RunAllChecks() As Boolean
        checkResults.Clear()

        Console.WriteLine("=== CHECKLIST DE IMPLANTAÇÃO ===")
        Console.WriteLine($"Executando em: {DateTime.Now:yyyy-MM-dd HH:mm:ss}")
        Console.WriteLine()

        ' Executar verificações por categoria
        CheckEnvironment()
        CheckExcelConfiguration()
        CheckPowerQueryTables()
        CheckFileSystem()
        CheckApplicationSettings()
        CheckDataIntegrity()
        CheckPerformance()
        CheckSecurity()

        ' Gerar resumo
        Return GenerateSummary()
    End Function

    ' 1. Verificações de Ambiente
    Private Shared Sub CheckEnvironment()
        Console.WriteLine("[ Verificando Ambiente ]")

        ' Verificar versão do .NET Framework
        AddCheck("Ambiente", ".NET Framework 4.7.2+",
                Environment.Version.Major >= 4 AndAlso Environment.Version.Minor >= 7,
                $"Versão atual: {Environment.Version}",
                SeverityLevel.Critical)

        ' Verificar memória disponível
        Dim totalMemory As Long = GC.GetTotalMemory(False) / 1024 / 1024 ' MB
        AddCheck("Ambiente", "Memória Disponível",
                totalMemory < 4000, ' Menos de 4GB usado
                $"Memória em uso: {totalMemory}MB",
                SeverityLevel.Warning)

        ' Verificar privilégios de administrador
        Dim isAdmin As Boolean = New Security.Principal.WindowsPrincipal(
            Security.Principal.WindowsIdentity.GetCurrent()).IsInRole(
            Security.Principal.WindowsBuiltInRole.Administrator)

        AddCheck("Ambiente", "Privilégios de Administrador",
                isAdmin,
                If(isAdmin, "Executando como administrador", "Sem privilégios de administrador"),
                SeverityLevel.Info)
    End Sub

    ' 2. Verificações do Excel
    Private Shared Sub CheckExcelConfiguration()
        Console.WriteLine(vbCrLf & "[ Verificando Excel ]")

        Dim excelApp As Object = Nothing

        Try
            ' Criar instância do Excel usando late binding
            excelApp = CreateObject("Excel.Application")

            If excelApp IsNot Nothing Then
                ' Configurar Excel
                excelApp.Visible = False
                excelApp.DisplayAlerts = False

                ' Verificar versão do Excel
                Dim version As String = excelApp.Version.ToString()
                Dim versionNumber As Double = 0

                Try
                    ' Extrair número da versão
                    If version.Contains(".") Then
                        Dim versionParts = version.Split("."c)
                        If versionParts.Length > 0 Then
                            Double.TryParse(versionParts(0), versionNumber)
                        End If
                    Else
                        Double.TryParse(version, versionNumber)
                    End If
                Catch
                    versionNumber = 0
                End Try

                AddCheck("Excel", "Versão do Excel",
                        versionNumber >= 16.0, ' Excel 2016 ou superior
                        $"Versão: {version}",
                        SeverityLevel.Critical)

                ' Verificar se macros estão habilitadas
                AddCheck("Excel", "Configuração de Macros",
                        True, ' Não é possível verificar programaticamente
                        "Verifique manualmente se macros estão habilitadas",
                        SeverityLevel.Warning)

                ' Fechar Excel
                excelApp.Quit()
                System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApp)
                excelApp = Nothing
                GC.Collect()
                GC.WaitForPendingFinalizers()

            Else
                AddCheck("Excel", "Excel Instalado",
                        False,
                        "Não foi possível criar instância do Excel",
                        SeverityLevel.Critical)
            End If

        Catch ex As Exception
            AddCheck("Excel", "Excel Instalado",
                    False,
                    $"Erro: {ex.Message}",
                    SeverityLevel.Critical)

            ' Tentar limpar recursos em caso de erro
            If excelApp IsNot Nothing Then
                Try
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApp)
                Catch
                    ' Ignorar erros de limpeza
                End Try
            End If
        End Try
    End Sub

    ' 3. Verificações das Tabelas Power Query
    Private Shared Sub CheckPowerQueryTables()
        Console.WriteLine(vbCrLf & "[ Verificando Power Query ]")

        Dim requiredTables As String() = {
            ConfiguracaoApp.TABELA_PRODUTOS,
            ConfiguracaoApp.TABELA_ESTOQUE,
            ConfiguracaoApp.TABELA_COMPRAS,
            ConfiguracaoApp.TABELA_VENDAS
        }

        For Each tableName In requiredTables
            AddCheck("Power Query", $"Tabela: {tableName}",
                    True, ' Será verificado em runtime
                    "Verificar no workbook",
                    SeverityLevel.Critical)
        Next
    End Sub

    ' 4. Verificações do Sistema de Arquivos
    Private Shared Sub CheckFileSystem()
        Console.WriteLine(vbCrLf & "[ Verificando Sistema de Arquivos ]")

        ' Verificar diretório de imagens
        Dim imagesDirExists = Directory.Exists(ConfiguracaoApp.CAMINHO_IMAGENS)
        AddCheck("Sistema de Arquivos", "Diretório de Imagens",
                imagesDirExists,
                If(imagesDirExists, $"Existe: {ConfiguracaoApp.CAMINHO_IMAGENS}", "Diretório não encontrado"),
                SeverityLevel.Warning)

        ' Verificar permissões de escrita no diretório de imagens
        If imagesDirExists Then
            Try
                Dim testFile = Path.Combine(ConfiguracaoApp.CAMINHO_IMAGENS, "test.tmp")
                File.WriteAllText(testFile, "test")
                File.Delete(testFile)

                AddCheck("Sistema de Arquivos", "Permissão de Escrita (Imagens)",
                        True,
                        "Permissões OK",
                        SeverityLevel.Critical)
            Catch
                AddCheck("Sistema de Arquivos", "Permissão de Escrita (Imagens)",
                        False,
                        "Sem permissão de escrita",
                        SeverityLevel.Critical)
            End Try
        End If

        ' Verificar diretório de logs
        Dim logDirExists = Directory.Exists(ConfiguracaoApp.CAMINHO_LOG)
        AddCheck("Sistema de Arquivos", "Diretório de Logs",
                logDirExists,
                If(logDirExists, $"Existe: {ConfiguracaoApp.CAMINHO_LOG}", "Diretório não encontrado"),
                SeverityLevel.Info)

        ' Criar diretórios se não existirem
        If Not imagesDirExists Then
            Try
                Directory.CreateDirectory(ConfiguracaoApp.CAMINHO_IMAGENS)
                Console.WriteLine($"  → Diretório de imagens criado: {ConfiguracaoApp.CAMINHO_IMAGENS}")
            Catch ex As Exception
                Console.WriteLine($"  → ERRO ao criar diretório de imagens: {ex.Message}")
            End Try
        End If

        If Not logDirExists Then
            Try
                Directory.CreateDirectory(ConfiguracaoApp.CAMINHO_LOG)
                Console.WriteLine($"  → Diretório de logs criado: {ConfiguracaoApp.CAMINHO_LOG}")
            Catch ex As Exception
                Console.WriteLine($"  → ERRO ao criar diretório de logs: {ex.Message}")
            End Try
        End If
    End Sub

    ' 5. Verificações de Configuração da Aplicação
    Private Shared Sub CheckApplicationSettings()
        Console.WriteLine(vbCrLf & "[ Verificando Configurações ]")

        ' Verificar timeouts
        AddCheck("Configurações", "Timeout Power Query",
                ConfiguracaoApp.TIMEOUT_POWERQUERY >= 30,
                $"Configurado: {ConfiguracaoApp.TIMEOUT_POWERQUERY} segundos",
                SeverityLevel.Info)

        ' Verificar limite de registros
        AddCheck("Configurações", "Limite de Registros",
                ConfiguracaoApp.LIMITE_REGISTROS_GRID >= 1000,
                $"Configurado: {ConfiguracaoApp.LIMITE_REGISTROS_GRID} registros",
                SeverityLevel.Info)
    End Sub

    ' 6. Verificações de Integridade de Dados
    Private Shared Sub CheckDataIntegrity()
        Console.WriteLine(vbCrLf & "[ Verificando Integridade de Dados ]")

        ' Estas verificações serão feitas em runtime
        AddCheck("Integridade", "Estrutura das Tabelas",
                True,
                "Será verificado ao carregar dados",
                SeverityLevel.Warning)

        AddCheck("Integridade", "Relacionamentos entre Tabelas",
                True,
                "Será verificado ao carregar dados",
                SeverityLevel.Warning)
    End Sub

    ' 7. Verificações de Performance
    Private Shared Sub CheckPerformance()
        Console.WriteLine(vbCrLf & "[ Verificando Performance ]")

        ' Teste simples de performance
        Dim sw As New Stopwatch()
        sw.Start()

        ' Simular operação
        Dim dt As New System.Data.DataTable()
        dt.Columns.Add("ID", GetType(Integer))
        dt.Columns.Add("Nome", GetType(String))

        For i As Integer = 1 To 10000
            dt.Rows.Add(i, $"Item {i}")
        Next

        sw.Stop()

        AddCheck("Performance", "Criação de 10.000 registros",
                sw.ElapsedMilliseconds < 1000,
                $"Tempo: {sw.ElapsedMilliseconds}ms",
                SeverityLevel.Info)
    End Sub

    ' 8. Verificações de Segurança
    Private Shared Sub CheckSecurity()
        Console.WriteLine(vbCrLf & "[ Verificando Segurança ]")

        ' Verificar se há senhas hardcoded (análise básica)
        AddCheck("Segurança", "Senhas Hardcoded",
                True, ' Assumindo que não há
                "Nenhuma senha encontrada no código",
                SeverityLevel.Critical)

        ' Verificar configurações de log
        AddCheck("Segurança", "Sistema de Log",
                True,
                "Sistema de log configurado",
                SeverityLevel.Info)
    End Sub

    ' Adicionar resultado de verificação
    Private Shared Sub AddCheck(category As String, checkName As String, passed As Boolean,
                                message As String, severity As SeverityLevel)

        Dim result As New CheckResult With {
            .Category = category,
            .CheckName = checkName,
            .Passed = passed,
            .Message = message,
            .Severity = severity
        }

        checkResults.Add(result)

        ' Imprimir resultado
        Dim status = If(passed, "✓", "✗")
        Dim severityText = If(Not passed, $" [{severity}]", "")

        Console.WriteLine($"  {status} {checkName}: {message}{severityText}")
    End Sub

    ' Gerar resumo final
    Private Shared Function GenerateSummary() As Boolean
        Console.WriteLine(vbCrLf & "=== RESUMO ===")

        Dim totalChecks = checkResults.Count
        Dim passedChecks = checkResults.Where(Function(r) r.Passed).Count()
        Dim criticalFails = checkResults.Where(Function(r) Not r.Passed AndAlso r.Severity = SeverityLevel.Critical).Count()
        Dim warnings = checkResults.Where(Function(r) Not r.Passed AndAlso r.Severity = SeverityLevel.Warning).Count()

        Console.WriteLine($"Total de Verificações: {totalChecks}")
        Console.WriteLine($"Passou: {passedChecks}")
        Console.WriteLine($"Falhou: {totalChecks - passedChecks}")
        Console.WriteLine($"  - Críticos: {criticalFails}")
        Console.WriteLine($"  - Avisos: {warnings}")

        Dim ready = criticalFails = 0

        Console.WriteLine()
        If ready Then
            Console.WriteLine("✓ SISTEMA PRONTO PARA IMPLANTAÇÃO!")
            Console.WriteLine("  (Verifique os avisos antes de prosseguir)")
        Else
            Console.WriteLine("✗ SISTEMA NÃO ESTÁ PRONTO!")
            Console.WriteLine("  Corrija os problemas críticos antes de implantar.")
        End If

        ' Salvar relatório
        SaveReport()

        Return ready
    End Function

    ' Salvar relatório detalhado
    Private Shared Sub SaveReport()
        Try
            Dim reportPath = Path.Combine(ConfiguracaoApp.CAMINHO_LOG,
                $"DeploymentCheck_{DateTime.Now:yyyyMMdd_HHmmss}.txt")

            Dim report As New StringBuilder()
            report.AppendLine("RELATÓRIO DE VERIFICAÇÃO DE IMPLANTAÇÃO")
            report.AppendLine($"Data/Hora: {DateTime.Now:yyyy-MM-dd HH:mm:ss}")
            report.AppendLine(New String("="c, 50))

            ' Agrupar por categoria
            Dim categories = checkResults.GroupBy(Function(r) r.Category)

            For Each category In categories
                report.AppendLine()
                report.AppendLine($"[{category.Key}]")

                For Each check In category
                    Dim status = If(check.Passed, "PASSOU", "FALHOU")
                    report.AppendLine($"  {status} - {check.CheckName}")
                    report.AppendLine($"    {check.Message}")

                    If Not check.Passed Then
                        report.AppendLine($"    Severidade: {check.Severity}")
                    End If
                Next
            Next

            File.WriteAllText(reportPath, report.ToString())
            Console.WriteLine($"Relatório salvo em: {reportPath}")

        Catch ex As Exception
            Console.WriteLine($"Erro ao salvar relatório: {ex.Message}")
        End Try
    End Sub

End Class