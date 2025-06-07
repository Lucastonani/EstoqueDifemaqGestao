Imports System.IO
Imports System.Text

Public Class LogErros
    Private Shared ReadOnly lockObject As New Object()
    Private Shared caminhoArquivoLog As String = ""

    Shared Sub New()
        Try
            ' Criar diretório de log se não existir
            If Not Directory.Exists(ConfiguracaoApp.CAMINHO_LOG) Then
                Directory.CreateDirectory(ConfiguracaoApp.CAMINHO_LOG)
            End If

            ' Definir nome do arquivo de log
            Dim nomeArquivo As String = $"GestaoEstoque_{DateTime.Now:yyyyMMdd}.log"
            caminhoArquivoLog = Path.Combine(ConfiguracaoApp.CAMINHO_LOG, nomeArquivo)

        Catch
            ' Se falhar ao criar o log, usar pasta temporária
            caminhoArquivoLog = Path.Combine(Path.GetTempPath(), "GestaoEstoque.log")
        End Try
    End Sub

    Public Shared Sub RegistrarErro(ex As Exception, contexto As String)
        Try
            SyncLock lockObject
                Dim logEntry As New StringBuilder()
                logEntry.AppendLine($"=== ERRO REGISTRADO ===")
                logEntry.AppendLine($"Data/Hora: {DateTime.Now:yyyy-MM-dd HH:mm:ss}")
                logEntry.AppendLine($"Contexto: {contexto}")
                logEntry.AppendLine($"Tipo da Exceção: {ex.GetType().Name}")
                logEntry.AppendLine($"Mensagem: {ex.Message}")

                If ex.InnerException IsNot Nothing Then
                    logEntry.AppendLine($"Exceção Interna: {ex.InnerException.Message}")
                End If

                logEntry.AppendLine($"Stack Trace: {ex.StackTrace}")
                logEntry.AppendLine($"========================")
                logEntry.AppendLine()

                ' Verificar tamanho do arquivo antes de escrever
                VerificarTamanhoArquivo()

                File.AppendAllText(caminhoArquivoLog, logEntry.ToString(), Encoding.UTF8)
            End SyncLock

        Catch
            ' Se falhar ao registrar no arquivo, não fazer nada para evitar loop infinito
        End Try
    End Sub

    Public Shared Sub RegistrarInfo(mensagem As String, contexto As String)
        Try
            SyncLock lockObject
                Dim logEntry As String = $"[INFO] {DateTime.Now:yyyy-MM-dd HH:mm:ss} - {contexto}: {mensagem}{Environment.NewLine}"

                VerificarTamanhoArquivo()
                File.AppendAllText(caminhoArquivoLog, logEntry, Encoding.UTF8)
            End SyncLock

        Catch
            ' Se falhar ao registrar no arquivo, não fazer nada
        End Try
    End Sub

    Private Shared Sub VerificarTamanhoArquivo()
        Try
            If File.Exists(caminhoArquivoLog) Then
                Dim tamanhoAtual As Long = New FileInfo(caminhoArquivoLog).Length

                If tamanhoAtual > ConfiguracaoApp.TAMANHO_MAX_LOG Then
                    ' Criar backup do arquivo atual
                    Dim nomeBackup As String = Path.ChangeExtension(caminhoArquivoLog, $".backup_{DateTime.Now:yyyyMMdd_HHmmss}.log")
                    File.Move(caminhoArquivoLog, nomeBackup)

                    ' Criar novo arquivo
                    File.WriteAllText(caminhoArquivoLog, $"=== LOG INICIADO EM {DateTime.Now:yyyy-MM-dd HH:mm:ss} ==={Environment.NewLine}", Encoding.UTF8)
                End If
            End If

        Catch
            ' Se falhar, continuar normalmente
        End Try
    End Sub

    Public Shared Function ObterCaminhoLog() As String
        Return caminhoArquivoLog
    End Function

End Class