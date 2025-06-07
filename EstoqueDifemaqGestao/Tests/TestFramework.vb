' TestFramework.vb
' Framework simples de testes para o projeto

Imports System.Text
Imports System.IO
Imports System.Diagnostics

Public Class TestFramework

    Private Shared testResults As New List(Of TestResult)
    Private Shared currentTestSuite As String = ""

    Public Class TestResult
        Public Property SuiteName As String
        Public Property TestName As String
        Public Property Passed As Boolean
        Public Property ErrorMessage As String
        Public Property ExecutionTime As TimeSpan
        Public Property Timestamp As DateTime
    End Class

    ' Iniciar uma suite de testes
    Public Shared Sub BeginTestSuite(suiteName As String)
        currentTestSuite = suiteName
        Console.WriteLine($"=== Iniciando Suite de Testes: {suiteName} ===")
    End Sub

    ' Executar um teste
    Public Shared Sub RunTest(testName As String, testAction As System.Action)
        Dim sw As New Stopwatch()
        Dim result As New TestResult With {
            .SuiteName = currentTestSuite,
            .TestName = testName,
            .Timestamp = DateTime.Now
        }

        Try
            sw.Start()
            testAction.Invoke()
            sw.Stop()

            result.Passed = True
            result.ExecutionTime = sw.Elapsed
            Console.WriteLine($"✓ {testName} - PASSOU ({sw.ElapsedMilliseconds}ms)")

        Catch ex As Exception
            sw.Stop()
            result.Passed = False
            result.ErrorMessage = ex.Message
            result.ExecutionTime = sw.Elapsed
            Console.WriteLine($"✗ {testName} - FALHOU: {ex.Message}")
        End Try

        testResults.Add(result)
    End Sub

    ' Assertions
    Public Shared Sub AssertTrue(condition As Boolean, Optional message As String = "")
        If Not condition Then
            Throw New Exception($"Assertion failed: Expected True. {message}")
        End If
    End Sub

    Public Shared Sub AssertFalse(condition As Boolean, Optional message As String = "")
        If condition Then
            Throw New Exception($"Assertion failed: Expected False. {message}")
        End If
    End Sub

    Public Shared Sub AssertEquals(expected As Object, actual As Object, Optional message As String = "")
        If Not Object.Equals(expected, actual) Then
            Throw New Exception($"Assertion failed: Expected [{expected}] but got [{actual}]. {message}")
        End If
    End Sub

    Public Shared Sub AssertNotNull(obj As Object, Optional message As String = "")
        If obj Is Nothing Then
            Throw New Exception($"Assertion failed: Expected not null. {message}")
        End If
    End Sub

    Public Shared Sub AssertNull(obj As Object, Optional message As String = "")
        If obj IsNot Nothing Then
            Throw New Exception($"Assertion failed: Expected null. {message}")
        End If
    End Sub

    Public Shared Sub AssertThrows(Of TException As Exception)(action As System.Action, Optional message As String = "")
        Try
            action.Invoke()
            Throw New Exception($"Assertion failed: Expected exception of type {GetType(TException).Name} but no exception was thrown. {message}")
        Catch ex As TException
            ' Sucesso - a exceção esperada foi lançada
        Catch ex As Exception
            Throw New Exception($"Assertion failed: Expected exception of type {GetType(TException).Name} but got {ex.GetType().Name}. {message}")
        End Try
    End Sub

    ' Gerar relatório
    Public Shared Function GenerateReport() As String
        Dim report As New StringBuilder()
        Dim totalTests = testResults.Count
        Dim passedTests = testResults.Where(Function(r) r.Passed).Count()
        Dim failedTests = totalTests - passedTests
        Dim totalTime = TimeSpan.FromMilliseconds(testResults.Sum(Function(r) r.ExecutionTime.TotalMilliseconds))

        report.AppendLine("=== RELATÓRIO DE TESTES ===")
        report.AppendLine($"Data/Hora: {DateTime.Now:yyyy-MM-dd HH:mm:ss}")
        report.AppendLine($"Total de Testes: {totalTests}")
        report.AppendLine($"Passou: {passedTests}")
        report.AppendLine($"Falhou: {failedTests}")
        report.AppendLine($"Taxa de Sucesso: {If(totalTests > 0, (passedTests * 100.0 / totalTests).ToString("F2"), "0")}%")
        report.AppendLine($"Tempo Total: {totalTime.TotalSeconds:F2}s")
        report.AppendLine()

        ' Agrupar por suite
        Dim suites = testResults.GroupBy(Function(r) r.SuiteName)

        For Each suite In suites
            report.AppendLine($"Suite: {suite.Key}")

            For Each test In suite
                Dim status = If(test.Passed, "PASSOU", "FALHOU")
                report.AppendLine($"  [{status}] {test.TestName} ({test.ExecutionTime.TotalMilliseconds}ms)")

                If Not test.Passed Then
                    report.AppendLine($"    Erro: {test.ErrorMessage}")
                End If
            Next

            report.AppendLine()
        Next

        Return report.ToString()
    End Function

    ' Salvar relatório em arquivo
    Public Shared Sub SaveReport(filePath As String)
        Try
            File.WriteAllText(filePath, GenerateReport())
            Console.WriteLine($"Relatório salvo em: {filePath}")
        Catch ex As Exception
            Console.WriteLine($"Erro ao salvar relatório: {ex.Message}")
        End Try
    End Sub

    ' Limpar resultados
    Public Shared Sub ClearResults()
        testResults.Clear()
        currentTestSuite = ""
    End Sub

End Class