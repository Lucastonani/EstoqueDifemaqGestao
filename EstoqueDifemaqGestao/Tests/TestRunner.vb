' TestRunner.vb
' Classe para executar testes de forma independente

Imports System.Drawing
Imports System.Windows.Forms

Public Class TestRunner

    ' Executar todos os testes de uma vez
    Public Shared Sub RunAllTestSuites()
        Console.WriteLine("=== EXECUTANDO TODOS OS TESTES ===")
        Console.WriteLine()

        ' Testes unitários
        UnitTests.RunAllTests()

        ' Testes de integração
        IntegrationTests.RunIntegrationTests()

        ' Testes de performance
        PerformanceTests.RunPerformanceTests()

        ' Checklist de deployment
        DeploymentChecklist.RunAllChecks()
    End Sub

    ' Executar testes em modo console
    Public Shared Sub RunConsoleTests()
        Console.WriteLine("=== EXECUTANDO TESTES DO SISTEMA ===")
        Console.WriteLine($"Início: {DateTime.Now:yyyy-MM-dd HH:mm:ss}")
        Console.WriteLine()

        Try
            ' Executar todos os testes
            UnitTests.RunAllTests()

            Console.WriteLine()
            Console.WriteLine("=== TESTES CONCLUÍDOS ===")
            Console.WriteLine("Pressione qualquer tecla para sair...")
            Console.ReadKey()

        Catch ex As Exception
            Console.WriteLine($"ERRO CRÍTICO: {ex.Message}")
            Console.WriteLine(ex.StackTrace)
            Console.ReadKey()
        End Try
    End Sub

    ' Executar testes com interface gráfica
    Public Shared Sub ShowTestUI()
        Dim form As New TestRunnerForm()
        Application.Run(form)
    End Sub

End Class

' Formulário para execução visual dos testes
Public Class TestRunnerForm
    Inherits Form

    Private WithEvents btnRunTests As Button
    Private WithEvents txtResults As TextBox
    Private WithEvents progressBar As ProgressBar
    Private lblStatus As Label

    Public Sub New()
        InitializeComponent()
    End Sub

    Private Sub InitializeComponent()
        Me.Text = "Executor de Testes - Gestão Estoque Difemaq"
        Me.Size = New Size(800, 600)
        Me.StartPosition = FormStartPosition.CenterScreen

        ' Botão executar testes unitários
        btnRunTests = New Button()
        btnRunTests.Text = "Executar Testes Unitários"
        btnRunTests.Location = New Point(10, 10)
        btnRunTests.Size = New Size(180, 30)
        btnRunTests.BackColor = Color.FromArgb(46, 134, 171)
        btnRunTests.ForeColor = Color.White
        btnRunTests.FlatStyle = FlatStyle.Flat
        btnRunTests.Font = New Font("Segoe UI", 10, FontStyle.Bold)

        ' Botão executar todos os testes
        Dim btnRunAllTests As New Button()
        btnRunAllTests.Text = "Executar TODOS os Testes"
        btnRunAllTests.Location = New Point(200, 10)
        btnRunAllTests.Size = New Size(180, 30)
        btnRunAllTests.BackColor = Color.FromArgb(139, 69, 19)
        btnRunAllTests.ForeColor = Color.White
        btnRunAllTests.FlatStyle = FlatStyle.Flat
        btnRunAllTests.Font = New Font("Segoe UI", 10, FontStyle.Bold)
        AddHandler btnRunAllTests.Click, AddressOf btnRunAllTests_Click

        ' Label status
        lblStatus = New Label()
        lblStatus.Text = "Pronto para executar testes"
        lblStatus.Location = New Point(220, 15)
        lblStatus.Size = New Size(300, 20)
        lblStatus.Font = New Font("Segoe UI", 9)

        ' Progress bar
        progressBar = New ProgressBar()
        progressBar.Location = New Point(10, 50)
        progressBar.Size = New Size(760, 20)
        progressBar.Style = ProgressBarStyle.Marquee
        progressBar.Visible = False

        ' TextBox resultados
        txtResults = New TextBox()
        txtResults.Location = New Point(10, 80)
        txtResults.Size = New Size(760, 470)
        txtResults.Multiline = True
        txtResults.ScrollBars = ScrollBars.Both
        txtResults.Font = New Font("Consolas", 9)
        txtResults.BackColor = Color.Black
        txtResults.ForeColor = Color.LightGreen
        txtResults.ReadOnly = True

        ' Adicionar controles
        Me.Controls.AddRange({btnRunTests, btnRunAllTests, lblStatus, progressBar, txtResults})
    End Sub

    Private Sub btnRunTests_Click(sender As Object, e As EventArgs) Handles btnRunTests.Click
        btnRunTests.Enabled = False
        progressBar.Visible = True
        lblStatus.Text = "Executando testes..."
        txtResults.Clear()

        ' Redirecionar console output
        Dim sw As New System.IO.StringWriter()
        Console.SetOut(sw)

        ' Executar em thread separada
        Dim testThread As New Threading.Thread(
            Sub()
                Try
                    UnitTests.RunAllTests()

                    Me.Invoke(Sub()
                                  txtResults.Text = sw.ToString()
                                  lblStatus.Text = "Testes concluídos!"
                                  progressBar.Visible = False
                                  btnRunTests.Enabled = True

                                  ' Rolar para o final
                                  txtResults.SelectionStart = txtResults.Text.Length
                                  txtResults.ScrollToCaret()
                              End Sub)

                Catch ex As Exception
                    Me.Invoke(Sub()
                                  txtResults.Text = $"ERRO: {ex.Message}{Environment.NewLine}{ex.StackTrace}"
                                  lblStatus.Text = "Erro durante execução!"
                                  progressBar.Visible = False
                                  btnRunTests.Enabled = True
                              End Sub)
                End Try

                ' Restaurar console
                Console.SetOut(Console.Out)
            End Sub
        )

        testThread.Start()
    End Sub

    Private Sub btnRunAllTests_Click(sender As Object, e As EventArgs)
        Dim btn As Button = CType(sender, Button)
        btn.Enabled = False
        progressBar.Visible = True
        lblStatus.Text = "Executando TODOS os testes..."
        txtResults.Clear()

        ' Redirecionar console output
        Dim sw As New System.IO.StringWriter()
        Console.SetOut(sw)

        ' Executar em thread separada
        Dim testThread As New Threading.Thread(
            Sub()
                Try
                    TestRunner.RunAllTestSuites()

                    Me.Invoke(Sub()
                                  txtResults.Text = sw.ToString()
                                  lblStatus.Text = "Todos os testes concluídos!"
                                  progressBar.Visible = False
                                  btn.Enabled = True

                                  ' Rolar para o final
                                  txtResults.SelectionStart = txtResults.Text.Length
                                  txtResults.ScrollToCaret()
                              End Sub)

                Catch ex As Exception
                    Me.Invoke(Sub()
                                  txtResults.Text = $"ERRO: {ex.Message}{Environment.NewLine}{ex.StackTrace}"
                                  lblStatus.Text = "Erro durante execução!"
                                  progressBar.Visible = False
                                  btn.Enabled = True
                              End Sub)
                End Try

                ' Restaurar console
                Console.SetOut(Console.Out)
            End Sub
        )

        testThread.Start()
    End Sub

End Class

' Módulo para executar testes via linha de comando
Module TestProgram

    Sub Main(args As String())
        If args.Length > 0 AndAlso args(0).ToLower() = "--gui" Then
            ' Executar com interface gráfica
            Application.EnableVisualStyles()
            Application.SetCompatibleTextRenderingDefault(False)
            TestRunner.ShowTestUI()
        Else
            ' Executar em modo console
            TestRunner.RunConsoleTests()
        End If
    End Sub

End Module