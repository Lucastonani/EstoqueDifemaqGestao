Imports System.Drawing
Imports System.Drawing.Drawing2D
Imports System.Windows.Forms
Imports System.Threading

Public Class SplashForm
    Private isClosing As Boolean = False

    Public Sub New()
        Try
            InitializeComponent()
            ConfigurarFormulario()
        Catch ex As Exception
            ' Se falhar na criação, continuar sem splash
            Try
                LogErros.RegistrarErro(ex, "SplashForm.New")
            Catch
                ' Ignorar se LogErros falhar
            End Try
        End Try
    End Sub

    Private Sub ConfigurarFormulario()
        Try
            ' Configurações básicas seguras
            Me.SetStyle(ControlStyles.AllPaintingInWmPaint Or ControlStyles.UserPaint Or ControlStyles.DoubleBuffer, True)
            Me.BackColor = Color.White
            Me.Opacity = 0.9 ' Começar quase opaco

            ' Aplicar bordas arredondadas (se possível)
            Try
                AplicarBordasArredondadas()
            Catch
                ' Se falhar, continuar sem bordas arredondadas
            End Try

            ' Configurar progress bar
            progressBar.Style = ProgressBarStyle.Continuous
            progressBar.Value = 0
            progressBar.Minimum = 0
            progressBar.Maximum = 100

            ' Fade in simples
            FadeIn()

        Catch ex As Exception
            Try
                LogErros.RegistrarErro(ex, "SplashForm.ConfigurarFormulario")
            Catch
                ' Ignorar se LogErros falhar
            End Try
        End Try
    End Sub

    Private Sub AplicarBordasArredondadas()
        Try
            Dim radius As Integer = 12
            Dim path As New GraphicsPath()

            path.AddArc(0, 0, radius, radius, 180, 90)
            path.AddArc(Me.Width - radius, 0, radius, radius, 270, 90)
            path.AddArc(Me.Width - radius, Me.Height - radius, radius, radius, 0, 90)
            path.AddArc(0, Me.Height - radius, radius, radius, 90, 90)
            path.CloseAllFigures()

            Me.Region = New Region(path)
        Catch
            ' Ignorar se não conseguir aplicar bordas arredondadas
        End Try
    End Sub

    Private Sub FadeIn()
        Try
            Me.Opacity = 0
            Dim fadeTimer As New System.Windows.Forms.Timer()
            fadeTimer.Interval = 30

            AddHandler fadeTimer.Tick, Sub()
                                           Try
                                               If Me.Opacity < 0.95 Then
                                                   Me.Opacity += 0.05
                                               Else
                                                   Me.Opacity = 1.0
                                                   fadeTimer.Stop()
                                                   fadeTimer.Dispose()
                                               End If
                                           Catch
                                               fadeTimer.Stop()
                                               Me.Opacity = 1.0
                                           End Try
                                       End Sub

            fadeTimer.Start()
        Catch
            Me.Opacity = 1.0
        End Try
    End Sub

    ' Método principal para atualizar status
    Public Sub AtualizarStatus(mensagem As String, Optional progresso As Integer = -1)
        Try
            If Me.IsDisposed OrElse isClosing Then Return

            If Me.InvokeRequired Then
                Me.Invoke(New Action(Of String, Integer)(AddressOf AtualizarStatusSeguro), mensagem, progresso)
            Else
                AtualizarStatusSeguro(mensagem, progresso)
            End If

        Catch ex As Exception
            ' Se falhar, tentar atualização direta
            Try
                AtualizarStatusSeguro(mensagem, progresso)
            Catch
                ' Ignorar se não conseguir atualizar
            End Try
        End Try
    End Sub

    Private Sub AtualizarStatusSeguro(mensagem As String, progresso As Integer)
        Try
            ' Atualizar mensagem
            If lblStatus IsNot Nothing Then
                lblStatus.Text = mensagem
            End If

            ' Atualizar progresso se especificado
            If progresso >= 0 AndAlso progresso <= 100 AndAlso progressBar IsNot Nothing Then
                progressBar.Value = progresso
            End If

            ' Forçar atualização
            Me.Refresh()
            Application.DoEvents()

        Catch ex As Exception
            Try
                LogErros.RegistrarErro(ex, "SplashForm.AtualizarStatusSeguro")
            Catch
                ' Ignorar se LogErros falhar
            End Try
        End Try
    End Sub

    ' Método para definir versão
    Public Sub DefinirVersao(versao As String)
        Try
            If Not Me.IsDisposed AndAlso lblVersion IsNot Nothing Then
                lblVersion.Text = $"Versão {versao}"
            End If
        Catch
            ' Ignorar erro
        End Try
    End Sub

    ' Método simplificado para fechar
    Public Sub FecharSuave()
        Try
            If Me.IsDisposed OrElse isClosing Then Return

            isClosing = True

            ' Fade out rápido
            Dim fadeTimer As New System.Windows.Forms.Timer()
            fadeTimer.Interval = 20

            AddHandler fadeTimer.Tick, Sub()
                                           Try
                                               If Me.Opacity > 0.1 Then
                                                   Me.Opacity -= 0.1
                                               Else
                                                   fadeTimer.Stop()
                                                   fadeTimer.Dispose()
                                                   Me.Close()
                                               End If
                                           Catch
                                               fadeTimer.Stop()
                                               Me.Close()
                                           End Try
                                       End Sub

            fadeTimer.Start()

        Catch
            Try
                Me.Close()
            Catch
                ' Ignorar erro final
            End Try
        End Try
    End Sub

    ' Evitar fechamento pelo usuário
    Protected Overrides Sub SetVisibleCore(value As Boolean)
        Try
            If Not isClosing AndAlso Not value Then
                Return
            End If
            MyBase.SetVisibleCore(value)
        Catch
            MyBase.SetVisibleCore(value)
        End Try
    End Sub

    Private Sub SplashForm_FormClosing(sender As Object, e As FormClosingEventArgs) Handles Me.FormClosing
        Try
            If e.CloseReason = CloseReason.UserClosing AndAlso Not isClosing Then
                e.Cancel = True
            End If
        Catch
            ' Ignorar erro
        End Try
    End Sub

End Class