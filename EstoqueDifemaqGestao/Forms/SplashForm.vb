Imports System.Drawing
Imports System.Drawing.Drawing2D
Imports System.Windows.Forms
Imports System.Threading

Public Class SplashForm
    Private isClosing As Boolean = False
    Private fadeTimer As System.Windows.Forms.Timer
    Private Const FADE_STEP As Double = 0.05
    Private Const FADE_INTERVAL As Integer = 5

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
            Me.Opacity = 0.0 ' Começar transparente para fade in

            ' Aplicar bordas arredondadas (se possível)
            Try
                AplicarBordasArredondadas()
            Catch
                ' Se falhar, continuar sem bordas arredondadas
            End Try

            ' Configurar progress bar
            If progressBar IsNot Nothing Then
                progressBar.Style = ProgressBarStyle.Continuous
                progressBar.Value = 0
                progressBar.Minimum = 0
                progressBar.Maximum = 100
            End If

            ' Fade in suave
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

            ' Reutilizar timer se possível ou criar novo
            If fadeTimer Is Nothing Then
                fadeTimer = New System.Windows.Forms.Timer()
            End If

            fadeTimer.Interval = FADE_INTERVAL

            ' Remover handlers anteriores
            RemoveHandler fadeTimer.Tick, AddressOf FadeInTick
            AddHandler fadeTimer.Tick, AddressOf FadeInTick

            fadeTimer.Start()
        Catch
            Me.Opacity = 1.0
        End Try
    End Sub

    Private Sub FadeInTick(sender As Object, e As EventArgs)
        Try
            If Me.Opacity < (1.0 - FADE_STEP) Then
                Me.Opacity += FADE_STEP
            Else
                Me.Opacity = 1.0
                fadeTimer.Stop()
                RemoveHandler fadeTimer.Tick, AddressOf FadeInTick
            End If
        Catch
            fadeTimer.Stop()
            Me.Opacity = 1.0
        End Try
    End Sub

    ' Método principal para atualizar status - MELHORADO
    Public Sub AtualizarStatus(mensagem As String, Optional progresso As Integer = -1)
        Try
            If Me.IsDisposed OrElse isClosing Then Return

            If Me.InvokeRequired Then
                ' Especificar System.Action para evitar conflito com Excel.Action
                Dim updateAction As New System.Action(Sub() AtualizarStatusSeguro(mensagem, progresso))
                Me.BeginInvoke(updateAction)
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
            If lblStatus IsNot Nothing AndAlso Not lblStatus.IsDisposed Then
                lblStatus.Text = mensagem
            End If

            ' Atualizar progresso se especificado
            If progresso >= 0 AndAlso progresso <= 100 AndAlso progressBar IsNot Nothing AndAlso Not progressBar.IsDisposed Then
                progressBar.Value = progresso
            End If

            ' Forçar atualização apenas se necessário
            If Me.Visible Then
                Me.Update() ' Mais eficiente que Refresh()
            End If

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
            If Not Me.IsDisposed AndAlso lblVersion IsNot Nothing AndAlso Not lblVersion.IsDisposed Then
                lblVersion.Text = $"Versão {versao}"
            End If
        Catch
            ' Ignorar erro
        End Try
    End Sub

    ' Método MELHORADO para fechar suavemente
    Public Sub FecharSuave()
        Try
            If Me.IsDisposed OrElse isClosing Then Return

            ' Limpar recursos antes de fechar
            LimparRecursos()

            ' Fade out rápido
            IniciarFadeOut()

        Catch
            Try
                Me.Close()
            Catch
                ' Ignorar erro final
            End Try
        End Try
    End Sub

    Private Sub IniciarFadeOut()
        Try
            ' Reutilizar timer
            If fadeTimer Is Nothing Then
                fadeTimer = New System.Windows.Forms.Timer()
            End If

            fadeTimer.Interval = 5 ' Fade out mais rápido

            ' Remover handlers anteriores e adicionar fade out
            RemoveHandler fadeTimer.Tick, AddressOf FadeInTick
            RemoveHandler fadeTimer.Tick, AddressOf FadeOutTick
            AddHandler fadeTimer.Tick, AddressOf FadeOutTick

            fadeTimer.Start()
        Catch
            Me.Close()
        End Try
    End Sub

    Private Sub FadeOutTick(sender As Object, e As EventArgs)
        Try
            If Me.Opacity > 0.1 Then
                Me.Opacity -= 0.1
            Else
                fadeTimer.Stop()
                RemoveHandler fadeTimer.Tick, AddressOf FadeOutTick
                Me.Close()
            End If
        Catch
            Me.Close()
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

    Private Sub SplashForm_FormClosed(sender As Object, e As FormClosedEventArgs) Handles Me.FormClosed
        ' Garantir limpeza de recursos quando form for fechado
        LimparRecursos()
    End Sub

    ' NOVO: Cleanup adequado - sobrescrever método do Designer
    Private Sub LimparRecursos()
        Try
            isClosing = True

            If fadeTimer IsNot Nothing Then
                fadeTimer.Stop()
                fadeTimer.Dispose()
                fadeTimer = Nothing
            End If
        Catch ex As Exception
            Try
                LogErros.RegistrarErro(ex, "SplashForm.LimparRecursos")
            Catch
                ' Ignorar se LogErros falhar
            End Try
        End Try
    End Sub

    ' NOVO: Método para mostrar splash sem travar UI
    Public Shared Function MostrarSplashAsync(versao As String) As SplashForm
        Try
            Dim splash As New SplashForm()
            splash.DefinirVersao(versao)
            splash.Show()
            Application.DoEvents()
            Return splash
        Catch ex As Exception
            Try
                LogErros.RegistrarErro(ex, "SplashForm.MostrarSplashAsync")
            Catch
            End Try
            Return Nothing
        End Try
    End Function

End Class