Public Class Form2
    Private Const WS_EX_NOACTIVATE As Integer = &H8000000
    Private Const WM_MOUSEACTIVATE As Integer = &H21
    Private Const MA_NOACTIVATE As Integer = &H3
    Private Cls = My.Resources.ResourceManager.GetObject("Close")
    Private offset As Point

    Protected Overrides ReadOnly Property CreateParams() As System.Windows.Forms.CreateParams
        Get
            Dim cp As System.Windows.Forms.CreateParams = MyBase.CreateParams
            cp.ExStyle = cp.ExStyle Or WS_EX_NOACTIVATE
            Return cp
        End Get
    End Property

    Protected Overrides Sub WndProc(ByRef m As Message)
        If (m.Msg = WM_MOUSEACTIVATE) Then
            m.Result = MA_NOACTIVATE
        Else
            MyBase.WndProc(m)
        End If
    End Sub

    Private Sub Label_MouseLeave(ByVal sender As Object, ByVal e As System.EventArgs) Handles Label1.MouseLeave, Label2.MouseLeave, Label3.MouseLeave, Label4.MouseLeave, Label5.MouseLeave, Label6.MouseLeave, Label7.MouseLeave, Label8.MouseLeave, Label9.MouseLeave, Label10.MouseLeave, Label11.MouseLeave, Label12.MouseLeave, Label13.MouseLeave, Label14.MouseLeave, Label15.MouseLeave, Label16.MouseLeave, Label17.MouseLeave, Label18.MouseLeave, Label19.MouseLeave, Label20.MouseLeave, Label21.MouseLeave, Label22.MouseLeave, Label23.MouseLeave
        sender.BorderStyle = BorderStyle.None
    End Sub

    Private Sub Label_MouseMove(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles Label1.MouseMove, Label2.MouseMove, Label3.MouseMove, Label4.MouseMove, Label5.MouseMove, Label6.MouseMove, Label7.MouseMove, Label8.MouseMove, Label9.MouseMove, Label10.MouseMove, Label11.MouseMove, Label12.MouseMove, Label13.MouseMove, Label14.MouseMove, Label15.MouseMove, Label16.MouseMove, Label17.MouseMove, Label18.MouseMove, Label19.MouseMove, Label20.MouseMove, Label21.MouseMove, Label22.MouseMove, Label23.MouseMove
        sender.BorderStyle = BorderStyle.FixedSingle
    End Sub

    Private Sub Label_Click(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles Label1.Click, Label2.Click, Label3.Click, Label4.Click, Label5.Click, Label6.Click, Label7.Click, Label8.Click, Label9.Click, Label10.Click, Label11.Click, Label12.Click, Label13.Click, Label14.Click, Label15.Click, Label16.Click, Label17.Click, Label18.Click, Label19.Click, Label20.Click, Label21.Click, Label22.Click, Label23.Click
        SendKeys.Send(sender.Tag)
    End Sub

    Private Sub Form2_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles Me.KeyPress
        If e.KeyChar = Chr(27) Then
            Me.Close()
        End If
    End Sub

    Private Sub Form2_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        For Each i In Me.Controls
            If i.GetType Is GetType(Label) And (Not i Is Label24) Then
                ToolTip1.SetToolTip(i, ToolTip1.GetToolTip(i) & vbCrLf & "Click to preview.")
            End If
        Next
    End Sub

    Private Sub Form2_MouseDown(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles Me.MouseDown
        If e.Button = Windows.Forms.MouseButtons.Left Then
            offset.X = MousePosition.X - Me.Left
            offset.Y = MousePosition.Y - Me.Top
        End If
    End Sub

    Private Sub Form2_MouseMove(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles Me.MouseMove
        If e.Button = Windows.Forms.MouseButtons.Left Then
            Me.Left = MousePosition.X - offset.X
            Me.Top = MousePosition.Y - offset.Y
        End If
    End Sub

    Private Sub Form2_MouseWheel(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles Me.MouseWheel
        Me.Opacity += 0.05 * e.Delta / Math.Abs(e.Delta)
    End Sub

    Private Sub Label24_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label24.Click
        Me.Close()
    End Sub

    Private Sub Label24_MouseLeave(ByVal sender As Object, ByVal e As System.EventArgs) Handles Label24.MouseLeave
        Label24.Image = Nothing
    End Sub

    Private Sub Label24_MouseMove(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles Label24.MouseMove
        Label24.Image = Cls
    End Sub
End Class