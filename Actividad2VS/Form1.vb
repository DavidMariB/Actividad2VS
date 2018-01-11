Public Class Form1

    Dim nRandom As Integer
    Dim contador As Integer = 0

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click

        Randomize()
        nRandom = Int(100 * Rnd() + 1)
        'MessageBox.Show("El numero aleatorio es: " + nRandom.ToString, "Numero Random")

        Button1.Enabled = False
    End Sub

    Private Sub TextBox1_TextChanged(sender As Object, e As KeyEventArgs) Handles TextBox1.KeyDown

        If e.KeyCode = Keys.Enter Then
            Try
                If Button1.Enabled Then
                    MessageBox.Show("Debes generar un numero aletorio antes de poder adivinarlo", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    TextBox1.Text = ""
                ElseIf Integer.Parse(TextBox1.Text) < 0 Then
                    MessageBox.Show("No puedes introducir un numero negativo", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    TextBox1.Text = ""
                ElseIf TextBox1.Text.Equals("") Then
                    MessageBox.Show("Tienes que introducir un numero", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    TextBox1.Text = ""
                ElseIf Integer.Parse(TextBox1.Text) < 1 Or Integer.Parse(TextBox1.Text) > 100 Then
                    MessageBox.Show("Debes introducir un numero entre 1 y 100", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    TextBox1.Text = ""
                ElseIf Integer.Parse(TextBox1.Text) < nRandom Then
                    ListBox1.Items.Add(TextBox1.Text + " El numero secreto es mayor")
                    ListBox1.TopIndex = ListBox1.Items.Count - 1
                    TextBox1.Text = ""
                    contador = contador + 1
                ElseIf Integer.Parse(TextBox1.Text) > nRandom Then
                    ListBox1.Items.Add(TextBox1.Text + " El numero secreto es menor")
                    ListBox1.TopIndex = ListBox1.Items.Count - 1
                    TextBox1.Text = ""
                    contador = contador + 1
                ElseIf Integer.Parse(TextBox1.Text) = nRandom Then
                    TextBox1.Text = ""
                    contador = contador + 1
                    Select Case MessageBox.Show("Has acertado el numero secreto. ¿Quieres volver a jugar? Intentos utilizados: " + contador.ToString, "¡Enhorabuena!", MessageBoxButtons.YesNo, MessageBoxIcon.Information)
                        Case DialogResult.Yes
                            funcionLimpiar()
                        Case DialogResult.No
                            funcionSalir()
                    End Select
                End If
            Catch ex As FormatException
                MessageBox.Show("Solo se pueden introducir numeros enteros y positivos", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                TextBox1.Text = ""
            End Try
        End If
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        funcionLimpiar()
    End Sub

    Sub funcionLimpiar()
        Button1.Enabled = True
        ListBox1.Items.Clear()
        TextBox1.Text = ""
        contador = 0
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        funcionSalir()
    End Sub

    Sub funcionSalir()
        Application.Exit()
    End Sub
End Class