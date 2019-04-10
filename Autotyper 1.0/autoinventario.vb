Public Class Form1



    Sub Button1_Click() Handles Button1.Click
        Delay(5)
        Dim okroadcase As String
        Dim nuevoroadcase As String
        Dim tick As String
        Dim linea As String()
        Dim codigos As String
        Dim i As Integer

        i = 0
        codigos = TextBox1.Text

        linea = Split(codigos, vbCrLf)
        okroadcase = "+{TAB 2}{ENTER}"
        nuevoroadcase = "{HOME}{DOWN}{TAB}{ENTER}{TAB 2}{ENTER}"
        tick = "{Enter}{TAB}{RIGHT} {LEFT 2}"

        For i = 0 To UBound(linea)
            If i = UBound(linea) Then
                SendKeys.Send(linea(i))
            Else
                If linea(i).Contains("RC") Then
                    SendKeys.Send(linea(i))
                    SendKeys.Send(tick)

                Else
                    If linea(i + 1).Contains("RC") Then
                        SendKeys.Send(linea(i) & "{Enter}")
                        SendKeys.Send(okroadcase & "{Enter}")
                        Delay(2)
                        SendKeys.Send(nuevoroadcase)
                    Else
                        SendKeys.Send(linea(i) & "{Enter}")
                    End If

                End If
            End If



        Next i

        'Dim TextArray() As String = TextBox1.Lines //otra forma de hacer un array

        'For Each strLine In TextArray

        '    SendKeys.Send(TextBox1.Text)

        'Next



        MessageBox.Show("El proceso ha terminado")

    End Sub




End Class
Public Class Form2
    Sub MakeOnTop()
        Form1.TopMost = True
    End Sub
End Class
