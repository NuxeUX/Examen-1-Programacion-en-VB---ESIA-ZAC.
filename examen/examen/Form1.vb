Public Class Form1

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Dim km, pago, viaje, redon As Integer
        Dim total, total1, tot, totre, puntos As Single

        km = Val(TextBox1.Text)
        redon = km * 2 'KILOMETROS REDONDOS
        tot = km * 1.48 'km x precio
        totre = redon * 1.48 'km redondos x precio
        pago = Val(TextBox2.Text)
        viaje = Val(TextBox3.Text)

        If km > 0 Then ' KM MAYORES A 0

            If pago = 1 Then ' PAGO EFECTIVO

                If viaje = 1 Then 'VIAJE SENCILLO


                    If km > 1200 Then
                        total = tot - (tot * 0.15)
                        puntos = km * 0.3
                    ElseIf km > 800 Then
                        total = tot - (tot * 0.09)
                        puntos = km * 0.18
                    ElseIf km > 600 Then
                        total = tot - (tot * 0.05)
                        puntos = km * 0.15
                    ElseIf km > 400 Then
                        total = tot - (tot * 0.04)
                        puntos = km * 0.1
                    ElseIf km <= 400 Then
                        total = tot - (tot * 0.03)
                        puntos = km * 0.08
                    End If



                        ElseIf viaje = 2 Then 'VIAJE REDONDO



                    If redon > 1200 Then
                        total = totre - (totre * 0.2) '20% 
                        puntos = redon * 0.3
                    ElseIf redon > 800 Then
                        total = (totre - (totre * 0.1)) '10% 
                        puntos = redon * 0.25
                    ElseIf redon > 600 Then
                        total1 = totre - (totre * 0.05)
                        total = total1 - (total1 * 0.07) '7% adicional
                        puntos = redon * 0.2
                    ElseIf redon > 400 Then
                        total1 = totre - (totre * 0.04)
                        total = (total1 - total1 * 0.06) '6% adicional
                        puntos = redon * 0.15
                    ElseIf redon <= 400 Then
                        total1 = totre - (totre * 0.03)
                        total = total1 - (total * 0.04) '4% adicional
                        puntos = redon * 0.08
                    End If


                    '///////////////////


                    Else
                        MsgBox("VIAJE NO VALIDO, ELIGE 1 ó 2", MsgBoxStyle.Critical)

                    End If ' VIN DE VIAJE


                    '////////////////////////////////////////////////////////////////////


                ElseIf pago = 2 Then ' PAGO TARJETA  ------ 2% ADICIONAL


                    If viaje = 1 Then 'VIAJE SENCILLO


                    If km > 1200 Then
                        total1 = tot - (tot * 0.15)
                        total = total1 - (total1 * 0.02)
                        puntos = (km * 0.30) + (km * 0.02)
                    ElseIf km > 800 Then
                        total1 = tot - (tot * 0.09)
                        total = total1 - (total1 * 0.02)
                        puntos = (km * 0.18) + (km * 0.02)
                    ElseIf km > 600 Then
                        total1 = tot - (tot * 0.05)
                        total = total1 - (total1 * 0.02)
                        puntos = (km * 0.15) + (km * 0.02)
                    ElseIf km > 400 Then
                        total1 = tot - (tot * 0.04)
                        total = total1 - (total1 * 0.02)
                        puntos = (km * 0.10) + (km * 0.02)
                    ElseIf km <= 400 Then
                        total1 = tot - (tot * 0.03)
                        total = total1 - (total1 * 0.02)
                        puntos = (km * 0.8) + (km * 0.02)
                    End If



                    ElseIf viaje = 2 Then 'VIAJE REDONDO



                    If redon > 1200 Then
                        total = totre - (totre * 0.1) - (totre - (totre * 0.2) * 0.02)
                        puntos = (redon * 0.30) + (redon * 0.02)
                    ElseIf redon > 800 Then
                        total = totre - (totre * 0.1) - (totre - (totre * 0.1) * 0.02)
                        puntos = (redon * 0.25) + (redon * 0.02)
                    ElseIf redon > 600 Then
                        total1 = (totre - (totre * 0.05)) - ((totre - (totre * 0.05)) * 0.07)
                        total = total1 - (total1 * 0.02)
                        puntos = (redon * 0.20) + (redon * 0.02)
                    ElseIf redon > 400 Then
                        total1 = (totre - (totre * 0.04)) - ((totre - (totre * 0.04)) * 0.06)
                        total = total1 - (total1 * 0.02)
                        puntos = (redon * 0.15) + (redon * 0.02)
                    ElseIf redon <= 400 Then
                        total1 = (totre - (totre * 0.03)) - ((totre - (totre * 0.03)) * 0.04)
                        total = total1 - (total1 * 0.02)
                        puntos = (redon * 0.8) + (redon * 0.02)
                    End If



                    '///////////////////


                    Else
                        MsgBox("VIAJE NO VALIDO, ELIGE 1 ó 2", MsgBoxStyle.Critical)

                    End If ' VIN DE VIAJE



                    '////////////////////////////////////////////



                Else
                    MsgBox("PAGO NO VALIDO, ELIGE 1 ó 2 ", MsgBoxStyle.Critical)


                End If ' FIN DE PAGO


        Else 'ELSE KM
            MsgBox("INGRESA KILOMETROS REALES ", MsgBoxStyle.Critical)
            End If ' FIN DE KM
            Label4.Text = "TOTAL A PAGAR:  " & total & "           PUNTOS PROXIMO VIAJE:  " & puntos
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        End
    End Sub

    Private Sub AcercaDeToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles AcercaDeToolStripMenuItem.Click
        MsgBox(" - Carlitrox Programmer -     Version 2.1.1 ", MsgBoxStyle.ApplicationModal)
    End Sub

    Private Sub ComoFuncionaToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComoFuncionaToolStripMenuItem.Click
        MsgBox(" 1.- Deberas Ingresar los kilometros desdeados.                                              2.- Ingresa el metodo de pago: Número 1 si es Efectivo, y 2 si es Tarjeta.  3.- Ingresa si tu viaje sera Sencillo o Redondo con numeros 1 o 2.", MsgBoxStyle.Information)
    End Sub

    Private Sub SalirToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles SalirToolStripMenuItem.Click
        End
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        TextBox1.Clear()
        TextBox2.Clear()
        TextBox3.Clear()
        Label4.Text = ""
    End Sub
End Class
