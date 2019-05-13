Imports System.Runtime.InteropServices
Imports System.Text
Public Class Form1

    <DllImport("user32.dll", EntryPoint:="GetWindowThreadProcessId")>
    Private Shared Function GetWindowThreadProcessId(<InAttribute()> ByVal hWnd As IntPtr, ByRef lpdwProcessId As Integer) As Integer
    End Function
    <DllImport("user32.dll", EntryPoint:="GetForegroundWindow")> Private Shared Function GetForegroundWindow() As IntPtr
    End Function
    <DllImport("user32.dll", SetLastError:=True, CharSet:=CharSet.Auto)> Private Shared Function GetWindowTextLength(ByVal hwnd As IntPtr) As Integer
    End Function
    <DllImport("user32.dll", EntryPoint:="GetWindowTextW")>
    Private Shared Function GetWindowTextW(<InAttribute()> ByVal hWnd As IntPtr, <OutAttribute(), MarshalAs(UnmanagedType.LPWStr)> ByVal lpString As StringBuilder, ByVal nMaxCount As Integer) As Integer
    End Function
    Dim codigoLote As String = ""
    Dim codigoLoteB As String = ""
    Dim limiteDeLotes As Integer
    Dim errorRC As Boolean
    Dim focoRoadcaseOK As Boolean
    Dim windowError As Boolean
    Dim paraloTodo As Boolean
    Dim limiteAparatos As Integer
    Dim codigoAparatos As Integer


    Function QuitaEspacios(ByVal Texto As String, ByVal DelComienzo As Boolean, ByVal DelFinal As Boolean) As String
        'On Local Error Resume Next

        If DelComienzo = True Then
            ' hacemos un bucle que se repetira hasta que la variable "Texto"
            ' no contenga ningun salto de linea con un espacio a continuacion.
            ' Comparacion: Si dentro de la cadena "Texto" encontramos la cadena "saltodelinea+espacio"...
            Do Until InStr(1, Texto, " ") = 0
                ' reemplazamos dentro de "Texto" todas las cadenas 
                ' "saltodelinea+espacio" por la cadena "saltodelinea" (osea vbcrlf).
                Texto = Replace(Texto, " ", "")
            Loop
        End If
        QuitaEspacios = Texto
    End Function
    Function LineTrim(ByVal sText As String) As String
        '-- add pre Lf and post Cr
        sText = vbLf & sText & vbCr '<--- unusual characters added
        '-- remove all spaces at the end of lines
        Do While InStr(sText, " " & vbCr)
            sText = Replace(sText, " " & vbCr, vbCr)
        Loop
        '-- remove all multiple CrLf's 
        sText = Replace(sText, vbLf & vbCr, "") '<--- westconn1's smart line
        '-- remove first and last added characters
        LineTrim = Mid$(sText, 2, Len(sText) - 2)
    End Function
    Function focoRoadcase()
        paraloTodo = False
        windowError = True
        While windowError And Not paraloTodo
            Dim hWnd As IntPtr = GetForegroundWindow()
            If Not hWnd.Equals(IntPtr.Zero) Then
                Dim lgth As Integer = GetWindowTextLength(hWnd)
                Dim wTitle As New System.Text.StringBuilder("", lgth + 1)
                If lgth > 0 Then
                    GetWindowTextW(hWnd, wTitle, wTitle.Capacity)
                End If
                Dim wProcID As Integer = Nothing
                GetWindowThreadProcessId(hWnd, wProcID)
                Dim Proc As Process = Process.GetProcessById(wProcID)
                Dim wFileName As String = ""
                Try
                    wFileName = Proc.MainModule.FileName
                Catch ex As Exception
                    wFileName = ""
                End Try
                If wTitle.ToString = "Pack a Road case" Then
                    windowError = False
                Else
                    windowError = True
                    Dim msg2 As String
                    Dim title2 As String
                    Dim style2 As MsgBoxStyle
                    Dim response2 As MsgBoxResult
                    msg2 = "La ventana 'Pack a Road case' no quedó seleccionada." & vbNewLine & "Selecciona la ventana 'Pack a Road case' y clicka 'YES'." & vbNewLine & "Selecciona NO para parar el proceso."   ' Define message.
                    style2 = MsgBoxStyle.MsgBoxSetForeground & MsgBoxStyle.YesNo
                    title2 = "¡¡¡La ventana perdió el foco!!!"   ' Define title.
                    ' Display message.
                    response2 = MsgBox(msg2, style2, title2)
                    If response2 = MsgBoxResult.No Then
                        paraloTodo = True
                    End If
                End If
            End If
        End While
        Return windowError
    End Function

    Function WarningWindow()
        paraloTodo = False
        windowError = True
        While windowError And Not paraloTodo
            Dim hWnd As IntPtr = GetForegroundWindow()
            If Not hWnd.Equals(IntPtr.Zero) Then
                Dim lgth As Integer = GetWindowTextLength(hWnd)
                Dim wTitle As New System.Text.StringBuilder("", lgth + 1)
                If lgth > 0 Then
                    GetWindowTextW(hWnd, wTitle, wTitle.Capacity)
                End If
                Dim wProcID As Integer = Nothing
                GetWindowThreadProcessId(hWnd, wProcID)
                Dim Proc As Process = Process.GetProcessById(wProcID)
                Dim wFileName As String = ""
                Try
                    wFileName = Proc.MainModule.FileName
                Catch ex As Exception
                    wFileName = ""
                End Try
                If wTitle.ToString = "Information" Then
                    windowError = True
                    paraloTodo = True
                Else
                    windowError = False
                End If
            End If
        End While
        Return windowError
    End Function

    Function volcado()
        Dim okroadcase As String
        Dim nuevoroadcase As String
        Dim tick As String
        Dim linea As String()
        Dim codigos As String
        Dim i As Integer
        Dim roadcasesHechos As String
        Dim lineaConRC As String
        Dim textSinEspacios As String = ""
        Dim lineaRaw As String()
        Dim espacio As Integer = 0
        Delay(5)
        lineaConRC = ""
        roadcasesHechos = ""
        i = 0
        codigos = TextBox1.Text
        lineaRaw = Split(codigos, vbCrLf)
        okroadcase = "+{TAB 2}{ENTER}"
        nuevoroadcase = "{HOME}{DOWN}{TAB}{ENTER}{TAB 2}{ENTER}"
        tick = "{Enter}{TAB}{RIGHT} {LEFT 2}"
        Dim num As Integer = 0
        focoRoadcase()
        If Not windowError Then
            For i = 0 To UBound(lineaRaw)
                If Not String.IsNullOrEmpty(lineaRaw(i)) And Not lineaRaw(i) = vbCrLf Then
                    textSinEspacios += lineaRaw(i) & vbCrLf
                End If
            Next i

            linea = Split(textSinEspacios, vbCrLf)
            For i = 0 To UBound(linea)

                If i = UBound(linea) Then
                    SendKeys.Send(linea(i) & "{Enter}")
                Else
                    If linea(i).Contains(".RC") Then
                        SendKeys.Send(linea(i))
                        SendKeys.Send(tick)
                    Else
                        If linea(i + 1).Contains(".RC") Then
                            SendKeys.Send(linea(i) & "{Enter}")
                            Delay(0.2)
                            WarningWindow()
                            If paraloTodo And windowError Then
                                Return False
                                Exit For
                            End If
                            Dim msg As String
                            Dim title As String
                            Dim style As MsgBoxStyle
                            Dim response As MsgBoxResult
                            msg = "Roadcase completo." & vbNewLine & "Si seleccionas 'YES' el proceso continuará." & vbNewLine & "Si seleccionas NO el proceso finalizará y deberás repasar lo que se pueda haber quedado a medias."   ' Define message.
                            style = MsgBoxStyle.MsgBoxSetForeground & MsgBoxStyle.YesNo
                            title = "Roadcase completo"   ' Define title.
                            ' Display message.
                            response = MsgBox(msg, style, title)
                            If response = MsgBoxResult.Yes Then   ' User chose Yes.
                                focoRoadcase()
                                If paraloTodo And windowError Then
                                    Exit For
                                Else
                                    SendKeys.Send(okroadcase & "{Enter}")
                                    Delay(2)
                                    SendKeys.Send(nuevoroadcase)
                                End If
                            Else
                                Exit For
                            End If
                        Else
                            SendKeys.Send(linea(i) & "{Enter}")
                            Delay(0.2)
                            WarningWindow()
                            If paraloTodo And windowError Then
                                Return False
                                Exit For
                            End If
                        End If
                    End If
                End If
                If linea(i).Contains(".RC") Then
                    lineaConRC = linea(i)
                    roadcasesHechos = roadcasesHechos & lineaConRC & vbCr
                End If

            Next i
            Return MsgBox("Incluyendo el que queda activo, los siguientes paquetes se han procesado:" & vbNewLine & roadcasesHechos & "Por favor repasa los paquetes hechos.", MsgBoxStyle.MsgBoxSetForeground & MsgBoxStyle.Information, "El proceso ha terminado")
        Else
            Return MsgBox("No se ha creado ningún paquete.", MsgBoxStyle.MsgBoxSetForeground & MsgBoxStyle.Information, "El proceso ha terminado")
        End If
    End Function
    Function comprobacion()

        Dim linea2 As String() = Split(TextBox1.Text, vbCrLf)
        Dim x As Integer = 0
        Dim y As Integer = 0
        errorRC = False

        For x = 0 To UBound(linea2)
            If linea2(x).Contains(".RC") Then
                If (codigoLote.Length > 0 And linea2(x).Contains(codigoLote)) Or
                    (codigoLoteB.Length > 0 And linea2(x).Contains(codigoLoteB)) Then
                    y = y + 1
                Else
                    errorRC = True
                    Exit For
                End If
            End If
        Next x
        If errorRC Then
            MessageBox.Show("Se ha encontrado algun paquete que no es de " & ComboBox1.SelectedItem & ". Revisa los códigos")
        End If

        If y > limiteDeLotes Then
            errorRC = True
            MessageBox.Show("No puedes hacer más de " & limiteDeLotes & " paquetes de golpe.")
        End If
        Return errorRC
    End Function
    Sub Button1_Click() Handles Button1.Click
        On Error GoTo ErrorHandler
        TextBox1.Text = Replace(TextBox1.Text, vbTab, "")
        TextBox1.Text = Trim$(QuitaEspacios(TextBox1.Text, True, True))
        TextBox1.Text = LineTrim(TextBox1.Text)
        Dim linea3 As String()
        Dim codigos3 As String
        Dim codigosDuplicados As String = ""
        codigos3 = TextBox1.Text
        linea3 = Split(codigos3, vbCrLf)

        'ABSEN A3
        If ComboBox1.Text = "Absen (6)" Then
            codigoLote = "A3"
            limiteDeLotes = 6
            Dim errComp As String = ""
            Dim h As Integer = 0
            Dim p As Integer = 0
            Dim paquetesQueHay As Integer = 0
            Dim absenEnLista As Integer = 0
            Dim absenQueHay As Integer = 0
            Dim absenHayPorPaquete As Double = 0
            Dim absenPorPaquete As Integer = 6
            Dim loteOK As Boolean = True
            Dim mensajeFinal As String = ""

            If Not linea3(0).Contains(".RC") Then
                mensajeFinal += "El primer código del paquete tiene que ser el de la caja" & vbNewLine & vbNewLine
                loteOK = False
            End If
            If loteOK = True Then
                For h = 0 To UBound(linea3)
                    For p = 0 To UBound(linea3)
                        If Not p = h Then
                            If linea3(p) = linea3(h) Then
                                loteOK = False
                                If Not codigosDuplicados.Contains(linea3(p)) Then
                                    codigosDuplicados += linea3(p) & vbNewLine
                                End If
                            End If
                        End If
                    Next p
                Next h
                If Not codigosDuplicados = "" Then
                    mensajeFinal += "Los siguientes códigos aparecen más de una vez: " & vbNewLine & codigosDuplicados & vbNewLine & vbNewLine
                End If
            End If

            If loteOK = True Then
                For h = 0 To UBound(linea3)
                    If linea3(h).Contains(".RC") Or linea3(h).Contains("SPA3/") Then

                    Else
                        loteOK = False
                        If Not errComp.Contains(linea3(h)) Then
                            errComp += linea3(h) & vbNewLine
                        End If
                    End If
                Next h
                If Not errComp = "" Then
                    mensajeFinal += "Los siguientes códigos no deben ir en un paquete de ABSEN: " & vbNewLine & errComp & vbNewLine & vbNewLine
                End If
            End If
            If loteOK = True Then
                For h = 0 To UBound(linea3)
                    If linea3(h).Contains(codigoLote + ".RC") Then
                        paquetesQueHay += 1
                    End If
                    If linea3(h).Contains("SPA3/") Then
                        absenQueHay += 1
                    End If
                Next h
                absenHayPorPaquete = absenQueHay / paquetesQueHay
                If absenHayPorPaquete <> absenPorPaquete Then
                    mensajeFinal += "En algún paquete, la cantidad de tiles Absen por paquete difiere de " & absenPorPaquete & vbNewLine
                    loteOK = False
                End If
            End If
            comprobacion()
            If errorRC = False And loteOK = True Then
                MsgBox("Después de darle al OK tendrás 5 segundos para seleccionar la ventana 'Pack a Road case'")
                volcado()
            ElseIf errorRC = False And loteOK = False Then
                MsgBox(mensajeFinal & vbNewLine & "Repasa los códigos.", MsgBoxStyle.MsgBoxSetForeground & MsgBoxStyle.Information, "Hay errores")
            End If
        End If

        'ABSEN BACK SUPPORT
        If ComboBox1.Text = "Absen Back Support (6)" Then
            codigoLote = "A3BS"
            limiteDeLotes = 6
            Dim errComp As String = ""
            Dim h As Integer = 0
            Dim p As Integer = 0
            Dim paquetesQueHay As Integer = 0
            Dim backSupportQueHay As Integer = 0
            Dim backSupportHayPorPaquete As Double = 0
            Dim backSupportPorPaquete As Integer = 6
            Dim bandejaQueHay As Integer = 0
            Dim bandejaHayPorPaquete As Double = 0
            Dim bandejaPorPaquete As Integer = 3
            Dim loteOK As Boolean = True
            Dim mensajeFinal As String = ""
            If Not linea3(0).Contains(".RC") Then
                mensajeFinal += "El primer código del paquete tiene que ser el de la caja" & vbNewLine & vbNewLine
                loteOK = False
            End If
            If loteOK = True Then
                For h = 0 To UBound(linea3)
                    For p = 0 To UBound(linea3)
                        If Not p = h Then
                            If linea3(p) = linea3(h) Then
                                loteOK = False
                                If Not codigosDuplicados.Contains(linea3(p)) Then
                                    codigosDuplicados += linea3(p) & vbNewLine
                                End If
                            End If
                        End If
                    Next p
                Next h
                If Not codigosDuplicados = "" Then
                    mensajeFinal += "Los siguientes códigos aparecen más de una vez: " & vbNewLine & codigosDuplicados & vbNewLine & vbNewLine
                End If
            End If

            If loteOK = True Then
                For h = 0 To UBound(linea3)
                    If linea3(h).Contains(".RC") Or
                        linea3(h).Contains("A3BS/") Or linea3(h).Contains("A3PBS1/") Or
                        linea3(h).Contains("A3TRAY/") Or linea3(h).Contains("A3PBT/") Then
                    Else
                        loteOK = False
                        If Not errComp.Contains(linea3(h)) Then
                            errComp += linea3(h) & vbNewLine
                        End If
                    End If
                Next h
                If Not errComp = "" Then
                    mensajeFinal += "Los siguientes códigos no deben ir en un paquete de back support: " & vbNewLine & errComp & vbNewLine & vbNewLine
                End If
            End If
            If loteOK = True Then
                For h = 0 To UBound(linea3)
                    If linea3(h).Contains(codigoLote + ".RC") Then
                        paquetesQueHay += 1
                    End If
                    If linea3(h).Contains("A3BS/") Or linea3(h).Contains("A3PBS1/") Then
                        backSupportQueHay += 1
                    End If
                    If linea3(h).Contains("A3TRAY/") Or linea3(h).Contains("A3PBT/") Then
                        bandejaQueHay += 1
                    End If
                Next h
                backSupportHayPorPaquete = backSupportQueHay / paquetesQueHay
                bandejaHayPorPaquete = bandejaQueHay / paquetesQueHay
                If backSupportHayPorPaquete <> backSupportPorPaquete Then
                    mensajeFinal += "En algún paquete, la cantidad de back support difiere de " & backSupportPorPaquete & vbNewLine
                    loteOK = False
                End If
                If bandejaHayPorPaquete <> bandejaPorPaquete Then
                    mensajeFinal += "En algún paquete, la cantidad de bandejas de 1m difiere de " & bandejaPorPaquete & vbNewLine
                    loteOK = False
                End If
            End If
            comprobacion()
            If errorRC = False And loteOK = True Then
                MsgBox("Después de darle al OK tendrás 5 segundos para seleccionar la ventana 'Pack a Road case'")
                volcado()
            ElseIf errorRC = False And loteOK = False Then
                MsgBox(mensajeFinal & vbNewLine & "Repasa los códigos.", MsgBoxStyle.MsgBoxSetForeground & MsgBoxStyle.Information, "Hay errores")
            End If
        End If

        'ABSEN GROUND BEAM 2 x 1m + 1 x 1.5m
        If ComboBox1.Text = "Absen Ground Beam 1 x 1.5m + 2 x 1m(4)" Then
            codigoLote = "A3GB1"
            limiteDeLotes = 4
            Dim errComp As String = ""
            Dim h As Integer = 0
            Dim p As Integer = 0
            Dim paquetesQueHay As Integer = 0
            Dim groundBeam1halfQueHay As Integer = 0
            Dim groundBeam1halfHayPorPaquete As Double = 0
            Dim groundBeam1halfPorPaquete As Integer = 1
            Dim groundBeam1mQueHay As Integer = 0
            Dim groundBeam1mHayPorPaquete As Double = 0
            Dim groundBeam1mPorPaquete As Integer = 2
            Dim outriggerQueHay As Integer = 0
            Dim outriggerHayPorPaquete As Double = 0
            Dim outriggerPorPaquete As Integer = 4
            Dim loteOK As Boolean = True
            Dim mensajeFinal As String = ""
            If Not linea3(0).Contains(".RC") Then
                mensajeFinal += "El primer código del paquete tiene que ser el de la caja" & vbNewLine & vbNewLine
                loteOK = False
            End If
            If loteOK = True Then
                For h = 0 To UBound(linea3)
                    For p = 0 To UBound(linea3)
                        If Not p = h Then
                            If linea3(p) = linea3(h) Then
                                loteOK = False
                                If Not codigosDuplicados.Contains(linea3(p)) Then
                                    codigosDuplicados += linea3(p) & vbNewLine
                                End If
                            End If
                        End If
                    Next p
                Next h
                If Not codigosDuplicados = "" Then
                    mensajeFinal += "Los siguientes códigos aparecen más de una vez: " & vbNewLine & codigosDuplicados & vbNewLine & vbNewLine
                End If
            End If

            If loteOK = True Then
                For h = 0 To UBound(linea3)
                    If linea3(h).Contains(".RC") Or
                        linea3(h).Contains("A3GB1.5/") Or
                        linea3(h).Contains("A3GB1/") Or
                        linea3(h).Contains("A3OUT/") Then
                    Else
                        loteOK = False
                        If Not errComp.Contains(linea3(h)) Then
                            errComp += linea3(h) & vbNewLine
                        End If
                    End If
                Next h
                If Not errComp = "" Then
                    mensajeFinal += "Los siguientes códigos no deben ir en un paquete de ground beam de 1.5m y 1m: " & vbNewLine & errComp & vbNewLine & vbNewLine
                End If
            End If
            If loteOK = True Then
                For h = 0 To UBound(linea3)
                    If linea3(h).Contains(codigoLote + ".RC") Then
                        paquetesQueHay += 1
                    End If
                    If linea3(h).Contains("A3GB1.5/") Then
                        groundBeam1halfQueHay += 1
                    End If
                    If linea3(h).Contains("A3GB1/") Then
                        groundBeam1mQueHay += 1
                    End If
                    If linea3(h).Contains("A3OUT/") Then
                        outriggerQueHay += 1
                    End If
                Next h
                groundBeam1halfHayPorPaquete = groundBeam1halfQueHay / paquetesQueHay
                groundBeam1mHayPorPaquete = groundBeam1mQueHay / paquetesQueHay
                outriggerHayPorPaquete = outriggerQueHay / paquetesQueHay
                If groundBeam1halfHayPorPaquete <> groundBeam1halfPorPaquete Then
                    mensajeFinal += "En algún paquete, la cantidad de ground beam de 1.5m difiere de " & groundBeam1halfPorPaquete & vbNewLine
                    loteOK = False
                End If
                If groundBeam1mHayPorPaquete <> groundBeam1mPorPaquete Then
                    mensajeFinal += "En algún paquete, la cantidad de ground beam de 1m difiere de " & groundBeam1mPorPaquete & vbNewLine
                    loteOK = False
                End If
                If outriggerHayPorPaquete <> outriggerPorPaquete Then
                    mensajeFinal += "En algún paquete, la cantidad de outriggers difiere de " & outriggerPorPaquete & vbNewLine
                    loteOK = False
                End If
            End If
            comprobacion()
            If errorRC = False And loteOK = True Then
                MsgBox("Después de darle al OK tendrás 5 segundos para seleccionar la ventana 'Pack a Road case'")
                volcado()
            ElseIf errorRC = False And loteOK = False Then
                MsgBox(mensajeFinal & vbNewLine & "Repasa los códigos.", MsgBoxStyle.MsgBoxSetForeground & MsgBoxStyle.Information, "Hay errores")
            End If
        End If

        'ABSEN GROUND BEAM 2 x 2m
        If ComboBox1.Text = "Absen Ground Beam 2 x 2m (4)" Then
            codigoLote = "A3GB2"
            limiteDeLotes = 4
            Dim errComp As String = ""
            Dim h As Integer = 0
            Dim p As Integer = 0
            Dim paquetesQueHay As Integer = 0
            Dim groundBeam2mQueHay As Integer = 0
            Dim groundBeam2mHayPorPaquete As Double = 0
            Dim groundBeam2mPorPaquete As Integer = 2
            Dim outriggerQueHay As Integer = 0
            Dim outriggerHayPorPaquete As Double = 0
            Dim outriggerPorPaquete As Integer = 4
            Dim loteOK As Boolean = True
            Dim mensajeFinal As String = ""
            If Not linea3(0).Contains(".RC") Then
                mensajeFinal += "El primer código del paquete tiene que ser el de la caja" & vbNewLine & vbNewLine
                loteOK = False
            End If
            If loteOK = True Then
                For h = 0 To UBound(linea3)
                    For p = 0 To UBound(linea3)
                        If Not p = h Then
                            If linea3(p) = linea3(h) Then
                                loteOK = False
                                If Not codigosDuplicados.Contains(linea3(p)) Then
                                    codigosDuplicados += linea3(p) & vbNewLine
                                End If
                            End If
                        End If
                    Next p
                Next h
                If Not codigosDuplicados = "" Then
                    mensajeFinal += "Los siguientes códigos aparecen más de una vez: " & vbNewLine & codigosDuplicados & vbNewLine & vbNewLine
                End If
            End If

            If loteOK = True Then
                For h = 0 To UBound(linea3)
                    If linea3(h).Contains(".RC") Or
                        linea3(h).Contains("A3GB2/") Or
                        linea3(h).Contains("A3OUT/") Then
                    Else
                        loteOK = False
                        If Not errComp.Contains(linea3(h)) Then
                            errComp += linea3(h) & vbNewLine
                        End If
                    End If
                Next h
                If Not errComp = "" Then
                    mensajeFinal += "Los siguientes códigos no deben ir en un paquete de ground beam de 2m: " & vbNewLine & errComp & vbNewLine & vbNewLine
                End If
            End If
            If loteOK = True Then
                For h = 0 To UBound(linea3)
                    If linea3(h).Contains(codigoLote + ".RC") Then
                        paquetesQueHay += 1
                    End If
                    If linea3(h).Contains("A3GB2/") Then
                        groundBeam2mQueHay += 1
                    End If
                    If linea3(h).Contains("A3OUT/") Then
                        outriggerQueHay += 1
                    End If
                Next h
                groundBeam2mHayPorPaquete = groundBeam2mQueHay / paquetesQueHay
                outriggerHayPorPaquete = outriggerQueHay / paquetesQueHay
                If groundBeam2mHayPorPaquete <> groundBeam2mPorPaquete Then
                    mensajeFinal += "En algún paquete, la cantidad de ground beam de 2m pulgadas difiere de " & groundBeam2mPorPaquete & vbNewLine
                    loteOK = False
                End If
                If outriggerHayPorPaquete <> outriggerPorPaquete Then
                    mensajeFinal += "En algún paquete, la cantidad de outriggers difiere de " & outriggerPorPaquete & vbNewLine
                    loteOK = False
                End If
            End If
            comprobacion()
            If errorRC = False And loteOK = True Then
                MsgBox("Después de darle al OK tendrás 5 segundos para seleccionar la ventana 'Pack a Road case'")
                volcado()
            ElseIf errorRC = False And loteOK = False Then
                MsgBox(mensajeFinal & vbNewLine & "Repasa los códigos.", MsgBoxStyle.MsgBoxSetForeground & MsgBoxStyle.Information, "Hay errores")
            End If
        End If

        '24 DigiLED MCK7
        If ComboBox1.Text = "24 x DigiLED MCK7 (4)" Then
            codigoLote = "24MC7"
            limiteDeLotes = 4
            Dim errComp As String = ""
            Dim h As Integer = 0
            Dim p As Integer = 0
            Dim paquetesQueHay As Integer = 0
            Dim mc7EnLista As Integer = 0
            Dim mc7QueHay As Integer = 0
            Dim mc7HayPorPaquete As Double = 0
            Dim mc7PorPaquete As Integer = 24
            Dim loteOK As Boolean = True
            Dim mensajeFinal As String = ""
            If Not linea3(0).Contains(".RC") Then
                mensajeFinal += "El primer código del paquete tiene que ser el de la caja" & vbNewLine & vbNewLine
                loteOK = False
            End If
            If loteOK = True Then
                For h = 0 To UBound(linea3)
                    For p = 0 To UBound(linea3)
                        If Not p = h Then
                            If linea3(p) = linea3(h) Then
                                loteOK = False
                                If Not codigosDuplicados.Contains(linea3(p)) Then
                                    codigosDuplicados += linea3(p) & vbNewLine
                                End If
                            End If
                        End If
                    Next p
                Next h
                If Not codigosDuplicados = "" Then
                    mensajeFinal += "Los siguientes códigos aparecen más de una vez: " & vbNewLine & codigosDuplicados & vbNewLine & vbNewLine
                End If
            End If

            If loteOK = True Then
                For h = 0 To UBound(linea3)
                    If linea3(h).Contains(".RC") Or linea3(h).Contains("MC7/") Then

                    Else
                        loteOK = False
                        If Not errComp.Contains(linea3(h)) Then
                            errComp += linea3(h) & vbNewLine
                        End If
                    End If
                Next h
                If Not errComp = "" Then
                    mensajeFinal += "Los siguientes códigos no deben en ir en la caja de 24 DigiLED MCK7: " & vbNewLine & errComp & vbNewLine & vbNewLine
                End If
            End If
            If loteOK = True Then
                For h = 0 To UBound(linea3)
                    If linea3(h).Contains(codigoLote + ".RC") Then
                        paquetesQueHay += 1
                    End If
                    If linea3(h).Contains("MC7/") Then
                        mc7QueHay += 1
                    End If
                Next h
                mc7HayPorPaquete = mc7QueHay / paquetesQueHay
                If mc7HayPorPaquete <> mc7PorPaquete Then
                    mensajeFinal += "La cantidad de tiles DigiLED MCK7 por paquete difiere de " & mc7PorPaquete & vbNewLine
                    loteOK = False
                End If
            End If
            comprobacion()
            If errorRC = False And loteOK = True Then
                MsgBox("Después de darle al OK tendrás 5 segundos para seleccionar la ventana 'Pack a Road case'")
                volcado()
            ElseIf errorRC = False And loteOK = False Then
                MsgBox(mensajeFinal & vbNewLine & "Repasa los códigos.", MsgBoxStyle.MsgBoxSetForeground & MsgBoxStyle.Information, "Hay errores")
            End If
        End If



        'PROYECTOR 21K
        If ComboBox1.Text = "21K (2)" Then
            codigoLote = "21K"
            limiteDeLotes = 2
            Dim errComp As String = ""
            Dim h As Integer = 0
            Dim p As Integer = 0
            Dim paquetesQueHay As Integer = 0
            Dim proy21kLista As Integer = 0
            Dim proy21kQueHay As Integer = 0
            Dim proy21kHayPorPaquete As Double = 0
            Dim proy21kPorPaquete As Integer = 1
            Dim mando21kEnLista As Integer = 0
            Dim mando21kQueHay As Integer = 0
            Dim mando21kHayPorPaquete As Double = 0
            Dim mando21kPorPaquete As Integer = 1
            Dim mono16aEnLista As Integer = 0
            Dim mono16aQueHay As Integer = 0
            Dim mono16aHayPorPaquete As Double = 0
            Dim mono16aPorPaquete As Integer = 1
            Dim rentalframe21kEnLista As Integer = 0
            Dim rentalframe21kQueHay As Integer = 0
            Dim rentalframe21kHayPorPaquete As Double = 0
            Dim rentalframe21kPorPaquete As Integer = 1
            Dim loteOK As Boolean = True
            Dim mensajeFinal As String = ""
            If Not linea3(0).Contains(".RC") Then
                mensajeFinal += "El primer código del paquete tiene que ser el de la caja" & vbNewLine & vbNewLine
                loteOK = False
            End If
            If loteOK = True Then
                For h = 0 To UBound(linea3)
                    For p = 0 To UBound(linea3)
                        If Not p = h Then
                            If linea3(p) = linea3(h) Then
                                loteOK = False
                                If Not codigosDuplicados.Contains(linea3(p)) Then
                                    codigosDuplicados += linea3(p) & vbNewLine
                                End If
                            End If
                        End If
                    Next p
                Next h
                If Not codigosDuplicados = "" Then
                    mensajeFinal += "Los siguientes códigos aparecen más de una vez: " & vbNewLine & codigosDuplicados & vbNewLine & vbNewLine
                End If
            End If

            If loteOK = True Then
                For h = 0 To UBound(linea3)
                    If linea3(h).Contains(".RC") Or
                        linea3(h).Contains("PAN21K/") Or
                        (linea3(h).Contains("RCPANMI/") Or linea3(h).Contains("RCPAN/")) Or
                        linea3(h).Contains("RF21K/") Or
                        linea3(h).Contains("SK16/") Then
                    Else
                        loteOK = False
                        If Not errComp.Contains(linea3(h)) Then
                            errComp += linea3(h) & vbNewLine
                        End If
                    End If
                Next h
                If Not errComp = "" Then
                    mensajeFinal += "Los siguientes códigos no deben ir en un paquete de proyector 21K: " & vbNewLine & errComp & vbNewLine & vbNewLine
                End If
            End If
            If loteOK = True Then
                For h = 0 To UBound(linea3)
                    If linea3(h).Contains(codigoLote + ".RC") Then
                        paquetesQueHay += 1
                    End If
                    If linea3(h).Contains("PAN21K/") Then
                        proy21kQueHay += 1
                    End If
                    If linea3(h).Contains("RCPANMI/") Or linea3(h).Contains("RCPAN/") Then
                        mando21kQueHay += 1
                    End If
                    If linea3(h).Contains("RF21K/") Then
                        rentalframe21kQueHay += 1
                    End If
                    If linea3(h).Contains("SK16/") Then
                        mono16aQueHay += 1
                    End If
                Next h
                proy21kHayPorPaquete = proy21kQueHay / paquetesQueHay
                mando21kHayPorPaquete = mando21kQueHay / paquetesQueHay
                rentalframe21kHayPorPaquete = rentalframe21kQueHay / paquetesQueHay
                mono16aHayPorPaquete = mono16aQueHay / paquetesQueHay

                If proy21kHayPorPaquete <> proy21kPorPaquete Then
                    mensajeFinal += "En algún paquete, la cantidad de proyectores de 21K difiere de " & proy21kPorPaquete & vbNewLine
                    loteOK = False
                End If
                If mando21kHayPorPaquete <> mando21kPorPaquete Then
                    mensajeFinal += "En algún paquete, la cantidad de mandos con hdmi de 21K difiere de " & mando21kPorPaquete & vbNewLine
                    loteOK = False
                End If
                If rentalframe21kHayPorPaquete <> rentalframe21kPorPaquete Then
                    mensajeFinal += "En algún paquete, la cantidad de rental frame de 21K difiere de " & rentalframe21kPorPaquete & vbNewLine
                    loteOK = False
                End If
                If mono16aHayPorPaquete <> mono16aPorPaquete Then
                    mensajeFinal += "En algún paquete, la cantidad de conversores schuko a 16A monofásico difiere de " & mono16aPorPaquete & vbNewLine
                    loteOK = False
                End If

            End If
            comprobacion()
            If errorRC = False And loteOK = True Then
                MsgBox("Después de darle al OK tendrás 5 segundos para seleccionar la ventana 'Pack a Road case'")
                volcado()
            ElseIf errorRC = False And loteOK = False Then
                MsgBox(mensajeFinal & vbNewLine & "Repasa los códigos.", MsgBoxStyle.MsgBoxSetForeground & MsgBoxStyle.Information, "Hay errores")
            End If
        End If

        'PROYECTOR 13K
        If ComboBox1.Text = "13K (2)" Then
            codigoLote = "13K"
            limiteDeLotes = 2
            Dim errComp As String = ""
            Dim h As Integer = 0
            Dim p As Integer = 0
            Dim paquetesQueHay As Integer = 0
            Dim proy13kLista As Integer = 0
            Dim proy13kQueHay As Integer = 0
            Dim proy13kHayPorPaquete As Double = 0
            Dim proy13kPorPaquete As Integer = 1
            Dim mando13kEnLista As Integer = 0
            Dim mando13kQueHay As Integer = 0
            Dim mando13kHayPorPaquete As Double = 0
            Dim mando13kPorPaquete As Integer = 1
            Dim rentalframe13kEnLista As Integer = 0
            Dim rentalframe13kQueHay As Integer = 0
            Dim rentalframe13kHayPorPaquete As Double = 0
            Dim rentalframe13kPorPaquete As Integer = 1
            Dim loteOK As Boolean = True
            Dim mensajeFinal As String = ""
            If Not linea3(0).Contains(".RC") Then
                mensajeFinal += "El primer código del paquete tiene que ser el de la caja" & vbNewLine & vbNewLine
                loteOK = False
            End If
            If loteOK = True Then
                For h = 0 To UBound(linea3)
                    For p = 0 To UBound(linea3)
                        If Not p = h Then
                            If linea3(p) = linea3(h) Then
                                loteOK = False
                                If Not codigosDuplicados.Contains(linea3(p)) Then
                                    codigosDuplicados += linea3(p) & vbNewLine
                                End If
                            End If
                        End If
                    Next p
                Next h
                If Not codigosDuplicados = "" Then
                    mensajeFinal += "Los siguientes códigos aparecen más de una vez: " & vbNewLine & codigosDuplicados & vbNewLine & vbNewLine
                End If
            End If

            If loteOK = True Then
                For h = 0 To UBound(linea3)
                    If linea3(h).Contains(".RC") Or
                        linea3(h).Contains("PAN13K/") Or
                        (linea3(h).Contains("RCPANMI/") Or linea3(h).Contains("RCPAN/")) Or
                        linea3(h).Contains("RF13K/") Then
                    Else
                        loteOK = False
                        If Not errComp.Contains(linea3(h)) Then
                            errComp += linea3(h) & vbNewLine
                        End If
                    End If
                Next h
                If Not errComp = "" Then
                    mensajeFinal += "Los siguientes códigos no deben ir en un paquete de proyector 13K: " & vbNewLine & errComp & vbNewLine & vbNewLine
                End If
            End If
            If loteOK = True Then
                For h = 0 To UBound(linea3)
                    If linea3(h).Contains(codigoLote + ".RC") Then
                        paquetesQueHay += 1
                    End If
                    If linea3(h).Contains("PAN13K/") Then
                        proy13kQueHay += 1
                    End If
                    If linea3(h).Contains("RCPANMI/") Or linea3(h).Contains("RCPAN/") Then
                        mando13kQueHay += 1
                    End If
                    If linea3(h).Contains("RF13K/") Then
                        rentalframe13kQueHay += 1
                    End If
                Next h
                proy13kHayPorPaquete = proy13kQueHay / paquetesQueHay
                mando13kHayPorPaquete = mando13kQueHay / paquetesQueHay
                rentalframe13kHayPorPaquete = rentalframe13kQueHay / paquetesQueHay

                If proy13kHayPorPaquete <> proy13kPorPaquete Then
                    mensajeFinal += "En algún paquete, la cantidad de proyectores de 13K difiere de " & proy13kPorPaquete & vbNewLine
                    loteOK = False
                End If
                If mando13kHayPorPaquete <> mando13kPorPaquete Then
                    mensajeFinal += "En algún paquete, la cantidad de mandos con hdmi de 13K difiere de " & mando13kPorPaquete & vbNewLine
                    loteOK = False
                End If
                If rentalframe13kHayPorPaquete <> rentalframe13kPorPaquete Then
                    mensajeFinal += "En algún paquete, la cantidad de rental frame de 13K difiere de " & rentalframe13kPorPaquete & vbNewLine
                    loteOK = False
                End If

            End If
            comprobacion()
            If errorRC = False And loteOK = True Then
                MsgBox("Después de darle al OK tendrás 5 segundos para seleccionar la ventana 'Pack a Road case'")
                volcado()
            ElseIf errorRC = False And loteOK = False Then
                MsgBox(mensajeFinal & vbNewLine & "Repasa los códigos.", MsgBoxStyle.MsgBoxSetForeground & MsgBoxStyle.Information, "Hay errores")
            End If
        End If

        'PROYECTOR 10K
        If ComboBox1.Text = "10K (2)" Then
            codigoLote = "10K"
            limiteDeLotes = 2
            Dim errComp As String = ""
            Dim h As Integer = 0
            Dim p As Integer = 0
            Dim paquetesQueHay As Integer = 0
            Dim proy10kLista As Integer = 0
            Dim proy10kQueHay As Integer = 0
            Dim proy10kHayPorPaquete As Double = 0
            Dim proy10kPorPaquete As Integer = 1
            Dim mando10kEnLista As Integer = 0
            Dim mando10kQueHay As Integer = 0
            Dim mando10kHayPorPaquete As Double = 0
            Dim mando10kPorPaquete As Integer = 1
            Dim rentalframe10kEnLista As Integer = 0
            Dim rentalframe10kQueHay As Integer = 0
            Dim rentalframe10kHayPorPaquete As Double = 0
            Dim rentalframe10kPorPaquete As Integer = 1
            Dim loteOK As Boolean = True
            Dim mensajeFinal As String = ""
            If Not linea3(0).Contains(".RC") Then
                mensajeFinal += "El primer código del paquete tiene que ser el de la caja" & vbNewLine & vbNewLine
                loteOK = False
            End If
            If loteOK = True Then
                For h = 0 To UBound(linea3)
                    For p = 0 To UBound(linea3)
                        If Not p = h Then
                            If linea3(p) = linea3(h) Then
                                loteOK = False
                                If Not codigosDuplicados.Contains(linea3(p)) Then
                                    codigosDuplicados += linea3(p) & vbNewLine
                                End If
                            End If
                        End If
                    Next p
                Next h
                If Not codigosDuplicados = "" Then
                    mensajeFinal += "Los siguientes códigos aparecen más de una vez: " & vbNewLine & codigosDuplicados & vbNewLine & vbNewLine
                End If
            End If

            If loteOK = True Then
                For h = 0 To UBound(linea3)
                    If linea3(h).Contains(".RC") Or
                        linea3(h).Contains("PAN10K/") Or
                        (linea3(h).Contains("RCPANMI/") Or linea3(h).Contains("RCPAN/")) Or
                        linea3(h).Contains("RF10K/") Then
                    Else
                        loteOK = False
                        If Not errComp.Contains(linea3(h)) Then
                            errComp += linea3(h) & vbNewLine
                        End If
                    End If
                Next h
                If Not errComp = "" Then
                    mensajeFinal += "Los siguientes códigos no deben ir en un paquete de proyector 10K: " & vbNewLine & errComp & vbNewLine & vbNewLine
                End If
            End If
            If loteOK = True Then
                For h = 0 To UBound(linea3)
                    If linea3(h).Contains(codigoLote + ".RC") Then
                        paquetesQueHay += 1
                    End If
                    If linea3(h).Contains("PAN10K/") Then
                        proy10kQueHay += 1
                    End If
                    If linea3(h).Contains("RCPANMI/") Or linea3(h).Contains("RCPAN/") Then
                        mando10kQueHay += 1
                    End If
                    If linea3(h).Contains("RF10K/") Then
                        rentalframe10kQueHay += 1
                    End If
                Next h
                proy10kHayPorPaquete = proy10kQueHay / paquetesQueHay
                mando10kHayPorPaquete = mando10kQueHay / paquetesQueHay
                rentalframe10kHayPorPaquete = rentalframe10kQueHay / paquetesQueHay

                If proy10kHayPorPaquete <> proy10kPorPaquete Then
                    mensajeFinal += "En algún paquete, la cantidad de proyectores de 10K difiere de " & proy10kPorPaquete & vbNewLine
                    loteOK = False
                End If
                If mando10kHayPorPaquete <> mando10kPorPaquete Then
                    mensajeFinal += "En algún paquete, la cantidad de mandos de hdmi de 10K difiere de " & mando10kPorPaquete & vbNewLine
                    loteOK = False
                End If
                If rentalframe10kHayPorPaquete <> rentalframe10kPorPaquete Then
                    mensajeFinal += "En algún paquete, la cantidad de rental frame de 10K difiere de " & rentalframe10kPorPaquete & vbNewLine
                    loteOK = False
                End If

            End If
            comprobacion()
            If errorRC = False And loteOK = True Then
                MsgBox("Después de darle al OK tendrás 5 segundos para seleccionar la ventana 'Pack a Road case'")
                volcado()
            ElseIf errorRC = False And loteOK = False Then
                MsgBox(mensajeFinal & vbNewLine & "Repasa los códigos.", MsgBoxStyle.MsgBoxSetForeground & MsgBoxStyle.Information, "Hay errores")
            End If
        End If

        'PROYECTOR 8K
        If ComboBox1.Text = "8.5K (2)" Then
            codigoLote = "8K"
            limiteDeLotes = 2
            Dim errComp As String = ""
            Dim h As Integer = 0
            Dim p As Integer = 0
            Dim paquetesQueHay As Integer = 0
            Dim proy8kLista As Integer = 0
            Dim proy8kQueHay As Integer = 0
            Dim proy8kHayPorPaquete As Double = 0
            Dim proy8kPorPaquete As Integer = 1
            Dim mando8kEnLista As Integer = 0
            Dim mando8kQueHay As Integer = 0
            Dim mando8kHayPorPaquete As Double = 0
            Dim mando8kPorPaquete As Integer = 1
            Dim rentalframe8kEnLista As Integer = 0
            Dim rentalframe8kQueHay As Integer = 0
            Dim rentalframe8kHayPorPaquete As Double = 0
            Dim rentalframe8kPorPaquete As Integer = 1
            Dim loteOK As Boolean = True
            Dim mensajeFinal As String = ""
            If Not linea3(0).Contains(".RC") Then
                mensajeFinal += "El primer código del paquete tiene que ser el de la caja" & vbNewLine & vbNewLine
                loteOK = False
            End If
            If loteOK = True Then
                For h = 0 To UBound(linea3)
                    For p = 0 To UBound(linea3)
                        If Not p = h Then
                            If linea3(p) = linea3(h) Then
                                loteOK = False
                                If Not codigosDuplicados.Contains(linea3(p)) Then
                                    codigosDuplicados += linea3(p) & vbNewLine
                                End If
                            End If
                        End If
                    Next p
                Next h
                If Not codigosDuplicados = "" Then
                    mensajeFinal += "Los siguientes códigos aparecen más de una vez: " & vbNewLine & codigosDuplicados & vbNewLine & vbNewLine
                End If
            End If

            If loteOK = True Then
                For h = 0 To UBound(linea3)
                    If linea3(h).Contains(".RC") Or
                        linea3(h).Contains("PT870/") Or
                        (linea3(h).Contains("RCPANMI/") Or linea3(h).Contains("RCPAN/")) Or
                        linea3(h).Contains("RF8K/") Then
                    Else
                        loteOK = False
                        If Not errComp.Contains(linea3(h)) Then
                            errComp += linea3(h) & vbNewLine
                        End If
                    End If
                Next h
                If Not errComp = "" Then
                    mensajeFinal += "Los siguientes códigos no deben ir en un paquete de proyector 8K: " & vbNewLine & errComp & vbNewLine & vbNewLine
                End If
            End If
            If loteOK = True Then
                For h = 0 To UBound(linea3)
                    If linea3(h).Contains(codigoLote + ".RC") Then
                        paquetesQueHay += 1
                    End If
                    If linea3(h).Contains("PT870/") Then
                        proy8kQueHay += 1
                    End If
                    If linea3(h).Contains("RCPANMI/") Or linea3(h).Contains("RCPAN/") Then
                        mando8kQueHay += 1
                    End If
                    If linea3(h).Contains("RF8K/") Then
                        rentalframe8kQueHay += 1
                    End If
                Next h
                proy8kHayPorPaquete = proy8kQueHay / paquetesQueHay
                mando8kHayPorPaquete = mando8kQueHay / paquetesQueHay
                rentalframe8kHayPorPaquete = rentalframe8kQueHay / paquetesQueHay

                If proy8kHayPorPaquete <> proy8kPorPaquete Then
                    mensajeFinal += "En algún paquete, la cantidad de proyectores de 8K difiere de " & proy8kPorPaquete & vbNewLine
                    loteOK = False
                End If
                If mando8kHayPorPaquete <> mando8kPorPaquete Then
                    mensajeFinal += "En algún paquete, la cantidad de mandos con hdmi de 8K difiere de " & mando8kPorPaquete & vbNewLine
                    loteOK = False
                End If
                If rentalframe8kHayPorPaquete <> rentalframe8kPorPaquete Then
                    mensajeFinal += "En algún paquete, la cantidad de rental frame de 8K difiere de " & rentalframe8kPorPaquete & vbNewLine
                    loteOK = False
                End If

            End If
            comprobacion()
            If errorRC = False And loteOK = True Then
                MsgBox("Después de darle al OK tendrás 5 segundos para seleccionar la ventana 'Pack a Road case'")
                volcado()
            ElseIf errorRC = False And loteOK = False Then
                MsgBox(mensajeFinal & vbNewLine & "Repasa los códigos.", MsgBoxStyle.MsgBoxSetForeground & MsgBoxStyle.Information, "Hay errores")
            End If
        End If



        'DISPLAY 85 MATE O BRILLO"
        If ComboBox1.Text = "85"" Mate o Brillo (4)" Then
            codigoLote = "85"
            limiteDeLotes = 4
            Dim errComp As String = ""
            Dim h As Integer = 0
            Dim p As Integer = 0
            Dim paquetesQueHay As Integer = 0
            Dim disp85EnLista As Integer = 0
            Dim disp85QueHay As Integer = 0
            Dim disp85HayPorPaquete As Double = 0
            Dim disp85PorPaquete As Integer = 1
            Dim mando85EnLista As Integer = 0
            Dim mando85QueHay As Integer = 0
            Dim mando85HayPorPaquete As Double = 0
            Dim mando85PorPaquete As Integer = 1
            Dim loteOK As Boolean = True
            Dim mensajeFinal As String = ""
            If Not linea3(0).Contains(".RC") Then
                mensajeFinal += "El primer código del paquete tiene que ser el de la caja" & vbNewLine & vbNewLine
                loteOK = False
            End If
            If loteOK = True Then
                For h = 0 To UBound(linea3)
                    For p = 0 To UBound(linea3)
                        If Not p = h Then
                            If linea3(p) = linea3(h) Then
                                loteOK = False
                                If Not codigosDuplicados.Contains(linea3(p)) Then
                                    codigosDuplicados += linea3(p) & vbNewLine
                                End If
                            End If
                        End If
                    Next p
                Next h
                If Not codigosDuplicados = "" Then
                    mensajeFinal += "Los siguientes códigos aparecen más de una vez " & vbNewLine & codigosDuplicados & vbNewLine & vbNewLine
                End If
            End If

            If loteOK = True Then
                For h = 0 To UBound(linea3)
                    If linea3(h).Contains(".RC") Or
                        linea3(h).Contains("QM85D/") Or
                        linea3(h).Contains("QM85NMAT/") Or
                        linea3(h).Contains("SRC1/") Then
                    Else
                        loteOK = False
                        If Not errComp.Contains(linea3(h)) Then
                            errComp += linea3(h) & vbNewLine
                        End If
                    End If
                Next h
                If Not errComp = "" Then
                    mensajeFinal += "Los siguientes códigos no deben ir en un paquete de display de 85 pulgadas " & vbNewLine & errComp & vbNewLine & vbNewLine
                End If
            End If
            If loteOK = True Then
                For h = 0 To UBound(linea3)
                    If linea3(h).Contains(codigoLote + ".RC") Then
                        paquetesQueHay += 1
                    End If
                    If linea3(h).Contains("QM85D/") Or linea3(h).Contains("QM85NMAT/") Then
                        disp85QueHay += 1
                    End If
                    If linea3(h).Contains("SRC1/") Then
                        mando85QueHay += 1
                    End If
                Next h
                disp85HayPorPaquete = disp85QueHay / paquetesQueHay
                mando85HayPorPaquete = mando85QueHay / paquetesQueHay
                If disp85HayPorPaquete <> disp85PorPaquete Then
                    mensajeFinal += "En algún paquete, la cantidad de displays de 85 pulgadas difiere de " & disp85PorPaquete & vbNewLine
                    loteOK = False
                End If
                If mando85HayPorPaquete <> mando85PorPaquete Then
                    mensajeFinal += "En algún paquete, la cantidad de mandos a distancia para display de 85 difiere de " & mando85PorPaquete & vbNewLine
                    loteOK = False
                End If
            End If
            comprobacion()
            If errorRC = False And loteOK = True Then
                MsgBox("Después de darle al OK tendrás 5 segundos para seleccionar la ventana 'Pack a Road case'")
                volcado()
            ElseIf errorRC = False And loteOK = False Then
                MsgBox(mensajeFinal & vbNewLine & "Repasa los códigos.", MsgBoxStyle.MsgBoxSetForeground & MsgBoxStyle.Information, "Hay errores")
            End If
        End If

        'DISPLAY 80"
        If ComboBox1.Text = "80"" (4)" Then
            codigoLote = "80"
            limiteDeLotes = 4
            Dim errComp As String = ""
            Dim h As Integer = 0
            Dim p As Integer = 0
            Dim paquetesQueHay As Integer = 0
            Dim disp80EnLista As Integer = 0
            Dim disp80QueHay As Integer = 0
            Dim disp80HayPorPaquete As Double = 0
            Dim disp80PorPaquete As Integer = 1
            Dim mando80EnLista As Integer = 0
            Dim mando80QueHay As Integer = 0
            Dim mando80HayPorPaquete As Double = 0
            Dim mando80PorPaquete As Integer = 1
            Dim soporte80EnLista As Integer = 0
            Dim soporte80QueHay As Integer = 0
            Dim soporte80HayPorPaquete As Double = 0
            Dim soporte80PorPaquete As Integer = 1
            Dim allen80EnLista As Integer = 0
            Dim allen80QueHay As Integer = 0
            Dim allen80HayPorPaquete As Double = 0
            Dim allen80PorPaquete As Integer = 1
            Dim loteOK As Boolean = True
            Dim mensajeFinal As String = ""
            If Not linea3(0).Contains(".RC") Then
                mensajeFinal += "El primer código del paquete tiene que ser el de la caja" & vbNewLine & vbNewLine
                loteOK = False
            End If
            If loteOK = True Then
                For h = 0 To UBound(linea3)
                    For p = 0 To UBound(linea3)
                        If Not p = h Then
                            If linea3(p) = linea3(h) Then
                                loteOK = False
                                If Not codigosDuplicados.Contains(linea3(p)) Then
                                    codigosDuplicados += linea3(p) & vbNewLine
                                End If
                            End If
                        End If
                    Next p
                Next h
                If Not codigosDuplicados = "" Then
                    mensajeFinal += "Los siguientes códigos aparecen más de una vez: " & vbNewLine & codigosDuplicados & vbNewLine & vbNewLine
                End If
            End If

            If loteOK = True Then
                For h = 0 To UBound(linea3)
                    If linea3(h).Contains(".RC") Or
                        linea3(h).Contains("SHARP80/") Or
                        linea3(h).Contains("HRC/") Or
                        linea3(h).Contains("80KEY/") Or
                        linea3(h).Contains("SOP80/") Then

                    Else
                        loteOK = False
                        If Not errComp.Contains(linea3(h)) Then
                            errComp += linea3(h) & vbNewLine
                        End If
                    End If
                Next h
                If Not errComp = "" Then
                    mensajeFinal += "Los siguientes códigos no deben ir en un paquete de display de 80 pulgadas: " & vbNewLine & errComp & vbNewLine & vbNewLine
                End If
            End If
            If loteOK = True Then
                For h = 0 To UBound(linea3)
                    If linea3(h).Contains(codigoLote + ".RC") Then
                        paquetesQueHay += 1
                    End If
                    If linea3(h).Contains("SHARP80/") Then
                        disp80QueHay += 1
                    End If
                    If linea3(h).Contains("HRC/") Then
                        mando80QueHay += 1
                    End If
                    If linea3(h).Contains("SOP80/") Then
                        soporte80QueHay += 1
                    End If
                    If linea3(h).Contains("80KEY/") Then
                        allen80QueHay += 1
                    End If
                Next h
                disp80HayPorPaquete = disp80QueHay / paquetesQueHay
                mando80HayPorPaquete = mando80QueHay / paquetesQueHay
                soporte80HayPorPaquete = soporte80QueHay / paquetesQueHay
                allen80HayPorPaquete = allen80QueHay / paquetesQueHay
                If disp80HayPorPaquete <> disp80PorPaquete Then
                    mensajeFinal += "En algún paquete, la cantidad de displays de 80 pulgadas difiere de " & disp80PorPaquete & vbNewLine
                    loteOK = False
                End If
                If mando80HayPorPaquete <> mando80PorPaquete Then
                    mensajeFinal += "En algún paquete, la cantidad de mandos a distancia para display de 80 difiere de " & mando80PorPaquete & vbNewLine
                    loteOK = False
                End If
                If soporte80HayPorPaquete <> soporte80PorPaquete Then
                    mensajeFinal += "En algún paquete, la cantidad de soportes para display de 80 pulgadas difiere de " & soporte80PorPaquete & vbNewLine
                    loteOK = False
                End If
                If allen80HayPorPaquete <> soporte80PorPaquete Then
                    mensajeFinal += "En algún paquete, la cantidad de llaves allen para el soporte de display de 80 pulgadas difiere de " & allen80PorPaquete & vbNewLine
                    loteOK = False
                End If
            End If
            comprobacion()
            If errorRC = False And loteOK = True Then
                MsgBox("Después de darle al OK tendrás 5 segundos para seleccionar la ventana 'Pack a Road case'")
                volcado()
            ElseIf errorRC = False And loteOK = False Then
                MsgBox(mensajeFinal & vbNewLine & "Repasa los códigos.", MsgBoxStyle.MsgBoxSetForeground & MsgBoxStyle.Information, "Hay errores")
            End If
        End If

        'DISPLAY DE 1 x 65"
        If ComboBox1.Text = "1 x 65"" (4)" Then
            codigoLote = "1U65"
            limiteDeLotes = 4
            Dim errComp As String = ""
            Dim h As Integer = 0
            Dim p As Integer = 0
            Dim paquetesQueHay As Integer = 0
            Dim disp65EnLista As Integer = 0
            Dim disp65QueHay As Integer = 0
            Dim disp65HayPorPaquete As Double = 0
            Dim disp65PorPaquete As Integer = 1
            Dim mando65EnLista As Integer = 0
            Dim mando65QueHay As Integer = 0
            Dim mando65HayPorPaquete As Double = 0
            Dim mando65PorPaquete As Integer = 1
            Dim loteOK As Boolean = True
            Dim mensajeFinal As String = ""
            If Not linea3(0).Contains(".RC") Then
                mensajeFinal += "El primer código del paquete tiene que ser el de la caja" & vbNewLine & vbNewLine
                loteOK = False
            End If
            If loteOK = True Then
                For h = 0 To UBound(linea3)
                    For p = 0 To UBound(linea3)
                        If Not p = h Then
                            If linea3(p) = linea3(h) Then
                                loteOK = False
                                If Not codigosDuplicados.Contains(linea3(p)) Then
                                    codigosDuplicados += linea3(p) & vbNewLine
                                End If
                            End If
                        End If
                    Next p
                Next h
                If Not codigosDuplicados = "" Then
                    mensajeFinal += "Los siguientes códigos aparecen más de una vez: " & vbNewLine & codigosDuplicados & vbNewLine & vbNewLine
                End If
            End If

            If loteOK = True Then
                For h = 0 To UBound(linea3)
                    If linea3(h).Contains(".RC") Or
                        linea3(h).Contains("QM65H/") Or
                        linea3(h).Contains("SRC1/") Then
                    Else
                        loteOK = False
                        If Not errComp.Contains(linea3(h)) Then
                            errComp += linea3(h) & vbNewLine
                        End If
                    End If
                Next h
                If Not errComp = "" Then
                    mensajeFinal += "Los siguientes códigos no deben ir en un paquete de 1 display de 65 pulgadas: " & vbNewLine & errComp & vbNewLine & vbNewLine
                End If
            End If
            If loteOK = True Then
                For h = 0 To UBound(linea3)
                    If linea3(h).Contains(codigoLote + ".RC") Then
                        paquetesQueHay += 1
                    End If
                    If linea3(h).Contains("QM65H/") Then
                        disp65QueHay += 1
                    End If
                    If linea3(h).Contains("SRC1/") Then
                        mando65QueHay += 1
                    End If
                Next h
                disp65HayPorPaquete = disp65QueHay / paquetesQueHay
                mando65HayPorPaquete = mando65QueHay / paquetesQueHay
                If disp65HayPorPaquete <> disp65PorPaquete Then
                    mensajeFinal += "En algún paquete, la cantidad de displays de 65 pulgadas difiere de " & disp65PorPaquete & vbNewLine
                    loteOK = False
                End If
                If mando65HayPorPaquete <> mando65PorPaquete Then
                    mensajeFinal += "En algún paquete, la cantidad de mandos a distancia para display de 65 difiere de " & mando65PorPaquete & vbNewLine
                    loteOK = False
                End If
            End If
            comprobacion()
            If errorRC = False And loteOK = True Then
                MsgBox("Después de darle al OK tendrás 5 segundos para seleccionar la ventana 'Pack a Road case'")
                volcado()
            ElseIf errorRC = False And loteOK = False Then
                MsgBox(mensajeFinal & vbNewLine & "Repasa los códigos.", MsgBoxStyle.MsgBoxSetForeground & MsgBoxStyle.Information, "Hay errores")
            End If
        End If

        'DISPLAY DE 2 x 65"
        If ComboBox1.Text = "2 x 65"" (4)" Then
            codigoLote = "2U65"
            limiteDeLotes = 4
            Dim errComp As String = ""
            Dim h As Integer = 0
            Dim p As Integer = 0
            Dim paquetesQueHay As Integer = 0
            Dim disp65EnLista As Integer = 0
            Dim disp65QueHay As Integer = 0
            Dim disp65HayPorPaquete As Double = 0
            Dim disp65PorPaquete As Integer = 2
            Dim mando65EnLista As Integer = 0
            Dim mando65QueHay As Integer = 0
            Dim mando65HayPorPaquete As Double = 0
            Dim mando65PorPaquete As Integer = 2
            Dim loteOK As Boolean = True
            Dim mensajeFinal As String = ""
            If Not linea3(0).Contains(".RC") Then
                mensajeFinal += "El primer código del paquete tiene que ser el de la caja" & vbNewLine & vbNewLine
                loteOK = False
            End If
            If loteOK = True Then
                For h = 0 To UBound(linea3)
                    For p = 0 To UBound(linea3)
                        If Not p = h Then
                            If linea3(p) = linea3(h) Then
                                loteOK = False
                                If Not codigosDuplicados.Contains(linea3(p)) Then
                                    codigosDuplicados += linea3(p) & vbNewLine
                                End If
                            End If
                        End If
                    Next p
                Next h
                If Not codigosDuplicados = "" Then
                    mensajeFinal += "Los siguientes códigos aparecen más de una vez: " & vbNewLine & codigosDuplicados & vbNewLine & vbNewLine
                End If
            End If

            If loteOK = True Then
                For h = 0 To UBound(linea3)
                    If linea3(h).Contains(".RC") Or
                        linea3(h).Contains("QM65H/") Or
                        linea3(h).Contains("SRC1/") Then
                    Else
                        loteOK = False
                        If Not errComp.Contains(linea3(h)) Then
                            errComp += linea3(h) & vbNewLine
                        End If
                    End If
                Next h
                If Not errComp = "" Then
                    mensajeFinal += "Los siguientes códigos no deben ir en un paquete de 1 display de 65 pulgadas: " & vbNewLine & errComp & vbNewLine & vbNewLine
                End If
            End If
            If loteOK = True Then
                For h = 0 To UBound(linea3)
                    If linea3(h).Contains(codigoLote + ".RC") Then
                        paquetesQueHay += 1
                    End If
                    If linea3(h).Contains("QM65H/") Then
                        disp65QueHay += 1
                    End If
                    If linea3(h).Contains("SRC1/") Then
                        mando65QueHay += 1
                    End If
                Next h
                disp65HayPorPaquete = disp65QueHay / paquetesQueHay
                mando65HayPorPaquete = mando65QueHay / paquetesQueHay
                If disp65HayPorPaquete <> disp65PorPaquete Then
                    mensajeFinal += "En algún paquete, la cantidad de displays de 65 pulgadas difiere de " & disp65PorPaquete & vbNewLine
                    loteOK = False
                End If
                If mando65HayPorPaquete <> mando65PorPaquete Then
                    mensajeFinal += "En algún paquete, la cantidad de mandos a distancia para display de 65 difiere de " & mando65PorPaquete & vbNewLine
                    loteOK = False
                End If
            End If
            comprobacion()
            If errorRC = False And loteOK = True Then
                MsgBox("Después de darle al OK tendrás 5 segundos para seleccionar la ventana 'Pack a Road case'")
                volcado()
            ElseIf errorRC = False And loteOK = False Then
                MsgBox(mensajeFinal & vbNewLine & "Repasa los códigos.", MsgBoxStyle.MsgBoxSetForeground & MsgBoxStyle.Information, "Hay errores")
            End If
        End If

        'DISPLAY DE 3 x 65"
        If ComboBox1.Text = "3 x 65"" (4)" Then
            codigoLote = "3U65"
            limiteDeLotes = 4
            Dim errComp As String = ""
            Dim h As Integer = 0
            Dim p As Integer = 0
            Dim paquetesQueHay As Integer = 0
            Dim disp65EnLista As Integer = 0
            Dim disp65QueHay As Integer = 0
            Dim disp65HayPorPaquete As Double = 0
            Dim disp65PorPaquete As Integer = 3
            Dim mando65EnLista As Integer = 0
            Dim mando65QueHay As Integer = 0
            Dim mando65HayPorPaquete As Double = 0
            Dim mando65PorPaquete As Integer = 1
            Dim loteOK As Boolean = True
            Dim mensajeFinal As String = ""
            If Not linea3(0).Contains(".RC") Then
                mensajeFinal += "El primer código del paquete tiene que ser el de la caja" & vbNewLine & vbNewLine
                loteOK = False
            End If
            If loteOK = True Then
                For h = 0 To UBound(linea3)
                    For p = 0 To UBound(linea3)
                        If Not p = h Then
                            If linea3(p) = linea3(h) Then
                                loteOK = False
                                If Not codigosDuplicados.Contains(linea3(p)) Then
                                    codigosDuplicados += linea3(p) & vbNewLine
                                End If
                            End If
                        End If
                    Next p
                Next h
                If Not codigosDuplicados = "" Then
                    mensajeFinal += "Los siguientes códigos aparecen más de una vez: " & vbNewLine & codigosDuplicados & vbNewLine & vbNewLine
                End If
            End If

            If loteOK = True Then
                For h = 0 To UBound(linea3)
                    If linea3(h).Contains(".RC") Or
                        linea3(h).Contains("QM65H/") Or
                        linea3(h).Contains("SRC1/") Then
                    Else
                        loteOK = False
                        If Not errComp.Contains(linea3(h)) Then
                            errComp += linea3(h) & vbNewLine
                        End If
                    End If
                Next h
                If Not errComp = "" Then
                    mensajeFinal += "Los siguientes códigos no deben ir en un paquete de 1 display de 65 pulgadas: " & vbNewLine & errComp & vbNewLine & vbNewLine
                End If
            End If
            If loteOK = True Then
                For h = 0 To UBound(linea3)
                    If linea3(h).Contains(codigoLote + ".RC") Then
                        paquetesQueHay += 1
                    End If
                    If linea3(h).Contains("QM65H/") Then
                        disp65QueHay += 1
                    End If
                    If linea3(h).Contains("SRC1/") Then
                        mando65QueHay += 1
                    End If
                Next h
                disp65HayPorPaquete = disp65QueHay / paquetesQueHay
                mando65HayPorPaquete = mando65QueHay / paquetesQueHay
                If disp65HayPorPaquete <> disp65PorPaquete Then
                    mensajeFinal += "En algún paquete, la cantidad de displays de 65 pulgadas difiere de " & disp65PorPaquete & vbNewLine
                    loteOK = False
                End If
                If mando65HayPorPaquete <> mando65PorPaquete Then
                    mensajeFinal += "En algún paquete, la cantidad de mandos a distancia para display de 65 difiere de " & mando65PorPaquete & vbNewLine
                    loteOK = False
                End If
            End If
            comprobacion()
            If errorRC = False And loteOK = True Then
                MsgBox("Después de darle al OK tendrás 5 segundos para seleccionar la ventana 'Pack a Road case'")
                volcado()
            ElseIf errorRC = False And loteOK = False Then
                MsgBox(mensajeFinal & vbNewLine & "Repasa los códigos.", MsgBoxStyle.MsgBoxSetForeground & MsgBoxStyle.Information, "Hay errores")
            End If
        End If

        'DISPLAY DE 2 x 55" 4K QM55H
        If ComboBox1.Text = "2 x 55"" 4K QM55H (4)" Then
            codigoLote = "2QM55H"
            limiteDeLotes = 4
            Dim errComp As String = ""
            Dim h As Integer = 0
            Dim p As Integer = 0
            Dim paquetesQueHay As Integer = 0
            Dim dispEnLista As Integer = 0
            Dim dispQueHay As Integer = 0
            Dim dispHayPorPaquete As Double = 0
            Dim dispPorPaquete As Integer = 2
            Dim mandoEnLista As Integer = 0
            Dim mandoQueHay As Integer = 0
            Dim mandoHayPorPaquete As Double = 0
            Dim mandoPorPaquete As Integer = 2
            Dim loteOK As Boolean = True
            Dim mensajeFinal As String = ""
            If Not linea3(0).Contains(".RC") Then
                mensajeFinal += "El primer código del paquete tiene que ser el de la caja" & vbNewLine & vbNewLine
                loteOK = False
            End If
            If loteOK = True Then
                For h = 0 To UBound(linea3)
                    For p = 0 To UBound(linea3)
                        If Not p = h Then
                            If linea3(p) = linea3(h) Then
                                loteOK = False
                                If Not codigosDuplicados.Contains(linea3(p)) Then
                                    codigosDuplicados += linea3(p) & vbNewLine
                                End If
                            End If
                        End If
                    Next p
                Next h
                If Not codigosDuplicados = "" Then
                    mensajeFinal += "Los siguientes códigos aparecen más de una vez: " & vbNewLine & codigosDuplicados & vbNewLine & vbNewLine
                End If
            End If

            If loteOK = True Then
                For h = 0 To UBound(linea3)
                    If linea3(h).Contains(".RC") Or
                        linea3(h).Contains("QM55H/") Or
                        linea3(h).Contains("SRC1/") Then
                    Else
                        loteOK = False
                        If Not errComp.Contains(linea3(h)) Then
                            errComp += linea3(h) & vbNewLine
                        End If
                    End If
                Next h
                If Not errComp = "" Then
                    mensajeFinal += "Los siguientes códigos no deben ir en un paquete de 2 display de 55 pulgadas modelo QM55H: " & vbNewLine & errComp & vbNewLine & vbNewLine
                End If
            End If
            If loteOK = True Then
                For h = 0 To UBound(linea3)
                    If linea3(h).Contains(codigoLote + ".RC") Then
                        paquetesQueHay += 1
                    End If
                    If linea3(h).Contains("QM55H/") Then
                        dispQueHay += 1
                    End If
                    If linea3(h).Contains("SRC1/") Then
                        mandoQueHay += 1
                    End If
                Next h
                dispHayPorPaquete = dispQueHay / paquetesQueHay
                mandoHayPorPaquete = mandoQueHay / paquetesQueHay
                If dispHayPorPaquete <> dispPorPaquete Then
                    mensajeFinal += "En algún paquete, la cantidad de displays de 55 pulgadas modelo QM55H difiere de " & dispPorPaquete & vbNewLine
                    loteOK = False
                End If
                If mandoHayPorPaquete <> mandoPorPaquete Then
                    mensajeFinal += "En algún paquete, la cantidad de mandos a distancia para display de 55 pulgadas modelo QM55H difiere de " & mandoPorPaquete & vbNewLine
                    loteOK = False
                End If
            End If
            comprobacion()
            If errorRC = False And loteOK = True Then
                MsgBox("Después de darle al OK tendrás 5 segundos para seleccionar la ventana 'Pack a Road case'")
                volcado()
            ElseIf errorRC = False And loteOK = False Then
                MsgBox(mensajeFinal & vbNewLine & "Repasa los códigos.", MsgBoxStyle.MsgBoxSetForeground & MsgBoxStyle.Information, "Hay errores")
            End If
        End If

        'DISPLAY DE 3 x 55" 4K QM55H
        If ComboBox1.Text = "3 x 55"" 4K QM55H (4)" Then
            codigoLote = "3QM55H"
            limiteDeLotes = 4
            Dim errComp As String = ""
            Dim h As Integer = 0
            Dim p As Integer = 0
            Dim paquetesQueHay As Integer = 0
            Dim dispEnLista As Integer = 0
            Dim dispQueHay As Integer = 0
            Dim dispHayPorPaquete As Double = 0
            Dim dispPorPaquete As Integer = 3
            Dim mandoEnLista As Integer = 0
            Dim mandoQueHay As Integer = 0
            Dim mandoHayPorPaquete As Double = 0
            Dim mandoPorPaquete As Integer = 1
            Dim loteOK As Boolean = True
            Dim mensajeFinal As String = ""
            If Not linea3(0).Contains(".RC") Then
                mensajeFinal += "El primer código del paquete tiene que ser el de la caja" & vbNewLine & vbNewLine
                loteOK = False
            End If
            If loteOK = True Then
                For h = 0 To UBound(linea3)
                    For p = 0 To UBound(linea3)
                        If Not p = h Then
                            If linea3(p) = linea3(h) Then
                                loteOK = False
                                If Not codigosDuplicados.Contains(linea3(p)) Then
                                    codigosDuplicados += linea3(p) & vbNewLine
                                End If
                            End If
                        End If
                    Next p
                Next h
                If Not codigosDuplicados = "" Then
                    mensajeFinal += "Los siguientes códigos aparecen más de una vez: " & vbNewLine & codigosDuplicados & vbNewLine & vbNewLine
                End If
            End If

            If loteOK = True Then
                For h = 0 To UBound(linea3)
                    If linea3(h).Contains(".RC") Or
                        linea3(h).Contains("QM55H/") Or
                        linea3(h).Contains("SRC1/") Then
                    Else
                        loteOK = False
                        If Not errComp.Contains(linea3(h)) Then
                            errComp += linea3(h) & vbNewLine
                        End If
                    End If
                Next h
                If Not errComp = "" Then
                    mensajeFinal += "Los siguientes códigos no deben ir en un paquete de 2 display de 55 pulgadas modelo QM55H: " & vbNewLine & errComp & vbNewLine & vbNewLine
                End If
            End If
            If loteOK = True Then
                For h = 0 To UBound(linea3)
                    If linea3(h).Contains(codigoLote + ".RC") Then
                        paquetesQueHay += 1
                    End If
                    If linea3(h).Contains("QM55H/") Then
                        dispQueHay += 1
                    End If
                    If linea3(h).Contains("SRC1/") Then
                        mandoQueHay += 1
                    End If
                Next h
                dispHayPorPaquete = dispQueHay / paquetesQueHay
                mandoHayPorPaquete = mandoQueHay / paquetesQueHay
                If dispHayPorPaquete <> dispPorPaquete Then
                    mensajeFinal += "En algún paquete, la cantidad de displays de 55 pulgadas modelo QM55H difiere de " & dispPorPaquete & vbNewLine
                    loteOK = False
                End If
                If mandoHayPorPaquete <> mandoPorPaquete Then
                    mensajeFinal += "En algún paquete, la cantidad de mandos a distancia para display de 55 pulgadas modelo QM55H difiere de " & mandoPorPaquete & vbNewLine
                    loteOK = False
                End If
            End If
            comprobacion()
            If errorRC = False And loteOK = True Then
                MsgBox("Después de darle al OK tendrás 5 segundos para seleccionar la ventana 'Pack a Road case'")
                volcado()
            ElseIf errorRC = False And loteOK = False Then
                MsgBox(mensajeFinal & vbNewLine & "Repasa los códigos.", MsgBoxStyle.MsgBoxSetForeground & MsgBoxStyle.Information, "Hay errores")
            End If
        End If

        'DISPLAY DE 2 x 55" ME55C
        If ComboBox1.Text = "2 x 55"" ME55C (4)" Then
            codigoLote = "ME55C"
            limiteDeLotes = 4
            Dim errComp As String = ""
            Dim h As Integer = 0
            Dim p As Integer = 0
            Dim paquetesQueHay As Integer = 0
            Dim dispME55CEnLista As Integer = 0
            Dim dispME55CQueHay As Integer = 0
            Dim dispME55CHayPorPaquete As Double = 0
            Dim dispME55CPorPaquete As Integer = 2
            Dim mandoME55CEnLista As Integer = 0
            Dim mandoME55CQueHay As Integer = 0
            Dim mandoME55CHayPorPaquete As Double = 0
            Dim mandoME55CPorPaquete As Integer = 2
            Dim loteOK As Boolean = True
            Dim mensajeFinal As String = ""
            If Not linea3(0).Contains(".RC") Then
                mensajeFinal += "El primer código del paquete tiene que ser el de la caja" & vbNewLine & vbNewLine
                loteOK = False
            End If
            If loteOK = True Then
                For h = 0 To UBound(linea3)
                    For p = 0 To UBound(linea3)
                        If Not p = h Then
                            If linea3(p) = linea3(h) Then
                                loteOK = False
                                If Not codigosDuplicados.Contains(linea3(p)) Then
                                    codigosDuplicados += linea3(p) & vbNewLine
                                End If
                            End If
                        End If
                    Next p
                Next h
                If Not codigosDuplicados = "" Then
                    mensajeFinal += "Los siguientes códigos aparecen más de una vez: " & vbNewLine & codigosDuplicados & vbNewLine & vbNewLine
                End If
            End If

            If loteOK = True Then
                For h = 0 To UBound(linea3)
                    If linea3(h).Contains(".RC") Or
                        linea3(h).Contains("ME55C/") Or
                        linea3(h).Contains("SRC1/") Then
                    Else
                        loteOK = False
                        If Not errComp.Contains(linea3(h)) Then
                            errComp += linea3(h) & vbNewLine
                        End If
                    End If
                Next h
                If Not errComp = "" Then
                    mensajeFinal += "Los siguientes códigos no deben ir en un paquete de 2 display de 55 pulgadas modelo ME55C: " & vbNewLine & errComp & vbNewLine & vbNewLine
                End If
            End If
            If loteOK = True Then
                For h = 0 To UBound(linea3)
                    If linea3(h).Contains(codigoLote + ".RC") Then
                        paquetesQueHay += 1
                    End If
                    If linea3(h).Contains("ME55C/") Then
                        dispME55CQueHay += 1
                    End If
                    If linea3(h).Contains("SRC1/") Then
                        mandoME55CQueHay += 1
                    End If
                Next h
                dispME55CHayPorPaquete = dispME55CQueHay / paquetesQueHay
                mandoME55CHayPorPaquete = mandoME55CQueHay / paquetesQueHay
                If dispME55CHayPorPaquete <> dispME55CPorPaquete Then
                    mensajeFinal += "En algún paquete, la cantidad de displays de 55 pulgadas modelo ME55C difiere de " & dispME55CPorPaquete & vbNewLine
                    loteOK = False
                End If
                If mandoME55CHayPorPaquete <> mandoME55CPorPaquete Then
                    mensajeFinal += "En algún paquete, la cantidad de mandos a distancia para display de 55 pulgadas modelo ME55C difiere de " & mandoME55CPorPaquete & vbNewLine
                    loteOK = False
                End If
            End If
            comprobacion()
            If errorRC = False And loteOK = True Then
                MsgBox("Después de darle al OK tendrás 5 segundos para seleccionar la ventana 'Pack a Road case'")
                volcado()
            ElseIf errorRC = False And loteOK = False Then
                MsgBox(mensajeFinal & vbNewLine & "Repasa los códigos.", MsgBoxStyle.MsgBoxSetForeground & MsgBoxStyle.Information, "Hay errores")
            End If
        End If

        'DISPLAY DE 2 x 55" DB55D
        If ComboBox1.Text = "2 x 55"" DB55D (4)" Then
            codigoLote = "DB55D"
            limiteDeLotes = 4
            Dim errComp As String = ""
            Dim h As Integer = 0
            Dim p As Integer = 0
            Dim paquetesQueHay As Integer = 0
            Dim dispDB55DEnLista As Integer = 0
            Dim dispDB55DQueHay As Integer = 0
            Dim dispDB55DHayPorPaquete As Double = 0
            Dim dispDB55DPorPaquete As Integer = 2
            Dim mandoDB55DEnLista As Integer = 0
            Dim mandoDB55DQueHay As Integer = 0
            Dim mandoDB55DHayPorPaquete As Double = 0
            Dim mandoDB55DPorPaquete As Integer = 2
            Dim loteOK As Boolean = True
            Dim mensajeFinal As String = ""
            If Not linea3(0).Contains(".RC") Then
                mensajeFinal += "El primer código del paquete tiene que ser el de la caja" & vbNewLine & vbNewLine
                loteOK = False
            End If
            If loteOK = True Then
                For h = 0 To UBound(linea3)
                    For p = 0 To UBound(linea3)
                        If Not p = h Then
                            If linea3(p) = linea3(h) Then
                                loteOK = False
                                If Not codigosDuplicados.Contains(linea3(p)) Then
                                    codigosDuplicados += linea3(p) & vbNewLine
                                End If
                            End If
                        End If
                    Next p
                Next h
                If Not codigosDuplicados = "" Then
                    mensajeFinal += "Los siguientes códigos aparecen más de una vez: " & vbNewLine & codigosDuplicados & vbNewLine & vbNewLine
                End If
            End If

            If loteOK = True Then
                For h = 0 To UBound(linea3)
                    If linea3(h).Contains(".RC") Or
                        linea3(h).Contains("DB55D/") Or
                        linea3(h).Contains("SRC1/") Then
                    Else
                        loteOK = False
                        If Not errComp.Contains(linea3(h)) Then
                            errComp += linea3(h) & vbNewLine
                        End If
                    End If
                Next h
                If Not errComp = "" Then
                    mensajeFinal += "Los siguientes códigos no deben ir en un paquete de 2 display de 55 pulgadas modelo DB55D: " & vbNewLine & errComp & vbNewLine & vbNewLine
                End If
            End If
            If loteOK = True Then
                For h = 0 To UBound(linea3)
                    If linea3(h).Contains(codigoLote + ".RC") Then
                        paquetesQueHay += 1
                    End If
                    If linea3(h).Contains("DB55D/") Then
                        dispDB55DQueHay += 1
                    End If
                    If linea3(h).Contains("SRC1/") Then
                        mandoDB55DQueHay += 1
                    End If
                Next h
                dispDB55DHayPorPaquete = dispDB55DQueHay / paquetesQueHay
                mandoDB55DHayPorPaquete = mandoDB55DQueHay / paquetesQueHay
                If dispDB55DHayPorPaquete <> dispDB55DPorPaquete Then
                    mensajeFinal += "En algún paquete, la cantidad de displays de 55 pulgadas modelo DB55D difiere de " & dispDB55DPorPaquete & vbNewLine
                    loteOK = False
                End If
                If mandoDB55DHayPorPaquete <> mandoDB55DPorPaquete Then
                    mensajeFinal += "En algún paquete, la cantidad de mandos a distancia para display de 55 pulgadas modelo DB55D difiere de " & mandoDB55DPorPaquete & vbNewLine
                    loteOK = False
                End If
            End If
            comprobacion()
            If errorRC = False And loteOK = True Then
                MsgBox("Después de darle al OK tendrás 5 segundos para seleccionar la ventana 'Pack a Road case'")
                volcado()
            ElseIf errorRC = False And loteOK = False Then
                MsgBox(mensajeFinal & vbNewLine & "Repasa los códigos.", MsgBoxStyle.MsgBoxSetForeground & MsgBoxStyle.Information, "Hay errores")
            End If
        End If

        'DISPLAY DE 2 x 49" QB49N
        If ComboBox1.Text = "2 x 49"" QB49N (4)" Then
            codigoLote = "2QB49N"
            limiteDeLotes = 4
            Dim errComp As String = ""
            Dim h As Integer = 0
            Dim p As Integer = 0
            Dim paquetesQueHay As Integer = 0
            Dim dispQB49NEnLista As Integer = 0
            Dim dispQB49NQueHay As Integer = 0
            Dim dispQB49NHayPorPaquete As Double = 0
            Dim dispQB49NPorPaquete As Integer = 2
            Dim mandoQB49NEnLista As Integer = 0
            Dim mandoQB49NQueHay As Integer = 0
            Dim mandoQB49NHayPorPaquete As Double = 0
            Dim mandoQB49NPorPaquete As Integer = 2
            Dim loteOK As Boolean = True
            Dim mensajeFinal As String = ""
            If Not linea3(0).Contains(".RC") Then
                mensajeFinal += "El primer código del paquete tiene que ser el de la caja" & vbNewLine & vbNewLine
                loteOK = False
            End If
            If loteOK = True Then
                For h = 0 To UBound(linea3)
                    For p = 0 To UBound(linea3)
                        If Not p = h Then
                            If linea3(p) = linea3(h) Then
                                loteOK = False
                                If Not codigosDuplicados.Contains(linea3(p)) Then
                                    codigosDuplicados += linea3(p) & vbNewLine
                                End If
                            End If
                        End If
                    Next p
                Next h
                If Not codigosDuplicados = "" Then
                    mensajeFinal += "Los siguientes códigos aparecen más de una vez: " & vbNewLine & codigosDuplicados & vbNewLine & vbNewLine
                End If
            End If

            If loteOK = True Then
                For h = 0 To UBound(linea3)
                    If linea3(h).Contains(".RC") Or
                        linea3(h).Contains("QB49N/") Or
                        linea3(h).Contains("SRC1/") Then
                    Else
                        loteOK = False
                        If Not errComp.Contains(linea3(h)) Then
                            errComp += linea3(h) & vbNewLine
                        End If
                    End If
                Next h
                If Not errComp = "" Then
                    mensajeFinal += "Los siguientes códigos no deben ir en un paquete de 2 displays de 49 pulgadas modelo QB49N: " & vbNewLine & errComp & vbNewLine & vbNewLine
                End If
            End If
            If loteOK = True Then
                For h = 0 To UBound(linea3)
                    If linea3(h).Contains(codigoLote + ".RC") Then
                        paquetesQueHay += 1
                    End If
                    If linea3(h).Contains("QB49N/") Then
                        dispQB49NQueHay += 1
                    End If
                    If linea3(h).Contains("SRC1/") Then
                        mandoQB49NQueHay += 1
                    End If
                Next h
                dispQB49NHayPorPaquete = dispQB49NQueHay / paquetesQueHay
                mandoQB49NHayPorPaquete = mandoQB49NQueHay / paquetesQueHay
                If dispQB49NHayPorPaquete <> dispQB49NPorPaquete Then
                    mensajeFinal += "En algún paquete, la cantidad de displays de 49 pulgadas modelo QB49N difiere de " & dispQB49NPorPaquete & vbNewLine
                    loteOK = False
                End If
                If mandoQB49NHayPorPaquete <> mandoQB49NPorPaquete Then
                    mensajeFinal += "En algún paquete, la cantidad de mandos a distancia para display de 49 pulgadas modelo QB49N difiere de " & mandoQB49NPorPaquete & vbNewLine
                    loteOK = False
                End If
            End If
            comprobacion()
            If errorRC = False And loteOK = True Then
                MsgBox("Después de darle al OK tendrás 5 segundos para seleccionar la ventana 'Pack a Road case'")
                volcado()
            ElseIf errorRC = False And loteOK = False Then
                MsgBox(mensajeFinal & vbNewLine & "Repasa los códigos.", MsgBoxStyle.MsgBoxSetForeground & MsgBoxStyle.Information, "Hay errores")
            End If
        End If

        'DISPLAY DE 3 x 49" QB49N
        If ComboBox1.Text = "3 x 49"" QB49N (4)" Then
            codigoLote = "3QB49N"
            limiteDeLotes = 4
            Dim errComp As String = ""
            Dim h As Integer = 0
            Dim p As Integer = 0
            Dim paquetesQueHay As Integer = 0
            Dim dispQB49NEnLista As Integer = 0
            Dim dispQB49NQueHay As Integer = 0
            Dim dispQB49NHayPorPaquete As Double = 0
            Dim dispQB49NPorPaquete As Integer = 3
            Dim mandoQB49NEnLista As Integer = 0
            Dim mandoQB49NQueHay As Integer = 0
            Dim mandoQB49NHayPorPaquete As Double = 0
            Dim mandoQB49NPorPaquete As Integer = 1
            Dim loteOK As Boolean = True
            Dim mensajeFinal As String = ""
            If Not linea3(0).Contains(".RC") Then
                mensajeFinal += "El primer código del paquete tiene que ser el de la caja" & vbNewLine & vbNewLine
                loteOK = False
            End If
            If loteOK = True Then
                For h = 0 To UBound(linea3)
                    For p = 0 To UBound(linea3)
                        If Not p = h Then
                            If linea3(p) = linea3(h) Then
                                loteOK = False
                                If Not codigosDuplicados.Contains(linea3(p)) Then
                                    codigosDuplicados += linea3(p) & vbNewLine
                                End If
                            End If
                        End If
                    Next p
                Next h
                If Not codigosDuplicados = "" Then
                    mensajeFinal += "Los siguientes códigos aparecen más de una vez: " & vbNewLine & codigosDuplicados & vbNewLine & vbNewLine
                End If
            End If

            If loteOK = True Then
                For h = 0 To UBound(linea3)
                    If linea3(h).Contains(".RC") Or
                        linea3(h).Contains("QB49N/") Or
                        linea3(h).Contains("SRC1/") Then
                    Else
                        loteOK = False
                        If Not errComp.Contains(linea3(h)) Then
                            errComp += linea3(h) & vbNewLine
                        End If
                    End If
                Next h
                If Not errComp = "" Then
                    mensajeFinal += "Los siguientes códigos no deben ir en un paquete de 3 displays de 49 pulgadas modelo QB49N: " & vbNewLine & errComp & vbNewLine & vbNewLine
                End If
            End If
            If loteOK = True Then
                For h = 0 To UBound(linea3)
                    If linea3(h).Contains(codigoLote + ".RC") Then
                        paquetesQueHay += 1
                    End If
                    If linea3(h).Contains("QB49N/") Then
                        dispQB49NQueHay += 1
                    End If
                    If linea3(h).Contains("SRC1/") Then
                        mandoQB49NQueHay += 1
                    End If
                Next h
                dispQB49NHayPorPaquete = dispQB49NQueHay / paquetesQueHay
                mandoQB49NHayPorPaquete = mandoQB49NQueHay / paquetesQueHay
                If dispQB49NHayPorPaquete <> dispQB49NPorPaquete Then
                    mensajeFinal += "En algún paquete, la cantidad de displays de 49 pulgadas modelo QB49N difiere de " & dispQB49NPorPaquete & vbNewLine
                    loteOK = False
                End If
                If mandoQB49NHayPorPaquete <> mandoQB49NPorPaquete Then
                    mensajeFinal += "En algún paquete, la cantidad de mandos a distancia para display de 49 pulgadas modelo QB49N difiere de " & mandoQB49NPorPaquete & vbNewLine
                    loteOK = False
                End If
            End If
            comprobacion()
            If errorRC = False And loteOK = True Then
                MsgBox("Después de darle al OK tendrás 5 segundos para seleccionar la ventana 'Pack a Road case'")
                volcado()
            ElseIf errorRC = False And loteOK = False Then
                MsgBox(mensajeFinal & vbNewLine & "Repasa los códigos.", MsgBoxStyle.MsgBoxSetForeground & MsgBoxStyle.Information, "Hay errores")
            End If
        End If

        'DISPLAY DE 1 x 48" DB48E
        If ComboBox1.Text = "1 x 48"" (4)" Then
            codigoLote = "1DB48E"
            limiteDeLotes = 4
            Dim errComp As String = ""
            Dim h As Integer = 0
            Dim p As Integer = 0
            Dim paquetesQueHay As Integer = 0
            Dim disp1DB48EEnLista As Integer = 0
            Dim disp1DB48EQueHay As Integer = 0
            Dim disp1DB48EHayPorPaquete As Double = 0
            Dim disp1DB48EPorPaquete As Integer = 1
            Dim mando1DB48EEnLista As Integer = 0
            Dim mando1DB48EQueHay As Integer = 0
            Dim mando1DB48EHayPorPaquete As Double = 0
            Dim mando1DB48EPorPaquete As Integer = 1
            Dim loteOK As Boolean = True
            Dim mensajeFinal As String = ""
            If Not linea3(0).Contains(".RC") Then
                mensajeFinal += "El primer código del paquete tiene que ser el de la caja" & vbNewLine & vbNewLine
                loteOK = False
            End If
            If loteOK = True Then
                For h = 0 To UBound(linea3)
                    For p = 0 To UBound(linea3)
                        If Not p = h Then
                            If linea3(p) = linea3(h) Then
                                loteOK = False
                                If Not codigosDuplicados.Contains(linea3(p)) Then
                                    codigosDuplicados += linea3(p) & vbNewLine
                                End If
                            End If
                        End If
                    Next p
                Next h
                If Not codigosDuplicados = "" Then
                    mensajeFinal += "Los siguientes códigos aparecen más de una vez: " & vbNewLine & codigosDuplicados & vbNewLine & vbNewLine
                End If
            End If

            If loteOK = True Then
                For h = 0 To UBound(linea3)
                    If linea3(h).Contains(".RC") Or
                        linea3(h).Contains("DB48E/") Or
                        linea3(h).Contains("SRC1/") Then
                    Else
                        loteOK = False
                        If Not errComp.Contains(linea3(h)) Then
                            errComp += linea3(h) & vbNewLine
                        End If
                    End If
                Next h
                If Not errComp = "" Then
                    mensajeFinal += "Los siguientes códigos no deben ir en un paquete de 1 display de 48 pulgadas modelo DB48E: " & vbNewLine & errComp & vbNewLine & vbNewLine
                End If
            End If
            If loteOK = True Then
                For h = 0 To UBound(linea3)
                    If linea3(h).Contains(codigoLote + ".RC") Then
                        paquetesQueHay += 1
                    End If
                    If linea3(h).Contains("DB48E/") Then
                        disp1DB48EQueHay += 1
                    End If
                    If linea3(h).Contains("SRC1/") Then
                        mando1DB48EQueHay += 1
                    End If
                Next h

                disp1DB48EHayPorPaquete = disp1DB48EQueHay / paquetesQueHay
                mando1DB48EHayPorPaquete = mando1DB48EQueHay / paquetesQueHay

                If disp1DB48EHayPorPaquete <> disp1DB48EPorPaquete Then
                    mensajeFinal += "En algún paquete, la cantidad de displays de 48 pulgadas modelo DB48E difiere de " & disp1DB48EPorPaquete & vbNewLine
                    loteOK = False
                End If

                If mando1DB48EHayPorPaquete <> mando1DB48EPorPaquete Then
                    mensajeFinal += "En algún paquete, la cantidad de mandos a distancia para display de 48 pulgadas modelo DB48E difiere de " & mando1DB48EPorPaquete & vbNewLine
                    loteOK = False
                End If

            End If
            comprobacion()
            If errorRC = False And loteOK = True Then
                MsgBox("Después de darle al OK tendrás 5 segundos para seleccionar la ventana 'Pack a Road case'")
                volcado()
            ElseIf errorRC = False And loteOK = False Then
                MsgBox(mensajeFinal & vbNewLine & "Repasa los códigos.", MsgBoxStyle.MsgBoxSetForeground & MsgBoxStyle.Information, "Hay errores")
            End If
        End If

        'DISPLAY DE 2 x 48" DB48E
        If ComboBox1.Text = "2 x 48"" (4)" Then
            codigoLote = "2DB48E"
            limiteDeLotes = 4
            Dim errComp As String = ""
            Dim h As Integer = 0
            Dim p As Integer = 0
            Dim paquetesQueHay As Integer = 0
            Dim disp2DB48EEnLista As Integer = 0
            Dim disp2DB48EQueHay As Integer = 0
            Dim disp2DB48EHayPorPaquete As Double = 0
            Dim disp2DB48EPorPaquete As Integer = 2
            Dim mando2DB48EEnLista As Integer = 0
            Dim mando2DB48EQueHay As Integer = 0
            Dim mando2DB48EHayPorPaquete As Double = 0
            Dim mando2DB48EPorPaquete As Integer = 2
            Dim loteOK As Boolean = True
            Dim mensajeFinal As String = ""
            If Not linea3(0).Contains(".RC") Then
                mensajeFinal += "El primer código del paquete tiene que ser el de la caja" & vbNewLine & vbNewLine
                loteOK = False
            End If
            If loteOK = True Then
                For h = 0 To UBound(linea3)
                    For p = 0 To UBound(linea3)
                        If Not p = h Then
                            If linea3(p) = linea3(h) Then
                                loteOK = False
                                If Not codigosDuplicados.Contains(linea3(p)) Then
                                    codigosDuplicados += linea3(p) & vbNewLine
                                End If
                            End If
                        End If
                    Next p
                Next h
                If Not codigosDuplicados = "" Then
                    mensajeFinal += "Los siguientes códigos aparecen más de una vez: " & vbNewLine & codigosDuplicados & vbNewLine & vbNewLine
                End If
            End If

            If loteOK = True Then
                For h = 0 To UBound(linea3)
                    If linea3(h).Contains(".RC") Or
                        linea3(h).Contains("DB48E/") Or
                        linea3(h).Contains("SRC1/") Then
                    Else
                        loteOK = False
                        If Not errComp.Contains(linea3(h)) Then
                            errComp += linea3(h) & vbNewLine
                        End If
                    End If
                Next h
                If Not errComp = "" Then
                    mensajeFinal += "Los siguientes códigos no deben ir en un paquete de 2 display de 48 pulgadas modelo DB48E: " & vbNewLine & errComp & vbNewLine & vbNewLine
                End If
            End If
            If loteOK = True Then
                For h = 0 To UBound(linea3)
                    If linea3(h).Contains(codigoLote + ".RC") Then
                        paquetesQueHay += 1
                    End If
                    If linea3(h).Contains("DB48E/") Then
                        disp2DB48EQueHay += 1
                    End If
                    If linea3(h).Contains("SRC1/") Then
                        mando2DB48EQueHay += 1
                    End If
                Next h
                disp2DB48EHayPorPaquete = disp2DB48EQueHay / paquetesQueHay
                mando2DB48EHayPorPaquete = mando2DB48EQueHay / paquetesQueHay
                If disp2DB48EHayPorPaquete <> disp2DB48EPorPaquete Then
                    mensajeFinal += "En algún paquete, la cantidad de displays de 48 pulgadas modelo DB48E difiere de " & disp2DB48EPorPaquete & vbNewLine
                    loteOK = False
                End If
                If mando2DB48EHayPorPaquete <> mando2DB48EPorPaquete Then
                    mensajeFinal += "En algún paquete, la cantidad de mandos a distancia para display de 48 pulgadas modelo DB48E difiere de " & mando2DB48EPorPaquete & vbNewLine
                    loteOK = False
                End If
            End If
            comprobacion()
            If errorRC = False And loteOK = True Then
                MsgBox("Después de darle al OK tendrás 5 segundos para seleccionar la ventana 'Pack a Road case'")
                volcado()
            ElseIf errorRC = False And loteOK = False Then
                MsgBox(mensajeFinal & vbNewLine & "Repasa los códigos.", MsgBoxStyle.MsgBoxSetForeground & MsgBoxStyle.Information, "Hay errores")
            End If
        End If

        'DISPLAY DE 2 x 46" ME46C
        If ComboBox1.Text = "2 x 46"" (4)" Then
            codigoLote = "ME46C"
            limiteDeLotes = 4
            Dim errComp As String = ""
            Dim h As Integer = 0
            Dim p As Integer = 0
            Dim paquetesQueHay As Integer = 0
            Dim dispME46CEnLista As Integer = 0
            Dim dispME46CQueHay As Integer = 0
            Dim dispME46CHayPorPaquete As Double = 0
            Dim dispME46CPorPaquete As Integer = 2
            Dim mandoME46CEnLista As Integer = 0
            Dim mandoME46CQueHay As Integer = 0
            Dim mandoME46CHayPorPaquete As Double = 0
            Dim mandoME46CPorPaquete As Integer = 2
            Dim loteOK As Boolean = True
            Dim mensajeFinal As String = ""
            If Not linea3(0).Contains(".RC") Then
                mensajeFinal += "El primer código del paquete tiene que ser el de la caja" & vbNewLine & vbNewLine
                loteOK = False
            End If
            If loteOK = True Then
                For h = 0 To UBound(linea3)
                    For p = 0 To UBound(linea3)
                        If Not p = h Then
                            If linea3(p) = linea3(h) Then
                                loteOK = False
                                If Not codigosDuplicados.Contains(linea3(p)) Then
                                    codigosDuplicados += linea3(p) & vbNewLine
                                End If
                            End If
                        End If
                    Next p
                Next h
                If Not codigosDuplicados = "" Then
                    mensajeFinal += "Los siguientes códigos aparecen más de una vez: " & vbNewLine & codigosDuplicados & vbNewLine & vbNewLine
                End If
            End If

            If loteOK = True Then
                For h = 0 To UBound(linea3)
                    If linea3(h).Contains(".RC") Or
                        linea3(h).Contains("ME46C/") Or
                        linea3(h).Contains("SRC1/") Then
                    Else
                        loteOK = False
                        If Not errComp.Contains(linea3(h)) Then
                            errComp += linea3(h) & vbNewLine
                        End If
                    End If
                Next h
                If Not errComp = "" Then
                    mensajeFinal += "Los siguientes códigos no deben ir en un paquete de 2 display de 46 pulgadas modelo ME46C: " & vbNewLine & errComp & vbNewLine & vbNewLine
                End If
            End If
            If loteOK = True Then
                For h = 0 To UBound(linea3)
                    If linea3(h).Contains(codigoLote + ".RC") Then
                        paquetesQueHay += 1
                    End If
                    If linea3(h).Contains("ME46C/") Then
                        dispME46CQueHay += 1
                    End If
                    If linea3(h).Contains("SRC1/") Then
                        mandoME46CQueHay += 1
                    End If
                Next h
                dispME46CHayPorPaquete = dispME46CQueHay / paquetesQueHay
                mandoME46CHayPorPaquete = mandoME46CQueHay / paquetesQueHay
                If dispME46CHayPorPaquete <> dispME46CPorPaquete Then
                    mensajeFinal += "En algún paquete, la cantidad de displays de 46 pulgadas modelo ME46C difiere de " & dispME46CPorPaquete & vbNewLine
                    loteOK = False
                End If
                If mandoME46CHayPorPaquete <> mandoME46CPorPaquete Then
                    mensajeFinal += "En algún paquete, la cantidad de mandos a distancia para display de 46 pulgadas modelo ME46C difiere de " & mandoME46CPorPaquete & vbNewLine
                    loteOK = False
                End If
            End If
            comprobacion()
            If errorRC = False And loteOK = True Then
                MsgBox("Después de darle al OK tendrás 5 segundos para seleccionar la ventana 'Pack a Road case'")
                volcado()
            ElseIf errorRC = False And loteOK = False Then
                MsgBox(mensajeFinal & vbNewLine & "Repasa los códigos.", MsgBoxStyle.MsgBoxSetForeground & MsgBoxStyle.Information, "Hay errores")
            End If
        End If

        'CLEVERTOUCH DE 55"
        If ComboBox1.Text = "Clevertouch 55"" (4)" Then
            codigoLote = "CLEV55"
            limiteDeLotes = 4
            Dim errComp As String = ""
            Dim h As Integer = 0
            Dim p As Integer = 0
            Dim paquetesQueHay As Integer = 0
            Dim clev55EnLista As Integer = 0
            Dim clev55QueHay As Integer = 0
            Dim clev55HayPorPaquete As Double = 0
            Dim clev55PorPaquete As Integer = 2
            Dim antenaEnLista As Integer = 0
            Dim antenaQueHay As Integer = 0
            Dim antenaHayPorPaquete As Double = 0
            Dim antenaPorPaquete As Integer = 2
            Dim mandoEnLista As Integer = 0
            Dim mandoQueHay As Integer = 0
            Dim mandoHayPorPaquete As Double = 0
            Dim mandoPorPaquete As Integer = 2
            Dim usbEnLista As Integer = 0
            Dim usbQueHay As Integer = 0
            Dim usbHayPorPaquete As Double = 0
            Dim usbPorPaquete As Integer = 2
            Dim penEnLista As Integer = 0
            Dim penQueHay As Integer = 0
            Dim penHayPorPaquete As Double = 0
            Dim penPorPaquete As Integer = 2
            Dim loteOK As Boolean = True
            Dim mensajeFinal As String = ""
            If Not linea3(0).Contains(".RC") Then
                mensajeFinal += "El primer código del paquete tiene que ser el de la caja" & vbNewLine & vbNewLine
                loteOK = False
            End If
            If loteOK = True Then
                For h = 0 To UBound(linea3)
                    For p = 0 To UBound(linea3)
                        If Not p = h Then
                            If linea3(p) = linea3(h) Then
                                loteOK = False
                                If Not codigosDuplicados.Contains(linea3(p)) Then
                                    codigosDuplicados += linea3(p) & vbNewLine
                                End If
                            End If
                        End If
                    Next p
                Next h
                If Not codigosDuplicados = "" Then
                    mensajeFinal += "Los siguientes códigos aparecen más de una vez: " & vbNewLine & codigosDuplicados & vbNewLine & vbNewLine
                End If
            End If

            If loteOK = True Then
                For h = 0 To UBound(linea3)
                    If linea3(h).Contains(".RC") Or
                        linea3(h).Contains("2458G/") Or
                        linea3(h).Contains("CLEV55P/") Or
                        linea3(h).Contains("CLEVERC/") Or
                        linea3(h).Contains("CLUS/") Or
                        linea3(h).Contains("RUBPEN/") Then

                    Else
                        loteOK = False
                        If Not errComp.Contains(linea3(h)) Then
                            errComp += linea3(h) & vbNewLine
                        End If
                    End If
                Next h
                If Not errComp = "" Then
                    mensajeFinal += "Los siguientes códigos no deben ir en un paquete de CLEVERTOUCH 55 PULGADAS: " & vbNewLine & errComp & vbNewLine & vbNewLine
                End If
            End If
            If loteOK = True Then
                For h = 0 To UBound(linea3)
                    If linea3(h).Contains(codigoLote + ".RC") Then
                        paquetesQueHay += 1
                    End If
                    If linea3(h).Contains("2458G/") Then
                        antenaQueHay += 1
                    End If
                    If linea3(h).Contains("CLEV55P/") Then
                        clev55QueHay += 1
                    End If
                    If linea3(h).Contains("CLEVERC/") Then
                        mandoQueHay += 1
                    End If
                    If linea3(h).Contains("CLUS/") Then
                        usbQueHay += 1
                    End If
                    If linea3(h).Contains("RUBPEN/") Then
                        penQueHay += 1
                    End If
                Next h
                clev55HayPorPaquete = clev55QueHay / paquetesQueHay
                antenaHayPorPaquete = antenaQueHay / paquetesQueHay
                mandoHayPorPaquete = mandoQueHay / paquetesQueHay
                usbHayPorPaquete = usbQueHay / paquetesQueHay
                penHayPorPaquete = penQueHay / paquetesQueHay
                If clev55HayPorPaquete <> clev55PorPaquete Then
                    mensajeFinal += "En algún paquete, la cantidad de Clevertouch de 55 pulgadas difiere de " & clev55PorPaquete & vbNewLine
                    loteOK = False
                End If
                If mandoHayPorPaquete <> mandoPorPaquete Then
                    mensajeFinal += "En algún paquete, la cantidad de mandos a distancias para Clevertouch 55 pulgadas difiere de " & mandoPorPaquete & vbNewLine
                    loteOK = False
                End If
                If antenaHayPorPaquete <> antenaPorPaquete Then
                    mensajeFinal += "En algún paquete, la cantidad de antenas para Clevertouch 55 pulgadas difiere de " & antenaPorPaquete & vbNewLine
                    loteOK = False
                End If
                If usbHayPorPaquete <> usbPorPaquete Then
                    mensajeFinal += "En algún paquete, la cantidad de cables usb para Clevertouch 55 pulgadas difiere de " & usbPorPaquete & vbNewLine
                    loteOK = False
                End If
                If penHayPorPaquete <> penPorPaquete Then
                    mensajeFinal += "En algún paquete, la cantidad de punteros para Clevrtouch 55 pulgadas difiere de " & penPorPaquete & vbNewLine
                    loteOK = False
                End If
            End If
            comprobacion()
            If errorRC = False And loteOK = True Then
                MsgBox("Después de darle al OK tendrás 5 segundos para seleccionar la ventana 'Pack a Road case'")
                volcado()
            ElseIf errorRC = False And loteOK = False Then
                MsgBox(mensajeFinal & vbNewLine & "Repasa los códigos.", MsgBoxStyle.MsgBoxSetForeground & MsgBoxStyle.Information, "Hay errores")
            End If
        End If

        'VIDEOWALL X55
        If ComboBox1.Text = "NEC X55 (5)" Then
            codigoLote = "X55"
            limiteDeLotes = 5
            Dim errComp As String = ""
            Dim h As Integer = 0
            Dim p As Integer = 0
            Dim paquetesQueHay As Integer = 0
            Dim x55EnLista As Integer = 0
            Dim x55QueHay As Integer = 0
            Dim x55HayPorPaquete As Double = 0
            Dim x55PorPaquete As Integer = 2
            Dim maletinEnLista As Integer = 0
            Dim maletinQueHay As Integer = 0
            Dim maletinHayPorPaquete As Double = 0
            Dim maletinPorPaquete As Integer = 2
            Dim frameEnLista As Integer = 0
            Dim frameQueHay As Integer = 0
            Dim frameHayPorPaquete As Double = 0
            Dim framePorPaquete As Integer = 2
            Dim stackerEnLista As Integer = 0
            Dim stackerQueHay As Integer = 0
            Dim stackerHayPorPaquete As Double = 0
            Dim stackerPorPaquete As Integer = 2
            Dim cuatroEnLista As Integer = 0
            Dim cuatroQueHay As Integer = 0
            Dim cuatroHayPorPaquete As Double = 0
            Dim cuatroPorPaquete As Integer = 2
            Dim dosEnLista As Integer = 0
            Dim dosQueHay As Integer = 0
            Dim dosHayPorPaquete As Double = 0
            Dim dosPorPaquete As Integer = 2
            Dim loteOK As Boolean = True
            Dim mensajeFinal As String = ""
            If Not linea3(0).Contains(".RC") Then
                mensajeFinal += "El primer código del paquete tiene que ser el de la caja" & vbNewLine & vbNewLine
                loteOK = False
            End If
            If loteOK = True Then
                For h = 0 To UBound(linea3)
                    For p = 0 To UBound(linea3)
                        If Not p = h Then
                            If linea3(p) = linea3(h) Then
                                loteOK = False
                                If Not codigosDuplicados.Contains(linea3(p)) Then
                                    codigosDuplicados += linea3(p) & vbNewLine
                                End If
                            End If
                        End If
                    Next p
                Next h
                If Not codigosDuplicados = "" Then
                    mensajeFinal += "Los siguientes códigos aparecen más de una vez: " & vbNewLine & codigosDuplicados & vbNewLine & vbNewLine
                End If
            End If

            If loteOK = True Then
                For h = 0 To UBound(linea3)
                    If linea3(h).Contains(".RC") Or
                        linea3(h).Contains("X55AS/") Or
                        linea3(h).Contains("X55/") Or
                        linea3(h).Contains("X55FR/") Or
                        linea3(h).Contains("X55SF/") Or
                        linea3(h).Contains("2M/") Or
                        linea3(h).Contains("4M/") Then

                    Else
                        loteOK = False
                        If Not errComp.Contains(linea3(h)) Then
                            errComp += linea3(h) & vbNewLine
                        End If
                    End If
                Next h
                If Not errComp = "" Then
                    mensajeFinal += "Los siguientes códigos no deben ir en un paquete de VIDEOWALL X55: " & vbNewLine & errComp & vbNewLine & vbNewLine
                End If
            End If
            If loteOK = True Then
                For h = 0 To UBound(linea3)
                    If linea3(h).Contains(codigoLote + ".RC") Then
                        paquetesQueHay += 1
                    End If
                    If linea3(h).Contains("X55AS/") Then
                        maletinQueHay += 1
                    End If
                    If linea3(h).Contains("X55/") Then
                        x55QueHay += 1
                    End If
                    If linea3(h).Contains("X55FR/") Then
                        frameQueHay += 1
                    End If
                    If linea3(h).Contains("X55SF/") Then
                        stackerQueHay += 1
                    End If
                    If linea3(h).Contains("2M/") Then
                        dosQueHay += 1
                    End If
                    If linea3(h).Contains("4M/") Then
                        cuatroQueHay += 1
                    End If
                Next h
                x55HayPorPaquete = x55QueHay / paquetesQueHay
                maletinHayPorPaquete = maletinQueHay / paquetesQueHay
                frameHayPorPaquete = frameQueHay / paquetesQueHay
                stackerHayPorPaquete = stackerQueHay / paquetesQueHay
                dosHayPorPaquete = dosQueHay / paquetesQueHay
                cuatroHayPorPaquete = cuatroQueHay / paquetesQueHay
                If x55HayPorPaquete <> x55PorPaquete Then
                    mensajeFinal += "En algún paquete, la cantidad de 'Full HD LCD Seamless Display NEC X555UNS' difiere de " & x55PorPaquete & vbNewLine
                    loteOK = False
                End If
                If frameHayPorPaquete <> framePorPaquete Then
                    mensajeFinal += "En algún paquete, la cantidad de 'Easyframe Frame for X55UNS' difiere de " & framePorPaquete & vbNewLine
                    loteOK = False
                End If
                If maletinHayPorPaquete <> maletinPorPaquete Then
                    mensajeFinal += "En algún paquete, la cantidad de 'Accessories Suitcase for X55UNS' difiere de " & maletinPorPaquete & vbNewLine
                    loteOK = False
                End If
                If stackerHayPorPaquete <> stackerPorPaquete Then
                    mensajeFinal += "En algún paquete, la cantidad de 'Easyframe Stacker frame for X555UNS' difiere de " & stackerPorPaquete & vbNewLine
                    loteOK = False
                End If
                If dosHayPorPaquete <> dosPorPaquete Then
                    mensajeFinal += "En algún paquete, la cantidad de 'Easyframe Double Magnetic Holder' difiere de " & dosPorPaquete & vbNewLine
                    loteOK = False
                End If
                If cuatroHayPorPaquete <> cuatroPorPaquete Then
                    mensajeFinal += "En algún paquete, la cantidad de 'Easyframe Quadruple Magnetic Holder' difiere de " & cuatroPorPaquete & vbNewLine
                    loteOK = False
                End If
            End If
            comprobacion()
            If errorRC = False And loteOK = True Then
                MsgBox("Después de darle al OK tendrás 5 segundos para seleccionar la ventana 'Pack a Road case'")
                volcado()
            ElseIf errorRC = False And loteOK = False Then
                MsgBox(mensajeFinal & vbNewLine & "Repasa los códigos.", MsgBoxStyle.MsgBoxSetForeground & MsgBoxStyle.Information, "Hay errores")
            End If
        End If



        'CUATRO T10
        If ComboBox1.Text = "4 x T10 (8)" Then
            codigoLote = "4T10"
            limiteDeLotes = 8
            Dim errComp As String = ""
            Dim h As Integer = 0
            Dim p As Integer = 0
            Dim paquetesQueHay As Integer = 0
            Dim t10EnLista As Integer = 0
            Dim t10QueHay As Integer = 0
            Dim t10HayPorPaquete As Double = 0
            Dim t10PorPaquete As Integer = 4
            Dim loteOK As Boolean = True
            Dim mensajeFinal As String = ""
            If Not linea3(0).Contains(".RC") Then
                mensajeFinal += "El primer código del paquete tiene que ser el de la caja" & vbNewLine & vbNewLine
                loteOK = False
            End If
            If loteOK = True Then
                For h = 0 To UBound(linea3)
                    For p = 0 To UBound(linea3)
                        If Not p = h Then
                            If linea3(p) = linea3(h) Then
                                loteOK = False
                                If Not codigosDuplicados.Contains(linea3(p)) Then
                                    codigosDuplicados += linea3(p) & vbNewLine
                                End If
                            End If
                        End If
                    Next p
                Next h
                If Not codigosDuplicados = "" Then
                    mensajeFinal += "Los siguientes códigos aparecen más de una vez: " & vbNewLine & codigosDuplicados & vbNewLine & vbNewLine
                End If
            End If

            If loteOK = True Then
                For h = 0 To UBound(linea3)
                    If linea3(h).Contains(".RC") Or linea3(h).Contains("T10/") Then

                    Else
                        loteOK = False
                        If Not errComp.Contains(linea3(h)) Then
                            errComp += linea3(h) & vbNewLine
                        End If
                    End If
                Next h
                If Not errComp = "" Then
                    mensajeFinal += "Los siguientes códigos no deben en ir en la caja de cuatro T10: " & vbNewLine & errComp & vbNewLine & vbNewLine
                End If
            End If
            If loteOK = True Then
                For h = 0 To UBound(linea3)
                    If linea3(h).Contains(codigoLote + ".RC") Then
                        paquetesQueHay += 1
                    End If
                    If linea3(h).Contains("T10/") Then
                        t10QueHay += 1
                    End If
                Next h
                t10HayPorPaquete = t10QueHay / paquetesQueHay
                If t10HayPorPaquete <> t10PorPaquete Then
                    mensajeFinal += "La cantidad de T10 por paquete difiere de " & t10PorPaquete & vbNewLine
                    loteOK = False
                End If
            End If
            comprobacion()
            If errorRC = False And loteOK = True Then
                MsgBox("Después de darle al OK tendrás 5 segundos para seleccionar la ventana 'Pack a Road case'")
                volcado()
            ElseIf errorRC = False And loteOK = False Then
                MsgBox(mensajeFinal & vbNewLine & "Repasa los códigos.", MsgBoxStyle.MsgBoxSetForeground & MsgBoxStyle.Information, "Hay errores")
            End If
        End If

        'MAUI
        If ComboBox1.Text = "MAUI (4)" Then
            codigoLote = "2G2"
            limiteDeLotes = 4
            Dim errComp As String = ""
            Dim h As Integer = 0
            Dim p As Integer = 0
            Dim paquetesQueHay As Integer = 0

            Dim subEnLista As Integer = 0
            Dim subQueHay As Integer = 0
            Dim subHayPorPaquete As Double = 0
            Dim subPorPaquete As Integer = 2
            Dim utopEnLista As Integer = 0
            Dim utopQueHay As Integer = 0
            Dim utopHayPorPaquete As Double = 0
            Dim utopPorPaquete As Integer = 2
            Dim ltopEnLista As Integer = 0
            Dim ltopQueHay As Integer = 0
            Dim ltopHayPorPaquete As Double = 0
            Dim ltopPorPaquete As Integer = 2

            Dim loteOK As Boolean = True
            Dim mensajeFinal As String = ""
            If Not linea3(0).Contains(".RC") Then
                mensajeFinal += "El primer código del paquete tiene que ser el de la caja" & vbNewLine & vbNewLine
                loteOK = False
            End If
            If loteOK = True Then
                For h = 0 To UBound(linea3)
                    For p = 0 To UBound(linea3)
                        If Not p = h Then
                            If linea3(p) = linea3(h) Then
                                loteOK = False
                                If Not codigosDuplicados.Contains(linea3(p)) Then
                                    codigosDuplicados += linea3(p) & vbNewLine
                                End If
                            End If
                        End If
                    Next p
                Next h
                If Not codigosDuplicados = "" Then
                    mensajeFinal += "Los siguientes códigos aparecen más de una vez: " & vbNewLine & codigosDuplicados & vbNewLine & vbNewLine
                End If
            End If

            If loteOK = True Then
                For h = 0 To UBound(linea3)
                    If linea3(h).Contains(".RC") Or
                        linea3(h).Contains("G2SUB/") Or
                        linea3(h).Contains("G2LTOP/") Or
                        linea3(h).Contains("G2UTOP/") Then
                    Else
                        loteOK = False
                        If Not errComp.Contains(linea3(h)) Then
                            errComp += linea3(h) & vbNewLine
                        End If
                    End If
                Next h
                If Not errComp = "" Then
                    mensajeFinal += "Los siguientes códigos no deben ir en un paquete de MAUI G2: " & vbNewLine & errComp & vbNewLine & vbNewLine
                End If
            End If
            If loteOK = True Then
                For h = 0 To UBound(linea3)
                    If linea3(h).Contains(codigoLote + ".RC") Then
                        paquetesQueHay += 1
                    End If
                    If linea3(h).Contains("G2SUB/") Then
                        subQueHay += 1
                    End If
                    If linea3(h).Contains("G2LTOP/") Then
                        ltopQueHay += 1
                    End If
                    If linea3(h).Contains("G2UTOP/") Then
                        utopQueHay += 1
                    End If

                Next h
                subHayPorPaquete = subQueHay / paquetesQueHay
                ltopHayPorPaquete = ltopQueHay / paquetesQueHay
                utopHayPorPaquete = utopQueHay / paquetesQueHay

                If subHayPorPaquete <> subPorPaquete Then
                    mensajeFinal += "En algún paquete, la cantidad de subgraves MAUI difiere de " & subPorPaquete & vbNewLine
                    loteOK = False
                End If
                If ltopHayPorPaquete <> ltopPorPaquete Then
                    mensajeFinal += "En algún paquete, la cantidad de top inferior Maui difiere de " & ltopPorPaquete & vbNewLine
                    loteOK = False
                End If
                If utopHayPorPaquete <> utopPorPaquete Then
                    mensajeFinal += "En algún paquete, la cantidad de top superior Maui difiere de " & utopPorPaquete & vbNewLine
                    loteOK = False
                End If
            End If
            comprobacion()
            If errorRC = False And loteOK = True Then
                MsgBox("Después de darle al OK tendrás 5 segundos para seleccionar la ventana 'Pack a Road case'")
                volcado()
            ElseIf errorRC = False And loteOK = False Then
                MsgBox(mensajeFinal & vbNewLine & "Repasa los códigos.", MsgBoxStyle.MsgBoxSetForeground & MsgBoxStyle.Information, "Hay errores")
            End If
        End If

        'MESA YAMAHA QL5
        If ComboBox1.Text = "QL5 (4)" Then
            codigoLote = "QL5"
            limiteDeLotes = 4
            Dim errComp As String = ""
            Dim h As Integer = 0
            Dim p As Integer = 0
            Dim paquetesQueHay As Integer = 0
            Dim ql5EnLista As Integer = 0
            Dim ql5QueHay As Integer = 0
            Dim ql5HayPorPaquete As Double = 0
            Dim ql5PorPaquete As Integer = 1
            Dim loteOK As Boolean = True
            Dim mensajeFinal As String = ""
            If Not linea3(0).Contains(".RC") Then
                mensajeFinal += "El primer código del paquete tiene que ser el de la caja" & vbNewLine & vbNewLine
                loteOK = False
            End If
            If loteOK = True Then
                For h = 0 To UBound(linea3)
                    For p = 0 To UBound(linea3)
                        If Not p = h Then
                            If linea3(p) = linea3(h) Then
                                loteOK = False
                                If Not codigosDuplicados.Contains(linea3(p)) Then
                                    codigosDuplicados += linea3(p) & vbNewLine
                                End If
                            End If
                        End If
                    Next p
                Next h
                If Not codigosDuplicados = "" Then
                    mensajeFinal += "Los siguientes códigos aparecen más de una vez: " & vbNewLine & codigosDuplicados & vbNewLine & vbNewLine
                End If
            End If

            If loteOK = True Then
                For h = 0 To UBound(linea3)
                    If linea3(h).Contains(".RC") Or linea3(h).Contains("YQL548/") Then

                    Else
                        loteOK = False
                        If Not errComp.Contains(linea3(h)) Then
                            errComp += linea3(h) & vbNewLine
                        End If
                    End If
                Next h
                If Not errComp = "" Then
                    mensajeFinal += "Los siguientes códigos no deben ir en un paquete de QL5: " & vbNewLine & errComp & vbNewLine & vbNewLine
                End If
            End If
            If loteOK = True Then
                For h = 0 To UBound(linea3)
                    If linea3(h).Contains(codigoLote + ".RC") Then
                        paquetesQueHay += 1
                    End If
                    If linea3(h).Contains("YQL548/") Then
                        ql5QueHay += 1
                    End If
                Next h
                ql5HayPorPaquete = ql5QueHay / paquetesQueHay
                If ql5HayPorPaquete <> ql5PorPaquete Then
                    mensajeFinal += "La cantidad de QL5 por paquete difiere de " & ql5PorPaquete & vbNewLine
                    loteOK = False
                End If
            End If
            comprobacion()
            If errorRC = False And loteOK = True Then
                MsgBox("Después de darle al OK tendrás 5 segundos para seleccionar la ventana 'Pack a Road case'")
                volcado()
            ElseIf errorRC = False And loteOK = False Then
                MsgBox(mensajeFinal & vbNewLine & "Repasa los códigos.", MsgBoxStyle.MsgBoxSetForeground & MsgBoxStyle.Information, "Hay errores")
            End If
        End If



        'AXIOM
        If ComboBox1.Text = "Axiom (6)" Then
            codigoLote = "AX"
            limiteDeLotes = 6
            Dim errComp As String = ""
            Dim h As Integer = 0
            Dim p As Integer = 0
            Dim paquetesQueHay As Integer = 0
            Dim axiomsEnLista As Integer = 0
            Dim axiomQueHay As Integer = 0
            Dim axiomHayPorPaquete As Double = 0
            Dim axiomPorPaquete As Integer = 2
            Dim loteOK As Boolean = True
            Dim mensajeFinal As String = ""
            If Not linea3(0).Contains(".RC") Then
                mensajeFinal += "El primer código del paquete tiene que ser el de la caja" & vbNewLine & vbNewLine
                loteOK = False
            End If
            If loteOK = True Then
                For h = 0 To UBound(linea3)
                    For p = 0 To UBound(linea3)
                        If Not p = h Then
                            If linea3(p) = linea3(h) Then
                                loteOK = False
                                If Not codigosDuplicados.Contains(linea3(p)) Then
                                    codigosDuplicados += linea3(p) & vbNewLine
                                End If
                            End If
                        End If
                    Next p
                Next h
                If Not codigosDuplicados = "" Then
                    mensajeFinal += "Los siguientes códigos aparecen más de una vez: " & vbNewLine & codigosDuplicados & vbNewLine & vbNewLine
                End If
            End If

            If loteOK = True Then
                For h = 0 To UBound(linea3)
                    If linea3(h).Contains(".RC") Or linea3(h).Contains("AXIOM/") Then

                    Else
                        loteOK = False
                        If Not errComp.Contains(linea3(h)) Then
                            errComp += linea3(h) & vbNewLine
                        End If
                    End If
                Next h
                If Not errComp = "" Then
                    mensajeFinal += "Los siguientes códigos no deben ir en un paquete de MAC AXIOM: " & vbNewLine & errComp & vbNewLine & vbNewLine
                End If
            End If
            If loteOK = True Then
                For h = 0 To UBound(linea3)
                    If linea3(h).Contains(codigoLote + ".RC") Then
                        paquetesQueHay += 1
                    End If
                    If linea3(h).Contains("AXIOM/") Then
                        axiomQueHay += 1
                    End If
                Next h
                axiomHayPorPaquete = axiomQueHay / paquetesQueHay
                If axiomHayPorPaquete <> axiomPorPaquete Then
                    mensajeFinal += "La cantidad de AXIOM por paquete difiere de " & axiomPorPaquete & vbNewLine
                    loteOK = False
                End If
            End If
            comprobacion()
            If errorRC = False And loteOK = True Then
                MsgBox("Después de darle al OK tendrás 5 segundos para seleccionar la ventana 'Pack a Road case'")
                volcado()
            ElseIf errorRC = False And loteOK = False Then
                MsgBox(mensajeFinal & vbNewLine & "Repasa los códigos.", MsgBoxStyle.MsgBoxSetForeground & MsgBoxStyle.Information, "Hay errores")
            End If
        End If

        'MARTIN MAC700
        If ComboBox1.Text = "MAC700 (6)" Then
            codigoLote = "700"
            limiteDeLotes = 6
            Dim errComp As String = ""
            Dim h As Integer = 0
            Dim p As Integer = 0
            Dim paquetesQueHay As Integer = 0
            Dim mac700EnLista As Integer = 0
            Dim mac700QueHay As Integer = 0
            Dim mac700HayPorPaquete As Double = 0
            Dim mac700PorPaquete As Integer = 2
            Dim loteOK As Boolean = True
            Dim mensajeFinal As String = ""
            If Not linea3(0).Contains(".RC") Then
                mensajeFinal += "El primer código del paquete tiene que ser el de la caja" & vbNewLine & vbNewLine
                loteOK = False
            End If
            If loteOK = True Then
                For h = 0 To UBound(linea3)
                    For p = 0 To UBound(linea3)
                        If Not p = h Then
                            If linea3(p) = linea3(h) Then
                                loteOK = False
                                If Not codigosDuplicados.Contains(linea3(p)) Then
                                    codigosDuplicados += linea3(p) & vbNewLine
                                End If
                            End If
                        End If
                    Next p
                Next h
                If Not codigosDuplicados = "" Then
                    mensajeFinal += "Los siguientes códigos aparecen más de una vez: " & vbNewLine & codigosDuplicados & vbNewLine & vbNewLine
                End If
            End If

            If loteOK = True Then
                For h = 0 To UBound(linea3)
                    If linea3(h).Contains(".RC") Or linea3(h).Contains("MAC700/") Then

                    Else
                        loteOK = False
                        If Not errComp.Contains(linea3(h)) Then
                            errComp += linea3(h) & vbNewLine
                        End If
                    End If
                Next h
                If Not errComp = "" Then
                    mensajeFinal += "Los siguientes códigos no deben ir en un paquete de MAC700: " & vbNewLine & errComp & vbNewLine & vbNewLine
                End If
            End If
            If loteOK = True Then
                For h = 0 To UBound(linea3)
                    If linea3(h).Contains(codigoLote + ".RC") Then
                        paquetesQueHay += 1
                    End If
                    If linea3(h).Contains("MAC700/") Then
                        mac700QueHay += 1
                    End If
                Next h
                mac700HayPorPaquete = mac700QueHay / paquetesQueHay
                If mac700HayPorPaquete <> mac700PorPaquete Then
                    mensajeFinal += "La cantidad de MAC700 por paquete difiere de " & mac700PorPaquete & vbNewLine
                    loteOK = False
                End If
            End If
            comprobacion()
            If errorRC = False And loteOK = True Then
                MsgBox("Después de darle al OK tendrás 5 segundos para seleccionar la ventana 'Pack a Road case'")
                volcado()
            ElseIf errorRC = False And loteOK = False Then
                MsgBox(mensajeFinal & vbNewLine & "Repasa los códigos.", MsgBoxStyle.MsgBoxSetForeground & MsgBoxStyle.Information, "Hay errores")
            End If
        End If

        'BARRAS DE LED Z8.
        If ComboBox1.Text = "Z8 Strip (4)" Then
            codigoLote = "Z8"
            limiteDeLotes = 4
            Dim errComp As String = ""
            Dim h As Integer = 0
            Dim p As Integer = 0
            Dim paquetesQueHay As Integer = 0
            Dim z8EnLista As Integer = 0
            Dim z8QueHay As Integer = 0
            Dim z8HayPorPaquete As Double = 0
            Dim z8PorPaquete As Integer = 6
            Dim loteOK As Boolean = True
            Dim mensajeFinal As String = ""
            If Not linea3(0).Contains(".RC") Then
                mensajeFinal += "El primer código del paquete tiene que ser el de la caja" & vbNewLine & vbNewLine
                loteOK = False
            End If
            If loteOK = True Then
                For h = 0 To UBound(linea3)
                    For p = 0 To UBound(linea3)
                        If Not p = h Then
                            If linea3(p) = linea3(h) Then
                                loteOK = False
                                If Not codigosDuplicados.Contains(linea3(p)) Then
                                    codigosDuplicados += linea3(p) & vbNewLine
                                End If
                            End If
                        End If
                    Next p
                Next h
                If Not codigosDuplicados = "" Then
                    mensajeFinal += "Los siguientes códigos aparecen más de una vez: " & vbNewLine & codigosDuplicados & vbNewLine & vbNewLine
                End If
            End If

            If loteOK = True Then
                For h = 0 To UBound(linea3)
                    If linea3(h).Contains(".RC") Or linea3(h).Contains("Z8STRIP/") Then

                    Else
                        loteOK = False
                        If Not errComp.Contains(linea3(h)) Then
                            errComp += linea3(h) & vbNewLine
                        End If
                    End If
                Next h
                If Not errComp = "" Then
                    mensajeFinal += "Los siguientes códigos no deben ir en un paquete de barras Z8: " & vbNewLine & errComp & vbNewLine & vbNewLine
                End If
            End If
            If loteOK = True Then
                For h = 0 To UBound(linea3)
                    If linea3(h).Contains(codigoLote + ".RC") Then
                        paquetesQueHay += 1
                    End If
                    If linea3(h).Contains("Z8STRIP/") Then
                        z8QueHay += 1
                    End If
                Next h
                z8HayPorPaquete = z8QueHay / paquetesQueHay
                If z8HayPorPaquete <> z8PorPaquete Then
                    mensajeFinal += "La cantidad de Z8STRIP por paquete difiere de " & z8PorPaquete & vbNewLine
                    loteOK = False
                End If
            End If
            comprobacion()
            If errorRC = False And loteOK = True Then
                MsgBox("Después de darle al OK tendrás 5 segundos para seleccionar la ventana 'Pack a Road case'")
                volcado()
            ElseIf errorRC = False And loteOK = False Then
                MsgBox(mensajeFinal & vbNewLine & "Repasa los códigos.", MsgBoxStyle.MsgBoxSetForeground & MsgBoxStyle.Information, "Hay errores")
            End If
        End If



        'LOTE SOCAPEX
        If ComboBox1.Text = "Socapex (1)" Then
            Dim errComp As String = ""
            Dim m As Integer = 0
            Dim h As Integer = 0
            Dim p As Integer = 0

            Dim soca20qty As Integer = 0
            Dim soca10qty As Integer = 0
            Dim soca05qty As Integer = 0
            Dim pulpoMqty As Integer = 0
            Dim pulpoHqty As Integer = 0
            Dim boxSocaqty As Integer = 0


            Dim soca20 As Integer = 4
            Dim soca10 As Integer = 4
            Dim soca05 As Integer = 4
            Dim pulpoM As Integer = 4
            Dim pulpoH As Integer = 2
            Dim boxSoca As Integer = 2
            Dim loteOK As Boolean = True
            Dim mensajeFinal As String = ""
            codigoLote = "SPX"
            limiteDeLotes = 1
            If Not linea3(0).Contains(".RC") Then
                mensajeFinal += "El primer código del paquete tiene que ser el de la caja" & vbNewLine & vbNewLine
                loteOK = False
            End If
            If loteOK = True Then
                For h = 0 To UBound(linea3)
                    For p = 0 To UBound(linea3)
                        If Not p = h Then
                            If linea3(p) = linea3(h) Then
                                loteOK = False
                                If Not codigosDuplicados.Contains(linea3(p)) Then
                                    codigosDuplicados += linea3(p) & vbNewLine
                                End If
                            End If
                        End If
                    Next p
                Next h
                If Not codigosDuplicados = "" Then
                    mensajeFinal += "Los siguientes códigos aparecen más de una vez: " & vbNewLine & codigosDuplicados & vbNewLine & vbNewLine
                End If
            End If

            If loteOK = True Then
                For h = 0 To UBound(linea3)
                    If linea3(h).Contains(".RC") Or
                                linea3(h).Contains("PB20/") Or linea3(h).Contains("SCPX20/") Or
                                linea3(h).Contains("PB10/") Or linea3(h).Contains("SCPX10/") Or
                                linea3(h).Contains("PB05/") Or linea3(h).Contains("SCPX05/") Or
                                linea3(h).Contains("SPXM/") Or linea3(h).Contains("SCPXM/") Or
                                linea3(h).Contains("SPXF/") Or linea3(h).Contains("SCPXF/") Or
                                linea3(h).Contains("SPX2SK/") Then
                    Else
                        loteOK = False
                        If Not errComp.Contains(linea3(h)) Then
                            errComp += linea3(h) & vbNewLine
                        End If
                    End If
                Next h
                If Not errComp = "" Then
                    mensajeFinal += "Los siguientes códigos no deben ir en el lote de soca: " & vbNewLine & errComp & vbNewLine & vbNewLine
                End If
            End If

            If loteOK = True Then
                For h = 0 To UBound(linea3)
                    If linea3(h).Contains("PB20/") Or linea3(h).Contains("SCPX20/") Then
                        soca20qty += 1
                    ElseIf linea3(h).Contains("PB10/") Or linea3(h).Contains("SCPX10/") Then
                        soca10qty += 1
                    ElseIf linea3(h).Contains("PB05/") Or linea3(h).Contains("SCPX05/") Then
                        soca05qty += 1
                    ElseIf linea3(h).Contains("SPXM/") Or linea3(h).Contains("SCPXM/") Then
                        pulpoMqty += 1
                    ElseIf linea3(h).Contains("SPXF/") Or linea3(h).Contains("SCPXF/") Then
                        pulpoHqty += 1
                    ElseIf linea3(h).Contains("SPX2SK/") Then
                        boxSocaqty += 1
                    End If
                Next h
                If soca20qty <> soca20 Then
                    mensajeFinal += "La cantidad de socapex de 20m difiere de " & soca20 & vbNewLine
                    loteOK = False
                End If
                If soca10qty <> soca10 Then
                    mensajeFinal += "La cantidad de socapex de 10m difiere de " & soca10 & vbNewLine
                    loteOK = False
                End If
                If soca05qty <> soca05 Then
                    mensajeFinal += "La cantidad de socapex de 5m difiere de " & soca05 & vbNewLine
                    loteOK = False
                End If
                If pulpoMqty <> pulpoM Then
                    mensajeFinal += "La cantidad de pulpos macho difiere de " & pulpoM & vbNewLine
                    loteOK = False
                End If
                If pulpoHqty <> pulpoH Then
                    mensajeFinal += "La cantidad de pulpos hembra difiere de " & pulpoH & vbNewLine
                    loteOK = False
                End If
                If boxSocaqty <> boxSoca Then
                    mensajeFinal += "La cantidad de cajetines hembra difiere de " & boxSoca & vbNewLine
                    loteOK = False
                End If
            End If
            comprobacion()
            If errorRC = False And loteOK = True Then
                MsgBox("Después de darle al OK tendrás 5 segundos para seleccionar la ventana 'Pack a Road case'")
                volcado()
            ElseIf errorRC = False And loteOK = False Then
                MsgBox(mensajeFinal & vbNewLine & "Repasa los códigos.", MsgBoxStyle.MsgBoxSetForeground & MsgBoxStyle.Information, "Hay errores")
            End If

        End If

        '6 x PAR LED 
        If ComboBox1.Text = "6 x Par Led (6)" Then
            codigoLote = "6PAR"
            limiteDeLotes = 6
            Dim errComp As String = ""
            Dim h As Integer = 0
            Dim p As Integer = 0
            Dim paquetesQueHay As Integer = 0
            Dim parledEnLista As Integer = 0
            Dim parledQueHay As Integer = 0
            Dim parledHayPorPaquete As Double = 0
            Dim parledPorPaquete As Integer = 6
            Dim loteOK As Boolean = True
            Dim mensajeFinal As String = ""
            If Not linea3(0).Contains(".RC") Then
                mensajeFinal += "El primer código del paquete tiene que ser el de la caja" & vbNewLine & vbNewLine
                loteOK = False
            End If
            If loteOK = True Then
                For h = 0 To UBound(linea3)
                    For p = 0 To UBound(linea3)
                        If Not p = h Then
                            If linea3(p) = linea3(h) Then
                                loteOK = False
                                If Not codigosDuplicados.Contains(linea3(p)) Then
                                    codigosDuplicados += linea3(p) & vbNewLine
                                End If
                            End If
                        End If
                    Next p
                Next h
                If Not codigosDuplicados = "" Then
                    mensajeFinal += "Los siguientes códigos aparecen más de una vez: " & vbNewLine & codigosDuplicados & vbNewLine & vbNewLine
                End If
            End If

            If loteOK = True Then
                For h = 0 To UBound(linea3)
                    If Not (linea3(h).Contains(".RC") Or linea3(h).Contains("PARPL/") Or linea3(h).Contains("PARSHL/") Or linea3(h).Contains("PARUV/")) Then
                        loteOK = False
                        If Not errComp.Contains(linea3(h)) Then
                            errComp += linea3(h) & vbNewLine
                        End If
                    End If
                Next h
                If Not errComp = "" Then
                    mensajeFinal += "Los siguientes códigos no deben ir en un paquete de 6 x Par Led: " & vbNewLine & errComp & vbNewLine & vbNewLine
                End If
            End If
            If loteOK = True Then
                For h = 0 To UBound(linea3)
                    If linea3(h).Contains(codigoLote + ".RC") Or linea3(h).Contains(codigoLoteB + ".RC") Then
                        paquetesQueHay += 1
                    End If
                    If linea3(h).Contains("PARPL/") Or linea3(h).Contains("PARSHL/") Or linea3(h).Contains("PARUV/") Then
                        parledQueHay += 1
                    End If
                Next h
                parledHayPorPaquete = parledQueHay / paquetesQueHay
                If parledHayPorPaquete <> parledPorPaquete Then
                    mensajeFinal += "La cantidad de Par Led por paquete difiere de " & parledPorPaquete & vbNewLine
                    loteOK = False
                End If
            End If
            comprobacion()
            If errorRC = False And loteOK = True Then
                MsgBox("Después de darle al OK tendrás 5 segundos para seleccionar la ventana 'Pack a Road case'")
                volcado()
            ElseIf errorRC = False And loteOK = False Then
                MsgBox(mensajeFinal & vbNewLine & "Repasa los códigos.", MsgBoxStyle.MsgBoxSetForeground & MsgBoxStyle.Information, "Hay errores")
            End If
        End If

        '8 x PAR LED 
        If ComboBox1.Text = "8 x Par Led (6)" Then
            codigoLote = "8PAR"
            limiteDeLotes = 6
            Dim errComp As String = ""
            Dim h As Integer = 0
            Dim p As Integer = 0
            Dim paquetesQueHay As Integer = 0
            Dim parledEnLista As Integer = 0
            Dim parledQueHay As Integer = 0
            Dim parledHayPorPaquete As Double = 0
            Dim parledPorPaquete As Integer = 8
            Dim loteOK As Boolean = True
            Dim mensajeFinal As String = ""
            If Not linea3(0).Contains(".RC") Then
                mensajeFinal += "El primer código del paquete tiene que ser el de la caja" & vbNewLine & vbNewLine
                loteOK = False
            End If
            If loteOK = True Then
                For h = 0 To UBound(linea3)
                    For p = 0 To UBound(linea3)
                        If Not p = h Then
                            If linea3(p) = linea3(h) Then
                                loteOK = False
                                If Not codigosDuplicados.Contains(linea3(p)) Then
                                    codigosDuplicados += linea3(p) & vbNewLine
                                End If
                            End If
                        End If
                    Next p
                Next h
                If Not codigosDuplicados = "" Then
                    mensajeFinal += "Los siguientes códigos aparecen más de una vez: " & vbNewLine & codigosDuplicados & vbNewLine & vbNewLine
                End If
            End If

            If loteOK = True Then
                For h = 0 To UBound(linea3)
                    If Not (linea3(h).Contains(".RC") Or linea3(h).Contains("PARPL/") Or linea3(h).Contains("PARSHL/") Or linea3(h).Contains("PARUV/")) Then
                        loteOK = False
                        If Not errComp.Contains(linea3(h)) Then
                            errComp += linea3(h) & vbNewLine
                        End If
                    End If
                Next h
                If Not errComp = "" Then
                    mensajeFinal += "Los siguientes códigos no deben ir en un paquete de 8 x Par Led: " & vbNewLine & errComp & vbNewLine & vbNewLine
                End If
            End If
            If loteOK = True Then
                For h = 0 To UBound(linea3)
                    If linea3(h).Contains(codigoLote + ".RC") Or linea3(h).Contains(codigoLoteB + ".RC") Then
                        paquetesQueHay += 1
                    End If
                    If linea3(h).Contains("PARPL/") Or linea3(h).Contains("PARSHL/") Or linea3(h).Contains("PARUV/") Then
                        parledQueHay += 1
                    End If
                Next h
                parledHayPorPaquete = parledQueHay / paquetesQueHay
                If parledHayPorPaquete <> parledPorPaquete Then
                    mensajeFinal += "La cantidad de Par Led por paquete difiere de " & parledPorPaquete & vbNewLine
                    loteOK = False
                End If
            End If
            comprobacion()
            If errorRC = False And loteOK = True Then
                MsgBox("Después de darle al OK tendrás 5 segundos para seleccionar la ventana 'Pack a Road case'")
                volcado()
            ElseIf errorRC = False And loteOK = False Then
                MsgBox(mensajeFinal & vbNewLine & "Repasa los códigos.", MsgBoxStyle.MsgBoxSetForeground & MsgBoxStyle.Information, "Hay errores")
            End If
        End If


ErrorHandler:
        Resume Next
    End Sub


End Class

