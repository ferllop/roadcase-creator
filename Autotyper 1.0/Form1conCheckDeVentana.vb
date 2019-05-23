Imports System.Runtime.InteropServices
Imports System.IO
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

    Dim array_paquetes() = {}

    Function PreparePaquetes(ByRef paquetesReadedFile As String)

        Dim paquetesFinalArray() = {}
        Dim arrayPaquetes As String() = paquetesReadedFile.Split(";")

        For i = 0 To arrayPaquetes.Length - 1

            arrayPaquetes(i) = arrayPaquetes(i).Replace(vbCr, "").Replace(vbLf, "").Replace(vbTab, "")
            Dim configPaqueteString As String = Strings.Left(arrayPaquetes(i), arrayPaquetes(i).IndexOf("#"))
            Dim componentesPaqueteString As String = Strings.Right(arrayPaquetes(i), arrayPaquetes(i).Length - arrayPaquetes(i).IndexOf("#"))

            Dim configPaqueteArray As String() = configPaqueteString.Split(",")


            Dim paqueteFinalArray(3)
            paqueteFinalArray(1) = configPaqueteArray(1).Replace(" ", "").ToUpper().Split("|") 'Códigos paquete
            paqueteFinalArray(2) = Convert.ToInt32(configPaqueteArray(2)) 'Limite paquetes

            Dim prepareShowLimiteLotes As String
            If paqueteFinalArray(2) = 0 Then
                prepareShowLimiteLotes = " (ilimitado)"
            Else
                prepareShowLimiteLotes = " (" & configPaqueteArray(2).Trim & ")"
            End If
            paqueteFinalArray(0) = configPaqueteArray(0).Trim & prepareShowLimiteLotes  'Descripción paquete

            Dim componentesPaqueteArray As String() = componentesPaqueteString.Trim("#").Split("#")

            Dim componentesFinalArray() = {}
            For j = 0 To componentesPaqueteArray.Length - 1
                Dim componenteArray As String() = componentesPaqueteArray(j).Split(",")

                Dim componenteFinalArray(2)
                componenteFinalArray(0) = componenteArray(0).Trim 'Descripción componente
                componenteFinalArray(1) = componenteArray(1).Replace(" ", "").ToUpper().Split("|") 'Códigos componente
                componenteFinalArray(2) = Convert.ToInt32(componenteArray(2)) 'Cantidad componente

                ReDim Preserve componentesFinalArray(componentesFinalArray.Length)
                componentesFinalArray(componentesFinalArray.Length - 1) = componenteFinalArray
            Next

            paqueteFinalArray(3) = componentesFinalArray

            ReDim Preserve paquetesFinalArray(paquetesFinalArray.Length)
            paquetesFinalArray(paquetesFinalArray.Length - 1) = paqueteFinalArray

        Next

        Return paquetesFinalArray

    End Function
    Function FillComboBox(ByVal paquetes())
        ComboBox1.MaxDropDownItems = paquetes.Length
        Dim dropdown_content() = {}

        For i = 0 To paquetes.Length - 1
            ReDim Preserve dropdown_content(dropdown_content.Length)
            dropdown_content(dropdown_content.Length - 1) = paquetes(i)(0)
        Next
        For i = 0 To dropdown_content.Length - 1
            ComboBox1.Items.Add(dropdown_content(i))
        Next
        Return True
    End Function

    Function Wait(ByVal dblSecs As Double)

        Const OneSec As Double = 1.0# / (1440.0# * 60.0#)
        Dim dblWaitTil As Date
        Now.AddSeconds(OneSec)
        dblWaitTil = Now.AddSeconds(OneSec).AddSeconds(dblSecs)
        Do Until Now > dblWaitTil
            Application.DoEvents() ' Allow windows messages to be processed
        Loop
        Return True
    End Function
    Function CleanLines(ByVal inputText As String) As String
        Return inputText.Replace(vbTab, "").Replace(" ", "").Replace(vbLf, vbCr).Replace(vbCr & vbCr & vbCr, vbCr).Replace(vbCr & vbCr, vbCr).Replace(vbCr, vbCrLf).Trim(vbCrLf.ToCharArray)
    End Function
    Function WindowByTitleCorrectlyFocused(ByVal windowTitle As String)
        While True
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
                If wTitle.ToString.Contains(windowTitle) And Not wTitle.ToString.Contains("Inventory Setup") Then
                    Return True
                Else
                    Dim response As MsgBoxResult = MsgBox("La ventana 'Pack a Road case' no quedó seleccionada." & vbNewLine & "Selecciona la ventana 'Pack a Road case' y clicka 'YES'." & vbNewLine & "Selecciona NO para parar el proceso.", MsgBoxStyle.MsgBoxSetForeground & MsgBoxStyle.YesNo, "¡¡¡La ventana perdió el foco!!!")
                    If response = MsgBoxResult.No Then
                        Return False
                    End If
                End If
            End If
        End While

        Return False

    End Function
    Function Volcado()
        Dim okroadcase As String = "+{TAB 2}{ENTER}"
        Dim nuevoroadcase As String = "{HOME}{DOWN}{TAB}{ENTER}{TAB 2}{ENTER}"
        Dim tick As String = "{Enter}{TAB}{RIGHT} {LEFT 2}"
        Dim linea As String()
        Dim codigos As String = TextBox1.Text
        Dim i As Integer = 0
        Dim roadcasesHechos As String = ""
        Dim lineaConRC As String = ""
        Dim textSinEspacios As String = ""
        Dim lineaRaw As String() = Split(codigos, vbCrLf)

        Wait(5)

        If WindowByTitleCorrectlyFocused("Pack a Road case") Then
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
                        If Not WindowByTitleCorrectlyFocused("Pack a Road") Then
                            Exit For
                        Else
                            SendKeys.Send(linea(i))
                            SendKeys.Send(tick)
                        End If
                    Else
                        If linea(i + 1).Contains(".RC") Then
                            If Not WindowByTitleCorrectlyFocused("Pack a Road") Then
                                Exit For
                            Else
                                SendKeys.Send(linea(i) & "{Enter}")
                                Wait(0.2)
                            End If


                            Dim response As MsgBoxResult = MsgBox("Roadcase completo." & vbNewLine & "Si seleccionas 'YES' el proceso continuará." & vbNewLine & "Si seleccionas NO el proceso finalizará y deberás repasar lo que se pueda haber quedado a medias.", MsgBoxStyle.MsgBoxSetForeground & MsgBoxStyle.YesNo, "Roadcase completo")
                            If response = MsgBoxResult.Yes Then

                                If Not WindowByTitleCorrectlyFocused("Pack a Road") Then
                                    Exit For
                                Else
                                    SendKeys.Send(okroadcase & "{Enter}")
                                    Wait(2)
                                End If
                                If Not WindowByTitleCorrectlyFocused("RentalPoint") Then
                                    Exit For
                                Else
                                    SendKeys.Send(nuevoroadcase)
                                    Wait(0.5)
                                End If
                            Else
                                Exit For
                            End If
                        Else
                            If Not WindowByTitleCorrectlyFocused("Pack a Road") Then
                                Exit For
                            Else
                                SendKeys.Send(linea(i) & "{Enter}")
                                Wait(0.2)
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

    Function Tipo_de_paquete_erroneo(ByRef listaDeCodigos, ByRef paquete)


        Dim x As Integer = 0

        For x = 0 To UBound(listaDeCodigos)
            If listaDeCodigos(x).Contains(".RC") Then
                Dim codigo As String = listaDeCodigos(x).Remove(listaDeCodigos(x).LastIndexOf("/"))
                If Array.IndexOf(paquete(1), codigo) < 0 Then
                    Return True
                End If
            End If
        Next x


        Return False
    End Function
    Function Cantidad_de_paquetes_erronea(ByRef listaDeCodigos, ByRef paquete)

        Dim h As Integer = 0
        Dim x As Integer = 0
        Dim y As Integer = 0
        Dim errorRC As Boolean = False

        If paquete(2) = 0 Then
            Return False
        End If

        For x = 0 To UBound(listaDeCodigos)
            If listaDeCodigos(x).Contains(".RC") Then
                Dim codigo = listaDeCodigos(x).Remove(listaDeCodigos(x).LastIndexOf("/"))
                If Array.IndexOf(paquete(1), codigo) > -1 Then
                    y = y + 1
                End If
            End If
        Next x


        If y > paquete(2) Then
            errorRC = True
        End If

        Return errorRC
    End Function
    Function Primer_codigo_es_de_caja(ByRef listaDeCodigos As String()) As Boolean
        If Not listaDeCodigos(0).Contains(".RC") Then
            Return True
        End If
        Return False
    End Function
    Function Listado_codigos_duplicados(ByRef listaDeCodigos As String()) As String
        Dim codigosDuplicados As String = ""

        Dim h As Integer = 0
        Dim p As Integer = 0
        For h = 0 To UBound(listaDeCodigos)
            For p = 0 To UBound(listaDeCodigos)
                If Not p = h Then
                    If listaDeCodigos(p) = listaDeCodigos(h) Then
                        If Not codigosDuplicados.Contains(listaDeCodigos(p)) Then
                            codigosDuplicados += listaDeCodigos(p) & vbNewLine
                        End If
                    End If
                End If
            Next p
        Next h
        Return codigosDuplicados
    End Function
    Function Listado_componentes_erroneos(ByRef listaDeCodigos, ByRef codigosCaja, ByRef contenido())
        Dim errComp As String = ""
        Dim i As Integer = 0

        For i = 0 To UBound(listaDeCodigos)
            If Not listaDeCodigos(i).contains(".RC") Then
                Dim codigo = listaDeCodigos(i)
                If codigo.contains("/") Then
                    codigo = listaDeCodigos(i).remove(listaDeCodigos(i).LastIndexOf("/"))
                End If

                Dim found As Boolean = False

                For Each product In contenido
                    If Array.IndexOf(codigosCaja, codigo) > -1 Or Array.IndexOf(product(1), codigo) > -1 Then
                        found = True
                        Exit For
                    End If
                Next
                If Not found And Not errComp.Contains(listaDeCodigos(i)) Then
                    errComp += listaDeCodigos(i) & vbNewLine
                End If
            End If
        Next i

        Return errComp
    End Function
    Function Cantidades_cajas(ByRef listaDeCodigos)
        Dim cajasQueHay As Integer = 0

        For h = 0 To UBound(listaDeCodigos)
            If listaDeCodigos(h).Contains(".RC") Then
                cajasQueHay += 1
            End If
        Next h

        Return cajasQueHay
    End Function
    Function Listado_cantidades_contenido_incorrectas(ByRef listaDeCodigos, ByRef codigosCaja, ByRef contenido())
        Dim cajasQueHay As Integer = Cantidades_cajas(listaDeCodigos)
        Dim productosQueHay() = {}
        Dim productoHayPorPaquete As Integer = 0
        Dim productos_incorrectos() = {}
        Dim j As Integer
        For Each producto In contenido
            ReDim Preserve productosQueHay(productosQueHay.Length)
            productosQueHay(productosQueHay.Length - 1) = {producto(0), 0, producto(2)}
        Next
        For i = 0 To UBound(listaDeCodigos)

            Dim codigo = listaDeCodigos(i)
            If codigo.contains("/") Then
                codigo = listaDeCodigos(i).remove(listaDeCodigos(i).LastIndexOf("/"))
            End If

            If Array.IndexOf(codigosCaja, codigo) < 0 Then

                For Each producto In contenido
                    If Array.IndexOf(producto(1), codigo) > -1 Then

                        Dim found As Boolean = False
                        If productosQueHay.Length > 0 And producto(2) > 0 Then
                            For j = 0 To productosQueHay.Length - 1
                                If producto(0) = productosQueHay(j)(0) Then
                                    productosQueHay(j)(1) += 1
                                    Exit For
                                End If
                            Next j
                        End If
                        Exit For
                    End If
                Next
            End If
        Next i

        If productosQueHay.Length > 0 Then
            For i = 0 To productosQueHay.Length - 1
                If productosQueHay(i)(1) / cajasQueHay <> productosQueHay(i)(2) Then
                    ReDim Preserve productos_incorrectos(productos_incorrectos.Length)
                    productos_incorrectos(productos_incorrectos.Length - 1) = productosQueHay(i)
                End If
            Next
        End If
        Return productos_incorrectos
    End Function

    Sub Form1_Load() Handles MyBase.Load
        Dim listado_de_paquetes As String = File.ReadAllText(".\Roadcase Creator Database.txt")
        array_paquetes = PreparePaquetes(listado_de_paquetes)
        FillComboBox(array_paquetes)
    End Sub

    Sub Button1_Click() Handles Button1.Click
        TextBox1.Text = CleanLines(TextBox1.Text)
        Dim linea As String() = Split(TextBox1.Text, vbCrLf)

        For Each element In array_paquetes
            If element(0) = ComboBox1.Text Then
                Dim mensajeFinal As String = ""

                If mensajeFinal = "" Then
                    If Tipo_de_paquete_erroneo(linea, element) Then
                        mensajeFinal += "Se ha encontrado algun paquete que no es de " & element(0) & vbNewLine & vbNewLine
                    End If
                End If

                If mensajeFinal = "" Then
                    Dim duplicados = Listado_codigos_duplicados(linea)
                    If Not duplicados = "" Then
                        mensajeFinal += "Los siguientes códigos aparecen más de una vez:  " & vbNewLine & duplicados & vbNewLine & vbNewLine
                    End If
                End If


                If mensajeFinal = "" Then
                    If Cantidad_de_paquetes_erronea(linea, element) Then
                        mensajeFinal += "No puedes hacer más de " & element(2) & " paquetes de " & element(0) & " de golpe." & vbNewLine & vbNewLine
                    End If
                End If

                If mensajeFinal = "" Then
                    If Primer_codigo_es_de_caja(linea) Then
                        mensajeFinal += "El primer código del paquete tiene que ser el de la caja" & vbNewLine & vbNewLine
                    End If
                End If

                If mensajeFinal = "" Then
                    Dim componentes_erroneos As String = Listado_componentes_erroneos(linea, element(1), element(3))
                    If Not componentes_erroneos = "" Then
                        mensajeFinal += "Los siguientes códigos no deben ir en un paquete de " & ComboBox1.Text & ":" & vbNewLine & componentes_erroneos & vbNewLine & vbNewLine
                    End If
                End If



                If mensajeFinal = "" Then
                    Dim productos_con_cantidades_erroneas() = {}
                    productos_con_cantidades_erroneas = Listado_cantidades_contenido_incorrectas(linea, element(1), element(3))
                    If productos_con_cantidades_erroneas.Length > 0 Then
                        For Each producto_con_cantidad_erronea In productos_con_cantidades_erroneas
                            mensajeFinal += "En algún paquete de " & element(0) & " la cantidad de " & producto_con_cantidad_erronea(0) & " difiere de " & producto_con_cantidad_erronea(2) & vbNewLine
                        Next
                    End If
                End If



                If mensajeFinal = "" Then
                    MsgBox("Después de darle al OK tendrás 5 segundos para seleccionar la ventana 'Pack a Road case'")
                    Volcado()
                Else
                    MsgBox(mensajeFinal & vbNewLine & "Repasa los códigos.", MsgBoxStyle.MsgBoxSetForeground & MsgBoxStyle.Information, "Hay errores")
                End If

            End If
        Next

    End Sub

End Class

