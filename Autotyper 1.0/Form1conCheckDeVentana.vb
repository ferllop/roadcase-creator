Imports System.Runtime.InteropServices
Imports System.IO
Imports System.Text

Public Class Form1
    Dim listado_de_paquetes() = {
        ({"Absen (6)", ({"A3.RC"}), 6,
            ({
                ({"TILES ABSEN", ({"SPA3"}), 6})
            })
        }),
        ({"Absen Back Support (6)", ({"A3BS.RC"}), 6,
            ({
                ({"BACK SUPPORTS", ({"A3BS", "A3PBS1"}), 6}),
                ({"BANDEJAS DE 1m", ({"A3TRAY", "A3PBT"}), 3})
            })
        }),
        ({"Absen Ground Beam 1 x 1.5m + 2 x 1m (4)", ({"A3GB1.RC"}), 4,
            ({
                ({"GROUND BEAM DE 1.5m", ({"A3GB1.5"}), 1}),
                ({"GROUND BEAM DE 1m", ({"A3GB1"}), 2}),
                ({"OUTRIGGERS", ({"A3OUT"}), 4})
            })
        }),
        ({"Absen Ground Beam 2 x 2m (4)", ({"A3GB2.RC"}), 4,
            ({
                ({"GROUND BEAM DE 2m", ({"A3GB2"}), 2}),
                ({"OUTRIGGERS", ({"A3OUT"}), 4})
            })
        }),
        ({"21K (4)", ({"21K.RC"}), 4,
            ({
                ({"PROYECTORES DE 21K", ({"PAN21K"}), 1}),
                ({"MANDOS CON HDMI PARA PROYECTOR PANASONIC", ({"RCPANMI", "RCPAN"}), 1}),
                ({"RENTAL FRAME PARA 21K", ({"RF21K"}), 1}),
                ({"CONVERSORES SCJUKO A 16A MONOFÁSICO", ({"SK16"}), 1})
            })
        }),
        ({"13K (4)", ({"13K.RC"}), 4,
            ({
                ({"PROYECTORES DE 13K", ({"PAN13K"}), 1}),
                ({"MANDOS CON HDMI PARA PROYECTOR PANASONIC", ({"RCPANMI", "RCPAN"}), 1}),
                ({"RENTAL FRAME PARA 13K", ({"RF13K"}), 1})
            })
        }),
        ({"10K (4)", ({"10K.RC"}), 4,
            ({
                ({"PROYECTORES DE 10K", ({"PAN10K"}), 1}),
                ({"MANDOS CON HDMI PARA PROYECTOR PANASONIC", ({"RCPANMI", "RCPAN"}), 1}),
                ({"RENTAL FRAME PARA 10K", ({"RF10K"}), 1})
            })
        }),
        ({"8.5K (4)", ({"8K.RC"}), 4,
            ({
                ({"PROYECTORES DE 8.5K", ({"PT870"}), 1}),
                ({"MANDOS CON HDMI PARA PROYECTOR PANASONIC", ({"RCPANMI", "RCPAN"}), 1}),
                ({"RENTAL FRAME PARA 8K", ({"RF8K"}), 1})
            })
        }),
        ({"85 pulgadas Mate o Brillo (4)", ({"85.RC"}), 4,
            ({
                ({"DISPLAYS DE 85 PULGADAS", ({"QM85D", "QM85NMAT"}), 1}),
                ({"MANDOS A DISTANCIA PARA DISPLAY SAMSUNG", ({"SRC1"}), 1})
            })
        }),
        ({"80 pulgadas (4)", ({"80.RC"}), 4,
            ({
                ({"DISPLAYS DE 80 PULGADAS", ({"SHARP80"}), 1}),
                ({"MANDOS A DISTANCIA PARA DISPLAY DE 80 PULGADAS", ({"HRC"}), 1}),
                ({"SOPORTES PARA DISPLAY DE 80 PULGADAS", ({"SOP80"}), 1}),
                ({"LLAVES ALLEN PARA EL SOPORTE DE DISPLAY DE 80 PULGADAS", ({"80KEY"}), 1})
            })
        }),
        ({"1 x 65 pulgadas (4)", ({"1U65.RC"}), 4,
            ({
                ({"DISPLAYS DE 65 PULGADAS", ({"QM65H"}), 1}),
                ({"MANDOS A DISTANCIA PARA DISPLAY SAMSUNG", ({"SRC1"}), 1})
            })
        }),
        ({"2 x 65 pulgadas (4)", ({"2U65.RC"}), 4,
            ({
                ({"DISPLAYS DE 65 PULGADAS", ({"QM65H"}), 2}),
                ({"MANDOS A DISTANCIA PARA DISPLAY SAMSUNG", ({"SRC1"}), 1})
            })
        }),
        ({"3 x 65 pulgadas (4)", ({"3U65.RC"}), 4,
            ({
                ({"DISPLAYS DE 65 PULGADAS", ({"QM65H"}), 3}),
                ({"MANDOS A DISTANCIA PARA DISPLAY SAMSUNG", ({"SRC1"}), 1})
            })
        }),
        ({"2 x 55 pulgadas 4K QM55H (4)", ({"2QM55H.RC"}), 4,
            ({
                ({"DISPLAYS DE 55 PULGADAS QM55H", ({"QM55H"}), 2}),
                ({"MANDOS A DISTANCIA PARA DISPLAY SAMSUNG", ({"SRC1"}), 1})
            })
        }),
        ({"3 x 55 pulgadas 4K QM55H (4)", ({"3QM55H.RC"}), 4,
            ({
                ({"DISPLAYS DE 55 PULGADAS QM55H", ({"QM55H"}), 3}),
                ({"MANDOS A DISTANCIA PARA DISPLAY SAMSUNG", ({"SRC1"}), 1})
            })
        }),
        ({"2 x 55 pulgadas ME55C (4)", ({"ME55C.RC"}), 4,
            ({
                ({"DISPLAYS DE 55 PULGADAS ME55C", ({"ME55C"}), 2}),
                ({"MANDOS A DISTANCIA PARA DISPLAY SAMSUNG", ({"SRC1"}), 1})
            })
        }),
        ({"2 x 55 pulgadas DB55D (4)", ({"DB55D.RC"}), 4,
            ({
                ({"DISPLAYS DE 55 PULGADAS DB55D", ({"DB55D"}), 2}),
                ({"MANDOS A DISTANCIA PARA DISPLAY SAMSUNG", ({"SRC1"}), 1})
            })
        }),
        ({"2 x 49 pulgadas QB49N (4)", ({"2QB49N.RC"}), 4,
            ({
                ({"DISPLAYS DE 49 PULGADAS QB49N", ({"QB49N"}), 2}),
                ({"MANDOS A DISTANCIA PARA DISPLAY SAMSUNG", ({"SRC1"}), 1})
            })
        }),
        ({"3 x 49 pulgadas QB49N (4)", ({"3QB49N.RC"}), 4,
            ({
                ({"DISPLAYS DE 49 PULGADAS QB49N", ({"QB49N"}), 3}),
                ({"MANDOS A DISTANCIA PARA DISPLAY SAMSUNG", ({"SRC1"}), 1})
            })
        }),
        ({"1 x 48 pulgadas DB48E (4)", ({"1DB48E.RC"}), 4,
            ({
                ({"DISPLAYS DE 48 PULGADAS DB48E", ({"DB48E"}), 1}),
                ({"MANDOS A DISTANCIA PARA DISPLAY SAMSUNG", ({"SRC1"}), 1})
            })
        }),
        ({"2 x 48 pulgadas DB48E (4)", ({"2DB48E.RC"}), 4,
            ({
                ({"DISPLAYS DE 48 PULGADAS DB48E", ({"DB48E"}), 2}),
                ({"MANDOS A DISTANCIA PARA DISPLAY SAMSUNG", ({"SRC1"}), 1})
            })
        }),
        ({"2 x 46 pulgadas ME46C (4)", ({"ME46C.RC"}), 4,
            ({
                ({"DISPLAYS DE 46 PULGADAS DB48E", ({"ME46C"}), 2}),
                ({"MANDOS A DISTANCIA PARA DISPLAY SAMSUNG", ({"SRC1"}), 1})
            })
        }),
        ({"Clevertouch 55 pulgadas (4)", ({"CLEV55.RC"}), 4,
            ({
                ({"CLEVERTOUCH 55 PULGADAS", ({"CLEV55P"}), 2}),
                ({"MANDOS A DISTANCIA PARA CLEVERTOUCH 55 PULGADAS", ({"CLEVERC"}), 2}),
                ({"PUNTEROS PARA CLEVERTOUCH 55 PULGADAS", ({"RUBPEN"}), 2}),
                ({"CABLES USB PARA CLEVERTOUCH 55 PULGADAS", ({"CLUS"}), 2}),
                ({"ANTENAS PARA CLEVERTOUCH 55 PULGADAS", ({"2458G"}), 2})
            })
        }),
        ({"NEC X55 (5)", ({"X55.RC"}), 5,
            ({
                ({"DISPLAY NEC X55", ({"X55"}), 2}),
                ({"MALETIN PARA X55", ({"X55AS"}), 2}),
                ({"EASYFRAME FRAME PARA X555UNS", ({"X55FR"}), 2}),
                ({"STACKER FRAME PARA X555UNS", ({"X55SF"}), 2}),
                ({"EASYFRAME DOUBLE MAGNETIC HOLDER", ({"2M"}), 2}),
                ({"EASYFRAME QUADRUPLE MAGNETIC HOLDER", ({"4M"}), 2})
            })
        }),
        ({"4 x T10 (8)", ({"4T10.RC"}), 8,
            ({
                ({"D&B T10", ({"T10"}), 4})
            })
        }),
        ({"MAUI (4)", ({"2G2.RC"}), 4,
            ({
                ({"SUBGRAVES MAUI", ({"G2SUB"}), 2}),
                ({"TOP INFERIOR MAUI", ({"G2LTOP"}), 2}),
                ({"TOP INFERIOR MAUI", ({"G2UTOP"}), 2})
            })
        }),
        ({"QL5 (4)", ({"QL5.RC"}), 4,
            ({
                ({"YAMAHA QL5", ({"YQL548"}), 1})
            })
        }),
        ({"Axiom (6)", ({"AX.RC"}), 6,
            ({
                ({"MARTIN MAC AXIOM", ({"AXIOM"}), 2})
            })
        }),
        ({"MAC700 (6)", ({"700.RC"}), 6,
            ({
                ({"MARTIN MAC 700", ({"MAC700"}), 2})
            })
        }),
        ({"Z8 Strip (4)", ({"Z8.RC"}), 4,
            ({
                ({"BARRAS Z8", ({"Z8STRIP"}), 2})
            })
        }),
        ({"Socapex (1)", ({"SPX.RC"}), 5,
            ({
                ({"SOCAPEX DE 20m", ({"PB20", "SCPX20"}), 4}),
                ({"SOCAPEX DE 10m", ({"PB10", "SCPX10"}), 4}),
                ({"SOCAPEX DE 5m", ({"PB05", "SCPX05"}), 4}),
                ({"PULPOS SOCAPEX MACHO", ({"SPXM", "SCPXM"}), 4}),
                ({"PULPOS SOCAPEX HEMBRA", ({"SPXF", "SCPXF"}), 2}),
                ({"CAJETINES HEMBRA", ({"SPX2SK"}), 2})
            })
        }),
        ({"6 x Par Led (6)", ({"6PAR.RC"}), 6,
            ({
                ({"Parleds", ({"PARPL", "PARSHL", "PARUV"}), 6})
            })
        }),
        ({"8 x Par Led (6)", ({"8PAR.RC"}), 6,
            ({
                ({"Parleds", ({"PARPL", "PARSHL", "PARUV"}), 8})
            })
        })
    }
    Dim windowError As Boolean
    Dim paraloTodo As Boolean
    'Dim listado_de_paquetes2 As String = File.ReadAllText("\\HPACTION\Almacen\macros-No-borrar-nunca\paquetes-test.txt")
    'Dim arrayPaquetes() = listado_de_paquetes2.Split(vbCrLf)
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

    Function fillComboBox(ByVal paquetes())
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
    Function cleanLine(ByVal inputText As String) As String
        inputText = Replace(inputText, vbTab, "")
        inputText = Replace(inputText, " ", "")
        inputText = Replace(inputText, vbCrLf & vbCrLf, vbCrLf)
        inputText = inputText.Trim(vbCrLf.ToCharArray)
        Return inputText
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
                            Wait(0.2)
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
                                    Wait(2)
                                    SendKeys.Send(nuevoroadcase)
                                End If
                            Else
                                Exit For
                            End If
                        Else
                            SendKeys.Send(linea(i) & "{Enter}")
                            Wait(0.2)
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

    Function tipo_de_paquete_erroneo(ByRef listaDeCodigos, ByRef paquete)


        Dim x As Integer = 0

        For x = 0 To UBound(listaDeCodigos)
            If listaDeCodigos(x).Contains(".RC") Then
                Dim codigo = listaDeCodigos(x).Remove(listaDeCodigos(x).LastIndexOf("/"))
                If Array.IndexOf(paquete(1), codigo) < 0 Then
                    Return True
                End If
            End If
        Next x


        Return False
    End Function
    Function cantidad_de_paquetes_erronea(ByRef listaDeCodigos, ByRef paquete)

        Dim h As Integer = 0
        Dim x As Integer = 0
        Dim y As Integer = 0
        Dim errorRC As Boolean = False

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
    Function primer_codigo_es_de_caja(ByRef listaDeCodigos As String()) As Boolean
        If Not listaDeCodigos(0).Contains(".RC") Then
            Return True
        End If
        Return False
    End Function
    Function listado_codigos_duplicados(ByRef listaDeCodigos As String()) As String
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
    Function listado_componentes_erroneos(ByRef listaDeCodigos, ByRef codigosCaja, ByRef contenido())
        Dim errComp As String = ""
        Dim i As Integer = 0

        For i = 0 To UBound(listaDeCodigos)
            If Not listaDeCodigos(i).contains(".RC") Then
                Dim codigo = listaDeCodigos(i).remove(listaDeCodigos(i).LastIndexOf("/"))
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
    Function cantidades_cajas(ByRef listaDeCodigos)
        Dim cajasQueHay As Integer = 0

        For h = 0 To UBound(listaDeCodigos)
            If listaDeCodigos(h).Contains(".RC") Then
                cajasQueHay += 1
            End If
        Next h

        Return cajasQueHay
    End Function
    Function listado_cantidades_contenido_incorrectas(ByRef listaDeCodigos, ByRef codigosCaja, ByRef contenido())
        Dim cajasQueHay As Integer = cantidades_cajas(listaDeCodigos)
        Dim productosQueHay() = {}
        Dim productoHayPorPaquete As Integer = 0
        Dim productos_incorrectos() = {}
        Dim j As Integer
        For Each producto In contenido
            ReDim Preserve productosQueHay(productosQueHay.Length)
            productosQueHay(productosQueHay.Length - 1) = {producto(0), 0, producto(2)}
        Next
        For i = 0 To UBound(listaDeCodigos)

            Dim codigo = listaDeCodigos(i).remove(listaDeCodigos(i).LastIndexOf("/"))
            If Array.IndexOf(codigosCaja, codigo) < 0 Then

                For Each producto In contenido
                    If Array.IndexOf(producto(1), codigo) > -1 Then

                        Dim found As Boolean = False
                        If productosQueHay.Length > 0 Then
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
        fillComboBox(listado_de_paquetes)
    End Sub

    Sub Button1_Click() Handles Button1.Click
        'paquetes = Replace(paquetes, "[", "({")
        'paquetes = Replace(paquetes, "]", "})")
        'paquetes = Replace(paquetes, vbCrLf, "")
        'paquetes = cleanLine(paquetes)
        'Dim listadoPaquetes() As Object = paquetes

        TextBox1.Text = cleanLine(TextBox1.Text)
        Dim linea As String() = Split(TextBox1.Text, vbCrLf)

        For Each element In listado_de_paquetes
            If element(0) = ComboBox1.Text Then
                Dim mensajeFinal As String = ""

                If mensajeFinal = "" Then
                    If tipo_de_paquete_erroneo(linea, element) Then
                        mensajeFinal += "Se ha encontrado algun paquete que no es de " & element(0) & vbNewLine & vbNewLine
                    End If
                End If

                If mensajeFinal = "" Then
                    Dim duplicados = listado_codigos_duplicados(linea)
                    If Not duplicados = "" Then
                        mensajeFinal += "Los siguientes códigos aparecen más de una vez:  " & vbNewLine & duplicados & vbNewLine & vbNewLine
                    End If
                End If


                If mensajeFinal = "" Then
                    If cantidad_de_paquetes_erronea(linea, element) Then
                        mensajeFinal += "No puedes hacer más de " & element(2) & " paquetes de " & element(0) & " de golpe." & vbNewLine & vbNewLine
                    End If
                End If

                If mensajeFinal = "" Then
                    If primer_codigo_es_de_caja(linea) Then
                        mensajeFinal += "El primer código del paquete tiene que ser el de la caja" & vbNewLine & vbNewLine
                    End If
                End If

                If mensajeFinal = "" Then
                    Dim componentes_erroneos As String = listado_componentes_erroneos(linea, element(1), element(3))
                    If Not componentes_erroneos = "" Then
                        mensajeFinal += "Los siguientes códigos no deben ir en un paquete de " & ComboBox1.Text & ":" & vbNewLine & componentes_erroneos & vbNewLine & vbNewLine
                    End If
                End If



                If mensajeFinal = "" Then
                    Dim productos_con_cantidades_erroneas() = {}
                    productos_con_cantidades_erroneas = listado_cantidades_contenido_incorrectas(linea, element(1), element(3))
                    If productos_con_cantidades_erroneas.Length > 0 Then
                        For Each producto_con_cantidad_erronea In productos_con_cantidades_erroneas
                            mensajeFinal += "En algún paquete de " & element(0) & " la cantidad de " & producto_con_cantidad_erronea(0) & " difiere de " & producto_con_cantidad_erronea(2) & vbNewLine
                        Next
                    End If
                End If



                If mensajeFinal = "" Then
                    MsgBox("Después de darle al OK tendrás 5 segundos para seleccionar la ventana 'Pack a Road case'")
                    volcado()
                Else
                    MsgBox(mensajeFinal & vbNewLine & "Repasa los códigos.", MsgBoxStyle.MsgBoxSetForeground & MsgBoxStyle.Information, "Hay errores")
                End If

            End If
        Next

    End Sub


End Class

