Imports ADODB
Imports System.Data.SqlClient
Imports System.Drawing.Printing


Public Class Form1

    Private DB_CONN As String
    Private TIPO_DOC As String
    Private SQLtext As String
    Private pd As PrintDocument

    Private CURRENT_IMPRESORA As String
    Private NUMERO_CARACTERES_X_LINEA_TICKET As Integer
    Private TAMANIO_FUENTE As Single
    Private NOMBRE_FUENTE As String
    Private FINAL_LINEA As String = Chr(13) & Chr(10)
    Private CURRENT_TEXT_TO_PRINT As String

    Private CURRENT_ESTACION As Recordset
    Private EMPRESA As Recordset

    Private Sub Form1_Load(sender As Object, e As System.EventArgs) Handles Me.Load
        Dim DB_BRUTO As String
        Dim NUM_TICKET As String
        Dim ESTACION As String

        If Environment.GetCommandLineArgs.Length >= 3 Then

            TIPO_DOC = Environment.GetCommandLineArgs(1)
            DB_BRUTO = Environment.GetCommandLineArgs(2)
            DB_CONN = Replace(DB_BRUTO, "%", " ")
            ESTACION = Environment.GetCommandLineArgs(3)
            NUM_TICKET = Environment.GetCommandLineArgs(4)

            TextBox1.AppendText(TIPO_DOC)
            TextBox1.AppendText(FINAL_LINEA)
            TextBox1.AppendText(DB_CONN)
            TextBox1.AppendText(FINAL_LINEA)
            TextBox1.AppendText(ESTACION)
            TextBox1.AppendText(FINAL_LINEA)
            TextBox1.AppendText(NUM_TICKET)


            MsgBox("Probando", MsgBoxStyle.Information, "Titulo")


            'Provider=SQLNCLI.1;Integrated Security=SSPI;Persist Security Info=True;Initial Catalog=C:\MyBusinessDatabase\MyBusinessPOS2012.mdf;Data Source=JONA_LAP;Use Procedure for Prepare=1;Auto Translate=True;Packet Size=4096;Workstation ID=JONA_LAP;Use Encryption for Data=False;Tag with column collation when possible=False;MARS Connection=False;DataTypeCompatibility=80;Trust Server Certificate=False


            'ticket
            'Provider=SQLNCLI.1;Password=12345678;Persist Security Info=True;User ID=sa;Initial Catalog=C:\MyBusinessDatabase\MyBusinessPOS2012.mdf;Data Source=TCP:.\SQLEXPRESS,1400;Use Procedure for Prepare=1;Auto Translate=True;Packet Size=4096;Workstation ID=SONOFGOD-PC;Use Encryption for Data=False;Tag with column collation when possible=False;MARS Connection=False;DataTypeCompatibility=0;Trust Server Certificate=False
            '19:

        Else
            'TextBox1.Text = "No se han indicado parámetros en la línea de comandos" & vbCrLf & _
            '                "El nombre (y path) del ejecutable es:" & vbCrLf & _
            ' Environment.GetCommandLineArgs(0)

            TIPO_DOC = "ticket"
            DB_CONN = "Provider=SQLNCLI.1;Password=12345678;Persist Security Info=True;User ID=sa;Initial Catalog=C:\MyBusinessDatabase\MyBusinessPOS2012.mdf;Data Source=SONOFGOD-PC\SQLEXPRESS,1433;Use Procedure for Prepare=1;Auto Translate=True;Packet Size=4096;Workstation ID=SONOFGOD-PC;Use Encryption for Data=False;Tag with column collation when possible=False;MARS Connection=False;DataTypeCompatibility=0;Trust Server Certificate=False;"
            ESTACION = "BODEGA01"
            NUM_TICKET = "25"

        End If

        EMPRESA = crearRecorset("SELECT * FROM econfig")
        CURRENT_ESTACION = crearRecorset("SELECT * FROM estaciones WHERE Estacion = '" & ESTACION & "'")

        If CURRENT_ESTACION.EOF Then
            MsgBox("No se encuentra la estación solicitada: " & ESTACION, vbInformation)
            Exit Sub
        End If

        CURRENT_IMPRESORA = CURRENT_ESTACION.Fields("pticket").Value.ToString


        If TIPO_DOC = "ticket" Then

            TAMANIO_FUENTE = 8.5
            NUMERO_CARACTERES_X_LINEA_TICKET = 46
            NOMBRE_FUENTE = "MS Mincho"
            CURRENT_TEXT_TO_PRINT = construirTicket(NUM_TICKET)

            PrintDocument1.PrinterSettings.PrinterName = CURRENT_IMPRESORA
            Dim margins As New Margins(0, 0, 0, 100)
            PrintDocument1.DefaultPageSettings.Margins = margins

            If PrintDocument1.PrinterSettings.IsValid Then
                PrintDocument1.DocumentName = TIPO_DOC
                ' Start printing
                PrintDocument1.Print()
            Else
                MessageBox.Show("La impresora no está correctamente configurada o no se encuentra en el sistema.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            End If

            'CURRENT_ESTACION.Close()
            ' EMPRESA.Close()
        End If
        'PrintDocument1.Print()

        Me.Close()

    End Sub


    Function construirTicket(ByVal numTicket As String) As String

        Dim TICKET_STRING As String = ""
        Dim LINEA_SEPARACION As String = StrDup(NUMERO_CARACTERES_X_LINEA_TICKET, "-")

        Dim rstTicket As Recordset = crearRecorset("SELECT * FROM ventas WHERE venta = " & numTicket)

        If rstTicket.EOF Then
            MsgBox("No existe la venta seleccionada", vbInformation)
            Return ""
        End If

        Dim rstPartidas As Recordset = crearRecorset("SELECT partvta.prcantidad, partvta.iespecial, partvta.id_salida, partvta.kit, partvta.articulo, prods.precio1,prods.descrip, partvta.precio, partvta.cantidad, partvta.descuento, partvta.impuesto, partvta.preciobase, partvta.prdescrip, prods.serie FROM partvta INNER JOIN prods ON prods.articulo = partvta.articulo WHERE venta = '" & rstTicket.Fields("venta").Value.ToString & " '")

        ' Traemos todos los datos del cliente
        Dim rstCliente As Recordset = crearRecorset("SELECT * FROM clients WHERE cliente = '" & rstTicket.Fields("Cliente").Value.ToString & "'")

        ' Traemos los datos de cobranza si es que existe
        Dim rstCobranza As Recordset = crearRecorset("SELECT * FROM cobranza WHERE venta = " & rstTicket.Fields("venta").Value.ToString)


        Dim rstVendedor As Recordset = crearRecorset("SELECT * from vends WHERE Vend = '" & rstTicket.Fields("VEND").Value.ToString & "'")

        'Dim primeraCadena As String = "000000000000000000000000000000000000000000000"
        'Dim segundaCadena As String = "The following is very fast and "
        'Dim terceraCadena As String = "useful for plain text as MeasureString calculates the text that can be fitted on an entire page. This is not that useful, however, for formatted text. In that case you would want to have word-level (vs page-level) control, which is more complicated."

        'TICKET_STRING = centrarTexto(primeraCadena) & FINAL_LINEA
        'TICKET_STRING = TICKET_STRING & centrarTexto(segundaCadena) & FINAL_LINEA
        'TICKET_STRING = TICKET_STRING & terceraCadena

        'Dim temp As Integer = CInt(CURRENT_ESTACION.Fields("Cajon").Value.ToString)
        'If temp > 0 Then
        '    TICKET_STRING += Chr(27) & Chr(112) & Chr(48) & Chr(20) & Chr(20)
        '    TICKET_STRING += Chr(7)
        'End If
        
        Dim rstTextoConfTicket As Recordset = crearRecorset("SELECT * FROM tickettext")

        'ESTO ES PARA DARLE LUGAR AL LOGOTIPO
        TICKET_STRING += FINAL_LINEA + FINAL_LINEA + FINAL_LINEA + FINAL_LINEA + FINAL_LINEA + FINAL_LINEA + FINAL_LINEA + FINAL_LINEA

        If rstTextoConfTicket.EOF Then
            If Not EMPRESA.EOF Then
                TICKET_STRING += centrarTexto(UCase(Trim(EMPRESA.Fields("EMPRESA").Value.ToString))) & FINAL_LINEA
                TICKET_STRING += centrarTexto(UCase(Trim(EMPRESA.Fields("Direccion1").Value.ToString))) & FINAL_LINEA
                TICKET_STRING += centrarTexto(UCase(Trim(EMPRESA.Fields("direccion2").Value.ToString))) & FINAL_LINEA
                TICKET_STRING += centrarTexto(UCase(Trim(EMPRESA.Fields("telefonos").Value.ToString))) & FINAL_LINEA
            End If
        Else
            Dim TestArray() As String = Split(rstTextoConfTicket.Fields("textheader").Value.ToString, Chr(13))
            For i As Integer = 0 To TestArray.Length - 1 Step 1


                If i = 2 Then
                    TICKET_STRING += FINAL_LINEA
                End If

                TICKET_STRING += centrarTexto(UCase(Replace(TestArray(i), Chr(10), ""))) & FINAL_LINEA
            Next
        End If


        If Not rstCliente.EOF Then
            Dim cliente As String = Trim("Cliente: " & rstCliente.Fields("NOMBRE").Value.ToString)
            TICKET_STRING += cliente & FINAL_LINEA
        End If

        Dim fecha As String = Trim("Fecha: " & DateValue(rstTicket.Fields("F_EMISION").Value.ToString))
        Dim hora As String = Trim("Hora: " & TimeValue(rstTicket.Fields("USUHORA").Value.ToString))

        'Dim widthTemp As Integer = (NUMERO_CARACTERES_X_LINEA_TICKET - (fecha.Length + hora.Length))

        'If widthTemp > 0 Then
        ' TICKET_STRING += fecha & Space(widthTemp) & hora & FINAL_LINEA
        ' Else
        ' TICKET_STRING += fecha & hora & FINAL_LINEA
        ' End If

        TICKET_STRING += insertarEspaciosEnMedio(fecha, hora) + FINAL_LINEA


        If Not rstVendedor.EOF Then
            Dim vendedor As String = Trim("Vendedor: " & rstVendedor.Fields("Nombre").Value.ToString)
            TICKET_STRING += vendedor & FINAL_LINEA
        End If

        Dim caja As String = Trim("Caja: " & rstTicket.Fields("Caja").Value.ToString)
        Dim ticket As String = Trim("Ticket: " & rstTicket.Fields("NO_REFEREN").Value.ToString)

        TICKET_STRING += insertarEspaciosEnMedio(caja, ticket) + FINAL_LINEA

        TICKET_STRING += LINEA_SEPARACION + FINAL_LINEA

        TICKET_STRING += "  CANT ARTICULO                 PRECIO   TOTAL" & FINAL_LINEA

        TICKET_STRING += LINEA_SEPARACION + FINAL_LINEA

        'Dim nImporteTotal As Double = 0
        Dim nImpuesto As Double = 0
        Dim nCantidadTotal As Double = 0
        Dim nDescuentoTotal As Double = 0
        Dim nImporteOrigen As Double = 0

        While Not rstPartidas.EOF

            'Dim nPrecio As Double = Double.Parse(rstPartidas.Fields("Precio").Value.ToString) * (1 - (Double.Parse(rstPartidas.Fields("Descuento").Value.ToString) / 100)) * (1 + (Double.Parse(rstPartidas.Fields("impuesto").Value.ToString) / 100))
            Dim nPrecio As Double = 0
            Dim nIEspecial As Double = CDbl(rstPartidas.Fields("Precio").Value) * (CDbl(rstPartidas.Fields("iespecial").Value) / 100)
            Dim nIVa As Double = CDbl(rstPartidas.Fields("Precio").Value) * (CDbl(rstPartidas.Fields("impuesto").Value) / 100)
            nPrecio = CDbl(rstPartidas.Fields("Precio").Value) + nIEspecial + nIVa

            Dim nPrecioBase As Double = Math.Round(Double.Parse(rstPartidas.Fields("PrecioBase").Value.ToString) * (1 + (Double.Parse(rstPartidas.Fields("impuesto").Value.ToString) / 100)), 2)
            'Dim nDescuento As Double = (nPrecioBase - nPrecio) * Double.Parse(rstPartidas.Fields("Cantidad").Value.ToString)

            Dim nDescPor As Double
            If nPrecioBase > 0 Then
                nDescPor = (nPrecio / nPrecioBase) * 100
            Else
                nDescPor = 0
            End If

            Dim nDescuento As Double = CDbl(rstPartidas.Fields("Descuento").Value)

            nDescuentoTotal += nDescuento

            Dim cCantidad As String = Format((CDbl(rstPartidas.Fields("cantidad").Value) / CDbl(rstPartidas.Fields("prCantidad").Value)), "#0.000")
            Dim cDescrip As String = rstPartidas.Fields("Descrip").Value.ToString
            Dim cPrecio As String = Format((nPrecio * (1 - (CDbl(rstPartidas.Fields("descuento").Value) / 100))), "###0.00")
            Dim nImporte As Double = nPrecio * CDbl(rstPartidas.Fields("Cantidad").Value)
            Dim cImporte As String = Format(nImporte, "###0.00")

            nImpuesto += (nImporte * (CDbl(rstPartidas.Fields("impuesto").Value) / 100))
            TICKET_STRING += crearlineaDeArticulos(cCantidad, cDescrip, cPrecio, cImporte) + FINAL_LINEA

            'TOTALES
            'nImporteTotal += nImporte
            nCantidadTotal += CDbl(rstPartidas.Fields("cantidad").Value) / CDbl(rstPartidas.Fields("prcantidad").Value)



            If CInt(rstPartidas.Fields("Kit").Value) <> 0 Then
                Dim rstOpciones As Recordset = crearRecorset("SELECT partvtaopciones.articulo, prods.descrip FROM partvtaopciones INNER JOIN prods ON partvtaopciones.articulo = prods.articulo WHERE id_salida = '" & rstPartidas.Fields("id_salida").Value.ToString & "'")

                While Not rstOpciones.EOF
                    TICKET_STRING += "Opcion:" & rstOpciones.Fields("descrip").Value.ToString & FINAL_LINEA
                    rstOpciones.MoveNext()
                End While
            End If

            If CInt(rstPartidas.Fields("serie").Value) <> 0 Then

                Dim rstSeries As Recordset

                If CInt(rstTicket.Fields("ticket").Value) <> 0 Then
                    rstSeries = crearRecorset("SELECT * FROM series WHERE documento = 'TICKET' AND numeroDocumento = '" & rstTicket.Fields("no_referen").Value.ToString & "' AND articulo = '" & rstPartidas.Fields("articulo").Value.ToString & "'")
                Else
                    rstSeries = crearRecorset("SELECT * FROM series WHERE documento = 'REMISION' AND numeroDocumento = '" & rstTicket.Fields("no_referen").Value.ToString & "' AND articulo = '" & rstPartidas.Fields("articulo").Value.ToString & "'")
                End If

                While Not rstSeries.EOF

                    Dim Serie As String = Trim(rstSeries.Fields("serie").Value.ToString)
                    'Serie = PadL(Serie, 11)
                    Serie = Mid(Serie, Len(Serie) - 10)

                    TICKET_STRING += "SERIE: " & Serie & FINAL_LINEA
                    rstSeries.MoveNext()
                End While

            End If

            rstPartidas.MoveNext()
        End While


        TICKET_STRING += LINEA_SEPARACION + FINAL_LINEA

        'TICKET_STRING += alinearALaDerecha("TOTAL:           " & Format(nImporteTotal, "###0.00")) & FINAL_LINEA
        Dim importeTotal As Double = CDbl(rstTicket.Fields("IMPORTE").Value)
        TICKET_STRING += alinearALaDerecha("TOTAL:           " & Format(importeTotal, "###0.00")) & FINAL_LINEA

        Dim impuesto As Double = CDbl(rstTicket.Fields("IMPUESTO").Value)
        If impuesto > 0 Then
            TICKET_STRING += alinearALaDerecha("+IMPUESTO:           " & Format(impuesto, "###0.00")) & FINAL_LINEA
        End If

        Dim impuestoEspecial As Double = CDbl(rstTicket.Fields("iespecial").Value)
        If impuestoEspecial > 0 Then
            TICKET_STRING += alinearALaDerecha("+IMPUESTO Especial:       " & Format(impuestoEspecial, "###0.00")) & FINAL_LINEA
        End If

        Dim redondeo As Double = CDbl(rstTicket.Fields("redondeo").Value)
        If redondeo > 0 Then
            TICKET_STRING += alinearALaDerecha("+REDONDEO:       " & Format(redondeo, "###0.00")) & FINAL_LINEA
        End If

        TICKET_STRING += FINAL_LINEA

        Dim nPagoTotal As Double = CDbl(rstTicket.Fields("Pago1").Value) + CDbl(rstTicket.Fields("Pago2").Value) + CDbl(rstTicket.Fields("Pago3").Value)

        If CDbl(rstTicket.Fields("Pago1").Value) > 0 Then
            TICKET_STRING += insertarEspaciosEnMedio("PAGO EN " & rstTicket.Fields("Concepto1").Value.ToString & ":", Format(CDbl(rstTicket.Fields("Pago1").Value), "$###0.00")) + FINAL_LINEA
        End If

        If CDbl(rstTicket.Fields("Pago2").Value) > 0 Then
            TICKET_STRING += insertarEspaciosEnMedio("PAGO EN " & rstTicket.Fields("Concepto2").Value.ToString & ":", Format(CDbl(rstTicket.Fields("Pago2").Value), "$###0.00")) & FINAL_LINEA
        End If

        If CDbl(rstTicket.Fields("Pago3").Value) > 0 Then
            TICKET_STRING += insertarEspaciosEnMedio("PAGO EN " & rstTicket.Fields("Concepto3").Value.ToString & ":", Format(CDbl(rstTicket.Fields("Pago3").Value), "$###0.00")) & FINAL_LINEA
        End If


        Dim nComision As Double = CDbl(rstTicket.Fields("Comision").Value)
        If nComision > 0 Then
            TICKET_STRING += "Comisión: " & Format(CDbl(rstTicket.Fields("comision").Value), "$###0.00") & FINAL_LINEA
        End If

        If nPagoTotal < (impuesto + importeTotal + impuestoEspecial + redondeo) Then
            TICKET_STRING += "Credito: " & Format(CDbl(rstTicket.Fields("impuesto").Value) + CDbl(rstTicket.Fields("iespecial").Value) + CDbl(rstTicket.Fields("importe").Value) + CDbl(rstTicket.Fields("redondeo").Value) - nPagoTotal, "$###0.00") & FINAL_LINEA
        Else
            TICKET_STRING += insertarEspaciosEnMedio("CAMBIO:", Format(nPagoTotal - nComision - CDbl(rstTicket.Fields("impuesto").Value) - CDbl(rstTicket.Fields("iespecial").Value) - CDbl(rstTicket.Fields("importe").Value) - CDbl(rstTicket.Fields("redondeo").Value), "$###0.00")) & FINAL_LINEA
        End If


        'If Not rstCliente.EOF Then
        '    TICKET_STRING += FINAL_LINEA

        '    If CDbl(rstCliente.Fields("saldo").Value) > 0 Then
        '        TICKET_STRING += "Su saldo: " & Format(CDbl(rstCliente.Fields("saldo").Value), "$##,##0.00") & FINAL_LINEA
        '    End If
        'End If

        If rstTextoConfTicket.EOF Then
            TICKET_STRING += FINAL_LINEA + FINAL_LINEA + FINAL_LINEA
            TICKET_STRING += centrarTexto("***** GRACIAS POR SU COMPRA *****") + FINAL_LINEA + FINAL_LINEA
        Else
            TICKET_STRING += centrarTexto(rstTextoConfTicket.Fields("textend").Value.ToString) + FINAL_LINEA + FINAL_LINEA
        End If

        ' Esto manda un corte de papel

        If CDbl(CURRENT_ESTACION.Fields("ticketcorte").Value) > 0 Then
            TICKET_STRING += Chr(27) & Chr(105)
        End If


        'rstTicket.Close()
        'rstCliente.Close()
        'rstCobranza.Close()
        'rstPartidas.Close()
        'rstTextoConfTicket.Close()
        'rstTicket.Close()
        'rstVendedor.Close()

        Return TICKET_STRING
    End Function

    Function crearlineaDeArticulos(ByRef cant As String, ByRef articulo As String, ByRef precio As String, total As String) As String
        Dim articulos As String = ""

        Dim espacioParaDescripcion As Integer = NUMERO_CARACTERES_X_LINEA_TICKET - 23

        Dim finalCant As String = Trim(cant)

        If finalCant.Length > 6 Then
            finalCant = Mid(finalCant, 1, 6)
        Else
            finalCant = Space(6 - (finalCant.Length)) + finalCant
        End If


        articulos += finalCant & " "

        Dim finalDescrip As String = Trim(articulo)
        If finalDescrip.Length > espacioParaDescripcion Then
            finalDescrip = Mid(finalDescrip, 1, espacioParaDescripcion)
        Else
            finalDescrip = finalDescrip & Space(espacioParaDescripcion - (finalDescrip.Length))
        End If

        articulos += finalDescrip & " "

        Dim finalPrecio As String = Trim(precio)

        If finalPrecio.Length > 7 Then
            finalPrecio = Mid(finalPrecio, 1, 7)
        Else
            finalPrecio = Space(7 - (finalPrecio.Length)) + finalPrecio
        End If

        articulos += finalPrecio & " "

        Dim finalImporte As String = Trim(total)

        If finalImporte.Length > 7 Then
            finalImporte = Mid(finalImporte, 1, 7)
        Else
            finalImporte = Space(7 - (finalImporte.Length)) + finalImporte
        End If

        articulos += finalImporte

        Return articulos
    End Function


    Function insertarEspaciosEnMedio(ByRef cadena1 As String, ByRef cadena2 As String, Optional ByVal prioridad As Integer = 1) As String
        Dim stringFinal As String = ""

        cadena1 = Trim(cadena1)
        cadena2 = Trim(cadena2)

        If ((cadena1.Length + cadena2.Length) > NUMERO_CARACTERES_X_LINEA_TICKET) Then

            If prioridad = 1 Then
                cadena2 = Mid(cadena2, 1, CInt((NUMERO_CARACTERES_X_LINEA_TICKET / 2) - 3))
                stringFinal = (Mid(cadena1, 1, CInt(NUMERO_CARACTERES_X_LINEA_TICKET / 2))) + "   " + cadena2
            Else
                cadena1 = Mid(cadena1, 1, CInt((NUMERO_CARACTERES_X_LINEA_TICKET / 2) - 3))
                stringFinal = cadena1 + "   " + (Mid(cadena2, 1, CInt(NUMERO_CARACTERES_X_LINEA_TICKET / 2)))
            End If
        Else

            Dim widthTemp As Integer = (NUMERO_CARACTERES_X_LINEA_TICKET - (cadena1.Length + cadena2.Length))

            If widthTemp > 0 Then
                stringFinal = cadena1 & Space(widthTemp) & cadena2
            Else
                stringFinal = cadena1 & cadena2
            End If

        End If

        Return stringFinal
    End Function

    Function alinearALaDerecha(ByRef texto As String) As String
        Dim textoAlineado As String

        textoAlineado = Trim(texto)

        If textoAlineado.Length > NUMERO_CARACTERES_X_LINEA_TICKET Then
            textoAlineado = Mid(textoAlineado, NUMERO_CARACTERES_X_LINEA_TICKET)
        Else
            textoAlineado = Space(NUMERO_CARACTERES_X_LINEA_TICKET - textoAlineado.Length) & textoAlineado
        End If

        Return textoAlineado
    End Function

    Function centrarTexto(ByRef lineaString As String) As String

        Dim tempStringSpacios As String = ""

        If lineaString.Length = NUMERO_CARACTERES_X_LINEA_TICKET Then
            tempStringSpacios = lineaString
        ElseIf Trim(lineaString).Length < NUMERO_CARACTERES_X_LINEA_TICKET Then

            Dim numeroDeSpacios As Double = Math.Truncate((NUMERO_CARACTERES_X_LINEA_TICKET - Trim(lineaString).Length) / 2)
            tempStringSpacios = Space(CInt(numeroDeSpacios))
            'tempStringSpacios = StrDup(CInt(numeroDeSpacios), "=")
            tempStringSpacios = tempStringSpacios & Trim(lineaString)

        End If

        Return tempStringSpacios
    End Function


    Function crearRecorset(ByVal SQLConsulta As String) As Recordset

        Dim recorset As Recordset = New Recordset

        Try
            recorset.Open(SQLConsulta, DB_CONN, CursorTypeEnum.adOpenForwardOnly, LockTypeEnum.adLockReadOnly)
        Catch SQLexc As SqlException
            MsgBox("Hubo un error al crear el recordSet" & SQLexc.Message)
        End Try
        Return recorset
    End Function


    Friend WithEvents pbImage As System.Windows.Forms.PictureBox
    Private Sub pdoc_PrintPage(ByVal sender As Object, ByVal e As System.Drawing.Printing.PrintPageEventArgs) Handles PrintDocument1.PrintPage
        ' Declare a variable to hold the position of the last printed char. Declare
        ' as static so that subsequent PrintPage events can reference it.
        Static intCurrentChar As Int32
        ' Initialize the font to be used for printing.
        'Dim font As New Font(NOMBRE_FUENTE, TAMANIO_FUENTE)
        Dim font As New Font(NOMBRE_FUENTE, TAMANIO_FUENTE, FontStyle.Regular, GraphicsUnit.Point)
        'Dim font As New Font("Microsoft Sans Serif", 10, FontStyle.Regular, GraphicsUnit.Point)

        Dim intPrintAreaHeight, intPrintAreaWidth, marginLeft, marginTop As Int32

        With PrintDocument1.DefaultPageSettings
            ' Initialize local variables that contain the bounds of the printing 
            ' area rectangle.
            intPrintAreaHeight = .PaperSize.Height - .Margins.Top - .Margins.Bottom
            intPrintAreaWidth = .PaperSize.Width - .Margins.Left - .Margins.Right

            ' Initialize local variables to hold margin values that will serve
            ' as the X and Y coordinates for the upper left corner of the printing 
            ' area rectangle.
            marginLeft = .Margins.Left ' X coordinate
            marginTop = .Margins.Top ' Y coordinate
        End With

        ' If the user selected Landscape mode, swap the printing area height 
        ' and width.
        If PrintDocument1.DefaultPageSettings.Landscape Then
            Dim intTemp As Int32
            intTemp = intPrintAreaHeight
            intPrintAreaHeight = intPrintAreaWidth
            intPrintAreaWidth = intTemp
        End If

        ' Calculate the total number of lines in the document based on the height of
        ' the printing area and the height of the font.
        Dim intLineCount As Int32 = CInt(intPrintAreaHeight / font.Height)
        ' Initialize the rectangle structure that defines the printing area.
        Dim rectPrintingArea As New RectangleF(marginLeft, marginTop, intPrintAreaWidth, intPrintAreaHeight)

        ' Instantiate the StringFormat class, which encapsulates text layout 
        ' information (such as alignment and line spacing), display manipulations 
        ' (such as ellipsis insertion and national digit substitution) and OpenType 
        ' features. Use of StringFormat causes MeasureString and DrawString to use
        ' only an integer number of lines when printing each page, ignoring partial
        ' lines that would otherwise likely be printed if the number of lines per 
        ' page do not divide up cleanly for each page (which is usually the case).
        ' See further discussion in the SDK documentation about StringFormatFlags.
        Dim fmt As New StringFormat(StringFormatFlags.LineLimit)

        ' Call MeasureString to determine the number of characters that will fit in
        ' the printing area rectangle. The CharFitted Int32 is passed ByRef and used
        ' later when calculating intCurrentChar and thus HasMorePages. LinesFilled 
        ' is not needed for this sample but must be passed when passing CharsFitted.
        ' Mid is used to pass the segment of remaining text left off from the 
        ' previous page of printing (recall that intCurrentChar was declared as 
        ' static.

        'e.Graphics.PageUnit = GraphicsUnit.Point

        ' Draw the bitmap
        pbImage = New System.Windows.Forms.PictureBox
        pbImage.Image = CType(My.Resources.logoAdame, System.Drawing.Image)
        Dim x, y As Integer
        x = 40
        y = 0
        e.Graphics.DrawImage(pbImage.Image, x, y, pbImage.Image.Width, pbImage.Image.Height )

        Dim intLinesFilled, intCharsFitted As Int32
        e.Graphics.MeasureString(Mid(CURRENT_TEXT_TO_PRINT, intCurrentChar + 1), font, New SizeF(intPrintAreaWidth, intPrintAreaHeight), fmt, intCharsFitted, intLinesFilled)

        ' Print the text to the page.
        e.Graphics.DrawString(Mid(CURRENT_TEXT_TO_PRINT, intCurrentChar + 1), font, Brushes.Black, rectPrintingArea, fmt)

        ' Advance the current char to the last char printed on this page. As 
        ' intCurrentChar is a static variable, its value can be used for the next
        ' page to be printed. It is advanced by 1 and passed to Mid() to print the
        ' next page (see above in MeasureString()).
        intCurrentChar += intCharsFitted

        ' HasMorePages tells the printing module whether another PrintPage event
        ' should be fired.
        If intCurrentChar < CURRENT_TEXT_TO_PRINT.Length Then
            e.HasMorePages = True
        Else
            e.HasMorePages = False
            ' You must explicitly reset intCurrentChar as it is static.
            intCurrentChar = 0
        End If
    End Sub
End Class

'ESTO SOLO ES PARA RECORRER EL OBJETO PRINTER Y SABER CUALES IMPRESORAS ESTAN LISTAS EN LA PC
'Dim i As Integer
'Dim pkInstalledPrinters As String

'        For i = 0 To PrinterSettings.InstalledPrinters.Count - 1
'            pkInstalledPrinters = PrinterSettings.InstalledPrinters.Item(i)
'            Console.WriteLine(pkInstalledPrinters)
'        Next




'Sub Main()
'    ' Es una funcion que puede ejecutar archivos junto
'    ' con su aplicacion asociada     
'    ShellRun(Parent.hWnd, "Open", "c:\daniel\adams.mp3")
'    ' ShellCommand puede ejecutar programas     
'    ' 0.- Oculto     ' 1.- Normal     ' 2.- Minimizado     ' 3.- Maximizado     
'    ShellCommand("notepad.exe", 1)
'    ' Ejecuta comandos de DOS y manda la salida a un archivo     
'CommandToFile "xcopy c:\clientes.txt clientes2.txt", "c:\salida.txt"    End Sub


'Sub Main()        
'ShellCommand Ambiente.Path & "\cassini\CassiniWebServer.Exe " & Chr(32) & Ambiente.Path & "\cassini\" & Chr(32) & " 4410 /", 1                 End Sub  


'    ShellCommand(Ambiente.Path & "\Web\CassiniWebServer.Exe " & Chr(34) & Ambiente.Path & "\web" & Chr(34) & " 4410 /", 0)


' Script.RunForm "RASTREOSERIE", Me, Ambiente,,

'    Script.RunProcess("BARTICULOS001", Parent, Ambiente)