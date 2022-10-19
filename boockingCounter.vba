Sub bookingCounter()


	'This is a second version of the first function witch counts a number of bookings cancelled 
	'by hearing the character '*' witch indicates the booking on the same date was cancelled 

    'Declaracion de Array
    Dim arrayValues As Variant
    'Declaracion de cadenas de texto
    Dim ArrayCommnents(), myRangeS As String
    'Declaración de rangos
    Dim myRange, commentsRange, mycell As Range
    ' Un contador
    Dim i As Integer
    
    
    
    'String que se asigna Apuntando a la direccion seleccionada
    'ESTO FUNCIONA BIEN
    myRangeS = Selection.Address
    
    
    'Tipo de dato variant Que da el valor de rango de myRangeS
    'arrayValues = Range(myRangeS).Value
    
    
    
    
    
    'Rango que apunta a myRangeS de la hoja activa
    'Set myRange = ActiveSheet.Range(myRangeS)
    'Para Dos o mas celdas
    'Set myRange = Cells.Range(myRangeS)
    'Para Una celda
    Set myRange = Range(myRangeS)
    
    
    
    
    
    
    
    'Rango de comentarios es un objeto de tipo rango que coincide con tipo y valor especificados
    'Set commentsRange = myRange.Cells.SpecialCells(xlCellTypeComments)
    Set commentsRange = myRange.Cells
    
    'Debug.Print (commentsRange)
    
    'Recoge la longitud del array
    arrayLenght = commentsRange.count
    
    'Reasigna el valor del String arraycoments
    ReDim arrayComments(arrayLenght)
    
    'MsgBox "el valor de I es " & i
    
    'Dim megaCadena As String
    
    'Declaración de contadores de mes
    Dim Enero As Integer
    Dim Febrero As Integer
    Dim Marzo As Integer
    Dim Abril As Integer
    Dim Mayo As Integer
    Dim Junio As Integer
    Dim Julio As Integer
    Dim Agosto As Integer
    Dim Septiembre As Integer
    Dim Octubre As Integer
    Dim Noviembre As Integer
    Dim Diciembre As Integer
    
    
    'Declaración de contador de cancelaciones
    Dim cEnero As Integer
    Dim cFebrero As Integer
    Dim cMarzo As Integer
    Dim cAbril As Integer
    Dim cMayo As Integer
    Dim cJunio As Integer
    Dim cJulio As Integer
    Dim cAgosto As Integer
    Dim cSeptiembre As Integer
    Dim cOctubre As Integer
    Dim cNoviembre As Integer
    Dim cDiciembre As Integer
    
    'Declaración de Booleano
    Dim interruptor As Boolean
    
    'Declaro un Booleano para valorar si es reserva o cancelación
    Dim cancelado As Boolean
    
    'Declaramos un Integer como variable de estado que solo podra tener 3 valores 0/1/2
    Dim estado As Integer
    estado = 0
    
    'Declaración de incrementer para recorrer la cadena
    Dim contador As Integer
    
    'Declaración de un Integer para evaluar el mes
    Dim evaluador As String
    'Declaro un almacen
    Dim almacen As String
    
    'MsgBox "Valor del Booleano se inicia en... " & cancelado
    ' Bucle que recorre casillas imprimiendo los valores de fechas
    For Each mycell In commentsRange
    
        arrayComments(i) = mycell.Comment.Text
        almacen = arrayComments(i)
        'Debug.Print (almacen)
        
        For contador = 1 To Len(almacen)
            'Debug.Print Mid(almacen, contador, 1)
            If Mid(almacen, contador, 1) = "*" Then
                cancelado = Not (cancelado)
            End If
            If Mid(almacen, contador, 1) = "/" Then
                interruptor = Not (interruptor)
                estado = estado + 1
            End If
            
            If cancelado = True And interruptor = True Then
                evaluador = evaluador & Mid(almacen, contador, 1)
            End If
            If cancelado = False And interruptor = True Then
                evaluador = evaluador & Mid(almacen, contador, 1)
            End If
            
            'Debug.Print (evaluador)
            If cancelado = True And interruptor = True Then
                If evaluador = "/01" Then cEnero = cEnero + 1
                If evaluador = "/02" Then cFebrero = cFebrero + 1
                If evaluador = "/03" Then cMarzo = cMarzo + 1
                If evaluador = "/04" Then cAbril = cAbril + 1
                If evaluador = "/05" Then cMayo = cMayo + 1
                If evaluador = "/06" Then cJunio = cJunio + 1
                If evaluador = "/07" Then cJulio = cJulio + 1
                If evaluador = "/08" Then cAgosto = cAgosto + 1
                If evaluador = "/09" Then cSeptiembre = cSeptiembre + 1
                If evaluador = "/10" Then cOctubre = cOctubre + 1
                If evaluador = "/11" Then cNoviembre = cNoviembre + 1
                If evaluador = "/12" Then cDiciembre = cDiciembre + 1
            ElseIf cancelado = True And estado = 2 Then
                evaluador = ""
                cancelado = Not (cancelado)
                estado = 0
            End If
            If cancelado = False And interruptor = False Then
                If evaluador = "/01" Then Enero = Enero + 1
                If evaluador = "/02" Then Febrero = Febrero + 1
                If evaluador = "/03" Then Marzo = Marzo + 1
                If evaluador = "/04" Then Abril = Abril + 1
                If evaluador = "/05" Then Mayo = Mayo + 1
                If evaluador = "/06" Then Junio = Junio + 1
                If evaluador = "/07" Then Julio = Julio + 1
                If evaluador = "/08" Then Agosto = Agosto + 1
                If evaluador = "/09" Then Septiembre = Septiembre + 1
                If evaluador = "/10" Then Octubre = Octubre + 1
                If evaluador = "/11" Then Noviembre = Noviembre + 1
                If evaluador = "/12" Then Diciembre = Diciembre + 1
                evaluador = ""
            End If
            If estado = 2 Then estado = 0
            'Debug.Print "estado" & (estado)
            
            
            
        Next
        
        i = i + 1
        
        
    Next mycell
    
    MsgBox " RESERVAS " & vbNewLine & "Enero:" & " " & Enero & vbNewLine & "Febrero:" & " " & Febrero & vbNewLine & "Marzo:" & " " & Marzo & vbNewLine & "Abril:" & " " & Abril & vbNewLine & "Mayo:" & " " & Mayo & vbNewLine & "Junio:" & " " & Junio & vbNewLine & "Julio:" & " " & Julio & vbNewLine & "Agosto:" & " " & Agosto & vbNewLine & "Septiembre:" & " " & Septiembre & vbNewLine & "Octubre:" & " " & Octubre & vbNewLine & "Noviembre:" & " " & Noviembre & vbNewLine & "Diciembre:" & " " & Diciembre
    
    MsgBox " CANCELACIONES " & vbNewLine & "Enero:" & " " & cEnero & vbNewLine & "Febrero:" & " " & cFebrero & vbNewLine & "Marzo:" & " " & cMarzo & vbNewLine & "Abril:" & " " & cAbril & vbNewLine & "Mayo:" & " " & cMayo & vbNewLine & "Junio:" & " " & cJunio & vbNewLine & "Julio:" & " " & cJulio & vbNewLine & "Agosto:" & " " & cAgosto & vbNewLine & "Septiembre:" & " " & cSeptiembre & vbNewLine & "Octubre:" & " " & cOctubre & vbNewLine & "Noviembre:" & " " & cNoviembre & vbNewLine & "Diciembre:" & " " & cDiciembre

    
End Sub
