Function getComment(xCell As Range) As String
'UpdatebyExtendoffice20180330
On Error Resume Next
getComment = xCell.Comment.Text
End Function


Public Sub test()
    Dim myString As String
    Dim count As Integer
    Dim vector() As Variant
    myString = Selection.Address
     'vector(count) = getComment(Selection)
     myString = Range(Selection).Comment()
    'myString = Selection.Range(myString).Comments
    
End Sub




Sub commentToArray()


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
    'Para Una celda y Selección de diferentes celdas
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
    
    'Declaración de Booleano
    Dim interruptor As Boolean  
    
    'Declaración de incrementer para recorrer la cadena
    Dim contador As Integer
    
    'Declaración de un Integer para evaluar el mes
    Dim evaluador As String
    'Declaro un almacen
    Dim almacen As String
    
    MsgBox "Valor del Booleano se inicia en... " & interruptor
    ' Bucle que recorre casillas imprimiendo los valores de fechas
    For Each mycell In commentsRange
    
        arrayComments(i) = mycell.Comment.Text
        almacen = arrayComments(i)
        'Debug.Print (arrayComments(i))
        
        For contador = 1 To Len(almacen)
            'Debug.Print Mid(Almacen, contador, 1)
          
            'MsgBox Mid(almacen, contador, 1)
            If Mid(almacen, contador, 1) = "/" Then
            interruptor = Not (interruptor)
            End If
            'MsgBox interruptor
            If interruptor = True Then
            evaluador = evaluador & Mid(almacen, contador, 1)
            Debug.Print (evaluador)
            Else
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
            
            
            
            
        Next
        
        ' megaCadena = megaCadena & arrayComments(i)
        i = i + 1
        
        
    Next mycell
    
    MsgBox "Enero:" & " " & Enero & vbNewLine & "Febrero:" & " " & Febrero & vbNewLine & "Marzo:" & " " & Marzo & vbNewLine & "Abril:" & " " & Abril & vbNewLine & "Mayo:" & " " & Mayo & vbNewLine & "Junio:" & " " & Junio & vbNewLine & "Julio:" & " " & Julio & vbNewLine & "Agosto:" & " " & Agosto & vbNewLine & "Septiembre:" & " " & Septiembre & vbNewLine & "Octubre:" & " " & Octubre & vbNewLine & "Noviembre:" & " " & Noviembre & vbNewLine & "Diciembre:" & " " & Diciembre
 
End Sub