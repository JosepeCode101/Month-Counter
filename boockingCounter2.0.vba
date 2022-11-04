Sub bookingCounter()

    Dim arrayValues As Variant
    
    Dim ArrayComments(), myRangeS As String
    
    Dim myRange, commentsRange, mycell As Range
    
    Dim i As Integer
    
    
    
    myRangeS = Selection.Address
    
    Set myRange = Range(myRangeS)
    
    Set commentsRange = myRange.Cells
    
    arrayLenght = commentsRange.Count
    
    ReDim ArrayComments(arrayLenght)
    
    
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
    
    
    
    Dim interruptor As Boolean
    Dim cancelado As Boolean
    
    Dim estado As Integer
    estado = 0
    
    Dim contador As Integer
    Dim evaluador As String
    
    Dim almacen As String
    
    'Declaro almacen para fechas
    Dim almacenFecha As String
    'Declaro numero de semana Reservados
    Dim s1, s2, s3, s4, s5, s6, s7, s8, s9, s10, s11, s12, s13, s14, s15, s16, s17, s18, s19, s20, s21, s22, s23, s25, s26, s27, s28, s29, s30, s31, s32, s33, s34, s35, s36, s37, s38, s39, s40, s41, s42, s43, s44, s45, s46, s47, s48, s49, s50, s51, s52, s53 As Integer
    'Declaro numero de semana Cancelado
    Dim cs1, cs2, cs3, cs4, cs5, cs6, cs7, cs8, cs9, cs10, cs11, cs12, cs13, cs14, cs15, cs16, cs17, cs18, cs19, cs20, cs21, cs22, cs23, cs25, cs26, cs27, cs28, cs29, cs30, cs31, cs32, cs33, cs34, cs35, cs36, cs37, cs38, cs39, cs40, cs41, cs42, cs43, cs44, cs45, cs46, cs47, cs48, cs49, cs50, cs51, cs52, cs53 As Integer
    'Declaro un evauador de nuemero de semana
    Dim weekEvaluator As Integer
    
    
    
    On Error Resume Next
    
    For Each mycell In commentsRange
        
        ArrayComments(i) = mycell.Comment.Text
        almacen = ArrayComments(i)
        

    
        'Debug.Print (almacen)
        
        For contador = 1 To Len(almacen)
            
            'Debug.Print Mid(almacen, contador, 1)
            
            If IsNumeric(Mid(almacen, contador, 1)) Or Mid(almacen, contador, 1) = "/" Then
                almacenFecha = almacenFecha & Mid(almacen, contador, 1)
            ElseIf Len(almacenFecha) = 10 Then
            Debug.Print almacenFecha
            Debug.Print WorksheetFunction.WeekNum(almacenFecha)
            End If
            
            'Debug.Print (almacenFecha)
            'Debug.Print Len(almacenFecha)
            'Evaluando Cancelaciones y Reservas
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
            
            'Contando cancelaciones y reservas por numero de semana
            If cancelado = False And interruptor = False And Len(almacenFecha) = 10 Then
                Debug.Print (almacenFecha)
                Debug.Print WorksheetFunction.WeekNum(almacenFecha)
                weekEvaluator = WorksheetFunction.WeekNum(almacenFecha)
                If weekEvaluator = 1 Then s1 = s1 + 1
                If weekEvaluator = 2 Then s2 = s2 + 1
                If weekEvaluator = 3 Then s3 = s3 + 1
                If weekEvaluator = 4 Then s4 = s4 + 1
                If weekEvaluator = 5 Then s5 = s5 + 1
                If weekEvaluator = 6 Then s6 = s6 + 1
                If weekEvaluator = 7 Then s7 = s7 + 1
                If weekEvaluator = 8 Then s8 = s8 + 1
                If weekEvaluator = 9 Then s9 = s9 + 1
                If weekEvaluator = 10 Then s10 = s10 + 1
                If weekEvaluator = 11 Then s11 = s11 + 1
                If weekEvaluator = 12 Then s12 = s12 + 1
                If weekEvaluator = 13 Then s13 = s13 + 1
                If weekEvaluator = 14 Then s14 = s14 + 1
                If weekEvaluator = 15 Then s15 = s15 + 1
                If weekEvaluator = 16 Then s16 = s16 + 1
                If weekEvaluator = 17 Then s17 = s17 + 1
                If weekEvaluator = 18 Then s18 = s18 + 1
                If weekEvaluator = 19 Then s19 = s19 + 1
                If weekEvaluator = 20 Then s20 = s20 + 1
                If weekEvaluator = 21 Then s21 = s21 + 1
                If weekEvaluator = 22 Then s22 = s22 + 1
                If weekEvaluator = 23 Then s23 = s23 + 1
                If weekEvaluator = 24 Then s24 = s24 + 1
                If weekEvaluator = 25 Then s25 = s25 + 1
                If weekEvaluator = 26 Then s26 = s26 + 1
                If weekEvaluator = 27 Then s27 = s27 + 1
                If weekEvaluator = 28 Then s28 = s28 + 1
                If weekEvaluator = 29 Then s29 = s29 + 1
                If weekEvaluator = 30 Then s30 = s30 + 1
                If weekEvaluator = 31 Then s31 = s31 + 1
                If weekEvaluator = 32 Then s32 = s32 + 1
                If weekEvaluator = 33 Then s33 = s33 + 1
                If weekEvaluator = 34 Then s34 = s34 + 1
                If weekEvaluator = 35 Then s35 = s35 + 1
                If weekEvaluator = 36 Then s36 = s36 + 1
                If weekEvaluator = 37 Then s37 = s37 + 1
                If weekEvaluator = 38 Then s38 = s38 + 1
                If weekEvaluator = 39 Then s39 = s39 + 1
                If weekEvaluator = 40 Then s40 = s40 + 1
                If weekEvaluator = 41 Then s41 = s41 + 1
                If weekEvaluator = 42 Then s42 = s42 + 1
                If weekEvaluator = 43 Then s43 = s43 + 1
                If weekEvaluator = 44 Then s44 = s44 + 1
                If weekEvaluator = 45 Then s45 = s45 + 1
                If weekEvaluator = 46 Then s46 = s46 + 1
                If weekEvaluator = 47 Then s47 = s47 + 1
                If weekEvaluator = 48 Then s48 = s48 + 1
                If weekEvaluator = 49 Then s49 = s49 + 1
                If weekEvaluator = 50 Then s50 = s50 + 1
                If weekEvaluator = 51 Then s51 = s51 + 1
                If weekEvaluator = 52 Then s52 = s52 + 1
                If weekEvaluator = 53 Then s53 = s53 + 1
            End If
            If cancelado = True And interruptor = True And Len(almacenFecha) = 10 Then
                weekEvaluator = WorksheetFunction.WeekNum(almacenFecha)
                If weekEvaluator = 1 Then cs1 = cs1 + 1
                If weekEvaluator = 2 Then cs2 = cs2 + 1
                If weekEvaluator = 3 Then cs3 = cs3 + 1
                If weekEvaluator = 4 Then cs4 = cs4 + 1
                If weekEvaluator = 5 Then cs5 = cs5 + 1
                If weekEvaluator = 6 Then cs6 = cs6 + 1
                If weekEvaluator = 7 Then cs7 = cs7 + 1
                If weekEvaluator = 8 Then cs8 = cs8 + 1
                If weekEvaluator = 9 Then cs9 = cs9 + 1
                If weekEvaluator = 10 Then cs10 = cs10 + 1
                If weekEvaluator = 11 Then cs11 = cs11 + 1
                If weekEvaluator = 12 Then cs12 = cs12 + 1
                If weekEvaluator = 13 Then cs13 = cs13 + 1
                If weekEvaluator = 14 Then cs14 = cs14 + 1
                If weekEvaluator = 15 Then cs15 = cs15 + 1
                If weekEvaluator = 16 Then cs16 = cs16 + 1
                If weekEvaluator = 17 Then cs17 = cs17 + 1
                If weekEvaluator = 18 Then cs18 = cs18 + 1
                If weekEvaluator = 19 Then cs19 = cs19 + 1
                If weekEvaluator = 20 Then cs20 = cs20 + 1
                If weekEvaluator = 21 Then cs21 = cs21 + 1
                If weekEvaluator = 22 Then cs22 = cs22 + 1
                If weekEvaluator = 23 Then cs23 = cs23 + 1
                If weekEvaluator = 24 Then cs24 = cs24 + 1
                If weekEvaluator = 25 Then cs25 = cs25 + 1
                If weekEvaluator = 26 Then cs26 = cs26 + 1
                If weekEvaluator = 27 Then cs27 = cs27 + 1
                If weekEvaluator = 28 Then cs28 = cs28 + 1
                If weekEvaluator = 29 Then cs29 = cs29 + 1
                If weekEvaluator = 30 Then cs30 = cs30 + 1
                If weekEvaluator = 31 Then cs31 = cs31 + 1
                If weekEvaluator = 32 Then cs32 = cs32 + 1
                If weekEvaluator = 33 Then cs33 = cs33 + 1
                If weekEvaluator = 34 Then cs34 = cs34 + 1
                If weekEvaluator = 35 Then cs35 = cs35 + 1
                If weekEvaluator = 36 Then cs36 = cs36 + 1
                If weekEvaluator = 37 Then cs37 = cs37 + 1
                If weekEvaluator = 38 Then cs38 = cs38 + 1
                If weekEvaluator = 39 Then cs39 = cs39 + 1
                If weekEvaluator = 40 Then cs40 = cs40 + 1
                If weekEvaluator = 41 Then cs41 = cs41 + 1
                If weekEvaluator = 42 Then cs42 = cs42 + 1
                If weekEvaluator = 43 Then cs43 = cs43 + 1
                If weekEvaluator = 44 Then cs44 = cs44 + 1
                If weekEvaluator = 45 Then cs45 = cs45 + 1
                If weekEvaluator = 46 Then cs46 = cs46 + 1
                If weekEvaluator = 47 Then cs47 = cs47 + 1
                If weekEvaluator = 48 Then cs48 = cs48 + 1
                If weekEvaluator = 49 Then cs49 = cs49 + 1
                If weekEvaluator = 50 Then cs50 = cs50 + 1
                If weekEvaluator = 51 Then cs51 = cs51 + 1
                If weekEvaluator = 52 Then cs52 = cs52 + 1
                If weekEvaluator = 53 Then cs53 = cs53 + 1
            End If
            
            
            
            'Contando Cancelaciones y reservas por meses
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
            
            
        Next
        
        i = i + 1
        

    
    Next mycell
    
    

    
    'MsgBox " RESERVAS " & vbNewLine & "Enero:" & " " & Enero & vbNewLine & "Febrero:" & " " & Febrero & vbNewLine & "Marzo:" & " " & Marzo & vbNewLine & "Abril:" & " " & Abril & vbNewLine & "Mayo:" & " " & Mayo & vbNewLine & "Junio:" & " " & Junio & vbNewLine & "Julio:" & " " & Julio & vbNewLine & "Agosto:" & " " & Agosto & vbNewLine & "Septiembre:" & " " & Septiembre & vbNewLine & "Octubre:" & " " & Octubre & vbNewLine & "Noviembre:" & " " & Noviembre & vbNewLine & "Diciembre:" & " " & Diciembre
    
    'MsgBox " CANCELACIONES " & vbNewLine & "Enero:" & " " & cEnero & vbNewLine & "Febrero:" & " " & cFebrero & vbNewLine & "Marzo:" & " " & cMarzo & vbNewLine & "Abril:" & " " & cAbril & vbNewLine & "Mayo:" & " " & cMayo & vbNewLine & "Junio:" & " " & cJunio & vbNewLine & "Julio:" & " " & cJulio & vbNewLine & "Agosto:" & " " & cAgosto & vbNewLine & "Septiembre:" & " " & cSeptiembre & vbNewLine & "Octubre:" & " " & cOctubre & vbNewLine & "Noviembre:" & " " & cNoviembre & vbNewLine & "Diciembre:" & " " & cDiciembre
    
    
    'Regresion de reservas por semana
    Grph.Range("P2").Value = s1
    Grph.Range("Q2").Value = s2
    Grph.Range("R2").Value = s3
    Grph.Range("S2").Value = s4
    Grph.Range("T2").Value = s5
    Grph.Range("U2").Value = s6
    Grph.Range("V2").Value = s7
    Grph.Range("W2").Value = s8
    Grph.Range("X2").Value = s9
    Grph.Range("Y2").Value = s10
    Grph.Range("Z2").Value = s11
    Grph.Range("AA2").Value = s12
    Grph.Range("AB2").Value = s13
    Grph.Range("AC2").Value = s14
    Grph.Range("AD2").Value = s15
    Grph.Range("AE2").Value = s16
    Grph.Range("AF2").Value = s17
    Grph.Range("AG2").Value = s18
    Grph.Range("AH2").Value = s19
    Grph.Range("AI2").Value = s20
    Grph.Range("AJ2").Value = s21
    Grph.Range("AK2").Value = s22
    Grph.Range("AL2").Value = s23
    Grph.Range("AM2").Value = s24
    Grph.Range("AN2").Value = s25
    Grph.Range("AO2").Value = s26
    Grph.Range("AP2").Value = s27
    Grph.Range("AQ2").Value = s28
    Grph.Range("AR2").Value = s29
    Grph.Range("AS2").Value = s30
    Grph.Range("AT2").Value = s31
    Grph.Range("AU2").Value = s32
    Grph.Range("AV2").Value = s33
    Grph.Range("AW2").Value = s34
    Grph.Range("AX2").Value = s35
    Grph.Range("AY2").Value = s36
    Grph.Range("AZ2").Value = s37
    Grph.Range("BA2").Value = s38
    Grph.Range("BB2").Value = s39
    Grph.Range("BC2").Value = s40
    Grph.Range("BD2").Value = s41
    Grph.Range("BE2").Value = s42
    Grph.Range("BF2").Value = s43
    Grph.Range("BG2").Value = s44
    Grph.Range("BH2").Value = s45
    Grph.Range("BI2").Value = s46
    Grph.Range("BJ2").Value = s47
    Grph.Range("BK2").Value = s48
    Grph.Range("BL2").Value = s49
    Grph.Range("BM2").Value = s50
    Grph.Range("BN2").Value = s51
    Grph.Range("BO2").Value = s52
    Grph.Range("BP2").Value = s53
    
    'Regresion de cancelaciones por semana
    Grph.Range("P3").Value = cs1
    Grph.Range("Q3").Value = cs2
    Grph.Range("R3").Value = cs3
    Grph.Range("S3").Value = cs4
    Grph.Range("T3").Value = cs5
    Grph.Range("U3").Value = cs6
    Grph.Range("V3").Value = cs7
    Grph.Range("W3").Value = cs8
    Grph.Range("X3").Value = cs9
    Grph.Range("Y3").Value = cs10
    Grph.Range("Z3").Value = cs11
    Grph.Range("AA3").Value = cs12
    Grph.Range("AB3").Value = cs13
    Grph.Range("AC3").Value = cs14
    Grph.Range("AD3").Value = cs15
    Grph.Range("AE3").Value = cs16
    Grph.Range("AF3").Value = cs17
    Grph.Range("AG3").Value = cs18
    Grph.Range("AH3").Value = cs19
    Grph.Range("AI3").Value = cs20
    Grph.Range("AJ3").Value = cs21
    Grph.Range("AK3").Value = cs22
    Grph.Range("AL3").Value = cs23
    Grph.Range("AM3").Value = cs24
    Grph.Range("AN3").Value = cs25
    Grph.Range("AO3").Value = cs26
    Grph.Range("AP3").Value = cs27
    Grph.Range("AQ3").Value = cs28
    Grph.Range("AR3").Value = cs29
    Grph.Range("AS3").Value = cs30
    Grph.Range("AT3").Value = cs31
    Grph.Range("AU3").Value = cs32
    Grph.Range("AV3").Value = cs33
    Grph.Range("AW3").Value = cs34
    Grph.Range("AX3").Value = cs35
    Grph.Range("AY3").Value = cs36
    Grph.Range("AZ3").Value = cs37
    Grph.Range("BA3").Value = cs38
    Grph.Range("BB3").Value = cs39
    Grph.Range("BC3").Value = cs40
    Grph.Range("BD3").Value = cs41
    Grph.Range("BE3").Value = cs42
    Grph.Range("BF3").Value = cs43
    Grph.Range("BG3").Value = cs44
    Grph.Range("BH3").Value = cs45
    Grph.Range("BI3").Value = cs46
    Grph.Range("BJ3").Value = cs47
    Grph.Range("BK3").Value = cs48
    Grph.Range("BL3").Value = cs49
    Grph.Range("BM3").Value = cs50
    Grph.Range("BN3").Value = cs51
    Grph.Range("BO3").Value = cs52
    Grph.Range("BP3").Value = cs53


    'Rregresion de reservas por mes
    Grph.Range("B2").Value = Enero
    Grph.Range("C2").Value = Febrero
    Grph.Range("D2").Value = Marzo
    Grph.Range("E2").Value = Abril
    Grph.Range("F2").Value = Mayo
    Grph.Range("G2").Value = Junio
    Grph.Range("H2").Value = Julio
    Grph.Range("I2").Value = Agosto
    Grph.Range("J2").Value = Septiembre
    Grph.Range("K2").Value = Octubre
    Grph.Range("L2").Value = Noviembre
    Grph.Range("M2").Value = Diciembre
    
    'Regresion de cancelaciones por mes
    Grph.Range("B3").Value = cEnero
    Grph.Range("C3").Value = cFebrero
    Grph.Range("D3").Value = cMarzo
    Grph.Range("E3").Value = cAbril
    Grph.Range("F3").Value = cMayo
    Grph.Range("G3").Value = cJunio
    Grph.Range("H3").Value = cJulio
    Grph.Range("I3").Value = cAgosto
    Grph.Range("J3").Value = cSeptiembre
    Grph.Range("K3").Value = cOctubre
    Grph.Range("L3").Value = cNoviembre
    Grph.Range("M3").Value = cDiciembre
    
    
    
    
    
    
    
    

End Sub