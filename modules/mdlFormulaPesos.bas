Attribute VB_Name = "FormulaPesos"
Public Function NumAlfa(numtotal As String)
    Dim nument, numfrac, cant, cant2 As String
    Dim cifras(3) As String
    Dim i As Integer
    
    nument = Mid$(numtotal, 1, IIf(InStr(1, numtotal, ".") = 0, Len(numtotal), InStr(1, numtotal, ".") - 1))
    nument = Right("000000000" + nument, 9)
    numfrac = IIf(InStr(1, numtotal, ".") = 0, "", Mid$(numtotal, InStr(1, numtotal, ".") + 1, 2))
    For i = Len(nument) - 1 To 0 Step -1
        cifras(i \ 3) = Mid$(nument, i + 1, 1) + cifras(i \ 3)
    Next i
    For i = 0 To 2
        cant2 = decodifica(cifras(i))
        cant = cant + cant2
        If Len(cant2) > 0 Then
            Select Case i
            Case 0
                cant = cant + IIf(InStr(1, cant, "un"), "millon ", "millones ")
            Case 1
                cant = cant + "mil "
            Case 2
                cant = cant + ""
            End Select
        End If
    Next i
    
    cant = cant & IIf(Len(numfrac) > 0, "." & numfrac, "")
    NumAlfa = UCase(cant)

End Function

Public Function PesosAlfa(numtotal As String)
    Dim nument, numfrac, cant, cant2 As String
    Dim cifras(3) As String
    Dim i As Integer
    
    nument = Mid$(numtotal, 1, IIf(InStr(1, numtotal, ".") = 0, Len(numtotal), InStr(1, numtotal, ".") - 1))
    nument = Right("000000000" + nument, 9)
    numfrac = IIf(InStr(1, numtotal, ".") = 0, "", Mid$(numtotal, InStr(1, numtotal, ".") + 1, 2))
    For i = Len(nument) - 1 To 0 Step -1
        cifras(i \ 3) = Mid$(nument, i + 1, 1) + cifras(i \ 3)
    Next i
    For i = 0 To 2
        cant2 = decodifica(cifras(i))
        cant = cant + cant2
        If Len(cant2) > 0 Then
            Select Case i
            Case 0
                cant = cant + IIf(InStr(1, cant, "un"), "millon ", "millones ")
            Case 1
                cant = cant + "mil "
            Case 2
                cant = cant + ""
            End Select
        End If
    Next i
    
    cant = cant & "pesos " & IIf(Len(numfrac) > 0, numfrac & "/100 M.N.", "00/100 M.N.")
    PesosAlfa = UCase(cant)
    
End Function

Public Function DolaresAlfa(numtotal As String)
    Dim nument, numfrac, cant, cant2 As String
    Dim cifras(3) As String
    Dim i As Integer
    
    nument = Mid$(numtotal, 1, IIf(InStr(1, numtotal, ".") = 0, Len(numtotal), InStr(1, numtotal, ".") - 1))
    nument = Right("000000000" + nument, 9)
    numfrac = IIf(InStr(1, numtotal, ".") = 0, "", Mid$(numtotal, InStr(1, numtotal, ".") + 1, 2))
    For i = Len(nument) - 1 To 0 Step -1
        cifras(i \ 3) = Mid$(nument, i + 1, 1) + cifras(i \ 3)
    Next i
    For i = 0 To 2
        cant2 = decodifica(cifras(i))
        cant = cant + cant2
        If Len(cant2) > 0 Then
            Select Case i
            Case 0
                cant = cant + IIf(InStr(1, cant, "un"), "millon ", "millones ")
            Case 1
                cant = cant + "mil "
            Case 2
                cant = cant + ""
            End Select
        End If
    Next i
    
    cant = cant & "dolares " & IIf(Len(numfrac) > 0, numfrac & "/100 DLS", "00/100 DLS")
    DolaresAlfa = UCase(cant)
    
End Function
Private Function decodifica(cifra As String)
    Dim i As Integer
    Dim cant As String
    
    For i = 1 To Len(cifra)
    Select Case i
    Case 1 'centenas
        Select Case Mid$(cifra, i, 1)
        Case "0"
        Case "1"
            cant = cant + IIf(Mid$(cifra, i + 1, 1) = "0" And Mid$(cifra, i + 2, 1) = "0", "cien ", "ciento ")
        Case "2"
            cant = "doscientos "
        Case "3"
            cant = "trescientos "
        Case "4"
            cant = "cuatrocientos "
        Case "5"
            cant = "quinientos "
        Case "6"
            cant = "seiscientos "
        Case "7"
            cant = "setecientos "
        Case "8"
            cant = "ochocientos "
        Case "9"
            cant = "novecientos "
        End Select
    Case 2 'decenas
        Select Case Mid$(cifra, i, 1)
        Case "0"
        Case "1"
            Select Case Mid$(cifra, i + 1, 1)
            Case "0"
                cant = cant + "diez "
            Case "1"
                cant = cant + "once "
            Case "2"
                cant = cant + "doce "
            Case "3"
                cant = cant + "trece "
            Case "4"
                cant = cant + "catorce "
            Case "5"
                cant = cant + "quince "
            Case Else
                cant = cant + "dieci"
            End Select
        Case "2"
            cant = cant + IIf(Mid$(cifra, i + 1, 1) = 0, "veinte ", "veinti")
        Case "3"
            cant = cant + IIf(Mid$(cifra, i + 1, 1) = 0, "treinta ", "treinta y ")
        Case "4"
            cant = cant + IIf(Mid$(cifra, i + 1, 1) = 0, "cuarenta ", "cuarenta y ")
        Case "5"
            cant = cant + IIf(Mid$(cifra, i + 1, 1) = 0, "cincuenta ", "cincuenta y ")
        Case "6"
            cant = cant + IIf(Mid$(cifra, i + 1, 1) = 0, "sesenta ", "sesenta y ")
        Case "7"
            cant = cant + IIf(Mid$(cifra, i + 1, 1) = 0, "setenta ", "setenta y ")
        Case "8"
            cant = cant + IIf(Mid$(cifra, i + 1, 1) = 0, "ochenta ", "ochenta y ")
        Case "9"
            cant = cant + IIf(Mid$(cifra, i + 1, 1) = 0, "noventa ", "noventa y ")
        End Select
    
    Case 3 'unidades
        If Not (Mid$(cifra, i - 1, 1) = "1" And Val(Mid$(cifra, i, 1)) <= 5) Then
        Select Case Mid$(cifra, i, 1)
        Case "0"
        Case "1"
             cant = cant + "un "
        Case "2"
             cant = cant + "dos "
        Case "3"
            cant = cant + "tres "
        Case "4"
            cant = cant + "cuatro "
        Case "5"
            cant = cant + "cinco "
        Case "6"
            cant = cant + "seis "
        Case "7"
            cant = cant + "siete "
        Case "8"
            cant = cant + "ocho "
        Case "9"
            cant = cant + "nueve "
        End Select
        End If
    End Select
    Next i
    decodifica = cant
End Function
