
Class MainWindow


    Private Sub Button_Click(sender As Object, e As RoutedEventArgs)
        Valida_CIF(Label1.Text)
    End Sub


    'Valida que un cif introducido sea correcto
    'según la Orden EHA/451/2008
    Public Function Valida_CIF(ByVal valor As String) As Boolean
        Dim A As Integer
        Dim B As Integer
        Dim C As Integer
        Dim CIF As String
        Dim CIFDIGITO As String
        Dim i As Integer

        A = 0
        B = 0

        Valida_CIF = False

        If Len(valor) <> 9 Then 'el CIF debe tener 9 cifras
            Exit Function
        End If

        CIF = Mid(valor, 2, 7) 'se obtienen los dígitos centrales

        CIFDIGITO = Right(valor, 1) 'dígito de control

        For i = 1 To 6 Step 2
            A = A + Mid(CIF, i + 1, 1)   'Suma de posiciones pares
            C = 2 * Mid(CIF, i, 1)       'Doble de posiciones impares
            B = B + (C Mod 10) + Int(C / 10)   'Suma de digitos de doble de pares
        Next i
        'para obtener el cálculo de la cifra de la séptima posición que no se trata
        'en el bucle
        B = B + ((2 * Mid(CIF, 7, 1)) Mod 10) + Int((2 * Mid(CIF, 7, 1)) / 10)

        'se obtiene la unidad de la cifra total
        C = (10 - ((A + B) Mod 10)) Mod 10

        Dim Digito As String

        Dim letras() As String = {"J", "A", "B", "C", "D", "E", "F", "G", "H", "I"}
        Select Case (Strings.Left(valor, 1))
        'los cif que comienzan por estas letras deben terminar en una letra
        'concreta de la lista anterior
            Case "K", "P", "R", "Q", "S", "W" : Digito = letras(C)

        'los cif que comienzan por estas letras deben terminar en un dígito
            Case "A", "B", "E", "H", "J", "U", "V" : Digito = C

            Case "X", "Y", "Z"
                ' Error: es un NIE

                'para el resto de cif, la terminación puede ser un número o una letra
            Case Else
                If IsNumeric(CIFDIGITO) Then
                    Digito = C
                Else
                    Digito = letras(C)
                End If
        End Select
        Valida_CIF = (CIFDIGITO = Digito)
        MessageBox.Show(Valida_CIF)

    End Function


End Class
