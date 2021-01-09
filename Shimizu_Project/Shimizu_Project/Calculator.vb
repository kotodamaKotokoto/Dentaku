Public Class Calculator

    Private m_calc_operator As String = ""


    ' 入力番号配列
    Private m_numbers As New List(Of String)
    ' 入力演算子配列
    Private m_calc_operators As New List(Of String)
    ' 計算結果
    Private m_calc_result As Decimal


    Public Property Calc_operator As String
        Get
            Return m_calc_operator
        End Get
        Set(value As String)
            m_calc_operator = value
        End Set
    End Property



    Public ReadOnly Property Calc_result As Decimal
        Get
            Return m_calc_result
        End Get
    End Property

    Public Property Numbers As List(Of String)
        Get
            Return m_numbers
        End Get
        Set(value As List(Of String))
            m_numbers = value
        End Set
    End Property

    Public Property Calc_operators As List(Of String)
        Get
            Return m_calc_operators
        End Get
        Set(value As List(Of String))
            m_calc_operators = value
        End Set
    End Property

    Public Sub Calc()
        Try
            'Select Case Calc_operator
            '    Case "+"
            '        m_calc_result = CInt(Number1) + CInt(Number2)
            '    Case "-"
            '        m_calc_result = CInt(Number1) - CInt(Number2)
            '    Case "*"
            '        m_calc_result = CInt(Number1) * CInt(Number2)
            '    Case "/"
            '        m_calc_result = Decimal.Parse(Number1) / Decimal.Parse(Number2) '0除算をキャッチするためにDecimal型に変換
            'End Select

        Catch dbze As DivideByZeroException
            MsgBox(dbze.Message)
        Catch ofe As OverflowException
            MsgBox(ofe.Message)
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    Public Function Formula() As String
        Dim m_formula As String = ""

        If Numbers.Count < 1 Or Calc_operators.Count < 1 Then Return m_formula

        For i As Integer = 0 To Calc_operators.Count - 1
            m_formula &= Numbers(i) & Calc_operators(i)
        Next

        ' 「＝」を入力時、演算子配列の要素数 < 数値配列の要素数 になるため
        If Calc_operators.Count < Numbers.Count Then
            m_formula &= Numbers.Last & "="
        End If

        Formula = m_formula
    End Function

    Public Function Calc_Result_Value() As Decimal
        Dim result As Decimal = 0
        Calc_Result_Value = result

        Try
            result = Numbers(0)
            For i As Integer = 0 To Calc_operators.Count - 1
                Select Case Calc_operators(i)
                    Case "+"
                        result += CInt(Numbers(i + 1))
                    Case "-"
                        result -= CInt(Numbers(i + 1))
                    Case "*"
                        result *= CInt(Numbers(i + 1))
                    Case "/"
                        result /= CInt(Numbers(i + 1))
                End Select
            Next

            Calc_Result_Value = result

        Catch ee As InvalidCastException
            MsgBox(ee.Message)
        Catch ex As ArgumentNullException
            MsgBox(ex.Message)
        Catch e As Exception
            MsgBox(e.Message)
        End Try

    End Function

    Public Sub Clear()
        Numbers.Clear()
        m_calc_result = 0
        Calc_operator = ""
        Calc_operators.Clear()

    End Sub

End Class
