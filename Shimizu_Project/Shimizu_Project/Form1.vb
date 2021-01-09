Public Class Form1

    Private calclator As Calculator

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Debug.Print("form_load")
        calclator = New Calculator()

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        MsgBox(TextBox1.Text)

    End Sub

    ' 番号入力処理
    Private Sub Button_Click(sender As Object, e As EventArgs) Handles btn0.Click, btn1.Click, btn2.Click, btn3.Click, btn4.Click, btn5.Click, btn6.Click, btn7.Click, btn8.Click, btn9.Click
        Calc_field.Text += sender.Text
        calclator.Calc_operator = ""
    End Sub

    ' 演算子入力処理
    Private Sub btn_operator_Click(sender As Object, e As EventArgs) Handles btn_plus.Click, btn_minus.Click, btn_multi.Click, btn_div.Click

        ' 
        If Not calclator.Calc_operator.Equals("") Then
            calclator.Calc_operators.RemoveAt(calclator.Calc_operators.Count - 1)
        End If
        calclator.Calc_operator = sender.Text

        ' 数値なら
        If IsNumeric(Calc_field.Text) Then
            calclator.Numbers.Add(Calc_field.Text)
        End If

        calclator.Calc_operators.Add(sender.Text)
        ' 入力項目をクリア
        Calc_field.Text = ""

        Label1.Text = calclator.Formula()

    End Sub

    ' イコール入力時
    Private Sub btn_equal_Click(sender As Object, e As EventArgs) Handles btn_equal.Click

        calclator.Numbers.Add(Calc_field.Text)
        Label1.Text = calclator.Formula()


        If Not Calc_field.Text.Equals("") Then
            calclator.Calc()


            calclator.Calc_operator = ""

        End If
        Calc_field.Text = calclator.Calc_Result_Value()

        calclator.Clear()
    End Sub


End Class
