Public Class percobaanloop

    Private Sub Button1_Click(sender As System.Object, e As System.EventArgs) Handles Button1.Click
        If RadioButton2.Checked = True Then
            Dim A, B, C As Integer
            ListBox1.Items.Clear()
            For A = TextBox1.Text To TextBox2.Text Step 1
                For B = 1 To TextBox3.Text Step 1
                    C = A * B
                    If A = 1 Then
                        ListBox1.Items.Add(A & "x" & B & "=" & C)
                    ElseIf A = 2 Then
                        ListBox1.Items.Add(A & "x" & B & "=" & C)
                    Else
                        ListBox1.Items.Add(A & "x" & B & "=" & C)
                    End If
                Next
                ListBox1.Items.Add("--------------")
            Next
        ElseIf RadioButton1.Checked = True Then
            Dim myArray(10) As Integer
            ListBox1.Items.Clear()
            For i As Integer = 1 To TextBox3.Text Step 1
                For j As Integer = 1 To TextBox3.Text Step 1
                    Dim C As Integer = i * j

                    ListBox1.Items.Add(i & "x" & j & "=" & C)

                Next
                ListBox1.Items.Add("--------------")
            Next
        End If
    End Sub

    Private Sub Button2_Click(sender As System.Object, e As System.EventArgs) Handles Button2.Click
        Dim myArray(10) As Integer
        For i As Integer = 2 To 10 Step 1
            myArray(i) = i * i
        Next

        Dim txt As String = ""
        For i As Integer = 0 To 10 Step 1
            txt &= myArray(i) & vbCrLf
        Next
        MsgBox(txt)
    End Sub

    
End Class
