Public Class PercobaanRumus1
    Public alfa, srumus, s1, s2, s3, at, ct, bt, ft, ftt, PE, APE, MPE, MAPE, Angmin, min, MUN As Double
    Public bulan As New ArrayList()
    Private Sub Button1_Click(sender As System.Object, e As System.EventArgs) Handles Button1.Click
        ListView1.Items.Clear()
        ListBox1.Items.Clear()
        ListBox2.Items.Clear()
        ListBox4.Items.Clear()
        ListView3.Items.Clear()
        ListView2.Items.Clear()

        Dim i As Integer
        Dim j As Integer
        

        For i = 1 To 9
            ListBox1.Items.Add("# ALFA " & i / 10 & " :")
            ListBox1.Items.Add("------------------------")
            ListBox2.Items.Add("# ALFA " & i / 10 & " :")
            ListBox2.Items.Add("------------------------")
            ListBox4.Items.Add("# ALFA " & i / 10 & " :")
            ListBox4.Items.Add("------------------------")
            MPE = 0
            MAPE = 0
            'PERULANGAN UNTUK T = 2
            For j = 2 To bulan.Count - 1 Step 1

                alfa = i / 10

                'srumus = setengah bagian dari rumus
                srumus = (1 - alfa) * bulan(j - 1)

                'S1 Adalah rumus S'T
                'S’t = aXt + (1-a)S’t -1
                'S’t = 0.1 (612)+0.9(570)
                s1 = FormatNumber((alfa * bulan(j)) + srumus, 2)

                'S2 Adalah rumus S'T
                'S” t = aS’ t + (1-a)S” t -1
                'S”2 = 0.1 (574.2)+0.9(570)
                s2 = FormatNumber((alfa * s1) + srumus, 2)

                'S3 Adalah rumus S'''T
                'S’” t = aS’’ t + (1-a)S’” t -1
                'S’”2 = 0.1 (570.42)+0.9(570)
                s3 = FormatNumber((alfa * s2) + srumus, 2)

                'a t = 3 S’ t – 3 S” t + S’” t
                at = FormatNumber((3 * s1) - (3 * s2) + s3, 2)

                'b t = {a/2(1-a)^^}*{(6 – 5a) S’ t – (10 – 8a) S” t + (4 – 3a)S’” t }
                bt = FormatNumber(((alfa / 2) * ((1 - alfa) * (1 - alfa))) * (((6 - (5 * alfa)) * s1) - ((10 - (8 * alfa)) * s2) + ((4 - (3 * alfa)) * s3)), 2)

                'c t = {a2/(1-a)2}*(S’t – 2S”t + S’”t )
                'c t = {(0.1)2/(0.9)2}*(574.2 – 2(570.42) + 570.04 )
                ct = FormatNumber(((alfa * alfa) / ((1 - alfa) * (1 - alfa))) * (s1 - (2 * s2) + s3), 2)

                'F t = a t + b t + ½ c t 2
                'F t = 581.38 + 0.785 + ½ (0.042) 2
                ft = FormatNumber(at + bt + ((1 / 2) * (ct * ct)), 2)

                'LISTBOX PERTAMA MENAMPILKAN HASIL RAMALAN FT 9 ALFA
                ListBox1.Items.Add(" ~ Bulan ke " & j)
                ListBox1.Items.Add("  S'T = " & s1)
                ListBox1.Items.Add("  S''T = " & s2)
                ListBox1.Items.Add("  S'''T = " & s3)
                ListBox1.Items.Add("  AT = " & at)
                ListBox1.Items.Add("  BT = " & bt)
                ListBox1.Items.Add("  CT = " & ct)
                ListBox1.Items.Add("")
                ListBox1.Items.Add("Hasil ramalan bulan ke " & j + 1 & " :")
                ListBox1.Items.Add("  FT = " & ft)
                ListBox1.Items.Add("")

                Dim lvvv As ListViewItem = ListView3.Items.Add(alfa)
                lvvv.SubItems.Add(j + 1)
                lvvv.SubItems.Add(FormatNumber(ft))

                'PERHITUNGAN PE APE 9 ALFA PADA T = 3
                If j + 1 < bulan.Count Then
                    PE = ((bulan(j + 1) - ft) / bulan(j + 1)) * 100
                    APE = FormatNumber(Math.Abs(PE), 2)
                    ListBox2.Items.Add(" ~ Bulan ke " & j + 1)
                    ListBox2.Items.Add("Jumlah Pengunjung : " & bulan(j + 1))
                    ListBox2.Items.Add("PE : " & PE)
                    ListBox2.Items.Add("APE : " & APE)

                    'RUMUS MENCARI TOTAL PE DAN APE UNTUK MPE DAN MAPE
                    MPE = MPE + PE
                    MAPE = MAPE + APE
                End If

            Next

            'PENAMBAHAN DATA PRESENTASE KESALAHAN PADA LISTVIEW
            ListBox4.Items.Add("MPE : " & MPE / 10)
            ListBox4.Items.Add("MAPE : " & MAPE / 10)
            Dim lv As ListViewItem = ListView1.Items.Add(i / 10)
            lv.SubItems.Add(ft)
            lv.SubItems.Add(FormatNumber(MPE, 2))
            lv.SubItems.Add(FormatNumber(MPE / 10, 2))
            lv.SubItems.Add(FormatNumber(MAPE / 10, 2))


            'PENAMBAHAN JARAK PADA LISTBOX MPE DAN MAPE
            ListBox4.Items.Add("")
            ListBox4.Items.Add("")
            ListBox4.Items.Add("")

            'PENAMBAHAN JARAK PADA LISTBOX FT
            ListBox1.Items.Add("")
            ListBox1.Items.Add("")
            ListBox1.Items.Add("")

            'PEAMBAHAN JARAK PADA LISTBOX PE APE
            ListBox2.Items.Add("")
            ListBox2.Items.Add("")
            ListBox2.Items.Add("")

        Next

        'MENCARI NILAI MAPE TERKECIL
        Angmin = ListView1.Items(0).SubItems(4).Text
        For l = 0 To 8
            If ListView1.Items(l).SubItems(4).Text < Angmin Then
                Angmin = ListView1.Items(l).SubItems(4).Text
                min = ListView1.Items(l).SubItems(0).Text
            End If
        Next
        Label6.Text = FormatNumber(Angmin, 2)
        Label8.Text = min

        'MENAMBAHKAN HASIL LAMARAN YANG BENAR KE LISTVIEW 2
        MUN = ListView3.Items(0).SubItems(0).Text
        For m = 0 To ListView3.Items.Count - 1
            If ListView3.Items(m).SubItems(0).Text = Label8.Text Then
                Dim lvv As ListViewItem = ListView2.Items.Add(ListView3.Items(m).SubItems(0).Text)
                lvv.SubItems.Add(ListView3.Items(m).SubItems(1).Text)
                lvv.SubItems.Add(ListView3.Items(m).SubItems(2).Text)
            End If
        Next




    End Sub

    Private Sub Button2_Click(sender As System.Object, e As System.EventArgs)

    End Sub

    Private Sub Button2_Click_1(sender As System.Object, e As System.EventArgs) Handles Button2.Click
        If bulan.Count = 0 Then
            bulan.Add("0")
            bulan.Add(TextBox1.Text)
            ListBox3.Items.Add(TextBox1.Text)
            TextBox1.Focus()
            TextBox1.Clear()
        Else
            Dim angka(100) As Integer
            Dim n As Integer
            'jumlah pengunjung sesungguhnya dari bulan ke 1 - 10
            n = n + 1
            angka(n) = CType(TextBox1.Text, Integer)
            'Tampilkan ke list array
            bulan.Add(angka(n))
            ListBox3.Items.Add(angka(n))
            TextBox1.Focus()
            TextBox1.Clear()
        End If
        
    End Sub
End Class