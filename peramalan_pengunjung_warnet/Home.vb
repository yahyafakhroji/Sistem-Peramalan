Public Class Home
    Public alfa, srumus, s1, s2, s3, at, ct, bt, ft, ftt, PE, APE, MPE, MAPE, Angmin, min, MUN As Double
    Public bulan As New ArrayList() 'pendeklarasian publik (untuk semua sub) bulan sebagai arraylist baru
    Public namabulan As New ArrayList() 'pendeklarasian publik (untuk semua sub) namabulan sebagai arraylist baru
    Public nbulan As Integer 'pendeklarasian publik (untuk semua sub) nbulan dengan tipe data integer
    Public nomor As Integer 'pendeklarasian publik (untuk semua sub) nomor dengan tipe data integer

    'Membuat sub namaaaa, yang berisi arraylist untuk nama bulan 
    'agar memudahkan saat melakukan perulangan pada penambahan bulan 
    'di list input data pengunjung awal.
    Private Sub namaaaa()
        namabulan.Add("Januari")
        namabulan.Add("Februari")
        namabulan.Add("Maret")
        namabulan.Add("April")
        namabulan.Add("Mei")
        namabulan.Add("Juni")
        namabulan.Add("Juli")
        namabulan.Add("Agustus")
        namabulan.Add("September")
        namabulan.Add("Oktober")
        namabulan.Add("November")
        namabulan.Add("Desember")
    End Sub

    'LOAD FORM
    Private Sub Home_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load
        ''Penamaan bulan, hasil sub namaaaa (untuk nama bulan) yang tadi telah dibuat,
        'akan dipanggil saat load form atau saat pertama kali form dibuka dan dijalankan.

        namaaaa() 'memanggil sub namaaaa
        namaaaa() 'memanggil sub namaaaa untuk kedua kalinya.
        '   Kegunaan memanggil sub namaaaa dua kali yaitu, -
        'karena didalam sub namaaaa diisikan data nama bulan sebanyak 12 bulan (jan - des), -
        'sedangkan pada saat menginputkan data pengunjung, nama bulan urut, -
        'Jadi setelah bulan desember kembali ke januari dan seterusnya dengan batas input 12 bulan. 
        'Agar setelah bulan desember bisa kembali ke bulan awal, - 
        'sub namaaaa dipanggil dua kali. => (Januari , ...... , desember, januari, ..... , desember)


        'Saat pertama kali form di load, tombol dibawah ini dinonaktifkan
        btn_tambah.Enabled = False
        txt_tambah.Enabled = False
        Btn_hapus.Enabled = False
        btn_update.Enabled = False

    End Sub

    '***** Setelah script Load form, akan ada script pembuatan sub untuk memudahkan pemanggilan script dalam sub. ******'

    Private Sub txt_tambah_focus_clr()
        'Sub yang digunakan untuk menempatkan kursor pada textbox tambah dan membersihkan text dari textbox
        txt_tambah.Focus()
        txt_tambah.Clear()
    End Sub
    Private Sub clear()
        'Sub yang digunakan untuk membersihkan isi/items/text dari :
        listview_NKTA.Items.Clear() 'ListView Nilai kesalahan tiap alfa, pada Tab Proses Perhitungan
        ListBox1.Items.Clear() ' Listbox Proses Peramalan pada Tab Proses View
        ListBox2.Items.Clear() ' Listbox Nilai Presentase Kesalahan pada Tab Proses View
        ListBox4.Items.Clear() ' Listbox Nilai kesalahan tiap alfa pada Tab Proses View
        Listview_PP.Items.Clear() 'Listview Proses Peramalan dan Nilai Presentase Kesalahan pada tab Proses Perhitungan
        ListView_hasil.Items.Clear() 'Listview Hasil Ramalan pada Tab Hasil Ramalan

        'label alfaminimum dan angka minimum pada tab proses perhitungan,
        'label hasil ramalan dan hasil bulan pada tab hasil ramalan
        lbl_alfamin.Text = "_ _ _ _ _"
        lbl_angkamin.Text = "_ _ _ _ _"
        lbl_hsl_rmln.Text = "_ _ _ _ _"
        Lbl_hsl_bln.Text = "_ _ _ _ _"
    End Sub
    Private Sub tombolhapus()
        'Sub yang digunakan untuk mengaktifkan dan menonaktifkan tombol hapus pada saat input data pengunjung awal,
        'dalam kondisi jika listview pengunjung kosong, dan tidak kosong.
        If ListView_pengunjung.Items.Count = 0 Then
            Btn_hapus.Enabled = False
        Else
            Btn_hapus.Enabled = True
        End If
    End Sub
    Private Sub tambahlvblnygdiramal()
        'Penambahan bulan yang diramal untuk chart, untuk penjelasan fungsi sub ini akan dijelaskan dibawah saat pemanggilan sub.

        Dim chartlv As ListViewItem = ListView_pengunjung.Items.Add(nomor + 1)
        chartlv.SubItems.Add(lbl_bln_yg_diramal.Text)
        chartlv.SubItems.Add("")

        ListView_hasil.Items(0).SubItems(0).Text = lbl_alfamin.Text
        ListView_hasil.Items(1).SubItems(0).Text = lbl_alfamin.Text

        For n = 2 To ListView_hasil.Items.Count - 1
            ListView_hasil.Items(n).SubItems(3).Text = FormatNumber(ListView_hasil.Items(n).SubItems(2).Text, 0).Replace(".", "")
        Next

    End Sub
    Private Sub reset()
        Dim tanya As String
        tanya = MsgBox("Yakin akan mereset data ?", vbYesNo, "konfirmasi")
        If tanya = vbYes Then
            ListView_pengunjung.Items.Clear()
            cmb_bulanawal.SelectedItem = Nothing
            cmb_bulanawal.Enabled = True
            btn_tambah.Enabled = False
            txt_tambah.Enabled = False
            Btn_hapus.Enabled = False
            btn_update.Enabled = False
            lbl_bln_yg_diramal.Text = "_ _ _ _ _"
            txt_tambah.Clear()
            clear()
            bulan.Clear()
            Chart1.Series("Pengunjung").Points.Clear()
            Chart1.Series("Ramalan").Points.Clear()
            TabControl1.SelectTab(0)

        End If
    End Sub
    Private Sub ComboBox1_SelectedIndexChanged(sender As System.Object, e As System.EventArgs) Handles cmb_bulanawal.SelectedIndexChanged
        txt_tambah.Enabled = True
        btn_tambah.Enabled = True
    End Sub

    Private Sub btn_tambah_Click(sender As System.Object, e As System.EventArgs) Handles btn_tambah.Click
        'ISI SCRIPT DARI TOMBOL TAMBAH saat input data pengunjung

        'saat tombol tambah di klik combobox bulan awal dan tombol update di nonaktifkan,
        'untuk memvalidasi agar pengguna tidak bisa memilih kembali bulan awal ketika tombol + / tambah di klik,
        'dan agar tidak memencet tombol update ketika data dari listview belum di klik untuk diupdate
        cmb_bulanawal.Enabled = False
        btn_update.Enabled = False


        nomor = ListView_pengunjung.Items.Count + 1 'variabel nomor sama dengan jumlah item listview pengunjung + 1,
        nbulan = cmb_bulanawal.SelectedIndex + 1 'variabel nbulan sama dengan nilai index dari bulan awal yang dipilih (urutan bulan) + 1,

        'Pengkondisian jika bulanawal tidak dipilih, atau ketika bulan awal dipilih.
        If cmb_bulanawal.SelectedItem = Nothing Then
            MsgBox("Pilihan bulan tidak boleh kosong")
        ElseIf txt_tambah.Text = "" Then
            MsgBox("Masukkan Data Pengunjung")
            txt_tambah.Focus()

        Else
            'dan jika bulan awal sudah dipilih,
            'melakukan pengkondisian kembali yaitu jumlah pengunjung harus diisi angka, dan banyak angka tidak boleh lebih dari 12
            If Not IsNumeric(txt_tambah.Text) Then
                MsgBox("Harus diisi dengan angka")
                txt_tambah_focus_clr()
            ElseIf Len(txt_tambah.Text) > 12 Then
                MsgBox("Tidak boleh lebih dari 12 digit")
                txt_tambah_focus_clr()
            Else
                'dan jika pengkondisian diatas sudah dilewati,
                'mengkondisikan kembali berdasarkan jumlah item listview pengunjung 
                If ListView_pengunjung.Items.Count < 12 Then 'jika item kurang dari 12, maka

                    'melakukan pengkondisian jumlah bulan
                    'jika jumlah bulan = 0 / list array masih kosong, maka
                    'pertama menambahkan bulan.Add("0") diawal
                    'kemudian diikuti menambahkan bulan.Add(txt_tambah.Text) / jumlah pengunjung yang diinputkan

                    If bulan.Count = 0 Then
                        bulan.Add("0")
                        bulan.Add(txt_tambah.Text)

                        'script untuk memasukkan data jumlah pengunjung pada listview
                        Dim lvvv As ListViewItem = ListView_pengunjung.Items.Add(nomor) 'nomor
                        lvvv.SubItems.Add(namabulan(cmb_bulanawal.SelectedIndex)) 'namabulan diambil dari arraylist di sub namaaaa pada index = index bulan awal yang dipilih di combobox
                        lvvv.SubItems.Add(txt_tambah.Text) 'jumlah pengunjung

                        'label bulan yang diramal diambil dari bulan setelah bulan terakhir yang diinputkan. ( .... + 1 )
                        lbl_bln_yg_diramal.Text = (namabulan(cmb_bulanawal.SelectedIndex + 1))

                        'memanggil sub txt_tambah_focus_clr untuk menempatkan kursor pada textbox tambah dan membersihkan isinya.
                        txt_tambah_focus_clr()

                    Else

                        'Sript untuk menambahkan item ke arraylist bulan dimana item tersebut adalah jumlah pengunjung yang ditambahkan
                        Dim angka(100) As Integer
                        Dim n As Integer
                        n = n + 1
                        angka(n) = CType(txt_tambah.Text, Integer)
                        'Tampilkan ke list array
                        bulan.Add(angka(n))

                        'Script untuk menambahkan data pengunjung yang telah diinputkan kedalam listview.
                        Dim lvvvv As ListViewItem = ListView_pengunjung.Items.Add(nomor) 'menambahkan nomor di kolom 1 listview pengunjung

                        'Melakukan perulangan untuk mengambil nama bulan melalui array list namabulan secara urut dari bulan awal yang dipilih, 
                        Dim m As Integer
                        For m = 0 To 24
                            'Jadi jika index bulan awal yang dipilih = m , 
                            'nama bulan yang ditambahkan kedalam listview adalah bulan selanjutnya setelah bulan awal dan -
                            'secara terus menerus berurutan setelah bulan ditambahkan

                            If cmb_bulanawal.SelectedIndex = m Then
                                'script untuk menambahkan nama bulan kedalam listview secara urut, 
                                'm - 1 artinya index bulan awal dikurangi 1 , karena index dimulai dari angka 0 maka harus dikurangi 1 agar mendapatkan nomor urut bulan awal yang tepat
                                'setelah itu ditambahkan "nomor", nomor = jumlah item di listview + 1 ( + 1 karena yg dicari index bulan selanjutnya dari bulan item terakhir di listview )
                                'Dan setelah index nya sudah ditentukan, maka akan mendapatkan nama bulan di arraylist namabulan sesuai dengan index yang ditemukan
                                lvvvv.SubItems.Add(namabulan((m - 1) + nomor))
                                'untuk mencari bulan yang diramal juga hampir sama, bedanya m tidak dikurangi 1
                                'artinya bulan selanjutnya dari bulan yang ditemukan diatas.
                                lbl_bln_yg_diramal.Text = (namabulan(m + nomor))
                            End If
                        Next
                        'menambahkan jumlah pengunjung yang diinputkan..
                        lvvvv.SubItems.Add(angka(n))
                        txt_tambah_focus_clr()
                    End If
                Else
                    MsgBox("Data pengunjung tidak boleh lebih dari 12 bulan")
                    txt_tambah.Clear()

                End If
            End If
        End If
        tombolhapus()

    End Sub
    Private Sub btn_hasil_Click(sender As System.Object, e As System.EventArgs) Handles btn_hasil.Click
        'menonaktifkan tombol update
        btn_update.Enabled = False


        'LANGKAH #1 mengkondisikan jika jumlah penungunjung yang diinputkan hanya sebanyak 3 bulan, 
        'jika iya maka tidak bisa menghitung hasil ramalan, -
        'jika tidak melanjutkan ke script perhitungan
        If ListView_pengunjung.Items.Count < 3 Then
            MsgBox("Data Pengunjung minimal 3 bulan")
        Else
            'LANGKAH #2 
            '1. membersihkan seluruh item terlebih dahulu (listview, listbox, label, textbox) menggunakan Sub Clear
            '2. setelah melakukan perulangan dari i = 1 ke 9, karena banyak nya alfa ada 9 alfa.
            '3. Perulangan digunakan untuk menghitung rumus pada setiap alfa (pada 9 alfa semuanya)
            'Jadi agar rumus secara berurutan bisa dihitung pada setiap alfa, dilakukan perulangan.
            clear()
            Dim i, j As Integer
            For i = 1 To 9
                'LISTBOX Pada Tab Proses View
                'script ini untuk memunculkan proses perhitungan,
                ' i / 10 artinya untuk menetapkan alfa secara desimal 0,...
                ListBox1.Items.Add("# ALFA " & i / 10 & " :")
                ListBox1.Items.Add("------------------------")
                ListBox2.Items.Add("# ALFA " & i / 10 & " :")
                ListBox2.Items.Add("------------------------")
                ListBox4.Items.Add("# ALFA " & i / 10 & " :")
                ListBox4.Items.Add("------------------------")

                'awal perhitungan MPE dan MAPE = 0
                MPE = 0
                MAPE = 0

                'LANGKAH #3
                'Penambahan DUA BULAN pertama pada listview Proses Peramalan
                'Script dibawah ini digunakan untuk menambahkan dua bulan pertama pada tabel proses perhitungan
                'Script dua bulan pertama dipisahkan seperti ini karena saat proses perhitungan dibawah nanti perulangan dimulai dari j = 2 (lihat perulangan kedua dibawah)
                ' yang artinya dimulai dari bulan ketiga. agar bulan secara lengkap masuk ke tabel proses perhitungan, maka dua bulan awal yang tidak ikut perhitungan di perulangan kedua, - 
                'akan ditambahkan pada script ini

                Dim lvhp1 As ListViewItem = Listview_PP.Items.Add(i / 10) 'Kolom pertama = Alfa = i/10 BULAN PERTAMA
                lvhp1.SubItems.Add(ListView_pengunjung.Items(0).SubItems(1).Text) 'Kolom kedua = nama bulan diambil dari listview Pengunjung (tabel saat input data awal)
                lvhp1.SubItems.Add("") 'Pengunjung dikosongkan karena dua bulan pertama masih tidak ada data pengunjung
                lvhp1.SubItems.Add(ListView_pengunjung.Items(0).SubItems(2).Text) 'Kolom ketiga jumlah pengunjung

                Dim lvhp2 As ListViewItem = Listview_PP.Items.Add(i / 10) 'Kolom pertama = Alfa = i/10 BULAN KEDUA
                lvhp2.SubItems.Add(ListView_pengunjung.Items(1).SubItems(1).Text)
                lvhp2.SubItems.Add("")
                lvhp2.SubItems.Add(ListView_pengunjung.Items(1).SubItems(2).Text)


                'LANGKAH #4
                'PERULANGAN UNTUK T = 2 
                For j = 2 To bulan.Count - 1 Step 1 'dimulai dari j = 2 ke banyaknya bulan yang telah diinputkan (ditentukan dari jumlah arraylist bulan)

                    alfa = i / 10

                    '! ArrayList bulan berisi jumlah pengunjung awal

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

                    'Listbox PROSES PERAMALAN di tab PROSES VIEW
                    'MENAMPILKAN HASIL RAMALAN 9 ALFA
                    'Yang ditampilkan adalah S'T, S''T, S'''T, AT, BT, CT, Hasil ramalan, FT
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

                    'Listview PROSES PERAMALAN DAN NILAI PRESENTASE KESALAHAN di tab PROSES PERHITUNGAN
                    'MENAMPILKAN HASIL RAMALAN 9 ALFA PADA listview
                    'yang ditampilkan adalah alfa, bulan, dan FT,
                    'untuk Pengunjung, PE, APE akan ditambahkan dibawah saat perhitungan PE APE
                    Dim lvvv As ListViewItem = Listview_PP.Items.Add(alfa)
                    lvvv.SubItems.Add(namabulan((j) + cmb_bulanawal.SelectedIndex))
                    lvvv.SubItems.Add(FormatNumber(ft))

                    'PERHITUNGAN PE APE 9 ALFA PADA T = 3
                    If j + 1 < bulan.Count Then
                        PE = ((bulan(j + 1) - ft) / bulan(j + 1)) * 100 'RUMUS PE
                        APE = FormatNumber(Math.Abs(PE), 2) 'RUMUS APE, Membulatkan nilai PE menjadi desimal dua angka dibelakang koma

                        'LISTBOX2 = Listbox NILAI PRESENTASE KESALAHAN di tab PROSES VIEW
                        ListBox2.Items.Add(" ~ Bulan ke " & j + 1)
                        ListBox2.Items.Add("Jumlah Pengunjung : " & bulan(j + 1))
                        ListBox2.Items.Add("PE : " & PE)
                        ListBox2.Items.Add("APE : " & APE)

                        'Listview PROSES PERAMALAN DAN NILAI PRESENTASE KESALAHAN di tab PROSES PERHITUNGAN
                        'Penambahan Pengunjung, PE, APE 
                        lvvv.SubItems.Add(bulan(j + 1))
                        lvvv.SubItems.Add(PE)
                        lvvv.SubItems.Add(APE)

                        'Rumus mencari MPE dan MAPE langkah pertama,
                        'MPE = Jumlah seluruh PE / 10
                        'MAPE = Jumlah seluruh APE / 10
                        'script ini menjumlahkan seluruh PE dan APE saja.
                        'Kemudian nanti akan dibagi 10 pada rumus selanjutnya
                        MPE = MPE + PE
                        MAPE = MAPE + APE

                    End If
                Next

                'PENAMBAHAN DATA  TIAP ALFA
                ListBox4.Items.Add("MPE : " & MPE / 10)
                ListBox4.Items.Add("MAPE : " & MAPE / 10)
                Dim lv As ListViewItem = listview_NKTA.Items.Add(i / 10)
                lv.SubItems.Add(ft)
                lv.SubItems.Add(FormatNumber(MPE, 2))
                lv.SubItems.Add(FormatNumber(MPE / 10, 2))
                lv.SubItems.Add(FormatNumber(MAPE / 10, 2))

                'PENAMBAHAN JARAK PADA LISTBOX MPE DAN MAPE
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
            Angmin = listview_NKTA.Items(0).SubItems(4).Text
            For l = 0 To 8
                If listview_NKTA.Items(l).SubItems(4).Text < Angmin Then
                    Angmin = listview_NKTA.Items(l).SubItems(4).Text
                    min = listview_NKTA.Items(l).SubItems(0).Text
                End If
            Next
            lbl_angkamin.Text = FormatNumber(Angmin, 2)
            lbl_alfamin.Text = min
            If lbl_alfamin.Text = "0" Then
                lbl_alfamin.Text = "0,1"
            End If

            'MENAMBAHKAN HASIL RAMALAN YANG BENAR KE LISTVIEW 2
            MUN = Listview_PP.Items(0).SubItems(0).Text
            For m = 0 To Listview_PP.Items.Count - 1
                If Listview_PP.Items(m).SubItems(0).Text = lbl_alfamin.Text Then
                    Dim lvv As ListViewItem = ListView_hasil.Items.Add(Listview_PP.Items(m).SubItems(0).Text)
                    lvv.SubItems.Add(Listview_PP.Items(m).SubItems(1).Text)
                    lvv.SubItems.Add(Listview_PP.Items(m).SubItems(2).Text)
                    'pembulatan untuk grafik
                    lvv.SubItems.Add(Listview_PP.Items(m).SubItems(2).Text)
                    Lbl_hsl_bln.Text = ListView_hasil.Items(ListView_hasil.Items.Count - 1).SubItems(1).Text
                    lbl_hsl_rmln.Text = ListView_hasil.Items(ListView_hasil.Items.Count - 1).SubItems(2).Text
                End If
            Next

            'MENAMPILKAN CHART / GRAFIK
            Chart1.Series("Pengunjung").Points.Clear()
            Chart1.Series("Ramalan").Points.Clear()
            tambahlvblnygdiramal()
            For o = 0 To ListView_pengunjung.Items.Count - 1
                Chart1.Series("Pengunjung").Points.AddXY(ListView_pengunjung.Items(o).SubItems(1).Text, ListView_pengunjung.Items(o).SubItems(2).Text)
            Next
            For p = 0 To ListView_hasil.Items.Count - 1
                Chart1.Series("Ramalan").Points.AddXY(ListView_hasil.Items(p).SubItems(1).Text, ListView_hasil.Items(p).SubItems(3).Text)
            Next

            'penghapusan kembali bulan yang diramal untuk chart
            Dim itemreplace1, itemreplace2 As String
            itemreplace1 = ListView_pengunjung.Items(ListView_pengunjung.Items.Count - 1).SubItems(0).Text
            itemreplace2 = ListView_pengunjung.Items(ListView_pengunjung.Items.Count - 1).SubItems(1).Text
            itemreplace1.Replace(itemreplace1, "")
            itemreplace2.Replace(itemreplace2, "")

            'penghapusan kembali pengunjung
            Dim a As Integer
            a = ListView_pengunjung.Items.Count - 1
            ListView_pengunjung.Items(a).Remove()

            'PINDAH TAB
            TabControl1.SelectTab(1)
        End If
    End Sub

    'Saat KLIK item di listview pengunjung (jika akan update data, maka harus klik item terlebih dahulu
    Private Sub ListView_pengunjung_Click(sender As System.Object, e As System.EventArgs) Handles ListView_pengunjung.Click
        'memindahkan item yang diklik ke textbox tambah
        txt_tambah.Text = ListView_pengunjung.SelectedItems.Item(0).SubItems(2).Text
        'tombol update diaktifkan
        btn_update.Enabled = True
    End Sub

    'Tombol HAPUS -> X di tab tambah data
    Private Sub Button7_Click(sender As System.Object, e As System.EventArgs) Handles Btn_hapus.Click
        Dim a As Integer
        a = ListView_pengunjung.Items.Count - 1 ' - 1 untuk mengubah menjadi urutan index (dimulai dari 0)

        'menghapus bulan terakhir, berdasarkan urutan a diatas
        ListView_pengunjung.Items(a).Remove()

        'juga menghapus arraylist bulan
        bulan.RemoveAt(bulan.Count - 1)

        Dim m As Integer
        'Looping untuk menentukan bulan yang diramal setelah melakukan penghapusan item diatas
        For m = 0 To 24 'to 24 karena jumlah item arraylist namabulan ada 24 (namabulan dua kali)
            If cmb_bulanawal.SelectedIndex = m Then '
                'bulan yang diramal diambil dari arraylist namabulan,
                'berdasarkan index yang dicari dari m (index bulan awal) + jumlah item pengunjung
                lbl_bln_yg_diramal.Text = (namabulan(m + ListView_pengunjung.Items.Count))
            End If
        Next
        txt_tambah_focus_clr()
        tombolhapus()
        'jika item di dalam listview pengunjung kosong, maka  tidak ada bulan yang akan diramal 
        If ListView_pengunjung.Items.Count = 0 Then
            lbl_bln_yg_diramal.Text = "_ _ _ _ _ "
        End If
        btn_update.Enabled = False
    End Sub

    'Tombol UPDATE di tab tambah data
    Private Sub Button8_Click(sender As System.Object, e As System.EventArgs) Handles btn_update.Click
        ListView_pengunjung.SelectedItems(0).SubItems(2).Text = txt_tambah.Text 'mengupdate item yang telah diselect dengan data baru di textbox tambah

        'mengupdate arraylist Bulan menggunakan insert dan remove
        'insert data baru, remove data lama = update
        bulan.Insert(ListView_pengunjung.SelectedItems(0).Index + 2, txt_tambah.Text) 'insert data baru
        bulan.RemoveAt(ListView_pengunjung.SelectedItems(0).Index + 1) 'remove data lama

        MsgBox("update berhasil")
        txt_tambah_focus_clr()
        btn_update.Enabled = False

    End Sub

    'Tombol BUAT DATA BARU
    Private Sub Button1_Click(sender As System.Object, e As System.EventArgs) Handles Button1.Click
        'memanggil sub reset
        reset()
    End Sub

    'Tombol RESET SEMUA DATA pada tab Grafik
    Private Sub Button7_Click_1(sender As System.Object, e As System.EventArgs) Handles Button7.Click
        'memanggil sub reset
        reset()
    End Sub

    '----- SCRIPT UNTUK GAYA GRAFIK ----------'

    'tombol radar pada tab grafik
    Private Sub Button13_Click(sender As System.Object, e As System.EventArgs) Handles Button13.Click
        Chart1.Series("Pengunjung").ChartType = DataVisualization.Charting.SeriesChartType.Radar
        Chart1.Series("Ramalan").ChartType = DataVisualization.Charting.SeriesChartType.Radar
    End Sub

    'tombol column pada tab grafik
    Private Sub Button11_Click(sender As System.Object, e As System.EventArgs) Handles Button11.Click
        Chart1.Series("Pengunjung").ChartType = DataVisualization.Charting.SeriesChartType.Column
        Chart1.Series("Ramalan").ChartType = DataVisualization.Charting.SeriesChartType.Column
    End Sub

    'tombol bar pada tab grafik
    Private Sub Button10_Click(sender As System.Object, e As System.EventArgs) Handles Button10.Click
        Chart1.Series("Pengunjung").ChartType = DataVisualization.Charting.SeriesChartType.Bar
        Chart1.Series("Ramalan").ChartType = DataVisualization.Charting.SeriesChartType.Bar
    End Sub

    'Tombol line pada tab grafik
    Private Sub Button9_Click(sender As System.Object, e As System.EventArgs) Handles Button9.Click
        Chart1.Series("Pengunjung").ChartType = DataVisualization.Charting.SeriesChartType.Line
        Chart1.Series("Ramalan").ChartType = DataVisualization.Charting.SeriesChartType.Line
    End Sub
    '----------------- [END] SCRIPT GAYA GRAFIK --------------'


    '------------------- SCRIPT UNTUK MEMBUKA TAB SELANJUTNYA, ATAU KEMBALI KE TAB SEBELUMNYA -------------'
    'Tombol HASIL RAMALAN pada tab PROSES PERHITUNGAN
    Private Sub Button2_Click(sender As System.Object, e As System.EventArgs) Handles Button2.Click
        'Membuka tab HASIL RAMALAN
        TabControl1.SelectTab(2)
    End Sub

    'Tombol VIEW PROSES pada tab Hasil Ramalan
    Private Sub Button3_Click(sender As System.Object, e As System.EventArgs) Handles Button3.Click
        'Membuka tab proses view
        TabControl1.SelectTab(3)
    End Sub

    'Tombol LIHAT GRAFIK pada tab Proses View
    Private Sub Button12_Click(sender As System.Object, e As System.EventArgs) Handles Button12.Click
        'Membuka tab grafik
        TabControl1.SelectTab(4)
    End Sub

    'Tombol BACK pada tab Grafik
    Private Sub Button8_Click_1(sender As System.Object, e As System.EventArgs) Handles Button8.Click
        'tombol back ini adalah kembali ke tab Proses View
        TabControl1.SelectTab(3)
    End Sub

    'Tombol BACK pada tab Proses Perhitungan 
    Private Sub Button4_Click(sender As System.Object, e As System.EventArgs) Handles Button4.Click
        'tombol back ini adalah kembali ke tab Tambah Data
        TabControl1.SelectTab(0)

        'Script tambahan, karena kembali ke tab awal, item terakhir yang ditambahkan (sebagai bantuan) akan dihapus kembali
        Dim a As Integer
        a = ListView_pengunjung.Items.Count - 1
        ListView_pengunjung.Items(a).Remove()
        clear()

    End Sub

    'Tombol BACK pada tab Hasil Ramalan
    Private Sub Button5_Click(sender As System.Object, e As System.EventArgs) Handles Button5.Click
        ''tombol back ini adalah kembali ke tab Proses Perhitungan
        TabControl1.SelectTab(1)
    End Sub

    'Tombol BACK pada tab Proses View
    Private Sub Button6_Click(sender As System.Object, e As System.EventArgs) Handles Button6.Click
        'tombol back ini adalah kembali ke tab hasil ramalan
        TabControl1.SelectTab(2)
    End Sub
    '-------------------[END] SCRIPT UNTUK MEMBUKA TAB SELANJUTNYA, ATAU KEMBALI KE TAB SEBELUMNYA -------------'
    
    'Tombol REPORT
    Private Sub Button14_Click(sender As System.Object, e As System.EventArgs) Handles Button14.Click
        'Mengkondisikan jika belum ada data ramalan maka tidak bisa membuka report
        If ListView_hasil.Items.Count = 0 Then
            MsgBox("Harus ada data hasil ramalan!")
        Else
            report.ShowDialog()
        End If
    End Sub

End Class