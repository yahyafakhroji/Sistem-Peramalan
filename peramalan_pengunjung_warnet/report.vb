Public Class report
    'mendeklarasikan nomor sebagai integer
    Public nomor As Integer
    'Membuat Sub

    Sub loadxml()
        'Nomor adalah jumlah item pada listview pengunjung (di Form Home pada tab tambah data) + 1
        ' + 1 karena yang ditampilkan dalam report akan ditambahkan satu bulan yang merupakan bulan yang diramal
        nomor = Home.ListView_pengunjung.Items.Count + 1

        'Script dibawah ini merupakan penambahan satu item listview pengunjung,
        'karena item yang akan dimasukkan ke dalam report adalah semua bulan yang diinputkan + bulan yang diramal,
        'sedangkan bulan yang diramal tidak ada didalam listview pengunjung, 
        'maka bulan yang diramal ditambahkan terlebih dahulu kedalam listview pengunjung 
        'agar mempermudah saat menampilkan ke dalam report
        '(sebagai bantuan)
        Dim chartlv As ListViewItem = Home.ListView_pengunjung.Items.Add(nomor + 1)
        chartlv.SubItems.Add(Home.lbl_bln_yg_diramal.Text)
        chartlv.SubItems.Add("")

        'Mendeklarasikan variabel sebagai parameter value
        'Paramater digunakan untuk menampilkan object selain item listview ke dalam report (contoh : textbox, label)
        Dim pvCollection, pvCollection2, pvCollection3, pvCollection4, pvCollection5, pvCollection6, pvCollection7 As New CrystalDecisions.Shared.ParameterValues()
        Dim pdvFrom, pdvFrom2, pdvFrom3, pdvFrom4, pdvFrom5, pdvFrom6, pdvFrom7 As New CrystalDecisions.Shared.ParameterDiscreteValue()


        'Objek dari report yang kita buat
        'mendeklarasikan sebagai crrpt (crrpt adalah report yang kita buat. crrpt.rpt)
        Dim MyRpt As New crrpt

        'Dataset dan Datarow objek yang diperlukan untuk membuat Data Source
        Dim row As DataRow = Nothing
        Dim DS As New DataSet

        'Add Table ke Dataset
        'menambahkan tabel ke dalam dataset dengan nama ListViewData
        DS.Tables.Add("ListViewData")

        'Add Kolom ke Table
        'Setelah ditambahkan tabel tadi, kemudian menambahkan kolom kedalam tabel
        'dibawah ini ada 6 field (kolom),
        'Tapi yang dipakai nanti adalah sebanyak 4 Field saja
        With DS.Tables(0).Columns
            .Add("Field1", Type.GetType("System.String")) 'Tipe String
            .Add("Field2", Type.GetType("System.String")) 'Tipe String
            .Add("Field3", Type.GetType("System.String")) 'Tipe String
            .Add("Field4", Type.GetType("System.String")) 'Tipe String
            .Add("Field5", Type.GetType("System.Number"))
            .Add("Field6", Type.GetType("System.String"))

        End With

        'Loop Listview dan Menambahkan sebuah Row ke Table untuk setiap Listviewitem
        'melakukan perulangan untuk membaca setiap item yang brada di listview Hasil Ramalan
        For i = 0 To Home.ListView_hasil.Items.Count - 1
            row = DS.Tables(0).NewRow
            row(0) = Home.ListView_hasil.Items(i).SubItems(0).Text 'Item nomor
            row(1) = Home.ListView_hasil.Items(i).SubItems(1).Text 'Item nama bulan
            row(2) = Home.ListView_hasil.Items(i).SubItems(2).Text  'Item hasil ramalan
            row(3) = Home.ListView_pengunjung.Items(i).SubItems(2).Text 'Item Pengunjung awal (diambil dari listview pengunjung)

            DS.Tables(0).Rows.Add(row) 'Menambahkan row diatas
        Next i

        'Set Report Source Ke Database
        'data source dari MyRpt (crrpt) adalah DS (dataset diatas)
        MyRpt.SetDataSource(DS)

        'Memasukan Report ke Crystal report viewer Control (berada di report.vb)
        'Dispose Dataset
        DS.Dispose()
        DS = Nothing

        'menambahkan Parameter Value, diambil dari objek dibawah ini

        pdvFrom2.Value = Home.lbl_angkamin.Text 'angka minimum (berada di Form Home tab Proses Perhitungan )
        pdvFrom3.Value = Home.Lbl_hsl_bln.Text  'Bulan yang diramal (berada di Form Home tab Hasil Ramalan)
        pdvFrom4.Value = Home.lbl_hsl_rmln.Text 'Hasil ramalan bulan yang diramal (berada di Form Home tab Hasil Ramalan)
        pvCollection2.Add(pdvFrom2)
        pvCollection3.Add(pdvFrom3)
        pvCollection4.Add(pdvFrom4)
        pvCollection.Add(pdvFrom)

        ' add value  ke crrpt.rpt / MyRpt (report yang telah kita buat
        'nilai, bulan, hasil adalah nama parameter (berada di crrpt.rpt)
        MyRpt.DataDefinition.ParameterFields("nilai").ApplyCurrentValues(pvCollection2)
        MyRpt.DataDefinition.ParameterFields("bulan").ApplyCurrentValues(pvCollection3)
        MyRpt.DataDefinition.ParameterFields("hasil").ApplyCurrentValues(pvCollection4)

        'Terakhir set report source dari MyRpt yang telah dibuat dari awal (item listview) sampai ke parameter
        'Crystal report viewer (berada di report.vb) report source = MyRpt
        CrystalReportViewer1.ReportSource = MyRpt


    End Sub
    Private Sub report_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load
        'Load sub LoadXml
        loadxml()
    End Sub
End Class