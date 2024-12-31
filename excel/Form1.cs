using System;
using System.Linq;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using System.Globalization;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using static System.Windows.Forms.VisualStyles.VisualStyleElement.Rebar;
using System.Data;
using System.Diagnostics;
using ClosedXML.Excel;
using System.Drawing.Printing;
using System.Drawing;
using DocumentFormat.OpenXml.Vml.Spreadsheet;
using System.Drawing.Drawing2D;
using System.Windows.Interop;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;
using DocumentFormat.OpenXml.Wordprocessing;
using System.ComponentModel;
using Microsoft.Office.Interop.Excel;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Collections;

namespace excel
{
    public partial class Form1 : Form
    {   
        private void ReleaseExcelResources()
        {
            if (xlWorkbook != null)
            {
                // Çalışma kitabını kapatın ve serbest bırakın
                xlWorkbook.Close(false);
                Marshal.ReleaseComObject(xlWorkbook);
                xlWorkbook = null;
            }

            if (xlApp != null)
            {
                // Excel uygulamasını kapatın ve serbest bırakın
                xlApp.Quit();
                Marshal.ReleaseComObject(xlApp);
                xlApp = null;
            }
        }
        private void Form1_Resize(object sender, EventArgs e)
        {
            int distanceFromLeft = dataGridView1.Location.X + dataGridView1.Width;
            button1.Location = new System.Drawing.Point(distanceFromLeft + 20, 10);
            button4.Location = new System.Drawing.Point(distanceFromLeft + 20, 50);
            button2.Location = new System.Drawing.Point(distanceFromLeft + 20, 90);
            button5.Location = new System.Drawing.Point(distanceFromLeft + 20, 130);
            button3.Location = new System.Drawing.Point(distanceFromLeft + 20, 170);
            progressBar.Location = new System.Drawing.Point(distanceFromLeft + 0, 200);

        }
        private Excel.Application xlApp;
        public Form1()
        {
            InitializeComponent();
            this.Resize += new EventHandler(Form1_Resize);
            DataGridView dataGridView1 = new DataGridView();
            DataGridView dataGridView2 = new DataGridView();
            DataGridView dataGridView3 = new DataGridView();
            dataGridView1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
            dataGridView1.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells;
            this.KeyPreview = true;
            this.KeyDown += new KeyEventHandler(Form1_KeyDown);
            //   xlApp = new Excel.Application();
            // DataGridView örneği oluşturuluyor ve formun üzerine ekleniyor (örnek verilerle)
            this.FormClosing += new FormClosingEventHandler(Form1_FormClosing);
        }
        private void eskisiparis()
        {
            dataGridView1.Rows.Clear();
            Excel.Workbook xlWorkbook = null;
            Excel.Application xlApp = new Excel.Application();
            try
            {
                OpenFileDialog openFileDialog1 = new OpenFileDialog
                {
                    InitialDirectory = @"C:\",
                    Title = "Excel Dosyası Seç",
                    CheckFileExists = true,
                    CheckPathExists = true,
                    DefaultExt = "xlsx",
                    Filter = "Excel Dosyaları (*.xlsx)|*.xlsx",
                    FilterIndex = 2,
                    RestoreDirectory = true,
                    ReadOnlyChecked = true,
                    ShowReadOnly = true

                };
                if (openFileDialog1.ShowDialog() == DialogResult.OK)
                {
                    List<string> excelHeaders = new List<string>();
                    ReleaseExcelResources();
                    string selectedFileName = openFileDialog1.FileName;
                    int i = 0;
                    xlWorkbook = xlApp.Workbooks.Open(selectedFileName);
                    Excel.Sheets sheets = xlWorkbook.Sheets;
                    Excel.Worksheet xlWorksheet = (Excel.Worksheet)sheets[1];
                    Excel.Range xlRange = xlWorksheet.UsedRange;
                    // Excel'den tüm başlıkları okuyun
                    for (int columnIndex = 1; columnIndex <= xlRange.Columns.Count; columnIndex++)
                    {
                        string header = xlWorksheet.Cells[1, columnIndex].Value != null ? xlWorksheet.Cells[1, columnIndex].Value.ToString().Trim() : string.Empty;
                        excelHeaders.Add(header);
                    }
                    // Beklenen başlıklar
                    string[] expectedHeaders = { "Tezgah Kodu", "Stok kod", "Malzeme Ad", "MEVCUT", "İstenen Miktar", "İstenen Tarih", "Firma Adı", "Kullanılan Malz." };
                    // Başlıkları kontrol edin
                    bool headersMatch = true;
                    foreach (var expectedHeader in expectedHeaders)
                    {
                        bool headerFound = false;
                        foreach (var excelHeader in excelHeaders)
                        {
                            if (excelHeader.Equals(expectedHeader, StringComparison.OrdinalIgnoreCase))
                            {
                                headerFound = true;
                                break;
                            }
                        }
                        if (!headerFound)
                        {
                            headersMatch = false;
                            MessageBox.Show($"Beklenen başlık bulunamadı: {expectedHeader}", "Başlık Eşleşmiyor", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            break;
                        }
                    }
                    // Eğer başlıklar eşleşiyorsa verileri getirin
                    if (headersMatch)
                    {
                        progressBar.Minimum = 0;
                        progressBar.Maximum = xlRange.Rows.Count - 1;
                        progressBar.Step = 1;
                        progressBar.Visible = true;

                        DataGridView dataGridView;
                        if (i == 0)
                        {
                            dataGridView = dataGridView1;
                        }
                        else if (i == 1)
                        {
                            dataGridView = dataGridView2;
                        }
                        else
                        {
                            dataGridView = dataGridView3;
                        }
                        // DataGridView'e sütunları ekle

                        Dictionary<string, double> istenenMiktarToplamlari = new Dictionary<string, double>();
                        HashSet<string> malzemeAdSet = new HashSet<string>();
                        for (int j = 2; j <= xlRange.Rows.Count; j++)
                        {
                            progressBar.PerformStep();

                            string siraNo = GetCellValue(xlWorksheet, j, excelHeaders, "Sıra No");
                            string tezgahKod = GetCellValue(xlWorksheet, j, excelHeaders, "No");
                            string firmaAdi = GetCellValue(xlWorksheet, j, excelHeaders, "Firma Adı");
                            string stokKod = GetCellValue(xlWorksheet, j, excelHeaders, "Stok kod");
                            string malzemeAd = GetCellValue(xlWorksheet, j, excelHeaders, "Malzeme Ad");
                            string isNo = GetCellValue(xlWorksheet, j, excelHeaders, "İş Emri No");
                            string mevcut = GetCellValue(xlWorksheet, j, excelHeaders, "MEVCUT");
                            string istenenMiktar = GetCellValue(xlWorksheet, j, excelHeaders, "İstenen Miktar");
                            string istenenTarih = GetCellValue(xlWorksheet, j, excelHeaders, "İstenen Tarih");
                            string kullanilanMalzeme = GetCellValue(xlWorksheet, j, excelHeaders, "Kullanılan Malz.");


                            if (!malzemeAdSet.Contains(malzemeAd))
                            {
                                malzemeAdSet.Add(malzemeAd);
                                if (istenenMiktarToplamlari.ContainsKey(malzemeAd))
                                {
                                    istenenMiktarToplamlari[malzemeAd] += double.Parse(istenenMiktar, CultureInfo.InvariantCulture);
                                }
                                else
                                {
                                    istenenMiktarToplamlari[malzemeAd] = double.Parse(istenenMiktar, CultureInfo.InvariantCulture);
                                }

                                dataGridView.Rows.Add(new string[] { siraNo, tezgahKod, firmaAdi, stokKod, malzemeAd, isNo, mevcut, istenenMiktar, istenenTarih, kullanilanMalzeme });
                            }
                            else
                            {
                                istenenMiktarToplamlari[malzemeAd] += double.Parse(istenenMiktar, CultureInfo.InvariantCulture);
                            }

                        }
                        ReleaseExcelResources();
                        // DataGridView'e yinelenenlerin istenen miktar toplamlarını yazdırma
                        int rowCount = dataGridView.Rows.Count;
                        foreach (var item in istenenMiktarToplamlari)
                        {
                            for (int rowIndex = 0; rowIndex < rowCount; rowIndex++)
                            {
                                if (dataGridView.Rows[rowIndex].Cells["MalzemeAd"].Value != null && dataGridView.Rows[rowIndex].Cells["MalzemeAd"].Value.ToString() == item.Key)
                                {
                                    dataGridView.Rows[rowIndex].Cells["IstenenMiktar"].Value = item.Value.ToString(CultureInfo.InvariantCulture);
                                }
                            }
                        }
                        this.Controls.Add(dataGridView);
                        i++;
                    }
                    else
                    {
                        MessageBox.Show("Başlıklar eşleşmiyor!", "Hata", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Hata oluştu: " + ex.Message, "Hata", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                ReleaseExcelResources();
            }
            if (xlWorkbook != null)
            {
                // Çalışma kitabını kapatın ve serbest bırakın
                xlWorkbook.Close(false);
                Marshal.ReleaseComObject(xlWorkbook);
                xlWorkbook = null;
            }
            if (xlApp != null)
            {
                // Excel uygulamasını kapatın ve serbest bırakın
                xlApp.Quit();
                Marshal.ReleaseComObject(xlApp);
                xlApp = null;
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            eskisiparis();
        }
     
        private string GetCellValue(Excel.Worksheet worksheet, int row, List<string> headers, string headerName)
        {
           ;
            int columnIndex = headers.IndexOf(headerName) + 1;
            if (columnIndex > 0)
            {
                if (headerName.Equals("Tezgah Kodu", StringComparison.OrdinalIgnoreCase) || headerName.Equals("Kullanılan Malz.", StringComparison.OrdinalIgnoreCase))
                {
                    return string.Empty;
                }

                var cellValue = worksheet.Cells[row, columnIndex].Value;
                if (cellValue != null)
                {
                    if (cellValue is double)
                    {
                        return Convert.ToString(cellValue, CultureInfo.InvariantCulture);
                    }
                    return cellValue.ToString().Trim();
                }
            }
            return string.Empty;
        }
        private void OpenAndProcessExcelFile()
        {
           

            Excel.Application xlApp = new Excel.Application();
            OpenFileDialog openFileDialog = new OpenFileDialog
            {
                InitialDirectory = @"C:\",
                Title = "Excel Dosyası Seç",
                CheckFileExists = true,
                CheckPathExists = true,
                DefaultExt = "xlsx",
                Filter = "Excel Dosyaları (*.xlsx)|*.xlsx",
                FilterIndex = 2,
                RestoreDirectory = true,
                ReadOnlyChecked = true,
                ShowReadOnly = true
            };

            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                string selectedFileName = openFileDialog.FileName;
                Excel.Workbook xlWorkbook = null;
                Excel._Worksheet xlWorksheet = null;
                Excel.Range xlRange = null;
                int totalRows = dataGridView1.Rows.Count; ;
                progressBar.Minimum = 0;
                progressBar.Maximum = totalRows;
                progressBar.Step = 1;
                progressBar.Value = 0;
                progressBar.Visible = true;
                try
                {

                    xlWorkbook = xlApp.Workbooks.Open(selectedFileName);
                    xlWorksheet = xlWorkbook.Sheets[1];
                    xlRange = xlWorksheet.UsedRange;

                    List<string> excelHeaders = new List<string>();
                    for (int columnIndex = 1; columnIndex <= xlRange.Columns.Count; columnIndex++)
                    {
                        string header = xlRange.Cells[1, columnIndex].Value?.ToString();
                        excelHeaders.Add(header);
                    }

                    for (int i = 2; i <= xlRange.Rows.Count; i++)
                    {

                        string malzemeAd = GetCellValue(xlWorksheet, i, excelHeaders, "Malzeme Ad");
                        string tezgahKod = GetCellValue(xlWorksheet, i, excelHeaders, "Tezgah Kodu");
                        string kullanilanMalzeme = GetCellValue(xlWorksheet, i, excelHeaders, "Kullanılan Malz.");

                        for (int rowIndex = 0; rowIndex < dataGridView1.Rows.Count; rowIndex++)
                        {

                            if (dataGridView1.Rows[rowIndex].Cells["MalzemeAd"].Value != null &&
                                dataGridView1.Rows[rowIndex].Cells["MalzemeAd"].Value.ToString() == malzemeAd)
                            {

                                dataGridView1.Rows[rowIndex].Cells["TezgahKodu"].Value = tezgahKod;
                                dataGridView1.Rows[rowIndex].Cells["KullanilanMalzeme"].Value = kullanilanMalzeme;
                                progressBar.PerformStep();


                            }

                        }

                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Hata oluştu: " + ex.Message, "Hata", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                finally
                {
                    // Excel nesnelerini serbest bırak
                    Marshal.ReleaseComObject(xlRange);
                    Marshal.ReleaseComObject(xlWorksheet);
                    xlWorkbook.Close();
                    Marshal.ReleaseComObject(xlWorkbook);
                    xlApp.Quit();
                    Marshal.ReleaseComObject(xlApp);
                }
            }
        }
        private void sirala()
        {
            var rows = dataGridView1.Rows.Cast<DataGridViewRow>();

            var sortedRows = rows.OrderBy(row =>
            {
                var tezgahKodValue = row.Cells["TezgahKodu"].Value?.ToString();
                if (string.IsNullOrEmpty(tezgahKodValue))
                {
                    return "zzzzzzzzz"; // Boş olanları listenin en sonuna eklemek için büyük bir değer kullanabilirsiniz
                }
                return tezgahKodValue;
            })
.ThenBy(row =>
{
    if (row.Cells["IstenenTarih"].Value != null && DateTime.TryParseExact(row.Cells["IstenenTarih"].Value.ToString().Trim('\''), "dd.MM.yyyy", CultureInfo.InvariantCulture, DateTimeStyles.None, out DateTime date))
    {
        return date;
    }
    else if (row.Cells["IstenenTarih"].Value != null && DateTime.TryParseExact(row.Cells["IstenenTarih"].Value.ToString().Trim('\''), "d.M.yyyy", CultureInfo.InvariantCulture, DateTimeStyles.None, out DateTime dateFallback))
    {
        return dateFallback;
    }
    return DateTime.MaxValue;
});
            List<object[]> temporaryRows = new List<object[]>();

            foreach (var sortedRow in sortedRows)
            {
                List<object> cellValues = new List<object>();
                foreach (DataGridViewCell cell in sortedRow.Cells)
                {
                    cellValues.Add(cell.Value);
                    progressBar.PerformStep();

                }
                temporaryRows.Add(cellValues.ToArray());
            }

            // DataGridView'i temizle
            dataGridView1.Rows.Clear();

            // DataGridView'e geçici listeden verileri ekleme
            foreach (var row in temporaryRows)
            {
                progressBar.PerformStep();

                dataGridView1.Rows.Add(row);
            }
            siraNo();
           
        }
        private void siraNo()
        {
            int siraNo = 1;
            string previousTezgahKod = string.Empty;
            foreach (DataGridViewRow row in dataGridView1.Rows)
            {
                string currentTezgahKod = row.Cells["TezgahKodu"].Value != null ? row.Cells["TezgahKodu"].Value.ToString() : string.Empty;

                if (currentTezgahKod != previousTezgahKod && currentTezgahKod != string.Empty)
                {
                    progressBar.PerformStep();

                    siraNo = 1;
                }

                row.Cells["SiraNo"].Value = siraNo;
                siraNo++;
                previousTezgahKod = currentTezgahKod;
              
            }

        }
        private string GetCellValue(Excel._Worksheet worksheet, int rowIndex, List<string> headers, string columnName)
        {
            int columnIndex = headers.IndexOf(columnName) + 1;
            if (columnIndex > 0)
            {
                Excel.Range targetCell = (Excel.Range)worksheet.Cells[rowIndex, columnIndex];
                if (targetCell != null && targetCell.Value2 != null)
                {
                    if (targetCell.Value2 is string)
                    {
                        return targetCell.Text;
                    }
                    else
                    {
                        return targetCell.Value2.ToString();
                    }
                }
            }
            return string.Empty;
        }
        private void button2_Click(object sender, EventArgs e)
        {
            OpenAndProcessExcelFile();
            sirala();
        }
        Excel.Workbook xlWorkbook;
        private void Form1_FormClosed_1(object sender, FormClosedEventArgs e)
        {
            if (xlWorkbook != null)
            {
                // Çalışma kitabını kapatın ve serbest bırakın
                xlWorkbook.Close(false);
                Marshal.ReleaseComObject(xlWorkbook);
                xlWorkbook = null;
            }
            if (xlApp != null)
            {
                // Excel uygulamasını kapatın ve serbest bırakın
                xlApp.Quit();
                Marshal.ReleaseComObject(xlApp);
                xlApp = null;
            }
        }
        private bool formLoaded = false;
        private void Form1_Load(object sender, EventArgs e)
        {
            formLoaded = true;
        }
        private void ExportToExcel()
        {
            ReleaseExcelResources();

            Excel.Application excelApp = new Excel.Application();
            Excel.Workbook excelWorkbook = excelApp.Workbooks.Add();
            Excel._Worksheet excelWorksheet = excelWorkbook.Sheets[1];
            excelWorksheet.Cells[1, 1] = "ÜRETİM PROGRAMI"; // İlk satırda başlık

            int rowIndex = 2; // Satır sayacı, sütun başlıklarının altına geçecek
            string previousTezgahKod = string.Empty; // Önceki TezgahKod'u takip etmek için bir değişken

            for (int i = 0; i < dataGridView1.Columns.Count; i++)
            {
                excelWorksheet.Cells[rowIndex, i + 1] = dataGridView1.Columns[i].HeaderText;
                Excel.Range cell = (Excel.Range)excelWorksheet.Cells[rowIndex, i + 1];
                cell.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                cell.Borders.Weight = Excel.XlBorderWeight.xlThin;
            }

            

            rowIndex++; // Verilerin altına geçmek için bir satır aşağı inin
            foreach (DataGridViewRow row in dataGridView1.Rows)
            {
                string currentTezgahKod = row.Cells["TezgahKodu"].Value != null ? row.Cells["TezgahKodu"].Value.ToString() : string.Empty;

                if (currentTezgahKod != previousTezgahKod && previousTezgahKod != string.Empty && currentTezgahKod != string.Empty)
                {
                    excelWorksheet.HPageBreaks.Add((Excel.Range)excelWorksheet.Cells[rowIndex, 1]);
                }

                for (int i = 0; i < dataGridView1.Columns.Count; i++)
                {
                    if (i == 6 || i == 7) // 7 ve 8. sütunları kontrol et
                    {
                        if (row.Cells[i].Value != null)
                        {
                            excelWorksheet.Cells[rowIndex, i + 1] = row.Cells[i].Value.ToString(); // ' işareti olmadan değeri ekle
                        }
                    }
                    else
                    {
                        if (row.Cells[i].Value != null)
                        {
                            excelWorksheet.Cells[rowIndex, i + 1] = "'" + row.Cells[i].Value; // Diğer durumlarda ' işaretiyle birlikte ekle
                        }
                    }
                    Excel.Range cell = (Excel.Range)excelWorksheet.Cells[rowIndex, i + 1];
                    cell.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                    cell.Borders.Weight = Excel.XlBorderWeight.xlThin;
                }
                previousTezgahKod = currentTezgahKod;
                rowIndex++;
            }

            int mergeCellCount = 10;
            Excel.Range startRange = excelWorksheet.Cells[1, 1];
            Excel.Range endRange = excelWorksheet.Cells[1, mergeCellCount];
            Excel.Range mergeRange = excelWorksheet.Range[startRange, endRange];
            mergeRange.Merge();
            mergeRange.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            mergeRange.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
            mergeRange.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Yellow);
            excelWorksheet.PageSetup.PrintTitleRows = "$1:$2";
            excelWorksheet.PageSetup.Orientation = Excel.XlPageOrientation.xlLandscape;

            for (int x = 1; x <= dataGridView1.Columns.Count; x++)
            {
                ((Excel.Range)excelWorksheet.Columns[x]).AutoFit();
            }

            Excel.Range boldRange = excelWorksheet.Range[excelWorksheet.Cells[1, 1], excelWorksheet.Cells[2, dataGridView1.Columns.Count]];
            boldRange.Font.Bold = true;

            for (int i = 1; i <= dataGridView1.Rows.Count; i++)
            {
                for (int j = 1; j <= 2; j++)
                {
                    Excel.Range cell = (Excel.Range)excelWorksheet.Cells[i + 2, j];
                    cell.Font.Bold = true;
                }
            }
            Excel.Range secondRow = (Excel.Range)excelWorksheet.Rows[2];
            secondRow.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
            secondRow.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            excelWorksheet.Rows["1"].RowHeight = 30;
            Excel.Range columnC = excelWorksheet.Columns["C:C"];
            columnC.ColumnWidth = 20;
            excelWorksheet.PageSetup.Zoom = 73;
            excelApp.Visible = true;
            Excel.Range fColumn = excelWorksheet.get_Range("G:G");
            fColumn.NumberFormat = "#,##0";
            excelWorksheet.PageSetup.CenterHorizontally = true;
            // G sütununun formatını sayı olarak ve binlik ayracıyla ayarla
            Excel.Range gColumn = excelWorksheet.get_Range("H:H");
            gColumn.NumberFormat = "#,##0";
            Marshal.ReleaseComObject(mergeRange);
            Marshal.ReleaseComObject(endRange);
            Marshal.ReleaseComObject(startRange);
            Marshal.ReleaseComObject(excelWorksheet);
            Marshal.ReleaseComObject(excelWorkbook);
            Marshal.ReleaseComObject(excelApp);
            GC.Collect();
            ReleaseExcelResources();



        }
        private bool excelExported = false;
        private void Form1_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.F8)
            {
                if (!excelExported)
                {
                    ExportToExcel();
                    excelExported = true; // Excel bir kez dışa aktarıldı
                }
                else
                {
                    excelExported = false; // Excel tekrar dışa aktarılmak üzere sıfırlandı
                }
            }
            if (e.KeyCode == Keys.F5)
            {
                if (formLoaded)
                {
                    dataGridView1.AutoResizeColumns();
                    dataGridView1.AutoResizeRows();
                }
            }
        }
        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            ReleaseExcelResources();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            ExportToExcel();
        }

        private void iegetir()

        {
            OpenFileDialog openFileDialog1 = new OpenFileDialog
            {
                Title = "Excel Dosyası Seçin",
                CheckFileExists = true,
                CheckPathExists = true,
                DefaultExt = "xlsx",
                Filter = "Excel Dosyaları|*.xls;*.xlsx;*.xlsm",
                FilterIndex = 2,
                RestoreDirectory = true,
                ReadOnlyChecked = true,
                ShowReadOnly = true
            };

            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                string filePath = openFileDialog1.FileName;
                Microsoft.Office.Interop.Excel.Application excel = new Microsoft.Office.Interop.Excel.Application();
                Microsoft.Office.Interop.Excel.Workbook wb = excel.Workbooks.Open(filePath);
                Microsoft.Office.Interop.Excel.Worksheet ws = (Microsoft.Office.Interop.Excel.Worksheet)wb.Sheets[1];

                Microsoft.Office.Interop.Excel.Range range = ws.UsedRange;
                bool foundMlzKod = false;
                bool foundIsEmriNo = false;
                bool foundMamulTanimi = false;
                int mlzKodColumnIndex = 0;
                int isEmriNoColumnIndex = 0;
                int mamulTanimiColumnIndex = 0;
                int totalRows = range.Rows.Count - 1;

                progressBar.Minimum = 0;
                progressBar.Maximum = totalRows;
                progressBar.Step = 1;
                progressBar.Value = 0;
                progressBar.Visible = true;
                try
                {
                    for (int i = 1; i <= range.Columns.Count; i++)
                    {
                        if (range.Cells[1, i].Value2 != null && range.Cells[1, i].Value2.ToString() == "MlzKodu")
                        {
                            foundMlzKod = true;
                            mlzKodColumnIndex = i;
                        }
                        if (range.Cells[1, i].Value2 != null && range.Cells[1, i].Value2.ToString() == "İş Emri No")
                        {
                            foundIsEmriNo = true;
                            isEmriNoColumnIndex = i;
                        }
                        if (range.Cells[1, i].Value2 != null && range.Cells[1, i].Value2.ToString() == "Mamül Tanımı")
                        {
                            foundMamulTanimi = true;
                            mamulTanimiColumnIndex = i;
                        }
                    }

                    if (foundMlzKod && foundIsEmriNo && foundMamulTanimi)
                    {

                        // Dtgr1'in olduğunu varsayarak;
                        for (int row = 2; row <= range.Rows.Count; row++)
                        {
                            progressBar.PerformStep();
                            string mlzKod = Convert.ToString((range.Cells[row, mlzKodColumnIndex] as Microsoft.Office.Interop.Excel.Range).Value2);
                            string isEmriNo = Convert.ToString((range.Cells[row, isEmriNoColumnIndex] as Microsoft.Office.Interop.Excel.Range).Value2);
                            string mamulTanimi = Convert.ToString((range.Cells[row, mamulTanimiColumnIndex] as Microsoft.Office.Interop.Excel.Range).Value2);
                            bool foundMatch = false;
                            foreach (DataGridViewRow dataGridViewRow in dataGridView1.Rows)
                            {
                                if (dataGridViewRow.Cells["Stokod"].Value != null && dataGridViewRow.Cells["Stokod"].Value.ToString() == mlzKod)
                                {

                                    foundMatch = true;
                                    dataGridViewRow.Cells["IsEmriNo"].Value = isEmriNo;
                                    dataGridViewRow.Cells["MalzemeAd"].Value = mamulTanimi;

                                }
                            }
                            progressBar.PerformStep();

                            if (!foundMatch)
                            {
                                progressBar.PerformStep();

                                DataGridViewRow newRow = (DataGridViewRow)dataGridView1.Rows[0].Clone();
                                newRow.Cells[3].Value = mlzKod;
                                newRow.Cells[4].Value = mamulTanimi;
                                newRow.Cells[5].Value = isEmriNo;
                                dataGridView1.Rows.Add(newRow);

                            }
                        }

                        MessageBox.Show("İşlem tamamlandı!");
                    }
                    else
                    {
                        string errorMessage = "Aşağıdaki sütun(lar) bulunamadı: ";
                        if (!foundMlzKod) errorMessage += "MlzKod, ";
                        if (!foundIsEmriNo) errorMessage += "İş Emri No, ";
                        if (!foundMamulTanimi) errorMessage += "Mamül Tanımı, ";
                        MessageBox.Show(errorMessage.TrimEnd(' ', ','));
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Bir hata oluştu: " + ex.Message);
                }
                finally
                {
                    Marshal.ReleaseComObject(range);
                    Marshal.ReleaseComObject(ws);
                    wb.Close(false);
                    Marshal.ReleaseComObject(wb);
                    excel.Quit();
                    Marshal.ReleaseComObject(excel);

                }
            }
        }
        private void stokgetir()
        {


           
            OpenFileDialog openFileDialog1 = new OpenFileDialog
            {
                Title = "Excel Dosyası Seçin",
                CheckFileExists = true,
                CheckPathExists = true,
                DefaultExt = "xlsx",
                Filter = "Excel Dosyaları|*.xls;*.xlsx;*.xlsm",
                FilterIndex = 2,
                RestoreDirectory = true,
                ReadOnlyChecked = true,
                ShowReadOnly = true
            };

            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                string filePath = openFileDialog1.FileName;
                Microsoft.Office.Interop.Excel.Application excel = new Microsoft.Office.Interop.Excel.Application();
                Microsoft.Office.Interop.Excel.Workbook wb = excel.Workbooks.Open(filePath);
                Microsoft.Office.Interop.Excel.Worksheet ws = (Microsoft.Office.Interop.Excel.Worksheet)wb.Sheets[1];

                Microsoft.Office.Interop.Excel.Range range = ws.UsedRange;
                bool foundMlzKod = false;
                bool foundstomevcut = false;
                bool foundfirmad = false;
                int mlzKodColumnIndex = 0;
                int stokmevcutColumnIndex = 0;

                int firmaAdColumnIndex = 0;
                int totalRows = range.Rows.Count - 1;

                progressBar.Minimum = 0;
                progressBar.Maximum = totalRows;
                progressBar.Step = 1;
                progressBar.Value = 0;
                progressBar.Visible = true;
                try
                {
                    

                    for (int i = 1; i <= range.Columns.Count; i++)
                    {

                        if (range.Cells[1, i].Value2 != null && range.Cells[1, i].Value2.ToString() == "Stok Kod")
                        {
                            foundMlzKod = true;
                            mlzKodColumnIndex = i;
                        }
                        if (range.Cells[1, i].Value2 != null && range.Cells[1, i].Value2.ToString() == "Mevcut")
                        {
                            foundstomevcut = true;
                            stokmevcutColumnIndex = i;
                        }
                        if (range.Cells[1, i].Value2 != null && range.Cells[1, i].Value2.ToString() == "Firma Adı")
                        {
                            foundfirmad = true;
                            firmaAdColumnIndex = i;
                        }
                    }

                    if (foundMlzKod && foundstomevcut && foundfirmad)
                    {

                        
                        for (int row = 2; row <= range.Rows.Count; row++)
                        {
                            string mlzKod = Convert.ToString((range.Cells[row, mlzKodColumnIndex] as Microsoft.Office.Interop.Excel.Range).Value2);
                            string stokmvct = Convert.ToString((range.Cells[row, stokmevcutColumnIndex] as Microsoft.Office.Interop.Excel.Range).Value2);
                            string firmad = Convert.ToString((range.Cells[row, firmaAdColumnIndex] as Microsoft.Office.Interop.Excel.Range).Value2);
                            bool foundMatch = false;
                            foreach (DataGridViewRow dataGridViewRow in dataGridView1.Rows)
                            {

                                if (dataGridViewRow.Cells["FirmaAdi"].Value == null || dataGridViewRow.Cells["FirmaAdi"].Value.ToString() == "")
                                {

                                    if (dataGridViewRow.Cells["Stokod"].Value != null && dataGridViewRow.Cells["Stokod"].Value.ToString() == mlzKod)
                                    {
                                        foundMatch = true;
                                        dataGridViewRow.Cells["Mevcut"].Value = stokmvct;
                                        dataGridViewRow.Cells["FirmaAdi"].Value = firmad;
                                        progressBar.PerformStep();

                                    }
                                }
                            }
                            if (!foundMatch)
                            {
                                        progressBar.PerformStep();


                            }
                        }
                        MessageBox.Show("İşlem tamamlandı!");

                    }

                }
                catch (Exception ex)
                {
                    MessageBox.Show("Bir hata oluştu: " + ex.Message);
                }
                finally
                {
                    Marshal.ReleaseComObject(range);
                    Marshal.ReleaseComObject(ws);
                    wb.Close(false);
                    Marshal.ReleaseComObject(wb);
                    excel.Quit();
                    Marshal.ReleaseComObject(excel);
                }
                
            }
        }
        private void button4_Click(object sender, EventArgs e)
        {
            iegetir();
        }

        private void button5_Click(object sender, EventArgs e)
        {
            stokgetir();
        }
    }
}