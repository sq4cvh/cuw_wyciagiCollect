using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Spire.Xls;
using static Spire.Xls.Core.Spreadsheet.XlsWorksheet;

namespace wyciagiCollect
{
    public partial class frmMain : Form
    {
        public frmMain()
        {
            InitializeComponent();
        }

        private string dateToDateWithLine(string input) { 
            return input.Substring(0, 4) + "-" + input.Substring(4, 2) + "-" + input.Substring(6, 2);
        }

        private void getDataFromFile(string fileName)
        {
            string[] allFileLines = File.ReadAllLines(fileName,Encoding.Default);
            int lineCount = allFileLines.Length;
            string lastLine = allFileLines[allFileLines.Length-1];
            string firstLine = allFileLines[0];
            if (firstLine.Split(',')[0] != "01" || lastLine.Split(',')[0] != "03")
            {
                MessageBox.Show("BŁĄD\r\nNiepoprawny format pliku", "Błąd", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            else
            {
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];
                sheet.PageSetup.TopMargin = 0.3;
                sheet.PageSetup.BottomMargin = 1;
                sheet.PageSetup.LeftMargin = 0.5;
                sheet.PageSetup.RightMargin = 0.5;
                sheet.PageSetup.HeaderMarginInch = 0.1;
                sheet.PageSetup.FooterMarginInch = 0.5;

                sheet.Name = "Wyciąg bankowy " + dateToDateWithLine(firstLine.Split(',')[3]);
                //Writes hello world to A1
                sheet.Range["A1"].Text = "Plik z Banku: " + Path.GetFileName(fileName)+" data księgowania: "+ dateToDateWithLine(firstLine.Split(',')[3]);
                sheet.DefaultRowHeight = 20;
                string newRange = "A2";
                int currentRowNumber = 3;
                newRange = "A" + currentRowNumber.ToString();
                sheet.Range[newRange].BorderAround(LineStyleType.Thin, ExcelColors.Black);
                sheet.Range[newRange].VerticalAlignment = VerticalAlignType.Center;
                sheet.Range[newRange].HorizontalAlignment = HorizontalAlignType.Center;
                sheet.Range[newRange].Text = "Lp";
                newRange = "B" + currentRowNumber.ToString();
                sheet.Range[newRange].BorderAround(LineStyleType.Thin, ExcelColors.Black);
                sheet.Range[newRange].VerticalAlignment = VerticalAlignType.Center;
                sheet.Range[newRange].HorizontalAlignment = HorizontalAlignType.Center;
                sheet.Range[newRange].Text = "Nr collect.";
                newRange = "C" + currentRowNumber.ToString();
                sheet.Range[newRange].BorderAround(LineStyleType.Thin, ExcelColors.Black);
                sheet.Range[newRange].VerticalAlignment = VerticalAlignType.Center;
                sheet.Range[newRange].HorizontalAlignment = HorizontalAlignType.Center;
                sheet.Range[newRange].Text = "Wpłacający";
                newRange = "D" + currentRowNumber.ToString();
                sheet.Range[newRange].BorderAround(LineStyleType.Thin, ExcelColors.Black);
                sheet.Range[newRange].VerticalAlignment = VerticalAlignType.Center;
                sheet.Range[newRange].HorizontalAlignment = HorizontalAlignType.Center;
                sheet.Range[newRange].Text = "Tytuł wpłaty";
                newRange = "E" + currentRowNumber.ToString();
                sheet.Range[newRange].BorderAround(LineStyleType.Thin, ExcelColors.Black);
                sheet.Range[newRange].VerticalAlignment = VerticalAlignType.Center;
                sheet.Range[newRange].HorizontalAlignment = HorizontalAlignType.Center;
                sheet.Range[newRange].Text = "Kwota";
                for (int i = 1; i < allFileLines.Length-1; i++)
                {

                
                    currentRowNumber = i+3;
                    newRange = "A" + currentRowNumber.ToString();
                    sheet.Range[newRange].ColumnWidth = 4;
                    sheet.Range[newRange].BorderAround(LineStyleType.Thin, ExcelColors.Black);
                    sheet.Range[newRange].VerticalAlignment = VerticalAlignType.Center;
                    sheet.Range[newRange].HorizontalAlignment = HorizontalAlignType.Center;
                    sheet.Range[newRange].Style.Font.Size = 8;
                    sheet.Range[newRange].Text = (currentRowNumber-3).ToString()+"."; //Lp
                    
                    newRange = "B" + currentRowNumber.ToString();
                    sheet.Range[newRange].ColumnWidth = 10;
                    sheet.Range[newRange].BorderAround(LineStyleType.Thin, ExcelColors.Black);
                    sheet.Range[newRange].VerticalAlignment = VerticalAlignType.Center;
                    sheet.Range[newRange].HorizontalAlignment = HorizontalAlignType.Center;
                    sheet.Range[newRange].Style.Font.Size = 16;
                    sheet.Range[newRange].Value = allFileLines[i].Split(',')[12].Substring(15, 11); //numer ewidencyjny (końcówka NRB)

                    newRange = "C" + currentRowNumber.ToString();
                    sheet.Range[newRange].ColumnWidth = 20;
                    sheet.Range[newRange].BorderAround(LineStyleType.Thin, ExcelColors.Black);
                    sheet.Range[newRange].VerticalAlignment = VerticalAlignType.Center;
                    sheet.Range[newRange].HorizontalAlignment = HorizontalAlignType.Left;
                    sheet.Range[newRange].Style.WrapText = true;
                    sheet.Range[newRange].Style.Font.Size = 8;
                    sheet.Range[newRange].Value = allFileLines[i].Split(',')[6]; //wpłacający

                    newRange = "D" + currentRowNumber.ToString();
                    sheet.Range[newRange].ColumnWidth = 40;
                    sheet.Range[newRange].BorderAround(LineStyleType.Thin, ExcelColors.Black);
                    sheet.Range[newRange].VerticalAlignment = VerticalAlignType.Center;
                    sheet.Range[newRange].HorizontalAlignment = HorizontalAlignType.Left;
                    sheet.Range[newRange].Style.WrapText = true;
                    sheet.Range[newRange].Style.Font.Size = 8;
                    sheet.Range[newRange].Value = allFileLines[i].Split(',')[14]+ allFileLines[i].Split(',')[15]+ allFileLines[i].Split(',')[16]+ allFileLines[i].Split(',')[17]; //tytuł wpłaty

                    newRange = "E" + currentRowNumber.ToString();
                    sheet.Range[newRange].ColumnWidth = 13;
                    sheet.Range[newRange].Style.NumberFormat = "0.00";
                    sheet.Range[newRange].BorderAround(LineStyleType.Thin, ExcelColors.Black);
                    sheet.Range[newRange].VerticalAlignment = VerticalAlignType.Center;
                    sheet.Range[newRange].HorizontalAlignment = HorizontalAlignType.Right;
                    sheet.Range[newRange].Style.Font.Size = 16;
                    sheet.Range[newRange].Value = allFileLines[i].Split(',')[3].Replace('.', ','); //kwota wpłaty
                }
                currentRowNumber++;
                newRange = "D" + currentRowNumber.ToString();
                sheet.Range[newRange].Text = "Łączna kwota wpłat w pliku:";
                sheet.Range[newRange].Style.Font.Size = 16;
                sheet.Range[newRange].Style.Font.Color = System.Drawing.Color.Red;
                newRange = "E" + currentRowNumber.ToString();
                sheet.Range[newRange].Style.NumberFormat = "0.00";
                sheet.Range[newRange].Formula= "=SUM(E4:E"+ (currentRowNumber-1).ToString()+")";
                sheet.Range[newRange].Style.Font.Size = 16;
                sheet.Range[newRange].Style.Font.Color = System.Drawing.Color.Red;
                
                string fileNameToSave = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) + @"\Bank_Pliki\wyciag_Collect_z_dnia_" + dateToDateWithLine(firstLine.Split(',')[3]) + ".xlsx";
                workbook.Worksheets[1].Remove();
                workbook.Worksheets[1].Remove();
                workbook.SaveToFile(fileNameToSave,ExcelVersion.Version2013);
                //workbook.PrintDocument.Print();

                try
                {
                    System.Diagnostics.Process.Start(workbook.FileName);
                }
                catch { }
            }
        }

        private void frmMain_Load(object sender, EventArgs e)
        {
            label1.Text= "Przeciągnij myszką i upuść \r\nw tym miejscu plik wynikowy \r\nz Banku";
            toolStripStatusLabel1.Text = "© 2020 CUW Suwałki v." + Assembly.GetExecutingAssembly().GetName().Version.ToString();
        }

        private void panel1_DragEnter(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop)) e.Effect = DragDropEffects.Move;
        }

        private void panel1_DragDrop(object sender, DragEventArgs e)
        {
            foreach (string path in (string[])e.Data.GetData(DataFormats.FileDrop))
            {
                if (Path.GetExtension(path) != ".csv")
                {
                    MessageBox.Show("BŁĄD\r\nNiepoprawny format pliku\r\nDopuszczone tylko pliki z rozszerzeniem CSV", "Błąd", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                else {
                    try
                    {
                        getDataFromFile(path);
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("BŁĄD\r\npodczas przetwarzania pliku z Banku \r\n"+ex.Message, "Błąd", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
            }
        }
    }
}
