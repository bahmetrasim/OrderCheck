using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data.OleDb;
using OfficeOpenXml;
using System.IO;
using System.Globalization;
using System.Xml;
using System.Text.RegularExpressions;
using OfficeOpenXml.Table;
using OfficeOpenXml.Utils;
using OfficeOpenXml.Style;
using System.Reflection;

namespace Siparis_Kontrol
{
    class Pivot
    {
        public Pivot() { }

        public double Pivotcalc(string pathins, string sheet, string pathout, bool allorders, bool nodate)
        {
            FileInfo file = new FileInfo(pathins);
            FileInfo output = new FileInfo(pathout);
            if (output.Exists)
            {
                output.Delete();
                output = new FileInfo(pathout);
            }
            using (ExcelPackage OrderCheck = new ExcelPackage(file))
            {
                ExcelWorksheet ws = OrderCheck.Workbook.Worksheets[sheet];
                ws.InsertColumn(8, 1);
                ws.InsertColumn(10, 1);
                ws.Cells[1, 8].Value = "KGrup";
                ws.Cells[1, 10].Value = "EGrup";

                using (var rng = ws.Cells[1, 8, 1, 10])
                {
                    //rng.Style.Font.Bold = true;
                    rng.Style.Font.SetFromFont(new Font("Tahoma", 8, FontStyle.Bold));
                    rng.Style.Font.Color.SetColor(Color.Black);
                    rng.Style.WrapText = false;
                    rng.Style.VerticalAlignment = ExcelVerticalAlignment.Bottom;
                    rng.Style.HorizontalAlignment = ExcelHorizontalAlignment.General;
                    rng.Style.Fill.PatternType = ExcelFillStyle.Solid;
                    rng.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(245, 245, 245));
                };
                int ColCnt = ws.Dimension.End.Column;
                int RowCnt = ws.Dimension.End.Row;
                for (int i = 1; i < RowCnt + 1; i++)
                {
                    if (ws.Cells[i, 1].Value == null) { break; }

                    else if (ws.Cells[i, 1].Value.ToString() == "Toplam" || ws.Cells[i, 2].Value.ToString() == "Toplam")
                    {
                        ws.DeleteRow(i);
                        i--;
                    }
                }
                RowCnt = ws.Dimension.End.Row;
                for (int i = 2; i < RowCnt + 1; i++)
                {
                    if (double.Parse(ws.Cells[i, 7].Value.ToString()) < 0.23)
                    {
                        ws.Cells[i, 8].Value = "Group1";
                    }
                    else if (double.Parse(ws.Cells[i, 7].Value.ToString()) <= 0.30)
                    {
                        ws.Cells[i, 8].Value = "Group2";
                    }
                    else if (double.Parse(ws.Cells[i, 7].Value.ToString()) < 0.45)
                    {
                        ws.Cells[i, 8].Value = "Group3";
                    }
                    else if (double.Parse(ws.Cells[i, 7].Value.ToString()) <= 0.70)
                    {
                        ws.Cells[i, 8].Value = "Group4";
                    }
                    else if (double.Parse(ws.Cells[i, 7].Value.ToString()) < 1)
                    {
                        ws.Cells[i, 8].Value = "Group5";
                    }
                    else if (double.Parse(ws.Cells[i, 7].Value.ToString()) < 1.5)
                    {
                        ws.Cells[i, 8].Value = "Group6";
                    }
                    else
                    {
                        ws.Cells[i, 8].Value = "Group7";
                    }
                    if (double.Parse(ws.Cells[i, 9].Value.ToString()) < 1031)
                    {
                        ws.Cells[i, 10].Value = "Group1";
                    }
                    else if (double.Parse(ws.Cells[i, 9].Value.ToString()) < 1271)
                    {
                        ws.Cells[i, 10].Value = "Group2";
                    }
                    else
                    {
                        ws.Cells[i, 10].Value = "Group3";
                    }
                }
                for (int i = 2; i < RowCnt + 1; i++)
                {
                    if (ws.Cells[i, 15].Value.ToString() == "-")
                    {
                        ws.Cells[i, 15].Value = "Terminsiz";
                    }
                    else
                    {
                        long dateNum = long.Parse(ws.Cells[i, 15].Value.ToString());
                        DateTime result = DateTime.FromOADate(dateNum);
                        ws.Cells[i, 15].Style.Numberformat.Format = "dd.MM.yyyy";
                        ws.Cells[i, 15].Value = result;
                    }
                }
                // Checked Şartları
                DateTime dt = DateTime.Now;
                DateTime value = dt;
                var lastDayOfMonth = dt.AddDays(1 - dt.Day).AddMonths(1).AddDays(-1).Date;

                if (allorders == true && nodate == true) // Tüm Siparişler ve Terminsizleri Hariç Tut
                {
                    for (int i = 2; i < RowCnt + 1; i++)
                    {
                        if (ws.Cells[i, 15].Value == null) { break; }

                        if (ws.Cells[i, 15].Value.ToString() == "Terminsiz")
                        {
                            ws.DeleteRow(i);
                            i--;
                        }
                    }
                }
                else if (allorders == false && nodate == false) // Backlog ve Terminsizler
                {

                    for (int i = 2; i < RowCnt + 1; i++)
                    {
                        if (ws.Cells[i, 15].Value == null) { break; }
                        if (ws.Cells[i, 15].Value.ToString() != "Terminsiz")
                        {
                            value = DateTime.Parse(ws.Cells[i, 15].Value.ToString());
                        }
                        else
                        {
                            value = DateTime.Parse("01.01.2000");
                        }
                        if (value > lastDayOfMonth)
                        {
                            ws.DeleteRow(i);
                            i--;
                        }
                    }
                }
                else if (allorders == false && nodate == true) // Backlog ve Terminsizleri Hariç Tut
                {
                    for (int i = 2; i < RowCnt + 1; i++)
                    {
                        if (ws.Cells[i, 15].Value == null) { break; }
                        if (ws.Cells[i, 15].Value.ToString() != "Terminsiz")
                        {
                           value = DateTime.Parse(ws.Cells[i, 15].Value.ToString());
                        }
                        else
                        {
                            value = DateTime.Parse("01.01.2999");
                        }
                        if (ws.Cells[i, 15].Value.ToString() == "Terminsiz" || value > lastDayOfMonth)
                        {
                            ws.DeleteRow(i);
                            i--;
                        }
                    }
                }
                // Saveas
                OrderCheck.SaveAs(output);
                Table tablo = new Table();
                double nethours = tablo.Tablecalc(pathout, sheet, RowCnt, ColCnt, ws.Cells[1, 8].Value.ToString(), ws.Cells[1, 10].Value.ToString());
                return nethours;
            }
        }
    }
}

