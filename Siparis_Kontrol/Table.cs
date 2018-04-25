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
using OfficeOpenXml.Table.PivotTable;

namespace Siparis_Kontrol
{
    class Table
    {
        public Table() { }
        public double Tablecalcqlik (string pathins, string sheet, int RwCnt, int ClmnCnt, string Rwfld, string Clmnfld)
        {
            FileInfo file = new FileInfo(pathins);
            double nethours = 0;

            using (ExcelPackage pck = new ExcelPackage(file))
            {
                ExcelWorksheet ws = pck.Workbook.Worksheets[sheet];
                var wsPivot = pck.Workbook.Worksheets.Add("PivotTest");
                var Pvt = wsPivot.PivotTables.Add(wsPivot.Cells[3, 1], ws.Cells[1, 1, RwCnt, ClmnCnt], "Test");
                Pvt.RowFields.Add(Pvt.Fields[Rwfld]);
                Pvt.RowFields[Pvt.RowFields.Count - 1].Sort = eSortType.Ascending;
                Pvt.ColumnFields.Add(Pvt.Fields[Clmnfld]);
                Pvt.ColumnFields[Pvt.ColumnFields.Count - 1].Sort = eSortType.Ascending;
                var values = Pvt.DataFields.Add(Pvt.Fields["Kalan Miktar"]);
                values.Function = DataFieldFunctions.Sum;
                values.Format = "#.##0";
                var filter = Pvt.Fields["Onaylanan Sevk Tarihi"];
                Pvt.PageFields.Add(filter);

                wsPivot.Cells[17, 1].Value = "Hız - Referans";
                wsPivot.Cells[26, 1].Value = "Kalınlık - Referans";
                wsPivot.Cells[35, 1].Value = "En - Referans";
                wsPivot.Cells[17, 7].Value = "Net Üretim Zamanı";

                wsPivot.Cells[18, 1].Value = "Group1";
                wsPivot.Cells[27, 1].Value = "Group1";
                wsPivot.Cells[36, 1].Value = "Group1";
                wsPivot.Cells[18, 7].Value = "Group1";

                wsPivot.Cells[19, 1].Value = "Group2";
                wsPivot.Cells[28, 1].Value = "Group2";
                wsPivot.Cells[37, 1].Value = "Group2";
                wsPivot.Cells[19, 7].Value = "Group2";

                wsPivot.Cells[20, 1].Value = "Group3";
                wsPivot.Cells[29, 1].Value = "Group3";
                wsPivot.Cells[38, 1].Value = "Group3";
                wsPivot.Cells[20, 7].Value = "Group3";

                wsPivot.Cells[21, 1].Value = "Group4";
                wsPivot.Cells[30, 1].Value = "Group4";
                wsPivot.Cells[39, 1].Value = "Group4";
                wsPivot.Cells[21, 7].Value = "Group4";

                wsPivot.Cells[22, 1].Value = "Group5";
                wsPivot.Cells[31, 1].Value = "Group5";
                wsPivot.Cells[40, 1].Value = "Group5";
                wsPivot.Cells[22, 7].Value = "Group5";

                wsPivot.Cells[23, 1].Value = "Group6";
                wsPivot.Cells[32, 1].Value = "Group6";
                wsPivot.Cells[41, 1].Value = "Group6";
                wsPivot.Cells[23, 7].Value = "Group6";

                wsPivot.Cells[24, 1].Value = "Group7";
                wsPivot.Cells[33, 1].Value = "Group7";
                wsPivot.Cells[42, 1].Value = "Group7";
                wsPivot.Cells[24, 7].Value = "Group7";

                wsPivot.Cells[17, 2].Value = "Group1";
                wsPivot.Cells[17, 3].Value = "Group2";
                wsPivot.Cells[17, 4].Value = "Group3";

                wsPivot.Cells[26, 2].Value = "Group1";
                wsPivot.Cells[26, 3].Value = "Group2";
                wsPivot.Cells[26, 4].Value = "Group3";

                wsPivot.Cells[35, 2].Value = "Group1";
                wsPivot.Cells[35, 3].Value = "Group2";
                wsPivot.Cells[35, 4].Value = "Group3";

                wsPivot.Cells[17, 8].Value = "Group1";
                wsPivot.Cells[17, 9].Value = "Group2";
                wsPivot.Cells[17, 10].Value = "Group3";

                // hız
                wsPivot.Cells[18, 2, 20, 2].Value = 90;
                wsPivot.Cells[21, 2, 21, 4].Value = 90;
                wsPivot.Cells[18, 3, 20, 3].Value = 85;
                wsPivot.Cells[19, 4, 20, 4].Value = 85;
                wsPivot.Cells[18, 4].Value = 80;
                wsPivot.Cells[22, 2].Value = 80;
                wsPivot.Cells[22, 3].Value = 70;
                wsPivot.Cells[22, 4].Value = 70;
                wsPivot.Cells[23, 2, 23, 4].Value = 50;
                wsPivot.Cells[24, 2, 24, 4].Value = 40;

                // Kalınlık
                wsPivot.Cells[27, 2].Value = 0.21;
                wsPivot.Cells[27, 3].Value = 0.20;
                wsPivot.Cells[27, 4].Value = 0.20;
                wsPivot.Cells[28, 2].Value = 0.23;
                wsPivot.Cells[28, 3].Value = 0.25;
                wsPivot.Cells[28, 4].Value = 0.27;
                wsPivot.Cells[29, 2, 29, 4].Value = 0.33;
                wsPivot.Cells[30, 2, 30, 4].Value = 0.50;
                wsPivot.Cells[31, 2, 31, 4].Value = 0.70;
                wsPivot.Cells[32, 2, 32, 4].Value = 0.88;
                wsPivot.Cells[33, 2, 33, 4].Value = 1.6;

                // EN
                wsPivot.Cells[36, 2, 42, 2].Value = 1000;
                wsPivot.Cells[36, 3, 42, 3].Value = 1250;
                wsPivot.Cells[36, 4, 42, 4].Value = 1500;

                // formuller
                // group1
                wsPivot.Cells[18, 8].Formula = "B5*1000/2.71/B27/(B36/1000)/B18/60";
                wsPivot.Cells[19, 8].Formula = "B6*1000/2.71/B28/(B37/1000)/B19/60";
                wsPivot.Cells[20, 8].Formula = "B7*1000/2.71/B29/(B38/1000)/B20/60";
                wsPivot.Cells[21, 8].Formula = "B8*1000/2.71/B30/(B39/1000)/B21/60";
                wsPivot.Cells[22, 8].Formula = "B9*1000/2.71/B31/(B40/1000)/B22/60";
                wsPivot.Cells[23, 8].Formula = "B10*1000/2.71/B32/(B41/1000)/B23/60";
                wsPivot.Cells[24, 8].Formula = "B11*1000/2.71/B33/(B42/1000)/B24/60";

                // group2
                wsPivot.Cells[18, 9].Formula = "C5*1000/2.71/C27/(C36/1000)/C18/60";
                wsPivot.Cells[19, 9].Formula = "C6*1000/2.71/C28/(C37/1000)/C19/60";
                wsPivot.Cells[20, 9].Formula = "C7*1000/2.71/C29/(C38/1000)/C20/60";
                wsPivot.Cells[21, 9].Formula = "C8*1000/2.71/C30/(C39/1000)/C21/60";
                wsPivot.Cells[22, 9].Formula = "C9*1000/2.71/C31/(C40/1000)/C22/60";
                wsPivot.Cells[23, 9].Formula = "C10*1000/2.71/C32/(C41/1000)/C23/60";
                wsPivot.Cells[24, 9].Formula = "C11*1000/2.71/C33/(C42/1000)/C24/60";

                // group3

                wsPivot.Cells[18, 10].Formula = "D5*1000/2.71/D27/(D36/1000)/D18/60";
                wsPivot.Cells[19, 10].Formula = "D6*1000/2.71/D28/(D37/1000)/D19/60";
                wsPivot.Cells[20, 10].Formula = "D7*1000/2.71/D29/(D38/1000)/D20/60";
                wsPivot.Cells[21, 10].Formula = "D8*1000/2.71/D30/(D39/1000)/D21/60";
                wsPivot.Cells[22, 10].Formula = "D9*1000/2.71/D31/(D40/1000)/D22/60";
                wsPivot.Cells[23, 10].Formula = "D10*1000/2.71/D32/(D41/1000)/D23/60";
                wsPivot.Cells[24, 10].Formula = "D11*1000/2.71/D33/(D42/1000)/D24/60";
                wsPivot.Cells[25, 11].Style.Numberformat.Format = "#.##0";
                // toplamlar
                wsPivot.Cells["H25"].Formula = "SUM(H18:H24)";
                wsPivot.Cells["I25"].Formula = "SUM(I18:I24)";
                wsPivot.Cells["J25"].Formula = "SUM(J18:J24)";
                wsPivot.Cells["K25"].Formula = "SUM(H25:J25)";
                pck.Workbook.Calculate();
                
                //wsPivot.Cells[25,11].Value = wsPivot.Cells[25,11].Value;
                //foreach (var cell in wsPivot.Cells.Where(cell => cell.Formula != null))
                //{
                //    cell.Value = cell.Value;
                //}

                wsPivot.Cells[25, 11].Style.Font.SetFromFont(new Font("Tahoma", 10, FontStyle.Bold));
                wsPivot.Cells[25, 11].Style.Font.Color.SetColor(Color.Black);
                wsPivot.Cells[25, 11].Style.WrapText = false;
                wsPivot.Cells[25, 11].Style.VerticalAlignment = ExcelVerticalAlignment.Bottom;
                wsPivot.Cells[25, 11].Style.HorizontalAlignment = ExcelHorizontalAlignment.General;
                wsPivot.Cells[25, 11].Style.Fill.PatternType = ExcelFillStyle.Solid;
                wsPivot.Cells[25, 11].Style.Fill.BackgroundColor.SetColor(Color.Yellow);

                pck.Save();

                nethours = double.Parse(wsPivot.Cells[25,11].Value.ToString());
                
            }
            return nethours;
        }

        public double Tablecalcaubt(string pathins, string sheet, int RwCnt, int ClmnCnt, string Rwfld, string Clmnfld, string datafld)
        {
            FileInfo file = new FileInfo(pathins);
            double nethours = 0;

            using (ExcelPackage pck = new ExcelPackage(file))
            {
                ExcelWorksheet ws = pck.Workbook.Worksheets[sheet];
                var wsPivot = pck.Workbook.Worksheets.Add("PivotTest");
                var Pvt = wsPivot.PivotTables.Add(wsPivot.Cells[3, 1], ws.Cells[1, 1, RwCnt, ClmnCnt], "Test");
                Pvt.RowFields.Add(Pvt.Fields[Rwfld]);
                Pvt.RowFields[Pvt.RowFields.Count - 1].Sort = eSortType.Ascending;
                Pvt.ColumnFields.Add(Pvt.Fields[Clmnfld]);
                Pvt.ColumnFields[Pvt.ColumnFields.Count - 1].Sort = eSortType.Ascending;
                var values = Pvt.DataFields.Add(Pvt.Fields[datafld]);
                values.Function = DataFieldFunctions.Sum;
                values.Format = "#.##0";
                wsPivot.Cells[1,1,12,5].Style.Numberformat.Format = "#.##0";

                wsPivot.Cells[17, 1].Value = "Hız - Referans";
                wsPivot.Cells[26, 1].Value = "Kalınlık - Referans";
                wsPivot.Cells[35, 1].Value = "En - Referans";
                wsPivot.Cells[17, 7].Value = "Net Üretim Zamanı";

                wsPivot.Cells[18, 1].Value = "Group1";
                wsPivot.Cells[27, 1].Value = "Group1";
                wsPivot.Cells[36, 1].Value = "Group1";
                wsPivot.Cells[18, 7].Value = "Group1";

                wsPivot.Cells[19, 1].Value = "Group2";
                wsPivot.Cells[28, 1].Value = "Group2";
                wsPivot.Cells[37, 1].Value = "Group2";
                wsPivot.Cells[19, 7].Value = "Group2";

                wsPivot.Cells[20, 1].Value = "Group3";
                wsPivot.Cells[29, 1].Value = "Group3";
                wsPivot.Cells[38, 1].Value = "Group3";
                wsPivot.Cells[20, 7].Value = "Group3";

                wsPivot.Cells[21, 1].Value = "Group4";
                wsPivot.Cells[30, 1].Value = "Group4";
                wsPivot.Cells[39, 1].Value = "Group4";
                wsPivot.Cells[21, 7].Value = "Group4";

                wsPivot.Cells[22, 1].Value = "Group5";
                wsPivot.Cells[31, 1].Value = "Group5";
                wsPivot.Cells[40, 1].Value = "Group5";
                wsPivot.Cells[22, 7].Value = "Group5";

                wsPivot.Cells[23, 1].Value = "Group6";
                wsPivot.Cells[32, 1].Value = "Group6";
                wsPivot.Cells[41, 1].Value = "Group6";
                wsPivot.Cells[23, 7].Value = "Group6";

                wsPivot.Cells[24, 1].Value = "Group7";
                wsPivot.Cells[33, 1].Value = "Group7";
                wsPivot.Cells[42, 1].Value = "Group7";
                wsPivot.Cells[24, 7].Value = "Group7";

                wsPivot.Cells[17, 2].Value = "Group1";
                wsPivot.Cells[17, 3].Value = "Group2";
                wsPivot.Cells[17, 4].Value = "Group3";

                wsPivot.Cells[26, 2].Value = "Group1";
                wsPivot.Cells[26, 3].Value = "Group2";
                wsPivot.Cells[26, 4].Value = "Group3";

                wsPivot.Cells[35, 2].Value = "Group1";
                wsPivot.Cells[35, 3].Value = "Group2";
                wsPivot.Cells[35, 4].Value = "Group3";

                wsPivot.Cells[17, 8].Value = "Group1";
                wsPivot.Cells[17, 9].Value = "Group2";
                wsPivot.Cells[17, 10].Value = "Group3";

                // hız
                wsPivot.Cells[18, 2, 20, 2].Value = 90;
                wsPivot.Cells[21, 2, 21, 4].Value = 90;
                wsPivot.Cells[18, 3, 20, 3].Value = 85;
                wsPivot.Cells[19, 4, 20, 4].Value = 85;
                wsPivot.Cells[18, 4].Value = 80;
                wsPivot.Cells[22, 2].Value = 80;
                wsPivot.Cells[22, 3].Value = 70;
                wsPivot.Cells[22, 4].Value = 70;
                wsPivot.Cells[23, 2, 23, 4].Value = 50;
                wsPivot.Cells[24, 2, 24, 4].Value = 40;

                // Kalınlık
                wsPivot.Cells[27, 2].Value = 0.21;
                wsPivot.Cells[27, 3].Value = 0.20;
                wsPivot.Cells[27, 4].Value = 0.20;
                wsPivot.Cells[28, 2].Value = 0.23;
                wsPivot.Cells[28, 3].Value = 0.25;
                wsPivot.Cells[28, 4].Value = 0.27;
                wsPivot.Cells[29, 2, 29, 4].Value = 0.33;
                wsPivot.Cells[30, 2, 30, 4].Value = 0.50;
                wsPivot.Cells[31, 2, 31, 4].Value = 0.70;
                wsPivot.Cells[32, 2, 32, 4].Value = 0.88;
                wsPivot.Cells[33, 2, 33, 4].Value = 1.6;

                // EN
                wsPivot.Cells[36, 2, 42, 2].Value = 1000;
                wsPivot.Cells[36, 3, 42, 3].Value = 1250;
                wsPivot.Cells[36, 4, 42, 4].Value = 1500;

                // formuller
                // group1
                wsPivot.Cells[18, 8].Formula = "B5/2.71/B27/(B36/1000)/B18/60";
                wsPivot.Cells[19, 8].Formula = "B6/2.71/B28/(B37/1000)/B19/60";
                wsPivot.Cells[20, 8].Formula = "B7/2.71/B29/(B38/1000)/B20/60";
                wsPivot.Cells[21, 8].Formula = "B8/2.71/B30/(B39/1000)/B21/60";
                wsPivot.Cells[22, 8].Formula = "B9/2.71/B31/(B40/1000)/B22/60";
                wsPivot.Cells[23, 8].Formula = "B10/2.71/B32/(B41/1000)/B23/60";
                wsPivot.Cells[24, 8].Formula = "B11/2.71/B33/(B42/1000)/B24/60";

                // group2
                wsPivot.Cells[18, 9].Formula = "C5/2.71/C27/(C36/1000)/C18/60";
                wsPivot.Cells[19, 9].Formula = "C6/2.71/C28/(C37/1000)/C19/60";
                wsPivot.Cells[20, 9].Formula = "C7/2.71/C29/(C38/1000)/C20/60";
                wsPivot.Cells[21, 9].Formula = "C8/2.71/C30/(C39/1000)/C21/60";
                wsPivot.Cells[22, 9].Formula = "C9/2.71/C31/(C40/1000)/C22/60";
                wsPivot.Cells[23, 9].Formula = "C10/2.71/C32/(C41/1000)/C23/60";
                wsPivot.Cells[24, 9].Formula = "C11/2.71/C33/(C42/1000)/C24/60";

                // group3

                wsPivot.Cells[18, 10].Formula = "D5/2.71/D27/(D36/1000)/D18/60";
                wsPivot.Cells[19, 10].Formula = "D6/2.71/D28/(D37/1000)/D19/60";
                wsPivot.Cells[20, 10].Formula = "D7/2.71/D29/(D38/1000)/D20/60";
                wsPivot.Cells[21, 10].Formula = "D8/2.71/D30/(D39/1000)/D21/60";
                wsPivot.Cells[22, 10].Formula = "D9/2.71/D31/(D40/1000)/D22/60";
                wsPivot.Cells[23, 10].Formula = "D10/2.71/D32/(D41/1000)/D23/60";
                wsPivot.Cells[24, 10].Formula = "D11/2.71/D33/(D42/1000)/D24/60";
                wsPivot.Cells[25, 11].Style.Numberformat.Format = "#.##0";
                // toplamlar
                wsPivot.Cells["H25"].Formula = "SUM(H18:H24)";
                wsPivot.Cells["I25"].Formula = "SUM(I18:I24)";
                wsPivot.Cells["J25"].Formula = "SUM(J18:J24)";
                wsPivot.Cells["K25"].Formula = "SUM(H25:J25)";
                pck.Workbook.Calculate();

                //wsPivot.Cells[25,11].Value = wsPivot.Cells[25,11].Value;
                //foreach (var cell in wsPivot.Cells.Where(cell => cell.Formula != null))
                //{
                //    cell.Value = cell.Value;
                //}

                wsPivot.Cells[25, 11].Style.Font.SetFromFont(new Font("Tahoma", 10, FontStyle.Bold));
                wsPivot.Cells[25, 11].Style.Font.Color.SetColor(Color.Black);
                wsPivot.Cells[25, 11].Style.WrapText = false;
                wsPivot.Cells[25, 11].Style.VerticalAlignment = ExcelVerticalAlignment.Bottom;
                wsPivot.Cells[25, 11].Style.HorizontalAlignment = ExcelHorizontalAlignment.General;
                wsPivot.Cells[25, 11].Style.Fill.PatternType = ExcelFillStyle.Solid;
                wsPivot.Cells[25, 11].Style.Fill.BackgroundColor.SetColor(Color.Yellow);

                pck.Save();

                nethours = double.Parse(wsPivot.Cells[25, 11].Value.ToString());

            }
            return nethours;
        }
    }
}



////the label row field 
//pivottable.rowfields.add(pivottable.fields["col1"]); 
//    pivottable.dataonrows = false;  
//    //the data fields 
//    var field = pivottable.datafields.add(pivottable.fields["col3"]);
//field.name = "sum of col2"; field.function = datafieldfunctions.sum; field.format = formatcurrency;  
//    field = pivottable.datafields.add(pivottable.fields["col4"]); 
//    field.name = "sum of col3"; field.function = datafieldfunctions.sum; 
//    field.format = formatcurrency;  
//    //the page field 
//    pivottable.pagefields.add(pivottable.fields["col2"]); 
//    var xdcachedefinition = pivottable.cachedefinition.cachedefinitionxml;
//var xecachefields = xdcachedefinition.firstchild["cachefields"]; 
//if (xecachefields == null)    
//    return;  
//    //to filter, add items cache definition via xml 
//    var count = 0;
//var assetfieldidx = -1;  
//foreach (xmlelement cfield in xecachefields) 
//    {     
//    var att = cfield.attributes["name"];     
//if (att != null && att.value == "col2" )    
//    {         assetfieldidx = count;          
//    var shareditems = cfield.getelementsbytagname("shareditems")[0] xmlelement;         
//    if(shareditems == null)           
//    continue;          
//    //set collection attributes        
//    shareditems.removeallattributes();        
//    att = xdcachedefinition.createattribute("count");      
//    att.value = "3";      
//    shareditems.attributes.append(att);       
//    //create , add item       
//    var item = xdcachedefinition.createelement("s", shareditems.namespaceuri);
//att = xdcachedefinition.createattribute("v");      
//    att.value = "group a";    
//    item.attributes.append(att);    
//    shareditems.appendchild(item);      
//    item = xdcachedefinition.createelement("s", shareditems.namespaceuri);     
//    att = xdcachedefinition.createattribute("v");    
//    att.value = "group b";        
//    item.attributes.append(att);   
//    shareditems.appendchild(item);   
//    item = xdcachedefinition.createelement("s", shareditems.namespaceuri);   
//    att = xdcachedefinition.createattribute("v");     
//    att.value = "group c";      
//    item.attributes.append(att);    
//    shareditems.appendchild(item);      
//    break;     
//    }      
//    count++; 
//    }  //now go main pivot table xml , add cross references complete filtering var xdpivottable = pivottable.pivottablexml; var xdpivotfields = xdpivottable.firstchild["pivotfields"]; if (xdpivotfields == null)     return;  count = 0; foreach (xmlelement pfield in xdpivotfields) {     //find asset type field     if (count == assetfieldidx)     {         var att = xdpivottable.createattribute("multipleitemselectionallowed");         att.value = "1";         pfield.attributes.append(att);          var items = pfield.getelementsbytagname("items")[0] xmlelement;         items.removeall();          att = xdpivottable.createattribute("count");         att.value = "4";         items.attributes.append(att);         pfield.appendchild(items);          //add classes fields item collection         (var = 0; < 3; i++)         {             var item = xdpivottable.createelement("item", items.namespaceuri);             att = xdpivottable.createattribute("x");             att.value = i.tostring(cultureinfo.invariantculture);             item.attributes.append(att);              //turn of cash class in fielder             if (i == 1)             {                 att = xdpivottable.createattribute("h");                 att.value = "1";                 item.attributes.append(att);             }              items.appendchild(item);          }          //add default         var defaultitem = xdpivottable.createelement("item", items.namespaceuri);         att = xdpivottable.createattribute("t");         att.value = "default";         defaultitem.attributes.append(att);         items.appendchild(defaultitem);          break;     }     count++; }  pck.save(); 
/*wsPivot.Cells[18, 2, 20, 2].Value = int.Parse("90");
                wsPivot.Cells[21, 2, 21, 4].Value = int.Parse("90");
                wsPivot.Cells[18, 3, 20, 3].Value = int.Parse("85");
                wsPivot.Cells[19, 4, 20, 4].Value = int.Parse("85");
                wsPivot.Cells[18, 4].Value = int.Parse("80");
                wsPivot.Cells[22, 2].Value = int.Parse("80");
                wsPivot.Cells[22, 3].Value = int.Parse("70");
                wsPivot.Cells[22, 4].Value = int.Parse("70");
                wsPivot.Cells[23, 2, 23, 4].Value = int.Parse("50");
                wsPivot.Cells[24, 2, 24, 4].Value = int.Parse("40");

                // Kalınlık
                wsPivot.Cells[27, 2].Value = double.Parse("0,21");
                wsPivot.Cells[27, 3].Value = double.Parse("0,20");
                wsPivot.Cells[27, 4].Value = double.Parse("0,20");
                wsPivot.Cells[28, 2].Value = double.Parse("0,23");
                wsPivot.Cells[28, 3].Value = double.Parse("0,25");
                wsPivot.Cells[28, 4].Value = double.Parse("0,27");
                wsPivot.Cells[29, 2, 29, 4].Value = double.Parse("0,33");
                wsPivot.Cells[30, 2, 30, 4].Value = double.Parse("0,50");
                wsPivot.Cells[31, 2, 31, 4].Value = double.Parse("0,70");
                wsPivot.Cells[32, 2, 32, 4].Value = double.Parse("0,88");
                wsPivot.Cells[33, 2, 33, 4].Value = double.Parse("1,6");

                // EN
                wsPivot.Cells[36, 2, 42, 2].Value = int.Parse("1000");
                wsPivot.Cells[36, 3, 42, 3].Value = int.Parse("1250");
                wsPivot.Cells[36, 4, 42, 4].Value = int.Parse("1500");



                // formuller
                // group1
                wsPivot.Cells[18, 8].Formula = "B5*1000/2,71/B27/(B36/1000)/B18/60";
                wsPivot.Cells[19, 8].Formula = "B6*1000/2,71/B28/(B37/1000)/B19/60";
                wsPivot.Cells[20, 8].Formula = "B7*1000/2,71/B29/(B38/1000)/B20/60";
                wsPivot.Cells[21, 8].Formula = "B8*1000/2,71/B30/(B39/1000)/B21/60";
                wsPivot.Cells[22, 8].Formula = "B9*1000/2,71/B31/(B40/1000)/B22/60";
                wsPivot.Cells[23, 8].Formula = "B10*1000/2,71/B32/(B41/1000)/B23/60";
                wsPivot.Cells[24, 8].Formula = "B11*1000/2,71/B33/(B42/1000)/B24/60";

                // group2
                wsPivot.Cells[18, 9].Formula = "C5*1000/2,71/C27/(C36/1000)/C18/60";
                wsPivot.Cells[19, 9].Formula = "C6*1000/2,71/C28/(C37/1000)/C19/60";
                wsPivot.Cells[20, 9].Formula = "C7*1000/2,71/C29/(C38/1000)/C20/60";
                wsPivot.Cells[21, 9].Formula = "C8*1000/2,71/C30/(C39/1000)/C21/60";
                wsPivot.Cells[22, 9].Formula = "C9*1000/2,71/C31/(C40/1000)/C22/60";
                wsPivot.Cells[23, 9].Formula = "C10*1000/2,71/C32/(C41/1000)/C23/60";
                wsPivot.Cells[24, 9].Formula = "C11*1000/2,71/C33/(C42/1000)/C24/60";

                // group3

                wsPivot.Cells[18, 10].Formula = "D5*1000/2,71/D27/(D36/1000)/D18/60";
                wsPivot.Cells[19, 10].Formula = "D6*1000/2,71/D28/(D37/1000)/D19/60";
                wsPivot.Cells[20, 10].Formula = "D7*1000/2,71/D29/(D38/1000)/D20/60";
                wsPivot.Cells[21, 10].Formula = "D8*1000/2,71/D30/(D39/1000)/D21/60";
                wsPivot.Cells[22, 10].Formula = "D9*1000/2,71/D31/(D40/1000)/D22/60";
                wsPivot.Cells[23, 10].Formula = "D10*1000/2,71/D32/(D41/1000)/D23/60";
                wsPivot.Cells[24, 10].Formula = "D11*1000/2,71/D33/(D42/1000)/D24/60";

                // toplamlar
                wsPivot.Cells["H25"].Formula = "SUM(H18:H24)";
                wsPivot.Cells["I25"].Formula = "SUM(I18:I24)";
                wsPivot.Cells["J25"].Formula = "SUM(J18:J24)";
                wsPivot.Cells["K25"].Formula = "SUM(H25:J25)";
*/