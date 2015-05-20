using System;
using System.Collections.Generic;
using System.Globalization;
using DocumentFormat.OpenXml.Office2010.ExcelAc;
using Microsoft.Examples.LtxOpenXml;
using DocumentFormat.OpenXml.Packaging;
using lol;

class Program
{
    static void IterateRowsAndCells(string filename)
    {
        //Console.WriteLine("Contents of Spreadsheet");
        //Console.WriteLine("=======================");

        List<Import> list = new List<Import>();

        using (SpreadsheetDocument doc =
            SpreadsheetDocument.Open(filename, false))
        {
            WorksheetPart worksheet =
                (WorksheetPart)doc.WorkbookPart.GetPartById("rId1");

            foreach (var row in worksheet.Rows())
            {
                //Console.WriteLine("  RowId:{0}", row.RowId);
                //Console.WriteLine("  Spans:{0}", row.Spans);

                //foreach (var cell in row.Cells())
                //{
                //      Console.WriteLine(cell.GetString());
                //}

              //  list.Add(row.GetImport(row.Cells()));  
                list.Add(row.GetImport());

                //var cells = row.Cells().ToArray();

                //string t=   cells[1].GetString();

                //foreach (var cell in row.Cells())//.Where(c => c.ColumnId == "A"))
                //{
                //    Console.WriteLine("    Column:{0}", cell.Column);
                //    Console.WriteLine("      ColumnId:{0}", cell.ColumnId);
                //    if (cell.Type != null)
                //        Console.WriteLine("      Type:{0}", cell.Type);
                //    if (cell.Value != null)
                //        Console.WriteLine("      Value:{0}", cell.Value);
                //    if (cell.Formula != null)
                //        Console.WriteLine("      Formula:>{0}<", cell.Formula);
                //    if (cell.SharedString != null)
                //        Console.WriteLine("      SharedString:>{0}<", cell.SharedString);
                //}
              //  Console.WriteLine("-----------------------");
            }
            Console.WriteLine();
        }
    }


    public static void Main(string[] args)
    {

       
        

        IterateRowsAndCells("work.xlsx");
       // IterateRowsAndCells("Book1.xlsx");
    }
}


