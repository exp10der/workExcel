using System;
using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using lol;

internal class Program
{
    private static void IterateRowsAndCells(string filename, List<Import> list)
    {
        using (var doc =
            SpreadsheetDocument.Open(filename, false))
        {
            var worksheet =
                (WorksheetPart) doc.WorkbookPart.GetPartById("rId1");

            list.AddRange(worksheet.Rows().Select(row => row.GetImport()));
        }
    }

    public static void Main()
    {
        var list = new List<Import>();
        IterateRowsAndCells("work.xlsx", list);

        var group = list.GroupBy(n => n.ObjectBilder).ToDictionary(n=>n.Key);


        foreach (var item in group)
        {
            Console.WriteLine(item.Key); 
        }

    }
}