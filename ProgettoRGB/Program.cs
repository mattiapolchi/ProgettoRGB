using System;
using System.Drawing;
using ClosedXML.Excel;
using DocumentFormat.OpenXml.Spreadsheet;

int red = 0, green = 0, blue = 0;

string filePath = "C:/DOCUMENTI/ProgettoRGB.xlsx";
using (var workbook = new XLWorkbook())
{
    var worksheet = workbook.Worksheets.Add("RGBfoglio");

    worksheet.Cell(2, 2).Value = ("Red");
    worksheet.Cell(2, 3).Value = ("Green");
    worksheet.Cell(2, 4).Value = ("Blue");

    Console.WriteLine("Inserisci il codice per il rosso");
    red = Convert.ToInt32(Console.ReadLine());

    Console.WriteLine("Inserisci il codice per il verde");
    green = Convert.ToInt32(Console.ReadLine());

    Console.WriteLine("Inserisci il codice per il blu");
    blue = Convert.ToInt32(Console.ReadLine());

    worksheet.Cell(2, 7).Value = ("Il colore ottenuto è il seguente: ");

    worksheet.Cell(3, 8).Style.Fill.BackgroundColor = XLColor.FromArgb(red, green, blue);
        
    workbook.SaveAs(filePath);
}