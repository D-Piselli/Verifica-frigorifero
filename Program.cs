using ClosedXML.Excel;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Reflection.Metadata.Ecma335;

string filePath = "C:/Documenti/VerificaFrigorifero.xlsx";
using (var workbook = new XLWorkbook())
{
    var worksheet = workbook.Worksheets.Add("Cella Frigorifica");
    var headerRow = worksheet.Row(1);
    headerRow.Style.Font.Bold = true;
    worksheet.Cell(1, 1).Value = "Revelazioni";
    worksheet.Cell(1, 1).Style.Fill.BackgroundColor = XLColor.Yellow;
    worksheet.Cell(1, 1).Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
    worksheet.Cell(1, 2).Value = "Nibble";
    worksheet.Cell(1, 2).Style.Fill.BackgroundColor = XLColor.Yellow;
    worksheet.Cell(1, 2).Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
    worksheet.Column(2).Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
    worksheet.Cell(1, 3).Value = "Hex";
    worksheet.Cell(1, 3).Style.Fill.BackgroundColor = XLColor.Yellow;
    worksheet.Cell(1, 3).Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
    worksheet.Column(3).Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
    worksheet.Cell(1, 4).Value = "Note";
    worksheet.Cell(1, 4).Style.Fill.BackgroundColor = XLColor.Yellow;
    worksheet.Cell(1, 4).Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
    worksheet.Column(1).Width = 13;
    worksheet.Column(2).Width = 6;
    worksheet.Column(3).Width = 6;
    worksheet.Column(4).Width = 30;

    int temp =0, revel,value;
    string binario = "", hex = "";
    Console.WriteLine("inserisci il numero di revelazioni");
    revel = Convert.ToInt16(Console.ReadLine());

    for (int i = 0; i < revel; i++)
    {
        Console.WriteLine("inserisci la temperatura della " + (i + 1) + "a revelazione in binario");
        for (int j = 3; j >= 0; j--)
        {
            Console.WriteLine("inserisci la " + (j + 1) + "a cifra binaria(da sinistra a destra)");
            value=Convert.ToInt16(Console.ReadLine());
            temp = temp + (value * (2^j));
            binario = binario + value;
        }
        if (temp < 10)
        {
            hex =""+ temp;
        }
        else
        {
            switch (temp)
            {
                case 10:
                    hex = "A";
                    break;
                case 11:
                    hex= "B";
                    break;
                case 12:
                    hex= "C";
                    break;
                case 13:
                    hex= "D";
                    break;
                case 14:
                    hex="E";
                    break;
                case 15:
                    hex="F";
                    break;
            }
        }
        if (temp >=7)
        {
            worksheet.Cell((i + 2), 1).Style.Fill.BackgroundColor = XLColor.LightPastelPurple;
            worksheet.Cell((i + 2), 2).Style.Fill.BackgroundColor = XLColor.LightPastelPurple;
            worksheet.Cell((i + 2), 3).Style.Fill.BackgroundColor = XLColor.LightPastelPurple;
            worksheet.Cell((i + 2), 4).Style.Fill.BackgroundColor = XLColor.LightPastelPurple;
            worksheet.Cell((i + 2), 1).Style.Font.Bold = true;
            worksheet.Cell((i + 2), 2).Style.Font.Bold = true;
            worksheet.Cell((i + 2), 3).Style.Font.Bold = true;
            worksheet.Cell((i + 2), 4).Style.Font.Bold = true;
            worksheet.Cell((i + 2), 4).Value = "Attenzione!Possibilita' avaria!";
        }
        else if (temp>2&&temp<7)
        {
            worksheet.Cell((i + 2), 4).Value = "Temperatura nella norma.";
        }

        else if(temp<=2)
        {
            worksheet.Cell((i + 2), 4).Value = "Attenzione!Rischio congelamento!";
        }
        worksheet.Cell((i + 2), 1).Value = "Rivelazione n" + (i + 1);
        worksheet.Cell((i + 2), 2).Value = "" + binario;
        worksheet.Cell((i + 2), 3).Value = "" + hex;
        worksheet.Cell((i+2), 1).Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
        worksheet.Cell((i + 2), 2).Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
        worksheet.Cell((i + 2), 3).Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
        worksheet.Cell((i + 2), 4).Style.Border.OutsideBorder = XLBorderStyleValues.Thin;

        binario = "";
        temp = 0;
    }
    workbook.SaveAs(filePath);
    Console.WriteLine("Il File "+filePath+" e' stato creato con successo!");
}