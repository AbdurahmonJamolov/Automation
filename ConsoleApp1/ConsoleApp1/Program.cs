using System;
using System.IO;
using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Interop.Word;

class Program
{
    static void Main(string[] args)
    {


        string path = @"E:\Автоматизации документов\19_Кулоб\21_Кулоб.xlsx";

        Microsoft.Office.Interop.Excel.Application excel = new Microsoft.Office.Interop.Excel.Application();
        Workbook wb = excel.Workbooks.Open(path);
        Worksheet excelSheet = wb.ActiveSheet;

        Microsoft.Office.Interop.Word.Application wordApp = new Microsoft.Office.Interop.Word.Application();


        Document wordDocument = wordApp.Documents.Add();
        Documents wDocs = wordApp.Documents;
        Microsoft.Office.Interop.Word.Document wDoc = wDocs.Open(@"E:\Автоматизации документов\19_Кулоб\Кулоб\ИНН.docx", ReadOnly: false); //querySelector выборка
        wDoc.Activate();

        Bookmarks wBookmarks = wDoc.Bookmarks; // Это записывает библиотеку

        Bookmark wDog = wBookmarks["wDog"];  // Закладка wDog
        Microsoft.Office.Interop.Word.Range wRangeDog = wDog.Range;

        Bookmark wData = wBookmarks["wDate"];  // Закладка wData
        Microsoft.Office.Interop.Word.Range wRangeData = wData.Range;

        Bookmark wName1 = wBookmarks["wName1"];  // Закладка wName
        Microsoft.Office.Interop.Word.Range wRangeName1 = wName1.Range;

        Bookmark wName2 = wBookmarks["wName2"];  // Закладка wName2
        Microsoft.Office.Interop.Word.Range wRangeName2 = wName2.Range;

        Bookmark wTax = wBookmarks["wTax"];  // Закладка wTax
        Microsoft.Office.Interop.Word.Range wRangewTax = wTax.Range;


        for (int i = 3; i < 196; i++)
        {
            string dog = excelSheet.Cells[i, 4].Value.ToString();// второй элемент номер ячейки в эксель
            string date = excelSheet.Cells[i, 5].Value.ToString().Substring(0, 10);//для вывода даты без времени и секунд
            string name = excelSheet.Cells[i, 2].Value.ToString();
            string tax = excelSheet.Cells[i, 3].Value.ToString();

            Console.WriteLine($"{dog}, {date}, {name}, {tax}");

            wRangeDog.Text = dog;
            wRangeData.Text = date;
            wRangeName1.Text = name;
            wRangeName2.Text = name;
            wRangewTax.Text = tax;

            wDoc.SaveAs2($@"E:\Автоматизации документов\19_Кулоб\Кулоб\ИНН-{tax}.docx");

        }

        wDoc.Close();
        wb.Close();
        Console.Read();
    }
}
