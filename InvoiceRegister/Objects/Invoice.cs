using Microsoft.Office.Interop.Excel;
using System.Diagnostics;
using System.Windows.Forms;

namespace InvoiceRegister.Objects
{
    class Invoice
    {
        public void FillTemplate(Template template)
        {
            //COM Exception managing vars & params
            bool doable = true; int counter = 0; int tryCount = 15; int waitBetweenTriesMs = 100; 

            Microsoft.Office.Interop.Excel.Application excelApp = new Microsoft.Office.Interop.Excel.Application();
            excelApp.Visible = true;

            Workbooks books = excelApp.Workbooks;
            Workbook book = books.Open(template.PathToFile);
            Worksheet sheet = book.ActiveSheet;

            Range excelRange = (Range)Globals.ThisAddIn.Application.Selection;

            while (doable)
            {
                try
                {
                    sheet.get_Range("Serija").Value2 = template.InvoiceSeries;
                    sheet.get_Range("Numeris").Value2 = template.InvoiceNumber;
                    sheet.get_Range("Uzsakovas").Value2 = template.ClientName;
                    sheet.get_Range("Adresas").Value2 = template.ClientAddress;
                    sheet.get_Range("Kodas").Value2 = template.ClientEntityCode;
                    sheet.get_Range("PVM_Kodas").Value2 = template.ClientVATCode;
                    sheet.get_Range("Data").Value2 = template.InvoiceDate;
                    doable = false;
                }
                catch (System.Runtime.InteropServices.COMException ex)
                {
                    counter++;
                    if (counter>=tryCount)
                    {
                        doable = false;
                        MessageBox.Show(ex.ToString());
                    }
                    System.Threading.Thread.Sleep(waitBetweenTriesMs);

                }
            }

            
        }

        public void OpenExistingInvoice(Template template)
        {
            Microsoft.Office.Interop.Excel.Application excelApp = new Microsoft.Office.Interop.Excel.Application();
            excelApp.Visible = true;

            Workbooks books = excelApp.Workbooks;
            Workbook book = books.Open(template.PathToFile);
            Worksheet sheet = book.ActiveSheet;
        }
    }
}
