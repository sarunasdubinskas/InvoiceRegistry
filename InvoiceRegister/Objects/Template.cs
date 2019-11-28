using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace InvoiceRegister.Objects
{
    class Template
    {
        public string PathToTemplate { get; } = ConfigurationManager.AppSettings["TemplatePath"];
        private string myUniqueFileName = $@"{Guid.NewGuid()}.xlsx";
        public string PathToFile { get; set; }
        public int InvoiceID { get; set; }
        public string ClientAddress { get; set; }
        public string PostalCode { get; set; }
        public string InvoiceDate { get; set; }
        public string ClientEntityCode { get; set; }
        public int InvoiceNumber { get; set; }
        public string ClientVATCode { get; set; }
        public string InvoiceSeries { get; set; }
        public int InvoiceSeriesID { get; set; }
        public string ClientName { get; set; }
        public int ClientID { get; set; }

        public void OpenTemplate()
        {
            Microsoft.Office.Interop.Excel.Application excelApp = new Microsoft.Office.Interop.Excel.Application();
            excelApp.Visible = true;
            Workbooks books = excelApp.Workbooks;
            Workbook book = books.Open(PathToTemplate);
        }


        /// <summary>
        /// default template
        /// </summary>
        public void SetPath()
        {
            PathToFile = ConfigurationManager.AppSettings["InvoiceDefaultDirectoryPath"] + myUniqueFileName;
            try
            {
                File.Copy(PathToTemplate, PathToFile);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }

        /// <summary>
        /// Overload for parent if not default template
        /// </summary>
        public void SetPath(string origin)
        {
            PathToFile = ConfigurationManager.AppSettings["InvoiceDefaultDirectoryPath"] + myUniqueFileName;
            try
            {
                File.Copy(origin, PathToFile);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }
    }
}
