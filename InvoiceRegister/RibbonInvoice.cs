using Microsoft.Office.Tools.Ribbon;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;
using System.Configuration;
using System.Data.SqlClient;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System;
using InvoiceRegister.Objects;
using System.Data;
using InvoiceRegister.Data;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;

namespace InvoiceRegister
{
    /*Duomenų validacija ir išvesti į ekraną kas negerai. pašto kodas. search and filtering įrankis*/

    public partial class RibbonInvoice
    {

        string connectionString = ConfigurationManager.ConnectionStrings["FoxVoice"].ConnectionString;
        DestinationDataWriter destinationDataWriter = new DestinationDataWriter();
        private int sql_element;
        private void RibbonInvoice_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private string ReadNextSQLAsString(SqlDataReader reader)
        {
            string temp;
            try
            {
                temp = reader.GetString(sql_element);
                sql_element++;
                return temp;
            }

            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
                sql_element++;
                return null;
            }
        }

        private int ReadNextSQLAsInt32(SqlDataReader reader)
        {
            int temp;
            try
            {
                temp = reader.GetInt32(sql_element);
                sql_element++;
                return temp;
            }

            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
                sql_element++;
                return 0;
            }
        }

        private string ReadNextSQLAsString_NotNull(SqlDataReader reader)
        {
            string temp = null;
            try
            {
                if (!reader.IsDBNull(sql_element))
                {
                    temp = reader.GetString(sql_element);
                }
                sql_element++;
                return temp;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
                sql_element++;
                return temp;
            }
        }

        private void WriteToExcelNextCellSQLAsString(SqlDataReader reader, int offset_rows, int offset_columns, ref Range excelRange)
        {
            try
            {
                excelRange = excelRange.Offset[offset_rows, offset_columns];
                excelRange.Select();
                excelRange.Value2 = reader.GetString(sql_element);
                sql_element++;
            }
            
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }

        private void WriteToExcelNextCellSQLAsString_NotNull(SqlDataReader reader, int offset_rows, int offset_columns, ref Range excelRange)
        {
            try
            {
                excelRange = excelRange.Offset[offset_rows, offset_columns];
                excelRange.Select();
                if (!reader.IsDBNull(sql_element))
                {
                    excelRange.Value2 = reader.GetString(sql_element);
                }
                sql_element++;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }

        }

        private void WriteToExcelNextCellSQLAsInt32(SqlDataReader reader, int offset_rows, int offset_columns, ref Range excelRange)
        {
            try
            {
                excelRange = excelRange.Offset[offset_rows, offset_columns];
                excelRange.Select();
                excelRange.Value2 = reader.GetInt32(sql_element);
                sql_element++;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }

        }

        private void WriteToExcelNextCellSQLAsDatetime(SqlDataReader reader, int offset_rows, int offset_columns, ref Range excelRange)
        {
            try
            {
                excelRange = excelRange.Offset[offset_rows, offset_columns];
                excelRange.Select();
                excelRange.Value2 = reader.GetDateTime(sql_element);
                sql_element++;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }
        private void PrepareTableForInvoices(ref Range excelRange)
        {
            WriteTextInOffsetCell(ref excelRange, 0, 0, "Name");
            WriteTextInOffsetCell(ref excelRange, 0, 1, "Adress");
            WriteTextInOffsetCell(ref excelRange, 0, 1, "Postal Code");
            WriteTextInOffsetCell(ref excelRange, 0, 1, "Entity Code");
            WriteTextInOffsetCell(ref excelRange, 0, 1, "VAT Code");
            WriteTextInOffsetCell(ref excelRange, 0, 1, "Invoice ID");
            WriteTextInOffsetCell(ref excelRange, 0, 1, "Serries ID");
            WriteTextInOffsetCell(ref excelRange, 0, 1, "Invoice number");
            WriteTextInOffsetCell(ref excelRange, 0, 1, "Date");
            WriteTextInOffsetCell(ref excelRange, 0, 1, "Series");
            WriteTextInOffsetCell(ref excelRange, 0, 1, "Client ID");
        }
        private void PrepareTableForClients(Range excelRange)
        {
            WriteTextInOffsetCell(ref excelRange, 0, 0, "ID");
            WriteTextInOffsetCell(ref excelRange, 0, 1, "Name");
            WriteTextInOffsetCell(ref excelRange, 0, 1, "Adress");
            WriteTextInOffsetCell(ref excelRange, 0, 1, "Postal Code");
            WriteTextInOffsetCell(ref excelRange, 0, 1, "Entity Code");
            WriteTextInOffsetCell(ref excelRange, 0, 1, "VAT Code");
        }
        private void WriteTextInOffsetCell(ref Range excelRange, int offset_rows, int offset_columns, string text)
        {
            excelRange = excelRange.Offset[offset_rows, offset_columns];
            excelRange.Select();
            excelRange.Value2 = text;
        }

        private void WriteTextInOffsetCell_NotNull(ref Range excelRange, int offset_rows, int offset_columns, string text)
        {
            excelRange = excelRange.Offset[offset_rows, offset_columns];
            excelRange.Select();
            if (text != null) { excelRange.Value2 = text; }
        }

        private void btnNewClient_Click(object sender, RibbonControlEventArgs e)
        {
            Client client = ClientInputDialog();

            if (client!=null)
            {
                destinationDataWriter.InsertNewClient(client);
            }
        }

        private Client ClientInputDialog()
        {
            Client client = new Client();
            Excel.Application ExcelApp = (Excel.Application)Marshal.GetActiveObject("Excel.Application");
            Excel.Workbook book = ExcelApp.ActiveWorkbook;
            try
            {
                client.Name = book.Application.InputBox("Entity Name:", "New Client Form", Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                client.PostalCode = book.Application.InputBox("Postal Code:", "New Client Form", Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                client.Adress = book.Application.InputBox("Adress:", "New Client Form", Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                client.EntityCode = book.Application.InputBox("Entity Code:", "New Client Form", Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                client.VATCode = book.Application.InputBox("VAT Code:", "New Client Form", Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            }
            catch (Exception ex)
            {
                client = null;
                MessageBox.Show(ex.ToString());
            }

            return client;
        }

        private void btnLoadClientList_Click(object sender, RibbonControlEventArgs e)
        {
            int row = 1;
            bool isSheetCreated = false;

            string querry = @"SELECT dbo.Client.ID, dbo.Client.Name, dbo.Client.Address, dbo.Client.PostalCode, dbo.Client.EntityCode, dbo.Client.VATCode
                                     FROM dbo.Client";

            Excel.Application ExcelApp = (Excel.Application)Marshal.GetActiveObject("Excel.Application");
            Excel.Workbook book = ExcelApp.ActiveWorkbook;
            Excel.Worksheet sheet;

            foreach (Excel.Worksheet sh in book.Sheets)
            {
                if (sh.Name == "Clients")
                {
                    isSheetCreated = true;
                    break;
                }
            }

            if (isSheetCreated)
            {
                sheet = book.Sheets["Clients"];
                sheet.Activate();
            }
            else
            {
                sheet = book.Application.Worksheets.Add();
                sheet.Name = "Clients";
            }

            sheet.Cells.ClearContents();

            Excel.Range excelRange = sheet.get_Range("A" + row);

            PrepareTableForClients(excelRange);
            row++;
            excelRange = sheet.get_Range("A" + row);

            using (SqlConnection connection = new SqlConnection(connectionString))
            {
                using (SqlCommand command = new SqlCommand(querry, connection))
                {
                    connection.Open();
                    using (SqlDataReader reader = command.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            sql_element = 0;
                            WriteToExcelNextCellSQLAsInt32(reader, 0, 0, ref excelRange);
                            WriteToExcelNextCellSQLAsString(reader, 0, 1, ref excelRange);
                            WriteToExcelNextCellSQLAsString(reader, 0, 1, ref excelRange);
                            WriteToExcelNextCellSQLAsString(reader, 0, 1, ref excelRange);
                            WriteToExcelNextCellSQLAsString(reader, 0, 1, ref excelRange);
                            WriteToExcelNextCellSQLAsString_NotNull(reader, 0, 1, ref excelRange);
                            row++;
                            excelRange = sheet.get_Range("A" + row);
                        }
                        sheet.Columns.AutoFit();
                    }
                }
                    
            }
        }

        private void btnUpdateClientInfo_Click(object sender, RibbonControlEventArgs e)
        {
            Client client = new Client();
            
            Range excelRange = (Range)Globals.ThisAddIn.Application.Selection;
            int row = excelRange.Row;
            
            Excel.Application excelApp = (Excel.Application)Marshal.GetActiveObject("Excel.Application");
            Workbook book = excelApp.ActiveWorkbook;
            Worksheet sheet = book.ActiveSheet;
            excelRange = sheet.get_Range("A" + row);

            client.Id = ReadExcelNextCellAsInt32(0, 0, ref excelRange);
            client.Name = ReadExcelNextCellAsString(0, 1, ref excelRange);
            client.Adress = ReadExcelNextCellAsString(0, 1, ref excelRange);
            client.PostalCode = ReadExcelNextCellAsString(0, 1, ref excelRange);
            client.EntityCode = ReadExcelNextCellAsString(0, 1, ref excelRange);
            client.VATCode = ReadExcelNextCellAsString(0, 1, ref excelRange);

            destinationDataWriter.UpdateClient(client);
        }


        private void btnGenerateInvoice_Click(object sender, RibbonControlEventArgs e)
        {
            Range excelRange = (Range)Globals.ThisAddIn.Application.Selection;
            int row = excelRange.Row;

            Template template = new Template();
            Invoice invoice = new Invoice();

            Excel.Application excelApp = (Excel.Application)Marshal.GetActiveObject("Excel.Application");
            Workbook book = excelApp.ActiveWorkbook;
            Worksheet sheet = book.ActiveSheet;
            excelRange = sheet.get_Range("A" + row);

            template.ClientName = ReadExcelNextCellAsString(0, 0, ref excelRange);
            template.ClientAddress = ReadExcelNextCellAsString(0,1, ref excelRange) + ", " + ReadExcelNextCellAsString(0,1, ref excelRange);
            template.ClientEntityCode = ReadExcelNextCellAsString(0,1, ref excelRange);
            template.ClientVATCode = ReadExcelNextCellAsString(0, 1, ref excelRange);
            template.InvoiceID = ReadExcelNextCellAsInt32(0, 1, ref excelRange);
            template.InvoiceSeriesID = ReadExcelNextCellAsInt32(0, 1, ref excelRange);
            template.InvoiceNumber = ReadExcelNextCellAsInt32(0, 1, ref excelRange);
            template.InvoiceDate = ReadExcelNextCellAsString(0, 1, ref excelRange);
            template.InvoiceSeries = ReadExcelNextCellAsString(0, 1, ref excelRange);
            template.ClientID = ReadExcelNextCellAsInt32(0, 1, ref excelRange);

            template.SetPath();
            invoice.FillTemplate(template);
            destinationDataWriter.InsertNewInvoice(template);            
        }

        private void btnOpenTemplate_Click(object sender, RibbonControlEventArgs e)
        {
            Template template = new Template();
            template.OpenTemplate();
        }


        private void btnNewFromThis_Click(object sender, RibbonControlEventArgs e)
        {
            Range excelRange = (Range)Globals.ThisAddIn.Application.Selection;
            int row = excelRange.Row;

            Excel.Application excelApp = (Excel.Application)Marshal.GetActiveObject("Excel.Application");
            Workbook book = excelApp.ActiveWorkbook;
            Worksheet sheet = book.ActiveSheet;

            excelRange = sheet.get_Range("A" + row);

            Template template = new Template();
            Invoice invoice = new Invoice();

            template.ClientName = ReadExcelNextCellAsString(0, 0, ref excelRange);
            template.ClientAddress = ReadExcelNextCellAsString(0, 1, ref excelRange) + ", " + ReadExcelNextCellAsString(0, 1, ref excelRange);
            template.ClientEntityCode = ReadExcelNextCellAsString(0, 1, ref excelRange);
            template.ClientVATCode = ReadExcelNextCellAsString(0, 1, ref excelRange);
            template.InvoiceID = ReadExcelNextCellAsInt32(0, 1, ref excelRange);
            template.InvoiceSeriesID = ReadExcelNextCellAsInt32(0, 1, ref excelRange);
            template.InvoiceNumber = ReadExcelNextCellAsInt32(0, 1,ref excelRange);
            template.InvoiceDate = ReadExcelNextCellAsString(0, 1, ref excelRange);
            template.InvoiceSeries = ReadExcelNextCellAsString(0, 1,ref  excelRange);
            template.ClientID = ReadExcelNextCellAsInt32(0, 1, ref excelRange);

            string querry = @"SELECT TOP (1) [Number] FROM[FoxVoice].[dbo].[Invoices] where Series = @SerriesID order by Number desc";
            using(SqlConnection connection = new SqlConnection(connectionString))
            {
                using (SqlCommand command = new SqlCommand(querry, connection))
                {
                    connection.Open();

                    SqlParameter param = new SqlParameter();
                    param.ParameterName = "@SerriesID";
                    param.Value = template.InvoiceSeriesID;
                    command.Parameters.Add(param);

                    using (SqlDataReader reader = command.ExecuteReader())
                    {
                        try
                        {
                            reader.Read();
                            template.InvoiceNumber = reader.GetInt32(0) + 1;
                        }
                        catch (Exception ex)
                        {
                            MessageBox.Show(Convert.ToString(ex));
                        }
                    }
                }
                    
            }

            template.InvoiceDate = DateTime.Now.ToString("yyyy-M-d");

            template.SetPath();
            invoice.FillTemplate(template);
            destinationDataWriter.InsertNewInvoice(template);

        }

        private void btnOpenInvoice_Click(object sender, RibbonControlEventArgs e)
        {
            Range excelRange = (Range)Globals.ThisAddIn.Application.Selection;
            int row = excelRange.Row;

            Excel.Application excelApp = (Excel.Application)Marshal.GetActiveObject("Excel.Application");
            Workbook book = excelApp.ActiveWorkbook;
            Worksheet sheet = book.ActiveSheet;

            excelRange = sheet.get_Range("A" + row);

            Template template = new Template();
            Invoice invoice = new Invoice();

            excelRange = excelRange.Offset[0, 5];
            excelRange.Select();
            template.InvoiceID = Convert.ToInt32(excelRange.Value2);
            
            string querry = @"SELECT Path_to_Invoice from dbo.Files Where Invoice_ID =@InvoiceID";
            
            using (SqlConnection connection = new SqlConnection(connectionString))
            {
                using (SqlCommand command = new SqlCommand(querry, connection))
                {
                    connection.Open();

                    SqlParameter param = new SqlParameter();
                    param.ParameterName = "@InvoiceID";
                    param.Value = template.InvoiceID;
                    command.Parameters.Add(param);

                    using (SqlDataReader reader = command.ExecuteReader())
                    {
                        try
                        {
                            reader.Read();
                            template.PathToFile = reader.GetString(0);
                        }
                        catch (Exception ex)
                        {
                            MessageBox.Show(ex.ToString());
                        }
                    }
                }                
            }
            invoice.OpenExistingInvoice(template);   
        }

        private string ReadExcelNextCellAsString(int offset_rows, int offset_columns, ref Range range)
        {
            range = range.Offset[offset_rows, offset_columns];
            range.Select();
            return Convert.ToString(range.Value);
        }
        private int ReadExcelNextCellAsInt32(int offset_rows, int offset_columns, ref Range range)
        {
            range = range.Offset[offset_rows, offset_columns];
            range.Select();
            return Convert.ToInt32(range.Value2);
        }

        private void btnLoadInvoiceList_Click(object sender, RibbonControlEventArgs e)
        {
            int row = 1;
            bool isSheetCreated = false;

            string querry = @"SELECT dbo.Client.Name, dbo.Client.Address, dbo.Client.PostalCode, dbo.Client.EntityCode, dbo.Client.VATCode, 
                                     dbo.Invoices.Id, dbo.Invoices.Series, dbo.Invoices.Number, dbo.Invoices.Date, dbo.Series_ID.Title, dbo.Client.ID
                                     FROM dbo.Client INNER JOIN
                                     dbo.Invoices ON dbo.Client.ID = dbo.Invoices.Client_ID INNER JOIN
                                     dbo.Series_ID ON dbo.Invoices.Series = dbo.Series_ID.Series_ID";

            Excel.Application ExcelApp = (Excel.Application)Marshal.GetActiveObject("Excel.Application");
            Excel.Workbook book = ExcelApp.ActiveWorkbook;
            Excel.Worksheet sheet = book.ActiveSheet;

            foreach (Excel.Worksheet sh in book.Sheets)
            {
                if (sh.Name == "Invoices")
                {
                    isSheetCreated = true;
                    break;
                }
            }

            if (isSheetCreated)
            {
                sheet = book.Sheets["Invoices"];
                sheet.Activate();
            }
            else
            {
                sheet = book.Application.Worksheets.Add();
                sheet.Name = "Invoices";
            }

            sheet.Cells.ClearContents();

            Excel.Range excelRange = sheet.get_Range("A" + row);

            PrepareTableForInvoices(ref excelRange);
            row++;
            excelRange = sheet.get_Range("A" + row);

            using (SqlConnection connection = new SqlConnection(connectionString))
            {
                using (SqlCommand command = new SqlCommand(querry, connection))
                {
                    connection.Open();
                    using (SqlDataReader reader = command.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            sql_element = 0;
                            WriteToExcelNextCellSQLAsString(reader, 0, 0, ref excelRange);
                            WriteToExcelNextCellSQLAsString(reader, 0, 1, ref excelRange);
                            WriteToExcelNextCellSQLAsString(reader, 0, 1, ref excelRange);
                            WriteToExcelNextCellSQLAsString(reader, 0, 1, ref excelRange);
                            WriteToExcelNextCellSQLAsString_NotNull(reader, 0, 1, ref excelRange);
                            WriteToExcelNextCellSQLAsInt32(reader, 0, 1, ref excelRange);
                            WriteToExcelNextCellSQLAsInt32(reader, 0, 1, ref excelRange);
                            WriteToExcelNextCellSQLAsInt32(reader, 0, 1, ref excelRange);
                            WriteToExcelNextCellSQLAsDatetime(reader, 0, 1, ref excelRange);
                            WriteToExcelNextCellSQLAsString(reader, 0, 1, ref excelRange);
                            WriteToExcelNextCellSQLAsInt32(reader, 0, 1, ref excelRange);
                            row++;
                            excelRange = sheet.get_Range("A" + row);
                        }
                        sheet.Columns.AutoFit();
                    }
                }                
            }
        }

        private void btn_FillClasificators_Click(object sender, RibbonControlEventArgs e)
        {
            Template template = new Template();
           
            Range excelRange = (Range)Globals.ThisAddIn.Application.Selection;
            int row = excelRange.Row;

            Excel.Application excelApp = (Excel.Application)Marshal.GetActiveObject("Excel.Application");
            Workbook book = excelApp.ActiveWorkbook;
            Worksheet sheet = book.ActiveSheet;

            excelRange = sheet.get_Range("A" + row);

            template.ClientName = ReadExcelNextCellAsString(0, 0, ref excelRange);
            template.ClientAddress = ReadExcelNextCellAsString(0, 1, ref excelRange) + ", " + ReadExcelNextCellAsString(0, 1, ref excelRange);
            template.ClientEntityCode = ReadExcelNextCellAsString(0, 1, ref excelRange);
            template.ClientVATCode = ReadExcelNextCellAsString(0, 1, ref excelRange);
            template.InvoiceID = ReadExcelNextCellAsInt32(0, 1, ref excelRange);
            template.InvoiceSeriesID = ReadExcelNextCellAsInt32(0, 1, ref excelRange);            
            template.InvoiceNumber = ReadExcelNextCellAsInt32(0, 1, ref excelRange);            
            template.InvoiceDate = ReadExcelNextCellAsString(0, 1, ref excelRange);
            template.InvoiceSeries = ReadExcelNextCellAsString(0, 1, ref excelRange);
            template.ClientID = ReadExcelNextCellAsInt32(0, 1, ref excelRange);

            GetAdditionalClassificators(template);
                        
            excelRange = sheet.get_Range("A" + row);
            try
            {
                WriteTextInOffsetCell(ref excelRange, 0, 1, template.ClientAddress.ToString());
                WriteTextInOffsetCell(ref excelRange, 0, 1, template.PostalCode == null ? "" : template.PostalCode.ToString());
                WriteTextInOffsetCell(ref excelRange, 0, 1, template.ClientEntityCode.ToString());
                WriteTextInOffsetCell(ref excelRange, 0, 1, template.ClientVATCode == null ? "" : template.ClientVATCode.ToString());
                WriteTextInOffsetCell(ref excelRange, 0, 2, template.InvoiceSeriesID.ToString());
                WriteTextInOffsetCell(ref excelRange, 0, 3, template.InvoiceSeries == null ? "" : template.InvoiceSeries.ToString());
                WriteTextInOffsetCell(ref excelRange, 0, 1, template.ClientID.ToString());
            }
            catch(Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }

            sheet.Columns.AutoFit();
        }

        private void GetAdditionalClassificators(Template template)
        {
            if (template.ClientID == 0)
            {
                string querry = @"Select Id, Address, PostalCode, EntityCode, VATCode from dbo.Client where Name = @Name";
                using (SqlConnection connection = new SqlConnection(connectionString))
                {
                    using (SqlCommand command = new SqlCommand(querry, connection))
                    {
                        connection.Open();
                        command.Parameters.AddWithValue("@Name", template.ClientName);
                        using (SqlDataReader reader = command.ExecuteReader())
                        {
                            sql_element = 0;
                            reader.Read();
                            template.ClientID = ReadNextSQLAsInt32(reader);
                            template.ClientAddress = ReadNextSQLAsString(reader);
                            template.PostalCode = ReadNextSQLAsString(reader);
                            template.ClientEntityCode = ReadNextSQLAsString(reader);
                            template.ClientVATCode = ReadNextSQLAsString_NotNull(reader);
                        }
                    }                    
                }
            }


            if (template.InvoiceSeriesID == 0&&template.InvoiceSeries!=null)
            {
                string querry = @"Select Series_ID from dbo.Series_ID where Title = @Series";
                using (SqlConnection connection = new SqlConnection(connectionString))
                {
                    using (SqlCommand command = new SqlCommand(querry, connection))
                    {
                        connection.Open();
                        command.Parameters.AddWithValue("@Series", template.InvoiceSeries);
                        using (SqlDataReader reader = command.ExecuteReader())
                        {
                            reader.Read();
                            template.InvoiceSeriesID = reader.GetInt32(0);
                        }
                    }                    
                }
            }
        }

        private void btnValidatePostalCode_Click(object sender, RibbonControlEventArgs e)
        {
            int row = 1;
            bool isSheetCreated = false;
            
            Excel.Application ExcelApp = (Excel.Application)Marshal.GetActiveObject("Excel.Application");
            Excel.Workbook book = ExcelApp.ActiveWorkbook;
            Excel.Worksheet sheet;

            foreach (Excel.Worksheet sh in book.Sheets)
            {
                if (sh.Name == "Clients")
                {
                    isSheetCreated = true;
                    break;
                }
            }

            if (isSheetCreated)
            {
                sheet = book.Sheets["Clients"];
                sheet.Activate();
            }
            else
            {
                sheet = book.Application.Worksheets.Add();
                sheet.Name = "Clients";
            }

            Excel.Range excelRange = sheet.get_Range("A" + row);

            int count = Count(excelRange);

            excelRange = sheet.get_Range("A" + row);
            excelRange = excelRange.Offset[0, 3];
            excelRange.Activate();

            
            CheckAllBelowPostalCode(ref excelRange, count);
        }

        private int Count(Range excelRange)
        {
            int counter = 0;
            while (excelRange.Value2 != null)
            {
                excelRange = excelRange.Offset[1, 0];
                excelRange.Activate();
                counter++;
            }
            return --counter;
        }

        private void CheckAllBelowPostalCode(ref Range excelRange, int count)
        {
            string val;
            for (int i = 0; i < count; i++)
            {
                excelRange = excelRange.Offset[1, 0];
                excelRange.Activate();
                val = Convert.ToString(excelRange.Value2);
                Parse(val, ref excelRange);
            }
        }

        private void Parse(string v, ref Range excelRange)
        {
            if (v == "" || v == null || v.Length==0)
            {
                excelRange.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red);
            }
            else if (!v.StartsWith("LT-"))
            {
                MessageBox.Show("Changing value from: " + v + " to: LT-"+v);
                excelRange.Value2 = "LT-" + v;
            }
        }
    }
}
