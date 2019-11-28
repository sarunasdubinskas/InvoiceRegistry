using InvoiceRegister.Objects;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace InvoiceRegister.Data
{
    class DestinationDataWriter
    {
        private string connectionString = ConfigurationManager.ConnectionStrings["FoxVoice"].ConnectionString;

        internal void InsertNewClient(Client client)
        {
            using (SqlConnection connection = new SqlConnection(connectionString))
            {

                string querry = @"INSERT INTO dbo.Client 
                               (dbo.Client.Name, dbo.Client.Address, dbo.Client.PostalCode, dbo.Client.EntityCode, dbo.Client.VATCode)
                                OUTPUT INSERTED.ID
                                VALUES (@Name,@Address,@PostalCode,@EntityCode,@VATCode)";

                using (SqlCommand command = new SqlCommand(querry, connection))
                {
                    command.Parameters.AddWithValue("@Name", client.Name);
                    command.Parameters.AddWithValue("@Address", client.Adress);
                    command.Parameters.AddWithValue("@PostalCode", client.PostalCode);
                    command.Parameters.AddWithValue("@EntityCode", client.EntityCode);
                    command.Parameters.AddWithValue("@VATCode", client.VATCode);

                    connection.Open();

                    int insertedID;

                    try
                    {
                        insertedID  = Convert.ToInt32(command.ExecuteScalar());

                        if (connection.State == ConnectionState.Open) { connection.Close(); }
                        MessageBox.Show(insertedID.ToString());
                    }
                    catch (System.Data.Common.DbException ex)
                    {
                        MessageBox.Show(ex.ToString());
                    }
                }
            }
        }

        internal void InsertNewInvoice(Template template)
        {
            using (SqlConnection connection = new SqlConnection(connectionString))
            {

                string insertQuerry = @"INSERT INTO dbo.Invoices 
                               (dbo.Series, dbo.Client_ID, dbo.Number, dbo.Date)
                                OUTPUT INSERTED.ID
                                VALUES (@Series,@Client_ID,@Number,@Date)";

                using (SqlCommand command = new SqlCommand(insertQuerry, connection))
                {
                    command.Parameters.AddWithValue("@Series", template.InvoiceSeriesID);
                    command.Parameters.AddWithValue("@Client_ID", template.ClientID);
                    command.Parameters.AddWithValue("@Number", template.InvoiceNumber);
                    command.Parameters.AddWithValue("@Date", template.InvoiceDate);

                    connection.Open();
                    
                    try
                    {
                        template.InvoiceID = Convert.ToInt32(command.ExecuteScalar());
                    }
                    catch (System.Data.Common.DbException ex)
                    {
                        MessageBox.Show(ex.ToString());
                    }

                    if (connection.State == ConnectionState.Open) { connection.Close(); }

                    string insertQuerry2 = @"INSERT INTO dbo.Files 
                               (dbo.Invoice_ID, dbo.Path_To_Invoice)
                                VALUES (@InvoiceID, @PathToInvoice)";

                    using (SqlCommand command2 = new SqlCommand(insertQuerry2, connection))
                    {
                        command2.Parameters.AddWithValue("@InvoiceID", template.InvoiceID);
                        command2.Parameters.AddWithValue("@PathToInvoice", template.PathToFile);

                        connection.Open();

                        try
                        {
                            command2.ExecuteScalar();
                        }
                        catch (System.Data.Common.DbException ex)
                        {
                            MessageBox.Show(ex.ToString());
                        }

                        if (connection.State == ConnectionState.Open) { connection.Close(); }
                    }
                }
            }
        }

        internal void UpdateClient(Client client)
        {
            using (SqlConnection connection = new SqlConnection(connectionString))
            {
                string querry = @"UPDATE dbo.Client 
                                  SET Name = @Name,
                                      Address = @Address,
                                      PostalCode = @PostalCode,
                                      EntityCode = @EntityCode,
                                      VATCode = @VATCode
                                  WHERE ID = @ID";

                using (SqlCommand command = new SqlCommand(querry,connection))
                {
                    command.Parameters.AddWithValue("@ID", client.Id);
                    command.Parameters.AddWithValue("@Name", client.Name);                    
                    command.Parameters.AddWithValue("@Address", client.Adress);
                    command.Parameters.AddWithValue("@PostalCode", client.PostalCode);
                    command.Parameters.AddWithValue("@EntityCode", client.EntityCode);
                    if (client.VATCode==null)
                    {
                        command.Parameters.AddWithValue("@VATCode", DBNull.Value);
                    }
                    else
                    {
                        command.Parameters.AddWithValue("@VATCode", client.VATCode);
                    }
                    
                    connection.Open();

                    try
                    {
                        command.ExecuteScalar();
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.ToString());
                    }
                }
            }
        }
    }
}
