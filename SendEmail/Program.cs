using System;
using System.Collections.Generic;
using System.ComponentModel.Design;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;


namespace SendEmail
{
    class Program
    {
        static void Main()
        {
            SqlConnection sqlconnection = new SqlConnection(ConfigurationManager.ConnectionStrings["UMS"].ConnectionString);
            SqlCommand sqlCommand = new SqlCommand("GetAll", sqlconnection);
            SqlCommand isSend = new SqlCommand("EmailIsSend", sqlconnection);
            isSend.CommandType = CommandType.StoredProcedure;
            sqlCommand.CommandType = CommandType.StoredProcedure;
            isSend.Parameters.AddWithValue("@bool",true);
            sqlconnection.Open();
            
            var reader = sqlCommand.ExecuteReader();
            while (reader.Read())
            {
                var name = reader.GetString(1);
                var username = reader.GetString(2);
                var email = reader.GetString(3);

                if (reader.GetBoolean(4) != true) 
                {
                    Microsoft.Office.Interop.Outlook.Application application = new Microsoft.Office.Interop.Outlook.Application();
                    Microsoft.Office.Interop.Outlook.MailItem mailItem = (Microsoft.Office.Interop.Outlook.MailItem)application.CreateItem(Microsoft.Office.Interop.Outlook.OlItemType.olMailItem);
                    mailItem.To = email;
                    mailItem.Subject = "Setup Password";
                    mailItem.HTMLBody = "Hello " + name + "," + "<br>" + "<br>" + "Your login username is: " + username + "<br>"+ "Your Temporary Password is : <b>Welcome12</b>"  + "<br>"+ "Please visit the link below to Login" + "<br>" + "<a href='https://umsklb.azurewebsites.net/Security/Login'> Login</a>";
                    mailItem.Send();
                }
                
                /* Console.WriteLine(name);
                 Console.WriteLine(username);
                 Console.WriteLine(email);
                 Console.WriteLine("");*/
            }
            sqlconnection.Close();
            sqlconnection.Open();
            isSend.ExecuteReader();
            sqlconnection.Close();
            Environment.Exit(0);
        }
    }
}
