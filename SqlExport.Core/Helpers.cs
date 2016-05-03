namespace SqlExport.Core
{
    using System;
    using System.Collections.Generic;
    using System.Configuration;
    using System.Data;
    using System.Data.SqlClient;
    using System.IO;
    using System.Linq;
    using System.Net.Mail;
    using System.Text;
    using System.Threading.Tasks;

    using SendGrid;

    public static class Helpers
    {
        public static string GetFileName(DateTime fromDateTime, DateTime toDateTime)
        {
            return $"Reporte Comidas {fromDateTime.ToShortDateString()} al {toDateTime.ToShortDateString()} ({DateTime.Now.ToString("s")}).csv"
                 .Replace("/", "-")
                 .Replace(":", "-");
        }

        public static string GetFileFullPath(DateTime fromDateTime, DateTime toDateTime)
        {
            return Path.Combine(Environment.CurrentDirectory, GetFileName(fromDateTime, toDateTime));
        }

        public static Task SendMailAsync(string[] mailTo, string fileName)
        {
            var myMessage = new SendGridMessage();
            myMessage.AddTo(mailTo);
            myMessage.From = new MailAddress(ConfigurationManager.AppSettings["MailMessageFromEmail"], ConfigurationManager.AppSettings["MailMessageFromName"]);
            myMessage.Subject = "Reporte Comidas";
            myMessage.Text = myMessage.Subject;
            myMessage.Attachments = new[]
                {
                    fileName
                };

            var transportWeb = new Web("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");
            return transportWeb.DeliverAsync(myMessage);
        }

        public static void GenerateReport(string fileName, string cs, DateTime fromDateTime, DateTime toDateTime)
        {
            using (var conn = new SqlConnection(cs))
            {
                var query = GetQuery(fromDateTime, toDateTime.AddDays(1));

                using (var command = new SqlCommand(query, conn))
                {
                    conn.Open();

                    using (var reader = command.ExecuteReader())
                    {
                        using (var outFile = File.CreateText(fileName))
                        {
                            var columnNames = GetColumnNames(reader)
                                .ToArray();

                            var numFields = columnNames.Length;

                            outFile.WriteLine(string.Join(",", columnNames));

                            if (!reader.HasRows)
                            {
                                return;
                            }

                            while (reader.Read())
                            {
                                string[] columnValues = Enumerable.Range(0, numFields)
                                    .Select(
                                        i => reader.GetValue(i)
                                                 .ToString())
                                    .Select(field => string.Concat("\"", field.Replace("\"", "\"\""), "\""))
                                    .ToArray();

                                outFile.WriteLine(string.Join(",", columnValues));
                            }
                        }
                    }
                }
            }
        }

        public static string GetConnectionString()
        {
            return ConfigurationManager.ConnectionStrings["AccessDb_Context"].ConnectionString;
        }

        public static string GetQuery(DateTime fromDateTime, DateTime toDateTime)
        {
            return $@"SELECT s.Name as Sector, 
                        p.Nombre, 
                        p.Apellido, 
                        Convert(date, [fechahora]) as Fecha, 
                        Convert(time, Min([fechahora])) as Primera, 
                        Convert(time, Max([fechahora])) as Ultima, 
                        LEFT(CONVERT(VARCHAR(10), Max([fechahora]) - Min([fechahora]), 108), 5) as Horas

                    FROM [HistoricoAccesos] ha
                        left join Personas p on ha.[PersonId] = p.ID_Persona

                        left join Sectors s on p.[SectorId] = s.Id

                    WHERE [fechahora] >= '{fromDateTime.ToString("yyyy/MM/dd")}' AND[fechahora] <= '{toDateTime.ToString("yyyy/MM/dd")}'
                    group by s.Name, p.Nombre, p.Apellido, Convert(date, [fechahora])
                    order by s.Name, p.Apellido, p.Nombre, Convert(date, [fechahora])";
        }

        public static IEnumerable<string> GetColumnNames(IDataReader reader)
        {
            foreach (DataRow row in reader.GetSchemaTable().Rows)
            {
                yield return (string)row["ColumnName"];
            }
        }
    }
}
