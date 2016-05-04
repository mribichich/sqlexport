namespace SqlExport.Console
{
    using System;
    using System.Collections.Generic;
    using System.Configuration;
    using System.Globalization;
    using System.Linq;
    using System.Text;
    using System.Threading.Tasks;

    using CommandLine;

    using SqlExport.Core;

    public class Program
    {
        public static void Main(string[] args)
        {
            try
            {
                var result = Parser.Default.ParseArguments<Options>(args);

                var exitCode = result.MapResult(
                    options =>
                        {
                            ExecuteReport(options);

                            return 0;
                        },
                    errors => 1);

                Environment.Exit(exitCode);
            }
            catch (Exception exception)
            {
                Console.WriteLine(exception.Message);

                Environment.Exit(1);
            }
        }

        private static void ExecuteReport(Options options)
        {
            Console.WriteLine("Generando y Guardando...");

            DateTime fromDateTime;
            DateTime toDateTime;

            if (string.IsNullOrWhiteSpace(options.Date))
            {
                fromDateTime = options.FromDate;
                toDateTime = options.ToDate;
            }
            else
            {
                ////if (DateTime.TryParse(options.Date, "yyyy-MM-dd", CultureInfo.CurrentCulture, DateTimeStyles.None, ))
                if (!DateTime.TryParse(options.Date, out fromDateTime))
                {
                    switch (options.Date.ToLower())
                    {
                        case "today":
                            fromDateTime = DateTime.Today;
                            break;

                        case "yesterday":
                            fromDateTime = DateTime.Today.AddDays(-1);
                            break;

                        default:
                            Console.WriteLine("Fecha en incorrecto formato");

                            Environment.Exit(1);

                            break;
                    }
                }

                toDateTime = fromDateTime;
            }

            var fileName = Helpers.GetFileName(fromDateTime, toDateTime);

            var cs = Helpers.GetConnectionString();

            Helpers.GenerateReport(fileName, cs, fromDateTime, toDateTime);

            Console.WriteLine("Guardado en: " + fileName);

            var mailTo = ConfigurationManager.AppSettings["MailMessageToEmail"];

            Console.WriteLine($"Enviando mail a {mailTo} ...");

            var mailsTo = mailTo.Split(
                new[]
                    {
                        ";"
                    },
                StringSplitOptions.RemoveEmptyEntries)
                .Select(s => s.Trim())
                .ToArray();

            var subject = Helpers.GetEmailSubject(fromDateTime);

            Helpers.SendMailAsync(mailsTo, subject, fileName)
                .Wait();

            Console.WriteLine("Mail enviado");
        }
    }
}