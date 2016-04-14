using System;
using System.Windows;

using Microsoft.Win32;

namespace SqlExport
{
    using System;
    using System.Collections.Generic;
    using System.Configuration;
    using System.Data;
    using System.Data.SqlClient;
    using System.IO;
    using System.Linq;
    using System.Net.Mail;
    using System.Net.Mime;
    using System.Threading.Tasks;

    using Microsoft.Win32;

    using SendGrid;

    using SqlExport.Core;

    public partial class MainWindow
    {
        private string fileName;

        public MainWindow()
        {
            InitializeComponent();

            this.DateFromPicker.SelectedDate = DateTime.Today;
            this.DateToPicker.SelectedDate = DateTime.Today;

            this.MailToTextBox.Text = ConfigurationManager.AppSettings["MailMessageToEmail"];
        }

        private void GenerateBtnOnClick(object sender, RoutedEventArgs e)
        {
            this.ReadyLabel.Visibility = Visibility.Collapsed;

            if (!this.DateFromPicker.SelectedDate.HasValue || !this.DateToPicker.SelectedDate.HasValue)
            {
                MessageBox.Show("Debe seleccionar las fechas para continuar");

                return;
            }

            var fromDateTime = this.DateFromPicker.SelectedDate.Value;
            var toDateTime = this.DateToPicker.SelectedDate.Value;

            this.fileName = GetFileName(fromDateTime, toDateTime);

            if (string.IsNullOrWhiteSpace(this.fileName))
            {
                MessageBox.Show("Debe seleccionar un archivo de destino");

                return;
            }

            this.ReadyLabel.Visibility = Visibility.Visible;
            this.ReadyLabel.Text = "Guardando...";

            Task.Run(
                 delegate ()
                    {
                        try
                        {
                            this.Dispatcher.Invoke(() => this.GenerateBtn.IsEnabled = false);

                            var cs = Helpers.GetConnectionString();

                            Helpers.GenerateReport(fileName, cs, fromDateTime, toDateTime);

                            this.Dispatcher.Invoke(() => this.ReadyLabel.Text = "Guardado en: " + fileName);
                        }
                        catch (Exception exception)
                        {
                            MessageBox.Show(exception.Message);

                            this.Dispatcher.Invoke(() => this.ReadyLabel.Text = "Error");
                        }
                        finally
                        {
                            this.Dispatcher.Invoke(() => this.GenerateBtn.IsEnabled = true);
                        }
                    });
        }

        private void SendMailBtnOnClick(object sender, RoutedEventArgs e)
        {
            if (string.IsNullOrWhiteSpace(this.fileName))
            {
                MessageBox.Show("Debe generar un reporte");

                return;
            }

            if (string.IsNullOrWhiteSpace(this.MailToTextBox.Text))
            {
                MessageBox.Show("Debe ingresar un destino para el mail");

                return;
            }

            this.SendMailBtn.IsEnabled = false;

            var mailTo = this.MailToTextBox.Text;

            this.ReadyLabel.Text = "Enviando mail...";

            Task.Run(
                 async delegate ()
                 {
                     try
                     {
                         await Helpers.SendMailAsync(mailTo, fileName);

                         this.Dispatcher.Invoke(() => this.ReadyLabel.Text = "Mail enviado");
                     }
                     catch (Exception exception)
                     {
                         MessageBox.Show(exception.Message);

                         this.Dispatcher.Invoke(() => this.ReadyLabel.Text = "Error");
                     }
                     finally
                     {
                         this.Dispatcher.Invoke(() => this.SendMailBtn.IsEnabled = true);
                     }
                 });
        }
        
        private static string GetFileName(DateTime fromDateTime, DateTime toDateTime)
        {
            var dlg = new SaveFileDialog
            {
                InitialDirectory = Environment.CurrentDirectory,
                FileName = Helpers.GetFileName(fromDateTime, toDateTime),
                DefaultExt = ".csv",
                Filter = "Separado por Comas (.csv)|*.csv"
            };

            var result = dlg.ShowDialog();

            return result == true ? dlg.FileName : null;
        }
    }
}