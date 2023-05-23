using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using Sagede.Shared.ReportViewerControl;
using System.Collections;
using Sagede.Shared.ControlCenter.Controller;
using System.Collections.Concurrent;
using Sagede.Shared.RealTimeData.Common;
using Sagede.Shared.ReportDesigner;
using BerichtTesting;
using Sagede.Shared.RealTimeData.Common.Constants;
using Sagede.Shared.ReportingEngine;
using Sagede.Core.Tools;
using Sagede.OfficeLine.Shared;
using Sagede.OfficeLine.Engine;
using System.Security;
using Sagede.OfficeLine.Shared.ServerConfiguration;
using Sagede.Shared.ApplicationServer;
using static System.Net.Mime.MediaTypeNames;

namespace BerichtTesting
{
    /// <summary>
    /// Interaktionslogik für MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        private Mandant _mandant;
        public Mandant mandant
        {
            get
            {
                return _mandant;
            }
            set
            {
                _mandant = value;
                //NotifyPropertyChanged(nameof(mandant));
            }
        }
        //public Session ErpSession { get; set; }
        public MainWindow()
        {
            //ErpSession = Sagede.OfficeLine.Engine.ApplicationEngine.CreateSession("OLReweAbf", ApplicationToken.Abf, null, new NamePasswordCredential());


            //this.Closing += (s, e) =>
            //{
            //    ErpSession.Close();
            //    ErpSession.Dispose();
            //};
            InitializeComponent();
            //mandant = ErpSession.CreateMandant((short)1); 
        }



        private static ConcurrentQueue<ReportViewerViewModel> _outQueue;
        public static IHost Host { get; private set; }
        private void EnqueueOut(ReportViewerViewModel vm)
        {
            vm.ReportName = "rptArtikel";
            //const string procedureName = "EnqueueOut";
            try
            {
                var language = "de";
                if (Host == null) return;
                if (Host.AppInfo == null) return;
                if (Host.AppInfo.Languages != null && Host.AppInfo.Languages.Length > 0)
                {
                    language = Host.AppInfo.Languages[0];
                }

                var user = Host.AppInfo.ServerInfo.Credential.Name;
                var password = Host.AppInfo.ServerInfo.Credential.Password;
                var applicationId = Host.AppInfo.ApplicationId;
                if (_outQueue == null || vm == null) return;
                _outQueue.Enqueue(vm);
                //Wenn der Report über eine eigene Culture Verfügt diese verwenden
                var reportCulture = vm.ReportNamedParameters.TryGetItem("reportCulture");
                if (reportCulture != null)
                {
                    language = reportCulture.Value;
                }

                vm.PrintReport(vm.TempEndpoint, user, password, applicationId, vm.ReportName, vm.ReportNamedParameters, language, true, true, string.Empty);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                //LogError(ex, procedureName);
            }
        }
        private string _endpoint;
        private ReportViewerViewModel CreateViewModel()
        {
            var vm = new ReportViewerViewModel { ReportNamedParameters = new NamedParameters() };
            
            return vm;
        }

        private string _Password = "Admin1234";
        public string Password
        {
            get
            {
                return _Password;
            }
            set
            {
                _Password = value;
            }
        }

        //private void Button_Click(object sender, RoutedEventArgs e)
        //{
        //    dynamic app = Marshal.GetActiveObject("Access.Application");
        //    app.Run("gbOpenDataEditPart", "ediArtikelstamm.170013662.berichtdrucken");
        //    app = null;

        //}


        //private void ExportReport(ReportViewerViewModel vm)
        //{
        //    try
        //    {
        //        //if (vm == null) return;
        //        //_activePrintingVms.Add(vm);
        //        //Log("ExportReport", "Bericht {0} wird exportiert.", CboReports.Text);

        //        vm.ReportName = "rptArtikel";
        //        vm.WaitForPrintCommand = true; //Bei Druck und Export immer true!

        //        var caption = vm.ReportName;
        //        var captionParameter = vm.ReportNamedParameters.TryGetItem(ReportingConstants.ParameterNameAccessCaption);
        //        if (captionParameter != null && !string.IsNullOrEmpty(captionParameter.Value))
        //            caption = captionParameter.Value;

        //        var dlg = new Microsoft.Win32.SaveFileDialog
        //        {
        //            CheckFileExists = false,
        //            ValidateNames = true,
        //            AddExtension = true,
        //            FileName = Constants.RemoveInvalidFileNameChars(caption),
        //            DefaultExt = ".pdf",
        //            OverwritePrompt = true,
        //        };

        //        //Achtung: Es funktionieren nicht alle Formate
        //        var extensions = new Dictionary<string, string>
        //        {
        //            {"pdf", "Pdf-Dokument"},
        //            {"csv", "Csv-Dokument"},
        //            {"docx", "Word2007-Dokument"},
        //            {"xlsx", "Excel2007-Dokument"},
        //            {"pptx", "Powerpoint2007-Dokument"},
        //            {"html", "Html-Dokument"},
        //            {"html5", "Html5-Dokument"},
        //            {"htmldiv", "HtmlDiv-Dokument"},
        //            {"htmlspan", "HtmlSpan-Dokument"},
        //            {"htmltable", "HtmlTable-Dokument"},
        //            {"bmp", "ImageBmp-Dokument"},
        //            {"emf", "ImageEmf-Dokument"},
        //            {"gif", "ImageGif-Dokument"},
        //            {"jpg", "ImageJpeg-Dokument"},
        //            {"pcx", "ImagePcx-Dokument"},
        //            {"png", "ImagePng-Dokument"},
        //            {"svg", "ImageSvg-Dokument"},
        //            {"svgz", "ImageSvgz-Dokument"},
        //            {"tif", "ImageTiff-Dokument"},
        //            {"tiff", "ImageTiff-Dokument"},
        //            {"mht", "Mht-Dokument"},
        //            {"ods", "Ods-Dokument"},
        //            {"odt", "Odt-Dokument"},
        //            {"rtf", "Rtf-Dokument"},
        //            {"rtfframe", "RtfFrame-Dokument"},
        //            {"rtftabbedtext", "RtfTabbedText-Dokument"},
        //            {"rtftable", "RtfTable-Dokument"},
        //            {"rtfwinword", "RtfWinWord-Dokument"},
        //            {"dat", "Data-Dokument"},
        //            {"dbf", "Dbf-Dokument"},
        //            {"dif", "Dif-Dokument"},
        //            {"txt", "Text-Dokument"},
        //            {"xml", "Xml-Dokument"},
        //            {"xps", "Xps-Dokument"}
        //        };
        //        string DataBase = "OLReweAbf";
        //        IEndpoint MyEndpoint = ServerConfigurationProxy.GetSDataHttpsEndpoint(true, 0, 0);
        //        string _endpoint = string.Join("/", MyEndpoint.Address, "ol", "controlcenterdata",
        //                DataBase + ";" + mandant);
        //        var filter = extensions.Aggregate(string.Empty, (current, extension) =>
        //            current + ("|" + extension.Value + " (." + extension.Key + ")|*." + extension.Key));
        //        filter = filter.Trim('|');
        //        dlg.Filter = filter;

               
        //        //Parameter siehe PrintReport
        //        vm.PrintReport(_endpoint, Properties.Settings.Default.UserName, Global.SecurePassword,
        //                        Properties.Settings.Default.ApplicationId, vm.ReportName, vm.ReportNamedParameters,
        //                        Properties.Settings.Default.Language, true, true, string.Empty);
        //    }
        //    catch (Exception ex)
        //    {
        //        MessageBox.Show(ex.Message);
        //    }
        //}
        private NamedParameters _namedParameters;
        //private void Button_Click(object sender, RoutedEventArgs e)
        //{
        //    var vm = CreateViewModel();
        //    ExportReport(vm);
        //}


        private void Button_Click(object sender, RoutedEventArgs e)
        {
            string DataBase = "OLReweAbf";
            string UserName = "Pandev";
            string ApplicationId = "Abf";
            string sprache = "de";
            string password = "Admin1234";
            IEndpoint MyEndpoint = ServerConfigurationProxy.GetSDataHttpsEndpoint(true, 0, 0);

            var vm = CreateViewModel();
            string _endpoint = String.Join("/", MyEndpoint.Address, "ol", "ControlCenterData",
                    DataBase + ";" + mandant);
            try
            {
                if (vm == null) return;

                Global.Password = inputPasswort.Text;
                Properties.Settings.Default.UserName = inputUser.Text;
                vm.ReportName = "rptTester.170013662.berichtdrucken";
                vm.WaitForPrintCommand = false; //Bei Druck und Export immer true!
                
                // <summary>
                // Druckt einen Bericht
                // </summary>
                // <param name="endpoint">Endpunktdes Applikations-Servers</param>
                // <param name="user">Benutzer</param>
                // <param name="password">Kennwort</param>
                // <param name="applicationId">Applikation-Token: Abf, Rewe, AppDesigner</param>
                // <param name="key">Name des Berichts mit Partnerkennung und Lösung</param>
                // <param name="namedParameters">Parameter zur Ausführung des Berichts</param>
                // <param name="language">Strache für fremdsprachige Office Line, nicht Kunden/Lieferanten spezifische Sprache</param>
                // <param name="isAsync">true = "Der Berichtsdruck erfolgt asynchron"</param>
                // <param name="isInitial">true = initialer Aufruf, für externe Aufrufe immer null</param>
                // <param name="trackingId">Tracking-Id zur Nachverfolgung, für externe Aufrufe immer leer. Bei leer wird sie automatisch vergeben.</param>
                vm.PrintReport(_endpoint, Properties.Settings.Default.UserName, Global.SecurePassword,
                                Properties.Settings.Default.ApplicationId, vm.ReportName, vm.ReportNamedParameters,
                                Properties.Settings.Default.Language, true, true, string.Empty);
                //Alternativ:
                //mandant.ApplicationStateInfo.



                //var factory = new TaskFactory(TaskCreationOptions.LongRunning, TaskContinuationOptions.None);
                //factory.StartNew(() => vm.PrintReport(_endpoint, UserName, Global.SecurePassword,
                //                ApplicationId, vm.ReportName, vm.ReportNamedParameters,
                //                sprache, true, true, string.Empty));
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            //dynamic app = Marshal.GetActiveObject("Access.Application");
            //app.Run("gbOpenDataEditPart", "ediLieferantenSuche.170013662.Wawi");
            //app = null;
            //dynamic app = Marshal.GetActiveObject("Access.Application");
            //app.Run("gbRptOpenForm", "rptartikel.170013662.berichtdrucken");
            //app = null;

            // StiReport report = new StiReport();
            //// Optional: Setzen Sie Parameterwerte für den Bericht
            ////report.Dictionary.Variables["ParameterName"].Value = "ParameterValue";
            //report.Load(@"C:\Program Files (x86)\Sage\Sage 100\9.0\Shared\Metadata\Data\ReportParts\rptTester.170013662.berichtdrucken.xml");
            //report.ShowWithWpf();

            //report.RegData("rptTester", sConnection, "rptTester");
            //report.Print(true);



            //report.dictionary.databases.clear();
            //report.dictionary.databases.add(new stimulsoft.report.dictionary.stioledbdatabase(
            //    "tomsql01",
            //    "provider=olrewabf",
            //    "server=tomsql01;initial catalog=olreweabf;;integrated security=true"
            //));

            //report.Print(true);

        }
    }
}
