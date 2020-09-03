using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.ServiceModel;
using System.Threading;
using System.Configuration;
using DMS.DBConnection;
using System.Data;

namespace SMEExportImportService
{

    class Program
    {
        public static Timer timer;
        public static void Main(string[] args)
        {
            //timer = new Timer(ScanningAlphabitFolder.Callback, timer, 0, long.Parse(ConfigurationManager.AppSettings["AlphabitScanningDownloadFolder"]));

            //ScanningAlphabitFolder.Scanning(ConfigurationManager.AppSettings["alfabitPathDownload"]);
            CallServices();
            Console.Read();
        }

        static void CallServices()
        {
            using (ServiceHost host = new ServiceHost(typeof(ExportWord)))
            {
                ServiceHost host2 = new ServiceHost(typeof(UploadToCore));
                host.Open();
                host2.Open();

                Console.WriteLine("Services is started !");
                Console.ReadLine();

                host2.Open();
                host.Close();
            }
        }
    }
}
