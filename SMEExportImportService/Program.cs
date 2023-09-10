using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.ServiceModel;
using System.Threading;
using System.Configuration;
using DMS.DBConnection;
using System.Data;
using log4net;

namespace SMEExportImportService
{

    class Program
    {
        public static Timer timer;
		private static readonly ILog log = LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);
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

				log.Info("Services is started ");
				Console.WriteLine("Services is started !");
                Console.ReadLine();

                host2.Open();
                host.Close();
            }
        }
    }
}
