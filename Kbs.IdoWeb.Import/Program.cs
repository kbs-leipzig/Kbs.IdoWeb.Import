using Kbs.IdoWeb.Import.Controllers;
using Microsoft.AspNetCore;
using Microsoft.AspNetCore.Hosting;
using Microsoft.Extensions.Configuration;
using NLog;
using System;
using System.IO;
using System.Linq;

namespace Kbs.IdoWeb.Import
{
    static class Constants
    {
        //@TODO: replace path via config
        //latest versions - to be used for import
        public const String LiImportFilePath_edit = @"importFiles/latest/Lithobiomorpha.xlsx";
        public const String ScImportFilePath_edit = @"importFiles/latest/Scolopendromorpha.xlsx";
        public const String BodenImportFilePath_edit = @"importFiles/latest/Bodentiere.xlsx";
        public const String ChiloImportFilePath_edit = @"importFiles/latest/Chilopoda-Ordnungen.xlsx";
        public const String GloImportFilePath_edit = @"importFiles/latest/Glomerida.xlsx";
        public const String GeophiloImportFilePath_edit = @"importFiles/latest/Geophilomorpha.xlsx";
        public const String IsopodaImportFilePath_edit = @"importFiles/latest/Isopoda.xlsx";
        public const String PolydesmidaImportFilePath_edit = @"importFiles/latest/Polydesmida.xlsx";
        public const String JulidaImportFilePath_edit = @"importFiles/latest/Julida.xlsx";
        public const String DiploOrdIFP = @"importFiles/latest/Diplopoda-Ordnungen.xlsx";
        public const String ChordeuIFP = @"importFiles/latest/Chordeumatida.xlsx";
        public const String importFilesPath = @"importFiles/latest";
        //not updated
        //Bilder ERLEBEN, Biotope now in Lithobiomorpha
        //public const String PicturesOnlyFilePath = @"importFiles/Bilder ERLEBEN Tiergruppen 2020-03-12.xlsx";
        //public const String BiotopesFilePathPicturesOnly = @"importFiles/Biotop Bildrechte 2020-05-22.xlsx";        
    }
    class Program
    {
        public static IWebHostBuilder CreateWebHostBuilder(string[] args) =>
            WebHost.CreateDefaultBuilder(args)
            .UseStartup<Kbs.IdoWeb.Api.Startup>();
        public static IConfigurationRoot Configuration;
        private static readonly NLog.Logger Logger = NLog.LogManager.GetCurrentClassLogger();

        static void Main(string[] args)
        {
            CreateWebHostBuilder(args);

            Logger.Info("Started Import ..");

            if (args.Length > 0)
            {
                Logger.Info("Found cli arguments");
                if (args[0] == "update-app-only")
                {
                    Logger.Info("Updating App Information only");
                    bool configSuccess = InitConfig();
                    ImportAppData();
                    Logger.Info(" Done Updating App Information.");
                }

            }
            else
            {
                try
                {
                    bool configSuccess = InitConfig();
                    _cleanUpLogs();
                    BackupAndTruncate();

                    if (configSuccess)
                    {
                        try
                        {
                            ImportTaxa();
                            //Description depending on successful Taxon?
                            ImportDescription();
                            ImportTaxonDescription();
                            ImportImages();
                            ImportMeta();
                            ImportAppData();
                        }
                        catch (Exception e)
                        {
                            Logger.Fatal(e, "Aborting due to Fatal Error In Import Process");
                            //Console.WriteLine("Error Import:");
                            //Console.WriteLine(e);
                        }
                    }
                }
                catch (Exception e)
                {
                    Logger.Error(e, "Error building Configuration");
                }
            }

            Logger.Info("###");
            Logger.Info("Import Finished");
            Logger.Info("More information at /Logs/-date-.log");
            Logger.Info("###");
            NLog.LogManager.Shutdown();
        }

        public static void BackupAndTruncate()
        {
            var importController = new ImportController();
            /**@TODO rewrite wo filepath**/
            importController.BackupAndTruncate(Constants.LiImportFilePath_edit);
        }

        public static bool InitConfig()
        {
            var builder = new ConfigurationBuilder()
                    .SetBasePath(Directory.GetCurrentDirectory())
                    .AddJsonFile("appsettings.json", optional: true, reloadOnChange: true);

            Configuration = builder.Build();
            //TODO: how to check for succesful Configuration? https://docs.microsoft.com/en-us/dotnet/api/microsoft.extensions.configuration.configurationbuilder.build?view=aspnetcore-2.2
            return Configuration.GetConnectionString("DatabaseConnection") == null ? false : true;
        }

        private static void _cleanUpLogs()
        {
            //var logPath = ${"basedir"} +"/Logs/";
            var date = DateTime.Now.AddHours(-24);
            var path = Path.Combine(Directory.GetCurrentDirectory(), "Logs");
            var logList = Directory.GetFiles(path)
             .Select(f => new FileInfo(f))
             .Where(f => f.CreationTime < date)
             .ToList();
            logList.ForEach(f => f.Delete());
        }

        public static void ImportTaxa()
        {
            var importController = new ImportController();
            DirectoryInfo di = new DirectoryInfo(Constants.importFilesPath);
            FileInfo[] files = di.GetFiles("*.xlsx");
            foreach (FileInfo file in files)
            {
                importController.ImportTaxonFile(file.FullName);
            }
        }

        public static void ImportDescription()
        {
            var importController = new ImportController();
            DirectoryInfo di = new DirectoryInfo(Constants.importFilesPath);
            FileInfo[] files = di.GetFiles("*.xlsx");
            foreach (FileInfo file in files)
            {
                importController.ImportDescriptionFile(file.FullName);
            }
        }
        public static void ImportTaxonDescription()
        {
            var importController = new ImportController();
            DirectoryInfo di = new DirectoryInfo(Constants.importFilesPath);
            FileInfo[] files = di.GetFiles("*.xlsx");
            foreach (FileInfo file in files)
            {
                importController.ImportTaxonDescriptionFile(file.FullName);
            }
        }

        public static void ImportImages()
        {
            var importController = new ImportController();
            DirectoryInfo di = new DirectoryInfo(Constants.importFilesPath);
            FileInfo[] files = di.GetFiles("*.xlsx");
            foreach (FileInfo file in files)
            {
                importController.ImportImageFile(file.FullName);
            }

        }

        public static void ImportMeta()
        {
            var importController = new ImportController();
            importController.ImportMeta();
        }

        public static void ImportAppData()
        {
            var importController = new ImportController();
            importController.ImportAppData();
        }

    }
}