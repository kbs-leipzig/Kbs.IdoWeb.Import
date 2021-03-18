using Kbs.IdoWeb.Import.Models;
using NLog;
using OfficeOpenXml;
using System;
using System.IO;

namespace Kbs.IdoWeb.Import.Controllers
{
    class ImportController
    {
        private static readonly NLog.Logger Logger = NLog.LogManager.GetCurrentClassLogger();

        public void ImportTaxonFile(String importFilePath)
        {
            //FILE INFO
            FileInfo fileInfo = new FileInfo(importFilePath);
            ExcelPackage package = new ExcelPackage(fileInfo);
            Logger.Info(@"- Reading Taxon File: " + importFilePath);

            //INIT MODEL
            TaxonImportModel taxonImportModel = new TaxonImportModel(package.Workbook.Worksheets[3]);
            Logger.Info("- " + fileInfo.FullName);

            Logger.Info("- Starting Taxon Import");
            taxonImportModel.StartTaxonImport();

            Logger.Info("- Imported Taxons: " + taxonImportModel.taxonImportCounter + " items imported");
            Logger.Info("- Taxon Import finished.");
            Logger.Info("--");

            package.Dispose();
        }

        /**@TODO: rewrite wo filepath when tested **/
        public void BackupAndTruncate(String importFilePath)
        {
            //FILE INFO
            FileInfo fileInfo = new FileInfo(importFilePath);
            ExcelPackage package = new ExcelPackage(fileInfo);
            Logger.Info(@"- Backup of TaxonIds @ ImageTable and Truncating Taxon Tables");
            
            //@TODO Rewrite constructor
            TaxonImportModel taxonImportModel = new TaxonImportModel();
            taxonImportModel.BackupAndTruncate();
        }

        internal void ImportDescriptionFile(string importFilePath)
        {

            Logger.Info("- Reading Description File:");

            DescriptionImportModel descriptionImportModel = new DescriptionImportModel();

            //FILE INFO
            FileInfo fileInfo = new FileInfo(importFilePath);
            ExcelPackage package = new ExcelPackage(fileInfo);
            Logger.Info(fileInfo.FullName);

            //READ Merkmale+Ampel+Abb ... WS
            descriptionImportModel.worksheetCharacteristics = package.Workbook.Worksheets[1];
            descriptionImportModel.StartDescriptionImport();
            Logger.Info("- Imported DescriptionKeyGroups: " + descriptionImportModel.descKeyGroupCounter + " items imported");
            Logger.Info("- Imported DescriptionKeys: " + descriptionImportModel.descKeyCounter + " items imported");

            package.Dispose();
            Logger.Info("- Description Import finished.");
            Logger.Info("--");

        }
        internal void ImportTaxonDescriptionFile(string importFilePath)
        {

            Logger.Info("- Reading TaxonDescription File:");

            TaxonDescriptionImportModel taxdescriptionImportModel = new TaxonDescriptionImportModel();

            //FILE INFO
            FileInfo fileInfo = new FileInfo(importFilePath);
            ExcelPackage package = new ExcelPackage(fileInfo);
            Logger.Info(fileInfo.FullName);

            //READ Matrix WS
            taxdescriptionImportModel.worksheetMatrix = package.Workbook.Worksheets[0];
            taxdescriptionImportModel.StartTaxDescriptionImport();
            Logger.Info("- Imported TaxDescriptions: " + taxdescriptionImportModel.taxDescKeyCounter + " TaxonDescription items imported");
            Logger.Info("- Imported TaxDescriptions: " + taxdescriptionImportModel.descKeyCounter + " DescriptionKeys imported");

            package.Dispose();
            Logger.Info("- TaxonDescription Import finished.");

        }

        internal void ImportImageFile(string importFilePath)
        {

            Logger.Info("- Reading Image File:");

            ImageImportModel imageImportModel = new ImageImportModel(importFilePath);
            Logger.Info(importFilePath);

            //READ Merkmale+Ampel+Abb ... WS
            imageImportModel.StartImageImport();
            Logger.Info("- Imported Image Information: " + imageImportModel.imageImportCounter + " items imported");
            Logger.Info("- Imported Image Information: " + imageImportModel.imageUpdateCounter + " items updated");

            imageImportModel.excelPackage.Dispose();
            Logger.Info("- Image Import finished.");
            Logger.Info("--");

        }

        internal void ImportMeta ()
        {
            Logger.Info("- Updating Import Meta Info");

            MetaInfoModel metaModel = new MetaInfoModel();

            //READ Merkmale+Ampel+Abb ... WS
            metaModel.UpdateTaxDescChildrenInfo();
            metaModel.UpdateDkgKeyTypeInfo();
            metaModel.UpdateEdaphobaseInfo();
            metaModel.updateTaxIdsInImageAndObs();

            Logger.Info("- Import Meta finished.");
            Logger.Info("--");
        }

        internal void ImportAppData()
        {
            Logger.Info("- Updating Import AppData ..");

            AppData appModel = new AppData();

            //READ Merkmale+Ampel+Abb ... WS
            appModel.GenerateAllFiles();

            Logger.Info("- Import AppData finished.");
            Logger.Info("--");
        }

    }
}