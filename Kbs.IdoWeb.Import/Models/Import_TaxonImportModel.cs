using Kbs.IdoWeb.Data.Information;
using Microsoft.EntityFrameworkCore;
using Microsoft.Extensions.Configuration;
using Newtonsoft.Json;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Data.SqlTypes;
using System.Linq;
using System.Security.Cryptography.X509Certificates;
using System.Text.RegularExpressions;

namespace Kbs.IdoWeb.Import.Models
{
    static class Constants
    {
        /** Get stateIds From DB? **/
        public const int stateId_kingdom = 124;
        public const int stateId_phylum = 123;
        public const int stateId_subphylum = 122;
        public const int stateId_class = 119;
        public const int stateId_subclass = 125;
        public const int stateId_order = 117;
        public const int stateId_suborder = 120;
        public const int stateId_family = 100;
        public const int stateId_subfamily = 101;
        public const int stateId_genus = 200;
        public const int stateId_species = 301;
        public const string _kingdomName = "Animalia";
        public const string _phylumName = "Gliederfüßer (Arthropoda)";
    }

    class TaxonImportModel
    {
        private static readonly NLog.Logger Logger = NLog.LogManager.GetCurrentClassLogger();
        public ExcelWorksheet worksheetGeneral;
        private Dictionary<string, int> _uniqueClassesDict = new Dictionary<string, int>();
        private Dictionary<string, int> _uniqueOrdersDict = new Dictionary<string, int>();
        private Dictionary<string, int> _uniqueFamiliesDict = new Dictionary<string, int>();
        private Dictionary<string, int> _uniqueGenereaDict = new Dictionary<string, int>();
        private Dictionary<string, int> _uniqueSpeciesDict = new Dictionary<string, int>();

        private InformationContext _infContext = new InformationContext();
        private Kbs.IdoWeb.Data.Determination.DeterminationContext _detContext = new Data.Determination.DeterminationContext();
        /** CURRENTLY KINGDOM NOT IN EXPORT FILE **/
        private int _kingdomId;
        private int _phylumId_arthropoda;
        public int taxonImportCounter = 0;
        //excel meta: column indices
        private int taxonCol;
        private int subfamilyCol;
        private int familyCol;
        private int suborderCol;
        private int orderCol;
        private int subclassCol;
        private int classCol;
        private int subphylumCol;
        private int phylumCol;
        private int distributionCol;
        private int biotopeCol;
        private int groupCol = 0;
        private int edaphoCol = 0;
        private int redListTypeCol = 0;
        private int redListSourceCol = 0;
        private int litSourceCol = 0;
        private int distEuropeCol = 0;
        private int diagnosisCol = 0;
        private int addInfoCol = 0;
        private int dispLengthCol = 0;
        private int sliderImageCol = 0;
        private int i18nNamesCol = 0;
        private int guidCol = 0;
        Dictionary<string, int> _redListMap;
        private const string _taxonRegex = @"^(?:\S[^\(\)]*\s\S[^\(\)]*\s)((?:\(?)(\S*|\S*\s\S*|\S*|\S{2}\s\S{2}\s\S*|\S{2,4}\s\S|\S\s\S\s\S)(?:\,\s?\d{4}\)?))$";
        private const string _taxonBracketRegex = @"^(?:\S*\s\S*\s)(\()(.*?)(?:(?:\,?\s\d{4}?\)?)|(\,\s\))|(\)))$";
        private const string _taxonNameExtractRegex = @"^(?:\S*\s\S*\s)(?:\(?)((\S*[^\,]\s*){1,4})(?:\,\s\d{0,4}\)?)$";
        public TaxonImportModel()
        {
            InitInformationContext();
        }

        public TaxonImportModel(ExcelWorksheet worksheetInstance)
        {
            try
            {
                worksheetGeneral = worksheetInstance;
                InitProperties();
                InitInformationContext();
                SetKingdomId();
                SetPhylumId();
                FillUniqueTaxonDicts();
            }
            catch (Exception e)
            {
                Logger.Error(e, "Error at Constructor of TaxonImportModel");
            }

        }

        public void InitProperties()
        {
            for (int i = 1; i <= worksheetGeneral.Dimension.Columns; i++)
            {
                if (worksheetGeneral.Cells[1, i].Value != null)
                {
                    var cellVal = worksheetGeneral.Cells[1, i].Value?.ToString().Trim();
                    switch (cellVal)
                    {
                        case "Taxon":
                            taxonCol = i;
                            break;
                        case "Unterfamilie":
                            subfamilyCol = i;
                            break;
                        case "Familie":
                            familyCol = i;
                            break;
                        case "Unterordnung":
                            suborderCol = i;
                            break;
                        case "Ordnung":
                            orderCol = i;
                            break;
                        case "Unterklasse":
                            subclassCol = i;
                            break;
                        case "Klasse":
                            classCol = i;
                            break;
                        case "Unterstamm":
                            subphylumCol = i;
                            break;
                        case "Stamm":
                            phylumCol = i;
                            break;
                        case "Verbreitung & Häufigkeit":
                            distributionCol = i;
                            break;
                        case "Lebensräume & Lebensweise":
                            biotopeCol = i;
                            break;
                        case "Gruppe":
                            groupCol = i;
                            break;
                        case "Edaphobase ID":
                            edaphoCol = i;
                            break;
                        case "Diagnose & ähnliche Arten":
                            diagnosisCol = i;
                            break;
                        case "Rote Liste D":
                            redListTypeCol = i;
                            break;
                        case "Rote Liste Quelle":
                            redListSourceCol = i;
                            break;
                        case "Verbreitung Europa":
                            distEuropeCol = i;
                            break;
                        case "Merkmale":
                            addInfoCol = i;
                            break;
                        case "Literatur":
                            litSourceCol = i;
                            break;
                        case "Körperlänge":
                            dispLengthCol = i;
                            break;
                        case "Slider":
                            sliderImageCol = i;
                            break;
                        case "Trivialname":
                            i18nNamesCol = i;
                            break;
                        case "GUID":
                            guidCol = i;
                            break;
                        default:
                            break;
                    }
                }
            }
        }

        public void FillUniqueTaxonDicts()
        {
            var classList = _infContext.Taxon.Where(s => s.TaxonomyStateId == Constants.stateId_class).Distinct().ToList();
            _uniqueClassesDict = classList.ToDictionary(item => item.TaxonName, item => item.TaxonId);
            var orderList = _infContext.Taxon.Where(s => s.TaxonomyStateId == Constants.stateId_order).Distinct().ToList();
            _uniqueOrdersDict = orderList.ToDictionary(item => item.TaxonName, item => item.TaxonId);
            var familyList = _infContext.Taxon.Where(s => s.TaxonomyStateId == Constants.stateId_family).Distinct().ToList();
            _uniqueFamiliesDict = familyList.ToDictionary(item => item.TaxonName, item => item.TaxonId);
            var genusList = _infContext.Taxon.Where(s => s.TaxonomyStateId == Constants.stateId_genus).Distinct().ToList();
            _uniqueGenereaDict = genusList.ToDictionary(item => item.TaxonName, item => item.TaxonId);
            var speciesList = _infContext.Taxon.Where(s => s.TaxonomyStateId == Constants.stateId_species).Distinct().ToList();
            _uniqueSpeciesDict = speciesList.ToDictionary(item => item.TaxonName, item => item.TaxonId);
            _redListMap = _infContext.RedListType.ToDictionary(item => item.RedListTypeName, item => item.RedListTypeId);
        }

        private void IncrementCounter()
        {
            taxonImportCounter++;
        }

        public void BackupAndTruncate()
        {

            _infContext.Database.ExecuteSqlCommand("TRUNCATE TABLE \"Inf\".\"Taxon\" CASCADE;");
            Logger.Info("Truncated Inf.Taxon");
            _infContext.SaveChanges();
            _detContext.Database.ExecuteSqlCommand("TRUNCATE TABLE \"Det\".\"DescriptionKeyGroup\" CASCADE;");
            Logger.Info("Truncated Det.DescriptionKeyGroup");
            _detContext.Database.ExecuteSqlCommand("TRUNCATE TABLE \"Det\".\"DescriptionKey\" CASCADE;");
            Logger.Info("Truncated Det.DescriptionKey");
            _detContext.Database.ExecuteSqlCommand("TRUNCATE TABLE \"Det\".\"TaxonDescription\" CASCADE;");
            Logger.Info("Truncated Det.TaxonDescription");
            _detContext.SaveChanges();
            //_infContext.Database.ExecuteSqlCommand(string.Format("DELETE FROM {0}", "[Taxon]"));
        }

        public void SetKingdomId()
        {
            //KINGDOM IN DB?
            Taxon kingdomTaxon = _infContext.Taxon.Where(s => s.TaxonName == Constants._kingdomName).FirstOrDefault();
            //no kingdom yet in DB, CREATE NEW
            if (kingdomTaxon == null)
            {
                try
                {
                    kingdomTaxon = new Taxon
                    {
                        TaxonName = Constants._kingdomName,
                        TaxonomyStateId = Constants.stateId_kingdom,
                        Group = "regular",
                    };
                    if (!_checkTaxonExists(kingdomTaxon.TaxonName))
                    {
                        _infContext.Taxon.Add(kingdomTaxon);
                        SaveToContext();
                        Logger.Info("-- Added Kingdom Taxon " + kingdomTaxon.TaxonName);
                        IncrementCounter();
                    }

                }
                catch (Exception e)
                {
                    Logger.Error(e, "Could not Save Kingdom Taxon");
                }
            }
            _kingdomId = kingdomTaxon.TaxonId;
        }

        public void SetPhylumId()
        {
            //phylum IN DB?
            Taxon phylumTaxon = _infContext.Taxon.Where(s => s.TaxonName == Constants._phylumName).FirstOrDefault();

            //no phylum yet in DB, CREATE NEW
            if (phylumTaxon == null)
            {
                try
                {
                    phylumTaxon = new Taxon
                    {
                        TaxonName = Constants._phylumName,
                        KingdomId = _kingdomId,
                        TaxonomyStateId = Constants.stateId_phylum,
                        Group = "regular"
                    };
                    if (!_checkTaxonExists(phylumTaxon.TaxonName))
                    {
                        _infContext.Taxon.Add(phylumTaxon);
                        SaveToContext();
                        IncrementCounter();
                        Logger.Info("-- Added Phylum Taxon " + phylumTaxon.TaxonName);
                    }
                }
                catch (Exception e)
                {
                    Logger.Error(e, "Could not Save Phylum Taxon");
                }
            }
            _phylumId_arthropoda = phylumTaxon.TaxonId;
        }

        private void SaveToContext()
        {
            try
            {
                _infContext.SaveChanges();
            }
            catch (Exception e)
            {
                Logger.Error(e.InnerException, "Error Saving Changes to Context");
            }
        }

        public void InitInformationContext()
        {
            try
            {
                var optionsBuilderInf = new DbContextOptionsBuilder<InformationContext>();
                optionsBuilderInf.UseNpgsql(Program.Configuration.GetConnectionString("DatabaseConnection"));
                _infContext = new InformationContext(optionsBuilderInf.Options);
                Logger.Debug("Init InformationContext successful");

                var optionsBuilderDet = new DbContextOptionsBuilder<Kbs.IdoWeb.Data.Determination.DeterminationContext>();
                optionsBuilderDet.UseNpgsql(Program.Configuration.GetConnectionString("DatabaseConnection"));
                _detContext = new Kbs.IdoWeb.Data.Determination.DeterminationContext(optionsBuilderDet.Options);
                Logger.Debug("Init DetContext successful");
            }
            catch (Exception e)
            {
                Logger.Error(e, "Could not init InformationContext");
            }
        }

        public void StartTaxonImport()
        {
            //Import Taxons from top to bottom hierarchy and add new ones to dictionaries
            ImportPhylum();
            ImportSubPhylums();
            ImportClasses();
            ImportSubclasses();
            ImportOrders();
            ImportSubOrders();
            ImportFamilies();
            ImportSubFamilies();
            ImportGenerea();
            ImportSpecies();
        }

        private string? _getGroupInfo(int row)
        {
            if (groupCol != 0)
            {
                if (worksheetGeneral.Cells[row, groupCol].Value != null)
                {
                    return worksheetGeneral.Cells[row, groupCol].Value.ToString().Trim();
                }
            }
            return "regular";
        }

        private void ImportSubPhylums()
        {
            if (subphylumCol != 0)
            {
                for (int i = 2; i <= worksheetGeneral.Dimension.Rows; i++)
                {
                    if (worksheetGeneral.Cells[i, subphylumCol].Value != null)
                    {
                        var subphylumName = worksheetGeneral.Cells[i, subphylumCol].Value.ToString().Trim();
                        Taxon subphylumTaxon = new Taxon();
                        subphylumTaxon.TaxonName = subphylumName;
                        subphylumTaxon.TaxonomyStateId = Constants.stateId_subphylum;
                        subphylumTaxon.PhylumId = _getPhylumIdFromExcel(i);
                        subphylumTaxon.KingdomId = _kingdomId;
                        subphylumTaxon.Group = _getGroupInfo(i);
                        if (!_checkTaxonExists(subphylumTaxon.TaxonName))
                        {
                            _infContext.Taxon.Add(subphylumTaxon);
                            _infContext.SaveChanges();
                        }
                        else
                        {
                            var alreadyTax = _infContext.Taxon.Where(tax => tax.TaxonName == subphylumTaxon.TaxonName).FirstOrDefault();
                            if (!alreadyTax.Group.Contains(subphylumTaxon.Group))
                            {
                                alreadyTax.Group += "," + subphylumTaxon.Group;
                                _infContext.Update(alreadyTax);
                            }
                        }
                    }
                }
            }
        }

        private void ImportSubFamilies()
        {
            if (subfamilyCol != 0)
            {
                for (int i = 2; i <= worksheetGeneral.Dimension.Rows; i++)
                {
                    if (worksheetGeneral.Cells[i, subfamilyCol].Value != null)
                    {
                        var subfamilyName = worksheetGeneral.Cells[i, subfamilyCol].Value.ToString().Trim();
                        Taxon newTaxon = new Taxon();
                        newTaxon.TaxonName = subfamilyName;
                        newTaxon.TaxonomyStateId = Constants.stateId_subfamily;
                        newTaxon.FamilyId = _getFamilyIdFromExcel(i);
                        newTaxon.SubphylumId = _getSubPhylumIdFromExcel(i);
                        newTaxon.PhylumId = _getPhylumIdFromExcel(i);
                        newTaxon.Group = _getGroupInfo(i);
                        newTaxon.KingdomId = _kingdomId;
                        if (!_checkTaxonExists(newTaxon.TaxonName))
                        {
                            _infContext.Taxon.Add(newTaxon);
                            _infContext.SaveChanges();
                        }
                        else
                        {
                            var alreadyTax = _infContext.Taxon.Where(tax => tax.TaxonName == newTaxon.TaxonName).FirstOrDefault();
                            if (!alreadyTax.Group.Contains(newTaxon.Group))
                            {
                                alreadyTax.Group += "," + newTaxon.Group;
                                _infContext.Update(alreadyTax);
                            }
                        }
                    }
                }
            }
        }

        private void ImportSubOrders()
        {
            if (suborderCol != 0)
            {
                for (int i = 2; i <= worksheetGeneral.Dimension.Rows; i++)
                {
                    if (worksheetGeneral.Cells[i, suborderCol].Value != null)
                    {
                        var suborderName = worksheetGeneral.Cells[i, suborderCol].Value.ToString().Trim();
                        Taxon newTaxon = new Taxon();
                        newTaxon.TaxonName = suborderName;
                        newTaxon.TaxonomyStateId = Constants.stateId_suborder;
                        newTaxon.OrderId = _getOrderIdFromExcel(i);
                        newTaxon.SubphylumId = _getSubPhylumIdFromExcel(i);
                        newTaxon.PhylumId = _getPhylumIdFromExcel(i);
                        newTaxon.Group = _getGroupInfo(i);
                        newTaxon.KingdomId = _kingdomId;
                        if (!_checkTaxonExists(newTaxon.TaxonName))
                        {
                            _infContext.Taxon.Add(newTaxon);
                            _infContext.SaveChanges();
                        }
                        else
                        {
                            var alreadyTax = _infContext.Taxon.Where(tax => tax.TaxonName == newTaxon.TaxonName).FirstOrDefault();
                            if (!alreadyTax.Group.Contains(newTaxon.Group))
                            {
                                alreadyTax.Group += "," + newTaxon.Group;
                                _infContext.Update(alreadyTax);
                            }
                        }
                    }
                }
            }
        }

        private void ImportSubclasses()
        {
            //Phylum column exists?
            if (subclassCol != 0)
            {
                for (int i = 2; i <= worksheetGeneral.Dimension.Rows; i++)
                {
                    if (worksheetGeneral.Cells[i, subclassCol].Value != null)
                    {
                        var taxonName = worksheetGeneral.Cells[i, subclassCol].Value.ToString().Trim();
                        Taxon newTaxon = new Taxon();
                        newTaxon.TaxonName = taxonName;
                        newTaxon.TaxonomyStateId = Constants.stateId_subclass;
                        newTaxon.ClassId = _getClassIdFromExcel(i);
                        newTaxon.SubphylumId = _getSubPhylumIdFromExcel(i);
                        newTaxon.PhylumId = _getPhylumIdFromExcel(i);
                        newTaxon.Group = _getGroupInfo(i);
                        newTaxon.KingdomId = _kingdomId;
                        if (!_checkTaxonExists(newTaxon.TaxonName))
                        {
                            _infContext.Taxon.Add(newTaxon);
                            _infContext.SaveChanges();
                        }
                        else
                        {
                            var alreadyTax = _infContext.Taxon.Where(tax => tax.TaxonName == newTaxon.TaxonName).FirstOrDefault();
                            if (!alreadyTax.Group.Contains(newTaxon.Group))
                            {
                                alreadyTax.Group += "," + newTaxon.Group;
                                _infContext.Update(alreadyTax);
                            }
                        }
                    }
                }
            }
        }

        public void ImportPhylum()
        {
            //Phylum column exists?
            if (phylumCol != 0)
            {
                for (int i = 2; i <= worksheetGeneral.Dimension.Rows; i++)
                {
                    if (worksheetGeneral.Cells[i, phylumCol].Value != null)
                    {
                        var phylumName = worksheetGeneral.Cells[i, phylumCol].Value.ToString().Trim();
                        Taxon phylumTaxon = new Taxon();
                        phylumTaxon.TaxonName = phylumName;
                        phylumTaxon.TaxonomyStateId = Constants.stateId_phylum;
                        phylumTaxon.KingdomId = _kingdomId;
                        phylumTaxon.Group = _getGroupInfo(i);
                        if (!_checkTaxonExists(phylumTaxon.TaxonName) && phylumTaxon.KingdomId != null)
                        {
                            _infContext.Taxon.Add(phylumTaxon);
                            _infContext.SaveChanges();
                        }
                        else
                        {
                            var alreadyTax = _infContext.Taxon.Where(tax => tax.TaxonName == phylumTaxon.TaxonName).FirstOrDefault();
                            if (!alreadyTax.Group.Contains(phylumTaxon.Group))
                            {
                                alreadyTax.Group += "," + phylumTaxon.Group;
                                _infContext.Update(alreadyTax);
                            }
                        }
                    }
                }
            }
        }

        public bool _checkTaxonExists(string taxonName)
        {
            var result = _infContext.Taxon.Where(tax => tax.TaxonName == taxonName).Select(tax => tax.TaxonId).FirstOrDefault();
            return result != 0 ? true : false;
        }

        /**TODO: refactor below **/
        public void ImportClasses()
        {
            if (classCol != 0)
            {
                for (int i = 2; i <= worksheetGeneral.Dimension.Rows; i++)
                {
                    if (worksheetGeneral.Cells[i, classCol].Value != null)
                    {
                        try
                        {
                            Taxon newTaxon = new Taxon
                            {
                                TaxonName = worksheetGeneral.Cells[i, classCol].Value?.ToString().Trim(),
                                TaxonomyStateId = Constants.stateId_class,
                                KingdomId = _kingdomId,
                                SubphylumId = _getSubPhylumIdFromExcel(i),
                                PhylumId = _getPhylumIdFromExcel(i),
                                Group = _getGroupInfo(i),
                            };
                            if (!_checkTaxonExists(newTaxon.TaxonName))
                            {
                                //hack because client cant keep his excels clean
                                 if(newTaxon.PhylumId != null)
                                {
                                    newTaxon.PhylumId = _infContext.Taxon.Where(tax => tax.TaxonName == "Gliederfüßer (Arthropoda)").Select(tax => tax.TaxonId).FirstOrDefault();
                                }

                                var newTaxonId = _infContext.Taxon.Add(newTaxon);
                                SaveToContext();
                                Logger.Debug("--- Added Taxon " + newTaxon.TaxonName);
                                IncrementCounter();
                            }
                            else
                            {
                                var alreadyTax = _infContext.Taxon.Where(tax => tax.TaxonName == newTaxon.TaxonName).FirstOrDefault();
                                if (!alreadyTax.Group.Contains(newTaxon.Group))
                                {
                                    alreadyTax.Group += "," + newTaxon.Group;
                                    _infContext.Update(alreadyTax);
                                }
                            }
                        }
                        catch (Exception e)
                        {
                            Logger.Error(e, "Error adding new Class Taxon " + worksheetGeneral.Cells[i, classCol].Value + " to Context");
                        }
                    }
                }
                Logger.Info("-- Finished Import of Classes");
            }
        }

        public int? _getPhylumIdFromExcel(int i)
        {
            if (phylumCol > 0)
            {
                var phylumName = worksheetGeneral.Cells[i, phylumCol].Value?.ToString().Trim();
                if (phylumName != null)
                {
                    int result = _infContext.Taxon.Where(tax => tax.TaxonName == phylumName).Select(tax => tax.TaxonId).FirstOrDefault();
                    if (result > 0)
                    {
                        return result;
                    }
                }
            }
            return null;
        }

        public int? _getSubPhylumIdFromExcel(int i)
        {
            if (subphylumCol > 0)
            {
                var subphylumName = worksheetGeneral.Cells[i, subphylumCol].Value?.ToString().Trim();
                if (subphylumName != null)
                {
                    int result = _infContext.Taxon.Where(tax => tax.TaxonName == subphylumName).Select(tax => tax.TaxonId).FirstOrDefault();
                    if (result > 0)
                    {
                        return result;
                    }

                }
            }
            return null;
        }

        public int? _getFamilyIdFromExcel(int i)
        {
            if (familyCol > 0)
            {
                var famName = worksheetGeneral.Cells[i, familyCol].Value?.ToString().Trim();
                if (famName != null)
                {
                    int result = _infContext.Taxon.Where(tax => tax.TaxonName == famName).Select(tax => tax.TaxonId).FirstOrDefault();
                    if (result > 0)
                    {
                        return result;
                    }

                }
            }
            return null;
        }

        public int? _getSubFamilyIdFromExcel(int i)
        {
            if (subfamilyCol > 0)
            {
                var subfamName = worksheetGeneral.Cells[i, subfamilyCol].Value?.ToString().Trim();
                if (subfamName != null)
                {
                    int result = _infContext.Taxon.Where(tax => tax.TaxonName == subfamName).Select(tax => tax.TaxonId).FirstOrDefault();
                    if (result > 0)
                    {
                        return result;
                    }

                }
            }
            return null;
        }

        public int? _getClassIdFromExcel(int i)
        {
            if (classCol > 0)
            {
                var className = worksheetGeneral.Cells[i, classCol].Value?.ToString().Trim();
                if (className != null)
                {
                    int result = _infContext.Taxon.Where(tax => tax.TaxonName == className).Select(tax => tax.TaxonId).FirstOrDefault();
                    if (result > 0)
                    {
                        return result;
                    }

                }
            }
            return null;
        }

        public int? _getSubClassIdFromExcel(int i)
        {
            if (subclassCol > 0)
            {
                var subclassName = worksheetGeneral.Cells[i, subclassCol].Value?.ToString().Trim();
                if (subclassName != null)
                {
                    int result = _infContext.Taxon.Where(tax => tax.TaxonName == subclassName).Select(tax => tax.TaxonId).FirstOrDefault();
                    if (result > 0)
                    {
                        return result;
                    }

                }
            }
            return null;
        }

        public int? _getOrderIdFromExcel(int i)
        {
            if (orderCol > 0)
            {
                var orderName = worksheetGeneral.Cells[i, orderCol].Value?.ToString().Trim();
                if (orderName != null)
                {
                    int result = _infContext.Taxon.Where(tax => tax.TaxonName == orderName).Select(tax => tax.TaxonId).FirstOrDefault();
                    if (result > 0)
                    {
                        return result;
                    }

                }
            }
            return null;
        }

        public int? _getSubOrderIdFromExcel(int i)
        {
            if (suborderCol > 0)
            {
                var suborderName = worksheetGeneral.Cells[i, suborderCol].Value?.ToString().Trim();
                if (suborderName != null)
                {
                    int result = _infContext.Taxon.Where(tax => tax.TaxonName == suborderName).Select(tax => tax.TaxonId).FirstOrDefault();
                    if (result > 0)
                    {
                        return result;
                    }
                }
            }
            return null;
        }

        public void ImportOrders()
        {
            //GET ALL FROM EXCEL
            for (int i = 2; i <= worksheetGeneral.Dimension.Rows; i++)
            {
                if (worksheetGeneral.Cells[i, orderCol].Value != null)
                {
                    string orderName = worksheetGeneral.Cells[i, orderCol].Value.ToString().Trim();
                    string className = worksheetGeneral.Cells[i, classCol].Value.ToString().Trim();
                    try
                    {
                        Taxon newTaxon = new Taxon
                        {
                            TaxonName = worksheetGeneral.Cells[i, orderCol].Value.ToString().Trim(),
                            TaxonomyStateId = Constants.stateId_order,
                            KingdomId = _kingdomId,
                            PhylumId = _getPhylumIdFromExcel(i),
                            SubphylumId = _getSubPhylumIdFromExcel(i),
                            ClassId = _getClassIdFromExcel(i),
                            SubclassId = _getSubClassIdFromExcel(i),
                            Group = _getGroupInfo(i)
                        };
                        if (!_checkTaxonExists(newTaxon.TaxonName))
                        {
                            if (newTaxon.PhylumId != null && newTaxon.ClassId != null)
                            {
                                newTaxon.PhylumId = _infContext.Taxon.Where(tax => tax.TaxonName == "Gliederfüßer (Arthropoda)").Select(tax => tax.TaxonId).FirstOrDefault();
                            }
                            _infContext.Taxon.Add(newTaxon);
                            SaveToContext();
                            _uniqueOrdersDict.Add(newTaxon.TaxonName, newTaxon.TaxonId);
                            Logger.Debug("--- Added Taxon " + newTaxon.TaxonName);
                            IncrementCounter();
                        }
                        else
                        {
                            var alreadyTax = _infContext.Taxon.Where(tax => tax.TaxonName == newTaxon.TaxonName).FirstOrDefault();
                            if (!alreadyTax.Group.Contains(newTaxon.Group))
                            {
                                alreadyTax.Group += "," + newTaxon.Group;
                                _infContext.Update(alreadyTax);
                            }
                        }
                    }
                    catch (Exception e)
                    {
                        Logger.Error(e, "Error adding new Order Taxon " + worksheetGeneral.Cells[i, orderCol].Value.ToString().Trim() + " to Context");
                    }
                }
            }
            Logger.Info("-- Finished Import of Orders");
        }

        public void ImportFamilies()
        {
            IDictionary<String, String> familyTaxonsExcel = new Dictionary<String, String>();

            //GET ALL FROM EXCEL
            for (int i = 2; i <= worksheetGeneral.Dimension.Rows; i++)
            {
                if (worksheetGeneral.Cells[i, familyCol].Value != null)
                {
                    try
                    {
                        var orderId = _getOrderIdFromExcel(i);
                        var classId = _getClassIdFromExcel(i);

                        Taxon newTaxon = new Taxon
                        {
                            TaxonName = worksheetGeneral.Cells[i, familyCol].Value.ToString().Trim(),
                            TaxonomyStateId = Constants.stateId_family,
                            KingdomId = _kingdomId,
                            PhylumId = _getPhylumIdFromExcel(i),
                            SubphylumId = _getSubPhylumIdFromExcel(i),
                            ClassId = _getClassIdFromExcel(i),
                            SubclassId = _getSubClassIdFromExcel(i),
                            OrderId = _getOrderIdFromExcel(i),
                            SuborderId = _getSubOrderIdFromExcel(i),
                            Group = _getGroupInfo(i),
                        };
                        if (!_checkTaxonExists(newTaxon.TaxonName) && newTaxon.PhylumId != null && newTaxon.ClassId != null && newTaxon.OrderId != null)
                        {
                            _infContext.Taxon.Add(newTaxon);
                            SaveToContext();
                            _uniqueFamiliesDict.Add(newTaxon.TaxonName, newTaxon.TaxonId);
                            Logger.Debug("--- Added Taxon " + newTaxon.TaxonName);
                            IncrementCounter();
                        }
                        else
                        {
                            var alreadyTax = _infContext.Taxon.Where(tax => tax.TaxonName == newTaxon.TaxonName).FirstOrDefault();
                            if (!alreadyTax.Group.Contains(newTaxon.Group))
                            {
                                alreadyTax.Group += "," + newTaxon.Group;
                                _infContext.Update(alreadyTax);
                            }
                        }
                    }
                    catch (Exception e)
                    {
                        Logger.Error(e, "Error adding new Family Taxon " + worksheetGeneral.Cells[i, 2].Value + " to Context");

                    }
                }
            }

            Logger.Info("-- Finished Import of Families");
        }

        public void ImportGenerea()
        {
            IDictionary<String, String> genusTaxonsExcel = new Dictionary<String, String>();

            //GET ALL FROM EXCEL
            for (int i = 2; i <= worksheetGeneral.Dimension.Rows; i++)
            {
                if (worksheetGeneral.Cells[i, taxonCol].Value != null)
                {
                    string genusName = worksheetGeneral.Cells[i, taxonCol].Value.ToString().Trim().Split(" ")[0];
                    string speciesName = worksheetGeneral.Cells[i, taxonCol].Value.ToString().Trim().Split(" ")[1];
                    try
                    {
                        Taxon newTaxon = new Taxon
                        {
                            TaxonName = genusName,
                            TaxonomyStateId = Constants.stateId_genus,
                            KingdomId = _kingdomId,
                            PhylumId = _getPhylumIdFromExcel(i),
                            SubphylumId = _getSubPhylumIdFromExcel(i),
                            ClassId = _getClassIdFromExcel(i),
                            SubclassId = _getSubClassIdFromExcel(i),
                            OrderId = _getOrderIdFromExcel(i),
                            SuborderId = _getSubOrderIdFromExcel(i),
                            FamilyId = _getFamilyIdFromExcel(i),
                            SubfamilyId = _getSubFamilyIdFromExcel(i),
                            Group = _getGroupInfo(i),
                        };
                        if (!_checkTaxonExists(newTaxon.TaxonName) && newTaxon.PhylumId != null && newTaxon.ClassId != null && newTaxon.OrderId != null && newTaxon.FamilyId != null)
                        {
                            _infContext.Taxon.Add(newTaxon);
                            SaveToContext();
                            Logger.Debug("--- Added Taxon " + newTaxon.TaxonName);
                            _uniqueGenereaDict.Add(newTaxon.TaxonName, newTaxon.TaxonId);
                            IncrementCounter();
                        }
                        else
                        {
                            var alreadyTax = _infContext.Taxon.Where(tax => tax.TaxonName == newTaxon.TaxonName).FirstOrDefault();
                            if (!alreadyTax.Group.Contains(newTaxon.Group))
                            {
                                alreadyTax.Group += "," + newTaxon.Group;
                                _infContext.Update(alreadyTax);
                            }
                        }
                    }
                    catch (Exception e)
                    {
                        Logger.Error(e, "Error adding new Genus Taxon " + worksheetGeneral.Cells[i, 2].Value + " to Context");
                    }
                }
            }

            Logger.Info("-- Finished Import of Generea");
        }

        private int? _extractYear(string genusString)
        {
            if (genusString.Length > 0)
            {
                var match = Regex.Match(genusString, @"s?\d{3,4}\)?$");

                if (Int32.TryParse(match.Value.TrimEnd(')').TrimStart(' '), out int yearInt))
                {
                    return yearInt;
                }
            }
            return null;
        }

        public void ImportSpecies()
        {
            IDictionary<String, String> speciesTaxonsExcel = new Dictionary<String, String>();
            IDictionary<String, String> speciesDescriptionExcel = new Dictionary<String, String>();

            //GET ALL INFO FROM EXCEL
            for (int i = 2; i <= worksheetGeneral.Dimension.Rows; i++)
            {
                if (worksheetGeneral.Cells[i, 1].Value != null)
                {
                    string taxonFullName = worksheetGeneral.Cells[i, 1].Value?.ToString().Trim();
                    //name is of type: "Metatrichoniscoides leydigii (Weber, 1880)"
                    if (Regex.IsMatch(taxonFullName, _taxonRegex))
                    {
                        _parseTaxonTypeName(taxonFullName, i);
                    }
                    else
                    {

                    }
                }
            }
            Logger.Info("-- Finished Import of Species");

        }

        private int? _getEdaphobaseIdFromExcel (int i)
        {
            if(edaphoCol > 0)
            {
                int edaphoId = Int32.Parse(worksheetGeneral.Cells[i, edaphoCol].Value?.ToString().Trim());
                return edaphoId;
            }
            return null;
        }

        private void _parseTaxonTypeName(string taxonFullName, int i)
        {
            var taxArray = taxonFullName.Split(' ');
            string genusName = taxArray[0];
            int genusId = GetGenusIdByName(genusName);
            string speciesName = taxArray[0] + " " + taxArray[1];
            //int? descriptionYear = null;
            int? descriptionYear = _extractYear(taxonFullName);
            string? descriptionBy = _extractDescriptionBy(taxonFullName);
            //regular name, eg: "Metatrichoniscoides leydigii (Weber, 1880)"
            //DateTime descriptionYear = DateTime.Parse(descriptionYearStr.Trim() + "-01-01");
            string taxonDistribution = worksheetGeneral.Cells[i, distributionCol].Value?.ToString().Trim();
            string taxonDescription = string.Join(" ", taxArray.Skip(2).ToList());
            string taxonBiotope = worksheetGeneral.Cells[i, biotopeCol].Value?.ToString().Trim();
            string taxonDiagnosis = diagnosisCol>0?worksheetGeneral.Cells[i, diagnosisCol].Value?.ToString().Trim():null;
            string taxonRedListType = redListTypeCol > 0? worksheetGeneral.Cells[i, redListTypeCol].Value?.ToString().Trim(): null;
            string taxonRedListSource = redListSourceCol>0? worksheetGeneral.Cells[i, redListSourceCol].Value?.ToString().Trim() : null;
            string taxonDistEurope = distEuropeCol>0?worksheetGeneral.Cells[i, distEuropeCol].Value?.ToString().Trim() : null;
            string taxonAddInfo = addInfoCol>0?worksheetGeneral.Cells[i, addInfoCol].Value?.ToString().Trim() : null;
            string taxonLitSource = litSourceCol>0?worksheetGeneral.Cells[i, litSourceCol].Value?.ToString().Trim():null;
            string taxonDispLength = dispLengthCol>0 ? worksheetGeneral.Cells[i, dispLengthCol].Value?.ToString().Trim() : null;
            string taxonSliderImages = sliderImageCol >0 ? worksheetGeneral.Cells[i, sliderImageCol].Value?.ToString().Trim() : null;
            bool hasBracketDesc = Regex.IsMatch(taxonFullName, _taxonBracketRegex);
            List<string> i18nNames = new List<string>();
            Guid? taxonGuid = null;
            if (i18nNamesCol > 0)
            {
                i18nNames = worksheetGeneral.Cells[i, i18nNamesCol].Value?.ToString().Trim().Split(',').ToList();
            }
            if (guidCol > 0)
            {
                if (worksheetGeneral.Cells[i, guidCol].Value != null) {
                    taxonGuid = Guid.Parse(worksheetGeneral.Cells[i, guidCol].Value.ToString());
                }
            }

            if (!SpeciesExistsInContext(speciesName, genusId))
            {
                try
                {
                    SaveSpeciesTaxonToContext(speciesName, genusId, taxonDescription, taxonDistribution, taxonBiotope, descriptionYear, descriptionBy, hasBracketDesc, taxonRedListSource, taxonRedListType, taxonAddInfo, taxonLitSource, taxonDistEurope, taxonDiagnosis, taxonDispLength, taxonSliderImages, i, i18nNames, taxonGuid);
                }
                //save species with info in excel row - TODO: rewrite below
                catch (Exception e)
                {
                    Logger.Error(e, "Error adding new Species Taxon " + speciesName + " to Context");
                }

            }
        }

        private string _extractDescriptionBy(string taxonFullName)
        {
            if (taxonFullName.Length > 0)
            {
                var match = Regex.Match(taxonFullName, _taxonNameExtractRegex);
                return match.Groups[1]?.ToString();
            }
            return null;
        }

        private void SaveSpeciesTaxonToContext(string speciesTaxonName, int genusId, string taxonDescription, string taxonBiotope, string taxonDistribution, int? descriptionYear, string descriptionBy, bool hasBracketDesc, string taxonRedListSource, string taxonRedListType, string taxonAddInfo, string taxonLitSource, string taxonDistEurope, string taxonDiagnosis, string taxonDispLength, string taxonSliderImages, int i, List<string> i18nNames, Guid? taxonGuid)
        {
            //var genusId = _uniqueGenereaDict.ContainsKey(speciesItemExcel.Value) ? _uniqueGenereaDict[speciesItemExcel.Value] : (int?)null;
            var familyId = _infContext.Taxon.Where(t => t.TaxonId == genusId).FirstOrDefault().FamilyId;
            var orderId = _infContext.Taxon.Where(t => t.TaxonId == familyId).FirstOrDefault().OrderId;
            var classId = _infContext.Taxon.Where(t => t.TaxonId == orderId).FirstOrDefault().ClassId;

            Taxon newTaxon = new Taxon
            {
                TaxonName = speciesTaxonName,
                TaxonomyStateId = Constants.stateId_species,
                TaxonDescription = taxonDescription,
                KingdomId = _kingdomId,
                PhylumId = _getPhylumIdFromExcel(i),
                SubphylumId = _getSubPhylumIdFromExcel(i),
                ClassId = _getClassIdFromExcel(i),
                SubclassId = _getSubClassIdFromExcel(i),
                OrderId = _getOrderIdFromExcel(i),
                SuborderId = _getSubOrderIdFromExcel(i),
                FamilyId = _getFamilyIdFromExcel(i),
                SubfamilyId = _getSubFamilyIdFromExcel(i),
                GenusId = genusId,
                DescriptionYear = descriptionYear,
                DescriptionBy = descriptionBy,
                TaxonBiotopeAndLifestyle = taxonBiotope,
                TaxonDistribution = taxonDistribution,
                HasBracketDescription = hasBracketDesc,
                RedListTypeId = _mapRedListType(taxonRedListType),
                RedListSource = taxonRedListSource,
                Diagnosis = taxonDiagnosis,
                LiteratureSource = taxonLitSource,
                DistributionEurope = taxonDistEurope,
                AdditionalInfo = taxonAddInfo,
                Group = _getGroupInfo(i),
                EdaphobaseId = _getEdaphobaseIdFromExcel(i),
                DisplayLength = taxonDispLength,
                SliderImages = JsonConvert.SerializeObject(taxonSliderImages?.Split(';').ToList()),
                I18nNames = JsonConvert.SerializeObject(i18nNames),
                Identifier = taxonGuid
            };
            if (!_checkTaxonExists(newTaxon.TaxonName) && newTaxon.PhylumId != null && newTaxon.ClassId != null && newTaxon.OrderId != null && newTaxon.FamilyId != null && newTaxon.GenusId != null)
            {
                _infContext.Taxon.Add(newTaxon);
                SaveToContext();
                Logger.Debug("--- Added Taxon " + newTaxon.TaxonName);
                _uniqueSpeciesDict.Add(newTaxon.TaxonName, newTaxon.TaxonId);
                IncrementCounter();
            }
            else
            {
                var alreadyTax = _infContext.Taxon.Where(tax => tax.TaxonName == newTaxon.TaxonName).FirstOrDefault();
                if (!alreadyTax.Group.Contains(newTaxon.Group))
                {
                    alreadyTax.Group += "," + newTaxon.Group;
                    _infContext.Update(alreadyTax);
                }
            }
        }

        private int? _mapRedListType (string redListTypeStr)
        {
            //string longOut = shortIn;
            if(redListTypeStr != null)
            {
                if (_redListMap.TryGetValue(redListTypeStr, out int redListTypId))
                {
                    return redListTypId;
                }
            }
            return null;
        }

        private int GetGenusIdByName(string genusName)
        {
            var genusTaxon = _infContext.Taxon.Where(tax => tax.TaxonName == genusName).FirstOrDefault();
            if (genusTaxon != null)
            {
                return genusTaxon.TaxonId;
            }
            return -1;
        }
        private bool SpeciesExistsInContext(string speciesName, int? genusId)
        {
            var speciesTaxon = _infContext.Taxon.Where(tax => tax.GenusId == genusId && tax.TaxonName == speciesName).FirstOrDefault();
            if (speciesTaxon != null)
            {
                return true;
            }
            return false;
        }

    }
}