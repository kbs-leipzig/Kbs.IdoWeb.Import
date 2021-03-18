using Kbs.IdoWeb.Data.Determination;
using Kbs.IdoWeb.Data.Information;
using Microsoft.EntityFrameworkCore;
using Microsoft.Extensions.Configuration;
using NLog;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text.RegularExpressions;

namespace Kbs.IdoWeb.Import.Models
{
    class TaxonDescriptionImportModel
    {
        //TODO: importmodel superclass?
        private static readonly NLog.Logger Logger = NLog.LogManager.GetCurrentClassLogger();
        public int taxDescKeyCounter;
        public int descKeyCounter;
        internal ExcelWorksheet worksheetMatrix;
        private DeterminationContext _contextDet;
        private InformationContext _contextInf;
        //Get dynamically?
        private int _descKeyGroupRowIndex = 4;
        private int? _mwKeyGroupRowIndex = null;
        private Dictionary<string, int> _keyTypeNameToIdDict = new Dictionary<string, int>();
        private const string _taxonRegex = @"^(?:\S[^\(\)]*\s\S[^\(\)]*\s)((?:\(?)(\S*|\S*\s\S*|\S*|\S{2}\s\S{2}\s\S*|\S{2,4}\s\S|\S\s\S\s\S)(?:\,\s?\d{4}\)?))$";
        private Dictionary<string, string> bundeslandLongNameMap = new Dictionary<string, string>()
        {
            {"BW", "Baden-Württemberg"},
            {"BY", "Bayern"},
            {"BE", "Berlin"},
            {"BB", "Brandenburg"},
            {"HB", "Bremen"},
            {"HH", "Hamburg"},
            {"HE", "Hessen"},
            {"NI", "Niedersachsen"},
            {"MV", "Mecklenburg-Vorpommern"},
            {"NW", "Nordrhein-Westfalen"},
            {"RP", "Rheinland-Pfalz"},
            {"SL", "Saarland"},
            {"SN", "Sachsen"},
            {"ST", "Sachsen-Anhalt"},
            {"SH", "Schleswig-Holstein"},
            {"TH", "Thüringen"},
            {"gesamt", "gesamt"}
        };

        public TaxonDescriptionImportModel()
        {
            InitContexts();

        }

        private void InitContexts()
        {
            var optionsBuilderDet = new DbContextOptionsBuilder<DeterminationContext>();
            optionsBuilderDet.UseNpgsql(Program.Configuration.GetConnectionString("DatabaseConnection"));
            _contextDet = new DeterminationContext(optionsBuilderDet.Options);

            var optionsBuilderInf = new DbContextOptionsBuilder<InformationContext>();
            optionsBuilderInf.UseNpgsql(Program.Configuration.GetConnectionString("DatabaseConnection"));
            _contextInf = new InformationContext(optionsBuilderInf.Options);
        }

        private void InitProperties()
        {
            for (int i = 1; i <= worksheetMatrix.Dimension.Rows; i++)
            {
                if (worksheetMatrix.Cells[i, 1].Value != null)
                {
                    if (worksheetMatrix.Cells[i, 1].Value.ToString().Trim() == "Taxon")
                    {
                        _descKeyGroupRowIndex = i;
                    }
                    if (worksheetMatrix.Cells[i, 1].Value.ToString().Trim() == "Geschlecht")
                    {
                        _mwKeyGroupRowIndex = i;
                    }
                }
            }
            _keyTypeNameToIdDict = _contextDet.DescriptionKeyType.ToDictionary(x => x.DescriptionKeyTypeName.ToString().Trim().ToLower().Replace(" ", ""), x => x.DescriptionKeyTypeId);
        }

        private int? MapDescKeyMWType(string mwTypeName)
        {
            if (mwTypeName != null)
            {
                mwTypeName = mwTypeName.Trim().Replace(" ", "").Replace("&", ";").ToLower();
                int mwTypeId = 0;
                if (_keyTypeNameToIdDict.TryGetValue(mwTypeName, out mwTypeId))
                {
                    return mwTypeId;
                }
            }
            return null;
        }

        //TODO: catch errors in elses
        internal void StartTaxDescriptionImport()
        {
            InitProperties();

            Logger.Info("-- Started TaxonDescriptionImport");
            for (int i = _descKeyGroupRowIndex + 1; i <= worksheetMatrix.Dimension.Rows; i++)
            {
                //Parse Region Col
                var cellVal_taxon = worksheetMatrix.Cells[i, 1].Value;
                if (cellVal_taxon != null)
                {
                    //GetTaxonId
                    var taxonId = GetTaxonIdFromExcelVal(cellVal_taxon);
                    if (taxonId != null)
                    {
                        //Parse rest of Row
                        for (int j = 2; j <= worksheetMatrix.Dimension.Columns; j++)
                        {
                            var cellVal_descKey = worksheetMatrix.Cells[i, j].Value;
                            if (cellVal_descKey != null)
                            {
                                int? mwTypeId = null;
                                if (_mwKeyGroupRowIndex != null)
                                {
                                    string mwType = worksheetMatrix.Cells[_mwKeyGroupRowIndex.GetValueOrDefault(), j].Value?.ToString().Trim();
                                    mwTypeId = MapDescKeyMWType(mwType);
                                }

                                var descKeyName = cellVal_descKey.ToString().Trim();
                                List<string> multipleDescKeyNames = null;

                                //check if ; present --> if so split into distinct desckeys
                                if (descKeyName.Contains(';'))
                                {
                                    multipleDescKeyNames = descKeyName.Split(';').Select(p => p.Trim()).ToList();
                                    multipleDescKeyNames.RemoveAll(s => String.IsNullOrEmpty(s.Trim()));
                                }

                                //Represents top 3 rows in Excel List
                                List<string> descKeyGroupList = GetDescKeyGroupFromExcel(j, mwTypeId);
                                var descKeyGroupId = GetDescKeyGroupIdFromKeyGroupList(descKeyGroupList);
                                if(descKeyGroupList.Count() > 1)
                                {
                                    doSomethingWithDK(descKeyGroupId, descKeyName, taxonId, mwTypeId, descKeyGroupList, multipleDescKeyNames, cellVal_descKey, j);
                                }
                            }
                        }
                    }
                }
                else
                {
                    Logger.Debug($"--- Did not parse Row {i} due to missing Taxon Value");
                }
                Logger.Info("-- Parsed Row: " + i);
            }
            Logger.Info("-- Finished Parsing TaxonDescription Excel Sheet");
        }

        private void doSomethingWithDK(int? descKeyGroupId, string descKeyName, int? taxonId, int? mwTypeId, List<string> descKeyGroupList, List<string> multipleDescKeyNames, object cellVal_descKey, int j)
        {
            string descKeyGroupDataType = GetDescKeyGroupDataType(descKeyGroupId);
            //GET GENDER/DescriptionKeyTypeId IF ROW IS SET IN EXCEL

            //Got descriptionKeyGroup from excel
            if (descKeyGroupId != null)
            {
                int? descKeyId = null;
                //KEYGROUP == VALUELIST || YESNO?
                //FOR VALUELIST AND YESNO: only add taxdesc IF descKey matches existing key
                if (descKeyGroupDataType == "VALUE")
                {
                    decimal? minVal = null;
                    decimal? maxVal = null;
                    var style = NumberStyles.AllowDecimalPoint | NumberStyles.AllowThousands;
                    var provider = new CultureInfo("de-DE");
                    //RANGE, e.g. 25-50 -> mapped to minvalue and maxvalue, dont create desckey (e.g. Länge bis(mm), Antenne: Anzahl Antennen)
                    if (descKeyName.Contains('-'))
                    {
                        //TODO: Länge vs Länge bis -> clarify: replace one?
                        var rangeArray = descKeyName.Split('-');
                        //FLOAT IT
                        minVal = decimal.Parse(Regex.Replace(rangeArray[0], "[^0-9.,]", ""), style, provider);
                        maxVal = decimal.Parse(Regex.Replace(rangeArray[1], "[^0-9.,]", ""), style, provider);
                    }
                    //EXACT VALUE -> set minval equal to maxval for easier search -- TODO: clarify
                    else
                    {
                        minVal = decimal.Parse(Regex.Replace(descKeyName, "[^0-9.,]", ""), style, provider);
                        maxVal = minVal;
                        //descKeyId = SaveDescriptionKeyToContext(descKeyName.Trim(), descKeyGroupId);
                    }

                    descKeyId = GetDescKeyIdByNameAndDescKeyGroupId(descKeyGroupList[0], descKeyGroupId.GetValueOrDefault());

                    if (descKeyId != null)
                    {
                        SaveTaxonDescriptionToContext(taxonId, descKeyId, mwTypeId, minVal, maxVal, descKeyName);
                    }
                    else
                    {
                        Logger.Debug($"--- Did not add TaxonDescr \'{cellVal_descKey}\' with TaxonId {taxonId} to Context, could not find matching DescKey");
                    }
                }
                else if (descKeyGroupDataType == "UNKNOWN" || descKeyGroupDataType == "VALUELIST" || descKeyGroupDataType == null || descKeyGroupDataType == "OPEN VAL" || descKeyGroupDataType == "PIC")
                {
                    if (multipleDescKeyNames != null)
                    {
                        foreach (string descKeyNameItem in multipleDescKeyNames)
                        {
                            //_mapBundeslandRealName(descKeyNameItem);
                            descKeyId = SaveDescriptionKeyToContext(_isBundeslandCol(j) ? _mapBundeslandRealName(descKeyNameItem.Trim()) : descKeyNameItem.Trim(), descKeyGroupId);

                            if (descKeyId != null)
                            {
                                SaveTaxonDescriptionToContext(taxonId, descKeyId, mwTypeId, null, null, _isBundeslandCol(j) ? _mapBundeslandRealName(descKeyNameItem.Trim()) : descKeyNameItem);
                            }
                            else
                            {
                                Logger.Debug($"--- Did not add TaxonDescr \'{cellVal_descKey}\' with TaxonId {taxonId} to Context, could not find matching DescKey");
                            }
                        }
                    }
                    else
                    {
                        //_mapBundeslandRealName(descKeyNameItem);
                        descKeyId = SaveDescriptionKeyToContext(_isBundeslandCol(j) ? _mapBundeslandRealName(descKeyName.Trim()) : descKeyName, descKeyGroupId);

                        if (descKeyId != null)
                        {
                            SaveTaxonDescriptionToContext(taxonId, descKeyId, mwTypeId, null, null, _isBundeslandCol(j) ? _mapBundeslandRealName(descKeyName.Trim()) : descKeyName);
                        }
                        else
                        {
                            Logger.Debug($"--- Did not add TaxonDescr \'{cellVal_descKey}\' with TaxonId {taxonId} to Context, could not find matching DescKey");
                        }
                    }


                }
                //VALUELIST || YESNO --> *dont* create new desckeys!
                else
                {
                    if (multipleDescKeyNames != null)
                    {
                        foreach (string descKeyNameItem in multipleDescKeyNames)
                        {
                            //descKeyId = SaveDescriptionKeyToContext(descKeyNameItem.Trim(), descKeyGroupId);
                            descKeyId = GetDescKeyIdByNameAndDescKeyGroupId(_isBundeslandCol(j) ? _mapBundeslandRealName(descKeyNameItem.Trim()) : descKeyNameItem.Trim(), descKeyGroupId.GetValueOrDefault());

                            if (descKeyId != null)
                            {
                                SaveTaxonDescriptionToContext(taxonId, descKeyId, mwTypeId, null, null, descKeyNameItem);
                            }
                            else
                            {
                                Logger.Debug($"--- Did not add TaxonDescr {cellVal_descKey} with ID {taxonId} to Context, could not find matching DescKey");
                            }
                        }
                    }
                    else
                    {
                        //descKeyId = SaveDescriptionKeyToContext(descKeyName.Trim(), descKeyGroupId);
                        descKeyId = GetDescKeyIdByNameAndDescKeyGroupId(descKeyName, descKeyGroupId.GetValueOrDefault());
                        if (descKeyId != null)
                        {
                            SaveTaxonDescriptionToContext(taxonId, descKeyId, mwTypeId, null, null, descKeyName);
                        }
                        else
                        {
                            Logger.Debug($"--- Did not add TaxonDescr {cellVal_descKey} with ID {taxonId} to Context, could not find matching DescKey");
                        }
                    }
                }
            }
            else
            {
                Logger.Debug("--- Did not add TaxonDescr: \'" + cellVal_descKey + "\' ({0}/{1}/{2}) to Context (TaxId/DescName/DescKeyType)", taxonId, descKeyName, mwTypeId);
                Logger.Debug("--- Did not find related DescKeyId");
            }
        }

        private bool _isBundeslandCol(int j)
        {
            try
            {
                var firstRowValue = worksheetMatrix.Cells[1, j].Value?.ToString().Trim();
                if (firstRowValue == "Bundesland")
                {
                    return true;
                }
                else if (firstRowValue == null)
                {
                    firstRowValue = worksheetMatrix.Cells[3, j].Value?.ToString().Trim();
                }

                if (firstRowValue != null)
                {
                    if (firstRowValue == "Bundesland" || firstRowValue == "Fundort")
                        return true;
                }

            }
            catch (Exception e)
            {
                Logger.Warn($"-- Could not determine Bundesland Column {e}");
            }
            return false;
        }

        private string _mapBundeslandRealName(string shortIn)
        {
            //string longOut = shortIn;
            if (bundeslandLongNameMap.TryGetValue(shortIn, out string longOut))
            {
                return longOut;
            }
            return shortIn;
        }

        private void SaveTaxonDescriptionToContext(int? taxonId, int? descKeyId, int? mwTypeId, decimal? minVal = null, decimal? maxVal = null, string descKeyName = null)
        {
            if (!TaxonDescriptionExists(taxonId, descKeyId, mwTypeId) || (minVal != null && maxVal == null))
            {
                TaxonDescription taxDescInstance = new TaxonDescription { TaxonId = taxonId.GetValueOrDefault(), DescriptionKeyId = descKeyId.GetValueOrDefault(), DescriptionKeyTypeId = mwTypeId, MinValue = minVal, MaxValue = maxVal };
                _contextDet.Add(taxDescInstance);
                Logger.Debug("--- Saving TaxonDescr {0} ({1}/{2}/{3}) to Context (TaxId/DescId/DescKeyType) ...", descKeyName, taxonId, descKeyId, mwTypeId);
                SaveToContext();
                taxDescKeyCounter++;
            }
            else
            {
                Logger.Debug("--- Did not add TaxonDescr {0} ({1}/{2}/{3}) to Context (TaxId/DescId/DescKeyType)", descKeyName, taxonId, descKeyId, mwTypeId);
                Logger.Debug("--- Entity already exists");
            }
        }

        private List<int> GetMultipleDescKeyGroupIdFromKeyGroupList(List<string> descKeyGroupList)
        {
            List<DescriptionKeyGroup> descKeyGroup = null;

            if (descKeyGroupList != null)
            {
                var descKeyArr = descKeyGroupList[1].Split(";").Select(p => p.Trim()).ToList();

                /**
                *Körper allgemein
                *Oberseite; Unterseite
                *Anzahl Körperringe
                **/
                if (descKeyGroupList.Count == 3)
                {
                    descKeyGroup = _contextDet.DescriptionKeyGroup
                        .Where(dkg => dkg.KeyGroupName == descKeyGroupList[0])
                        .Include(dkg => dkg.ParentDescriptionKeyGroup)
                            .ThenInclude(dkgp => dkgp.ParentDescriptionKeyGroup)
                        .ToList();

                    descKeyGroup.RemoveAll(dkg => !descKeyArr.Contains(dkg.ParentDescriptionKeyGroup.KeyGroupName));

                    if (descKeyGroup != null)
                    {
                        
                        if (descKeyGroup.Count == descKeyArr.Count)
                        {
                            return descKeyGroup.Select(dkg => dkg.DescriptionKeyGroupId).ToList();
                        }
                        else if (descKeyGroup.Count < descKeyArr.Count)
                        {
                            Logger.Error("--- Could not find Multiple DescKeyGroupId [{0}]", string.Join(", ", descKeyGroupList));
                            return null;
                        }
                        else
                        {
                            try
                            {
                                //descKeyGroup.RemoveAll(dkg => dkg.ParentDescriptionKeyGroup == null);
                                List<DescriptionKeyGroup> descKeyGroupParentFilterList = descKeyGroup
                                    .Where(dkg => dkg.ParentDescriptionKeyGroup?.KeyGroupName == descKeyGroupList[1])
                                    .ToList();

                                if (descKeyGroupParentFilterList != null)
                                {
                                    if (descKeyGroupParentFilterList.Count < descKeyArr.Count)
                                    {
                                        Logger.Debug("--- Could not find Multiple DescKeyGroupId [{0}]", string.Join(", ", descKeyGroupList));
                                        return null;

                                    }
                                    else if (descKeyGroupParentFilterList.Count == descKeyArr.Count)
                                    {
                                        return descKeyGroupParentFilterList.Select(dkg => dkg.DescriptionKeyGroupId).ToList();
                                    }
                                    else
                                    {
                                        var descKeyGroupGrandParentFilterList = descKeyGroupParentFilterList
                                            .Where(dkgp => dkgp.ParentDescriptionKeyGroup.KeyGroupName == descKeyGroupList[2])
                                            .ToList();

                                        if(descKeyGroupGrandParentFilterList.Count == descKeyArr.Count)
                                        {
                                            return descKeyGroupGrandParentFilterList.Select(dkg => dkg.DescriptionKeyGroupId).ToList();

                                        } 
                                        else
                                        {
                                            Logger.Debug("--- Could not find Multiple DescKeyGroupId [{0}]", string.Join(", ", descKeyGroupList));
                                            return null;
                                        }
                                    }
                                }
                            }
                            catch (Exception e)
                            {
                                Logger.Error(e, $"Error Getting Parent DKG Info for DKG {descKeyGroup.FirstOrDefault().KeyGroupName}");
                            }
                        }
                    }
                }
            }
            return null;
        }

        private int? GetDescKeyGroupIdFromKeyGroupList(List<string> descKeyGroupList)
        {
            List<DescriptionKeyGroup> descKeyGroup = null;

            if (descKeyGroupList != null)
            {
                if (descKeyGroupList.Count == 1)
                {
                    descKeyGroup = _contextDet.DescriptionKeyGroup
                        .Where(dkg => dkg.KeyGroupName == descKeyGroupList[0])
                        .ToList();
                    //TODO: clarify what to do when more than 1 result but excel only 1 desckeygroup element?
                    if (descKeyGroup != null)
                    {
                        if (descKeyGroup.Count == 0)
                        {
                            Logger.Debug("--- Could not find DescKeyGroupId [{0}]", string.Join(", ", descKeyGroupList));
                        }
                        else if (descKeyGroup.Count == 1)
                        {
                            return descKeyGroup.FirstOrDefault().DescriptionKeyGroupId;
                        }
                        else
                        {
                            Logger.Debug("--- Could not determine DescKeyGroupId for combination: [{0}]", string.Join(", ", descKeyGroupList));
                        }
                    }
                }

                else if (descKeyGroupList.Count == 2)
                {
                    descKeyGroup = _contextDet.DescriptionKeyGroup
                            .Where(dkg => dkg.KeyGroupName == descKeyGroupList[0])
                            .Include(dkg => dkg.ParentDescriptionKeyGroup).ToList();

                    if (descKeyGroup != null)
                    {

                        if (descKeyGroup.Count == 0)
                        {
                            Logger.Debug("--- Could not find DescKeyGroupId [{0}]", string.Join(", ", descKeyGroupList));
                        }
                        else if (descKeyGroup.Count == 1)
                        {
                            return descKeyGroup.First().DescriptionKeyGroupId;
                        }
                        else
                        {
                            DescriptionKeyGroup descKeyGroupParentFilter = descKeyGroup.Where(dkg => dkg.ParentDescriptionKeyGroup.KeyGroupName == descKeyGroupList[1]).FirstOrDefault();
                            if (descKeyGroupParentFilter != null)
                            {
                                return descKeyGroupParentFilter.DescriptionKeyGroupId;
                            }
                        }
                    }

                }

                else if (descKeyGroupList.Count == 3)
                {
                    if (descKeyGroupList[1].Contains(";"))
                    {
                        var descKeyArr = descKeyGroupList[1].Split(";");
                        foreach (string dkName in descKeyArr)
                        {

                        }
                    }

                    descKeyGroup = _contextDet.DescriptionKeyGroup
                        .Where(dkg => dkg.KeyGroupName == descKeyGroupList[0])
                        .Include(dkg => dkg.ParentDescriptionKeyGroup)
                            .ThenInclude(dkgp => dkgp.ParentDescriptionKeyGroup)
                        .ToList();

                    if (descKeyGroup != null)
                    {

                        if (descKeyGroup.Count == 0)
                        {
                            Logger.Debug("--- Could not find DescKeyGroupId [{0}]", string.Join(", ", descKeyGroupList));
                            return null;
                        }
                        else if (descKeyGroup.Count == 1)
                        {
                            return descKeyGroup.FirstOrDefault().DescriptionKeyGroupId;
                        }
                        else
                        {
                            try
                            {

                                //descKeyGroup.RemoveAll(dkg => dkg.ParentDescriptionKeyGroup == null);

                                List<DescriptionKeyGroup> descKeyGroupParentFilterList = descKeyGroup
                                    .Where(dkg => dkg.ParentDescriptionKeyGroup?.KeyGroupName == descKeyGroupList[1])
                                    .ToList();
                                if (descKeyGroupParentFilterList != null)
                                {
                                    if (descKeyGroupParentFilterList.Count == 0)
                                    {
                                        Logger.Debug("--- Could not find DescKeyGroupId [{0}]", string.Join(", ", descKeyGroupList));
                                        return null;

                                    }
                                    else if (descKeyGroupParentFilterList.Count == 1)
                                    {
                                        return descKeyGroupParentFilterList.FirstOrDefault().DescriptionKeyGroupId;
                                    }
                                    else
                                    {
                                        var descKeyGroupGrandParentFilterList = descKeyGroupParentFilterList.Where(dkgp => dkgp.ParentDescriptionKeyGroup.KeyGroupName == descKeyGroupList[2]).FirstOrDefault();
                                        Logger.Debug("--- Could not find DescKeyGroupId [{0}]", string.Join(", ", descKeyGroupList));
                                        return null;
                                    }
                                }
                            }
                            catch (Exception e)
                            {
                                Logger.Error(e, $"Error Getting Parent DKG Info for DKG {descKeyGroup.FirstOrDefault().KeyGroupName}");
                            }
                        }
                    }
                }
            }
            return null;
        }

        private int? SaveDescriptionKeyToContext(string descKeyName, int? descKeyGroupId)
        {
            int? descKeyId = null;
            //FOR VALUES and UNKNOWN: currently all values are added: TODO: how to validate?
            //NEW DESCRIPTIONKEY NEEDS TO BE CREATED
            if (!DescriptionKeyExists(descKeyName, descKeyGroupId))
            {
                DescriptionKey newDescKey = new DescriptionKey { KeyName = descKeyName, DescriptionKeyGroupId = descKeyGroupId.GetValueOrDefault() };
                _contextDet.Add(newDescKey);
                SaveToContext();
                Logger.Debug("--- Added new DescKey \'" + descKeyName + "\' to Context ID: " + descKeyGroupId);
                descKeyCounter++;
                descKeyId = newDescKey.DescriptionKeyId;
            }
            else
            {
                //TODO: error handling
                Logger.Debug("--- Did not add new DescKey \'" + descKeyName + "\' to Context for KeyType VALUE");
                Logger.Debug("--- Description Key already Exists");
                descKeyId = GetDescKeyIdByNameAndDescKeyGroupId(descKeyName, descKeyGroupId.GetValueOrDefault());
            }

            return descKeyId;

        }

        private bool DescriptionKeyExists(string descKeyName, int? descKeyGroupId)
        {
            if (descKeyGroupId != null)
            {
                var descKey = _contextDet.DescriptionKey.Where(dk => (dk.KeyName.ToLower() == descKeyName.ToLower()) && (dk.DescriptionKeyGroupId == descKeyGroupId)).FirstOrDefault();
                if (descKey != null)
                {
                    return true;
                }
            }
            return false;
        }

        private string GetDescKeyGroupDataType(int? descKeyGroupId)
        {
            var descKeyGroup = _contextDet.DescriptionKeyGroup.Where(dkg => dkg.DescriptionKeyGroupId == descKeyGroupId).FirstOrDefault();
            if (descKeyGroup != null)
            {
                return descKeyGroup.DescriptionKeyGroupDataType;
            }
            return null;
        }

        private void SaveToContext()
        {
            try
            {
                _contextDet.SaveChanges();
                Logger.Debug(".. Saved to Context");

            }
            catch (Exception e)
            {
                Logger.Error(e.InnerException, "Error Saving Changes to Context");
            }
        }
        private int? GetDescKeyIdByNameAndDescKeyGroupId(string dkName, int dkgId)
        {
            var descKey = _contextDet.DescriptionKey.Where(dk => dk.KeyName.ToLower() == dkName.ToLower() && dk.DescriptionKeyGroupId == dkgId).FirstOrDefault();
            if (descKey != null)
            {
                return descKey.DescriptionKeyId;
            }
            return null;
        }

        /** TODO: refactor if-else **/
        private List<string> GetDescKeyGroupFromExcel(int colIndex, int? mwTypeId = null)
        {
            List<string> descKeyGroups = new List<string>();
            if (_mwKeyGroupRowIndex != null)
            {
                for (int j = 0; j < 4; j++)
                {
                    if (j != 1)
                    {
                        descKeyGroups.Add(worksheetMatrix.Cells[_descKeyGroupRowIndex - j, colIndex].Value?.ToString().Trim());
                    }
                }
            }
            else
            {
                for (int j = 0; j < 3; j++)
                {
                    descKeyGroups.Add(worksheetMatrix.Cells[_descKeyGroupRowIndex - j, colIndex].Value?.ToString().Trim());
                }
            }

            descKeyGroups.RemoveAll(item => item == null);
            return descKeyGroups;
        }

        private int? GetTaxonIdFromExcelVal(object cellVal)
        {
            string taxonFullName = cellVal?.ToString().Trim();
            if (taxonFullName != null)
            {
                //check if sth like "Metatrichoniscoides leydigii (Weber, 1880)"
                Taxon taxonInstance = null;
                if (Regex.IsMatch(taxonFullName, _taxonRegex))
                {
                    var taxArray = taxonFullName.Split(' ');
                    string genusName = taxArray[0];
                    string speciesName = taxArray[0] + " " + taxArray[1];
                    string taxonDescription = string.Join(" ", taxArray.Skip(2));
                    taxonDescription = taxonDescription.ToString().TrimEnd(')').TrimStart('(');
                    //disabled below as of 25.11.19; data in excel file does not always have a descriptionBy and descriptionYear
                    //taxonInstance = _contextInf.Taxon.Where(tax => tax.TaxonName == speciesName && tax.TaxonDescription == taxonDescription).FirstOrDefault();
                    taxonInstance = _contextInf.Taxon.Where(tax => tax.TaxonName == speciesName).FirstOrDefault();
                }
                else
                {
                    taxonInstance = _contextInf.Taxon.Where(tax => tax.TaxonName == taxonFullName).FirstOrDefault();
                }
                if (taxonInstance != null)
                {
                    return taxonInstance.TaxonId;
                }
            }
            return null;
        }

        private bool TaxonDescriptionExists(int? taxonId, int? descriptionId, int? mwTypeId)
        {
            var taxDesc = _contextDet.TaxonDescription.Where(td => td.TaxonId == taxonId && td.DescriptionKeyId == descriptionId && td.DescriptionKeyTypeId == mwTypeId).FirstOrDefault();
            if (taxDesc != null)
            {
                return true;
            }
            return false;
        }

    }

}
