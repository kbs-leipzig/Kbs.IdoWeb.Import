using Kbs.IdoWeb.Data.Determination;
using Microsoft.EntityFrameworkCore;
using Microsoft.Extensions.Configuration;
using Newtonsoft.Json;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

namespace Kbs.IdoWeb.Import.Models
{
    internal class DescriptionImportModel
    {
        private static readonly NLog.Logger Logger = NLog.LogManager.GetCurrentClassLogger();
        public ExcelWorksheet worksheetCharacteristics;
        private DeterminationContext _context = new DeterminationContext();

        //Indices of header row - may vary in worksheet
        private int _regionColIndexExcel;

        private int _lageColIndexExcel;
        private int _merkmaleColIndexExcel;
        private int? _untermerkmaleColIndexExcel = null;
        private int _typColIndexExcel;
        private bool _mwColExists = false;

        //TODO: add heading and get dynamically?
        private int _mwTypeColIndexExcel = 3;

        private int _ampelColIndexExcel = 0;
        private int _orderPrioColIndexExcel = 0;
        private int _keyDesciptionColIndexExcel = 0;
        private int _abbColIndexExcel;
        public int descKeyCounter = 0;
        public int descKeyGroupCounter = 0;
        private Dictionary<string, int> _ampelTypeToIdDict = new Dictionary<string, int>();
        //Array to save ExcelRow - keyGroupId to avoid searching context for parentKeyGroups
        /**TODO: clarify how much items will be imported -> performance issues? **/
        private int[] _keyGroupToExcelRowMapping;
        private int[] _descKeyToExcelRowMapping;
        private string _dkNameRegex = @"^[a-zA-Z0-9äÄöÖüÜß_ .,\)\(\/-:-]*$";
        private Dictionary<string, int> _keyTypeNameToIdDict = new Dictionary<string, int>();

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

        public DescriptionImportModel()
        {
            InitDeterminationContext();
        }

        public void StartDescriptionImport()
        {
            //DEBUGGING ONLY - TBD
            //IterateWS();
            InitProperties();
            FillContext();
            ParseHeaderRow();
            //DIFFERENT APPROACH? - SAVE ROWS SEQUENTIALLY?
            ImportDescKeyGroups();
            ImportDescKeys();
        }

        private void InitProperties()
        {
            _keyGroupToExcelRowMapping = new int[worksheetCharacteristics.Dimension.Rows];
            _descKeyToExcelRowMapping = new int[worksheetCharacteristics.Dimension.Rows];
            //lower and remove whitespaces to avoid typos causing mismatches
            _keyTypeNameToIdDict = _context.DescriptionKeyType.ToDictionary(x => x.DescriptionKeyTypeName.ToString().Trim().ToLower().Replace(" ", ""), x => x.DescriptionKeyTypeId);
            _ampelTypeToIdDict = _context.VisibilityCategory.ToDictionary(x => x.VisibilityCategoryName.ToString().Trim().ToLower().Replace(" ", ""), x => x.VisibilityCategoryId);
        }

        private void InitDeterminationContext()
        {
            var optionsBuilder = new DbContextOptionsBuilder<DeterminationContext>();
            optionsBuilder.UseNpgsql(Program.Configuration.GetConnectionString("DatabaseConnection"));
            _context = new DeterminationContext(optionsBuilder.Options);
        }

        /** TODO: Relation Kopf - Unterseite? **/

        private void ImportDescKeyGroups()
        {
            //Skip Header Row and empty Row
            //First get "Region" KeyGroups
            for (int i = 3; i <= worksheetCharacteristics.Dimension.Rows; i++)
            {
                //Parse Region Col
                var cellVal = worksheetCharacteristics.Cells[i, _regionColIndexExcel].Value;
                if (cellVal != null)
                {
                    //TODO: switch back when clarified how to import desckeygroups
                    var keyGroupName = cellVal.ToString().Trim();

                    //var keyGroupName = ConcatKeyGroupName(i);

                    if (!DescKeyGroupNameExists(keyGroupName))
                    {
                        string keyGroupDataType = null;
                        if (isRegionOnlyKeyGroup(i))
                        {
                            keyGroupDataType = GetDescKeyGroupDataType(i);
                        }

                        DescriptionKeyGroup descriptionKeyGroup = new DescriptionKeyGroup { KeyGroupName = keyGroupName, DescriptionKeyGroupDataType = keyGroupDataType, VisibilityCategoryId = _MapAmpelType(i) };
                        _context.Add(descriptionKeyGroup);
                        SaveToContext();
                        descKeyGroupCounter++;
                    }
                }
            }

            //Lage -> Region
            Logger.Info("-- Started Lage-Region Import");
            ParseAndSaveExcelColumn(_lageColIndexExcel, _regionColIndexExcel, false);
            Logger.Info("-- Finished Lage-Region Import");
            //Merkmale -> Lage
            Logger.Info("-- Started DescKeyGroups Merkmal-Lage Import");
            ParseAndSaveExcelColumn(_merkmaleColIndexExcel, _lageColIndexExcel, true);
            Logger.Info("-- Finished DescKeyGroups: Merkmal-Lage Import");
        }

        private void ParseUntermerkmale()
        {
            for (int i = 3; i <= worksheetCharacteristics.Dimension.Rows; i++)
            {
                var cellVal = worksheetCharacteristics.Cells[i, _untermerkmaleColIndexExcel.GetValueOrDefault()].Value;
                if (cellVal != null)
                {
                    var merkmalKeyGroup = worksheetCharacteristics.Cells[i, _merkmaleColIndexExcel].Value?.ToString().Trim();
                    if (merkmalKeyGroup != null)
                    {
                        string untermerkmalKeyName = cellVal.ToString().Trim();
                        string merkmalKeyGroupName = merkmalKeyGroup?.ToString().Trim();
                        //GET ROWINDEX WHERE KEYGROUP,PARENTKEYGROUP AND PARENTPARENTKEYGROUP ARE STORED
                        int keyGroupRowIndex = GetKeyGroupRowIndexInMerkmaleCol(i);

                        string abbString = JsonConvert.SerializeObject(null);
                        //Get DescriptionKeyType From Excel
                        /**
                        if (worksheetCharacteristics.Cells[rowIndex, _typColIndexExcel].Value != null)
                        {
                            descKeyTypeCell = worksheetCharacteristics.Cells[rowIndex, _typColIndexExcel].Value.ToString().Trim();
                        }
                        **/
                        if (worksheetCharacteristics.Cells[i, _abbColIndexExcel].Value != null)
                        {
                            abbString = ConvertAbbStringToJson(worksheetCharacteristics.Cells[i, _abbColIndexExcel].Value.ToString().Trim());
                        }

                        //USE KeyGroupRowIndex to Find all KeyGroups and ParentKeyGroups
                        string keyGroupName = worksheetCharacteristics.Cells[keyGroupRowIndex, _merkmaleColIndexExcel].Value?.ToString().Trim();
                        string supKeyGroupVal = worksheetCharacteristics.Cells[keyGroupRowIndex, _lageColIndexExcel].Value?.ToString().Trim();
                        string supSupKeyGroupVal = worksheetCharacteristics.Cells[keyGroupRowIndex, _regionColIndexExcel].Value?.ToString().Trim();
                        var parentDescKeyId = GetParentKeyDescIdForUntermerkmal(merkmalKeyGroupName, keyGroupName, supKeyGroupVal, supSupKeyGroupVal);

                        if (parentDescKeyId != null)
                        {
                            int? keyGroupId = GetDescKeyGroupIdFromDescKeyId(parentDescKeyId);

                            if (keyGroupId != null)
                            {
                                if (!UntermerkmalExistsInContext(untermerkmalKeyName, keyGroupId, parentDescKeyId))
                                {
                                    DescriptionKey dk = new DescriptionKey { KeyName = untermerkmalKeyName, DescriptionKeyGroupId = keyGroupId.GetValueOrDefault(), ListSourceJson = abbString, ParentDescriptionKeyId = parentDescKeyId.GetValueOrDefault() };
                                    _context.Add(dk);
                                    SaveToContext();
                                }
                            }
                        }
                        //int? keyGroupId = GetCurrentKeyGroupIdFromExcel(i);
                        //int? descKeyId = GetCurrentDescriptionKeyIdFromExcel(i);
                    }
                }
            }
        }

        private bool UntermerkmalExistsInContext(string untermerkmalKeyName, int? keyGroupId, int? parentDescKeyId)
        {
            var dk_untermerkmal = _context.DescriptionKey.Where(dk => dk.KeyName == untermerkmalKeyName && dk.DescriptionKeyGroupId == keyGroupId && dk.ParentDescriptionKeyId == parentDescKeyId).FirstOrDefault();
            if (dk_untermerkmal != null)
            {
                return true;
            }
            return false;
        }

        private int? GetDescKeyGroupIdFromDescKeyId(int? parentDescKeyId)
        {
            var parentDk = _context.DescriptionKey.Where(dk => dk.DescriptionKeyId == parentDescKeyId).FirstOrDefault();
            if (parentDk != null)
            {
                return parentDk.DescriptionKeyGroupId;
            }
            return null;
        }

        private int? GetParentKeyDescIdForUntermerkmal(string merkmalKeyName, string keyGroupName, string supKeyGroupName, string supSupKeyGroupName = null)
        {
            //TODO: how to get this done with one query?
            if (merkmalKeyName != null && keyGroupName != null && supKeyGroupName != null)
            {
                if (supSupKeyGroupName == null)
                {
                    //TODO: test
                    var dk = _context.DescriptionKey.Where(dk => dk.KeyName == merkmalKeyName)
                        .Include(dk => dk.DescriptionKeyGroup)
                        .Where(dk => dk.DescriptionKeyGroup.KeyGroupName == keyGroupName).FirstOrDefault();

                    if (dk != null)
                    {
                        return dk.DescriptionKeyId;
                    }
                }
                else
                {
                    var dk = _context.DescriptionKey
                        .Where(dk => dk.KeyName == merkmalKeyName)
                        .Include(dk => dk.DescriptionKeyGroup)
                            .ThenInclude(dkg => dkg.ParentDescriptionKeyGroup).ToList();

                    if (dk != null)
                    {
                        if (dk.Count > 1)
                        {
                            var dkg_filtered = dk.Where(dk => dk.DescriptionKeyGroup.KeyGroupName == keyGroupName);
                            if (dkg_filtered != null)
                            {
                                if (dkg_filtered.Count() > 1)
                                {
                                    var parent_dkg_filtered = dkg_filtered.Where(dkg_filtered => dkg_filtered.DescriptionKeyGroup.KeyGroupName == supKeyGroupName);
                                    if (parent_dkg_filtered != null)
                                    {
                                        if (parent_dkg_filtered.Count() > 1)
                                        {
                                            Logger.Warn("--- Could not find unique DescriptionKey Id for Merkmal: \'" + merkmalKeyName + "\' KeyGroup: \'" + keyGroupName + "\' ParentKeyGroupName: \'" + supKeyGroupName + "\' ParentKeyGroupName: \'" + supSupKeyGroupName + "\'");
                                            Logger.Warn("--- Please Check Excel and Database");
                                        }
                                        else
                                        {
                                            return dkg_filtered.FirstOrDefault().DescriptionKeyId;
                                        }
                                    }
                                }
                                else
                                {
                                    return dkg_filtered.FirstOrDefault().DescriptionKeyId;
                                }
                            }
                        }
                        else if (dk.Count < 1)
                        {
                            try
                            {
                                return dk.FirstOrDefault().DescriptionKeyId;
                            }
                            catch (Exception e)
                            {
                                Logger.Warn("--- Could not find unique DescriptionKey Id for Merkmal: \'" + merkmalKeyName + "\' KeyGroup: \'" + keyGroupName + "\' ParentKeyGroupName: \'" + supKeyGroupName + "\' ParentKeyGroupName: \'" + supSupKeyGroupName + "\'");
                                Logger.Warn("--- Please Check Excel and Database");
                            }
                        }
                        else
                        {
                            dk.Select(dki => dki.DescriptionKeyId);
                        }
                    }
                }
            }

            return null;
        }

        public void ParseAndSaveExcelColumn(int currentKeyGroupRow, int superiorKeyGroupRow, bool isMerkmaleRow)
        {
            for (int i = 3; i <= worksheetCharacteristics.Dimension.Rows; i++)
            {
                //Parse Region Col
                var cellVal = worksheetCharacteristics.Cells[i, currentKeyGroupRow].Value;
                //Empty row above -> new DescKeyGroup --> ignore desckeys in merkmale-row
                if (cellVal != null && worksheetCharacteristics.Cells[i - 1, currentKeyGroupRow].Value == null)
                {
                    var keyGroupName = worksheetCharacteristics.Cells[i, currentKeyGroupRow].Value.ToString().Trim();

                    //Type has to be set only on merkmale level?
                    var keyGroupDataType = isMerkmaleRow ? GetDescKeyGroupDataType(i) : null;
                    //Männchen;Weibchen --> moved to TaxDescr
                    //var keyGroupType = isMerkmaleRow && _mwColExists ? GetDescKeyGroupMWType(i) : null;

                    //Parent KeyGroup?
                    var supKeyGroupVal = worksheetCharacteristics.Cells[i, superiorKeyGroupRow].Value;

                    //If Lage empty -> get Value from Region Column and reference it
                    //cf. Row 40 "Antenne: Anzahl Antennenglieder
                    if (isMerkmaleRow && supKeyGroupVal == null)
                    {
                        //supKeyGroupVal = worksheetCharacteristics.Cells[i, _regionColIndexExcel].Value;
                    }
                    int? supSupKeyGroupId = null;
                    //Merkmal -> Lage
                    if (supKeyGroupVal != null)
                    {
                        string supKeyGroupName = supKeyGroupVal.ToString().Trim();
                        if (supKeyGroupName == "Bundesland")
                        {
                            keyGroupName = _mapBundeslandRealName(keyGroupName);
                        }
                        //Problem: Unterseite -> Körperende || Unterseite -> Kopf
                        //check first column desckeygroup
                        if (isMerkmaleRow)
                        {
                            var supSupKeyGroupName = worksheetCharacteristics.Cells[i, _regionColIndexExcel].Value?.ToString().Trim();
                            supSupKeyGroupId = GetDescKeyGroupIdByName(supSupKeyGroupName).GetValueOrDefault();
                        }
                        int? supKeyGroupId = null;
                        supKeyGroupId = GetDescKeyGroupIdByName(supKeyGroupName.ToString().Trim(), supSupKeyGroupId).GetValueOrDefault();
                        SaveDKGWithParent(keyGroupName, supKeyGroupId, supSupKeyGroupId, supKeyGroupName, isMerkmaleRow, i, currentKeyGroupRow, keyGroupDataType);
                    }
                    //Lage empty, check Region and relate merkmal directly to Region
                    else if (worksheetCharacteristics.Cells[i, _regionColIndexExcel].Value != null)
                    {
                        //tbd
                        string supKeyGroupName = worksheetCharacteristics.Cells[i, _regionColIndexExcel].Value.ToString().Trim();
                        var supKeyGroupId = GetDescKeyGroupIdByName(supKeyGroupName);
                        if (!SupDescKeyGroupNameExists(keyGroupName, supKeyGroupId))
                        {
                            DescriptionKeyGroup descriptionKeyGroup = new DescriptionKeyGroup { KeyGroupName = keyGroupName, ParentDescriptionKeyGroupId = supKeyGroupId, DescriptionKeyGroupDataType = keyGroupDataType, VisibilityCategoryId = _MapAmpelType(i), OrderPriority = _GetOrderPriority(i) };
                            _context.Add(descriptionKeyGroup);
                            SaveToContext();
                            Logger.Debug("--- Added DescKeyGroup \'" + keyGroupName + "\' With ParentKeyGroup: \'" + supKeyGroupName);
                            //Save merkmale keyGroupId to row for reference when parsing untermerkmale
                            if (isMerkmaleRow)
                            {
                                _keyGroupToExcelRowMapping[i] = descriptionKeyGroup.DescriptionKeyGroupId;
                                if (worksheetCharacteristics.Cells[i, _typColIndexExcel].Value?.ToString().Trim() == "Zahlenfeld" && worksheetCharacteristics.Cells[i + 1, currentKeyGroupRow].Value == null)
                                {
                                    SaveDescriptionKey(descriptionKeyGroup.KeyGroupName, descriptionKeyGroup.KeyGroupName, descriptionKeyGroup.DescriptionKeyGroupId, i);
                                }
                            }
                            descKeyGroupCounter++;
                        }
                        else
                        {
                            Logger.Debug("--- Did not add DescKeyGroup \'" + keyGroupName + "\' With ParentKeyGroup: \'" + supKeyGroupName);
                            Logger.Debug("--- Entity already exists");
                        }
                    }
                    //merkmal as keygroup without Lage nor Region
                    else
                    {
                        DescriptionKeyGroup descriptionKeyGroup = new DescriptionKeyGroup { KeyGroupName = keyGroupName, DescriptionKeyGroupDataType = keyGroupDataType, VisibilityCategoryId = _MapAmpelType(i), OrderPriority =  _GetOrderPriority(i)};
                        _context.Add(descriptionKeyGroup);
                        SaveToContext();
                        Logger.Debug("--- Added DescKeyGroup \'" + keyGroupName + " with no ParentKeyGroup");
                        //Save merkmale keyGroupId to row for reference when parsing untermerkmale
                        if (isMerkmaleRow)
                        {
                            _keyGroupToExcelRowMapping[i] = descriptionKeyGroup.DescriptionKeyGroupId;
                            if (worksheetCharacteristics.Cells[i, _typColIndexExcel].Value?.ToString().Trim() == "Zahlenfeld" && worksheetCharacteristics.Cells[i + 1, currentKeyGroupRow].Value == null)
                            {
                                SaveDescriptionKey(descriptionKeyGroup.KeyGroupName, descriptionKeyGroup.KeyGroupName, descriptionKeyGroup.DescriptionKeyGroupId, i);
                            }
                        }
                        descKeyGroupCounter++;
                    }
                }
            }
        }

        private int? _GetOrderPriority(int i)
        {
            if(i > 0 && _orderPrioColIndexExcel > 0)
            {
                if (Int32.TryParse(worksheetCharacteristics.Cells[i, _orderPrioColIndexExcel].Value?.ToString(), out var result)) {
                    return result;
                }
            } else
            {
                return 2;
            }
            return 2;
        }

        private void SaveDKGWithParent(string keyGroupName, int? supKeyGroupId, int? supSupKeyGroupId, string supKeyGroupName, bool isMerkmaleRow, int i, int currentKeyGroupRow, string? keyGroupDataType)
        {
            if (!SupDescKeyGroupNameExists(keyGroupName, supKeyGroupId, supSupKeyGroupId))
            {
                DescriptionKeyGroup descriptionKeyGroup = null;
                //eg "Oberseite;Unterseite"
                descriptionKeyGroup = new DescriptionKeyGroup { KeyGroupName = keyGroupName, ParentDescriptionKeyGroupId = supKeyGroupId, DescriptionKeyGroupDataType = isMerkmaleRow ? keyGroupDataType : null, VisibilityCategoryId = _MapAmpelType(i), OrderPriority = _GetOrderPriority(i) };
                _context.Add(descriptionKeyGroup);
                Logger.Debug("--- Saving DescKeyGroup \'" + keyGroupName + "\' with ParentDescKeyGroup \'" + supKeyGroupName + "\' to Context ...");
                SaveToContext();

                //Save merkmale keyGroupId to row for reference when parsing untermerkmale
                if (isMerkmaleRow)
                {
                    _keyGroupToExcelRowMapping[i] = descriptionKeyGroup.DescriptionKeyGroupId;

                    //Check if DataType = Value (Zahlenfeld) AND no DescriptionKey DataCell below --> Create DescKey with same name (eg. Einzelaugen: Anzahl Augen)
                    if (worksheetCharacteristics.Cells[i, _typColIndexExcel].Value?.ToString().Trim() == "Zahlenfeld" && worksheetCharacteristics.Cells[i + 1, currentKeyGroupRow].Value == null)
                    {
                        try
                        {
                            SaveDescriptionKey(descriptionKeyGroup.KeyGroupName, descriptionKeyGroup.KeyGroupName, descriptionKeyGroup.DescriptionKeyGroupId, i);
                        }
                        catch (Exception e)
                        {
                            Logger.Error("--- Error saving \'" + keyGroupName + "\' with ParentDescKeyGroup \'" + supKeyGroupName + "\'");
                            Logger.Error(e);
                        }
                    }
                }
                descKeyGroupCounter++;
            }
            else
            {
                Logger.Debug("--- Did not add KeyGroup: \'" + keyGroupName + "\' With ParentKeyGroup: \'" + supKeyGroupName + "\'");
                Logger.Debug("--- Entity already exists.");
            }
            //TODO: clarify - save Lage-Key groups without Region Key Groups as top-level keyGroups?
        }

        private bool isRegionOnlyKeyGroup(int rowIndex)
        {
            //Check if region column is set and all other empty -> get datatype (cf. land,bundesland,biotop)
            //4 empty cols until datatype
            for (int i = _lageColIndexExcel; i < _typColIndexExcel; i++)
            {
                if (worksheetCharacteristics.Cells[rowIndex, i].Value != null)
                {
                    return false;
                }
            }

            return true;
        }

        private string GetDescKeyGroupDataType(int rowIndex)
        {
            //Currently keyTypes in excel next to descriptionKeys - may be modified / moved up
            //TODO: how far to search for type, ie. how many cells shall be checked?
            string descKeyGroupType = null;
            for (int i = 0; i < 3; i++)
            {
                if (worksheetCharacteristics.Cells[rowIndex + i, _typColIndexExcel].Value != null)
                {
                    descKeyGroupType = worksheetCharacteristics.Cells[rowIndex + i, _typColIndexExcel].Value.ToString().Trim();
                    return MapDescKeyDataType(descKeyGroupType);
                }
            }
            return MapDescKeyDataType(descKeyGroupType);
        }

        private string MapDescKeyDataType(string mapInput)
        {
            //TODO: what default value?
            string mapOutput = "UNKNOWN";
            switch (mapInput)
            {
                case "Auswahl":
                    mapOutput = "VALUELIST";
                    break;

                case "Auswahl Offen":
                    mapOutput = "OPEN VAL";
                    break;

                case "Bildauswahl":
                    mapOutput = "PIC";
                    break;

                case "Zahlenfeld":
                    mapOutput = "VALUE";
                    break;
                //TODO: clarify what will be mapped to "YESNO" & replace below
                case "Checkbox":
                    mapOutput = "YESNO";
                    break;

                default:
                    break;
            }
            return mapOutput;
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

        private void ImportDescKeys()
        {
            string currentKeyGroupName = null;

            //Skip Header Row and empty Row
            for (int i = 3; i <= worksheetCharacteristics.Dimension.Rows; i++)
            {
                //Parse Merkmale Col
                var cellVal = worksheetCharacteristics.Cells[i, _merkmaleColIndexExcel].Value;
                //speaking names!
                //we found a descriptionKeyGroup cell!
                if (cellVal != null && worksheetCharacteristics.Cells[i - 1, _merkmaleColIndexExcel].Value == null)
                {
                    currentKeyGroupName = cellVal.ToString().Trim();
                    //currentKeyGroupName = ConcatKeyGroupName(i);
                }
                //We found a descriptionKey cell!
                else if (cellVal != null && CheckLeftColsEmpty(i))
                {
                    var descKeyName = cellVal.ToString().Trim();
                    if (currentKeyGroupName == "Bundesland")
                    {
                        descKeyName = _mapBundeslandRealName(descKeyName);
                    }
                    SaveDescriptionKeyFromExcelValues(descKeyName, currentKeyGroupName, i);
                }
                //end of desckeygroup (empty row) --> reset keygroupname
                else
                {
                    currentKeyGroupName = null;
                }
            }
            if (_untermerkmaleColIndexExcel != null)
            {
                Logger.Info("-- Started Untermerkmale Import");
                ParseUntermerkmale();
                Logger.Info("-- Finished Untermerkmale Import");
            }
        }

        private bool SupDescKeyGroupNameExists(string descKeyGroupName, int? supDescKeyGroupId, int? supSupKeyGroupId = null)
        {
            if (supSupKeyGroupId != null)
            {
                var descKeyGroup = _context.DescriptionKeyGroup
                    .Where(dkg => dkg.KeyGroupName == descKeyGroupName && dkg.ParentDescriptionKeyGroupId == supDescKeyGroupId)
                    .Include(dkg => dkg.ParentDescriptionKeyGroup)
                    .FirstOrDefault();
                if (descKeyGroup != null && descKeyGroup.ParentDescriptionKeyGroup.ParentDescriptionKeyGroupId == supSupKeyGroupId)
                {
                    return true;
                }
            }
            else
            {
                if (_context.DescriptionKeyGroup.Where(dkg => dkg.KeyGroupName == descKeyGroupName && dkg.ParentDescriptionKeyGroupId == supDescKeyGroupId).FirstOrDefault() != null)
                {
                    return true;
                }
            }
            return false;
        }

        public void SaveDescriptionKey(string descKeyName, string keyGroupName, int? descKeyGroupId, int rowIndex)
        {
            if (descKeyGroupId != null)
            {
                if (IsNewDescKeyEntity(descKeyName, descKeyGroupId) && descKeyName.Count() > 2)
                {
                    if (Regex.IsMatch(descKeyName, _dkNameRegex) && !String.IsNullOrEmpty(descKeyName))
                    {
                        DescriptionKey descKey = new DescriptionKey { KeyName = descKeyName, DescriptionKeyGroupId = descKeyGroupId.GetValueOrDefault() };
                        _context.Add(descKey);
                        Logger.Debug("--- Saving descKeyName \'" + descKeyName + "\' descKeyGroup \'" + keyGroupName + "\' to Context ..");
                        SaveToContext();
                        _descKeyToExcelRowMapping[rowIndex] = descKey.DescriptionKeyId;
                        descKeyCounter++;
                    }
                    else
                    {
                        Logger.Warn("--- Did not save descKeyName \'" + descKeyName + "\' descKeyGroup \'" + keyGroupName + "\'. Please review Name.");
                    }
                }
                else
                {
                    Logger.Debug("--- Did not add Key: \'" + descKeyName + "\' AND KeyGroup: \'" + keyGroupName + "\'");
                    Logger.Debug("--- Entity already exists.");
                }
            }
        }

        public int? _MapAmpelType(int i)
        {
            if (_ampelColIndexExcel > 0)
            {
                string ampelString = worksheetCharacteristics.Cells[i, _ampelColIndexExcel].Value?.ToString().Trim();
                if (ampelString != null)
                {
                    /*@TODO: only temporary until orange -> gelb */
                    if (ampelString.Trim() == "orange")
                    {
                        return 2;
                    }
                    else if (ampelString.Length > 0 && _ampelTypeToIdDict.TryGetValue(ampelString, out int ampelId))
                    {
                        return ampelId;
                    }
                }
            }
            return null;
        }

        public void SaveDescriptionKeyFromExcelValues(string descKeyName, string keyGroupName, int rowIndex)
        {
            try
            {
                int? descKeyGroupId = GetDescKeyGroupIdByName(keyGroupName);
                //KeyType now at KeyGroup Level - can be commented in if needed on Key level
                //var descKeyTypeCell = "UNKNOWN";
                string abbString = JsonConvert.SerializeObject(null);

                if (_untermerkmaleColIndexExcel != null)
                {
                    if (worksheetCharacteristics.Cells[rowIndex, _abbColIndexExcel].Value != null && worksheetCharacteristics.Cells[rowIndex, _untermerkmaleColIndexExcel.GetValueOrDefault()].Value == null)
                    {
                        abbString = ConvertAbbStringToJson(worksheetCharacteristics.Cells[rowIndex, _abbColIndexExcel].Value.ToString().Trim());
                    }
                }
                //var descKeyType = MapDescKeyDataType(descKeyTypeCell);

                if (descKeyGroupId != null)
                {
                    if (IsNewDescKeyEntity(descKeyName, descKeyGroupId) && descKeyName.Count() > 0)
                    {
                        string keyDescription = null;
                        if(_keyDesciptionColIndexExcel != 0)
                        {
                            keyDescription = worksheetCharacteristics.Cells[rowIndex, _keyDesciptionColIndexExcel].Value != null? worksheetCharacteristics.Cells[rowIndex, _keyDesciptionColIndexExcel].Value.ToString():null;
                        }
                        DescriptionKey descKey = new DescriptionKey { KeyName = descKeyName, DescriptionKeyGroupId = descKeyGroupId.GetValueOrDefault(), ListSourceJson = abbString, KeyDescription = keyDescription };
                        _context.Add(descKey);
                        Logger.Debug("--- Saving descKeyName \'" + descKeyName + "\' descKeyGroup \'" + keyGroupName + "\' to Context ..");
                        SaveToContext();
                        _descKeyToExcelRowMapping[rowIndex] = descKey.DescriptionKeyId;
                        descKeyCounter++;
                    }
                    else
                    {
                        Logger.Debug("--- Did not add Key: \'" + descKeyName + "\' AND KeyGroup: \'" + keyGroupName + "\'");
                        Logger.Debug("--- Entity already exists.");
                    }
                }
                else
                {
                    //TODO: proper error handling
                    Logger.Warn("--- Not able to get DescKeyGroupId for Key: \'" + descKeyName + "\' AND KeyGroup: \'" + keyGroupName + "\'");
                }
            }
            catch (Exception e)
            {
                Logger.Error(e, "-- DescKeyGroup Not Found for DescKey " + descKeyName);
                Logger.Error("-- At Row:" + rowIndex + " " + " Column: " + _merkmaleColIndexExcel);
                Logger.Error("-- Please review the ExcelSheet: " + worksheetCharacteristics.Name);
            }
        }

        private bool IsNewDescKeyEntity(string descKeyName, int? descKeyGroupId)
        {
            var test = _context.DescriptionKey.Where(dk => dk.KeyName == descKeyName && dk.DescriptionKeyGroupId == descKeyGroupId).FirstOrDefault();
            if (_context.DescriptionKey.Where(dk => dk.KeyName == descKeyName && dk.DescriptionKeyGroupId == descKeyGroupId).FirstOrDefault() == null)
            {
                return true;
            }
            return false;
        }

        private int GetKeyGroupRowIndexInMerkmaleCol(int rowIndex)
        {
            //go back up to find last not-null cell --> keygroupname!
            int notNullRowIndex = rowIndex;
            int i = 0;
            while (worksheetCharacteristics.Cells[rowIndex - i, _merkmaleColIndexExcel].Value != null)
            {
                notNullRowIndex = rowIndex - i;
                i++;
            }

            return notNullRowIndex;
        }

        private bool CheckLeftColsEmpty(int rowIndex)
        {
            if ((worksheetCharacteristics.Cells[rowIndex, _merkmaleColIndexExcel - 1].Value == null)
                && (worksheetCharacteristics.Cells[rowIndex, _lageColIndexExcel].Value == null)
                && (worksheetCharacteristics.Cells[rowIndex, _regionColIndexExcel].Value == null))
            {
                return true;
            }

            return false;
        }

        private int? GetDescKeyGroupIdByName(string descKeyGroupName, int? superDescKeyGroupid = null)
        {
            DescriptionKeyGroup dkg = null;
            if (superDescKeyGroupid != null)
            {
                dkg = _context.DescriptionKeyGroup.Where(dkg => dkg.KeyGroupName == descKeyGroupName && dkg.ParentDescriptionKeyGroupId == superDescKeyGroupid).FirstOrDefault();
            }
            else
            {
                dkg = _context.DescriptionKeyGroup.Where(dkg => dkg.KeyGroupName == descKeyGroupName).FirstOrDefault();
            }
            if (dkg != null)
            {
                return dkg.DescriptionKeyGroupId;
            }
            return null;
        }

        private bool DescKeyGroupNameExists(string keyGroupName)
        {
            var result = _context.DescriptionKeyGroup.Where(dkg => dkg.KeyGroupName == keyGroupName).FirstOrDefault();
            return result == null ? false : true;
        }

        private bool DescKeyNameExists(string keyName)
        {
            var result = _context.DescriptionKey.Where(dk => dk.KeyName == keyName).FirstOrDefault();
            return result == null ? false : true;
        }

        private void FillContext()
        {
            //GET ALL DESCRKEYGROUPS
            var descKeyGroupEntity = _context.DescriptionKeyGroup.ToList();//   DescriptionKeyGroup. Where(dkg => dkg.DescriptionKeyGroupId != null).ToList();
            //
            if (descKeyGroupEntity.Count < 1)
            {
                Logger.Info("-- Description Key Group Table Empty");
            }
        }

        private void ParseHeaderRow()
        {
            for (int i = 1; i <= worksheetCharacteristics.Dimension.Columns; i++)
            {
                var cellVal = worksheetCharacteristics.Cells[1, i].Value;
                //"empty" cells -> null
                if (cellVal != null)
                {
                    AssignHeaderIndexToProperty(cellVal.ToString().Trim(), i);
                }
            }
            _mwColExists = _merkmaleColIndexExcel - _lageColIndexExcel < 1 ? false : true;
        }

        private void AssignHeaderIndexToProperty(string cellValueString, int index)
        {
            switch (cellValueString)
            {
                case "Region":
                    _regionColIndexExcel = index;
                    break;

                case "Lage":
                    _lageColIndexExcel = index;
                    break;

                case "Merkmale":
                    _merkmaleColIndexExcel = index;
                    break;

                case "Untermerkmal":
                    _untermerkmaleColIndexExcel = index;
                    break;

                case "Typ":
                    _typColIndexExcel = index;
                    break;

                case "Ampel":
                    _ampelColIndexExcel = index;
                    break;

                case "Abb.":
                    _abbColIndexExcel = index;
                    break;

                case "Priorität 1":
                    _orderPrioColIndexExcel = index;
                    break;

                case "Erklärung":
                    _keyDesciptionColIndexExcel = index;
                    break;

                default:
                    break;
            }
        }

        private void SaveToContext()
        {
            try
            {
                _context.SaveChanges();
                Logger.Debug(".. Saved successfully.");
            }
            catch (Exception e)
            {
                Logger.Error(e.InnerException, "Error Saving Changes to Context");
            }
        }

        private int? GetCurrentKeyGroupIdFromExcel(int rowIndex)
        {
            int? keyGroupId = null;
            keyGroupId = GetClosestValue(_keyGroupToExcelRowMapping, rowIndex);
            return keyGroupId;
        }

        private int? GetCurrentDescriptionKeyIdFromExcel(int rowIndex)
        {
            int? keyGroupId = null;
            keyGroupId = GetClosestValue(_descKeyToExcelRowMapping, rowIndex);
            return keyGroupId;
        }

        private string ConvertAbbStringToJson(string abbString)
        {
            var abbArray = abbString.Split(';');
            //TODO: key?
            var abbJsonString = JsonConvert.SerializeObject(abbArray.ToList());
            return abbJsonString;
        }

        public static int GetClosestValue(int[] myArray, int myValue)
        {
            //optional
            int i = 0;

            while (myArray[++i] < myValue) ;

            return myArray[--i];
        }

        /**TEMP HELPER - tbd **/

        public static string Truncate(string value, int maxLength)
        {
            if (!string.IsNullOrEmpty(value) && value.Length > maxLength)
            {
                return value.Substring(0, maxLength);
            }

            return value;
        }
    }
}