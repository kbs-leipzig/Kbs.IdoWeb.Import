using Kbs.IdoWeb.Data.Determination;
using Kbs.IdoWeb.Data.Information;
using Kbs.IdoWeb.Data.Observation;
using Microsoft.EntityFrameworkCore;
using Microsoft.Extensions.Configuration;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace Kbs.IdoWeb.Import.Models
{
    class AppData
    {
        private DeterminationContext _detContext = new DeterminationContext();
        private InformationContext _infContext = new InformationContext();
        private ObservationContext _obsContext = new ObservationContext();
        private static readonly NLog.Logger Logger = NLog.LogManager.GetCurrentClassLogger();
        private static string edaphoApi = @"https://api.edaphobase.org/taxon/";
        private static string _filePath = Path.Combine("AppFiles/");
        private List<string> taxonTagsFilterGroups = new List<string> { "Bodentiere", "Doppelfüßer (Diplopoda)", "Samenfüßer (Chordeumatida)", "Bandfüßer (Polydesmida)", "Schnurfüßer (Julida)", "Saftkugler (Glomerida)", "Pinselfüßer (Polyxenida)", "Bohrfüßer (Polyzoniida)", "Hundertfüßer (Chilopoda)", "Steinläufer (Lithobiomorpha)", "Skolopender (Scolopendromorpha)", "Erdkriecher (Geophilomorpha)", "Spinnenläufer (Scutigeromorpha)", "Asseln (Isopoda)" };
        private object config;


        public AppData()
        {
            Init();
        }

        private void Init()
        {
            try
            {
                Logger.Info("Initializing AppData ..");
                var optionsBuilderDet = new DbContextOptionsBuilder<DeterminationContext>();
                optionsBuilderDet.UseNpgsql(Program.Configuration.GetConnectionString("DatabaseConnection"));
                _detContext = new DeterminationContext(optionsBuilderDet.Options);

                var optionsBuilderInf = new DbContextOptionsBuilder<InformationContext>();
                optionsBuilderInf.UseNpgsql(Program.Configuration.GetConnectionString("DatabaseConnection"));
                _infContext = new InformationContext(optionsBuilderInf.Options);

                var optionsBuilderObs = new DbContextOptionsBuilder<ObservationContext>();
                optionsBuilderObs.UseNpgsql(Program.Configuration.GetConnectionString("DatabaseConnection"));
                _obsContext = new ObservationContext(optionsBuilderObs.Options);

                Logger.Info(".. Success");
            }
            catch (Exception e)
            {
                Logger.Error(e, "Init Import Meta failed");
            }
        }

        public void GenerateAllFiles()
        {
            try
            {
                GenerateTaxaFile();
                GenerateTaxonTagFile();
                GenerateTaxonTagFilterGroups();
                GenerateTaxonImageFile();
                GenerateTaxonImageTypeFile();
                GenerateTaxonProtectionFile();
                GenerateTaxonSynonymsFile();
                GenerateVersionsFile();
                Logger.Debug("-- Succesfully Created All Files");
            }
            catch (Exception e)
            {
                Logger.Error("-- ERROR: " + e.Message);
                throw e;
            }
        }

        public void GenerateTaxaFile()
        {
            /*
             * {"TaxonId":402930,"TaxonTypeId":20005,"HasDiagnosis":false,"Diagnosis":"","TaxonName":"Nemobius sylvestris","LocalName":"Waldgrille","FamilyName":"Gryllidae","FamilyLocalName":"Grillen","OrderName":"Orthoptera","OrderLocalName":"Heuschrecken","TaxonAuthor":"(Bosc, 1792)","IdentificationLevelFemale":1,"IdentificationLevelMale":1},
            */
            /**Get Species,Orders,Classes,Families
            **/
            var taxInfo = _infContext.Taxon
                .Include(tax => tax.Family).AsNoTracking()
                .Include(tax => tax.Order).AsNoTracking()
                .Include(tax => tax.RedListType).AsNoTracking()
                .Include(tax => tax.TaxonomyState)
                .Where(tax => (tax.TaxonomyStateId == 100 || (tax.TaxonomyStateId == 301 && tax.OrderId != null) || tax.TaxonomyStateId == 117 || tax.TaxonomyStateId == 119) || tax.Group.Contains("Bodentiere"))
                .Select(tax => new
                {
                    //TaxonDistribution -> Lebensräume; Biotope --> Verbreitung & Häufigkeit; beschreibung -> AdditionalInfo
                    TaxonId = tax.TaxonId,
                    Synonyms = tax.Synonyms != null ? "<h4>Synonyme</h4><p><em>" + String.Join("</em></p><p><em>", ConvertSynonymsToList(tax.Synonyms)) + "</em></p>" : "",
                    DisplayLength = tax.DisplayLength != null ? "<p><b>L&auml;nge:</b><br/>" + tax.DisplayLength + "</p>" : "",
                    TaxonBiotopeAndLifestyle = tax.TaxonBiotopeAndLifestyle != null ? "<p><b>Verbreitung &amp; H&auml;ufigkeit:</b><br/>" + tax.TaxonBiotopeAndLifestyle + "</p>" : "",
                    TaxonDistribution = tax.TaxonDistribution != null ? "<p><b>Lebensr&auml;ume &amp; Lebensweise:</b><br/>" + tax.TaxonDistribution + "</p>" : "",
                    AdditionalInfo = tax.AdditionalInfo != null ? "<p><b>Beschreibung:</b><br/>" + tax.AdditionalInfo + "</p>" : "",
                    Lit = tax.LiteratureSource != null ? "<p><b>Literatur:</b><br/>" + tax.LiteratureSource + "</p>" : "",
                    RedListTypeId = tax.RedListTypeId != null ? "<p><b>Rote Liste Deutschland:</b><br/>" + tax.RedListType.RedListTypeName + "</p>" : "",
                    RedListSource = tax.RedListSource != null ? "<p><b>Rote Liste Quelle:</b><br/>" + tax.RedListSource + "</p>" : "",
                    TaxonomyStateName = tax.TaxonomyState.StateName,
                    TaxonName = tax.TaxonName,
                    SliderImages = tax.SliderImages,
                    FamilyName = tax.Family.TaxonName,
                    OrderName = tax.Order.TaxonName,
                    OrderId = tax.OrderId,
                    TaxonAuthor = tax.TaxonDescription,
                }).AsNoTracking().ToList();

            List<TaxonInfo> result = new List<TaxonInfo>();

            foreach (var tInfo in taxInfo)
            {
                TaxonInfo newTaxonInfo = new TaxonInfo();
                string content = "";
                content += "<h4>Allgemeines</h4>";
                content += tInfo.DisplayLength + tInfo.TaxonBiotopeAndLifestyle + tInfo.TaxonDistribution + tInfo.AdditionalInfo + tInfo.Lit + tInfo.RedListTypeId + tInfo.RedListSource;
                content += tInfo.Synonyms;
                newTaxonInfo.Diagnosis = content;
                newTaxonInfo.TaxonId = tInfo.TaxonId;
                newTaxonInfo.TaxonName = tInfo.TaxonName;
                newTaxonInfo.LocalName = tInfo.TaxonName;
                newTaxonInfo.TaxonomyStateName = tInfo.TaxonomyStateName;
                newTaxonInfo.FamilyName = tInfo.FamilyName;
                newTaxonInfo.FamilyLocalName = tInfo.FamilyName;
                newTaxonInfo.OrderName = tInfo.OrderName;
                newTaxonInfo.OrderLocalName = tInfo.OrderName;
                newTaxonInfo.OrderId = tInfo.OrderId != null ? tInfo.OrderId : 0;
                newTaxonInfo.TaxonAuthor = tInfo.TaxonAuthor != null ? tInfo.TaxonAuthor : "";
                newTaxonInfo.SliderImages = tInfo.SliderImages != null ? JsonConvert.DeserializeObject(tInfo.SliderImages) : "";
                result.Add(newTaxonInfo);
            }

            //var taxaJson = JsonConvert.SerializeObject(taxa, Formatting.None);
            _writeFile(result, "Taxa.json");

        }

        public List<TaxonDescriptionToTagFilterObject> GetTaxDescStateLevels(string toplevel, string lowLevel, string groupFilter)
        {
            try
            {
                /**TODO: enable below when switching to taxon level pre-selection**/
                //List<int> taxonDescTaxIds = _infContext.Taxon.Where(td => td.HasTaxDescChildren.HasValue).Select(tax => tax.TaxonId).Distinct().ToList();
                List<int> _uniqueTaxonIdList = _detContext.TaxonDescription.Select(tax => tax.TaxonId).Distinct().ToList();

                //get all taxondescriptions available in table
                if (toplevel != null)
                {
                    var topLevelTaxonId = _infContext.Taxon.Where(tax => tax.TaxonName == toplevel).Select(tax => tax.TaxonId).FirstOrDefault();
                    var topLevelStateId = _infContext.Taxon.Where(tax => tax.TaxonName == toplevel).Select(tax => tax.TaxonomyStateId).FirstOrDefault();
                    /** Default TaxonomyStateId?? **/
                    int targetLevelStateId = 0;
                    switch (topLevelStateId)
                    {
                        case 119:
                            targetLevelStateId = 117;
                            break;
                        case 117:
                            targetLevelStateId = 301;
                            break;
                        default:
                            targetLevelStateId = 301;
                            break;

                    }

                    if (topLevelTaxonId != 0)
                    {
                        var taxonIds = _infContext.Taxon
                            .Where(tax => (tax.KingdomId == topLevelTaxonId || tax.PhylumId == topLevelTaxonId || tax.ClassId == topLevelTaxonId || tax.OrderId == topLevelTaxonId || tax.FamilyId == topLevelTaxonId) && tax.TaxonomyStateId == targetLevelStateId)
                            .Select(tax => tax.TaxonId).ToList();
                        //taxonIds.Add(topLevelTaxonId);
                        //List<int> taxonTopLevelFiltered = _infContext.Taxon.Where(tax => taxonIds.Contains(tax.TaxonId)).Select(tax => tax.TaxonId).Distinct().ToList();
                        _uniqueTaxonIdList = _uniqueTaxonIdList.Intersect(taxonIds).ToList();
                    }
                }
                if (lowLevel != null)
                {
                    var hierarchyLevel = _infContext.TaxonomyState.Where(ts => lowLevel.Trim().ToLower() == ts.StateDescription).Select(ts => ts.HierarchyLevel).FirstOrDefault();
                    var higherHierarchyLevels = _infContext.TaxonomyState.Where(ts => hierarchyLevel > ts.HierarchyLevel).OrderByDescending(ts => ts.HierarchyLevel).Select(ts => ts.StateId).ToList();
                    if (higherHierarchyLevels.Count > 1)
                    {
                        higherHierarchyLevels.RemoveAt(0);
                    }

                    var taxon_lowLevelFiltered = _infContext.Taxon
                        .Where(tax => higherHierarchyLevels.Contains(tax.TaxonomyStateId.Value))
                        .Select(tx => tx.TaxonId).ToList();
                    _uniqueTaxonIdList = _uniqueTaxonIdList.Intersect(taxon_lowLevelFiltered).ToList();
                }
                if (groupFilter != null)
                {
                    var taxon_dropDFiltered = _infContext.Taxon
                        .Where(tax => tax.Group.Contains(groupFilter.Trim()))
                        .Select(tx => tx.TaxonId).ToList();
                    _uniqueTaxonIdList = _uniqueTaxonIdList.Intersect(taxon_dropDFiltered).ToList();
                }
                else
                {
                    var taxon_dropDFiltered = _infContext.Taxon
                        .Where(tax => tax.Group.Contains("regular"))
                        .Select(tx => tx.TaxonId).ToList();
                    _uniqueTaxonIdList = _uniqueTaxonIdList.Intersect(taxon_dropDFiltered).ToList();
                }

                var result = _detContext.TaxonDescription.Where(tax => _uniqueTaxonIdList.Contains(tax.TaxonId)).Select(tax => new TaxonDescriptionToTagFilterObject { TaxonId = tax.TaxonId, DescriptionKeyId = tax.DescriptionKeyId, DescriptionKeyGroupId = tax.DescriptionKey.DescriptionKeyGroupId }).Distinct().ToList();
                //result.OrderBy(res => res.TaxonomyStateId);
                return result;

            }
            catch (Exception e)
            {
                var exp = e;
                throw (e);
            }
        }

        public void GenerateTaxonTagFilterGroups()
        {
            List<TaxonTagFilterGroup> result = new List<TaxonTagFilterGroup>();
            foreach (string topFilterName in taxonTagsFilterGroups)
            {
                TaxonTagFilterGroup temp = new TaxonTagFilterGroup();
                temp.GroupName = topFilterName;
                List<TaxonDescriptionToTagFilterObject> tempTdList = new List<TaxonDescriptionToTagFilterObject>();
                if (topFilterName != "Bodentiere")
                {
                    tempTdList = GetTaxDescStateLevels(topFilterName, null, null);
                }
                else
                {
                    tempTdList = GetTaxDescStateLevels(null, null, topFilterName);
                }
                temp.TaxonIds = tempTdList.Select(td => td.TaxonId).Distinct().ToList();
                temp.DKIds = tempTdList.Select(td => td.DescriptionKeyId).Distinct().ToList();
                temp.DKGIds = tempTdList.Select(td => td.DescriptionKeyGroupId).Distinct().ToList();
                result.Add(temp);
            }

            _writeFile(result, "TaxonTagFilterFilterGroups.json");
        }


        public void GenerateTaxonTagFile()
        {
            //{"TagId":101, "TaxonId":441582, "TaxonTypeId":20005,"TagValue":null },

            var hitListPerDk = _detContext.TaxonDescription
                .Where(td => td.DescriptionKeyId != null)
                .Include(td => td.DescriptionKey)
                    .ThenInclude(dk => dk.DescriptionKeyGroup)
                .Where(td => td.DescriptionKey.DescriptionKeyGroup.DescriptionKeyGroupDataType == "VALUELIST" && td.DescriptionKey.DescriptionKeyGroup.VisibilityCategoryId > 1)
                .Select(td => new { td.DescriptionKeyId, td.TaxonId }).GroupBy(td => td.DescriptionKeyId, v => v.TaxonId)
                .ToDictionary(x => x.Key, x => x.ToList());

            var taxInfo = _detContext.TaxonDescription
                .Where(td => td.DescriptionKeyId != null)
                .Include(td => td.DescriptionKey)
                    .ThenInclude(dk => dk.DescriptionKeyGroup)
                .Where(td => td.DescriptionKey.DescriptionKeyGroup != null && td.DescriptionKey.DescriptionKeyGroup.VisibilityCategoryId > 1)
                .Distinct()
                .Select(td => new TaxonFilter
                {
                    TagId = td.DescriptionKeyId,
                    TagParentName = td.DescriptionKey.DescriptionKeyGroup.KeyGroupName,
                    TagParentId = td.DescriptionKey.DescriptionKeyGroup.DescriptionKeyGroupId,
                    TaxonId = td.TaxonId,
                    TaxonTypeId = 20005,
                    OrderPriority = td.DescriptionKey.DescriptionKeyGroup != null ? td.DescriptionKey.DescriptionKeyGroup.OrderPriority : 2,
                    ListSourceJson = td.DescriptionKey.ListSourceJson != null ? (List<string>)JsonConvert.DeserializeObject<List<string>>(td.DescriptionKey.ListSourceJson) : null,
                    KeyDataType = td.DescriptionKey.DescriptionKeyGroup.DescriptionKeyGroupDataType,
                    TagValue = td.DescriptionKey.KeyName,
                    MinValue = (int?)td.MinValue,
                    MaxValue = (int?)td.MaxValue,
                    TaxonHits = hitListPerDk.ContainsKey(td.DescriptionKeyId) ? hitListPerDk[td.DescriptionKeyId].ToList() : null,
                    VisibilityCategoryId = (int)td.DescriptionKey.DescriptionKeyGroup.VisibilityCategoryId != null ? td.DescriptionKey.DescriptionKeyGroup.VisibilityCategoryId : 0,
                })
                //disabled per client request 2020-07-28
                //re-enabled per client request 2021-02-14
                .Where(td => td.VisibilityCategoryId >= 2)
                .AsNoTracking();
            List<TaxonFilter> result = new List<TaxonFilter>();
            foreach (TaxonFilter uten in taxInfo)
            {
                result.Add(uten);
                if (uten.ListSourceJson != null)
                {
                    uten.ListSourceJson = uten.ListSourceJson.Take(3).ToList();
                }
                result.Add(uten);
            }
            //var taxDJson = JsonConvert.SerializeObject(taxonDesc, Formatting.None);
            _writeFile(result, "TaxonFilterItems.json");
        }


        private void GenerateTaxonImageFile()
        {
            //{"TaxonId":440867,"ImageId":4,"Index":3,"Title":"Anthocharis cardamines","Copyright":"Matthias Hartung","Description":"Weibchen des Aurorafalters im Mai 2008 bei Wei&szlig;ensand","Male":false,"Female":true,"Downside":false,"SecondForm":false },
            //Add sliderImages to TaxonImage.json
            List<TaxonImage> sliderInfo_temp = _infContext.Taxon
                .Where(t => t.SliderImages != null && t.TaxonId != null && t.TaxonName != null)
                .Select(t => new TaxonImage
                {
                    TaxonId = t.TaxonId,
                    SliderImgArray = (JArray)JsonConvert.DeserializeObject(t.SliderImages),
                    Index = 3,
                    Title = t.TaxonName,
                    Description = null,
                    Male = true,
                    Female = true,
                    Downside = false,
                    SecondForm = false
                }).AsNoTracking().ToList();

            List<TaxonImage> sliderImgs = new List<TaxonImage>();

            foreach (var sliderArray in sliderInfo_temp)
            {
                if (sliderArray.SliderImgArray != null)
                {
                    int cnt = 0;
                    //List<string> slides = sliderArray.SliderImgArray;
                    foreach (string slideName in sliderArray.SliderImgArray)
                    {
                        if (cnt < 3)
                        {
                            string slideNameClean = slideName.Trim(' ', ',');
                            sliderImgs.Add(new TaxonImage
                            {
                                TaxonId = sliderArray.TaxonId,
                                //ImageId = slideName.Trim([' ', ","]),
                                ImageId = slideNameClean,
                                Index = sliderArray.Index,
                                Title = slideNameClean,
                                Description = null,
                                Male = sliderArray.Male,
                                Female = sliderArray.Female,
                                Downside = sliderArray.Downside,
                                SecondForm = sliderArray.SecondForm,
                            });
                        }
                        else
                        {
                            break;
                        }
                        cnt++;
                    }
                }
            }

            List<TaxonImage> taxInfo = _obsContext.Image
                .Where(td => td.TaxonId != null && td.LicenseId != null && td.IsApproved)
                .Include(td => td.License)
                .Select(td => new TaxonImage
                {
                    TaxonId = td.TaxonId.HasValue ? td.TaxonId : 0,
                    ImageId = td.ImageId.ToString(),
                    Index = td.ImagePriority,
                    Title = td.ImagePath,
                    Description = $"<p>{td.Description}<br/>{td.Author}<br/>{td.CopyrightText}<br/><a href='{td.License.LicenseLink}' target='_blank'>&#169;&nbsp;{td.License.LicenseName}</a><p>",
                    Male = true,
                    Female = true,
                    Downside = false,
                    SecondForm = false
                }).OrderBy(i => i.TaxonId).ThenBy(i => i.Index).AsNoTracking().ToList();

            taxInfo.AddRange(sliderImgs);
            var test = taxInfo.GroupBy(t => t.TaxonId).ToList();
            List<TaxonImage> result = new List<TaxonImage>();
            foreach (var item in test)
            {
                result.AddRange(item.Take(3));
            }
            result.OrderBy(i => i.ImageId);

            Dictionary<string, string> imgListById = _obsContext.Image.Where(i => i.Description != null && i.LicenseId != null).Include(i => i.License).Select(i => new { ImageId = i.ImageId.ToString(), DescriptionStr = $"{i.Description}<br/>{i.Author}<br/>{i.CopyrightText}<br/><a href='{i.License.LicenseLink}' target='_blank'>{i.License.LicenseName}</a>" }).ToDictionary(x => x.ImageId, x => x.DescriptionStr);
            Dictionary<string, string> imgListByTitle = _obsContext.Image.Where(i => i.Description != null && i.LicenseId != null).Include(i => i.License).Select(i => new { ImagePath = i.ImagePath, DescriptionStr = $"{i.Description}<br/>{i.Author}<br/>{i.CopyrightText}<br/><a href='{i.License.LicenseLink}' target='_blank'>{i.License.LicenseName}</a>" }).ToDictionary(x => x.ImagePath, x => x.DescriptionStr);
            foreach (TaxonImage rItem in result)
            {
                if (rItem.Description == null)
                {
                    if (imgListById.TryGetValue(rItem.ImageId, out string desc))
                    {
                        rItem.Description = desc;
                    }
                    else if (imgListByTitle.TryGetValue(rItem.ImageId, out string desc2))
                    {
                        rItem.Description = desc2;
                    }
                    rItem.Description = "Beschreibung folgt";
                }
            }
            _writeFile(result.OrderBy(i => i.Index), "TaxonImages.json");
        }

        private List<string> ConvertSynonymsToList(string synJson)
        {
            var step1 = JArray.FromObject(JsonConvert.DeserializeObject(synJson));
            var step2 = step1.ToObject<List<string>>();
            List<string> step3 = new List<string>();
            foreach (string s in step2)
            {
                step3.Add(s.Trim());
            }
            return step3;
        }

        private void GenerateTaxonProtectionFile()
        {
            //{ "ClassId":11,"TaxonId":441347,"ClassValue":"gefährdet"},
            var taxInfo = _infContext.Taxon
                .Where(t => t.TaxonomyStateId == 301 && t.RedListTypeId != null)
                .Include(t => t.RedListType).AsNoTracking()
                .Select(td => new
                {
                    ClassId = td.RedListTypeId,
                    TaxonId = td.TaxonId,
                    ClassValue = td.RedListType.RedListTypeName
                }).Distinct().AsNoTracking().ToList();

            ///var taxDJson = JsonConvert.SerializeObject(taxonImg, Formatting.None);
            _writeFile(taxInfo, "TaxonProtectionClasses.json");
        }

        private void GenerateTaxonSynonymsFile()
        {
            //{"TaxonId":85987,"Pattern":"[[AS_Rhabdiopteryxalpina]]","Text":"Rhabdiopteryx alpina"},
            //TODO: Pattern & Text Format
            var taxInfo = _infContext.Taxon
                .Where(td => td.TaxonomyStateId == 301 && td.Synonyms != null)
                .Select(td => new TaxonSynonymsJson
                {
                    TaxonId = td.TaxonId,
                    SynArr = td.Synonyms != null ? JArray.FromObject(JsonConvert.DeserializeObject(td.Synonyms)) : null,
                })
                .AsNoTracking()
                .ToList();

            List<TaxonSynonymsJson> taxInfoNew = taxInfo
                .SelectMany(t => t.SynArr.ToObject<List<string>>()
                    .Select(tp => new TaxonSynonymsJson
                    {
                        TaxonId = t.TaxonId,
                        Text = tp.Trim(),
                        Pattern = "[[AS_" + tp.Trim() + "]]",
                        SynArr = null
                    })).ToList();

            //var taxDJson = JsonConvert.SerializeObject(taxInfoNew, Formatting.None);
            _writeFile(taxInfoNew, "TaxonSynonyms.json");
        }

        private void GenerateVersionsFile()
        {
            /*  "Taxa.json": "2020-09-03",
                  "TaxonFilterItems.json": "2020-09-03",
                  "TaxonImages.json": "2020-09-03",
                  "TaxonImageTypes.json": "2020-09-03",
                  "TaxonProtectionClasses.json": "2020-09-03",
                  "TaxonSynonyms.json": "2020-09-03",
                  "Versions.json": "2020-09-03",
            */
            var versionDate = DateTime.Today.ToString("yyyy-MM-dd");
            Dictionary<string, string> taxInfo = new Dictionary<string, string>();
            taxInfo.Add($"Taxa.json", $"{versionDate}");
            taxInfo.Add($"TaxonFilterItems.json", $"{versionDate}");
            taxInfo.Add($"TaxonImages.json", $"{versionDate}");
            taxInfo.Add($"TaxonProtectionClasses.json", $"{versionDate}");
            //taxInfo.Add($"TaxonDesc.json", $"{versionDate}");
            taxInfo.Add($"TaxonSynonyms.json", $"{versionDate}");
            taxInfo.Add($"Versions.json", $"{versionDate}");
            _writeFile(taxInfo, "Versions.json");
        }


        private void GenerateTaxonImageTypeFile()
        {
            //{"TaxonId":441092,"ImageId":25630,"Index":2,"TaxonTypeId":20001 },
            var taxInfo = _obsContext.Image
                .Where(td => td.TaxonId != null)
                .Select(td => new
                {
                    TaxonId = td.TaxonId,
                    ImageId = td.ImageId,
                    Index = td.ImagePriority,
                    TaxonTypeId = 20005,
                }).AsNoTracking().ToList();

            _writeFile(taxInfo, "TaxonImageTypes.json");
        }

        private void _writeFile(object jsonObj, string fileName)
        {
            try
            {
                if (!Directory.Exists(_filePath))
                {
                    Directory.CreateDirectory(_filePath);
                }
                string pathString = System.IO.Path.Combine(_filePath, fileName);
                if (File.Exists(pathString))
                {
                    File.Delete(pathString);
                }
                using (StreamWriter file = File.CreateText(pathString))
                {
                    JsonSerializer serializer = new JsonSerializer();
                    //serialize object directly into file stream
                    if (jsonObj != null)
                    {
                        serializer.Serialize(file, jsonObj);
                    }
                }
            }
            catch (Exception e)
            {
                Logger.Error("-- ERROR: " + e.Message);
                throw e;
            }
        }
    }

    internal class TaxonTagFilterGroup
    {
        public string GroupName { get; set; }
        public List<int> TaxonIds { get; set; }
        public List<int> DKIds { get; set; }
        public List<int> DKGIds { get; set; }
    }

    internal class TaxonDescriptionToTagFilterObject
    {
        public int TaxonId { get; set; }
        public int DescriptionKeyId { get; set; }
        public int DescriptionKeyGroupId { get; set; }
    }

    internal class TaxonInfo
    {
        public string Diagnosis;
        public int TaxonId;
        public int TaxonTypeId = 20005;
        public bool HasDiagnosis = true;
        public int? OrderId = 0;
        public string TaxonName;
        public string LocalName;
        public string FamilyName;
        public string FamilyLocalName;
        public string OrderName;
        public string OrderLocalName;
        public string TaxonAuthor;
        public string TaxonomyStateName;
        public int IdentificationLevelFemale = 1;
        public int IdentificationLevelMale = 1;
        public object SliderImages { get; internal set; }
    }

    internal class TaxonImage
    {
        public int? TaxonId;
        //ImageId -> ImageName @bodentier4
        public string ImageId;
        //only temporary for conversion
        public JArray SliderImgArray;
        public int? Index;
        public string Title;
        public string Description;
        public bool Male = true;
        public bool Female = true;
        public bool Downside = false;
        public bool SecondForm = false;
    }

    internal class TaxonFilter
    {
        public int TagId;
        public string TagParentName;
        public int TagParentId;
        public int TaxonId;
        public int TaxonTypeId = 20005;
        public int? OrderPriority;
        public List<string> ListSourceJson;
        public string KeyDataType;
        public string TagValue;
        public int? MinValue;
        public int? MaxValue;
        public List<int> TaxonHits;
        public List<int> OrderIds;
        public int? VisibilityCategoryId;
    }

    internal class TaxonSynonymsJson
    {
        public int TaxonId;
        public string Pattern;
        public JArray SynArr;
        public string Text;
    }
}