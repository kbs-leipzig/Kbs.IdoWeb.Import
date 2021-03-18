using Kbs.IdoWeb.Data.Determination;
using Kbs.IdoWeb.Data.Information;
using Kbs.IdoWeb.Data.Observation;
using Microsoft.EntityFrameworkCore;
using Microsoft.Extensions.Configuration;
using Newtonsoft.Json;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Logical;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text.RegularExpressions;

namespace Kbs.IdoWeb.Import.Models
{
    class MetaInfoModel
    {
        private DeterminationContext _detContext = new DeterminationContext();
        private InformationContext _infContext = new InformationContext();
        private ObservationContext _obsContext = new ObservationContext();
        private static readonly NLog.Logger Logger = NLog.LogManager.GetCurrentClassLogger();
        private static string edaphoApi = @"https://api.edaphobase.org/taxon/";

        public MetaInfoModel()
        {
            Init();
        }

        private List<string>? doEdaphoRequest (string URL, string urlParameter)
        {
            HttpClient client = new HttpClient();
            client.BaseAddress = new Uri(URL+urlParameter);

            // Add an Accept header for JSON format.
            client.DefaultRequestHeaders.Accept.Add(
            new MediaTypeWithQualityHeaderValue("application/json"));

            // List data response.
            HttpResponseMessage response = client.GetAsync("").Result;  // Blocking call! Program will wait here until a response is received or a timeout occurs.
            if (response.IsSuccessStatusCode)
            {
                // Parse the response body.
                var jsonString = response.Content.ReadAsStringAsync().Result;  //Make sure to add a reference to System.Net.Http.Formatting.dll
                if(jsonString != null)
                {
                    //TODO: fix JsonConvert Error!
                    var dataObject = JsonConvert.DeserializeObject<EdaphobaseEntry>(jsonString);
                    if(dataObject.synonyms != null)
                    {
                        List<string> synonyms = dataObject.synonyms.ToList().Select(syns => syns.displayName).ToList();
                        return synonyms;
                    }
                }
            }
            else
            {
                Console.WriteLine("{0} ({1})", (int) response.StatusCode, response.ReasonPhrase);
            }

            //Make any other calls using HttpClient here.
            //Dispose once all HttpClient calls are complete. This is not necessary if the containing object will be disposed of; for example in this case the HttpClient instance will be disposed automatically when the application terminates so the following call is superfluous.
            client.Dispose();
            return null;
        }

        private void Init()
        {
            try
            {
                Logger.Info("Initializing Import Meta ..");
                var optionsBuilderDet = new DbContextOptionsBuilder<DeterminationContext>();
                optionsBuilderDet.UseNpgsql(Program.Configuration.GetConnectionString("DatabaseConnection"));
                _detContext = new DeterminationContext(optionsBuilderDet.Options);

                var optionsBuilderInf = new DbContextOptionsBuilder<InformationContext>();
                optionsBuilderInf.UseNpgsql(Program.Configuration.GetConnectionString("DatabaseConnection"));
                _infContext = new InformationContext(optionsBuilderInf.Options);
                Logger.Info(".. Success");

                var optionsBuilderObs = new DbContextOptionsBuilder<ObservationContext>();
                optionsBuilderObs.UseNpgsql(Program.Configuration.GetConnectionString("DatabaseConnection"));
                _obsContext = new ObservationContext(optionsBuilderObs.Options);
                Logger.Info(".. Success");
            }
            catch (Exception e)
            {
                Logger.Error(e, "Could not init Import Meta");
            }
        }

        public void UpdateDkgKeyTypeInfo()
        {
            try
            {
                var dks = _detContext.DescriptionKey.Include(dk => dk.TaxonDescription).Include(dk => dk.DescriptionKeyGroup).AsNoTracking().ToList();
                dks.ForEach(dk =>
                {
                    var dkTypeList = dk.TaxonDescription.ToList().Where(td => td.DescriptionKeyTypeId.HasValue).Select(td => td.DescriptionKeyTypeId.Value).Distinct().ToList();
                    var dkg = dk.DescriptionKeyGroup;
                    dkg.DescriptionKeyGroupType = JsonConvert.SerializeObject(dkTypeList);
                    _detContext.DescriptionKeyGroup.Update(dkg);
                });
                _detContext.SaveChanges();
            }
            catch (Exception e)
            {
                Logger.Error(e, "Error updating DescriptionKeyGroupDataType");
            }
        }

        public void UpdateEdaphobaseInfo ()
        {
            try
            {
                var taxons = _infContext.Taxon.Where(tax => tax.EdaphobaseId != null).ToList();
                foreach (Taxon tax in taxons)
                {
                    if (tax.EdaphobaseId != null)
                    {
                        List<string> synonyms = doEdaphoRequest(edaphoApi, tax.EdaphobaseId.ToString());
                        if(synonyms != null)
                        {
                            synonyms.RemoveAll(syn => syn == tax.Genus.TaxonName);
                            //sanitize edaphobase input double whitespaces
                            synonyms.ForEach(syn => Regex.Replace(syn, @"\s+", " "));
                            tax.Synonyms = JsonConvert.SerializeObject(synonyms);
                        }
                    }
                }
                _infContext.UpdateRange(taxons);
                _infContext.SaveChanges();
                Logger.Info("-- Done Updating Edaphobase Info.");
            } catch (Exception e)
            {
                Logger.Error(e, "-- Error Updating Edaphobase Info");
            }
        }

        public void UpdateTaxDescChildrenInfo()
        {
            var taxons = _infContext.Taxon
                .Include(tax => tax.TaxonomyState);
            taxons.ToList().ForEach(taxItem =>
                taxItem.HasTaxDescChildren = _taxonCheckChildren(taxItem.TaxonId, taxItem.TaxonomyState.StateDescription)
            );
            _infContext.UpdateRange(taxons);
            _infContext.SaveChanges();
        }

        private bool _taxonCheckChildren(int taxonId, string? stateDesc)
        {
            if (stateDesc != null)
            {
                var taxDesc = _detContext.TaxonDescription.Select(taxD => taxD.TaxonId).Distinct().ToList();
                List<int> childrenTaxIds;
                //disabled certain states; check import files (bodentiere) for enabling
                switch (stateDesc)
                {
                    /** disabled, too detailed **/
                    case "family":
                        childrenTaxIds = _infContext.Taxon
                                .Where(tax => tax.FamilyId == taxonId && (tax.SubfamilyId == null || tax.GenusId == null))
                                .Select(tx => tx.TaxonId).ToList();
                        break;
                    /** disabled, too detailed **/
                    /**
                    case "subfamily":
                        childrenTaxIds = _infContext.Taxon
                                .Where(tax => tax.SubfamilyId == taxonId && tax.GenusId == null)
                                .Select(tx => tx.TaxonId).ToList();
                        break;
                    **/
                    case "order":
                        childrenTaxIds = _infContext.Taxon
                                .Where(tax => tax.OrderId == taxonId && (tax.SuborderId == null || tax.FamilyId == null))
                                .Select(tx => tx.TaxonId).ToList();
                        break;
                    case "class":
                        childrenTaxIds = _infContext.Taxon
                                .Where(tax => tax.ClassId == taxonId && (tax.Subclass == null || tax.OrderId == null))
                                .Select(tx => tx.TaxonId).ToList();
                        break;
                    /** disabled, too detailed **/
                    /**
                    case "suborder":
                        childrenTaxIds = _infContext.Taxon
                                .Where(tax => tax.SuborderId == taxonId && tax.FamilyId == null)
                                .Select(tx => tx.TaxonId).ToList();
                        break;
                    **/
                    /** disabled, too detailed **/
                    /**
                    case "subphylum":
                        childrenTaxIds = _infContext.Taxon
                                .Where(tax => tax.SubphylumId == taxonId && tax.ClassId == null)
                                .Select(tx => tx.TaxonId).ToList();
                        break;
                    **/
                    case "phylum":
                        childrenTaxIds = _infContext.Taxon
                                .Where(tax => tax.PhylumId == taxonId && (tax.SubphylumId == null || tax.ClassId == null))
                                .Select(tx => tx.TaxonId).ToList();
                        break;
                    case "kingdom":
                        childrenTaxIds = _infContext.Taxon
                                .Where(tax => tax.KingdomId == taxonId && tax.PhylumId == null)
                                .Select(tx => tx.TaxonId).ToList();
                        break;
                    /** disabled, too detailed **/
                    /**
                    case "genus":
                        childrenTaxIds = _infContext.Taxon
                                .Where(tax => tax.GenusId == taxonId && tax.SpeciesId == null)
                                .Select(tx => tx.TaxonId).ToList();
                        break;
                    **/
                    case "sub class":
                        childrenTaxIds = _infContext.Taxon
                                .Where(tax => tax.SubclassId == taxonId && tax.OrderId == null)
                                .Select(tx => tx.TaxonId).ToList();
                        break;
                    case "species":
                        Logger.Debug("Skipping because Species for TaxonId " + taxonId);
                        childrenTaxIds = null;
                        break;
                    default:
                        Logger.Debug("Could not determine State for TaxonId " + taxonId);
                        childrenTaxIds = null;
                        break;
                }

                if (childrenTaxIds != null)
                {
                    var result = taxDesc.Where(taxD => childrenTaxIds.Contains(taxD)).ToList();
                    if (result.Count > 0)
                    {
                        return true;
                    }
                }
            }

            return false;
        }

        public void updateTaxIdsInImageAndObs ()
        {
            Dictionary<string, int> taxIdName_dict = _infContext.Taxon.Where(o => o.TaxonId != null && o.TaxonName != null && o.TaxonomyStateId == 301).Select(t => new { t.TaxonId, t.TaxonName }).ToDictionary(x => x.TaxonName, y => y.TaxonId);
            List<Observation> obs = _obsContext.Observation.Where(o => o.TaxonName != null).ToList();
            List<Image> img = _obsContext.Image.Where(o => o.TaxonName != null).ToList();
            foreach(Observation o in obs)
            {
                if(taxIdName_dict.TryGetValue(o.TaxonName, out int val))
                {
                    o.TaxonId = val;
                }
            }
            foreach (Image i in img)
            {
                if (taxIdName_dict.TryGetValue(i.TaxonName, out int val))
                {
                    i.TaxonId = val;
                }
            }
            _obsContext.SaveChanges();
        }

        private List<int> _getTaxonIdsByHierarchyLevel(Taxon taxonGroupInst, string hierLevel)
        {
            List<int> taxon_groupFiltered = null;
            switch (hierLevel)
            {
                case "KingdomId":
                    taxon_groupFiltered = _infContext.Taxon
                        .Where(tax => tax.KingdomId == taxonGroupInst.TaxonId && tax.PhylumId == null)
                        .Select(tx => tx.TaxonId)
                        .ToList();
                    break;
                case "PhylumId":
                    taxon_groupFiltered = _infContext.Taxon
                        .Where(tax => tax.PhylumId == taxonGroupInst.TaxonId && tax.ClassId == null)
                        .Select(tx => tx.TaxonId)
                        .ToList();
                    break;
                case "SubphylumId":
                    taxon_groupFiltered = _infContext.Taxon
                        .Where(tax => tax.PhylumId == taxonGroupInst.TaxonId && tax.ClassId == null)
                        .Select(tx => tx.TaxonId)
                        .ToList();
                    break;
                case "ClassId":
                    taxon_groupFiltered = _infContext.Taxon
                        .Where(tax => tax.ClassId == taxonGroupInst.TaxonId && tax.OrderId == null)
                        .Select(tx => tx.TaxonId)
                        .ToList();
                    break;
                case "SubclassId":
                    taxon_groupFiltered = _infContext.Taxon
                        .Where(tax => tax.ClassId == taxonGroupInst.TaxonId && tax.OrderId == null)
                        .Select(tx => tx.TaxonId)
                        .ToList();
                    break;
                //currently order is the second to last level with taxonDescriptions --> only species underneath
                case "OrderId":
                    taxon_groupFiltered = _infContext.Taxon
                        .Where(tax => tax.OrderId == taxonGroupInst.TaxonId && tax.SpeciesId == null)
                        .Select(tx => tx.TaxonId)
                        .ToList();
                    break;
                case "SuborderId":
                    taxon_groupFiltered = _infContext.Taxon
                        .Where(tax => tax.OrderId == taxonGroupInst.TaxonId && tax.SpeciesId == null)
                        .Select(tx => tx.TaxonId)
                        .ToList();
                    break;
                case "FamilyId":
                    //enable Subfamily when available in excel import
                    taxon_groupFiltered = _infContext.Taxon
                        .Where(tax => tax.FamilyId == taxonGroupInst.TaxonId && tax.SpeciesId == null)
                        .Select(tx => tx.TaxonId)
                        .ToList();
                    break;
                case "SubfamilyId":
                    //enable Subfamily when available in excel import
                    taxon_groupFiltered = _infContext.Taxon
                        .Where(tax => tax.FamilyId == taxonGroupInst.TaxonId && tax.SpeciesId == null)
                        .Select(tx => tx.TaxonId)
                        .ToList();
                    break;
                default:
                    taxon_groupFiltered = _infContext.Taxon
                        .Select(tx => tx.TaxonId)
                        .ToList();
                    break;
            }
            return taxon_groupFiltered;
        }
    }

    internal class EdaphobaseEntry
    {
        public int? id;
        public string? name = null;
        public int? rank;
        public int? year;
        public bool? valid;
        public string? author;
        public string? displayName;
        public int? valid_taxon_id;
        //disabled for shitty api response format
        //public Dictionary<int, EdaphobaseEntry> children;
        public List<EdaphobaseEntry> synonyms;
    }
}
