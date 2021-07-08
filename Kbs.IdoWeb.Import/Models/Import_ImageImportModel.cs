using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using Kbs.IdoWeb.Data.Determination;
using Kbs.IdoWeb.Data.Information;
using Kbs.IdoWeb.Data.Observation;
using Microsoft.EntityFrameworkCore;
using Microsoft.Extensions.Configuration;
using NLog;
using OfficeOpenXml;

namespace Kbs.IdoWeb.Import.Models
{
    class ImageImportModel
    {
        private ObservationContext _contextObs;
        private InformationContext _contextInf;
        private static readonly NLog.Logger Logger = NLog.LogManager.GetCurrentClassLogger();
        private int abbNameCol;
        private int descriptionCol;
        private int taxonCol;
        private int sourceCol;
        private int rightsCol;
        private int ccrCol;
        private int ccrLinkCol;
        public ExcelPackage excelPackage;
        private const string _taxonRegex = @"^(?:\S*\s\S*\s)(?:\(?)(\S*|\S*\s\S*|\S*|\S*\s\&\s\S*|\S{2}\s\S{2}\s\S*|\S{2}\s\S*)(?:(?:\,\s\d{4}?\)?)|(?:\,\s\)?))$";
        private const string _ccrRegex = @"^(((CC(\s|\-))*?[A-Z]{2}((\-)|(\s)))(\d{1}.\d{1}|(([A-Z]{2})(\s|\-)(\d{1}.\d{1}))|([A-Z]{2})(\s|\-)(([A-Z]{2})|(\d{1}.\d{1}))\s(\d{1}.\d{1})?)|(CC0\s\d{1}.\d{1}))$";
        private const string _ccUrlRegex = @"^(?:http(s)?:\/\/)?[creativecommons.-]+(?:\.[\w\.-]+)+[\w\-\._~:/?#[\]@!\$&'\(\)\*\+,;=.]+$";


        public ExcelWorksheet worksheetImages { get; internal set; }
        public int imageImportCounter { get; internal set; }
        public int imageUpdateCounter { get; internal set; }
        public int licenseImportCounter { get; private set; }

        public ImageImportModel(string importFilePath)
        {
            FileInfo fileInfo = new FileInfo(importFilePath);
            excelPackage = new ExcelPackage(fileInfo);
            worksheetImages = excelPackage.Workbook.Worksheets[2];
            if (worksheetImages != null)
            {
                InitContexts();
                InitProperties();
            }
        }

        public void InitProperties()
        {
            imageImportCounter = 0;
            imageUpdateCounter = 0;
            for (int i = 1; i <= worksheetImages.Dimension.Columns; i++)
            {
                if (worksheetImages.Cells[1, i].Value != null)
                {
                    var cellVal = worksheetImages.Cells[1, i].Value?.ToString().Trim();
                    switch (cellVal)
                    {
                        case "Abb. Name":
                            abbNameCol = i;
                            break;
                        case "Erklärung":
                            descriptionCol = i;
                            break;
                        case "Art":
                            taxonCol = i;
                            break;
                        case "Quelle":
                            sourceCol = i;
                            break;
                        case "Rechte":
                            rightsCol = i;
                            break;
                        case "CCR":
                            ccrCol = i;
                            break;
                        case "CCR-Link":
                            ccrLinkCol = i;
                            break;
                        default:
                            break;
                    }
                }
            }
        }

        private void InitContexts()
        {
            var optionsBuilderInf = new DbContextOptionsBuilder<InformationContext>();
            optionsBuilderInf.UseNpgsql(Program.Configuration.GetConnectionString("DatabaseConnection"));
            _contextInf = new InformationContext(optionsBuilderInf.Options);

            var optionsBuilderObs = new DbContextOptionsBuilder<ObservationContext>();
            optionsBuilderObs.UseNpgsql(Program.Configuration.GetConnectionString("DatabaseConnection"));
            _contextObs = new ObservationContext(optionsBuilderObs.Options);
        }

        internal void StartImageImport()
        {
            if (worksheetImages != null)
            {
                try
                {
                    for (int i = 2; i <= worksheetImages.Dimension.Rows; i++)
                    {
                        if (worksheetImages.Cells[i, 1].Value != null)
                        {
                            Image imageInst = _parseImageRow(i);
                            if (imageInst != null)
                            {
                                _saveImage(imageInst);
                                _contextObs.SaveChanges();
                            }
                        }
                    }
                }
                catch (Exception e)
                {
                    Logger.Error(e, "Error @ Image Import");
                }
            }
        }


        private Image _parseImageRow(int i)
        {
            Image img = new Image();
            if (worksheetImages.Cells[i, taxonCol].Value?.ToString().Trim() != "")
            {
                var taxId = _taxonLookup(worksheetImages.Cells[i, taxonCol].Value?.ToString().Trim());
                if (taxId != null)
                {
                    img.TaxonId = taxId.Value;
                }
                else
                {
                    Logger.Debug("-- Not getting TaxonId for " + worksheetImages.Cells[i, taxonCol].Value?.ToString().Trim());
                    img.TaxonId = null;
                }
                img.Author = worksheetImages.Cells[i, rightsCol].Value?.ToString().Trim();
                img.CopyrightText = worksheetImages.Cells[i, sourceCol].Value?.ToString().Trim();
                img.LicenseId = _ccrLookup(worksheetImages.Cells[i, ccrCol].Value?.ToString().Trim(), i);
                img.Description = worksheetImages.Cells[i, descriptionCol].Value?.ToString().Trim();
                img.ImagePath = worksheetImages.Cells[i, abbNameCol].Value?.ToString().Trim();
                img.TaxonName = worksheetImages.Cells[i, taxonCol].Value?.ToString().Trim();
                img.ImagePriority = i;
                img.IsApproved = true;
                return img;
            }
            return null;
        }

        private int? _taxonLookup(string taxonName)
        {
            if (taxonName != null)
            {

                if (Regex.IsMatch(taxonName, _taxonRegex))
                {
                    string genusName = taxonName.Split(" ")[0].Trim();
                    string speciesName = taxonName.Split(" ")[1].Trim();
                    taxonName = $"{genusName} {speciesName}";
                }

                var taxonId = _contextInf.Taxon.Where(tax => EF.Functions.Like(tax.TaxonName, $"%{taxonName}%")).Select(img => img.TaxonId).FirstOrDefault();
                if (taxonId != 0)
                {
                    return taxonId;
                }
            }
            return null;
        }

        private int? _ccrLookup(string inString, int rowIdx)
        {
            int licenseId = 7;
            if (inString != null)
            {
                var inString_clean = inString.TrimStart("CC ".ToCharArray()).Replace("-", " ");
                licenseId = _contextObs.ImageLicense.Where(img => EF.Functions.Like(img.LicenseName.Replace("-", " "), $"%{inString_clean}%")).Select(img => img.LicenseId).FirstOrDefault();
                if (licenseId == 0)
                {
                    //not in db - valid format?
                    if (Regex.IsMatch(inString, _ccrRegex))
                    {
                        //TODO LocalisationJson??
                        var ccrLink = worksheetImages.Cells[rowIdx, ccrLinkCol].Value?.ToString().Trim();
                        if (Regex.IsMatch(inString, _ccUrlRegex))
                        {
                            ImageLicense imgLic = new ImageLicense();
                            imgLic.LicenseName = inString;
                            imgLic.LicenseLink = ccrLink;
                            licenseId = _saveImageLicense(imgLic);
                            if (licenseId == 0) return null;
                        }
                        else
                        {
                            //regular copyright as fallback
                            licenseId = 7;
                        }
                    }
                }
            }
            return licenseId;
        }

        private void _saveImage(Image img)
        {
            try
            {
                if (!_imageAlreadyExists(img))
                {
                    Logger.Debug("Saving new " + img.ImagePath + " to Context");
                    _contextObs.Add(img);
                    Logger.Debug(".. Saved to Context");
                    imageImportCounter++;
                }
                else
                {
                    Image exImg = _getExistingImage(img);
                    if (exImg != null)
                    {
                        _updateImageProperties(img, ref exImg);
                        _contextObs.Update(exImg);
                        imageUpdateCounter++;
                        Logger.Debug($"{img.ImagePath} updated");
                    }
                }
            }
            catch (Exception e)
            {
                Logger.Error(e.InnerException, "Error Adding Image to Context");
            }
        }

        private static void _updateImageProperties(Image img, ref Image exImg)
        {
            exImg.IsApproved = img.IsApproved;
            exImg.LicenseId = img.LicenseId;
            exImg.TaxonId = img.TaxonId;
            exImg.UserId = img.UserId;
            exImg.Author = img.Author;
            exImg.CopyrightText = img.CopyrightText;
            exImg.Description = img.Description;
            exImg.ImagePriority = img.ImagePriority;
            if (exImg.TaxonName == null)
            {
                exImg.TaxonName = img.TaxonName;
            }
            exImg.CmsId = null;
        }

        private Image _getExistingImage(Image img)
        {
            var result = _contextObs.Image.Where(image => image.ImagePath == img.ImagePath).FirstOrDefault();
            if (result != null)
            {
                return result;
            }
            return null;
        }

        public int _saveImageLicense(ImageLicense imgLic)
        {
            try
            {
                Logger.Debug("Saving new" + imgLic.LicenseName + " to Context");
                _contextObs.Add(imgLic);
                _contextObs.SaveChanges();
                Logger.Debug(".. Saved to Context");
                licenseImportCounter++;
                return imgLic.LicenseId;
            }
            catch (Exception e)
            {
                Logger.Error(e.InnerException, "Error Adding ImageLicense to Context");
                return 0;
            }
        }


        private bool _imageAlreadyExists(Image img)
        {
            var result = _contextObs.Image.FirstOrDefault(image => image.ImagePath == img.ImagePath);
            if (result != null)
            {
                return true;
            }
            return false;
        }

    }
}
