using CsvHelper;
using ExcelDataReader;
using System;
using System.Collections.Generic;
using System.Data;
using System.Globalization;
using System.IO;
using System.Linq;

namespace CalculateREFscores
{
    class Program
    {
        static void Main(string[] args)
        {
            // Needed to make ExcelReader work
            System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);

            // This methodology is largey a port from R to C# .NET of the excellent code at https://github.com/oscci/REFvsWellcome
            // Additionally I assign Univerisities to UK NUTS1 regions and create regional sum scores

            // The REF results spreadsheet is available at https://results.ref.ac.uk/(S(5azlsbi3iciltmzxzr4wfez5))/DownloadResults
            List<UniversityRegion> UniversityRegions = new List<UniversityRegion>();
            using (StreamReader textReader = new StreamReader("Assets/UniversityRegionsLink.csv"))
            {
                using (CsvReader csvReader = new CsvReader(textReader))
                {
                    UniversityRegions = csvReader.GetRecords<UniversityRegion>().ToList();
                }
            }

            // Load panel weightings
            Dictionary<string, double> PanelWeightings = new Dictionary<string, double>();
            using (StreamReader textReader = new StreamReader("Assets/PanelWeightings.csv"))
            {
                using (CsvReader csvReader = new CsvReader(textReader))
                {
                    PanelWeightings = csvReader.GetRecords<PanelWeighting>().ToDictionary(x => x.Panel, x => x.Weighting);
                }
            }

            // Load the Excel sheet's data
            string ExcelFilePath = @"Assets/REF2014 Results.xlsx";
            string SheetName = @"REF2014 Profiles";

            DataTable REFEXCELTABLE;
            using (var stream = File.Open(ExcelFilePath, FileMode.Open, FileAccess.Read))
            {
                using (var reader = ExcelReaderFactory.CreateReader(stream))
                {
                    var result = reader.AsDataSet();
                    DataTableCollection Worksheets = result.Tables;
                    REFEXCELTABLE = Worksheets[Worksheets.IndexOf(SheetName)];
                }
            }

            // 
            List<REFRow> REFRows = new List<REFRow>();
            for (int i = 8; i < REFEXCELTABLE.Rows.Count; i++)
            {
                REFRow refRow = new REFRow();
                refRow.InstitutionCode_UKPRN = (string)REFEXCELTABLE.Rows[i].ItemArray[0];
                refRow.InstitutionName = (string)REFEXCELTABLE.Rows[i].ItemArray[1];
                refRow.MainPanel = (string)REFEXCELTABLE.Rows[i].ItemArray[3];
                refRow.UnitOfAssessmentNumber = (double)REFEXCELTABLE.Rows[i].ItemArray[4];
                refRow.UnitOfAssessmentName = (string)REFEXCELTABLE.Rows[i].ItemArray[5];
                refRow.MultipleSubmissionLetter = (string)REFEXCELTABLE.Rows[i].ItemArray[6];
                refRow.MultipleSubmissionName = (string)REFEXCELTABLE.Rows[i].ItemArray[7];

                refRow.Profile = (string)REFEXCELTABLE.Rows[i].ItemArray[9];
                refRow.FTECategoryAStaffSubmitted = (double)REFEXCELTABLE.Rows[i].ItemArray[10];
                if (REFEXCELTABLE.Rows[i].ItemArray[11].GetType() != typeof(string))
                {
                    refRow.percent4star = (double)REFEXCELTABLE.Rows[i].ItemArray[11];
                }
                if (REFEXCELTABLE.Rows[i].ItemArray[12].GetType() != typeof(string))
                {
                    refRow.percent3star = (double)REFEXCELTABLE.Rows[i].ItemArray[12];
                }
                if (REFEXCELTABLE.Rows[i].ItemArray[13].GetType() != typeof(string))
                {
                    refRow.percent2star = (double)REFEXCELTABLE.Rows[i].ItemArray[13];
                }
                if (REFEXCELTABLE.Rows[i].ItemArray[14].GetType() != typeof(string))
                {
                    refRow.percent1star = (double)REFEXCELTABLE.Rows[i].ItemArray[14];
                }
                if (REFEXCELTABLE.Rows[i].ItemArray[15].GetType() != typeof(string))
                {
                    refRow.unclassified = (double)REFEXCELTABLE.Rows[i].ItemArray[15];
                }
                REFRows.Add(refRow);
            }

            // Now calculate the strength measure
            // This is weighted as 4x4*+1x3*+0x2*+0x1*+0xNC
            // How QR is assigned is here -- https://re.ukri.org/research/how-we-fund-research/
            // Document showing this is at https://www.hefcw.ac.uk/documents/policy_areas/research/Explanation%20of%20QR%20Funding%20Method.pdf

            // Weightings by profile are "65 per cent for outputs, 20 per cent for impact and 15 per cent for environment"
            // But we care about excellence so we're just going to look at outputs

            List<StrengthByRegion> StrengthByRegions = new List<StrengthByRegion>();
            foreach(REFRow refrow in REFRows.Where(x => x.Profile == "Outputs"))
            {
                if (null != UniversityRegions.Where(x => x.University == refrow.InstitutionName).FirstOrDefault())
                {
                    // I'm not taking London weighting into accounts, we're supposed to be funding excellence not regional distribution. But weights could be added here
                    string UniversityRegion = UniversityRegions.Where(x => x.University == refrow.InstitutionName).FirstOrDefault().Region;

                    double ScoreAltA = ((refrow.percent4star * 3 + refrow.percent3star * 2) * refrow.FTECategoryAStaffSubmitted.Value) * PanelWeightings[refrow.MainPanel];
                    double ExistingScore = ((refrow.percent4star * 4 + refrow.percent3star * 1) * refrow.FTECategoryAStaffSubmitted.Value) * PanelWeightings[refrow.MainPanel];
                    double ScoreAltB = ((refrow.percent4star * 5 + refrow.percent3star * 0) * refrow.FTECategoryAStaffSubmitted.Value) * PanelWeightings[refrow.MainPanel];

                    if (StrengthByRegions.Where(x => x.Region == UniversityRegion).FirstOrDefault() == null)
                    {
                        StrengthByRegion strengthByRegion = new StrengthByRegion()
                        {
                            Region = UniversityRegion
                        };
                        StrengthByRegions.Add(strengthByRegion);
                    }
                    StrengthByRegions.Where(x => x.Region == UniversityRegion).FirstOrDefault().ExistingStrength += ExistingScore;
                    StrengthByRegions.Where(x => x.Region == UniversityRegion).FirstOrDefault().StrenghtAltA += ScoreAltA;
                    StrengthByRegions.Where(x => x.Region == UniversityRegion).FirstOrDefault().StrenghtAltB += ScoreAltB;
                    
                    if (refrow.MainPanel == "A")
                    {
                        StrengthByRegions.Where(x => x.Region == UniversityRegion).FirstOrDefault().PanelAOnly += ExistingScore;
                    }

                }
                else
                {
                    Console.WriteLine($"No region was found for {refrow.InstitutionName} in UniversityRegionsLink.csv");
                }
            }          
            
            using (TextWriter TextWriter = File.CreateText(@"REFOutputStrengthByRegion.csv"))
            {
                CsvWriter CSVwriter = new CsvWriter(TextWriter);
                CSVwriter.WriteRecords(StrengthByRegions);
            }
        }
    }

    public class StrengthByRegion
    {
        public string Region { get; set; }
        public double ExistingStrength { get; set; }
        public double StrenghtAltA { get; set; }
        public double StrenghtAltB { get; set; }
        public double PanelAOnly { get; set; }
    }

    public class PanelWeighting
    {
        public string Panel { get; set; }
        public double Weighting { get; set; }
    }

    public class UniversityRegion
    {
        public string University { get; set; }
        public string Region { get; set;}
        public string SubRegion { get; set; }
    }

    public class REFRow
    {
        public string InstitutionCode_UKPRN { get; set; }
        public string InstitutionName { get; set; }
        public string MainPanel { get; set; }
        public double UnitOfAssessmentNumber { get; set; }
        public string UnitOfAssessmentName { get; set; }
        public string MultipleSubmissionLetter { get; set; }
        public string MultipleSubmissionName { get; set; }
        public double? FTECategoryAStaffSubmitted { get; set; }
        public string Profile { get; set; }
        public double percent4star { get; set; }
        public double percent3star { get; set; }
        public double percent2star { get; set; }
        public double percent1star { get; set; }
        public double unclassified { get; set; }
    }
}
