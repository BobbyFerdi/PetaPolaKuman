using ClosedXML.Excel;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Logging;
using PetaPolaKuman.Models;
using PetaPolaKuman.Utilities;
using Travewell.Library.Extensions;
using Travewell.Library.Utilities;

var builder = new ServiceCollection();
builder.AddLibrary();
var app = builder.BuildServiceProvider();
var logger = app.GetService<ILogger<Program>>();
var fileName = "..\\..\\..\\source.xlsx";
var targetFileName = $"Peta Pola Kuman RSMA 2";
var targetFilePath = $"..\\..\\..\\{targetFileName}.xlsx";
var year = 2022;
var antibioticsLineNumber = 7;
var antibioticsStartColumn = "M".ExcelColumnNameToNumber();
var antibioticsEndColumn = "BA".ExcelColumnNameToNumber();
var sourceWorkbook = new XLWorkbook(fileName);
var antibioticsRow = sourceWorkbook.Worksheet(1).RangeUsed().RowsUsed().Skip(antibioticsLineNumber - 1).FirstOrDefault();
var baseData = new BaseData
{
    Antibiotics = new List<string>(),
    Records = new List<Record>(),
    Organisms = new List<string>(),
    Locations = new List<string>(),
    Specimens = new List<string>()
};
var resistanceRates = new ResistanceRates();

for (var a = antibioticsStartColumn; a <= antibioticsEndColumn; a++)
{
    baseData.Antibiotics.Add(antibioticsRow.GetCellValue(a));
}

var recordsStartLineNumber = 11;
var recordsRows = sourceWorkbook.Worksheet(1).RangeUsed().RowsUsed().Skip(recordsStartLineNumber - 1);
var locationColumn = "H";
var specimenColumn = "I";
var organismColumn = "J";

foreach (var recordsRow in recordsRows)
{
    var organism = recordsRow.GetCellValue(organismColumn).Trim();

    if (organism.ToLower() == "negatif") continue;
    if (organism.ToLower() == "staphylococcus hemolyticus") organism = "Staphylococcus haemolyticus";

    var specimen = recordsRow.GetCellValue(specimenColumn).Trim();

    if (specimen.Contains('(')) specimen = specimen.Split('(', ')')[1].ToTitleCase();

    var location = recordsRow.GetCellValue(locationColumn).Trim().Replace("Rawat", "R.");
    var record = new Record
    {
        Location = location,
        Specimen = specimen,
        Organism = organism,
        ResistanceRates = new List<ResistanceRate>()
    };

    if (!baseData.Organisms.Contains(organism)) baseData.Organisms.Add(organism);
    if (!baseData.Specimens.Contains(specimen)) baseData.Specimens.Add(specimen);
    if (!baseData.Locations.Contains(location)) baseData.Locations.Add(location);

    var antibioticsCounter = 0;

    for (var a = antibioticsStartColumn; a <= antibioticsEndColumn; a++)
    {
        if (!recordsRow.GetCellValue(a).IsNullOrEmpty())
        {
            var resistance = resistanceRates.Rates.FirstOrDefault(r => r.Code == recordsRow.GetCellValue(a));

            if (resistance != null) record.ResistanceRates.Add(new ResistanceRate(resistance.Code, baseData.Antibiotics[antibioticsCounter], resistance.Value));
        }

        antibioticsCounter++;
    }

    baseData.Records.Add(record);
}

baseData.Organisms = baseData.Organisms.DistinctAndOrder();
baseData.Specimens = baseData.Specimens.DistinctAndOrder();
baseData.Locations = baseData.Locations.DistinctAndOrder();

var baseDataJson = StaticJsonHelper.Serialize(baseData);
logger.LogInformation(StaticJsonHelper.Serialize(baseData));

using var resultWorkbook = new XLWorkbook();

foreach (var specimen in baseData.Specimens)
{
    #region Specimen

    var specimenRowCounter = 1;
    var specimenCellCounter = 1;
    var specimenSheet = resultWorkbook.Worksheets.Add(specimen.Replace("/", ""));
    specimenSheet.SetCellValue(specimenRowCounter, specimenCellCounter, $"PETA POLA KUMAN {year}");
    specimenRowCounter++;
    specimenSheet.SetCellValue(specimenRowCounter, specimenCellCounter, "Organism");
    specimenCellCounter++;

    foreach (var organism in baseData.Organisms)
    {
        specimenSheet
            .SetCellValue(specimenRowCounter, specimenCellCounter, organism)
            .Style
            .Alignment
            .SetTextRotation(90)
            .Alignment
            .SetHorizontal(XLAlignmentHorizontalValues.Center);
        specimenCellCounter++;
    }

    specimenRowCounter++;
    specimenCellCounter = 1;
    specimenSheet.SetCellValue(specimenRowCounter, specimenCellCounter, "Number of isolates");
    specimenCellCounter++;

    var specimenData = baseData.Records.Where(r => r.Specimen.ToLower() == specimen.ToLower());

    foreach (var organism in baseData.Organisms)
    {
        var organismCounter = specimenData.Where(s => s.Organism.ToLower() == organism.ToLower()).Sum(s => s.ResistanceRates.Count);

        if (organismCounter > 0)
            specimenSheet
                .SetCellValue(specimenRowCounter, specimenCellCounter, organismCounter)
                .Style
                .Alignment
                .SetHorizontal(XLAlignmentHorizontalValues.Center);

        specimenCellCounter++;
    }

    specimenRowCounter++;
    var staticRowCounter = specimenRowCounter;
    specimenCellCounter = 1;

    foreach (var antibiotic in baseData.Antibiotics)
    {
        specimenSheet.SetCellValue(specimenRowCounter, specimenCellCounter, antibiotic);
        specimenRowCounter++;
    }

    specimenCellCounter++;

    foreach (var organism in baseData.Organisms)
    {
        specimenRowCounter = staticRowCounter;
        var recordResistanceRates = specimenData.Where(a => a.Organism.ToLower() == organism.ToLower()).SelectMany(o => o.ResistanceRates);

        if (recordResistanceRates.CheckNotNullAndAny())
        {
            foreach (var antibiotics in baseData.Antibiotics)
            {
                var antibioticsData = recordResistanceRates.Where(a => a.Antibiotic.ToLower() == antibiotics.ToLower());

                if (antibioticsData != null && antibioticsData.Any())
                {
                    var sum = antibioticsData.Where(a => a.Antibiotic.ToLower() == antibiotics.ToLower()).Sum(s => s.Value);
                    var average = antibioticsData.Any() ? ((double)(antibioticsData.Sum(s => s.Value) / antibioticsData.Count())).ToInt() : 0;
                    specimenSheet
                        .SetCellValue(specimenRowCounter, specimenCellCounter, average)
                        .Style
                        .Alignment
                        .SetHorizontal(XLAlignmentHorizontalValues.Center)
                        .Fill
                        .SetBackgroundColor(Formatter.GetNumberColor(average));
                }

                specimenRowCounter++;
            }
        }

        specimenCellCounter++;
    }

    #endregion Specimen

    #region Specimen-Location

    foreach (var location in baseData.Locations)
    {
        var specimenLocationRowCounter = 1;
        var specimenLocationCellCounter = 1;
        var sheetName = $"{specimen.Replace("/", "")}-{location}";

        if (sheetName.Length > 31) sheetName = sheetName[..31];

        var specimenLocationSheet = resultWorkbook.Worksheets.Add(sheetName);
        specimenLocationSheet.SetCellValue(specimenLocationRowCounter, specimenLocationCellCounter, $"PETA POLA KUMAN {year}");
        specimenLocationRowCounter++;
        specimenLocationSheet.SetCellValue(specimenLocationRowCounter, specimenLocationCellCounter, "Organism");
        specimenLocationCellCounter++;

        foreach (var organism in baseData.Organisms)
        {
            specimenLocationSheet
                .SetCellValue(specimenLocationRowCounter, specimenLocationCellCounter, organism)
                .Style.Alignment
                .SetTextRotation(90)
                .Alignment
                .SetHorizontal(XLAlignmentHorizontalValues.Center);
            specimenLocationCellCounter++;
        }

        specimenLocationRowCounter++;
        specimenLocationCellCounter = 1;
        specimenLocationSheet.SetCellValue(specimenLocationRowCounter, specimenLocationCellCounter, "Number of isolates");
        specimenLocationCellCounter++;

        var specimenLocationData = baseData.Records.Where(r => r.Specimen.ToLower() == specimen.ToLower() && r.Location.ToLower() == location.ToLower());

        foreach (var organism in baseData.Organisms)
        {
            var organismCounter = specimenLocationData.Where(s => s.Organism.ToLower() == organism.ToLower()).Sum(s => s.ResistanceRates.Count);

            if (organismCounter > 0)
                specimenLocationSheet
                    .SetCellValue(specimenLocationRowCounter, specimenLocationCellCounter, organismCounter)
                    .Style
                    .Alignment
                    .SetHorizontal(XLAlignmentHorizontalValues.Center);

            specimenLocationCellCounter++;
        }

        specimenLocationRowCounter++;
        var staticLocationRowCounter = specimenLocationRowCounter;
        specimenLocationCellCounter = 1;

        foreach (var antibiotic in baseData.Antibiotics)
        {
            specimenLocationSheet.SetCellValue(specimenLocationRowCounter, specimenLocationCellCounter, antibiotic);
            specimenLocationRowCounter++;
        }

        specimenLocationCellCounter++;

        foreach (var organism in baseData.Organisms)
        {
            specimenLocationRowCounter = staticLocationRowCounter;
            var recordResistanceRates = specimenLocationData.Where(a => a.Organism.ToLower() == organism.ToLower()).SelectMany(o => o.ResistanceRates);

            if (recordResistanceRates.CheckNotNullAndAny())
            {
                foreach (var antibiotics in baseData.Antibiotics)
                {
                    var antibioticsData = recordResistanceRates.Where(a => a.Antibiotic.ToLower() == antibiotics.ToLower());

                    if (antibioticsData != null && antibioticsData.Any())
                    {
                        var sum = antibioticsData.Where(a => a.Antibiotic.ToLower() == antibiotics.ToLower()).Sum(s => s.Value);
                        var average = antibioticsData.Any() ? ((double)(antibioticsData.Sum(s => s.Value) / antibioticsData.Count())).ToInt() : 0;
                        specimenLocationSheet
                            .SetCellValue(specimenLocationRowCounter, specimenLocationCellCounter, average)
                            .Style
                            .Alignment
                            .SetHorizontal(XLAlignmentHorizontalValues.Center)
                            .Fill
                            .SetBackgroundColor(Formatter.GetNumberColor(average));
                    }

                    specimenLocationRowCounter++;
                }
            }

            specimenLocationCellCounter++;
        }
    }

    #endregion Specimen-Location

    specimenSheet.Columns().AdjustToContents();
}

if (File.Exists(targetFilePath)) File.Delete(targetFilePath);

var base64String = new ExcelWriter().WriteToBase64(resultWorkbook);
File.WriteAllBytes(targetFilePath, Convert.FromBase64String(base64String));