using ClosedXML.Excel;
using PetaPolaKuman.Models;
using PetaPolaKuman.Utilities;
using Travewell.Library.Extensions;

string fileName = "..\\..\\..\\source.xlsx";
var antibioticsLineNumber = 7;
var antibioticsStartColumn = "M".ExcelColumnNameToNumber();
var antibioticsEndColumn = "BA".ExcelColumnNameToNumber();
var sourceWorkbook = new XLWorkbook(fileName);
var antibioticsRow = sourceWorkbook.Worksheet(1).RangeUsed().RowsUsed().Skip(antibioticsLineNumber - 1).FirstOrDefault();
var baseData = new BaseData
{
    Antibiotics = new List<string>(),
    Records = new List<Record>()
};

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
    var record = new Record
    {
        Location = recordsRow.GetCellValue(locationColumn),
        Specimen = recordsRow.GetCellValue(specimenColumn),
        Organism = recordsRow.GetCellValue(organismColumn),
        Resistances = new List<AntibioticResistance>()
    };

    var antibioticsCounter = 0;

    for (var a = antibioticsStartColumn; a <= antibioticsEndColumn; a++)
    {
        if (!recordsRow.GetCellValue(a).IsNullOrEmpty())
        {
            record.Resistances.Add(new AntibioticResistance() {
                Antibiotic = baseData.Antibiotics[antibioticsCounter],
                ResistanceRate = recordsRow.GetCellValue(a).IsNullOrEmpty()
            });
        }

        antibioticsCounter++;
    }

    baseData.Records.Add(record);
}

using var resultWorkbook = new XLWorkbook();