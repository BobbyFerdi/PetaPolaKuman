using ClosedXML.Excel;

namespace PetaPolaKuman.Models;

public class BaseData
{
    public List<string> Antibiotics { get; set; }
    public List<string> Organisms { get; set; }
    public List<string> Specimens { get; set; }
    public List<string> Locations { get; set; }
    public List<Record> Records { get; set; }
}

public class Record
{
    public string Location { get; set; }
    public string Specimen { get; set; }
    public string Organism { get; set; }
    public List<ResistanceRate> ResistanceRates { get; set; }
}

public class ResistanceRate
{
    public ResistanceRate(string code, string antibiotic, int value)
    {
        Code = code;
        Antibiotic = antibiotic;
        Value = value;
    }

    public ResistanceRate(string code, int value)
    {
        Code = code;
        Value = value;
    }

    public string Code { get; set; }
    public string Antibiotic { get; set; }
    public int Value { get; set; }
}

public class ResistanceRates
{
    public ResistanceRates()
    {
        Rates = new List<ResistanceRate>
        {
            new ResistanceRate("tidak tercantum dlm taksonomi CLSI", 0),
            new ResistanceRate("tidak ada data AST di CLSI", 0),
            new ResistanceRate("R", 0),
            new ResistanceRate("R", 0),
            new ResistanceRate("I", 75),
            new ResistanceRate("S", 100),
        };
    }

    public List<ResistanceRate> Rates { get; set; }
}