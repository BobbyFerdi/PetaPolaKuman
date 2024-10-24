namespace PetaPolaKuman.Models;

public class BaseData
{
    public List<string> Antibiotics { get; set; } = new();
    public List<string> Organisms { get; set; } = new();
    public List<string> Specimens { get; set; } = new();
    public List<string> Locations { get; set; } = new();
    public List<Record> Records { get; set; } = new();
}

public class Record
{
    public string Location { get; set; }
    public string Specimen { get; set; }
    public string Organism { get; set; }
    public List<ResistanceRate> ResistanceRates { get; set; } = new();
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
            new("tidak tercantum dlm taksonomi CLSI", 0),
            new("tidak ada data AST di CLSI", 0),
            new("R", 0),
            new("R", 0),
            new("I", 75),
            new("S", 100)
        };
    }

    public List<ResistanceRate> Rates { get; set; } = new();
}