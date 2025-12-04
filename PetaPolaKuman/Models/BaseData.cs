namespace PetaPolaKuman.Models;

public class BaseData
{
    public List<string> Antibiotics { get; set; } = [];
    public List<string> Organisms { get; set; } = [];
    public List<string> Specimens { get; set; } = [];
    public List<string> Locations { get; set; } = [];
    public List<Record> Records { get; set; } = [];
}

public class Record
{
    public string Location { get; set; }
    public string Specimen { get; set; }
    public string Organism { get; set; }
    public List<ResistanceRate> ResistanceRates { get; set; } = [];
}

public class ResistanceRate(string code, string antibiotic, int value)
{
    public ResistanceRate(string code, int value) : this(code, null, value)
    {
    }

    public string Code { get; set; } = code;
    public string Antibiotic { get; set; } = antibiotic;
    public int Value { get; set; } = value;
}

public class ResistanceRates
{
    public ResistanceRates()
    {
        Rates =
        [
            new("tidak tercantum dlm taksonomi CLSI", 0),
            new("tidak ada data AST di CLSI", 0),
            new("R", 0),
            new("R", 0),
            new("I", 75),
            new("S", 100)
        ];
    }

    public List<ResistanceRate> Rates { get; set; } = [];
}