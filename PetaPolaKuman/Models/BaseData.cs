using static PetaPolaKuman.Models.Enums;

namespace PetaPolaKuman.Models;

internal class BaseData
{
    public List<string> Antibiotics { get; set; }
    public List<Record> Records { get; set; }
}

internal class Record
{
    public string Location { get; set; }
    public string Specimen { get; set; }
    public string Organism { get; set; }
    public List<AntibioticResistance> Resistances { get; set; }
}

internal class AntibioticResistance
{
    public string Antibiotic { get; set; }
    public Rate ResistanceRate { get; set; }
}