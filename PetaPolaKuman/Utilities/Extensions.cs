namespace PetaPolaKuman.Utilities;

public static class Extensions
{
    public static string OrganismTranslator(this string input) => input.ToLower() switch
    {
        "staphylococcus hemolyticus" => "Staphylococcus haemolyticus",
        "eschericia coli" => "Escherichia coli",
        "staphyloccous hominis" => "Staphylococcus hominis",
        _ => input
    };
}