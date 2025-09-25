using Autodesk.Revit.Attributes;

namespace ATP_Common_Plugin.Commands.Calculation.SpacesEnvelopeExport
{
    [Regeneration(RegenerationOption.Manual)]
    [Transaction(TransactionMode.Manual)]
    /// <summary>
    /// Placeholder for future filtering (levels, phases, categories, etc.).
    /// </summary>
    public sealed class ExportSpacesEnvelopeOptions
    {
        public bool IncludeAllLevels { get; set; } = true;
        public bool UseTrueNorth { get; set; } = true;
        public bool SkipTinyFaces { get; set; } = true;
    }
}