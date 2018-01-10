using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System.Threading.Tasks;

namespace EarthML.Temply.Core
{
    public interface IProcessorProvider
    {
        string Name { get; }

        Task UpdateElement(MainDocumentPart mainPart, SdtElement element, TemplateReplacement tag);
    }
}
