using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System.Threading.Tasks;

namespace EarthML.Temply.Core
{
    public class BaseProcessorProvider : IProcessorProvider
    {
        public string Name { get; set; }

        public virtual async Task UpdateElement(MainDocumentPart mainPart, SdtElement element, TemplateReplacement tag)
        {
             

        }
    }
}
