using System.Collections.Generic;
using System.IO;

namespace Core.Interfaces
{
    public interface ICSVExportService
    {
        void GenericExport(object[] dList, string FileName, bool SpaceCapitals = false);
        string GenerateCSVData(object[] dList, bool SpaceCapitals = false);
    }
}
