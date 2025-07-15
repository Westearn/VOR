using System.Collections.Generic;

namespace VOR.Helpers.Import
{
    public interface IImporter<T>
    {
        List<T> Import(string filePath);
    }
}
