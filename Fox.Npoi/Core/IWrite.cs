using System.Collections.Generic;

namespace Fox.Npoi.Core
{
    public interface IWrite
    {
        void CreateRow(int sheetIndex, int rowIndex);

        void CreateSheet(string name);

        void WriteTitle(string[] titles, int sheetIndex, int rowIndex, int cellIndex = 0, int styleIndex = 1);

        void WriteValue(int sheetIndex, int rowIndex, int cellIndex, dynamic value, int styleIndex, string formula = null);

        int WriteProperty<T>(T entity, int sheetIndex, int rowIndex, int cellIndex = 0, int styleIndex = 2);

        void WriteEnumerable<T>(IEnumerable<T> entities, int sheetIndex, int rowIndex);

        void WriteObject<T>(ICollection<T> entities, int sheetIndex, int rowIndex);

        byte[] WriteStream();

        void WriteFile(string filePath);
    }
}