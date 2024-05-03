using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml.Wordprocessing;

namespace GenTimeSheet.Core
{
    internal class Serialize
    {
        private readonly Table _table;

        internal Serialize(Table table)
        {
            _table = table;
        }

        internal IEnumerable<IEnumerable<string>> GetTable()        
            => _table.Elements<TableRow>().Select(r => r.Elements<TableCell>().Select(c => c.InnerText));        
    }
}
