using System.Windows.Documents;
using CodeReason.Reports.Interfaces;

namespace CodeReason.Reports.Document
{
    public class TableRowConditional : TableRow, ITableRowConditional
    {
        public bool Visible { get; set; }

        public string TableName { get; set; }
    }
}