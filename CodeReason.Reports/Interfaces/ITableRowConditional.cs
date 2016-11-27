namespace CodeReason.Reports.Interfaces
{
    public interface ITableRowConditional
    {
        bool Visible { get; set; }

        string TableName { get; set; }
    }
}