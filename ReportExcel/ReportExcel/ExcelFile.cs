
namespace ReportExcel
{
    class ExcelFile
    {
        public int RowName;
        public int ColumnName;
        public int RowPrice;
        public int ColumnPrice;
        public string FullfileName;

        public ExcelFile()
        {
            RowName = -1;
            RowPrice = -1;
            ColumnName = -1;
            ColumnPrice = -1;
            FullfileName = null;
        }
    }
}
