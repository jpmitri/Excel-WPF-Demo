using ClosedXML.Excel;

namespace Excel_Functions
{
    public partial class Excel
    {
    #region Properties
        private readonly List<string> Path_Strings = new();
    #endregion
    #region Read Functions
        public bool Read(out List<Excel_Data> Data)
        {
            List<Excel_Data> return_data = new();
            Data = return_data;
            try
            {
                foreach (string key in Path_Strings)
                {
                    Excel_Data eData = new() {FileName = key};
                    XLWorkbook wb    = new(key);
                    eData.Data = new Dictionary<string, List<cell_Data>>();
                    foreach (IXLWorksheet ws in wb.Worksheets)
                    {
                        List<cell_Data> datares = new();
                        IXLCells        cells   = ws.CellsUsed();
                        foreach (IXLCell cell in cells)
                        {
                            IXLAddress address = cell.Address;
                            cell_Data cd = new()
                            {
                                Col       = address.ColumnNumber,
                                Row       = address.RowNumber,
                                Value     = cell.Value.ToString(),
                                IsFormula = cell.HasFormula
                            };
                            if (cd.IsFormula)
                            {
                                cd.Formula = cell.FormulaA1;
                            }
                            datares.Add(cd);
                        }
                        eData.Data.Add(ws.Name, datares);
                    }
                    return_data.Add(eData);
                }
                Data = return_data;
                return true;
            }
            catch
            {
                return false;
            }
        }
    #endregion
    #region Constructors
        public Excel() { }
        public Excel(string path) => Path_Strings.Add(path ?? throw new ArgumentNullException(nameof(path)));
        public Excel(List<string> path_Strings) =>
            Path_Strings = path_Strings ?? throw new ArgumentNullException(nameof(path_Strings));
    #endregion
    }
}
