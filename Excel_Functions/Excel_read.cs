using ClosedXML.Excel;

namespace Excel_Functions
{
    public partial class Excel
    {
        #region Properties
        private readonly List<String> Path_Strings = new();
        #endregion
        #region Constructors
        public Excel() { }
        public Excel(String path) => Path_Strings.Add(path ?? throw new ArgumentNullException(nameof(path)));
        public Excel(List<String> path_Strings) => Path_Strings = path_Strings ?? throw new ArgumentNullException(nameof(path_Strings));
        #endregion
        #region Read Functions
        public Boolean Read(out List<Excel_Data> Data)
        {
            List<Excel_Data> return_data = new();
            Data = return_data;
            foreach (String key in Path_Strings)
            {
                Excel_Data eData = new()
                {
                    FileName = key
                };
                XLWorkbook wb = new(key);
                eData.Data = new();
                foreach (IXLWorksheet ws in wb.Worksheets)
                {
                    List<cell_Data> datares = new();
                    IXLCells cells = ws.CellsUsed();
                    foreach (IXLCell cell in cells)
                    {
                        IXLAddress address = cell.Address;
                        cell_Data cd = new()
                        {
                            Col = address.ColumnNumber,
                            Row = address.RowNumber,
                            Value = cell.Value.ToString(),
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
            Boolean return_result = true;
            Data = return_data;
            return return_result;
        }
        #endregion
    }
}