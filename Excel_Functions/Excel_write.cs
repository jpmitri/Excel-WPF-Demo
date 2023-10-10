using ClosedXML.Excel;

namespace Excel_Functions
{
    public partial class Excel
    {

        #region Inbound
        public Boolean writeInbound(String Dest, String Joker = "")
        {
            if (Read(out List<Excel_Data>? Data))
            {
                List<Inbound_Data> inbound_Data = new();
                foreach (Excel_Data Item in Data)
                {
                    XLWorkbook wb = new();
                    IXLWorksheet ws = wb.AddWorksheet($"{Item!.Data!.First().Key.Trim()[..^2]}");
                    ws.Cell("A1").SetValue("Inbound").Style.Font.Bold = true;
                    List<String> IbVal = new();
                    Int32 palletCounter = 0;
                    List<ExcelFunction> ef = new();
                    foreach (KeyValuePair<String, List<cell_Data>> sheet in Item!.Data!)
                    {
                        foreach (cell_Data item in sheet.Value)
                        {
                            Int32 colItemVal = sheet.Value[0].Col;
                            if (
                                (!IbVal.Contains(item!.Value!) &&
                                item.Col == colItemVal &&
                                !item!.Value!.ToLower().Trim().Equals("inbound") &&
                                !item.Value.ToLower().Trim().Contains("comments"))
                                || (item.Value == "Pallet Supplied" && palletCounter != 2)
                                )
                            {
                                if (item.Value == "Pallet Supplied")
                                {
                                    palletCounter++;
                                }
                                IbVal.Add(item.Value);
                            }
                        }
                    }
                    Int32 col = 2;
                    Int32 row = 2;
                    Int32 CostRow = 0;
                    foreach (String item in IbVal)
                    {
                        if (item.ToLower().Trim().Equals("Management - Daily".ToLower().Trim()))
                        {
                            row++;
                            CostRow = row;
                        }
                        _ = ws.Cell(row, 1).SetValue(item);
                        Inbound_Data toadd = new()
                        {
                            Row = row,
                            Name = item,
                            isRevenue = CostRow != 0
                        };
                        try
                        {
                            if (Double.TryParse(Item.Data.First().Value.First(x => x.Row == row && x.Col == 4).Value, out Double res))
                            {
                                toadd.Price = res;
                            }
                        }
                        catch (Exception)
                        {
                            Int32 adjusterRow = row;
                            while (true)
                            {
                                adjusterRow++;
                                if (Double.TryParse(Item.Data.First().Value.First(x => x.Row == adjusterRow && x.Col == 4).Value, out Double res))
                                {
                                    toadd.Price = res;
                                    break;
                                }
                            }
                        }
                        inbound_Data.Add(toadd);
                        row++;
                    }
                    row++;
                    Int32 row1 = row;
                    ws.Cell(row, 1).SetValue("Daily Revenue").Style.Font.Bold = true;
                    row++;
                    Int32 row2 = row;
                    ws.Cell(row, 1).SetValue("Daily Cost").Style.Font.Bold = true;
                    row++;
                    ws.Cell(row, 1).SetValue("Daily Summary").Style.Font.Bold = true;
                    foreach (KeyValuePair<String, List<cell_Data>> sheet in Item.Data)
                    {
                        Int32 col2 = col + 1;
                        Int32 qtyCol = col;
                        Int32 MergedCol = col;
                        _ = ws.Range(row, MergedCol, row, col2).Row(1).Merge();
                        _ = ws.Range(row1, MergedCol, row1, col2).Row(1).Merge();
                        _ = ws.Range(row2, MergedCol, row2, col2).Row(1).Merge();
                        ws.Cell(1, col).SetValue($"Qty {sheet.Key.Trim()}").Style.Font.Bold = true;
                        col++;
                        Int32 hrsCol = col;
                        ws.Cell(1, col).SetValue($"Hrs {sheet.Key.Trim()}").Style.Font.Bold = true;
                        col++;
                        Int32 itemOffset = 1;
                        foreach (cell_Data item in sheet.Value)
                        {

                            if (item!.Value!.ToLower().Trim().Contains("comments"))
                            {
                                break;
                            }
                            if (item.Col == 2)
                            {
                                try
                                {
                                    Inbound_Data id = inbound_Data.First(x => x.Name == sheet.Value[itemOffset].Value);
                                    Int32 loc = inbound_Data.LastIndexOf(
                                        inbound_Data.First(x => x.Name == sheet.Value[itemOffset].Value
                                        ));
                                    inbound_Data[loc].isQty = true;
                                    inbound_Data[loc].TotalUA += Double.Parse(item.Value);
                                    _ = ws.Cell(inbound_Data[loc].Row, qtyCol).SetValue(Double.Parse(item.Value));
                                    _ = ws.Cell(inbound_Data[loc].Row, hrsCol).Style.Fill.SetBackgroundColor(XLColor.Gray);
                                }
                                catch { continue; }
                            }
                            else if (item.Col == 3)
                            {
                                try
                                {
                                    Inbound_Data id = inbound_Data.Last(x => x.Name == sheet.Value[itemOffset].Value);
                                    Int32 loc = inbound_Data.LastIndexOf(
                                        inbound_Data.Last(x => x.Name == sheet.Value[itemOffset].Value
                                        ));
                                    inbound_Data[loc].isQty = false;
                                    inbound_Data[loc].TotalUA += Double.Parse(item.Value);
                                    _ = ws.Cell(inbound_Data[loc].Row, hrsCol).SetValue(Double.Parse(item.Value));
                                    _ = ws.Cell(inbound_Data[loc].Row, qtyCol).Style.Fill.SetBackgroundColor(XLColor.Gray);
                                }
                                catch { continue; }
                            }
                            itemOffset++;
                        }
                        ef.Add(new()
                        {
                            Row = row1,
                            Col = MergedCol,
                            Function = $"SUMPRODUCT({IntToLeters[MergedCol]}2:{IntToLeters[MergedCol]}16,@CUST@2:@CUST@16)+SUMPRODUCT({IntToLeters[MergedCol + 1]}2:{IntToLeters[MergedCol + 1]}16,@CUST@2:@CUST@16)"

                        });
                        ef.Add(new()
                        {
                            Row = row2,
                            Col = MergedCol,
                            Function = $"SUMPRODUCT({IntToLeters[MergedCol]}18:{IntToLeters[MergedCol]}22,@CUST@18:@CUST@22)+SUMPRODUCT({IntToLeters[MergedCol + 1]}18:{IntToLeters[MergedCol + 1]}22,@CUST@18:@CUST@22)"
                        });
                        ws.Cell(row, MergedCol).FormulaA1 = $"{IntToLeters[MergedCol]}{row1}-{IntToLeters[MergedCol]}{row2}";
                        ws.Cell(row, MergedCol).Style.NumberFormat.Format = "$#,##0.00";
                    }
                    ws.Cell(1, col).SetValue($"Price").Style.Font.Bold = true;
                    Int32 PriceOffset = 2;
                    foreach (Inbound_Data item in inbound_Data)
                    {
                        if (item.Name == "Management - Daily")
                        {
                            PriceOffset++;
                        }
                        ws.Cell(PriceOffset, col).SetValue(item.Price).Style.NumberFormat.Format = "$#,##0.00";
                        PriceOffset++;
                    }

                    foreach (ExcelFunction item in ef)
                    {
                        item.Function = item!.Function!.Replace("@CUST@", $"{IntToLeters[col]}");
                        ws.Cell(item.Row, item.Col).FormulaA1 = item.Function;
                        ws.Cell(item.Row, item.Col).Style.NumberFormat.Format = "$#,##0.00";
                    }
                    col++;
                    ws.Cell(1, col).SetValue($"Monthly Revenue").Style.Font.Bold = true;
                    col += 2;
                    ws.Cell(1, col).SetValue($"Monthly Inbound Summary").Style.Font.Bold = true;
                    Int32 col3 = col + 2;
                    _ = ws.Range(1, col, 1, col3).Row(1).Merge();
                    _ = ws.Columns().AdjustToContents();

                    wb.SaveAs($"{Dest}\\{Joker}_{Item!.FileName![(Item!.FileName!.LastIndexOf("\\") + 1)..]}");
                }

                return true;
            }
            return false;
        }

        #endregion
        #region Outbound
        #endregion
        #region Special Project
        #endregion
    }
}
