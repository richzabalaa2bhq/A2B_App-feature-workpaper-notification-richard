using A2B_App.Server.Data;
using A2B_App.Shared.Sox;
using AutoMapper.Configuration;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Linq;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace A2B_App.Server.Services
{
    public class ExcelService
    {
        //private SoxContext _soxContext;
        //private readonly IConfiguration _config;
        //public ExcelService(SoxContext soxContext, IConfiguration config)
        //{
        //    _soxContext = soxContext;
        //    _config = config;
        //}

        public List<Population> ExcelPopulation(ExcelWorksheet ws, int type)
        {
            List<Population> listPopulation = new List<Population>();
            int colCount = ws.Dimension.End.Column;
            int rowCount = ws.Dimension.End.Row;
            //int type = 1;
            for (int row = 1; row <= rowCount; row++)
            {
                //if (ws.Cells[row, 2].Value != null){}

                Population population = new Population();
                for (int col = 1; col <= colCount; col++)
                {

                    population.PopType = type;

                    switch (col)
                    {

                        case 1:
                            if (ws.Cells[row, col].Value?.ToString() != string.Empty)
                            {
                                population.UniqueId = ws.Cells[row, col].Value?.ToString();
                            }
                            break;
                        case 2:
                            if (ws.Cells[row, col].Value?.ToString() != string.Empty)
                            {
                                population.PurchaseOrder = ws.Cells[row, col].Value?.ToString();
                            }
                            break;
                        case 3:
                            if (ws.Cells[row, col].Value?.ToString() != string.Empty)
                            {
                                population.SupplierSched = ws.Cells[row, col].Value?.ToString();
                            }
                            break;
                        case 4:
                            if (ws.Cells[row, col].Value?.ToString() != string.Empty)
                            {
                                population.PoRev = ws.Cells[row, col].Value?.ToString();
                            }
                            break;
                        case 5:
                            if (ws.Cells[row, col].Value?.ToString() != string.Empty)
                            {
                                population.PoLine = ws.Cells[row, col].Value?.ToString();
                            }
                            break;
                        case 6:
                            if (ws.Cells[row, col].Value?.ToString() != string.Empty)
                            {
                                population.Requisition = ws.Cells[row, col].Value?.ToString();
                            }
                            break;
                        case 7:
                            if (ws.Cells[row, col].Value?.ToString() != string.Empty)
                            {
                                population.RequisitionLine = ws.Cells[row, col].Value?.ToString();
                            }
                            break;
                        case 8:
                            if (ws.Cells[row, col].Value?.ToString() != string.Empty)
                            {
                                population.EnteredBy = ws.Cells[row, col].Value?.ToString();
                            }
                            break;
                        case 9:
                            if (ws.Cells[row, col].Value?.ToString() != string.Empty)
                            {
                                population.Status = ws.Cells[row, col].Value?.ToString();
                            }
                            break;
                        case 10:
                            if (ws.Cells[row, col].Value?.ToString() != string.Empty)
                            {
                                population.Buyer = ws.Cells[row, col].Value?.ToString();
                            }
                            break;
                        case 11:
                            if (ws.Cells[row, col].Value?.ToString() != string.Empty)
                            {
                                population.Contact = ws.Cells[row, col].Value?.ToString();
                            }
                            break;
                        case 12:
                            if (ws.Cells[row, col].Value?.ToString() != string.Empty)
                            {
                                population.OrderDate = ws.Cells[row, col].Value?.ToString();
                            }
                            break;
                        case 13:
                            if (ws.Cells[row, col].Value?.ToString() != string.Empty)
                            {
                                population.Supplier = ws.Cells[row, col].Value?.ToString();
                            }
                            break;
                        case 14:
                            if (ws.Cells[row, col].Value?.ToString() != string.Empty)
                            {
                                population.ShipTo = ws.Cells[row, col].Value?.ToString();
                            }
                            break;
                        case 15:
                            if (ws.Cells[row, col].Value?.ToString() != string.Empty)
                            {
                                population.SortName = ws.Cells[row, col].Value?.ToString();
                            }
                            break;
                        case 16:
                            if (ws.Cells[row, col].Value?.ToString() != string.Empty)
                            {
                                population.Telephone = ws.Cells[row, col].Value?.ToString();
                            }
                            break;
                        case 17:
                            if (ws.Cells[row, col].Value?.ToString() != string.Empty)
                            {
                                population.ItemNumber = ws.Cells[row, col].Value?.ToString();
                            }
                            break;
                        case 18:
                            if (ws.Cells[row, col].Value?.ToString() != string.Empty)
                            {
                                population.ProdLine = ws.Cells[row, col].Value?.ToString();
                            }
                            break;
                        case 19:
                            if (ws.Cells[row, col].Value?.ToString() != string.Empty)
                            {
                                population.ProdDescription = ws.Cells[row, col].Value?.ToString();
                            }
                            break;
                        case 20:
                            if (ws.Cells[row, col].Value?.ToString() != string.Empty)
                            {
                                population.Site = ws.Cells[row, col].Value?.ToString();
                            }
                            break;
                        case 21:
                            if (ws.Cells[row, col].Value?.ToString() != string.Empty)
                            {
                                population.Location = ws.Cells[row, col].Value?.ToString();
                            }
                            break;
                        case 22:
                            if (ws.Cells[row, col].Value?.ToString() != string.Empty)
                            {
                                population.ItemRevision = ws.Cells[row, col].Value?.ToString();
                            }
                            break;
                        case 23:
                            if (ws.Cells[row, col].Value?.ToString() != string.Empty)
                            {
                                population.SupplierItem = ws.Cells[row, col].Value?.ToString();
                            }
                            break;
                        case 24:
                            if (ws.Cells[row, col].Value?.ToString() != string.Empty)
                            {
                                population.QuantityOrdered = ws.Cells[row, col].Value?.ToString();
                            }
                            break;
                        case 25:
                            if (ws.Cells[row, col].Value?.ToString() != string.Empty)
                            {
                                population.UnitOfMeasure = ws.Cells[row, col].Value?.ToString();
                            }
                            break;
                        case 26:
                            if (ws.Cells[row, col].Value?.ToString() != string.Empty)
                            {
                                population.UMConversion = ws.Cells[row, col].Value?.ToString();
                            }
                            break;
                        case 27:
                            if (ws.Cells[row, col].Value?.ToString() != string.Empty)
                            {
                                population.QtyOrderedXPOCost = ws.Cells[row, col].Value?.ToString();
                            }
                            break;
                        case 28:
                            if (ws.Cells[row, col].Value?.ToString() != string.Empty)
                            {
                                population.QuantityReceived = ws.Cells[row, col].Value?.ToString();
                            }
                            break;
                        case 29:
                            if (ws.Cells[row, col].Value?.ToString() != string.Empty)
                            {
                                population.QtyOpen = ws.Cells[row, col].Value?.ToString();
                            }
                            break;
                        case 30:
                            if (ws.Cells[row, col].Value?.ToString() != string.Empty)
                            {
                                population.QtyReturned = ws.Cells[row, col].Value?.ToString();
                            }
                            break;
                        case 31:
                            if (ws.Cells[row, col].Value?.ToString() != string.Empty)
                            {
                                population.DueDate = ws.Cells[row, col].Value?.ToString();
                            }
                            break;
                        case 32:
                            if (ws.Cells[row, col].Value?.ToString() != string.Empty)
                            {
                                population.OverDue = ws.Cells[row, col].Value?.ToString();
                            }
                            break;
                        case 33:
                            if (ws.Cells[row, col].Value?.ToString() != string.Empty)
                            {
                                population.PerformanceDate = ws.Cells[row, col].Value?.ToString();
                            }
                            break;
                        case 34:
                            if (ws.Cells[row, col].Value?.ToString() != string.Empty)
                            {
                                population.Currency = ws.Cells[row, col].Value?.ToString();
                            }
                            break;
                        case 35:
                            if (ws.Cells[row, col].Value?.ToString() != string.Empty)
                            {
                                population.StandardCost = ws.Cells[row, col].Value?.ToString();
                            }
                            break;
                        case 36:
                            if (ws.Cells[row, col].Value?.ToString() != string.Empty)
                            {
                                population.PurchasedCost = ws.Cells[row, col].Value?.ToString();
                            }
                            break;
                        case 37:
                            if (ws.Cells[row, col].Value?.ToString() != string.Empty)
                            {
                                population.PurCostBC = ws.Cells[row, col].Value?.ToString();
                            }
                            break;
                        case 38:
                            if (ws.Cells[row, col].Value?.ToString() != string.Empty)
                            {
                                population.OpenPoCost = ws.Cells[row, col].Value?.ToString();
                            }
                            break;
                        case 39:
                            if (ws.Cells[row, col].Value?.ToString() != string.Empty)
                            {
                                population.PpvPerUnit = ws.Cells[row, col].Value?.ToString();
                            }
                            break;
                        case 40:
                            if (ws.Cells[row, col].Value?.ToString() != string.Empty)
                            {
                                population.Type = ws.Cells[row, col].Value?.ToString();
                            }
                            break;
                        case 41:
                            if (ws.Cells[row, col].Value?.ToString() != string.Empty)
                            {
                                population.StdMtlCostNow = ws.Cells[row, col].Value?.ToString();
                            }
                            break;
                        case 42:
                            if (ws.Cells[row, col].Value?.ToString() != string.Empty)
                            {
                                population.WorkOrderId = ws.Cells[row, col].Value?.ToString();
                            }
                            break;
                        case 43:
                            if (ws.Cells[row, col].Value?.ToString() != string.Empty)
                            {
                                population.Operation = ws.Cells[row, col].Value?.ToString();
                            }
                            break;
                        case 44:
                            if (ws.Cells[row, col].Value?.ToString() != string.Empty)
                            {
                                population.PurchAcct = ws.Cells[row, col].Value?.ToString();
                            }
                            break;
                        case 45:
                            if (ws.Cells[row, col].Value?.ToString() != string.Empty)
                            {
                                population.GlAccountDesc = ws.Cells[row, col].Value?.ToString();
                            }
                            break;
                        case 46:
                            if (ws.Cells[row, col].Value?.ToString() != string.Empty)
                            {
                                population.CostCenter = ws.Cells[row, col].Value?.ToString();
                            }
                            break;
                        case 47:
                            if (ws.Cells[row, col].Value?.ToString() != string.Empty)
                            {
                                population.GlDescription = ws.Cells[row, col].Value?.ToString();
                            }
                            break;
                        case 48:
                            if (ws.Cells[row, col].Value?.ToString() != string.Empty)
                            {
                                population.Project = ws.Cells[row, col].Value?.ToString();
                            }
                            break;
                        case 49:
                            if (ws.Cells[row, col].Value?.ToString() != string.Empty)
                            {
                                population.Description = ws.Cells[row, col].Value?.ToString();
                            }
                            break;
                        case 50:
                            if (ws.Cells[row, col].Value?.ToString() != string.Empty)
                            {
                                population.Taxable = ws.Cells[row, col].Value?.ToString();
                            }
                            break;
                        case 51:
                            if (ws.Cells[row, col].Value?.ToString() != string.Empty)
                            {
                                population.Comments = ws.Cells[row, col].Value?.ToString();
                            }
                            break;

                    }
                }

                if (population != null)
                {
                    listPopulation.Add(population);
                }


            }

            return listPopulation;

        }

        public List<List<string>> ExcelPopulation2(ExcelWorksheet ws, int type)
        {
            int colCount = ws.Dimension.End.Column;
            int rowCount = ws.Dimension.End.Row;
            //List<string>[] listPopulation = new List<string>[];
            List<List<string>> parentPopulation = new List<List<string>>();
            bool isEmpty;
            //int type = 1;
            for (int row = 1; row <= rowCount; row++)
            {
                isEmpty = true;
                List<string> childPopulation = new List<string>();
                for (int col = 1; col <= colCount; col++)
                {
                    bool isDateHeader = IsDateHeader(ws.Cells[1, col].Value?.ToString());
                    bool isDateTime = IsDateTime(ws.Cells[row, col].Value?.ToString());
                    bool isDecimal = IsDecimal(ws.Cells[row, col].Value?.ToString());
                    bool isIdentificationHeader = IsIdentificationHeader(ws.Cells[1, col].Value?.ToString());

                    Debug.WriteLine($"Cell[{row},{col}] : {ws.Cells[row, col].Value?.ToString()} - IdentificationHeader({isIdentificationHeader}) | DateHeader({isDateHeader}) | Datetime({isDateTime}) | Decimal({isDecimal})");
                    //Debug.WriteLine($"Cell[{row},{col}] : {ws.Cells[row, col].Value?.ToString()}");

                    if (isDateTime && isDateHeader)
                    {
                        DateTime dtVal = DateTime.Parse(ws.Cells[row, col].Value?.ToString());
                        childPopulation.Add(dtVal.ToString("MM/dd/yyyy"));
                    }
                    else if (isDecimal && !isIdentificationHeader && !isDateHeader)
                    {
                        decimal decVal = decimal.Parse(ws.Cells[row, col].Value?.ToString());
                        childPopulation.Add(decVal.ToString("0.00"));
                    }
                    else
                    {
                        childPopulation.Add(ws.Cells[row, col].Value?.ToString());
                    }

                    if (ws.Cells[row, col].Value?.ToString() != string.Empty && ws.Cells[row, col].Value?.ToString() != null)
                    {
                        isEmpty = false;
                    }
                }

                if (childPopulation != null && !isEmpty)
                {
                    parentPopulation.Add(childPopulation);
                }

            }

            return parentPopulation;

        }

        //Excel function to set boarder range
        public void ExcelSetBorder(ExcelWorksheet ws, string range)
        {
            ws.Cells[range].Style.Border.Top.Style = ExcelBorderStyle.Thin;
            ws.Cells[range].Style.Border.Left.Style = ExcelBorderStyle.Thin;
            ws.Cells[range].Style.Border.Right.Style = ExcelBorderStyle.Thin;
            ws.Cells[range].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
        }

        //Excel function to set align center
        public void ExcelSetAlignCenter(ExcelWorksheet ws, string range)
        {
            ws.Cells[range].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            ws.Cells[range].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
        }

        //Excel function to set arial 12
        public void ExcelSetArialSize12(ExcelWorksheet ws, string range)
        {
            ws.Cells[range].Style.Font.SetFromFont(new Font("Arial", 12));
        }

        //Excel function to set background #ccffcc
        public void ExcelSetBackgroundGreen(ExcelWorksheet ws, string range)
        {
            ws.Cells[range].Style.Fill.PatternType = ExcelFillStyle.Solid;
            ws.Cells[range].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#ccffcc"));
        }

        //Excel function to set background #d9e1f2
        public void ExcelSetBackgroundBlue(ExcelWorksheet ws, string range)
        {
            ws.Cells[range].Style.Fill.PatternType = ExcelFillStyle.Solid;
            ws.Cells[range].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#d9e1f2"));
        }

        //Excel function to set font color #c00000
        public void ExcelSetFontColorRed(ExcelWorksheet ws, string range)
        {
            ws.Cells[range].Style.Font.Color.SetColor(ColorTranslator.FromHtml("#c00000"));
        }

        //Excel function to set font color #375623
        public void ExcelSetFontColorGreen(ExcelWorksheet ws, string range)
        {
            ws.Cells[range].Style.Font.Color.SetColor(ColorTranslator.FromHtml("#375623"));
        }

        public void ExcelSetFontColorYellow(ExcelWorksheet ws, string range)
        {
            ws.Cells[range].Style.Font.Color.SetColor(ColorTranslator.FromHtml("#fffb21"));
        }

        //Excel function to set font bold
        public void ExcelSetFontBold(ExcelWorksheet ws, string range)
        {
            ws.Cells[range].Style.Font.Bold = true;
        }

        //Excel function to set wraptext true
        public void ExcelWrapText(ExcelWorksheet ws, string range)
        {
            ws.Cells[range].Style.WrapText = true;
        }

        public void ExcelDailyWeeklyMonty(ExcelWorksheet ws, SampleSelection sampleSelection)
        {
            #region Version 1
            int row = 14;
            int colTestRounds = 1;
            int colA2Q2Samples = 2;
            int colDate = 3;
            int colWeekOnly = 4;
            int colStatus = 5;
            int colComment = 6;

            ws.Cells[row, colTestRounds].Value = "Testing Round";
            ws.Cells[row, colA2Q2Samples].Value = "A2Q2 Samples";
            ws.Cells[row, colDate].Value = "Date";
            ws.Cells[row, colWeekOnly].Value = "Weekly Only";
            ws.Cells[row, colStatus].Value = "Status";
            ws.Cells[row, colComment].Value = "A2Q2 Comments";
            ws.Cells["F" + row + ":H" + row].Merge = true;
            ws.Cells["A" + row + ":H" + row].Style.WrapText = true;

            ws.Cells["A" + row + ":H" + row].Style.Border.Top.Style = ExcelBorderStyle.Thin;
            ws.Cells["A" + row + ":H" + row].Style.Border.Left.Style = ExcelBorderStyle.Thin;
            ws.Cells["A" + row + ":H" + row].Style.Border.Right.Style = ExcelBorderStyle.Thin;
            ws.Cells["A" + row + ":H" + row].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
            ws.Cells["A" + row + ":H" + row].Style.Font.SetFromFont(new Font("Arial", 12));
            ws.Cells["A" + row + ":H" + row].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            ws.Cells["A" + row + ":H" + row].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
            ws.Cells["A" + row + ":D" + row].Style.Fill.PatternType = ExcelFillStyle.Solid;
            ws.Cells["A" + row + ":D" + row].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#d9e1f2"));
            ws.Cells["E" + row + ":H" + row].Style.Fill.PatternType = ExcelFillStyle.Solid;
            ws.Cells["E" + row + ":H" + row].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#ccffcc"));
            ws.Cells["D" + row + ":H" + row].Style.Font.Color.SetColor(ColorTranslator.FromHtml("#c00000"));
            ws.Cells["A" + row + ":H" + row].Style.Font.Bold = true;
            ws.Row(row).Height = 42;

            row++;
            foreach (var item in sampleSelection.ListTestRound)
            {
                ws.Cells[row, colTestRounds].Value = item.TestingRound;
                ws.Cells[row, colA2Q2Samples].Value = item.A2Q2Samples;
                if (sampleSelection.Version != "3")
                {
                    ws.Cells[row, colDate].Value = (sampleSelection.Frequency == "Monthly") ? item.MonthOnly : item.Date.Value.DateTime.ToString("MM/dd/yyyy");
                }
                else
                {
                    ws.Cells[row, colDate].Value = item.ContentDisplay1;
                }

                ws.Cells[row, colWeekOnly].Value = (sampleSelection.Frequency == "Weekly") ? item.WeeklyOnly : string.Empty;
                ws.Cells[row, colStatus].Value = item.Status;
                ws.Cells[row, colComment].Value = item.Comment;

                #region Format
                ws.Cells["F" + row + ":H" + row].Merge = true;
                ws.Cells["A" + row + ":H" + row].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                ws.Cells["A" + row + ":H" + row].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                ws.Cells["A" + row + ":H" + row].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                ws.Cells["A" + row + ":H" + row].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                ws.Cells["A" + row + ":H" + row].Style.Font.SetFromFont(new Font("Arial", 12));
                ws.Cells["A" + row + ":H" + row].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                ws.Cells["A" + row + ":H" + row].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                ws.Cells["A" + row + ":D" + row].Style.Fill.PatternType = ExcelFillStyle.Solid;
                ws.Cells["A" + row + ":D" + row].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#d9e1f2"));
                ws.Cells["E" + row + ":H" + row].Style.Fill.PatternType = ExcelFillStyle.Solid;
                ws.Cells["E" + row + ":H" + row].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#ccffcc"));
                #endregion

                row++;
            }
            #endregion

        }

        public void ExcelTransactional(ExcelWorksheet ws, SampleSelection sampleSelection)
        {
            #region Round 1
            int row = 15;

            List<int> listColumn = new List<int>();
            List<TestRound> listTestRound = new List<TestRound>();
            string[] header1 = new string[] { string.Empty, string.Empty, string.Empty, string.Empty, string.Empty };
            int column = 1;
            listColumn.Add(column);
            column++;
            listColumn.Add(column);

            if (sampleSelection.ListTestRound != null && sampleSelection.ListTestRound.Any(item => item.TestingRound == "Round 1"))
            {
                listTestRound.AddRange(sampleSelection.ListTestRound);

                column++;
                header1[0] = listTestRound.Where(x => x.TestingRound == "Round 1").Select(x => x.HeaderRoundDisplay1).FirstOrDefault();
                if (header1[0] != string.Empty && header1[0] != null)
                {
                    listColumn.Add(column);
                }


                column++;
                header1[1] = listTestRound.Where(x => x.TestingRound == "Round 1").Select(x => x.HeaderRoundDisplay2).FirstOrDefault();
                if (header1[1] != string.Empty && header1[1] != null)
                {
                    listColumn.Add(column);
                }


                column++;
                header1[2] = listTestRound.Where(x => x.TestingRound == "Round 1").Select(x => x.HeaderRoundDisplay3).FirstOrDefault();
                if (header1[2] != string.Empty && header1[2] != null)
                {
                    listColumn.Add(column);
                }


                column++;
                header1[3] = listTestRound.Where(x => x.TestingRound == "Round 1").Select(x => x.HeaderRoundDisplay4).FirstOrDefault();
                if (header1[3] != string.Empty && header1[3] != null)
                {
                    listColumn.Add(column);
                }

                column++;
                header1[4] = listTestRound.Where(x => x.TestingRound == "Round 1").Select(x => x.HeaderRoundDisplay5).FirstOrDefault();
                if (header1[4] != string.Empty && header1[4] != null)
                {
                    listColumn.Add(column);
                }


                column++; //status
                listColumn.Add(column);
                column++; //for a2q2 comments
                listColumn.Add(column);

                //Console.WriteLine($"listColumn.Count: {listColumn.Count}");

                ws.Cells[row, listColumn[0]].Value = "Testing Round";
                ws.Cells[row, listColumn[1]].Value = "A2Q2 Samples";

                int columnHeaderStart = 3;
                //int countHeader = 0;
                if (listColumn.Count - 4 > 0)
                {
                    for (int i = 0; i < listColumn.Count - 4; i++)
                    {
                        if (header1[i] != string.Empty)
                        {
                            //countHeader++;
                            ws.Cells[row, columnHeaderStart].Value = header1[i];
                            columnHeaderStart++;
                        }
                    }
                }

                ws.Cells[row, listColumn[listColumn.Count - 2]].Value = "Status";
                ws.Cells[row, listColumn[listColumn.Count - 1]].Value = "A2Q2 Comments";
                ws.Cells[row, listColumn[listColumn.Count - 1], row, listColumn[listColumn.Count - 1] + 2].Merge = true;
                ws.Cells[row, 1, row, listColumn[listColumn.Count - 1] + 2].Style.WrapText = true;
                ws.Cells[row, 1, row, listColumn[listColumn.Count - 1] + 2].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                ws.Cells[row, 1, row, listColumn[listColumn.Count - 1] + 2].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                ws.Cells[row, 1, row, listColumn[listColumn.Count - 1] + 2].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                ws.Cells[row, 1, row, listColumn[listColumn.Count - 1] + 2].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                ws.Cells[row, 1, row, listColumn[listColumn.Count - 1] + 2].Style.Font.SetFromFont(new Font("Arial", 12));
                ws.Cells[row, 1, row, listColumn[listColumn.Count - 1] + 2].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                ws.Cells[row, 1, row, listColumn[listColumn.Count - 1] + 2].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                ws.Cells[row, 1, row, listColumn[listColumn.Count - 1] + 2].Style.Fill.PatternType = ExcelFillStyle.Solid;
                ws.Cells[row, 1, row, listColumn[listColumn.Count - 1] + 2].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#d9e1f2"));
                ws.Cells[row, listColumn[listColumn.Count - 2], row, listColumn[listColumn.Count - 1] + 2].Style.Fill.PatternType = ExcelFillStyle.Solid;
                ws.Cells[row, listColumn[listColumn.Count - 2], row, listColumn[listColumn.Count - 1] + 2].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#ccffcc"));
                ws.Cells[row, 1, row, listColumn[listColumn.Count - 1]].Style.Font.Bold = true;
                ws.Row(row).Height = 42;

                row++;
                foreach (var item in sampleSelection.ListTestRound)
                {
                    if (item.TestingRound == "Round 1")
                    {
                        int columnIndex = 0;
                        ws.Cells[row, listColumn[columnIndex]].Value = item.TestingRound;
                        columnIndex++;
                        ws.Cells[row, listColumn[columnIndex]].Value = item.A2Q2Samples;

                        columnIndex++;
                        if (item.ContentDisplay1 != string.Empty && item.ContentDisplay1 != null)
                        {
                            ws.Cells[row, listColumn[columnIndex]].Value = item.ContentDisplay1;
                        }

                        columnIndex++;
                        if (item.ContentDisplay2 != string.Empty && item.ContentDisplay2 != null)
                        {
                            ws.Cells[row, listColumn[columnIndex]].Value = item.ContentDisplay2;
                        }

                        columnIndex++;
                        if (item.ContentDisplay3 != string.Empty && item.ContentDisplay3 != null)
                        {
                            ws.Cells[row, listColumn[columnIndex]].Value = item.ContentDisplay3;
                        }

                        columnIndex++;
                        if (item.ContentDisplay4 != string.Empty && item.ContentDisplay4 != null)
                        {
                            ws.Cells[row, listColumn[columnIndex]].Value = item.ContentDisplay4;
                        }

                        columnIndex++;
                        if (item.ContentDisplay5 != string.Empty && item.ContentDisplay5 != null)
                        {
                            ws.Cells[row, listColumn[columnIndex]].Value = item.ContentDisplay5;
                        }

                        columnIndex++;
                        ws.Cells[row, columnIndex + 1].Value = item.Status;
                        columnIndex++;
                        ws.Cells[row, columnIndex + 1].Value = item.Comment;

                        #region Format
                        ws.Cells[row, listColumn[listColumn.Count - 1], row, listColumn[listColumn.Count - 1] + 2].Merge = true;
                        ws.Cells[row, 1, row, listColumn[listColumn.Count - 1] + 2].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                        ws.Cells[row, 1, row, listColumn[listColumn.Count - 1] + 2].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                        ws.Cells[row, 1, row, listColumn[listColumn.Count - 1] + 2].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                        ws.Cells[row, 1, row, listColumn[listColumn.Count - 1] + 2].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                        ws.Cells[row, 1, row, listColumn[listColumn.Count - 1] + 2].Style.Font.SetFromFont(new Font("Arial", 12));
                        ws.Cells[row, 1, row, listColumn[listColumn.Count - 1] + 2].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        ws.Cells[row, 1, row, listColumn[listColumn.Count - 1] + 2].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                        ws.Cells[row, 1, row, listColumn[listColumn.Count - 1] + 2].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        ws.Cells[row, 1, row, listColumn[listColumn.Count - 1] + 2].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#d9e1f2"));
                        ws.Cells[row, listColumn[listColumn.Count - 2], row, listColumn[listColumn.Count - 1] + 2].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        ws.Cells[row, listColumn[listColumn.Count - 2], row, listColumn[listColumn.Count - 1] + 2].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#ccffcc"));
                        #endregion

                        row++;
                    }

                }

            }





            #endregion

            #region Round 2
            row += 2;
            listColumn = new List<int>();
            listTestRound = new List<TestRound>();
            header1 = new string[] { string.Empty, string.Empty, string.Empty, string.Empty, string.Empty };
            column = 1;
            listColumn.Add(column);
            column++;
            listColumn.Add(column);

            if (sampleSelection.ListTestRound != null && sampleSelection.ListTestRound.Any(item => item.TestingRound == "Round 2"))
            {
                listTestRound.AddRange(sampleSelection.ListTestRound);

                header1[0] = listTestRound.Where(x => x.TestingRound == "Round 2").Select(x => x.HeaderRoundDisplay1).FirstOrDefault();
                column++;
                if (header1[0] != string.Empty && header1[0] != null)
                {
                    listColumn.Add(column);
                }


                header1[1] = listTestRound.Where(x => x.TestingRound == "Round 2").Select(x => x.HeaderRoundDisplay2).FirstOrDefault();
                column++;
                if (header1[1] != string.Empty && header1[1] != null)
                {
                    listColumn.Add(column);
                }


                header1[2] = listTestRound.Where(x => x.TestingRound == "Round 2").Select(x => x.HeaderRoundDisplay3).FirstOrDefault();
                column++;
                if (header1[2] != string.Empty && header1[2] != null)
                {
                    listColumn.Add(column);
                }


                header1[3] = listTestRound.Where(x => x.TestingRound == "Round 2").Select(x => x.HeaderRoundDisplay4).FirstOrDefault();
                column++;
                if (header1[3] != string.Empty && header1[3] != null)
                {
                    listColumn.Add(column);
                }


                header1[4] = listTestRound.Where(x => x.TestingRound == "Round 2").Select(x => x.HeaderRoundDisplay5).FirstOrDefault();
                column++;
                if (header1[4] != string.Empty && header1[4] != null)
                {
                    listColumn.Add(column);
                }

                column++; //status
                listColumn.Add(column);
                column++; //for a2q2 comments
                listColumn.Add(column);

                //Console.WriteLine($"listColumn.Count: {listColumn.Count}");

                ws.Cells[row, listColumn[0]].Value = "Testing Round";
                ws.Cells[row, listColumn[1]].Value = "A2Q2 Samples";

                int columnHeaderStart = 3;
                //int countHeader = 0;
                if (listColumn.Count - 4 > 0)
                {
                    for (int i = 0; i < listColumn.Count - 4; i++)
                    {
                        if (header1[i] != string.Empty)
                        {
                            //countHeader++;
                            ws.Cells[row, columnHeaderStart].Value = header1[i];
                            columnHeaderStart++;
                        }
                    }
                }

                ws.Cells[row, listColumn[listColumn.Count - 2]].Value = "Status";
                ws.Cells[row, listColumn[listColumn.Count - 1]].Value = "A2Q2 Comments";
                ws.Cells[row, listColumn[listColumn.Count - 1], row, listColumn[listColumn.Count - 1] + 2].Merge = true;
                ws.Cells[row, 1, row, listColumn[listColumn.Count - 1] + 2].Style.WrapText = true;
                ws.Cells[row, 1, row, listColumn[listColumn.Count - 1] + 2].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                ws.Cells[row, 1, row, listColumn[listColumn.Count - 1] + 2].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                ws.Cells[row, 1, row, listColumn[listColumn.Count - 1] + 2].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                ws.Cells[row, 1, row, listColumn[listColumn.Count - 1] + 2].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                ws.Cells[row, 1, row, listColumn[listColumn.Count - 1] + 2].Style.Font.SetFromFont(new Font("Arial", 12));
                ws.Cells[row, 1, row, listColumn[listColumn.Count - 1] + 2].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                ws.Cells[row, 1, row, listColumn[listColumn.Count - 1] + 2].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                ws.Cells[row, 1, row, listColumn[listColumn.Count - 1] + 2].Style.Fill.PatternType = ExcelFillStyle.Solid;
                ws.Cells[row, 1, row, listColumn[listColumn.Count - 1] + 2].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#d9e1f2"));
                ws.Cells[row, listColumn[listColumn.Count - 2], row, listColumn[listColumn.Count - 1] + 2].Style.Fill.PatternType = ExcelFillStyle.Solid;
                ws.Cells[row, listColumn[listColumn.Count - 2], row, listColumn[listColumn.Count - 1] + 2].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#ccffcc"));
                ws.Cells[row, 1, row, listColumn[listColumn.Count - 1]].Style.Font.Bold = true;
                ws.Row(row).Height = 42;

                row++;
                foreach (var item in sampleSelection.ListTestRound)
                {
                    if (item.TestingRound == "Round 2")
                    {
                        int columnIndex = 0;
                        ws.Cells[row, listColumn[columnIndex]].Value = item.TestingRound;
                        columnIndex++;
                        ws.Cells[row, listColumn[columnIndex]].Value = item.A2Q2Samples;

                        columnIndex++;
                        if (item.ContentDisplay1 != string.Empty && item.ContentDisplay1 != null)
                        {
                            ws.Cells[row, listColumn[columnIndex]].Value = item.ContentDisplay1;
                        }

                        columnIndex++;
                        if (item.ContentDisplay2 != string.Empty && item.ContentDisplay2 != null)
                        {
                            ws.Cells[row, listColumn[columnIndex]].Value = item.ContentDisplay2;
                        }

                        columnIndex++;
                        if (item.ContentDisplay3 != string.Empty && item.ContentDisplay3 != null)
                        {
                            ws.Cells[row, listColumn[columnIndex]].Value = item.ContentDisplay3;
                        }

                        columnIndex++;
                        if (item.ContentDisplay4 != string.Empty && item.ContentDisplay4 != null)
                        {
                            ws.Cells[row, listColumn[columnIndex]].Value = item.ContentDisplay4;
                        }

                        columnIndex++;
                        if (item.ContentDisplay5 != string.Empty && item.ContentDisplay5 != null)
                        {
                            ws.Cells[row, listColumn[columnIndex]].Value = item.ContentDisplay5;
                        }

                        columnIndex++;
                        ws.Cells[row, columnIndex + 1].Value = item.Status;
                        columnIndex++;
                        ws.Cells[row, columnIndex + 1].Value = item.Comment;

                        #region Format
                        ws.Cells[row, listColumn[listColumn.Count - 1], row, listColumn[listColumn.Count - 1] + 2].Merge = true;
                        ws.Cells[row, 1, row, listColumn[listColumn.Count - 1] + 2].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                        ws.Cells[row, 1, row, listColumn[listColumn.Count - 1] + 2].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                        ws.Cells[row, 1, row, listColumn[listColumn.Count - 1] + 2].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                        ws.Cells[row, 1, row, listColumn[listColumn.Count - 1] + 2].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                        ws.Cells[row, 1, row, listColumn[listColumn.Count - 1] + 2].Style.Font.SetFromFont(new Font("Arial", 12));
                        ws.Cells[row, 1, row, listColumn[listColumn.Count - 1] + 2].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        ws.Cells[row, 1, row, listColumn[listColumn.Count - 1] + 2].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                        ws.Cells[row, 1, row, listColumn[listColumn.Count - 1] + 2].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        ws.Cells[row, 1, row, listColumn[listColumn.Count - 1] + 2].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#d9e1f2"));
                        ws.Cells[row, listColumn[listColumn.Count - 2], row, listColumn[listColumn.Count - 1] + 2].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        ws.Cells[row, listColumn[listColumn.Count - 2], row, listColumn[listColumn.Count - 1] + 2].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#ccffcc"));
                        #endregion

                        row++;
                    }

                }
            }

            #endregion

            #region Round 3
            row += 2;
            listColumn = new List<int>();
            listTestRound = new List<TestRound>();
            header1 = new string[] { string.Empty, string.Empty, string.Empty, string.Empty, string.Empty };
            column = 1;
            listColumn.Add(column);
            column++;
            listColumn.Add(column);

            if (sampleSelection.ListTestRound != null && sampleSelection.ListTestRound.Any(item => item.TestingRound == "Round 3"))
            {
                listTestRound.AddRange(sampleSelection.ListTestRound);

                header1[0] = listTestRound.Where(x => x.TestingRound == "Round 3").Select(x => x.HeaderRoundDisplay1).FirstOrDefault();
                column++;
                if (header1[0] != string.Empty && header1[0] != null)
                {
                    listColumn.Add(column);
                }


                header1[1] = listTestRound.Where(x => x.TestingRound == "Round 3").Select(x => x.HeaderRoundDisplay2).FirstOrDefault();
                column++;
                if (header1[1] != string.Empty && header1[1] != null)
                {
                    listColumn.Add(column);
                }


                header1[2] = listTestRound.Where(x => x.TestingRound == "Round 3").Select(x => x.HeaderRoundDisplay3).FirstOrDefault();
                column++;
                if (header1[2] != string.Empty && header1[2] != null)
                {
                    listColumn.Add(column);
                }


                header1[3] = listTestRound.Where(x => x.TestingRound == "Round 3").Select(x => x.HeaderRoundDisplay4).FirstOrDefault();
                column++;
                if (header1[3] != string.Empty && header1[3] != null)
                {
                    listColumn.Add(column);
                }


                header1[4] = listTestRound.Where(x => x.TestingRound == "Round 3").Select(x => x.HeaderRoundDisplay5).FirstOrDefault();
                column++;
                if (header1[4] != string.Empty && header1[4] != null)
                {
                    listColumn.Add(column);
                }

                column++; //status
                listColumn.Add(column);
                column++; //for a2q2 comments
                listColumn.Add(column);

                //Console.WriteLine($"listColumn.Count: {listColumn.Count}");

                ws.Cells[row, listColumn[0]].Value = "Testing Round";
                ws.Cells[row, listColumn[1]].Value = "A2Q2 Samples";

                int columnHeaderStart = 3;
                int countHeader = 0;
                if (listColumn.Count - 4 > 0)
                {
                    for (int i = 0; i < listColumn.Count - 4; i++)
                    {
                        if (header1[i] != string.Empty)
                        {
                            countHeader++;
                            ws.Cells[row, columnHeaderStart].Value = header1[i];
                            columnHeaderStart++;
                        }
                    }
                }

                ws.Cells[row, listColumn[listColumn.Count - 2]].Value = "Status";
                ws.Cells[row, listColumn[listColumn.Count - 1]].Value = "A2Q2 Comments";
                ws.Cells[row, listColumn[listColumn.Count - 1], row, listColumn[listColumn.Count - 1] + 2].Merge = true;
                ws.Cells[row, 1, row, listColumn[listColumn.Count - 1] + 2].Style.WrapText = true;
                ws.Cells[row, 1, row, listColumn[listColumn.Count - 1] + 2].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                ws.Cells[row, 1, row, listColumn[listColumn.Count - 1] + 2].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                ws.Cells[row, 1, row, listColumn[listColumn.Count - 1] + 2].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                ws.Cells[row, 1, row, listColumn[listColumn.Count - 1] + 2].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                ws.Cells[row, 1, row, listColumn[listColumn.Count - 1] + 2].Style.Font.SetFromFont(new Font("Arial", 12));
                ws.Cells[row, 1, row, listColumn[listColumn.Count - 1] + 2].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                ws.Cells[row, 1, row, listColumn[listColumn.Count - 1] + 2].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                ws.Cells[row, 1, row, listColumn[listColumn.Count - 1] + 2].Style.Fill.PatternType = ExcelFillStyle.Solid;
                ws.Cells[row, 1, row, listColumn[listColumn.Count - 1] + 2].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#d9e1f2"));
                ws.Cells[row, listColumn[listColumn.Count - 2], row, listColumn[listColumn.Count - 1] + 2].Style.Fill.PatternType = ExcelFillStyle.Solid;
                ws.Cells[row, listColumn[listColumn.Count - 2], row, listColumn[listColumn.Count - 1] + 2].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#ccffcc"));
                ws.Cells[row, 1, row, listColumn[listColumn.Count - 1]].Style.Font.Bold = true;
                ws.Row(row).Height = 42;

                row++;
                foreach (var item in sampleSelection.ListTestRound)
                {
                    if (item.TestingRound == "Round 3")
                    {
                        int columnIndex = 0;
                        ws.Cells[row, listColumn[columnIndex]].Value = item.TestingRound;
                        columnIndex++;
                        ws.Cells[row, listColumn[columnIndex]].Value = item.A2Q2Samples;

                        columnIndex++;
                        if (item.ContentDisplay1 != string.Empty && item.ContentDisplay1 != null)
                        {
                            ws.Cells[row, listColumn[columnIndex]].Value = item.ContentDisplay1;
                        }

                        columnIndex++;
                        if (item.ContentDisplay2 != string.Empty && item.ContentDisplay2 != null)
                        {
                            ws.Cells[row, listColumn[columnIndex]].Value = item.ContentDisplay2;
                        }

                        columnIndex++;
                        if (item.ContentDisplay3 != string.Empty && item.ContentDisplay3 != null)
                        {
                            ws.Cells[row, listColumn[columnIndex]].Value = item.ContentDisplay3;
                        }

                        columnIndex++;
                        if (item.ContentDisplay4 != string.Empty && item.ContentDisplay4 != null)
                        {
                            ws.Cells[row, listColumn[columnIndex]].Value = item.ContentDisplay4;
                        }

                        columnIndex++;
                        if (item.ContentDisplay5 != string.Empty && item.ContentDisplay5 != null)
                        {
                            ws.Cells[row, listColumn[columnIndex]].Value = item.ContentDisplay5;
                        }

                        columnIndex++;
                        ws.Cells[row, columnIndex + 1].Value = item.Status;
                        columnIndex++;
                        ws.Cells[row, columnIndex + 1].Value = item.Comment;

                        #region Format
                        ws.Cells[row, listColumn[listColumn.Count - 1], row, listColumn[listColumn.Count - 1] + 2].Merge = true;
                        ws.Cells[row, 1, row, listColumn[listColumn.Count - 1] + 2].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                        ws.Cells[row, 1, row, listColumn[listColumn.Count - 1] + 2].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                        ws.Cells[row, 1, row, listColumn[listColumn.Count - 1] + 2].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                        ws.Cells[row, 1, row, listColumn[listColumn.Count - 1] + 2].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                        ws.Cells[row, 1, row, listColumn[listColumn.Count - 1] + 2].Style.Font.SetFromFont(new Font("Arial", 12));
                        ws.Cells[row, 1, row, listColumn[listColumn.Count - 1] + 2].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        ws.Cells[row, 1, row, listColumn[listColumn.Count - 1] + 2].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                        ws.Cells[row, 1, row, listColumn[listColumn.Count - 1] + 2].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        ws.Cells[row, 1, row, listColumn[listColumn.Count - 1] + 2].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#d9e1f2"));
                        ws.Cells[row, listColumn[listColumn.Count - 2], row, listColumn[listColumn.Count - 1] + 2].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        ws.Cells[row, listColumn[listColumn.Count - 2], row, listColumn[listColumn.Count - 1] + 2].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#ccffcc"));
                        #endregion

                        row++;
                    }

                }

            }


            #endregion

        }

        public void ExcelSetBackgroundColorGray(ExcelWorksheet ws, string range)
        {
            ws.Cells[range].Style.Fill.PatternType = ExcelFillStyle.Solid;
            ws.Cells[range].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#d9d9d9"));
        }

        public void ExcelSetBackgroundColorYellowPaper(ExcelWorksheet ws, string range)
        {
            ws.Cells[range].Style.Fill.PatternType = ExcelFillStyle.Solid;
            ws.Cells[range].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#f8F7D1"));
        }

        //BorderAround
        public void ExcelBorderAround(ExcelWorksheet ws, string range)
        {
            ws.Cells[range].Style.Border.BorderAround(ExcelBorderStyle.Thin);
        }

        public void ExcelSetBackgroundColorLBluePaper(ExcelWorksheet ws, string range)
        {
            ws.Cells[range].Style.Fill.PatternType = ExcelFillStyle.Solid;
            ws.Cells[range].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#CCFFFF"));
        }

        public void ExcelSetBackgroundColorGray(ExcelWorksheet ws, int fromRow, int fromCol, int toRow, int toCol)
        {
            ws.Cells[fromRow, fromCol, toRow, toCol].Style.Fill.PatternType = ExcelFillStyle.Solid;
            ws.Cells[fromRow, fromCol, toRow, toCol].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#d9d9d9"));
        }

        public void ExcelSetBackgroundColorRed(ExcelWorksheet ws, int fromRow, int fromCol, int toRow, int toCol)
        {
            ws.Cells[fromRow, fromCol, toRow, toCol].Style.Fill.PatternType = ExcelFillStyle.Solid;
            ws.Cells[fromRow, fromCol, toRow, toCol].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#c00000"));
        }

        public void ExcelSetBackgroundColorDarkRed(ExcelWorksheet ws, int fromRow, int fromCol, int toRow, int toCol)
        {
            ws.Cells[fromRow, fromCol, toRow, toCol].Style.Fill.PatternType = ExcelFillStyle.Solid;
            ws.Cells[fromRow, fromCol, toRow, toCol].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#7A1818"));
        }

        public void ExcelSetBackgroundColorGreen(ExcelWorksheet ws, int fromRow, int fromCol, int toRow, int toCol)
        {
            ws.Cells[fromRow, fromCol, toRow, toCol].Style.Fill.PatternType = ExcelFillStyle.Solid;
            ws.Cells[fromRow, fromCol, toRow, toCol].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#92d050"));
        }

        public void ExcelSetBackgroundColorOrange(ExcelWorksheet ws, int fromRow, int fromCol, int toRow, int toCol)
        {
            ws.Cells[fromRow, fromCol, toRow, toCol].Style.Fill.PatternType = ExcelFillStyle.Solid;
            ws.Cells[fromRow, fromCol, toRow, toCol].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#ffc000"));
        }

        public void ExcelSetBackgroundColorLightBlue(ExcelWorksheet ws, int fromRow, int fromCol, int toRow, int toCol)
        {
            ws.Cells[fromRow, fromCol, toRow, toCol].Style.Fill.PatternType = ExcelFillStyle.Solid;
            ws.Cells[fromRow, fromCol, toRow, toCol].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#8ea9db"));
        }

        public void ExcelSetBackgroundColorYellow2(ExcelWorksheet ws, int fromRow, int fromCol, int toRow, int toCol)
        {
            ws.Cells[fromRow, fromCol, toRow, toCol].Style.Fill.PatternType = ExcelFillStyle.Solid;
            ws.Cells[fromRow, fromCol, toRow, toCol].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#fffb21"));
        }

        public void ExcelSetBackgroundColorLightGray(ExcelWorksheet ws, string range)
        {
            ws.Cells[range].Style.Fill.PatternType = ExcelFillStyle.Solid;
            ws.Cells[range].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#f2f2f2"));
        }

        public void ExcelSetBackgroundColorLightGray(ExcelWorksheet ws, int fromRow, int fromCol, int toRow, int toCol)
        {
            ws.Cells[fromRow, fromCol, toRow, toCol].Style.Fill.PatternType = ExcelFillStyle.Solid;
            ws.Cells[fromRow, fromCol, toRow, toCol].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#f2f2f2"));
        }

        public void ExcelSetBackgroundColorLightRed(ExcelWorksheet ws, int fromRow, int fromCol, int toRow, int toCol)
        {
            ws.Cells[fromRow, fromCol, toRow, toCol].Style.Fill.PatternType = ExcelFillStyle.Solid;
            ws.Cells[fromRow, fromCol, toRow, toCol].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#c65853"));
        }

        //KeyReport format
        public void ExcelSetBackgroundColorDarkBlue(ExcelWorksheet ws, int fromRow, int fromCol, int toRow, int toCol)
        {
            ws.Cells[fromRow, fromCol, toRow, toCol].Style.Fill.PatternType = ExcelFillStyle.Solid;
            ws.Cells[fromRow, fromCol, toRow, toCol].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#00316E"));
        }

        public void ExcelSetBackgroundColorDarkGray(ExcelWorksheet ws, int fromRow, int fromCol, int toRow, int toCol)
        {
            ws.Cells[fromRow, fromCol, toRow, toCol].Style.Fill.PatternType = ExcelFillStyle.Solid;
            ws.Cells[fromRow, fromCol, toRow, toCol].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#536878"));
        }

        public void ExcelSetBackgroundColorBlueGray(ExcelWorksheet ws, int fromRow, int fromCol, int toRow, int toCol)
        {
            ws.Cells[fromRow, fromCol, toRow, toCol].Style.Fill.PatternType = ExcelFillStyle.Solid;
            ws.Cells[fromRow, fromCol, toRow, toCol].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#0F3F4F"));
        }

        public void ExcelSetBackgroundColorAshGray(ExcelWorksheet ws, int fromRow, int fromCol, int toRow, int toCol)
        {
            ws.Cells[fromRow, fromCol, toRow, toCol].Style.Fill.PatternType = ExcelFillStyle.Solid;
            ws.Cells[fromRow, fromCol, toRow, toCol].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#333F4F"));
        }

        public void ExcelSetBackgroundColorMintGreen(ExcelWorksheet ws, int fromRow, int fromCol, int toRow, int toCol)
        {
            ws.Cells[fromRow, fromCol, toRow, toCol].Style.Fill.PatternType = ExcelFillStyle.Solid;
            ws.Cells[fromRow, fromCol, toRow, toCol].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#C6E0B4"));
        }

        public void ExcelSetBackgroundColorLightBlue1(ExcelWorksheet ws, int fromRow, int fromCol, int toRow, int toCol)
        {
            ws.Cells[fromRow, fromCol, toRow, toCol].Style.Fill.PatternType = ExcelFillStyle.Solid;
            ws.Cells[fromRow, fromCol, toRow, toCol].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#b7e2f3"));
        }

        public void ExcelSetBackgroundColorAmber(ExcelWorksheet ws, int fromRow, int fromCol, int toRow, int toCol)
        {
            ws.Cells[fromRow, fromCol, toRow, toCol].Style.Fill.PatternType = ExcelFillStyle.Solid;
            ws.Cells[fromRow, fromCol, toRow, toCol].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#FDE64B"));
        }

        public void ExcelSetBackgroundColorLightGreen(ExcelWorksheet ws, int fromRow, int fromCol, int toRow, int toCol)
        {
            ws.Cells[fromRow, fromCol, toRow, toCol].Style.Fill.PatternType = ExcelFillStyle.Solid;
            ws.Cells[fromRow, fromCol, toRow, toCol].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#BDFFA4"));
        }

        public void ExcelSetBackgroundColorYellow(ExcelWorksheet ws, int fromRow, int fromCol, int toRow, int toCol)
        {
            ws.Cells[fromRow, fromCol, toRow, toCol].Style.Fill.PatternType = ExcelFillStyle.Solid;
            ws.Cells[fromRow, fromCol, toRow, toCol].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#FFFF00"));
        }

        public void ExcelSetBackgroundColorSkinTone(ExcelWorksheet ws, int fromRow, int fromCol, int toRow, int toCol)
        {
            ws.Cells[fromRow, fromCol, toRow, toCol].Style.Fill.PatternType = ExcelFillStyle.Solid;
            ws.Cells[fromRow, fromCol, toRow, toCol].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#FFE0C4"));
        }

        public void ExcelSetBackgroundColorMidGray(ExcelWorksheet ws, int fromRow, int fromCol, int toRow, int toCol)
        {
            ws.Cells[fromRow, fromCol, toRow, toCol].Style.Fill.PatternType = ExcelFillStyle.Solid;
            ws.Cells[fromRow, fromCol, toRow, toCol].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#D9D9D9"));
        }

        //End Key Report


        public void ExcelSetBackgroundColorCustom(ExcelWorksheet ws, int fromRow, int fromCol, int toRow, int toCol, string color)
        {
            ws.Cells[fromRow, fromCol, toRow, toCol].Style.Fill.PatternType = ExcelFillStyle.Solid;
            ws.Cells[fromRow, fromCol, toRow, toCol].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml($"#{color}"));
        }

        public void ExcelSetFontColorCustom(ExcelWorksheet ws, int fromRow, int fromCol, int toRow, int toCol, string color)
        {
            //ws.Cells[fromRow, fromCol, toRow, toCol].Style.Fill.PatternType = ExcelFillStyle.Solid;
            ws.Cells[fromRow, fromCol, toRow, toCol].Style.Font.Color.SetColor(ColorTranslator.FromHtml($"#{color}"));
        }


        public void ExcelSetFontColorRed(ExcelWorksheet ws, int fromRow, int fromCol, int toRow, int toCol)
        {
            //ws.Cells[fromRow, fromCol, toRow, toCol].Style.Fill.PatternType = ExcelFillStyle.Solid;
            ws.Cells[fromRow, fromCol, toRow, toCol].Style.Font.Color.SetColor(ColorTranslator.FromHtml("#ed1500"));
        }

        public void ExcelSetFontColorWhite(ExcelWorksheet ws, int fromRow, int fromCol, int toRow, int toCol)
        {
            //ws.Cells[fromRow, fromCol, toRow, toCol].Style.Fill.PatternType = ExcelFillStyle.Solid;
            ws.Cells[fromRow, fromCol, toRow, toCol].Style.Font.Color.SetColor(ColorTranslator.FromHtml("#ffffff"));
        }

        public void ExcelSetFontBold(ExcelWorksheet ws, int fromRow, int fromCol, int toRow, int toCol)
        {
            ws.Cells[fromRow, fromCol, toRow, toCol].Style.Font.Bold = true;
        }

        public void ExcelSetFontItalic(ExcelWorksheet ws, int fromRow, int fromCol, int toRow, int toCol)
        {
            ws.Cells[fromRow, fromCol, toRow, toCol].Style.Font.Italic = true;
        }

        public void ExcelWrapText(ExcelWorksheet ws, int fromRow, int fromCol, int toRow, int toCol)
        {
            ws.Cells[fromRow, fromCol, toRow, toCol].Style.WrapText = true;
        }

        public void ExcelSetArialSize10(ExcelWorksheet ws, string range)
        {
            ws.Cells[range].Style.Font.SetFromFont(new Font("Arial", 10));
        }

        public void ExcelSetFontSize(ExcelWorksheet ws, int fromRow, int fromCol, int toRow, int toCol, int fontSize)
        {
            ws.Cells[fromRow, fromCol, toRow, toCol].Style.Font.Size = fontSize;
        }

        public void ExcelSetFontName(ExcelWorksheet ws, int fromRow, int fromCol, int toRow, int toCol, string fontname)
        {
            ws.Cells.Style.Font.Name = fontname;
        }

        public void ExcelSetArialSize8(ExcelWorksheet ws, int fromRow, int fromCol, int toRow, int toCol)
        {
            ws.Cells[fromRow, fromCol, toRow, toCol].Style.Font.SetFromFont(new Font("Arial", 8));
        }

        public void ExcelSetFontCustom(ExcelWorksheet ws, int fromRow, int fromCol, int toRow, int toCol, string fontName, int size)
        {
            ws.Cells[fromRow, fromCol, toRow, toCol].Style.Font.SetFromFont(new Font(fontName, size));
        }

        public void ExcelSetArialSize10(ExcelWorksheet ws, int fromRow, int fromCol, int toRow, int toCol)
        {
            ws.Cells[fromRow, fromCol, toRow, toCol].Style.Font.SetFromFont(new Font("Arial", 10));
        }

        public void ExcelSetCalibriLight10(ExcelWorksheet ws, int fromRow, int fromCol, int toRow, int toCol)
        {
            ws.Cells[fromRow, fromCol, toRow, toCol].Style.Font.SetFromFont(new Font("Calibri Light", 10));
        }

        public void ExcelSetArialSize16(ExcelWorksheet ws, int fromRow, int fromCol, int toRow, int toCol)
        {
            ws.Cells[fromRow, fromCol, toRow, toCol].Style.Font.SetFromFont(new Font("Arial", 16));
        }

        public void ExcelSetArialSize12(ExcelWorksheet ws, int fromRow, int fromCol, int toRow, int toCol)
        {
            ws.Cells[fromRow, fromCol, toRow, toCol].Style.Font.SetFromFont(new Font("Arial", 12));
        }

        public void ExcelSetCalibriLight12(ExcelWorksheet ws, int fromRow, int fromCol, int toRow, int toCol)
        {
            ws.Cells[fromRow, fromCol, toRow, toCol].Style.Font.SetFromFont(new Font("Calibri Light", 12));
        }

        public void ExcelSetHorizontalAlignCenter(ExcelWorksheet ws, string range)
        {
            ws.Cells[range].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
        }

        public void ExcelSetHorizontalAlignCenter(ExcelWorksheet ws, int fromRow, int fromCol, int toRow, int toCol)
        {
            ws.Cells[fromRow, fromCol, toRow, toCol].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
        }

        public void ExcelSetHorizontalAlignLeft(ExcelWorksheet ws, string range)
        {
            ws.Cells[range].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
        }

        public void ExcelSetHorizontalAlignLeft(ExcelWorksheet ws, int fromRow, int fromCol, int toRow, int toCol)
        {
            ws.Cells[fromRow, fromCol, toRow, toCol].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
        }

        public void ExcelSetVerticalAlignCenter(ExcelWorksheet ws, string range)
        {
            ws.Cells[range].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
        }

        public void ExcelSetVerticalAlignCenter(ExcelWorksheet ws, int fromRow, int fromCol, int toRow, int toCol)
        {
            ws.Cells[fromRow, fromCol, toRow, toCol].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
        }

        public void ExcelSetVerticalAlignTop(ExcelWorksheet ws, string range)
        {
            ws.Cells[range].Style.VerticalAlignment = ExcelVerticalAlignment.Top;
        }

        public void ExcelSetVerticalAlignTop(ExcelWorksheet ws, int fromRow, int fromCol, int toRow, int toCol)
        {
            ws.Cells[fromRow, fromCol, toRow, toCol].Style.VerticalAlignment = ExcelVerticalAlignment.Top;
        }

        public void ExcelSetBorder(ExcelWorksheet ws, int fromRow, int fromCol, int toRow, int toCol)
        {
            ws.Cells[fromRow, fromCol, toRow, toCol].Style.Border.Top.Style = ExcelBorderStyle.Thin;
            ws.Cells[fromRow, fromCol, toRow, toCol].Style.Border.Left.Style = ExcelBorderStyle.Thin;
            ws.Cells[fromRow, fromCol, toRow, toCol].Style.Border.Right.Style = ExcelBorderStyle.Thin;
            ws.Cells[fromRow, fromCol, toRow, toCol].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
        }

        



        public void ExcelSetBorderRed(ExcelWorksheet ws, int fromRow, int fromCol, int toRow, int toCol)
        {
            ws.Cells[fromRow, fromCol, toRow, toCol].Style.Border.Top.Style = ExcelBorderStyle.Medium;
            ws.Cells[fromRow, fromCol, toRow, toCol].Style.Border.Left.Style = ExcelBorderStyle.Medium;
            ws.Cells[fromRow, fromCol, toRow, toCol].Style.Border.Right.Style = ExcelBorderStyle.Medium;
            ws.Cells[fromRow, fromCol, toRow, toCol].Style.Border.Bottom.Style = ExcelBorderStyle.Medium;
            ws.Cells[fromRow, fromCol, toRow, toCol].Style.Border.Top.Color.SetColor(Color.Red);
            ws.Cells[fromRow, fromCol, toRow, toCol].Style.Border.Left.Color.SetColor(Color.Red);
            ws.Cells[fromRow, fromCol, toRow, toCol].Style.Border.Right.Color.SetColor(Color.Red);
            ws.Cells[fromRow, fromCol, toRow, toCol].Style.Border.Bottom.Color.SetColor(Color.Red);

        }

        public void ExcelSetBorderThick(ExcelWorksheet ws, int fromRow, int fromCol, int toRow, int toCol)
        {
            ws.Cells[fromRow, fromCol, toRow, toCol].Style.Border.Top.Style = ExcelBorderStyle.Thick;
            ws.Cells[fromRow, fromCol, toRow, toCol].Style.Border.Left.Style = ExcelBorderStyle.Thick;
            ws.Cells[fromRow, fromCol, toRow, toCol].Style.Border.Right.Style = ExcelBorderStyle.Thick;
            ws.Cells[fromRow, fromCol, toRow, toCol].Style.Border.Bottom.Style = ExcelBorderStyle.Thick;
        }

        public void XlsSetRow(ExcelWorksheet ws, int row, string value, int maxLength, int defaultRow)
        {
            int counttxt;
            double newline = 1;
            if (value != null && value != string.Empty)
            {
                counttxt = value.Length;

                if (counttxt > maxLength && counttxt != 0)
                {
                    newline = counttxt / maxLength;
                    ws.Row(row).Height = defaultRow * newline;
                }
                else
                    ws.Row(row).Height = defaultRow;
            }
            else
                ws.Row(row).Height = defaultRow;
        }

        public void TestingAttributeFormat(ExcelWorksheet ws, int row, int column, string answer, string note)
        {
            using (ExcelRange Rng = ws.Cells[row, column])
            {
                ExcelRichTextCollection RichTxtCollection = Rng.RichText;
                ExcelRichText RichText;

                if (answer != null && answer != string.Empty)
                {
                    RichText = RichTxtCollection.Add($"{answer} ");
                    RichText.FontName = "Arial";
                    RichText.Size = 10;
                    RichText.Color = Color.Black;
                }

                if (note != null && note != string.Empty)
                {
                    RichText = RichTxtCollection.Add("{" + note + "}");
                    RichText.Color = Color.Red;
                    RichText.FontName = "Arial";
                    RichText.Size = 10;
                    RichText.Bold = true;
                }

            }
        }

        public bool IsDateTime(string dateTime)
        {
            bool result = false;
            DateTime dtValue;
            DateTime.TryParse(dateTime, out dtValue);
            if (dtValue.Year != 0001)
                result = true;

            return result;
        }

        public bool IsDecimal(string decimalValue)
        {
            bool result = false;
            decimal intValue;
            decimal.TryParse(decimalValue, out intValue);
            if (intValue != 0)
                result = true;

            return result;
        }

        public bool IsDateHeader(string decimalValue)
        {
            bool result = false;
            if (decimalValue != null)
            {
                if (decimalValue.Contains("date", StringComparison.OrdinalIgnoreCase))
                    result = true;
            }
            return result;
        }

        public bool IsIdentificationHeader(string decimalValue)
        {
            bool result = false;
            if (decimalValue != null)
            {
                if (decimalValue.Contains("id", StringComparison.OrdinalIgnoreCase))
                    result = true;
            }
            return result;
        }

    }
}
