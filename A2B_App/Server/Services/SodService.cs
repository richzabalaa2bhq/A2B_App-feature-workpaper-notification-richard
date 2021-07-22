using A2B_App.Server.Data;
using A2B_App.Shared.Sox;
using Microsoft.Extensions.Configuration;
using OfficeOpenXml;
using PodioAPI.Models.Request;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Threading.Tasks;

namespace A2B_App.Server.Services
{
    
    public class SodService
    {
        private SoxContext _soxContext;
        private readonly IConfiguration _config;

        public SodService(SoxContext soxContext, IConfiguration config)
        {
            _soxContext = soxContext;
            _config = config;
        }

        /// <summary>
        /// Process SOD File (use for manual process)
        /// </summary>
        /// <param name="fileName"></param>
        /// <param name="fileImport"></param>
        /// <returns></returns>
        public async Task<ServiceResponse> ProcessSodFile(string fileName, FileImport fileImport)
        {
            ServiceResponse response = new ServiceResponse();
            ExcelService xlsService = new ExcelService();
            WriteLog writeLog = new WriteLog();
            try
            {
                List<ConflictDefinition> listConflictDefinition = new List<ConflictDefinition>();
                List<SodUser> listSodUser = new List<SodUser>();
                string startupPath = Directory.GetCurrentDirectory();
                string path = Path.Combine(startupPath, "include", "upload", fileName);

                FileInfo fi = new FileInfo(path);
                using (ExcelPackage p = new ExcelPackage(fi))
                {
                    // If you use EPPlus in a noncommercial context
                    // according to the Polyform Noncommercial license:
                    ////ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

                    #region Process Definition of Conflict
                    ExcelWorksheet wsDefConflict = p.Workbook.Worksheets["1. Definition of Conflict"];
                    int colCountSheet1 = wsDefConflict.Dimension.End.Column;
                    int rowCountSheet1 = wsDefConflict.Dimension.End.Row;
                    for (int row = 4; row <= rowCountSheet1; row++)
                    {
                        ConflictDefinition conflictDef = new ConflictDefinition();
                        for (int col = 1; col <= colCountSheet1; col++)
                        {
                            switch (col)
                            {
                                case 1:
                                    conflictDef.Duty1 = wsDefConflict.Cells[row, col].Value?.ToString();
                                    break;
                                case 2:
                                    conflictDef.Duty2 = wsDefConflict.Cells[row, col].Value?.ToString();
                                    break;
                                case 3:
                                    conflictDef.RiskLevel = wsDefConflict.Cells[row, col].Value?.ToString();
                                    break;
                                case 4:
                                    conflictDef.Definition = wsDefConflict.Cells[row, col].Value?.ToString();
                                    break;
                                default:
                                    break;
                            }
                        }
                        listConflictDefinition.Add(conflictDef);
                    }
                    #endregion

                    #region Process Definition of Conflict
                    ExcelWorksheet wsUser = p.Workbook.Worksheets["2. Users"];
                    int colCountSheet2 = wsUser.Dimension.End.Column;
                    int rowCountSheet2 = wsUser.Dimension.End.Row;
                    string lastProcess = string.Empty;
                    for (int row = 2; row <= rowCountSheet2; row++)
                    {
                        SodUser sodUser = new SodUser();
                        for (int col = 1; col <= colCountSheet2; col++)
                        {
                            switch (col)
                            {
                                case 1:
                                    sodUser.OwnersForManual = wsUser.Cells[row, col].Value?.ToString();
                                    break;
                                case 2:

                                    if (wsUser.Cells[row, col].Value != null)
                                    {
                                        sodUser.Process = wsUser.Cells[row, col].Value?.ToString();
                                        lastProcess = sodUser.Process;
                                    }
                                    else
                                        sodUser.Process = lastProcess;
                                    break;
                                case 3:
                                    sodUser.DutyNum = wsUser.Cells[row, col].Value?.ToString();
                                    break;
                                case 4:
                                    sodUser.Function = wsUser.Cells[row, col].Value?.ToString();
                                    break;
                                case 5:
                                    sodUser.Users = wsUser.Cells[row, col].Value?.ToString();
                                    break;
                                default:
                                    break;
                            }
                        }
                        listSodUser.Add(sodUser);
                    }
                    #endregion

                    #region Tab Duties
                    ExcelWorksheet wsDuties = p.Workbook.Worksheets["Duties"];
                    //disable grid
                    wsDuties.View.ShowGridLines = false;
                    int rowDuties = 2;
                    if (listSodUser.Count > 0)
                    {
                        foreach (var item in listSodUser)
                        {
                            wsDuties.Cells[rowDuties, 1].Value = item.DutyNum;
                            wsDuties.Cells[rowDuties, 2].Value = item.Process;
                            wsDuties.Cells[rowDuties, 3].Value = item.Function;
                            //xlsService.ExcelWrapText(wsDuties, rowDuties, 3, rowDuties, 3);
                            xlsService.ExcelSetVerticalAlignCenter(wsDuties, rowDuties, 1, rowDuties, 1);
                            xlsService.ExcelSetHorizontalAlignCenter(wsDuties, rowDuties, 1, rowDuties, 1);
                            xlsService.ExcelSetBorder(wsDuties, rowDuties, 1, rowDuties, 3);
                            rowDuties++;
                        }
                    }


                    #endregion

                    #region Tab 3.Conflict
                    ExcelWorksheet wsConflict = p.Workbook.Worksheets["3.Conflict"];
                    //disable grid
                    wsConflict.View.ShowGridLines = false;
                    int rowConflict = 2;
                    string conflictUsers = string.Empty;
                    
                    if (listConflictDefinition.Count > 0 && listSodUser.Count > 0)
                    {

                        foreach (var item in listConflictDefinition)
                        {
                            var allDuty1 = listSodUser.Where(x => x.DutyNum.Equals(item.Duty1)).FirstOrDefault();
                            var allDuty2 = listSodUser.Where(x => x.DutyNum.Equals(item.Duty2)).FirstOrDefault();
                            List<string> listUser1 = new List<string>();
                            List<string> listUser2 = new List<string>();
                            List<string> listConflict = new List<string>();
                            conflictUsers = string.Empty;

                            if (allDuty1.Users != string.Empty)
                            {
                                listUser1 = allDuty1.Users.Split(",").Select(x => x.TrimStart(' ')).ToList();
                            }

                            if (allDuty2.Users != string.Empty)
                            {
                                listUser2 = allDuty2.Users.Split(",").Select(x => x.TrimStart(' ')).ToList();
                            }

                            if(listUser1.Count > 0 && listUser2.Count > 0)
                            {
                                listConflict = listUser1.Intersect(listUser2).ToList();
                                conflictUsers = string.Join(",", listConflict);
                            }

                            wsConflict.Cells[rowConflict, 1].Value = item.Duty1;
                            wsConflict.Cells[rowConflict, 2].Value = allDuty1 != null ? allDuty1.Users : string.Empty;
                            wsConflict.Cells[rowConflict, 3].Value = item.Duty2;
                            wsConflict.Cells[rowConflict, 4].Value = allDuty2 != null ? allDuty2.Users : string.Empty;
                            wsConflict.Cells[rowConflict, 5].Value = conflictUsers;
                            xlsService.ExcelWrapText(wsConflict, rowConflict, 1, rowConflict, 5);
                            xlsService.ExcelSetBorder(wsConflict, rowConflict, 1, rowConflict, 5);
                            xlsService.ExcelSetVerticalAlignCenter(wsConflict, rowConflict, 1, rowConflict, 5);
                            xlsService.ExcelSetHorizontalAlignCenter(wsConflict, rowConflict, 1, rowConflict, 1);
                            xlsService.ExcelSetHorizontalAlignCenter(wsConflict, rowConflict, 3, rowConflict, 3);
                            if (conflictUsers != string.Empty)
                                xlsService.ExcelSetBackgroundColorYellow2(wsConflict, rowConflict, 5, rowConflict, 5);
                            rowConflict++;
                        }
                    }

                    #endregion

                    #region Tab 4. SOD Matrix

                    ExcelWorksheet wsMatrix = p.Workbook.Worksheets["4. SOD Matrix"];
                    //disable grid
                    wsMatrix.View.ShowGridLines = false;
                    int rowMatrix = 1;
                    string conflictMatrix = string.Empty;

                    
                    wsMatrix.Row(rowMatrix).Height = 30;
                    wsMatrix.Column(1).Width = 5;
                    wsMatrix.Column(2).Width = 5;
                    wsMatrix.Column(3).Width = 25;
                    wsMatrix.Column(4).Width = 43;

                    //row 1 client name
                    wsMatrix.Cells[rowMatrix, 1].Value = $"{fileImport.ClientName} Segregation of Duties Matrix";
                    xlsService.ExcelSetArialSize8(wsMatrix, rowMatrix, 1, rowMatrix, 1); //set arial 8
                    xlsService.ExcelSetFontBold(wsMatrix, rowMatrix, 1, rowMatrix, 1);
                    rowMatrix++;


                    //get unique process
                    List<string> uniqueProcess = new List<string>();
                    foreach (var item in listSodUser)
                    {
                        if (!uniqueProcess.Contains(item.Process))
                            uniqueProcess.Add(item.Process);
                    }

                    //loop through unique process and set header
                    int tempColStart = 5;
                    int tempColEnd = 0;
                    foreach (var item in uniqueProcess)
                    {
                        int totalCount = listSodUser.Where(x => x.Process.Equals(item)).Count();
                        wsMatrix.Cells[rowMatrix, tempColStart].Value = item;
                        tempColEnd = tempColStart + totalCount - 1;
                        wsMatrix.Cells[rowMatrix, tempColStart, rowMatrix, tempColEnd].Merge = true; //Merge header
                        xlsService.ExcelSetBorderThick(wsMatrix, rowMatrix, tempColStart, rowMatrix, tempColEnd); //set border thick
                        xlsService.ExcelSetVerticalAlignCenter(wsMatrix, rowMatrix, tempColStart, rowMatrix, tempColEnd);
                        xlsService.ExcelSetHorizontalAlignCenter(wsMatrix, rowMatrix, tempColStart, rowMatrix, tempColEnd);
                        xlsService.ExcelWrapText(wsMatrix, rowMatrix, tempColStart, rowMatrix, tempColEnd); //wrap text
                        xlsService.ExcelSetArialSize8(wsMatrix, rowMatrix, tempColStart, rowMatrix, tempColEnd); //set arial 8
                        tempColStart = tempColEnd + 1;
                    }

                    //header function
                    rowMatrix++;
                    wsMatrix.Row(rowMatrix).Height = 120;

                    string legend = "Legend:" + Environment.NewLine +
                        "X(no color) = No conflicting roles" + Environment.NewLine +
                        "Y(yellow) = Conflict with mitigating controls(Conflicts identified with formula; all users have conflicts)";

                    wsMatrix.Cells[rowMatrix, 1].Value = legend;
                    wsMatrix.Cells[rowMatrix, 1, rowMatrix, 4].Merge = true; //Merge 
                    xlsService.ExcelWrapText(wsMatrix, rowMatrix, 1, rowMatrix, 4); //wrap text
                    xlsService.ExcelSetVerticalAlignCenter(wsMatrix, rowMatrix, 1, rowMatrix, 4); //vertical align center
                    xlsService.ExcelSetFontBold(wsMatrix, rowMatrix, 1, rowMatrix, 4); //vertical align center
                    xlsService.ExcelSetArialSize8(wsMatrix, rowMatrix, 1, rowMatrix, 4); //set arial 8
                    xlsService.ExcelSetFontBold(wsMatrix, rowMatrix, 1, rowMatrix, 4); //set font bold

                    tempColStart = 5;
                    foreach (var item in listSodUser)
                    {
                        wsMatrix.Cells[rowMatrix, tempColStart].Value = item.Function;
                        xlsService.ExcelSetBorderThick(wsMatrix, rowMatrix, tempColStart, rowMatrix, tempColStart); //set border thick
                        wsMatrix.Cells[rowMatrix, tempColStart].Style.TextRotation = 90;
                        xlsService.ExcelSetHorizontalAlignCenter(wsMatrix, rowMatrix, tempColStart, rowMatrix, tempColStart);
                        xlsService.ExcelWrapText(wsMatrix, rowMatrix, tempColStart, rowMatrix, tempColStart); //wrap text
                        xlsService.ExcelSetArialSize8(wsMatrix, rowMatrix, tempColStart, rowMatrix, tempColStart); //set arial 8
                        tempColStart++;
                    }

                    //header user
                    rowMatrix++;
                    wsMatrix.Row(rowMatrix).Height = 45;
                    tempColStart = 5;
                    foreach (var item in listSodUser)
                    {
                        wsMatrix.Cells[rowMatrix, tempColStart].Value = item.Users;
                        xlsService.ExcelSetBorderThick(wsMatrix, rowMatrix, tempColStart, rowMatrix, tempColStart); //set border thick
                        wsMatrix.Cells[rowMatrix, tempColStart].Style.TextRotation = 90;
                        xlsService.ExcelSetHorizontalAlignCenter(wsMatrix, rowMatrix, tempColStart, rowMatrix, tempColStart);
                        xlsService.ExcelWrapText(wsMatrix, rowMatrix, tempColStart, rowMatrix, tempColStart); //wrap text
                        xlsService.ExcelSetArialSize8(wsMatrix, rowMatrix, tempColStart, rowMatrix, tempColStart); //set arial 8
                        tempColStart++;
                    }

                    //description
                    rowMatrix++;
                    wsMatrix.Cells[rowMatrix, 2].Value = "Duty";
                    wsMatrix.Cells[rowMatrix, 3].Value = "Description";
                    wsMatrix.Cells[rowMatrix, 4].Value = "Name (Highest Level of Authority)";
                    xlsService.ExcelSetHorizontalAlignCenter(wsMatrix, rowMatrix, 2, rowMatrix, 4);
                    xlsService.ExcelSetVerticalAlignCenter(wsMatrix, rowMatrix, 2, rowMatrix, 4);
                    xlsService.ExcelWrapText(wsMatrix, rowMatrix, 2, rowMatrix, 4); //wrap text
                    xlsService.ExcelSetFontBold(wsMatrix, rowMatrix, 2, rowMatrix, 4);
                    xlsService.ExcelSetBackgroundColorCustom(wsMatrix, rowMatrix, 2, rowMatrix, 4, "c0c0c0");
                    xlsService.ExcelSetFontBold(wsMatrix, rowMatrix, 2, rowMatrix, 4);
                    xlsService.ExcelSetArialSize8(wsMatrix, rowMatrix, 2, rowMatrix, 4); //set arial 8

                    tempColStart = 5;
                    foreach (var item in listSodUser)
                    {
                        wsMatrix.Cells[rowMatrix, tempColStart].Value = item.DutyNum;
                        xlsService.ExcelSetBackgroundColorCustom(wsMatrix, rowMatrix, tempColStart, rowMatrix, tempColStart, "c0c0c0");
                        xlsService.ExcelSetFontBold(wsMatrix, rowMatrix, tempColStart, rowMatrix, tempColStart);
                        xlsService.ExcelSetArialSize8(wsMatrix, rowMatrix, tempColStart, rowMatrix, tempColStart); //set arial 8
                        xlsService.ExcelSetHorizontalAlignCenter(wsMatrix, rowMatrix, tempColStart, rowMatrix, tempColStart);
                        xlsService.ExcelSetVerticalAlignCenter(wsMatrix, rowMatrix, tempColStart, rowMatrix, tempColStart);
                        xlsService.ExcelSetFontBold(wsMatrix, rowMatrix, tempColStart, rowMatrix, tempColStart);
                        tempColStart++;
                    }

                    //items
                    
                    foreach (var item in listSodUser)
                    {
                        
                        rowMatrix++;
                        wsMatrix.Row(rowMatrix).Height = 15;
                        tempColStart = 1;

                        wsMatrix.Cells[rowMatrix, tempColStart].Value = item.Process;
                        tempColStart++;

                        wsMatrix.Cells[rowMatrix, tempColStart].Value = item.DutyNum;
                        xlsService.ExcelSetHorizontalAlignCenter(wsMatrix, rowMatrix, tempColStart, rowMatrix, tempColStart);
                        xlsService.ExcelSetVerticalAlignCenter(wsMatrix, rowMatrix, tempColStart, rowMatrix, tempColStart);
                        xlsService.ExcelSetFontBold(wsMatrix, rowMatrix, tempColStart, rowMatrix, tempColStart);
                        xlsService.ExcelSetBackgroundColorCustom(wsMatrix, rowMatrix, tempColStart, rowMatrix, tempColStart, "c0c0c0");

                        tempColStart++;
                        wsMatrix.Cells[rowMatrix, tempColStart].Value = item.Function;
                        xlsService.ExcelWrapText(wsMatrix, rowMatrix, tempColStart, rowMatrix, tempColStart); //wrap text


                        tempColStart++;
                        wsMatrix.Cells[rowMatrix, tempColStart].Value = item.Users;
                        xlsService.ExcelWrapText(wsMatrix, rowMatrix, tempColStart, rowMatrix, tempColStart); //wrap text

                        xlsService.ExcelSetArialSize8(wsMatrix, rowMatrix, 1, rowMatrix, tempColStart); //set arial 8

                        tempColStart++;
                        foreach (var itemInner in listSodUser)
                        {
                            //check if it is included in the definition of conflict
                            var checkDefConflict = listConflictDefinition.Where(x => 
                                (x.Duty1.Equals(item.DutyNum) && x.Duty2.Equals(itemInner.DutyNum)) ||
                                (x.Duty2.Equals(item.DutyNum) && x.Duty1.Equals(itemInner.DutyNum))
                            ).FirstOrDefault();
                            if(checkDefConflict != null)
                            {
                                //check if it has conflict, Y - yes, X - no
                                var conflictRes = await CheckConflict(listSodUser, item.DutyNum, itemInner.DutyNum);
                                if(conflictRes.Status)
                                {
                                    wsMatrix.Cells[rowMatrix, tempColStart].Value = "Y";
                                    if (int.Parse(item.DutyNum) < int.Parse(itemInner.DutyNum))
                                    {
                                        xlsService.ExcelSetBackgroundColorCustom(wsMatrix, rowMatrix, tempColStart, rowMatrix, tempColStart, "858203");
                                        xlsService.ExcelSetFontColorCustom(wsMatrix, rowMatrix, tempColStart, rowMatrix, tempColStart, "ffffff");
                                    }
                                    else if (int.Parse(item.DutyNum) > int.Parse(itemInner.DutyNum))
                                    {
                                        xlsService.ExcelSetBackgroundColorCustom(wsMatrix, rowMatrix, tempColStart, rowMatrix, tempColStart, "fffb21");
                                    }

                                }
                                else
                                {
                                    wsMatrix.Cells[rowMatrix, tempColStart].Value = "X";
                                    xlsService.ExcelSetBackgroundColorCustom(wsMatrix, rowMatrix, tempColStart, rowMatrix, tempColStart, "ffffff");
                                }

                                xlsService.ExcelSetHorizontalAlignCenter(wsMatrix, rowMatrix, tempColStart, rowMatrix, tempColStart);
                                xlsService.ExcelSetVerticalAlignCenter(wsMatrix, rowMatrix, tempColStart, rowMatrix, tempColStart);
                                xlsService.ExcelSetFontBold(wsMatrix, rowMatrix, tempColStart, rowMatrix, tempColStart);
                            }

                            else
                            {
                                if (item.DutyNum.Equals(itemInner.DutyNum))
                                {
                                    xlsService.ExcelSetBackgroundColorCustom(wsMatrix, rowMatrix, tempColStart, rowMatrix, tempColStart, "000000");
                                }
                                else if (int.Parse(item.DutyNum) < int.Parse(itemInner.DutyNum))
                                {
                                    xlsService.ExcelSetBackgroundColorCustom(wsMatrix, rowMatrix, tempColStart, rowMatrix, tempColStart, "757578");
                                }
                                else if (int.Parse(item.DutyNum) > int.Parse(itemInner.DutyNum))
                                {
                                    xlsService.ExcelSetBackgroundColorCustom(wsMatrix, rowMatrix, tempColStart, rowMatrix, tempColStart, "bbbbbf");
                                }
                            }

                            xlsService.ExcelSetFontCustom(wsMatrix, rowMatrix, tempColStart, rowMatrix, tempColStart, "Arial", 6);
                            xlsService.ExcelSetFontBold(wsMatrix, rowMatrix, tempColStart, rowMatrix, tempColStart);
                            tempColStart++;
                        }

                        
                        xlsService.ExcelSetBorder(wsMatrix, rowMatrix, 1, rowMatrix, tempColStart - 1); //set border
                    }



                    #endregion

                    #region Tab 5. SOD Analysis

                    ExcelWorksheet wsConflictAnalysis = p.Workbook.Worksheets["5. SOD Analysis"];
                    //disable grid
                    wsConflictAnalysis.View.ShowGridLines = false;
                    int rowSodConflict = 2;
                    string conflictSodUsers = string.Empty;

                    if (listConflictDefinition.Count > 0 && listSodUser.Count > 0)
                    {

                        foreach (var item in listConflictDefinition)
                        {
                            var allDuty1 = listSodUser.Where(x => x.DutyNum.Equals(item.Duty1)).FirstOrDefault();
                            var allDuty2 = listSodUser.Where(x => x.DutyNum.Equals(item.Duty2)).FirstOrDefault();
                            List<string> listUser1 = new List<string>();
                            List<string> listUser2 = new List<string>();
                            List<string> listConflict = new List<string>();
                            conflictSodUsers = string.Empty;

                            if (allDuty1.Users != string.Empty)
                            {
                                listUser1 = allDuty1.Users.Split(",").Select(x => x.TrimStart(' ')).ToList();
                            }

                            if (allDuty2.Users != string.Empty)
                            {
                                listUser2 = allDuty2.Users.Split(",").Select(x => x.TrimStart(' ')).ToList();
                            }

                            if (listUser1.Count > 0 && listUser2.Count > 0)
                            {
                                listConflict = listUser1.Intersect(listUser2).ToList();
                                conflictSodUsers = string.Join(",", listConflict);
                            }

                            wsConflictAnalysis.Cells[rowSodConflict, 1].Value = conflictSodUsers != string.Empty ? "Y" : string.Empty;
                            wsConflictAnalysis.Cells[rowSodConflict, 2].Value = item.Duty1;
                            wsConflictAnalysis.Cells[rowSodConflict, 3].Value = item.Duty2;
                            wsConflictAnalysis.Cells[rowSodConflict, 4].Value = item.Definition;
                            wsConflictAnalysis.Cells[rowSodConflict, 7].Value = conflictSodUsers;
                            xlsService.ExcelWrapText(wsConflictAnalysis, rowSodConflict, 1, rowSodConflict, 11);
                            xlsService.ExcelSetBorder(wsConflictAnalysis, rowSodConflict, 1, rowSodConflict, 11);
                            xlsService.ExcelSetVerticalAlignCenter(wsConflictAnalysis, rowSodConflict, 1, rowSodConflict, 11);
                            xlsService.ExcelSetHorizontalAlignCenter(wsConflictAnalysis, rowSodConflict, 1, rowSodConflict, 11);
                            if (conflictSodUsers != string.Empty)
                            {
                                xlsService.ExcelSetBackgroundColorYellow2(wsConflictAnalysis, rowSodConflict, 1, rowSodConflict, 4);
                                xlsService.ExcelSetBackgroundColorYellow2(wsConflictAnalysis, rowSodConflict, 7, rowSodConflict, 7);
                            }
                                
                            rowSodConflict++;
                        }
                    }

                    #endregion



                    string strSourceDownload = Path.Combine(startupPath, "include", "sod");

                    if (!Directory.Exists(strSourceDownload))
                    {
                        Directory.CreateDirectory(strSourceDownload);
                    }
                    var ts = DateTime.Now.ToString("yyyyMMdd_HHmmss");
                    string filename = $"{fileImport.ClientName}-Sod-{ts}.xlsx";
                    string strOutput = Path.Combine(strSourceDownload, filename);

                    //Check if file not exists
                    if (System.IO.File.Exists(strOutput))
                    {
                        System.IO.File.Delete(strOutput);
                    }

                    p.SaveAs(new FileInfo(strOutput));

                    response.Status = true;
                    response.Content = filename;
                    //writeLog.Display(listConflictDefinition);
                    //writeLog.Display(listSodUser);
                }

               
            }
            catch (Exception ex)
            {
                response.Status = false;
                response.Content = ex.ToString();
                FileLog.Write($"Error ProcessSodFile {ex}", "ErrorProcessSodFile");
                AdminService adminService = new AdminService(_config);
                adminService.SendAlert(true, true, ex.ToString(), "ProcessSodFile");
            }


            return response;
        }

        public Task<List<RoleUser>> ReadFileSodSoxRoxRoleUser(List<SoxRoxFile> listSoxFile, string clientName)
        {
            ServiceResponse response = new ServiceResponse();
            ExcelService xlsService = new ExcelService();
            WriteLog writeLog = new WriteLog();
            //SodSoxRox sodSoxRox = new SodSoxRox();
            List<RoleUser> listRoleUser = new List<RoleUser>();
            try
            {
                string startupPath = Directory.GetCurrentDirectory();
                string path = string.Empty;

                //Read Input File
                //position 0 - role user
                //position 1 - role permission
                //position 2 - conflict permission
                //position 3 - description to permission

                #region Read Role User
                Debug.WriteLine($"Reading Role User");
                var checkRoleUser = listSoxFile.Where(x => x.Position.Equals(0)).FirstOrDefault();
                if(checkRoleUser != null && checkRoleUser.NewFileName != string.Empty)
                {
                    path = Path.Combine(startupPath, "include", "upload", checkRoleUser.NewFileName);
                    FileInfo fi = new FileInfo(path);
                    using (ExcelPackage p = new ExcelPackage(fi))
                    {
                        // If you use EPPlus in a noncommercial context
                        // according to the Polyform Noncommercial license:
                        ////ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
                        ExcelWorksheet wsRoleUser = p.Workbook.Worksheets[0];
                        int colCountSheet1 = wsRoleUser.Dimension.End.Column;
                        int rowCountSheet1 = wsRoleUser.Dimension.End.Row;
                        for (int row = 2; row <= rowCountSheet1; row++)
                        {
                            RoleUser roleUser = new RoleUser();
                            for (int col = 1; col <= colCountSheet1; col++)
                            {
                                //Debug.WriteLine($"Col {col} : Value {wsRoleUser.Cells[row, col].Value?.ToString()}");
                                //Debug.WriteLine($"USER - Row:{row} | Col:{col} | Value:{wsRoleUser.Cells[row, col].Value?.ToString()}");
                                switch (col)
                                {
                                    case 1:
                                        roleUser.Name = wsRoleUser.Cells[row, col].Value?.ToString();
                                        break;
                                    case 2:
                                        roleUser.Email = wsRoleUser.Cells[row, col].Value?.ToString();
                                        break;
                                    case 3:
                                        roleUser.Phone = wsRoleUser.Cells[row, col].Value?.ToString();
                                        break;
                                    case 4:
                                        roleUser.Role = wsRoleUser.Cells[row, col].Value?.ToString();
                                        break;
                                    default:
                                        break;
                                }
                            }
                            
                            roleUser.Client = clientName;
                            roleUser.Row = row;
                            listRoleUser.Add(roleUser);
                            Debug.WriteLine($"Added ListRoleUser");
                            //writeLog.Display(listRoleUser);
                        }
                    }
                }

                #endregion

            }
            catch (Exception ex)
            {
                response.Status = false;
                response.Content = ex.ToString();
                FileLog.Write($"Error ProcessSodFile {ex}", "ErrorProcessSodFile");
                AdminService adminService = new AdminService(_config);
                adminService.SendAlert(true, true, ex.ToString(), "ProcessSodFile");
            }

            return Task.FromResult(listRoleUser);
        }

        public Task<List<RolePerm>> ReadFileSodSoxRoxRolePerm(List<SoxRoxFile> listSoxFile, string clientName)
        {
            ServiceResponse response = new ServiceResponse();
            ExcelService xlsService = new ExcelService();
            WriteLog writeLog = new WriteLog();
            //SodSoxRox sodSoxRox = new SodSoxRox();
            List<RolePerm> listRolePerm = new List<RolePerm>();
            try
            {
                string startupPath = Directory.GetCurrentDirectory();
                string path = string.Empty;

                //Read Input File
                //position 0 - role user
                //position 1 - role permission
                //position 2 - conflict permission
                //position 3 - description to permission
                #region Read Role Permission

                var checkRolePermission = listSoxFile.Where(x => x.Position.Equals(1)).FirstOrDefault();
                if (checkRolePermission != null && checkRolePermission.NewFileName != string.Empty)
                {
                    path = Path.Combine(startupPath, "include", "upload", checkRolePermission.NewFileName);
                    FileInfo fi = new FileInfo(path);
                    using (ExcelPackage p = new ExcelPackage(fi))
                    {
                        // If you use EPPlus in a noncommercial context
                        // according to the Polyform Noncommercial license:
                        ////ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
                        ExcelWorksheet wsRolePerm = p.Workbook.Worksheets[0];
                        int colCountSheet1 = wsRolePerm.Dimension.End.Column;
                        int rowCountSheet1 = wsRolePerm.Dimension.End.Row;
                        //int rowCountSheet1 = 50;
                        List<string> listHeader = new List<string>();

                        //Get Row Headers at row 1
                        for (int row = 1; row <= 1; row++)
                        {
                            for (int col = 1; col <= colCountSheet1; col++)
                            {
                                listHeader.Add(wsRolePerm.Cells[row, col].Value?.ToString());
                            }
                        }

                        //Get items - start at row 2
                        for (int row = 2; row <= rowCountSheet1; row++)
                        {
                            List<Perm> listPerm = new List<Perm>();
                            RolePerm rolePerm = new RolePerm();
                            for (int col = 1; col <= colCountSheet1; col++)
                            {

                                if(col == 2)
                                {
                                    rolePerm.Permission = wsRolePerm.Cells[row, col].Value?.ToString();
                                }

                                //Debug.WriteLine($"PERMISSION - Row:{row} | Col:{col} | Value:{wsRolePerm.Cells[row, col].Value?.ToString()}");
                                Perm permission = new Perm();
                                permission.Column = col;
                                permission.Header = listHeader[col - 1] ?? string.Empty;
                                permission.Value = wsRolePerm.Cells[row, col].Value?.ToString();
                                listPerm.Add(permission);
                            }
                            rolePerm.Client = clientName;
                            rolePerm.ListPerm = listPerm;
                            rolePerm.Row = row;
                            listRolePerm.Add(rolePerm);
                            Debug.WriteLine($"Added ListRolePerm");
                            //writeLog.Display(listRolePerm);
                        }


                    }
                }

                #endregion



            }
            catch (Exception ex)
            {
                response.Status = false;
                response.Content = ex.ToString();
                FileLog.Write($"Error ReadSodSoxRoxRolePerm {ex}", "ErrorReadSodSoxRoxRolePerm");
                AdminService adminService = new AdminService(_config);
                adminService.SendAlert(true, true, ex.ToString(), "ReadSodSoxRoxRolePerm");
            }
            //sodSoxRox.ListRolePerm = listRolePerm;
            return Task.FromResult(listRolePerm);
        }

        public Task<List<DescriptionToPerm>> ReadFileSodSoxRoxDescToPerm(List<SoxRoxFile> listSoxFile, string clientName)
        {
            ServiceResponse response = new ServiceResponse();
            ExcelService xlsService = new ExcelService();
            WriteLog writeLog = new WriteLog();
            //SodSoxRox sodSoxRox = new SodSoxRox();
            List<DescriptionToPerm> listDesc = new List<DescriptionToPerm>();
            try
            {
                string startupPath = Directory.GetCurrentDirectory();
                string path = string.Empty;

                //Read Input File
                //position 0 - role user
                //position 1 - role permission
                //position 2 - conflict permission
                //position 3 - description to permission
                #region Read Role Permission

                var checkDesc = listSoxFile.Where(x => x.Position.Equals(3)).FirstOrDefault();
                if (checkDesc != null && checkDesc.NewFileName != string.Empty)
                {
                    path = Path.Combine(startupPath, "include", "upload", checkDesc.NewFileName);
                    FileInfo fi = new FileInfo(path);
                    using (ExcelPackage p = new ExcelPackage(fi))
                    {
                        // If you use EPPlus in a noncommercial context
                        // according to the Polyform Noncommercial license:
                        ////ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
                        ExcelWorksheet wsDesc = p.Workbook.Worksheets[0];
                        int colCountSheet1 = wsDesc.Dimension.End.Column;
                        int rowCountSheet1 = wsDesc.Dimension.End.Row;
                        //int rowCountSheet1 = 50;
                        List<string> listHeader = new List<string>();

                        //Get items - start at row 2
                        for (int row = 2; row <= rowCountSheet1; row++)
                        {
                            
                            DescriptionToPerm descToPerm = new DescriptionToPerm();
                            for (int col = 1; col <= colCountSheet1; col++)
                            {
                                //Debug.WriteLine($"DESCIPTION - Row:{row} | Col:{col} | Value:{wsDesc.Cells[row, col].Value?.ToString()}");
                                switch (col)
                                {
                                    case 1:
                                        descToPerm.Description = wsDesc.Cells[row, col].Value?.ToString();
                                        break;
                                    case 2:
                                        descToPerm.Permission = wsDesc.Cells[row, col].Value?.ToString();
                                        break;
                                }
                            }
                            descToPerm.Client = clientName;
                            descToPerm.Row = row;
                            listDesc.Add(descToPerm);
                            Debug.WriteLine($"Added ListDescriptionToPerm");
                            //writeLog.Display(listDesc);
                        }


                    }
                }

                #endregion



            }
            catch (Exception ex)
            {
                response.Status = false;
                response.Content = ex.ToString();
                FileLog.Write($"Error ReadSodSoxRoxRolePerm {ex}", "ErrorReadSodSoxRoxRolePerm");
                AdminService adminService = new AdminService(_config);
                adminService.SendAlert(true, true, ex.ToString(), "ReadSodSoxRoxRolePerm");
            }
            //sodSoxRox.ListRolePerm = listRolePerm;
            return Task.FromResult(listDesc);
        }

        public Task<List<ConflictPerm>> ReadFileSodSoxRoxConflictPerm(List<SoxRoxFile> listSoxFile, string clientName)
        {
            ServiceResponse response = new ServiceResponse();
            ExcelService xlsService = new ExcelService();
            WriteLog writeLog = new WriteLog();
            //SodSoxRox sodSoxRox = new SodSoxRox();
            List<ConflictPerm> listDesc = new List<ConflictPerm>();
            try
            {
                string startupPath = Directory.GetCurrentDirectory();
                string path = string.Empty;

                //Read Input File
                //position 0 - role user
                //position 1 - role permission
                //position 2 - conflict permission
                //position 3 - description to permission
                #region Read Role Permission

                var checkDesc = listSoxFile.Where(x => x.Position.Equals(2)).FirstOrDefault();
                if (checkDesc != null && checkDesc.NewFileName != string.Empty)
                {
                    path = Path.Combine(startupPath, "include", "upload", checkDesc.NewFileName);
                    FileInfo fi = new FileInfo(path);
                    using (ExcelPackage p = new ExcelPackage(fi))
                    {
                        // If you use EPPlus in a noncommercial context
                        // according to the Polyform Noncommercial license:
                        //ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
                        ExcelWorksheet wsConflict = p.Workbook.Worksheets["Map_03_ConfltPerm"];
                        int colCountSheet1 = wsConflict.Dimension.End.Column;
                        int rowCountSheet1 = wsConflict.Dimension.End.Row;
                        //int rowCountSheet1 = 50;
                        List<string> listHeader = new List<string>();

                        //Get items - start at row 2
                        for (int row = 28; row <= rowCountSheet1; row++) //starts at row 28
                        {

                            ConflictPerm conflictPerm = new ConflictPerm();
                            for (int col = 1; col <= colCountSheet1; col++)
                            {
                                //Debug.WriteLine($"CONFLICT - Row:{row} | Col:{col} | Value:{wsConflict.Cells[row, col].Value?.ToString()}");
                                switch (col)
                                {
                                    case 1:
                                        conflictPerm.RefNum = wsConflict.Cells[row, col].Value?.ToString();
                                        break;
                                    case 2:
                                        conflictPerm.ProcessA = wsConflict.Cells[row, col].Value?.ToString();
                                        break;
                                    case 3:
                                        conflictPerm.ProcessB = wsConflict.Cells[row, col].Value?.ToString();
                                        break;
                                    case 4:
                                        conflictPerm.SODDescriptionA = wsConflict.Cells[row, col].Value?.ToString();
                                        break;
                                    case 5:
                                        conflictPerm.SODDescriptionB = wsConflict.Cells[row, col].Value?.ToString();
                                        break;
                                    case 6:
                                        conflictPerm.DescriptionOfConflict = wsConflict.Cells[row, col].Value?.ToString();
                                        break;
                                    case 7:
                                        conflictPerm.RiskComments = wsConflict.Cells[row, col].Value?.ToString();
                                        break;
                                    case 8:
                                        conflictPerm.System = wsConflict.Cells[row, col].Value?.ToString();
                                        break;
                                    case 9:
                                        conflictPerm.Manual = wsConflict.Cells[row, col].Value?.ToString();
                                        break;
                                    case 10:
                                        conflictPerm.Comment = wsConflict.Cells[row, col].Value?.ToString();
                                        break;
                                    case 11:
                                        conflictPerm.NsPermPairA = wsConflict.Cells[row, col].Value?.ToString();
                                        break;
                                    case 12:
                                        conflictPerm.NsPermPairB = wsConflict.Cells[row, col].Value?.ToString();
                                        break;
                                }
                            }
                            conflictPerm.Client = clientName;
                            conflictPerm.Row = row;
                            listDesc.Add(conflictPerm);
                            Debug.WriteLine($"Added ListConflict");
                            //writeLog.Display(listDesc);
                        }


                    }
                }

                #endregion



            }
            catch (Exception ex)
            {
                response.Status = false;
                response.Content = ex.ToString();
                FileLog.Write($"Error ReadSodSoxRoxRolePerm {ex}", "ErrorReadSodSoxRoxRolePerm");
                AdminService adminService = new AdminService(_config);
                adminService.SendAlert(true, true, ex.ToString(), "ReadSodSoxRoxRolePerm");
            }
            //sodSoxRox.ListRolePerm = listRolePerm;
            return Task.FromResult(listDesc);
        }

        public Task<List<SodSoxRoxDescriptionPermission>> ProcessSoxRoxDescriptionOutput(SodSoxRoxInput sodSoxRox)
        {
            SodSoxRoxOutputFile sodSoxRoxOutput = new SodSoxRoxOutputFile();

            List<SodSoxRoxDescriptionPermission> listSoxRoxPerm = new List<SodSoxRoxDescriptionPermission>();

            
                if (sodSoxRox.ListConflictPerm != null && sodSoxRox.ListConflictPerm.Count > 0)
                {
                    //Get SODDescriptionA
                    foreach (var item in sodSoxRox.ListConflictPerm.GroupBy(x => x.SODDescriptionA).Select(x => x).OrderBy(x => x.Key))
                    {
                        SodSoxRoxDescriptionPermission soxDescPerm = new SodSoxRoxDescriptionPermission();
                        soxDescPerm.Description = item.Key;
                        soxDescPerm.Permission = sodSoxRox.ListDescriptionToPerm.Where(x => x.Description.Equals(soxDescPerm.Description)).Select(x => x.Permission).FirstOrDefault();
                        soxDescPerm.SummaryCount = sodSoxRox.ListConflictPerm.Where(x => x.SODDescriptionA.Equals(soxDescPerm.Description)).Count();
                        soxDescPerm.Count = 1;
                        listSoxRoxPerm.Add(soxDescPerm);
                    }

                    //Get SODDescriptionB
                    foreach (var item in sodSoxRox.ListConflictPerm.GroupBy(x => x.SODDescriptionB).Select(x => x).OrderBy(x => x.Key))
                    {
                        SodSoxRoxDescriptionPermission soxDescPerm = new SodSoxRoxDescriptionPermission();
                        soxDescPerm.Description = item.Key;
                        soxDescPerm.Permission = sodSoxRox.ListDescriptionToPerm.Where(x => x.Description.Equals(soxDescPerm.Description)).Select(x => x.Permission).FirstOrDefault();
                        soxDescPerm.SummaryCount = sodSoxRox.ListConflictPerm.Where(x => x.SODDescriptionB.Equals(soxDescPerm.Description)).Count();
                        soxDescPerm.Count = 1;
                        listSoxRoxPerm.Add(soxDescPerm);
                    }


                }

            
            


            return Task.FromResult(listSoxRoxPerm);
        }

        public Task<string> CreateSoxRoxDescriptionFile(List<SodSoxRoxDescriptionPermission> listPerm, string clientName)
        {
            ExcelService xlsService = new ExcelService();
            WriteLog writeLog = new WriteLog();
    
            using (ExcelPackage p = new ExcelPackage())
            {

                // If you use EPPlus in a noncommercial context
                // according to the Polyform Noncommercial license:
                //ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

                var ws = p.Workbook.Worksheets.Add("Description_Perm_List");
                int row = 1;

                //Title
                ws.Cells[row, 1].Value = "Description";
                ws.Cells[row, 2].Value = "Permission";
                ws.Cells[row, 3].Value = "Summary_Count";
                ws.Cells[row, 4].Value = "COUNT1";
                row++;
                foreach (var item in listPerm)
                {
                    ws.Cells[row, 1].Value = item.Description;
                    ws.Cells[row, 2].Value = item.Permission;
                    ws.Cells[row, 3].Value = item.SummaryCount;
                    ws.Cells[row, 4].Value = item.Count;
                    row++;
                }



                string startupPath = Directory.GetCurrentDirectory();
                string strSourceDownload = Path.Combine(startupPath, "include", "sod");
                if (!Directory.Exists(strSourceDownload))
                {
                    Directory.CreateDirectory(strSourceDownload);
                }
                var ts = DateTime.Now.ToString("yyyyMMdd_HHmmss");
                string filename = $"{ts}_Description_to_Permission_List.xlsx";
                string strOutput = Path.Combine(strSourceDownload, filename);

                //Check if file not exists
                if (System.IO.File.Exists(strOutput))
                {
                    System.IO.File.Delete(strOutput);
                }

                p.SaveAs(new FileInfo(strOutput));
                return Task.FromResult(filename);
            }
           
           
           

        }

        public Task<string> CreateSoxRoxFile(List<SodSoxRoxDescriptionPermission> listPerm, string clientName)
        {
            ExcelService xlsService = new ExcelService();
            WriteLog writeLog = new WriteLog();

            using (ExcelPackage p = new ExcelPackage())
            {

                // If you use EPPlus in a noncommercial context
                // according to the Polyform Noncommercial license:
                //ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

                var ws = p.Workbook.Worksheets.Add("Description_Perm_List");
                int row = 1;

                //Title
                ws.Cells[row, 1].Value = "Description";
                ws.Cells[row, 2].Value = "Permission";
                ws.Cells[row, 3].Value = "Summary_Count";
                ws.Cells[row, 4].Value = "COUNT1";
                row++;
                foreach (var item in listPerm)
                {
                    ws.Cells[row, 1].Value = item.Description;
                    ws.Cells[row, 2].Value = item.Permission;
                    ws.Cells[row, 3].Value = item.SummaryCount;
                    ws.Cells[row, 4].Value = item.Count;
                    row++;
                }



                string startupPath = Directory.GetCurrentDirectory();
                string strSourceDownload = Path.Combine(startupPath, "include", "sod");
                if (!Directory.Exists(strSourceDownload))
                {
                    Directory.CreateDirectory(strSourceDownload);
                }
                var ts = DateTime.Now.ToString("yyyyMMdd_HHmmss");
                string filename = $"{ts}_Description to Permission List.xlsx";
                string strOutput = Path.Combine(strSourceDownload, filename);

                //Check if file not exists
                if (System.IO.File.Exists(strOutput))
                {
                    System.IO.File.Delete(strOutput);
                }

                p.SaveAs(new FileInfo(strOutput));
                return Task.FromResult(filename);
            }




        }

        public List<SodSoxRoxRoleUser> GetSodSoxRoxRoleUser(SodSoxRoxInput sodSoxRox)
        {

            List<SodSoxRoxRoleUser> listRoleUser = new List<SodSoxRoxRoleUser>();
            //we only check acccess level Create, Edit, Full
            List<string> listAccessLevel = new List<string> { "Create", "Edit", "Full" };
            //Get unique item from DescriptionToPerm 
            List<string> listUniqueDesc = new List<string>();
            listUniqueDesc = GetUniqueDescriptionToPerm(sodSoxRox.ListDescriptionToPerm);

            //Get list of roles
            List<string> listRoles = new List<string>();
            listRoles = GetMap02Roles(sodSoxRox.ListRolePerm);

            //loop through all user with role in Map_01_RoleUser and save in listRoleUser
            foreach (var item in sodSoxRox.ListRoleUser.GroupBy(x => x.Name))
            {
                SodSoxRoxRoleUser sodSoxUser = new SodSoxRoxRoleUser();
                sodSoxUser.Name = item.Key;
                var userDetail = sodSoxRox.ListRoleUser.Where(x => x.Name.Equals(sodSoxUser.Name)).Select(y => new { y.Email, y.Phone}).FirstOrDefault();
                if(userDetail != null)
                {
                    sodSoxUser.Email = userDetail.Email;
                    sodSoxUser.Phone = userDetail.Phone;
                }
                sodSoxUser.Phone = sodSoxRox.ListRoleUser.Where(x => x.Name.Equals(sodSoxUser.Name)).Select(y => y.Phone).FirstOrDefault();

                //Get list of role 
                var userRole = sodSoxRox.ListRoleUser.Where(x => x.Name.Equals(sodSoxUser.Name)).Select(y => y.Role);
                if(userRole != null)
                {
                    sodSoxUser.Role = userRole.ToList();
                }

                //Get list of permission
                if(sodSoxUser.Role != null && sodSoxUser.Role.Count > 0)
                {
                    List<string> tempPerm = new List<string>();
                    foreach (var itemRole in sodSoxUser.Role)
                    {
                        int roleIndex = listRoles.FindIndex(x => x.Equals(itemRole));
                        if (roleIndex > 0)
                        {
                            foreach (var itemPerm in sodSoxRox.ListRolePerm)
                            {
                                var accessLevel = itemPerm.ListPerm[roleIndex].Value; //Get Access level for this role
                                if (listUniqueDesc.Contains(itemPerm.Permission) && listAccessLevel.Contains(accessLevel) && !tempPerm.Contains(itemPerm.Permission))
                                {
                                    tempPerm.Add(itemPerm.Permission);
                                }
                            }
                        }
                    }
                    sodSoxUser.Permission = tempPerm;
                }

                listRoleUser.Add(sodSoxUser);
            }

            return listRoleUser;
        }

        public Task<List<SODSoxRoxReportRaw2>> ProcessSoxRoxDataRaw2(SodSoxRoxInput sodSoxRox)
        {
            List<SODSoxRoxReportRaw2> listSoxRoxRaw2 = new List<SODSoxRoxReportRaw2>();
            List<Perm> listPerm = new List<Perm>();

            List<string> listNsPermA = new List<string>();
            List<string> listNsPermB = new List<string>();

            //Get unique item from DescriptionToPerm 
            List<string> listUniqueDesc = new List<string>();
            listUniqueDesc = GetUniqueDescriptionToPerm(sodSoxRox.ListDescriptionToPerm);

            //Get list of roles
            List<string> listRoles = new List<string>();
            listRoles = GetMap02Roles(sodSoxRox.ListRolePerm);

            //we only check acccess level Create, Edit, Full
            List<string> listAccessLevel = new List<string> { "Create", "Edit", "Full" };

            //loop through all list of Role User Trim
            foreach (var itemUser in sodSoxRox.ListRoleUserTrim)
            {
                List<string> listRoleUser = new List<string>();

                //Loop role to determine self_conflict and role_conflict
                foreach (var itemRole1 in itemUser.Role)
                {
                    foreach (var itemRole2 in itemUser.Role)
                    {

                        //Debug.WriteLine($"Role_A: {itemRole1} : Role_B: {itemRole2} - {itemRole1.Equals(itemRole2)}");
                        //Debug.WriteLine($"--------------------------------------------------------------------------");

                        int roleIndexA = listRoles.FindIndex(x => x.Equals(itemRole2));
                        //if its greater than zero then role is found in Map_02_RolePerm
                        if (roleIndexA > 0)
                        {
                            foreach (var itemPerm in sodSoxRox.ListRolePerm)
                            {
                                var accessLevelA = itemPerm.ListPerm[roleIndexA].Value; //Get Access level for this role

                                //check if permission contains in DescriptionToPerm file
                                if (listUniqueDesc.Contains(itemPerm.Permission) && listAccessLevel.Contains(accessLevelA))
                                {

                                    //Debug.WriteLine($"Name: {itemUser.Name} | Role: {itemRole2} | Index: {roleIndex} | Permission: {itemPerm.Permission} | Access: {accessLevel} | Row: {itemPerm.Row} | Column: {itemPerm.ListPerm[roleIndex].Column}");

                                    //start analysis
                                    //Get conflict configuration from Map_03_ConflictPerm
                                    foreach (var itemConflictPerm in sodSoxRox.ListConflictPerm.Where(x => x.NsPermPairA.Contains(itemPerm.Permission)))
                                    {
                                        //Get NS PermA and NS PermB in Map_03_ConflictPerm
                                        listNsPermA = itemConflictPerm.NsPermPairA.Split('|').ToList();
                                        listNsPermB = itemConflictPerm.NsPermPairB.Split('|').ToList();
                                        if (listNsPermA.Contains(itemPerm.Permission)) //if permission is found in A, then loop through NSPermB
                                        {
                                            
                                            foreach (var itemPermB in listNsPermB)
                                            {
                                                //int roleIndexB = listRoles.FindIndex(x => x.Equals(itemRole2));
                                                string accessLevelB = "None";
                                                var rawlevel = sodSoxRox.ListRolePerm.Where(x => x.Permission.Equals(itemPermB)).FirstOrDefault();
                                                if (rawlevel != null)
                                                {
                                                    accessLevelB = rawlevel.ListPerm.Where(x => x.Header.Equals(itemRole2)).Select(y => y.Value).FirstOrDefault();
                                                }

                                                if (itemUser.Permission.Contains(itemPermB) && listAccessLevel.Contains(accessLevelB))
                                                {

                                                    //Debug.WriteLine($"Name: {itemUser.Name} | Role: {itemRole2} | Index: {roleIndex} | Permission: {itemPerm.Permission} | Access: {accessLevel} | Row: {itemPerm.Row} | Column: {itemPerm.ListPerm[roleIndex].Column}");
                                                    SODSoxRoxReportRaw2 sodSoxRoxReportRaw2 = new SODSoxRoxReportRaw2();
                                                    sodSoxRoxReportRaw2.Name = itemUser.Name;
                                                    sodSoxRoxReportRaw2.Email = itemUser.Email;
                                                    sodSoxRoxReportRaw2.Phone = itemUser.Phone;
                                                    sodSoxRoxReportRaw2.SODRef = itemConflictPerm.RefNum;
                                                    sodSoxRoxReportRaw2.RoleA = itemRole1;
                                                    sodSoxRoxReportRaw2.RoleB = itemRole2;
                                                    sodSoxRoxReportRaw2.PermA = itemPerm.Permission;
                                                    sodSoxRoxReportRaw2.PermB = itemPermB;
                                                    sodSoxRoxReportRaw2.ProcessA = itemConflictPerm.ProcessA;
                                                    sodSoxRoxReportRaw2.ProcessB = itemConflictPerm.ProcessB;
                                                    sodSoxRoxReportRaw2.FunctionTypeA = accessLevelA;
                                                    sodSoxRoxReportRaw2.FunctionTypeB = accessLevelB;
                                                    sodSoxRoxReportRaw2.DescOfConflict = itemConflictPerm.DescriptionOfConflict;
                                                    sodSoxRoxReportRaw2.DescriptionA = itemConflictPerm.SODDescriptionA;
                                                    sodSoxRoxReportRaw2.DescriptionB = itemConflictPerm.SODDescriptionB;
                                                    sodSoxRoxReportRaw2.RiskComments = itemConflictPerm.RiskComments;
                                                    switch (itemConflictPerm.System)
                                                    {
                                                        case string s when s.ToLower().Equals("low"):
                                                            sodSoxRoxReportRaw2.SODPriority = "3-Low";
                                                            break;
                                                        case string s when s.ToLower().Equals("med") || s.ToLower().Equals("medium"):
                                                            sodSoxRoxReportRaw2.SODPriority = "2-Medium";
                                                            break;
                                                        case string s when s.ToLower().Equals("hig") || s.ToLower().Equals("high"):
                                                            sodSoxRoxReportRaw2.SODPriority = "1-High";
                                                            break;
                                                        default:
                                                            sodSoxRoxReportRaw2.SODPriority = "4-no";
                                                            break;
                                                    };
                                                    sodSoxRoxReportRaw2.ConflictRolePair = $"{itemRole1} and {itemRole2}";
                                                    sodSoxRoxReportRaw2.ConflictType = (itemRole1.Equals(itemRole2) ? "Self_Conflict" : "Conflict_Roles");
                                                    listSoxRoxRaw2.Add(sodSoxRoxReportRaw2);

                                                }
                                            
                                            }
                                        }

                                    }

                                }
                            }
                        }

                        //Debug.WriteLine($"--------------------------------------------------------------------------");
                    }
                }

            }


            return Task.FromResult(listSoxRoxRaw2);
        }


        public Task<List<SODSoxRoxReportRaw2>> ProcessSoxRoxDataRaw2_2(SodSoxRoxInput sodSoxRox)
        {
            List<SODSoxRoxReportRaw2> listSoxRoxRaw2 = new List<SODSoxRoxReportRaw2>();
            List<Perm> listPerm = new List<Perm>();

            List<string> listNsPermA = new List<string>();
            List<string> listNsPermB = new List<string>();

            //Get unique item from DescriptionToPerm 
            List<string> listUniqueDesc = new List<string>();
            listUniqueDesc = GetUniqueDescriptionToPerm(sodSoxRox.ListDescriptionToPerm);

            //Get list of roles
            List<string> listRoles = new List<string>();
            listRoles = GetMap02Roles(sodSoxRox.ListRolePerm);

            //we only check acccess level Create, Edit, Full
            List<string> listAccessLevel = new List<string> { "Create", "Edit", "Full" };

            //loop through all list of Role User Trim
            foreach (var itemUser in sodSoxRox.ListRoleUserTrim)
            {
                List<string> listRoleUser = new List<string>();

                //Loop role to determine self_conflict and role_conflict
                foreach (var itemRole1 in itemUser.Role)
                {
                    foreach (var itemRole2 in itemUser.Role)
                    {

                        //Debug.WriteLine($"Role_A: {itemRole1} : Role_B: {itemRole2} - {itemRole1.Equals(itemRole2)}");
                        //Debug.WriteLine($"--------------------------------------------------------------------------");

                        
                        foreach (var itemPerm in sodSoxRox.ListRolePerm.Where(x => listUniqueDesc.Contains(x.Permission)))
                        {
                            var accessLevelA = itemPerm.ListPerm.Where(x => x.Header.Equals(itemRole1)).Select(y => y.Value).FirstOrDefault() ?? "None"; //Get Access level for this role

                            //check if permission contains in DescriptionToPerm file
                            if (listAccessLevel.Contains(accessLevelA))
                            {

                                //Debug.WriteLine($"Name: {itemUser.Name} | Role: {itemRole2} | Index: {roleIndex} | Permission: {itemPerm.Permission} | Access: {accessLevel} | Row: {itemPerm.Row} | Column: {itemPerm.ListPerm[roleIndex].Column}");

                                //start analysis
                                //Get conflict configuration from Map_03_ConflictPerm
                                foreach (var itemConflictPerm in sodSoxRox.ListConflictPerm.Where(x => x.NsPermPairA.Contains(itemPerm.Permission)))
                                {
                                    //Get NS PermA and NS PermB in Map_03_ConflictPerm
                                    listNsPermA = itemConflictPerm.NsPermPairA.Split('|').ToList();
                                    listNsPermB = itemConflictPerm.NsPermPairB.Split('|').ToList();
                                    if (listNsPermA.Contains(itemPerm.Permission)) //if permission is found in A, then loop through NSPermB
                                    {

                                        foreach (var itemPermB in listNsPermB)
                                        {
                                            //int roleIndexB = listRoles.FindIndex(x => x.Equals(itemRole2));
                                            string accessLevelB = "None";
                                            var rawlevelB = sodSoxRox.ListRolePerm.Where(x => x.Permission.Equals(itemPermB)).FirstOrDefault();
                                            if (rawlevelB != null)
                                            {
                                                accessLevelB = rawlevelB.ListPerm.Where(x => x.Header.Equals(itemRole2)).Select(y => y.Value).FirstOrDefault();
                                            }

                                            if (itemUser.Permission.Contains(itemPermB) && listAccessLevel.Contains(accessLevelB))
                                            {

                                                //Debug.WriteLine($"Name: {itemUser.Name} | Role: {itemRole2} | Index: {roleIndex} | Permission: {itemPerm.Permission} | Access: {accessLevel} | Row: {itemPerm.Row} | Column: {itemPerm.ListPerm[roleIndex].Column}");
                                                SODSoxRoxReportRaw2 sodSoxRoxReportRaw2 = new SODSoxRoxReportRaw2();
                                                sodSoxRoxReportRaw2.Name = itemUser.Name;
                                                sodSoxRoxReportRaw2.Email = itemUser.Email;
                                                sodSoxRoxReportRaw2.Phone = itemUser.Phone;
                                                sodSoxRoxReportRaw2.SODRef = itemConflictPerm.RefNum;
                                                sodSoxRoxReportRaw2.RoleA = itemRole1;
                                                sodSoxRoxReportRaw2.RoleB = itemRole2;
                                                sodSoxRoxReportRaw2.PermA = itemPerm.Permission;
                                                sodSoxRoxReportRaw2.PermB = itemPermB;
                                                sodSoxRoxReportRaw2.ProcessA = itemConflictPerm.ProcessA;
                                                sodSoxRoxReportRaw2.ProcessB = itemConflictPerm.ProcessB;
                                                sodSoxRoxReportRaw2.FunctionTypeA = accessLevelA;
                                                sodSoxRoxReportRaw2.FunctionTypeB = accessLevelB;
                                                sodSoxRoxReportRaw2.DescOfConflict = itemConflictPerm.DescriptionOfConflict;
                                                sodSoxRoxReportRaw2.DescriptionA = itemConflictPerm.SODDescriptionA;
                                                sodSoxRoxReportRaw2.DescriptionB = itemConflictPerm.SODDescriptionB;
                                                sodSoxRoxReportRaw2.RiskComments = itemConflictPerm.RiskComments;
                                                switch (itemConflictPerm.System)
                                                {
                                                    case string s when s.ToLower().Equals("low"):
                                                        sodSoxRoxReportRaw2.SODPriority = "3-Low";
                                                        break;
                                                    case string s when s.ToLower().Equals("med") || s.ToLower().Equals("medium"):
                                                        sodSoxRoxReportRaw2.SODPriority = "2-Medium";
                                                        break;
                                                    case string s when s.ToLower().Equals("hig") || s.ToLower().Equals("high"):
                                                        sodSoxRoxReportRaw2.SODPriority = "1-High";
                                                        break;
                                                    default:
                                                        sodSoxRoxReportRaw2.SODPriority = "4-no";
                                                        break;
                                                };
                                                sodSoxRoxReportRaw2.ConflictRolePair = $"{itemRole1} and {itemRole2}";
                                                sodSoxRoxReportRaw2.ConflictType = (itemRole1.Equals(itemRole2) ? "Self_Conflict" : "Conflict_Roles");

                                                bool checkExists = listSoxRoxRaw2.Any(x =>
                                                    x.Name.Equals(sodSoxRoxReportRaw2.Name) &&
                                                    x.Email.Equals(sodSoxRoxReportRaw2.Email) &&

                                                    x.RoleA.Equals(sodSoxRoxReportRaw2.RoleA) &&
                                                    x.RoleB.Equals(sodSoxRoxReportRaw2.RoleB) &&
                                                        
                                                    x.PermA.Equals(sodSoxRoxReportRaw2.PermA) &&
                                                    x.PermB.Equals(sodSoxRoxReportRaw2.PermB) &&

                                                    x.ProcessA.Equals(sodSoxRoxReportRaw2.ProcessA) &&
                                                    x.ProcessB.Equals(sodSoxRoxReportRaw2.ProcessB) &&

                                                    x.FunctionTypeA.Equals(sodSoxRoxReportRaw2.FunctionTypeA) &&
                                                    x.FunctionTypeB.Equals(sodSoxRoxReportRaw2.FunctionTypeB) &&

                                                    x.ConflictType.Equals(sodSoxRoxReportRaw2.ConflictType) &&
                                                    x.SODPriority.Equals(sodSoxRoxReportRaw2.SODPriority) &&
                                                    x.SODRef.Equals(sodSoxRoxReportRaw2.SODRef)
                                                );
                                                if(!checkExists)
                                                {
                                                    listSoxRoxRaw2.Add(sodSoxRoxReportRaw2);
                                                }
                                                    

                                            }

                                        }
                                    }

                                }

                            }
                        }
                        

                        //Debug.WriteLine($"--------------------------------------------------------------------------");
                    }
                }

            }

            return Task.FromResult(listSoxRoxRaw2);
        }


        public Task<List<SODSoxRoxReportRaw2>> ProcessSoxRoxDataRaw2_3(SodSoxRoxInput sodSoxRox)
        {
            List<SODSoxRoxReportRaw2> listSoxRoxRaw2 = new List<SODSoxRoxReportRaw2>();
            List<Perm> listPerm = new List<Perm>();

            List<string> listNsPermA = new List<string>();
            List<string> listNsPermB = new List<string>();

            //Get unique item from DescriptionToPerm 
            List<string> listUniqueDesc = new List<string>();
            listUniqueDesc = GetUniqueDescriptionToPerm(sodSoxRox.ListDescriptionToPerm);

            //Get list of roles
            List<string> listRoles = new List<string>();
            listRoles = GetMap02Roles(sodSoxRox.ListRolePerm);

            //we only check acccess level Create, Edit, Full
            List<string> listAccessLevel = new List<string> { "Create", "Edit", "Full" };

            //loop through all list of Role User Trim
            foreach (var itemUser in sodSoxRox.ListRoleUserTrim)
            {
                List<string> listRoleUser = new List<string>();

                //Loop role to determine self_conflict and role_conflict
                foreach (var itemRole1 in itemUser.Role) //loop through its available role
                {
                    foreach (var itemPermA in sodSoxRox.ListRolePerm) //loop list perm A
                    {
                        var accessLevelA = itemPermA.ListPerm.Where(x => x.Header.Equals(itemRole1)).Select(y => y.Value).FirstOrDefault() ?? "None"; //Get Access level for this role
                        if (listUniqueDesc.Contains(itemPermA.Permission) && listAccessLevel.Contains(accessLevelA)) //check if itemPermA exists in DescriptionToPerm and access level is full, edit, create
                        {
                            foreach (var itemRole2 in itemUser.Role) //loop through its available role
                            {

                                foreach (var itemPermB in sodSoxRox.ListRolePerm) //loop list perm A
                                {
                                    var accessLevelB = itemPermB.ListPerm.Where(x => x.Header.Equals(itemRole2)).Select(y => y.Value).FirstOrDefault() ?? "None"; //Get Access level for this role
                                    if (listUniqueDesc.Contains(itemPermB.Permission) && listAccessLevel.Contains(accessLevelB)) //check if itemPermB exists in DescriptionToPerm
                                    {
                                        //check permission in ConflictToPerm
                                        var listConflict = sodSoxRox.ListConflictPerm.Where(x => x.NsPermPairA.Contains(itemPermA.Permission) && x.NsPermPairB.Contains(itemPermB.Permission));
                                        if (listConflict.Any())
                                        {
                                            foreach (var itemConflictPerm in listConflict)
                                            {
                                                //Get NS PermA and NS PermB in Map_03_ConflictPerm
                                                listNsPermA = itemConflictPerm.NsPermPairA.Split('|').ToList();
                                                listNsPermB = itemConflictPerm.NsPermPairB.Split('|').ToList();

                                                if (listNsPermA.Contains(itemPermA.Permission) && listNsPermB.Contains(itemPermB.Permission)) //if permission is found in A and B
                                                {
                                                    //Debug.WriteLine($"Name: {itemUser.Name} | Role: {itemRole2} | Index: {roleIndex} | Permission: {itemPerm.Permission} | Access: {accessLevel} | Row: {itemPerm.Row} | Column: {itemPerm.ListPerm[roleIndex].Column}");
                                                    SODSoxRoxReportRaw2 sodSoxRoxReportRaw2 = new SODSoxRoxReportRaw2();
                                                    sodSoxRoxReportRaw2.Name = itemUser.Name;
                                                    sodSoxRoxReportRaw2.Email = itemUser.Email;
                                                    sodSoxRoxReportRaw2.Phone = itemUser.Phone;
                                                    sodSoxRoxReportRaw2.SODRef = itemConflictPerm.RefNum;
                                                    sodSoxRoxReportRaw2.RoleA = itemRole1;
                                                    sodSoxRoxReportRaw2.RoleB = itemRole2;
                                                    sodSoxRoxReportRaw2.PermA = itemPermA.Permission;
                                                    sodSoxRoxReportRaw2.PermB = itemPermB.Permission;
                                                    sodSoxRoxReportRaw2.ProcessA = itemConflictPerm.ProcessA;
                                                    sodSoxRoxReportRaw2.ProcessB = itemConflictPerm.ProcessB;
                                                    sodSoxRoxReportRaw2.FunctionTypeA = accessLevelA;
                                                    sodSoxRoxReportRaw2.FunctionTypeB = accessLevelB;
                                                    sodSoxRoxReportRaw2.DescOfConflict = itemConflictPerm.DescriptionOfConflict;
                                                    sodSoxRoxReportRaw2.DescriptionA = itemConflictPerm.SODDescriptionA;
                                                    sodSoxRoxReportRaw2.DescriptionB = itemConflictPerm.SODDescriptionB;
                                                    sodSoxRoxReportRaw2.RiskComments = itemConflictPerm.RiskComments;
                                                    //if(itemConflictPerm.RefNum == "SOD0229")
                                                    //{
                                                    //    Console.WriteLine($"itemConflictPerm.System: {itemConflictPerm.System} | itemConflictPerm.Manual: {itemConflictPerm.Manual}");
                                                    //}
                                                    if(itemConflictPerm.System != string.Empty && itemConflictPerm.System != null)
                                                    {
                                                        switch (itemConflictPerm.System)
                                                        {
                                                            case string s when s.ToLower().Equals("low"):
                                                                sodSoxRoxReportRaw2.SODPriority = "3-Low";
                                                                break;
                                                            case string s when s.ToLower().Equals("med") || s.ToLower().Equals("medium"):
                                                                sodSoxRoxReportRaw2.SODPriority = "2-Medium";
                                                                break;
                                                            case string s when s.ToLower().Equals("hig") || s.ToLower().Equals("high"):
                                                                sodSoxRoxReportRaw2.SODPriority = "1-High";
                                                                break;
                                                            default:
                                                                sodSoxRoxReportRaw2.SODPriority = "4-no";
                                                                break;
                                                        };
                                                    }
                                                    else
                                                    {
                                                        switch (itemConflictPerm.Manual)
                                                        {
                                                            case string s2 when s2.ToLower().Equals("low"):
                                                                sodSoxRoxReportRaw2.SODPriority = "3-Low";
                                                                break;
                                                            case string s2 when s2.ToLower().Equals("med") || s2.ToLower().Equals("medium"):
                                                                sodSoxRoxReportRaw2.SODPriority = "2-Medium";
                                                                break;
                                                            case string s2 when s2.ToLower().Equals("hig") || s2.ToLower().Equals("high"):
                                                                sodSoxRoxReportRaw2.SODPriority = "1-High";
                                                                break;
                                                            default:
                                                                sodSoxRoxReportRaw2.SODPriority = "4-no";
                                                                break;
                                                        };
                                                    }



                                                    sodSoxRoxReportRaw2.ConflictRolePair = $"{itemRole1} and {itemRole2}";
                                                    sodSoxRoxReportRaw2.ConflictType = (itemRole1.Equals(itemRole2) ? "Self_Conflict" : "Conflict_Roles");

                                                    bool checkExists = listSoxRoxRaw2.Any(x =>
                                                        x.Name.Equals(sodSoxRoxReportRaw2.Name) &&
                                                        (sodSoxRoxReportRaw2.Email != null ? x.Email.Equals(sodSoxRoxReportRaw2.Email) : false) &&

                                                        x.RoleA.Equals(sodSoxRoxReportRaw2.RoleA) &&
                                                        x.RoleB.Equals(sodSoxRoxReportRaw2.RoleB) &&

                                                        x.PermA.Equals(sodSoxRoxReportRaw2.PermA) &&
                                                        x.PermB.Equals(sodSoxRoxReportRaw2.PermB) &&

                                                        x.ProcessA.Equals(sodSoxRoxReportRaw2.ProcessA) &&
                                                        x.ProcessB.Equals(sodSoxRoxReportRaw2.ProcessB) &&

                                                        x.FunctionTypeA.Equals(sodSoxRoxReportRaw2.FunctionTypeA) &&
                                                        x.FunctionTypeB.Equals(sodSoxRoxReportRaw2.FunctionTypeB) &&

                                                        x.ConflictType.Equals(sodSoxRoxReportRaw2.ConflictType) &&
                                                        x.SODPriority.Equals(sodSoxRoxReportRaw2.SODPriority) &&
                                                        x.SODRef.Equals(sodSoxRoxReportRaw2.SODRef)
                                                    );
                                                    if (!checkExists)
                                                    {
                                                        listSoxRoxRaw2.Add(sodSoxRoxReportRaw2);
                                                    }
                                                }


                                            }
                                        }
                                    }
                                }
                            }
                        }
                    }
                        
                }


            }

            return Task.FromResult(listSoxRoxRaw2);
        }



        public Task<List<SODSoxRoxReportRaw3>> ProcessSoxRoxDataRaw3(SodSoxRoxInput sodSoxRox)
        {
            List<SODSoxRoxReportRaw3> listSoxRoxRaw3 = new List<SODSoxRoxReportRaw3>();
            List<Perm> listPerm = new List<Perm>();

            List<string> listNsPermA = new List<string>();
            List<string> listNsPermB = new List<string>();

            //Get unique item from DescriptionToPerm 
            List<string> listUniqueDesc = new List<string>();
            listUniqueDesc = GetUniqueDescriptionToPerm(sodSoxRox.ListDescriptionToPerm);

            //Get list of roles
            List<string> listRoles = new List<string>();
            listRoles = GetMap02Roles(sodSoxRox.ListRolePerm);

            //we only check acccess level Create, Edit, Full
            List<string> listAccessLevel = new List<string> { "Create", "Edit", "Full" };



            if (listRoles.Count > 0)
            {
                foreach (var itemRole1 in listRoles) //loop through role A
                {
                    if(!itemRole1.Equals("Category") && !itemRole1.Equals("Permission")) //exclude category and permission. First and second column in map02
                    {

                        List<string> listPermA = new List<string>();
                        int roleIndexMain = listRoles.FindIndex(x => x.Equals(itemRole1));
                        if (roleIndexMain > 0)
                        {
                            foreach (var item in sodSoxRox.ListRolePerm)
                            {
                                if(listAccessLevel.Contains(item.ListPerm[roleIndexMain].Value) && !listPermA.Contains(item.Permission))
                                {
                                    listPermA.Add(item.Permission);
                                }
                            }
                        }


                        foreach (var itemRole2 in listRoles) //loop through role B
                        {
                            if (!itemRole2.Equals("Category") && !itemRole2.Equals("Permission")) //exclude category and permission. First and second column in map02
                            {
                                int roleIndexA = listRoles.FindIndex(x => x.Equals(itemRole2));
                                if(roleIndexA > 0)
                                {
                                    foreach (var itemPerm in sodSoxRox.ListRolePerm) //loop permission
                                    {
                                        var accessLevelA = itemPerm.ListPerm[roleIndexA].Value; //Get Access level for this role

                                        //check if permission contains in DescriptionToPerm file
                                        if (listUniqueDesc.Contains(itemPerm.Permission) && listAccessLevel.Contains(accessLevelA) && listPermA.Contains(itemPerm.Permission))
                                        {

                                            //Debug.WriteLine($"RoleA:{itemRole1} | RoleB:{itemRole2} - PERM: {itemPerm.Permission}");

                                            //start analysis 1
                                            //Get conflict configuration from Map_03_ConflictPerm Permission A
                                            var listConflictA = sodSoxRox.ListConflictPerm.Where(x => x.NsPermPairA.Contains(itemPerm.Permission));
                                            if(listConflictA.Count() > 0)
                                            {
                                                foreach (var itemConflictPerm in listConflictA)
                                                {
                                                    //Get NS PermA and NS PermB in Map_03_ConflictPerm
                                                    listNsPermA = itemConflictPerm.NsPermPairA.Split('|').ToList();
                                                    listNsPermB = itemConflictPerm.NsPermPairB.Split('|').ToList();
                                                    if (listNsPermA.Contains(itemPerm.Permission)) //if permission is found, then loop through NSPermB
                                                    {
                                                        foreach (var itemPermB in listNsPermB)
                                                        {
                                                            //int roleIndexB = listRoles.FindIndex(x => x.Equals(itemRole2));
                                                            string accessLevelB = "None";
                                                            var rawlevel = sodSoxRox.ListRolePerm.Where(x => x.Permission.Equals(itemPermB)).FirstOrDefault();
                                                            if (rawlevel != null)
                                                            {
                                                                accessLevelB = rawlevel.ListPerm.Where(x => x.Header.Equals(itemRole2)).Select(y => y.Value).FirstOrDefault();
                                                            }

                                                            if (listAccessLevel.Contains(accessLevelB))
                                                            {

                                                                //Debug.WriteLine($"Name: {itemUser.Name} | Role: {itemRole2} | Index: {roleIndex} | Permission: {itemPerm.Permission} | Access: {accessLevel} | Row: {itemPerm.Row} | Column: {itemPerm.ListPerm[roleIndex].Column}");
                                                                SODSoxRoxReportRaw3 sodSoxRoxReportRaw3 = new SODSoxRoxReportRaw3();
                                                                sodSoxRoxReportRaw3.SODRef = itemConflictPerm.RefNum;
                                                                sodSoxRoxReportRaw3.RoleA = itemRole1;
                                                                sodSoxRoxReportRaw3.RoleB = itemRole2;
                                                                sodSoxRoxReportRaw3.PermA = itemPerm.Permission;
                                                                sodSoxRoxReportRaw3.PermB = itemPermB;
                                                                sodSoxRoxReportRaw3.ProcessA = itemConflictPerm.ProcessA;
                                                                sodSoxRoxReportRaw3.ProcessB = itemConflictPerm.ProcessB;
                                                                sodSoxRoxReportRaw3.DescOfConflict = itemConflictPerm.DescriptionOfConflict;
                                                                sodSoxRoxReportRaw3.DescriptionA = itemConflictPerm.SODDescriptionA;
                                                                sodSoxRoxReportRaw3.DescriptionB = itemConflictPerm.SODDescriptionB;
                                                                sodSoxRoxReportRaw3.RiskComments = itemConflictPerm.RiskComments;
                                                                switch (itemConflictPerm.System)
                                                                {
                                                                    case string s when s.ToLower().Equals("low"):
                                                                        sodSoxRoxReportRaw3.SODPriority = "3-Low";
                                                                        break;
                                                                    case string s when s.ToLower().Equals("med") || s.ToLower().Equals("medium"):
                                                                        sodSoxRoxReportRaw3.SODPriority = "2-Medium";
                                                                        break;
                                                                    case string s when s.ToLower().Equals("hig") || s.ToLower().Equals("high"):
                                                                        sodSoxRoxReportRaw3.SODPriority = "1-High";
                                                                        break;
                                                                    case string s when s == string.Empty: //if itemConflictPerm.System is blank then we check the manual column
                                                                        switch (itemConflictPerm.Manual)
                                                                        {
                                                                            case string s2 when s2.ToLower().Equals("low"):
                                                                                sodSoxRoxReportRaw3.SODPriority = "3-Low";
                                                                                break;
                                                                            case string s2 when s2.ToLower().Equals("med") || s2.ToLower().Equals("medium"):
                                                                                sodSoxRoxReportRaw3.SODPriority = "2-Medium";
                                                                                break;
                                                                            case string s2 when s2.ToLower().Equals("hig") || s2.ToLower().Equals("high"):
                                                                                sodSoxRoxReportRaw3.SODPriority = "1-High";
                                                                                break;
                                                                            default:
                                                                                sodSoxRoxReportRaw3.SODPriority = "4-no";
                                                                                break;
                                                                        }
                                                                        break;
                                                                    default:
                                                                        sodSoxRoxReportRaw3.SODPriority = "4-no";
                                                                        break;
                                                                };
                                                                sodSoxRoxReportRaw3.ConflictRolePair = $"{itemRole1} and {itemRole2}";
                                                                sodSoxRoxReportRaw3.ConflictType = (itemRole1.Equals(itemRole2) ? "Self_Conflict" : "Role_Conflict");
                                                                sodSoxRoxReportRaw3.AssignedA = string.Empty;
                                                                sodSoxRoxReportRaw3.AssignedB = string.Empty;

                                                                listSoxRoxRaw3.Add(sodSoxRoxReportRaw3);

                                                            }

                                                        }
                                                    }
                                                }
                                            }
                                            //else
                                            //{
                                            //    //start analysis 2
                                            //    //Get conflict configuration from Map_03_ConflictPerm Permission B
                                            //    var listConflictB = sodSoxRox.ListConflictPerm.Where(x => x.NsPermPairB.Contains(itemPerm.Permission));
                                            //    if (listConflictB.Count() > 0)
                                            //    {
                                            //        foreach (var itemConflictPerm in listConflictB)
                                            //        {
                                            //            //Get NS PermA and NS PermB in Map_03_ConflictPerm
                                            //            listNsPermA = itemConflictPerm.NsPermPairA.Split('|').ToList();
                                            //            listNsPermB = itemConflictPerm.NsPermPairB.Split('|').ToList();
                                            //            if (listNsPermB.Contains(itemPerm.Permission)) //if permission is found, then loop through NSPermB
                                            //            {
                                            //                foreach (var itemPermA in listNsPermA)
                                            //                {
                                            //                    if (listPermA.Contains(itemPermA))
                                            //                    {

                                            //                        //Debug.WriteLine($"Name: {itemUser.Name} | Role: {itemRole2} | Index: {roleIndex} | Permission: {itemPerm.Permission} | Access: {accessLevel} | Row: {itemPerm.Row} | Column: {itemPerm.ListPerm[roleIndex].Column}");
                                            //                        SODSoxRoxReportRaw3 sodSoxRoxReportRaw3 = new SODSoxRoxReportRaw3();
                                            //                        sodSoxRoxReportRaw3.SODRef = itemConflictPerm.RefNum;
                                            //                        sodSoxRoxReportRaw3.RoleA = itemRole1;
                                            //                        sodSoxRoxReportRaw3.RoleB = itemRole2;
                                            //                        sodSoxRoxReportRaw3.PermA = itemPermA;
                                            //                        sodSoxRoxReportRaw3.PermB = itemPerm.Permission;
                                            //                        sodSoxRoxReportRaw3.ProcessA = itemConflictPerm.ProcessA;
                                            //                        sodSoxRoxReportRaw3.ProcessB = itemConflictPerm.ProcessB;
                                            //                        sodSoxRoxReportRaw3.DescOfConflict = itemConflictPerm.DescriptionOfConflict;
                                            //                        sodSoxRoxReportRaw3.DescriptionA = itemConflictPerm.SODDescriptionA;
                                            //                        sodSoxRoxReportRaw3.DescriptionB = itemConflictPerm.SODDescriptionB;
                                            //                        sodSoxRoxReportRaw3.RiskComments = itemConflictPerm.RiskComments;
                                            //                        switch (itemConflictPerm.System)
                                            //                        {
                                            //                            case string s when s.ToLower().Equals("low"):
                                            //                                sodSoxRoxReportRaw3.SODPriority = "3-Low";
                                            //                                break;
                                            //                            case string s when s.ToLower().Equals("med") || s.ToLower().Equals("medium"):
                                            //                                sodSoxRoxReportRaw3.SODPriority = "2-Medium";
                                            //                                break;
                                            //                            case string s when s.ToLower().Equals("hig") || s.ToLower().Equals("high"):
                                            //                                sodSoxRoxReportRaw3.SODPriority = "1-High";
                                            //                                break;
                                            //                            default:
                                            //                                sodSoxRoxReportRaw3.SODPriority = "4-no";
                                            //                                break;
                                            //                        };
                                            //                        sodSoxRoxReportRaw3.ConflictRolePair = $"{itemRole1} and {itemRole2}";
                                            //                        sodSoxRoxReportRaw3.ConflictType = (itemRole1.Equals(itemRole2) ? "Self_Conflict" : "Role_Conflict");
                                            //                        sodSoxRoxReportRaw3.AssignedA = string.Empty;
                                            //                        sodSoxRoxReportRaw3.AssignedB = string.Empty;

                                            //                        listSoxRoxRaw3.Add(sodSoxRoxReportRaw3);

                                            //                    }

                                            //                }
                                            //            }
                                            //        }
                                            //    }
                                            //}
                                         

                                        }
                                    }
                                }
                                
                            }
                        }
                    
                    
                    }
                }
            }
            

            


            return Task.FromResult(listSoxRoxRaw3);
        }


        public Task<List<SODSoxRoxReportRaw3>> ProcessSoxRoxDataRaw3_2(SodSoxRoxInput sodSoxRox)
        {
            List<SODSoxRoxReportRaw3> listSoxRoxRaw3 = new List<SODSoxRoxReportRaw3>();
            List<Perm> listPerm = new List<Perm>();

            List<string> listNsPermA = new List<string>();
            List<string> listNsPermB = new List<string>();

            //Get unique item from DescriptionToPerm 
            List<string> listUniqueDesc = new List<string>();
            listUniqueDesc = GetUniqueDescriptionToPerm(sodSoxRox.ListDescriptionToPerm);

            //Get list of roles
            List<string> listRoles = new List<string>();
            listRoles = GetMap02Roles(sodSoxRox.ListRolePerm);

            //we only check acccess level Create, Edit, Full
            List<string> listAccessLevel = new List<string> { "Create", "Edit", "Full" };

            foreach (var itemRole1 in listRoles) //loop through role A
            {
                if (!itemRole1.Equals("Category") && !itemRole1.Equals("Permission")) //exclude category and permission. First and second column in map02
                {
                    List<string> listPermA = new List<string>();
                    int roleIndexMain = listRoles.FindIndex(x => x.Equals(itemRole1));
                    if (roleIndexMain > 0)
                    {
                        foreach (var item in sodSoxRox.ListRolePerm)
                        {
                            if (listAccessLevel.Contains(item.ListPerm[roleIndexMain].Value) && listUniqueDesc.Contains(item.Permission))
                            {
                                listPermA.Add(item.Permission);
                            }
                        }
                    }

                    foreach (var itemPerm in listPermA)
                    {
                        
                            //NS Perm A and NS Perm B
                            var listConflict = sodSoxRox.ListConflictPerm.Where(x => x.NsPermPairA.Contains(itemPerm) || x.NsPermPairB.Contains(itemPerm));
                            if (listConflict.Count() > 0)
                            {
                                foreach (var itemConflictPerm in listConflict)
                                {
                                    //Get NS PermA and NS PermB in Map_03_ConflictPerm
                                    listNsPermA = itemConflictPerm.NsPermPairA.Split('|').ToList();
                                    listNsPermB = itemConflictPerm.NsPermPairB.Split('|').ToList();
                                    
                                    if (listNsPermA.Contains(itemPerm)) //if permission is found in A, then loop through NSPermB
                                    {
                                        foreach (var itemPermB in listNsPermB)
                                        {
                                            if (listUniqueDesc.Contains(itemPermB))
                                            {
                                                //search for permission in roles
                                                foreach (var item in sodSoxRox.ListRolePerm.Where(x => x.Permission.Equals(itemPermB)))
                                                {
                                                    foreach (var itemRolePerm in item.ListPerm.Where(x => listAccessLevel.Contains(x.Value)))
                                                    {
                                                        //Debug.WriteLine($"Name: {itemUser.Name} | Role: {itemRole2} | Index: {roleIndex} | Permission: {itemPerm.Permission} | Access: {accessLevel} | Row: {itemPerm.Row} | Column: {itemPerm.ListPerm[roleIndex].Column}");
                                                        SODSoxRoxReportRaw3 sodSoxRoxReportRaw3 = new SODSoxRoxReportRaw3();
                                                        sodSoxRoxReportRaw3.SODRef = itemConflictPerm.RefNum;
                                                        sodSoxRoxReportRaw3.RoleA = itemRole1;
                                                        sodSoxRoxReportRaw3.RoleB = itemRolePerm.Header;
                                                        sodSoxRoxReportRaw3.PermA = itemPerm;
                                                        sodSoxRoxReportRaw3.PermB = itemPermB;
                                                        sodSoxRoxReportRaw3.ProcessA = itemConflictPerm.ProcessA;
                                                        sodSoxRoxReportRaw3.ProcessB = itemConflictPerm.ProcessB;
                                                        sodSoxRoxReportRaw3.DescOfConflict = itemConflictPerm.DescriptionOfConflict;
                                                        sodSoxRoxReportRaw3.DescriptionA = itemConflictPerm.SODDescriptionA;
                                                        sodSoxRoxReportRaw3.DescriptionB = itemConflictPerm.SODDescriptionB;
                                                        sodSoxRoxReportRaw3.RiskComments = itemConflictPerm.RiskComments;
                                                        switch (itemConflictPerm.System)
                                                        {
                                                            case string s when s.ToLower().Equals("low"):
                                                                sodSoxRoxReportRaw3.SODPriority = "3-Low";
                                                                break;
                                                            case string s when s.ToLower().Equals("med") || s.ToLower().Equals("medium"):
                                                                sodSoxRoxReportRaw3.SODPriority = "2-Medium";
                                                                break;
                                                            case string s when s.ToLower().Equals("hig") || s.ToLower().Equals("high"):
                                                                sodSoxRoxReportRaw3.SODPriority = "1-High";
                                                                break;
                                                            default:
                                                                sodSoxRoxReportRaw3.SODPriority = "4-no";
                                                                break;
                                                        };
                                                        sodSoxRoxReportRaw3.ConflictRolePair = $"{itemRole1} and {itemRolePerm.Header}";
                                                        sodSoxRoxReportRaw3.ConflictType = (itemRole1.Equals(itemRolePerm.Header) ? "Self_Conflict" : "Conflict_Roles");
                                                        sodSoxRoxReportRaw3.AssignedA = sodSoxRox.ListRoleUser.Where(x => x.Role.Equals(itemRole1)).FirstOrDefault() != null ? "YES" : string.Empty;
                                                        sodSoxRoxReportRaw3.AssignedB = sodSoxRox.ListRoleUser.Where(x => x.Role.Equals(itemRolePerm.Header)).FirstOrDefault() != null ? "YES" : string.Empty;

                                                        var checkExisting = listSoxRoxRaw3.Any(x =>
                                                            x.SODRef.Equals(sodSoxRoxReportRaw3.SODRef) &&
                                                            x.RoleA.Equals(sodSoxRoxReportRaw3.RoleA) &&
                                                            x.RoleB.Equals(sodSoxRoxReportRaw3.RoleB) &&
                                                            x.PermA.Equals(sodSoxRoxReportRaw3.PermA) &&
                                                            x.PermB.Equals(sodSoxRoxReportRaw3.PermB) &&
                                                            x.ProcessA.Equals(sodSoxRoxReportRaw3.ProcessA) &&
                                                            x.ProcessB.Equals(sodSoxRoxReportRaw3.ProcessB) &&
                                                            x.SODPriority.Equals(sodSoxRoxReportRaw3.SODPriority) &&
                                                            x.ConflictType.Equals(sodSoxRoxReportRaw3.ConflictType)
                                                        );
                                                        //listSoxRoxRaw3.Add(sodSoxRoxReportRaw3);
                                                        if (!checkExisting)
                                                        {
                                                            listSoxRoxRaw3.Add(sodSoxRoxReportRaw3);
                                                        }
                                                    }
                                                }

                                            }
                                        }
                                    }

                                    if(listNsPermB.Contains(itemPerm)) //if permission is found in B, then loop through NSPermA
                                    {
                                        foreach (var itemPermA in listNsPermA)
                                        {
                                            if (listUniqueDesc.Contains(itemPermA))
                                            {
                                                //search for permission in roles
                                                foreach (var item in sodSoxRox.ListRolePerm.Where(x => x.Permission.Equals(itemPermA)))
                                                {
                                                    foreach (var itemRolePerm in item.ListPerm.Where(x => listAccessLevel.Contains(x.Value)))
                                                    {
                                                        //Debug.WriteLine($"Name: {itemUser.Name} | Role: {itemRole2} | Index: {roleIndex} | Permission: {itemPerm.Permission} | Access: {accessLevel} | Row: {itemPerm.Row} | Column: {itemPerm.ListPerm[roleIndex].Column}");
                                                        SODSoxRoxReportRaw3 sodSoxRoxReportRaw3 = new SODSoxRoxReportRaw3();
                                                        sodSoxRoxReportRaw3.SODRef = itemConflictPerm.RefNum;
                                                        sodSoxRoxReportRaw3.RoleA = itemRole1;
                                                        sodSoxRoxReportRaw3.RoleB = itemRolePerm.Header;
                                                        sodSoxRoxReportRaw3.PermA = itemPermA;
                                                        sodSoxRoxReportRaw3.PermB = itemPerm;
                                                        sodSoxRoxReportRaw3.ProcessA = itemConflictPerm.ProcessA;
                                                        sodSoxRoxReportRaw3.ProcessB = itemConflictPerm.ProcessB;
                                                        sodSoxRoxReportRaw3.DescOfConflict = itemConflictPerm.DescriptionOfConflict;
                                                        sodSoxRoxReportRaw3.DescriptionA = itemConflictPerm.SODDescriptionA;
                                                        sodSoxRoxReportRaw3.DescriptionB = itemConflictPerm.SODDescriptionB;
                                                        sodSoxRoxReportRaw3.RiskComments = itemConflictPerm.RiskComments;
                                                        switch (itemConflictPerm.System)
                                                        {
                                                            case string s when s.ToLower().Equals("low"):
                                                                sodSoxRoxReportRaw3.SODPriority = "3-Low";
                                                                break;
                                                            case string s when s.ToLower().Equals("med") || s.ToLower().Equals("medium"):
                                                                sodSoxRoxReportRaw3.SODPriority = "2-Medium";
                                                                break;
                                                            case string s when s.ToLower().Equals("hig") || s.ToLower().Equals("high"):
                                                                sodSoxRoxReportRaw3.SODPriority = "1-High";
                                                                break;
                                                            default:
                                                                sodSoxRoxReportRaw3.SODPriority = "4-no";
                                                                break;
                                                        };
                                                        sodSoxRoxReportRaw3.ConflictRolePair = $"{itemRole1} and {itemRolePerm.Header}";
                                                        sodSoxRoxReportRaw3.ConflictType = (itemRole1.Equals(itemRolePerm.Header) ? "Self_Conflict" : "Conflict_Roles");
                                                        sodSoxRoxReportRaw3.AssignedA = sodSoxRox.ListRoleUser.Where(x => x.Role.Equals(itemRole1)).FirstOrDefault() != null ? "YES" : string.Empty;
                                                        sodSoxRoxReportRaw3.AssignedB = sodSoxRox.ListRoleUser.Where(x => x.Role.Equals(itemRolePerm.Header)).FirstOrDefault() != null ? "YES" : string.Empty;

                                                        var checkExisting = listSoxRoxRaw3.Any(x =>
                                                            x.SODRef.Equals(sodSoxRoxReportRaw3.SODRef) &&
                                                            x.RoleA.Equals(sodSoxRoxReportRaw3.RoleA) &&
                                                            x.RoleB.Equals(sodSoxRoxReportRaw3.RoleB) &&
                                                            x.PermA.Equals(sodSoxRoxReportRaw3.PermA) &&
                                                            x.PermB.Equals(sodSoxRoxReportRaw3.PermB) &&
                                                            x.ProcessA.Equals(sodSoxRoxReportRaw3.ProcessA) &&
                                                            x.ProcessB.Equals(sodSoxRoxReportRaw3.ProcessB) &&
                                                            x.SODPriority.Equals(sodSoxRoxReportRaw3.SODPriority) &&
                                                            x.ConflictType.Equals(sodSoxRoxReportRaw3.ConflictType)
                                                        );
                                                        //listSoxRoxRaw3.Add(sodSoxRoxReportRaw3);
                                                        if (!checkExisting)
                                                        {
                                                            listSoxRoxRaw3.Add(sodSoxRoxReportRaw3);
                                                        }

                                                    }
                                                }

                                            }
                                        }
                                    }
                                }
                            }
                       
                    }

                }
            }

            return Task.FromResult(listSoxRoxRaw3);
        }


        public Task<List<SODSoxRoxReportRaw3>> ProcessSoxRoxDataRaw3_3(SodSoxRoxInput sodSoxRox)
        {
            List<SODSoxRoxReportRaw3> listSoxRoxRaw3 = new List<SODSoxRoxReportRaw3>();
            List<Perm> listPerm = new List<Perm>();

            List<string> listNsPermA = new List<string>();
            List<string> listNsPermB = new List<string>();

            //Get unique item from DescriptionToPerm 
            List<string> listUniqueDesc = new List<string>();
            listUniqueDesc = GetUniqueDescriptionToPerm(sodSoxRox.ListDescriptionToPerm);

            //Get list of roles
            List<string> listRoles = new List<string>();
            listRoles = GetMap02Roles(sodSoxRox.ListRolePerm);

            //we only check acccess level Create, Edit, Full
            List<string> listAccessLevel = new List<string> { "Create", "Edit", "Full" };

            foreach (var itemRole1 in listRoles) //loop through role A
            {

                foreach (var itemPermA in sodSoxRox.ListRolePerm) //loop list perm A
                {

                    var accessLevelA = itemPermA.ListPerm.Where(x => x.Header.Equals(itemRole1)).Select(y => y.Value).FirstOrDefault() ?? "None"; //Get Access level for this role
                    if (listUniqueDesc.Contains(itemPermA.Permission) && listAccessLevel.Contains(accessLevelA)) //check if itemPermA exists in DescriptionToPerm and access level is full, edit, create
                    {
                     
                        foreach (var itemRole2 in listRoles) //loop through role B
                        {
                            foreach (var itemPermB in sodSoxRox.ListRolePerm) // loop list perm B
                            {
                                var accessLevelB = itemPermB.ListPerm.Where(x => x.Header.Equals(itemRole2)).Select(y => y.Value).FirstOrDefault() ?? "None"; //Get Access level for this role
                                if (listUniqueDesc.Contains(itemPermB.Permission) && listAccessLevel.Contains(accessLevelB)) //check if itemPermB exists in DescriptionToPerm
                                {

                                    //check permission in ConflictToPerm
                                    var listConflict = sodSoxRox.ListConflictPerm.Where(x => x.NsPermPairA.Contains(itemPermA.Permission) && x.NsPermPairB.Contains(itemPermB.Permission));
                                    if(listConflict.Any())
                                    {
                                        foreach (var itemConflictPerm in listConflict)
                                        {
                                            //Get NS PermA and NS PermB in Map_03_ConflictPerm
                                            listNsPermA = itemConflictPerm.NsPermPairA.Split('|').ToList();
                                            listNsPermB = itemConflictPerm.NsPermPairB.Split('|').ToList();

                                            if (listNsPermA.Contains(itemPermA.Permission) && listNsPermB.Contains(itemPermB.Permission)) //if permission is found in A and B
                                            {
                                                //Debug.WriteLine($"Name: {itemUser.Name} | Role: {itemRole2} | Index: {roleIndex} | Permission: {itemPerm.Permission} | Access: {accessLevel} | Row: {itemPerm.Row} | Column: {itemPerm.ListPerm[roleIndex].Column}");
                                                SODSoxRoxReportRaw3 sodSoxRoxReportRaw3 = new SODSoxRoxReportRaw3();
                                                sodSoxRoxReportRaw3.SODRef = itemConflictPerm.RefNum;
                                                sodSoxRoxReportRaw3.RoleA = itemRole1;
                                                sodSoxRoxReportRaw3.RoleB = itemRole2;
                                                sodSoxRoxReportRaw3.PermA = itemPermA.Permission;
                                                sodSoxRoxReportRaw3.PermB = itemPermB.Permission;
                                                sodSoxRoxReportRaw3.ProcessA = itemConflictPerm.ProcessA;
                                                sodSoxRoxReportRaw3.ProcessB = itemConflictPerm.ProcessB;
                                                sodSoxRoxReportRaw3.DescOfConflict = itemConflictPerm.DescriptionOfConflict;
                                                sodSoxRoxReportRaw3.DescriptionA = itemConflictPerm.SODDescriptionA;
                                                sodSoxRoxReportRaw3.DescriptionB = itemConflictPerm.SODDescriptionB;
                                                sodSoxRoxReportRaw3.RiskComments = itemConflictPerm.RiskComments;
                                                if(itemConflictPerm.System != string.Empty && itemConflictPerm.System != null)
                                                {
                                                    switch (itemConflictPerm.System)
                                                    {
                                                        case string s when s.ToLower().Equals("low"):
                                                            sodSoxRoxReportRaw3.SODPriority = "3-Low";
                                                            break;
                                                        case string s when s.ToLower().Equals("med") || s.ToLower().Equals("medium"):
                                                            sodSoxRoxReportRaw3.SODPriority = "2-Medium";
                                                            break;
                                                        case string s when s.ToLower().Equals("hig") || s.ToLower().Equals("high"):
                                                            sodSoxRoxReportRaw3.SODPriority = "1-High";
                                                            break;
                                                        default:
                                                            sodSoxRoxReportRaw3.SODPriority = "4-no";
                                                            break;
                                                    };
                                                }
                                                else
                                                {
                                                    switch (itemConflictPerm.Manual)
                                                    {
                                                        case string s2 when s2.ToLower().Equals("low"):
                                                            sodSoxRoxReportRaw3.SODPriority = "3-Low";
                                                            break;
                                                        case string s2 when s2.ToLower().Equals("med") || s2.ToLower().Equals("medium"):
                                                            sodSoxRoxReportRaw3.SODPriority = "2-Medium";
                                                            break;
                                                        case string s2 when s2.ToLower().Equals("hig") || s2.ToLower().Equals("high"):
                                                            sodSoxRoxReportRaw3.SODPriority = "1-High";
                                                            break;
                                                        default:
                                                            sodSoxRoxReportRaw3.SODPriority = "4-no";
                                                            break;

                                                    }
                                                }
                                                
                                                sodSoxRoxReportRaw3.ConflictRolePair = $"{itemRole1} and {itemRole2}";
                                                sodSoxRoxReportRaw3.ConflictType = (itemRole1.Equals(itemRole2) ? "Self_Conflict" : "Conflict_Roles");
                                                sodSoxRoxReportRaw3.AssignedA = sodSoxRox.ListRoleUser.Where(x => x.Role.Equals(itemRole1)).FirstOrDefault() != null ? "YES" : string.Empty;
                                                sodSoxRoxReportRaw3.AssignedB = sodSoxRox.ListRoleUser.Where(x => x.Role.Equals(itemRole2)).FirstOrDefault() != null ? "YES" : string.Empty;


                                                //Console.WriteLine($"==========================================================================");
                                                //Console.WriteLine($"sodSoxRoxReportRaw3.SODRef: {sodSoxRoxReportRaw3.SODRef} ");
                                                //Console.WriteLine($"sodSoxRoxReportRaw3.RoleA: {sodSoxRoxReportRaw3.RoleA} ");
                                                //Console.WriteLine($"sodSoxRoxReportRaw3.RoleB: {sodSoxRoxReportRaw3.RoleB} ");
                                                //Console.WriteLine($"sodSoxRoxReportRaw3.PermA: {sodSoxRoxReportRaw3.PermA} ");
                                                //Console.WriteLine($"sodSoxRoxReportRaw3.PermB: {sodSoxRoxReportRaw3.PermB} ");
                                                //Console.WriteLine($"sodSoxRoxReportRaw3.ProcessA: {sodSoxRoxReportRaw3.ProcessA} ");
                                                //Console.WriteLine($"sodSoxRoxReportRaw3.ProcessB: {sodSoxRoxReportRaw3.ProcessB} ");
                                                //Console.WriteLine($"sodSoxRoxReportRaw3.SODPriority: {sodSoxRoxReportRaw3.SODPriority} ");
                                                //Console.WriteLine($"sodSoxRoxReportRaw3.ConflictType: {sodSoxRoxReportRaw3.ConflictType} ");
                                
                                            

                                                var checkExisting = listSoxRoxRaw3.Any(x =>
                                                    x.SODRef.Equals(sodSoxRoxReportRaw3.SODRef) &&
                                                    x.RoleA.Equals(sodSoxRoxReportRaw3.RoleA) &&
                                                    x.RoleB.Equals(sodSoxRoxReportRaw3.RoleB) &&
                                                    x.PermA.Equals(sodSoxRoxReportRaw3.PermA) &&
                                                    x.PermB.Equals(sodSoxRoxReportRaw3.PermB) &&
                                                    x.ProcessA.Equals(sodSoxRoxReportRaw3.ProcessA) &&
                                                    x.ProcessB.Equals(sodSoxRoxReportRaw3.ProcessB) &&
                                                    x.SODPriority.Equals(sodSoxRoxReportRaw3.SODPriority) &&
                                                    x.ConflictType.Equals(sodSoxRoxReportRaw3.ConflictType)
                                                );
                                                //listSoxRoxRaw3.Add(sodSoxRoxReportRaw3);
                                                if (!checkExisting)
                                                {
                                                    listSoxRoxRaw3.Add(sodSoxRoxReportRaw3);
                                                }
                                            }

                                            
                                        }
                                    }
                                

                                }
                            }
                        }
                    
                    }

                    
                }
                
            }

            return Task.FromResult(listSoxRoxRaw3);
        }


        


        public Task<string> CreateSodSoxRoxFile(List<SODSoxRoxReportRaw2> sodRaw2, List<SODSoxRoxReportRaw3> sodRaw3, string clientName)
        {
            //string fileName = string.Empty;

            ExcelService xlsService = new ExcelService();
            WriteLog writeLog = new WriteLog();
            // If you use EPPlus in a noncommercial context
            // according to the Polyform Noncommercial license:
            ////ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            using (ExcelPackage p = new ExcelPackage())
            {

                #region Tab 1
                var ws = p.Workbook.Worksheets.Add("_2_Raw_2");
                int row = 1;

                //Title
                ws.Cells[row, 1].Value = "Name";
                ws.Cells[row, 2].Value = "Email";
                ws.Cells[row, 3].Value = "Phone";
                ws.Cells[row, 4].Value = "SOD_Ref";
                ws.Cells[row, 5].Value = "Role_A";
                ws.Cells[row, 6].Value = "Role_B";
                ws.Cells[row, 7].Value = "Perm_A";
                ws.Cells[row, 8].Value = "Perm_B";
                ws.Cells[row, 9].Value = "Process_A_Descrpt";
                ws.Cells[row, 10].Value = "Process_B_Descrpt";
                ws.Cells[row, 11].Value = "FunctionType_A";
                ws.Cells[row, 12].Value = "FunctionType_B";
                ws.Cells[row, 13].Value = "Description_of_Conflict";
                ws.Cells[row, 14].Value = "Description_A";
                ws.Cells[row, 15].Value = "Description_B";
                ws.Cells[row, 16].Value = "Risks_Comments";
                ws.Cells[row, 17].Value = "SOD_priority";
                ws.Cells[row, 18].Value = "Conflicting_Role_Pair";
                ws.Cells[row, 19].Value = "Conflict_Type";
                row++;

                foreach (var item in sodRaw2)
                {
                    ws.Cells[row, 1].Value = item.Name; //"Name";
                    ws.Cells[row, 2].Value = item.Email; //"Email";
                    ws.Cells[row, 3].Value = item.Phone; //"Phone";
                    ws.Cells[row, 4].Value = item.SODRef; //"SOD_Ref";
                    ws.Cells[row, 5].Value = item.RoleA; //"Role_A";
                    ws.Cells[row, 6].Value = item.RoleB; //"Role_B";
                    ws.Cells[row, 7].Value = item.PermA; //"Perm_A";
                    ws.Cells[row, 8].Value = item.PermB; //"Perm_B";
                    ws.Cells[row, 9].Value = item.ProcessA; //"Process_A_Descrpt";
                    ws.Cells[row, 10].Value = item.ProcessB; //"Process_B_Descrpt";
                    ws.Cells[row, 11].Value = item.FunctionTypeA; //"FunctionType_A";
                    ws.Cells[row, 12].Value = item.FunctionTypeB; //"FunctionType_B";
                    ws.Cells[row, 13].Value = item.DescOfConflict; //"Description_of_Conflict";
                    ws.Cells[row, 14].Value = item.DescriptionA; //"Description_A";
                    ws.Cells[row, 15].Value = item.DescriptionB; //"Description_B";
                    ws.Cells[row, 16].Value = item.RiskComments; //"Risks_Comments";
                    ws.Cells[row, 17].Value = item.SODPriority; //"SOD_priority";
                    ws.Cells[row, 18].Value = item.ConflictRolePair; //"Conflicting_Role_Pair";
                    ws.Cells[row, 19].Value = item.ConflictType; //"Conflict_Type";
                    row++;
                }

                #endregion

                #region Tab 2
                var ws2 = p.Workbook.Worksheets.Add("_3_Raw_2");
                row = 1;

                //Title
                ws2.Cells[row, 1].Value = "SOD_Ref";
                ws2.Cells[row, 2].Value = "ROLE_A";
                ws2.Cells[row, 3].Value = "ROLE_B";
                ws2.Cells[row, 4].Value = "Perm_A";
                ws2.Cells[row, 5].Value = "Perm_B";
                ws2.Cells[row, 6].Value = "Process_A_Descrpt";
                ws2.Cells[row, 7].Value = "Process_B_Descrpt";
                ws2.Cells[row, 8].Value = "Description_of_Conflict";
                ws2.Cells[row, 9].Value = "Description_A";
                ws2.Cells[row, 10].Value = "Description_B";
                ws2.Cells[row, 11].Value = "Risks_Comments";
                ws2.Cells[row, 12].Value = "SOD_priority";
                ws2.Cells[row, 13].Value = "Conflicting_Role_Pair";
                ws2.Cells[row, 14].Value = "Conflict_Type";
                ws2.Cells[row, 15].Value = "Assigned_A";
                ws2.Cells[row, 16].Value = "Assigned_B";
                row++;

                foreach (var item in sodRaw3)
                {
                    ws2.Cells[row, 1].Value = item.SODRef; //"SOD_Ref";
                    ws2.Cells[row, 2].Value = item.RoleA; //"ROLE_A";
                    ws2.Cells[row, 3].Value = item.RoleB; //"ROLE_B";
                    ws2.Cells[row, 4].Value = item.PermA; //"Perm_A";
                    ws2.Cells[row, 5].Value = item.PermB; //"Perm_B";
                    ws2.Cells[row, 6].Value = item.ProcessA; //"Process_A_Descrpt";
                    ws2.Cells[row, 7].Value = item.ProcessB; //"Process_B_Descrpt";
                    ws2.Cells[row, 8].Value = item.DescOfConflict; //"Description_of_Conflict";
                    ws2.Cells[row, 9].Value = item.DescriptionA; //"Description_A";
                    ws2.Cells[row, 10].Value = item.DescriptionB; //"Description_B";
                    ws2.Cells[row, 11].Value = item.RiskComments; //"Risks_Comments";
                    ws2.Cells[row, 12].Value = item.SODPriority; //"SOD_priority";
                    ws2.Cells[row, 13].Value = item.ConflictRolePair; //"Conflicting_Role_Pair";
                    ws2.Cells[row, 14].Value = item.ConflictType; //"Conflict_Type";
                    ws2.Cells[row, 15].Value = item.AssignedA; //"Assigned_A";
                    ws2.Cells[row, 16].Value = item.AssignedB; //"Assigned_B";
                    row++;
                }
                #endregion

                string startupPath = Directory.GetCurrentDirectory();
                string strSourceDownload = Path.Combine(startupPath, "include", "sod");
                if (!Directory.Exists(strSourceDownload))
                {
                    Directory.CreateDirectory(strSourceDownload);
                }
                var ts = DateTime.Now.ToString("yyyyMMdd_HHmmss");
                string filename = $"{ts}_{clientName}_SOD_Report.xlsx";
                string strOutput = Path.Combine(strSourceDownload, filename);

                //Check if file not exists
                if (System.IO.File.Exists(strOutput))
                {
                    System.IO.File.Delete(strOutput);
                }

                p.SaveAs(new FileInfo(strOutput));
                return Task.FromResult(filename);
            }


            //return Task.FromResult(fileName);
        }

        public List<string> GetUniqueDescriptionToPerm(List<DescriptionToPerm> ListDescriptionToPerm)
        {
            List<string> listUniqueDesc = new List<string>();

            foreach (var item in ListDescriptionToPerm)
            {
                if(item.Permission != string.Empty)
                {
                    //split string that has character | and store in list
                    //sample permission = "Purchase Order|Blanket Purchase Order|Requisition|Purchase Contract"
                    var listPerm = item.Permission.Split('|').ToList();
                    foreach (var perm in listPerm)
                    {
                        //if permission is not added in list, we add the item in list
                        if(!listUniqueDesc.Contains(perm))
                        {
                            listUniqueDesc.Add(perm);
                        }
                    }
                }
            }

            return listUniqueDesc;
        }

        /// <summary>
        /// Extract all roles in RolePerm
        /// </summary>
        /// <param name="ListRolePerm">List<RolePerm></param>
        /// <returns>List<string></returns>
        public List<string> GetMap02Roles(List<RolePerm> ListRolePerm)
        {
            List<string> listRole = new List<string>();

            foreach (var item in ListRolePerm)
            {
                foreach (var itemPerm in item.ListPerm)
                {
                    listRole.Add(itemPerm.Header);
                }
                break;
            }

            return listRole;
        }

        private Task<ServiceResponse> CheckConflict(List<SodUser> listSodUser, string duty1, string duty2)
        {
            ServiceResponse res = new ServiceResponse();
            var allDuty1 = listSodUser.Where(x => x.DutyNum.Equals(duty1)).FirstOrDefault();
            var allDuty2 = listSodUser.Where(x => x.DutyNum.Equals(duty2)).FirstOrDefault();
            List<string> listUser1 = new List<string>();
            List<string> listUser2 = new List<string>();
            List<string> listConflict = new List<string>();
            string conflictUsers = string.Empty;
            bool hasConflict = false;

            if (allDuty1.Users != string.Empty)
            {
                listUser1 = allDuty1.Users.Split(",").Select(x => x.TrimStart(' ')).ToList();
            }

            if (allDuty2.Users != string.Empty)
            {
                listUser2 = allDuty2.Users.Split(",").Select(x => x.TrimStart(' ')).ToList();
            }

            if (listUser1.Count > 0 && listUser2.Count > 0)
            {                
                listConflict = listUser1.Intersect(listUser2).ToList();
                conflictUsers = string.Join(",", listConflict);
            }

            if (conflictUsers != string.Empty)
            {
                hasConflict = true;
            }
            res.Status = hasConflict;
            res.Content = conflictUsers;
            //return (hasConflict, conflictUsers);
            return Task.FromResult(res);
        }
    
    }
}
