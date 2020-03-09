using etaxOneth_Process.ControlAPI;
using etaxOneth_Process.DataModel;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace etaxOneth_Process
{
    public partial class etaxOnethProcess : Form
    {
        int mouseX = 0;
        int mouseY = 0;
        bool mouseDown;
        int iCountRunFile;
        DtGetParameters dtParam = new DtGetParameters();
        ManageAPI conAPIClass = new ManageAPI();
        DataOutput strOutputFile = new DataOutput();
        string strTempSumAmount = string.Empty;
        List<string> lstTempSumAmount = new List<string>();
        string strTempLogTime = string.Empty;
        int iSumFail = 0;
        int iSumSuccess = 0;
        bool chkOption = false;
        string strDateTimeStamp = string.Empty;
        System.Reflection.Assembly assembly = System.Reflection.Assembly.GetExecutingAssembly();
        //string strTempPath = AppDomain.CurrentDomain.BaseDirectory + "Automate\\Input\\Temp";
        //string strInputPath = AppDomain.CurrentDomain.BaseDirectory + "Automate\\Input";
        //string strLogPath = AppDomain.CurrentDomain.BaseDirectory + "Automate\\Result\\LogTime";
        //string strFailSourcePath = AppDomain.CurrentDomain.BaseDirectory + "Automate\\Result\\Fail\\Source";
        //string strFailErrPath = AppDomain.CurrentDomain.BaseDirectory + "Automate\\Result\\Fail\\Error";
        //string strSuccessSourcePath = AppDomain.CurrentDomain.BaseDirectory + "Automate\\Result\\Success\\Source";
        //string strSuccessOutputPath = AppDomain.CurrentDomain.BaseDirectory + "Automate\\Result\\Success\\Output";
        //string strLogFileRunPath = AppDomain.CurrentDomain.BaseDirectory + "Automate\\Input\\LogFileRun";
        //string strFileRunPath = AppDomain.CurrentDomain.BaseDirectory + "Automate\\Input\\FileRun";

        Dictionary<string, string> DocType = new Dictionary<string, string>()
        {
            { "ใบกำกับภาษี", "388" },
            { "ใบแจ้งหนี้/ใบกำกับภาษี", "T02" },
            { "ใบเสร็จรับเงิน/ใบกำกับภาษี", "T03" },
            { "ใบส่งของ/ใบกำกับภาษี", "T04" },
            { "ใบรับ(ใบเสร็จรับเงิน)", "T01" },
            { "ใบเพิ่มหนี้", "80" },
            { "ใบลดหนี้", "81" },
            { "ใบแจ้งหนี้", "380" }
        };

        Dictionary<string, string> BuyerTaxType = new Dictionary<string, string>()
        {
            { "1", "TXID" },
            { "2", "NIDN" },
            { "3", "CCPT" },
            { "4", "OTHR" }
        };

        Dictionary<string, string> Month = new Dictionary<string, string>()
        {
            { "Jan","01" },
            { "Feb","02" },
            { "Mar","03" },
            { "Apr","04" },
            { "May","05" },
            { "Jun","06" },
            { "Jul","07" },
            { "Aug","08" },
            { "Sep","09" },
            { "Oct","10" },
            { "Nov","11" },
            { "Dec","12" }
        };

        Dictionary<string, string> Day = new Dictionary<string, string>()
        {
            { "1","01" },
            { "2","02" },
            { "3","03" },
            { "4","04" },
            { "5","05" },
            { "6","06" },
            { "7","07" },
            { "8","08" },
            { "9","09" },
            { "10","10" },
            { "11","11" },
            { "12","12" },
            { "13","13" },
            { "14","14" },
            { "15","15" },
            { "16","16" },
            { "17","17" },
            { "18","18" },
            { "19","19" },
            { "20","20" },
            { "21","21" },
            { "22","22" },
            { "23","23" },
            { "24","24" },
            { "25","25" },
            { "26","26" },
            { "27","27" },
            { "28","28" },
            { "29","29" },
            { "30","30" },
            { "31","31" },
        };

        public etaxOnethProcess()
        {
            InitializeComponent();
            
        }

        private void pbClose_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }

        private void pbRestoreDown_Click(object sender, EventArgs e)
        {
            this.WindowState = FormWindowState.Normal;
            pbRestoreDown.Visible = false;
            pbMaximize.Visible = true;
        }

        private void pbMaximize_Click(object sender, EventArgs e)
        {
            this.WindowState = FormWindowState.Maximized;
            pbMaximize.Visible = false;
            pbRestoreDown.Visible = true;
        }

        private void pbMinimize_Click(object sender, EventArgs e)
        {
            this.WindowState = FormWindowState.Minimized;
        }

        public void StopWorking(string AmountFile)
        {
            iCountRunFile = Int32.Parse(AmountFile) + 1;
        }

        public ValueReturnForm RunProcess(PathFilesIO pathFileIO, string strDateTime, string strSellerTaxID, string strBranchID, string strAPIKey, string strUserCode, string strAccessKey, string strServiceCode, string strAmountFile, string strServiceURL, bool firstStatus, ValueReturnForm valueReturnChk)
        {
            strDateTimeStamp = strDateTime;
            chkOption = false;
            dtParam = new DtGetParameters();
            dtParam.PathInput = pathFileIO.PathInput;
            dtParam.PathOutput = pathFileIO.PathOutput;
            dtParam.SellerTaxID = strSellerTaxID;
            dtParam.BranchID = strBranchID;
            dtParam.APIKey = strAPIKey;
            dtParam.UserCode = strUserCode;
            dtParam.AccessKey = strAccessKey;
            dtParam.ServiceCode = strServiceCode;
            dtParam.AmountFile = strAmountFile;
            dtParam.ServiceURL = strServiceURL;
            valueReturnChk.StatusFindPDF = true;

            try
            {
                RefreshForm();

                if (pathFileIO.TypeRunning.Equals("M"))
                {
                    chkOption = true;
                    string strFileName = Path.GetFileNameWithoutExtension(pathFileIO.PathInput);
                    string strFileNameEx = Path.GetFileName(pathFileIO.PathInput);
                    int iFail = 0;
                    string[] arrFilesPDF = System.IO.Directory.GetFiles(Path.GetDirectoryName(pathFileIO.PathInput), "*.pdf");

                    if (strServiceCode.Equals("S03"))
                    {
                        WorkProcess(pathFileIO, strFileName, strFileNameEx, "", out iFail);
                    }
                    else
                    {
                        var lookupName = arrFilesPDF.ToLookup(x => Path.GetFileNameWithoutExtension(x));

                        if (lookupName.Contains(strFileName))
                        {
                            var resultJoin = lookupName[strFileName];
                            foreach (var itemPDF in resultJoin)
                            {
                                WorkProcess(pathFileIO, strFileName, strFileNameEx, itemPDF, out iFail);
                            }
                        }
                        else
                        {
                            CreateTextFile(pathFileIO.PathOutput + "\\LogPDFNotFound_" + strDateTimeStamp + ".txt", "Not Found PDF File: " + strFileNameEx);
                            //MessageBox.Show("Not Found PDF File!");
                            valueReturnChk.StatusFindPDF = false;
                            return valueReturnChk;
                        }
                    }

                    if (iFail > 0)
                    {
                        iSumFail++;
                    }
                    else
                    {
                        iSumSuccess++;
                        string pathPrint = dtParam.PathOutput + "\\" + strFileName + "_" + strDateTimeStamp + "_PDF.pdf";
                        valueReturnChk.pathPrint = pathPrint;
                    }
                }
                else
                {
                    string[] arrFilesXlsx = System.IO.Directory.GetFiles(pathFileIO.PathInput, "*.xlsx");
                    string[] arrFilesTxt = System.IO.Directory.GetFiles(pathFileIO.PathInput, "*.txt");
                    string[] arrFilesPDF = System.IO.Directory.GetFiles(pathFileIO.PathInput, "*.pdf");
                    string strFileNameRun = string.Empty;
                    dtParam.PathOutput = string.Empty;
                    dtParam.PathOutput = pathFileIO.PathTemp;
                    iCountRunFile = 1;
                    
                    if (arrFilesXlsx.Count() == 0 && arrFilesTxt.Count() == 0)
                    {
                        valueReturnChk.StatusRunning = true;
                        return valueReturnChk;
                    }

                    if (firstStatus == true)
                    {
                        strFileNameRun = String.Join(",", arrFilesXlsx) + "," + String.Join(",", arrFilesTxt);
                        File.WriteAllText(pathFileIO.PathLogFileRun + "\\LogFileRunPerTime.txt", String.Empty);
                        CreateTextFile(pathFileIO.PathLogFileRun + "\\LogFileRunPerTime.txt", strFileNameRun);
                        //MessageBox.Show(arrFilesXlsx.Count()+"");
                        valueReturnChk.AmountAllFile = arrFilesXlsx.Count() + arrFilesTxt.Count();
                    }
                    string strFileRunning = System.IO.File.ReadAllText(pathFileIO.PathLogFileRun + "\\LogFileRunPerTime.txt");
                    string[] arrSplitFileRunning = strFileRunning.Split(',');
                    var lookupFileLog = arrSplitFileRunning.ToLookup(x => Path.GetFileName(x));
                    //valueReturnChk.CountFileRun += Int32.Parse(dtParam.AmountFile);

                    foreach (var item in arrFilesXlsx)
                    {
                        if (lookupFileLog.Contains(Path.GetFileName(item)))
                        {
                            if (iCountRunFile > Int32.Parse(dtParam.AmountFile))
                            {
                                break;
                            }
                            dtParam.PathInput = string.Empty;
                            dtParam.PathInput = item;
                            string strFileName = Path.GetFileNameWithoutExtension(item);
                            string strFileNameEx = Path.GetFileName(item);
                            int iFail = 0;

                            System.Threading.Thread.Sleep(500);

                            if (strServiceCode.Equals("S03"))
                            {
                                WorkProcess(pathFileIO, strFileName, strFileNameEx, "", out iFail);
                            }
                            else
                            {
                                var lookupName = arrFilesPDF.ToLookup(x => Path.GetFileNameWithoutExtension(x));

                                if (lookupName.Contains(strFileName))
                                {
                                    var resultJoin = lookupName[strFileName];
                                    foreach (var itemPDF in resultJoin)
                                    {
                                        try
                                        {
                                            WorkProcess(pathFileIO, strFileName, strFileNameEx, itemPDF, out iFail);
                                            //File.Move(itemPDF, pathFileIO.PathFileRun + "\\" + Path.GetFileNameWithoutExtension(itemPDF)+"_"+strDateTimeStamp+".pdf");
                                            if (File.Exists(pathFileIO.PathFileRun + "\\" + Path.GetFileName(item)))
                                            {
                                                File.Delete(pathFileIO.PathFileRun + "\\" + Path.GetFileName(item));
                                                File.Move(item, pathFileIO.PathFileRun + "\\" + Path.GetFileName(item));
                                            }
                                            else
                                            {
                                                File.Move(item, pathFileIO.PathFileRun + "\\" + Path.GetFileName(item));
                                            }
                                        }
                                        catch(Exception e)
                                        {
                                            MessageBox.Show("a");
                                        }
                                    }
                                }
                                else
                                {
                                    if (File.Exists(pathFileIO.PathOutput + "\\LogPDFNotFound_" + strDateTimeStamp + ".txt"))
                                    {
                                        // create a filestream w/ append
                                        //File.AppendAllText(pathFileIO.PathOutput + "\\" + pathFileIO.DateTimeFolderName + "\\LogPDFNotFound_" + strDateTimeStamp + ".txt", "Not Found PDF File: " + strFileNameEx + Environment.NewLine);
                                        File.AppendAllText(pathFileIO.PathOutput + "\\LogPDFNotFound_" + strDateTimeStamp + ".txt", "Not Found PDF File: " + strFileNameEx + Environment.NewLine);
                                    }
                                    else
                                    {
                                        // create a filestream for new.
                                        //CreateTextFile(pathFileIO.PathOutput + "\\" + pathFileIO.DateTimeFolderName + "\\LogPDFNotFound_" + strDateTimeStamp + ".txt", "Not Found PDF File: " + strFileNameEx);
                                        CreateTextFile(pathFileIO.PathOutput + "\\LogPDFNotFound_" + strDateTimeStamp + ".txt", "Not Found PDF File: " + strFileNameEx);
                                    }
                                    //File.Move(item, pathFileIO.PathFileRun + "\\" + Path.GetFileNameWithoutExtension(item) + "_" + strDateTimeStamp + ".xlsx");
                                    if (File.Exists(pathFileIO.PathFileRun + "\\" + Path.GetFileName(item)))
                                    {
                                        File.Delete(pathFileIO.PathFileRun + "\\" + Path.GetFileName(item));
                                        File.Move(item, pathFileIO.PathFileRun + "\\" + Path.GetFileName(item));
                                    }
                                    else
                                    {
                                        File.Move(item, pathFileIO.PathFileRun + "\\" + Path.GetFileName(item));
                                    }
                                    continue;
                                }
                            }

                            if (iFail > 0)
                            {
                                iSumFail++;
                            }
                            else
                            {
                                iSumSuccess++;
                            }

                            iCountRunFile++;
                            valueReturnChk.CountFileRun += 1;
                        }
                        try
                        {
                            //File.Move(item, pathFileIO.PathFileRun + "\\" + Path.GetFileNameWithoutExtension(item) + "_" + strDateTimeStamp + ".xlsx");
                            if (File.Exists(pathFileIO.PathFileRun + "\\" + Path.GetFileName(item)))
                            {
                                File.Delete(pathFileIO.PathFileRun + "\\" + Path.GetFileName(item));
                                File.Move(item, pathFileIO.PathFileRun + "\\" + Path.GetFileName(item));
                            }
                            else
                            {
                                File.Move(item, pathFileIO.PathFileRun + "\\" + Path.GetFileName(item));
                            }
                        }
                        catch(Exception e)
                        {
                            MessageBox.Show("b");
                        }
                    }

                    foreach (var item in arrFilesTxt)
                    {
                        if (lookupFileLog.Contains(Path.GetFileName(item)))
                        {
                            if (iCountRunFile > Int32.Parse(dtParam.AmountFile))
                            {
                                break;
                            }
                           
                            dtParam.PathInput = string.Empty;
                            dtParam.PathInput = item;
                            string strFileName = Path.GetFileNameWithoutExtension(item);
                            string strFileNameEx = Path.GetFileName(item);
                            int iFail = 0;

                            System.Threading.Thread.Sleep(500);

                            if (strServiceCode.Equals("S03"))
                            {
                                WorkProcess(pathFileIO, strFileName, strFileNameEx, "", out iFail);
                            }
                            else
                            {
                                var lookupName = arrFilesPDF.ToLookup(x => Path.GetFileNameWithoutExtension(x));
                                //MessageBox.Show(lookupName[strFileName].ToString());
                                if (lookupName.Contains(strFileName))
                                {
                                    var resultJoin = lookupName[strFileName];
                                    foreach (var itemPDF in resultJoin)
                                    {
                                        //MessageBox.Show(itemPDF);
                                        try
                                        {
                                            WorkProcess(pathFileIO, strFileName, strFileNameEx, itemPDF, out iFail);
                                            //File.Move(itemPDF, pathFileIO.PathFileRun + "\\" + Path.GetFileNameWithoutExtension(itemPDF) + "_" + strDateTimeStamp + ".pdf");
                                            if (File.Exists(pathFileIO.PathFileRun + "\\" + Path.GetFileName(itemPDF)))
                                            {
                                                File.Delete(pathFileIO.PathFileRun + "\\" + Path.GetFileName(itemPDF));
                                                File.Move(itemPDF, pathFileIO.PathFileRun + "\\" + Path.GetFileName(itemPDF));
                                            }
                                            else
                                            {
                                                File.Move(itemPDF, pathFileIO.PathFileRun + "\\" + Path.GetFileName(itemPDF));
                                            }
                                        }
                                        catch(Exception e)
                                        {
                                            MessageBox.Show("C");
                                        }
                                    }
                                }
                                else
                                {
                                    if (File.Exists(pathFileIO.PathOutput + "\\LogPDFNotFound_" + strDateTimeStamp + ".txt"))
                                    {
                                        // create a filestream w/ append
                                        //File.AppendAllText(pathFileIO.PathOutput + "\\" + pathFileIO.DateTimeFolderName + "\\LogPDFNotFound_" + strDateTimeStamp + ".txt", "Not Found PDF File: " + strFileNameEx + Environment.NewLine);
                                        File.AppendAllText(pathFileIO.PathOutput + "\\LogPDFNotFound_" + strDateTimeStamp + ".txt", "Not Found PDF File: " + strFileNameEx + Environment.NewLine);
                                    }
                                    else
                                    {
                                        // create a filestream for new.
                                        //CreateTextFile(pathFileIO.PathOutput + "\\" + pathFileIO.DateTimeFolderName + "\\LogPDFNotFound_" + strDateTimeStamp + ".txt", "Not Found PDF File: " + strFileNameEx);
                                        CreateTextFile(pathFileIO.PathOutput + "\\LogPDFNotFound_" + strDateTimeStamp + ".txt", "Not Found PDF File: " + strFileNameEx);
                                    }
                                    try
                                    {
                                        //File.Move(item, pathFileIO.PathFileRun + "\\" + Path.GetFileNameWithoutExtension(item) + "_" + strDateTimeStamp + ".txt");
                                        if (File.Exists(pathFileIO.PathFileRun + "\\" + Path.GetFileName(item)))
                                        {
                                            File.Delete(pathFileIO.PathFileRun + "\\" + Path.GetFileName(item));
                                            File.Move(item, pathFileIO.PathFileRun + "\\" + Path.GetFileName(item));
                                        }
                                        else
                                        {
                                            File.Move(item, pathFileIO.PathFileRun + "\\" + Path.GetFileName(item));
                                        }

                                    }
                                    catch(Exception edwa)
                                    {
                                        MessageBox.Show("D");
                                    }
                                    continue;
                                }
                            }
                           
                            if (iFail > 0)
                            {
                                iSumFail++;
                            }
                            else
                            {
                                iSumSuccess++;
                            }

                            iCountRunFile++;
                            valueReturnChk.CountFileRun += 1;
                        }
                        try
                        {
                            //File.Move(item, pathFileIO.PathFileRun + "\\" + Path.GetFileNameWithoutExtension(item) + "_" + strDateTimeStamp + ".txt");
                            if (File.Exists(pathFileIO.PathFileRun + "\\" + Path.GetFileName(item)))
                            {
                                File.Delete(pathFileIO.PathFileRun + "\\" + Path.GetFileName(item));
                                File.Move(item, pathFileIO.PathFileRun + "\\" + Path.GetFileName(item));
                            }
                            else
                            {
                                File.Move(item, pathFileIO.PathFileRun + "\\" + Path.GetFileName(item));
                            }
                        }
                        catch(Exception e)
                        {
                            MessageBox.Show(e.ToString());
                        }
                    }
                }

                if (valueReturnChk.CountFileRun == valueReturnChk.AmountAllFile)
                {
                    valueReturnChk.StatusRunning = true;
                }
                else
                {
                    valueReturnChk.StatusRunning = false;
                }
                FileVersionInfo fvi = FileVersionInfo.GetVersionInfo(assembly.Location);
                string version = fvi.FileVersion;
                txtStatus.Text += Environment.NewLine + "Summary:" + Environment.NewLine + "   -Success " + iSumSuccess + Environment.NewLine + "   -Fail " + iSumFail;
                txtStatus.Refresh();

                strTempLogTime += Environment.NewLine + "Summary: Success " + iSumSuccess + " Fail " + iSumFail;

                if (chkOption == true)
                {
                    //Write log time
                    CreateTextFile(pathFileIO.PathOutput + "\\LogTime Version." + version + strDateTimeStamp + ".txt", strTempLogTime);
                }
                else if (chkOption == false)
                {
                    //Write log time
                    CreateTextFile(pathFileIO.PathOutput + "\\LogTime Version." + version + strDateTimeStamp + ".txt", strTempLogTime);
                }
            }
            catch (Exception exc)
            {
                MessageBox.Show("Method btnExport_Click Error: " + exc.ToString());
            }
            finally
            {
                iSumFail = 0;
                iSumSuccess = 0;
                strTempLogTime = string.Empty;
                GC.Collect();
                GC.WaitForPendingFinalizers();
            }

            return valueReturnChk;
        }

        public void WorkProcess(PathFilesIO pfIO, string strFileName, string strFileNameExtension, string strFileNamePDF, out int cntFail)
        {
            cntFail = 0;

            try
            {
                string pathText = string.Empty;

                if (Path.GetExtension(dtParam.PathInput).Equals(".xlsx"))
                {
                    pgbLoad.Value = 0;
                    lbPercent.Text = "Export Data: 0%";
                    lbPercent.Refresh();

                    List<string> lstDataRow = new List<string>();
                    List<string> lstDataMenu = new List<string>();
                    string strSheetName = string.Empty;
                    BGroup grpB = new BGroup();
                    CGroup grpC = new CGroup();
                    LGroup grpL = new LGroup();
                    HGroup grpH = new HGroup();

                    var package = new ExcelPackage(new FileInfo(dtParam.PathInput));
                    var oSheetName = package.Workbook.Worksheets.Select(n => n.Name);

                    foreach (var iName in oSheetName)
                    {
                        strSheetName = iName;
                    }

                    ExcelWorksheet workSheet = package.Workbook.Worksheets[strSheetName];
                    var start = workSheet.Dimension.Start;
                    var end = workSheet.Dimension.End;

                    //Keep All
                    for (int row = start.Row; row <= end.Row; row++)
                    { // Row by row...  
                        for (int col = start.Column; col <= end.Column; col++)
                        { // ... Cell by cell...  
                            object cellValue = workSheet.Cells[row, col].Text;
                            //if (cellValue != null && !cellValue.Equals(""))
                            //{
                            lstDataRow.Add(cellValue.ToString());
                            //}
                        }
                    }

                    string[] arrDateSplit = lstDataRow[269].Split('-');
                    string strYear = DateTime.Now.Year.ToString();
                    string strYearFront = strYear.Substring(0, 2);
                    string strDocID = strYearFront + arrDateSplit[2] + "-" + Month[arrDateSplit[1].Replace(" ", string.Empty)] + "-" + Day[arrDateSplit[0].Replace(" ", string.Empty)];

                    if (txtStatus != null && !txtStatus.Text.Equals(""))
                    {
                        txtStatus.Text += Environment.NewLine + "เลขที่เอกสาร " + lstDataRow[289] + " วันที่ " + strDocID + " ชื่อไฟล์ " + strFileNameExtension + ":";
                        strTempLogTime += Environment.NewLine + "เลขที่เอกสาร " + lstDataRow[289] + " วันที่ " + strDocID + " ชื่อไฟล์ " + strFileNameExtension + ":";
                    }
                    else
                    {
                        txtStatus.Text = "เลขที่เอกสาร " + lstDataRow[289] + " วันที่ " + strDocID + " ชื่อไฟล์ " + strFileNameExtension + " :";
                        strTempLogTime = "เลขที่เอกสาร " + lstDataRow[289] + " วันที่ " + strDocID + " ชื่อไฟล์ " + strFileNameExtension + " :";
                    }

                    txtStatus.Refresh();

                    pgbLoad.Value = 10;
                    lbPercent.Text = "Export Data: 10%";
                    lbPercent.Refresh();
                    

                    //Keep Menu
                    string[] arrCol = new string[] { "E", "F", "G", "H", "I", "J" };
                    for (int row = 20; row <= end.Row; row++)
                    { // Row by row... 
                        for (int col = 0; col < arrCol.Length; col++)
                        {
                            object cellValue = workSheet.Cells[arrCol[col] + row].Text;
                            lstDataMenu.Add(cellValue.ToString());
                        }
                    }

                    pgbLoad.Value = 20;
                    lbPercent.Text = "Export Data: 20%";
                    lbPercent.Refresh();

                    //Type C
                    grpC.Data_Type = "C";
                    grpC.Seller_Tax_ID = lstDataRow[146].Replace("Tax ID :  ", string.Empty); //เลขประจำตัวผู้เสียภาษี  
                    grpC.Seller_Tax_ID = grpC.Seller_Tax_ID.Replace(" ", string.Empty);
                    grpC.Seller_Branch_ID = lstDataRow[166].Replace("รหัสสาขา : ", string.Empty); //เลขสาขาประกอบการ
                    grpC.Seller_Branch_ID = grpC.Seller_Branch_ID.Replace(" ", string.Empty);
                    grpC.File_Name = lstDataRow[146].Replace("Tax ID :  ", string.Empty) + ".txt"; //ชื่อไฟล์     
                    grpC.File_Name = grpC.File_Name.Replace(" ", string.Empty);

                    pgbLoad.Value = 30;
                    lbPercent.Text = "Export Data: 30%";
                    lbPercent.Refresh();

                    //Type B
                    int iComSplit = lstDataRow[265].IndexOf("(");

                    if (iComSplit != -1)
                    {
                        grpB.Buyer_Name = lstDataRow[265].Substring(0, iComSplit - 1); //CompanyName
                    }
                    else
                    {
                        grpB.Buyer_Name = lstDataRow[265].Replace(" ", string.Empty); //CompanyName
                    }

                    grpB.Buyer_Phone_No = lstDataRow[345].Replace("Tel.", string.Empty); //Tel.
                    grpB.Buyer_Phone_No = grpB.Buyer_Phone_No.Replace(" ", string.Empty);
                    string strTaxID = lstDataRow[225].Replace("Tax ID :", string.Empty); //ประเภทผู้เสียภาษี
                    strTaxID = strTaxID.Replace(" ", string.Empty);
                    grpB.Buyer_Branch_ID = lstDataRow[245].Replace("รหัสสาขา : ", string.Empty); ;//เลขที่สาขา
                    grpB.Buyer_Branch_ID = grpB.Buyer_Branch_ID.Replace(" ", string.Empty);
                    if (lstDataRow[305] == "" || lstDataRow[305].Replace("รหัสไปรษณีย์ : ", string.Empty) == "")
                    {
                        grpB.Buyer_Post_Code = ("00000");
                    }
                    else
                    {
                        grpB.Buyer_Post_Code = lstDataRow[305].Replace("รหัสไปรษณีย์ : ", string.Empty);
                    }
                    string keyType = string.Empty;

                    if (strTaxID.Equals("N/A"))
                    {
                        keyType = "4";
                    }
                    else
                    {
                        int countTaxNum = strTaxID.Length;

                        if (countTaxNum == 13 && (grpB.Buyer_Branch_ID != null && !grpB.Buyer_Branch_ID.Equals("")))
                        {
                            keyType = "1";
                        }
                        else if (countTaxNum == 13 && (grpB.Buyer_Branch_ID == null || grpB.Buyer_Branch_ID.Equals("")))
                        {
                            keyType = "2";
                        }
                        //else if () //อนาคตหากมีเลขที่ PassPort
                        //{
                        //    keyType = "3";
                        //}
                    }

                    grpB.Buyer_Tax_ID_Type = BuyerTaxType[keyType];
                    grpB.Buyer_Tax_ID = strTaxID; //เลขที่ประจำตัวผู้เสียภาษี
                    grpB.Buyer_URIID = lstDataRow[346].Replace("Email : ", string.Empty);
                    grpB.Buyer_URIID = grpB.Buyer_URIID.Replace(" ", string.Empty);
                    grpB.Buyer_Add_Line1 = lstDataRow[285].Replace("ที่อยู่ : ", string.Empty);
                    grpB.Buyer_Add_Line2 = "";

                    pgbLoad.Value = 40;
                    lbPercent.Text = "Export Data: 40%";
                    lbPercent.Refresh();

                    //Type L
                    int iCountRound = 0;
                    List<LGroup> lstGrpL = new List<LGroup>();

                    for (int x = 0; x < lstDataMenu.Count; x++)
                    {
                        bool chkSting = false;
                        bool chkNum = false;
                        int value = 0;
                        string patternChkString = @"([a-zA-Zก-๙])";

                        if (!lstDataMenu[x].Equals(""))
                        {
                            chkSting = Regex.IsMatch(lstDataMenu[x + 1], patternChkString);
                            chkNum = int.TryParse(lstDataMenu[x], out value);
                        }

                        if (chkSting == true && chkNum == true)
                        {
                            if (iCountRound > 0)
                            {
                                if (grpL.Product_Desc == null || grpL.Product_Desc.Equals(""))
                                {
                                    grpL.Product_Desc = DoubleQuote("");
                                }
                                else
                                {
                                    grpL.Product_Desc = DoubleQuote(grpL.Product_Desc);
                                }

                                lstGrpL.Add(grpL);
                            }

                            grpL = new LGroup();
                            grpL.Data_Type = DoubleQuote("L"); //ประเภทรายการ
                            grpL.Line_ID = DoubleQuote(lstDataMenu[x]); //ลำดับรายการ
                            grpL.Product_ID = DoubleQuote(""); //รหัสสินค้า
                            grpL.Product_Name = lstDataMenu[x + 1].Replace("สั่งซื้อ : ", string.Empty); //ชื่อสินค้า
                            grpL.Product_Name = DoubleQuote(grpL.Product_Name.Replace(" ", string.Empty));
                            grpL.Product_Batch_ID = DoubleQuote(""); //ครั้งที่ผลิต
                            grpL.Product_Expire_Dtm = DoubleQuote(""); //วันหมดอายุ
                            grpL.Product_Class_Code = DoubleQuote(""); //รหัสหมวดหมู่สินค้า
                            grpL.Product_Class_Name = DoubleQuote(""); //ชื่อหมวดหมู่สินค้า
                            grpL.Product_OriCountry_ID = DoubleQuote(""); //รหัสประเทศกำเนิด
                            grpL.Product_Charge_Amount = DoubleQuote(RemoveComma(lstDataMenu[x + 3])); //ราคาต่อหน่วย
                            grpL.Product_Charge_Curr_Code = DoubleQuote("THB"); //รหัสสกุลเงิน (ราคาต่อหน่วย)
                            grpL.Product_Al_Charge_IND = DoubleQuote(""); //ตัวบอกส่วนลดหรือค่าธรรมเนียม
                            grpL.Product_Al_Actual_Amount = DoubleQuote(""); //มูลค่าส่วนลดหรือค่าธรรมเนียม
                            grpL.Product_Al_Actual_Curr_Code = DoubleQuote(""); //รหัสสกุลเงิน (มูลค่าส่วนลดหรือค่าธรรมเนียม)
                            grpL.Product_Al_Reason_Code = DoubleQuote(""); //รหัสเหตุผลในการคิดส่วนลดหรือค่าธรรมเนียม
                            grpL.Product_Al_Reason = DoubleQuote(""); //เหตุผลในการคิดสวนลดหรือค่าธรรมเนียม
                            grpL.Product_Quantity = DoubleQuote(lstDataMenu[x + 4]); //จำนวนสินค้า
                            grpL.Product_Unit_Code = DoubleQuote(""); //รหัสหน่วยสินค้า
                            grpL.Product_Quan_Per_Unit = DoubleQuote("1"); //ขนาดบรรจุต่อหน่วยขาย
                            grpL.Line_Tax_Type_Code = DoubleQuote("VAT"); //รหัสประเภทภาษี
                            grpL.Line_Tax_Cal_Rate = DoubleQuote("7.00"); //อัตราภาษี
                            grpL.Line_Basis_Amount = CalSumItem(lstDataMenu[x + 3], lstDataMenu[x + 4]); //มูลค่าสินค้า/บริการ (ไม่รวมภาษีมูลค่าเพิ่ม)
                            grpL.Line_Basis_Curr_Code = DoubleQuote("THB"); //รหัสสกุลเงิน (มูลค่าสินค้า/บริการ)
                            grpL.Line_Tax_Cal_Amount = CalVatItem(grpL.Line_Basis_Amount); //มูลค่าภาษีมูลค่าเพิ่ม
                            grpL.Line_Tax_Cal_Curr_Code = DoubleQuote("THB"); //รหัสสกุลเงิน (มูลค่าภาษีมูลค่าเพิ่ม)
                            grpL.Line_AL_Charge_IND = DoubleQuote(""); //ตัวบอกส่วนลดหรือค่าธรรมเนียม
                            grpL.Line_AL_Actual_Amount = DoubleQuote(""); //มูลค่าส่วนลดหรือค่าธรรมเนียม
                            grpL.Line_AL_Actual_Curr_Code = DoubleQuote(""); //มูลค่าส่วนลดหรือค่าธรรมเนียม
                            grpL.Line_AL_Reason_Code = DoubleQuote(""); //รหัสเหตุผลในการคิดส่วนลดหรือค่าธรรมเนียม
                            grpL.Line_AL_Reason = DoubleQuote(""); //เหตุผลในการคิดสวนลดหรือค่าธรรมเนียม
                            grpL.Line_Tax_Total_Amount = DoubleQuote(RemoveComma(grpL.Line_Tax_Cal_Amount)); //ภาษีมูลค่าเพิ่ม
                            grpL.Line_Tax_Total_Curr_Code = DoubleQuote("THB"); //รหัสสกุลเงิน (ภาษีมูลค่าเพิ่ม)
                            grpL.Line_Net_Total_Amount = DoubleQuote(RemoveComma(grpL.Line_Basis_Amount)); //จำนวนเงินรวมก่อนภาษี
                            grpL.Line_Net_Total_Curr_Code = DoubleQuote("THB"); //รหัสสกุลเงิน (จำนวนเงินรวมก่อนภาษี)
                            grpL.Line_Net_Include_Amount = DoubleQuote(RemoveComma(CalSumPlusVatItem(grpL.Line_Basis_Amount))); //จำนวนเงินรวม
                            strTempSumAmount = string.Empty;
                            strTempSumAmount = grpL.Line_Basis_Amount;
                            lstTempSumAmount.Add(strTempSumAmount);
                            grpL.Line_Net_Include_Curr_Code = DoubleQuote("THB"); //รหัสสกุลเงิน (จำนวนเงินรวม)
                            grpL.Line_Basis_Amount = DoubleQuote(RemoveComma(grpL.Line_Basis_Amount));
                            grpL.Line_Tax_Cal_Amount = DoubleQuote(RemoveComma(grpL.Line_Tax_Cal_Amount));
                            grpL.Product_Remark = ""; //หมายเหตุท้ายสินค้า
                            iCountRound++;
                            x += 2;
                        }
                        else
                        {
                            if (lstDataMenu[x].Contains(":") && !lstDataMenu[x].Contains("Remark : "))
                            {
                                string[] arrSplit = lstDataMenu[x].Split(':');
                                if (arrSplit[1].Equals(""))
                                {
                                    if (grpL.Product_Desc != null && !grpL.Product_Desc.Equals(""))
                                    {
                                        grpL.Product_Desc += "," + "" /*lstDataMenu[x + 1]*/; //รายละเอียดสินค้า
                                    }
                                    else
                                    {
                                        grpL.Product_Desc = ""/*lstDataMenu[x + 1]*/; //รายละเอียดสินค้า
                                    }
                                }
                            }
                        }
                    }

                    if (grpL.Product_Desc == null || grpL.Product_Desc.Equals(""))
                    {
                        grpL.Product_Desc = DoubleQuote("");
                    }
                    else
                    {
                        grpL.Product_Desc = DoubleQuote(grpL.Product_Desc);
                    }

                    lstGrpL.Add(grpL);

                    pgbLoad.Value = 50;
                    lbPercent.Text = "Export Data: 50%";
                    lbPercent.Refresh();

                    //Type F
                    double sumAmount = 0.0;

                    for (int j = 0; j < lstTempSumAmount.Count; j++)
                    {
                        sumAmount = sumAmount + double.Parse(lstTempSumAmount[j]);
                    }

                    double sumTaxAmount = sumAmount * 0.07;
                    double sumGrandTotal = sumTaxAmount + sumAmount;

                    pgbLoad.Value = 60;
                    lbPercent.Text = "Export Data: 60%";

                    //Type H
                    string[] arrKey = new string[] { "เลขที่ใบสั่งซื้อ :", "วันที่ใบสั่งซื้อ :" };
                    int[] arrIndex = new int[2];
                    int countArr = 0;
                    grpH.Doc_Type_Code = DocType[lstDataRow[248].Replace(" ", string.Empty)];
                    grpH.Doc_Name = lstDataRow[248].Replace(" ", string.Empty);
                    grpH.Doc_ID = lstDataRow[289].Replace(" ", string.Empty);
                    string[] arrDate = lstDataRow[269].Split('-');
                    string year = DateTime.Now.Year.ToString();
                    string yearFront = strYear.Substring(0, 2);
                    grpH.Doc_Issue_Dtm = strYearFront + arrDateSplit[2] + "-" + Month[arrDateSplit[1].Replace(" ", string.Empty)] + "-" + Day[arrDateSplit[0].Replace(" ", string.Empty)] + "T00:00:00";
                    grpH.Add_Ref_Assign_ID = lstDataRow[309];
                    arrDateSplit = lstDataRow[329].Split('-');
                    grpH.Add_Ref_Issue_Dtm = strYearFront + arrDateSplit[2] + "-" + Month[arrDateSplit[1].Replace(" ", string.Empty)] + "-" + Day[arrDateSplit[0].Replace(" ", string.Empty)] + "T00:00:00";

                    foreach (var key in arrKey)
                    {
                        arrIndex[countArr] = lstDataRow.IndexOf(key);
                        countArr++;
                    }

                    grpH.Buyer_Order_Assign_ID = lstDataRow[arrIndex[0] + 1].Replace(" ", string.Empty);
                    arrDateSplit = lstDataRow[arrIndex[1] + 1].Split('-');
                    grpH.Buyer_Order_Issue_Dtm = strYearFront + arrDateSplit[2] + "-" + Month[arrDateSplit[1].Replace(" ", string.Empty)] + "-" + Day[arrDateSplit[0].Replace(" ", string.Empty)] + "T00:00:00";

                    if (grpH.Buyer_Order_Assign_ID.Equals(""))
                    {
                        grpH.Buyer_Order_Ref_Type_Code = "";
                    }
                    else
                    {
                        grpH.Buyer_Order_Ref_Type_Code = "ON";
                    }

                    List<string> lstC = new List<string> { DoubleQuote("C"),
                                                DoubleQuote(grpC.Seller_Tax_ID), //เลขที่ประจำตัวผู้เสียภาษี
                                                DoubleQuote(grpC.Seller_Branch_ID), //เลขสาขาประกอบการ
                                                DoubleQuote(grpC.File_Name), //ชื่อไฟล์  
                                                };

                    List<string> lstH = new List<string> { DoubleQuote("H"),
                                                DoubleQuote(grpH.Doc_Type_Code), //ประเภทเอกสาร 
                                                DoubleQuote(grpH.Doc_Name), //ชื่อเอกสาร
                                                DoubleQuote(grpH.Doc_ID), // เลขที่เอกสาร
                                                DoubleQuote(grpH.Doc_Issue_Dtm), //วันที่
                                                DoubleQuote(""), //สาเหตุการออกเอกสาร
                                                DoubleQuote(""), //กรณีระบุสาเหตุเอกสาร
                                                DoubleQuote(grpH.Add_Ref_Assign_ID), //เลขที่เอกสารอ้างอิง
                                                DoubleQuote(grpH.Add_Ref_Issue_Dtm), //เอกสารอ้างอิงลงวันที่
                                                DoubleQuote(grpH.Doc_Type_Code), //ประเภทเอกสารอ้างอิง
                                                DoubleQuote(""), //ชื่อเอกสารอ้างอิง 
                                                DoubleQuote(""), //เงื่อนไขการส่งของ
                                                DoubleQuote(grpH.Buyer_Order_Assign_ID), //เลขที่ใบสั่งซื้อ
                                                DoubleQuote(grpH.Buyer_Order_Issue_Dtm), //วันเดือนปีที่ออกใบสั่งซื้อ
                                                DoubleQuote(grpH.Buyer_Order_Ref_Type_Code), //ประเภทเอกสารอ้างอิงการสั่งซื้อ
                                                DoubleQuote("") //หมายเหตุท้ายเอกสาร
                                                };

                    pgbLoad.Value = 70;
                    lbPercent.Text = "Export Data: 70%";
                    lbPercent.Refresh();

                    List<string> lstB = new List<string> { DoubleQuote("B"),
                                                DoubleQuote(""), //รหัสผู้ซื้อ
                                                DoubleQuote(grpB.Buyer_Name), //ชื่อผู้ซื้อ
                                                DoubleQuote(grpB.Buyer_Tax_ID_Type), //ประเภทผู้เสียภาษี
                                                DoubleQuote(grpB.Buyer_Tax_ID), //เลขประจำตัวผู้เสียภาษี
                                                DoubleQuote(grpB.Buyer_Branch_ID), //เลขที่สาขา
                                                DoubleQuote(""), //ชื่อผู้ติดต่อ
                                                DoubleQuote(""), //ชื่อแผนก
                                                DoubleQuote(grpB.Buyer_URIID), //อีเมลล์
                                                DoubleQuote(grpB.Buyer_Phone_No), //เบอร์โทรศัพท์
                                                DoubleQuote(grpB.Buyer_Post_Code), //รหัสไปรษณีย์ กรณี th
                                                DoubleQuote(""), //ชื่ออาคาร
                                                DoubleQuote(""), //บ้านเลขที่
                                                DoubleQuote(grpB.Buyer_Add_Line1), //ที่อยู่บรรทัด 1
                                                DoubleQuote(grpB.Buyer_Add_Line2), //ที่อยู่บรรทัด 2
                                                DoubleQuote(""), //ซอย
                                                DoubleQuote(""), //หมู่บ้าน
                                                DoubleQuote(""), //หมู่ที่
                                                DoubleQuote(""), //ถนน
                                                DoubleQuote(""), //รหัสตำบล
                                                DoubleQuote(""),//
                                                DoubleQuote(""), //รหัสอำเภอ
                                                DoubleQuote(""),//
                                                DoubleQuote(""), //รหัสจังหวัด
                                                DoubleQuote(""),//
                                                DoubleQuote("TH") //รหัสประเทศ
                                                };

                    List<string> lstF = new List<string> { DoubleQuote("F"),
                                                    DoubleQuote(String.Format("{0:00000}", lstGrpL.Count).ToString()), //จำนวนรายการสินค้า
                                                    DoubleQuote(""), //วันเวลานัดส่งสินค้า
                                                    DoubleQuote("THB"), //รหัสสกุลเงินตรา
                                                    DoubleQuote("VAT"), //รหัสประเภทภาษี
                                                    DoubleQuote("7.00"), //อัตราภาษี
                                                    DoubleQuote(RemoveComma(sumAmount.ToString("N2"))), //มูลค่าสินค้า(ไม่รวมภาษีมูลค่าเพิ่ม)2350
                                                    DoubleQuote("THB"),
                                                    DoubleQuote(RemoveComma(sumTaxAmount.ToString("N2"))), //มูลค่าภาษีมูลค่าเพิ่ม
                                                    DoubleQuote("THB"),
                                                    DoubleQuote(""), //ตัวบอกส่วนลดหรือค่าธรรมเนียม
                                                    DoubleQuote(""), //มูลค่าส่วนลดหรือค่าธรรมเนียม
                                                    DoubleQuote(""),
                                                    DoubleQuote(""), //รหัสเหตุผลในการคิดส่วนลดหรือค่าธรรมเนียม
                                                    DoubleQuote(""), //เหตุผลในการคิดส่วนลดหรือค่าธรรมเนียม
                                                    DoubleQuote(""), //รหัสประเภทส่วนลด     
                                                    DoubleQuote(""), //รายละเอียดเงื่อนไขการชำระเงิน
                                                    DoubleQuote(""), //วันครบกำหนดชำระเงิน
                                                    DoubleQuote(""), //รวมมูลค่าตามเอกสารเดิม
                                                    DoubleQuote(""),
                                                    DoubleQuote(RemoveComma(sumAmount.ToString("N2"))),
                                                    DoubleQuote("THB"),
                                                    DoubleQuote(""), //มูลค่าผลต่าง
                                                    DoubleQuote(""),
                                                    DoubleQuote(""), //ส่วนลดทั้งหมด
                                                    DoubleQuote(""),
                                                    DoubleQuote(""), //ค่าธรรมเนียมทั้งหมด
                                                    DoubleQuote(""),
                                                    DoubleQuote(RemoveComma(sumAmount.ToString("N2"))), //มูลค่าที่นำมาคิดภาษีมูลค่าเพิ่ม
                                                    DoubleQuote("THB"),
                                                    DoubleQuote(RemoveComma(sumTaxAmount.ToString("N2"))), //จำนวนภาษีมูลค่าเพิ่ม
                                                    DoubleQuote("THB"),
                                                    DoubleQuote(RemoveComma(sumGrandTotal.ToString("N2"))), //จำนวนเงินรวม(รวมภาษีมูลค่าเพิ่ม)
                                                    DoubleQuote("THB")
                                                    };

                    List<string> lstT = new List<string> { DoubleQuote("T"),
                                                    DoubleQuote("1") //จำนวนเอกสารทั้งหมด
                                                    };

                    pgbLoad.Value = 80;
                    lbPercent.Text = "Export Data: 80%";
                    lbPercent.Refresh();

                    string messageText = String.Join(",", lstC) + "\r"
                            + String.Join(",", lstH) + "\r"
                            + String.Join(",", lstB) + "\r";

                    for (int k = 0; k < lstGrpL.Count; k++)
                    {
                        messageText += lstGrpL[k].Data_Type + "," + lstGrpL[k].Line_ID + "," + lstGrpL[k].Product_ID + "," + lstGrpL[k].Product_Name + "," + lstGrpL[k].Product_Desc + ","
                            + lstGrpL[k].Product_Batch_ID + "," + lstGrpL[k].Product_Expire_Dtm + "," + lstGrpL[k].Product_Class_Code + "," + lstGrpL[k].Product_Class_Name + "," + lstGrpL[k].Product_OriCountry_ID + ","
                            + lstGrpL[k].Product_Charge_Amount + "," + lstGrpL[k].Product_Charge_Curr_Code + "," + lstGrpL[k].Product_Al_Charge_IND + "," + lstGrpL[k].Product_Al_Actual_Amount + "," + lstGrpL[k].Product_Al_Actual_Curr_Code + ","
                            + lstGrpL[k].Product_Al_Reason_Code + "," + lstGrpL[k].Product_Al_Reason + "," + lstGrpL[k].Product_Quantity + "," + lstGrpL[k].Product_Unit_Code + "," + lstGrpL[k].Product_Quan_Per_Unit + ","
                            + lstGrpL[k].Line_Tax_Type_Code + "," + lstGrpL[k].Line_Tax_Cal_Rate + "," + lstGrpL[k].Line_Basis_Amount + "," + lstGrpL[k].Line_Basis_Curr_Code + "," + lstGrpL[k].Line_Tax_Cal_Amount + ","
                            + lstGrpL[k].Line_Tax_Cal_Curr_Code + "," + lstGrpL[k].Line_AL_Charge_IND + "," + lstGrpL[k].Line_AL_Actual_Amount + "," + lstGrpL[k].Line_AL_Actual_Curr_Code + "," + lstGrpL[k].Line_AL_Reason_Code + ","
                            + lstGrpL[k].Line_AL_Reason + "," + lstGrpL[k].Line_Tax_Total_Amount + "," + lstGrpL[k].Line_Tax_Total_Curr_Code + "," + lstGrpL[k].Line_Net_Total_Amount + "," + lstGrpL[k].Line_Net_Total_Curr_Code + ","
                            + lstGrpL[k].Line_Net_Include_Amount + "," + lstGrpL[k].Line_Net_Include_Curr_Code + lstGrpL[k].Product_Remark + "\r";
                    }

                    messageText += String.Join(",", lstF) + "\r"
                            + String.Join(",", lstT);

                    string txtFileName = grpC.Seller_Tax_ID.Replace("Tax ID :  ", string.Empty);
                    pathText = dtParam.PathOutput + "\\" + strFileName + "_" + strDateTimeStamp + ".txt";

                    CreateTextFile(pathText, messageText);
                    txtStatus.Text += Environment.NewLine + "   -Convert Success!";
                    strTempLogTime += " Convert Success!";
                    txtStatus.Refresh();

                    System.Threading.Thread.Sleep(50);
                    strOutputFile = conAPIClass.CallAPI(dtParam, pathText, strFileNamePDF);
                }
                else
                {
                    pgbLoad.Value = 0;
                    lbPercent.Text = "Export Data: 0%";
                    lbPercent.Refresh();

                    string strText = System.IO.File.ReadAllText(dtParam.PathInput);
                    string[] arrSplit = strText.Split(',');
                    string strID = arrSplit[6].Replace("\"", string.Empty);
                    string[] arrDate = Regex.Split(arrSplit[7], "T");
                    string strDate = arrDate[0].Replace("\"", string.Empty);

                    if (txtStatus != null && !txtStatus.Text.Equals(""))
                    {
                        txtStatus.Text += Environment.NewLine + "เลขที่เอกสาร " + strID + " วันที่ " + strDate + " ชื่อไฟล์ " + strFileNameExtension + ":";
                        strTempLogTime += Environment.NewLine + "เลขที่เอกสาร " + strID + " วันที่ " + strDate + " ชื่อไฟล์ " + strFileNameExtension + ":";
                    }
                    else
                    {
                        txtStatus.Text = "เลขที่เอกสาร " + strID + " วันที่ " + strDate + " ชื่อไฟล์ " + strFileNameExtension + " :";
                        strTempLogTime = "เลขที่เอกสาร " + strID + " วันที่ " + strDate + " ชื่อไฟล์ " + strFileNameExtension + " :";
                    }

                    System.Threading.Thread.Sleep(50);
                    strOutputFile = conAPIClass.CallAPI(dtParam, dtParam.PathInput, strFileNamePDF);
                    pathText = Path.GetFileNameWithoutExtension(dtParam.PathInput) + "_" + strDateTimeStamp + ".txt";
                }

                pgbLoad.Value = 90;
                lbPercent.Text = "Export Data: 90%";
                lbPercent.Refresh();

                if (strOutputFile.MessageResultError != null && !strOutputFile.MessageResultError.Equals(""))
                {
                    if(strOutputFile.MessageResultError == "{}")
                    {
                        strOutputFile.MessageResultError = "กรุณาตรวจสอบอินเตอร์เน็ต!!";
                    }
                    txtStatus.Text += Environment.NewLine + "   -**********etax.one.th Fail!" + " (" + strOutputFile.MessageLogTime + ")" + "**********";
                    strTempLogTime += " etax.one.th Fail!" + " (" + strOutputFile.MessageLogTime + ")";
                    cntFail++;
                }
                else
                {
                    txtStatus.Text += Environment.NewLine + "   -etax.one.th Success!" + " (" + strOutputFile.MessageLogTime + ")";
                    strTempLogTime += ", etax.one.th Success!" + " (" + strOutputFile.MessageLogTime + ")";
                }

                txtStatus.Refresh();

                if (chkOption == true)
                {
                    if (strOutputFile.StatusCallAPI == false)
                    {
                        string pathErr = dtParam.PathOutput + "\\" + Path.GetFileNameWithoutExtension(pathText) + "_Error.txt";
                        CreateTextFile(pathErr, strOutputFile.MessageResultError);
                    }
                    else
                    {
                        DownloadFile(strOutputFile.MessageResultPDF, dtParam.PathOutput, Path.GetFileNameWithoutExtension(pathText) + "_PDF.pdf");
                        DownloadFile(strOutputFile.MessageResultXML, dtParam.PathOutput, Path.GetFileNameWithoutExtension(pathText) + "_XML.xml");
                    }
                }
                else if (chkOption == false)
                {
                    if (strOutputFile.StatusCallAPI == false)
                    {
                        string pathErr = pfIO.PathErr + "\\" + Path.GetFileNameWithoutExtension(pathText) + "_Error.txt";
                        CreateTextFile(pathErr, strOutputFile.MessageResultError);

                        string[] arrFiles = System.IO.Directory.GetFiles(pfIO.PathTemp, "*.txt");
                        string[] arrFilesSource = System.IO.Directory.GetFiles(pfIO.PathSource_F, "*.txt");

                        foreach (var item in arrFiles)
                        {
                            string fileName = Path.GetFileName(item);
                            string pathTxtNew = pfIO.PathSource_F + "\\" + fileName;
                            var lookupFileName = arrFilesSource.ToLookup(x => Path.GetFileName(x));

                            if (lookupFileName.Contains(fileName))
                            {
                                var resultJoin = lookupFileName[fileName];
                                foreach (var itemDelete in resultJoin)
                                {
                                    File.Exists(itemDelete);
                                    File.Delete(itemDelete);
                                }
                            }
                            File.Move(item, pathTxtNew);
                        }
                    }
                    else
                    {
                        string fileNameWithoutExtension = string.Empty;
                        string[] arrFiles = System.IO.Directory.GetFiles(pfIO.PathTemp, "*.txt");
                        //string[] arrFilesSource = System.IO.Directory.GetFiles(pfIO.PathSource_S, "*.txt");

                        //foreach (var item in arrFiles)
                        //{
                        //    string fileName = Path.GetFileName(item);
                        //    //string pathTxtNew = pfIO.PathSource_S + "\\" + fileName;
                        //    var lookupFileName = arrFilesSource.ToLookup(x => Path.GetFileName(x));

                        //    if (lookupFileName.Contains(fileName))
                        //    {
                        //        var resultJoin = lookupFileName[fileName];
                        //        foreach (var itemDelete in resultJoin)
                        //        {
                        //            File.Exists(itemDelete);
                        //            File.Delete(itemDelete);
                        //        }
                        //    }

                        //    File.Move(item, pathTxtNew);
                        //}

                        DownloadFile(strOutputFile.MessageResultPDF, pfIO.PathSuccess_O, Path.GetFileNameWithoutExtension(pathText) + "_PDF.pdf");
                        DownloadFile(strOutputFile.MessageResultXML, pfIO.PathSuccess_O, Path.GetFileNameWithoutExtension(pathText) + "_XML.xml");
                    }
                }

                pgbLoad.Value = 100;
                lbPercent.Text = "Export Data: 100%";
                lbPercent.Refresh();
            }
            catch (FileNotFoundException ex)
            {
                txtStatus.Text += Environment.NewLine + "   -**********ไม่พบ Path File ที่คุณระบุ กรุณาเลือก Path File ให้ถูกต้อง!**********";
                strTempLogTime += "ไม่พบ Path File ที่คุณระบุ กรุณาเลือก Path File ให้ถูกต้อง!";

                if (chkOption == true)
                {
                    CreateTextFile(dtParam.PathOutput + "\\" + strFileName + "_" + strDateTimeStamp + "_Error.txt", "ไม่พบ Path File ที่คุณระบุ กรุณาเลือก Path File ให้ถูกต้อง");
                }
                else
                {
                    CreateTextFile(pfIO.PathErr + "\\" + strFileName + "_" + strDateTimeStamp + "_Error.txt", "ไม่พบ Path File ที่คุณระบุ กรุณาเลือก Path File ให้ถูกต้อง");
                }

                txtStatus.Refresh();
                cntFail++;
                MessageBox.Show("ไม่พบ Path File ที่คุณระบุ กรุณาเลือก Path File ให้ถูกต้อง");
            }
            catch (System.IndexOutOfRangeException e)
            {
                txtStatus.Text += Environment.NewLine + "   -**********ไฟล์ของคุณมีข้อผิดพลาดในข้อมูลที่ใส่!**********";
                strTempLogTime += "กรุณาตรวจสอบไฟล์ของคุณ!";

                if (chkOption == true)
                {
                    CreateTextFile(dtParam.PathOutput + "\\" + strFileName + "_" + strDateTimeStamp + "_Error.txt", "ไฟล์ของคุณมีข้อผิดพลาด กรุณาตรวจสอบและใส่ข้อมูลให้ถูกต้อง");
                }
                else
                {
                    CreateTextFile(pfIO.PathErr + "\\" + strFileName + "_" + strDateTimeStamp + "_Error.txt", "ไฟล์ของคุณมีข้อผิดพลาด กรุณาตรวจสอบและใส่ข้อมูลให้ถูกต้อง");
                }

                txtStatus.Refresh();
                cntFail++;
            }
            catch (Exception ex)
            {
                txtStatus.Text += Environment.NewLine + "   -**********Convert Fail!**********";
                strTempLogTime += " Convert Fail!";

                if (chkOption == true)
                {
                    CreateTextFile(dtParam.PathOutput + "\\" + strFileName + "_" + strDateTimeStamp + "_Error.txt", "Convert Fail");
                }
                else
                {
                    CreateTextFile(pfIO.PathErr + "\\" + strFileName + "_" + strDateTimeStamp + "_Error.txt", "Convert Fail");
                }

                txtStatus.Refresh();
                cntFail++;
                //MessageBox.Show("Method RunExport Error: " + ex.ToString());
            }
            finally
            {
                lstTempSumAmount.Clear();
                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
        }

        public void DownloadFile(string strUrlFile, string pathOutputFile, string fileName)
        {
            using (var client = new WebClient())
            {
                client.DownloadFile(strUrlFile, pathOutputFile + "/" + fileName);
            }
        }

        /*""*/
        public string DoubleQuote(string str)
        {
            str = '"' + str + '"';
            return str;
        }

        //Method Delete Comma 
        public string RemoveComma(string valueNum)
        {
            valueNum = valueNum.Replace(",", string.Empty);
            return valueNum;
        }

        /*substring*/
        public string RemoveWord(string text, string word)
        {
            List<string> list = new List<string>() { word };
            foreach (string l in list)
            {
                text = text.Replace(l, "");
            }
            return text;
        }


        /*จำนวนเงินแต่ละสินค้า*/
        public string CalSumItem(string strPrice, string strQuan)
        {
            double price;
            int quantity;
            double rate;
            int.TryParse(strQuan, out quantity);
            double.TryParse(strPrice, out rate);
            price = quantity * rate;

            return price.ToString("N2");
        }

        /*จำนวนเงินแต่ละสินค้า+vat*/
        public string CalSumPlusVatItem(string strAmount)
        {
            double ItemPrice;
            double VatItem;
            double price = 0;
            double.TryParse(strAmount, out ItemPrice);
            double.TryParse(CalVatItem(strAmount), out VatItem);
            price = ItemPrice + VatItem;

            return price.ToString("N2");
        }

        /*vat ของแต่ละสินค้า*/
        public string CalVatItem(string strAmount)
        {
            double total;
            double vatItem;
            double.TryParse(strAmount, out total);
            vatItem = total * 0.07;

            return vatItem.ToString("N2");
        }

        public void RefreshForm()
        {
            etaxOnethProcess frm = new etaxOnethProcess();
            frm.Refresh();
            pbLogo.Refresh();
            lbName.Refresh();
            pbMinimize.Refresh();
            pbRestoreDown.Refresh();
            pbClose.Refresh();
            pbMaximize.Refresh();
            lbHeadName.Refresh();
            pgbLoad.Refresh();
        }

        public void CreateTextFile(string pathConcatFileName, string strData)
        {
            //Write Text
            TextWriter txtWrite = new StreamWriter(pathConcatFileName);
            txtWrite.Write(strData);
            txtWrite.Close();
        }

        private void pnHead_MouseDown(object sender, MouseEventArgs e)
        {
            mouseDown = true;
        }

        private void pnHead_MouseMove(object sender, MouseEventArgs e)
        {
            if (mouseDown)
            {
                mouseX = MousePosition.X - 200;
                mouseY = MousePosition.Y - 40;

                this.SetDesktopLocation(mouseX, mouseY);
            }
        }

        private void pnHead_MouseUp(object sender, MouseEventArgs e)
        {
            mouseDown = false;
        }
    }
}
