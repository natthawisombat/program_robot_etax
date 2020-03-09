using etaxOneth_Process.ControlAPI;
using etaxOneth_Process.DataModel;
using etaxOnethVersion2.API;
using etaxOnethVersion2.Model;
using Microsoft.Win32.SafeHandles;
using Newtonsoft.Json.Linq;
using Spire.Xls;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Net;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.RegularExpressions;
//using Spire.Pdf;
using System.Threading;
using System.Windows.Forms;
//using Spire.Pdf.Graphics.Fonts;
//using Spire.Pdf.Graphics;
using System.Xml;

namespace etaxOnethVersion2
{
    class ProcressETAX : IDisposable
    {
        int iCountRunFile;
        private bool disposed;
        private SafeRegistryHandle hExtHandle, hAppIdHandle;
        DtGetParameters dtParam = new DtGetParameters();
        ManageAPI conAPIClass = new ManageAPI();
        APIManage conAPIETAX_Viladate = new APIManage();
        DataOutput strOutputFile = new DataOutput();
        getModelOutPutViladateSign BCP_Output = new getModelOutPutViladateSign();
        API.GetSocket sock = new API.GetSocket();
        Worksheet sheet;
        API.APImail _apimail = new API.APImail();
        string strTempSumAmount = string.Empty;
        List<string> lstTempSumAmount = new List<string>();
        string strTempLogTime = string.Empty;
        int iSumFail = 0;
        int iSumSuccess = 0;
        bool chkOption = false;
        string strDateTimeStamp = string.Empty;
        public string txtstr = null;
        string RealValue, RealValue2;
        string ErrorMessage;
        string pathText, pathIn;
        string AllChar;
        string[] arrDateSplit;
        string strYear, strYearFront;
        string[] numberRange;
        string[] charRange, typeOfvalue;
        List<string> charRange_ = new List<string>();
        public string nameFilePDF { get; set; }
        int DiffOfYears, years;
        string strDocID;
        string UATorPROD;
        string itemPDF;
        DateTime dateValue;
        string pathOutput;
        string a_with_b;
        string a;
        string[] b;
        System.Reflection.Assembly assembly = System.Reflection.Assembly.GetExecutingAssembly();
        PreviewPDF s;
        string pathPreviewPDF;
        string DrResultOfPreviewPDF; // ใช้สำหรับเก็บค่าที่ปุ่ม OK หรือ Cancel ของ PreviewPDF
        public int timest_process { get; set; }
        string[] DocType =
        {
        "ใบแจ้งหนี้/ใบกำกับภาษี", "T02" ,
        "ใบส่งสินค้า/ใบแจ้งหนี้/ใบกำกับภาษี", "T02" ,
        "ใบเสร็จรับเงิน/ใบกำกับภาษี", "T03" ,
        "ใบส่งของ/ใบกำกับภาษี", "T04" ,
        "ใบรับ(ใบเสร็จรับเงิน)", "T01" ,
        "ใบเสร็จรับเงิน", "T01" ,
        "ใบกำกับภาษี","388" ,
        "ใบเพิ่มหนี้", "80" ,
        "ใบลดหนี้", "81" ,
        "ใบแจ้งหนี้", "380" ,
        "TAXINVOICE","388",
        ""
        };
        string[] DocType_ENG_AND_CODE =
        {
        "TAXINVOICE", "388" ,
        "RECEIVE", "T01" ,
        "DEBIT", "80" ,
        "CREDIT", "81" ,
        "INVOICE", "380" ,
        "T02", "380" ,
        "380", "380" ,
        "388", "388" ,
        "80", "80" ,
        "81", "81" ,
        ""
        };

        string[] TIVCPurpose =
        {
            "ชื่อผิด", "TIVC01" ,
            "ที่อยู่ผิด", "TIVC02",
            "TIVC99"
        };

        string[] DBNGPurpose =
        {
            "มีการเพิ่มราคาค่าสินค้า(สินค้าเกินกว่าจำนวนที่ตกลงกัน)", "DBNG01" ,
            "คำนวณราคาสินค้า ผิดพลาดต่ำกว่าที่เป็นจริง", "DBNG02" ,
            "การเพิ่มราคาค่าบริการ(บริการเกินกว่าข้อกำหนดที่ตกลงกัน)", "DBNS01",
            "คำนวณราคาค่าบริการ ผิดพลาดต่ำกว่าที่เป็นจริง" , "DBNS02",
            "DBNG99"
        };

        string[] CDNGPurpose =
        {
            "ลดราคาสินค้าที่ขาย(สินค้าผิดข้อกาหนดที่ตกลงกัน)", "CDNG01" ,
            "สินค้าชารุดเสียหาย", "CDNG02" ,
            "สินค้าขาดจานวนตามที่ตกลงซื้อขาย", "CDNG03",
            "คำนวณราคาสินค้าผิดพลาดสูงกว่าที่เป็นจริง" , "CDNG04",
            "รับคืนสินค้า(ไม่ตรงตามคาพรรณนา)" , "CDNG05",
            "ลดราคาค่าบริการ(บริการผิดข้อกาหนดที่ตกลงกัน)" , "CDNS01",
            "ค่าบริการขาดจานวน" , "CDNS02",
            "คำนวณราคาค่าบริการผิดพลาดสูงกว่าที่เป็นจริง", "CDNS03",
            "บอกเลิกสัญญาบริการ", "CDNS04",
            "CDNG99"
        };

        string[] RCTCPurpose =
        {
            "ชื่อผิด", "RCTC01" ,
            "ที่อยู่ผิด", "RCTC02" ,
            "รับคืนสินค้า/ยกเลิกบริการทั้งจำนวน(ระบุจำนวนเงิน)บาท", "RCTC03",
            "รับคืนสินค้า/ยกเลิกบริการ บางส่วนจำนวน(ระบุจำนวนเงิน) บาท" , "RCTC04",
            "RCTC99"
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
            { "01","01" },
            { "02","02" },
            { "03","03" },
            { "04","04" },
            { "05","05" },
            { "06","06" },
            { "07","07" },
            { "08","08" },
            { "09","09" },
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
            { "01","01" },
            { "02","02" },
            { "03","03" },
            { "04","04" },
            { "05","05" },
            { "06","06" },
            { "07","07" },
            { "08","08" },
            { "09","09" },
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


        public void StopWorking(string AmountFile, etaxOneth form)
        {

            iCountRunFile = Int32.Parse(AmountFile) + 1;

        }

        public void Dispose()
        {
            this.Dispose(true);
            GC.SuppressFinalize(this);
        }

        /// <summary>
        /// The virtual dispose method that allows
        /// classes inherithed from this one to dispose their resources.
        /// </summary>
        /// <param name="disposing"></param>
        protected virtual void Dispose(bool disposing)
        {
            if (!disposed)
            {
                if (disposing)
                {
                    // Dispose managed resources here.
                }

                // Dispose unmanaged resources here.
            }

            disposed = true;
        }
        private void PreviewpdfThread()
        {
            s = new PreviewPDF();
            s.PathPreviewPdf = this.pathPreviewPDF;
            DialogResult dr = s.ShowDialog();
            if (dr == DialogResult.Cancel)
            {
                s.Close();
                DrResultOfPreviewPDF = "Cancel";
            }
            else if (dr == DialogResult.OK)
            {
                //textBox1.Text = frm2.getText(); 
                s.Close();
                DrResultOfPreviewPDF = "OK";
            }
        }
        public ValueReturnForm RunProcess(PathFilesIO pathFileIO, string strDateTime, string strSellerTaxID, string strBranchID, string strAPIKey, string strUserCode, string strAccessKey, string strServiceCode, string strAmountFile, string strServiceURL, bool firstStatus, ValueReturnForm valueReturnChk, etaxOneth form)
        {
        Loop:
            strDateTimeStamp = strDateTime;
            chkOption = false;
            dtParam = new DtGetParameters();
            dtParam.PathInput = pathFileIO.PathInput;
            dtParam.PathConfigExcel = pathFileIO.PathConfigExcel;
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
            if (dtParam.ServiceURL == "https://uatetaxsp.one.th/etaxdocumentws/etaxsigndocument")
            {
                UATorPROD = "UAT";
            }
            else
            {
                UATorPROD = "PROD";
            }
            string fileContent;
            string[] content;
            string[] lengthOfconnectC;
            string BathText = "";
            try
            {
                RefreshForm();

                if (pathFileIO.TypeRunning.Equals("M"))
                {

                    chkOption = true;
                    string strFileName = Path.GetFileNameWithoutExtension(pathFileIO.PathInput);
                    string strFileNameEx = Path.GetFileName(pathFileIO.PathInput);
                    if (strFileNameEx.Split('.')[1] == "txt")
                    {

                        fileContent = File.ReadAllText(pathFileIO.PathInput);
                        content = fileContent.Split('\n');
                        lengthOfconnectC = content[0].Split(',');
                        if (lengthOfconnectC.Length == 5 || lengthOfconnectC.Length == 4)
                        {
                            string[] contentH = content[1].Split(',');
                            string patternChkString = @"([a-zA-Zก-๙0-9/])";
                            bool chkSting = false;
                            chkSting = Regex.IsMatch(contentH[2], patternChkString);
                            if (chkSting == false)
                            {
                                ConvertAnsiToUTF8(pathFileIO.PathInput, pathFileIO.PathInput);
                            }
                            Console.WriteLine(chkSting);
                        }
                        else
                        {
                            fileContent = File.ReadAllText(pathFileIO.PathInput);
                            content = fileContent.Split('\r');
                            lengthOfconnectC = content[0].Split(',');
                            if (lengthOfconnectC.Length == 5 || lengthOfconnectC.Length == 4)
                            {
                                string[] contentH = content[1].Split(',');
                                string patternChkString = @"([a-zA-Zก-๙0-9/])";
                                bool chkSting = false;
                                chkSting = Regex.IsMatch(contentH[2], patternChkString);
                                if (chkSting == false)
                                {
                                    ConvertAnsiToUTF8(pathFileIO.PathInput, pathFileIO.PathInput);
                                }
                                //Console.WriteLine(chkSting);
                            }
                        }
                    }
                    int iFail = 0;
                    string[] arrFilesPDF = System.IO.Directory.GetFiles(Path.GetDirectoryName(pathFileIO.PathInput), "*.pdf");
                    var regexItem = new Regex("^[a-zA-Z0-9 ]*$");
                    if (strServiceCode.Equals("S03"))
                    {

                        WorkProcess(pathFileIO, strFileName, strFileNameEx, "", out iFail, form);

                        Dispose();
                    }
                    else if (strServiceCode.Equals("S03(Excel Only)"))
                    {
                        dtParam.ServiceCode = "S03";
                        try
                        {
                            WorkProcess(pathFileIO, strFileName, strFileNameEx, "", out iFail, form);
                            Dispose();
                        }
                        catch (Exception ex)
                        {

                            MessageBox.Show(ex.Message + " 369");
                        }
                        strServiceCode = "S03(Excel Only)";
                    }
                    else if (strServiceCode.Equals("S06(Excel Only)"))
                    {
                        dtParam.ServiceCode = "S06";
                        Workbook workbook = new Workbook();
                        //string getnamefile = workbook.FileName;
                        //Console.WriteLine(getnamefile);

                        try
                        {

                            workbook.LoadFromFile(pathFileIO.PathInput);

                            using (var reader = new StreamReader(dtParam.PathConfigExcel))
                            {
                                List<string> listA = new List<string>();
                                List<string> listB = new List<string>();
                                while (!reader.EndOfStream)
                                {
                                    var line = reader.ReadLine();
                                    var values = line.Split(';');
                                    listA.Add(values[0]);
                                    //listB.Add(values[1]);
                                }
                                foreach (var items in listA)
                                {
                                    string[] arrItem = items.Split(',');
                                    //Console.WriteLine(arrItem[0]);
                                    switch (arrItem[0].ToLower().Trim(' '))
                                    {
                                        case "bathtext":
                                            Console.WriteLine("BathText : " + arrItem[2] + "," + arrItem[3]);
                                            if (!arrItem[2].Equals(""))
                                                BathText = arrItem[2] + "," + arrItem[3];
                                            else
                                                BathText = arrItem[2];
                                            break;

                                    }
                                }
                            }
                            Worksheet sheet = workbook.Worksheets[0];

                            try
                            {
                                if (!BathText.Split(',')[0].Equals(""))
                                {
                                    sheet.Range[BathText.Split(',')[0]].Text = sheet.Range[BathText.Split(',')[0]].FormulaValue.ToString();
                                }
                            }
                            catch (NullReferenceException e)
                            {

                            }
                            workbook.SaveToFile(Path.GetDirectoryName(pathFileIO.PathInput) + "\\" + Path.GetFileNameWithoutExtension(pathFileIO.PathInput) + ".pdf", Spire.Xls.FileFormat.PDF);
                            workbook.Dispose();
                        }
                        catch (IOException e)
                        {
                            MessageBox.Show("กรุณาปิดไฟล์ excel ทั้งหมด");
                            //MessageBox.Show(e.Message + " ==>c");
                            goto Loop;
                        }
                        catch (XmlException ex)
                        {
                            MessageBox.Show("ไม่สามารถทำการ Convert PDF ได้กรุณาตรวจสอบไฟล์ของคุณ");
                            strTempLogTime += " Convert Fail!" + ex.Message;
                            //MessageBox.Show(ErrorMessage);
                            if (chkOption == true)
                            {
                                CreateTextFile(dtParam.PathOutput + "\\" + strFileName + "_" + strDateTimeStamp + "_Error.txt", "ไม่สามารถทำการ Convert PDF ได้กรุณาตรวจสอบไฟล์ของคุณ");

                            }
                            else
                            {

                                CreateTextFile(dtParam.PathOutput + "\\" + strFileName + "_" + strDateTimeStamp + "_Error.txt", "ไม่สามารถทำการ Convert PDF ได้กรุณาตรวจสอบไฟล์ของคุณ");

                            }

                            form.txtStatus.Refresh();
                        }
                        if (pathFileIO.TypePrintPreview == "M")
                        {
                            try
                            {
                                var thread = new Thread(PreviewpdfThread);
                                pathPreviewPDF = Path.GetDirectoryName(pathFileIO.PathInput) + "\\" + Path.GetFileNameWithoutExtension(pathFileIO.PathInput) + ".pdf";
                                thread.SetApartmentState(ApartmentState.STA);
                                thread.Start();
                                thread.Join();
                                GC.Collect();
                                if (DrResultOfPreviewPDF == "Cancel")
                                {
                                    Dispose();
                                    Thread.MemoryBarrier();
                                    thread.Abort();
                                    if (form.txtStatus != null && !form.txtStatus.Text.Equals(""))
                                    {
                                        form.txtStatus.Text += Environment.NewLine + " ชื่อไฟล์ " + Path.GetFileName(pathFileIO.PathInput) + ":" + " ยกเลิกไฟล์ ";
                                        strTempLogTime += Environment.NewLine + " ชื่อไฟล์ " + Path.GetFileName(pathFileIO.PathInput) + ":" + " ยกเลิกไฟล์ ";
                                    }
                                    else
                                    {
                                        form.txtStatus.Text = " ชื่อไฟล์ " + Path.GetFileName(pathFileIO.PathInput) + " :" + " ยกเลิกไฟล์ ";
                                        strTempLogTime = " ชื่อไฟล์ " + Path.GetFileName(pathFileIO.PathInput) + " :" + " ยกเลิกไฟล์ ";
                                    }
                                    iSumFail++;
                                    valueReturnChk.StatusFindPDF = false;
                                    return valueReturnChk;
                                }
                                else
                                {
                                    Dispose();
                                    Thread.MemoryBarrier();
                                    thread.Abort();
                                }
                            }
                            catch (Exception e)
                            {
                                MessageBox.Show(e.Message);
                            }
                        }

                        arrFilesPDF = System.IO.Directory.GetFiles(Path.GetDirectoryName(pathFileIO.PathInput), "*.pdf");
                        var lookupName = arrFilesPDF.ToLookup(x => Path.GetFileNameWithoutExtension(x));
                        if (lookupName.Contains(strFileName))
                        {
                            var resultJoin = lookupName[strFileName];
                            foreach (var itemPDF in resultJoin)
                            {
                                try
                                {
                                    WorkProcess(pathFileIO, strFileName, strFileNameEx, itemPDF, out iFail, form);
                                    Dispose();
                                }
                                catch (Exception e)
                                {
                                    Console.WriteLine(e.Message);
                                    //MessageBox.Show("a");
                                }
                            }
                        }
                    }
                    else if (strServiceCode.Equals("S06(Excel Only & List Item)"))
                    {
                        dtParam.ServiceCode = "S06";
                        Workbook workbook = new Workbook();
                        S06ListItemModel res = new S06ListItemModel();
                        try
                        {
                            workbook.LoadFromFile(pathFileIO.PathInput);

                            using (var reader = new StreamReader(dtParam.PathConfigExcel))
                            {
                                List<string> listA = new List<string>();
                                List<string> listB = new List<string>();
                                while (!reader.EndOfStream)
                                {
                                    var line = reader.ReadLine();
                                    var values = line.Split(';');
                                    listA.Add(values[0]);
                                }
                                foreach (var items in listA)
                                {
                                    string[] arrItem = items.Split(',');

                                }
                            }

                            Worksheet sheet = workbook.Worksheets[0];

                            workbook.SaveToFile(Path.GetDirectoryName(pathFileIO.PathInput) + "\\" + Path.GetFileNameWithoutExtension(pathFileIO.PathInput) + ".pdf", Spire.Xls.FileFormat.PDF);
                            workbook.Dispose();
                        }
                        catch (IOException e)
                        {
                            MessageBox.Show("กรุณาปิดไฟล์ excel ทั้งหมด");
                            //MessageBox.Show(e.Message + " ==>c");
                            goto Loop;
                        }
                        catch (XmlException ex)
                        {
                            MessageBox.Show("ไม่สามารถทำการ Convert PDF ได้กรุณาตรวจสอบไฟล์ของคุณ");
                            strTempLogTime += " Convert Fail!" + ex.Message;
                            //MessageBox.Show(ErrorMessage);
                            if (chkOption == true)
                            {
                                CreateTextFile(dtParam.PathOutput + "\\" + strFileName + "_" + strDateTimeStamp + "_Error.txt", "ไม่สามารถทำการ Convert PDF ได้กรุณาตรวจสอบไฟล์ของคุณ");

                            }
                            else
                            {

                                CreateTextFile(dtParam.PathOutput + "\\" + strFileName + "_" + strDateTimeStamp + "_Error.txt", "ไม่สามารถทำการ Convert PDF ได้กรุณาตรวจสอบไฟล์ของคุณ");

                            }

                            form.txtStatus.Refresh();
                        }
                        if (pathFileIO.TypePrintPreview == "M")
                        {
                            try
                            {
                                var thread = new Thread(PreviewpdfThread);
                                pathPreviewPDF = Path.GetDirectoryName(pathFileIO.PathInput) + "\\" + Path.GetFileNameWithoutExtension(pathFileIO.PathInput) + ".pdf";
                                thread.SetApartmentState(ApartmentState.STA);
                                thread.Start();
                                thread.Join();
                                GC.Collect();
                                if (DrResultOfPreviewPDF == "Cancel")
                                {
                                    Dispose();
                                    Thread.MemoryBarrier();
                                    thread.Abort();
                                    if (form.txtStatus != null && !form.txtStatus.Text.Equals(""))
                                    {
                                        form.txtStatus.Text += Environment.NewLine + " ชื่อไฟล์ " + Path.GetFileName(pathFileIO.PathInput) + ":" + " ยกเลิกไฟล์ ";
                                        strTempLogTime += Environment.NewLine + " ชื่อไฟล์ " + Path.GetFileName(pathFileIO.PathInput) + ":" + " ยกเลิกไฟล์ ";
                                    }
                                    else
                                    {
                                        form.txtStatus.Text = " ชื่อไฟล์ " + Path.GetFileName(pathFileIO.PathInput) + " :" + " ยกเลิกไฟล์ ";
                                        strTempLogTime = " ชื่อไฟล์ " + Path.GetFileName(pathFileIO.PathInput) + " :" + " ยกเลิกไฟล์ ";
                                    }
                                    iSumFail++;
                                    valueReturnChk.StatusFindPDF = false;
                                    return valueReturnChk;
                                }
                                else
                                {
                                    Dispose();
                                    Thread.MemoryBarrier();
                                    thread.Abort();
                                }
                            }
                            catch (Exception e)
                            {
                                MessageBox.Show(e.Message);
                            }
                        }

                        arrFilesPDF = System.IO.Directory.GetFiles(Path.GetDirectoryName(pathFileIO.PathInput), "*.pdf");
                        var lookupName = arrFilesPDF.ToLookup(x => Path.GetFileNameWithoutExtension(x));
                        if (lookupName.Contains(strFileName))
                        {
                            var resultJoin = lookupName[strFileName];
                            foreach (var itemPDF in resultJoin)
                            {
                                try
                                {
                                    WorkProcess_forS06_ListItem(pathFileIO, strFileName, strFileNameEx, itemPDF, out iFail, form);
                                    Dispose();
                                }
                                catch (Exception e)
                                {
                                    Console.WriteLine(e.Message);
                                    //MessageBox.Show("a");
                                }
                            }
                        }
                    }
                    else if (strServiceCode.Equals("BCP Service"))
                    {
                        dtParam.ServiceCode = "S06";
                        Workbook workbook = new Workbook();
                        S06ListItemModel res = new S06ListItemModel();
                        try
                        {
                            workbook.LoadFromFile(pathFileIO.PathInput);

                            using (var reader = new StreamReader(dtParam.PathConfigExcel))
                            {
                                List<string> listA = new List<string>();
                                List<string> listB = new List<string>();
                                while (!reader.EndOfStream)
                                {
                                    var line = reader.ReadLine();
                                    var values = line.Split(';');
                                    listA.Add(values[0]);
                                }
                                foreach (var items in listA)
                                {
                                    string[] arrItem = items.Split(',');

                                }
                            }

                            Worksheet sheet = workbook.Worksheets[0];

                            workbook.SaveToFile(Path.GetDirectoryName(pathFileIO.PathInput) + "\\" + Path.GetFileNameWithoutExtension(pathFileIO.PathInput) + ".pdf", Spire.Xls.FileFormat.PDF);
                            workbook.Dispose();
                        }
                        catch (IOException e)
                        {
                            MessageBox.Show("กรุณาปิดไฟล์ excel ทั้งหมด");
                            //MessageBox.Show(e.Message + " ==>c");
                            goto Loop;
                        }
                        catch (XmlException ex)
                        {
                            MessageBox.Show("ไม่สามารถทำการ Convert PDF ได้กรุณาตรวจสอบไฟล์ของคุณ");
                            strTempLogTime += " Convert Fail!";
                            //MessageBox.Show(ErrorMessage);
                            if (chkOption == true)
                            {
                                CreateTextFile(dtParam.PathOutput + "\\" + strFileName + "_" + strDateTimeStamp + "_Error.txt", "ไม่สามารถทำการ Convert PDF ได้กรุณาตรวจสอบไฟล์ของคุณ");

                            }
                            else
                            {

                                CreateTextFile(dtParam.PathOutput + "\\" + strFileName + "_" + strDateTimeStamp + "_Error.txt", "ไม่สามารถทำการ Convert PDF ได้กรุณาตรวจสอบไฟล์ของคุณ");

                            }

                            form.txtStatus.Refresh();
                        }
                        //if (pathFileIO.TypePrintPreview == "M")
                        //{
                        //    try
                        //    {
                        //        var thread = new Thread(PreviewpdfThread);
                        //        pathPreviewPDF = Path.GetDirectoryName(pathFileIO.PathInput) + "\\" + Path.GetFileNameWithoutExtension(pathFileIO.PathInput) + ".pdf";
                        //        thread.SetApartmentState(ApartmentState.STA);
                        //        thread.Start();
                        //        thread.Join();
                        //        GC.Collect();
                        //        if (DrResultOfPreviewPDF == "Cancel")
                        //        {
                        //            Dispose();
                        //            Thread.MemoryBarrier();
                        //            thread.Abort();
                        //            if (form.txtStatus != null && !form.txtStatus.Text.Equals(""))
                        //            {
                        //                form.txtStatus.Text += Environment.NewLine + " ชื่อไฟล์ " + Path.GetFileName(pathFileIO.PathInput) + ":" + " ยกเลิกไฟล์ ";
                        //                strTempLogTime += Environment.NewLine + " ชื่อไฟล์ " + Path.GetFileName(pathFileIO.PathInput) + ":" + " ยกเลิกไฟล์ ";
                        //            }
                        //            else
                        //            {
                        //                form.txtStatus.Text = " ชื่อไฟล์ " + Path.GetFileName(pathFileIO.PathInput) + " :" + " ยกเลิกไฟล์ ";
                        //                strTempLogTime = " ชื่อไฟล์ " + Path.GetFileName(pathFileIO.PathInput) + " :" + " ยกเลิกไฟล์ ";
                        //            }
                        //            iSumFail++;
                        //            valueReturnChk.StatusFindPDF = false;
                        //            return valueReturnChk;
                        //        }
                        //        else
                        //        {
                        //            Dispose();
                        //            Thread.MemoryBarrier();
                        //            thread.Abort();
                        //        }
                        //    }
                        //    catch (Exception e)
                        //    {
                        //        MessageBox.Show(e.Message);
                        //    }
                        //}

                        //arrFilesPDF = System.IO.Directory.GetFiles(Path.GetDirectoryName(pathFileIO.PathInput), "*.pdf");
                        //var lookupName = arrFilesPDF.ToLookup(x => Path.GetFileNameWithoutExtension(x));
                        //if (lookupName.Contains(strFileName))
                        //{
                        //    var resultJoin = lookupName[strFileName];
                        //    foreach (var itemPDF in resultJoin)
                        //    {
                        //        try
                        //        {
                        //            //WorkProcess_forS06_ListItem(pathFileIO, strFileName, strFileNameEx, itemPDF, out iFail, form);
                        //            Dispose();
                        //        }
                        //        catch (Exception e)
                        //        {
                        //            Console.WriteLine(e.Message);
                        //            //MessageBox.Show("a");
                        //        }
                        //    }
                        //}
                    }
                    else
                    {
                        var lookupName = arrFilesPDF.ToLookup(x => Path.GetFileNameWithoutExtension(x));

                        if (lookupName.Contains(strFileName))
                        {
                            var resultJoin = lookupName[strFileName];
                            foreach (var itemPDF in resultJoin)
                            {
                                WorkProcess(pathFileIO, strFileName, strFileNameEx, itemPDF, out iFail, form);
                                Dispose();
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
                        string pathPrint = dtParam.PathOutput + "\\" + this.pathOutput + "_PDF.pdf";
                        valueReturnChk.pathPrint = pathPrint;
                    }
                }
                else
                //auto
                {

                    bool checkfilename;
                    //string[] arrFilespcfg = System.IO.Directory.GetFiles(pathFileIO.PathInput, "*.pcfg");
                    //string[] arrFilespcfg = System.IO.Directory.GetFiles(pathFileIO.PathInput, "*.pcfg");
                    //List<string> listpcfg = new List<string>(arrFilespcfg);
                    //if (form.check___copies.Checked == true)
                    //{
                    //    for (var i = 0; i < listpcfg.Count(); i++)
                    //    {
                    //        string strFileName = Path.GetFileNameWithoutExtension(listpcfg[i]);
                    //        var pattern = @",";
                    //        checkfilename = Regex.IsMatch(strFileName, pattern);
                    //        Console.WriteLine(checkfilename);
                    //        if (checkfilename == false)
                    //        {
                    //            Console.WriteLine(strFileName);
                    //        }
                    //        else
                    //        {
                    //            listpcfg.Remove(listpcfg[i]);
                    //        }
                    //    }
                    //}
                    //string[] arrFilespCfg = listpcfg.ToArray();
                    /*--------------------------------------------------------*/
                    //if (form.check___copies.Checked == true)
                    //{
                    //    for (var i = 0; i < arrFilespcfg.Count(); i++)
                    //    {

                    //        try
                    //        {
                    //            File.Move(arrFilespcfg[i], arrPathIO[0] + "\\" + "InputTemp" + "\\" + this.RandomNameOfFolder + "\\" + Path.GetFileName(arrFilespcfg[i]));
                    //        }
                    //        catch (Exception e)
                    //        {
                    //            Console.WriteLine(e.Message);
                    //        }
                    //    }
                    //}
                    string[] arrFilesXlsxcheck = System.IO.Directory.GetFiles(pathFileIO.PathInput, "*.xlsx");
                    List<string> listxlsx = new List<string>(arrFilesXlsxcheck);
                    for (var i = 0; i < listxlsx.Count(); i++)
                    {
                        string fileExt = System.IO.Path.GetExtension(listxlsx[i]);
                        if (fileExt == ".xlsx")
                        {
                            string strFileName = Path.GetFileNameWithoutExtension(listxlsx[i]);
                            var pattern = @",";
                            checkfilename = Regex.IsMatch(strFileName, pattern);
                            Console.WriteLine(checkfilename);
                            if (checkfilename == false)
                            {
                                Console.WriteLine(strFileName);
                            }
                            else
                            {
                                listxlsx.Remove(listxlsx[i]);
                            }
                        }

                    }
                    string[] arrFilesXlsx = listxlsx.ToArray();
                    /*-----------------------------*/

                    string[] arrFilesXlscheck = System.IO.Directory.GetFiles(pathFileIO.PathInput, "*.xls");
                    List<string> listXls = new List<string>(arrFilesXlscheck);
                    for (var i = 0; i < listXls.Count(); i++)
                    {
                        string fileExt = System.IO.Path.GetExtension(listXls[i]);
                        if (fileExt == ".xls" && listXls[i].Split('.')[listXls[i].Split('.').Length - 1] == "xls")
                        {
                            string strFileName = Path.GetFileNameWithoutExtension(listXls[i]);
                            var pattern = @",";
                            checkfilename = Regex.IsMatch(strFileName, pattern);
                            Console.WriteLine(checkfilename);
                            if (checkfilename == false)
                            {
                                Console.WriteLine(strFileName);
                            }
                            else
                            {
                                listXls.Remove(listXls[i]);
                            }
                        }
                        else
                        {
                            listXls.Remove(listXls[i]);
                        }

                    }
                    string[] arrFilesXls = listXls.ToArray();
                    /*-----------------------------*/

                    string[] arrFilesTxtcheck = System.IO.Directory.GetFiles(pathFileIO.PathInput, "*.txt");
                    List<string> listtxt = new List<string>(arrFilesTxtcheck);
                    for (var i = 0; i < listtxt.Count(); i++)
                    {
                        string strFileName = Path.GetFileNameWithoutExtension(listtxt[i]);
                        var pattern = @",";
                        checkfilename = Regex.IsMatch(strFileName, pattern);
                        Console.WriteLine(checkfilename);
                        if (checkfilename == false)
                        {
                            Console.WriteLine(strFileName);
                        }
                        else
                        {
                            listtxt.Remove(listtxt[i]);
                        }
                    }
                    string[] arrFilesTxt = listtxt.ToArray();
                    /*-----------------------------*/

                    string[] arrFilesCsvcheck = System.IO.Directory.GetFiles(pathFileIO.PathInput, "*.csv");
                    List<string> listcsv = new List<string>(arrFilesCsvcheck);
                    for (var i = 0; i < listcsv.Count(); i++)
                    {
                        string strFileName = Path.GetFileNameWithoutExtension(listcsv[i]);
                        var pattern = @",";
                        checkfilename = Regex.IsMatch(strFileName, pattern);
                        Console.WriteLine(checkfilename);
                        if (checkfilename == false)
                        {
                            Console.WriteLine(strFileName);
                        }
                        else
                        {
                            listcsv.Remove(listcsv[i]);
                        }
                    }
                    string[] arrFilesCsv = listcsv.ToArray();
                    /*-----------------------------*/

                    string[] arrFilesPDFcheck = System.IO.Directory.GetFiles(pathFileIO.PathInput, "*.pdf");
                    List<string> listpdf = new List<string>(arrFilesPDFcheck);
                    for (var i = 0; i < listpdf.Count(); i++)
                    {
                        string strFileName = Path.GetFileNameWithoutExtension(listpdf[i]);
                        var pattern = @",";
                        checkfilename = Regex.IsMatch(strFileName, pattern);
                        Console.WriteLine(checkfilename);
                        if (checkfilename == false)
                        {
                            Console.WriteLine(strFileName);
                        }
                        else
                        {
                            listpdf.Remove(listpdf[i]);
                        }
                    }
                    string[] arrFilesPDF = listpdf.ToArray();
                    /*-----------------------------*/

                    for (var i = 0; i < arrFilesXls.Length; i++)
                    {
                        if (Path.GetExtension(arrFilesXls[i]) == ".xls")
                        {
                            arrFilesXls[i] = arrFilesXls[i];
                        }
                        else
                        {
                            arrFilesXls[i] = "";
                        }
                    }
                    string strFileNameRun = string.Empty;
                    dtParam.PathOutput = string.Empty;
                    dtParam.PathOutput = pathFileIO.PathTemp;
                    iCountRunFile = 1;

                    if (arrFilesXlsx.Count() == 0 && arrFilesTxt.Count() == 0 && arrFilesXls.Count() == 0 && arrFilesCsv.Count() == 0)
                    {
                        valueReturnChk.StatusRunning = true;
                        return valueReturnChk;
                    }

                    if (firstStatus == true)
                    {
                        var Timestamp = new DateTimeOffset(DateTime.UtcNow).ToUnixTimeSeconds();
                        Console.WriteLine(Timestamp.ToString() + " Timestamp");
                        timest_process = Int32.Parse(Timestamp.ToString());
                        Console.WriteLine(timest_process + " timest_process");
                        ////CreateTextFile(pathFileIO.LogTimeProcess + "\\LogProcess_ " + strDateTimeStamp + ".txt", DateTime.Today.ToString());
                        strFileNameRun = String.Join(",", arrFilesXlsx) + "," + String.Join(",", arrFilesTxt) + "," + String.Join(",", arrFilesXls) + "," + string.Join(",", arrFilesCsv);
                        File.WriteAllText(pathFileIO.PathLogFileRun + "\\LogFileRunPerTime.txt", String.Empty);
                        CreateTextFile(pathFileIO.PathLogFileRun + "\\LogFileRunPerTime.txt", strFileNameRun);
                        //MessageBox.Show(arrFilesXlsx.Count()+"");
                        valueReturnChk.AmountAllFile = arrFilesXlsx.Count() + arrFilesTxt.Count() + arrFilesXls.Count() + arrFilesCsv.Count();
                    }
                    string strFileRunning = System.IO.File.ReadAllText(pathFileIO.PathLogFileRun + "\\LogFileRunPerTime.txt");
                    string[] arrSplitFileRunning = strFileRunning.Split(',');
                    var lookupFileLog = arrSplitFileRunning.ToLookup(x => Path.GetFileName(x));
                    var Thread_num = Environment.ProcessorCount;

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

                                try
                                {

                                    if (File.Exists(pathFileIO.PathFileRun + "\\" + Path.GetFileName(item)))
                                    {
                                        File.Delete(pathFileIO.PathFileRun + "\\" + Path.GetFileName(item));
                                        File.Move(item, pathFileIO.PathFileRun + "\\" + Path.GetFileName(item));
                                    }
                                    else
                                    {
                                        File.Move(item, pathFileIO.PathFileRun + "\\" + Path.GetFileName(item));
                                    }
                                    dtParam.PathInput = pathFileIO.PathFileRun + "\\" + Path.GetFileName(item);
                                    WorkProcess(pathFileIO, strFileName, strFileNameEx, "", out iFail, form);
                                    Dispose();
                                }
                                catch (Exception ex)
                                {
                                    pathIn = "";
                                    string[] path = pathFileIO.PathInput.Split('\\');
                                    for (int i = 0; i < path.Length; i++)
                                    {
                                        pathIn += path[i] + "/";
                                    }
                                    Console.WriteLine(pathIn);
                                    this.ErrorMessage = "ไม่สามารถย้ายไฟล์" + pathIn + Path.GetFileName(item).Replace("~$", string.Empty) + "!!!" + strDateTimeStamp;

                                    string itmsg = "ไม่สามารถย้ายไฟล์ ";
                                    _apimail.err_code = "F01";
                                    _apimail.actionmsg = pathIn + Path.GetFileName(item).Replace("~$", string.Empty);
                                    _apimail.err_msg = pathIn + Path.GetFileName(item).Replace("~$", string.Empty);
                                    _apimail.input = pathFileIO.PathInput;
                                    _apimail.path = pathFileIO.PathErr;
                                    _apimail.email = form.emailtxt.Text;
                                    _apimail.taxseller = form.txtSellerTaxID.Text;
                                    if (form.pingeng)
                                    {
                                        _apimail.send_err_service();
                                    }
                                    //sock.SendMailAlert(pathFileIO.PathInput, pathFileIO.PathErr, "F01", form.emailtxt.Text, form.txtSellerTaxID.Text, pathIn + Path.GetFileName(item).Replace("~$", string.Empty), pathIn + Path.GetFileName(item).Replace("~$", string.Empty));
                                    string pathErr = pathFileIO.PathErr + "\\" + Path.GetFileNameWithoutExtension(item) + "_" + strDateTimeStamp + "_Error.txt";
                                    CreateTextFile(pathErr, ErrorMessage);
                                    MessageBox.Show(ErrorMessage);
                                }

                            }

                            else if (strServiceCode.Equals("S03(Excel Only)"))
                            {
                                dtParam.ServiceCode = "S03";
                                try
                                {
                                    if (File.Exists(pathFileIO.PathFileRun + "\\" + Path.GetFileName(item)))
                                    {
                                        File.Delete(pathFileIO.PathFileRun + "\\" + Path.GetFileName(item));
                                        File.Move(item, pathFileIO.PathFileRun + "\\" + Path.GetFileName(item));
                                    }
                                    else
                                    {
                                        File.Move(item, pathFileIO.PathFileRun + "\\" + Path.GetFileName(item));
                                    }
                                    dtParam.PathInput = pathFileIO.PathFileRun + "\\" + Path.GetFileName(item);
                                    WorkProcess(pathFileIO, strFileName, strFileNameEx, "", out iFail, form);
                                    Dispose();
                                }
                                catch (Exception ex)
                                {
                                    pathIn = "";
                                    string[] path = pathFileIO.PathInput.Split('\\');
                                    for (int i = 0; i < path.Length; i++)
                                    {
                                        pathIn += path[i] + "/";
                                    }
                                    Console.WriteLine(pathIn);
                                    this.ErrorMessage = "ไม่สามารถย้ายไฟล์" + pathIn + Path.GetFileName(item).Replace("~$", string.Empty) + "!!!" + strDateTimeStamp;

                                    string itmsg = "ไม่สามารถย้ายไฟล์ ";
                                    _apimail.err_code = "F01";
                                    _apimail.actionmsg = pathIn + Path.GetFileName(item).Replace("~$", string.Empty);
                                    _apimail.err_msg = pathIn + Path.GetFileName(item).Replace("~$", string.Empty);
                                    _apimail.input = pathFileIO.PathInput;
                                    _apimail.path = pathFileIO.PathErr;
                                    _apimail.email = form.emailtxt.Text;
                                    _apimail.taxseller = form.txtSellerTaxID.Text;
                                    if (form.pingeng)
                                    {
                                        _apimail.send_err_service();
                                    }
                                    //sock.SendMailAlert(pathFileIO.PathInput, pathFileIO.PathErr, "F01", form.emailtxt.Text, form.txtSellerTaxID.Text, pathIn + Path.GetFileName(item).Replace("~$", string.Empty), pathIn + Path.GetFileName(item).Replace("~$", string.Empty));
                                    string pathErr = pathFileIO.PathErr + "\\" + Path.GetFileNameWithoutExtension(item) + "_" + strDateTimeStamp + "_Error.txt";
                                    CreateTextFile(pathErr, ErrorMessage);
                                }
                                strServiceCode = "S03(Excel Only)";
                            }
                            else if (strServiceCode.Equals("BCP Service"))
                            {
                                dtParam.ServiceCode = "S03";
                                Workbook workbook = new Workbook();
                                //S06ListItemModel res = new S06ListItemModel();
                                try
                                {
                                    try
                                    {
                                        workbook.LoadFromFile(item);
                                    }
                                    catch (IOException e)
                                    {
                                        pathIn = "";
                                        string[] path = pathFileIO.PathInput.Split('\\');
                                        for (int i = 0; i < path.Length; i++)
                                        {
                                            pathIn += path[i] + "/";
                                        }
                                        _apimail.err_code = "F01";
                                        _apimail.actionmsg = pathIn + Path.GetFileName(item).Replace("~$", string.Empty);
                                        _apimail.err_msg = pathIn + Path.GetFileName(item).Replace("~$", string.Empty);
                                        _apimail.input = pathFileIO.PathInput;
                                        _apimail.path = pathFileIO.PathErr;
                                        _apimail.email = form.emailtxt.Text;
                                        _apimail.taxseller = form.txtSellerTaxID.Text;
                                        if (form.pingeng)
                                        {
                                            _apimail.send_err_service();
                                        }
                                        //sock.SendMailAlert(pathFileIO.PathInput, pathFileIO.PathErr, "F01", form.emailtxt.Text, form.txtSellerTaxID.Text, pathIn + Path.GetFileName(item).Replace("~$", string.Empty), pathIn + Path.GetFileName(item).Replace("~$", string.Empty));
                                        MessageBox.Show("กรุณาปิดไฟล์ excel ทั้งหมด");
                                        //MessageBox.Show(e.Message + " ==>d");
                                        goto Loop;
                                    }
                                    try
                                    {
                                        using (var reader = new StreamReader(dtParam.PathConfigExcel))
                                        {
                                            List<string> listA = new List<string>();
                                            List<string> listB = new List<string>();
                                            while (!reader.EndOfStream)
                                            {
                                                var line = reader.ReadLine();
                                                var values = line.Split(';');
                                                listA.Add(values[0]);
                                                //listB.Add(values[1]);
                                            }
                                            foreach (var items in listA)
                                            {
                                                string[] arrItem = items.Split(',');
                                                //Console.WriteLine(arrItem[0]);
                                                switch (arrItem[0].ToLower().Trim(' '))
                                                {
                                                    case "bathtext":
                                                        Console.WriteLine("BathText : " + arrItem[2] + "," + arrItem[3]);
                                                        if (!arrItem[2].Equals(""))
                                                            BathText = arrItem[2] + "," + arrItem[3];
                                                        else
                                                            BathText = arrItem[2];
                                                        break;

                                                }
                                            }
                                        }
                                    }
                                    catch (FileNotFoundException e)
                                    {
                                        MessageBox.Show("File Not Found ConfigExcel => " + dtParam.PathConfigExcel);
                                        goto Loop;
                                    }

                                    Worksheet sheet = workbook.Worksheets[0];
                                    try
                                    {
                                        if (!BathText.Split(',')[0].Equals(""))
                                        {
                                            sheet.Range[BathText.Split(',')[0]].Text = sheet.Range[BathText.Split(',')[0]].FormulaValue.ToString();
                                            //MessageBox.Show(sheet.Range[BathText.Split(',')[0]].Value.ToString());
                                        }
                                    }
                                    catch (NullReferenceException e)
                                    {

                                    }
                                    //ทำการ convert file to pdf 
                                    workbook.SaveToFile(Path.GetDirectoryName(item) + "\\" + Path.GetFileNameWithoutExtension(item) + ".pdf", Spire.Xls.FileFormat.PDF);
                                    workbook.Dispose();
                                }
                                catch (IOException e)
                                {
                                    MessageBox.Show("กรุณาปิดไฟล์ excel ทั้งหมด");
                                    //MessageBox.Show(e.Message + " ==>c");
                                    goto Loop;
                                }
                                catch (XmlException ex)
                                {
                                    MessageBox.Show("ไม่สามารถทำการ Convert PDF ได้กรุณาตรวจสอบไฟล์ของคุณ");
                                    strTempLogTime += " Convert Fail!" + ex.Message;
                                    //MessageBox.Show(ErrorMessage);
                                    if (chkOption == true)
                                    {
                                        CreateTextFile(dtParam.PathOutput + "\\" + strFileName + "_" + strDateTimeStamp + "_Error.txt", "ไม่สามารถทำการ Convert PDF ได้กรุณาตรวจสอบไฟล์ของคุณ");
                                    }
                                    else
                                    {
                                        CreateTextFile(dtParam.PathOutput + "\\" + strFileName + "_" + strDateTimeStamp + "_Error.txt", "ไม่สามารถทำการ Convert PDF ได้กรุณาตรวจสอบไฟล์ของคุณ");
                                    }
                                    form.txtStatus.Refresh();
                                }
                                finally
                                {
                                    for (int i = 0; i <= GC.MaxGeneration; i++)
                                    {
                                        int count = GC.CollectionCount(i);
                                        GC.Collect();
                                    }
                                    GC.WaitForPendingFinalizers();
                                    GC.SuppressFinalize(this);
                                }

                                if (pathFileIO.TypePrintPreview == "M")
                                {
                                    try
                                    {
                                        var thread = new Thread(PreviewpdfThread);
                                        pathPreviewPDF = Path.GetDirectoryName(item) + "\\" + Path.GetFileNameWithoutExtension(item) + ".pdf";
                                        thread.SetApartmentState(ApartmentState.STA);
                                        thread.Start();
                                        thread.Join();
                                        if (DrResultOfPreviewPDF == "Cancel")
                                        {
                                            Dispose();
                                            thread.Abort();
                                            Thread.MemoryBarrier();
                                            try
                                            {
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
                                            catch (IOException ex)
                                            {
                                                Dispose();
                                                thread.Abort();
                                                Thread.MemoryBarrier();
                                                pathIn = "";
                                                string[] path = pathFileIO.PathInput.Split('\\');
                                                for (int i = 0; i < path.Length; i++)
                                                {
                                                    pathIn += path[i] + "/";
                                                }
                                                Console.WriteLine(pathIn);
                                                this.ErrorMessage = "ไม่สามารถย้ายไฟล์" + pathIn + Path.GetFileName(item).Replace("~$", string.Empty) + "!!!" + strDateTimeStamp;

                                                string itmsg = "ไม่สามารถย้ายไฟล์ ";
                                                _apimail.err_code = "F01";
                                                _apimail.actionmsg = pathIn + Path.GetFileName(item).Replace("~$", string.Empty);
                                                _apimail.err_msg = pathIn + Path.GetFileName(item).Replace("~$", string.Empty);
                                                _apimail.input = pathFileIO.PathInput;
                                                _apimail.path = pathFileIO.PathErr;
                                                _apimail.email = form.emailtxt.Text;
                                                _apimail.taxseller = form.txtSellerTaxID.Text;
                                                if (form.pingeng)
                                                {
                                                    _apimail.send_err_service();
                                                }
                                                //sock.SendMailAlert(pathFileIO.PathInput, pathFileIO.PathErr, "F01", form.emailtxt.Text, form.txtSellerTaxID.Text, pathIn + Path.GetFileName(item).Replace("~$", string.Empty), pathIn + Path.GetFileName(item).Replace("~$", string.Empty));
                                                string pathErr = pathFileIO.PathErr + "\\" + Path.GetFileNameWithoutExtension(item) + "_" + strDateTimeStamp + "_Error.txt";
                                                //MessageBox.Show(e.Message); 
                                                CreateTextFile(pathErr, ErrorMessage);
                                            }
                                            if (form.txtStatus != null && !form.txtStatus.Text.Equals(""))
                                            {
                                                form.txtStatus.Text += Environment.NewLine + " ชื่อไฟล์ " + Path.GetFileName(item) + ":" + " ยกเลิกไฟล์ ";
                                                strTempLogTime += Environment.NewLine + " ชื่อไฟล์ " + Path.GetFileName(item) + ":" + " ยกเลิกไฟล์ ";
                                            }
                                            else
                                            {
                                                form.txtStatus.Text = " ชื่อไฟล์ " + Path.GetFileName(item) + " :" + " ยกเลิกไฟล์ ";
                                                strTempLogTime = " ชื่อไฟล์ " + Path.GetFileName(item) + " :" + " ยกเลิกไฟล์ ";
                                            }
                                            iSumFail++;
                                            continue;
                                        }
                                        else
                                        {
                                            Dispose();
                                            thread.Abort();
                                            Thread.MemoryBarrier();
                                        }
                                    }
                                    catch (Exception e)
                                    {
                                        MessageBox.Show(e.Message);
                                    }
                                }

                                arrFilesPDF = System.IO.Directory.GetFiles(pathFileIO.PathInput, "*.pdf");
                                var lookupName = arrFilesPDF.ToLookup(x => Path.GetFileNameWithoutExtension(x));
                                if (lookupName.Contains(strFileName))
                                {
                                    var resultJoin = lookupName[strFileName];
                                    foreach (var itemPDF in resultJoin)
                                    {
                                        this.itemPDF = itemPDF;
                                        try
                                        {
                                            try
                                            {
                                                if (File.Exists(pathFileIO.PathFileRun + "\\" + Path.GetFileName(item)))
                                                {
                                                    File.Delete(pathFileIO.PathFileRun + "\\" + Path.GetFileName(item));
                                                    File.Move(item, pathFileIO.PathFileRun + "\\" + Path.GetFileName(item));
                                                }
                                                else
                                                {
                                                    File.Move(item, pathFileIO.PathFileRun + "\\" + Path.GetFileName(item));
                                                }
                                                dtParam.PathInput = pathFileIO.PathFileRun + "\\" + Path.GetFileName(item);
                                            }
                                            catch (IOException ex)
                                            {
                                                pathIn = "";
                                                string[] path = pathFileIO.PathInput.Split('\\');
                                                for (int i = 0; i < path.Length; i++)
                                                {
                                                    pathIn += path[i] + "/";
                                                }
                                                Console.WriteLine(pathIn);
                                                this.ErrorMessage = "ไม่สามารถย้ายไฟล์" + pathIn + Path.GetFileName(item).Replace("~$", string.Empty) + "!!!" + strDateTimeStamp;

                                                string itmsg = "ไม่สามารถย้ายไฟล์ ";
                                                _apimail.err_code = "F01";
                                                _apimail.actionmsg = pathIn + Path.GetFileName(item).Replace("~$", string.Empty);
                                                _apimail.err_msg = pathIn + Path.GetFileName(item).Replace("~$", string.Empty);
                                                _apimail.input = pathFileIO.PathInput;
                                                _apimail.path = pathFileIO.PathErr;
                                                _apimail.email = form.emailtxt.Text;
                                                _apimail.taxseller = form.txtSellerTaxID.Text;
                                                if (form.pingeng)
                                                {
                                                    _apimail.send_err_service();
                                                }
                                                //sock.SendMailAlert(pathFileIO.PathInput, pathFileIO.PathErr, "F01", form.emailtxt.Text, form.txtSellerTaxID.Text, pathIn + Path.GetFileName(item).Replace("~$", string.Empty), pathIn + Path.GetFileName(item).Replace("~$", string.Empty));
                                                string pathErr = pathFileIO.PathErr + "\\" + Path.GetFileNameWithoutExtension(item) + "_" + strDateTimeStamp + "_Error.txt";
                                                //MessageBox.Show(e.Message);
                                                CreateTextFile(pathErr, ErrorMessage);

                                            }
                                            //File.Move(itemPDF, pathFileIO.PathFileRun + "\\" + Path.GetFileNameWithoutExtension(itemPDF)+"_"+strDateTimeStamp+".pdf");
                                            if (File.Exists(pathFileIO.PathFileRun + "\\" + Path.GetFileName(itemPDF)))
                                            {
                                                File.Delete(pathFileIO.PathFileRun + "\\" + Path.GetFileName(itemPDF));
                                                File.Move(itemPDF, pathFileIO.PathFileRun + "\\" + Path.GetFileName(itemPDF));
                                            }
                                            else
                                            {
                                                File.Move(itemPDF, pathFileIO.PathFileRun + "\\" + Path.GetFileName(itemPDF));
                                            }
                                            this.itemPDF = pathFileIO.PathFileRun + "\\" + Path.GetFileName(itemPDF);
                                            WorkProcess_BCP(pathFileIO, strFileName, strFileNameEx, out iFail, form);
                                            //WorkProcess(pathFileIO, strFileName, strFileNameEx, this.itemPDF, out iFail, form);
                                            Dispose();
                                        }
                                        catch (IOException e)
                                        {
                                            //MessageBox.Show("a");
                                            pathIn = "";
                                            string[] path = pathFileIO.PathInput.Split('\\');
                                            for (int i = 0; i < path.Length; i++)
                                            {
                                                pathIn += path[i] + "/";
                                            }
                                            Console.WriteLine(pathIn);
                                            this.ErrorMessage = "ไม่สามารถย้ายไฟล์" + pathIn + Path.GetFileName(itemPDF).Replace("~$", string.Empty) + "!!!" + strDateTimeStamp;
                                            string itmsg = "ไม่สามารถย้ายไฟล์ ";
                                            _apimail.err_code = "F02";
                                            _apimail.actionmsg = pathIn + Path.GetFileName(itemPDF).Replace("~$", string.Empty);
                                            _apimail.err_msg = pathIn + Path.GetFileName(itemPDF).Replace("~$", string.Empty);
                                            _apimail.input = pathFileIO.PathInput;
                                            _apimail.path = pathFileIO.PathErr;
                                            _apimail.email = form.emailtxt.Text;
                                            _apimail.taxseller = form.txtSellerTaxID.Text;
                                            if (form.pingeng)
                                            {
                                                _apimail.send_err_service();
                                            }
                                            //sock.SendMailAlert(pathFileIO.PathInput, pathFileIO.PathErr, "F02", form.emailtxt.Text, form.txtSellerTaxID.Text, pathIn + Path.GetFileName(itemPDF).Replace("~$", string.Empty), pathIn + Path.GetFileName(itemPDF).Replace("~$", string.Empty));
                                            string pathErr = pathFileIO.PathErr + "\\" + Path.GetFileNameWithoutExtension(itemPDF) + "_" + strDateTimeStamp + "_Error.txt";
                                            //MessageBox.Show(e.Message);
                                            CreateTextFile(pathErr, ErrorMessage);
                                        }
                                        catch (Exception e)
                                        {
                                            MessageBox.Show("Error1" + e.Message);
                                        }
                                    }
                                }
                                else
                                {
                                    this.itemPDF = "";
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
                                    try
                                    {
                                        if (File.Exists(pathFileIO.PathFileRun + "\\" + Path.GetFileName(item)))
                                        {
                                            File.Delete(pathFileIO.PathFileRun + "\\" + Path.GetFileName(item));
                                            File.Move(item, pathFileIO.PathFileRun + "\\" + Path.GetFileName(item));
                                        }
                                        else
                                        {
                                            File.Move(item, pathFileIO.PathFileRun + "\\" + Path.GetFileName(item));
                                        }
                                        dtParam.PathInput = pathFileIO.PathFileRun + "\\" + Path.GetFileName(item);
                                    }
                                    catch (IOException ex)
                                    {
                                        pathIn = "";
                                        string[] path = pathFileIO.PathInput.Split('\\');
                                        for (int i = 0; i < path.Length; i++)
                                        {
                                            pathIn += path[i] + "/";
                                        }
                                        Console.WriteLine(pathIn);
                                        this.ErrorMessage = "ไม่สามารถย้ายไฟล์" + pathIn + Path.GetFileName(item).Replace("~$", string.Empty) + "!!!" + strDateTimeStamp;

                                        string itmsg = "ไม่สามารถย้ายไฟล์ ";
                                        _apimail.err_code = "F01";
                                        _apimail.actionmsg = pathIn + Path.GetFileName(item).Replace("~$", string.Empty);
                                        _apimail.err_msg = pathIn + Path.GetFileName(item).Replace("~$", string.Empty);
                                        _apimail.input = pathFileIO.PathInput;
                                        _apimail.path = pathFileIO.PathErr;
                                        _apimail.email = form.emailtxt.Text;
                                        _apimail.taxseller = form.txtSellerTaxID.Text;
                                        if (form.pingeng)
                                        {
                                            _apimail.send_err_service();
                                        }
                                        //sock.SendMailAlert(pathFileIO.PathInput, pathFileIO.PathErr, "F01", form.emailtxt.Text, form.txtSellerTaxID.Text, pathIn + Path.GetFileName(item).Replace("~$", string.Empty), pathIn + Path.GetFileName(item).Replace("~$", string.Empty));
                                        string pathErr = pathFileIO.PathErr + "\\" + Path.GetFileNameWithoutExtension(item) + "_" + strDateTimeStamp + "_Error.txt";
                                        //MessageBox.Show(e.Message);
                                        CreateTextFile(pathErr, ErrorMessage);
                                    }
                                    continue;
                                }
                                strServiceCode = "S06(Excel Only)";

                            }
                            else if (strServiceCode.Equals("S06(Excel Only)"))
                            {
                                dtParam.ServiceCode = "S06";
                                Workbook workbook = new Workbook();
                                try
                                {
                                    try
                                    {
                                        workbook.LoadFromFile(item);

                                    }
                                    catch (IOException e)
                                    {
                                        pathIn = "";
                                        string[] path = pathFileIO.PathInput.Split('\\');
                                        for (int i = 0; i < path.Length; i++)
                                        {
                                            pathIn += path[i] + "/";
                                        }
                                        _apimail.err_code = "F01";
                                        _apimail.actionmsg = pathIn + Path.GetFileName(item).Replace("~$", string.Empty);
                                        _apimail.err_msg = pathIn + Path.GetFileName(item).Replace("~$", string.Empty);
                                        _apimail.input = pathFileIO.PathInput;
                                        _apimail.path = pathFileIO.PathErr;
                                        _apimail.email = form.emailtxt.Text;
                                        _apimail.taxseller = form.txtSellerTaxID.Text;
                                        if (form.pingeng)
                                        {
                                            _apimail.send_err_service();
                                        }
                                        //sock.SendMailAlert(pathFileIO.PathInput, pathFileIO.PathErr, "F01", form.emailtxt.Text, form.txtSellerTaxID.Text, pathIn + Path.GetFileName(item).Replace("~$", string.Empty), pathIn + Path.GetFileName(item).Replace("~$", string.Empty));
                                        MessageBox.Show("กรุณาปิดไฟล์ excel ทั้งหมด");
                                        //MessageBox.Show(e.Message + " ==>d");
                                        goto Loop;
                                    }
                                    //catch (Exception e)
                                    //{
                                    //    MessageBox.Show("Error2" + e.Message);
                                    //}
                                    try
                                    {
                                        using (var reader = new StreamReader(dtParam.PathConfigExcel))
                                        {
                                            List<string> listA = new List<string>();
                                            List<string> listB = new List<string>();
                                            while (!reader.EndOfStream)
                                            {
                                                var line = reader.ReadLine();
                                                var values = line.Split(';');
                                                listA.Add(values[0]);
                                                //listB.Add(values[1]);
                                            }
                                            foreach (var items in listA)
                                            {
                                                string[] arrItem = items.Split(',');
                                                //Console.WriteLine(arrItem[0]);
                                                switch (arrItem[0].ToLower().Trim(' '))
                                                {
                                                    case "bathtext":
                                                        Console.WriteLine("BathText : " + arrItem[2] + "," + arrItem[3]);
                                                        if (!arrItem[2].Equals(""))
                                                            BathText = arrItem[2] + "," + arrItem[3];
                                                        else
                                                            BathText = arrItem[2];
                                                        break;

                                                }
                                            }
                                        }
                                    }
                                    catch (FileNotFoundException e)
                                    {
                                        MessageBox.Show("File Not Found ConfigExcel => " + dtParam.PathConfigExcel);
                                        goto Loop;
                                    }

                                    Worksheet sheet = workbook.Worksheets[0];
                                    try
                                    {
                                        if (!BathText.Split(',')[0].Equals(""))
                                        {
                                            sheet.Range[BathText.Split(',')[0]].Text = sheet.Range[BathText.Split(',')[0]].FormulaValue.ToString();
                                            //MessageBox.Show(sheet.Range[BathText.Split(',')[0]].Value.ToString());
                                        }
                                    }
                                    catch (NullReferenceException e)
                                    {
                                        Console.WriteLine("NullReference => " + e.Message);
                                    }
                                    catch (Exception e)
                                    {
                                        MessageBox.Show("Error3" + e.Message);
                                    }
                                    //ทำการ convert file to pdf 
                                    workbook.SaveToFile(Path.GetDirectoryName(item) + "\\" + Path.GetFileNameWithoutExtension(item) + ".pdf", Spire.Xls.FileFormat.PDF);
                                    workbook.Dispose();
                                }
                                catch (DirectoryNotFoundException ex)
                                {
                                    Console.WriteLine("DirectoryNotFound => " + ex.Message);
                                }
                                catch (IOException e)
                                {
                                    pathIn = "";
                                    string[] path = pathFileIO.PathInput.Split('\\');
                                    for (int i = 0; i < path.Length; i++)
                                    {
                                        pathIn += path[i] + "/";
                                    }
                                    _apimail.err_code = "F01";
                                    _apimail.actionmsg = pathIn + Path.GetFileName(item).Replace("~$", string.Empty);
                                    _apimail.err_msg = pathIn + Path.GetFileName(item).Replace("~$", string.Empty);
                                    _apimail.input = pathFileIO.PathInput;
                                    _apimail.path = pathFileIO.PathErr;
                                    _apimail.email = form.emailtxt.Text;
                                    _apimail.taxseller = form.txtSellerTaxID.Text;
                                    if (form.pingeng)
                                    {
                                        _apimail.send_err_service();
                                    }
                                    //sock.SendMailAlert(pathFileIO.PathInput, pathFileIO.PathErr, "F01", form.emailtxt.Text, form.txtSellerTaxID.Text, pathIn + Path.GetFileName(item).Replace("~$", string.Empty), pathIn + Path.GetFileName(item).Replace("~$", string.Empty));
                                    MessageBox.Show("กรุณาปิดไฟล์ excel ทั้งหมด");
                                    //MessageBox.Show(e.Message + " ==>d");
                                    goto Loop;
                                }
                                catch (XmlException e)
                                {
                                    pathIn = "";
                                    string[] path = pathFileIO.PathInput.Split('\\');
                                    for (int i = 0; i < path.Length; i++)
                                    {
                                        pathIn += path[i] + "/";
                                    }
                                    _apimail.err_code = "F01";
                                    _apimail.actionmsg = pathIn + Path.GetFileName(item).Replace("~$", string.Empty);
                                    _apimail.err_msg = pathIn + Path.GetFileName(item).Replace("~$", string.Empty);
                                    _apimail.input = pathFileIO.PathInput;
                                    _apimail.path = pathFileIO.PathErr;
                                    _apimail.email = form.emailtxt.Text;
                                    _apimail.taxseller = form.txtSellerTaxID.Text;
                                    if (form.pingeng)
                                    {
                                        _apimail.send_err_service();
                                    }
                                    //sock.SendMailAlert(pathFileIO.PathInput, pathFileIO.PathErr, "F01", form.emailtxt.Text, form.txtSellerTaxID.Text, pathIn + Path.GetFileName(item).Replace("~$", string.Empty), pathIn + Path.GetFileName(item).Replace("~$", string.Empty));
                                    MessageBox.Show("ไม่สามารถทำการ Convert PDF ได้กรุณาตรวจสอบไฟล์ของคุณ");
                                }
                                catch (Exception e)
                                {
                                    MessageBox.Show("Error =>" + e.Message);
                                }
                                finally
                                {
                                    for (int i = 0; i <= GC.MaxGeneration; i++)
                                    {
                                        int count = GC.CollectionCount(i);
                                        GC.Collect();
                                    }
                                    GC.WaitForPendingFinalizers();
                                    GC.SuppressFinalize(this);
                                }
                                if (pathFileIO.TypePrintPreview == "M")
                                {
                                    try
                                    {
                                        var thread = new Thread(PreviewpdfThread);
                                        pathPreviewPDF = Path.GetDirectoryName(item) + "\\" + Path.GetFileNameWithoutExtension(item) + ".pdf";
                                        thread.SetApartmentState(ApartmentState.STA);
                                        thread.Start();
                                        thread.Join();
                                        if (DrResultOfPreviewPDF == "Cancel")
                                        {
                                            Dispose();
                                            thread.Abort();
                                            Thread.MemoryBarrier();
                                            try
                                            {
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
                                            catch (IOException ex)
                                            {
                                                Dispose();
                                                thread.Abort();
                                                Thread.MemoryBarrier();
                                                pathIn = "";
                                                string[] path = pathFileIO.PathInput.Split('\\');
                                                for (int i = 0; i < path.Length; i++)
                                                {
                                                    pathIn += path[i] + "/";
                                                }
                                                Console.WriteLine(pathIn);
                                                this.ErrorMessage = "ไม่สามารถย้ายไฟล์" + pathIn + Path.GetFileName(item).Replace("~$", string.Empty) + "!!!" + strDateTimeStamp;

                                                string itmsg = "ไม่สามารถย้ายไฟล์ ";
                                                _apimail.err_code = "F01";
                                                _apimail.actionmsg = pathIn + Path.GetFileName(item).Replace("~$", string.Empty);
                                                _apimail.err_msg = pathIn + Path.GetFileName(item).Replace("~$", string.Empty);
                                                _apimail.input = pathFileIO.PathInput;
                                                _apimail.path = pathFileIO.PathErr;
                                                _apimail.email = form.emailtxt.Text;
                                                _apimail.taxseller = form.txtSellerTaxID.Text;
                                                if (form.pingeng)
                                                {
                                                    _apimail.send_err_service();
                                                }
                                                //sock.SendMailAlert(pathFileIO.PathInput, pathFileIO.PathErr, "F01", form.emailtxt.Text, form.txtSellerTaxID.Text, pathIn + Path.GetFileName(item).Replace("~$", string.Empty), pathIn + Path.GetFileName(item).Replace("~$", string.Empty));
                                                string pathErr = pathFileIO.PathErr + "\\" + Path.GetFileNameWithoutExtension(item) + "_" + strDateTimeStamp + "_Error.txt";
                                                //MessageBox.Show(e.Message); 
                                                CreateTextFile(pathErr, ErrorMessage);
                                            }
                                            if (form.txtStatus != null && !form.txtStatus.Text.Equals(""))
                                            {
                                                form.txtStatus.Text += Environment.NewLine + " ชื่อไฟล์ " + Path.GetFileName(item) + ":" + " ยกเลิกไฟล์ ";
                                                strTempLogTime += Environment.NewLine + " ชื่อไฟล์ " + Path.GetFileName(item) + ":" + " ยกเลิกไฟล์ ";
                                            }
                                            else
                                            {
                                                form.txtStatus.Text = " ชื่อไฟล์ " + Path.GetFileName(item) + " :" + " ยกเลิกไฟล์ ";
                                                strTempLogTime = " ชื่อไฟล์ " + Path.GetFileName(item) + " :" + " ยกเลิกไฟล์ ";
                                            }
                                            iSumFail++;
                                            continue;
                                        }
                                        else
                                        {
                                            Dispose();
                                            thread.Abort();
                                            Thread.MemoryBarrier();
                                        }
                                    }
                                    catch (Exception e)
                                    {
                                        MessageBox.Show(e.Message);
                                    }
                                }

                                arrFilesPDF = System.IO.Directory.GetFiles(pathFileIO.PathInput, "*.pdf");
                                var lookupName = arrFilesPDF.ToLookup(x => Path.GetFileNameWithoutExtension(x));
                                if (lookupName.Contains(strFileName))
                                {
                                    var resultJoin = lookupName[strFileName];
                                    foreach (var itemPDF in resultJoin)
                                    {
                                        this.itemPDF = itemPDF;
                                        try
                                        {
                                            try
                                            {
                                                if (File.Exists(pathFileIO.PathFileRun + "\\" + Path.GetFileName(item)))
                                                {
                                                    File.Delete(pathFileIO.PathFileRun + "\\" + Path.GetFileName(item));
                                                    File.Move(item, pathFileIO.PathFileRun + "\\" + Path.GetFileName(item));
                                                }
                                                else
                                                {
                                                    File.Move(item, pathFileIO.PathFileRun + "\\" + Path.GetFileName(item));
                                                }
                                                dtParam.PathInput = pathFileIO.PathFileRun + "\\" + Path.GetFileName(item);
                                            }
                                            catch (IOException ex)
                                            {
                                                pathIn = "";
                                                string[] path = pathFileIO.PathInput.Split('\\');
                                                for (int i = 0; i < path.Length; i++)
                                                {
                                                    pathIn += path[i] + "/";
                                                }
                                                Console.WriteLine(pathIn);
                                                this.ErrorMessage = "ไม่สามารถย้ายไฟล์" + pathIn + Path.GetFileName(item).Replace("~$", string.Empty) + "!!!" + strDateTimeStamp;

                                                string itmsg = "ไม่สามารถย้ายไฟล์ ";
                                                _apimail.err_code = "F01";
                                                _apimail.actionmsg = pathIn + Path.GetFileName(item).Replace("~$", string.Empty);
                                                _apimail.err_msg = pathIn + Path.GetFileName(item).Replace("~$", string.Empty);
                                                _apimail.input = pathFileIO.PathInput;
                                                _apimail.path = pathFileIO.PathErr;
                                                _apimail.email = form.emailtxt.Text;
                                                _apimail.taxseller = form.txtSellerTaxID.Text;
                                                if (form.pingeng)
                                                {
                                                    _apimail.send_err_service();
                                                }
                                                //sock.SendMailAlert(pathFileIO.PathInput, pathFileIO.PathErr, "F01", form.emailtxt.Text, form.txtSellerTaxID.Text, pathIn + Path.GetFileName(item).Replace("~$", string.Empty), pathIn + Path.GetFileName(item).Replace("~$", string.Empty));
                                                string pathErr = pathFileIO.PathErr + "\\" + Path.GetFileNameWithoutExtension(item) + "_" + strDateTimeStamp + "_Error.txt";
                                                //MessageBox.Show(e.Message);
                                                CreateTextFile(pathErr, ErrorMessage);

                                            }
                                            //File.Move(itemPDF, pathFileIO.PathFileRun + "\\" + Path.GetFileNameWithoutExtension(itemPDF)+"_"+strDateTimeStamp+".pdf");
                                            if (File.Exists(pathFileIO.PathFileRun + "\\" + Path.GetFileName(itemPDF)))
                                            {
                                                File.Delete(pathFileIO.PathFileRun + "\\" + Path.GetFileName(itemPDF));
                                                File.Move(itemPDF, pathFileIO.PathFileRun + "\\" + Path.GetFileName(itemPDF));
                                            }
                                            else
                                            {
                                                File.Move(itemPDF, pathFileIO.PathFileRun + "\\" + Path.GetFileName(itemPDF));
                                            }
                                            this.itemPDF = pathFileIO.PathFileRun + "\\" + Path.GetFileName(itemPDF);
                                            WorkProcess(pathFileIO, strFileName, strFileNameEx, this.itemPDF, out iFail, form);
                                            Dispose();
                                        }
                                        catch (IOException e)
                                        {
                                            //MessageBox.Show("a");
                                            pathIn = "";
                                            string[] path = pathFileIO.PathInput.Split('\\');
                                            for (int i = 0; i < path.Length; i++)
                                            {
                                                pathIn += path[i] + "/";
                                            }
                                            Console.WriteLine(pathIn);
                                            this.ErrorMessage = "ไม่สามารถย้ายไฟล์" + pathIn + Path.GetFileName(itemPDF).Replace("~$", string.Empty) + "!!!" + strDateTimeStamp;
                                            MessageBox.Show(ErrorMessage + " 1162");
                                            string itmsg = "ไม่สามารถย้ายไฟล์ ";
                                            _apimail.err_code = "F02";
                                            _apimail.actionmsg = pathIn + Path.GetFileName(itemPDF).Replace("~$", string.Empty);
                                            _apimail.err_msg = pathIn + Path.GetFileName(itemPDF).Replace("~$", string.Empty);
                                            _apimail.input = pathFileIO.PathInput;
                                            _apimail.path = pathFileIO.PathErr;
                                            _apimail.email = form.emailtxt.Text;
                                            _apimail.taxseller = form.txtSellerTaxID.Text;
                                            if (form.pingeng)
                                            {
                                                _apimail.send_err_service();
                                            }
                                            //sock.SendMailAlert(pathFileIO.PathInput, pathFileIO.PathErr, "F02", form.emailtxt.Text, form.txtSellerTaxID.Text, pathIn + Path.GetFileName(itemPDF).Replace("~$", string.Empty), pathIn + Path.GetFileName(itemPDF).Replace("~$", string.Empty));
                                            string pathErr = pathFileIO.PathErr + "\\" + Path.GetFileNameWithoutExtension(itemPDF) + "_" + strDateTimeStamp + "_Error.txt";
                                            //MessageBox.Show(e.Message);
                                            CreateTextFile(pathErr, ErrorMessage);
                                        }
                                        catch (Exception e)
                                        {

                                        }
                                    }
                                }
                                else
                                {
                                    this.itemPDF = "";
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
                                    try
                                    {
                                        if (File.Exists(pathFileIO.PathFileRun + "\\" + Path.GetFileName(item)))
                                        {
                                            File.Delete(pathFileIO.PathFileRun + "\\" + Path.GetFileName(item));
                                            File.Move(item, pathFileIO.PathFileRun + "\\" + Path.GetFileName(item));
                                        }
                                        else
                                        {
                                            File.Move(item, pathFileIO.PathFileRun + "\\" + Path.GetFileName(item));
                                        }
                                        dtParam.PathInput = pathFileIO.PathFileRun + "\\" + Path.GetFileName(item);
                                    }
                                    catch (IOException ex)
                                    {
                                        pathIn = "";
                                        string[] path = pathFileIO.PathInput.Split('\\');
                                        for (int i = 0; i < path.Length; i++)
                                        {
                                            pathIn += path[i] + "/";
                                        }
                                        Console.WriteLine(pathIn);
                                        this.ErrorMessage = "ไม่สามารถย้ายไฟล์" + pathIn + Path.GetFileName(item).Replace("~$", string.Empty) + "!!!" + strDateTimeStamp;

                                        string itmsg = "ไม่สามารถย้ายไฟล์ ";
                                        _apimail.err_code = "F01";
                                        _apimail.actionmsg = pathIn + Path.GetFileName(item).Replace("~$", string.Empty);
                                        _apimail.err_msg = pathIn + Path.GetFileName(item).Replace("~$", string.Empty);
                                        _apimail.input = pathFileIO.PathInput;
                                        _apimail.path = pathFileIO.PathErr;
                                        _apimail.email = form.emailtxt.Text;
                                        _apimail.taxseller = form.txtSellerTaxID.Text;
                                        if (form.pingeng)
                                        {
                                            _apimail.send_err_service();
                                        }
                                        //sock.SendMailAlert(pathFileIO.PathInput, pathFileIO.PathErr, "F01", form.emailtxt.Text, form.txtSellerTaxID.Text, pathIn + Path.GetFileName(item).Replace("~$", string.Empty), pathIn + Path.GetFileName(item).Replace("~$", string.Empty));
                                        string pathErr = pathFileIO.PathErr + "\\" + Path.GetFileNameWithoutExtension(item) + "_" + strDateTimeStamp + "_Error.txt";
                                        //MessageBox.Show(e.Message);
                                        CreateTextFile(pathErr, ErrorMessage);
                                        MessageBox.Show(ErrorMessage + " 1242");
                                    }
                                    continue;
                                }
                                strServiceCode = "S06(Excel Only)";
                            }
                            else if (strServiceCode.Equals("S06(Excel Only & List Item)"))
                            {
                                dtParam.ServiceCode = "S06";
                                Workbook workbook = new Workbook();
                                try
                                {
                                    try
                                    {
                                        workbook.LoadFromFile(item);
                                    }
                                    catch (IOException e)
                                    {
                                        pathIn = "";
                                        string[] path = pathFileIO.PathInput.Split('\\');
                                        for (int i = 0; i < path.Length; i++)
                                        {
                                            pathIn += path[i] + "/";
                                        }
                                        _apimail.err_code = "F01";
                                        _apimail.actionmsg = pathIn + Path.GetFileName(item).Replace("~$", string.Empty);
                                        _apimail.err_msg = pathIn + Path.GetFileName(item).Replace("~$", string.Empty);
                                        _apimail.input = pathFileIO.PathInput;
                                        _apimail.path = pathFileIO.PathErr;
                                        _apimail.email = form.emailtxt.Text;
                                        _apimail.taxseller = form.txtSellerTaxID.Text;
                                        if (form.pingeng)
                                        {
                                            _apimail.send_err_service();
                                        }
                                        //sock.SendMailAlert(pathFileIO.PathInput, pathFileIO.PathErr, "F01", form.emailtxt.Text, form.txtSellerTaxID.Text, pathIn + Path.GetFileName(item).Replace("~$", string.Empty), pathIn + Path.GetFileName(item).Replace("~$", string.Empty));
                                        MessageBox.Show("กรุณาปิดไฟล์ excel ทั้งหมด");
                                        //MessageBox.Show(e.Message + " ==>d");
                                        goto Loop;
                                    }
                                    try
                                    {
                                        using (var reader = new StreamReader(dtParam.PathConfigExcel))
                                        {
                                            List<string> listA = new List<string>();
                                            List<string> listB = new List<string>();
                                            while (!reader.EndOfStream)
                                            {
                                                var line = reader.ReadLine();
                                                var values = line.Split(';');
                                                listA.Add(values[0]);
                                                //listB.Add(values[1]);
                                            }
                                            foreach (var items in listA)
                                            {
                                                string[] arrItem = items.Split(',');
                                                //Console.WriteLine(arrItem[0]);
                                                switch (arrItem[0].ToLower().Trim(' '))
                                                {
                                                    case "bathtext":
                                                        Console.WriteLine("BathText : " + arrItem[2] + "," + arrItem[3]);
                                                        if (!arrItem[2].Equals(""))
                                                            BathText = arrItem[2] + "," + arrItem[3];
                                                        else
                                                            BathText = arrItem[2];
                                                        break;

                                                }
                                            }
                                        }
                                    }
                                    catch (FileNotFoundException e)
                                    {
                                        MessageBox.Show("File Not Found ConfigExcel => " + dtParam.PathConfigExcel);
                                        goto Loop;
                                    }

                                    Worksheet sheet = workbook.Worksheets[0];
                                    try
                                    {
                                        if (!BathText.Split(',')[0].Equals(""))
                                        {
                                            sheet.Range[BathText.Split(',')[0]].Text = sheet.Range[BathText.Split(',')[0]].FormulaValue.ToString();
                                            //MessageBox.Show(sheet.Range[BathText.Split(',')[0]].Value.ToString());
                                        }
                                    }
                                    catch (NullReferenceException e)
                                    {

                                    }
                                    //ทำการ convert file to pdf 
                                    workbook.SaveToFile(Path.GetDirectoryName(item) + "\\" + Path.GetFileNameWithoutExtension(item) + ".pdf", Spire.Xls.FileFormat.PDF);
                                    workbook.Dispose();
                                }
                                catch (DirectoryNotFoundException ex)
                                {
                                    Console.WriteLine(ex.Message);
                                }
                                catch (IOException e)
                                {
                                    pathIn = "";
                                    string[] path = pathFileIO.PathInput.Split('\\');
                                    for (int i = 0; i < path.Length; i++)
                                    {
                                        pathIn += path[i] + "/";
                                    }
                                    _apimail.err_code = "F01";
                                    _apimail.actionmsg = pathIn + Path.GetFileName(item).Replace("~$", string.Empty);
                                    _apimail.err_msg = pathIn + Path.GetFileName(item).Replace("~$", string.Empty);
                                    _apimail.input = pathFileIO.PathInput;
                                    _apimail.path = pathFileIO.PathErr;
                                    _apimail.email = form.emailtxt.Text;
                                    _apimail.taxseller = form.txtSellerTaxID.Text;
                                    if (form.pingeng)
                                    {
                                        _apimail.send_err_service();
                                    }
                                    //sock.SendMailAlert(pathFileIO.PathInput, pathFileIO.PathErr, "F01", form.emailtxt.Text, form.txtSellerTaxID.Text, pathIn + Path.GetFileName(item).Replace("~$", string.Empty), pathIn + Path.GetFileName(item).Replace("~$", string.Empty));
                                    MessageBox.Show("กรุณาปิดไฟล์ excel ทั้งหมด");
                                    //MessageBox.Show(e.Message + " ==>d");
                                    goto Loop;
                                }
                                catch (XmlException e)
                                {
                                    pathIn = "";
                                    string[] path = pathFileIO.PathInput.Split('\\');
                                    for (int i = 0; i < path.Length; i++)
                                    {
                                        pathIn += path[i] + "/";
                                    }
                                    _apimail.err_code = "F01";
                                    _apimail.actionmsg = pathIn + Path.GetFileName(item).Replace("~$", string.Empty);
                                    _apimail.err_msg = pathIn + Path.GetFileName(item).Replace("~$", string.Empty);
                                    _apimail.input = pathFileIO.PathInput;
                                    _apimail.path = pathFileIO.PathErr;
                                    _apimail.email = form.emailtxt.Text;
                                    _apimail.taxseller = form.txtSellerTaxID.Text;
                                    if (form.pingeng)
                                    {
                                        _apimail.send_err_service();
                                    }
                                    //sock.SendMailAlert(pathFileIO.PathInput, pathFileIO.PathErr, "F01", form.emailtxt.Text, form.txtSellerTaxID.Text, pathIn + Path.GetFileName(item).Replace("~$", string.Empty), pathIn + Path.GetFileName(item).Replace("~$", string.Empty));
                                    MessageBox.Show("ไม่สามารถทำการ Convert PDF ได้กรุณาตรวจสอบไฟล์ของคุณ");
                                }
                                finally
                                {
                                    for (int i = 0; i <= GC.MaxGeneration; i++)
                                    {
                                        int count = GC.CollectionCount(i);
                                        GC.Collect();
                                    }
                                    GC.WaitForPendingFinalizers();
                                    GC.SuppressFinalize(this);
                                }
                                if (pathFileIO.TypePrintPreview == "M")
                                {
                                    try
                                    {
                                        var thread = new Thread(PreviewpdfThread);
                                        pathPreviewPDF = Path.GetDirectoryName(item) + "\\" + Path.GetFileNameWithoutExtension(item) + ".pdf";
                                        thread.SetApartmentState(ApartmentState.STA);
                                        thread.Start();
                                        thread.Join();
                                        if (DrResultOfPreviewPDF == "Cancel")
                                        {
                                            Dispose();
                                            thread.Abort();
                                            Thread.MemoryBarrier();
                                            try
                                            {
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
                                            catch (IOException ex)
                                            {
                                                Dispose();
                                                thread.Abort();
                                                Thread.MemoryBarrier();
                                                pathIn = "";
                                                string[] path = pathFileIO.PathInput.Split('\\');
                                                for (int i = 0; i < path.Length; i++)
                                                {
                                                    pathIn += path[i] + "/";
                                                }
                                                Console.WriteLine(pathIn);
                                                this.ErrorMessage = "ไม่สามารถย้ายไฟล์" + pathIn + Path.GetFileName(item).Replace("~$", string.Empty) + "!!!" + strDateTimeStamp;

                                                string itmsg = "ไม่สามารถย้ายไฟล์ ";
                                                _apimail.err_code = "F01";
                                                _apimail.actionmsg = pathIn + Path.GetFileName(item).Replace("~$", string.Empty);
                                                _apimail.err_msg = pathIn + Path.GetFileName(item).Replace("~$", string.Empty);
                                                _apimail.input = pathFileIO.PathInput;
                                                _apimail.path = pathFileIO.PathErr;
                                                _apimail.email = form.emailtxt.Text;
                                                _apimail.taxseller = form.txtSellerTaxID.Text;
                                                if (form.pingeng)
                                                {
                                                    _apimail.send_err_service();
                                                }
                                                //sock.SendMailAlert(pathFileIO.PathInput, pathFileIO.PathErr, "F01", form.emailtxt.Text, form.txtSellerTaxID.Text, pathIn + Path.GetFileName(item).Replace("~$", string.Empty), pathIn + Path.GetFileName(item).Replace("~$", string.Empty));
                                                string pathErr = pathFileIO.PathErr + "\\" + Path.GetFileNameWithoutExtension(item) + "_" + strDateTimeStamp + "_Error.txt";
                                                //MessageBox.Show(e.Message); 
                                                CreateTextFile(pathErr, ErrorMessage);
                                                MessageBox.Show(ErrorMessage + " 1056");
                                            }
                                            if (form.txtStatus != null && !form.txtStatus.Text.Equals(""))
                                            {
                                                form.txtStatus.Text += Environment.NewLine + " ชื่อไฟล์ " + Path.GetFileName(item) + ":" + " ยกเลิกไฟล์ ";
                                                strTempLogTime += Environment.NewLine + " ชื่อไฟล์ " + Path.GetFileName(item) + ":" + " ยกเลิกไฟล์ ";
                                            }
                                            else
                                            {
                                                form.txtStatus.Text = " ชื่อไฟล์ " + Path.GetFileName(item) + " :" + " ยกเลิกไฟล์ ";
                                                strTempLogTime = " ชื่อไฟล์ " + Path.GetFileName(item) + " :" + " ยกเลิกไฟล์ ";
                                            }
                                            iSumFail++;
                                            continue;
                                        }
                                        else
                                        {
                                            Dispose();
                                            thread.Abort();
                                            Thread.MemoryBarrier();
                                        }
                                    }
                                    catch (Exception e)
                                    {
                                        MessageBox.Show(e.Message);
                                    }
                                }

                                arrFilesPDF = System.IO.Directory.GetFiles(pathFileIO.PathInput, "*.pdf");
                                var lookupName = arrFilesPDF.ToLookup(x => Path.GetFileNameWithoutExtension(x));
                                if (lookupName.Contains(strFileName))
                                {
                                    var resultJoin = lookupName[strFileName];
                                    foreach (var itemPDF in resultJoin)
                                    {
                                        this.itemPDF = itemPDF;
                                        try
                                        {
                                            try
                                            {
                                                if (File.Exists(pathFileIO.PathFileRun + "\\" + Path.GetFileName(item)))
                                                {
                                                    File.Delete(pathFileIO.PathFileRun + "\\" + Path.GetFileName(item));
                                                    File.Move(item, pathFileIO.PathFileRun + "\\" + Path.GetFileName(item));
                                                }
                                                else
                                                {
                                                    File.Move(item, pathFileIO.PathFileRun + "\\" + Path.GetFileName(item));
                                                }
                                                dtParam.PathInput = pathFileIO.PathFileRun + "\\" + Path.GetFileName(item);
                                            }
                                            catch (IOException ex)
                                            {
                                                pathIn = "";
                                                string[] path = pathFileIO.PathInput.Split('\\');
                                                for (int i = 0; i < path.Length; i++)
                                                {
                                                    pathIn += path[i] + "/";
                                                }
                                                Console.WriteLine(pathIn);
                                                this.ErrorMessage = "ไม่สามารถย้ายไฟล์" + pathIn + Path.GetFileName(item).Replace("~$", string.Empty) + "!!!" + strDateTimeStamp;

                                                string itmsg = "ไม่สามารถย้ายไฟล์ ";
                                                _apimail.err_code = "F01";
                                                _apimail.actionmsg = pathIn + Path.GetFileName(item).Replace("~$", string.Empty);
                                                _apimail.err_msg = pathIn + Path.GetFileName(item).Replace("~$", string.Empty);
                                                _apimail.input = pathFileIO.PathInput;
                                                _apimail.path = pathFileIO.PathErr;
                                                _apimail.email = form.emailtxt.Text;
                                                _apimail.taxseller = form.txtSellerTaxID.Text;
                                                if (form.pingeng)
                                                {
                                                    _apimail.send_err_service();
                                                }
                                                //sock.SendMailAlert(pathFileIO.PathInput, pathFileIO.PathErr, "F01", form.emailtxt.Text, form.txtSellerTaxID.Text, pathIn + Path.GetFileName(item).Replace("~$", string.Empty), pathIn + Path.GetFileName(item).Replace("~$", string.Empty));
                                                string pathErr = pathFileIO.PathErr + "\\" + Path.GetFileNameWithoutExtension(item) + "_" + strDateTimeStamp + "_Error.txt";
                                                //MessageBox.Show(e.Message);
                                                CreateTextFile(pathErr, ErrorMessage);
                                                MessageBox.Show(ErrorMessage + " 1134");

                                            }
                                            //File.Move(itemPDF, pathFileIO.PathFileRun + "\\" + Path.GetFileNameWithoutExtension(itemPDF)+"_"+strDateTimeStamp+".pdf");
                                            if (File.Exists(pathFileIO.PathFileRun + "\\" + Path.GetFileName(itemPDF)))
                                            {
                                                File.Delete(pathFileIO.PathFileRun + "\\" + Path.GetFileName(itemPDF));
                                                File.Move(itemPDF, pathFileIO.PathFileRun + "\\" + Path.GetFileName(itemPDF));
                                            }
                                            else
                                            {
                                                File.Move(itemPDF, pathFileIO.PathFileRun + "\\" + Path.GetFileName(itemPDF));
                                            }
                                            this.itemPDF = pathFileIO.PathFileRun + "\\" + Path.GetFileName(itemPDF);
                                            WorkProcess_forS06_ListItem(pathFileIO, strFileName, strFileNameEx, this.itemPDF, out iFail, form);
                                            Dispose();
                                        }
                                        catch (IOException e)
                                        {
                                            //MessageBox.Show("a");
                                            pathIn = "";
                                            string[] path = pathFileIO.PathInput.Split('\\');
                                            for (int i = 0; i < path.Length; i++)
                                            {
                                                pathIn += path[i] + "/";
                                            }
                                            Console.WriteLine(pathIn);
                                            this.ErrorMessage = "ไม่สามารถย้ายไฟล์" + pathIn + Path.GetFileName(itemPDF).Replace("~$", string.Empty) + "!!!" + strDateTimeStamp;
                                            MessageBox.Show(ErrorMessage + " 1162");
                                            string itmsg = "ไม่สามารถย้ายไฟล์ ";
                                            _apimail.err_code = "F02";
                                            _apimail.actionmsg = pathIn + Path.GetFileName(itemPDF).Replace("~$", string.Empty);
                                            _apimail.err_msg = pathIn + Path.GetFileName(itemPDF).Replace("~$", string.Empty);
                                            _apimail.input = pathFileIO.PathInput;
                                            _apimail.path = pathFileIO.PathErr;
                                            _apimail.email = form.emailtxt.Text;
                                            _apimail.taxseller = form.txtSellerTaxID.Text;
                                            if (form.pingeng)
                                            {
                                                _apimail.send_err_service();
                                            }
                                            //sock.SendMailAlert(pathFileIO.PathInput, pathFileIO.PathErr, "F02", form.emailtxt.Text, form.txtSellerTaxID.Text, pathIn + Path.GetFileName(itemPDF).Replace("~$", string.Empty), pathIn + Path.GetFileName(itemPDF).Replace("~$", string.Empty));
                                            string pathErr = pathFileIO.PathErr + "\\" + Path.GetFileNameWithoutExtension(itemPDF) + "_" + strDateTimeStamp + "_Error.txt";
                                            //MessageBox.Show(e.Message);
                                            CreateTextFile(pathErr, ErrorMessage);
                                        }
                                        catch (Exception e)
                                        {

                                        }
                                    }
                                }
                                else
                                {
                                    this.itemPDF = "";
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
                                    try
                                    {
                                        if (File.Exists(pathFileIO.PathFileRun + "\\" + Path.GetFileName(item)))
                                        {
                                            File.Delete(pathFileIO.PathFileRun + "\\" + Path.GetFileName(item));
                                            File.Move(item, pathFileIO.PathFileRun + "\\" + Path.GetFileName(item));
                                        }
                                        else
                                        {
                                            File.Move(item, pathFileIO.PathFileRun + "\\" + Path.GetFileName(item));
                                        }
                                        dtParam.PathInput = pathFileIO.PathFileRun + "\\" + Path.GetFileName(item);
                                    }
                                    catch (IOException ex)
                                    {
                                        pathIn = "";
                                        string[] path = pathFileIO.PathInput.Split('\\');
                                        for (int i = 0; i < path.Length; i++)
                                        {
                                            pathIn += path[i] + "/";
                                        }
                                        Console.WriteLine(pathIn);
                                        this.ErrorMessage = "ไม่สามารถย้ายไฟล์" + pathIn + Path.GetFileName(item).Replace("~$", string.Empty) + "!!!" + strDateTimeStamp;

                                        string itmsg = "ไม่สามารถย้ายไฟล์ ";
                                        _apimail.err_code = "F01";
                                        _apimail.actionmsg = pathIn + Path.GetFileName(item).Replace("~$", string.Empty);
                                        _apimail.err_msg = pathIn + Path.GetFileName(item).Replace("~$", string.Empty);
                                        _apimail.input = pathFileIO.PathInput;
                                        _apimail.path = pathFileIO.PathErr;
                                        _apimail.email = form.emailtxt.Text;
                                        _apimail.taxseller = form.txtSellerTaxID.Text;
                                        if (form.pingeng)
                                        {
                                            _apimail.send_err_service();
                                        }
                                        //sock.SendMailAlert(pathFileIO.PathInput, pathFileIO.PathErr, "F01", form.emailtxt.Text, form.txtSellerTaxID.Text, pathIn + Path.GetFileName(item).Replace("~$", string.Empty), pathIn + Path.GetFileName(item).Replace("~$", string.Empty));
                                        string pathErr = pathFileIO.PathErr + "\\" + Path.GetFileNameWithoutExtension(item) + "_" + strDateTimeStamp + "_Error.txt";
                                        //MessageBox.Show(e.Message);
                                        CreateTextFile(pathErr, ErrorMessage);
                                        MessageBox.Show(ErrorMessage + " 1242");
                                    }
                                    continue;
                                }
                                strServiceCode = "S06(Excel Only & List Item)";
                            }
                            else if (!strServiceCode.Equals("S03(Excel Only)"))
                            {
                                //MessageBox.Show("S06");
                                var lookupName = arrFilesPDF.ToLookup(x => Path.GetFileNameWithoutExtension(x));
                                if (lookupName.Contains(strFileName))
                                {
                                    var resultJoin = lookupName[strFileName];
                                    foreach (var itemPDF in resultJoin)
                                    {
                                        try
                                        {
                                            try
                                            {
                                                if (File.Exists(pathFileIO.PathFileRun + "\\" + Path.GetFileName(item)))
                                                {
                                                    File.Delete(pathFileIO.PathFileRun + "\\" + Path.GetFileName(item));
                                                    File.Move(item, pathFileIO.PathFileRun + "\\" + Path.GetFileName(item));
                                                }
                                                else
                                                {
                                                    File.Move(item, pathFileIO.PathFileRun + "\\" + Path.GetFileName(item));
                                                }
                                                dtParam.PathInput = pathFileIO.PathFileRun + "\\" + Path.GetFileName(item);
                                            }
                                            catch (IOException ex)
                                            {
                                                pathIn = "";
                                                string[] path = pathFileIO.PathInput.Split('\\');
                                                for (int i = 0; i < path.Length; i++)
                                                {
                                                    pathIn += path[i] + "/";
                                                }
                                                Console.WriteLine(pathIn);
                                                this.ErrorMessage = "ไม่สามารถย้ายไฟล์" + pathIn + Path.GetFileName(item).Replace("~$", string.Empty) + "!!!" + strDateTimeStamp;

                                                string itmsg = "ไม่สามารถย้ายไฟล์ ";
                                                _apimail.err_code = "F01";
                                                _apimail.actionmsg = pathIn + Path.GetFileName(item).Replace("~$", string.Empty);
                                                _apimail.err_msg = pathIn + Path.GetFileName(item).Replace("~$", string.Empty);
                                                _apimail.input = pathFileIO.PathInput;
                                                _apimail.path = pathFileIO.PathErr;
                                                _apimail.email = form.emailtxt.Text;
                                                _apimail.taxseller = form.txtSellerTaxID.Text;
                                                if (form.pingeng)
                                                {
                                                    _apimail.send_err_service();
                                                }
                                                //sock.SendMailAlert(pathFileIO.PathInput, pathFileIO.PathErr, "F01", form.emailtxt.Text, form.txtSellerTaxID.Text, pathIn + Path.GetFileName(item).Replace("~$", string.Empty), pathIn + Path.GetFileName(item).Replace("~$", string.Empty));
                                                string pathErr = pathFileIO.PathErr + "\\" + Path.GetFileNameWithoutExtension(item) + "_" + strDateTimeStamp + "_Error.txt";
                                                //MessageBox.Show(e.Message);
                                                CreateTextFile(pathErr, ErrorMessage);
                                                MessageBox.Show(ErrorMessage + " 1299");
                                            }
                                            //File.Move(itemPDF, pathFileIO.PathFileRun + "\\" + Path.GetFileNameWithoutExtension(itemPDF)+"_"+strDateTimeStamp+".pdf");
                                            if (File.Exists(pathFileIO.PathFileRun + "\\" + Path.GetFileName(itemPDF)))
                                            {
                                                File.Delete(pathFileIO.PathFileRun + "\\" + Path.GetFileName(itemPDF));
                                                File.Move(itemPDF, pathFileIO.PathFileRun + "\\" + Path.GetFileName(itemPDF));
                                            }
                                            else
                                            {
                                                File.Move(itemPDF, pathFileIO.PathFileRun + "\\" + Path.GetFileName(itemPDF));
                                            }
                                            this.itemPDF = pathFileIO.PathFileRun + "\\" + Path.GetFileName(itemPDF);
                                            WorkProcess(pathFileIO, strFileName, strFileNameEx, this.itemPDF, out iFail, form);
                                            Dispose();
                                        }
                                        catch (IOException e)
                                        {
                                            //MessageBox.Show("a");
                                            pathIn = "";
                                            string[] path = pathFileIO.PathInput.Split('\\');
                                            for (int i = 0; i < path.Length; i++)
                                            {
                                                pathIn += path[i] + "/";
                                            }
                                            Console.WriteLine(pathIn);
                                            this.ErrorMessage = "ไม่สามารถย้ายไฟล์" + pathIn + Path.GetFileName(itemPDF).Replace("~$", string.Empty) + "!!!" + strDateTimeStamp;
                                            //MessageBox.Show(ErrorMessage);
                                            string itmsg = "ไม่สามารถย้ายไฟล์ ";
                                            _apimail.err_code = "F02";
                                            _apimail.actionmsg = pathIn + Path.GetFileName(itemPDF).Replace("~$", string.Empty);
                                            _apimail.err_msg = pathIn + Path.GetFileName(itemPDF).Replace("~$", string.Empty);
                                            _apimail.input = pathFileIO.PathInput;
                                            _apimail.path = pathFileIO.PathErr;
                                            _apimail.email = form.emailtxt.Text;
                                            _apimail.taxseller = form.txtSellerTaxID.Text;
                                            if (form.pingeng)
                                            {
                                                _apimail.send_err_service();
                                            }
                                            //sock.SendMailAlert(pathFileIO.PathInput, pathFileIO.PathErr, "F02", form.emailtxt.Text, form.txtSellerTaxID.Text, pathIn + Path.GetFileName(itemPDF).Replace("~$", string.Empty), pathIn + Path.GetFileName(itemPDF).Replace("~$", string.Empty));
                                            string pathErr = pathFileIO.PathErr + "\\" + Path.GetFileNameWithoutExtension(itemPDF) + "_" + strDateTimeStamp + "_Error.txt";
                                            //MessageBox.Show(e.Message);
                                            CreateTextFile(pathErr, ErrorMessage);
                                        }
                                        catch (Exception e)
                                        {

                                        }
                                    }
                                    //MessageBox.Show(dtParam.PathInput);
                                }
                                else
                                {
                                    this.itemPDF = "";
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
                                    try
                                    {
                                        if (File.Exists(pathFileIO.PathFileRun + "\\" + Path.GetFileName(item)))
                                        {
                                            File.Delete(pathFileIO.PathFileRun + "\\" + Path.GetFileName(item));
                                            File.Move(item, pathFileIO.PathFileRun + "\\" + Path.GetFileName(item));
                                        }
                                        else
                                        {
                                            File.Move(item, pathFileIO.PathFileRun + "\\" + Path.GetFileName(item));
                                        }
                                        dtParam.PathInput = pathFileIO.PathFileRun + "\\" + Path.GetFileName(item);
                                    }
                                    catch (Exception ex)
                                    {
                                        pathIn = "";
                                        string[] path = pathFileIO.PathInput.Split('\\');
                                        for (int i = 0; i < path.Length; i++)
                                        {
                                            pathIn += path[i] + "/";
                                        }
                                        Console.WriteLine(pathIn);
                                        this.ErrorMessage = "ไม่สามารถย้ายไฟล์" + pathIn + Path.GetFileName(item).Replace("~$", string.Empty) + "!!!" + strDateTimeStamp;

                                        string itmsg = "ไม่สามารถย้ายไฟล์ ";
                                        _apimail.err_code = "F01";
                                        _apimail.actionmsg = pathIn + Path.GetFileName(item).Replace("~$", string.Empty);
                                        _apimail.err_msg = pathIn + Path.GetFileName(item).Replace("~$", string.Empty);
                                        _apimail.input = pathFileIO.PathInput;
                                        _apimail.path = pathFileIO.PathErr;
                                        _apimail.email = form.emailtxt.Text;
                                        _apimail.taxseller = form.txtSellerTaxID.Text;
                                        if (form.pingeng)
                                        {
                                            _apimail.send_err_service();
                                        }
                                        //sock.SendMailAlert(pathFileIO.PathInput, pathFileIO.PathErr, "F01", form.emailtxt.Text, form.txtSellerTaxID.Text, pathIn + Path.GetFileName(item).Replace("~$", string.Empty), pathIn + Path.GetFileName(item).Replace("~$", string.Empty));
                                        string pathErr = pathFileIO.PathErr + "\\" + Path.GetFileNameWithoutExtension(item) + "_" + strDateTimeStamp + "_Error.txt";
                                        //MessageBox.Show(e.Message);
                                        CreateTextFile(pathErr, ErrorMessage);
                                        MessageBox.Show(ErrorMessage + " 1407");
                                    }
                                    iSumFail++;
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

                    }

                    foreach (var item in arrFilesXls)
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
                                try
                                {
                                    if (File.Exists(pathFileIO.PathFileRun + "\\" + Path.GetFileName(item)))
                                    {
                                        File.Delete(pathFileIO.PathFileRun + "\\" + Path.GetFileName(item));
                                        File.Move(item, pathFileIO.PathFileRun + "\\" + Path.GetFileName(item));
                                    }
                                    else
                                    {
                                        File.Move(item, pathFileIO.PathFileRun + "\\" + Path.GetFileName(item));
                                    }
                                    dtParam.PathInput = pathFileIO.PathFileRun + "\\" + Path.GetFileName(item);
                                    WorkProcess(pathFileIO, strFileName, strFileNameEx, "", out iFail, form);
                                    Dispose();
                                }
                                catch (Exception ex)
                                {
                                    pathIn = "";
                                    string[] path = pathFileIO.PathInput.Split('\\');
                                    for (int i = 0; i < path.Length; i++)
                                    {
                                        pathIn += path[i] + "/";
                                    }
                                    Console.WriteLine(pathIn);
                                    this.ErrorMessage = "ไม่สามารถย้ายไฟล์" + pathIn + Path.GetFileName(item).Replace("~$", string.Empty) + "!!!" + strDateTimeStamp;

                                    string itmsg = "ไม่สามารถย้ายไฟล์ ";
                                    _apimail.err_code = "F01";
                                    _apimail.actionmsg = pathIn + Path.GetFileName(item).Replace("~$", string.Empty);
                                    _apimail.err_msg = pathIn + Path.GetFileName(item).Replace("~$", string.Empty);
                                    _apimail.input = pathFileIO.PathInput;
                                    _apimail.path = pathFileIO.PathErr;
                                    _apimail.email = form.emailtxt.Text;
                                    _apimail.taxseller = form.txtSellerTaxID.Text;
                                    if (form.pingeng)
                                    {
                                        _apimail.send_err_service();
                                    }
                                    //sock.SendMailAlert(pathFileIO.PathInput, pathFileIO.PathErr, "F01", form.emailtxt.Text, form.txtSellerTaxID.Text, pathIn + Path.GetFileName(item).Replace("~$", string.Empty), pathIn + Path.GetFileName(item).Replace("~$", string.Empty));
                                    string pathErr = pathFileIO.PathErr + "\\" + Path.GetFileNameWithoutExtension(item) + "_" + strDateTimeStamp + "_Error.txt";
                                    CreateTextFile(pathErr, ErrorMessage);
                                    MessageBox.Show(ErrorMessage + " 1517");
                                }

                            }
                            else if (strServiceCode.Equals("S03(Excel Only)"))
                            {
                                dtParam.ServiceCode = "S03";
                                try
                                {
                                    if (File.Exists(pathFileIO.PathFileRun + "\\" + Path.GetFileName(item)))
                                    {
                                        File.Delete(pathFileIO.PathFileRun + "\\" + Path.GetFileName(item));
                                        File.Move(item, pathFileIO.PathFileRun + "\\" + Path.GetFileName(item));
                                    }
                                    else
                                    {
                                        File.Move(item, pathFileIO.PathFileRun + "\\" + Path.GetFileName(item));
                                    }
                                    dtParam.PathInput = pathFileIO.PathFileRun + "\\" + Path.GetFileName(item);
                                    WorkProcess(pathFileIO, strFileName, strFileNameEx, "", out iFail, form);
                                    Dispose();
                                }
                                catch (Exception ex)
                                {
                                    pathIn = "";
                                    string[] path = pathFileIO.PathInput.Split('\\');
                                    for (int i = 0; i < path.Length; i++)
                                    {
                                        pathIn += path[i] + "/";
                                    }
                                    Console.WriteLine(pathIn);
                                    this.ErrorMessage = "ไม่สามารถย้ายไฟล์" + pathIn + Path.GetFileName(item).Replace("~$", string.Empty) + "!!!" + strDateTimeStamp;

                                    string itmsg = "ไม่สามารถย้ายไฟล์ ";
                                    _apimail.err_code = "F01";
                                    _apimail.actionmsg = pathIn + Path.GetFileName(item).Replace("~$", string.Empty);
                                    _apimail.err_msg = pathIn + Path.GetFileName(item).Replace("~$", string.Empty);
                                    _apimail.input = pathFileIO.PathInput;
                                    _apimail.path = pathFileIO.PathErr;
                                    _apimail.email = form.emailtxt.Text;
                                    _apimail.taxseller = form.txtSellerTaxID.Text;
                                    if (form.pingeng)
                                    {
                                        _apimail.send_err_service();
                                    }
                                    //sock.SendMailAlert(pathFileIO.PathInput, pathFileIO.PathErr, "F01", form.emailtxt.Text, form.txtSellerTaxID.Text, pathIn + Path.GetFileName(item).Replace("~$", string.Empty), pathIn + Path.GetFileName(item).Replace("~$", string.Empty));
                                    string pathErr = pathFileIO.PathErr + "\\" + Path.GetFileNameWithoutExtension(item) + "_" + strDateTimeStamp + "_Error.txt";
                                    CreateTextFile(pathErr, ErrorMessage);
                                    MessageBox.Show(ErrorMessage + " 848");
                                }
                                strServiceCode = "S03(Excel Only)";
                            }
                            else if (strServiceCode.Equals("S06(Excel Only)"))
                            {
                                dtParam.ServiceCode = "S06";
                                Workbook workbook = new Workbook();
                                try
                                {
                                    try
                                    {
                                        workbook.LoadFromFile(item);
                                    }
                                    catch (IOException e)
                                    {
                                        pathIn = "";
                                        string[] path = pathFileIO.PathInput.Split('\\');
                                        for (int i = 0; i < path.Length; i++)
                                        {
                                            pathIn += path[i] + "/";
                                        }
                                        _apimail.err_code = "F01";
                                        _apimail.actionmsg = pathIn + Path.GetFileName(item).Replace("~$", string.Empty);
                                        _apimail.err_msg = pathIn + Path.GetFileName(item).Replace("~$", string.Empty);
                                        _apimail.input = pathFileIO.PathInput;
                                        _apimail.path = pathFileIO.PathErr;
                                        _apimail.email = form.emailtxt.Text;
                                        _apimail.taxseller = form.txtSellerTaxID.Text;
                                        if (form.pingeng)
                                        {
                                            _apimail.send_err_service();
                                        }
                                        //sock.SendMailAlert(pathFileIO.PathInput, pathFileIO.PathErr, "F01", form.emailtxt.Text, form.txtSellerTaxID.Text, pathIn + Path.GetFileName(item).Replace("~$", string.Empty), pathIn + Path.GetFileName(item).Replace("~$", string.Empty));
                                        MessageBox.Show("กรุณาปิดไฟล์ excel ทั้งหมด");
                                        //MessageBox.Show(e.Message + " ==>d");
                                        goto Loop;
                                    }
                                    try
                                    {
                                        using (var reader = new StreamReader(dtParam.PathConfigExcel))
                                        {
                                            List<string> listA = new List<string>();
                                            List<string> listB = new List<string>();
                                            while (!reader.EndOfStream)
                                            {
                                                var line = reader.ReadLine();
                                                var values = line.Split(';');
                                                listA.Add(values[0]);
                                                //listB.Add(values[1]);
                                            }
                                            foreach (var items in listA)
                                            {
                                                string[] arrItem = items.Split(',');
                                                //Console.WriteLine(arrItem[0]);
                                                switch (arrItem[0].ToLower().Trim(' '))
                                                {
                                                    case "bathtext":
                                                        Console.WriteLine("BathText : " + arrItem[2] + "," + arrItem[3]);
                                                        if (!arrItem[2].Equals(""))
                                                            BathText = arrItem[2] + "," + arrItem[3];
                                                        else
                                                            BathText = arrItem[2];
                                                        break;

                                                }
                                            }
                                        }
                                    }
                                    catch (FileNotFoundException e)
                                    {
                                        MessageBox.Show("File Not Found ConfigExcel => " + dtParam.PathConfigExcel);
                                        goto Loop;
                                    }
                                    Worksheet sheet = workbook.Worksheets[0];
                                    try
                                    {
                                        if (!BathText.Split(',')[0].Equals(""))
                                        {
                                            sheet.Range[BathText.Split(',')[0]].Text = sheet.Range[BathText.Split(',')[0]].FormulaValue.ToString();
                                            //MessageBox.Show(sheet.Range[BathText.Split(',')[0]].Value.ToString());
                                        }
                                    }
                                    catch (NullReferenceException e)
                                    {

                                    }
                                    //ทำการ convert file to pdf 
                                    workbook.SaveToFile(Path.GetDirectoryName(item) + "\\" + Path.GetFileNameWithoutExtension(item) + ".pdf", Spire.Xls.FileFormat.PDF);
                                    workbook.Dispose();
                                }
                                catch (DirectoryNotFoundException ex)
                                {
                                    Console.WriteLine(ex.Message);
                                }
                                catch (IOException e)
                                {
                                    pathIn = "";
                                    string[] path = pathFileIO.PathInput.Split('\\');
                                    for (int i = 0; i < path.Length; i++)
                                    {
                                        pathIn += path[i] + "/";
                                    }
                                    _apimail.err_code = "F01";
                                    _apimail.actionmsg = pathIn + Path.GetFileName(item).Replace("~$", string.Empty);
                                    _apimail.err_msg = pathIn + Path.GetFileName(item).Replace("~$", string.Empty);
                                    _apimail.input = pathFileIO.PathInput;
                                    _apimail.path = pathFileIO.PathErr;
                                    _apimail.email = form.emailtxt.Text;
                                    _apimail.taxseller = form.txtSellerTaxID.Text;
                                    if (form.pingeng)
                                    {
                                        _apimail.send_err_service();
                                    }

                                    //sock.SendMailAlert(pathFileIO.PathInput, pathFileIO.PathErr, "F01", form.emailtxt.Text, form.txtSellerTaxID.Text, pathIn + Path.GetFileName(item).Replace("~$", string.Empty), pathIn + Path.GetFileName(item).Replace("~$", string.Empty));
                                    MessageBox.Show("กรุณาปิดไฟล์ excel ทั้งหมด");
                                    //MessageBox.Show(e.Message + " ==>d");
                                    goto Loop;
                                }
                                catch (XmlException e)
                                {
                                    pathIn = "";
                                    string[] path = pathFileIO.PathInput.Split('\\');
                                    for (int i = 0; i < path.Length; i++)
                                    {
                                        pathIn += path[i] + "/";
                                    }
                                    _apimail.err_code = "F01";
                                    _apimail.actionmsg = pathIn + Path.GetFileName(item).Replace("~$", string.Empty);
                                    _apimail.err_msg = pathIn + Path.GetFileName(item).Replace("~$", string.Empty);
                                    _apimail.input = pathFileIO.PathInput;
                                    _apimail.path = pathFileIO.PathErr;
                                    _apimail.email = form.emailtxt.Text;
                                    _apimail.taxseller = form.txtSellerTaxID.Text;
                                    if (form.pingeng)
                                    {
                                        _apimail.send_err_service();
                                    }
                                    //sock.SendMailAlert(pathFileIO.PathInput, pathFileIO.PathErr, "F01", form.emailtxt.Text, form.txtSellerTaxID.Text, pathIn + Path.GetFileName(item).Replace("~$", string.Empty), pathIn + Path.GetFileName(item).Replace("~$", string.Empty));
                                    MessageBox.Show("ไม่สามารถทำการ Convert PDF ได้กรุณาตรวจสอบไฟล์ของคุณ");
                                }
                                if (pathFileIO.TypePrintPreview == "M")
                                {
                                    try
                                    {
                                        var thread = new Thread(PreviewpdfThread);
                                        pathPreviewPDF = Path.GetDirectoryName(item) + "\\" + Path.GetFileNameWithoutExtension(item) + ".pdf";
                                        thread.SetApartmentState(ApartmentState.STA);
                                        thread.Start();
                                        thread.Join();
                                        if (DrResultOfPreviewPDF == "Cancel")
                                        {
                                            Dispose();
                                            Thread.MemoryBarrier();
                                            thread.Abort();
                                            try
                                            {
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
                                            catch (IOException ex)
                                            {
                                                Dispose();
                                                Thread.MemoryBarrier();
                                                thread.Abort();
                                                pathIn = "";
                                                string[] path = pathFileIO.PathInput.Split('\\');
                                                for (int i = 0; i < path.Length; i++)
                                                {
                                                    pathIn += path[i] + "/";
                                                }
                                                Console.WriteLine(pathIn);
                                                this.ErrorMessage = "ไม่สามารถย้ายไฟล์" + pathIn + Path.GetFileName(item).Replace("~$", string.Empty) + "!!!" + strDateTimeStamp;

                                                string itmsg = "ไม่สามารถย้ายไฟล์ ";
                                                _apimail.err_code = "F01";
                                                _apimail.actionmsg = pathIn + Path.GetFileName(item).Replace("~$", string.Empty);
                                                _apimail.err_msg = pathIn + Path.GetFileName(item).Replace("~$", string.Empty);
                                                _apimail.input = pathFileIO.PathInput;
                                                _apimail.path = pathFileIO.PathErr;
                                                _apimail.email = form.emailtxt.Text;
                                                _apimail.taxseller = form.txtSellerTaxID.Text;
                                                if (form.pingeng)
                                                {
                                                    _apimail.send_err_service();
                                                }
                                                //sock.SendMailAlert(pathFileIO.PathInput, pathFileIO.PathErr, "F01", form.emailtxt.Text, form.txtSellerTaxID.Text, pathIn + Path.GetFileName(item).Replace("~$", string.Empty), pathIn + Path.GetFileName(item).Replace("~$", string.Empty));
                                                string pathErr = pathFileIO.PathErr + "\\" + Path.GetFileNameWithoutExtension(item) + "_" + strDateTimeStamp + "_Error.txt";
                                                //MessageBox.Show(e.Message); 
                                                CreateTextFile(pathErr, ErrorMessage);
                                                MessageBox.Show(ErrorMessage + " 1715");
                                            }
                                            if (form.txtStatus != null && !form.txtStatus.Text.Equals(""))
                                            {
                                                form.txtStatus.Text += Environment.NewLine + " ชื่อไฟล์ " + Path.GetFileName(item) + ":" + " ยกเลิกไฟล์ ";
                                                strTempLogTime += Environment.NewLine + " ชื่อไฟล์ " + Path.GetFileName(item) + ":" + " ยกเลิกไฟล์ ";
                                            }
                                            else
                                            {
                                                form.txtStatus.Text = " ชื่อไฟล์ " + Path.GetFileName(item) + " :" + " ยกเลิกไฟล์ ";
                                                strTempLogTime = " ชื่อไฟล์ " + Path.GetFileName(item) + " :" + " ยกเลิกไฟล์ ";
                                            }
                                            iSumFail++;
                                            continue;
                                        }
                                        else
                                        {
                                            Dispose();
                                            Thread.MemoryBarrier();
                                            thread.Abort();
                                        }
                                    }
                                    catch (Exception e)
                                    {
                                        MessageBox.Show(e.Message);
                                    }
                                }

                                arrFilesPDF = System.IO.Directory.GetFiles(pathFileIO.PathInput, "*.pdf");
                                var lookupName = arrFilesPDF.ToLookup(x => Path.GetFileNameWithoutExtension(x));
                                if (lookupName.Contains(strFileName))
                                {
                                    var resultJoin = lookupName[strFileName];
                                    foreach (var itemPDF in resultJoin)
                                    {
                                        this.itemPDF = itemPDF;
                                        //MessageBox.Show("wait");
                                        try
                                        {
                                            try
                                            {
                                                if (File.Exists(pathFileIO.PathFileRun + "\\" + Path.GetFileName(item)))
                                                {
                                                    File.Delete(pathFileIO.PathFileRun + "\\" + Path.GetFileName(item));
                                                    File.Move(item, pathFileIO.PathFileRun + "\\" + Path.GetFileName(item));
                                                }
                                                else
                                                {
                                                    File.Move(item, pathFileIO.PathFileRun + "\\" + Path.GetFileName(item));
                                                }
                                                dtParam.PathInput = pathFileIO.PathFileRun + "\\" + Path.GetFileName(item);
                                            }
                                            catch (IOException ex)
                                            {
                                                pathIn = "";
                                                string[] path = pathFileIO.PathInput.Split('\\');
                                                for (int i = 0; i < path.Length; i++)
                                                {
                                                    pathIn += path[i] + "/";
                                                }
                                                Console.WriteLine(pathIn);
                                                this.ErrorMessage = "ไม่สามารถย้ายไฟล์" + pathIn + Path.GetFileName(item).Replace("~$", string.Empty) + "!!!" + strDateTimeStamp;

                                                string itmsg = "ไม่สามารถย้ายไฟล์ ";
                                                _apimail.err_code = "F01";
                                                _apimail.actionmsg = pathIn + Path.GetFileName(item).Replace("~$", string.Empty);
                                                _apimail.err_msg = pathIn + Path.GetFileName(item).Replace("~$", string.Empty);
                                                _apimail.input = pathFileIO.PathInput;
                                                _apimail.path = pathFileIO.PathErr;
                                                _apimail.email = form.emailtxt.Text;
                                                _apimail.taxseller = form.txtSellerTaxID.Text;
                                                if (form.pingeng)
                                                {
                                                    _apimail.send_err_service();
                                                }
                                                //sock.SendMailAlert(pathFileIO.PathInput, pathFileIO.PathErr, "F01", form.emailtxt.Text, form.txtSellerTaxID.Text, pathIn + Path.GetFileName(item).Replace("~$", string.Empty), pathIn + Path.GetFileName(item).Replace("~$", string.Empty));
                                                string pathErr = pathFileIO.PathErr + "\\" + Path.GetFileNameWithoutExtension(item) + "_" + strDateTimeStamp + "_Error.txt";
                                                //MessageBox.Show(e.Message);
                                                CreateTextFile(pathErr, ErrorMessage);
                                                MessageBox.Show(ErrorMessage + " 1794");
                                            }
                                            //File.Move(itemPDF, pathFileIO.PathFileRun + "\\" + Path.GetFileNameWithoutExtension(itemPDF)+"_"+strDateTimeStamp+".pdf");
                                            if (File.Exists(pathFileIO.PathFileRun + "\\" + Path.GetFileName(itemPDF)))
                                            {
                                                File.Delete(pathFileIO.PathFileRun + "\\" + Path.GetFileName(itemPDF));
                                                File.Move(itemPDF, pathFileIO.PathFileRun + "\\" + Path.GetFileName(itemPDF));
                                            }
                                            else
                                            {
                                                File.Move(itemPDF, pathFileIO.PathFileRun + "\\" + Path.GetFileName(itemPDF));
                                            }
                                            this.itemPDF = pathFileIO.PathFileRun + "\\" + Path.GetFileName(itemPDF);
                                            WorkProcess(pathFileIO, strFileName, strFileNameEx, this.itemPDF, out iFail, form);
                                            Dispose();
                                        }
                                        catch (IOException e)
                                        {
                                            //MessageBox.Show("a");
                                            pathIn = "";
                                            string[] path = pathFileIO.PathInput.Split('\\');
                                            for (int i = 0; i < path.Length; i++)
                                            {
                                                pathIn += path[i] + "/";
                                            }
                                            Console.WriteLine(pathIn);
                                            this.ErrorMessage = "ไม่สามารถย้ายไฟล์" + pathIn + Path.GetFileName(itemPDF).Replace("~$", string.Empty) + "!!!" + strDateTimeStamp;
                                            MessageBox.Show(ErrorMessage + " 1821");
                                            string itmsg = "ไม่สามารถย้ายไฟล์ ";
                                            _apimail.err_code = "F02";
                                            _apimail.actionmsg = pathIn + Path.GetFileName(itemPDF).Replace("~$", string.Empty);
                                            _apimail.err_msg = pathIn + Path.GetFileName(itemPDF).Replace("~$", string.Empty);
                                            _apimail.input = pathFileIO.PathInput;
                                            _apimail.path = pathFileIO.PathErr;
                                            _apimail.email = form.emailtxt.Text;
                                            _apimail.taxseller = form.txtSellerTaxID.Text;
                                            if (form.pingeng)
                                            {
                                                _apimail.send_err_service();
                                            }
                                            //sock.SendMailAlert(pathFileIO.PathInput, pathFileIO.PathErr, "F02", form.emailtxt.Text, form.txtSellerTaxID.Text, pathIn + Path.GetFileName(itemPDF).Replace("~$", string.Empty), pathIn + Path.GetFileName(itemPDF).Replace("~$", string.Empty));
                                            string pathErr = pathFileIO.PathErr + "\\" + Path.GetFileNameWithoutExtension(itemPDF) + "_" + strDateTimeStamp + "_Error.txt";
                                            //MessageBox.Show(e.Message);
                                            CreateTextFile(pathErr, ErrorMessage);
                                            iSumFail++;
                                            continue;
                                        }
                                        catch (Exception e)
                                        {

                                        }
                                    }
                                }
                                else
                                {
                                    this.itemPDF = "";
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
                                    try
                                    {
                                        if (File.Exists(pathFileIO.PathFileRun + "\\" + Path.GetFileName(item)))
                                        {
                                            File.Delete(pathFileIO.PathFileRun + "\\" + Path.GetFileName(item));
                                            File.Move(item, pathFileIO.PathFileRun + "\\" + Path.GetFileName(item));
                                        }
                                        else
                                        {
                                            File.Move(item, pathFileIO.PathFileRun + "\\" + Path.GetFileName(item));
                                        }
                                        dtParam.PathInput = pathFileIO.PathFileRun + "\\" + Path.GetFileName(item);
                                    }
                                    catch (IOException ex)
                                    {
                                        pathIn = "";
                                        string[] path = pathFileIO.PathInput.Split('\\');
                                        for (int i = 0; i < path.Length; i++)
                                        {
                                            pathIn += path[i] + "/";
                                        }
                                        Console.WriteLine(pathIn);
                                        this.ErrorMessage = "ไม่สามารถย้ายไฟล์" + pathIn + Path.GetFileName(item).Replace("~$", string.Empty) + "!!!" + strDateTimeStamp;

                                        string itmsg = "ไม่สามารถย้ายไฟล์ ";
                                        _apimail.err_code = "F01";
                                        _apimail.actionmsg = pathIn + Path.GetFileName(item).Replace("~$", string.Empty);
                                        _apimail.err_msg = pathIn + Path.GetFileName(item).Replace("~$", string.Empty);
                                        _apimail.input = pathFileIO.PathInput;
                                        _apimail.path = pathFileIO.PathErr;
                                        _apimail.email = form.emailtxt.Text;
                                        _apimail.taxseller = form.txtSellerTaxID.Text;
                                        if (form.pingeng)
                                        {
                                            _apimail.send_err_service();
                                        }
                                        //sock.SendMailAlert(pathFileIO.PathInput, pathFileIO.PathErr, "F01", form.emailtxt.Text, form.txtSellerTaxID.Text, pathIn + Path.GetFileName(item).Replace("~$", string.Empty), pathIn + Path.GetFileName(item).Replace("~$", string.Empty));
                                        string pathErr = pathFileIO.PathErr + "\\" + Path.GetFileNameWithoutExtension(item) + "_" + strDateTimeStamp + "_Error.txt";
                                        //MessageBox.Show(e.Message);
                                        CreateTextFile(pathErr, ErrorMessage);
                                        MessageBox.Show(ErrorMessage + " 1902");
                                    }
                                    continue;
                                }
                                strServiceCode = "S06(Excel Only)";
                            }

                            else if (!strServiceCode.Equals("S03(Excel Only)"))
                            {
                                //MessageBox.Show("S06");
                                var lookupName = arrFilesPDF.ToLookup(x => Path.GetFileNameWithoutExtension(x));
                                if (lookupName.Contains(strFileName))
                                {
                                    var resultJoin = lookupName[strFileName];
                                    foreach (var itemPDF in resultJoin)
                                    {
                                        try
                                        {
                                            try
                                            {
                                                if (File.Exists(pathFileIO.PathFileRun + "\\" + Path.GetFileName(item)))
                                                {
                                                    File.Delete(pathFileIO.PathFileRun + "\\" + Path.GetFileName(item));
                                                    File.Move(item, pathFileIO.PathFileRun + "\\" + Path.GetFileName(item));
                                                }
                                                else
                                                {
                                                    File.Move(item, pathFileIO.PathFileRun + "\\" + Path.GetFileName(item));
                                                }
                                                dtParam.PathInput = pathFileIO.PathFileRun + "\\" + Path.GetFileName(item);
                                            }
                                            catch (IOException ex)
                                            {
                                                pathIn = "";
                                                string[] path = pathFileIO.PathInput.Split('\\');
                                                for (int i = 0; i < path.Length; i++)
                                                {
                                                    pathIn += path[i] + "/";
                                                }
                                                Console.WriteLine(pathIn);
                                                this.ErrorMessage = "ไม่สามารถย้ายไฟล์" + pathIn + Path.GetFileName(item).Replace("~$", string.Empty) + "!!!" + strDateTimeStamp;

                                                string itmsg = "ไม่สามารถย้ายไฟล์ ";
                                                _apimail.err_code = "F01";
                                                _apimail.actionmsg = pathIn + Path.GetFileName(item).Replace("~$", string.Empty);
                                                _apimail.err_msg = pathIn + Path.GetFileName(item).Replace("~$", string.Empty);
                                                _apimail.input = pathFileIO.PathInput;
                                                _apimail.path = pathFileIO.PathErr;
                                                _apimail.email = form.emailtxt.Text;
                                                _apimail.taxseller = form.txtSellerTaxID.Text;
                                                if (form.pingeng)
                                                {
                                                    _apimail.send_err_service();
                                                }
                                                //sock.SendMailAlert(pathFileIO.PathInput, pathFileIO.PathErr, "F01", form.emailtxt.Text, form.txtSellerTaxID.Text, pathIn + Path.GetFileName(item).Replace("~$", string.Empty), pathIn + Path.GetFileName(item).Replace("~$", string.Empty));
                                                string pathErr = pathFileIO.PathErr + "\\" + Path.GetFileNameWithoutExtension(item) + "_" + strDateTimeStamp + "_Error.txt";
                                                //MessageBox.Show(e.Message);
                                                CreateTextFile(pathErr, ErrorMessage);
                                                MessageBox.Show(ErrorMessage + " 1961");
                                            }
                                            //File.Move(itemPDF, pathFileIO.PathFileRun + "\\" + Path.GetFileNameWithoutExtension(itemPDF)+"_"+strDateTimeStamp+".pdf");
                                            if (File.Exists(pathFileIO.PathFileRun + "\\" + Path.GetFileName(itemPDF)))
                                            {
                                                File.Delete(pathFileIO.PathFileRun + "\\" + Path.GetFileName(itemPDF));
                                                File.Move(itemPDF, pathFileIO.PathFileRun + "\\" + Path.GetFileName(itemPDF));
                                            }
                                            else
                                            {
                                                File.Move(itemPDF, pathFileIO.PathFileRun + "\\" + Path.GetFileName(itemPDF));
                                            }
                                            this.itemPDF = pathFileIO.PathFileRun + "\\" + Path.GetFileName(itemPDF);
                                            WorkProcess(pathFileIO, strFileName, strFileNameEx, this.itemPDF, out iFail, form);
                                            Dispose();
                                        }
                                        catch (IOException e)
                                        {
                                            //MessageBox.Show("a");
                                            pathIn = "";
                                            string[] path = pathFileIO.PathInput.Split('\\');
                                            for (int i = 0; i < path.Length; i++)
                                            {
                                                pathIn += path[i] + "/";
                                            }
                                            Console.WriteLine(pathIn);
                                            this.ErrorMessage = "ไม่สามารถย้ายไฟล์" + pathIn + Path.GetFileName(itemPDF).Replace("~$", string.Empty) + "!!!" + strDateTimeStamp;
                                            //MessageBox.Show(ErrorMessage);
                                            string itmsg = "ไม่สามารถย้ายไฟล์ ";
                                            _apimail.err_code = "F02";
                                            _apimail.actionmsg = pathIn + Path.GetFileName(itemPDF).Replace("~$", string.Empty);
                                            _apimail.err_msg = pathIn + Path.GetFileName(itemPDF).Replace("~$", string.Empty);
                                            _apimail.input = pathFileIO.PathInput;
                                            _apimail.path = pathFileIO.PathErr;
                                            _apimail.email = form.emailtxt.Text;
                                            _apimail.taxseller = form.txtSellerTaxID.Text;
                                            if (form.pingeng)
                                            {
                                                _apimail.send_err_service();
                                            }
                                            //sock.SendMailAlert(pathFileIO.PathInput, pathFileIO.PathErr, "F02", form.emailtxt.Text, form.txtSellerTaxID.Text, pathIn + Path.GetFileName(itemPDF).Replace("~$", string.Empty), pathIn + Path.GetFileName(itemPDF).Replace("~$", string.Empty));
                                            string pathErr = pathFileIO.PathErr + "\\" + Path.GetFileNameWithoutExtension(itemPDF) + "_" + strDateTimeStamp + "_Error.txt";
                                            //MessageBox.Show(e.Message);
                                            CreateTextFile(pathErr, ErrorMessage);

                                        }
                                        catch (Exception e)
                                        {

                                        }
                                    }
                                    //MessageBox.Show(dtParam.PathInput);
                                }
                                else if (strServiceCode.Equals("S03(Excel Only)"))
                                {


                                }
                                else
                                {
                                    this.itemPDF = "";
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
                                    try
                                    {
                                        if (File.Exists(pathFileIO.PathFileRun + "\\" + Path.GetFileName(item)))
                                        {
                                            File.Delete(pathFileIO.PathFileRun + "\\" + Path.GetFileName(item));
                                            File.Move(item, pathFileIO.PathFileRun + "\\" + Path.GetFileName(item));
                                        }
                                        else
                                        {
                                            File.Move(item, pathFileIO.PathFileRun + "\\" + Path.GetFileName(item));
                                        }
                                        dtParam.PathInput = pathFileIO.PathFileRun + "\\" + Path.GetFileName(item);
                                    }
                                    catch (Exception ex)
                                    {
                                        pathIn = "";
                                        string[] path = pathFileIO.PathInput.Split('\\');
                                        for (int i = 0; i < path.Length; i++)
                                        {
                                            pathIn += path[i] + "/";
                                        }
                                        Console.WriteLine(pathIn);
                                        this.ErrorMessage = "ไม่สามารถย้ายไฟล์" + pathIn + Path.GetFileName(item).Replace("~$", string.Empty) + "!!!" + strDateTimeStamp;

                                        string itmsg = "ไม่สามารถย้ายไฟล์ ";
                                        _apimail.err_code = "F01";
                                        _apimail.actionmsg = pathIn + Path.GetFileName(item).Replace("~$", string.Empty);
                                        _apimail.err_msg = pathIn + Path.GetFileName(item).Replace("~$", string.Empty);
                                        _apimail.input = pathFileIO.PathInput;
                                        _apimail.path = pathFileIO.PathErr;
                                        _apimail.email = form.emailtxt.Text;
                                        _apimail.taxseller = form.txtSellerTaxID.Text;
                                        if (form.pingeng)
                                        {
                                            _apimail.send_err_service();
                                        }
                                        //sock.SendMailAlert(pathFileIO.PathInput, pathFileIO.PathErr, "F01", form.emailtxt.Text, form.txtSellerTaxID.Text, pathIn + Path.GetFileName(item).Replace("~$", string.Empty), pathIn + Path.GetFileName(item).Replace("~$", string.Empty));
                                        string pathErr = pathFileIO.PathErr + "\\" + Path.GetFileNameWithoutExtension(item) + "_" + strDateTimeStamp + "_Error.txt";
                                        //MessageBox.Show(e.Message);
                                        CreateTextFile(pathErr, ErrorMessage);
                                        MessageBox.Show(ErrorMessage + " 2070");
                                    }
                                    iSumFail++;
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
                    }
                    //string toDisplay = string.Join(Environment.NewLine, arrFilesTxt);
                    //int limit = arrFilesTxt.Count() / Thread_num;
                    ////MessageBox.Show(Thread_num.ToString());
                    ////MessageBox.Show(arrFilesTxt.Count().ToString());
                    //Thread[] objThread = new Thread[Thread_num];
                    //if (arrFilesTxt.Count() > 0)
                    //{
                    //    for (int i = 0; i < Thread_num; i++)
                    //    {
                    //        int first = i * limit;
                    //        int second;
                    //        if (i != (Thread_num - 1))
                    //        {
                    //            second = limit;
                    //        }
                    //        else
                    //        {
                    //            second = arrFilesTxt.Count() - (limit * i);
                    //        }
                    //        objThread[i] = new Thread(() => delegates(arrFilesTxt, dtParam, pathFileIO, lookupFileLog, arrFilesPDF, form, valueReturnChk, strServiceCode, first, second));
                    //        objThread[i].Priority = ThreadPriority.AboveNormal;
                    //        objThread[i].Start();
                    //    }
                    //}
                    //for (int i = 0; i < objThread.Length; i++)
                    //{
                    //    // Wait until thread is finished.
                    //    objThread[i].Join();
                    //}
                    foreach (var item in arrFilesTxt)
                    {
                        int threadID = (int)AppDomain.GetCurrentThreadId();
                        fileContent = File.ReadAllText(item);
                        content = fileContent.Split('\n');
                        lengthOfconnectC = content[0].Split(',');
                        //Console.WriteLine(lengthOfconnectC.Length);
                        //Console.WriteLine(content[2]);
                        if (lengthOfconnectC.Length == 5 || lengthOfconnectC.Length == 4)
                        {
                            string[] contentH = content[1].Split(',');
                            string patternChkString = @"([a-zA-Zก-๙0-9/])";
                            bool chkSting = false;
                            chkSting = Regex.IsMatch(contentH[2], patternChkString);
                            if (chkSting == false)
                            {
                                ConvertAnsiToUTF8(item, item);
                            }
                            Console.WriteLine(chkSting);
                        }
                        else
                        {
                            fileContent = File.ReadAllText(item);
                            content = fileContent.Split('\r');
                            lengthOfconnectC = content[0].Split(',');
                            if (lengthOfconnectC.Length == 5 || lengthOfconnectC.Length == 4)
                            {
                                string[] contentH = content[1].Split(',');
                                string patternChkString = @"([a-zA-Zก-๙0-9/])";
                                bool chkSting = false;
                                chkSting = Regex.IsMatch(contentH[2], patternChkString);
                                if (chkSting == false)
                                {
                                    ConvertAnsiToUTF8(item, item);
                                }
                                Console.WriteLine(chkSting);
                            }
                        }
                        if (lookupFileLog.Contains(Path.GetFileName(item)))
                        {
                            if (iCountRunFile > Int32.Parse(dtParam.AmountFile))
                            {
                                break;
                            }
                            dtParam.PathInput = string.Empty;
                            dtParam.PathInput = item;
                            arrFilesPDF = System.IO.Directory.GetFiles(pathFileIO.PathInput, "*.pdf");
                            string strFileName = Path.GetFileNameWithoutExtension(item);
                            string strFileNameEx = Path.GetFileName(item);
                            int iFail = 0;

                            //System.Threading.Thread.Sleep(500);
                            if (strServiceCode.Equals("S03"))
                            {
                                try
                                {
                                    Console.WriteLine(pathFileIO.PathFileRun);
                                    if (File.Exists(pathFileIO.PathFileRun + "\\" + Path.GetFileName(item)))
                                    {
                                        File.Delete(pathFileIO.PathFileRun + "\\" + Path.GetFileName(item));
                                        File.Move(item, pathFileIO.PathFileRun + "\\" + Path.GetFileName(item));
                                    }
                                    else
                                    {
                                        File.Move(item, pathFileIO.PathFileRun + "\\" + Path.GetFileName(item));
                                    }
                                    dtParam.PathInput = pathFileIO.PathFileRun + "\\" + Path.GetFileName(item);
                                    WorkProcess(pathFileIO, strFileName, strFileNameEx, "", out iFail, form);
                                    Dispose();
                                }
                                catch (Exception ex)
                                {
                                    pathIn = "";
                                    string[] path = pathFileIO.PathInput.Split('\\');
                                    for (int i = 0; i < path.Length; i++)
                                    {
                                        pathIn += path[i] + "/";
                                    }
                                    Console.WriteLine(pathIn);
                                    this.ErrorMessage = "ไม่สามารถย้ายไฟล์" + pathIn + Path.GetFileName(item).Replace("~$", string.Empty) + "!!!" + strDateTimeStamp;

                                    string itmsg = "ไม่สามารถย้ายไฟล์ ";
                                    _apimail.err_code = "F01";
                                    _apimail.actionmsg = pathIn + Path.GetFileName(item).Replace("~$", string.Empty);
                                    _apimail.err_msg = pathIn + Path.GetFileName(item).Replace("~$", string.Empty);
                                    _apimail.input = pathFileIO.PathInput;
                                    _apimail.path = pathFileIO.PathErr;
                                    _apimail.email = form.emailtxt.Text;
                                    _apimail.taxseller = form.txtSellerTaxID.Text;
                                    if (form.pingeng)
                                    {
                                        _apimail.send_err_service();
                                    }
                                    //sock.SendMailAlert(pathFileIO.PathInput, pathFileIO.PathErr, "F01", form.emailtxt.Text, form.txtSellerTaxID.Text, pathIn + Path.GetFileName(item).Replace("~$", string.Empty), pathIn + Path.GetFileName(item).Replace("~$", string.Empty));
                                    string pathErr = pathFileIO.PathErr + "\\" + Path.GetFileNameWithoutExtension(item) + "_" + strDateTimeStamp + "_Error.txt";
                                    //MessageBox.Show(e.Message);
                                    CreateTextFile(pathErr, ErrorMessage);
                                    MessageBox.Show(ErrorMessage);
                                }
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
                                                dtParam.PathInput = pathFileIO.PathFileRun + "\\" + Path.GetFileName(item);
                                            }
                                            catch (Exception e)
                                            {
                                                pathIn = "";
                                                string[] path = pathFileIO.PathInput.Split('\\');
                                                for (int i = 0; i < path.Length; i++)
                                                {
                                                    pathIn += path[i] + "/";
                                                }
                                                Console.WriteLine(pathIn);
                                                this.ErrorMessage = "ไม่สามารถย้ายไฟล์" + pathIn + Path.GetFileName(item).Replace("~$", string.Empty) + "!!!" + strDateTimeStamp;

                                                string itmsg = "ไม่สามารถย้ายไฟล์ ";
                                                _apimail.err_code = "F01";
                                                _apimail.actionmsg = pathIn + Path.GetFileName(item).Replace("~$", string.Empty);
                                                _apimail.err_msg = pathIn + Path.GetFileName(item).Replace("~$", string.Empty);
                                                _apimail.input = pathFileIO.PathInput;
                                                _apimail.path = pathFileIO.PathErr;
                                                _apimail.email = form.emailtxt.Text;
                                                _apimail.taxseller = form.txtSellerTaxID.Text;
                                                if (form.pingeng)
                                                {
                                                    _apimail.send_err_service();
                                                }
                                                //sock.SendMailAlert(pathFileIO.PathInput, pathFileIO.PathErr, "F01", form.emailtxt.Text, form.txtSellerTaxID.Text, pathIn + Path.GetFileName(item).Replace("~$", string.Empty), pathIn + Path.GetFileName(item).Replace("~$", string.Empty));
                                                string pathErr = pathFileIO.PathErr + "\\" + Path.GetFileNameWithoutExtension(item) + "_" + strDateTimeStamp + "_Error.txt";
                                                //MessageBox.Show(e.Message);
                                                CreateTextFile(pathErr, ErrorMessage);
                                                //goto Loop;
                                            }
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
                                            this.itemPDF = pathFileIO.PathFileRun + "\\" + Path.GetFileName(itemPDF);
                                            WorkProcess(pathFileIO, strFileName, strFileNameEx, this.itemPDF, out iFail, form);
                                            Dispose();
                                        }
                                        catch (Exception e)
                                        {
                                            pathIn = "";
                                            string[] path = pathFileIO.PathInput.Split('\\');
                                            for (int i = 0; i < path.Length; i++)
                                            {
                                                pathIn += path[i] + "/";
                                            }
                                            Console.WriteLine(pathIn);
                                            this.ErrorMessage = "ไม่สามารถย้ายไฟล์" + pathIn + Path.GetFileName(itemPDF).Replace("~$", string.Empty) + "!!!" + strDateTimeStamp;
                                            string itmsg = "ไม่สามารถย้ายไฟล์ ";
                                            _apimail.err_code = "F02";
                                            _apimail.actionmsg = pathIn + Path.GetFileName(itemPDF).Replace("~$", string.Empty);
                                            _apimail.err_msg = pathIn + Path.GetFileName(itemPDF).Replace("~$", string.Empty);
                                            _apimail.input = pathFileIO.PathInput;
                                            _apimail.path = pathFileIO.PathErr;
                                            _apimail.email = form.emailtxt.Text;
                                            _apimail.taxseller = form.txtSellerTaxID.Text;
                                            if (form.pingeng)
                                            {
                                                _apimail.send_err_service();
                                            }
                                            //sock.SendMailAlert(pathFileIO.PathInput, pathFileIO.PathErr, "F02", form.emailtxt.Text, form.txtSellerTaxID.Text, pathIn + Path.GetFileName(itemPDF).Replace("~$", string.Empty), pathIn + Path.GetFileName(itemPDF).Replace("~$", string.Empty));
                                            string pathErr = pathFileIO.PathErr + "\\" + Path.GetFileNameWithoutExtension(itemPDF) + "_" + strDateTimeStamp + "_Error.txt";
                                            //MessageBox.Show(e.Message);
                                            CreateTextFile(pathErr, ErrorMessage);
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
                                    catch (Exception e)
                                    {
                                        pathIn = "";
                                        string[] path = pathFileIO.PathInput.Split('\\');
                                        for (int i = 0; i < path.Length; i++)
                                        {
                                            pathIn += path[i] + "/";
                                        }
                                        Console.WriteLine(pathIn);
                                        this.ErrorMessage = "ไม่สามารถย้ายไฟล์" + pathIn + Path.GetFileName(item).Replace("~$", string.Empty) + "!!!" + strDateTimeStamp;

                                        string itmsg = "ไม่สามารถย้ายไฟล์ ";
                                        _apimail.err_code = "F01";
                                        _apimail.actionmsg = pathIn + Path.GetFileName(item).Replace("~$", string.Empty);
                                        _apimail.err_msg = pathIn + Path.GetFileName(item).Replace("~$", string.Empty);
                                        _apimail.input = pathFileIO.PathInput;
                                        _apimail.path = pathFileIO.PathErr;
                                        _apimail.email = form.emailtxt.Text;
                                        _apimail.taxseller = form.txtSellerTaxID.Text;
                                        if (form.pingeng)
                                        {
                                            _apimail.send_err_service();
                                        }
                                        //sock.SendMailAlert(pathFileIO.PathInput, pathFileIO.PathErr, "F01", form.emailtxt.Text, form.txtSellerTaxID.Text, pathIn + Path.GetFileName(item).Replace("~$", string.Empty), pathIn + Path.GetFileName(item).Replace("~$", string.Empty));
                                        string pathErr = pathFileIO.PathErr + "\\" + Path.GetFileNameWithoutExtension(item) + "_" + strDateTimeStamp + "_Error.txt";
                                        //MessageBox.Show(e.Message);
                                        CreateTextFile(pathErr, ErrorMessage);
                                    }
                                    iSumFail++;
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
                         
                    }

                    foreach (var item in arrFilesCsv)
                    {
                        if (lookupFileLog.Contains(Path.GetFileName(item)))
                        {
                            if (iCountRunFile > Int32.Parse(dtParam.AmountFile))
                            {
                                break;
                            }
                            dtParam.PathInput = string.Empty;
                            dtParam.PathInput = item;
                            arrFilesPDF = System.IO.Directory.GetFiles(pathFileIO.PathInput, "*.pdf");
                            string strFileName = Path.GetFileNameWithoutExtension(item);
                            string strFileNameEx = Path.GetFileName(item);
                            int iFail = 0;

                            System.Threading.Thread.Sleep(500);

                            if (strServiceCode.Equals("S03"))
                            {
                                try
                                {
                                    Console.WriteLine(pathFileIO.PathFileRun);
                                    if (File.Exists(pathFileIO.PathFileRun + "\\" + Path.GetFileName(item)))
                                    {
                                        File.Delete(pathFileIO.PathFileRun + "\\" + Path.GetFileName(item));
                                        File.Move(item, pathFileIO.PathFileRun + "\\" + Path.GetFileName(item));
                                    }
                                    else
                                    {
                                        File.Move(item, pathFileIO.PathFileRun + "\\" + Path.GetFileName(item));
                                    }
                                    dtParam.PathInput = pathFileIO.PathFileRun + "\\" + Path.GetFileName(item);
                                    WorkProcess(pathFileIO, strFileName, strFileNameEx, "", out iFail, form);
                                    Dispose();
                                }
                                catch (Exception ex)
                                {
                                    pathIn = "";
                                    string[] path = pathFileIO.PathInput.Split('\\');
                                    for (int i = 0; i < path.Length; i++)
                                    {
                                        pathIn += path[i] + "/";
                                    }
                                    Console.WriteLine(pathIn);
                                    this.ErrorMessage = "ไม่สามารถย้ายไฟล์" + pathIn + Path.GetFileName(item).Replace("~$", string.Empty) + "!!!" + strDateTimeStamp;

                                    string itmsg = "ไม่สามารถย้ายไฟล์ ";
                                    _apimail.err_code = "F01";
                                    _apimail.actionmsg = pathIn + Path.GetFileName(item).Replace("~$", string.Empty);
                                    _apimail.err_msg = pathIn + Path.GetFileName(item).Replace("~$", string.Empty);
                                    _apimail.input = pathFileIO.PathInput;
                                    _apimail.path = pathFileIO.PathErr;
                                    _apimail.email = form.emailtxt.Text;
                                    _apimail.taxseller = form.txtSellerTaxID.Text;
                                    if (form.pingeng)
                                    {
                                        _apimail.send_err_service();
                                    }
                                    //sock.SendMailAlert(pathFileIO.PathInput, pathFileIO.PathErr, "F01", form.emailtxt.Text, form.txtSellerTaxID.Text, pathIn + Path.GetFileName(item).Replace("~$", string.Empty), pathIn + Path.GetFileName(item).Replace("~$", string.Empty));
                                    string pathErr = pathFileIO.PathErr + "\\" + Path.GetFileNameWithoutExtension(item) + "_" + strDateTimeStamp + "_Error.txt";
                                    //MessageBox.Show(e.Message);
                                    CreateTextFile(pathErr, ErrorMessage);
                                }
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
                                                dtParam.PathInput = pathFileIO.PathFileRun + "\\" + Path.GetFileName(item);
                                            }
                                            catch (Exception e)
                                            {
                                                pathIn = "";
                                                string[] path = pathFileIO.PathInput.Split('\\');
                                                for (int i = 0; i < path.Length; i++)
                                                {
                                                    pathIn += path[i] + "/";
                                                }
                                                Console.WriteLine(pathIn);
                                                this.ErrorMessage = "ไม่สามารถย้ายไฟล์" + pathIn + Path.GetFileName(item).Replace("~$", string.Empty) + "!!!" + strDateTimeStamp;

                                                string itmsg = "ไม่สามารถย้ายไฟล์ ";
                                                _apimail.err_code = "F01";
                                                _apimail.actionmsg = pathIn + Path.GetFileName(item).Replace("~$", string.Empty);
                                                _apimail.err_msg = pathIn + Path.GetFileName(item).Replace("~$", string.Empty);
                                                _apimail.input = pathFileIO.PathInput;
                                                _apimail.path = pathFileIO.PathErr;
                                                _apimail.email = form.emailtxt.Text;
                                                _apimail.taxseller = form.txtSellerTaxID.Text;
                                                if (form.pingeng)
                                                {
                                                    _apimail.send_err_service();
                                                }
                                                //sock.SendMailAlert(pathFileIO.PathInput, pathFileIO.PathErr, "F01", form.emailtxt.Text, form.txtSellerTaxID.Text, pathIn + Path.GetFileName(item).Replace("~$", string.Empty), pathIn + Path.GetFileName(item).Replace("~$", string.Empty));
                                                string pathErr = pathFileIO.PathErr + "\\" + Path.GetFileNameWithoutExtension(item) + "_" + strDateTimeStamp + "_Error.txt";
                                                //MessageBox.Show(e.Message);
                                                CreateTextFile(pathErr, ErrorMessage);
                                                MessageBox.Show(ErrorMessage);
                                                goto Loop;
                                            }
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
                                            this.itemPDF = pathFileIO.PathFileRun + "\\" + Path.GetFileName(itemPDF);
                                            WorkProcess(pathFileIO, strFileName, strFileNameEx, this.itemPDF, out iFail, form);
                                            Dispose();
                                        }
                                        catch (Exception e)
                                        {
                                            pathIn = "";
                                            string[] path = pathFileIO.PathInput.Split('\\');
                                            for (int i = 0; i < path.Length; i++)
                                            {
                                                pathIn += path[i] + "/";
                                            }
                                            Console.WriteLine(pathIn);
                                            this.ErrorMessage = "ไม่สามารถย้ายไฟล์" + pathIn + Path.GetFileName(itemPDF).Replace("~$", string.Empty) + "!!!" + strDateTimeStamp;
                                            string itmsg = "ไม่สามารถย้ายไฟล์ ";
                                            _apimail.err_code = "F02";
                                            _apimail.actionmsg = pathIn + Path.GetFileName(itemPDF).Replace("~$", string.Empty);
                                            _apimail.err_msg = pathIn + Path.GetFileName(itemPDF).Replace("~$", string.Empty);
                                            _apimail.input = pathFileIO.PathInput;
                                            _apimail.path = pathFileIO.PathErr;
                                            _apimail.email = form.emailtxt.Text;
                                            _apimail.taxseller = form.txtSellerTaxID.Text;
                                            if (form.pingeng)
                                            {
                                                _apimail.send_err_service();
                                            }
                                            //sock.SendMailAlert(pathFileIO.PathInput, pathFileIO.PathErr, "F02", form.emailtxt.Text, form.txtSellerTaxID.Text, pathIn + Path.GetFileName(itemPDF).Replace("~$", string.Empty), pathIn + Path.GetFileName(itemPDF).Replace("~$", string.Empty));
                                            string pathErr = pathFileIO.PathErr + "\\" + Path.GetFileNameWithoutExtension(itemPDF) + "_" + strDateTimeStamp + "_Error.txt";
                                            //MessageBox.Show(e.Message);
                                            CreateTextFile(pathErr, ErrorMessage);
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
                                    catch (Exception e)
                                    {
                                        pathIn = "";
                                        string[] path = pathFileIO.PathInput.Split('\\');
                                        for (int i = 0; i < path.Length; i++)
                                        {
                                            pathIn += path[i] + "/";
                                        }
                                        Console.WriteLine(pathIn);
                                        this.ErrorMessage = "ไม่สามารถย้ายไฟล์" + pathIn + Path.GetFileName(item).Replace("~$", string.Empty) + "!!!" + strDateTimeStamp;

                                        string itmsg = "ไม่สามารถย้ายไฟล์ ";
                                        _apimail.err_code = "F01";
                                        _apimail.actionmsg = pathIn + Path.GetFileName(item).Replace("~$", string.Empty);
                                        _apimail.err_msg = pathIn + Path.GetFileName(item).Replace("~$", string.Empty);
                                        _apimail.input = pathFileIO.PathInput;
                                        _apimail.path = pathFileIO.PathErr;
                                        _apimail.email = form.emailtxt.Text;
                                        _apimail.taxseller = form.txtSellerTaxID.Text;
                                        if (form.pingeng)
                                        {
                                            _apimail.send_err_service();
                                        }
                                        //sock.SendMailAlert(pathFileIO.PathInput, pathFileIO.PathErr, "F01", form.emailtxt.Text, form.txtSellerTaxID.Text, pathIn + Path.GetFileName(item).Replace("~$", string.Empty), pathIn + Path.GetFileName(item).Replace("~$", string.Empty));
                                        string pathErr = pathFileIO.PathErr + "\\" + Path.GetFileNameWithoutExtension(item) + "_" + strDateTimeStamp + "_Error.txt";
                                        //MessageBox.Show(e.Message);
                                        CreateTextFile(pathErr, ErrorMessage);
                                        MessageBox.Show(ErrorMessage);
                                    }
                                    iSumFail++;
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
                        //try
                        //{
                        //    //File.Move(item, pathFileIO.PathFileRun + "\\" + Path.GetFileNameWithoutExtension(item) + "_" + strDateTimeStamp + ".txt");
                        //    if (File.Exists(pathFileIO.PathFileRun + "\\" + Path.GetFileName(item)))
                        //    {
                        //        File.Delete(pathFileIO.PathFileRun + "\\" + Path.GetFileName(item));
                        //        File.Move(item, pathFileIO.PathFileRun + "\\" + Path.GetFileName(item));
                        //    }
                        //    else
                        //    {
                        //        File.Move(item, pathFileIO.PathFileRun + "\\" + Path.GetFileName(item));
                        //    }
                        //}
                        //catch (Exception e)
                        //{
                        //    pathIn = "";
                        //    string[] path = pathFileIO.PathInput.Split('\\');
                        //    for (int i = 0; i < path.Length; i++)
                        //    {
                        //        pathIn += path[i] + "/";
                        //    }
                        //    Console.WriteLine(pathIn);
                        //    this.ErrorMessage = "ไม่สามารถย้ายไฟล์" + pathIn + Path.GetFileName(item).Replace("~$", string.Empty) + "!!!" + strDateTimeStamp;
                        //    string itmsg = "ไม่สามารถย้ายไฟล์ ";
                        //    sock.SendMailAlert(pathFileIO.PathInput, pathFileIO.PathErr, "F01", form.emailtxt.Text, form.txtSellerTaxID.Text,Path.GetFileName(item).Replace("~$",string.Empty), pathIn + Path.GetFileName(item).Replace("~$", string.Empty));
                        //    string pathErr = pathFileIO.PathErr + "\\" + Path.GetFileNameWithoutExtension(item) + "_" + strDateTimeStamp + "_ErrorFile.txt";
                        //    //MessageBox.Show(e.Message);
                        //    CreateTextFile(pathErr, ErrorMessage);
                        //}
                    }

                }


            }
            catch (Exception exc)
            {
                //MessageBox.Show("Method btnExport_Click Error: " + exc.ToString());
            }
            finally
            {
                for (int i = 0; i <= GC.MaxGeneration; i++)
                {
                    int count = GC.CollectionCount(i);
                    GC.Collect();
                }
                GC.WaitForPendingFinalizers();
                GC.SuppressFinalize(this);
            }
            if (valueReturnChk.CountFileRun == valueReturnChk.AmountAllFile)
            {
                valueReturnChk.StatusRunning = true;
            }
            else
            {
                valueReturnChk.StatusRunning = false;
            }

            form.txtStatus.Text += Environment.NewLine + "Summary:" + Environment.NewLine + "   -Success " + iSumSuccess + Environment.NewLine + "   -Fail " + iSumFail;
            strTempLogTime += Environment.NewLine + "Summary: Success " + iSumSuccess + " Fail " + iSumFail;

            if (chkOption == true)
            {
                //Write log time
                CreateTextFile(pathFileIO.PathLogTime + "\\LogTime" + strDateTimeStamp + ".txt", strTempLogTime);
                //Console.WriteLine(pathFileIO.PathLogTime + "\\LogTime" + strDateTimeStamp + ".txt", strTempLogTime);
            }
            else if (chkOption == false)
            {
                //Write log time
                CreateTextFile(pathFileIO.PathLogTime + "\\LogTime" + strDateTimeStamp + ".txt", strTempLogTime);
                //Console.WriteLine(pathFileIO.PathLogTime + "\\LogTime" + strDateTimeStamp + ".txt", strTempLogTime);
            }

            iSumFail = 0;
            iSumSuccess = 0;
            strTempLogTime = string.Empty;
            return valueReturnChk;
        }

        public void delegates (string[] item, DtGetParameters dtParam, PathFilesIO pathFileIO, dynamic lookupFileLog, string[] arrFilesPDF, etaxOneth form, ValueReturnForm valueReturnChk, string strServiceCode,int round_start,int round_end)
        {
            int round_starts = round_start;
            int round_ends = round_start + round_end;
            //MessageBox.Show("round_start => " + round_starts.ToString() + " round_end => " + round_ends.ToString());
            if (round_ends > round_starts)
            {
                for (int p = round_starts; p < round_ends; p++)
                {
                    //MessageBox.Show("DO");
                    runtxt(item[p], dtParam, pathFileIO, lookupFileLog, arrFilesPDF, form, valueReturnChk, strServiceCode);
                    Thread.Sleep(500);
                }
            }
        }
        public void runtxt(string item, DtGetParameters dtParam, PathFilesIO pathFileIO, dynamic lookupFileLog,string[] arrFilesPDF, etaxOneth form, ValueReturnForm valueReturnChk, string strServiceCode)
        {
            string fileContent;
            string[] content;
            string[] lengthOfconnectC;
            string BathText = "";
            
                int threadID = (int)AppDomain.GetCurrentThreadId();
                fileContent = File.ReadAllText(item);
                content = fileContent.Split('\n');
                lengthOfconnectC = content[0].Split(',');
                //Console.WriteLine(lengthOfconnectC.Length);
                //Console.WriteLine(content[2]);
                if (lengthOfconnectC.Length == 5 || lengthOfconnectC.Length == 4)
                {
                    string[] contentH = content[1].Split(',');
                    string patternChkString = @"([a-zA-Zก-๙0-9/])";
                    bool chkSting = false;
                    chkSting = Regex.IsMatch(contentH[2], patternChkString);
                    if (chkSting == false)
                    {
                        ConvertAnsiToUTF8(item, item);
                    }
                    Console.WriteLine(chkSting);
                }
                else
                {
                    fileContent = File.ReadAllText(item);
                    content = fileContent.Split('\r');
                    lengthOfconnectC = content[0].Split(',');
                    if (lengthOfconnectC.Length == 5 || lengthOfconnectC.Length == 4)
                    {
                        string[] contentH = content[1].Split(',');
                        string patternChkString = @"([a-zA-Zก-๙0-9/])";
                        bool chkSting = false;
                        chkSting = Regex.IsMatch(contentH[2], patternChkString);
                        if (chkSting == false)
                        {
                            ConvertAnsiToUTF8(item, item);
                        }
                        Console.WriteLine(chkSting);
                    }
                }
                if (lookupFileLog.Contains(Path.GetFileName(item)))
                {
                    //if (iCountRunFile > Int32.Parse(dtParam.AmountFile))
                    //{
                    //    break;
                    //}
                    dtParam.PathInput = string.Empty;
                    dtParam.PathInput = item;
                    arrFilesPDF = System.IO.Directory.GetFiles(pathFileIO.PathInput, "*.pdf");
                    string strFileName = Path.GetFileNameWithoutExtension(item);
                    string strFileNameEx = Path.GetFileName(item);
                    int iFail = 0;

                    System.Threading.Thread.Sleep(500);
                    if (strServiceCode.Equals("S03"))
                    {
                        try
                        {
                            Console.WriteLine(pathFileIO.PathFileRun);
                            if (File.Exists(pathFileIO.PathFileRun + "\\" + Path.GetFileName(item)))
                            {
                                File.Delete(pathFileIO.PathFileRun + "\\" + Path.GetFileName(item));
                                File.Move(item, pathFileIO.PathFileRun + "\\" + Path.GetFileName(item));
                            }
                            else
                            {
                                File.Move(item, pathFileIO.PathFileRun + "\\" + Path.GetFileName(item));
                            }
                            dtParam.PathInput = pathFileIO.PathFileRun + "\\" + Path.GetFileName(item);
                            WorkProcess(pathFileIO, strFileName, strFileNameEx, "", out iFail, form);
                            Dispose();
                        }
                        catch (Exception ex)
                        {
                            pathIn = "";
                            string[] path = pathFileIO.PathInput.Split('\\');
                            for (int i = 0; i < path.Length; i++)
                            {
                                pathIn += path[i] + "/";
                            }
                            Console.WriteLine(pathIn);
                            this.ErrorMessage = "ไม่สามารถย้ายไฟล์" + pathIn + Path.GetFileName(item).Replace("~$", string.Empty) + "!!!" + strDateTimeStamp;

                            string itmsg = "ไม่สามารถย้ายไฟล์ ";
                            _apimail.err_code = "F01";
                            _apimail.actionmsg = pathIn + Path.GetFileName(item).Replace("~$", string.Empty);
                            _apimail.err_msg = pathIn + Path.GetFileName(item).Replace("~$", string.Empty);
                            _apimail.input = pathFileIO.PathInput;
                            _apimail.path = pathFileIO.PathErr;
                            _apimail.email = form.emailtxt.Text;
                            _apimail.taxseller = form.txtSellerTaxID.Text;
                            if (form.pingeng)
                            {
                                _apimail.send_err_service();
                            }
                            //sock.SendMailAlert(pathFileIO.PathInput, pathFileIO.PathErr, "F01", form.emailtxt.Text, form.txtSellerTaxID.Text, pathIn + Path.GetFileName(item).Replace("~$", string.Empty), pathIn + Path.GetFileName(item).Replace("~$", string.Empty));
                            string pathErr = pathFileIO.PathErr + "\\" + Path.GetFileNameWithoutExtension(item) + "_" + strDateTimeStamp + "_Error.txt";
                            //MessageBox.Show(e.Message);
                            CreateTextFile(pathErr, ErrorMessage);
                            MessageBox.Show(ErrorMessage);
                        }
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
                                        dtParam.PathInput = pathFileIO.PathFileRun + "\\" + Path.GetFileName(item);
                                    }
                                    catch (Exception e)
                                    {
                                        pathIn = "";
                                        string[] path = pathFileIO.PathInput.Split('\\');
                                        for (int i = 0; i < path.Length; i++)
                                        {
                                            pathIn += path[i] + "/";
                                        }
                                        Console.WriteLine(pathIn);
                                        this.ErrorMessage = "ไม่สามารถย้ายไฟล์" + pathIn + Path.GetFileName(item).Replace("~$", string.Empty) + "!!!" + strDateTimeStamp;

                                        string itmsg = "ไม่สามารถย้ายไฟล์ ";
                                        _apimail.err_code = "F01";
                                        _apimail.actionmsg = pathIn + Path.GetFileName(item).Replace("~$", string.Empty);
                                        _apimail.err_msg = pathIn + Path.GetFileName(item).Replace("~$", string.Empty);
                                        _apimail.input = pathFileIO.PathInput;
                                        _apimail.path = pathFileIO.PathErr;
                                        _apimail.email = form.emailtxt.Text;
                                        _apimail.taxseller = form.txtSellerTaxID.Text;
                                        if (form.pingeng)
                                        {
                                            _apimail.send_err_service();
                                        }
                                        //sock.SendMailAlert(pathFileIO.PathInput, pathFileIO.PathErr, "F01", form.emailtxt.Text, form.txtSellerTaxID.Text, pathIn + Path.GetFileName(item).Replace("~$", string.Empty), pathIn + Path.GetFileName(item).Replace("~$", string.Empty));
                                        string pathErr = pathFileIO.PathErr + "\\" + Path.GetFileNameWithoutExtension(item) + "_" + strDateTimeStamp + "_Error.txt";
                                        //MessageBox.Show(e.Message);
                                        CreateTextFile(pathErr, ErrorMessage);
                                        //goto Loop;
                                    }
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
                                    this.itemPDF = pathFileIO.PathFileRun + "\\" + Path.GetFileName(itemPDF);
                                    WorkProcess(pathFileIO, strFileName, strFileNameEx, this.itemPDF, out iFail, form);
                                    Dispose();
                                }
                                catch (Exception e)
                                {
                                    pathIn = "";
                                    string[] path = pathFileIO.PathInput.Split('\\');
                                    for (int i = 0; i < path.Length; i++)
                                    {
                                        pathIn += path[i] + "/";
                                    }
                                    Console.WriteLine(pathIn);
                                    this.ErrorMessage = "ไม่สามารถย้ายไฟล์" + pathIn + Path.GetFileName(itemPDF).Replace("~$", string.Empty) + "!!!" + strDateTimeStamp;
                                    string itmsg = "ไม่สามารถย้ายไฟล์ ";
                                    _apimail.err_code = "F02";
                                    _apimail.actionmsg = pathIn + Path.GetFileName(itemPDF).Replace("~$", string.Empty);
                                    _apimail.err_msg = pathIn + Path.GetFileName(itemPDF).Replace("~$", string.Empty);
                                    _apimail.input = pathFileIO.PathInput;
                                    _apimail.path = pathFileIO.PathErr;
                                    _apimail.email = form.emailtxt.Text;
                                    _apimail.taxseller = form.txtSellerTaxID.Text;
                                    if (form.pingeng)
                                    {
                                        _apimail.send_err_service();
                                    }
                                    //sock.SendMailAlert(pathFileIO.PathInput, pathFileIO.PathErr, "F02", form.emailtxt.Text, form.txtSellerTaxID.Text, pathIn + Path.GetFileName(itemPDF).Replace("~$", string.Empty), pathIn + Path.GetFileName(itemPDF).Replace("~$", string.Empty));
                                    string pathErr = pathFileIO.PathErr + "\\" + Path.GetFileNameWithoutExtension(itemPDF) + "_" + strDateTimeStamp + "_Error.txt";
                                    //MessageBox.Show(e.Message);
                                    CreateTextFile(pathErr, ErrorMessage);
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
                            catch (Exception e)
                            {
                                pathIn = "";
                                string[] path = pathFileIO.PathInput.Split('\\');
                                for (int i = 0; i < path.Length; i++)
                                {
                                    pathIn += path[i] + "/";
                                }
                                Console.WriteLine(pathIn);
                                this.ErrorMessage = "ไม่สามารถย้ายไฟล์" + pathIn + Path.GetFileName(item).Replace("~$", string.Empty) + "!!!" + strDateTimeStamp;

                                string itmsg = "ไม่สามารถย้ายไฟล์ ";
                                _apimail.err_code = "F01";
                                _apimail.actionmsg = pathIn + Path.GetFileName(item).Replace("~$", string.Empty);
                                _apimail.err_msg = pathIn + Path.GetFileName(item).Replace("~$", string.Empty);
                                _apimail.input = pathFileIO.PathInput;
                                _apimail.path = pathFileIO.PathErr;
                                _apimail.email = form.emailtxt.Text;
                                _apimail.taxseller = form.txtSellerTaxID.Text;
                                if (form.pingeng)
                                {
                                    _apimail.send_err_service();
                                }
                                //sock.SendMailAlert(pathFileIO.PathInput, pathFileIO.PathErr, "F01", form.emailtxt.Text, form.txtSellerTaxID.Text, pathIn + Path.GetFileName(item).Replace("~$", string.Empty), pathIn + Path.GetFileName(item).Replace("~$", string.Empty));
                                string pathErr = pathFileIO.PathErr + "\\" + Path.GetFileNameWithoutExtension(item) + "_" + strDateTimeStamp + "_Error.txt";
                                //MessageBox.Show(e.Message);
                                CreateTextFile(pathErr, ErrorMessage);
                            }
                            iSumFail++;
                            //continue;
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
            
        }
        private static void ConvertAnsiToUTF8(string inputFilePath, string outputFilePath)
        {
            string fileContent = File.ReadAllText(inputFilePath, Encoding.GetEncoding(874));
            File.WriteAllText(outputFilePath, fileContent, Encoding.UTF8);
        }
        public void WorkProcess(PathFilesIO pfIO, string strFileName, string strFileNameExtension, string strFileNamePDF, out int cntFail, etaxOneth form)
        {
            a = null;
            b = null;
            a_with_b = null;
        loop:

            cntFail = 0;
            strOutputFile = new DataOutput();
            string strTaxID;
            string typedoc = "", sellertaxid = "", sellerbranchid = "", document_name = "", document_id = "", document_issue_dtm = "", create_purpose_code = "",
                create_purpose = "", additional_ref_assign_id = "", additional_ref_issue_dtm = "", buyer_name = "", buyer_branch_id = "",
                buyer_tax_id = "", buyer_uriid = "", buyer_address = "", buyer_countrypostcode = "", rangeoflistitem = "", noitem = "", description = "",
                priceunit = "", quanlity = "", amount = "", totalamount = "", vat = "", total = "", document_remark = "",
                discount = "", totaldiscount = "", original_total_amount = "", line_total_amount = "", adjusted_information_amount = "", allowance_total_amount = "",
                tax_basis_total_amount = "", countrybuyer = "", typebuyer = "", buyer_order_assign_id = "", buyer_order_issue_dtm = "", vat_rate = "";

            if (dtParam.ServiceURL == "https://uatetaxsp.one.th/etaxdocumentws/etaxsigndocument")
            {
                UATorPROD = "UAT";
            }
            else
            {
                UATorPROD = "PROD";
            }
            try
            {
                pathText = string.Empty;
                if (Path.GetExtension(dtParam.PathInput).Equals(".xlsx") || Path.GetExtension(dtParam.PathInput).Equals(".xls"))
                {
                    try
                    {
                        using (var reader = new StreamReader(dtParam.PathConfigExcel))
                        {
                            List<string> listA = new List<string>();
                            List<string> listB = new List<string>();
                            while (!reader.EndOfStream)
                            {
                                var line = reader.ReadLine();
                                var values = line.Split(';');
                                listA.Add(values[0]);
                                //listB.Add(values[1]);
                            }
                            foreach (var item in listA)
                            {
                                string[] arrItem = item.Split(',');
                                //Console.WriteLine(arrItem[0]);
                                switch (arrItem[0].ToLower().Trim(' '))
                                {
                                    case "typedoc":
                                        if (!arrItem[2].Equals(""))
                                            typedoc = arrItem[2];
                                        else
                                            typedoc = arrItem[2];
                                        break;
                                    case "discount":
                                        if (!arrItem[2].Equals(""))
                                            discount = arrItem[2] + "," + arrItem[3];
                                        else
                                            discount = arrItem[2];
                                        break;
                                    case "sellertaxid":
                                        if (!arrItem[2].Equals(""))
                                            sellertaxid = arrItem[2] + "," + arrItem[3];
                                        else
                                            sellertaxid = arrItem[2];
                                        break;
                                    case "sellerbranchid":
                                        if (!arrItem[2].Equals(""))
                                            sellerbranchid = arrItem[2] + "," + arrItem[3];
                                        else
                                            sellerbranchid = arrItem[2];
                                        break;
                                    case "document_name":
                                        if (!arrItem[2].Equals(""))
                                            document_name = arrItem[2] + "," + arrItem[3];
                                        else
                                            document_name = arrItem[2];
                                        break;
                                    case "document_id":
                                        if (!arrItem[2].Equals(""))
                                            document_id = arrItem[2] + "," + arrItem[3];
                                        else
                                            document_id = arrItem[2];
                                        break;
                                    case "document_remark":
                                        if (!arrItem[2].Equals(""))
                                            document_remark = arrItem[2] + "," + arrItem[3];
                                        else
                                            document_remark = arrItem[2];
                                        break;
                                    case "document_issue_dtm":
                                        if (!arrItem[2].Equals(""))
                                        {
                                            document_issue_dtm = arrItem[2] + "," + arrItem[3];
                                        }
                                        else
                                            document_issue_dtm = arrItem[2];
                                        break;
                                    case "create_purpose_code":
                                        if (!arrItem[2].Equals(""))
                                            create_purpose_code = arrItem[2] + "," + arrItem[3];
                                        else
                                            create_purpose_code = arrItem[2];
                                        break;
                                    case "create_purpose":
                                        if (!arrItem[2].Equals(""))
                                            create_purpose = arrItem[2] + "," + arrItem[3];
                                        else
                                            create_purpose = arrItem[2];
                                        break;
                                    case "additional_ref_assign_id":
                                        if (!arrItem[2].Equals(""))
                                            additional_ref_assign_id = arrItem[2] + "," + arrItem[3];
                                        else
                                            additional_ref_assign_id = arrItem[2];
                                        break;
                                    case "additional_ref_issue_dtm":
                                        if (!arrItem[2].Equals(""))
                                            additional_ref_issue_dtm = arrItem[2] + "," + arrItem[3];
                                        else
                                            additional_ref_issue_dtm = arrItem[2];
                                        break;
                                    case "buyer_name":
                                        if (!arrItem[2].Equals(""))
                                            buyer_name = arrItem[2] + "," + arrItem[3];
                                        else
                                            buyer_name = arrItem[2];
                                        break;
                                    case "buyer_tax_id":
                                        if (!arrItem[2].Equals(""))
                                            buyer_tax_id = arrItem[2] + "," + arrItem[3];
                                        else
                                            buyer_tax_id = arrItem[2];
                                        break;
                                    case "buyer_branch_id":
                                        if (!arrItem[2].Equals(""))
                                            buyer_branch_id = arrItem[2] + "," + arrItem[3];
                                        else
                                            buyer_branch_id = arrItem[2];
                                        break;
                                    case "buyer_uriid":
                                        if (!arrItem[2].Equals(""))
                                            buyer_uriid = arrItem[2] + "," + arrItem[3];
                                        else
                                            buyer_uriid = arrItem[2];
                                        break;
                                    case "buyer_address":
                                        if (!arrItem[2].Equals(""))
                                            buyer_address = arrItem[2] + "," + arrItem[3];
                                        else
                                            buyer_address = arrItem[2];
                                        break;
                                    case "buyer_country&postcode":
                                        if (!arrItem[2].Equals(""))
                                            buyer_countrypostcode = arrItem[2] + "," + arrItem[3];
                                        else
                                            buyer_countrypostcode = arrItem[2];
                                        break;
                                    case "rangeoflistitem":
                                        rangeoflistitem = arrItem[2];
                                        string[] splitData = arrItem[2].Split('_');
                                        numberRange = splitData[0].Split('-');
                                        charRange = splitData[1].Split('-');

                                        foreach (var k in charRange)
                                        {
                                            AllChar = AllChar + k;
                                        }

                                        break;
                                    case "no.item":
                                        if (!arrItem[2].Equals(""))
                                            noitem = arrItem[2] + "," + arrItem[3];
                                        else
                                            noitem = arrItem[2];
                                        break;
                                    case "description":
                                        if (!arrItem[2].Equals(""))
                                            description = arrItem[2] + "," + arrItem[3];
                                        else
                                            description = arrItem[2];
                                        break;
                                    case "price/unit":
                                        if (!arrItem[2].Equals(""))
                                            priceunit = arrItem[2] + "," + arrItem[3];
                                        else
                                            priceunit = arrItem[2];
                                        break;
                                    case "quanlity":
                                        if (!arrItem[2].Equals(""))
                                            quanlity = arrItem[2] + "," + arrItem[3];
                                        else
                                            quanlity = arrItem[2];
                                        break;
                                    case "amount":
                                        if (!arrItem[2].Equals(""))
                                            amount = arrItem[2] + "," + arrItem[3];
                                        else
                                            amount = arrItem[2];
                                        break;
                                    case "totalamount":
                                        if (!arrItem[2].Equals(""))
                                            totalamount = arrItem[2] + "," + arrItem[3];
                                        else
                                            totalamount = arrItem[2];
                                        break;
                                    case "totaldiscount":
                                        if (!arrItem[2].Equals(""))
                                            totaldiscount = arrItem[2] + "," + arrItem[3];
                                        else
                                            totaldiscount = arrItem[2];
                                        break;
                                    case "vat":
                                        if (!arrItem[2].Equals(""))
                                            vat = arrItem[2] + "," + arrItem[3];
                                        else
                                            vat = arrItem[2];
                                        break;
                                    case "total":
                                        if (!arrItem[2].Equals(""))
                                            total = arrItem[2] + "," + arrItem[3];
                                        else
                                            total = arrItem[2];
                                        break;
                                    case "original_total_amount":
                                        if (!arrItem[2].Equals(""))
                                            original_total_amount = arrItem[2] + "," + arrItem[3];
                                        else
                                            original_total_amount = arrItem[2];
                                        break;
                                    case "line_total_amount":
                                        if (!arrItem[2].Equals(""))
                                            line_total_amount = arrItem[2] + "," + arrItem[3];
                                        else
                                            line_total_amount = arrItem[2];
                                        break;
                                    case "adjusted_information_amount":
                                        if (!arrItem[2].Equals(""))
                                            adjusted_information_amount = arrItem[2] + "," + arrItem[3];
                                        else
                                            adjusted_information_amount = arrItem[2];
                                        break;
                                    case "allowance_total_amount":
                                        if (!arrItem[2].Equals(""))
                                            allowance_total_amount = arrItem[2] + "," + arrItem[3];
                                        else
                                            allowance_total_amount = arrItem[2];
                                        break;
                                    case "tax_basis_total_amount":
                                        if (!arrItem[2].Equals(""))
                                            tax_basis_total_amount = arrItem[2] + "," + arrItem[3];
                                        else
                                            tax_basis_total_amount = arrItem[2];
                                        break;
                                    case "countrybuyer":
                                        if (!arrItem[2].Equals(""))
                                            countrybuyer = arrItem[2] + "," + arrItem[3];
                                        else
                                            countrybuyer = arrItem[2];
                                        break;
                                    case "typebuyer":
                                        if (!arrItem[2].Equals(""))
                                            typebuyer = arrItem[2] + "," + arrItem[3];
                                        else
                                            typebuyer = arrItem[2];
                                        break;
                                    case "buyer_order_assign_id":
                                        if (!arrItem[2].Equals(""))
                                            buyer_order_assign_id = arrItem[2] + "," + arrItem[3];
                                        else
                                            buyer_order_assign_id = arrItem[2];
                                        break;
                                    case "buyer_order_issue_dtm":
                                        if (!arrItem[2].Equals(""))
                                            buyer_order_issue_dtm = arrItem[2] + "," + arrItem[3];
                                        else
                                            buyer_order_issue_dtm = arrItem[2];
                                        break;
                                    case "vat_rate":
                                        if (!arrItem[3].Equals(""))
                                            vat_rate = arrItem[3];
                                        else
                                            vat_rate = arrItem[3];
                                        break;
                                    default:
                                        break;
                                }
                            }
                            //foreach (var item in listA)
                            //{
                            //    Console.WriteLine(item);
                            //}
                            //Console.WriteLine(listA[0]);
                            //Console.WriteLine(listB);
                        } //End of Using for Read ConfigExcel
                        form.pgbLoad.Value = 0;
                        form.OutputPrc(0, "Export Data: 0%");
                        //pgbLoad.Value = 0;
                        //lbPercent.Text = "Export Data: 0%";
                        //lbPercent.Refresh();

                        List<string> lstDataRow = new List<string>();
                        List<string> lstDataMenu = new List<string>();
                        string strSheetName = string.Empty;
                        BGroup grpB = new BGroup();
                        CGroup grpC = new CGroup();
                        LGroup grpL = new LGroup();
                        HGroup grpH = new HGroup();
                        FGroup grpF = new FGroup();
                        Workbook workbook = new Workbook();
                        workbook.LoadFromFile(dtParam.PathInput);
                        sheet = workbook.Worksheets[0];
                        try
                        {
                            //DateTime dateValue;
                            if (DateTime.TryParse(getvalue(document_issue_dtm), out dateValue))
                            {
                                arrDateSplit = getvalue(document_issue_dtm).Split('/');
                                strYear = DateTime.Now.Year.ToString();
                                strYearFront = strYear.Substring(0, 2);
                                DiffOfYears = int.Parse(strYear) - (int.Parse(arrDateSplit[2].Split(' ')[0]) - 543); //ต้องลบ543เพราะ โปรแกรม+543ให้เองอัตโนมัติจึงลบออกเพื่อให้ได้ค่าที่ถูกต้อง
                                if (DiffOfYears < 0)
                                {
                                    DiffOfYears = 543;
                                }
                                else
                                {
                                    DiffOfYears = 0;
                                }
                                if (((int.Parse(arrDateSplit[2].Split(' ')[0]) - 543) - DiffOfYears) < 2000)
                                {
                                    years = (int.Parse(arrDateSplit[2].Split(' ')[0])) - DiffOfYears;
                                }
                                else
                                {
                                    years = (int.Parse(arrDateSplit[2].Split(' ')[0]) - 543) - DiffOfYears;
                                }
                                Console.WriteLine("YearsNow: " + strYear + " Years : " + arrDateSplit[2].Split(' ')[0]);
                                strDocID = arrDateSplit[2].Split(' ')[0] + "-" + Month[arrDateSplit[1].Replace(" ", string.Empty)] + "-" + Day[arrDateSplit[0].Replace(" ", string.Empty)];
                            }
                            else
                            {
                                MessageBox.Show("document_issue_dtm ไม่ถูกต้อง => " + getvalue(document_issue_dtm));
                            }
                        }
                        catch (ArgumentException e)
                        {
                            Console.WriteLine(e.Message);
                        }
                        catch (NullReferenceException e)
                        {
                            Console.WriteLine(e.Message);
                        }
                        catch (IndexOutOfRangeException e)
                        {
                            Console.WriteLine(e.Message);
                        }
                        catch (KeyNotFoundException e)
                        {
                            Console.WriteLine("Date Wrong");
                        }
                        //MessageBox.Show(strDocID);

                        if (form.txtStatus != null && !form.txtStatus.Text.Equals(""))
                        {
                            form.txtStatus.Text += Environment.NewLine + "เลขที่เอกสาร " + getvalue(document_id) + " วันที่ " + strDocID + " ชื่อไฟล์ " + strFileNameExtension + ":";
                            strTempLogTime += Environment.NewLine + "เลขที่เอกสาร " + getvalue(document_id) + " วันที่ " + strDocID + " ชื่อไฟล์ " + strFileNameExtension + ":";
                        }
                        else
                        {
                            form.txtStatus.Text = "เลขที่เอกสาร " + getvalue(document_id) + " วันที่ " + strDocID + " ชื่อไฟล์ " + strFileNameExtension + " :";
                            strTempLogTime = "เลขที่เอกสาร " + getvalue(document_id) + " วันที่ " + strDocID + " ชื่อไฟล์ " + strFileNameExtension + " :";
                        }

                        form.pgbLoad.Value = 10;
                        form.OutputPrc(10, "Export Data: 10%");

                        for (int x = int.Parse(numberRange[0]); x <= int.Parse(numberRange[1]); x++)
                        {
                            for (int k = 0; k < charRange.Length; k++)
                            {
                                string value;
                                if (charRange[k] == noitem.Split(',')[0])
                                {
                                    value = getvalue(charRange[k] + x + "," + noitem.Split(',')[1]);
                                    Console.WriteLine("j =>" + value);
                                    object cellValue = value;
                                    lstDataMenu.Add(cellValue.ToString());
                                }
                                else if (charRange[k] == description.Split(',')[0])
                                {
                                    value = getvalue(charRange[k] + x + "," + description.Split(',')[1]);
                                    Console.WriteLine("k =>" + value);
                                    object cellValue = value;
                                    lstDataMenu.Add(cellValue.ToString());
                                }
                                else if (charRange[k] == priceunit.Split(',')[0])
                                {
                                    value = getvalue(charRange[k] + x + "," + priceunit.Split(',')[1]);
                                    Console.WriteLine("p =>" + value);
                                    if (!value.Equals(""))
                                    {
                                        object cellValue = value;
                                        lstDataMenu.Add(cellValue.ToString());
                                    }
                                    else
                                    {
                                        object cellValue = "";
                                        lstDataMenu.Add(cellValue.ToString());
                                    }
                                }
                                else if (charRange[k] == quanlity.Split(',')[0])
                                {
                                    value = getvalue(charRange[k] + x + "," + quanlity.Split(',')[1]);
                                    if (!value.Equals(""))
                                    {
                                        object cellValue = value;
                                        Console.WriteLine("t =>" + cellValue);
                                        lstDataMenu.Add(cellValue.ToString());
                                    }
                                    else
                                    {
                                        object cellValue = "";
                                        lstDataMenu.Add(cellValue.ToString());
                                    }
                                }
                                else if (charRange[k] == discount.Split(',')[0])
                                {
                                    value = getvalue(charRange[k] + x + "," + discount.Split(',')[1]);
                                    if (!value.Equals(""))
                                    {
                                        object cellValue = value;
                                        lstDataMenu.Add(cellValue.ToString());
                                    }
                                    else
                                    {
                                        object cellValue = "";
                                        lstDataMenu.Add(cellValue.ToString());
                                    }
                                }
                                else if (charRange[k] == amount.Split(',')[0])
                                {
                                    if (!sheet.Range[charRange[k] + x].Value.Equals(""))
                                    {
                                        value = getvalue(charRange[k] + x + "," + amount.Split(',')[1]);
                                        if (!value.Equals(""))
                                        {
                                            object cellValue = value;
                                            try
                                            {
                                                lstDataMenu.Add(Double.Parse(cellValue.ToString()).ToString("0.00"));
                                            }
                                            catch (FormatException e)
                                            {
                                                Console.WriteLine(e);
                                            }
                                        }
                                        else
                                        {
                                            object cellValue = "";
                                            lstDataMenu.Add(cellValue.ToString());
                                        }

                                    }
                                    else
                                    {
                                        object cellValue = "";
                                        lstDataMenu.Add(cellValue.ToString());
                                    }
                                }
                                else
                                {
                                    object cellValue = "";
                                    lstDataMenu.Add(cellValue.ToString());
                                }
                            }
                        }

                        form.pgbLoad.Value = 20;
                        form.OutputPrc(20, "Export Data: 20%");
                        //pgbLoad.Value = 20;
                        //lbPercent.Text = "Export Data: 20%";
                        //lbPercent.Refresh();

                        //Type C
                        grpC.Data_Type = "C";
                        //MessageBox.Show(sheet.Range["F7"].Value.Replace(" ", string.Empty));


                        //Console.WriteLine(RecursionTaxid(" " + sheet.Range[sellertaxid].Value.Replace(" ", string.Empty) + " "));

                        //grpC.Seller_Tax_ID = RecursionTaxid(" " + getvalue(sellertaxid).Replace(" ", string.Empty).Replace("-",string.Empty) + " ").Replace(" ", string.Empty); //เลขประจำตัวผู้เสียภาษี
                        //Console.WriteLine(sheet.Range[sellerbranchid.Split(',')[0]].Value.Replace(" ", string.Empty) + " sellerbranchid");
                        //MessageBox.Show("check");
                        grpC.Seller_Tax_ID = dtParam.SellerTaxID;
                        //grpC.Seller_Branch_ID = sheet.Range["L8"].Value.Replace(" ", string.Empty); //เลขสาขาประกอบการ
                        //grpC.Seller_Branch_ID = dtParam.BranchID;
                        //Console.WriteLine(getvalue(sellerbranchid).Replace(" ", string.Empty));

                        if (!sellerbranchid.Equals(""))
                        {
                            try
                            {
                                if (!getvalue(sellerbranchid).Equals(""))
                                {

                                    grpC.Seller_Branch_ID = branch_seller((sheet.Range[sellerbranchid.Split(',')[0]].Value.Replace(" ", string.Empty)).ToString());
                                }
                                else
                                {
                                    grpC.Seller_Branch_ID = "00000";
                                }
                            }
                            catch (IndexOutOfRangeException e)
                            {
                                Console.WriteLine(e.Message);
                            }
                            catch (NullReferenceException e)
                            {
                                Console.WriteLine(e.Message);
                            }

                        }
                        else
                        {
                            grpC.Seller_Branch_ID = "00000";
                        }

                        Console.WriteLine(grpC.Seller_Branch_ID + "sellerbranchid");

                        //grpC.File_Name = RecursionTaxid(" " + getvalue(sellertaxid).Replace(" ", string.Empty) + " ").Replace(" ", string.Empty) + ".txt"; //ชื่อไฟล์
                        grpC.File_Name = grpC.Seller_Tax_ID + ".txt";
                        form.pgbLoad.Value = 30;
                        form.OutputPrc(30, "Export Data: 30%");
                        grpB.Data_Type = "B";
                        //int iComSplit = getvalue(buyer_name).IndexOf("(");
                        //if (iComSplit != -1)
                        //{
                        //    grpB.Buyer_Name = getvalue(buyer_name).Substring(0, iComSplit - 1); //CompanyName
                        //}
                        //else
                        //{
                        //    grpB.Buyer_Name = getvalue(buyer_name).Replace(" ", string.Empty); //CompanyName
                        //}
                        grpB.Buyer_Name = getvalue(buyer_name);
                        grpB.Buyer_Phone_No = "";
                        try
                        {
                            strTaxID = RecursionTaxid(" " + getvalue(buyer_tax_id).Replace(" ", string.Empty).Replace("-", string.Empty) + " "); //ประเภทผู้เสียภาษี

                            strTaxID = strTaxID.Replace(" ", string.Empty);
                        }
                        catch (NullReferenceException e)
                        {
                            strTaxID = "N/A";
                        }
                        catch (Exception e)
                        {
                            strTaxID = "";
                        }
                        if (!buyer_branch_id.Equals(""))
                        {
                            try
                            {
                                if (!getvalue(buyer_branch_id).Equals(""))
                                {
                                    string buyerbrach_String = branch_buyyer(getvalue(buyer_branch_id).Replace(" ", string.Empty));
                                    grpB.Buyer_Branch_ID = buyerbrach_String;
                                }
                                else
                                {
                                    grpB.Buyer_Branch_ID = "";
                                }
                            }
                            catch (IndexOutOfRangeException e)
                            {
                                grpB.Buyer_Branch_ID = "";
                            }
                            catch (Exception e)
                            {
                                grpB.Buyer_Branch_ID = "";
                            }
                        }
                        else
                        {
                            grpB.Buyer_Branch_ID = "";
                        }

                        try
                        {
                            if (getvalue(buyer_countrypostcode) == "")
                            {
                                grpB.Buyer_Post_Code = ("00000");
                            }
                            else
                            {
                                try
                                {
                                    RealValue = "";
                                    grpB.Buyer_Post_Code = RecursionPostCode(" " + getvalue(buyer_countrypostcode) + " ");
                                }
                                catch (IndexOutOfRangeException e)
                                {
                                    grpB.Buyer_Post_Code = "00000";
                                }
                                catch (Exception e)
                                {
                                    grpB.Buyer_Post_Code = "00000";
                                }
                            }
                        }
                        catch (Exception e)
                        {
                            grpB.Buyer_Post_Code = "";
                        }
                        string keyType = string.Empty;
                        if (strTaxID.Equals("N/A"))
                        {
                            keyType = "4";
                        }
                        else if (strTaxID.Equals(""))
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
                            else if (Char.TryParse(countTaxNum.ToString().Substring(0, 1), out char data) && (grpB.Buyer_Branch_ID == null || grpB.Buyer_Branch_ID.Equals(""))) //อนาคตหากมีเลขที่ PassPort
                            {
                                keyType = "3";
                            }
                            else
                            {
                                keyType = "4";
                                strTaxID = getvalue(buyer_tax_id).Replace(" ", string.Empty).Replace("-", string.Empty);
                            }

                        }
                        grpB.Buyer_Tax_ID_Type = BuyerTaxType[keyType];
                        if (!typebuyer.Equals(""))
                        {
                            if (!getvalue(typebuyer).Equals(""))
                            {
                                grpB.Buyer_Tax_ID_Type = getvalue(typebuyer).Replace(" ", string.Empty).Replace("-", string.Empty);
                                strTaxID = getvalue(buyer_tax_id).Replace(" ", string.Empty).Replace("-", string.Empty);
                            }
                        }
                        if (strTaxID.Equals(""))
                        {
                            grpB.Buyer_Tax_ID = DoubleQuote(" "); //เลขที่ประจำตัวผู้เสียภาษี
                        }
                        else
                        {
                            grpB.Buyer_Tax_ID = strTaxID; //เลขที่ประจำตัวผู้เสียภาษี
                        }

                        if (buyer_uriid.Equals(""))
                        {
                            grpB.Buyer_URIID = "";
                        }
                        else
                        {
                            grpB.Buyer_URIID = getvalue(buyer_uriid).Replace(" ", string.Empty);
                        }
                        grpB.Buyer_Add_Line1 = getvalue(buyer_address);
                        grpB.Buyer_Add_Line2 = "";
                        form.pgbLoad.Value = 40;
                        form.OutputPrc(40, "Export Data: 40%");
                        int iCountRound = 0;

                        List<LGroup> lstGrpL = new List<LGroup>();
                        int distantNameItem = (AllChar.IndexOf(description.Split(',')[0]) - AllChar.IndexOf(noitem.Split(',')[0]));
                        int distantPriceUnit = (AllChar.IndexOf(priceunit.Split(',')[0]) - AllChar.IndexOf(noitem.Split(',')[0]));
                        int distantQuanlity = (AllChar.IndexOf(quanlity.Split(',')[0]) - AllChar.IndexOf(noitem.Split(',')[0]));
                        int distantAmount = (AllChar.IndexOf(amount.Split(',')[0]) - AllChar.IndexOf(noitem.Split(',')[0]));
                        int distantDiscount = (AllChar.IndexOf(discount.Split(',')[0]) - AllChar.IndexOf(noitem.Split(',')[0]));
                        for (int x = 0; x < lstDataMenu.Count; x++)
                        {
                            bool chkSting = false;
                            bool chkNum = false;
                            Double value = 0;
                            string patternChkString = @"([a-zA-Zก-๙0-9])";
                            if (!lstDataMenu[x].Equals(""))
                            {
                                chkSting = Regex.IsMatch(lstDataMenu[x], patternChkString);
                                chkNum = Double.TryParse(lstDataMenu[x], out value);
                            }
                            if (chkSting == true && chkNum == true)
                            {
                                if (iCountRound > 0)
                                {
                                    Console.WriteLine("LengthOfProduct_Desc => " + grpL.Product_Desc.Length);
                                    if (grpL.Product_Desc == null || grpL.Product_Desc.Equals(""))
                                    {
                                        grpL.Product_Desc = "";
                                    }
                                    else
                                    {

                                        if (grpL.Product_Desc.Length > 256)
                                        {
                                            string a = grpL.Product_Desc.Substring(0, 256);
                                            Console.WriteLine("a => " + a);
                                            string[] b = a.Split(' ');
                                            Console.WriteLine(b.Length);
                                            for (int i = 0; i < b.Length - 1; i++)
                                            {
                                                a_with_b += b[i] + " ";
                                            }
                                            Console.WriteLine("a_With_b => " + a_with_b);
                                            Console.WriteLine("LengthOfa_with_b => " + a_with_b.Length);
                                            grpL.Product_Remark = DoubleQuote(grpL.Product_Desc.Substring(a_with_b.Length));
                                            Console.WriteLine("Product_Remark => " + grpL.Product_Remark);
                                            grpL.Product_Desc = a_with_b;
                                            Console.WriteLine("Product_Desc => " + grpL.Product_Desc);
                                        }
                                        grpL.Product_Desc = (grpL.Product_Desc).Replace(",", DoubleQuote(","));
                                    }
                                    lstGrpL.Add(grpL);
                                }

                                grpL = new LGroup();
                                try
                                {
                                    grpL.Data_Type = DoubleQuote("L"); //ประเภทรายการ
                                    grpL.Line_ID = DoubleQuote(lstDataMenu[x]); //ลำดับรายการ
                                    grpL.Product_ID = DoubleQuote(""); //รหัสสินค้า
                                    grpL.Product_Name = lstDataMenu[x + distantNameItem].Replace(" ", string.Empty).Replace(",", DoubleQuote(",")); //ชื่อสินค้า
                                    grpL.Product_Desc = "";
                                    grpL.Product_Batch_ID = DoubleQuote(""); //ครั้งที่ผลิต
                                    grpL.Product_Expire_Dtm = DoubleQuote(""); //วันหมดอายุ
                                    grpL.Product_Class_Code = DoubleQuote(""); //รหัสหมวดหมู่สินค้า
                                    grpL.Product_Class_Name = DoubleQuote(""); //ชื่อหมวดหมู่สินค้า
                                    grpL.Product_OriCountry_ID = DoubleQuote(""); //รหัสประเทศกำเนิด
                                    try
                                    {
                                        grpL.Product_Charge_Amount = DoubleQuote(Double.Parse(RemoveComma(lstDataMenu[x + distantPriceUnit])).ToString("0.00")); //ราคาต่อหน่วย
                                    }
                                    catch (FormatException e)
                                    {
                                        grpL.Product_Charge_Amount = DoubleQuote("");
                                    }
                                    grpL.Product_Charge_Curr_Code = DoubleQuote("THB"); //รหัสสกุลเงิน (ราคาต่อหน่วย)
                                    grpL.Product_Al_Charge_IND = DoubleQuote(""); //ตัวบอกส่วนลดหรือค่าธรรมเนียม
                                    if (!discount.Equals(""))
                                    {
                                        try
                                        {
                                            grpL.Product_Al_Actual_Amount = DoubleQuote(Double.Parse(lstDataMenu[x + distantQuanlity]).ToString("0.00")); //มูลค่าส่วนลดหรือค่าธรรมเนียม
                                        }
                                        catch (FormatException e)
                                        {
                                            grpL.Product_Al_Actual_Amount = DoubleQuote("");
                                        }
                                        grpL.Product_Al_Actual_Curr_Code = DoubleQuote("THB"); //รหัสสกุลเงิน (มูลค่าส่วนลดหรือค่าธรรมเนียม)
                                    }
                                    else
                                    {
                                        grpL.Product_Al_Actual_Amount = DoubleQuote(""); //มูลค่าส่วนลดหรือค่าธรรมเนียม
                                        grpL.Product_Al_Actual_Curr_Code = DoubleQuote(""); //รหัสสกุลเงิน (มูลค่าส่วนลดหรือค่าธรรมเนียม)
                                    }
                                    grpL.Product_Al_Reason_Code = DoubleQuote(""); //รหัสเหตุผลในการคิดส่วนลดหรือค่าธรรมเนียม
                                    grpL.Product_Al_Reason = DoubleQuote(""); //เหตุผลในการคิดสวนลดหรือค่าธรรมเนียม

                                    try
                                    {
                                        grpL.Product_Quantity = DoubleQuote(Double.Parse(lstDataMenu[x + distantQuanlity]).ToString("0.00")); //จำนวนสินค้า
                                    }
                                    catch (FormatException e)
                                    {
                                        grpL.Product_Quantity = DoubleQuote("");
                                    }
                                    grpL.Product_Unit_Code = DoubleQuote(""); //รหัสหน่วยสินค้า
                                    grpL.Product_Quan_Per_Unit = DoubleQuote("1"); //ขนาดบรรจุต่อหน่วยขาย
                                    grpL.Line_Tax_Type_Code = DoubleQuote("VAT"); //รหัสประเภทภาษี
                                    grpL.Line_Tax_Cal_Rate = DoubleQuote("7.00"); //อัตราภาษี
                                                                                  //MessageBox.Show(lstDataMenu[x + 6]);
                                    grpL.Line_Basis_Amount = lstDataMenu[x + distantAmount]; //มูลค่าสินค้า/บริการ (ไม่รวมภาษีมูลค่าเพิ่ม)
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
                                    x = x + (AllChar.Length - 1);
                                }
                                catch (ArgumentOutOfRangeException e)
                                {

                                }
                                catch (NullReferenceException e)
                                {

                                }
                                catch (IndexOutOfRangeException e)
                                {

                                }
                            }
                            else
                            {
                                if (!lstDataMenu[x].Equals(""))
                                {
                                    grpL.Product_Desc += " " + lstDataMenu[x] /*lstDataMenu[x + 1]*/;
                                }
                            }
                        }

                        if (grpL.Product_Desc == null || grpL.Product_Desc.Equals(""))
                        {
                            grpL.Product_Desc = DoubleQuote("");
                        }
                        else
                        {


                            if (grpL.Product_Desc.Length > 256)
                            {
                                string a = grpL.Product_Desc.Substring(0, 256);
                                Console.WriteLine("a => " + a);
                                string[] b = a.Split(' ');
                                Console.WriteLine(b.Length);
                                for (int i = 0; i < b.Length - 1; i++)
                                {
                                    a_with_b += b[i] + " ";
                                }
                                Console.WriteLine("a_With_b => " + a_with_b);
                                Console.WriteLine("LengthOfa_with_b => " + a_with_b.Length);
                                grpL.Product_Remark = DoubleQuote(grpL.Product_Desc.Substring(a_with_b.Length));
                                Console.WriteLine("Product_Remark => " + grpL.Product_Remark);
                                grpL.Product_Desc = a_with_b;
                            }
                            grpL.Product_Desc = (grpL.Product_Desc).Replace(",", DoubleQuote(","));
                        }

                        lstGrpL.Add(grpL);

                        form.pgbLoad.Value = 50;
                        form.OutputPrc(50, "Export Data: 50%");

                        //Type F
                        form.pgbLoad.Value = 60;
                        form.OutputPrc(60, "Export Data: 60%");
                        //Type H
                        string[] arrKey = new string[] { "เลขที่ใบสั่งซื้อ :", "วันที่ใบสั่งซื้อ :" };
                        int[] arrIndex = new int[2];
                        int countArr = 0;
                        try
                        {
                            try
                            {
                                if (!pfIO.TypeDoc.Equals(""))
                                {
                                    grpH.Doc_Type_Code = pfIO.TypeDoc;
                                }
                                else
                                {

                                    if (!typedoc.Equals(""))
                                    {

                                        grpH.Doc_Type_Code = instring(DocType_ENG_AND_CODE, typedoc.Replace(" ", string.Empty));
                                    }
                                    else
                                    {
                                        Console.WriteLine(typedoc + " typedoc");
                                        Console.WriteLine(getvalue(document_name) + " getvalue(document_name)");
                                        grpH.Doc_Type_Code = instring(DocType, getvalue(document_name).Replace(" ", string.Empty));
                                    }
                                }


                            }
                            catch (KeyNotFoundException e)
                            {
                                Console.WriteLine("ไม่พบชื่อตัวแปรที่ส่งมา จึงเกิด error ");
                            }
                            Console.WriteLine(getvalue(document_name).Replace(" ", string.Empty) + " getvalue(document_name).Replace");
                            grpH.Doc_Name = getvalue(document_name).Replace(" ", string.Empty);
                            grpH.Doc_ID = getvalue(document_id).Replace(" ", string.Empty);
                            if (DateTime.TryParse(getvalue(document_issue_dtm), out dateValue))
                            {
                                string[] arrDate = getvalue(document_issue_dtm).Split('/');
                                string year = DateTime.Now.Year.ToString();
                                string yearFront = strYear.Substring(0, 2);
                                grpH.Doc_Issue_Dtm = years + "-" + Month[arrDateSplit[1].Replace(" ", string.Empty)] + "-" + Day[arrDateSplit[0].Replace(" ", string.Empty)] + "T00:00:00";
                            }
                            else
                            {
                                grpH.Doc_Issue_Dtm = getvalue(document_issue_dtm);
                            }

                        }
                        catch (ArgumentException e)
                        {
                            Console.WriteLine(e.Message);
                        }
                        catch (NullReferenceException e)
                        {
                            Console.WriteLine(e.Message);
                        }
                        catch (IndexOutOfRangeException e)
                        {
                            Console.WriteLine(e.Message);
                        }
                        catch (KeyNotFoundException e)
                        {
                            Console.WriteLine(e.Message + "=>" + "KeyNotFoundException");
                        }
                        if (!additional_ref_assign_id.Equals(""))
                        {
                            grpH.Add_Ref_Assign_ID = getvalue(additional_ref_assign_id).Replace(" ", string.Empty);
                        }
                        else
                        {
                            grpH.Add_Ref_Assign_ID = "";

                        }
                        if (!additional_ref_issue_dtm.Equals(""))
                        {
                            try
                            {
                                if (DateTime.TryParse(getvalue(document_issue_dtm), out dateValue))
                                {
                                    arrDateSplit = getvalue(additional_ref_issue_dtm).Split('/');
                                    grpH.Add_Ref_Issue_Dtm = years + "-" + Month[arrDateSplit[1].Replace(" ", string.Empty)] + "-" + Day[arrDateSplit[0].Replace(" ", string.Empty)] + "T00:00:00";
                                }
                                else
                                {
                                    grpH.Add_Ref_Issue_Dtm = getvalue(additional_ref_issue_dtm);
                                }
                            }
                            catch (ArgumentException e)
                            {
                                Console.WriteLine(e.Message);
                            }
                            catch (NullReferenceException e)
                            {
                                Console.WriteLine(e.Message);
                            }
                            catch (IndexOutOfRangeException e)
                            {
                                Console.WriteLine(e.Message);
                            }
                            catch (KeyNotFoundException e)
                            {
                                Console.WriteLine("aasdsasd");
                            }
                        }
                        else
                        {
                            grpH.Add_Ref_Issue_Dtm = "";
                        }

                        //MessageBox.Show(grpH.Add_Ref_Assign_ID);
                        if (!grpH.Add_Ref_Assign_ID.Equals(""))
                        {
                            grpH.Add_Ref_Type_Code = grpH.Doc_Type_Code;
                        }
                        else
                        {
                            grpH.Add_Ref_Type_Code = "";
                        }
                        try
                        {
                            if (create_purpose_code.Equals(""))
                            {
                                switch (grpH.Doc_Type_Code)
                                {
                                    case "388":
                                        try
                                        {
                                            grpH.Create_Purpose_Code = instring(TIVCPurpose, getvalue(create_purpose).Replace(" ", string.Empty));
                                            if (grpH.Create_Purpose_Code.Substring(grpH.Create_Purpose_Code.Length - 2) == "99")
                                            {
                                                grpH.Create_Purpose = getvalue(create_purpose).Replace(" ", string.Empty);
                                            }
                                        }
                                        catch (KeyNotFoundException ex)
                                        {
                                            grpH.Create_Purpose_Code = "TIVC99";
                                            grpH.Create_Purpose = getvalue(create_purpose).Replace(" ", string.Empty);
                                        }
                                        break;
                                    case "T02":
                                        try
                                        {
                                            grpH.Create_Purpose_Code = instring(TIVCPurpose, getvalue(create_purpose).Replace(" ", string.Empty));
                                            if (grpH.Create_Purpose_Code.Substring(grpH.Create_Purpose_Code.Length - 2) == "99")
                                            {
                                                grpH.Create_Purpose = getvalue(create_purpose).Replace(" ", string.Empty);
                                            }
                                        }
                                        catch (KeyNotFoundException ex)
                                        {
                                            grpH.Create_Purpose_Code = "TIVC99";
                                            grpH.Create_Purpose = getvalue(create_purpose).Replace(" ", string.Empty);
                                        }
                                        break;
                                    case "T03":
                                        try
                                        {
                                            grpH.Create_Purpose_Code = instring(TIVCPurpose, getvalue(create_purpose).Replace(" ", string.Empty));
                                            if (grpH.Create_Purpose_Code.Substring(grpH.Create_Purpose_Code.Length - 2) == "99")
                                            {
                                                grpH.Create_Purpose = getvalue(create_purpose).Replace(" ", string.Empty);
                                            }
                                        }
                                        catch (KeyNotFoundException ex)
                                        {
                                            grpH.Create_Purpose_Code = "TIVC99";
                                            grpH.Create_Purpose = getvalue(create_purpose).Replace(" ", string.Empty);
                                        }
                                        break;
                                    case "T04":
                                        try
                                        {
                                            grpH.Create_Purpose_Code = instring(TIVCPurpose, getvalue(create_purpose).Replace(" ", string.Empty));
                                            if (grpH.Create_Purpose_Code.Substring(grpH.Create_Purpose_Code.Length - 2) == "99")
                                            {
                                                grpH.Create_Purpose = getvalue(create_purpose).Replace(" ", string.Empty);
                                            }
                                        }
                                        catch (KeyNotFoundException ex)
                                        {
                                            grpH.Create_Purpose_Code = "TIVC99";
                                            grpH.Create_Purpose = getvalue(create_purpose).Replace(" ", string.Empty);
                                        }
                                        break;
                                    case "T01":
                                        try
                                        {
                                            grpH.Create_Purpose_Code = instring(RCTCPurpose, getvalue(create_purpose).Replace(" ", string.Empty));
                                            if (grpH.Create_Purpose_Code.Substring(grpH.Create_Purpose_Code.Length - 2) == "99")
                                            {
                                                grpH.Create_Purpose = getvalue(create_purpose).Replace(" ", string.Empty);
                                            }

                                        }
                                        catch (KeyNotFoundException ex)
                                        {
                                            grpH.Create_Purpose_Code = "RCTC99";
                                            grpH.Create_Purpose = getvalue(create_purpose).Replace(" ", string.Empty);
                                        }
                                        break;
                                    case "80":
                                        try
                                        {
                                            grpH.Create_Purpose_Code = instring(DBNGPurpose, getvalue(create_purpose).Replace(" ", string.Empty));
                                            if (grpH.Create_Purpose_Code.Substring(grpH.Create_Purpose_Code.Length - 2) == "99")
                                            {
                                                grpH.Create_Purpose = getvalue(create_purpose).Replace(" ", string.Empty);
                                            }
                                        }
                                        catch (KeyNotFoundException ex)
                                        {
                                            grpH.Create_Purpose_Code = "DBNG99";
                                            grpH.Create_Purpose = getvalue(create_purpose).Replace(" ", string.Empty);
                                        }
                                        break;
                                    case "81":
                                        try
                                        {
                                            grpH.Create_Purpose_Code = instring(CDNGPurpose, getvalue(create_purpose).Replace(" ", string.Empty));
                                            if (grpH.Create_Purpose_Code.Substring(grpH.Create_Purpose_Code.Length - 2) == "99")
                                            {
                                                grpH.Create_Purpose = getvalue(create_purpose).Replace(" ", string.Empty);
                                            }
                                        }
                                        catch (KeyNotFoundException ex)
                                        {
                                            grpH.Create_Purpose_Code = "CDNG99";
                                            grpH.Create_Purpose = getvalue(create_purpose).Replace(" ", string.Empty);
                                        }
                                        break;
                                    default:
                                        grpH.Create_Purpose_Code = "";
                                        grpH.Create_Purpose = "";
                                        break;
                                }
                            }
                            else
                            {
                                grpH.Create_Purpose_Code = getvalue(create_purpose_code);
                                grpH.Create_Purpose = getvalue(create_purpose);
                            }
                        }
                        catch (NullReferenceException e)
                        {
                            Console.WriteLine(e.Message);
                        }
                        catch (IndexOutOfRangeException e)
                        {
                            grpH.Create_Purpose_Code = "";
                            grpH.Create_Purpose = "";
                        }
                        catch (KeyNotFoundException e)
                        {
                            grpH.Create_Purpose_Code = "";
                            grpH.Create_Purpose = "";
                        }

                        if (!document_remark.Equals(""))
                        {
                            grpH.DOCUMENT_REMARK = getvalue(document_remark);
                        }
                        else
                        {
                            grpH.DOCUMENT_REMARK = "";
                        }


                        if (!buyer_order_assign_id.Equals(""))
                        {
                            //grpH.Buyer_Order_Assign_ID = getvalue(additional_ref_assign_id).Replace(" ", string.Empty);
                            grpH.Buyer_Order_Assign_ID = getvalue(buyer_order_assign_id).Replace(" ", string.Empty);
                        }
                        else
                        {
                            grpH.Buyer_Order_Assign_ID = "";
                        }


                        if (!buyer_order_issue_dtm.Equals(""))
                        {
                            try
                            {
                                arrDateSplit = getvalue(buyer_order_issue_dtm).Split('/');
                                grpH.Buyer_Order_Issue_Dtm = years + "-" + Month[arrDateSplit[1].Replace(" ", string.Empty)] + "-" + Day[arrDateSplit[0].Replace(" ", string.Empty)] + "T00:00:00";
                            }
                            catch (ArgumentException e)
                            {
                                Console.WriteLine(e.Message);
                            }
                            catch (NullReferenceException e)
                            {
                                Console.WriteLine(e.Message);
                            }
                            catch (IndexOutOfRangeException e)
                            {
                                Console.WriteLine(e.Message);
                            }
                            catch (KeyNotFoundException e)
                            {
                                Console.WriteLine(e.Message);
                            }
                            catch (Exception e)
                            {
                                Console.WriteLine("Exception => " + e.Message);
                            }
                        }
                        else
                        {
                            grpH.Buyer_Order_Issue_Dtm = "";
                        }


                        if (grpH.Buyer_Order_Assign_ID.Equals(""))
                        {
                            grpH.Buyer_Order_Ref_Type_Code = "";
                        }
                        else
                        {
                            grpH.Buyer_Order_Ref_Type_Code = "ON";
                        }

                        try
                        {
                            if (!original_total_amount.Equals(""))
                            {

                                grpF.Original_Total_Amount = Double.Parse(getvalue(original_total_amount)).ToString("0.00");
                                grpF.Original_Total_Curr_Code = "THB";
                            }
                            else
                            {
                                grpF.Original_Total_Amount = "";
                                grpF.Original_Total_Curr_Code = "";
                            }
                        }
                        catch (FormatException e)
                        {
                            Console.WriteLine(e.Message);
                        }
                        catch (NullReferenceException e)
                        {
                            Console.WriteLine(e.Message);
                        }
                        catch (Exception e)
                        {
                            Console.WriteLine("Exception => " + e.Message);
                            
                        }
                        try
                        {
                            if (!line_total_amount.Equals(""))
                            {
                                grpF.LINE_TOTAL_AMOUNT = Double.Parse(getvalue(line_total_amount)).ToString("0.00");
                                grpF.LINE_TOTAL_CURRENCY_CODE = "THB";
                            }
                            else
                            {
                                grpF.LINE_TOTAL_AMOUNT = "";
                                grpF.LINE_TOTAL_CURRENCY_CODE = "";
                            }
                        }
                        catch (FormatException e)
                        {
                            Console.WriteLine(e.Message);
                        }
                        catch (NullReferenceException e)
                        {
                            Console.WriteLine(e.Message);
                        }
                        catch (Exception e)
                        {
                            Console.WriteLine("Exception => " + e.Message);
                            
                        }
                        try
                        {
                            if (!adjusted_information_amount.Equals(""))
                            {
                                grpF.Adjusted_Inform_Amount = Double.Parse(getvalue(adjusted_information_amount)).ToString("0.00");
                                grpF.Adjusted_Inform_Curr_Code = "THB";
                            }
                            else
                            {
                                grpF.Adjusted_Inform_Amount = "";
                                grpF.Adjusted_Inform_Curr_Code = "";
                            }
                        }
                        catch (FormatException e)
                        {
                            Console.WriteLine(e.Message);
                        }
                        catch (NullReferenceException e)
                        {
                            Console.WriteLine(e.Message);
                        }
                        catch (Exception e)
                        {
                            Console.WriteLine("Exception => " + e.Message);
                            
                        }
                        try
                        {
                            if (!allowance_total_amount.Equals(""))
                            {
                                grpF.Al_Total_Amount = Double.Parse(getvalue(allowance_total_amount)).ToString("0.00");
                                grpF.Al_Total_Curr_Code = "THB";
                            }
                            else
                            {
                                grpF.Al_Total_Amount = "";
                                grpF.Al_Total_Curr_Code = "";
                            }
                        }
                        catch (FormatException e)
                        {
                            Console.WriteLine(e.Message);
                        }
                        catch (NullReferenceException e)
                        {
                            Console.WriteLine(e.Message);
                        }
                        catch (Exception e)
                        {
                            Console.WriteLine("Exception => " + e.Message);
                            
                        }
                        if (!countrybuyer.Equals(""))
                        {
                            if (!getvalue(countrybuyer).Equals(""))
                            {
                                grpB.Buyer_Country_ID = getvalue(countrybuyer).Replace(" ", string.Empty).Replace("-", string.Empty);
                                if (grpB.Buyer_Country_ID != "TH")
                                {
                                    grpB.Buyer_Post_Code = "";
                                }
                            }
                            else
                            {
                                grpB.Buyer_Country_ID = "TH";
                            }
                        }
                        else
                        {
                            grpB.Buyer_Country_ID = "TH";
                        }
                        List<string> lstC = new List<string> { DoubleQuote("C"),
                                            DoubleQuote(grpC.Seller_Tax_ID.Replace(" ",string.Empty)), //เลขที่ประจำตัวผู้เสียภาษี
                                            DoubleQuote(grpC.Seller_Branch_ID), //เลขสาขาประกอบการ
                                            DoubleQuote(grpC.File_Name.Replace(" ",string.Empty)), //ชื่อไฟล์  
                                            };
                        //MessageBox.Show("lstC:Success");
                        List<string> lstH = new List<string> { DoubleQuote("H"),
                                            DoubleQuote(grpH.Doc_Type_Code), //ประเภทเอกสาร 
                                            DoubleQuote(grpH.Doc_Name), //ชื่อเอกสาร
                                            DoubleQuote(grpH.Doc_ID), // เลขที่เอกสาร
                                            DoubleQuote(grpH.Doc_Issue_Dtm), //วันที่
                                            DoubleQuote(grpH.Create_Purpose_Code), //สาเหตุการออกเอกสาร
                                            DoubleQuote(grpH.Create_Purpose), //กรณีระบุสาเหตุเอกสาร
                                            DoubleQuote(grpH.Add_Ref_Assign_ID), //เลขที่เอกสารอ้างอิง
                                            DoubleQuote(grpH.Add_Ref_Issue_Dtm), //เอกสารอ้างอิงลงวันที่
                                            DoubleQuote(grpH.Add_Ref_Type_Code), //ประเภทเอกสารอ้างอิง
                                            DoubleQuote(""), //ชื่อเอกสารอ้างอิง 
                                            DoubleQuote(""), //เงื่อนไขการส่งของ
                                            DoubleQuote(grpH.Buyer_Order_Assign_ID), //เลขที่ใบสั่งซื้อ
                                            DoubleQuote(grpH.Buyer_Order_Issue_Dtm), //วันเดือนปีที่ออกใบสั่งซื้อ
                                            DoubleQuote(grpH.Buyer_Order_Ref_Type_Code), //ประเภทเอกสารอ้างอิงการสั่งซื้อ
                                            DoubleQuote(grpH.DOCUMENT_REMARK) //หมายเหตุท้ายเอกสาร
                                            };
                        form.pgbLoad.Value = 70;
                        form.OutputPrc(70, "Export Data: 70%");

                        List<string> lstB = new List<string> { DoubleQuote("B"),
                                            DoubleQuote(""), //รหัสผู้ซื้อ
                                            DoubleQuote(grpB.Buyer_Name), //ชื่อผู้ซื้อ
                                            DoubleQuote(grpB.Buyer_Tax_ID_Type), //ประเภทผู้เสียภาษี
                                            DoubleQuote(grpB.Buyer_Tax_ID.Replace(" ",string.Empty)), //เลขประจำตัวผู้เสียภาษี
                                            DoubleQuote(grpB.Buyer_Branch_ID.Replace(" ",string.Empty)), //เลขที่สาขา
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
                                            DoubleQuote(grpB.Buyer_Country_ID) //รหัสประเทศ
                                            };

                        try
                        {
                            if (!totalamount.Equals(""))
                            {
                                totalamount = RemoveComma(double.Parse(getvalue(totalamount)).ToString("0.00"));
                            }
                            else
                            {
                                totalamount = "";
                            }
                        }
                        catch (IndexOutOfRangeException e)
                        {
                            Console.WriteLine(e.Message);
                        }
                        catch (NullReferenceException e)
                        {
                            Console.WriteLine(e.Message);
                        }
                        catch (Exception e)
                        {
                            Console.WriteLine("Exception => " + e.Message);
                            
                        }
                        try
                        {
                            if (!vat.Equals(""))
                            {
                                vat = RemoveComma(double.Parse(getvalue(vat)).ToString("0.00"));
                            }
                            else
                            {
                                vat = "";
                            }
                        }
                        catch (IndexOutOfRangeException e)
                        {
                            Console.WriteLine(e.Message);
                        }
                        catch (NullReferenceException e)
                        {
                            Console.WriteLine(e.Message);
                        }
                        catch (Exception e)
                        {
                            Console.WriteLine("Exception => " + e.Message);
                            
                        }
                        try
                        {
                            if (!tax_basis_total_amount.Equals(""))
                            {
                                tax_basis_total_amount = RemoveComma(double.Parse(getvalue(tax_basis_total_amount)).ToString("0.00"));
                            }
                            else
                            {
                                tax_basis_total_amount = "";
                            }
                        }
                        catch (IndexOutOfRangeException e)
                        {
                            Console.WriteLine(e.Message);
                        }
                        catch (NullReferenceException e)
                        {
                            Console.WriteLine(e.Message);
                        }
                        catch (Exception e)
                        {
                            Console.WriteLine("Exception => " + e.Message);
                            
                        }
                        try
                        {
                            if (!total.Equals(""))
                            {
                                total = RemoveComma(double.Parse(getvalue(total)).ToString("0.00"));
                            }
                            else
                            {
                                total = "";
                            }
                        }
                        catch (IndexOutOfRangeException e)
                        {
                            Console.WriteLine(e.Message);
                        }
                        catch (NullReferenceException e)
                        {
                            Console.WriteLine(e.Message);
                        }
                        catch (Exception e)
                        {
                            Console.WriteLine("Exception => " + e.Message);
                           
                        }
                        try
                        {
                            if (vat_rate == "" || vat_rate == null)
                            {
                                vat_rate = "7.00";
                            }
                            else
                            {
                                vat_rate = ConvertNumber(vat_rate);
                            }
                        }
                        catch (Exception ex)
                        {
                            vat_rate = "7.00";
                        }
                        List<string> lstF = new List<string> { DoubleQuote("F"),
                                                DoubleQuote(String.Format("{0:00000}", lstGrpL.Count).ToString()), //จำนวนรายการสินค้า
                                                DoubleQuote(""), //วันเวลานัดส่งสินค้า
                                                DoubleQuote("THB"), //รหัสสกุลเงินตรา
                                                DoubleQuote("VAT"), //รหัสประเภทภาษี
                                                DoubleQuote(vat_rate), //อัตราภาษี
                                                //DoubleQuote(RemoveComma(sumAmount.ToString("N2"))), //มูลค่าสินค้า(ไม่รวมภาษีมูลค่าเพิ่ม)2350
                                                DoubleQuote(totalamount),
                                                DoubleQuote("THB"),
                                                //DoubleQuote(RemoveComma(sumTaxAmount.ToString("N2"))), //มูลค่าภาษีมูลค่าเพิ่ม
                                                DoubleQuote(vat),
                                                DoubleQuote("THB"),
                                                DoubleQuote(""), //ตัวบอกส่วนลดหรือค่าธรรมเนียม
                                                DoubleQuote(""), //มูลค่าส่วนลดหรือค่าธรรมเนียม
                                                DoubleQuote(""),
                                                DoubleQuote(""), //รหัสเหตุผลในการคิดส่วนลดหรือค่าธรรมเนียม
                                                DoubleQuote(""), //เหตุผลในการคิดส่วนลดหรือค่าธรรมเนียม
                                                DoubleQuote(""), //รหัสประเภทส่วนลด     
                                                DoubleQuote(""), //รายละเอียดเงื่อนไขการชำระเงิน
                                                DoubleQuote(""), //วันครบกำหนดชำระเงิน
                                                DoubleQuote(grpF.Original_Total_Amount), //รวมมูลค่าตามเอกสารเดิม
                                                DoubleQuote(grpF.Original_Total_Curr_Code),
                                                //DoubleQuote(RemoveComma(sumAmount.ToString("N2"))),
                                                DoubleQuote(totalamount),
                                                DoubleQuote("THB"),
                                                DoubleQuote(grpF.Adjusted_Inform_Amount), //มูลค่าผลต่าง
                                                DoubleQuote(grpF.Adjusted_Inform_Curr_Code),
                                                DoubleQuote(grpF.Al_Total_Amount), //ส่วนลดทั้งหมด
                                                DoubleQuote(grpF.Al_Total_Curr_Code),
                                                DoubleQuote(""), //ค่าธรรมเนียมทั้งหมด
                                                DoubleQuote(""),
                                                //DoubleQuote(RemoveComma(sumAmount.ToString("N2"))), //มูลค่าที่นำมาคิดภาษีมูลค่าเพิ่ม
                                                DoubleQuote(tax_basis_total_amount),
                                                DoubleQuote("THB"),
                                                //DoubleQuote(RemoveComma(sumTaxAmount.ToString("N2"))), //จำนวนภาษีมูลค่าเพิ่ม
                                                DoubleQuote(vat),
                                                DoubleQuote("THB"),
                                                //DoubleQuote(RemoveComma(sumGrandTotal.ToString("N2"))), //จำนวนเงินรวม(รวมภาษีมูลค่าเพิ่ม)
                                                DoubleQuote(total),
                                                DoubleQuote("THB")
                                                };

                        List<string> lstT = new List<string> { DoubleQuote("T"),
                                                DoubleQuote("1") //จำนวนเอกสารทั้งหมด
                                                };
                        form.pgbLoad.Value = 80;
                        form.OutputPrc(80, "Export Data: 80%");
                        Console.WriteLine("a");
                        string messageText = String.Join(",", lstC) + "\r"
                                + String.Join(",", lstH) + "\r"
                                + String.Join(",", lstB) + "\r";
                        Console.WriteLine("b");
                        for (int k = 0; k < lstGrpL.Count; k++)
                        {
                            messageText += lstGrpL[k].Data_Type + "," + lstGrpL[k].Line_ID + "," + lstGrpL[k].Product_ID + "," + lstGrpL[k].Product_Name + "," + lstGrpL[k].Product_Desc + ","
                                + lstGrpL[k].Product_Batch_ID + "," + lstGrpL[k].Product_Expire_Dtm + "," + lstGrpL[k].Product_Class_Code + "," + lstGrpL[k].Product_Class_Name + "," + lstGrpL[k].Product_OriCountry_ID + ","
                                + lstGrpL[k].Product_Charge_Amount + "," + lstGrpL[k].Product_Charge_Curr_Code + "," + lstGrpL[k].Product_Al_Charge_IND + "," + lstGrpL[k].Product_Al_Actual_Amount + "," + lstGrpL[k].Product_Al_Actual_Curr_Code + ","
                                + lstGrpL[k].Product_Al_Reason_Code + "," + lstGrpL[k].Product_Al_Reason + "," + lstGrpL[k].Product_Quantity + "," + lstGrpL[k].Product_Unit_Code + "," + lstGrpL[k].Product_Quan_Per_Unit + ","
                                + lstGrpL[k].Line_Tax_Type_Code + "," + lstGrpL[k].Line_Tax_Cal_Rate + "," + lstGrpL[k].Line_Basis_Amount + "," + lstGrpL[k].Line_Basis_Curr_Code + "," + lstGrpL[k].Line_Tax_Cal_Amount + ","
                                + lstGrpL[k].Line_Tax_Cal_Curr_Code + "," + lstGrpL[k].Line_AL_Charge_IND + "," + lstGrpL[k].Line_AL_Actual_Amount + "," + lstGrpL[k].Line_AL_Actual_Curr_Code + "," + lstGrpL[k].Line_AL_Reason_Code + ","
                                + lstGrpL[k].Line_AL_Reason + "," + lstGrpL[k].Line_Tax_Total_Amount + "," + lstGrpL[k].Line_Tax_Total_Curr_Code + "," + lstGrpL[k].Line_Net_Total_Amount + "," + lstGrpL[k].Line_Net_Total_Curr_Code + ","
                                + lstGrpL[k].Line_Net_Include_Amount + "," + lstGrpL[k].Line_Net_Include_Curr_Code + "," + lstGrpL[k].Product_Remark + "\r";
                        }
                        messageText += String.Join(",", lstF) + "\r"
                                + String.Join(",", lstT);
                        pathText = dtParam.PathOutput + "\\" + UATorPROD + "_" + strFileName + "_" + strDateTimeStamp + ".txt";
                        CreateTextFile(pathText, messageText);
                        form.txtStatus.Text += Environment.NewLine + "   -Convert Success!";
                        strTempLogTime += " Convert Success!";
                        //txtStatus.Refresh();
                        form.Outputmessage(txtstr);
                        System.Threading.Thread.Sleep(500);
                        if (dtParam.ServiceCode == "S06")
                        {
                            strOutputFile = conAPIClass.CallAPI(dtParam, pathText, strFileNamePDF);
                        }
                        else if (dtParam.ServiceCode == "S03")
                        {
                            strOutputFile = conAPIClass.CallAPI(dtParam, pathText, "");
                        }


                    }
                    catch (FileNotFoundException e)
                    {
                        MessageBox.Show("File Not Found ConfigExcel => " + dtParam.PathConfigExcel);
                        goto loop;
                    }
                    catch (IOException e)
                    {
                        MessageBox.Show("กรุณาปิดไฟล์ Excel ทั้งหมด");
                        goto loop;
                    }
                    catch (Exception e)
                    {
                        MessageBox.Show("Exception => " + e.Message);
                        goto loop;
                    }
                }
                else if (Path.GetExtension(dtParam.PathInput).Equals(".csv"))
                {
                    form.pgbLoad.Value = 0;
                    form.OutputPrc(0, "Export Data: 0%");

                    using (var reader = new StreamReader(dtParam.PathInput))
                    {
                        List<string> listA = new List<string>();
                        List<string> listB = new List<string>();
                        while (!reader.EndOfStream)
                        {
                            var line = reader.ReadLine();
                            var values = line.Split(';');
                            listA.Add(values[0]);
                            //listB.Add(values[1]);
                        }
                        string[] arrSplit = listA[1].Split(',');
                        string strID = arrSplit[3].Replace("\"", string.Empty);
                        string[] arrDate = Regex.Split(arrSplit[4], "T");
                        string strDate = arrDate[0].Replace("\"", string.Empty);
                        //MessageBox.Show(strDate);
                        if (form.txtStatus.Text != null && !form.txtStatus.Text.Equals(""))
                        {
                            form.txtStatus.Text += Environment.NewLine + "เลขที่เอกสาร " + strID + " วันที่ " + strDate + " ชื่อไฟล์ " + strFileNameExtension + ":";
                            strTempLogTime += Environment.NewLine + "เลขที่เอกสาร " + strID + " วันที่ " + strDate + " ชื่อไฟล์ " + strFileNameExtension + ":";
                        }
                        else
                        {
                            form.txtStatus.Text = "เลขที่เอกสาร " + strID + " วันที่ " + strDate + " ชื่อไฟล์ " + strFileNameExtension + " :";
                            strTempLogTime = "เลขที่เอกสาร " + strID + " วันที่ " + strDate + " ชื่อไฟล์ " + strFileNameExtension + " :";
                        }
                    }
                    System.Threading.Thread.Sleep(500);
                    form.pgbLoad.Value = 50;
                    form.OutputPrc(50, "Export Data: 50%");
                    //MessageBox.Show(dtParam.PathInput);
                    System.Threading.Thread.Sleep(500);

                    strOutputFile = conAPIClass.CallAPI(dtParam, dtParam.PathInput, strFileNamePDF);
                    pathText = UATorPROD + "_" + Path.GetFileNameWithoutExtension(dtParam.PathInput) + "_" + strDateTimeStamp + ".txt";
                }
                else
                {
                    try
                    {
                        form.pgbLoad.Value = 0;
                        form.OutputPrc(0, "Export Data: 0%");

                        string strText = System.IO.File.ReadAllText(dtParam.PathInput);
                        string[] arrSplit = strText.Split(',');
                        string strID = arrSplit[6].Replace("\"", string.Empty);
                        string[] arrDate = Regex.Split(arrSplit[7], "T");
                        string strDate = arrDate[0].Replace("\"", string.Empty);
                        //MessageBox.Show(strDate);
                        if (form.txtStatus.Text != null && !form.txtStatus.Text.Equals(""))
                        {
                            form.txtStatus.Text += Environment.NewLine + "เลขที่เอกสาร " + strID + " วันที่ " + strDate + " ชื่อไฟล์ " + strFileNameExtension + ":";
                            strTempLogTime += Environment.NewLine + "เลขที่เอกสาร " + strID + " วันที่ " + strDate + " ชื่อไฟล์ " + strFileNameExtension + ":";
                        }
                        else
                        {
                            form.txtStatus.Text = "เลขที่เอกสาร " + strID + " วันที่ " + strDate + " ชื่อไฟล์ " + strFileNameExtension + " :";
                            strTempLogTime = "เลขที่เอกสาร " + strID + " วันที่ " + strDate + " ชื่อไฟล์ " + strFileNameExtension + " :";
                        }
                        form.pgbLoad.Value = 50;
                        form.OutputPrc(50, "Export Data: 50%");

                        //System.Threading.Thread.Sleep(1);
                        strOutputFile = conAPIClass.CallAPI(dtParam, dtParam.PathInput, strFileNamePDF);
                        pathText = UATorPROD + "_" + Path.GetFileNameWithoutExtension(dtParam.PathInput) + "_" + strDateTimeStamp + ".txt";
                        Console.WriteLine(strOutputFile.Message_Content + " Console.WriteLine(strOutputFile.MessageError)");
                    }
                    catch (Exception e)
                    {
                        MessageBox.Show("Exception => " + e.Message);
                        goto loop;
                    }

                }
                form.pgbLoad.Value = 90;
                form.OutputPrc(90, "Export Data: 90%");
                FileVersionInfo fvi = FileVersionInfo.GetVersionInfo(assembly.Location);
                string version = fvi.FileVersion;
                if (strOutputFile.MessageResultError != null && !strOutputFile.MessageResultError.Equals(""))
                //if (strOutputFile.MessageResultError == "")
                {
                    JObject oKeepResponeExecute = new JObject();
                    //MessageBox.Show(strOutputFile.MessageResultError);

                    if (strOutputFile.MessageResultError == "{}")
                    {
                        strOutputFile.MessageResultError = "กรุณาตรวจสอบอินเตอร์เน็ต!!";
                    }
                    form.txtStatus.Text += Environment.NewLine + "   -**********etax.one.th Fail!" + " (" + strOutputFile.MessageLogTime + ")" + "**********";
                    strTempLogTime += " etax.one.th Fail!" + " (" + strOutputFile.MessageLogTime + ") Version " + version;
                    //Console.WriteLine(pathIn + " " + pathOutput);
                    oKeepResponeExecute = JObject.Parse(strOutputFile.MessageResultError.ToString());

                    //sock.SendMailAlert(dtParam.PathInput, dtParam.PathOutput, "FE99", form.emailtxt.Text, form.txtSellerTaxID.Text, Path.GetFileName(dtParam.PathInput).Replace("~$", string.Empty) + "-//-" + oKeepResponeExecute["errorCode"], oKeepResponeExecute["errorMessage"].ToString().Replace(" ",string.Empty));
                    //sock.SendMailAlert(pathFileIO.PathInput, pathFileIO.PathErr, "F01", form.emailtxt.Text, form.txtSellerTaxID.Text, pathIn + Path.GetFileName(item).Replace("~$", string.Empty), pathIn + Path.GetFileName(item).Replace("~$", string.Empty));
                    cntFail++;
                }
                else
                {

                    form.txtStatus.Text += Environment.NewLine + "   -etax.one.th Success!" + " (" + strOutputFile.MessageLogTime + ")";
                    strTempLogTime += ", etax.one.th Success!" + " (" + strOutputFile.MessageLogTime + ") Version " + version;
                }


                //txtStatus.Refresh();
                if (chkOption == true)
                {
                    if (strOutputFile.StatusCallAPI == false)
                    {

                        string pathErr = dtParam.PathOutput + "\\" + Path.GetFileNameWithoutExtension(pathText) + "_Error.txt";
                        JObject oKeepResponeExecute = new JObject();
                        oKeepResponeExecute = JObject.Parse(strOutputFile.MessageResultError.ToString());
                        Console.WriteLine(strOutputFile + " oKeepResponeExecute");
                        _apimail.err_code = "FE99";
                        _apimail.actionmsg = oKeepResponeExecute["errorMessage"].ToString().Replace(" ", string.Empty).Replace("\n", string.Empty).Replace(",", string.Empty).Replace("'", string.Empty);
                        _apimail.err_msg = Path.GetFileName(dtParam.PathInput).Replace("~$", string.Empty) + "-//-" + oKeepResponeExecute["errorCode"];
                        _apimail.input = dtParam.PathInput;
                        _apimail.path = dtParam.PathOutput;
                        _apimail.email = form.emailtxt.Text;
                        _apimail.taxseller = form.txtSellerTaxID.Text;
                        if (form.pingeng && oKeepResponeExecute["errorCode"].ToString() != "ER011")
                        {
                            _apimail.send_err_service();

                        }
                        if (oKeepResponeExecute["errorCode"].ToString() == "ER011")
                        {

                        }
                        else
                        {
                            Console.WriteLine(dtParam.PathInput);
                            CreateTextFile(pathErr, strOutputFile.MessageResultError);
                        }




                        //if (form.pingeng == true)
                        //{
                        //    sock.SendMailAlert(dtParam.PathInput, dtParam.PathOutput, "FE99", form.emailtxt.Text, form.txtSellerTaxID.Text, Path.GetFileName(dtParam.PathInput).Replace("~$", string.Empty) + "-//-" + oKeepResponeExecute["errorCode"], oKeepResponeExecute["errorMessage"].ToString());
                        //}


                    }
                    else
                    {
                        this.pathOutput = Path.GetFileNameWithoutExtension(pathText);
                        try
                        {
                            if (!Directory.Exists(dtParam.PathOutput + "\\" + "LogSucces"))
                            {
                                Directory.CreateDirectory(dtParam.PathOutput + "\\" + "LogSucces");
                            }
                        }
                        catch (Exception ea)
                        {
                            Console.WriteLine(ea);
                        }
                        finally
                        {
                            for (int i = 0; i <= GC.MaxGeneration; i++)
                            {
                                int count = GC.CollectionCount(i);
                                GC.Collect();
                            }
                            GC.WaitForPendingFinalizers();
                            GC.SuppressFinalize(this);
                        }
                        try
                        {
                            if (!Directory.Exists(dtParam.PathOutput + "\\" + "Log"))
                            {
                                Directory.CreateDirectory(dtParam.PathOutput + "\\" + "Log");
                            }
                        }
                        catch (Exception ea)
                        {
                            Console.WriteLine(ea);
                        }
                        finally
                        {
                            for (int i = 0; i <= GC.MaxGeneration; i++)
                            {
                                int count = GC.CollectionCount(i);
                                GC.Collect();
                            }
                            GC.WaitForPendingFinalizers();
                            GC.SuppressFinalize(this);
                        }
                        try
                        {
                            if (!Directory.Exists(dtParam.PathOutput + "\\" + "Temp_Succes"))
                            {
                                Directory.CreateDirectory(dtParam.PathOutput + "\\" + "Temp_Succes");
                            }
                        }
                        catch (Exception ea)
                        {
                            Console.WriteLine(ea);
                        }
                        finally
                        {
                            for (int i = 0; i <= GC.MaxGeneration; i++)
                            {
                                int count = GC.CollectionCount(i);
                                GC.Collect();
                            }
                            GC.WaitForPendingFinalizers();
                            GC.SuppressFinalize(this);
                        }
                        try
                        {
                            if (!Directory.Exists(dtParam.PathOutput + "\\" + "Log_Resend"))
                            {
                                Directory.CreateDirectory(dtParam.PathOutput + "\\" + "Log_Resend");
                            }
                        }
                        catch (Exception ea)
                        {
                            Console.WriteLine(ea);
                        }
                        finally
                        {
                            for (int i = 0; i <= GC.MaxGeneration; i++)
                            {
                                int count = GC.CollectionCount(i);
                                GC.Collect();
                            }
                            GC.WaitForPendingFinalizers();
                            GC.SuppressFinalize(this);
                        }

                        _apimail.err_code = "FE91";
                        _apimail.err_msg = strFileNameExtension.Replace("~$", string.Empty) + "-//-" + "";
                        _apimail.input = dtParam.PathInput;
                        _apimail.path = dtParam.PathOutput;
                        _apimail.email = form.emailtxt.Text;
                        _apimail.taxseller = form.txtSellerTaxID.Text;
                        int counttimes = 0;
                        bool string_check__pdf;
                        bool string_check__xml;
                        JObject json_respo = new JObject();
                        json_respo = JObject.Parse(strOutputFile.Message_Content);
                        Console.WriteLine(json_respo + " json_respo");
                        if (json_respo["status"].ToString() != "ER")
                        {
                            string Temp_succes = dtParam.PathOutput + "\\" + "Temp_Succes";
                            CreateTextFile(dtParam.PathOutput + "\\LogSucces\\" + "Success_" + Path.GetFileNameWithoutExtension(pathText) + ".txt", strOutputFile.Message_Content);
                        download_pdfandxml:
                            counttimes = counttimes + 1;
                            DownloadFile(strOutputFile.MessageResultPDF, Temp_succes, this.pathOutput + "_PDF.pdf");
                            DownloadFile(strOutputFile.MessageResultXML, Temp_succes, this.pathOutput + "_XML.xml");
                            string_check__pdf = _checkfolder_pdf(this.pathOutput + "_PDF.pdf", Temp_succes);
                            string_check__xml = _checkfolder_xml(this.pathOutput + "_XML.xml", Temp_succes);
                            if (string_check__pdf == false && string_check__xml == false && counttimes <= 3)
                            {
                                goto download_pdfandxml;
                            }
                            else if (string_check__pdf == false && counttimes <= 3)
                            {
                                goto download_pdfandxml;
                            }
                            else if (string_check__xml == false && counttimes <= 3)
                            {
                                goto download_pdfandxml;
                            }

                            if (string_check__pdf == false && string_check__xml == false)
                            {

                                _apimail.actionmsg = "ไม่สามารถดาวน์โหลดไฟล์ PDF และ XML ได้";
                                if (form.pingeng)
                                {
                                    _apimail.send_err_service();
                                }
                                CreateTextFile(dtParam.PathOutput + "\\Log\\" + "Log_" + Path.GetFileNameWithoutExtension(pathText) + ".txt", "ไม่สามารถดาวน์โหลดไฟล์ PDF และ XML ได้");
                            }
                            else if (string_check__pdf == false)
                            {
                                _apimail.actionmsg = "ไม่สามารถดาวน์โหลดไฟล์ PDF ได้";
                                if (form.pingeng)
                                {
                                    _apimail.send_err_service();
                                }
                                CreateTextFile(dtParam.PathOutput + "\\Log\\" + "Log_" + Path.GetFileNameWithoutExtension(pathText) + ".txt", "ไม่สามารถดาวน์โหลดไฟล์ PDF ได้");
                            }
                            else if (string_check__xml == false)
                            {
                                _apimail.actionmsg = "ไม่สามารถดาวน์โหลดไฟล์ XML ได้";
                                if (form.pingeng)
                                {
                                    _apimail.send_err_service();
                                }
                                CreateTextFile(dtParam.PathOutput + "\\Log\\" + "Log_" + Path.GetFileNameWithoutExtension(pathText) + ".txt", "ไม่สามารถดาวน์โหลดไฟล์ XML ได้");
                            }
                            try
                            {
                                var move_file_dis_form_pdf = Temp_succes + "\\" + this.pathOutput + "_PDF.pdf";
                                var move_file_dis_to_pdf = dtParam.PathOutput + "\\" + this.pathOutput + "_PDF.pdf";
                                var move_file_dis_form_xml = Temp_succes + "\\" + this.pathOutput + "_XML.xml";
                                var move_file_dis_to_xml = dtParam.PathOutput + "\\" + this.pathOutput + "_XML.xml";
                                if (string_check__pdf == true && string_check__xml == true)
                                {
                                    File.Move(move_file_dis_form_pdf, move_file_dis_to_pdf);
                                    File.Move(move_file_dis_form_xml, move_file_dis_to_xml);
                                }
                                else if (string_check__pdf == true && string_check__xml == false)
                                {
                                    File.Move(move_file_dis_form_pdf, move_file_dis_to_pdf);
                                }
                                else if (string_check__pdf == false && string_check__xml == true)
                                {
                                    File.Move(move_file_dis_form_xml, move_file_dis_to_xml);
                                }
                            }
                            catch (Exception ex)
                            {
                                Console.WriteLine(ex.Message);
                            }

                        }
                        else if (json_respo["errorCode"].ToString() == "ER011")
                        {
                            string Temp_succes = dtParam.PathOutput + "\\" + "Temp_Succes";
                            CreateTextFile(dtParam.PathOutput + "\\" + "Log_Resend\\" + "Resend_Success_" + Path.GetFileNameWithoutExtension(pathText) + ".txt", strOutputFile.Message_Content);
                            //MessageBox.Show(json_respo.ToString());
                            int counttimes_resend = 0;
                            string pathText_resend = "RESEND_" + UATorPROD + "_" + Path.GetFileNameWithoutExtension(dtParam.PathInput) + "_" + strDateTimeStamp + ".txt";
                            string pathOutput_resend = Path.GetFileNameWithoutExtension(pathText_resend);
                        download_pdf_resend:
                            counttimes_resend = counttimes_resend + 1;
                            DownloadFile(strOutputFile.MessageResultPDF, Temp_succes, pathOutput_resend + "_PDF.pdf");
                            DownloadFile(strOutputFile.MessageResultXML, Temp_succes, pathOutput_resend + "_XML.xml");
                            bool string_check_pdf_resend = _checkfolder_pdf(pathOutput_resend + "_PDF.pdf", Temp_succes);
                            bool string_check_xml_resend = _checkfolder_xml(pathOutput_resend + "_XML.xml", Temp_succes);
                            if (string_check_pdf_resend == false && string_check_xml_resend == false && counttimes_resend <= 3)
                            {
                                goto download_pdf_resend;
                            }
                            else if (string_check_pdf_resend == false && counttimes_resend <= 3)
                            {
                                goto download_pdf_resend;
                            }
                            else if (string_check_xml_resend == false && counttimes_resend <= 3)
                            {
                                goto download_pdf_resend;
                            }
                            if (string_check_pdf_resend == false && string_check_xml_resend == false)
                            {
                                CreateTextFile(dtParam.PathOutput + "\\Log\\" + "Log_" + Path.GetFileNameWithoutExtension(pathText) + ".txt", "ไม่สามารถดาวน์โหลดไฟล์ PDF และ XML ได้");
                            }
                            else if (string_check_pdf_resend == false)
                            {
                                CreateTextFile(dtParam.PathOutput + "\\Log\\" + "Log_" + Path.GetFileNameWithoutExtension(pathText) + ".txt", "ไม่สามารถดาวน์โหลดไฟล์ PDF ได้");
                            }
                            else if (string_check_xml_resend == false)
                            {
                                CreateTextFile(dtParam.PathOutput + "\\Log\\" + "Log_" + Path.GetFileNameWithoutExtension(pathText) + ".txt", "ไม่สามารถดาวน์โหลดไฟล์ XML ได้");
                            }

                            try
                            {
                                var move_file_dis_form_pdf = Temp_succes + "\\" + pathOutput_resend + "_PDF.pdf";
                                var move_file_dis_to_pdf = dtParam.PathOutput + "\\" + pathOutput_resend + "_PDF.pdf";
                                var move_file_dis_form_xml = Temp_succes + "\\" + pathOutput_resend + "_XML.xml";
                                var move_file_dis_to_xml = dtParam.PathOutput + "\\" + pathOutput_resend + "_XML.xml";
                                if (string_check_pdf_resend == true && string_check_xml_resend == true)
                                {
                                    File.Move(move_file_dis_form_pdf, move_file_dis_to_pdf);
                                    File.Move(move_file_dis_form_xml, move_file_dis_to_xml);
                                }
                                else if (string_check_pdf_resend == true && string_check_xml_resend == false)
                                {
                                    File.Move(move_file_dis_form_pdf, move_file_dis_to_pdf);
                                }
                                else if (string_check_pdf_resend == false && string_check_xml_resend == true)
                                {
                                    File.Move(move_file_dis_form_xml, move_file_dis_to_xml);
                                }
                            }
                            catch (Exception ex)
                            {
                                Console.WriteLine(ex.Message);
                            }

                        }

                        //CreateTextFile(dtParam.PathOutput, strOutputFile.Message_Content);
                        Console.WriteLine(dtParam.PathOutput + " 3872");
                    }
                }
                else if (chkOption == false)
                {
                    if (strOutputFile.StatusCallAPI == false)
                    {
                        string pathErr = pfIO.PathErr + "\\" + Path.GetFileNameWithoutExtension(pathText) + "_Error.txt";
                        CreateTextFile(pathErr, strOutputFile.MessageResultError);
                        JObject oKeepResponeExecute1 = new JObject();
                        oKeepResponeExecute1 = JObject.Parse(strOutputFile.MessageResultError.ToString());
                        Console.WriteLine(strOutputFile + " oKeepResponeExecute1");
                        Console.WriteLine(Path.GetFileName(dtParam.PathInput).Replace("~$", string.Empty) + "-//-" + oKeepResponeExecute1["errorCode"]);
                        try
                        {
                            _apimail.err_code = "FE99";
                            _apimail.actionmsg = oKeepResponeExecute1["errorMessage"].ToString().Replace(" ", string.Empty).Replace("\n", string.Empty).Replace(",", string.Empty).Replace("'", string.Empty);
                            _apimail.err_msg = Path.GetFileName(dtParam.PathInput).Replace("~$", string.Empty) + "-//-" + oKeepResponeExecute1["errorCode"];
                            _apimail.input = dtParam.PathInput;
                            _apimail.path = dtParam.PathOutput;
                            _apimail.email = form.emailtxt.Text;
                            _apimail.taxseller = form.txtSellerTaxID.Text;

                            if (form.pingeng)
                            {
                                _apimail.send_err_service();
                            }
                            Console.WriteLine(dtParam.PathInput);
                            CreateTextFile(pathErr, strOutputFile.MessageResultError);

                            //if (form.metroToggle1.Checked == true && form.pingeng == true)
                            //{
                            //    sock.SendMailAlert(dtParam.PathInput, dtParam.PathOutput, "FE99", form.emailtxt.Text, form.txtSellerTaxID.Text, Path.GetFileName(dtParam.PathInput).Replace("~$", string.Empty) + "-//-" + oKeepResponeExecute1["errorCode"], oKeepResponeExecute1["errorMessage"].ToString().Replace(" ", string.Empty).Replace("\n", string.Empty).Replace(",", string.Empty).Replace("'", string.Empty));
                            //}
                        }
                        catch (Exception ea)
                        {
                            Console.WriteLine(ea);
                        }

                        string[] arrFiles = System.IO.Directory.GetFiles(pfIO.PathTemp, "*.txt");
                        string[] arrFilesSource = System.IO.Directory.GetFiles(pfIO.PathSource_F, "*.txt");

                        foreach (var item in arrFiles)
                        {
                            string fileName = Path.GetFileName(item);
                            this.nameFilePDF = item;
                            string pathTxtNew = pfIO.PathSource_F + "\\" + fileName;
                            string pathTxtNew_S = pfIO.PathSource_S + "\\" + fileName;
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
                            File.Copy(item, pathTxtNew_S);
                            File.Move(item, pathTxtNew);
                        }
                    }
                    else
                    {
                        string fileNameWithoutExtension = string.Empty;
                        string[] arrFiles = System.IO.Directory.GetFiles(pfIO.PathTemp, "*.txt");
                        string[] arrFiles__pcfg = System.IO.Directory.GetFiles(pfIO.PathInput, "*.pcfg");

                        this.pathOutput = Path.GetFileNameWithoutExtension(pathText);
                        try
                        {
                            if (!Directory.Exists(pfIO.PathSuccess_O + "\\" + "LogSucces"))
                            {
                                Directory.CreateDirectory(pfIO.PathSuccess_O + "\\" + "LogSucces");
                            }
                        }
                        catch (Exception ea)
                        {
                            Console.WriteLine(ea);
                        }
                        finally
                        {
                            for (int i = 0; i <= GC.MaxGeneration; i++)
                            {
                                int count = GC.CollectionCount(i);
                                GC.Collect();
                            }
                            GC.WaitForPendingFinalizers();
                            GC.SuppressFinalize(this);
                        }
                        try
                        {
                            if (!Directory.Exists(pfIO.PathSuccess_O + "\\" + "Log"))
                            {
                                Directory.CreateDirectory(pfIO.PathSuccess_O + "\\" + "Log");
                            }
                        }
                        catch (Exception ea)
                        {
                            Console.WriteLine(ea);
                        }
                        finally
                        {
                            for (int i = 0; i <= GC.MaxGeneration; i++)
                            {
                                int count = GC.CollectionCount(i);
                                GC.Collect();
                            }
                            GC.WaitForPendingFinalizers();
                            GC.SuppressFinalize(this);
                        }
                        try
                        {
                            if (!Directory.Exists(pfIO.PathErr + "\\" + "Log_Resend"))
                            {
                                Directory.CreateDirectory(pfIO.PathErr + "\\" + "Log_Resend");
                            }
                        }
                        catch (Exception ea)
                        {
                            Console.WriteLine(ea);
                        }
                        finally
                        {
                            for (int i = 0; i <= GC.MaxGeneration; i++)
                            {
                                int count = GC.CollectionCount(i);
                                GC.Collect();
                            }
                            GC.WaitForPendingFinalizers();
                            GC.SuppressFinalize(this);
                        }
                        try
                        {
                            if (!Directory.Exists(pfIO.PathSuccess_O + "\\" + "Temp_Succes"))
                            {
                                Directory.CreateDirectory(pfIO.PathSuccess_O + "\\" + "Temp_Succes");
                            }
                        }
                        catch (Exception ea)
                        {
                            Console.WriteLine(ea);
                        }
                        finally
                        {
                            for (int i = 0; i <= GC.MaxGeneration; i++)
                            {
                                int count = GC.CollectionCount(i);
                                GC.Collect();
                            }
                            GC.WaitForPendingFinalizers();
                            GC.SuppressFinalize(this);
                        }
                        int counttimes = 0;
                        bool string_check__pdf;
                        bool string_check__xml;
                        _apimail.err_code = "FE91";
                        _apimail.err_msg = strFileNameExtension.Replace("~$", string.Empty) + "-//-" + "";
                        _apimail.input = dtParam.PathInput;
                        _apimail.path = dtParam.PathOutput;
                        _apimail.email = form.emailtxt.Text;
                        _apimail.taxseller = form.txtSellerTaxID.Text;
                        JObject ok_json = new JObject();
                        ok_json = JObject.Parse(strOutputFile.Message_Content);
                        Console.WriteLine(ok_json + " strOutputFile.Message_Content");

                        if (ok_json["status"].ToString() != "ER")
                        {
                            string Temp_succes = pfIO.PathSuccess_O + "\\" + "Temp_Succes";
                            CreateTextFile(pfIO.PathSuccess_O + "\\LogSucces\\" + "Success_" + Path.GetFileNameWithoutExtension(pathText) + ".txt", strOutputFile.Message_Content);
                        downloadpdfandxml:
                            counttimes = counttimes + 1;
                            DownloadFile(strOutputFile.MessageResultPDF, Temp_succes, this.pathOutput + "_PDF.pdf");
                            DownloadFile(strOutputFile.MessageResultXML, Temp_succes, this.pathOutput + "_XML.xml");
                            string_check__pdf = _checkfolder_pdf(this.pathOutput + "_PDF.pdf", Temp_succes);
                            string_check__xml = _checkfolder_xml(this.pathOutput + "_XML.xml", Temp_succes);
                            if (string_check__pdf == false && string_check__xml == false && counttimes <= 3)
                            {
                                goto downloadpdfandxml;
                            }
                            else if (string_check__pdf == false && string_check__xml == true && counttimes <= 3)
                            {
                                goto downloadpdfandxml;
                            }
                            else if (string_check__xml == false && string_check__pdf == true && counttimes <= 3)
                            {
                                goto downloadpdfandxml;
                            }
                            //MessageBox.Show(string_check__pdf.ToString());
                            if (string_check__pdf == false && string_check__xml == false)
                            {
                                _apimail.actionmsg = "ไม่สามารถดาวน์โหลดไฟล์ PDF และ XML ได้";
                                if (form.pingeng)
                                {
                                    _apimail.send_err_service();
                                }
                                CreateTextFile(pfIO.PathSuccess_O + "\\Log\\" + "Log_" + Path.GetFileNameWithoutExtension(pathText) + ".txt", "ไม่สามารถดาวน์โหลดไฟล์ PDF และ XML ได้");
                            }
                            else if (string_check__pdf == false && string_check__xml == true)
                            {
                                _apimail.actionmsg = "ไม่สามารถดาวน์โหลดไฟล์ PDF ได้";
                                if (form.pingeng)
                                {
                                    _apimail.send_err_service();
                                }
                                CreateTextFile(pfIO.PathSuccess_O + "\\Log\\" + "Log_" + Path.GetFileNameWithoutExtension(pathText) + ".txt", "ไม่สามารถดาวน์โหลดไฟล์ PDF ได้");
                            }
                            else if (string_check__xml == false && string_check__pdf == true)
                            {
                                _apimail.actionmsg = "ไม่สามารถดาวน์โหลดไฟล์ XML ได้";
                                if (form.pingeng)
                                {
                                    _apimail.send_err_service();
                                }
                                CreateTextFile(pfIO.PathSuccess_O + "\\Log\\" + "Log_" + Path.GetFileNameWithoutExtension(pathText) + ".txt", "ไม่สามารถดาวน์โหลดไฟล์ XML ได้");
                            }
                            try
                            {
                                var move_file_dis_form_pdf = Temp_succes + "\\" + this.pathOutput + "_PDF.pdf";
                                var move_file_dis_to_pdf = pfIO.PathSuccess_O + "\\" + this.pathOutput + "_PDF.pdf";
                                var move_file_dis_form_xml = Temp_succes + "\\" + this.pathOutput + "_XML.xml";
                                var move_file_dis_to_xml = pfIO.PathSuccess_O + "\\" + this.pathOutput + "_XML.xml";
                                if (string_check__pdf == true && string_check__xml == true)
                                {
                                    File.Move(move_file_dis_form_pdf, move_file_dis_to_pdf);
                                    File.Move(move_file_dis_form_xml, move_file_dis_to_xml);
                                }
                                else if (string_check__pdf == true && string_check__xml == false)
                                {
                                    File.Move(move_file_dis_form_pdf, move_file_dis_to_pdf);
                                }
                                else if (string_check__pdf == false && string_check__xml == true)
                                {
                                    File.Move(move_file_dis_form_xml, move_file_dis_to_xml);
                                }
                            }
                            catch (Exception ex)
                            {
                                Console.WriteLine(ex.Message);
                            }


                        }
                        else if (ok_json["errorCode"].ToString() == "ER011")
                        {
                            string Temp_succes = pfIO.PathSuccess_O + "\\" + "Temp_Succes";
                            CreateTextFile(pfIO.PathErr + "\\" + "Log_Resend\\" + "Resend_Success_" + Path.GetFileNameWithoutExtension(pathText) + ".txt", strOutputFile.Message_Content);
                            int counttimes_resend = 0;
                            string pathText_resend = "RESEND_" + UATorPROD + "_" + Path.GetFileNameWithoutExtension(dtParam.PathInput) + "_" + strDateTimeStamp + ".txt";
                            string pathOutput_resend = Path.GetFileNameWithoutExtension(pathText_resend);
                        downloadpdfandxml_resend:
                            counttimes_resend = counttimes_resend + 1;
                            DownloadFile(strOutputFile.MessageResultPDF, Temp_succes, pathOutput_resend + "_PDF.pdf");
                            DownloadFile(strOutputFile.MessageResultXML, Temp_succes, pathOutput_resend + "_XML.xml");
                            bool string_check__pdf_resend = _checkfolder_pdf(pathOutput_resend + "_PDF.pdf", Temp_succes);
                            bool string_check__xml_resend = _checkfolder_xml(pathOutput_resend + "_XML.xml", Temp_succes);
                            if (string_check__pdf_resend == false && string_check__xml_resend == false && counttimes_resend <= 3)
                            {
                                goto downloadpdfandxml_resend;
                            }
                            else if (string_check__pdf_resend == false && string_check__xml_resend == true && counttimes_resend <= 3)
                            {
                                goto downloadpdfandxml_resend;
                            }
                            else if (string_check__xml_resend == false && string_check__pdf_resend == true && counttimes_resend <= 3)
                            {
                                goto downloadpdfandxml_resend;
                            }
                            if (string_check__pdf_resend == false && string_check__xml_resend == false)
                            {
                                CreateTextFile(pfIO.PathSuccess_O + "\\" + "Log\\" + "Log_" + Path.GetFileNameWithoutExtension(pathText) + ".txt", "ไม่สามารถดาวน์โหลดไฟล์ PDF และ XML ได้");
                            }
                            else if (string_check__pdf_resend == false && string_check__xml_resend == true)
                            {
                                CreateTextFile(pfIO.PathSuccess_O + "\\Log\\" + "Log_" + Path.GetFileNameWithoutExtension(pathText) + ".txt", "ไม่สามารถดาวน์โหลดไฟล์ PDF ได้");
                            }
                            else if (string_check__xml_resend == false && string_check__pdf_resend == true)
                            {
                                CreateTextFile(pfIO.PathSuccess_O + "\\Log\\" + "Log_" + Path.GetFileNameWithoutExtension(pathText) + ".txt", "ไม่สามารถดาวน์โหลดไฟล์ XML ได้");
                            }
                            try
                            {
                                var move_file_dis_form_pdf = Temp_succes + "\\" + pathOutput_resend + "_PDF.pdf";
                                var move_file_dis_to_pdf = pfIO.PathSuccess_O + "\\" + pathOutput_resend + "_PDF.pdf";
                                var move_file_dis_form_xml = Temp_succes + "\\" + pathOutput_resend + "_XML.xml";
                                var move_file_dis_to_xml = pfIO.PathSuccess_O + "\\" + pathOutput_resend + "_XML.xml";
                                if (string_check__pdf_resend == true && string_check__xml_resend == true)
                                {
                                    File.Move(move_file_dis_form_pdf, move_file_dis_to_pdf);
                                    File.Move(move_file_dis_form_xml, move_file_dis_to_xml);
                                }
                                else if (string_check__pdf_resend == true && string_check__xml_resend == false)
                                {
                                    File.Move(move_file_dis_form_pdf, move_file_dis_to_pdf);
                                }
                                else if (string_check__pdf_resend == false && string_check__xml_resend == true)
                                {
                                    File.Move(move_file_dis_form_xml, move_file_dis_to_xml);
                                }
                            }
                            catch (Exception ex)
                            {
                                Console.WriteLine(ex.Message);
                            }

                        }



                        //Thread.Sleep(1000);
                        string namefilepdf = Path.GetFileName(dtParam.PathInput);
                        etaxOneth_Printer.Class1 _printer = new etaxOneth_Printer.Class1();
                        if (pfIO.TypePrinting == "A" && form.check___copies.Checked == false)
                        {
                            var Timestamp = new DateTimeOffset(DateTime.UtcNow).ToUnixTimeSeconds();
                            timest_process = Int32.Parse(Timestamp.ToString()) - timest_process;
                            Console.WriteLine(timest_process + " timest_process");
                            CreateTextFile(pfIO.LogTimeProcess + "\\LogProcess_Print_" + Path.GetFileNameWithoutExtension(pathText) + ".txt", "เอกสาร " + Path.GetFileNameWithoutExtension(pathText) + " ใช้เวลาการประมวลผลทั้งหมดประมาณ " + timest_process + " วินาที");
                            //PrinterSettings.SetDefaultPrinter(pfIO.Printer);
                            //ProcessStartInfo printProcessInfo = new ProcessStartInfo()
                            //{
                            //    UseShellExecute = true,
                            //    Verb = "print",
                            //    CreateNoWindow = true,
                            //    FileName = pfIO.PathSuccess_O + "\\" + this.pathOutput + "_PDF.pdf",
                            //    //Arguments = printDialog1.PrinterSettings.PrinterName.ToString(),
                            //    WindowStyle = ProcessWindowStyle.Hidden
                            //};
                            //_printer.PrintMethod("C:\\Users\\JIRAYU-NB\\Documents\\FillTEST\\output\\Success\\UAT_0105561072420_03-4-62T17-32-15_PDF.pdf", "ApeosPort-IV C5570 16", 1);
                            //Console.WriteLine(pfIO.PathSuccess_O + "\\" + this.pathOutput + "_PDF.pdf" + " " + pfIO.Printer + " " + short.Parse(form.input_copies.Text));
                            _printer.PrintMethod(pfIO.PathSuccess_O + "\\" + this.pathOutput + "_PDF.pdf", pfIO.Printer, short.Parse(form.input_copies.Text));

                            //try
                            //{
                            //    Process printProcess = new Process();
                            //    printProcess.StartInfo = printProcessInfo;
                            //    printProcess.Start();
                            //    //Thread.Sleep(3000);
                            //    //if (printProcess.HasExited == false)
                            //    //{
                            //    //    printProcess.Kill();
                            //    //}
                            //}
                            //catch (Exception ex)
                            //{
                            //    //MessageBox.Show(ex.ToString());
                            //    //MessageBox.Show("ไม่พบตัวอ่านไฟล์ของคุณ");
                            //}
                        }
                        else if (pfIO.TypePrinting == "A" && form.check___copies.Checked == true)
                        {
                            //MessageBox.Show(pfIO.PathInput.Split('\\')[pfIO.PathInput.Split('\\').Length -1]);
                            if (arrFiles__pcfg.Count() != 0)
                            {
                                foreach (var item in arrFiles__pcfg)
                                {
                                    string namefile__ = Path.GetFileName(item);
                                    var Timestamp = new DateTimeOffset(DateTime.UtcNow).ToUnixTimeSeconds();
                                    timest_process = Int32.Parse(Timestamp.ToString()) - timest_process;
                                    Console.WriteLine(timest_process + " timest_process");
                                    CreateTextFile(pfIO.LogTimeProcess + "\\LogProcess_Print_" + Path.GetFileNameWithoutExtension(pathText) + ".txt", "เอกสาร " + Path.GetFileNameWithoutExtension(pathText) + " ใช้เวลาการประมวลผลทั้งหมดประมาณ " + timest_process + " วินาที");
                                    if (Path.GetFileNameWithoutExtension(namefilepdf) == Path.GetFileNameWithoutExtension(namefile__))
                                    {
                                        // Open the text file using a stream reader.
                                        using (StreamReader sr = new StreamReader(item))
                                        {
                                            // Read the stream to a string, and write the string to the console.
                                            String line = sr.ReadToEnd();
                                            bool string___checkinpcfg = checkcopiesin__pcfg(line.Replace(" ", string.Empty).Replace(Environment.NewLine, string.Empty).Replace("\t", string.Empty));
                                            if (string___checkinpcfg == true)
                                            {
                                                try
                                                {
                                                    _printer.PrintMethod(pfIO.PathSuccess_O + "\\" + this.pathOutput + "_PDF.pdf", pfIO.Printer, short.Parse(line));
                                                }
                                                catch (Exception ex)
                                                {
                                                    Console.WriteLine(ex);
                                                }
                                                finally
                                                {
                                                    sr.Close();
                                                    File.Delete(item);
                                                }
                                            }
                                            else if (string___checkinpcfg == false)
                                            {
                                                sr.Close();
                                                File.Delete(item);
                                                MessageBox.Show("ไม่สามารถปริ้นได้เนื่องจาก จำนวน Copies ไม่ถูกต้อง *ควรระบุ 1-99*",
                                                                            "แจ้งเตือน",
                                                                MessageBoxButtons.OK,
                                                                MessageBoxIcon.Error);
                                            }
                                        }
                                    }
                                    else
                                    {
                                        File.Delete(item);
                                        MessageBox.Show("ชื่อไฟล์ .pcfg ไม่ตรงกับไฟล์ที่นำเข้า",
                                                        "แจ้งเตือน",
                                                        MessageBoxButtons.OK,
                                                        MessageBoxIcon.Warning);
                                    }
                                }

                            }
                            else
                            {
                                MessageBox.Show("ไม่พบไฟล์ .pcfg",
                                                    "แจ้งเตือน",
                                                    MessageBoxButtons.OK,
                                                    MessageBoxIcon.Warning);
                            }



                            //for (int i = 0; i < arrFiles__pcfg.Length; i++)
                            //{
                            //    string namefile__cfg = Path.GetFileName(arrFiles__pcfg[i]);
                            //    MessageBox.Show(namefilepdf);
                            //    if(Path.GetFileNameWithoutExtension(namefilepdf) == Path.GetFileNameWithoutExtension(namefile__cfg))
                            //    {
                            //        try
                            //        {   // Open the text file using a stream reader.
                            //            using (StreamReader sr = new StreamReader(arrFiles__pcfg[i]))
                            //            {
                            //                // Read the stream to a string, and write the string to the console.
                            //                String line = sr.ReadToEnd();
                            //                bool string___checkinpcfg = checkcopiesin__pcfg(line.Replace(" ", string.Empty).Replace(Environment.NewLine,string.Empty));
                            //                if(string___checkinpcfg == true)
                            //                {
                            //                    try
                            //                    {
                            //                        _printer.PrintMethod(pfIO.PathSuccess_O + "\\" + this.pathOutput + "_PDF.pdf", pfIO.Printer, short.Parse(line));
                            //                    }
                            //                    catch(Exception ex)
                            //                    {
                            //                        Console.WriteLine(ex);
                            //                    }
                            //                    finally
                            //                    {
                            //                        sr.Close();
                            //                        File.Delete(arrFiles__pcfg[i]);
                            //                    }


                            //                }
                            //                else if(string___checkinpcfg == false)
                            //                {
                            //                    sr.Close();
                            //                    File.Delete(arrFiles__pcfg[i]);
                            //                    MessageBox.Show("ไม่สามารถปริ้นได้เนื่องจาก จำนวน Copies เกิน 99 แผ่น",
                            //                                                "แจ้งเตือน",
                            //                                                MessageBoxButtons.OK,
                            //                                                MessageBoxIcon.Error);

                            //                }
                            //                //MessageBox.Show(string___checkinpcfg);

                            //            }
                            //        }
                            //        catch (Exception e)
                            //        {
                            //            Console.WriteLine("The file could not be read:");
                            //            Console.WriteLine(e.Message);
                            //        }
                            //    }
                            //    else
                            //    {

                            //    }

                            //}

                            //if(namefilepdf.Split(',')[namefilepdf.Split(',').Length - 1] + ".pcfg" == )
                            //bool check__print = _checkfolder(namefilepdf.Split(',')[namefilepdf.Split(',').Length - 1] + ".pcfg", pfIO.PathSuccess_O);
                            //MessageBox.Show(check__print.ToString());
                        }
                    }

                }

                form.pgbLoad.Value = 100;
                form.OutputPrc(100, "Export Data: 100%");
                //pgbLoad.Value = 100;
                //lbPercent.Text = "Export Data: 100%";
                //lbPercent.Refresh();
            }
            catch (FileNotFoundException ex)
            {
                form.txtStatus.Text += Environment.NewLine + "   -**********ไม่พบ Path File ที่คุณระบุ กรุณาเลือก Path File ให้ถูกต้อง!**********";
                strTempLogTime += "ไม่พบ Path File ที่คุณระบุ กรุณาเลือก Path File ให้ถูกต้อง!";

                if (chkOption == true)
                {
                    CreateTextFile(dtParam.PathOutput + "\\" + strFileName + "_" + strDateTimeStamp + "_Error.txt", "ไม่พบ Path File ที่คุณระบุ กรุณาเลือก Path File ให้ถูกต้อง");
                }
                else
                {
                    CreateTextFile(pfIO.PathErr + "\\" + strFileName + "_" + strDateTimeStamp + "_Error.txt", "ไม่พบ Path File ที่คุณระบุ กรุณาเลือก Path File ให้ถูกต้อง");
                }

                form.txtStatus.Refresh();
                cntFail++;
                //MessageBox.Show("ไม่พบ Path File ที่คุณระบุ กรุณาเลือก Path File ให้ถูกต้อง");
            }
            catch (System.IndexOutOfRangeException e)
            {
                form.txtStatus.Text += Environment.NewLine + "   -**********ไฟล์ของคุณมีข้อผิดพลาดในข้อมูลที่ใส่!**********";
                strTempLogTime += "กรุณาตรวจสอบไฟล์ของคุณ!";

                if (chkOption == true)
                {
                    CreateTextFile(dtParam.PathOutput + "\\" + strFileName + "_" + strDateTimeStamp + "_Error.txt", "ไฟล์ของคุณมีข้อผิดพลาด กรุณาตรวจสอบและใส่ข้อมูลให้ถูกต้อง");
                    CreateTextFile(dtParam.PathOutput + "\\" + strFileName + "_" + strDateTimeStamp + "_ErrorServiceOrProgram.txt", e.Message + " " + e.Data);
                }
                else
                {
                    CreateTextFile(pfIO.PathErr + "\\" + strFileName + "_" + strDateTimeStamp + "_Error.txt", "ไฟล์ของคุณมีข้อผิดพลาด กรุณาตรวจสอบและใส่ข้อมูลให้ถูกต้อง");
                    CreateTextFile(pfIO.PathErr + "\\" + strFileName + "_" + strDateTimeStamp + "_ErrorServiceOrProgram.txt", e.Message + " " + e.Data);
                }

                form.txtStatus.Refresh();
                cntFail++;
            }
            catch (DirectoryNotFoundException ex)
            {
                Console.WriteLine(ex.Message);
            }
            catch (XmlException ex)
            {
                FileVersionInfo fvi = FileVersionInfo.GetVersionInfo(assembly.Location);
                string version = fvi.FileVersion;
                //MessageBox.Show("ไฟล์มีปัญหา!!!");
                Console.WriteLine(ex.Message);
                form.txtStatus.Text += Environment.NewLine + "   -**********Convert Fail!**********";
                strTempLogTime += " Convert Fail!" + ex.Message + "Version : " + version;
                //MessageBox.Show(ErrorMessage);
                if (chkOption == true)
                {
                    CreateTextFile(dtParam.PathOutput + "\\" + strFileName + "_" + strDateTimeStamp + "_Error.txt", "Convert Fail" + ex.Message);

                }
                else
                {

                    CreateTextFile(dtParam.PathOutput + "\\" + strFileName + "_" + strDateTimeStamp + "_Error.txt", "Convert Fail" + ex.Message);

                }

                form.txtStatus.Refresh();
                cntFail++;
            }
            catch (Exception ex)
            {

                //MessageBox.Show(ex.Message);
                FileVersionInfo fvi = FileVersionInfo.GetVersionInfo(assembly.Location);
                string version = fvi.FileVersion;
                form.txtStatus.Text += Environment.NewLine + "   -**********Convert Fail!**********";
                strTempLogTime += " Convert Fail!"+ex.Message+" Version : " + version;
                if (chkOption == true)
                {
                    CreateTextFile(dtParam.PathOutput + "\\" + strFileName + "_" + strDateTimeStamp + "_Error.txt", "Convert Fail" + ex.Message);

                }
                else
                {

                    CreateTextFile(dtParam.PathOutput + "\\" + strFileName + "_" + strDateTimeStamp + "_Error.txt", "Convert Fail" + ex.Message);

                }

                form.txtStatus.Refresh();
                cntFail++;
            }
            finally
            {
                lstTempSumAmount.Clear();
                for (int i = 0; i <= GC.MaxGeneration; i++)
                {
                    int count = GC.CollectionCount(i);
                    GC.Collect();
                }
                GC.WaitForPendingFinalizers();
                GC.SuppressFinalize(this);
            }
        }
        public void WorkProcess_forS06_ListItem(PathFilesIO pfIO, string strFileName, string strFileNameExtension, string strFileNamePDF, out int cntFail, etaxOneth form)
        {
            a = null;
            b = null;
            a_with_b = null;
        loop:

            cntFail = 0;
            strOutputFile = new DataOutput();
            string strTaxID;
            string typedoc = "", sellertaxid = "", sellerbranchid = "", document_name = "", document_id = "", document_issue_dtm = "", create_purpose_code = "",
                create_purpose = "", additional_ref_assign_id = "", additional_ref_issue_dtm = "", buyer_name = "", buyer_branch_id = "",
                buyer_tax_id = "", buyer_uriid = "", buyer_address = "", buyer_countrypostcode = "", rangeoflistitem = "", noitem = "", description = "",
                priceunit = "", quanlity = "", amount = "", totalamount = "", vat = "", total = "", document_remark = "",
                discount = "", totaldiscount = "", original_total_amount = "", line_total_amount = "", adjusted_information_amount = "", allowance_total_amount = "",
                tax_basis_total_amount = "", countrybuyer = "", typebuyer = "", buyer_order_assign_id = "", buyer_order_issue_dtm = "", buyer_address2 = "", vat_rate = "", amount_deposit = "";
            string listheader = "", listfooter = "";
            string type_Value_Sum = "";
            if (dtParam.ServiceURL == "https://uatetaxsp.one.th/etaxdocumentws/etaxsigndocument")
            {
                UATorPROD = "UAT";
            }
            else
            {
                UATorPROD = "PROD";
            }
            try
            {
                pathText = string.Empty;
                if (Path.GetExtension(dtParam.PathInput).Equals(".xlsx") || Path.GetExtension(dtParam.PathInput).Equals(".xls"))
                {
                    try
                    {
                        using (var reader = new StreamReader(dtParam.PathConfigExcel))
                        {
                            List<string> listA = new List<string>();
                            List<string> listB = new List<string>();
                            while (!reader.EndOfStream)
                            {
                                var line = reader.ReadLine();
                                var values = line.Split(';');
                                listA.Add(values[0]);
                                //listB.Add(values[1]);
                            }
                            foreach (var item in listA)
                            {
                                string[] arrItem = item.Split(',');
                                //Console.WriteLine(arrItem[0]);
                                switch (arrItem[0].ToLower().Trim(' '))
                                {
                                    case "header":
                                        if (!arrItem[2].Equals(""))
                                            listheader = arrItem[2];
                                        else
                                            listheader = "";
                                        break;
                                    case "typedoc":
                                        if (!arrItem[2].Equals(""))
                                            typedoc = arrItem[2];
                                        else
                                            typedoc = arrItem[2];
                                        break;
                                    case "discount":
                                        if (!arrItem[2].Equals(""))
                                        {
                                            discount = arrItem[2] + "," + arrItem[3];
                                            AllChar = AllChar + arrItem[2];
                                            charRange_.Add(arrItem[2]);
                                        }
                                        else
                                        {
                                            discount = arrItem[2];
                                        }
                                        break;
                                    case "sellertaxid":
                                        if (!arrItem[2].Equals(""))
                                            sellertaxid = arrItem[2] + "," + arrItem[3];
                                        else
                                            sellertaxid = arrItem[2];
                                        break;
                                    case "sellerbranchid":
                                        if (!arrItem[2].Equals(""))
                                            sellerbranchid = arrItem[2] + "," + arrItem[3];
                                        else
                                            sellerbranchid = arrItem[2];
                                        break;
                                    case "document_name":
                                        if (!arrItem[2].Equals(""))
                                            document_name = arrItem[2] + "," + arrItem[3];
                                        else
                                            document_name = arrItem[2];
                                        break;
                                    case "document_id":
                                        if (!arrItem[2].Equals(""))
                                            document_id = arrItem[2] + "," + arrItem[3];
                                        else
                                            document_id = arrItem[2];
                                        break;
                                    case "document_remark":
                                        if (!arrItem[2].Equals(""))
                                            document_remark = arrItem[2] + "," + arrItem[3];
                                        else
                                            document_remark = arrItem[2];
                                        break;
                                    case "document_issue_dtm":
                                        if (!arrItem[2].Equals(""))
                                            document_issue_dtm = arrItem[2] + "," + arrItem[3];
                                        else
                                            document_issue_dtm = arrItem[2];
                                        break;
                                    case "create_purpose_code":
                                        if (!arrItem[2].Equals(""))
                                            create_purpose_code = arrItem[2] + "," + arrItem[3];
                                        else
                                            create_purpose_code = arrItem[2];
                                        break;
                                    case "create_purpose":
                                        if (!arrItem[2].Equals(""))
                                            create_purpose = arrItem[2] + "," + arrItem[3];
                                        else
                                            create_purpose = arrItem[2];
                                        break;
                                    case "additional_ref_assign_id":
                                        if (!arrItem[2].Equals(""))
                                            additional_ref_assign_id = arrItem[2] + "," + arrItem[3];
                                        else
                                            additional_ref_assign_id = arrItem[2];
                                        break;
                                    case "additional_ref_issue_dtm":
                                        if (!arrItem[2].Equals(""))
                                            additional_ref_issue_dtm = arrItem[2] + "," + arrItem[3];
                                        else
                                            additional_ref_issue_dtm = arrItem[2];
                                        break;
                                    case "buyer_name":
                                        if (!arrItem[2].Equals(""))
                                            buyer_name = arrItem[2] + "," + arrItem[3];
                                        else
                                            buyer_name = arrItem[2];
                                        break;
                                    case "buyer_tax_id":
                                        if (!arrItem[2].Equals(""))
                                            buyer_tax_id = arrItem[2] + "," + arrItem[3];
                                        else
                                            buyer_tax_id = arrItem[2];
                                        break;
                                    case "buyer_branch_id":
                                        if (!arrItem[2].Equals(""))
                                            buyer_branch_id = arrItem[2] + "," + arrItem[3];
                                        else
                                            buyer_branch_id = arrItem[2];
                                        break;
                                    case "buyer_uriid":
                                        if (!arrItem[2].Equals(""))
                                            buyer_uriid = arrItem[2] + "," + arrItem[3];
                                        else
                                            buyer_uriid = arrItem[2];
                                        break;
                                    case "buyer_address":
                                        if (!arrItem[2].Equals(""))
                                            buyer_address = arrItem[2] + "," + arrItem[3];
                                        else
                                            buyer_address = arrItem[2];
                                        break;
                                    case "buyer_address2":
                                        if (!arrItem[2].Equals(""))
                                            buyer_address2 = arrItem[2] + "," + arrItem[3];
                                        else
                                            buyer_address2 = arrItem[2];
                                        break;
                                    case "buyer_postcode":
                                        if (!arrItem[2].Equals(""))
                                            buyer_countrypostcode = arrItem[2] + "," + arrItem[3];
                                        else
                                            buyer_countrypostcode = arrItem[2];
                                        break;
                                    case "amount_deposit":
                                        if (!arrItem[2].Equals(""))
                                            amount_deposit = arrItem[2] + "," + arrItem[3];
                                        else
                                            amount_deposit = arrItem[2];
                                        break;
                                    //case "rangeoflistitem":
                                    //    rangeoflistitem = arrItem[2];
                                    //    string[] splitData = arrItem[2].Split('_');
                                    //    numberRange = splitData[0].Split('-');
                                    //    charRange = splitData[1].Split('-');

                                    //    foreach (var k in charRange)
                                    //    {
                                    //        AllChar = AllChar + k;
                                    //    }

                                    //    break;
                                    case "no.item":
                                        if (!arrItem[2].Equals(""))
                                        {
                                            noitem = arrItem[2] + "," + arrItem[3];
                                            AllChar = AllChar + arrItem[2];
                                            charRange_.Add(arrItem[2]);
                                        }
                                        else
                                        {
                                            noitem = arrItem[2];
                                        }
                                        break;
                                    case "description":
                                        if (!arrItem[2].Equals(""))
                                        {
                                            description = arrItem[2] + "," + arrItem[3];
                                            AllChar = AllChar + arrItem[2];
                                            charRange_.Add(arrItem[2]);
                                        }
                                        else
                                        {
                                            description = arrItem[2];
                                        }
                                        break;
                                    case "price/unit":
                                        if (!arrItem[2].Equals(""))
                                        {
                                            priceunit = arrItem[2] + "," + arrItem[3];
                                            AllChar = AllChar + arrItem[2];
                                            charRange_.Add(arrItem[2]);
                                        }
                                        else
                                        {
                                            priceunit = arrItem[2];
                                        }
                                        break;
                                    case "quanlity":
                                        if (!arrItem[2].Equals(""))
                                        {
                                            quanlity = arrItem[2] + "," + arrItem[3];
                                            AllChar = AllChar + arrItem[2];
                                            charRange_.Add(arrItem[2]);
                                        }
                                        else
                                        {
                                            quanlity = arrItem[2];
                                        }
                                        break;
                                    case "amount":
                                        if (!arrItem[2].Equals(""))
                                        {
                                            amount = arrItem[2] + "," + arrItem[3];
                                            AllChar = AllChar + arrItem[2];
                                            charRange_.Add(arrItem[2]);
                                        }
                                        else
                                        {
                                            amount = arrItem[2];
                                        }
                                        break;
                                    case "totalamount":
                                        if (!arrItem[2].Equals(""))
                                            totalamount = arrItem[2] + "," + arrItem[3];
                                        else
                                            totalamount = arrItem[2];
                                        break;
                                    //case "totaldiscount":
                                    //    if (!arrItem[2].Equals(""))
                                    //        totaldiscount = arrItem[2] + "," + arrItem[3];
                                    //    else
                                    //        totaldiscount = arrItem[2];
                                    //    break;
                                    case "vat":
                                        if (!arrItem[2].Equals(""))
                                            vat = arrItem[2] + "," + arrItem[3];
                                        else
                                            vat = arrItem[2];
                                        break;
                                    case "total":
                                        if (!arrItem[2].Equals(""))
                                            total = arrItem[2] + "," + arrItem[3];
                                        else
                                            total = arrItem[2];
                                        break;
                                    case "original_total_amount":
                                        if (!arrItem[2].Equals(""))
                                            original_total_amount = arrItem[2] + "," + arrItem[3];
                                        else
                                            original_total_amount = arrItem[2];
                                        break;
                                    case "line_total_amount":
                                        if (!arrItem[2].Equals(""))
                                            line_total_amount = arrItem[2] + "," + arrItem[3];
                                        else
                                            line_total_amount = arrItem[2];
                                        break;
                                    case "adjusted_information_amount":
                                        if (!arrItem[2].Equals(""))
                                            adjusted_information_amount = arrItem[2] + "," + arrItem[3];
                                        else
                                            adjusted_information_amount = arrItem[2];
                                        break;
                                    case "allowance_total_amount":
                                        if (!arrItem[2].Equals(""))
                                            allowance_total_amount = arrItem[2] + "," + arrItem[3];
                                        else
                                            allowance_total_amount = arrItem[2];
                                        break;
                                    case "tax_basis_total_amount":
                                        if (!arrItem[2].Equals(""))
                                            tax_basis_total_amount = arrItem[2] + "," + arrItem[3];
                                        else
                                            tax_basis_total_amount = arrItem[2];
                                        break;
                                    case "countrybuyer":
                                        if (!arrItem[2].Equals(""))
                                            countrybuyer = arrItem[2] + "," + arrItem[3];
                                        else
                                            countrybuyer = arrItem[2];
                                        break;
                                    case "typebuyer":
                                        if (!arrItem[2].Equals(""))
                                            typebuyer = arrItem[2] + "," + arrItem[3];
                                        else
                                            typebuyer = arrItem[2];
                                        break;
                                    case "buyer_order_assign_id":
                                        if (!arrItem[2].Equals(""))
                                            buyer_order_assign_id = arrItem[2] + "," + arrItem[3];
                                        else
                                            buyer_order_assign_id = arrItem[2];
                                        break;
                                    case "buyer_order_issue_dtm":
                                        if (!arrItem[2].Equals(""))
                                            buyer_order_issue_dtm = arrItem[2] + "," + arrItem[3];
                                        else
                                            buyer_order_issue_dtm = arrItem[2];
                                        break;
                                    case "vat_rate":
                                        if (!arrItem[2].Equals(""))
                                            vat_rate = arrItem[2] + "," + arrItem[3];
                                        else
                                            vat_rate = arrItem[2];
                                        break;
                                    case "footer":
                                        if (!arrItem[2].Equals(""))
                                            listfooter = arrItem[2];
                                        else
                                            listfooter = "";
                                        break;
                                    default:
                                        break;
                                }
                            }
                            //foreach (var item in listA)
                            //{
                            //    Console.WriteLine(item);
                            //}
                            //Console.WriteLine(listA[0]);
                            //Console.WriteLine(listB);
                        } //End of Using for Read ConfigExcel
                        form.pgbLoad.Value = 0;
                        form.OutputPrc(0, "Export Data: 0%");
                        //pgbLoad.Value = 0;
                        //lbPercent.Text = "Export Data: 0%";
                        //lbPercent.Refresh();

                        List<string> lstDataRow = new List<string>();
                        List<string> lstDataMenu = new List<string>();
                        string strSheetName = string.Empty;
                        BGroup grpB = new BGroup();
                        CGroup grpC = new CGroup();
                        LGroup grpL = new LGroup();
                        HGroup grpH = new HGroup();
                        FGroup grpF = new FGroup();
                        Workbook workbook = new Workbook();

                        workbook.LoadFromFile(dtParam.PathInput);
                        sheet = workbook.Worksheets[0];
                        int totel_Rows_Header = sheet.Rows.Count() - int.Parse(listheader);
                        int totel_Rows_Footer = sheet.Rows.Count() - int.Parse(listfooter);
                        int totel_Header = int.Parse(listheader);
                        int totel_Rows = sheet.Rows.Count();
                        Console.WriteLine("totel_Rows_unFooter : " + totel_Rows_Footer + " totel_Rows_unHeader :" + totel_Rows_Header);
                        try
                        {
                            if (DateTime.TryParse(getvalue(document_issue_dtm), out dateValue))
                            {
                                arrDateSplit = getvalue(document_issue_dtm).Split('/');
                                strYear = DateTime.Now.Year.ToString();
                                strYearFront = strYear.Substring(0, 2);
                                DiffOfYears = int.Parse(strYear) - (int.Parse(arrDateSplit[2].Split(' ')[0]) - 543); //ต้องลบ543เพราะ โปรแกรม+543ให้เองอัตโนมัติจึงลบออกเพื่อให้ได้ค่าที่ถูกต้อง
                                if (DiffOfYears < 0)
                                {
                                    DiffOfYears = 543;
                                }
                                else
                                {
                                    DiffOfYears = 0;
                                }
                                if (((int.Parse(arrDateSplit[2].Split(' ')[0]) - 543) - DiffOfYears) < 2000)
                                {
                                    years = (int.Parse(arrDateSplit[2].Split(' ')[0])) - DiffOfYears;
                                }
                                else
                                {
                                    years = (int.Parse(arrDateSplit[2].Split(' ')[0]) - 543) - DiffOfYears;
                                }
                                Console.WriteLine("YearsNow: " + strYear + " Years : " + arrDateSplit[2].Split(' ')[0]);
                                strDocID = arrDateSplit[2].Split(' ')[0] + "-" + Month[arrDateSplit[1].Replace(" ", string.Empty)] + "-" + Day[arrDateSplit[0].Replace(" ", string.Empty)];

                            }
                            else
                            {
                                MessageBox.Show("document_issue_dtm ไม่ถูกต้อง => " + getvalue(document_issue_dtm));
                            }
                        }
                        catch (ArgumentException e)
                        {
                            Console.WriteLine(e.Message);
                        }
                        catch (NullReferenceException e)
                        {
                            Console.WriteLine(e.Message);
                        }
                        catch (IndexOutOfRangeException e)
                        {
                            Console.WriteLine(e.Message);
                        }
                        catch (KeyNotFoundException e)
                        {
                            Console.WriteLine("Date Wrong");
                        }
                        //MessageBox.Show(strDocID);

                        if (form.txtStatus != null && !form.txtStatus.Text.Equals(""))
                        {
                            form.txtStatus.Text += Environment.NewLine + "เลขที่เอกสาร " + getvalue(document_id) + " วันที่ " + strDocID + " ชื่อไฟล์ " + strFileNameExtension + ":";
                            strTempLogTime += Environment.NewLine + "เลขที่เอกสาร " + getvalue(document_id) + " วันที่ " + strDocID + " ชื่อไฟล์ " + strFileNameExtension + ":";
                        }
                        else
                        {
                            form.txtStatus.Text = "เลขที่เอกสาร " + getvalue(document_id) + " วันที่ " + strDocID + " ชื่อไฟล์ " + strFileNameExtension + " :";
                            strTempLogTime = "เลขที่เอกสาร " + getvalue(document_id) + " วันที่ " + strDocID + " ชื่อไฟล์ " + strFileNameExtension + " :";
                        }

                        form.pgbLoad.Value = 10;
                        form.OutputPrc(10, "Export Data: 10%");

                        for (int x = totel_Header; x <= totel_Rows_Footer; x++)
                        {
                            int y = 0;
                            for (int k = 0; k < charRange_.ToArray().Length; k++)
                            {

                                //if (getvalue(charRange_[k] + (totel_Header + 1) + "," + "value").Replace(" ", string.Empty) == "จำนวนเงิน")
                                //{
                                //    type_Value_Sum = "Sum_NO_VAT";
                                //}
                                //else if (getvalue(charRange_[k] + (totel_Header + 1) + "," + "value").Replace(" ", string.Empty) == "ราคารวมภาษี")
                                //{
                                //    type_Value_Sum = "Sum_VAT";
                                //}
                                Console.WriteLine(type_Value_Sum + " type_Value_Sum");
                                string value;
                                if (charRange_[k] == noitem.Split(',')[0])
                                {
                                    Regex regex = new Regex(@"^[0-9]*$");
                                    value = getvalue(charRange_[k] + x + "," + noitem.Split(',')[1]);
                                    if (regex.IsMatch(value.Replace(" ", string.Empty)) && !value.Replace(" ", string.Empty).Equals(""))
                                    {
                                        Console.WriteLine(value);
                                        object cellValue = value;
                                        y = x;
                                        lstDataMenu.Add(cellValue.ToString());
                                    }
                                }
                                else if (charRange_[k] == description.Split(',')[0])
                                {
                                    if (y != 0)
                                    {
                                        value = getvalue(charRange_[k] + y + "," + description.Split(',')[1]);
                                        Console.WriteLine(value);
                                        object cellValue = value;
                                        lstDataMenu.Add(cellValue.ToString());
                                    }
                                }
                                else if (charRange_[k] == priceunit.Split(',')[0])
                                {
                                    if (y != 0)
                                    {
                                        value = getvalue(charRange_[k] + y + "," + priceunit.Split(',')[1]);
                                        Console.WriteLine(value);
                                        if (!value.Equals(""))
                                        {
                                            object cellValue = value;
                                            lstDataMenu.Add(cellValue.ToString());
                                        }
                                        else
                                        {
                                            object cellValue = "";
                                            lstDataMenu.Add(cellValue.ToString());
                                        }
                                    }

                                }
                                else if (charRange_[k] == quanlity.Split(',')[0])
                                {
                                    if (y != 0)
                                    {
                                        value = getvalue(charRange_[k] + y + "," + quanlity.Split(',')[1]);
                                        if (!value.Equals(""))
                                        {
                                            object cellValue = Regex.Match(RemoveComma(value), @"\d+").Value;
                                            Console.WriteLine(cellValue);
                                            lstDataMenu.Add(cellValue.ToString());
                                        }
                                        else
                                        {
                                            object cellValue = "";
                                            lstDataMenu.Add(cellValue.ToString());
                                        }
                                    }
                                }
                                else if (charRange_[k] == discount.Split(',')[0])
                                {
                                    if (y != 0)
                                    {
                                        value = getvalue(charRange_[k] + y + "," + discount.Split(',')[1]);
                                        if (!value.Equals(""))
                                        {
                                            object cellValue = value;
                                            lstDataMenu.Add(cellValue.ToString());
                                        }
                                        else
                                        {
                                            object cellValue = "";
                                            lstDataMenu.Add(cellValue.ToString());
                                        }
                                    }
                                }
                                else if (charRange_[k] == amount.Split(',')[0])
                                {
                                    if (y != 0)
                                    {
                                        if (!sheet.Range[charRange_[k] + y].Value.Equals(""))
                                        {
                                            value = getvalue(charRange_[k] + y + "," + amount.Split(',')[1]);
                                            if (!value.Equals(""))
                                            {
                                                object cellValue = value;
                                                try
                                                {
                                                    lstDataMenu.Add(Double.Parse(cellValue.ToString()).ToString("0.00"));
                                                }
                                                catch (FormatException e)
                                                {
                                                    Console.WriteLine(e);
                                                }
                                            }
                                            else
                                            {
                                                object cellValue = "";
                                                lstDataMenu.Add(cellValue.ToString());
                                            }

                                        }
                                        else
                                        {
                                            object cellValue = "";
                                            lstDataMenu.Add(cellValue.ToString());
                                        }
                                    }
                                }
                                else
                                {
                                    object cellValue = "";
                                    lstDataMenu.Add(cellValue.ToString());
                                }
                            }
                        }
                        form.pgbLoad.Value = 20;
                        form.OutputPrc(20, "Export Data: 20%");
                        //pgbLoad.Value = 20;
                        //lbPercent.Text = "Export Data: 20%";
                        //lbPercent.Refresh();

                        //Type C
                        grpC.Data_Type = "C";
                        //MessageBox.Show(sheet.Range["F7"].Value.Replace(" ", string.Empty));


                        //Console.WriteLine(RecursionTaxid(" " + sheet.Range[sellertaxid].Value.Replace(" ", string.Empty) + " "));

                        //grpC.Seller_Tax_ID = RecursionTaxid(" " + getvalue(sellertaxid).Replace(" ", string.Empty).Replace("-",string.Empty) + " ").Replace(" ", string.Empty); //เลขประจำตัวผู้เสียภาษี
                        //Console.WriteLine(sheet.Range[sellerbranchid.Split(',')[0]].Value.Replace(" ", string.Empty) + " sellerbranchid");
                        grpC.Seller_Tax_ID = dtParam.SellerTaxID;                                                                    //grpC.Seller_Branch_ID = sheet.Range["L8"].Value.Replace(" ", string.Empty); //เลขสาขาประกอบการ
                                                                                                                                     //grpC.Seller_Branch_ID = dtParam.BranchID;
                        Console.WriteLine(getvalue(sellerbranchid).Replace(" ", string.Empty));
                        if (!sellerbranchid.Equals(""))
                        {
                            try
                            {
                                if (!getvalue(sellerbranchid).Equals(""))
                                {

                                    grpC.Seller_Branch_ID = branch_seller((sheet.Range[sellerbranchid.Split(',')[0]].Value.Replace(" ", string.Empty)).ToString());
                                }
                                else
                                {
                                    grpC.Seller_Branch_ID = "00000";
                                }
                            }
                            catch (IndexOutOfRangeException e)
                            {
                                Console.WriteLine(e.Message);
                            }
                            catch (NullReferenceException e)
                            {
                                Console.WriteLine(e.Message);
                            }

                        }
                        else
                        {
                            grpC.Seller_Branch_ID = "00000";
                        }
                        Console.WriteLine(grpC.Seller_Branch_ID + "sellerbranchid");

                        //grpC.File_Name = RecursionTaxid(" " + getvalue(sellertaxid).Replace(" ", string.Empty) + " ").Replace(" ", string.Empty) + ".txt"; //ชื่อไฟล์
                        grpC.File_Name = grpC.Seller_Tax_ID + ".txt";
                        form.pgbLoad.Value = 30;
                        form.OutputPrc(30, "Export Data: 30%");
                        grpB.Data_Type = "B";
                        int iComSplit = getvalue(buyer_name).IndexOf("(");
                        if (iComSplit != -1)
                        {
                            grpB.Buyer_Name = getvalue(buyer_name).Substring(0, iComSplit - 1); //CompanyName
                        }
                        else
                        {
                            grpB.Buyer_Name = getvalue(buyer_name).Replace(" ", string.Empty); //CompanyName
                        }
                        grpB.Buyer_Name = getvalue(buyer_name);
                        grpB.Buyer_Phone_No = "";
                        try
                        {
                            strTaxID = RecursionTaxid(" " + getvalue(buyer_tax_id).Replace(" ", string.Empty).Replace("-", string.Empty) + " "); //ประเภทผู้เสียภาษี

                            strTaxID = strTaxID.Replace(" ", string.Empty);
                        }
                        catch (NullReferenceException e)
                        {
                            strTaxID = "N/A";
                        }
                        catch (Exception e)
                        {
                            strTaxID = "";
                        }
                        if (!buyer_branch_id.Equals(""))
                        {
                            try
                            {
                                if (!getvalue(buyer_branch_id).Equals(""))
                                {
                                    string buyerbrach_String = branch_buyyer(getvalue(buyer_branch_id).Replace(" ", string.Empty));
                                    grpB.Buyer_Branch_ID = buyerbrach_String;
                                }
                                else
                                {
                                    grpB.Buyer_Branch_ID = "";
                                }
                            }
                            catch (IndexOutOfRangeException e)
                            {
                                grpB.Buyer_Branch_ID = "";
                            }
                            catch (Exception e)
                            {
                                grpB.Buyer_Branch_ID = "";
                            }
                        }
                        else
                        {
                            grpB.Buyer_Branch_ID = "";
                        }

                        try
                        {
                            if (getvalue(buyer_countrypostcode) == "")
                            {
                                grpB.Buyer_Post_Code = ("00000");
                            }
                            else
                            {
                                try
                                {
                                    RealValue = "";
                                    grpB.Buyer_Post_Code = RecursionPostCode(" " + getvalue(buyer_countrypostcode) + " ");
                                }
                                catch (IndexOutOfRangeException e)
                                {
                                    grpB.Buyer_Post_Code = "00000";
                                }
                                catch (Exception e)
                                {
                                    grpB.Buyer_Post_Code = "00000";
                                }
                            }
                        }
                        catch (Exception e)
                        {
                            grpB.Buyer_Post_Code = "";
                        }
                        string keyType = string.Empty;
                        if (strTaxID.Equals("N/A"))
                        {
                            keyType = "4";
                        }
                        else if (strTaxID.Equals(""))
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
                            else if (Char.TryParse(countTaxNum.ToString().Substring(0, 1), out char data) && (grpB.Buyer_Branch_ID == null || grpB.Buyer_Branch_ID.Equals(""))) //อนาคตหากมีเลขที่ PassPort
                            {
                                keyType = "3";
                            }
                            else
                            {
                                keyType = "4";
                                strTaxID = getvalue(buyer_tax_id).Replace(" ", string.Empty).Replace("-", string.Empty);
                            }

                        }
                        grpB.Buyer_Tax_ID_Type = BuyerTaxType[keyType];
                        if (!typebuyer.Equals(""))
                        {
                            if (!getvalue(typebuyer).Equals(""))
                            {
                                grpB.Buyer_Tax_ID_Type = getvalue(typebuyer).Replace(" ", string.Empty).Replace("-", string.Empty);
                                strTaxID = getvalue(buyer_tax_id).Replace(" ", string.Empty).Replace("-", string.Empty);
                            }
                        }
                        if (strTaxID.Equals(""))
                        {
                            grpB.Buyer_Tax_ID = DoubleQuote(" "); //เลขที่ประจำตัวผู้เสียภาษี
                        }
                        else
                        {
                            grpB.Buyer_Tax_ID = strTaxID; //เลขที่ประจำตัวผู้เสียภาษี
                        }

                        if (buyer_uriid.Equals(""))
                        {
                            grpB.Buyer_URIID = "";
                        }
                        else
                        {
                            grpB.Buyer_URIID = getvalue(buyer_uriid).Replace(" ", string.Empty);
                        }
                        grpB.Buyer_Add_Line1 = getvalue(buyer_address);
                        grpB.Buyer_Add_Line2 = getvalue(buyer_address2);
                        form.pgbLoad.Value = 40;
                        form.OutputPrc(40, "Export Data: 40%");
                        int iCountRound = 0;
                        List<LGroup> lstGrpL = new List<LGroup>();
                        int coun = sheet.Rows.Count();
                        int distantNameItem = (AllChar.IndexOf(description.Split(',')[0]) - AllChar.IndexOf(noitem.Split(',')[0]));
                        int distantPriceUnit = (AllChar.IndexOf(priceunit.Split(',')[0]) - AllChar.IndexOf(noitem.Split(',')[0]));
                        int distantQuanlity = (AllChar.IndexOf(quanlity.Split(',')[0]) - AllChar.IndexOf(noitem.Split(',')[0]));
                        int distantAmount = (AllChar.IndexOf(amount.Split(',')[0]) - AllChar.IndexOf(noitem.Split(',')[0]));
                        int distantDiscount = (AllChar.IndexOf(discount.Split(',')[0]) - AllChar.IndexOf(noitem.Split(',')[0]));
                        for (int x = 0; x < lstDataMenu.Count; x++)
                        {
                            bool chkSting = false;
                            bool chkNum = false;
                            Double value = 0;
                            string patternChkString = @"([a-zA-Zก-๙0-9])";
                            if (!lstDataMenu[x].Equals(""))
                            {
                                chkSting = Regex.IsMatch(lstDataMenu[x], patternChkString);
                                chkNum = Double.TryParse(lstDataMenu[x], out value);
                            }
                            if (chkSting == true && chkNum == true)
                            {
                                if (iCountRound > 0)
                                {
                                    Console.WriteLine("LengthOfProduct_Desc => " + grpL.Product_Desc.Length);
                                    if (grpL.Product_Desc == null || grpL.Product_Desc.Equals(""))
                                    {
                                        grpL.Product_Desc = "";
                                    }
                                    else
                                    {

                                        if (grpL.Product_Desc.Length > 256)
                                        {
                                            string a = grpL.Product_Desc.Substring(0, 256);
                                            Console.WriteLine("a => " + a);
                                            string[] b = a.Split(' ');
                                            Console.WriteLine(b.Length);
                                            for (int i = 0; i < b.Length - 1; i++)
                                            {
                                                a_with_b += b[i] + " ";
                                            }
                                            Console.WriteLine("a_With_b => " + a_with_b);
                                            Console.WriteLine("LengthOfa_with_b => " + a_with_b.Length);
                                            grpL.Product_Remark = DoubleQuote(grpL.Product_Desc.Substring(a_with_b.Length));
                                            Console.WriteLine("Product_Remark => " + grpL.Product_Remark);
                                            grpL.Product_Desc = a_with_b;
                                            Console.WriteLine("Product_Desc => " + grpL.Product_Desc);
                                        }
                                        grpL.Product_Desc = (grpL.Product_Desc).Replace(",", DoubleQuote(","));
                                    }
                                    lstGrpL.Add(grpL);
                                }
                                string Item_not_vat = "";
                                string Item_Vat = "";
                                string Item_Sum = "";
                                try
                                {
                                    //if (type_Value_Sum == "Sum_VAT")
                                    //{
                                    //Item_not_vat = ConvertNumber((Double.Parse(lstDataMenu[x + distantAmount]) / 1.07).ToString());
                                    //Item_Vat = ConvertNumber((Double.Parse(lstDataMenu[x + distantAmount]) - Double.Parse(Item_not_vat)).ToString());
                                    //Item_Sum = RemoveComma(ConvertNumber((Double.Parse(lstDataMenu[x + distantAmount])).ToString()));
                                    //Console.WriteLine(type_Value_Sum);
                                    //}
                                    //else if (type_Value_Sum == "Sum_NO_VAT")
                                    //{
                                    //    Item_not_vat = ConvertNumber((Double.Parse(lstDataMenu[x + distantAmount])).ToString());
                                    //    Item_Vat = CalVatItem_ListItem(Double.Parse(lstDataMenu[x + distantAmount]).ToString());
                                    //    Item_Sum = RemoveComma(ConvertNumber((Double.Parse(lstDataMenu[x + distantAmount]) + Double.Parse(Item_Vat)).ToString()));
                                    //    Console.WriteLine(type_Value_Sum);
                                    //}
                                    Item_not_vat = ConvertNumber((Double.Parse(lstDataMenu[x + distantAmount]) / 1.07).ToString());
                                    Item_Vat = ConvertNumber((Double.Parse(lstDataMenu[x + distantAmount]) - Double.Parse(Item_not_vat)).ToString());
                                    Item_Sum = RemoveComma(ConvertNumber((Double.Parse(lstDataMenu[x + distantAmount])).ToString()));
                                    Console.WriteLine(type_Value_Sum);
                                }
                                catch (Exception ex)
                                {
                                    Item_not_vat = ConvertNumber((Double.Parse(lstDataMenu[x + distantAmount]) / 1.07).ToString());
                                    Item_Vat = ConvertNumber((Double.Parse(lstDataMenu[x + distantAmount]) - Double.Parse(Item_not_vat)).ToString());
                                    Item_Sum = RemoveComma(ConvertNumber((Double.Parse(lstDataMenu[x + distantAmount])).ToString()));
                                    Console.WriteLine(type_Value_Sum);
                                }

                                grpL = new LGroup();
                                try
                                {
                                    grpL.Data_Type = DoubleQuote("L"); //ประเภทรายการ
                                    grpL.Line_ID = DoubleQuote(lstDataMenu[x]); //ลำดับรายการ
                                    grpL.Product_ID = DoubleQuote(""); //รหัสสินค้า
                                    grpL.Product_Name = lstDataMenu[x + distantNameItem].Replace(" ", string.Empty).Replace(",", DoubleQuote(",")); //ชื่อสินค้า
                                    grpL.Product_Desc = "";
                                    grpL.Product_Batch_ID = DoubleQuote(""); //ครั้งที่ผลิต
                                    grpL.Product_Expire_Dtm = DoubleQuote(""); //วันหมดอายุ
                                    grpL.Product_Class_Code = DoubleQuote(""); //รหัสหมวดหมู่สินค้า
                                    grpL.Product_Class_Name = DoubleQuote(""); //ชื่อหมวดหมู่สินค้า
                                    grpL.Product_OriCountry_ID = DoubleQuote(""); //รหัสประเทศกำเนิด
                                    try
                                    {
                                        grpL.Product_Charge_Amount = DoubleQuote((Double.Parse(RemoveComma(lstDataMenu[x + distantPriceUnit]))).ToString("0.00")); //ราคาต่อหน่วย
                                    }
                                    catch (FormatException e)
                                    {
                                        grpL.Product_Charge_Amount = DoubleQuote("");
                                    }
                                    grpL.Product_Charge_Curr_Code = DoubleQuote("THB"); //รหัสสกุลเงิน (ราคาต่อหน่วย)
                                    grpL.Product_Al_Charge_IND = DoubleQuote(""); //ตัวบอกส่วนลดหรือค่าธรรมเนียม
                                    if (!discount.Equals(""))
                                    {
                                        try
                                        {
                                            grpL.Product_Al_Actual_Amount = DoubleQuote(Double.Parse(lstDataMenu[x + distantQuanlity]).ToString("0.00")); //มูลค่าส่วนลดหรือค่าธรรมเนียม
                                        }
                                        catch (FormatException e)
                                        {
                                            grpL.Product_Al_Actual_Amount = DoubleQuote("");
                                        }
                                        grpL.Product_Al_Actual_Curr_Code = DoubleQuote("THB"); //รหัสสกุลเงิน (มูลค่าส่วนลดหรือค่าธรรมเนียม)
                                    }
                                    else
                                    {
                                        grpL.Product_Al_Actual_Amount = DoubleQuote(""); //มูลค่าส่วนลดหรือค่าธรรมเนียม
                                        grpL.Product_Al_Actual_Curr_Code = DoubleQuote(""); //รหัสสกุลเงิน (มูลค่าส่วนลดหรือค่าธรรมเนียม)
                                    }
                                    grpL.Product_Al_Reason_Code = DoubleQuote(""); //รหัสเหตุผลในการคิดส่วนลดหรือค่าธรรมเนียม
                                    grpL.Product_Al_Reason = DoubleQuote(""); //เหตุผลในการคิดสวนลดหรือค่าธรรมเนียม

                                    try
                                    {
                                        grpL.Product_Quantity = DoubleQuote(Double.Parse(lstDataMenu[x + distantQuanlity]).ToString("0.00")); //จำนวนสินค้า
                                    }
                                    catch (FormatException e)
                                    {
                                        grpL.Product_Quantity = DoubleQuote("");
                                    }
                                    grpL.Product_Unit_Code = DoubleQuote(""); //รหัสหน่วยสินค้า
                                    grpL.Product_Quan_Per_Unit = DoubleQuote("1"); //ขนาดบรรจุต่อหน่วยขาย
                                    grpL.Line_Tax_Type_Code = DoubleQuote("VAT"); //รหัสประเภทภาษี
                                    grpL.Line_Tax_Cal_Rate = DoubleQuote("7.00"); //อัตราภาษี
                                                                                  //MessageBox.Show(lstDataMenu[x + 6]);
                                    grpL.Line_Basis_Amount = Item_not_vat; //มูลค่าสินค้า/บริการ (ไม่รวมภาษีมูลค่าเพิ่ม)
                                    grpL.Line_Basis_Curr_Code = DoubleQuote("THB"); //รหัสสกุลเงิน (มูลค่าสินค้า/บริการ)
                                    grpL.Line_Tax_Cal_Amount = Item_Vat;//มูลค่าภาษีมูลค่าเพิ่ม
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
                                    grpL.Line_Net_Include_Amount = DoubleQuote(Item_Sum); //จำนวนเงินรวม
                                    strTempSumAmount = string.Empty;
                                    strTempSumAmount = grpL.Line_Basis_Amount;
                                    lstTempSumAmount.Add(strTempSumAmount);
                                    grpL.Line_Net_Include_Curr_Code = DoubleQuote("THB"); //รหัสสกุลเงิน (จำนวนเงินรวม)
                                    grpL.Line_Basis_Amount = DoubleQuote(RemoveComma(grpL.Line_Basis_Amount));
                                    grpL.Line_Tax_Cal_Amount = DoubleQuote(RemoveComma(grpL.Line_Tax_Cal_Amount));
                                    grpL.Product_Remark = ""; //หมายเหตุท้ายสินค้า
                                    iCountRound++;
                                    x = x + (AllChar.Length - 1);
                                }
                                catch (ArgumentOutOfRangeException e)
                                {

                                }
                                catch (NullReferenceException e)
                                {

                                }
                                catch (IndexOutOfRangeException e)
                                {

                                }
                            }
                            else
                            {
                                if (!lstDataMenu[x].Equals(""))
                                {
                                    grpL.Product_Desc += " " + lstDataMenu[x] /*lstDataMenu[x + 1]*/;
                                }
                            }
                        }



                        if (grpL.Product_Desc == null || grpL.Product_Desc.Equals(""))
                        {
                            grpL.Product_Desc = DoubleQuote("");
                        }
                        else
                        {
                            if (grpL.Product_Desc.Length > 256)
                            {
                                string a = grpL.Product_Desc.Substring(0, 256);
                                Console.WriteLine("a => " + a);
                                string[] b = a.Split(' ');
                                Console.WriteLine(b.Length);
                                for (int i = 0; i < b.Length - 1; i++)
                                {
                                    a_with_b += b[i] + " ";
                                }
                                Console.WriteLine("a_With_b => " + a_with_b);
                                Console.WriteLine("LengthOfa_with_b => " + a_with_b.Length);
                                grpL.Product_Remark = DoubleQuote(grpL.Product_Desc.Substring(a_with_b.Length));
                                Console.WriteLine("Product_Remark => " + grpL.Product_Remark);
                                grpL.Product_Desc = a_with_b;
                            }
                            grpL.Product_Desc = (grpL.Product_Desc).Replace(",", DoubleQuote(","));
                        }

                        lstGrpL.Add(grpL);

                        form.pgbLoad.Value = 50;
                        form.OutputPrc(50, "Export Data: 50%");

                        ////Type F
                        form.pgbLoad.Value = 60;
                        form.OutputPrc(60, "Export Data: 60%");
                        ////Type H
                        //string[] arrKey = new string[] { "เลขที่ใบสั่งซื้อ :", "วันที่ใบสั่งซื้อ :" };
                        //int[] arrIndex = new int[2];
                        //int countArr = 0;
                        try
                        {
                            try
                            {
                                if (!pfIO.TypeDoc.Equals(""))
                                {
                                    grpH.Doc_Type_Code = pfIO.TypeDoc;
                                }
                                else
                                {

                                    if (!typedoc.Equals(""))
                                    {

                                        grpH.Doc_Type_Code = instring(DocType_ENG_AND_CODE, typedoc.Replace(" ", string.Empty));
                                    }
                                    else
                                    {
                                        Console.WriteLine(typedoc + " typedoc");
                                        Console.WriteLine(getvalue(document_name) + " getvalue(document_name)");
                                        grpH.Doc_Type_Code = instring(DocType, getvalue(document_name).Replace(" ", string.Empty));
                                    }
                                }


                            }
                            catch (KeyNotFoundException e)
                            {
                                Console.WriteLine("ไม่พบชื่อตัวแปรที่ส่งมา จึงเกิด error ");
                            }
                            Console.WriteLine(getvalue(document_name).Replace(" ", string.Empty) + " getvalue(document_name).Replace");
                            grpH.Doc_Name = getvalue(document_name).Replace(" ", string.Empty);
                            grpH.Doc_ID = getvalue(document_id).Replace(" ", string.Empty);
                            if (DateTime.TryParse(getvalue(document_issue_dtm), out dateValue))
                            {
                                string[] arrDate = getvalue(document_issue_dtm).Split('/');
                                string year = DateTime.Now.Year.ToString();
                                string yearFront = strYear.Substring(0, 2);
                                grpH.Doc_Issue_Dtm = years + "-" + Month[arrDateSplit[1].Replace(" ", string.Empty)] + "-" + Day[arrDateSplit[0].Replace(" ", string.Empty)] + "T00:00:00";
                            }
                            else
                            {
                                grpH.Doc_Issue_Dtm = getvalue(document_issue_dtm);
                            }

                        }
                        catch (ArgumentException e)
                        {
                            Console.WriteLine(e.Message);
                        }
                        catch (NullReferenceException e)
                        {
                            Console.WriteLine(e.Message);
                        }
                        catch (IndexOutOfRangeException e)
                        {
                            Console.WriteLine(e.Message);
                        }
                        catch (KeyNotFoundException e)
                        {
                            Console.WriteLine(e.Message + "=>" + "KeyNotFoundException");
                        }
                        if (!additional_ref_assign_id.Equals(""))
                        {
                            grpH.Add_Ref_Assign_ID = getvalue(additional_ref_assign_id).Replace(" ", string.Empty);
                        }
                        else
                        {
                            grpH.Add_Ref_Assign_ID = "";

                        }
                        if (!additional_ref_issue_dtm.Equals(""))
                        {
                            try
                            {
                                if (DateTime.TryParse(getvalue(document_issue_dtm), out dateValue))
                                {
                                    arrDateSplit = getvalue(additional_ref_issue_dtm).Split('/');
                                    grpH.Add_Ref_Issue_Dtm = years + "-" + Month[arrDateSplit[1].Replace(" ", string.Empty)] + "-" + Day[arrDateSplit[0].Replace(" ", string.Empty)] + "T00:00:00";
                                }
                                else
                                {
                                    grpH.Add_Ref_Issue_Dtm = getvalue(additional_ref_issue_dtm);
                                }
                            }
                            catch (ArgumentException e)
                            {
                                Console.WriteLine(e.Message);
                            }
                            catch (NullReferenceException e)
                            {
                                Console.WriteLine(e.Message);
                            }
                            catch (IndexOutOfRangeException e)
                            {
                                Console.WriteLine(e.Message);
                            }
                            catch (KeyNotFoundException e)
                            {
                                Console.WriteLine("aasdsasd");
                            }
                        }
                        else
                        {
                            grpH.Add_Ref_Issue_Dtm = "";
                        }

                        //MessageBox.Show(grpH.Add_Ref_Assign_ID);
                        if (!grpH.Add_Ref_Assign_ID.Equals(""))
                        {
                            grpH.Add_Ref_Type_Code = grpH.Doc_Type_Code;
                        }
                        else
                        {
                            grpH.Add_Ref_Type_Code = "";
                        }
                        try
                        {
                            if (create_purpose_code.Equals(""))
                            {
                                switch (grpH.Doc_Type_Code)
                                {
                                    case "388":
                                        try
                                        {
                                            grpH.Create_Purpose_Code = instring(TIVCPurpose, getvalue(create_purpose).Replace(" ", string.Empty));
                                            if (grpH.Create_Purpose_Code.Substring(grpH.Create_Purpose_Code.Length - 2) == "99")
                                            {
                                                grpH.Create_Purpose = getvalue(create_purpose).Replace(" ", string.Empty);
                                            }
                                        }
                                        catch (KeyNotFoundException ex)
                                        {
                                            grpH.Create_Purpose_Code = "TIVC99";
                                            grpH.Create_Purpose = getvalue(create_purpose).Replace(" ", string.Empty);
                                        }
                                        break;
                                    case "T02":
                                        try
                                        {
                                            grpH.Create_Purpose_Code = instring(TIVCPurpose, getvalue(create_purpose).Replace(" ", string.Empty));
                                            if (grpH.Create_Purpose_Code.Substring(grpH.Create_Purpose_Code.Length - 2) == "99")
                                            {
                                                grpH.Create_Purpose = getvalue(create_purpose).Replace(" ", string.Empty);
                                            }
                                        }
                                        catch (KeyNotFoundException ex)
                                        {
                                            grpH.Create_Purpose_Code = "TIVC99";
                                            grpH.Create_Purpose = getvalue(create_purpose).Replace(" ", string.Empty);
                                        }
                                        break;
                                    case "T03":
                                        try
                                        {
                                            grpH.Create_Purpose_Code = instring(TIVCPurpose, getvalue(create_purpose).Replace(" ", string.Empty));
                                            if (grpH.Create_Purpose_Code.Substring(grpH.Create_Purpose_Code.Length - 2) == "99")
                                            {
                                                grpH.Create_Purpose = getvalue(create_purpose).Replace(" ", string.Empty);
                                            }
                                        }
                                        catch (KeyNotFoundException ex)
                                        {
                                            grpH.Create_Purpose_Code = "TIVC99";
                                            grpH.Create_Purpose = getvalue(create_purpose).Replace(" ", string.Empty);
                                        }
                                        break;
                                    case "T04":
                                        try
                                        {
                                            grpH.Create_Purpose_Code = instring(TIVCPurpose, getvalue(create_purpose).Replace(" ", string.Empty));
                                            if (grpH.Create_Purpose_Code.Substring(grpH.Create_Purpose_Code.Length - 2) == "99")
                                            {
                                                grpH.Create_Purpose = getvalue(create_purpose).Replace(" ", string.Empty);
                                            }
                                        }
                                        catch (KeyNotFoundException ex)
                                        {
                                            grpH.Create_Purpose_Code = "TIVC99";
                                            grpH.Create_Purpose = getvalue(create_purpose).Replace(" ", string.Empty);
                                        }
                                        break;
                                    case "T01":
                                        try
                                        {
                                            grpH.Create_Purpose_Code = instring(RCTCPurpose, getvalue(create_purpose).Replace(" ", string.Empty));
                                            if (grpH.Create_Purpose_Code.Substring(grpH.Create_Purpose_Code.Length - 2) == "99")
                                            {
                                                grpH.Create_Purpose = getvalue(create_purpose).Replace(" ", string.Empty);
                                            }

                                        }
                                        catch (KeyNotFoundException ex)
                                        {
                                            grpH.Create_Purpose_Code = "RCTC99";
                                            grpH.Create_Purpose = getvalue(create_purpose).Replace(" ", string.Empty);
                                        }
                                        break;
                                    case "80":
                                        try
                                        {
                                            grpH.Create_Purpose_Code = instring(DBNGPurpose, getvalue(create_purpose).Replace(" ", string.Empty));
                                            if (grpH.Create_Purpose_Code.Substring(grpH.Create_Purpose_Code.Length - 2) == "99")
                                            {
                                                grpH.Create_Purpose = getvalue(create_purpose).Replace(" ", string.Empty);
                                            }
                                        }
                                        catch (KeyNotFoundException ex)
                                        {
                                            grpH.Create_Purpose_Code = "DBNG99";
                                            grpH.Create_Purpose = getvalue(create_purpose).Replace(" ", string.Empty);
                                        }
                                        break;
                                    case "81":
                                        try
                                        {
                                            grpH.Create_Purpose_Code = instring(CDNGPurpose, getvalue(create_purpose).Replace(" ", string.Empty));
                                            if (grpH.Create_Purpose_Code.Substring(grpH.Create_Purpose_Code.Length - 2) == "99")
                                            {
                                                grpH.Create_Purpose = getvalue(create_purpose).Replace(" ", string.Empty);
                                            }
                                        }
                                        catch (KeyNotFoundException ex)
                                        {
                                            grpH.Create_Purpose_Code = "CDNG99";
                                            grpH.Create_Purpose = getvalue(create_purpose).Replace(" ", string.Empty);
                                        }
                                        break;
                                    default:
                                        grpH.Create_Purpose_Code = "";
                                        grpH.Create_Purpose = "";
                                        break;
                                }
                            }
                            else
                            {
                                grpH.Create_Purpose_Code = getvalue(create_purpose_code);
                                grpH.Create_Purpose = getvalue(create_purpose);
                            }
                        }
                        catch (NullReferenceException e)
                        {
                            Console.WriteLine(e.Message);
                        }
                        catch (IndexOutOfRangeException e)
                        {
                            grpH.Create_Purpose_Code = "";
                            grpH.Create_Purpose = "";
                        }
                        catch (KeyNotFoundException e)
                        {
                            grpH.Create_Purpose_Code = "";
                            grpH.Create_Purpose = "";
                        }
                        Console.WriteLine(counext(document_remark, (totel_Rows_Footer + int.Parse(document_remark.Split(',')[0].Split('-')[1]))));
                        Console.WriteLine("เงินค่ามัดจำ : " + getvalue(counext(amount_deposit, (totel_Rows_Footer + int.Parse(amount_deposit.Split(',')[0].Split('-')[1])))));
                        if (!document_remark.Equals(""))
                        {
                            grpH.DOCUMENT_REMARK = getvalue(counext(document_remark, (totel_Rows_Footer + int.Parse(document_remark.Split(',')[0].Split('-')[1]))));
                            if (getvalue(counext(amount_deposit, (totel_Rows_Footer + int.Parse(amount_deposit.Split(',')[0].Split('-')[1])))) != "0")
                            {
                                grpH.DOCUMENT_REMARK = grpH.DOCUMENT_REMARK + " หักเงินค่ามัดจำอยู่ในฟิลหักเงินค่าส่วนลดในไฟล์ TEXT";
                            }

                        }
                        else
                        {
                            grpH.DOCUMENT_REMARK = "";
                            if (getvalue(counext(amount_deposit, (totel_Rows_Footer + int.Parse(amount_deposit.Split(',')[0].Split('-')[1])))) != "0")
                            {
                                grpH.DOCUMENT_REMARK = grpH.DOCUMENT_REMARK + " หักเงินค่ามัดจำอยู่ในฟิลหักเงินค่าส่วนลดในไฟล์ TEXT";
                            }
                        }


                        if (!buyer_order_assign_id.Equals(""))
                        {
                            //grpH.Buyer_Order_Assign_ID = getvalue(additional_ref_assign_id).Replace(" ", string.Empty);
                            grpH.Buyer_Order_Assign_ID = getvalue(buyer_order_assign_id).Replace(" ", string.Empty);
                        }
                        else
                        {
                            grpH.Buyer_Order_Assign_ID = "";
                        }


                        if (!buyer_order_issue_dtm.Equals(""))
                        {
                            try
                            {
                                arrDateSplit = getvalue(buyer_order_issue_dtm).Split('/');
                                grpH.Buyer_Order_Issue_Dtm = years + "-" + Month[arrDateSplit[1].Replace(" ", string.Empty)] + "-" + Day[arrDateSplit[0].Replace(" ", string.Empty)] + "T00:00:00";
                            }
                            catch (ArgumentException e)
                            {
                                Console.WriteLine(e.Message);
                            }
                            catch (NullReferenceException e)
                            {
                                Console.WriteLine(e.Message);
                            }
                            catch (IndexOutOfRangeException e)
                            {
                                Console.WriteLine(e.Message);
                            }
                            catch (KeyNotFoundException e)
                            {
                                Console.WriteLine(e.Message);
                            }
                        }
                        else
                        {
                            grpH.Buyer_Order_Issue_Dtm = "";
                        }


                        if (grpH.Buyer_Order_Assign_ID.Equals(""))
                        {
                            grpH.Buyer_Order_Ref_Type_Code = "";
                        }
                        else
                        {
                            grpH.Buyer_Order_Ref_Type_Code = "ON";
                        }

                        try
                        {
                            if (!original_total_amount.Equals(""))
                            {

                                grpF.Original_Total_Amount = Double.Parse(getvalue(counext(totalamount, totel_Rows_Footer - int.Parse(totalamount.Split(',')[0].Split('-')[1])))).ToString("0.00");
                                grpF.Original_Total_Curr_Code = "THB";
                            }
                            else
                            {
                                grpF.Original_Total_Amount = "";
                                grpF.Original_Total_Curr_Code = "";
                            }
                        }
                        catch (FormatException e)
                        {
                            Console.WriteLine(e.Message);
                        }
                        catch (NullReferenceException e)
                        {
                            Console.WriteLine(e.Message);
                        }

                        try
                        {
                            if (!line_total_amount.Equals(""))
                            {
                                grpF.LINE_TOTAL_AMOUNT = Double.Parse(getvalue(line_total_amount)).ToString("0.00");
                                grpF.LINE_TOTAL_CURRENCY_CODE = "THB";
                            }
                            else
                            {
                                grpF.LINE_TOTAL_AMOUNT = "";
                                grpF.LINE_TOTAL_CURRENCY_CODE = "";
                            }
                        }
                        catch (FormatException e)
                        {
                            Console.WriteLine(e.Message);
                        }
                        catch (NullReferenceException e)
                        {
                            Console.WriteLine(e.Message);
                        }

                        try
                        {
                            if (!adjusted_information_amount.Equals(""))
                            {
                                grpF.Adjusted_Inform_Amount = Double.Parse(getvalue(adjusted_information_amount)).ToString("0.00");
                                grpF.Adjusted_Inform_Curr_Code = "THB";
                            }
                            else
                            {
                                grpF.Adjusted_Inform_Amount = "";
                                grpF.Adjusted_Inform_Curr_Code = "";
                            }
                        }
                        catch (FormatException e)
                        {
                            Console.WriteLine(e.Message);
                        }
                        catch (NullReferenceException e)
                        {
                            Console.WriteLine(e.Message);
                        }
                        try
                        {
                            if (!allowance_total_amount.Equals("") || !amount_deposit.Equals(""))
                            {
                                grpF.Al_Total_Amount = RemoveComma(ConvertNumber((float.Parse(getvalue(counext(allowance_total_amount, (totel_Rows_Footer + int.Parse(allowance_total_amount.Split(',')[0].Split('-')[1]))))) + float.Parse(getvalue(counext(amount_deposit, (totel_Rows_Footer + int.Parse(amount_deposit.Split(',')[0].Split('-')[1])))))).ToString()));
                                grpF.Al_Total_Curr_Code = "THB";
                            }
                            else
                            {
                                grpF.Al_Total_Amount = "";
                                grpF.Al_Total_Curr_Code = "";
                            }
                        }
                        catch (FormatException e)
                        {
                            Console.WriteLine(e.Message);
                        }
                        catch (NullReferenceException e)
                        {
                            Console.WriteLine(e.Message);
                        }
                        if (!countrybuyer.Equals(""))
                        {
                            if (!getvalue(countrybuyer).Equals(""))
                            {
                                grpB.Buyer_Country_ID = getvalue(countrybuyer).Replace(" ", string.Empty).Replace("-", string.Empty);
                                if (grpB.Buyer_Country_ID != "TH")
                                {
                                    grpB.Buyer_Post_Code = "";
                                }
                            }
                            else
                            {
                                grpB.Buyer_Country_ID = "TH";
                            }
                        }
                        else
                        {
                            grpB.Buyer_Country_ID = "TH";
                        }
                        List<string> lstC = new List<string> { DoubleQuote("C"),
                                                    DoubleQuote(grpC.Seller_Tax_ID.Replace(" ",string.Empty)), //เลขที่ประจำตัวผู้เสียภาษี
                                                    DoubleQuote(grpC.Seller_Branch_ID), //เลขสาขาประกอบการ
                                                    DoubleQuote(grpC.File_Name.Replace(" ",string.Empty)), //ชื่อไฟล์  
                                                    };
                        ////MessageBox.Show("lstC:Success");
                        List<string> lstH = new List<string> { DoubleQuote("H"),
                                                    DoubleQuote(grpH.Doc_Type_Code), //ประเภทเอกสาร 
                                                    DoubleQuote(grpH.Doc_Name), //ชื่อเอกสาร
                                                    DoubleQuote(grpH.Doc_ID), // เลขที่เอกสาร
                                                    DoubleQuote(grpH.Doc_Issue_Dtm), //วันที่
                                                    DoubleQuote(grpH.Create_Purpose_Code), //สาเหตุการออกเอกสาร
                                                    DoubleQuote(grpH.Create_Purpose), //กรณีระบุสาเหตุเอกสาร
                                                    DoubleQuote(grpH.Add_Ref_Assign_ID), //เลขที่เอกสารอ้างอิง
                                                    DoubleQuote(grpH.Add_Ref_Issue_Dtm), //เอกสารอ้างอิงลงวันที่
                                                    DoubleQuote(grpH.Add_Ref_Type_Code), //ประเภทเอกสารอ้างอิง
                                                    DoubleQuote(""), //ชื่อเอกสารอ้างอิง 
                                                    DoubleQuote(""), //เงื่อนไขการส่งของ
                                                    DoubleQuote(grpH.Buyer_Order_Assign_ID), //เลขที่ใบสั่งซื้อ
                                                    DoubleQuote(grpH.Buyer_Order_Issue_Dtm), //วันเดือนปีที่ออกใบสั่งซื้อ
                                                    DoubleQuote(grpH.Buyer_Order_Ref_Type_Code), //ประเภทเอกสารอ้างอิงการสั่งซื้อ
                                                    DoubleQuote(grpH.DOCUMENT_REMARK) //หมายเหตุท้ายเอกสาร
                                                    };
                        form.pgbLoad.Value = 70;
                        form.OutputPrc(70, "Export Data: 70%");

                        List<string> lstB = new List<string> { DoubleQuote("B"),
                                                    DoubleQuote(""), //รหัสผู้ซื้อ
                                                    DoubleQuote(grpB.Buyer_Name), //ชื่อผู้ซื้อ
                                                    DoubleQuote(grpB.Buyer_Tax_ID_Type), //ประเภทผู้เสียภาษี
                                                    DoubleQuote(grpB.Buyer_Tax_ID.Replace(" ",string.Empty)), //เลขประจำตัวผู้เสียภาษี
                                                    DoubleQuote(grpB.Buyer_Branch_ID.Replace(" ",string.Empty)), //เลขที่สาขา
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
                                                    DoubleQuote(grpB.Buyer_Country_ID) //รหัสประเทศ
                                                    };

                        try
                        {

                            if (!totalamount.Equals(""))
                            {
                                totalamount = RemoveComma(double.Parse(getvalue(counext(totalamount, (totel_Rows_Footer + int.Parse(totalamount.Split(',')[0].Split('-')[1]))))).ToString("0.00"));
                            }
                            else
                            {
                                totalamount = "";
                            }
                        }
                        catch (IndexOutOfRangeException e)
                        {
                            Console.WriteLine(e.Message);
                        }
                        catch (NullReferenceException e)
                        {
                            Console.WriteLine(e.Message);
                        }

                        try
                        {
                            if (!vat.Equals(""))
                            {
                                vat = RemoveComma(double.Parse(getvalue(counext(vat, (totel_Rows_Footer + int.Parse(vat.Split(',')[0].Split('-')[1]))))).ToString("0.00"));
                            }
                            else
                            {
                                vat = "";
                            }
                        }
                        catch (IndexOutOfRangeException e)
                        {
                            Console.WriteLine(e.Message);
                        }
                        catch (NullReferenceException e)
                        {
                            Console.WriteLine(e.Message);
                        }

                        try
                        {
                            if (!tax_basis_total_amount.Equals(""))
                            {
                                tax_basis_total_amount = RemoveComma(double.Parse(getvalue(counext(tax_basis_total_amount, (totel_Rows_Footer + int.Parse(tax_basis_total_amount.Split(',')[0].Split('-')[1]))))).ToString("0.00"));
                            }
                            else
                            {
                                tax_basis_total_amount = "";
                            }
                        }
                        catch (IndexOutOfRangeException e)
                        {
                            Console.WriteLine(e.Message);
                        }
                        catch (NullReferenceException e)
                        {
                            Console.WriteLine(e.Message);
                        }

                        try
                        {
                            if (!total.Equals(""))
                            {
                                total = RemoveComma(double.Parse(getvalue(counext(total, (totel_Rows_Footer + int.Parse(total.Split(',')[0].Split('-')[1]))))).ToString("0.00"));
                            }
                            else
                            {
                                total = "";
                            }
                        }
                        catch (IndexOutOfRangeException e)
                        {
                            Console.WriteLine(e.Message);
                        }
                        catch (NullReferenceException e)
                        {
                            Console.WriteLine(e.Message);
                        }

                        try
                        {

                            vat_rate = Regex.Match(RemoveComma(getvalue(counext(vat_rate, (totel_Rows_Footer + int.Parse(vat_rate.Split(',')[0].Split('-')[1]))))), @"\d+").Value + ".00";
                            if (vat_rate != "0.00")
                            {
                                vat = RemoveComma(ConvertNumber(getvalue(counext(vat, (totel_Rows_Footer + int.Parse(vat.Split(',')[0].Split('-')[1])))).ToString()));
                            }
                            else
                            {
                                vat = "0.00";
                            }
                        }
                        catch (Exception ex)
                        {
                            vat_rate = "7.00";
                        }
                        try
                        {
                            if (getvalue(counext(amount_deposit, (totel_Rows_Footer + int.Parse(amount_deposit.Split(',')[0].Split('-')[1])))) != "0")
                            {
                                grpF.Al_Charge_IND = "true";
                                grpF.Al_Actual_Amount = RemoveComma(ConvertNumber(getvalue(counext(amount_deposit, (totel_Rows_Footer + int.Parse(amount_deposit.Split(',')[0].Split('-')[1]))))));
                                grpF.Al_Actual_Curr_Code = "THB";
                            }
                            else
                            {
                                grpF.Al_Charge_IND = "";
                            }
                        }
                        catch (Exception ex)
                        {
                            grpF.Al_Charge_IND = "";
                        }
                        List<string> lstF = new List<string> { DoubleQuote("F"),
                                                        DoubleQuote(String.Format("{0:0}", lstGrpL.Count).ToString()), //จำนวนรายการสินค้า
                                                        DoubleQuote(""), //วันเวลานัดส่งสินค้า
                                                        DoubleQuote("THB"), //รหัสสกุลเงินตรา
                                                        DoubleQuote("VAT"), //รหัสประเภทภาษี
                                                        DoubleQuote(vat_rate), //อัตราภาษี
                                                        //DoubleQuote(RemoveComma(sumAmount.ToString("N2"))), //มูลค่าสินค้า(ไม่รวมภาษีมูลค่าเพิ่ม)2350
                                                        DoubleQuote(totalamount),
                                                        DoubleQuote("THB"),
                                                        //DoubleQuote(RemoveComma(sumTaxAmount.ToString("N2"))), //มูลค่าภาษีมูลค่าเพิ่ม
                                                        DoubleQuote(vat),
                                                        DoubleQuote("THB"),
                                                        DoubleQuote(grpF.Al_Charge_IND), //ตัวบอกส่วนลดหรือค่าธรรมเนียม
                                                        DoubleQuote(grpF.Al_Actual_Amount), //มูลค่าส่วนลดหรือค่าธรรมเนียม
                                                        DoubleQuote(grpF.Al_Actual_Curr_Code), //มูลค่าส่วนลดหรือค่าธรรมเนียม
                                                        DoubleQuote(""), //รหัสเหตุผลในการคิดส่วนลดหรือค่าธรรมเนียม
                                                        DoubleQuote(""), //เหตุผลในการคิดส่วนลดหรือค่าธรรมเนียม
                                                        DoubleQuote(""), //รหัสประเภทส่วนลด     
                                                        DoubleQuote(""), //รายละเอียดเงื่อนไขการชำระเงิน
                                                        DoubleQuote(""), //วันครบกำหนดชำระเงิน
                                                        DoubleQuote(grpF.Original_Total_Amount), //รวมมูลค่าตามเอกสารเดิม
                                                        DoubleQuote(grpF.Original_Total_Curr_Code),
                                                        //DoubleQuote(RemoveComma(sumAmount.ToString("N2"))),
                                                        DoubleQuote(totalamount),
                                                        DoubleQuote("THB"),
                                                        DoubleQuote(grpF.Adjusted_Inform_Amount), //มูลค่าผลต่าง
                                                        DoubleQuote(grpF.Adjusted_Inform_Curr_Code),
                                                        DoubleQuote(grpF.Al_Total_Amount), //ส่วนลดทั้งหมด
                                                        DoubleQuote(grpF.Al_Total_Curr_Code),
                                                        DoubleQuote(""), //ค่าธรรมเนียมทั้งหมด
                                                        DoubleQuote(""),
                                                        //DoubleQuote(RemoveComma(sumAmount.ToString("N2"))), //มูลค่าที่นำมาคิดภาษีมูลค่าเพิ่ม
                                                        DoubleQuote(tax_basis_total_amount),
                                                        DoubleQuote("THB"),
                                                        //DoubleQuote(RemoveComma(sumTaxAmount.ToString("N2"))), //จำนวนภาษีมูลค่าเพิ่ม
                                                        DoubleQuote(vat),
                                                        DoubleQuote("THB"),
                                                        //DoubleQuote(RemoveComma(sumGrandTotal.ToString("N2"))), //จำนวนเงินรวม(รวมภาษีมูลค่าเพิ่ม)
                                                        DoubleQuote(total),
                                                        DoubleQuote("THB")
                                                        };

                        List<string> lstT = new List<string> { DoubleQuote("T"),
                                                        DoubleQuote("1") //จำนวนเอกสารทั้งหมด
                                                        };
                        form.pgbLoad.Value = 80;
                        form.OutputPrc(80, "Export Data: 80%");
                        Console.WriteLine("a");
                        string messageText = String.Join(",", lstC) + "\r" + String.Join(",", lstH) + "\r" + String.Join(",", lstB) + "\r";

                        for (int k = 0; k < lstGrpL.Count; k++)
                        {
                            messageText += lstGrpL[k].Data_Type + "," + lstGrpL[k].Line_ID + "," + lstGrpL[k].Product_ID + "," + lstGrpL[k].Product_Name + "," + lstGrpL[k].Product_Desc + ","
                                + lstGrpL[k].Product_Batch_ID + "," + lstGrpL[k].Product_Expire_Dtm + "," + lstGrpL[k].Product_Class_Code + "," + lstGrpL[k].Product_Class_Name + "," + lstGrpL[k].Product_OriCountry_ID + ","
                                + lstGrpL[k].Product_Charge_Amount + "," + lstGrpL[k].Product_Charge_Curr_Code + "," + lstGrpL[k].Product_Al_Charge_IND + "," + lstGrpL[k].Product_Al_Actual_Amount + "," + lstGrpL[k].Product_Al_Actual_Curr_Code + ","
                                + lstGrpL[k].Product_Al_Reason_Code + "," + lstGrpL[k].Product_Al_Reason + "," + lstGrpL[k].Product_Quantity + "," + lstGrpL[k].Product_Unit_Code + "," + lstGrpL[k].Product_Quan_Per_Unit + ","
                                + lstGrpL[k].Line_Tax_Type_Code + "," + lstGrpL[k].Line_Tax_Cal_Rate + "," + lstGrpL[k].Line_Basis_Amount + "," + lstGrpL[k].Line_Basis_Curr_Code + "," + lstGrpL[k].Line_Tax_Cal_Amount + ","
                                + lstGrpL[k].Line_Tax_Cal_Curr_Code + "," + lstGrpL[k].Line_AL_Charge_IND + "," + lstGrpL[k].Line_AL_Actual_Amount + "," + lstGrpL[k].Line_AL_Actual_Curr_Code + "," + lstGrpL[k].Line_AL_Reason_Code + ","
                                + lstGrpL[k].Line_AL_Reason + "," + lstGrpL[k].Line_Tax_Total_Amount + "," + lstGrpL[k].Line_Tax_Total_Curr_Code + "," + lstGrpL[k].Line_Net_Total_Amount + "," + lstGrpL[k].Line_Net_Total_Curr_Code + ","
                                + lstGrpL[k].Line_Net_Include_Amount + "," + lstGrpL[k].Line_Net_Include_Curr_Code + "," + lstGrpL[k].Product_Remark + "\r";
                        }
                        messageText += String.Join(",", lstF) + "\r" + String.Join(",", lstT);
                        pathText = dtParam.PathOutput + "\\" + UATorPROD + "_" + strFileName + "_" + strDateTimeStamp + ".txt";
                        CreateTextFile(pathText, messageText);
                        form.txtStatus.Text += Environment.NewLine + "   -Convert Success!";
                        strTempLogTime += " Convert Success!";
                        //txtStatus.Refresh();
                        form.Outputmessage(txtstr);
                        System.Threading.Thread.Sleep(500);
                        if (dtParam.ServiceCode == "S06")
                        {
                            strOutputFile = conAPIClass.CallAPI(dtParam, pathText, strFileNamePDF);
                        }
                        else if (dtParam.ServiceCode == "S03")
                        {
                            strOutputFile = conAPIClass.CallAPI(dtParam, pathText, "");
                        }


                    }
                    catch (FileNotFoundException e)
                    {
                        MessageBox.Show("File Not Found ConfigExcel => " + dtParam.PathConfigExcel);
                        goto loop;
                    }
                }


                else if (Path.GetExtension(dtParam.PathInput).Equals(".csv"))
                {
                    form.pgbLoad.Value = 0;
                    form.OutputPrc(0, "Export Data: 0%");

                    using (var reader = new StreamReader(dtParam.PathInput))
                    {
                        List<string> listA = new List<string>();
                        List<string> listB = new List<string>();
                        while (!reader.EndOfStream)
                        {
                            var line = reader.ReadLine();
                            var values = line.Split(';');
                            listA.Add(values[0]);
                            //listB.Add(values[1]);
                        }
                        string[] arrSplit = listA[1].Split(',');
                        string strID = arrSplit[3].Replace("\"", string.Empty);
                        string[] arrDate = Regex.Split(arrSplit[4], "T");
                        string strDate = arrDate[0].Replace("\"", string.Empty);
                        //MessageBox.Show(strDate);
                        if (form.txtStatus.Text != null && !form.txtStatus.Text.Equals(""))
                        {
                            form.txtStatus.Text += Environment.NewLine + "เลขที่เอกสาร " + strID + " วันที่ " + strDate + " ชื่อไฟล์ " + strFileNameExtension + ":";
                            strTempLogTime += Environment.NewLine + "เลขที่เอกสาร " + strID + " วันที่ " + strDate + " ชื่อไฟล์ " + strFileNameExtension + ":";
                        }
                        else
                        {
                            form.txtStatus.Text = "เลขที่เอกสาร " + strID + " วันที่ " + strDate + " ชื่อไฟล์ " + strFileNameExtension + " :";
                            strTempLogTime = "เลขที่เอกสาร " + strID + " วันที่ " + strDate + " ชื่อไฟล์ " + strFileNameExtension + " :";
                        }
                    }
                    System.Threading.Thread.Sleep(500);
                    form.pgbLoad.Value = 50;
                    form.OutputPrc(50, "Export Data: 50%");
                    //MessageBox.Show(dtParam.PathInput);
                    System.Threading.Thread.Sleep(500);

                    strOutputFile = conAPIClass.CallAPI(dtParam, dtParam.PathInput, strFileNamePDF);
                    pathText = UATorPROD + "_" + Path.GetFileNameWithoutExtension(dtParam.PathInput) + "_" + strDateTimeStamp + ".txt";
                }
                else
                {
                    form.pgbLoad.Value = 0;
                    form.OutputPrc(0, "Export Data: 0%");

                    string strText = System.IO.File.ReadAllText(dtParam.PathInput);
                    string[] arrSplit = strText.Split(',');
                    string strID = arrSplit[6].Replace("\"", string.Empty);
                    string[] arrDate = Regex.Split(arrSplit[7], "T");
                    string strDate = arrDate[0].Replace("\"", string.Empty);
                    //MessageBox.Show(strDate);
                    if (form.txtStatus.Text != null && !form.txtStatus.Text.Equals(""))
                    {
                        form.txtStatus.Text += Environment.NewLine + "เลขที่เอกสาร " + strID + " วันที่ " + strDate + " ชื่อไฟล์ " + strFileNameExtension + ":";
                        strTempLogTime += Environment.NewLine + "เลขที่เอกสาร " + strID + " วันที่ " + strDate + " ชื่อไฟล์ " + strFileNameExtension + ":";
                    }
                    else
                    {
                        form.txtStatus.Text = "เลขที่เอกสาร " + strID + " วันที่ " + strDate + " ชื่อไฟล์ " + strFileNameExtension + " :";
                        strTempLogTime = "เลขที่เอกสาร " + strID + " วันที่ " + strDate + " ชื่อไฟล์ " + strFileNameExtension + " :";
                    }
                    form.pgbLoad.Value = 50;
                    form.OutputPrc(50, "Export Data: 50%");

                    System.Threading.Thread.Sleep(1);
                    strOutputFile = conAPIClass.CallAPI(dtParam, dtParam.PathInput, strFileNamePDF);
                    pathText = UATorPROD + "_" + Path.GetFileNameWithoutExtension(dtParam.PathInput) + "_" + strDateTimeStamp + ".txt";
                    Console.WriteLine(strOutputFile.Message_Content + " Console.WriteLine(strOutputFile.MessageError)");

                }
                form.pgbLoad.Value = 90;
                form.OutputPrc(90, "Export Data: 90%");
                FileVersionInfo fvi = FileVersionInfo.GetVersionInfo(assembly.Location);
                string version = fvi.FileVersion;
                if (strOutputFile.MessageResultError != null && !strOutputFile.MessageResultError.Equals(""))
                //if (strOutputFile.MessageResultError == "")
                {
                    JObject oKeepResponeExecute = new JObject();
                    //MessageBox.Show(strOutputFile.MessageResultError);

                    if (strOutputFile.MessageResultError == "{}")
                    {
                        strOutputFile.MessageResultError = "กรุณาตรวจสอบอินเตอร์เน็ต!!";
                    }
                    form.txtStatus.Text += Environment.NewLine + "   -**********etax.one.th Fail!" + " (" + strOutputFile.MessageLogTime + ")" + "**********";
                    strTempLogTime += " etax.one.th Fail!" + " (" + strOutputFile.MessageLogTime + ") Version " + version;
                    //Console.WriteLine(pathIn + " " + pathOutput);
                    oKeepResponeExecute = JObject.Parse(strOutputFile.MessageResultError.ToString());

                    //sock.SendMailAlert(dtParam.PathInput, dtParam.PathOutput, "FE99", form.emailtxt.Text, form.txtSellerTaxID.Text, Path.GetFileName(dtParam.PathInput).Replace("~$", string.Empty) + "-//-" + oKeepResponeExecute["errorCode"], oKeepResponeExecute["errorMessage"].ToString().Replace(" ",string.Empty));
                    //sock.SendMailAlert(pathFileIO.PathInput, pathFileIO.PathErr, "F01", form.emailtxt.Text, form.txtSellerTaxID.Text, pathIn + Path.GetFileName(item).Replace("~$", string.Empty), pathIn + Path.GetFileName(item).Replace("~$", string.Empty));
                    cntFail++;
                }
                else
                {

                    form.txtStatus.Text += Environment.NewLine + "   -etax.one.th Success!" + " (" + strOutputFile.MessageLogTime + ")";
                    strTempLogTime += ", etax.one.th Success!" + " (" + strOutputFile.MessageLogTime + ") Version " + version;
                }


                //txtStatus.Refresh();
                if (chkOption == true)
                {
                    if (strOutputFile.StatusCallAPI == false)
                    {

                        string pathErr = dtParam.PathOutput + "\\" + Path.GetFileNameWithoutExtension(pathText) + "_Error.txt";
                        JObject oKeepResponeExecute = new JObject();
                        oKeepResponeExecute = JObject.Parse(strOutputFile.MessageResultError.ToString());
                        Console.WriteLine(strOutputFile + " oKeepResponeExecute");
                        _apimail.err_code = "FE99";
                        _apimail.actionmsg = oKeepResponeExecute["errorMessage"].ToString().Replace(" ", string.Empty).Replace("\n", string.Empty).Replace(",", string.Empty).Replace("'", string.Empty);
                        _apimail.err_msg = Path.GetFileName(dtParam.PathInput).Replace("~$", string.Empty) + "-//-" + oKeepResponeExecute["errorCode"];
                        _apimail.input = dtParam.PathInput;
                        _apimail.path = dtParam.PathOutput;
                        _apimail.email = form.emailtxt.Text;
                        _apimail.taxseller = form.txtSellerTaxID.Text;
                        if (form.pingeng && oKeepResponeExecute["errorCode"].ToString() != "ER011")
                        {
                            _apimail.send_err_service();

                        }
                        if (oKeepResponeExecute["errorCode"].ToString() == "ER011")
                        {

                        }
                        else
                        {
                            Console.WriteLine(dtParam.PathInput);
                            CreateTextFile(pathErr, strOutputFile.MessageResultError);
                        }




                        //if (form.pingeng == true)
                        //{
                        //    sock.SendMailAlert(dtParam.PathInput, dtParam.PathOutput, "FE99", form.emailtxt.Text, form.txtSellerTaxID.Text, Path.GetFileName(dtParam.PathInput).Replace("~$", string.Empty) + "-//-" + oKeepResponeExecute["errorCode"], oKeepResponeExecute["errorMessage"].ToString());
                        //}


                    }
                    else
                    {
                        this.pathOutput = Path.GetFileNameWithoutExtension(pathText);
                        try
                        {
                            if (!Directory.Exists(dtParam.PathOutput + "\\" + "LogSucces"))
                            {
                                Directory.CreateDirectory(dtParam.PathOutput + "\\" + "LogSucces");
                            }
                        }
                        catch (Exception ea)
                        {
                            Console.WriteLine(ea);
                        }
                        finally
                        {
                            for (int i = 0; i <= GC.MaxGeneration; i++)
                            {
                                int count = GC.CollectionCount(i);
                                GC.Collect();
                            }
                            GC.WaitForPendingFinalizers();
                            GC.SuppressFinalize(this);
                        }
                        try
                        {
                            if (!Directory.Exists(dtParam.PathOutput + "\\" + "Log"))
                            {
                                Directory.CreateDirectory(dtParam.PathOutput + "\\" + "Log");
                            }
                        }
                        catch (Exception ea)
                        {
                            Console.WriteLine(ea);
                        }
                        finally
                        {
                            for (int i = 0; i <= GC.MaxGeneration; i++)
                            {
                                int count = GC.CollectionCount(i);
                                GC.Collect();
                            }
                            GC.WaitForPendingFinalizers();
                            GC.SuppressFinalize(this);
                        }
                        try
                        {
                            if (!Directory.Exists(dtParam.PathOutput + "\\" + "Temp_Succes"))
                            {
                                Directory.CreateDirectory(dtParam.PathOutput + "\\" + "Temp_Succes");
                            }
                        }
                        catch (Exception ea)
                        {
                            Console.WriteLine(ea);
                        }
                        finally
                        {
                            for (int i = 0; i <= GC.MaxGeneration; i++)
                            {
                                int count = GC.CollectionCount(i);
                                GC.Collect();
                            }
                            GC.WaitForPendingFinalizers();
                            GC.SuppressFinalize(this);
                        }
                        try
                        {
                            if (!Directory.Exists(dtParam.PathOutput + "\\" + "Log_Resend"))
                            {
                                Directory.CreateDirectory(dtParam.PathOutput + "\\" + "Log_Resend");
                            }
                        }
                        catch (Exception ea)
                        {
                            Console.WriteLine(ea);
                        }
                        finally
                        {
                            for (int i = 0; i <= GC.MaxGeneration; i++)
                            {
                                int count = GC.CollectionCount(i);
                                GC.Collect();
                            }
                            GC.WaitForPendingFinalizers();
                            GC.SuppressFinalize(this);
                        }

                        _apimail.err_code = "FE91";
                        _apimail.err_msg = strFileNameExtension.Replace("~$", string.Empty) + "-//-" + "";
                        _apimail.input = dtParam.PathInput;
                        _apimail.path = dtParam.PathOutput;
                        _apimail.email = form.emailtxt.Text;
                        _apimail.taxseller = form.txtSellerTaxID.Text;
                        int counttimes = 0;
                        bool string_check__pdf;
                        bool string_check__xml;
                        JObject json_respo = new JObject();
                        json_respo = JObject.Parse(strOutputFile.Message_Content);
                        Console.WriteLine(json_respo + " json_respo");
                        if (json_respo["status"].ToString() != "ER")
                        {
                            string Temp_succes = dtParam.PathOutput + "\\" + "Temp_Succes";
                            CreateTextFile(dtParam.PathOutput + "\\LogSucces\\" + "Success_" + Path.GetFileNameWithoutExtension(pathText) + ".txt", strOutputFile.Message_Content);
                        download_pdfandxml:
                            counttimes = counttimes + 1;
                            DownloadFile(strOutputFile.MessageResultPDF, Temp_succes, this.pathOutput + "_PDF.pdf");
                            DownloadFile(strOutputFile.MessageResultXML, Temp_succes, this.pathOutput + "_XML.xml");
                            string_check__pdf = _checkfolder_pdf(this.pathOutput + "_PDF.pdf", Temp_succes);
                            string_check__xml = _checkfolder_xml(this.pathOutput + "_XML.xml", Temp_succes);
                            if (string_check__pdf == false && string_check__xml == false && counttimes <= 3)
                            {
                                goto download_pdfandxml;
                            }
                            else if (string_check__pdf == false && counttimes <= 3)
                            {
                                goto download_pdfandxml;
                            }
                            else if (string_check__xml == false && counttimes <= 3)
                            {
                                goto download_pdfandxml;
                            }

                            if (string_check__pdf == false && string_check__xml == false)
                            {

                                _apimail.actionmsg = "ไม่สามารถดาวน์โหลดไฟล์ PDF และ XML ได้";
                                if (form.pingeng)
                                {
                                    _apimail.send_err_service();
                                }
                                CreateTextFile(dtParam.PathOutput + "\\Log\\" + "Log_" + Path.GetFileNameWithoutExtension(pathText) + ".txt", "ไม่สามารถดาวน์โหลดไฟล์ PDF และ XML ได้");
                            }
                            else if (string_check__pdf == false)
                            {
                                _apimail.actionmsg = "ไม่สามารถดาวน์โหลดไฟล์ PDF ได้";
                                if (form.pingeng)
                                {
                                    _apimail.send_err_service();
                                }
                                CreateTextFile(dtParam.PathOutput + "\\Log\\" + "Log_" + Path.GetFileNameWithoutExtension(pathText) + ".txt", "ไม่สามารถดาวน์โหลดไฟล์ PDF ได้");
                            }
                            else if (string_check__xml == false)
                            {
                                _apimail.actionmsg = "ไม่สามารถดาวน์โหลดไฟล์ XML ได้";
                                if (form.pingeng)
                                {
                                    _apimail.send_err_service();
                                }
                                CreateTextFile(dtParam.PathOutput + "\\Log\\" + "Log_" + Path.GetFileNameWithoutExtension(pathText) + ".txt", "ไม่สามารถดาวน์โหลดไฟล์ XML ได้");
                            }
                            try
                            {
                                var move_file_dis_form_pdf = Temp_succes + "\\" + this.pathOutput + "_PDF.pdf";
                                var move_file_dis_to_pdf = dtParam.PathOutput + "\\" + this.pathOutput + "_PDF.pdf";
                                var move_file_dis_form_xml = Temp_succes + "\\" + this.pathOutput + "_XML.xml";
                                var move_file_dis_to_xml = dtParam.PathOutput + "\\" + this.pathOutput + "_XML.xml";
                                if (string_check__pdf == true && string_check__xml == true)
                                {
                                    File.Move(move_file_dis_form_pdf, move_file_dis_to_pdf);
                                    File.Move(move_file_dis_form_xml, move_file_dis_to_xml);
                                }
                                else if (string_check__pdf == true && string_check__xml == false)
                                {
                                    File.Move(move_file_dis_form_pdf, move_file_dis_to_pdf);
                                }
                                else if (string_check__pdf == false && string_check__xml == true)
                                {
                                    File.Move(move_file_dis_form_xml, move_file_dis_to_xml);
                                }
                            }
                            catch (Exception ex)
                            {
                                Console.WriteLine(ex.Message);
                            }

                        }
                        else if (json_respo["errorCode"].ToString() == "ER011")
                        {
                            string Temp_succes = dtParam.PathOutput + "\\" + "Temp_Succes";
                            CreateTextFile(dtParam.PathOutput + "\\" + "Log_Resend\\" + "Resend_Success_" + Path.GetFileNameWithoutExtension(pathText) + ".txt", strOutputFile.Message_Content);
                            //MessageBox.Show(json_respo.ToString());
                            int counttimes_resend = 0;
                            string pathText_resend = "RESEND_" + UATorPROD + "_" + Path.GetFileNameWithoutExtension(dtParam.PathInput) + "_" + strDateTimeStamp + ".txt";
                            string pathOutput_resend = Path.GetFileNameWithoutExtension(pathText_resend);
                        download_pdf_resend:
                            counttimes_resend = counttimes_resend + 1;
                            DownloadFile(strOutputFile.MessageResultPDF, Temp_succes, pathOutput_resend + "_PDF.pdf");
                            DownloadFile(strOutputFile.MessageResultXML, Temp_succes, pathOutput_resend + "_XML.xml");
                            bool string_check_pdf_resend = _checkfolder_pdf(pathOutput_resend + "_PDF.pdf", Temp_succes);
                            bool string_check_xml_resend = _checkfolder_xml(pathOutput_resend + "_XML.xml", Temp_succes);
                            if (string_check_pdf_resend == false && string_check_xml_resend == false && counttimes_resend <= 3)
                            {
                                goto download_pdf_resend;
                            }
                            else if (string_check_pdf_resend == false && counttimes_resend <= 3)
                            {
                                goto download_pdf_resend;
                            }
                            else if (string_check_xml_resend == false && counttimes_resend <= 3)
                            {
                                goto download_pdf_resend;
                            }
                            if (string_check_pdf_resend == false && string_check_xml_resend == false)
                            {
                                CreateTextFile(dtParam.PathOutput + "\\Log\\" + "Log_" + Path.GetFileNameWithoutExtension(pathText) + ".txt", "ไม่สามารถดาวน์โหลดไฟล์ PDF และ XML ได้");
                            }
                            else if (string_check_pdf_resend == false)
                            {
                                CreateTextFile(dtParam.PathOutput + "\\Log\\" + "Log_" + Path.GetFileNameWithoutExtension(pathText) + ".txt", "ไม่สามารถดาวน์โหลดไฟล์ PDF ได้");
                            }
                            else if (string_check_xml_resend == false)
                            {
                                CreateTextFile(dtParam.PathOutput + "\\Log\\" + "Log_" + Path.GetFileNameWithoutExtension(pathText) + ".txt", "ไม่สามารถดาวน์โหลดไฟล์ XML ได้");
                            }

                            try
                            {
                                var move_file_dis_form_pdf = Temp_succes + "\\" + pathOutput_resend + "_PDF.pdf";
                                var move_file_dis_to_pdf = dtParam.PathOutput + "\\" + pathOutput_resend + "_PDF.pdf";
                                var move_file_dis_form_xml = Temp_succes + "\\" + pathOutput_resend + "_XML.xml";
                                var move_file_dis_to_xml = dtParam.PathOutput + "\\" + pathOutput_resend + "_XML.xml";
                                if (string_check_pdf_resend == true && string_check_xml_resend == true)
                                {
                                    File.Move(move_file_dis_form_pdf, move_file_dis_to_pdf);
                                    File.Move(move_file_dis_form_xml, move_file_dis_to_xml);
                                }
                                else if (string_check_pdf_resend == true && string_check_xml_resend == false)
                                {
                                    File.Move(move_file_dis_form_pdf, move_file_dis_to_pdf);
                                }
                                else if (string_check_pdf_resend == false && string_check_xml_resend == true)
                                {
                                    File.Move(move_file_dis_form_xml, move_file_dis_to_xml);
                                }
                            }
                            catch (Exception ex)
                            {
                                Console.WriteLine(ex.Message);
                            }

                        }

                        //CreateTextFile(dtParam.PathOutput, strOutputFile.Message_Content);
                        Console.WriteLine(dtParam.PathOutput + " 3872");
                    }
                }
                else if (chkOption == false)
                {
                    if (strOutputFile.StatusCallAPI == false)
                    {
                        string pathErr = pfIO.PathErr + "\\" + Path.GetFileNameWithoutExtension(pathText) + "_Error.txt";
                        CreateTextFile(pathErr, strOutputFile.MessageResultError);
                        JObject oKeepResponeExecute1 = new JObject();
                        oKeepResponeExecute1 = JObject.Parse(strOutputFile.MessageResultError.ToString());
                        Console.WriteLine(strOutputFile + " oKeepResponeExecute1");
                        Console.WriteLine(Path.GetFileName(dtParam.PathInput).Replace("~$", string.Empty) + "-//-" + oKeepResponeExecute1["errorCode"]);
                        try
                        {
                            _apimail.err_code = "FE99";
                            _apimail.actionmsg = oKeepResponeExecute1["errorMessage"].ToString().Replace(" ", string.Empty).Replace("\n", string.Empty).Replace(",", string.Empty).Replace("'", string.Empty);
                            _apimail.err_msg = Path.GetFileName(dtParam.PathInput).Replace("~$", string.Empty) + "-//-" + oKeepResponeExecute1["errorCode"];
                            _apimail.input = dtParam.PathInput;
                            _apimail.path = dtParam.PathOutput;
                            _apimail.email = form.emailtxt.Text;
                            _apimail.taxseller = form.txtSellerTaxID.Text;

                            if (form.pingeng)
                            {
                                _apimail.send_err_service();
                            }
                            Console.WriteLine(dtParam.PathInput);
                            CreateTextFile(pathErr, strOutputFile.MessageResultError);

                            //if (form.metroToggle1.Checked == true && form.pingeng == true)
                            //{
                            //    sock.SendMailAlert(dtParam.PathInput, dtParam.PathOutput, "FE99", form.emailtxt.Text, form.txtSellerTaxID.Text, Path.GetFileName(dtParam.PathInput).Replace("~$", string.Empty) + "-//-" + oKeepResponeExecute1["errorCode"], oKeepResponeExecute1["errorMessage"].ToString().Replace(" ", string.Empty).Replace("\n", string.Empty).Replace(",", string.Empty).Replace("'", string.Empty));
                            //}
                        }
                        catch (Exception ea)
                        {
                            Console.WriteLine(ea);
                        }

                        string[] arrFiles = System.IO.Directory.GetFiles(pfIO.PathTemp, "*.txt");
                        string[] arrFilesSource = System.IO.Directory.GetFiles(pfIO.PathSource_F, "*.txt");

                        foreach (var item in arrFiles)
                        {
                            string fileName = Path.GetFileName(item);
                            this.nameFilePDF = item;
                            string pathTxtNew = pfIO.PathSource_F + "\\" + fileName;
                            string pathTxtNew_S = pfIO.PathSource_S + "\\" + fileName;
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
                            File.Copy(item, pathTxtNew_S);
                            File.Move(item, pathTxtNew);
                        }
                    }
                    else
                    {
                        string fileNameWithoutExtension = string.Empty;
                        string[] arrFiles = System.IO.Directory.GetFiles(pfIO.PathTemp, "*.txt");
                        string[] arrFiles__pcfg = System.IO.Directory.GetFiles(pfIO.PathInput, "*.pcfg");

                        this.pathOutput = Path.GetFileNameWithoutExtension(pathText);
                        try
                        {
                            if (!Directory.Exists(pfIO.PathSuccess_O + "\\" + "LogSucces"))
                            {
                                Directory.CreateDirectory(pfIO.PathSuccess_O + "\\" + "LogSucces");
                            }
                        }
                        catch (Exception ea)
                        {
                            Console.WriteLine(ea);
                        }
                        finally
                        {
                            for (int i = 0; i <= GC.MaxGeneration; i++)
                            {
                                int count = GC.CollectionCount(i);
                                GC.Collect();
                            }
                            GC.WaitForPendingFinalizers();
                            GC.SuppressFinalize(this);
                        }
                        try
                        {
                            if (!Directory.Exists(pfIO.PathSuccess_O + "\\" + "Log"))
                            {
                                Directory.CreateDirectory(pfIO.PathSuccess_O + "\\" + "Log");
                            }
                        }
                        catch (Exception ea)
                        {
                            Console.WriteLine(ea);
                        }
                        finally
                        {
                            for (int i = 0; i <= GC.MaxGeneration; i++)
                            {
                                int count = GC.CollectionCount(i);
                                GC.Collect();
                            }
                            GC.WaitForPendingFinalizers();
                            GC.SuppressFinalize(this);
                        }
                        try
                        {
                            if (!Directory.Exists(pfIO.PathErr + "\\" + "Log_Resend"))
                            {
                                Directory.CreateDirectory(pfIO.PathErr + "\\" + "Log_Resend");
                            }
                        }
                        catch (Exception ea)
                        {
                            Console.WriteLine(ea);
                        }
                        finally
                        {
                            for (int i = 0; i <= GC.MaxGeneration; i++)
                            {
                                int count = GC.CollectionCount(i);
                                GC.Collect();
                            }
                            GC.WaitForPendingFinalizers();
                            GC.SuppressFinalize(this);
                        }
                        try
                        {
                            if (!Directory.Exists(pfIO.PathSuccess_O + "\\" + "Temp_Succes"))
                            {
                                Directory.CreateDirectory(pfIO.PathSuccess_O + "\\" + "Temp_Succes");
                            }
                        }
                        catch (Exception ea)
                        {
                            Console.WriteLine(ea);
                        }
                        finally
                        {
                            for (int i = 0; i <= GC.MaxGeneration; i++)
                            {
                                int count = GC.CollectionCount(i);
                                GC.Collect();
                            }
                            GC.WaitForPendingFinalizers();
                            GC.SuppressFinalize(this);
                        }
                        int counttimes = 0;
                        bool string_check__pdf;
                        bool string_check__xml;
                        _apimail.err_code = "FE91";
                        _apimail.err_msg = strFileNameExtension.Replace("~$", string.Empty) + "-//-" + "";
                        _apimail.input = dtParam.PathInput;
                        _apimail.path = dtParam.PathOutput;
                        _apimail.email = form.emailtxt.Text;
                        _apimail.taxseller = form.txtSellerTaxID.Text;
                        JObject ok_json = new JObject();
                        ok_json = JObject.Parse(strOutputFile.Message_Content);
                        Console.WriteLine(ok_json + " strOutputFile.Message_Content");

                        if (ok_json["status"].ToString() != "ER")
                        {
                            string Temp_succes = pfIO.PathSuccess_O + "\\" + "Temp_Succes";
                            CreateTextFile(pfIO.PathSuccess_O + "\\LogSucces\\" + "Success_" + Path.GetFileNameWithoutExtension(pathText) + ".txt", strOutputFile.Message_Content);
                        downloadpdfandxml:
                            counttimes = counttimes + 1;
                            DownloadFile(strOutputFile.MessageResultPDF, Temp_succes, this.pathOutput + "_PDF.pdf");
                            DownloadFile(strOutputFile.MessageResultXML, Temp_succes, this.pathOutput + "_XML.xml");
                            string_check__pdf = _checkfolder_pdf(this.pathOutput + "_PDF.pdf", Temp_succes);
                            string_check__xml = _checkfolder_xml(this.pathOutput + "_XML.xml", Temp_succes);
                            if (string_check__pdf == false && string_check__xml == false && counttimes <= 3)
                            {
                                goto downloadpdfandxml;
                            }
                            else if (string_check__pdf == false && string_check__xml == true && counttimes <= 3)
                            {
                                goto downloadpdfandxml;
                            }
                            else if (string_check__xml == false && string_check__pdf == true && counttimes <= 3)
                            {
                                goto downloadpdfandxml;
                            }
                            //MessageBox.Show(string_check__pdf.ToString());
                            if (string_check__pdf == false && string_check__xml == false)
                            {
                                _apimail.actionmsg = "ไม่สามารถดาวน์โหลดไฟล์ PDF และ XML ได้";
                                if (form.pingeng)
                                {
                                    _apimail.send_err_service();
                                }
                                CreateTextFile(pfIO.PathSuccess_O + "\\Log\\" + "Log_" + Path.GetFileNameWithoutExtension(pathText) + ".txt", "ไม่สามารถดาวน์โหลดไฟล์ PDF และ XML ได้");
                            }
                            else if (string_check__pdf == false && string_check__xml == true)
                            {
                                _apimail.actionmsg = "ไม่สามารถดาวน์โหลดไฟล์ PDF ได้";
                                if (form.pingeng)
                                {
                                    _apimail.send_err_service();
                                }
                                CreateTextFile(pfIO.PathSuccess_O + "\\Log\\" + "Log_" + Path.GetFileNameWithoutExtension(pathText) + ".txt", "ไม่สามารถดาวน์โหลดไฟล์ PDF ได้");
                            }
                            else if (string_check__xml == false && string_check__pdf == true)
                            {
                                _apimail.actionmsg = "ไม่สามารถดาวน์โหลดไฟล์ XML ได้";
                                if (form.pingeng)
                                {
                                    _apimail.send_err_service();
                                }
                                CreateTextFile(pfIO.PathSuccess_O + "\\Log\\" + "Log_" + Path.GetFileNameWithoutExtension(pathText) + ".txt", "ไม่สามารถดาวน์โหลดไฟล์ XML ได้");
                            }
                            try
                            {
                                var move_file_dis_form_pdf = Temp_succes + "\\" + this.pathOutput + "_PDF.pdf";
                                var move_file_dis_to_pdf = pfIO.PathSuccess_O + "\\" + this.pathOutput + "_PDF.pdf";
                                var move_file_dis_form_xml = Temp_succes + "\\" + this.pathOutput + "_XML.xml";
                                var move_file_dis_to_xml = pfIO.PathSuccess_O + "\\" + this.pathOutput + "_XML.xml";
                                if (string_check__pdf == true && string_check__xml == true)
                                {
                                    File.Move(move_file_dis_form_pdf, move_file_dis_to_pdf);
                                    File.Move(move_file_dis_form_xml, move_file_dis_to_xml);
                                }
                                else if (string_check__pdf == true && string_check__xml == false)
                                {
                                    File.Move(move_file_dis_form_pdf, move_file_dis_to_pdf);
                                }
                                else if (string_check__pdf == false && string_check__xml == true)
                                {
                                    File.Move(move_file_dis_form_xml, move_file_dis_to_xml);
                                }
                            }
                            catch (Exception ex)
                            {
                                Console.WriteLine(ex.Message);
                            }


                        }
                        else if (ok_json["errorCode"].ToString() == "ER011")
                        {
                            string Temp_succes = pfIO.PathSuccess_O + "\\" + "Temp_Succes";
                            CreateTextFile(pfIO.PathErr + "\\" + "Log_Resend\\" + "Resend_Success_" + Path.GetFileNameWithoutExtension(pathText) + ".txt", strOutputFile.Message_Content);
                            int counttimes_resend = 0;
                            string pathText_resend = "RESEND_" + UATorPROD + "_" + Path.GetFileNameWithoutExtension(dtParam.PathInput) + "_" + strDateTimeStamp + ".txt";
                            string pathOutput_resend = Path.GetFileNameWithoutExtension(pathText_resend);
                        downloadpdfandxml_resend:
                            counttimes_resend = counttimes_resend + 1;
                            DownloadFile(strOutputFile.MessageResultPDF, Temp_succes, pathOutput_resend + "_PDF.pdf");
                            DownloadFile(strOutputFile.MessageResultXML, Temp_succes, pathOutput_resend + "_XML.xml");
                            bool string_check__pdf_resend = _checkfolder_pdf(pathOutput_resend + "_PDF.pdf", Temp_succes);
                            bool string_check__xml_resend = _checkfolder_xml(pathOutput_resend + "_XML.xml", Temp_succes);
                            if (string_check__pdf_resend == false && string_check__xml_resend == false && counttimes_resend <= 3)
                            {
                                goto downloadpdfandxml_resend;
                            }
                            else if (string_check__pdf_resend == false && string_check__xml_resend == true && counttimes_resend <= 3)
                            {
                                goto downloadpdfandxml_resend;
                            }
                            else if (string_check__xml_resend == false && string_check__pdf_resend == true && counttimes_resend <= 3)
                            {
                                goto downloadpdfandxml_resend;
                            }
                            if (string_check__pdf_resend == false && string_check__xml_resend == false)
                            {
                                CreateTextFile(pfIO.PathSuccess_O + "\\" + "Log\\" + "Log_" + Path.GetFileNameWithoutExtension(pathText) + ".txt", "ไม่สามารถดาวน์โหลดไฟล์ PDF และ XML ได้");
                            }
                            else if (string_check__pdf_resend == false && string_check__xml_resend == true)
                            {
                                CreateTextFile(pfIO.PathSuccess_O + "\\Log\\" + "Log_" + Path.GetFileNameWithoutExtension(pathText) + ".txt", "ไม่สามารถดาวน์โหลดไฟล์ PDF ได้");
                            }
                            else if (string_check__xml_resend == false && string_check__pdf_resend == true)
                            {
                                CreateTextFile(pfIO.PathSuccess_O + "\\Log\\" + "Log_" + Path.GetFileNameWithoutExtension(pathText) + ".txt", "ไม่สามารถดาวน์โหลดไฟล์ XML ได้");
                            }
                            try
                            {
                                var move_file_dis_form_pdf = Temp_succes + "\\" + pathOutput_resend + "_PDF.pdf";
                                var move_file_dis_to_pdf = pfIO.PathSuccess_O + "\\" + pathOutput_resend + "_PDF.pdf";
                                var move_file_dis_form_xml = Temp_succes + "\\" + pathOutput_resend + "_XML.xml";
                                var move_file_dis_to_xml = pfIO.PathSuccess_O + "\\" + pathOutput_resend + "_XML.xml";
                                if (string_check__pdf_resend == true && string_check__xml_resend == true)
                                {
                                    File.Move(move_file_dis_form_pdf, move_file_dis_to_pdf);
                                    File.Move(move_file_dis_form_xml, move_file_dis_to_xml);
                                }
                                else if (string_check__pdf_resend == true && string_check__xml_resend == false)
                                {
                                    File.Move(move_file_dis_form_pdf, move_file_dis_to_pdf);
                                }
                                else if (string_check__pdf_resend == false && string_check__xml_resend == true)
                                {
                                    File.Move(move_file_dis_form_xml, move_file_dis_to_xml);
                                }
                            }
                            catch (Exception ex)
                            {
                                Console.WriteLine(ex.Message);
                            }

                        }



                        Thread.Sleep(1000);
                        string namefilepdf = Path.GetFileName(dtParam.PathInput);
                        etaxOneth_Printer.Class1 _printer = new etaxOneth_Printer.Class1();
                        if (pfIO.TypePrinting == "A" && form.check___copies.Checked == false)
                        {
                            var Timestamp = new DateTimeOffset(DateTime.UtcNow).ToUnixTimeSeconds();
                            timest_process = Int32.Parse(Timestamp.ToString()) - timest_process;
                            Console.WriteLine(timest_process + " timest_process");
                            CreateTextFile(pfIO.LogTimeProcess + "\\LogProcess_Print_" + Path.GetFileNameWithoutExtension(pathText) + ".txt", "เอกสาร " + Path.GetFileNameWithoutExtension(pathText) + " ใช้เวลาการประมวลผลทั้งหมดประมาณ " + timest_process + " วินาที");
                            //PrinterSettings.SetDefaultPrinter(pfIO.Printer);
                            //ProcessStartInfo printProcessInfo = new ProcessStartInfo()
                            //{
                            //    UseShellExecute = true,
                            //    Verb = "print",
                            //    CreateNoWindow = true,
                            //    FileName = pfIO.PathSuccess_O + "\\" + this.pathOutput + "_PDF.pdf",
                            //    //Arguments = printDialog1.PrinterSettings.PrinterName.ToString(),
                            //    WindowStyle = ProcessWindowStyle.Hidden
                            //};
                            //_printer.PrintMethod("C:\\Users\\JIRAYU-NB\\Documents\\FillTEST\\output\\Success\\UAT_0105561072420_03-4-62T17-32-15_PDF.pdf", "ApeosPort-IV C5570 16", 1);
                            //Console.WriteLine(pfIO.PathSuccess_O + "\\" + this.pathOutput + "_PDF.pdf" + " " + pfIO.Printer + " " + short.Parse(form.input_copies.Text));
                            _printer.PrintMethod(pfIO.PathSuccess_O + "\\" + this.pathOutput + "_PDF.pdf", pfIO.Printer, short.Parse(form.input_copies.Text));

                            //try
                            //{
                            //    Process printProcess = new Process();
                            //    printProcess.StartInfo = printProcessInfo;
                            //    printProcess.Start();
                            //    //Thread.Sleep(3000);
                            //    //if (printProcess.HasExited == false)
                            //    //{
                            //    //    printProcess.Kill();
                            //    //}
                            //}
                            //catch (Exception ex)
                            //{
                            //    //MessageBox.Show(ex.ToString());
                            //    //MessageBox.Show("ไม่พบตัวอ่านไฟล์ของคุณ");
                            //}
                        }
                        else if (pfIO.TypePrinting == "A" && form.check___copies.Checked == true)
                        {
                            //MessageBox.Show(pfIO.PathInput.Split('\\')[pfIO.PathInput.Split('\\').Length -1]);
                            if (arrFiles__pcfg.Count() != 0)
                            {
                                foreach (var item in arrFiles__pcfg)
                                {
                                    string namefile__ = Path.GetFileName(item);
                                    var Timestamp = new DateTimeOffset(DateTime.UtcNow).ToUnixTimeSeconds();
                                    timest_process = Int32.Parse(Timestamp.ToString()) - timest_process;
                                    Console.WriteLine(timest_process + " timest_process");
                                    CreateTextFile(pfIO.LogTimeProcess + "\\LogProcess_Print_" + Path.GetFileNameWithoutExtension(pathText) + ".txt", "เอกสาร " + Path.GetFileNameWithoutExtension(pathText) + " ใช้เวลาการประมวลผลทั้งหมดประมาณ " + timest_process + " วินาที");
                                    if (Path.GetFileNameWithoutExtension(namefilepdf) == Path.GetFileNameWithoutExtension(namefile__))
                                    {
                                        // Open the text file using a stream reader.
                                        using (StreamReader sr = new StreamReader(item))
                                        {
                                            // Read the stream to a string, and write the string to the console.
                                            String line = sr.ReadToEnd();
                                            bool string___checkinpcfg = checkcopiesin__pcfg(line.Replace(" ", string.Empty).Replace(Environment.NewLine, string.Empty).Replace("\t", string.Empty));
                                            if (string___checkinpcfg == true)
                                            {
                                                try
                                                {
                                                    _printer.PrintMethod(pfIO.PathSuccess_O + "\\" + this.pathOutput + "_PDF.pdf", pfIO.Printer, short.Parse(line));
                                                }
                                                catch (Exception ex)
                                                {
                                                    Console.WriteLine(ex);
                                                }
                                                finally
                                                {
                                                    sr.Close();
                                                    File.Delete(item);
                                                }
                                            }
                                            else if (string___checkinpcfg == false)
                                            {
                                                sr.Close();
                                                File.Delete(item);
                                                MessageBox.Show("ไม่สามารถปริ้นได้เนื่องจาก จำนวน Copies ไม่ถูกต้อง *ควรระบุ 1-99*",
                                                                            "แจ้งเตือน",
                                                                MessageBoxButtons.OK,
                                                                MessageBoxIcon.Error);
                                            }
                                        }
                                    }
                                    else
                                    {
                                        File.Delete(item);
                                        MessageBox.Show("ชื่อไฟล์ .pcfg ไม่ตรงกับไฟล์ที่นำเข้า",
                                                        "แจ้งเตือน",
                                                        MessageBoxButtons.OK,
                                                        MessageBoxIcon.Warning);
                                    }
                                }

                            }
                            else
                            {
                                MessageBox.Show("ไม่พบไฟล์ .pcfg",
                                                    "แจ้งเตือน",
                                                    MessageBoxButtons.OK,
                                                    MessageBoxIcon.Warning);
                            }



                            //for (int i = 0; i < arrFiles__pcfg.Length; i++)
                            //{
                            //    string namefile__cfg = Path.GetFileName(arrFiles__pcfg[i]);
                            //    MessageBox.Show(namefilepdf);
                            //    if(Path.GetFileNameWithoutExtension(namefilepdf) == Path.GetFileNameWithoutExtension(namefile__cfg))
                            //    {
                            //        try
                            //        {   // Open the text file using a stream reader.
                            //            using (StreamReader sr = new StreamReader(arrFiles__pcfg[i]))
                            //            {
                            //                // Read the stream to a string, and write the string to the console.
                            //                String line = sr.ReadToEnd();
                            //                bool string___checkinpcfg = checkcopiesin__pcfg(line.Replace(" ", string.Empty).Replace(Environment.NewLine,string.Empty));
                            //                if(string___checkinpcfg == true)
                            //                {
                            //                    try
                            //                    {
                            //                        _printer.PrintMethod(pfIO.PathSuccess_O + "\\" + this.pathOutput + "_PDF.pdf", pfIO.Printer, short.Parse(line));
                            //                    }
                            //                    catch(Exception ex)
                            //                    {
                            //                        Console.WriteLine(ex);
                            //                    }
                            //                    finally
                            //                    {
                            //                        sr.Close();
                            //                        File.Delete(arrFiles__pcfg[i]);
                            //                    }


                            //                }
                            //                else if(string___checkinpcfg == false)
                            //                {
                            //                    sr.Close();
                            //                    File.Delete(arrFiles__pcfg[i]);
                            //                    MessageBox.Show("ไม่สามารถปริ้นได้เนื่องจาก จำนวน Copies เกิน 99 แผ่น",
                            //                                                "แจ้งเตือน",
                            //                                                MessageBoxButtons.OK,
                            //                                                MessageBoxIcon.Error);

                            //                }
                            //                //MessageBox.Show(string___checkinpcfg);

                            //            }
                            //        }
                            //        catch (Exception e)
                            //        {
                            //            Console.WriteLine("The file could not be read:");
                            //            Console.WriteLine(e.Message);
                            //        }
                            //    }
                            //    else
                            //    {

                            //    }

                            //}

                            //if(namefilepdf.Split(',')[namefilepdf.Split(',').Length - 1] + ".pcfg" == )
                            //bool check__print = _checkfolder(namefilepdf.Split(',')[namefilepdf.Split(',').Length - 1] + ".pcfg", pfIO.PathSuccess_O);
                            //MessageBox.Show(check__print.ToString());
                        }
                    }

                }

                form.pgbLoad.Value = 100;
                form.OutputPrc(100, "Export Data: 100%");
                //pgbLoad.Value = 100;
                //lbPercent.Text = "Export Data: 100%";
                //lbPercent.Refresh();
            }
            catch (FileNotFoundException ex)
            {
                form.txtStatus.Text += Environment.NewLine + "   -**********ไม่พบ Path File ที่คุณระบุ กรุณาเลือก Path File ให้ถูกต้อง!**********";
                strTempLogTime += "ไม่พบ Path File ที่คุณระบุ กรุณาเลือก Path File ให้ถูกต้อง!";

                if (chkOption == true)
                {
                    CreateTextFile(dtParam.PathOutput + "\\" + strFileName + "_" + strDateTimeStamp + "_Error.txt", "ไม่พบ Path File ที่คุณระบุ กรุณาเลือก Path File ให้ถูกต้อง");
                }
                else
                {
                    CreateTextFile(pfIO.PathErr + "\\" + strFileName + "_" + strDateTimeStamp + "_Error.txt", "ไม่พบ Path File ที่คุณระบุ กรุณาเลือก Path File ให้ถูกต้อง");
                }

                form.txtStatus.Refresh();
                cntFail++;
                //MessageBox.Show("ไม่พบ Path File ที่คุณระบุ กรุณาเลือก Path File ให้ถูกต้อง");
            }
            catch (System.IndexOutOfRangeException e)
            {
                form.txtStatus.Text += Environment.NewLine + "   -**********ไฟล์ของคุณมีข้อผิดพลาดในข้อมูลที่ใส่!**********";
                strTempLogTime += "กรุณาตรวจสอบไฟล์ของคุณ!";

                if (chkOption == true)
                {
                    CreateTextFile(dtParam.PathOutput + "\\" + strFileName + "_" + strDateTimeStamp + "_Error.txt", "ไฟล์ของคุณมีข้อผิดพลาด กรุณาตรวจสอบและใส่ข้อมูลให้ถูกต้อง");
                }
                else
                {
                    CreateTextFile(pfIO.PathErr + "\\" + strFileName + "_" + strDateTimeStamp + "_Error.txt", "ไฟล์ของคุณมีข้อผิดพลาด กรุณาตรวจสอบและใส่ข้อมูลให้ถูกต้อง");
                }

                form.txtStatus.Refresh();
                cntFail++;
            }
            catch (DirectoryNotFoundException ex)
            {
                Console.WriteLine(ex.Message);
            }
            catch (XmlException ex)
            {
                FileVersionInfo fvi = FileVersionInfo.GetVersionInfo(assembly.Location);
                string version = fvi.FileVersion;
                //MessageBox.Show("ไฟล์มีปัญหา!!!");
                Console.WriteLine(ex.Message);
                form.txtStatus.Text += Environment.NewLine + "   -**********Convert Fail!**********";
                strTempLogTime += " Convert Fail!" +ex.Message+" Version : " + version;
                //MessageBox.Show(ErrorMessage);
                if (chkOption == true)
                {
                    CreateTextFile(dtParam.PathOutput + "\\" + strFileName + "_" + strDateTimeStamp + "_Error.txt", "Convert Fail");
                    CreateTextFile(pfIO.PathErr + "\\" + strFileName + "_" + strDateTimeStamp + "_ErrorServiceOrProgram.txt", ex.Message);

                }
                else
                {

                    CreateTextFile(dtParam.PathOutput + "\\" + strFileName + "_" + strDateTimeStamp + "_Error.txt", "Convert Fail");
                    CreateTextFile(pfIO.PathErr + "\\" + strFileName + "_" + strDateTimeStamp + "_ErrorServiceOrProgram.txt", ex.Message);

                }

                form.txtStatus.Refresh();
                cntFail++;
            }
            catch (Exception ex)
            {

                //MessageBox.Show(ex.Message);
                FileVersionInfo fvi = FileVersionInfo.GetVersionInfo(assembly.Location);
                string version = fvi.FileVersion;
                form.txtStatus.Text += Environment.NewLine + "   -**********Convert Fail!**********";
                strTempLogTime += " Convert Fail! " +ex.Message+" Version : " + version;
                if (chkOption == true)
                {
                    CreateTextFile(dtParam.PathOutput + "\\" + strFileName + "_" + strDateTimeStamp + "_Error.txt", "Convert Fail");
                    CreateTextFile(pfIO.PathErr + "\\" + strFileName + "_" + strDateTimeStamp + "_ErrorServiceOrProgram.txt", ex.Message);
                }
                else
                {
                    CreateTextFile(dtParam.PathOutput + "\\" + strFileName + "_" + strDateTimeStamp + "_Error.txt", "Convert Fail");
                    CreateTextFile(pfIO.PathErr + "\\" + strFileName + "_" + strDateTimeStamp + "_ErrorServiceOrProgram.txt", ex.Message);
                }

                form.txtStatus.Refresh();
                cntFail++;
            }
            finally
            {
                lstTempSumAmount.Clear();
                for (int i = 0; i <= GC.MaxGeneration; i++)
                {
                    int count = GC.CollectionCount(i);
                    GC.Collect();
                }
                GC.WaitForPendingFinalizers();
                GC.SuppressFinalize(this);
            }
        }
        public void WorkProcess_BCP(PathFilesIO pfIO, string strFileName, string strFileNameExtension, out int cntFail, etaxOneth form)
        {
            a = null;
            b = null;
            a_with_b = null;
        loop:

            cntFail = 0;
            BCP_Output = new getModelOutPutViladateSign();
            string strTaxID;
            string typedoc = "", sellertaxid = "", sellerbranchid = "", document_name = "", document_id = "", document_issue_dtm = "", create_purpose_code = "",
                create_purpose = "", additional_ref_assign_id = "", additional_ref_issue_dtm = "", buyer_name = "", buyer_branch_id = "",
                buyer_tax_id = "", buyer_uriid = "", buyer_address = "", buyer_countrypostcode = "", rangeoflistitem = "", noitem = "", description = "",
                priceunit = "", quanlity = "", amount = "", totalamount = "", vat = "", total = "", document_remark = "",
                discount = "", totaldiscount = "", original_total_amount = "", line_total_amount = "", adjusted_information_amount = "", allowance_total_amount = "",
                tax_basis_total_amount = "", countrybuyer = "", typebuyer = "", buyer_order_assign_id = "", buyer_order_issue_dtm = "", vat_rate = "";
            string pathPdf;

            if (dtParam.ServiceURL == "https://uatetaxsp.one.th/etaxdocumentws/etaxsigndocument")
            {
                UATorPROD = "UAT";
            }
            else
            {
                UATorPROD = "PROD";
            }
            try
            {
                pathText = string.Empty;
                if (Path.GetExtension(dtParam.PathInput).Equals(".xlsx") || Path.GetExtension(dtParam.PathInput).Equals(".xls"))
                {
                    try
                    {
                        using (var reader = new StreamReader(dtParam.PathConfigExcel))
                        {
                            List<string> listA = new List<string>();
                            List<string> listB = new List<string>();
                            while (!reader.EndOfStream)
                            {
                                var line = reader.ReadLine();
                                var values = line.Split(';');
                                listA.Add(values[0]);
                                //listB.Add(values[1]);
                            }
                            foreach (var item in listA)
                            {
                                string[] arrItem = item.Split(',');
                                //Console.WriteLine(arrItem[0]);
                                switch (arrItem[0].ToLower().Trim(' '))
                                {
                                    case "typedoc":
                                        if (!arrItem[2].Equals(""))
                                            typedoc = arrItem[2];
                                        else
                                            typedoc = arrItem[2];
                                        break;
                                    case "discount":
                                        if (!arrItem[2].Equals(""))
                                            discount = arrItem[2] + "," + arrItem[3];
                                        else
                                            discount = arrItem[2];
                                        break;
                                    case "sellertaxid":
                                        if (!arrItem[2].Equals(""))
                                            sellertaxid = arrItem[2] + "," + arrItem[3];
                                        else
                                            sellertaxid = arrItem[2];
                                        break;
                                    case "sellerbranchid":
                                        if (!arrItem[2].Equals(""))
                                            sellerbranchid = arrItem[2] + "," + arrItem[3];
                                        else
                                            sellerbranchid = arrItem[2];
                                        break;
                                    case "document_name":
                                        if (!arrItem[2].Equals(""))
                                            document_name = arrItem[2] + "," + arrItem[3];
                                        else
                                            document_name = arrItem[2];
                                        break;
                                    case "document_id":
                                        if (!arrItem[2].Equals(""))
                                            document_id = arrItem[2] + "," + arrItem[3];
                                        else
                                            document_id = arrItem[2];
                                        break;
                                    case "document_remark":
                                        if (!arrItem[2].Equals(""))
                                            document_remark = arrItem[2] + "," + arrItem[3];
                                        else
                                            document_remark = arrItem[2];
                                        break;
                                    case "document_issue_dtm":
                                        if (!arrItem[2].Equals(""))
                                        {
                                            document_issue_dtm = arrItem[2] + "," + arrItem[3];
                                        }
                                        else
                                            document_issue_dtm = arrItem[2];
                                        break;
                                    case "create_purpose_code":
                                        if (!arrItem[2].Equals(""))
                                            create_purpose_code = arrItem[2] + "," + arrItem[3];
                                        else
                                            create_purpose_code = arrItem[2];
                                        break;
                                    case "create_purpose":
                                        if (!arrItem[2].Equals(""))
                                            create_purpose = arrItem[2] + "," + arrItem[3];
                                        else
                                            create_purpose = arrItem[2];
                                        break;
                                    case "additional_ref_assign_id":
                                        if (!arrItem[2].Equals(""))
                                            additional_ref_assign_id = arrItem[2] + "," + arrItem[3];
                                        else
                                            additional_ref_assign_id = arrItem[2];
                                        break;
                                    case "additional_ref_issue_dtm":
                                        if (!arrItem[2].Equals(""))
                                            additional_ref_issue_dtm = arrItem[2] + "," + arrItem[3];
                                        else
                                            additional_ref_issue_dtm = arrItem[2];
                                        break;
                                    case "buyer_name":
                                        if (!arrItem[2].Equals(""))
                                            buyer_name = arrItem[2] + "," + arrItem[3];
                                        else
                                            buyer_name = arrItem[2];
                                        break;
                                    case "buyer_tax_id":
                                        if (!arrItem[2].Equals(""))
                                            buyer_tax_id = arrItem[2] + "," + arrItem[3];
                                        else
                                            buyer_tax_id = arrItem[2];
                                        break;
                                    case "buyer_branch_id":
                                        if (!arrItem[2].Equals(""))
                                            buyer_branch_id = arrItem[2] + "," + arrItem[3];
                                        else
                                            buyer_branch_id = arrItem[2];
                                        break;
                                    case "buyer_uriid":
                                        if (!arrItem[2].Equals(""))
                                            buyer_uriid = arrItem[2] + "," + arrItem[3];
                                        else
                                            buyer_uriid = arrItem[2];
                                        break;
                                    case "buyer_address":
                                        if (!arrItem[2].Equals(""))
                                            buyer_address = arrItem[2] + "," + arrItem[3];
                                        else
                                            buyer_address = arrItem[2];
                                        break;
                                    case "buyer_country&postcode":
                                        if (!arrItem[2].Equals(""))
                                            buyer_countrypostcode = arrItem[2] + "," + arrItem[3];
                                        else
                                            buyer_countrypostcode = arrItem[2];
                                        break;
                                    case "rangeoflistitem":
                                        rangeoflistitem = arrItem[2];
                                        string[] splitData = arrItem[2].Split('_');
                                        numberRange = splitData[0].Split('-');
                                        charRange = splitData[1].Split('-');

                                        foreach (var k in charRange)
                                        {
                                            AllChar = AllChar + k;
                                        }

                                        break;
                                    case "no.item":
                                        if (!arrItem[2].Equals(""))
                                            noitem = arrItem[2] + "," + arrItem[3];
                                        else
                                            noitem = arrItem[2];
                                        break;
                                    case "description":
                                        if (!arrItem[2].Equals(""))
                                            description = arrItem[2] + "," + arrItem[3];
                                        else
                                            description = arrItem[2];
                                        break;
                                    case "price/unit":
                                        if (!arrItem[2].Equals(""))
                                            priceunit = arrItem[2] + "," + arrItem[3];
                                        else
                                            priceunit = arrItem[2];
                                        break;
                                    case "quanlity":
                                        if (!arrItem[2].Equals(""))
                                            quanlity = arrItem[2] + "," + arrItem[3];
                                        else
                                            quanlity = arrItem[2];
                                        break;
                                    case "amount":
                                        if (!arrItem[2].Equals(""))
                                            amount = arrItem[2] + "," + arrItem[3];
                                        else
                                            amount = arrItem[2];
                                        break;
                                    case "totalamount":
                                        if (!arrItem[2].Equals(""))
                                            totalamount = arrItem[2] + "," + arrItem[3];
                                        else
                                            totalamount = arrItem[2];
                                        break;
                                    case "totaldiscount":
                                        if (!arrItem[2].Equals(""))
                                            totaldiscount = arrItem[2] + "," + arrItem[3];
                                        else
                                            totaldiscount = arrItem[2];
                                        break;
                                    case "vat":
                                        if (!arrItem[2].Equals(""))
                                            vat = arrItem[2] + "," + arrItem[3];
                                        else
                                            vat = arrItem[2];
                                        break;
                                    case "total":
                                        if (!arrItem[2].Equals(""))
                                            total = arrItem[2] + "," + arrItem[3];
                                        else
                                            total = arrItem[2];
                                        break;
                                    case "original_total_amount":
                                        if (!arrItem[2].Equals(""))
                                            original_total_amount = arrItem[2] + "," + arrItem[3];
                                        else
                                            original_total_amount = arrItem[2];
                                        break;
                                    case "line_total_amount":
                                        if (!arrItem[2].Equals(""))
                                            line_total_amount = arrItem[2] + "," + arrItem[3];
                                        else
                                            line_total_amount = arrItem[2];
                                        break;
                                    case "adjusted_information_amount":
                                        if (!arrItem[2].Equals(""))
                                            adjusted_information_amount = arrItem[2] + "," + arrItem[3];
                                        else
                                            adjusted_information_amount = arrItem[2];
                                        break;
                                    case "allowance_total_amount":
                                        if (!arrItem[2].Equals(""))
                                            allowance_total_amount = arrItem[2] + "," + arrItem[3];
                                        else
                                            allowance_total_amount = arrItem[2];
                                        break;
                                    case "tax_basis_total_amount":
                                        if (!arrItem[2].Equals(""))
                                            tax_basis_total_amount = arrItem[2] + "," + arrItem[3];
                                        else
                                            tax_basis_total_amount = arrItem[2];
                                        break;
                                    case "countrybuyer":
                                        if (!arrItem[2].Equals(""))
                                            countrybuyer = arrItem[2] + "," + arrItem[3];
                                        else
                                            countrybuyer = arrItem[2];
                                        break;
                                    case "typebuyer":
                                        if (!arrItem[2].Equals(""))
                                            typebuyer = arrItem[2] + "," + arrItem[3];
                                        else
                                            typebuyer = arrItem[2];
                                        break;
                                    case "buyer_order_assign_id":
                                        if (!arrItem[2].Equals(""))
                                            buyer_order_assign_id = arrItem[2] + "," + arrItem[3];
                                        else
                                            buyer_order_assign_id = arrItem[2];
                                        break;
                                    case "buyer_order_issue_dtm":
                                        if (!arrItem[2].Equals(""))
                                            buyer_order_issue_dtm = arrItem[2] + "," + arrItem[3];
                                        else
                                            buyer_order_issue_dtm = arrItem[2];
                                        break;
                                    case "vat_rate":
                                        if (!arrItem[3].Equals(""))
                                            vat_rate = arrItem[3];
                                        else
                                            vat_rate = arrItem[3];
                                        break;
                                    default:
                                        break;
                                }
                            }
                            //foreach (var item in listA)
                            //{
                            //    Console.WriteLine(item);
                            //}
                            //Console.WriteLine(listA[0]);
                            //Console.WriteLine(listB);
                        } //End of Using for Read ConfigExcel
                        form.pgbLoad.Value = 0;
                        form.OutputPrc(0, "Export Data: 0%");
                        //pgbLoad.Value = 0;
                        //lbPercent.Text = "Export Data: 0%";
                        //lbPercent.Refresh();

                        List<string> lstDataRow = new List<string>();
                        List<string> lstDataMenu = new List<string>();
                        string strSheetName = string.Empty;
                        BGroup grpB = new BGroup();
                        CGroup grpC = new CGroup();
                        LGroup grpL = new LGroup();
                        HGroup grpH = new HGroup();
                        FGroup grpF = new FGroup();
                        Workbook workbook = new Workbook();
                        workbook.LoadFromFile(dtParam.PathInput);
                        sheet = workbook.Worksheets[0];
                        try
                        {
                            //DateTime dateValue;
                            if (DateTime.TryParse(getvalue(document_issue_dtm), out dateValue))
                            {
                                arrDateSplit = getvalue(document_issue_dtm).Split('/');
                                strYear = DateTime.Now.Year.ToString();
                                strYearFront = strYear.Substring(0, 2);
                                DiffOfYears = int.Parse(strYear) - (int.Parse(arrDateSplit[2].Split(' ')[0]) - 543); //ต้องลบ543เพราะ โปรแกรม+543ให้เองอัตโนมัติจึงลบออกเพื่อให้ได้ค่าที่ถูกต้อง
                                if (DiffOfYears < 0)
                                {
                                    DiffOfYears = 543;
                                }
                                else
                                {
                                    DiffOfYears = 0;
                                }
                                if (((int.Parse(arrDateSplit[2].Split(' ')[0]) - 543) - DiffOfYears) < 2000)
                                {
                                    years = (int.Parse(arrDateSplit[2].Split(' ')[0])) - DiffOfYears;
                                }
                                else
                                {
                                    years = (int.Parse(arrDateSplit[2].Split(' ')[0]) - 543) - DiffOfYears;
                                }
                                Console.WriteLine("YearsNow: " + strYear + " Years : " + arrDateSplit[2].Split(' ')[0]);
                                strDocID = arrDateSplit[2].Split(' ')[0] + "-" + Month[arrDateSplit[1].Replace(" ", string.Empty)] + "-" + Day[arrDateSplit[0].Replace(" ", string.Empty)];
                            }
                            else
                            {
                                MessageBox.Show("document_issue_dtm ไม่ถูกต้อง => " + getvalue(document_issue_dtm));
                            }
                        }
                        catch (ArgumentException e)
                        {
                            Console.WriteLine(e.Message);
                        }
                        catch (NullReferenceException e)
                        {
                            Console.WriteLine(e.Message);
                        }
                        catch (IndexOutOfRangeException e)
                        {
                            Console.WriteLine(e.Message);
                        }
                        catch (KeyNotFoundException e)
                        {
                            Console.WriteLine("Date Wrong");
                        }
                        //MessageBox.Show(strDocID);

                        if (form.txtStatus != null && !form.txtStatus.Text.Equals(""))
                        {
                            form.txtStatus.Text += Environment.NewLine + "เลขที่เอกสาร " + getvalue(document_id) + " วันที่ " + strDocID + " ชื่อไฟล์ " + strFileNameExtension + ":";
                            strTempLogTime += Environment.NewLine + "เลขที่เอกสาร " + getvalue(document_id) + " วันที่ " + strDocID + " ชื่อไฟล์ " + strFileNameExtension + ":";
                        }
                        else
                        {
                            form.txtStatus.Text = "เลขที่เอกสาร " + getvalue(document_id) + " วันที่ " + strDocID + " ชื่อไฟล์ " + strFileNameExtension + " :";
                            strTempLogTime = "เลขที่เอกสาร " + getvalue(document_id) + " วันที่ " + strDocID + " ชื่อไฟล์ " + strFileNameExtension + " :";
                        }

                        form.pgbLoad.Value = 10;
                        form.OutputPrc(10, "Export Data: 10%");

                        for (int x = int.Parse(numberRange[0]); x <= int.Parse(numberRange[1]); x++)
                        {
                            for (int k = 0; k < charRange.Length; k++)
                            {
                                string value;
                                if (charRange[k] == noitem.Split(',')[0])
                                {
                                    value = getvalue(charRange[k] + x + "," + noitem.Split(',')[1]);
                                    Console.WriteLine(value);
                                    object cellValue = value;
                                    lstDataMenu.Add(cellValue.ToString());
                                }
                                else if (charRange[k] == description.Split(',')[0])
                                {
                                    value = getvalue(charRange[k] + x + "," + description.Split(',')[1]);
                                    Console.WriteLine(value);
                                    object cellValue = value;
                                    lstDataMenu.Add(cellValue.ToString());
                                }
                                else if (charRange[k] == priceunit.Split(',')[0])
                                {
                                    value = getvalue(charRange[k] + x + "," + priceunit.Split(',')[1]);
                                    Console.WriteLine(value);
                                    if (!value.Equals(""))
                                    {
                                        object cellValue = value;
                                        lstDataMenu.Add(cellValue.ToString());
                                    }
                                    else
                                    {
                                        object cellValue = "";
                                        lstDataMenu.Add(cellValue.ToString());
                                    }
                                }
                                else if (charRange[k] == quanlity.Split(',')[0])
                                {
                                    value = getvalue(charRange[k] + x + "," + quanlity.Split(',')[1]);
                                    if (!value.Equals(""))
                                    {
                                        object cellValue = value;
                                        Console.WriteLine(cellValue);
                                        lstDataMenu.Add(cellValue.ToString());
                                    }
                                    else
                                    {
                                        object cellValue = "";
                                        lstDataMenu.Add(cellValue.ToString());
                                    }
                                }
                                else if (charRange[k] == discount.Split(',')[0])
                                {
                                    value = getvalue(charRange[k] + x + "," + discount.Split(',')[1]);
                                    if (!value.Equals(""))
                                    {
                                        object cellValue = value;
                                        lstDataMenu.Add(cellValue.ToString());
                                    }
                                    else
                                    {
                                        object cellValue = "";
                                        lstDataMenu.Add(cellValue.ToString());
                                    }
                                }
                                else if (charRange[k] == amount.Split(',')[0])
                                {
                                    if (!sheet.Range[charRange[k] + x].Value.Equals(""))
                                    {
                                        value = getvalue(charRange[k] + x + "," + amount.Split(',')[1]);
                                        if (!value.Equals(""))
                                        {
                                            object cellValue = value;
                                            try
                                            {
                                                lstDataMenu.Add(Double.Parse(cellValue.ToString()).ToString("0.00"));
                                            }
                                            catch (FormatException e)
                                            {
                                                Console.WriteLine(e);
                                            }
                                        }
                                        else
                                        {
                                            object cellValue = "";
                                            lstDataMenu.Add(cellValue.ToString());
                                        }

                                    }
                                    else
                                    {
                                        object cellValue = "";
                                        lstDataMenu.Add(cellValue.ToString());
                                    }
                                }
                                else
                                {
                                    object cellValue = "";
                                    lstDataMenu.Add(cellValue.ToString());
                                }
                            }
                        }

                        form.pgbLoad.Value = 20;
                        form.OutputPrc(20, "Export Data: 20%");
                        //pgbLoad.Value = 20;
                        //lbPercent.Text = "Export Data: 20%";
                        //lbPercent.Refresh();

                        //Type C
                        grpC.Data_Type = "C";
                        //MessageBox.Show(sheet.Range["F7"].Value.Replace(" ", string.Empty));


                        //Console.WriteLine(RecursionTaxid(" " + sheet.Range[sellertaxid].Value.Replace(" ", string.Empty) + " "));

                        //grpC.Seller_Tax_ID = RecursionTaxid(" " + getvalue(sellertaxid).Replace(" ", string.Empty).Replace("-",string.Empty) + " ").Replace(" ", string.Empty); //เลขประจำตัวผู้เสียภาษี
                        //Console.WriteLine(sheet.Range[sellerbranchid.Split(',')[0]].Value.Replace(" ", string.Empty) + " sellerbranchid");
                        //MessageBox.Show("check");
                        grpC.Seller_Tax_ID = dtParam.SellerTaxID;
                        //grpC.Seller_Branch_ID = sheet.Range["L8"].Value.Replace(" ", string.Empty); //เลขสาขาประกอบการ
                        //grpC.Seller_Branch_ID = dtParam.BranchID;
                        //Console.WriteLine(getvalue(sellerbranchid).Replace(" ", string.Empty));

                        if (!sellerbranchid.Equals(""))
                        {
                            try
                            {
                                if (!getvalue(sellerbranchid).Equals(""))
                                {

                                    grpC.Seller_Branch_ID = branch_seller((sheet.Range[sellerbranchid.Split(',')[0]].Value.Replace(" ", string.Empty)).ToString());
                                }
                                else
                                {
                                    grpC.Seller_Branch_ID = "00000";
                                }
                            }
                            catch (IndexOutOfRangeException e)
                            {
                                Console.WriteLine(e.Message);
                            }
                            catch (NullReferenceException e)
                            {
                                Console.WriteLine(e.Message);
                            }

                        }
                        else
                        {
                            grpC.Seller_Branch_ID = "00000";
                        }

                        Console.WriteLine(grpC.Seller_Branch_ID + "sellerbranchid");

                        //grpC.File_Name = RecursionTaxid(" " + getvalue(sellertaxid).Replace(" ", string.Empty) + " ").Replace(" ", string.Empty) + ".txt"; //ชื่อไฟล์
                        grpC.File_Name = grpC.Seller_Tax_ID + ".txt";
                        form.pgbLoad.Value = 30;
                        form.OutputPrc(30, "Export Data: 30%");
                        grpB.Data_Type = "B";
                        //int iComSplit = getvalue(buyer_name).IndexOf("(");
                        //if (iComSplit != -1)
                        //{
                        //    grpB.Buyer_Name = getvalue(buyer_name).Substring(0, iComSplit - 1); //CompanyName
                        //}
                        //else
                        //{
                        //    grpB.Buyer_Name = getvalue(buyer_name).Replace(" ", string.Empty); //CompanyName
                        //}
                        grpB.Buyer_Name = getvalue(buyer_name);
                        grpB.Buyer_Phone_No = "";
                        try
                        {
                            strTaxID = RecursionTaxid(" " + getvalue(buyer_tax_id).Replace(" ", string.Empty).Replace("-", string.Empty) + " "); //ประเภทผู้เสียภาษี

                            strTaxID = strTaxID.Replace(" ", string.Empty);
                        }
                        catch (NullReferenceException e)
                        {
                            strTaxID = "N/A";
                        }
                        catch (Exception e)
                        {
                            strTaxID = "";
                        }
                        if (!buyer_branch_id.Equals(""))
                        {
                            try
                            {
                                if (!getvalue(buyer_branch_id).Equals(""))
                                {
                                    string buyerbrach_String = branch_buyyer(getvalue(buyer_branch_id).Replace(" ", string.Empty));
                                    grpB.Buyer_Branch_ID = buyerbrach_String;
                                }
                                else
                                {
                                    grpB.Buyer_Branch_ID = "";
                                }
                            }
                            catch (IndexOutOfRangeException e)
                            {
                                grpB.Buyer_Branch_ID = "";
                            }
                            catch (Exception e)
                            {
                                grpB.Buyer_Branch_ID = "";
                            }
                        }
                        else
                        {
                            grpB.Buyer_Branch_ID = "";
                        }

                        try
                        {
                            if (getvalue(buyer_countrypostcode) == "")
                            {
                                grpB.Buyer_Post_Code = ("00000");
                            }
                            else
                            {
                                try
                                {
                                    RealValue = "";
                                    grpB.Buyer_Post_Code = RecursionPostCode(" " + getvalue(buyer_countrypostcode) + " ");
                                }
                                catch (IndexOutOfRangeException e)
                                {
                                    grpB.Buyer_Post_Code = "00000";
                                }
                                catch (Exception e)
                                {
                                    grpB.Buyer_Post_Code = "00000";
                                }
                            }
                        }
                        catch (Exception e)
                        {
                            grpB.Buyer_Post_Code = "";
                        }
                        string keyType = string.Empty;
                        if (strTaxID.Equals("N/A"))
                        {
                            keyType = "4";
                        }
                        else if (strTaxID.Equals(""))
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
                            else if (Char.TryParse(countTaxNum.ToString().Substring(0, 1), out char data) && (grpB.Buyer_Branch_ID == null || grpB.Buyer_Branch_ID.Equals(""))) //อนาคตหากมีเลขที่ PassPort
                            {
                                keyType = "3";
                            }
                            else
                            {
                                keyType = "4";
                                strTaxID = getvalue(buyer_tax_id).Replace(" ", string.Empty).Replace("-", string.Empty);
                            }

                        }
                        grpB.Buyer_Tax_ID_Type = BuyerTaxType[keyType];
                        if (!typebuyer.Equals(""))
                        {
                            if (!getvalue(typebuyer).Equals(""))
                            {
                                grpB.Buyer_Tax_ID_Type = getvalue(typebuyer).Replace(" ", string.Empty).Replace("-", string.Empty);
                                strTaxID = getvalue(buyer_tax_id).Replace(" ", string.Empty).Replace("-", string.Empty);
                            }
                        }
                        if (strTaxID.Equals(""))
                        {
                            grpB.Buyer_Tax_ID = DoubleQuote(" "); //เลขที่ประจำตัวผู้เสียภาษี
                        }
                        else
                        {
                            grpB.Buyer_Tax_ID = strTaxID; //เลขที่ประจำตัวผู้เสียภาษี
                        }

                        if (buyer_uriid.Equals(""))
                        {
                            grpB.Buyer_URIID = "";
                        }
                        else
                        {
                            grpB.Buyer_URIID = getvalue(buyer_uriid).Replace(" ", string.Empty);
                        }
                        grpB.Buyer_Add_Line1 = getvalue(buyer_address);
                        grpB.Buyer_Add_Line2 = "";
                        form.pgbLoad.Value = 40;
                        form.OutputPrc(40, "Export Data: 40%");
                        int iCountRound = 0;

                        List<LGroup> lstGrpL = new List<LGroup>();
                        int distantNameItem = (AllChar.IndexOf(description.Split(',')[0]) - AllChar.IndexOf(noitem.Split(',')[0]));
                        int distantPriceUnit = (AllChar.IndexOf(priceunit.Split(',')[0]) - AllChar.IndexOf(noitem.Split(',')[0]));
                        int distantQuanlity = (AllChar.IndexOf(quanlity.Split(',')[0]) - AllChar.IndexOf(noitem.Split(',')[0]));
                        int distantAmount = (AllChar.IndexOf(amount.Split(',')[0]) - AllChar.IndexOf(noitem.Split(',')[0]));
                        int distantDiscount = (AllChar.IndexOf(discount.Split(',')[0]) - AllChar.IndexOf(noitem.Split(',')[0]));
                        for (int x = 0; x < lstDataMenu.Count; x++)
                        {
                            bool chkSting = false;
                            bool chkNum = false;
                            Double value = 0;
                            string patternChkString = @"([a-zA-Zก-๙0-9])";
                            if (!lstDataMenu[x].Equals(""))
                            {
                                chkSting = Regex.IsMatch(lstDataMenu[x], patternChkString);
                                chkNum = Double.TryParse(lstDataMenu[x], out value);
                            }
                            if (chkSting == true && chkNum == true)
                            {
                                if (iCountRound > 0)
                                {
                                    Console.WriteLine("LengthOfProduct_Desc => " + grpL.Product_Desc.Length);
                                    if (grpL.Product_Desc == null || grpL.Product_Desc.Equals(""))
                                    {
                                        grpL.Product_Desc = "";
                                    }
                                    else
                                    {

                                        if (grpL.Product_Desc.Length > 256)
                                        {
                                            string a = grpL.Product_Desc.Substring(0, 256);
                                            Console.WriteLine("a => " + a);
                                            string[] b = a.Split(' ');
                                            Console.WriteLine(b.Length);
                                            for (int i = 0; i < b.Length - 1; i++)
                                            {
                                                a_with_b += b[i] + " ";
                                            }
                                            Console.WriteLine("a_With_b => " + a_with_b);
                                            Console.WriteLine("LengthOfa_with_b => " + a_with_b.Length);
                                            grpL.Product_Remark = DoubleQuote(grpL.Product_Desc.Substring(a_with_b.Length));
                                            Console.WriteLine("Product_Remark => " + grpL.Product_Remark);
                                            grpL.Product_Desc = a_with_b;
                                            Console.WriteLine("Product_Desc => " + grpL.Product_Desc);
                                        }
                                        grpL.Product_Desc = (grpL.Product_Desc).Replace(",", DoubleQuote(","));
                                    }
                                    lstGrpL.Add(grpL);
                                }

                                grpL = new LGroup();
                                try
                                {
                                    grpL.Data_Type = DoubleQuote("L"); //ประเภทรายการ
                                    grpL.Line_ID = DoubleQuote(lstDataMenu[x]); //ลำดับรายการ
                                    grpL.Product_ID = DoubleQuote(""); //รหัสสินค้า
                                    grpL.Product_Name = lstDataMenu[x + distantNameItem].Replace(" ", string.Empty).Replace(",", DoubleQuote(",")); //ชื่อสินค้า
                                    grpL.Product_Desc = "";
                                    grpL.Product_Batch_ID = DoubleQuote(""); //ครั้งที่ผลิต
                                    grpL.Product_Expire_Dtm = DoubleQuote(""); //วันหมดอายุ
                                    grpL.Product_Class_Code = DoubleQuote(""); //รหัสหมวดหมู่สินค้า
                                    grpL.Product_Class_Name = DoubleQuote(""); //ชื่อหมวดหมู่สินค้า
                                    grpL.Product_OriCountry_ID = DoubleQuote(""); //รหัสประเทศกำเนิด
                                    try
                                    {
                                        grpL.Product_Charge_Amount = DoubleQuote(Double.Parse(RemoveComma(lstDataMenu[x + distantPriceUnit])).ToString("0.00")); //ราคาต่อหน่วย
                                    }
                                    catch (FormatException e)
                                    {
                                        grpL.Product_Charge_Amount = DoubleQuote("");
                                    }
                                    grpL.Product_Charge_Curr_Code = DoubleQuote("THB"); //รหัสสกุลเงิน (ราคาต่อหน่วย)
                                    grpL.Product_Al_Charge_IND = DoubleQuote(""); //ตัวบอกส่วนลดหรือค่าธรรมเนียม
                                    if (!discount.Equals(""))
                                    {
                                        try
                                        {
                                            grpL.Product_Al_Actual_Amount = DoubleQuote(Double.Parse(lstDataMenu[x + distantQuanlity]).ToString("0.00")); //มูลค่าส่วนลดหรือค่าธรรมเนียม
                                        }
                                        catch (FormatException e)
                                        {
                                            grpL.Product_Al_Actual_Amount = DoubleQuote("");
                                        }
                                        grpL.Product_Al_Actual_Curr_Code = DoubleQuote("THB"); //รหัสสกุลเงิน (มูลค่าส่วนลดหรือค่าธรรมเนียม)
                                    }
                                    else
                                    {
                                        grpL.Product_Al_Actual_Amount = DoubleQuote(""); //มูลค่าส่วนลดหรือค่าธรรมเนียม
                                        grpL.Product_Al_Actual_Curr_Code = DoubleQuote(""); //รหัสสกุลเงิน (มูลค่าส่วนลดหรือค่าธรรมเนียม)
                                    }
                                    grpL.Product_Al_Reason_Code = DoubleQuote(""); //รหัสเหตุผลในการคิดส่วนลดหรือค่าธรรมเนียม
                                    grpL.Product_Al_Reason = DoubleQuote(""); //เหตุผลในการคิดสวนลดหรือค่าธรรมเนียม

                                    try
                                    {
                                        grpL.Product_Quantity = DoubleQuote(Double.Parse(lstDataMenu[x + distantQuanlity]).ToString("0.00")); //จำนวนสินค้า
                                    }
                                    catch (FormatException e)
                                    {
                                        grpL.Product_Quantity = DoubleQuote("");
                                    }
                                    grpL.Product_Unit_Code = DoubleQuote(""); //รหัสหน่วยสินค้า
                                    grpL.Product_Quan_Per_Unit = DoubleQuote("1"); //ขนาดบรรจุต่อหน่วยขาย
                                    grpL.Line_Tax_Type_Code = DoubleQuote("VAT"); //รหัสประเภทภาษี
                                    grpL.Line_Tax_Cal_Rate = DoubleQuote("7.00"); //อัตราภาษี
                                                                                  //MessageBox.Show(lstDataMenu[x + 6]);
                                    grpL.Line_Basis_Amount = lstDataMenu[x + distantAmount]; //มูลค่าสินค้า/บริการ (ไม่รวมภาษีมูลค่าเพิ่ม)
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
                                    x = x + (AllChar.Length - 1);
                                }
                                catch (ArgumentOutOfRangeException e)
                                {

                                }
                                catch (NullReferenceException e)
                                {

                                }
                                catch (IndexOutOfRangeException e)
                                {

                                }
                            }
                            else
                            {
                                if (!lstDataMenu[x].Equals(""))
                                {
                                    grpL.Product_Desc += " " + lstDataMenu[x] /*lstDataMenu[x + 1]*/;
                                }
                            }
                        }

                        if (grpL.Product_Desc == null || grpL.Product_Desc.Equals(""))
                        {
                            grpL.Product_Desc = DoubleQuote("");
                        }
                        else
                        {


                            if (grpL.Product_Desc.Length > 256)
                            {
                                string a = grpL.Product_Desc.Substring(0, 256);
                                Console.WriteLine("a => " + a);
                                string[] b = a.Split(' ');
                                Console.WriteLine(b.Length);
                                for (int i = 0; i < b.Length - 1; i++)
                                {
                                    a_with_b += b[i] + " ";
                                }
                                Console.WriteLine("a_With_b => " + a_with_b);
                                Console.WriteLine("LengthOfa_with_b => " + a_with_b.Length);
                                grpL.Product_Remark = DoubleQuote(grpL.Product_Desc.Substring(a_with_b.Length));
                                Console.WriteLine("Product_Remark => " + grpL.Product_Remark);
                                grpL.Product_Desc = a_with_b;
                            }
                            grpL.Product_Desc = (grpL.Product_Desc).Replace(",", DoubleQuote(","));
                        }

                        lstGrpL.Add(grpL);

                        form.pgbLoad.Value = 50;
                        form.OutputPrc(50, "Export Data: 50%");

                        //Type F
                        form.pgbLoad.Value = 60;
                        form.OutputPrc(60, "Export Data: 60%");
                        //Type H
                        string[] arrKey = new string[] { "เลขที่ใบสั่งซื้อ :", "วันที่ใบสั่งซื้อ :" };
                        int[] arrIndex = new int[2];
                        int countArr = 0;
                        try
                        {
                            try
                            {
                                if (!pfIO.TypeDoc.Equals(""))
                                {
                                    grpH.Doc_Type_Code = pfIO.TypeDoc;
                                }
                                else
                                {

                                    if (!typedoc.Equals(""))
                                    {

                                        grpH.Doc_Type_Code = instring(DocType_ENG_AND_CODE, typedoc.Replace(" ", string.Empty));
                                    }
                                    else
                                    {
                                        Console.WriteLine(typedoc + " typedoc");
                                        Console.WriteLine(getvalue(document_name) + " getvalue(document_name)");
                                        grpH.Doc_Type_Code = instring(DocType, getvalue(document_name).Replace(" ", string.Empty));
                                    }
                                }


                            }
                            catch (KeyNotFoundException e)
                            {
                                Console.WriteLine("ไม่พบชื่อตัวแปรที่ส่งมา จึงเกิด error ");
                            }
                            Console.WriteLine(getvalue(document_name).Replace(" ", string.Empty) + " getvalue(document_name).Replace");
                            grpH.Doc_Name = getvalue(document_name).Replace(" ", string.Empty);
                            grpH.Doc_ID = getvalue(document_id).Replace(" ", string.Empty);
                            if (DateTime.TryParse(getvalue(document_issue_dtm), out dateValue))
                            {
                                string[] arrDate = getvalue(document_issue_dtm).Split('/');
                                string year = DateTime.Now.Year.ToString();
                                string yearFront = strYear.Substring(0, 2);
                                grpH.Doc_Issue_Dtm = years + "-" + Month[arrDateSplit[1].Replace(" ", string.Empty)] + "-" + Day[arrDateSplit[0].Replace(" ", string.Empty)] + "T00:00:00";
                            }
                            else
                            {
                                grpH.Doc_Issue_Dtm = getvalue(document_issue_dtm);
                            }

                        }
                        catch (ArgumentException e)
                        {
                            Console.WriteLine(e.Message);
                        }
                        catch (NullReferenceException e)
                        {
                            Console.WriteLine(e.Message);
                        }
                        catch (IndexOutOfRangeException e)
                        {
                            Console.WriteLine(e.Message);
                        }
                        catch (KeyNotFoundException e)
                        {
                            Console.WriteLine(e.Message + "=>" + "KeyNotFoundException");
                        }
                        if (!additional_ref_assign_id.Equals(""))
                        {
                            grpH.Add_Ref_Assign_ID = getvalue(additional_ref_assign_id).Replace(" ", string.Empty);
                        }
                        else
                        {
                            grpH.Add_Ref_Assign_ID = "";

                        }
                        if (!additional_ref_issue_dtm.Equals(""))
                        {
                            try
                            {
                                if (DateTime.TryParse(getvalue(document_issue_dtm), out dateValue))
                                {
                                    arrDateSplit = getvalue(additional_ref_issue_dtm).Split('/');
                                    grpH.Add_Ref_Issue_Dtm = years + "-" + Month[arrDateSplit[1].Replace(" ", string.Empty)] + "-" + Day[arrDateSplit[0].Replace(" ", string.Empty)] + "T00:00:00";
                                }
                                else
                                {
                                    grpH.Add_Ref_Issue_Dtm = getvalue(additional_ref_issue_dtm);
                                }
                            }
                            catch (ArgumentException e)
                            {
                                Console.WriteLine(e.Message);
                            }
                            catch (NullReferenceException e)
                            {
                                Console.WriteLine(e.Message);
                            }
                            catch (IndexOutOfRangeException e)
                            {
                                Console.WriteLine(e.Message);
                            }
                            catch (KeyNotFoundException e)
                            {
                                Console.WriteLine("aasdsasd");
                            }
                        }
                        else
                        {
                            grpH.Add_Ref_Issue_Dtm = "";
                        }

                        //MessageBox.Show(grpH.Add_Ref_Assign_ID);
                        if (!grpH.Add_Ref_Assign_ID.Equals(""))
                        {
                            grpH.Add_Ref_Type_Code = grpH.Doc_Type_Code;
                        }
                        else
                        {
                            grpH.Add_Ref_Type_Code = "";
                        }
                        try
                        {
                            if (create_purpose_code.Equals(""))
                            {
                                switch (grpH.Doc_Type_Code)
                                {
                                    case "388":
                                        try
                                        {
                                            grpH.Create_Purpose_Code = instring(TIVCPurpose, getvalue(create_purpose).Replace(" ", string.Empty));
                                            if (grpH.Create_Purpose_Code.Substring(grpH.Create_Purpose_Code.Length - 2) == "99")
                                            {
                                                grpH.Create_Purpose = getvalue(create_purpose).Replace(" ", string.Empty);
                                            }
                                        }
                                        catch (KeyNotFoundException ex)
                                        {
                                            grpH.Create_Purpose_Code = "TIVC99";
                                            grpH.Create_Purpose = getvalue(create_purpose).Replace(" ", string.Empty);
                                        }
                                        break;
                                    case "T02":
                                        try
                                        {
                                            grpH.Create_Purpose_Code = instring(TIVCPurpose, getvalue(create_purpose).Replace(" ", string.Empty));
                                            if (grpH.Create_Purpose_Code.Substring(grpH.Create_Purpose_Code.Length - 2) == "99")
                                            {
                                                grpH.Create_Purpose = getvalue(create_purpose).Replace(" ", string.Empty);
                                            }
                                        }
                                        catch (KeyNotFoundException ex)
                                        {
                                            grpH.Create_Purpose_Code = "TIVC99";
                                            grpH.Create_Purpose = getvalue(create_purpose).Replace(" ", string.Empty);
                                        }
                                        break;
                                    case "T03":
                                        try
                                        {
                                            grpH.Create_Purpose_Code = instring(TIVCPurpose, getvalue(create_purpose).Replace(" ", string.Empty));
                                            if (grpH.Create_Purpose_Code.Substring(grpH.Create_Purpose_Code.Length - 2) == "99")
                                            {
                                                grpH.Create_Purpose = getvalue(create_purpose).Replace(" ", string.Empty);
                                            }
                                        }
                                        catch (KeyNotFoundException ex)
                                        {
                                            grpH.Create_Purpose_Code = "TIVC99";
                                            grpH.Create_Purpose = getvalue(create_purpose).Replace(" ", string.Empty);
                                        }
                                        break;
                                    case "T04":
                                        try
                                        {
                                            grpH.Create_Purpose_Code = instring(TIVCPurpose, getvalue(create_purpose).Replace(" ", string.Empty));
                                            if (grpH.Create_Purpose_Code.Substring(grpH.Create_Purpose_Code.Length - 2) == "99")
                                            {
                                                grpH.Create_Purpose = getvalue(create_purpose).Replace(" ", string.Empty);
                                            }
                                        }
                                        catch (KeyNotFoundException ex)
                                        {
                                            grpH.Create_Purpose_Code = "TIVC99";
                                            grpH.Create_Purpose = getvalue(create_purpose).Replace(" ", string.Empty);
                                        }
                                        break;
                                    case "T01":
                                        try
                                        {
                                            grpH.Create_Purpose_Code = instring(RCTCPurpose, getvalue(create_purpose).Replace(" ", string.Empty));
                                            if (grpH.Create_Purpose_Code.Substring(grpH.Create_Purpose_Code.Length - 2) == "99")
                                            {
                                                grpH.Create_Purpose = getvalue(create_purpose).Replace(" ", string.Empty);
                                            }

                                        }
                                        catch (KeyNotFoundException ex)
                                        {
                                            grpH.Create_Purpose_Code = "RCTC99";
                                            grpH.Create_Purpose = getvalue(create_purpose).Replace(" ", string.Empty);
                                        }
                                        break;
                                    case "80":
                                        try
                                        {
                                            grpH.Create_Purpose_Code = instring(DBNGPurpose, getvalue(create_purpose).Replace(" ", string.Empty));
                                            if (grpH.Create_Purpose_Code.Substring(grpH.Create_Purpose_Code.Length - 2) == "99")
                                            {
                                                grpH.Create_Purpose = getvalue(create_purpose).Replace(" ", string.Empty);
                                            }
                                        }
                                        catch (KeyNotFoundException ex)
                                        {
                                            grpH.Create_Purpose_Code = "DBNG99";
                                            grpH.Create_Purpose = getvalue(create_purpose).Replace(" ", string.Empty);
                                        }
                                        break;
                                    case "81":
                                        try
                                        {
                                            grpH.Create_Purpose_Code = instring(CDNGPurpose, getvalue(create_purpose).Replace(" ", string.Empty));
                                            if (grpH.Create_Purpose_Code.Substring(grpH.Create_Purpose_Code.Length - 2) == "99")
                                            {
                                                grpH.Create_Purpose = getvalue(create_purpose).Replace(" ", string.Empty);
                                            }
                                        }
                                        catch (KeyNotFoundException ex)
                                        {
                                            grpH.Create_Purpose_Code = "CDNG99";
                                            grpH.Create_Purpose = getvalue(create_purpose).Replace(" ", string.Empty);
                                        }
                                        break;
                                    default:
                                        grpH.Create_Purpose_Code = "";
                                        grpH.Create_Purpose = "";
                                        break;
                                }
                            }
                            else
                            {
                                grpH.Create_Purpose_Code = getvalue(create_purpose_code);
                                grpH.Create_Purpose = getvalue(create_purpose);
                            }
                        }
                        catch (NullReferenceException e)
                        {
                            Console.WriteLine(e.Message);
                        }
                        catch (IndexOutOfRangeException e)
                        {
                            grpH.Create_Purpose_Code = "";
                            grpH.Create_Purpose = "";
                        }
                        catch (KeyNotFoundException e)
                        {
                            grpH.Create_Purpose_Code = "";
                            grpH.Create_Purpose = "";
                        }

                        if (!document_remark.Equals(""))
                        {
                            grpH.DOCUMENT_REMARK = getvalue(document_remark);
                        }
                        else
                        {
                            grpH.DOCUMENT_REMARK = "";
                        }


                        if (!buyer_order_assign_id.Equals(""))
                        {
                            //grpH.Buyer_Order_Assign_ID = getvalue(additional_ref_assign_id).Replace(" ", string.Empty);
                            grpH.Buyer_Order_Assign_ID = getvalue(buyer_order_assign_id).Replace(" ", string.Empty);
                        }
                        else
                        {
                            grpH.Buyer_Order_Assign_ID = "";
                        }


                        if (!buyer_order_issue_dtm.Equals(""))
                        {
                            try
                            {
                                arrDateSplit = getvalue(buyer_order_issue_dtm).Split('/');
                                grpH.Buyer_Order_Issue_Dtm = years + "-" + Month[arrDateSplit[1].Replace(" ", string.Empty)] + "-" + Day[arrDateSplit[0].Replace(" ", string.Empty)] + "T00:00:00";
                            }
                            catch (ArgumentException e)
                            {
                                Console.WriteLine(e.Message);
                            }
                            catch (NullReferenceException e)
                            {
                                Console.WriteLine(e.Message);
                            }
                            catch (IndexOutOfRangeException e)
                            {
                                Console.WriteLine(e.Message);
                            }
                            catch (KeyNotFoundException e)
                            {
                                Console.WriteLine(e.Message);
                            }
                        }
                        else
                        {
                            grpH.Buyer_Order_Issue_Dtm = "";
                        }


                        if (grpH.Buyer_Order_Assign_ID.Equals(""))
                        {
                            grpH.Buyer_Order_Ref_Type_Code = "";
                        }
                        else
                        {
                            grpH.Buyer_Order_Ref_Type_Code = "ON";
                        }

                        try
                        {
                            if (!original_total_amount.Equals(""))
                            {

                                grpF.Original_Total_Amount = Double.Parse(getvalue(original_total_amount)).ToString("0.00");
                                grpF.Original_Total_Curr_Code = "THB";
                            }
                            else
                            {
                                grpF.Original_Total_Amount = "";
                                grpF.Original_Total_Curr_Code = "";
                            }
                        }
                        catch (FormatException e)
                        {
                            Console.WriteLine(e.Message);
                        }
                        catch (NullReferenceException e)
                        {
                            Console.WriteLine(e.Message);
                        }

                        try
                        {
                            if (!line_total_amount.Equals(""))
                            {
                                grpF.LINE_TOTAL_AMOUNT = Double.Parse(getvalue(line_total_amount)).ToString("0.00");
                                grpF.LINE_TOTAL_CURRENCY_CODE = "THB";
                            }
                            else
                            {
                                grpF.LINE_TOTAL_AMOUNT = "";
                                grpF.LINE_TOTAL_CURRENCY_CODE = "";
                            }
                        }
                        catch (FormatException e)
                        {
                            Console.WriteLine(e.Message);
                        }
                        catch (NullReferenceException e)
                        {
                            Console.WriteLine(e.Message);
                        }

                        try
                        {
                            if (!adjusted_information_amount.Equals(""))
                            {
                                grpF.Adjusted_Inform_Amount = Double.Parse(getvalue(adjusted_information_amount)).ToString("0.00");
                                grpF.Adjusted_Inform_Curr_Code = "THB";
                            }
                            else
                            {
                                grpF.Adjusted_Inform_Amount = "";
                                grpF.Adjusted_Inform_Curr_Code = "";
                            }
                        }
                        catch (FormatException e)
                        {
                            Console.WriteLine(e.Message);
                        }
                        catch (NullReferenceException e)
                        {
                            Console.WriteLine(e.Message);
                        }

                        try
                        {
                            if (!allowance_total_amount.Equals(""))
                            {
                                grpF.Al_Total_Amount = Double.Parse(getvalue(allowance_total_amount)).ToString("0.00");
                                grpF.Al_Total_Curr_Code = "THB";
                            }
                            else
                            {
                                grpF.Al_Total_Amount = "";
                                grpF.Al_Total_Curr_Code = "";
                            }
                        }
                        catch (FormatException e)
                        {
                            Console.WriteLine(e.Message);
                        }
                        catch (NullReferenceException e)
                        {
                            Console.WriteLine(e.Message);
                        }
                        if (!countrybuyer.Equals(""))
                        {
                            if (!getvalue(countrybuyer).Equals(""))
                            {
                                grpB.Buyer_Country_ID = getvalue(countrybuyer).Replace(" ", string.Empty).Replace("-", string.Empty);
                                if (grpB.Buyer_Country_ID != "TH")
                                {
                                    grpB.Buyer_Post_Code = "";
                                }
                            }
                            else
                            {
                                grpB.Buyer_Country_ID = "TH";
                            }
                        }
                        else
                        {
                            grpB.Buyer_Country_ID = "TH";
                        }
                        List<string> lstC = new List<string> { DoubleQuote("C"),
                                            DoubleQuote(grpC.Seller_Tax_ID.Replace(" ",string.Empty)), //เลขที่ประจำตัวผู้เสียภาษี
                                            DoubleQuote(grpC.Seller_Branch_ID), //เลขสาขาประกอบการ
                                            DoubleQuote(grpC.File_Name.Replace(" ",string.Empty)), //ชื่อไฟล์  
                                            };
                        //MessageBox.Show("lstC:Success");
                        List<string> lstH = new List<string> { DoubleQuote("H"),
                                            DoubleQuote(grpH.Doc_Type_Code), //ประเภทเอกสาร 
                                            DoubleQuote(grpH.Doc_Name), //ชื่อเอกสาร
                                            DoubleQuote(grpH.Doc_ID), // เลขที่เอกสาร
                                            DoubleQuote(grpH.Doc_Issue_Dtm), //วันที่
                                            DoubleQuote(grpH.Create_Purpose_Code), //สาเหตุการออกเอกสาร
                                            DoubleQuote(grpH.Create_Purpose), //กรณีระบุสาเหตุเอกสาร
                                            DoubleQuote(grpH.Add_Ref_Assign_ID), //เลขที่เอกสารอ้างอิง
                                            DoubleQuote(grpH.Add_Ref_Issue_Dtm), //เอกสารอ้างอิงลงวันที่
                                            DoubleQuote(grpH.Add_Ref_Type_Code), //ประเภทเอกสารอ้างอิง
                                            DoubleQuote(""), //ชื่อเอกสารอ้างอิง 
                                            DoubleQuote(""), //เงื่อนไขการส่งของ
                                            DoubleQuote(grpH.Buyer_Order_Assign_ID), //เลขที่ใบสั่งซื้อ
                                            DoubleQuote(grpH.Buyer_Order_Issue_Dtm), //วันเดือนปีที่ออกใบสั่งซื้อ
                                            DoubleQuote(grpH.Buyer_Order_Ref_Type_Code), //ประเภทเอกสารอ้างอิงการสั่งซื้อ
                                            DoubleQuote(grpH.DOCUMENT_REMARK) //หมายเหตุท้ายเอกสาร
                                            };
                        form.pgbLoad.Value = 70;
                        form.OutputPrc(70, "Export Data: 70%");

                        List<string> lstB = new List<string> { DoubleQuote("B"),
                                            DoubleQuote(""), //รหัสผู้ซื้อ
                                            DoubleQuote(grpB.Buyer_Name), //ชื่อผู้ซื้อ
                                            DoubleQuote(grpB.Buyer_Tax_ID_Type), //ประเภทผู้เสียภาษี
                                            DoubleQuote(grpB.Buyer_Tax_ID.Replace(" ",string.Empty)), //เลขประจำตัวผู้เสียภาษี
                                            DoubleQuote(grpB.Buyer_Branch_ID.Replace(" ",string.Empty)), //เลขที่สาขา
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
                                            DoubleQuote(grpB.Buyer_Country_ID) //รหัสประเทศ
                                            };

                        try
                        {
                            if (!totalamount.Equals(""))
                            {
                                totalamount = RemoveComma(double.Parse(getvalue(totalamount)).ToString("0.00"));
                            }
                            else
                            {
                                totalamount = "";
                            }
                        }
                        catch (IndexOutOfRangeException e)
                        {
                            Console.WriteLine(e.Message);
                        }
                        catch (NullReferenceException e)
                        {
                            Console.WriteLine(e.Message);
                        }

                        try
                        {
                            if (!vat.Equals(""))
                            {
                                vat = RemoveComma(double.Parse(getvalue(vat)).ToString("0.00"));
                            }
                            else
                            {
                                vat = "";
                            }
                        }
                        catch (IndexOutOfRangeException e)
                        {
                            Console.WriteLine(e.Message);
                        }
                        catch (NullReferenceException e)
                        {
                            Console.WriteLine(e.Message);
                        }

                        try
                        {
                            if (!tax_basis_total_amount.Equals(""))
                            {
                                tax_basis_total_amount = RemoveComma(double.Parse(getvalue(tax_basis_total_amount)).ToString("0.00"));
                            }
                            else
                            {
                                tax_basis_total_amount = "";
                            }
                        }
                        catch (IndexOutOfRangeException e)
                        {
                            Console.WriteLine(e.Message);
                        }
                        catch (NullReferenceException e)
                        {
                            Console.WriteLine(e.Message);
                        }

                        try
                        {
                            if (!total.Equals(""))
                            {
                                total = RemoveComma(double.Parse(getvalue(total)).ToString("0.00"));
                            }
                            else
                            {
                                total = "";
                            }
                        }
                        catch (IndexOutOfRangeException e)
                        {
                            Console.WriteLine(e.Message);
                        }
                        catch (NullReferenceException e)
                        {
                            Console.WriteLine(e.Message);
                        }
                        try
                        {
                            if (vat_rate == "" || vat_rate == null)
                            {
                                vat_rate = "7.00";
                            }
                            else
                            {
                                vat_rate = ConvertNumber(vat_rate);
                            }
                        }
                        catch (Exception ex)
                        {
                            vat_rate = "7.00";
                        }
                        List<string> lstF = new List<string> { DoubleQuote("F"),
                                                DoubleQuote(String.Format("{0:00000}", lstGrpL.Count).ToString()), //จำนวนรายการสินค้า
                                                DoubleQuote(""), //วันเวลานัดส่งสินค้า
                                                DoubleQuote("THB"), //รหัสสกุลเงินตรา
                                                DoubleQuote("VAT"), //รหัสประเภทภาษี
                                                DoubleQuote(vat_rate), //อัตราภาษี
                                                //DoubleQuote(RemoveComma(sumAmount.ToString("N2"))), //มูลค่าสินค้า(ไม่รวมภาษีมูลค่าเพิ่ม)2350
                                                DoubleQuote(totalamount),
                                                DoubleQuote("THB"),
                                                //DoubleQuote(RemoveComma(sumTaxAmount.ToString("N2"))), //มูลค่าภาษีมูลค่าเพิ่ม
                                                DoubleQuote(vat),
                                                DoubleQuote("THB"),
                                                DoubleQuote(""), //ตัวบอกส่วนลดหรือค่าธรรมเนียม
                                                DoubleQuote(""), //มูลค่าส่วนลดหรือค่าธรรมเนียม
                                                DoubleQuote(""),
                                                DoubleQuote(""), //รหัสเหตุผลในการคิดส่วนลดหรือค่าธรรมเนียม
                                                DoubleQuote(""), //เหตุผลในการคิดส่วนลดหรือค่าธรรมเนียม
                                                DoubleQuote(""), //รหัสประเภทส่วนลด     
                                                DoubleQuote(""), //รายละเอียดเงื่อนไขการชำระเงิน
                                                DoubleQuote(""), //วันครบกำหนดชำระเงิน
                                                DoubleQuote(grpF.Original_Total_Amount), //รวมมูลค่าตามเอกสารเดิม
                                                DoubleQuote(grpF.Original_Total_Curr_Code),
                                                //DoubleQuote(RemoveComma(sumAmount.ToString("N2"))),
                                                DoubleQuote(totalamount),
                                                DoubleQuote("THB"),
                                                DoubleQuote(grpF.Adjusted_Inform_Amount), //มูลค่าผลต่าง
                                                DoubleQuote(grpF.Adjusted_Inform_Curr_Code),
                                                DoubleQuote(grpF.Al_Total_Amount), //ส่วนลดทั้งหมด
                                                DoubleQuote(grpF.Al_Total_Curr_Code),
                                                DoubleQuote(""), //ค่าธรรมเนียมทั้งหมด
                                                DoubleQuote(""),
                                                //DoubleQuote(RemoveComma(sumAmount.ToString("N2"))), //มูลค่าที่นำมาคิดภาษีมูลค่าเพิ่ม
                                                DoubleQuote(tax_basis_total_amount),
                                                DoubleQuote("THB"),
                                                //DoubleQuote(RemoveComma(sumTaxAmount.ToString("N2"))), //จำนวนภาษีมูลค่าเพิ่ม
                                                DoubleQuote(vat),
                                                DoubleQuote("THB"),
                                                //DoubleQuote(RemoveComma(sumGrandTotal.ToString("N2"))), //จำนวนเงินรวม(รวมภาษีมูลค่าเพิ่ม)
                                                DoubleQuote(total),
                                                DoubleQuote("THB")
                                                };

                        List<string> lstT = new List<string> { DoubleQuote("T"),
                                                DoubleQuote("1") //จำนวนเอกสารทั้งหมด
                                                };
                        form.pgbLoad.Value = 80;
                        form.OutputPrc(80, "Export Data: 80%");
                        Console.WriteLine("a");
                        string messageText = String.Join(",", lstC) + "\r"
                                + String.Join(",", lstH) + "\r"
                                + String.Join(",", lstB) + "\r";
                        Console.WriteLine("b");
                        for (int k = 0; k < lstGrpL.Count; k++)
                        {
                            messageText += lstGrpL[k].Data_Type + "," + lstGrpL[k].Line_ID + "," + lstGrpL[k].Product_ID + "," + lstGrpL[k].Product_Name + "," + lstGrpL[k].Product_Desc + ","
                                + lstGrpL[k].Product_Batch_ID + "," + lstGrpL[k].Product_Expire_Dtm + "," + lstGrpL[k].Product_Class_Code + "," + lstGrpL[k].Product_Class_Name + "," + lstGrpL[k].Product_OriCountry_ID + ","
                                + lstGrpL[k].Product_Charge_Amount + "," + lstGrpL[k].Product_Charge_Curr_Code + "," + lstGrpL[k].Product_Al_Charge_IND + "," + lstGrpL[k].Product_Al_Actual_Amount + "," + lstGrpL[k].Product_Al_Actual_Curr_Code + ","
                                + lstGrpL[k].Product_Al_Reason_Code + "," + lstGrpL[k].Product_Al_Reason + "," + lstGrpL[k].Product_Quantity + "," + lstGrpL[k].Product_Unit_Code + "," + lstGrpL[k].Product_Quan_Per_Unit + ","
                                + lstGrpL[k].Line_Tax_Type_Code + "," + lstGrpL[k].Line_Tax_Cal_Rate + "," + lstGrpL[k].Line_Basis_Amount + "," + lstGrpL[k].Line_Basis_Curr_Code + "," + lstGrpL[k].Line_Tax_Cal_Amount + ","
                                + lstGrpL[k].Line_Tax_Cal_Curr_Code + "," + lstGrpL[k].Line_AL_Charge_IND + "," + lstGrpL[k].Line_AL_Actual_Amount + "," + lstGrpL[k].Line_AL_Actual_Curr_Code + "," + lstGrpL[k].Line_AL_Reason_Code + ","
                                + lstGrpL[k].Line_AL_Reason + "," + lstGrpL[k].Line_Tax_Total_Amount + "," + lstGrpL[k].Line_Tax_Total_Curr_Code + "," + lstGrpL[k].Line_Net_Total_Amount + "," + lstGrpL[k].Line_Net_Total_Curr_Code + ","
                                + lstGrpL[k].Line_Net_Include_Amount + "," + lstGrpL[k].Line_Net_Include_Curr_Code + "," + lstGrpL[k].Product_Remark + "\r";
                        }
                        messageText += String.Join(",", lstF) + "\r"
                                + String.Join(",", lstT);
                        pathText = dtParam.PathOutput + "\\" + "BCP" + "_" + strFileName + "_" + strDateTimeStamp + ".txt";
                        CreateTextFile(pathText, messageText);
                        form.txtStatus.Text += Environment.NewLine + "   -Convert Success!";

                        strTempLogTime += " Convert Success!";
                        //txtStatus.Refresh();
                        form.Outputmessage(txtstr);
                        /*- BCPSERVICE -*/
                        getModelViladateSign iteminput = new getModelViladateSign();
                        iteminput.AccessKey = dtParam.AccessKey;
                        iteminput.APIKey = dtParam.APIKey;
                        iteminput.SellerBranchId = dtParam.BranchID;
                        iteminput.ServiceCode = dtParam.ServiceCode;
                        iteminput.UserCode = dtParam.UserCode;
                        iteminput.SellerTaxId = dtParam.SellerTaxID;
                        iteminput.TextContent = pathText;
                        BCP_Output = conAPIETAX_Viladate.ViladateSignAPI(iteminput);
                        Console.WriteLine(BCP_Output);


                    }
                    catch (FileNotFoundException e)
                    {
                        MessageBox.Show("File Not Found ConfigExcel => " + dtParam.PathConfigExcel);
                        goto loop;
                    }
                    catch (IOException e)
                    {
                        MessageBox.Show("กรุณาปิดไฟล์ Excel ทั้งหมด");
                        goto loop;
                    }

                }

                form.pgbLoad.Value = 90;
                form.OutputPrc(90, "Export Data: 90%");
                FileVersionInfo fvi = FileVersionInfo.GetVersionInfo(assembly.Location);
                string version = fvi.FileVersion;
                if (BCP_Output.MessageResultError != null && !BCP_Output.MessageResultError.Equals(""))
                //if (strOutputFile.MessageResultError == "")
                {
                    JObject oKeepResponeExecute = new JObject();
                    //MessageBox.Show(strOutputFile.MessageResultError);

                    if (BCP_Output.MessageResultError == "{}")
                    {
                        BCP_Output.MessageResultError = "กรุณาตรวจสอบอินเตอร์เน็ต!!";
                    }
                    form.txtStatus.Text += Environment.NewLine + "   -********** Viladate Fail!" + " **********";
                    strTempLogTime += " Viladate etax.one.th Fail!" + " Version " + version;
                    //Console.WriteLine(pathIn + " " + pathOutput);
                    oKeepResponeExecute = JObject.Parse(BCP_Output.MessageResultError.ToString());

                    //sock.SendMailAlert(dtParam.PathInput, dtParam.PathOutput, "FE99", form.emailtxt.Text, form.txtSellerTaxID.Text, Path.GetFileName(dtParam.PathInput).Replace("~$", string.Empty) + "-//-" + oKeepResponeExecute["errorCode"], oKeepResponeExecute["errorMessage"].ToString().Replace(" ",string.Empty));
                    //sock.SendMailAlert(pathFileIO.PathInput, pathFileIO.PathErr, "F01", form.emailtxt.Text, form.txtSellerTaxID.Text, pathIn + Path.GetFileName(item).Replace("~$", string.Empty), pathIn + Path.GetFileName(item).Replace("~$", string.Empty));
                    cntFail++;
                }
                else
                {

                    form.txtStatus.Text += Environment.NewLine + "   -Viladate etax.one.th Success!" + "";
                    strTempLogTime += ", Viladate etax.one.th Success!" + " Version " + version;
                }

                Console.WriteLine(chkOption);
                //txtStatus.Refresh();
                if (chkOption == true)
                {
                    if (BCP_Output.StatusCallAPI == false)
                    {

                        string pathErr = dtParam.PathOutput + "\\" + Path.GetFileNameWithoutExtension(pathText) + "_Error.txt";
                        JObject oKeepResponeExecute = new JObject();
                        oKeepResponeExecute = JObject.Parse(strOutputFile.MessageResultError.ToString());
                        Console.WriteLine(strOutputFile + " oKeepResponeExecute");
                        _apimail.err_code = "FE99";
                        _apimail.actionmsg = oKeepResponeExecute["errorMessage"].ToString().Replace(" ", string.Empty).Replace("\n", string.Empty).Replace(",", string.Empty).Replace("'", string.Empty);
                        _apimail.err_msg = Path.GetFileName(dtParam.PathInput).Replace("~$", string.Empty) + "-//-" + oKeepResponeExecute["errorCode"];
                        _apimail.input = dtParam.PathInput;
                        _apimail.path = dtParam.PathOutput;
                        _apimail.email = form.emailtxt.Text;
                        _apimail.taxseller = form.txtSellerTaxID.Text;
                        if (form.pingeng && oKeepResponeExecute["errorCode"].ToString() != "ER011")
                        {
                            _apimail.send_err_service();

                        }
                        if (oKeepResponeExecute["errorCode"].ToString() == "ER011")
                        {

                        }
                        else
                        {
                            Console.WriteLine(dtParam.PathInput);
                            CreateTextFile(pathErr, BCP_Output.MessageResultError);
                        }




                        //if (form.pingeng == true)
                        //{
                        //    sock.SendMailAlert(dtParam.PathInput, dtParam.PathOutput, "FE99", form.emailtxt.Text, form.txtSellerTaxID.Text, Path.GetFileName(dtParam.PathInput).Replace("~$", string.Empty) + "-//-" + oKeepResponeExecute["errorCode"], oKeepResponeExecute["errorMessage"].ToString());
                        //}


                    }
                    else
                    {
                        this.pathOutput = Path.GetFileNameWithoutExtension(pathText);
                        try
                        {
                            if (!Directory.Exists(dtParam.PathOutput + "\\" + "LogSucces"))
                            {
                                Directory.CreateDirectory(dtParam.PathOutput + "\\" + "LogSucces");
                            }
                        }
                        catch (Exception ea)
                        {
                            Console.WriteLine(ea);
                        }
                        finally
                        {
                            for (int i = 0; i <= GC.MaxGeneration; i++)
                            {
                                int count = GC.CollectionCount(i);
                                GC.Collect();
                            }
                            GC.WaitForPendingFinalizers();
                            GC.SuppressFinalize(this);
                        }
                        try
                        {
                            if (!Directory.Exists(dtParam.PathOutput + "\\" + "Log"))
                            {
                                Directory.CreateDirectory(dtParam.PathOutput + "\\" + "Log");
                            }
                        }
                        catch (Exception ea)
                        {
                            Console.WriteLine(ea);
                        }
                        finally
                        {
                            for (int i = 0; i <= GC.MaxGeneration; i++)
                            {
                                int count = GC.CollectionCount(i);
                                GC.Collect();
                            }
                            GC.WaitForPendingFinalizers();
                            GC.SuppressFinalize(this);
                        }
                        try
                        {
                            if (!Directory.Exists(dtParam.PathOutput + "\\" + "Temp_Succes"))
                            {
                                Directory.CreateDirectory(dtParam.PathOutput + "\\" + "Temp_Succes");
                            }
                        }
                        catch (Exception ea)
                        {
                            Console.WriteLine(ea);
                        }
                        finally
                        {
                            for (int i = 0; i <= GC.MaxGeneration; i++)
                            {
                                int count = GC.CollectionCount(i);
                                GC.Collect();
                            }
                            GC.WaitForPendingFinalizers();
                            GC.SuppressFinalize(this);
                        }
                        try
                        {
                            if (!Directory.Exists(pfIO.BCP_Folder + "\\" + "Succes"))
                            {
                                Directory.CreateDirectory(pfIO.BCP_Folder + "\\" + "Succes");
                            }
                        }
                        catch (Exception ea)
                        {
                            Console.WriteLine(ea);
                        }
                        finally
                        {
                            for (int i = 0; i <= GC.MaxGeneration; i++)
                            {
                                int count = GC.CollectionCount(i);
                                GC.Collect();
                            }
                            GC.WaitForPendingFinalizers();
                            GC.SuppressFinalize(this);
                        }
                        try
                        {
                            if (!Directory.Exists(pfIO.BCP_Folder + "\\" + "Fail"))
                            {
                                Directory.CreateDirectory(pfIO.BCP_Folder + "\\" + "Fail");
                            }
                        }
                        catch (Exception ea)
                        {
                            Console.WriteLine(ea);
                        }
                        finally
                        {
                            for (int i = 0; i <= GC.MaxGeneration; i++)
                            {
                                int count = GC.CollectionCount(i);
                                GC.Collect();
                            }
                            GC.WaitForPendingFinalizers();
                            GC.SuppressFinalize(this);
                        }
                        try
                        {
                            if (!Directory.Exists(dtParam.PathOutput + "\\" + "Log_Resend"))
                            {
                                Directory.CreateDirectory(dtParam.PathOutput + "\\" + "Log_Resend");
                            }
                        }
                        catch (Exception ea)
                        {
                            Console.WriteLine(ea);
                        }
                        finally
                        {
                            for (int i = 0; i <= GC.MaxGeneration; i++)
                            {
                                int count = GC.CollectionCount(i);
                                GC.Collect();
                            }
                            GC.WaitForPendingFinalizers();
                            GC.SuppressFinalize(this);
                        }

                        int counttimes = 0;
                        bool string_check__pdf;
                        bool string_check__xml;
                        JObject json_respo = new JObject();
                        json_respo = JObject.Parse(BCP_Output.ResponseMessage);
                        Console.WriteLine(json_respo + " json_respo");
                        pathPdf = Path.GetDirectoryName(dtParam.PathInput) + "\\" + Path.GetFileNameWithoutExtension(dtParam.PathInput) + ".pdf";
                        if (json_respo["status"].ToString() != "ER")
                        {
                            Console.WriteLine(pathPdf);
                        }
                        //    string Temp_succes = dtParam.PathOutput + "\\" + "Temp_Succes";
                        //    CreateTextFile(dtParam.PathOutput + "\\LogSucces\\" + "Success_" + Path.GetFileNameWithoutExtension(pathText) + ".txt", strOutputFile.Message_Content);
                        //download_pdfandxml:
                        //    counttimes = counttimes + 1;
                        //    DownloadFile(strOutputFile.MessageResultPDF, Temp_succes, this.pathOutput + "_PDF.pdf");
                        //    DownloadFile(strOutputFile.MessageResultXML, Temp_succes, this.pathOutput + "_XML.xml");
                        //    string_check__pdf = _checkfolder_pdf(this.pathOutput + "_PDF.pdf", Temp_succes);
                        //    string_check__xml = _checkfolder_xml(this.pathOutput + "_XML.xml", Temp_succes);
                        //    if (string_check__pdf == false && string_check__xml == false && counttimes <= 3)
                        //    {
                        //        goto download_pdfandxml;
                        //    }
                        //    else if (string_check__pdf == false && counttimes <= 3)
                        //    {
                        //        goto download_pdfandxml;
                        //    }
                        //    else if (string_check__xml == false && counttimes <= 3)
                        //    {
                        //        goto download_pdfandxml;
                        //    }

                        //    if (string_check__pdf == false && string_check__xml == false)
                        //    {

                        //        _apimail.actionmsg = "ไม่สามารถดาวน์โหลดไฟล์ PDF และ XML ได้";
                        //        if (form.pingeng)
                        //        {
                        //            _apimail.send_err_service();
                        //        }
                        //        CreateTextFile(dtParam.PathOutput + "\\Log\\" + "Log_" + Path.GetFileNameWithoutExtension(pathText) + ".txt", "ไม่สามารถดาวน์โหลดไฟล์ PDF และ XML ได้");
                        //    }
                        //    else if (string_check__pdf == false)
                        //    {
                        //        _apimail.actionmsg = "ไม่สามารถดาวน์โหลดไฟล์ PDF ได้";
                        //        if (form.pingeng)
                        //        {
                        //            _apimail.send_err_service();
                        //        }
                        //        CreateTextFile(dtParam.PathOutput + "\\Log\\" + "Log_" + Path.GetFileNameWithoutExtension(pathText) + ".txt", "ไม่สามารถดาวน์โหลดไฟล์ PDF ได้");
                        //    }
                        //    else if (string_check__xml == false)
                        //    {
                        //        _apimail.actionmsg = "ไม่สามารถดาวน์โหลดไฟล์ XML ได้";
                        //        if (form.pingeng)
                        //        {
                        //            _apimail.send_err_service();
                        //        }
                        //        CreateTextFile(dtParam.PathOutput + "\\Log\\" + "Log_" + Path.GetFileNameWithoutExtension(pathText) + ".txt", "ไม่สามารถดาวน์โหลดไฟล์ XML ได้");
                        //    }
                        //    try
                        //    {
                        //        var move_file_dis_form_pdf = Temp_succes + "\\" + this.pathOutput + "_PDF.pdf";
                        //        var move_file_dis_to_pdf = dtParam.PathOutput + "\\" + this.pathOutput + "_PDF.pdf";
                        //        var move_file_dis_form_xml = Temp_succes + "\\" + this.pathOutput + "_XML.xml";
                        //        var move_file_dis_to_xml = dtParam.PathOutput + "\\" + this.pathOutput + "_XML.xml";
                        //        if (string_check__pdf == true && string_check__xml == true)
                        //        {
                        //            File.Move(move_file_dis_form_pdf, move_file_dis_to_pdf);
                        //            File.Move(move_file_dis_form_xml, move_file_dis_to_xml);
                        //        }
                        //        else if (string_check__pdf == true && string_check__xml == false)
                        //        {
                        //            File.Move(move_file_dis_form_pdf, move_file_dis_to_pdf);
                        //        }
                        //        else if (string_check__pdf == false && string_check__xml == true)
                        //        {
                        //            File.Move(move_file_dis_form_xml, move_file_dis_to_xml);
                        //        }
                        //    }
                        //    catch (Exception ex)
                        //    {
                        //        Console.WriteLine(ex.Message);
                        //    }

                        //}
                        //else if (json_respo["errorCode"].ToString() == "ER011")
                        //{
                        //    string Temp_succes = dtParam.PathOutput + "\\" + "Temp_Succes";
                        //    CreateTextFile(dtParam.PathOutput + "\\" + "Log_Resend\\" + "Resend_Success_" + Path.GetFileNameWithoutExtension(pathText) + ".txt", strOutputFile.Message_Content);
                        //    //MessageBox.Show(json_respo.ToString());
                        //    int counttimes_resend = 0;
                        //    string pathText_resend = "RESEND_" + UATorPROD + "_" + Path.GetFileNameWithoutExtension(dtParam.PathInput) + "_" + strDateTimeStamp + ".txt";
                        //    string pathOutput_resend = Path.GetFileNameWithoutExtension(pathText_resend);
                        //download_pdf_resend:
                        //    counttimes_resend = counttimes_resend + 1;
                        //    DownloadFile(strOutputFile.MessageResultPDF, Temp_succes, pathOutput_resend + "_PDF.pdf");
                        //    DownloadFile(strOutputFile.MessageResultXML, Temp_succes, pathOutput_resend + "_XML.xml");
                        //    bool string_check_pdf_resend = _checkfolder_pdf(pathOutput_resend + "_PDF.pdf", Temp_succes);
                        //    bool string_check_xml_resend = _checkfolder_xml(pathOutput_resend + "_XML.xml", Temp_succes);
                        //    if (string_check_pdf_resend == false && string_check_xml_resend == false && counttimes_resend <= 3)
                        //    {
                        //        goto download_pdf_resend;
                        //    }
                        //    else if (string_check_pdf_resend == false && counttimes_resend <= 3)
                        //    {
                        //        goto download_pdf_resend;
                        //    }
                        //    else if (string_check_xml_resend == false && counttimes_resend <= 3)
                        //    {
                        //        goto download_pdf_resend;
                        //    }
                        //    if (string_check_pdf_resend == false && string_check_xml_resend == false)
                        //    {
                        //        CreateTextFile(dtParam.PathOutput + "\\Log\\" + "Log_" + Path.GetFileNameWithoutExtension(pathText) + ".txt", "ไม่สามารถดาวน์โหลดไฟล์ PDF และ XML ได้");
                        //    }
                        //    else if (string_check_pdf_resend == false)
                        //    {
                        //        CreateTextFile(dtParam.PathOutput + "\\Log\\" + "Log_" + Path.GetFileNameWithoutExtension(pathText) + ".txt", "ไม่สามารถดาวน์โหลดไฟล์ PDF ได้");
                        //    }
                        //    else if (string_check_xml_resend == false)
                        //    {
                        //        CreateTextFile(dtParam.PathOutput + "\\Log\\" + "Log_" + Path.GetFileNameWithoutExtension(pathText) + ".txt", "ไม่สามารถดาวน์โหลดไฟล์ XML ได้");
                        //    }

                        //    try
                        //    {
                        //        var move_file_dis_form_pdf = Temp_succes + "\\" + pathOutput_resend + "_PDF.pdf";
                        //        var move_file_dis_to_pdf = dtParam.PathOutput + "\\" + pathOutput_resend + "_PDF.pdf";
                        //        var move_file_dis_form_xml = Temp_succes + "\\" + pathOutput_resend + "_XML.xml";
                        //        var move_file_dis_to_xml = dtParam.PathOutput + "\\" + pathOutput_resend + "_XML.xml";
                        //        if (string_check_pdf_resend == true && string_check_xml_resend == true)
                        //        {
                        //            File.Move(move_file_dis_form_pdf, move_file_dis_to_pdf);
                        //            File.Move(move_file_dis_form_xml, move_file_dis_to_xml);
                        //        }
                        //        else if (string_check_pdf_resend == true && string_check_xml_resend == false)
                        //        {
                        //            File.Move(move_file_dis_form_pdf, move_file_dis_to_pdf);
                        //        }
                        //        else if (string_check_pdf_resend == false && string_check_xml_resend == true)
                        //        {
                        //            File.Move(move_file_dis_form_xml, move_file_dis_to_xml);
                        //        }
                        //    }
                        //    catch (Exception ex)
                        //    {
                        //        Console.WriteLine(ex.Message);
                        //    }

                        //}

                        ////CreateTextFile(dtParam.PathOutput, strOutputFile.Message_Content);
                        //Console.WriteLine(dtParam.PathOutput + " 3872");
                    }
                }
                else if (chkOption == false)
                {
                    if (BCP_Output.StatusCallAPI == false)
                    {
                        string pathErr = pfIO.PathErr + "\\" + Path.GetFileNameWithoutExtension(pathText) + "_Error.txt";
                        CreateTextFile(pathErr, strOutputFile.MessageResultError);
                        JObject oKeepResponeExecute1 = new JObject();
                        oKeepResponeExecute1 = JObject.Parse(strOutputFile.MessageResultError.ToString());
                        Console.WriteLine(strOutputFile + " oKeepResponeExecute1");
                        Console.WriteLine(Path.GetFileName(dtParam.PathInput).Replace("~$", string.Empty) + "-//-" + oKeepResponeExecute1["errorCode"]);
                        try
                        {
                            _apimail.err_code = "FE99";
                            _apimail.actionmsg = oKeepResponeExecute1["errorMessage"].ToString().Replace(" ", string.Empty).Replace("\n", string.Empty).Replace(",", string.Empty).Replace("'", string.Empty);
                            _apimail.err_msg = Path.GetFileName(dtParam.PathInput).Replace("~$", string.Empty) + "-//-" + oKeepResponeExecute1["errorCode"];
                            _apimail.input = dtParam.PathInput;
                            _apimail.path = dtParam.PathOutput;
                            _apimail.email = form.emailtxt.Text;
                            _apimail.taxseller = form.txtSellerTaxID.Text;

                            if (form.pingeng)
                            {
                                _apimail.send_err_service();
                            }
                            Console.WriteLine(dtParam.PathInput);
                            CreateTextFile(pathErr, strOutputFile.MessageResultError);

                            //if (form.metroToggle1.Checked == true && form.pingeng == true)
                            //{
                            //    sock.SendMailAlert(dtParam.PathInput, dtParam.PathOutput, "FE99", form.emailtxt.Text, form.txtSellerTaxID.Text, Path.GetFileName(dtParam.PathInput).Replace("~$", string.Empty) + "-//-" + oKeepResponeExecute1["errorCode"], oKeepResponeExecute1["errorMessage"].ToString().Replace(" ", string.Empty).Replace("\n", string.Empty).Replace(",", string.Empty).Replace("'", string.Empty));
                            //}
                        }
                        catch (Exception ea)
                        {
                            Console.WriteLine(ea);
                        }

                        string[] arrFiles = System.IO.Directory.GetFiles(pfIO.PathTemp, "*.txt");
                        string[] arrFilesSource = System.IO.Directory.GetFiles(pfIO.PathSource_F, "*.txt");

                        foreach (var item in arrFiles)
                        {
                            string fileName = Path.GetFileName(item);
                            this.nameFilePDF = item;
                            string pathTxtNew = pfIO.PathSource_F + "\\" + fileName;
                            string pathTxtNew_S = pfIO.PathSource_S + "\\" + fileName;
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
                            File.Copy(item, pathTxtNew_S);
                            File.Move(item, pathTxtNew);
                        }
                    }
                    else
                    {
                        string fileNameWithoutExtension = string.Empty;
                        string[] arrFiles = System.IO.Directory.GetFiles(pfIO.PathTemp, "*.txt");
                        string[] arrFiles__pcfg = System.IO.Directory.GetFiles(pfIO.PathInput, "*.pcfg");

                        this.pathOutput = Path.GetFileNameWithoutExtension(pathText);
                        try
                        {
                            if (!Directory.Exists(pfIO.PathSuccess_O + "\\" + "LogSucces"))
                            {
                                Directory.CreateDirectory(pfIO.PathSuccess_O + "\\" + "LogSucces");
                            }
                        }
                        catch (Exception ea)
                        {
                            Console.WriteLine(ea);
                        }
                        finally
                        {
                            for (int i = 0; i <= GC.MaxGeneration; i++)
                            {
                                int count = GC.CollectionCount(i);
                                GC.Collect();
                            }
                            GC.WaitForPendingFinalizers();
                            GC.SuppressFinalize(this);
                        }
                        try
                        {
                            if (!Directory.Exists(pfIO.PathSuccess_O + "\\" + "Log"))
                            {
                                Directory.CreateDirectory(pfIO.PathSuccess_O + "\\" + "Log");
                            }
                        }
                        catch (Exception ea)
                        {
                            Console.WriteLine(ea);
                        }
                        finally
                        {
                            for (int i = 0; i <= GC.MaxGeneration; i++)
                            {
                                int count = GC.CollectionCount(i);
                                GC.Collect();
                            }
                            GC.WaitForPendingFinalizers();
                            GC.SuppressFinalize(this);
                        }
                        try
                        {
                            if (!Directory.Exists(pfIO.BCP_Folder + "\\" + "Success"))
                            {
                                Directory.CreateDirectory(pfIO.BCP_Folder + "\\" + "Success");
                            }
                        }
                        catch (Exception ea)
                        {
                            Console.WriteLine(ea);
                        }
                        finally
                        {
                            for (int i = 0; i <= GC.MaxGeneration; i++)
                            {
                                int count = GC.CollectionCount(i);
                                GC.Collect();
                            }
                            GC.WaitForPendingFinalizers();
                            GC.SuppressFinalize(this);
                        }
                        try
                        {
                            if (!Directory.Exists(pfIO.BCP_Folder + "\\" + "Fail"))
                            {
                                Directory.CreateDirectory(pfIO.BCP_Folder + "\\" + "Fail");
                            }
                        }
                        catch (Exception ea)
                        {
                            Console.WriteLine(ea);
                        }
                        finally
                        {
                            for (int i = 0; i <= GC.MaxGeneration; i++)
                            {
                                int count = GC.CollectionCount(i);
                                GC.Collect();
                            }
                            GC.WaitForPendingFinalizers();
                            GC.SuppressFinalize(this);
                        }
                        try
                        {
                            if (!Directory.Exists(pfIO.PathErr + "\\" + "Log_Resend"))
                            {
                                Directory.CreateDirectory(pfIO.PathErr + "\\" + "Log_Resend");
                            }
                        }
                        catch (Exception ea)
                        {
                            Console.WriteLine(ea);
                        }
                        finally
                        {
                            for (int i = 0; i <= GC.MaxGeneration; i++)
                            {
                                int count = GC.CollectionCount(i);
                                GC.Collect();
                            }
                            GC.WaitForPendingFinalizers();
                            GC.SuppressFinalize(this);
                        }
                        try
                        {
                            if (!Directory.Exists(pfIO.PathSuccess_O + "\\" + "Temp_Succes"))
                            {
                                Directory.CreateDirectory(pfIO.PathSuccess_O + "\\" + "Temp_Succes");
                            }
                        }
                        catch (Exception ea)
                        {
                            Console.WriteLine(ea);
                        }
                        finally
                        {
                            for (int i = 0; i <= GC.MaxGeneration; i++)
                            {
                                int count = GC.CollectionCount(i);
                                GC.Collect();
                            }
                            GC.WaitForPendingFinalizers();
                            GC.SuppressFinalize(this);
                        }
                        int counttimes = 0;
                        bool string_check__pdf;
                        bool string_check__xml;
                        _apimail.err_code = "FE91";
                        _apimail.err_msg = strFileNameExtension.Replace("~$", string.Empty) + "-//-" + "";
                        _apimail.input = dtParam.PathInput;
                        _apimail.path = dtParam.PathOutput;
                        _apimail.email = form.emailtxt.Text;
                        _apimail.taxseller = form.txtSellerTaxID.Text;
                        JObject ok_json = new JObject();
                        ok_json = JObject.Parse(BCP_Output.ResponseMessage);
                        pathPdf = Path.GetDirectoryName(dtParam.PathInput) + "\\" + Path.GetFileNameWithoutExtension(dtParam.PathInput) + ".pdf";

                        string path_BCP_Success = pfIO.BCP_Folder + "\\" + "Success";
                        if (ok_json["status"].ToString() != "ER")
                        {
                            try
                            {
                                File.Move(pathPdf, path_BCP_Success + "\\" + "BCP_" + Path.GetFileNameWithoutExtension(dtParam.PathInput) + ".pdf");
                            }
                            catch (Exception ex)
                            {
                                Console.WriteLine(ex.Message);
                            }
                            try
                            {
                                File.Move(pathText, path_BCP_Success + "\\" + "BCP_" + Path.GetFileNameWithoutExtension(dtParam.PathInput) + ".txt");
                            }
                            catch (Exception ex)
                            {
                                Console.WriteLine(ex.Message);
                            }
                        }

                        //if (ok_json["status"].ToString() != "ER")
                        //{
                        //    string Temp_succes = pfIO.PathSuccess_O + "\\" + "Temp_Succes";
                        //    CreateTextFile(pfIO.PathSuccess_O + "\\LogSucces\\" + "Success_" + Path.GetFileNameWithoutExtension(pathText) + ".txt", strOutputFile.Message_Content);
                        //downloadpdfandxml:
                        //    counttimes = counttimes + 1;
                        //    DownloadFile(strOutputFile.MessageResultPDF, Temp_succes, this.pathOutput + "_PDF.pdf");
                        //    DownloadFile(strOutputFile.MessageResultXML, Temp_succes, this.pathOutput + "_XML.xml");
                        //    string_check__pdf = _checkfolder_pdf(this.pathOutput + "_PDF.pdf", Temp_succes);
                        //    string_check__xml = _checkfolder_xml(this.pathOutput + "_XML.xml", Temp_succes);
                        //    if (string_check__pdf == false && string_check__xml == false && counttimes <= 3)
                        //    {
                        //        goto downloadpdfandxml;
                        //    }
                        //    else if (string_check__pdf == false && string_check__xml == true && counttimes <= 3)
                        //    {
                        //        goto downloadpdfandxml;
                        //    }
                        //    else if (string_check__xml == false && string_check__pdf == true && counttimes <= 3)
                        //    {
                        //        goto downloadpdfandxml;
                        //    }
                        //    //MessageBox.Show(string_check__pdf.ToString());
                        //    if (string_check__pdf == false && string_check__xml == false)
                        //    {
                        //        _apimail.actionmsg = "ไม่สามารถดาวน์โหลดไฟล์ PDF และ XML ได้";
                        //        if (form.pingeng)
                        //        {
                        //            _apimail.send_err_service();
                        //        }
                        //        CreateTextFile(pfIO.PathSuccess_O + "\\Log\\" + "Log_" + Path.GetFileNameWithoutExtension(pathText) + ".txt", "ไม่สามารถดาวน์โหลดไฟล์ PDF และ XML ได้");
                        //    }
                        //    else if (string_check__pdf == false && string_check__xml == true)
                        //    {
                        //        _apimail.actionmsg = "ไม่สามารถดาวน์โหลดไฟล์ PDF ได้";
                        //        if (form.pingeng)
                        //        {
                        //            _apimail.send_err_service();
                        //        }
                        //        CreateTextFile(pfIO.PathSuccess_O + "\\Log\\" + "Log_" + Path.GetFileNameWithoutExtension(pathText) + ".txt", "ไม่สามารถดาวน์โหลดไฟล์ PDF ได้");
                        //    }
                        //    else if (string_check__xml == false && string_check__pdf == true)
                        //    {
                        //        _apimail.actionmsg = "ไม่สามารถดาวน์โหลดไฟล์ XML ได้";
                        //        if (form.pingeng)
                        //        {
                        //            _apimail.send_err_service();
                        //        }
                        //        CreateTextFile(pfIO.PathSuccess_O + "\\Log\\" + "Log_" + Path.GetFileNameWithoutExtension(pathText) + ".txt", "ไม่สามารถดาวน์โหลดไฟล์ XML ได้");
                        //    }
                        //    try
                        //    {
                        //        var move_file_dis_form_pdf = Temp_succes + "\\" + this.pathOutput + "_PDF.pdf";
                        //        var move_file_dis_to_pdf = pfIO.PathSuccess_O + "\\" + this.pathOutput + "_PDF.pdf";
                        //        var move_file_dis_form_xml = Temp_succes + "\\" + this.pathOutput + "_XML.xml";
                        //        var move_file_dis_to_xml = pfIO.PathSuccess_O + "\\" + this.pathOutput + "_XML.xml";
                        //        if (string_check__pdf == true && string_check__xml == true)
                        //        {
                        //            File.Move(move_file_dis_form_pdf, move_file_dis_to_pdf);
                        //            File.Move(move_file_dis_form_xml, move_file_dis_to_xml);
                        //        }
                        //        else if (string_check__pdf == true && string_check__xml == false)
                        //        {
                        //            File.Move(move_file_dis_form_pdf, move_file_dis_to_pdf);
                        //        }
                        //        else if (string_check__pdf == false && string_check__xml == true)
                        //        {
                        //            File.Move(move_file_dis_form_xml, move_file_dis_to_xml);
                        //        }
                        //    }
                        //    catch (Exception ex)
                        //    {
                        //        Console.WriteLine(ex.Message);
                        //    }


                        //}
                        //else if (ok_json["errorCode"].ToString() == "ER011")
                        //{
                        //    string Temp_succes = pfIO.PathSuccess_O + "\\" + "Temp_Succes";
                        //    CreateTextFile(pfIO.PathErr + "\\" + "Log_Resend\\" + "Resend_Success_" + Path.GetFileNameWithoutExtension(pathText) + ".txt", strOutputFile.Message_Content);
                        //    int counttimes_resend = 0;
                        //    string pathText_resend = "RESEND_" + UATorPROD + "_" + Path.GetFileNameWithoutExtension(dtParam.PathInput) + "_" + strDateTimeStamp + ".txt";
                        //    string pathOutput_resend = Path.GetFileNameWithoutExtension(pathText_resend);
                        //downloadpdfandxml_resend:
                        //    counttimes_resend = counttimes_resend + 1;
                        //    DownloadFile(strOutputFile.MessageResultPDF, Temp_succes, pathOutput_resend + "_PDF.pdf");
                        //    DownloadFile(strOutputFile.MessageResultXML, Temp_succes, pathOutput_resend + "_XML.xml");
                        //    bool string_check__pdf_resend = _checkfolder_pdf(pathOutput_resend + "_PDF.pdf", Temp_succes);
                        //    bool string_check__xml_resend = _checkfolder_xml(pathOutput_resend + "_XML.xml", Temp_succes);
                        //    if (string_check__pdf_resend == false && string_check__xml_resend == false && counttimes_resend <= 3)
                        //    {
                        //        goto downloadpdfandxml_resend;
                        //    }
                        //    else if (string_check__pdf_resend == false && string_check__xml_resend == true && counttimes_resend <= 3)
                        //    {
                        //        goto downloadpdfandxml_resend;
                        //    }
                        //    else if (string_check__xml_resend == false && string_check__pdf_resend == true && counttimes_resend <= 3)
                        //    {
                        //        goto downloadpdfandxml_resend;
                        //    }
                        //    if (string_check__pdf_resend == false && string_check__xml_resend == false)
                        //    {
                        //        CreateTextFile(pfIO.PathSuccess_O + "\\" + "Log\\" + "Log_" + Path.GetFileNameWithoutExtension(pathText) + ".txt", "ไม่สามารถดาวน์โหลดไฟล์ PDF และ XML ได้");
                        //    }
                        //    else if (string_check__pdf_resend == false && string_check__xml_resend == true)
                        //    {
                        //        CreateTextFile(pfIO.PathSuccess_O + "\\Log\\" + "Log_" + Path.GetFileNameWithoutExtension(pathText) + ".txt", "ไม่สามารถดาวน์โหลดไฟล์ PDF ได้");
                        //    }
                        //    else if (string_check__xml_resend == false && string_check__pdf_resend == true)
                        //    {
                        //        CreateTextFile(pfIO.PathSuccess_O + "\\Log\\" + "Log_" + Path.GetFileNameWithoutExtension(pathText) + ".txt", "ไม่สามารถดาวน์โหลดไฟล์ XML ได้");
                        //    }
                        //    try
                        //    {
                        //        var move_file_dis_form_pdf = Temp_succes + "\\" + pathOutput_resend + "_PDF.pdf";
                        //        var move_file_dis_to_pdf = pfIO.PathSuccess_O + "\\" + pathOutput_resend + "_PDF.pdf";
                        //        var move_file_dis_form_xml = Temp_succes + "\\" + pathOutput_resend + "_XML.xml";
                        //        var move_file_dis_to_xml = pfIO.PathSuccess_O + "\\" + pathOutput_resend + "_XML.xml";
                        //        if (string_check__pdf_resend == true && string_check__xml_resend == true)
                        //        {
                        //            File.Move(move_file_dis_form_pdf, move_file_dis_to_pdf);
                        //            File.Move(move_file_dis_form_xml, move_file_dis_to_xml);
                        //        }
                        //        else if (string_check__pdf_resend == true && string_check__xml_resend == false)
                        //        {
                        //            File.Move(move_file_dis_form_pdf, move_file_dis_to_pdf);
                        //        }
                        //        else if (string_check__pdf_resend == false && string_check__xml_resend == true)
                        //        {
                        //            File.Move(move_file_dis_form_xml, move_file_dis_to_xml);
                        //        }
                        //    }
                        //    catch (Exception ex)
                        //    {
                        //        Console.WriteLine(ex.Message);
                        //    }

                        //}



                        Thread.Sleep(1000);
                        string namefilepdf = Path.GetFileName(dtParam.PathInput);
                        etaxOneth_Printer.Class1 _printer = new etaxOneth_Printer.Class1();
                        if (pfIO.TypePrinting == "A" && form.check___copies.Checked == false)
                        {
                            var Timestamp = new DateTimeOffset(DateTime.UtcNow).ToUnixTimeSeconds();
                            timest_process = Int32.Parse(Timestamp.ToString()) - timest_process;
                            Console.WriteLine(timest_process + " timest_process");
                            CreateTextFile(pfIO.LogTimeProcess + "\\LogProcess_Print_" + Path.GetFileNameWithoutExtension(pathText) + ".txt", "เอกสาร " + Path.GetFileNameWithoutExtension(pathText) + " ใช้เวลาการประมวลผลทั้งหมดประมาณ " + timest_process + " วินาที");
                            //PrinterSettings.SetDefaultPrinter(pfIO.Printer);
                            //ProcessStartInfo printProcessInfo = new ProcessStartInfo()
                            //{
                            //    UseShellExecute = true,
                            //    Verb = "print",
                            //    CreateNoWindow = true,
                            //    FileName = pfIO.PathSuccess_O + "\\" + this.pathOutput + "_PDF.pdf",
                            //    //Arguments = printDialog1.PrinterSettings.PrinterName.ToString(),
                            //    WindowStyle = ProcessWindowStyle.Hidden
                            //};
                            //_printer.PrintMethod("C:\\Users\\JIRAYU-NB\\Documents\\FillTEST\\output\\Success\\UAT_0105561072420_03-4-62T17-32-15_PDF.pdf", "ApeosPort-IV C5570 16", 1);
                            //Console.WriteLine(pfIO.PathSuccess_O + "\\" + this.pathOutput + "_PDF.pdf" + " " + pfIO.Printer + " " + short.Parse(form.input_copies.Text));
                            _printer.PrintMethod(pfIO.PathSuccess_O + "\\" + this.pathOutput + "_PDF.pdf", pfIO.Printer, short.Parse(form.input_copies.Text));

                            //try
                            //{
                            //    Process printProcess = new Process();
                            //    printProcess.StartInfo = printProcessInfo;
                            //    printProcess.Start();
                            //    //Thread.Sleep(3000);
                            //    //if (printProcess.HasExited == false)
                            //    //{
                            //    //    printProcess.Kill();
                            //    //}
                            //}
                            //catch (Exception ex)
                            //{
                            //    //MessageBox.Show(ex.ToString());
                            //    //MessageBox.Show("ไม่พบตัวอ่านไฟล์ของคุณ");
                            //}
                        }
                        else if (pfIO.TypePrinting == "A" && form.check___copies.Checked == true)
                        {
                            if (arrFiles__pcfg.Count() != 0)
                            {
                                foreach (var item in arrFiles__pcfg)
                                {
                                    string namefile__ = Path.GetFileName(item);
                                    var Timestamp = new DateTimeOffset(DateTime.UtcNow).ToUnixTimeSeconds();
                                    timest_process = Int32.Parse(Timestamp.ToString()) - timest_process;
                                    Console.WriteLine(timest_process + " timest_process");
                                    CreateTextFile(pfIO.LogTimeProcess + "\\LogProcess_Print_" + Path.GetFileNameWithoutExtension(pathText) + ".txt", "เอกสาร " + Path.GetFileNameWithoutExtension(pathText) + " ใช้เวลาการประมวลผลทั้งหมดประมาณ " + timest_process + " วินาที");
                                    if (Path.GetFileNameWithoutExtension(namefilepdf) == Path.GetFileNameWithoutExtension(namefile__))
                                    {
                                        // Open the text file using a stream reader.
                                        using (StreamReader sr = new StreamReader(item))
                                        {
                                            // Read the stream to a string, and write the string to the console.
                                            String line = sr.ReadToEnd();
                                            bool string___checkinpcfg = checkcopiesin__pcfg(line.Replace(" ", string.Empty).Replace(Environment.NewLine, string.Empty).Replace("\t", string.Empty));
                                            if (string___checkinpcfg == true)
                                            {
                                                try
                                                {
                                                    _printer.PrintMethod(pfIO.PathSuccess_O + "\\" + this.pathOutput + "_PDF.pdf", pfIO.Printer, short.Parse(line));
                                                }
                                                catch (Exception ex)
                                                {
                                                    Console.WriteLine(ex);
                                                }
                                                finally
                                                {
                                                    sr.Close();
                                                    File.Delete(item);
                                                }
                                            }
                                            else if (string___checkinpcfg == false)
                                            {
                                                sr.Close();
                                                File.Delete(item);
                                                MessageBox.Show("ไม่สามารถปริ้นได้เนื่องจาก จำนวน Copies ไม่ถูกต้อง *ควรระบุ 1-99*",
                                                                            "แจ้งเตือน",
                                                                MessageBoxButtons.OK,
                                                                MessageBoxIcon.Error);
                                            }
                                        }
                                    }
                                    else
                                    {
                                        File.Delete(item);
                                        MessageBox.Show("ชื่อไฟล์ .pcfg ไม่ตรงกับไฟล์ที่นำเข้า",
                                                        "แจ้งเตือน",
                                                        MessageBoxButtons.OK,
                                                        MessageBoxIcon.Warning);
                                    }
                                }

                            }
                            else
                            {
                                MessageBox.Show("ไม่พบไฟล์ .pcfg",
                                                    "แจ้งเตือน",
                                                    MessageBoxButtons.OK,
                                                    MessageBoxIcon.Warning);
                            }



                            //for (int i = 0; i < arrFiles__pcfg.Length; i++)
                            //{
                            //    string namefile__cfg = Path.GetFileName(arrFiles__pcfg[i]);
                            //    MessageBox.Show(namefilepdf);
                            //    if(Path.GetFileNameWithoutExtension(namefilepdf) == Path.GetFileNameWithoutExtension(namefile__cfg))
                            //    {
                            //        try
                            //        {   // Open the text file using a stream reader.
                            //            using (StreamReader sr = new StreamReader(arrFiles__pcfg[i]))
                            //            {
                            //                // Read the stream to a string, and write the string to the console.
                            //                String line = sr.ReadToEnd();
                            //                bool string___checkinpcfg = checkcopiesin__pcfg(line.Replace(" ", string.Empty).Replace(Environment.NewLine,string.Empty));
                            //                if(string___checkinpcfg == true)
                            //                {
                            //                    try
                            //                    {
                            //                        _printer.PrintMethod(pfIO.PathSuccess_O + "\\" + this.pathOutput + "_PDF.pdf", pfIO.Printer, short.Parse(line));
                            //                    }
                            //                    catch(Exception ex)
                            //                    {
                            //                        Console.WriteLine(ex);
                            //                    }
                            //                    finally
                            //                    {
                            //                        sr.Close();
                            //                        File.Delete(arrFiles__pcfg[i]);
                            //                    }


                            //                }
                            //                else if(string___checkinpcfg == false)
                            //                {
                            //                    sr.Close();
                            //                    File.Delete(arrFiles__pcfg[i]);
                            //                    MessageBox.Show("ไม่สามารถปริ้นได้เนื่องจาก จำนวน Copies เกิน 99 แผ่น",
                            //                                                "แจ้งเตือน",
                            //                                                MessageBoxButtons.OK,
                            //                                                MessageBoxIcon.Error);

                            //                }
                            //                //MessageBox.Show(string___checkinpcfg);

                            //            }
                            //        }
                            //        catch (Exception e)
                            //        {
                            //            Console.WriteLine("The file could not be read:");
                            //            Console.WriteLine(e.Message);
                            //        }
                            //    }
                            //    else
                            //    {

                            //    }

                            //}

                            //if(namefilepdf.Split(',')[namefilepdf.Split(',').Length - 1] + ".pcfg" == )
                            //bool check__print = _checkfolder(namefilepdf.Split(',')[namefilepdf.Split(',').Length - 1] + ".pcfg", pfIO.PathSuccess_O);
                            //MessageBox.Show(check__print.ToString());
                        }
                    }

                }

                form.pgbLoad.Value = 100;
                form.OutputPrc(100, "Export Data: 100%");
                //pgbLoad.Value = 100;
                //lbPercent.Text = "Export Data: 100%";
                //lbPercent.Refresh();
            }
            catch (FileNotFoundException ex)
            {
                form.txtStatus.Text += Environment.NewLine + "   -**********ไม่พบ Path File ที่คุณระบุ กรุณาเลือก Path File ให้ถูกต้อง!**********";
                strTempLogTime += "ไม่พบ Path File ที่คุณระบุ กรุณาเลือก Path File ให้ถูกต้อง!";

                if (chkOption == true)
                {
                    CreateTextFile(dtParam.PathOutput + "\\" + strFileName + "_" + strDateTimeStamp + "_Error.txt", "ไม่พบ Path File ที่คุณระบุ กรุณาเลือก Path File ให้ถูกต้อง");
                }
                else
                {
                    CreateTextFile(pfIO.PathErr + "\\" + strFileName + "_" + strDateTimeStamp + "_Error.txt", "ไม่พบ Path File ที่คุณระบุ กรุณาเลือก Path File ให้ถูกต้อง");
                }

                form.txtStatus.Refresh();
                cntFail++;
                //MessageBox.Show("ไม่พบ Path File ที่คุณระบุ กรุณาเลือก Path File ให้ถูกต้อง");
            }
            catch (System.IndexOutOfRangeException e)
            {
                form.txtStatus.Text += Environment.NewLine + "   -**********ไฟล์ของคุณมีข้อผิดพลาดในข้อมูลที่ใส่!**********";
                strTempLogTime += "กรุณาตรวจสอบไฟล์ของคุณ!";

                if (chkOption == true)
                {
                    CreateTextFile(dtParam.PathOutput + "\\" + strFileName + "_" + strDateTimeStamp + "_Error.txt", "ไฟล์ของคุณมีข้อผิดพลาด กรุณาตรวจสอบและใส่ข้อมูลให้ถูกต้อง");
                    CreateTextFile(dtParam.PathOutput + "\\" + strFileName + "_" + strDateTimeStamp + "_ErrorServiceOrProgram.txt", e.Message + " " + e.Data);
                }
                else
                {
                    CreateTextFile(pfIO.PathErr + "\\" + strFileName + "_" + strDateTimeStamp + "_Error.txt", "ไฟล์ของคุณมีข้อผิดพลาด กรุณาตรวจสอบและใส่ข้อมูลให้ถูกต้อง");
                    CreateTextFile(pfIO.PathErr + "\\" + strFileName + "_" + strDateTimeStamp + "_ErrorServiceOrProgram.txt", e.Message + " " + e.Data);
                }

                form.txtStatus.Refresh();
                cntFail++;
            }
            catch (DirectoryNotFoundException ex)
            {
                Console.WriteLine(ex.Message);
            }
            catch (XmlException ex)
            {
                FileVersionInfo fvi = FileVersionInfo.GetVersionInfo(assembly.Location);
                string version = fvi.FileVersion;
                //MessageBox.Show("ไฟล์มีปัญหา!!!");
                Console.WriteLine(ex.Message);
                form.txtStatus.Text += Environment.NewLine + "   -**********Convert Fail!**********";
                strTempLogTime += " Convert Fail! " +ex.Message+" Version : " + version;
                //MessageBox.Show(ErrorMessage);
                if (chkOption == true)
                {
                    CreateTextFile(dtParam.PathOutput + "\\" + strFileName + "_" + strDateTimeStamp + "_Error.txt", "Convert Fail");

                }
                else
                {

                    CreateTextFile(dtParam.PathOutput + "\\" + strFileName + "_" + strDateTimeStamp + "_Error.txt", "Convert Fail");

                }

                form.txtStatus.Refresh();
                cntFail++;
            }
            catch (Exception ex)
            {

                //MessageBox.Show(ex.Message);
                FileVersionInfo fvi = FileVersionInfo.GetVersionInfo(assembly.Location);
                string version = fvi.FileVersion;
                form.txtStatus.Text += Environment.NewLine + "   -**********Convert Fail!**********";
                strTempLogTime += " Convert Fail! "+ex.Message+" Version : " + version;
                if (chkOption == true)
                {
                    CreateTextFile(dtParam.PathOutput + "\\" + strFileName + "_" + strDateTimeStamp + "_Error.txt", "Convert Fail");

                }
                else
                {

                    CreateTextFile(dtParam.PathOutput + "\\" + strFileName + "_" + strDateTimeStamp + "_Error.txt", "Convert Fail");

                }

                form.txtStatus.Refresh();
                cntFail++;
            }
            finally
            {
                lstTempSumAmount.Clear();
                for (int i = 0; i <= GC.MaxGeneration; i++)
                {
                    int count = GC.CollectionCount(i);
                    GC.Collect();
                }
                GC.WaitForPendingFinalizers();
                GC.SuppressFinalize(this);
            }
        }
        public bool _checkfolder_pdf(string namefile, string path)
        {
            DirectoryInfo di = new DirectoryInfo(path);
            FileInfo[] pdffiles = di.GetFiles("*.pdf");
            for (int i = 0; i < pdffiles.Length; i++)
            {
                if (namefile == pdffiles[i].ToString())
                {
                    return true;
                }
                Console.WriteLine();
            }
            return false;
        }
        public bool _checkfolder_xml(string namefile, string path)
        {
            DirectoryInfo di = new DirectoryInfo(path);
            FileInfo[] xmlfiles = di.GetFiles("*.xml");
            for (int i = 0; i < xmlfiles.Length; i++)
            {
                if (namefile == xmlfiles[i].ToString())
                {
                    return true;
                }
                Console.WriteLine();
            }
            return false;
        }
        public bool checkcopiesin__pcfg(string line)
        {
            Regex regex = new Regex(@"^(([1-8][0-9]?|9[0-9]?))$");
            if (regex.IsMatch(line))
            {
                return true;
            }
            else
            {
                return false;
            }
        }
        public static class PrinterSettings
        {
            [DllImport("winspool.drv",
              CharSet = CharSet.Auto,
              SetLastError = true)]
            [return: MarshalAs(UnmanagedType.Bool)]
            public static extern Boolean SetDefaultPrinter(String name);
        }
        public void DownloadFile(string strUrlFile, string pathOutputFile, string fileName)
        {
            using (var client = new WebClient())
            {
                try
                {
                    client.DownloadFile(strUrlFile, pathOutputFile + "/" + fileName);
                }
                catch (Exception ea)
                {
                    Console.WriteLine(ea);
                }
                finally
                {
                    for (int i = 0; i <= GC.MaxGeneration; i++)
                    {
                        int count = GC.CollectionCount(i);
                        GC.Collect();
                    }
                    GC.WaitForPendingFinalizers();
                    GC.SuppressFinalize(this);
                }

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
        public string FormatDecimal(string value)
        {
            value = value.Split('.')[0] + "." + value.Split('.')[1].Substring(0, 2);
            return value;
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

            return price.ToString("0.00");
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

            return price.ToString("0.00");
        }

        /*vat ของแต่ละสินค้า*/
        public string CalVatItem(string strAmount)
        {
            double total;
            double vatItem;
            double.TryParse(strAmount, out total);
            vatItem = total * 0.07;

            return vatItem.ToString("0.00");
        }

        public string CalVatItem_ListItem(string strAmount)
        {
            double total;
            double vatItem;
            double.TryParse(strAmount, out total);
            string total_sum;
            vatItem = (total * 7) / 100;
            total_sum = vatItem.ToString();
            total_sum = string.Format("{0:#,##0.00}", float.Parse(total_sum));
            return total_sum;
        }

        public void RefreshForm()
        {
            //etaxOnethProcess frm = new etaxOnethProcess();
            //frm.Refresh();
            //pbLogo.Refresh();
            //lbName.Refresh();
            //pbMinimize.Refresh();
            //pbRestoreDown.Refresh();
            //pbClose.Refresh();
            //pbMaximize.Refresh();
            //lbHeadName.Refresh();
            //pgbLoad.Refresh();
        }

        public void CreateTextFile(string pathConcatFileName, string strData)
        {
            //Write Text
            TextWriter txtWrite = new StreamWriter(pathConcatFileName);
            txtWrite.Write(strData);
            txtWrite.Close();
        }
        public string RecursionPostCode(string address)
        {
            int value;
            if (address.Length > 0 && address.Length >= 7)
            {
                string start = address.Substring(address.Length - address.Length, 1);
                string mid = address.Substring(1, 5);
                string end = address.Substring(address.Length - (address.Length - 6), 1);

                //Console.WriteLine("start = " + start + " && " + "mid = " + mid + " && " + "end = " + end + " && " + "adress.length = " + address.Length);
                //Console.Read();
                bool checkstart = int.TryParse(start, out value);
                bool checkend = int.TryParse(end, out value);
                bool checkmid = int.TryParse(mid, out value);
                if (checkstart == false && checkend == false && checkmid == true)
                {
                    RealValue = value.ToString();
                }

                return RecursionPostCode(address.Substring(1, address.Length - 1));
            }
            else
                return RealValue;
        }
        public string RecursionTaxid(string taxid)
        {
            long value;
            if (taxid.Length > 0 && taxid.Length >= 15)
            {
                string start = taxid.Substring(taxid.Length - taxid.Length, 1);
                string mid = taxid.Substring(1, 13);
                string end = taxid.Substring(taxid.Length - (taxid.Length - 14), 1);

                //Console.WriteLine("start = " + start + " && " + "mid = " + mid + " && " + "end = " + end + " && " + "taxid.length = " + taxid.Length);
                //Console.Read();
                bool checkstart = long.TryParse(start, out value);
                bool checkend = long.TryParse(end, out value);
                bool checkmid = long.TryParse(mid, out value);
                if (checkstart == false && checkend == false && checkmid == true)
                {
                    string values;
                    if (value.ToString().Length == 12)
                    {
                        values = "0" + value;
                        RealValue2 = values.ToString();
                    }
                    else
                    {
                        RealValue2 = value.ToString();
                    }
                }

                return RecursionTaxid(taxid.Substring(1, taxid.Length - 1));

            }
            else if (taxid.Length == 0)
                return RealValue2;
            else
                return taxid;
        }

        public string getvalue(string item)
        {
            string value = "";
            switch (item.Split(',')[1].ToLower())
            {
                case "value":
                    value = sheet.Range[item.Split(',')[0]].Value.ToString();
                    return value;
                case "formula":
                    value = sheet.Range[item.Split(',')[0]].FormulaValue.ToString();
                    return value;
            }
            return "";
        }

        public string instring(string[] data, string input)
        {
            for (int i = 0; i < data.Length - 2; i = i + 2)
            {
                Console.WriteLine(data[i]);
                if (input.Contains(data[i]))
                {
                    return data[i + 1];
                }
            }
            return data[data.Length - 1];
        }

        public string branch_seller(string branch_text)
        {
            string rgx = @"[0-9]{5}";
            string ret;
            Match result = Regex.Match(branch_text, rgx);
            if (result.Success)
            {
                ret = result.Value;
            }
            else
            {
                ret = "00000";
            }
            Console.WriteLine(ret);
            return ret;
        }

        public string branch_buyyer(string branch_text)
        {
            string rgx = @"[0-9]{5}";
            string ret;
            Match result = Regex.Match(branch_text, rgx);
            if (result.Success)
            {
                ret = result.Value;
            }
            else
            {
                ret = "00000";
            }
            Console.WriteLine(ret);
            return ret;
        }
        public string ColumnIndexToColumnLetter(int colIndex)
        {
            int div = colIndex;
            string colLetter = String.Empty;
            int mod = 0;

            while (div > 0)
            {
                mod = (div - 1) % 26;
                colLetter = (char)(65 + mod) + colLetter;
                div = (int)((div - mod) / 26);
            }
            return colLetter;
        }
        public string ConvertNumber(string numberic)
        {
            if (numberic == null)
            {
                return "";
            }
            else if (numberic == "0.00" || numberic == "0")
            {
                numberic = "0.00";
                return numberic;
            }
            else
            {
                numberic = string.Format("{0:#,##0.00}", float.Parse(numberic));
                return numberic;
            }
        }
        public string counext(string value_, int rows)
        {
            string res;
            res = value_.Split(',')[0].Split('-')[0] + (rows - 1) + "," + value_.Split(',')[1];
            return res;
        }
    }
}
