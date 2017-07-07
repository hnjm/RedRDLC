using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using SQL02;
using Microsoft.Reporting.WinForms;
using System.IO;

namespace BillingCreateAuto
{
    public class BillingCreate
    {
        // *************************************************************************************************
        // ****  This process will get a list of all reports to process for current date            ****
        // *************************************************************************************************
        public void createBill()
        {
            // get list of special reports to run for current day
            BillingEntities rptBilling = new BillingEntities();
            string rptBillPeriod = DateTime.Now.ToString("dd");
            var cltList = (from c in rptBilling.BillDeliveries where c.BillPeriod == rptBillPeriod && c.Active == true select c).ToList<BillDelivery>();
            if (cltList != null)
            {
                foreach (var clt in cltList)
                {
                    //DateTime? rptBillDate = GetBillDate(clt.CltID.Value);
                    DateTime? rptBillDate = DateTime.Now;
                    if (clt.Current == "Ahead")
                    {
                        rptBillDate = DateTime.Now.AddMonths(1);
                    }
                    string rptBillMonth = rptBillDate.Value.ToString("MM");
                    string rptBillYear = rptBillDate.Value.Year.ToString();
                    GetRptData(clt.CltID.Value, rptBillMonth, rptBillYear, clt.ID);
                }
            }
            // get list of master reports to run for current day
            var masterList = (from m in rptBilling.BillDeliveries where m.BillPeriod == rptBillPeriod && m.Active == true && m.ReportGroup == "Master" select m).ToList<BillDelivery>();
            if (masterList != null)
            {
                foreach (var mlst in masterList)
                {
                    //DateTime? rptBillDate = GetBillDate(mlst.CltID.Value);
                    DateTime? rptBillDate = DateTime.Now;
                    string rptBillMonth = rptBillDate.Value.ToString("MM");
                    string rptBillYear = rptBillDate.Value.Year.ToString();
                    GetRptData(mlst.CltID.Value, rptBillMonth, rptBillYear, mlst.ID);
                }
            }
        }
        // *************************************************************************************************
        // Process report for selected client for specific billdate
        // *************************************************************************************************
        public void createBill(int cltID, DateTime? billDate)
        {
            // get list of reports to run for Client
            BillingEntities cltBilling = new BillingEntities();
            string rptBillPeriod = billDate.Value.ToString("MM");
            DateTime? rptBillDate = billDate;
            string rptBillMonth = rptBillDate.Value.ToString("MM");
            string rptBillYear = rptBillDate.Value.Year.ToString();
            var rptList = (from c in cltBilling.BillDeliveries where c.CltID == cltID && c.Active == true select c).ToList<BillDelivery>();
            if (rptList != null)
            {
                foreach (var rpt in rptList)
                {
                    GetRptData(cltID, rptBillMonth, rptBillYear, rpt.ID);
                }
            }
            // get list of master reports to run for current day
            var masterList = (from m in cltBilling.BillDeliveries where m.CltID == cltID && m.Active == true && m.ReportGroup == "Master" select m).ToList<BillDelivery>();
            if (masterList != null)
            {
                foreach (var mlst in masterList)
                {
                    GetRptData(mlst.CltID.Value, rptBillMonth, rptBillYear, mlst.ID);
                }
            }
        }
        // *************************************************************************************************
        // ****  This process will get the data for the billing report                                  ****
        // ****  The bills that are seperated into seperate files must create empty bills               ****
        // ****  There are some bill where the Datasource name and storedProc used are different        ****
        // ****  need to find a way to process these bills for clients                                  ****
        // ****  165,633,177
        // *************************************************************************************************
        private void GetRptData(int rptcltID, string billMonth, string billYear, int billDeliveryID)
        {
            fileInfo rptfileInfo = GetBillFileInfo(billDeliveryID, rptcltID, billMonth, billYear);
            List<ReportParameter> rptparam = GetPBSParams(rptcltID, billMonth, billYear);
            var sepFile = rptfileInfo.rptReportGroup;
            BillingEntities billData = new BillingEntities();
            int qbillMonth = Convert.ToInt32(billMonth);
            int qbillYear = Convert.ToInt32(billYear);
            var tmprds = (from brpt in billData.tblBillReportDatas
                          where
                            brpt.CltID == rptcltID &&
                            brpt.BillDate.Value.Month == qbillMonth &&
                            brpt.BillDate.Value.Year == qbillYear
                          select brpt).ToList<tblBillReportData>();
            if (sepFile == "Location")
            {
                var cltFileSep = (from c in billData.v_CltLoc where c.CltID == rptcltID select c).ToList<v_CltLoc>();
                foreach (var l in cltFileSep)
                {
                    rptfileInfo.rptFileName = cleanFileName(l.LocName) + "_" + rptfileInfo.billToSend + Convert.ToString(billYear) + Convert.ToString(billMonth);
                    var temp = from c in tmprds where c.LocID == l.FileSepCD select c;
                    ReportDataSource rds = new ReportDataSource("BillReportData", temp);
                    if (temp.Count() > 0)
                    {
                        createRptFile(rds, rptcltID, rptfileInfo, rptparam);
                    }
                }
            }
            if (sepFile == "Class")
            {
                var cltFileSep = (from c in billData.v_CltClass where c.CltID == rptcltID select c).ToList<v_CltClass>();
                foreach (var l in cltFileSep)
                {
                    rptfileInfo.rptFileName = cleanFileName(l.ClassName);
                    var temp = from c in tmprds where c.ClassCD == l.FileSepCD select c;
                    ReportDataSource rds = new ReportDataSource("BillReportData", temp);
                    if (temp.Count() > 0)
                    {
                        createRptFile(rds, rptcltID, rptfileInfo, rptparam);
                    }

                }
            }
            if (sepFile == "Department")
            {
                var cltFileSep = (from c in billData.v_CltDept where c.CltID == rptcltID select c).ToList<v_CltDept>();
                foreach (var l in cltFileSep)
                {
                    rptfileInfo.rptFileName = cleanFileName(l.DeptNum);
                    var temp = from c in tmprds where c.DeptNum == l.FileSepCD select c;
                    ReportDataSource rds = new ReportDataSource("BillReportData", temp);
                    if (temp.Count() > 0)
                    {
                        createRptFile(rds, rptcltID, rptfileInfo, rptparam);
                    }
                }
            }
            if (sepFile == "Carrier")
            {
                var cltFileSep = (from c in billData.v_CltCar where c.CltID == rptcltID select c).ToList<v_CltCar>();
                foreach (var l in cltFileSep)
                {
                    rptfileInfo.rptFileName = cleanFileName(l.CarName) + "_" + cleanFileName(l.CltShortName) + "_" + Convert.ToString(billYear) + Convert.ToString(billMonth);
                    var temp = from c in tmprds where c.CarID == l.FileSepCD select c;
                    ReportDataSource rds = new ReportDataSource("BillReportData", temp);
                    if (temp.Count() > 0)
                    {
                        createRptFile(rds, rptcltID, rptfileInfo, rptparam);
                    }
                }
            }
            if (sepFile == "Master")
            {
                rptfileInfo.rptFileName = rptfileInfo.billToSend + Convert.ToString(billYear) + Convert.ToString(billMonth);
                var temp = (from brpt in billData.tblBillReportDatas
                            where
                              brpt.MasterCltID == rptcltID &&
                              brpt.BillDate.Value.Month == qbillMonth &&
                              brpt.BillDate.Value.Year == qbillYear
                            select brpt).ToList<tblBillReportData>();
                ReportDataSource rds = new ReportDataSource("BillReportData", temp);
                if (temp.Count() > 0)
                {
                    createRptFile(rds, rptcltID, rptfileInfo, rptparam);
                }
            }
            if (string.IsNullOrEmpty(sepFile))
            {
                if (tmprds.Any())
                {
                    // for regular bills if there is no data do not produce
                    rptfileInfo.rptFileName = rptfileInfo.billToSend + Convert.ToString(billYear) + Convert.ToString(billMonth);
                    ReportDataSource rds = new ReportDataSource("BillReportData", tmprds);
                    createRptFile(rds, rptcltID, rptfileInfo, rptparam);
                }
            }
        }
        // *************************************************************************************************
        // ****  This process will get the information on how to process the billing report             ****
        // *************************************************************************************************
        private fileInfo GetBillFileInfo(int BillDeliveryID, int rptcltID, string billMonth, string billYear)
        {
            SQL02Billing billRptfile = new SQL02Billing();
            var rptInfo = (from c in billRptfile.GetBillDeliveryPrintCtrl(BillDeliveryID) select c).FirstOrDefault();
            fileInfo finfo = new fileInfo();
            finfo.rdlcFilename = rptInfo.ReportName; //rdlc template name
            finfo.rptPath = Properties.Settings.Default.reportPath; // path of the rdlc template
            finfo.rptOutPath = Properties.Settings.Default.exportedReports + Convert.ToString(rptcltID) + "\\" + Convert.ToString(billYear) + Convert.ToString(billMonth) + "\\"; // path of the created file 
            finfo.rptReportGroup = rptInfo.ReportGroup; // how the files will be seperated
            finfo.billToSend = rptInfo.BillToSend;
            finfo.createNewFilename = rptInfo.ReportSeperateFile; // bool if the file should be in sep files
            finfo.rptFileExt = cleanFileName(rptInfo.FileExt);
            if (!Directory.Exists(@finfo.rptOutPath))
            {
                Directory.CreateDirectory(@finfo.rptOutPath);
            }
            return finfo;
        }
        // *************************************************************************************************
        // Set Params for the report
        // *************************************************************************************************
        private List<ReportParameter> GetPBSParams(int cltID, string billMonth, string billYear)
        {
            string rptLogoPath = Properties.Settings.Default.logoPath;
            string logoPath = string.Concat("file:///", rptLogoPath, cltID.ToString(), ".jpg");
            string BillMonth = string.Concat(System.Globalization.CultureInfo.CurrentCulture.DateTimeFormat.GetMonthName(Convert.ToInt32(billMonth)), " ", billYear);
            string AdjustmentMonth = BillMonth;
            List<ReportParameter> params1 = new List<ReportParameter>();
            params1.Add(new ReportParameter("Logo", logoPath));
            params1.Add(new ReportParameter("BillMonth", BillMonth));
            params1.Add(new ReportParameter("AdjustmentMonth", AdjustmentMonth));
            return params1;

        }
        // *************************************************************************************************
        // ****  This will get the last bill date for client                            ****
        // *************************************************************************************************
        private DateTime? GetBillDate(int rptcltID)
        {
            BillingEntities billRptDate = new BillingEntities();
            var rptBillDateInfo = (from x in billRptDate.BillingLogs
                                   where x.CltID == rptcltID
                                   orderby x.BatchID descending
                                   select x.BillDate).FirstOrDefault();
            return rptBillDateInfo;
        }
        // *************************************************************************************************
        // Create the Billing Report
        // *************************************************************************************************
        private void createRptFile(ReportDataSource rds, int cltID, fileInfo finfo, List<ReportParameter> rptParams)
        {
            if (!string.IsNullOrEmpty(finfo.rdlcFilename))
            {
                try
                {
                    Warning[] warnings;
                    string[] streamIds;
                    string mimeType = string.Empty;
                    string encoding = string.Empty;
                    string extension = string.Empty;

                    string rptReportPath = finfo.rptPath;
                    string rptOutPath = finfo.rptOutPath; // *FI*
                    ReportViewer viewer = new ReportViewer();
                    viewer.LocalReport.EnableExternalImages = true;
                    viewer.ProcessingMode = ProcessingMode.Local;
                    viewer.LocalReport.ReportPath = string.Concat(@rptReportPath, finfo.rdlcFilename); // *FI*
                    viewer.LocalReport.DataSources.Add(rds); // Add datasource here
                    viewer.LocalReport.SetParameters(rptParams);
                    viewer.LocalReport.Refresh();
                    byte[] bytes = viewer.LocalReport.Render(rptType(finfo.rptFileExt), null, out mimeType, out encoding, out extension, out streamIds, out warnings);
                    string filename = string.Concat(@rptOutPath, finfo.rptFileName + "." + finfo.rptFileExt);
                    using (var fs = new FileStream(filename, FileMode.Create))
                    {
                        fs.Write(bytes, 0, bytes.Length);
                        fs.Close();
                    }
                }
                catch
                {

                }
            }
        }
        // *************************************************************************************************
        // Remove special char from string
        // *************************************************************************************************
        private string cleanFileName(string cfilename)
        {
            String result = new String(cfilename.Where(ch => Char.IsLetterOrDigit(ch)).ToArray());
            return result;
        }
        // *************************************************************************************************
        //  
        // *************************************************************************************************
        private string rptType(string fileext)
        {
            string rType = "PDF";
            switch (fileext.ToLower())
            {
                case "pdf":
                    rType = "PDF";
                    break;
                case "xls":
                    rType = "EXCEL";
                    break;
            }
            return rType;
        }
        // *************************************************************************************************
        // 
        // *************************************************************************************************
        public class fileInfo
        {
            private string rptoutpath;
            private string rptpath;
            private string rptfilename;
            private string rptreportgroup;
            private bool? createnewfilename;
            private string rdlcfilename;
            private string billtosend;
            private string rptfileext;
            public string rptOutPath
            {
                get { return this.rptoutpath; }

                set { rptoutpath = value; }
            }
            public string rptPath
            {
                get { return rptpath; }

                set { rptpath = value; }
            }
            public string rptFileName
            {
                get { return rptfilename; }

                set { rptfilename = value; }
            }
            public string rptReportGroup
            {
                get { return rptreportgroup; }

                set { rptreportgroup = value; }
            }
            public bool? createNewFilename
            {
                get { return createnewfilename; }

                set { createnewfilename = value; }
            }
            public string rdlcFilename
            {
                get { return this.rdlcfilename; }

                set { rdlcfilename = value; }
            }
            public string billToSend
            {
                get { return this.billtosend; }

                set { billtosend = value; }
            }
            public string rptFileExt
            {
                get { return this.rptfileext; }

                set { rptfileext = value; }
            }
        }
    }
    class StartBilling
    {
        static void Main()
        {
            // Call the constructor that has no parameters.
            BillingCreate nBill = new BillingCreate();
            nBill.createBill();
            //Console.ReadLine();
        }

    }

}
