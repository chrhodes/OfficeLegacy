using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Xml.Linq;

using Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;
using ExcelHlp = AddinHelper.Excel;

// TODO
// Do something clever so only have to walk the worksheet one time looking for matching rows.

namespace SupportTools_Excel.User_Interface.Task_Panes
{
    public partial class TaskPane_LTC : UserControl
    {

        #region Constants

        // Worksheet names

        private const string cFILE_CONTROL_WS_NAME = "File Control";
        private const string cCLAIMS_WS_NAME = "Claims";
        private const string cPAYMENTS_WS_NAME = "Payments";
        private const string cXOBINFO_WS_NAME = "XOB Info";

        private const string cDEFAULT_OUTPUTFILE_NAME = @"LTC Claims.xml";
        private const string cDEFAULT_INPUTFILE_NAME = @"Test Claims Data File";
        private const string cDEFAULT_INPUT_FOLDER = @"C:\TFS\Systems\Office\SupportTools_Excel\LTC";

        // FileControl Worksheet


        // Claims Worksheet

        private const int cWS_CLAIMS_SEQID_COLUMN = 1;
        private const int cWS_CLAIMS_POLICYNUMBER_COLUMN = 3;

        // Payments Worksheet

        private const int cWS_PAYMENTS_POLICYNUMBER_COLUMN = 1;
        private const int cWS_PAYMENTS_CLAIMSEQID_COLUMN = 2;
        private const int cWS_PAYMENTS_SEQID_COLUMN = 3;
        private const int cWS_PAYMENTS_BENEFITPAYMENTID_COLUMN = 25;

        // XOB Worksheet

        private const int cWS_XOB_BENEFITPAYMENTID_COLUMN = 1;

        #endregion

        #region Fields and Properties

        private static Worksheet _wsFileControl = null;
        private static Worksheet _wsClaims = null;
        private static Worksheet _wsPayments = null;
        private static Worksheet _wsXOBInfo = null;

        public static int CountFC_ClaimRecord
        {
            get;
            set;
        }
        public static int CountWS_ClaimRecord
        {
            get;
            set;
        }

        public static int CountFC_ClaimPayment
        {
            get;
            set;
        }
        public static int CountWS_ClaimPayment
        {
            get;
            set;
        }

        public static int CountFC_XOBUDSCode
        {
            get;
            set;
        }
        public static int CountWS_XOBUDSCode
        {
            get;
            set;
        }
        
        public static float ClaimPaymentTotalAmount
        {
            get;
            set;
        }
        public static float LoanPaymentTotalAmount
        {
            get;
            set;
        }
        public static float InterestPaymentTotalAmount
        {
            get;
            set;
        }
        public static float CareGiverTrainingTotal
        {
            get;
            set;
        }
        public static float TransitionPaymentTotalAmount
        {
            get;
            set;
        }
        public static float FederalWithholdingTotalAmount
        {
            get;
            set;
        }
        public static float StateWitholdingTotalAmount
        {
            get;
            set;
        }
        public static float CityWithholdingTotalAmount
        {
            get;
            set;
        }
        public static float CheckAmountTotalAmount
        {
            get;
            set;
        }
        public static float XOBUDSCodeAmount
        {
            get;
            set;
        }
         
        #endregion

        #region Initialization

        public TaskPane_LTC()
        {
            InitializeComponent();
        }

        #endregion

        #region Event Handlers

        private void btnDebug_Click(object sender, EventArgs e)
        {

        }

        private void btnEvaluateWorkbook_Click(object sender, EventArgs e)
        {
            EvaluateWorkbookAndGatherProcessingInfo();
        }

        private void btnOpenSourceFile_Click(object sender, EventArgs e)
        {
            openFileDialog1.Filter = @"XLS files (*.xls*)|*.xls*|All files (*.*)|*.*";
            openFileDialog1.InitialDirectory = cDEFAULT_INPUT_FOLDER;
            openFileDialog1.FileName = cDEFAULT_INPUTFILE_NAME;

            if(DialogResult.OK == openFileDialog1.ShowDialog())
            {
                Globals.ThisAddIn.Application.Workbooks.Open(openFileDialog1.FileName);
            }
        }

        private void btnSaveXMLFile_Click(object sender, EventArgs e)
        {
            SaveXMLFile(ProduceXML());    
        }

        #endregion

        #region Main Function Routines

        private static XElement AddElement_Claim(int claimRow, int claimSeqId)
        {
            Common.WriteToDebugWindow(string.Format("AddElement_Claim(claimRow:{0}, claimSeqId:{1})", claimRow, claimSeqId));

            string policyNumber;
            
            try
            {
                policyNumber = _wsClaims.Cells[claimRow, cWS_CLAIMS_POLICYNUMBER_COLUMN].Value.ToString().ToUpper();

            }
            catch(Exception)
            {
                MessageBox.Show(string.Format("Unable to locate PolicyNumber on Claims Worksheet row {0}", claimRow));
                throw;
            }

            XElement claimElement = new XElement("Claim", new XAttribute("id", string.Format("Claim_{0}", claimSeqId)),
                new XElement("SeqID",                claimSeqId),
                new XElement("Company",              _wsClaims.Cells[claimRow, 2].Value.ToString().ToUpper()),
                new XElement("PolicyNumber",         policyNumber),
                new XElement("ClaimEndDate",         _wsClaims.Cells[claimRow, 4].Value.ToString()),
                new XElement("ClaimStartDate",       _wsClaims.Cells[claimRow, 5].Value.ToString()),
                new XElement("EliminationDaysUsed",  _wsClaims.Cells[claimRow, 6].Value.ToString()),
                new XElement("EOBStatus",            _wsClaims.Cells[claimRow, 7].Value.ToString().ToUpper())
                );

            // Add the Payment Elements for the Claim

            var matchingRows = FindMatchingPayments(policyNumber, claimSeqId);

            for(int i = 0 ; i < matchingRows.Count ; i++)
            {
                int paymentRow = matchingRows[i];
                int paymentSeqID = int.Parse(_wsPayments.Cells[paymentRow, cWS_PAYMENTS_SEQID_COLUMN].Value.ToString());
                claimElement.Add(AddElement_Payment(paymentRow, paymentSeqID));
            }

            return claimElement;
        }

        private static XElement AddElement_FileControl()
        {
            Common.WriteToDebugWindow(string.Format("AddElement_FileControl()"));

            // This assumes the SystemControl field and the CycleDate are the same for each row.

            XElement fileControlElement = new XElement("FileControl",
                new XElement("SystemControl",                _wsFileControl.Range["A2"].Value.ToString().ToUpper()),
                new XElement("CycleDate",                    _wsFileControl.Range["B2"].Value),
                new XElement("ClaimRecordCount",             CountFC_ClaimRecord),
                new XElement("ClaimPaymentCount",            CountFC_ClaimPayment),
                new XElement("XOBUDSCodeCount",              CountFC_XOBUDSCode),
                new XElement("ClaimPaymentTotalAmount",      string.Format("{0:F2}", ClaimPaymentTotalAmount)),
                new XElement("LoanPaymentTotalAmount",       string.Format("{0:F2}", LoanPaymentTotalAmount)),
                new XElement("InterestPaymentTotalAmount",   string.Format("{0:F2}", InterestPaymentTotalAmount)),
                new XElement("CareGiverTrainingTotal",       string.Format("{0:F2}", CareGiverTrainingTotal)),
                new XElement("TransitionPaymentTotalAmount", string.Format("{0:F2}", TransitionPaymentTotalAmount)),
                new XElement("FederalWitholdingTotalAmount", string.Format("{0:F2}", FederalWithholdingTotalAmount)),
                new XElement("StateWithholdingTotalAmount",  string.Format("{0:F2}", StateWitholdingTotalAmount)),
                new XElement("CityWithholdingTotalAmount",   string.Format("{0:F2}", CityWithholdingTotalAmount)),
                new XElement("CheckAmountTotalAmount",       string.Format("{0:F2}", CheckAmountTotalAmount)),
                new XElement("XOBUDSCodeAmount",             string.Format("{0:F2}", XOBUDSCodeAmount))
                );

            return fileControlElement;
        }

        private static XElement AddElement_Payment(int paymentRow, int paymentSeqId)
        {
            Common.WriteToDebugWindow(string.Format("AddElement_Payment(paymentRow:{0}, paymentSeqId:{1})", paymentRow, paymentSeqId));

            // Not all cells have data.  Handle the cases where it is not required.  Did not iterate across the column headers
            // in case the file does not have header row that matches <Element> name, e.g. FileControl worksheet.

            string payeeLast;
            string payeeFirst;
            string payeeOrg;
            string payeeAddressLine1;
            string payeeAddressLine2;
            string payeeAddressLine3;
            string payeeCity;
            string payeeState;
            string payeeZip;
            string payeeAddressCountry;
            string payeeTaxID;
            string payeeTaxIDType;
            string claimAmount;
            string checkAmount;
            string loanPaymentAmount;
            string interestAmount;
            string careGiverTrainingAmount;
            string transitionPaymentAmount;
            string federalWithholdingAmount;
            string stateWithholdingAmount;
            string cityWithholdingAmount;
            int benefitPaymentID = int.Parse(_wsPayments.Cells[paymentRow, cWS_PAYMENTS_BENEFITPAYMENTID_COLUMN].Value.ToString());
            string paymentType;
            string checkNumber;
            string reversal;
            string reversalReason;

            try { payeeLast                 = _wsPayments.Cells[paymentRow, 4].Value.ToString(); }              catch(Exception) { payeeLast = ""; }
            try { payeeFirst                = _wsPayments.Cells[paymentRow, 5].Value.ToString(); }              catch(Exception) { payeeFirst = ""; }
            try { payeeOrg                  = _wsPayments.Cells[paymentRow, 6].Value.ToString(); }              catch(Exception) { payeeOrg = ""; }
            try { payeeAddressLine1         = _wsPayments.Cells[paymentRow, 7].Value.ToString(); }              catch(Exception) { payeeAddressLine1 = ""; }
            try { payeeAddressLine2         = _wsPayments.Cells[paymentRow, 8].Value.ToString(); }              catch(Exception) { payeeAddressLine2 = ""; }
            try { payeeAddressLine3         = _wsPayments.Cells[paymentRow, 9].Value.ToString(); }              catch(Exception) { payeeAddressLine3 = ""; }
            try { payeeCity                 = _wsPayments.Cells[paymentRow, 10].Value.ToString(); }             catch(Exception) { payeeCity = ""; }
            try { payeeState                = _wsPayments.Cells[paymentRow, 11].Value.ToString().ToUpper(); }   catch(Exception) { payeeState = ""; }
            try { payeeZip                  = _wsPayments.Cells[paymentRow, 12].Value.ToString();   }           catch(Exception) { payeeZip = ""; }
            try { payeeAddressCountry       = _wsPayments.Cells[paymentRow, 13].Value.ToString(); }             catch(Exception) { payeeAddressCountry = "0"; }
            try { payeeTaxID                = _wsPayments.Cells[paymentRow, 14].Value.ToString(); }             catch(Exception) { payeeTaxID = ""; }
            try { payeeTaxIDType            = _wsPayments.Cells[paymentRow, 15].Value.ToString().ToUpper(); }   catch(Exception) { payeeTaxIDType = ""; }
            try { claimAmount               = _wsPayments.Cells[paymentRow, 16].Value.ToString("F2"); }         catch(Exception) { claimAmount = "0"; }
            try { checkAmount               = _wsPayments.Cells[paymentRow, 17].Value.ToString("F2"); }         catch(Exception) { checkAmount = "0"; }
            try { loanPaymentAmount         = _wsPayments.Cells[paymentRow, 18].Value.ToString("F2"); }         catch(Exception) { loanPaymentAmount = "0"; }
            try { interestAmount            = _wsPayments.Cells[paymentRow, 19].Value.ToString("F2"); }         catch(Exception) { interestAmount = "0"; }
            try { careGiverTrainingAmount   = _wsPayments.Cells[paymentRow, 20].Value.ToString("F2"); }         catch(Exception) { careGiverTrainingAmount = "0"; }
            try { transitionPaymentAmount   = _wsPayments.Cells[paymentRow, 21].Value.ToString("F2"); }         catch(Exception) { transitionPaymentAmount = "0"; }
            try { federalWithholdingAmount  = _wsPayments.Cells[paymentRow, 22].Value.ToString("F2"); }         catch(Exception) { federalWithholdingAmount = "0"; }
            try { stateWithholdingAmount    = _wsPayments.Cells[paymentRow, 23].Value.ToString("F2"); }         catch(Exception) { stateWithholdingAmount = "0"; }
            try { cityWithholdingAmount     = _wsPayments.Cells[paymentRow, 24].Value.ToString("F2"); }         catch(Exception) { cityWithholdingAmount = "0"; }

            try { paymentType               = _wsPayments.Cells[paymentRow, 26].Value.ToString().ToUpper(); }             catch(Exception) { paymentType = ""; }
            try { checkNumber               = _wsPayments.Cells[paymentRow, 27].Value.ToString(); }             catch(Exception) { checkNumber = ""; }
            try { reversal                  = _wsPayments.Cells[paymentRow, 28].Value.ToString().ToUpper(); }   catch(Exception) { reversal = ""; }
            try { reversalReason            = _wsPayments.Cells[paymentRow, 29].Value.ToString().ToUpper(); }   catch(Exception) { reversalReason = ""; }

            XElement paymentElement = new XElement("Payment", new XAttribute("id", string.Format("Payment_{0}", paymentSeqId)),
                new XElement("SeqID",                    paymentSeqId),
                new XElement("PayeeLast",                payeeLast),
                new XElement("PayeeFirst",               payeeFirst),
                new XElement("PayeeOrg",                 payeeOrg),
                new XElement("PayeeAddressLine1",        payeeAddressLine1),
                new XElement("PayeeAddressLine2",        payeeAddressLine2),
                new XElement("PayeeAddressLine3",        payeeAddressLine3),
                new XElement("PayeeCity",                payeeCity),
                new XElement("PayeeState",               payeeState),
                new XElement("PayeeZip",                 payeeZip),
                new XElement("PayeeAddressCountry",      payeeAddressCountry),
                new XElement("PayeeTaxID",               payeeTaxID),
                new XElement("PayeeTaxIDType",           payeeTaxIDType),
                new XElement("ClaimAmount",              claimAmount),
                new XElement("CheckAmount",              checkAmount),
                new XElement("LoanPaymentAmount",        loanPaymentAmount),
                new XElement("InterestAmount",           interestAmount),
                new XElement("CareGiverTrainingAmount",  careGiverTrainingAmount),
                new XElement("TransitionPaymentAmount",  transitionPaymentAmount),
                new XElement("FederalWithholdingAmount", federalWithholdingAmount),
                new XElement("StateWithholdingAmount",   stateWithholdingAmount),
                new XElement("CityWithholdingAmount",    cityWithholdingAmount),
                new XElement("BenefitPaymentID",         benefitPaymentID), // Y2
                new XElement("PaymentType",              paymentType),
                new XElement("CheckNumber",              checkNumber),
                new XElement("Reversal",                 reversal),
                new XElement("ReversalReason",           reversalReason)
                );

            var matchingRows = FindMatchingXOB(benefitPaymentID);

            for (int i = 0; i < matchingRows.Count; i++)
            {
                int xobRow = matchingRows[i];
                int xobSeqID = i + 1;   // Sequence IDs start with 1
                paymentElement.Add(AddElement_XobInfo(xobRow, xobSeqID));
            }

            return paymentElement;
        }

        private static XElement AddElement_XobInfo(int xobRow, int xobSeqId)
        {
            Common.WriteToDebugWindow(string.Format("AddElement_XobInfo(xobRow:{0}, xobSeqId:{1})", xobRow, xobSeqId));

            XElement xobElement = new XElement("XOBInfo", new XAttribute("id", string.Format("XOB_Info_{0}", xobSeqId)),
                new XElement("XOBUDSCode",          _wsXOBInfo.Cells[xobRow, 2].Value.ToString().ToUpper()),
                new XElement("Amount",              _wsXOBInfo.Cells[xobRow, 3].Value.ToString("F2")),
                new XElement("ServicePeriodStart",  _wsXOBInfo.Cells[xobRow, 4].Value),
                new XElement("ServicePeriodEnd",    _wsXOBInfo.Cells[xobRow, 5].Value),
                new XElement("ServiceUnits",        _wsXOBInfo.Cells[xobRow, 6].Value),
                new XElement("ServiceUnitsMeasure", _wsXOBInfo.Cells[xobRow, 7].Value)
                );

            return xobElement;
        }

        private void EvaluateWorkbookAndGatherProcessingInfo()
        {
            if ( ! WorksheetsAreValid())
            {
                return;
            }

            // Extract some info off the worksheets and place in properties.
            LoadInfoFromWorksheet_FileControl();
            LoadInfoFromWorksheet_Claims();
            LoadInfoFromWorksheet_Payments();
            LoadInfoFromWorksheet_XOBInfo();

            txtCountFC_ClaimPayment.Text = CountFC_ClaimPayment.ToString();
            txtCountWS_ClaimPayment.Text = CountWS_ClaimPayment.ToString();

            txtCountFC_ClaimRecord.Text = CountFC_ClaimRecord.ToString();
            txtCountWS_ClaimRecord.Text = CountWS_ClaimRecord.ToString();

            txtCountFC_XOBUDSCode.Text = CountFC_XOBUDSCode.ToString();
            txtCountWS_XOBUDSCode.Text = CountWS_XOBUDSCode.ToString();

            if (CountFC_ClaimRecord == CountWS_ClaimRecord &&
                CountFC_ClaimPayment == CountWS_ClaimPayment &&
                CountFC_XOBUDSCode == CountWS_XOBUDSCode)
            {
            	lblFileIsValid.Text = "Workbook appears to be valid";
                lblFileIsValid.ForeColor = Color.Green;
            }
            else
            {
                lblFileIsValid.Text = "Workbook has inconsistent record counts";
                lblFileIsValid.ForeColor = Color.Red;
            }
        }

        private static List<int> FindMatchingPayments(string policyNumber, int claimSeqId)
        {
            Common.WriteToDebugWindow(string.Format("FindMatchingPayments(policyNumber:{0} claimSeqId:{1}", policyNumber, claimSeqId));
            List<int> matchingRows = new List<int>();

            // Look on Payment worksheet for matching rows.  First row is header.  Data starts on second row (i = 2)
            for(int i = 2 ; i < CountFC_ClaimPayment + 2 ; i++)
            {
                if(policyNumber.ToUpper() == _wsPayments.Cells[i, cWS_PAYMENTS_POLICYNUMBER_COLUMN].Value.ToUpper()
                    && claimSeqId == int.Parse(_wsPayments.Cells[i, cWS_PAYMENTS_CLAIMSEQID_COLUMN].Value.ToString()))
                {
                    Common.WriteToDebugWindow(string.Format("FindMatchingPayments - matchingRow:{0}", i));
                    matchingRows.Add(i);
                }
            }

  
            return matchingRows;
        }

        private static List<int> FindMatchingXOB(int benefitPaymentID)
        {
            Common.WriteToDebugWindow(string.Format("FindMatchingXOB(benefitPaymentID:{0}", benefitPaymentID));
            List<int> matchingRows = new List<int>();

            // Look on XOB worksheet for matching rows.  First row is header.  Data starts on second row (i = 2)
            for(int i = 2 ; i < CountFC_XOBUDSCode + 2; i++)
            {
                // NB.  Need to change this to string if benefitPaymentIDs are not numeric
                if(benefitPaymentID == int.Parse(_wsXOBInfo.Cells[i, cWS_XOB_BENEFITPAYMENTID_COLUMN].Value.ToString()))
                {
                    Common.WriteToDebugWindow(string.Format("FindMatchingXOB - matchingRow:{0}", i));
                    matchingRows.Add(i);
                }                
            }

            return matchingRows;
        }

        private void LoadInfoFromWorksheet_Claims()
        {
            CountWS_ClaimRecord = ExcelHlp.FindLast_PopulatedRow_InColumn(_wsClaims.Range["A1"]) - 1;
        }

        private void LoadInfoFromWorksheet_FileControl()
        {
            CountFC_ClaimRecord = ExcelHlp.FindLast_PopulatedRow_InColumn(_wsFileControl.Range["A1"]) - 1;

            // Need populate these as sum of values.
            Range sumRange;
            //string sumRangeString;

            //sumRangeString = string.Format("D2:D{0}", CountFC_ClaimRecord + 1);
            sumRange = _wsFileControl.Range[string.Format("D2:D{0}", CountFC_ClaimRecord + 1)];

            CountFC_ClaimPayment = (int)Globals.ThisAddIn.Application.WorksheetFunction.Sum(sumRange);

            //sumRangeString = string.Format("E2:E{0}", CountFC_ClaimRecord + 1);
            //sumRange = _wsFileControl.Range[sumRangeString];

            sumRange = _wsFileControl.Range[string.Format("E2:E{0}", CountFC_ClaimRecord + 1)];
            CountFC_XOBUDSCode = (int)Globals.ThisAddIn.Application.WorksheetFunction.Sum(sumRange);

            sumRange = _wsFileControl.Range[string.Format("F2:F{0}", CountFC_ClaimRecord + 1)];
            ClaimPaymentTotalAmount = (float)Globals.ThisAddIn.Application.WorksheetFunction.Sum(sumRange);
            
            sumRange = _wsFileControl.Range[string.Format("G2:G{0}", CountFC_ClaimRecord + 1)];
            LoanPaymentTotalAmount = (float)Globals.ThisAddIn.Application.WorksheetFunction.Sum(sumRange);

            sumRange = _wsFileControl.Range[string.Format("H2:H{0}", CountFC_ClaimRecord + 1)];
            InterestPaymentTotalAmount = (float)Globals.ThisAddIn.Application.WorksheetFunction.Sum(sumRange);

            sumRange = _wsFileControl.Range[string.Format("I2:I{0}", CountFC_ClaimRecord + 1)];
            CareGiverTrainingTotal = (float)Globals.ThisAddIn.Application.WorksheetFunction.Sum(sumRange);

            sumRange = _wsFileControl.Range[string.Format("J2:J{0}", CountFC_ClaimRecord + 1)];
            TransitionPaymentTotalAmount = (float)Globals.ThisAddIn.Application.WorksheetFunction.Sum(sumRange);

            sumRange = _wsFileControl.Range[string.Format("K2:K{0}", CountFC_ClaimRecord + 1)];
            FederalWithholdingTotalAmount = (float)Globals.ThisAddIn.Application.WorksheetFunction.Sum(sumRange);

            sumRange = _wsFileControl.Range[string.Format("L2:L{0}", CountFC_ClaimRecord + 1)];
            StateWitholdingTotalAmount = (float)Globals.ThisAddIn.Application.WorksheetFunction.Sum(sumRange);

            sumRange = _wsFileControl.Range[string.Format("M2:M{0}", CountFC_ClaimRecord + 1)];
            CityWithholdingTotalAmount = (float)Globals.ThisAddIn.Application.WorksheetFunction.Sum(sumRange);

            sumRange = _wsFileControl.Range[string.Format("N2:N{0}", CountFC_ClaimRecord + 1)];
            CheckAmountTotalAmount = (float)Globals.ThisAddIn.Application.WorksheetFunction.Sum(sumRange);

            sumRange = _wsFileControl.Range[string.Format("O2:O{0}", CountFC_ClaimRecord + 1)];
            XOBUDSCodeAmount = (float)Globals.ThisAddIn.Application.WorksheetFunction.Sum(sumRange);
        }

        private void LoadInfoFromWorksheet_Payments()
        {
            CountWS_ClaimPayment = ExcelHlp.FindLast_PopulatedRow_InColumn(_wsPayments.Range["A1"]) - 1;
        }

        private void LoadInfoFromWorksheet_XOBInfo()
        {
            CountWS_XOBUDSCode = ExcelHlp.FindLast_PopulatedRow_InColumn(_wsXOBInfo.Range["A1"]) - 1;
        }

        private XElement ProduceXML()
        {
            XElement outputXml = new XElement("ClaimFile");

            outputXml.Add(AddElement_FileControl());

            for(int i = 2 ; i < CountFC_ClaimRecord + 2 ; i++)
            {
                int claimSeqId = int.Parse(_wsClaims.Cells[i, cWS_CLAIMS_SEQID_COLUMN].Value.ToString());
                outputXml.Add(AddElement_Claim(i, claimSeqId));                
            }

            Common.WriteToDebugWindow("SaveXMLFile()");
            Common.WriteToDebugWindow(outputXml.ToString());

            return outputXml;        
        }

        private void SaveXMLFile(XElement outputXml)
        {
            saveFileDialog1.Filter = "XML files (*.xml)|*.xml|All files (*.*)|*.*";
            saveFileDialog1.FileName = cDEFAULT_OUTPUTFILE_NAME;

            if(DialogResult.OK == saveFileDialog1.ShowDialog())
            {
                StreamWriter writer = new StreamWriter(saveFileDialog1.FileName);
                writer.Write(outputXml.ToString());
                writer.Flush();
                writer.Close();
            }
        }

        private bool WorksheetIsValid_FileControl()
        {
            // Can do something if necessary.
            return true;
        }

        private bool WorksheetIsValid_Claims()
        {
            // Can do something if necessary.
            return true;
        }

        private bool WorkSheetIsValid_Payments()
        {
            // Can do something if necessary.
            return true;
        }

        private bool WorksheetIsValid_XOBInfo()
        {
            // Can do something if necessary.
            return true;
        }

        /// <summary>
        /// Validate current workbook has valid worksheets.
        /// </summary>
        /// <returns></returns>
        private bool WorksheetsAreValid()
        {
            bool result = false;

            Workbook wb = Globals.ThisAddIn.Application.ActiveWorkbook;

            if (wb == null)
            {
            	MessageBox.Show("Cannot proceed.  Open source workbook first");
                return result;
            }

            try
            {
                _wsFileControl = wb.Sheets[cFILE_CONTROL_WS_NAME];
                result = WorksheetIsValid_FileControl();
            }
            catch
            {
                MessageBox.Show(string.Format("Invalid or missing worksheet: {0}", cFILE_CONTROL_WS_NAME));
                return result;
            }

            try
            {
                _wsClaims = wb.Sheets[cCLAIMS_WS_NAME];
                result = WorksheetIsValid_Claims();
            }
            catch
            {
                MessageBox.Show(string.Format("Invalid or missing worksheet: {0}", cCLAIMS_WS_NAME));
                return result;
            }

            try
            {
                _wsPayments = wb.Sheets[cPAYMENTS_WS_NAME];
                result = WorkSheetIsValid_Payments();
            }
            catch
            {
                MessageBox.Show(string.Format("Invalid or missing worksheet: {0}", cPAYMENTS_WS_NAME));
                return result;
            }

            try
            {
                _wsXOBInfo = wb.Sheets[cXOBINFO_WS_NAME];
                result = WorksheetIsValid_XOBInfo();
            }
            catch
            {
                Common.WriteToDebugWindow(string.Format("Invalid or missing worksheet: {0}", cXOBINFO_WS_NAME));
                return result;
            }

            return result;
        }
        #endregion

    }
}
