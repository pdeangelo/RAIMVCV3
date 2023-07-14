using ClosedXML.Excel;
using RAIMVCV3.Models;
using RAIMVCV3.Repository;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using Spire.Doc;
using System.Net;
using System.Linq;
using System.Data.Entity;

namespace RAIMVCV3.Controllers
{
    public class HomeController : Controller
    {
        private ClientsRepository _clientsRepository = null;
        private LoanRepository _loanRepository = null;
        private InvestorRepository _investorRepository = null;
        private EntityRepository _entityRepository = null;
        public HomeController()
        {
             _loanRepository = new LoanRepository();
            _clientsRepository = new ClientsRepository();
            _investorRepository = new InvestorRepository();
            _entityRepository = new EntityRepository();
        }
        [Authorize]
        [AuthAttribute(Roles = "Accounting,Underwriting")]
        public ActionResult Index()
        {
            var raiLoans = _loanRepository.GetLoans().Where(x => x.LoanStatus.LoanStatusName.Equals("1 - Underwriting")
                    || x.LoanStatus.LoanStatusName.Equals("2 - Funding")
                    || x.LoanStatus.LoanStatusName.Equals("3 - Open")).ToList();
            SetupSelectListItems();
            return View(raiLoans);
        }
        [HttpPost]
        public ActionResult Index(string SearchText, string StatusSelectListItems, string EntitiesSelectListItems, string ClientsSelectListItems, string chkShowCompleted)
        {
            var raiLoans = _loanRepository.GetLoans();

            if (chkShowCompleted == "false")
                raiLoans = raiLoans.Where(x => x.LoanStatus.LoanStatusName.Equals("1 - Underwriting")
                    || x.LoanStatus.LoanStatusName.Equals("2 - Funding")
                    || x.LoanStatus.LoanStatusName.Equals("3 - Open")).ToList();

            if (StatusSelectListItems != "")
                raiLoans = raiLoans.Where(x => x.LoanStatusID == Convert.ToInt32(StatusSelectListItems)).ToList();

            if (EntitiesSelectListItems != "")
                raiLoans = raiLoans.Where(x => x.EntityID == Convert.ToInt32(EntitiesSelectListItems)).ToList();

            if (ClientsSelectListItems != "")
                raiLoans = raiLoans.Where(x => x.ClientID == Convert.ToInt32(ClientsSelectListItems)).ToList();

            if (SearchText != "")
                raiLoans = raiLoans.Where(loan => loan.LoanMortgagee.ToLower().Contains(SearchText.ToLower()) 
                    || loan.LoanMortgageeBusiness.ToLower().Contains(SearchText.ToLower())
                    || loan.LoanNumber.ToLower().Contains(SearchText.ToLower())
                    || loan.LoanPropertyAddress.ToLower().Contains(SearchText.ToLower())).ToList();

            SetupSelectListItems();
            return View(raiLoans);
        }
        [HttpPost]
        public ActionResult Delete(int id)
        {
            //_loanRepository.DeleteLoan(id);

            //TempData["Message"] = "Loan Successfully Deleted";
            return RedirectToAction("Index");
        }
        public ActionResult Edit(int? id)
        {
            if (id == null)
            {
                return new HttpStatusCodeResult(HttpStatusCode.BadRequest);
            }

            Loan loan = _loanRepository.GetLoan((int)id);
            if (loan == null)
            {
                return HttpNotFound();
            }


            SetupSelectListItems();
            return View(loan);
        }
        [HttpPost]
        public ActionResult Edit(Loan loan)
        {
            ValidateLoan(loan);
            if (ModelState.IsValid)
            {
                _loanRepository.UpdateLoan(loan);

                TempData["Message"] = "Loan Successfully Updated";
                return RedirectToAction("Index");
            }

            SetupSelectListItems();
            return View(loan);
        }

        public ActionResult RunBaileeLetter(string loanlist)
        {
            //TempData.Keep("Message");
            try
            {
               
                List<string> loansSplit = loanlist.Split(',').ToList();

                List<Int32> modLoansSplit = new List<Int32>();
                foreach (string loanID in loansSplit)
                {
                    int test;
                    int.TryParse(loanID, out test);
                    if (loanID.Length > 0 && test > 0)
                        modLoansSplit.Add(Convert.ToInt32(loanID));
                }
                if (modLoansSplit.Count() == 0)
                {
                    TempData["Error"] = "Error - Please select at least one loan";
                    return Json("");
                }

                List<Loan> loans = new List<Loan>();
                foreach (Int32 loanID in modLoansSplit)
                {                    
                    Loan loan = _loanRepository.GetLoan(loanID);
                    loans.Add(loan);
                 
                }
                

                Client client = _loanRepository.GetClient(loans[0].ClientID.Value);
                Entity entity = _loanRepository.GetEntity(loans[0].EntityID.Value);

                Spire.Doc.Document document = new Spire.Doc.Document();
                //Add Section

                Spire.Doc.Section section1 = document.AddSection();

                section1.PageSetup.Orientation = Spire.Doc.Documents.PageOrientation.Portrait;

                //Add Paragraph

                Spire.Doc.Documents.Paragraph parExhibitA = section1.AddParagraph();

                parExhibitA.Format.HorizontalAlignment = Spire.Doc.Documents.HorizontalAlignment.Center;
                //Add Paragraph

                Spire.Doc.Documents.Paragraph parSchedofInv = section1.AddParagraph();
                parSchedofInv.Format.HorizontalAlignment = Spire.Doc.Documents.HorizontalAlignment.Center;

                Spire.Doc.Documents.Paragraph parDate = section1.AddParagraph();
                parDate.Format.HorizontalAlignment = Spire.Doc.Documents.HorizontalAlignment.Center;

                Spire.Doc.Table tblInv = section1.AddTable(true);
                Spire.Doc.Fields.TextRange trExhibitA = parExhibitA.AppendText("EXHIBIT A\v");
                trExhibitA.CharacterFormat.UnderlineStyle = Spire.Doc.Documents.UnderlineStyle.Single;


                parSchedofInv.AppendText("SCHEDULE OF ASSETS\v");

                parDate.AppendText("Date:" + DateTime.Now.ToShortDateString() + "\v");

                //Create Header and Data

                String[] Header = { "Loan ID", "Borrower", "Address" };

                //
                // Table Logic
                //
                //Add Cells
                tblInv.ResetCells(loans.Count() + 1, Header.Length);
                //Header Row
                Spire.Doc.TableRow FRow = tblInv.Rows[0];
                FRow.IsHeader = true;

                //Row Height
                FRow.Height = 23;

                //Header Format
                //FRow.RowFormat.BackColor = Color.AliceBlue;

                for (int i = 0; i < Header.Length; i++)
                {
                    //Cell Alignment
                    Spire.Doc.Documents.Paragraph p = FRow.Cells[i].AddParagraph();
                    FRow.Cells[i].CellFormat.VerticalAlignment = Spire.Doc.Documents.VerticalAlignment.Middle;
                    p.Format.HorizontalAlignment = Spire.Doc.Documents.HorizontalAlignment.Center;
                    //Data Format
                    Spire.Doc.Fields.TextRange TR = p.AppendText(Header[i]);
                    //TR.CharacterFormat.FontName = "Calibri";
                    TR.CharacterFormat.FontSize = 14;
                    //TR.CharacterFormat.TextColor = Color.Teal;                
                    TR.CharacterFormat.Bold = true;

                }

                //Data Row

                string loanNumberList = "";
                int r = 0;
                //for (int r = 0; r < data.Length; r++)
                foreach (Loan rptloan in loans)
                {

                    Spire.Doc.TableRow DataRow = tblInv.Rows[r + 1];

                    //Row Height

                    DataRow.Height = 20;

                    //C Represents Column.

                    //for (int c = 0; c < data[r].Length; c++)                    
                    //{                    
                    //Cell Alignment
                    DataRow.Cells[0].CellFormat.VerticalAlignment = Spire.Doc.Documents.VerticalAlignment.Middle;
                    DataRow.Cells[1].CellFormat.VerticalAlignment = Spire.Doc.Documents.VerticalAlignment.Middle;
                    DataRow.Cells[2].CellFormat.VerticalAlignment = Spire.Doc.Documents.VerticalAlignment.Middle;

                    //Fill Data in Rows

                    loanNumberList += rptloan.LoanNumber + " "; ;

                    Spire.Doc.Documents.Paragraph pLoanNumber = DataRow.Cells[0].AddParagraph();
                    Spire.Doc.Fields.TextRange TRLoanNumber = pLoanNumber.AppendText(rptloan.LoanNumber);

                    Spire.Doc.Documents.Paragraph pBorrower = DataRow.Cells[1].AddParagraph();
                    Spire.Doc.Fields.TextRange TRBorrower = pBorrower.AppendText(rptloan.LoanMortgagee);

                    Spire.Doc.Documents.Paragraph pAddress = DataRow.Cells[2].AddParagraph();
                    Spire.Doc.Fields.TextRange TRAddress = pAddress.AppendText(rptloan.LoanPropertyAddress);

                    //Format Cells

                    pLoanNumber.Format.HorizontalAlignment = Spire.Doc.Documents.HorizontalAlignment.Center;
                    pBorrower.Format.HorizontalAlignment = Spire.Doc.Documents.HorizontalAlignment.Center;
                    pAddress.Format.HorizontalAlignment = Spire.Doc.Documents.HorizontalAlignment.Center;

                    //TRLoanNumber.CharacterFormat.FontName = "Calibri";
                    //TRBorrower.CharacterFormat.FontName = "Calibri";
                    //TRAddress.CharacterFormat.FontName = "Calibri";

                    TRLoanNumber.CharacterFormat.FontSize = 12;
                    TRBorrower.CharacterFormat.FontSize = 12;
                    TRAddress.CharacterFormat.FontSize = 12;

                    //TR2.CharacterFormat.TextColor = Color.Brown;

                    //}
                    r++;

                }
                //
                // End Table Logic
                //
                Spire.Doc.Documents.Paragraph parWireInfo = section1.AddParagraph();
                parWireInfo.Format.HorizontalAlignment = Spire.Doc.Documents.HorizontalAlignment.Left;

                if (loans[0].BaileeLetterDate == null)
                {
                    TempData["Error"] = "Error - Bailee Letter Date Missing";
                    return Json("");
                }
                if (loans[0].ClosingDate == null)
                {
                    TempData["Error"] = "Error - Closing Date Date Missing";
                    return Json("");
                }
                DateTime baileeLetterDate = loans[0].BaileeLetterDate == null ? DateTime.MinValue : (DateTime)loans[0].BaileeLetterDate;
                DateTime closingDate = loans[0].ClosingDate == null ? DateTime.MinValue : (DateTime)loans[0].ClosingDate;

                parWireInfo.AppendText("Purchase Advice Date:" + baileeLetterDate.ToShortDateString() + "\v\v");
                parWireInfo.AppendText("Closing Date:" + closingDate.ToShortDateString() + "\v\v");
                parWireInfo.AppendText("Wiring Instructions:\v");

                Spire.Doc.Documents.Paragraph parBankInfo = section1.AddParagraph();
                parBankInfo.Format.HorizontalAlignment = Spire.Doc.Documents.HorizontalAlignment.Left;
                parBankInfo.Format.LeftIndent = 80;

                parBankInfo.AppendText("Bank: " + entity.EntityBank + "\v");
                parBankInfo.AppendText("ABA: " + entity.ABA + "\v");
                parBankInfo.AppendText("Account: " + entity.Account + "\v");
                parBankInfo.AppendText("Account Name: " + entity.EntityName + "\v");
                parBankInfo.AppendText("Ref: " + client.ClientName + "\v");
                Spire.Doc.Break pageBreak = new Spire.Doc.Break(document, Spire.Doc.Documents.BreakType.PageBreak);

                Spire.Doc.Document DocOne = new Spire.Doc.Document();

                //DocOne.LoadFromStream(FileUploadControlBailee.PostedFile.InputStream, FileFormat.Docx);
                DocOne.LoadFromFile(@"C:\Users\pdean\LetterTemplates\Bailee_Aloha_RAI Funding, LLC.docx", Spire.Doc.FileFormat.Docx);

                DateTime today = DateTime.Today;
                string todayStr;
                todayStr = string.Format("{0:MMMM dd, yyyy}", today);

                DocOne.Replace("TODAYSDATE", todayStr, false, true);

                foreach (Spire.Doc.Section sec in document.Sections)
                {
                    DocOne.Sections.Add(sec.Clone());
                }
                string fileName;
                string todayStrFile = string.Format("{0:yyyy_MM_dd}", today);

                fileName = @"C:\Users\pdean\Reports\\Bailee_" + client.ClientName + "_" + entity.EntityName + "_" + loanNumberList + ".docx";
                //CommonOpenFileDialog dialog = new CommonOpenFileDialog();
                //dialog.InitialDirectory = "C:\\Users";
                //dialog.IsFolderPicker = true;

                //if (dialog.ShowDialog() == CommonFileDialogResult.Ok)
                //{
                //    fileName = dialog.FileName + "\\" + fileName;

                //}

                DocOne.SaveToFile(fileName, FileFormat.Docx);

                System.Diagnostics.Process.Start(fileName);
                TempData["Message"] = "Bailee Report Successfully Created";
                
            }
            catch (Exception ex)
            {
                TempData["Error"] = "Bailee Report Error";
            }

            return Json("", JsonRequestBehavior.AllowGet);
        }

        public ActionResult AdvanceReport(string loanlist)
        {
            try
            {
                List<string> loansSplit = loanlist.Split(',').ToList();

                List<Int32> modLoansSplit = new List<Int32>();
                foreach (string loanID in loansSplit)
                {
                    int test;
                    int.TryParse(loanID, out test);
                    if (loanID.Length > 0 && test > 0)
                        modLoansSplit.Add(Convert.ToInt32(loanID));
                }
                if (modLoansSplit.Count() == 0)
                {
                    TempData["Error"] = "Error - Please select at least one loan";
                    return Json("");
                }

                List<Loan> loans = new List<Loan>();
                foreach (Int32 loanID in modLoansSplit)
                {
                    Loan loan = _loanRepository.GetLoan(loanID);
                    loans.Add(loan);

                }
                Client client = _loanRepository.GetClient(loans[0].ClientID.Value);
                Entity entity = _loanRepository.GetEntity(loans[0].EntityID.Value);

                Spire.Doc.Document document = new Spire.Doc.Document();
                //Add Section

                Spire.Doc.Section section1 = document.AddSection();

                section1.PageSetup.Orientation = Spire.Doc.Documents.PageOrientation.Landscape;
                //Add Paragraph

                Spire.Doc.Documents.Paragraph parSchedule1 = section1.AddParagraph();
                parSchedule1.Format.HorizontalAlignment = Spire.Doc.Documents.HorizontalAlignment.Center;
                Spire.Doc.Fields.TextRange trSchedule1 = parSchedule1.AppendText(entity.EntityName);
                //trSchedule1.CharacterFormat.UnderlineStyle = Spire.Doc.Documents.UnderlineStyle.Single;

                Spire.Doc.Documents.Paragraph parMort = section1.AddParagraph();
                parMort.Format.HorizontalAlignment = Spire.Doc.Documents.HorizontalAlignment.Center;

                parMort.AppendText("Advance Report");

                Spire.Doc.Documents.Paragraph parClient = section1.AddParagraph();
                parClient.Format.HorizontalAlignment = Spire.Doc.Documents.HorizontalAlignment.Center;
                parClient.AppendText(client.ClientName);

                Spire.Doc.Documents.Paragraph parRemitReport = section1.AddParagraph();
                parRemitReport.Format.HorizontalAlignment = Spire.Doc.Documents.HorizontalAlignment.Left;
                parRemitReport.AppendText("Advance Date " + DateTime.Today.ToShortDateString() + "\v");

                Spire.Doc.Table tblSch = section1.AddTable(true);
                //
                // Table Logic
                //

                String[] Header = { "Business Name", "Client Name", "Loan Number", "Collateral Description", "Mortgage Amount", "Reserve Amount", "Total Transfer" };

                //Add Cells
                // Number of rows is # of loans + header + footer + line for wire fee
                tblSch.ResetCells(loans.Count() + 3, Header.Length);
                int footerRow;
                int wireRow;

                footerRow = loans.Count() + 2;
                wireRow = loans.Count() + 1;
                //Header Row
                Spire.Doc.TableRow FRow2 = tblSch.Rows[0];
                FRow2.IsHeader = true;

                //Row Height
                FRow2.Height = 30;

                //Wire Fee Row
                Spire.Doc.TableRow FRowWire = tblSch.Rows[wireRow];

                FRowWire.Height = 30;
                Spire.Doc.Documents.Paragraph pWire = FRowWire.Cells[0].AddParagraph();
                Spire.Doc.Fields.TextRange TRWire = pWire.AppendText("Wire Transfer Fee If Applicable");
                pWire.Format.HorizontalAlignment = Spire.Doc.Documents.HorizontalAlignment.Left;

                Spire.Doc.Documents.Paragraph pWireFee = FRowWire.Cells[6].AddParagraph();
                Spire.Doc.Fields.TextRange TRWireFee = pWireFee.AppendText("(25.00)");
                pWireFee.Format.HorizontalAlignment = Spire.Doc.Documents.HorizontalAlignment.Right;

                //Merge first 4 cells
                tblSch.ApplyHorizontalMerge((wireRow), 0, 4);

                Spire.Doc.TableRow FRowFooter = tblSch.Rows[footerRow];

                //Merge first 3 cells
                tblSch.ApplyHorizontalMerge((footerRow), 0, 2);

                Spire.Doc.Documents.Paragraph pFooter = FRowFooter.Cells[0].AddParagraph();
                Spire.Doc.Fields.TextRange TRFooter = pFooter.AppendText("Total Due " + client.ClientName);
                pFooter.Format.HorizontalAlignment = Spire.Doc.Documents.HorizontalAlignment.Left;

                //Row Height
                FRowFooter.Height = 23;

                //Header Format
                //FRow.RowFormat.BackColor = Color.AliceBlue;

                for (int i = 0; i < Header.Length; i++)
                {
                    //Cell Alignment
                    Spire.Doc.Documents.Paragraph p = FRow2.Cells[i].AddParagraph();
                    FRow2.Cells[i].CellFormat.VerticalAlignment = Spire.Doc.Documents.VerticalAlignment.Middle;

                    p.Format.HorizontalAlignment = Spire.Doc.Documents.HorizontalAlignment.Center;
                    //Data Format
                    Spire.Doc.Fields.TextRange TR = p.AppendText(Header[i]);
                    //TR.CharacterFormat.FontName = "Calibri";
                    TR.CharacterFormat.FontSize = 14;
                    //TR.CharacterFormat.TextColor = Color.Teal;                
                    TR.CharacterFormat.Bold = true;

                }

                //Data Row

                double totalMortgage = 0;
                double totalAdvance = -25;  // -25 for wire fee
                double totalReserve = 0;

                int r = 0;
                //for (int r = 0; r < data.Length; r++)
                foreach (Loan rptloan in loans)
                {

                    Spire.Doc.TableRow DataRow = tblSch.Rows[r + 1];

                    //Row Height

                    DataRow.Height = 30;

                    //C Represents Column.

                    //for (int c = 0; c < data[r].Length; c++)                    
                    //{                    
                    //Cell Alignment
                    DataRow.Cells[0].CellFormat.VerticalAlignment = Spire.Doc.Documents.VerticalAlignment.Middle;

                    DataRow.Cells[1].CellFormat.VerticalAlignment = Spire.Doc.Documents.VerticalAlignment.Middle;
                    DataRow.Cells[2].CellFormat.VerticalAlignment = Spire.Doc.Documents.VerticalAlignment.Middle;
                    DataRow.Cells[3].CellFormat.VerticalAlignment = Spire.Doc.Documents.VerticalAlignment.Middle;
                    DataRow.Cells[4].CellFormat.VerticalAlignment = Spire.Doc.Documents.VerticalAlignment.Middle;
                    DataRow.Cells[5].CellFormat.VerticalAlignment = Spire.Doc.Documents.VerticalAlignment.Middle;
                    DataRow.Cells[6].CellFormat.VerticalAlignment = Spire.Doc.Documents.VerticalAlignment.Middle;



                    //Fill Data in Rows

                    Spire.Doc.Documents.Paragraph pBusiness = DataRow.Cells[0].AddParagraph();
                    Spire.Doc.Fields.TextRange TRBusiness = pBusiness.AppendText(rptloan.LoanMortgageeBusiness);

                    Spire.Doc.Documents.Paragraph pBorrower = DataRow.Cells[1].AddParagraph();
                    Spire.Doc.Fields.TextRange TRBorrower = pBorrower.AppendText(rptloan.LoanMortgagee);

                    Spire.Doc.Documents.Paragraph pLoanNumber = DataRow.Cells[2].AddParagraph();
                    Spire.Doc.Fields.TextRange TRLoanNumber = pLoanNumber.AppendText(rptloan.LoanNumber);

                    Spire.Doc.Documents.Paragraph pAddress = DataRow.Cells[3].AddParagraph();
                    Spire.Doc.Fields.TextRange TRAddress = pAddress.AppendText(rptloan.LoanPropertyAddress);

                    Spire.Doc.Documents.Paragraph pMortgageAmount = DataRow.Cells[4].AddParagraph();
                    Spire.Doc.Fields.TextRange trMortgageAmount = pMortgageAmount.AppendText(FormatNumberCommas2dec(rptloan.LoanMortgageAmount.ToString()));

                    Spire.Doc.Documents.Paragraph pReserveAmount = DataRow.Cells[5].AddParagraph();
                    Spire.Doc.Fields.TextRange trReserveAmount = pReserveAmount.AppendText(FormatNumberCommas2dec(rptloan.LoanReserveAmount.ToString()));

                    Spire.Doc.Documents.Paragraph pAdvanceAmount = DataRow.Cells[6].AddParagraph();
                    Spire.Doc.Fields.TextRange trAdvanceAmount = pAdvanceAmount.AppendText(FormatNumberCommas2dec(rptloan.LoanAdvanceAmount.ToString()));

                    totalMortgage += rptloan.LoanMortgageAmount.Value;
                    totalAdvance += rptloan.LoanAdvanceAmount.Value;
                    totalReserve += rptloan.LoanReserveAmount.Value;

                    //Format Cells

                    pLoanNumber.Format.HorizontalAlignment = Spire.Doc.Documents.HorizontalAlignment.Center;
                    pBorrower.Format.HorizontalAlignment = Spire.Doc.Documents.HorizontalAlignment.Center;
                    pBusiness.Format.HorizontalAlignment = Spire.Doc.Documents.HorizontalAlignment.Center;
                    pAddress.Format.HorizontalAlignment = Spire.Doc.Documents.HorizontalAlignment.Center;
                    pMortgageAmount.Format.HorizontalAlignment = Spire.Doc.Documents.HorizontalAlignment.Right;
                    pReserveAmount.Format.HorizontalAlignment = Spire.Doc.Documents.HorizontalAlignment.Right;
                    pAdvanceAmount.Format.HorizontalAlignment = Spire.Doc.Documents.HorizontalAlignment.Right;
                    //TRLoanNumber.CharacterFormat.FontName = "Calibri";
                    //TRBorrower.CharacterFormat.FontName = "Calibri";
                    //TRAddress.CharacterFormat.FontName = "Calibri";

                    TRLoanNumber.CharacterFormat.FontSize = 12;
                    TRBorrower.CharacterFormat.FontSize = 12;
                    TRAddress.CharacterFormat.FontSize = 12;

                    TRAddress.CharacterFormat.FontSize = 12;
                    trMortgageAmount.CharacterFormat.FontSize = 12;
                    trReserveAmount.CharacterFormat.FontSize = 12;
                    trAdvanceAmount.CharacterFormat.FontSize = 12;
                    //TR2.CharacterFormat.TextColor = Color.Brown;

                    //}
                    r++;
                }

                Spire.Doc.TableRow FRowFooterMortgage = tblSch.Rows[footerRow];
                Spire.Doc.Documents.Paragraph pFooterMortgage = FRowFooterMortgage.Cells[4].AddParagraph();
                Spire.Doc.Fields.TextRange TRFooterMortgage = pFooterMortgage.AppendText(FormatNumberCommas2dec(totalMortgage.ToString()));
                pFooterMortgage.Format.HorizontalAlignment = Spire.Doc.Documents.HorizontalAlignment.Right;
                TRFooterMortgage.CharacterFormat.Bold = true;

                Spire.Doc.TableRow FRowFooterReserve = tblSch.Rows[footerRow];
                Spire.Doc.Documents.Paragraph pFooterReserve = FRowFooterReserve.Cells[5].AddParagraph();
                Spire.Doc.Fields.TextRange TRFooterReserve = pFooterReserve.AppendText(FormatNumberCommas2dec(totalReserve.ToString()));
                pFooterReserve.Format.HorizontalAlignment = Spire.Doc.Documents.HorizontalAlignment.Right;
                TRFooterReserve.CharacterFormat.Bold = true;

                Spire.Doc.TableRow FRowFooterAdvance = tblSch.Rows[footerRow];
                Spire.Doc.Documents.Paragraph pFooterAdvance = FRowFooterAdvance.Cells[6].AddParagraph();
                Spire.Doc.Fields.TextRange TRFooterAdvance = pFooterAdvance.AppendText(FormatNumberCommas2dec(totalAdvance.ToString()));
                pFooterAdvance.Format.HorizontalAlignment = Spire.Doc.Documents.HorizontalAlignment.Right;
                TRFooterAdvance.CharacterFormat.Bold = true;

                //
                // End Table Logic
                //

                Spire.Doc.Documents.Paragraph parEscrow = section1.AddParagraph();
                parEscrow.Format.HorizontalAlignment = Spire.Doc.Documents.HorizontalAlignment.Left;

                parEscrow.AppendText("Wire Sent to LaRocca Escrow Account");

                DateTime today = DateTime.Today;
                string todayStr;
                todayStr = string.Format("{0:MMMM dd, yyyy}", today);

                string fileName;
                string todayStrFile = string.Format("{0:yyyy_MM_dd}", today);

                fileName = "Advance_" + client.ClientName + "_" + entity.EntityName + "_" + todayStrFile + ".docx";
                //CommonOpenFileDialog dialog = new CommonOpenFileDialog();
                //dialog.InitialDirectory = "C:\\Users";
                //dialog.IsFolderPicker = true;

                //if (dialog.ShowDialog() == CommonFileDialogResult.Ok)
                //{
                //    fileName = dialog.FileName + "\\" + fileName;

                //}
                fileName = @"C:\Users\pdean\Reports\\Advance_" + client.ClientName + "_" + entity.EntityName + "_" + loanlist + ".docx";

                document.SaveToFile(fileName, FileFormat.Docx);

                System.Diagnostics.Process.Start(fileName);
            }
            catch (Exception ex)
            {
                TempData["Error"] = "Error - Advance Report Error";
                return Json("");
            }
            TempData["Message"] = "Advance Report Successfully Created";
            return Json("");
        }
        public ActionResult RemittanceReport(string loanlist)
        {
            try
            {
                List<string> loansSplit = loanlist.Split(',').ToList();

                List<Int32> modLoansSplit = new List<Int32>();
                foreach (string loanID in loansSplit)
                {
                    int test;
                    int.TryParse(loanID, out test);
                    if (loanID.Length > 0 && test > 0)
                        modLoansSplit.Add(Convert.ToInt32(loanID));
                }
                if (modLoansSplit.Count() == 0)
                {
                    TempData["Error"] = "Error - Please select at least one loan";
                    return Json("");
                }

                List<Loan> loans = new List<Loan>();
                foreach (Int32 loanID in modLoansSplit)
                {
                    Loan loan = _loanRepository.GetLoan(loanID);
                    loans.Add(loan);

                }
                Client client = _loanRepository.GetClient(loans[0].ClientID.Value);
                Entity entity = _loanRepository.GetEntity(loans[0].EntityID.Value);


                if (loans[0].BaileeLetterDate == null)
                {
                    //ErrorLabel.Content = "Bailee Letter Date Missing";
                    //ErrorLabel.Foreground = new SolidColorBrush(Colors.Red);
                    //return;
                }
                DateTime baileeLetterDate = (DateTime)loans[0].BaileeLetterDate;

                string baileeLetterDateStr = string.Format("{0:MMMM dd, yyyy}", baileeLetterDate);
                Spire.Doc.Document document = new Spire.Doc.Document();
                //Add Section

                Spire.Doc.Section section1 = document.AddSection();

                section1.PageSetup.Orientation = Spire.Doc.Documents.PageOrientation.Landscape;
                //Add Paragraph

                Spire.Doc.Documents.Paragraph parSchedule1 = section1.AddParagraph();
                parSchedule1.Format.HorizontalAlignment = Spire.Doc.Documents.HorizontalAlignment.Center;
                Spire.Doc.Fields.TextRange trSchedule1 = parSchedule1.AppendText(entity.EntityName);
                //trSchedule1.CharacterFormat.UnderlineStyle = Spire.Doc.Documents.UnderlineStyle.Single;

                Spire.Doc.Documents.Paragraph parMort = section1.AddParagraph();
                parMort.Format.HorizontalAlignment = Spire.Doc.Documents.HorizontalAlignment.Center;

                parMort.AppendText("Remittance Report");

                Spire.Doc.Documents.Paragraph parClient = section1.AddParagraph();
                parClient.Format.HorizontalAlignment = Spire.Doc.Documents.HorizontalAlignment.Center;
                parClient.AppendText(client.ClientName);

                Spire.Doc.Documents.Paragraph parRemitReport = section1.AddParagraph();
                parRemitReport.Format.HorizontalAlignment = Spire.Doc.Documents.HorizontalAlignment.Left;
                parRemitReport.AppendText("Remittance Date " + DateTime.Today.ToShortDateString() + "\v");

                Spire.Doc.Table tblSch = section1.AddTable(true);
                //
                // Table Logic
                //

                String[] Header = { "Business Name", "Client Name", "Loan Number", "Interest Percentage", "Mortgage Amount", "Proceeds", "Loan Amount", "Interest", "Origination Fee", "Underwriting/ Administrative Fee", "Total Transfer" };

                //Add Cells
                // Number of rows is # of loans + header + footer + line for wire fee
                tblSch.ResetCells(loans.Count() + 3, Header.Length);
                int footerRow;
                int wireRow;

                footerRow = loans.Count() + 2;
                wireRow = loans.Count() + 1;
                //Header Row
                Spire.Doc.TableRow FRow2 = tblSch.Rows[0];
                FRow2.IsHeader = true;

                //Row Height
                FRow2.Height = 23;

                //Wire Fee Row
                Spire.Doc.TableRow FRowWire = tblSch.Rows[wireRow];

                Spire.Doc.Documents.Paragraph pWire = FRowWire.Cells[0].AddParagraph();
                Spire.Doc.Fields.TextRange TRWire = pWire.AppendText("Wire Transfer Fee If Applicable");
                pWire.Format.HorizontalAlignment = Spire.Doc.Documents.HorizontalAlignment.Left;
                TRWire.CharacterFormat.FontSize = 12;

                Spire.Doc.Documents.Paragraph pWireFee = FRowWire.Cells[10].AddParagraph();
                Spire.Doc.Fields.TextRange TRWireFee = pWireFee.AppendText("(25.00)");
                pWireFee.Format.HorizontalAlignment = Spire.Doc.Documents.HorizontalAlignment.Right;
                TRWireFee.CharacterFormat.FontSize = 12;

                //Merge first 4 cells
                tblSch.ApplyHorizontalMerge((wireRow), 0, 8);

                Spire.Doc.TableRow FRowFooter = tblSch.Rows[footerRow];

                //Merge first 3 cells
                tblSch.ApplyHorizontalMerge((footerRow), 0, 2);

                Spire.Doc.Documents.Paragraph pFooter = FRowFooter.Cells[0].AddParagraph();
                Spire.Doc.Fields.TextRange TRFooter = pFooter.AppendText("Total Transfer Due " + client.ClientName);
                pFooter.Format.HorizontalAlignment = Spire.Doc.Documents.HorizontalAlignment.Left;
                TRFooter.CharacterFormat.FontSize = 12;

                //Row Height
                FRowFooter.Height = 23;

                //Header Format
                //FRow.RowFormat.BackColor = Color.AliceBlue;

                for (int i = 0; i < Header.Length; i++)
                {
                    //Cell Alignment
                    Spire.Doc.Documents.Paragraph p = FRow2.Cells[i].AddParagraph();
                    FRow2.Cells[i].CellFormat.VerticalAlignment = Spire.Doc.Documents.VerticalAlignment.Middle;

                    p.Format.HorizontalAlignment = Spire.Doc.Documents.HorizontalAlignment.Center;
                    //Data Format
                    Spire.Doc.Fields.TextRange TR = p.AppendText(Header[i]);
                    //TR.CharacterFormat.FontName = "Calibri";
                    TR.CharacterFormat.FontSize = 12;
                    //TR.CharacterFormat.TextColor = Color.Teal;                
                    TR.CharacterFormat.Bold = true;

                }

                //Data Row

                double totalMortgage = 0;
                double totalProceeds = 0;
                double totalLoanAmount = 0;
                double totalInterest = 0;
                double totalOriginationFee = 0;
                double totalUnderwritingFee = 0;
                double totalTransfer = -25; // -25 for wire fee

                int r = 0;
                //for (int r = 0; r < data.Length; r++)
                foreach (Loan rptloan in loans)
                {

                    Spire.Doc.TableRow DataRow = tblSch.Rows[r + 1];

                    //Row Height

                    DataRow.Height = 20;

                    //C Represents Column.

                    //for (int c = 0; c < data[r].Length; c++)                    
                    //{                    
                    //Cell Alignment
                    DataRow.Cells[0].CellFormat.VerticalAlignment = Spire.Doc.Documents.VerticalAlignment.Middle;

                    DataRow.Cells[1].CellFormat.VerticalAlignment = Spire.Doc.Documents.VerticalAlignment.Middle;
                    DataRow.Cells[2].CellFormat.VerticalAlignment = Spire.Doc.Documents.VerticalAlignment.Middle;
                    DataRow.Cells[3].CellFormat.VerticalAlignment = Spire.Doc.Documents.VerticalAlignment.Middle;
                    DataRow.Cells[4].CellFormat.VerticalAlignment = Spire.Doc.Documents.VerticalAlignment.Middle;
                    DataRow.Cells[5].CellFormat.VerticalAlignment = Spire.Doc.Documents.VerticalAlignment.Middle;
                    DataRow.Cells[6].CellFormat.VerticalAlignment = Spire.Doc.Documents.VerticalAlignment.Middle;
                    DataRow.Cells[7].CellFormat.VerticalAlignment = Spire.Doc.Documents.VerticalAlignment.Middle;
                    DataRow.Cells[8].CellFormat.VerticalAlignment = Spire.Doc.Documents.VerticalAlignment.Middle;
                    DataRow.Cells[9].CellFormat.VerticalAlignment = Spire.Doc.Documents.VerticalAlignment.Middle;
                    DataRow.Cells[10].CellFormat.VerticalAlignment = Spire.Doc.Documents.VerticalAlignment.Middle;



                    //Fill Data in Rows
                    Spire.Doc.Documents.Paragraph pBusiness = DataRow.Cells[0].AddParagraph();
                    Spire.Doc.Fields.TextRange TRBusiness = pBusiness.AppendText(rptloan.LoanMortgageeBusiness);

                    Spire.Doc.Documents.Paragraph pBorrower = DataRow.Cells[1].AddParagraph();
                    Spire.Doc.Fields.TextRange TRBorrower = pBorrower.AppendText(rptloan.LoanMortgagee);

                    Spire.Doc.Documents.Paragraph pLoanNumber = DataRow.Cells[2].AddParagraph();
                    Spire.Doc.Fields.TextRange TRLoanNumber = pLoanNumber.AppendText(rptloan.LoanNumber);

                    Spire.Doc.Documents.Paragraph pAddress = DataRow.Cells[3].AddParagraph();
                    Spire.Doc.Fields.TextRange TRAddress = pAddress.AppendText(FormatPcnt(rptloan.LoanMinimumInterest.ToString()));

                    Spire.Doc.Documents.Paragraph pMortgageAmount = DataRow.Cells[4].AddParagraph();
                    Spire.Doc.Fields.TextRange trMortgageAmount = pMortgageAmount.AppendText(FormatNumberCommas2dec(rptloan.LoanMortgageAmount.ToString()));

                    Spire.Doc.Documents.Paragraph pProceeds = DataRow.Cells[5].AddParagraph();
                    Spire.Doc.Fields.TextRange trProceeds = pProceeds.AppendText(FormatNumberCommas2dec(rptloan.InvestorProceeds.ToString()));

                    Spire.Doc.Documents.Paragraph pLoanAmount = DataRow.Cells[6].AddParagraph();
                    Spire.Doc.Fields.TextRange trLoanAmount = pLoanAmount.AppendText(FormatNumberCommas2dec(rptloan.LoanAdvanceAmount.ToString()));

                    double interestFee = rptloan.InterestFee.Value + rptloan.CustSvcInterestDiscount.Value;
                    Spire.Doc.Documents.Paragraph pInterest = DataRow.Cells[7].AddParagraph();
                    Spire.Doc.Fields.TextRange trInterest = pInterest.AppendText(FormatNumberCommas2dec(interestFee.ToString()));

                    Spire.Doc.Documents.Paragraph pOrig = DataRow.Cells[8].AddParagraph();
                    double originationFee = rptloan.OriginationFee.Value + rptloan.CustSvcOriginationDiscount.Value;
                    Spire.Doc.Fields.TextRange trOrig = pOrig.AppendText(FormatNumberCommas2dec(originationFee.ToString()));

                    double underwritingFee = rptloan.UnderwritingFee.Value + rptloan.CustSvcUnderwritingDiscount.Value;
                    Spire.Doc.Documents.Paragraph pUW = DataRow.Cells[9].AddParagraph();
                    Spire.Doc.Fields.TextRange trUW = pUW.AppendText(FormatNumberCommas2dec(underwritingFee.ToString()));

                    Spire.Doc.Documents.Paragraph pTotalTransfer = DataRow.Cells[10].AddParagraph();
                    //Spire.Doc.Fields.TextRange trTotalTransfer = pTotalTransfer.AppendText(FormatNumberCommas2dec(loan.TotalTransfer.ToString()));

                    totalMortgage += rptloan.LoanMortgageAmount.Value;
                    totalProceeds += rptloan.InvestorProceeds.Value;
                    totalLoanAmount += rptloan.LoanAdvanceAmount.Value;
                    totalInterest += rptloan.InterestFee.Value + rptloan.CustSvcInterestDiscount.Value;
                    totalOriginationFee += rptloan.OriginationFee.Value + rptloan.CustSvcOriginationDiscount.Value;
                    totalUnderwritingFee += rptloan.UnderwritingFee.Value + rptloan.CustSvcUnderwritingDiscount.Value;
                    totalTransfer += rptloan.TotalTransfer.Value;

                    //Format Cells

                    pLoanNumber.Format.HorizontalAlignment = Spire.Doc.Documents.HorizontalAlignment.Center;
                    pBorrower.Format.HorizontalAlignment = Spire.Doc.Documents.HorizontalAlignment.Center;
                    pAddress.Format.HorizontalAlignment = Spire.Doc.Documents.HorizontalAlignment.Center;
                    pMortgageAmount.Format.HorizontalAlignment = Spire.Doc.Documents.HorizontalAlignment.Right;
                    pProceeds.Format.HorizontalAlignment = Spire.Doc.Documents.HorizontalAlignment.Right;
                    pLoanAmount.Format.HorizontalAlignment = Spire.Doc.Documents.HorizontalAlignment.Right;
                    pInterest.Format.HorizontalAlignment = Spire.Doc.Documents.HorizontalAlignment.Right;
                    pOrig.Format.HorizontalAlignment = Spire.Doc.Documents.HorizontalAlignment.Right;
                    pUW.Format.HorizontalAlignment = Spire.Doc.Documents.HorizontalAlignment.Right;
                    pTotalTransfer.Format.HorizontalAlignment = Spire.Doc.Documents.HorizontalAlignment.Right;
                    //TRLoanNumber.CharacterFormat.FontName = "Calibri";
                    //TRBorrower.CharacterFormat.FontName = "Calibri";
                    //TRAddress.CharacterFormat.FontName = "Calibri";

                    TRLoanNumber.CharacterFormat.FontSize = 12;
                    TRBorrower.CharacterFormat.FontSize = 12;
                    //TRAddress.CharacterFormat.FontSize = 12;

                    //TRAddress.CharacterFormat.FontSize = 12;
                    trMortgageAmount.CharacterFormat.FontSize = 12;
                    trProceeds.CharacterFormat.FontSize = 12;
                    //trLoanAmount.CharacterFormat.FontSize = 12;
                    trInterest.CharacterFormat.FontSize = 12;
                    trOrig.CharacterFormat.FontSize = 12;
                    trUW.CharacterFormat.FontSize = 12;
                    //trTotalTransfer.CharacterFormat.FontSize = 12;
                    //TR2.CharacterFormat.TextColor = Color.Brown;

                    //}
                    r++;
                }

                Spire.Doc.TableRow FRowFooterMortgage = tblSch.Rows[footerRow];
                Spire.Doc.Documents.Paragraph pFooterMortgage = FRowFooterMortgage.Cells[4].AddParagraph();
                Spire.Doc.Fields.TextRange TRFooterMortgage = pFooterMortgage.AppendText(FormatNumberCommas2dec(totalMortgage.ToString()));
                pFooterMortgage.Format.HorizontalAlignment = Spire.Doc.Documents.HorizontalAlignment.Right;
                TRFooterMortgage.CharacterFormat.Bold = true;
                TRFooterMortgage.CharacterFormat.FontSize = 12;

                Spire.Doc.TableRow FRowFooterProceeds = tblSch.Rows[footerRow];
                Spire.Doc.Documents.Paragraph pFooterProceeds = FRowFooterProceeds.Cells[5].AddParagraph();
                Spire.Doc.Fields.TextRange TRFooterProceeds = pFooterProceeds.AppendText(FormatNumberCommas2dec(totalProceeds.ToString()));
                pFooterProceeds.Format.HorizontalAlignment = Spire.Doc.Documents.HorizontalAlignment.Right;
                TRFooterProceeds.CharacterFormat.Bold = true;
                TRFooterProceeds.CharacterFormat.FontSize = 12;

                Spire.Doc.TableRow FRowFooterLoanAmount = tblSch.Rows[footerRow];
                Spire.Doc.Documents.Paragraph pFooterLoanAmount = FRowFooterLoanAmount.Cells[6].AddParagraph();
                Spire.Doc.Fields.TextRange TRFooterLoanAmount = pFooterLoanAmount.AppendText(FormatNumberCommas2dec(totalLoanAmount.ToString()));
                pFooterLoanAmount.Format.HorizontalAlignment = Spire.Doc.Documents.HorizontalAlignment.Right;
                TRFooterLoanAmount.CharacterFormat.Bold = true;
                TRFooterLoanAmount.CharacterFormat.FontSize = 12;

                Spire.Doc.TableRow FRowFooterInterest = tblSch.Rows[footerRow];
                Spire.Doc.Documents.Paragraph pFooterInterest = FRowFooterInterest.Cells[7].AddParagraph();
                Spire.Doc.Fields.TextRange TRFooterInterest = pFooterInterest.AppendText(FormatNumberCommas2dec(totalInterest.ToString()));
                pFooterInterest.Format.HorizontalAlignment = Spire.Doc.Documents.HorizontalAlignment.Right;
                TRFooterInterest.CharacterFormat.Bold = true;
                TRFooterInterest.CharacterFormat.FontSize = 12;

                Spire.Doc.TableRow FRowFooterOrig = tblSch.Rows[footerRow];
                Spire.Doc.Documents.Paragraph pFooterOrig = FRowFooterOrig.Cells[8].AddParagraph();
                Spire.Doc.Fields.TextRange TRFooterOrig = pFooterOrig.AppendText(FormatNumberCommas2dec(totalOriginationFee.ToString()));
                pFooterOrig.Format.HorizontalAlignment = Spire.Doc.Documents.HorizontalAlignment.Right;
                TRFooterOrig.CharacterFormat.Bold = true;
                TRFooterOrig.CharacterFormat.FontSize = 12;

                Spire.Doc.TableRow FRowFooterUW = tblSch.Rows[footerRow];
                Spire.Doc.Documents.Paragraph pFooterUW = FRowFooterUW.Cells[9].AddParagraph();
                Spire.Doc.Fields.TextRange TRFooterUW = pFooterUW.AppendText(FormatNumberCommas2dec(totalUnderwritingFee.ToString()));
                pFooterUW.Format.HorizontalAlignment = Spire.Doc.Documents.HorizontalAlignment.Right;
                TRFooterUW.CharacterFormat.Bold = true;
                TRFooterUW.CharacterFormat.FontSize = 12;

                Spire.Doc.TableRow FRowFooterTotalTransfer = tblSch.Rows[footerRow];
                Spire.Doc.Documents.Paragraph pFooterTotalTransfer = FRowFooterTotalTransfer.Cells[10].AddParagraph();
                Spire.Doc.Fields.TextRange TRFooterTotalTransfer = pFooterTotalTransfer.AppendText(FormatNumberCommas2dec(totalTransfer.ToString()));
                pFooterTotalTransfer.Format.HorizontalAlignment = Spire.Doc.Documents.HorizontalAlignment.Right;
                TRFooterTotalTransfer.CharacterFormat.Bold = true;
                TRFooterTotalTransfer.CharacterFormat.FontSize = 12;

                //
                // End Table Logic
                //

                Spire.Doc.Documents.Paragraph parEscrow = section1.AddParagraph();
                parEscrow.Format.HorizontalAlignment = Spire.Doc.Documents.HorizontalAlignment.Left;

                parEscrow.AppendText("Wire Sent to " + client.ClientName);

                //document.SaveToFile("C:/Temp/result.docx", FileFormat.Docx);

                //Spire.Doc.Document DocOne = new Spire.Doc.Document();

                //DocOne.LoadFromFile("C:/Temp/Release Text.docx", FileFormat.Docx);

                //DocOne.Replace("TODAYSDATE", todayStr, false, true);

                //DocOne.Replace("BAILEELETTERDATE", baileeLetterDateStr, false, true);

                //foreach (Spire.Doc.Section sec in document.Sections)
                //{
                //    DocOne.Sections.Add(sec.Clone());
                //}
                DateTime today = DateTime.Today;
                string todayStr;
                todayStr = string.Format("{0:MMMM dd, yyyy}", today);

                string fileName;
                string todayStrFile = string.Format("{0:yyyy_MM_dd}", today);

                //CommonOpenFileDialog dialog = new CommonOpenFileDialog();
                //dialog.InitialDirectory = "C:\\Users";
                //dialog.IsFolderPicker = true;

                //if (dialog.ShowDialog() == CommonFileDialogResult.Ok)
                //{
                //    fileName = dialog.FileName + "\\" + fileName;

                //}

                fileName = @"C:\Users\pdean\Reports\\Remittance_" + client.ClientName + "_" + entity.EntityName + "_" + loanlist + ".docx";

                document.SaveToFile(fileName, FileFormat.Docx);

                System.Diagnostics.Process.Start(fileName);
            }
            catch (Exception ex)
            {
                TempData["Error"] = "Error - Remittance Report Error";
                return Json("");
            }
            TempData["Message"] = "Remittance Report Successfully Created";
            return Json("");
        }
        public ActionResult ReleaseReport(string loanlist)
        {
            try
            {
                List<string> loansSplit = loanlist.Split(',').ToList();

                List<Int32> modLoansSplit = new List<Int32>();
                foreach (string loanID in loansSplit)
                {
                    int test;
                    int.TryParse(loanID, out test);
                    if (loanID.Length > 0 && test > 0)
                        modLoansSplit.Add(Convert.ToInt32(loanID));
                }
                if (modLoansSplit.Count() == 0)
                {
                    TempData["Error"] = "Error - Please select at least one loan";
                    return Json("");
                }

                List<Loan> loans = new List<Loan>();
                foreach (Int32 loanID in modLoansSplit)
                {
                    Loan loan = _loanRepository.GetLoan(loanID);
                    loans.Add(loan);

                }
                Client client = _loanRepository.GetClient(loans[0].ClientID.Value);
                Entity entity = _loanRepository.GetEntity(loans[0].EntityID.Value);



                if (loans[0].BaileeLetterDate == null)
                {
                   
                    TempData["Error"] = "Error - Bailee Letter Date Missing";
                    return Json("");
                }

                DateTime baileeLetterDate = (DateTime)loans[0].BaileeLetterDate;

                string baileeLetterDateStr = string.Format("{0:MMMM dd, yyyy}", baileeLetterDate);
                Spire.Doc.Document document = new Spire.Doc.Document();
                //Add Section

                Spire.Doc.Section section1 = document.AddSection();
                section1.PageSetup.Orientation = Spire.Doc.Documents.PageOrientation.Portrait;

                //Add Paragraph

                Spire.Doc.Documents.Paragraph parSchedule1 = section1.AddParagraph();
                parSchedule1.Format.HorizontalAlignment = Spire.Doc.Documents.HorizontalAlignment.Center;
                Spire.Doc.Fields.TextRange trSchedule1 = parSchedule1.AppendText("Schedule 1\v");
                trSchedule1.CharacterFormat.UnderlineStyle = Spire.Doc.Documents.UnderlineStyle.Single;

                Spire.Doc.Documents.Paragraph parMort = section1.AddParagraph();
                parMort.Format.HorizontalAlignment = Spire.Doc.Documents.HorizontalAlignment.Center;

                parMort.AppendText("Mortgage Loan Schedule\v");

                Spire.Doc.Table tblSch = section1.AddTable(true);
                //
                // Table Logic
                //

                String[] Header = { "Loan ID", "Borrower", "Address" };

                //Add Cells
                tblSch.ResetCells(loans.Count() + 1, Header.Length);
                //Header Row
                Spire.Doc.TableRow FRow2 = tblSch.Rows[0];
                FRow2.IsHeader = true;

                //Row Height
                FRow2.Height = 23;

                //Header Format
                //FRow.RowFormat.BackColor = Color.AliceBlue;

                for (int i = 0; i < Header.Length; i++)
                {
                    //Cell Alignment
                    Spire.Doc.Documents.Paragraph p = FRow2.Cells[i].AddParagraph();
                    FRow2.Cells[i].CellFormat.VerticalAlignment = Spire.Doc.Documents.VerticalAlignment.Middle;
                    p.Format.HorizontalAlignment = Spire.Doc.Documents.HorizontalAlignment.Center;
                    //Data Format
                    Spire.Doc.Fields.TextRange TR = p.AppendText(Header[i]);
                    //TR.CharacterFormat.FontName = "Calibri";
                    TR.CharacterFormat.FontSize = 14;
                    //TR.CharacterFormat.TextColor = Color.Teal;                
                    TR.CharacterFormat.Bold = true;

                }

                //Data Row
                string loanNumberList = "";
                int r = 0;
                //for (int r = 0; r < data.Length; r++)
                foreach (Loan rptLoan in loans)
                {

                    Spire.Doc.TableRow DataRow = tblSch.Rows[r + 1];

                    //Row Height

                    DataRow.Height = 20;

                    //C Represents Column.

                    //for (int c = 0; c < data[r].Length; c++)                    
                    //{                    
                    //Cell Alignment
                    DataRow.Cells[0].CellFormat.VerticalAlignment = Spire.Doc.Documents.VerticalAlignment.Middle;
                    DataRow.Cells[1].CellFormat.VerticalAlignment = Spire.Doc.Documents.VerticalAlignment.Middle;
                    DataRow.Cells[2].CellFormat.VerticalAlignment = Spire.Doc.Documents.VerticalAlignment.Middle;

                    //Fill Data in Rows
                    loanNumberList += rptLoan.LoanNumber + " ";

                    Spire.Doc.Documents.Paragraph pLoanNumber = DataRow.Cells[0].AddParagraph();
                    Spire.Doc.Fields.TextRange TRLoanNumber = pLoanNumber.AppendText(rptLoan.LoanNumber);

                    Spire.Doc.Documents.Paragraph pBorrower = DataRow.Cells[1].AddParagraph();
                    Spire.Doc.Fields.TextRange TRBorrower = pBorrower.AppendText(rptLoan.LoanMortgagee);

                    Spire.Doc.Documents.Paragraph pAddress = DataRow.Cells[2].AddParagraph();
                    Spire.Doc.Fields.TextRange TRAddress = pAddress.AppendText(rptLoan.LoanPropertyAddress);

                    //Format Cells

                    pLoanNumber.Format.HorizontalAlignment = Spire.Doc.Documents.HorizontalAlignment.Center;
                    pBorrower.Format.HorizontalAlignment = Spire.Doc.Documents.HorizontalAlignment.Center;
                    pAddress.Format.HorizontalAlignment = Spire.Doc.Documents.HorizontalAlignment.Center;

                    //TRLoanNumber.CharacterFormat.FontName = "Calibri";
                    //TRBorrower.CharacterFormat.FontName = "Calibri";
                    //TRAddress.CharacterFormat.FontName = "Calibri";

                    TRLoanNumber.CharacterFormat.FontSize = 12;
                    TRBorrower.CharacterFormat.FontSize = 12;
                    TRAddress.CharacterFormat.FontSize = 12;

                    //TR2.CharacterFormat.TextColor = Color.Brown;

                    //}
                    r++;
                }
                //
                // End Table Logic
                //

                //document.SaveToFile(MyVariables.letterFilePath + "result.docx", FileFormat.Docx);

                Spire.Doc.Document DocOne = new Spire.Doc.Document();

                //OpenFileDialog fileDialog = new OpenFileDialog();
                //string releaseFileName = "";
                ////releaseFileName = "Release_" + client.ClientName + "_" + entity.EntityName + ".docx";

                //if (fileDialog.ShowDialog() == true)
                //{
                //    releaseFileName = fileDialog.FileName;
                //}

                //DocOne.LoadFromFile(releaseFileName, FileFormat.Docx);
                DocOne.LoadFromFile(@"C:\Users\pdean\LetterTemplates\Release_Northwind Financial_RAI Funding, LLC.docx", FileFormat.Docx);

                DateTime today = DateTime.Today;
                string todayStr;
                todayStr = string.Format("{0:MMMM dd, yyyy}", today);

                DocOne.Replace("TODAYSDATE", todayStr, false, true);

                DocOne.Replace("BAILEELETTERDATE", baileeLetterDateStr, false, true);

                foreach (Spire.Doc.Section sec in document.Sections)
                {
                    DocOne.Sections.Add(sec.Clone());
                }
                string fileName;
                string todayStrFile = string.Format("{0:yyyy_MM_dd}", today);


                //CommonOpenFileDialog dialog = new CommonOpenFileDialog();
                //dialog.InitialDirectory = "C:\\Users";
                //dialog.IsFolderPicker = true;

                //if (dialog.ShowDialog() == CommonFileDialogResult.Ok)
                //{
                //    fileName = dialog.FileName + "\\" + fileName;

                //}
                fileName = @"C:\Users\pdean\Reports\\Release_" + client.ClientName + "_" + entity.EntityName + "_" + loanlist + ".docx";

                DocOne.SaveToFile(fileName, FileFormat.Docx);

                System.Diagnostics.Process.Start(fileName);
            }
            catch (Exception ex)
            {
                TempData["Error"] = "Error - Release Report Error";
                return Json("");
            }
            TempData["Message"] = "Release Report Successfully Created";
            return Json("");
        }
        public ActionResult Add()
        {
            var loan = new Loan();

            SetupSelectListItems();
            return View(loan);
        }
        [HttpPost]
        public ActionResult Add(Loan loan)
        {
            ValidateLoan(loan);
            if (ModelState.IsValid)
            {
                _loanRepository.AddLoan(loan);
                TempData["Message"] = "Loan Successfully Added";
                return Redirect("Index");
            }
            SetupSelectListItems();
            return View(loan);
        }
        public ActionResult AddClient()
        {
            var client = new Client();

            SetupSelectListItems();
            return View(client);
        }
        [HttpPost]
        public ActionResult AddClient(Client client)
        {
            if (ModelState.IsValid)
            {
                _clientsRepository.AddClient(client);
                TempData["Message"] = "Client Successfully Added";
                return Redirect("Clients");
            }
            SetupSelectListItems();
            return View(client);
        }

        public ActionResult AddEntity()
        {
            var entity = new Entity();

            SetupSelectListItems();
            return View(entity);
        }
        [HttpPost]
        public ActionResult AddEntity(Entity entity)
        {
            if (ModelState.IsValid)
            {
                _entityRepository.AddEntity(entity);
                TempData["Message"] = "Entity Successfully Added";
                return Redirect("Entities");
            }
            SetupSelectListItems();
            return View(entity);
        }

        public ActionResult AddInvestor()
        {
            var investor = new Investor();

            SetupSelectListItems();
            return View(investor);
        }
        [HttpPost]
        public ActionResult AddInvestor(Investor investor)
        {
            if (ModelState.IsValid)
            {
                _investorRepository.AddInvestor(investor);
                TempData["Message"] = "Investor Successfully Added";
                return Redirect("Investors");
            }
            SetupSelectListItems();
            return View(investor);
        }
        //public ActionResult AddRole()
        //{
        //    var role = new Role();

        //    SetupSelectListItems();
        //    return View(role);
        //}
        //[HttpPost]
        //public ActionResult AddRole(Role role)
        //{
        //    if (ModelState.IsValid)
        //    {
        //        _usersRepository.AddRole(role);
        //        TempData["Message"] = "Role Successfully Added";
        //        return Redirect("Roles");
        //    }
        //    SetupSelectListItems();
        //    return View(role);
        //}
        //public ActionResult AddUser()
        //{
        //    var user = new User();

        //    SetupSelectListItems();
        //    return View(user);
        //}
        //[HttpPost]
        //public ActionResult AddUser(User user)
        //{
        //    if (ModelState.IsValid)
        //    {
        //        _usersRepository.AddUser(user);
        //        TempData["Message"] = "User Successfully Added";
        //        return Redirect("USers");
        //    }
        //    SetupSelectListItems();
        //    return View(user);
        //}

        public ActionResult Investors()
        {
            var investors = _investorRepository.GetInvestors();
            return View(investors);
        }
        public ActionResult EditInvestor(int? id)
        {
            if (id == null)
            {
                return new HttpStatusCodeResult(HttpStatusCode.BadRequest);
            }

            Investor investor = _investorRepository.GetInvestor((int)id);
            if (investor == null)
            {
                return HttpNotFound();
            }
            SetupSelectListItems();
            return View(investor);
        }
        [HttpPost]
        public ActionResult EditInvestor(Investor investor)
        {
            if (ModelState.IsValid)
            {
                _investorRepository.UpdateInvestor(investor);

                TempData["Message"] = "Investor Successfully Updated";
                return RedirectToAction("Investors");
            }

            return View(investor);
        }
        public ActionResult DeleteInvestor(int id)
        {
            _investorRepository.DeleteInvestor(id);

            TempData["Message"] = "Investor Successfully Deleted";
            return RedirectToAction("Investors");
        }

        public ActionResult Entities()
        {
            var entities = _entityRepository.GetEntities();
            return View(entities);
        }
        public ActionResult EditEntity(int? id)
        {
            if (id == null)
            {
                return new HttpStatusCodeResult(HttpStatusCode.BadRequest);
            }

            Entity entity = _entityRepository.GetEntity((int)id);
            if (entity == null)
            {
                return HttpNotFound();
            }
            SetupSelectListItems();
            return View(entity);
        }
        [HttpPost]
        public ActionResult EditEntity(Entity entity)
        {
            if (ModelState.IsValid)
            {
                _entityRepository.UpdateEntity(entity);

                TempData["Message"] = "Entity Successfully Updated";
                return RedirectToAction("Entities");
            }

            return View(entity);
        }
        public ActionResult DeleteEntity(int id)
        {
            _entityRepository.DeleteEntity(id);

            TempData["Message"] = "Entity Successfully Deleted";
            return RedirectToAction("Entities");
        }

        public ActionResult Clients()
        {
            var clients = _clientsRepository.GetClients();
            return View(clients);
        }
        public ActionResult EditClient(int? id)
        {
            if (id == null)
            {
                return new HttpStatusCodeResult(HttpStatusCode.BadRequest);
            }

            Client client = _clientsRepository.GetClient((int)id);
            if (client == null)
            {
                return HttpNotFound();
            }
            SetupSelectListItems();
            return View(client);
        }
        [HttpPost]
        public ActionResult EditClient(Client client)
        {
            if (ModelState.IsValid)
            {
                _clientsRepository.UpdateClient(client);

                TempData["Message"] = "Client Successfully Updated";
                return RedirectToAction("Clients");
            }

            return View(client);
        }
        public ActionResult DeleteClient(int id)
        {
            _clientsRepository.DeleteClient(id);

            TempData["Message"] = "Client Successfully Deleted";
            return RedirectToAction("Clients");
        }

        //public ActionResult Users()
        //{
        //    var clients = _usersRepository.GetUsers();
        //    return View(clients);
        //}

        //public ActionResult EditUser(int? id)
        //{
        //    if (id == null)
        //    {
        //        return new HttpStatusCodeResult(HttpStatusCode.BadRequest);
        //    }

        //    User user = _usersRepository.GetUser((int)id);
        //    if (user == null)
        //    {
        //        return HttpNotFound();
        //    }
        //    SetupRoleSelectListItems();
        //    return View(user);
        //}
        //[HttpPost]
        //public ActionResult EditUser(User user)
        //{
        //    if (ModelState.IsValid)
        //    {
        //        _usersRepository.UpdateUser(user);

        //        TempData["Message"] = "User Successfully Updated";
        //        return RedirectToAction("Users");
        //    }

        //    return View(user);
        //}
        //public ActionResult DeleteUser(int id)
        //{
        //    _usersRepository.DeleteUser(id);

        //    TempData["Message"] = "User Successfully Deleted";
        //    return RedirectToAction("Users");
        //}
        //

        //public ActionResult Roles()
        //{
        //    var clients = _usersRepository.GetRoles();
        //    return View(clients);
        //}

        //public ActionResult EditRole(int? id)
        //{
        //    if (id == null)
        //    {
        //        return new HttpStatusCodeResult(HttpStatusCode.BadRequest);
        //    }

        //    Role role = _usersRepository.GetRole((int)id);
        //    if (role == null)
        //    {
        //        return HttpNotFound();
        //    }
        //    SetupRoleSelectListItems();
        //    return View(role);
        //}
        //[HttpPost]
        //public ActionResult EditRole(Role role)
        //{
        //    if (ModelState.IsValid)
        //    {
        //        _usersRepository.UpdateRole(role);

        //        TempData["Message"] = "Role Successfully Updated";
        //        return RedirectToAction("Roles");
        //    }

        //    return View(role);
        //}
        //public ActionResult DeleteRole(int id)
        //{
        //    _usersRepository.DeleteRole(id);

        //    TempData["Message"] = "Role Successfully Deleted";
        //    return RedirectToAction("Roles");
        //}

        public ActionResult SalesReport()
        {
            var fromToDate = new FromToDate();
            return View(fromToDate);
        }
        public ActionResult TrackingReport()
        {
            var fromToDate = new FromToDate();
            return View(fromToDate);
        }
        public ActionResult CollectionReport()
        {
            var fromToDate = new FromToDate();
            return View(fromToDate);
        }
        public ActionResult BaileeLetter()
        {
            //var raiLoans = _loanRepository.GetLoans().Where(x => x.LoanStatus.LoanStatusName.Equals("1 - Underwriting")
            //       || x.LoanStatus.LoanStatusName.Equals("2 - Funding")
            //       || x.LoanStatus.LoanStatusName.Equals("3 - Open")).ToList();
            //var loanList = new LoanList();
            //loanList.Loans = raiLoans;
            //var context = new ApplicationDbContext();
            var loans = new List<LoanForLetters>();
            var loan1 = new LoanForLetters();
            loan1.LoanNumber = "XXX";
            loan1.LoanID = 1;
            loans.Add(loan1);
            return View(loans);
          
        }

        public ActionResult AgingReport()
        {
            var fromToDate = new FromToDate();
            return View(fromToDate);
        }
        public ActionResult Dashboard()
        {
            return View("");
        }
        [HttpPost]
        public JsonResult RunAgingReport(DateTime? dtFromDate, DateTime? dtToDate)
        {
            try
            {

                DateTime fromDate = new DateTime();
                DateTime toDate = new DateTime();

                //if (!dtFromDate.SelectedDate.HasValue)
                //{
                //    ErrorLabel.Content = "Please enter a valid date";
                //    return;
                //}
                //else
                //    fromDate = dtFromDate.SelectedDate.Value;
                //if (!dtToDate.SelectedDate.HasValue)
                //{
                //    ErrorLabel.Content = "Please enter a valid date";
                //    return;
                //}
                //else
                //    toDate = dtToDate.SelectedDate.Value;
                //fromDate = dtFromDate.Value;
                //toDate = dtToDate.Value;
                var rptData = _loanRepository.GetLoans();


                string client = "";
                string entity = "";
                using (var workbook = new XLWorkbook())
                {
                    var ws = workbook.Worksheets.Add("Aging Report");
                    ws.Cell("A2").Value = "Master Receivable Report";
                    ws.Cell("A3").FormulaA1 = "=Today()";

                    ws.Cell("A3").Style.NumberFormat.Format = "yyyy-mm-dd";

                    ws.Range("E5:F200").Style.NumberFormat.Format = "#,##0.00";

                    ws.Range("E5:F200").Style.NumberFormat.Format = "#,##0.00";
                    ws.Range("H5:H200").Style.NumberFormat.Format = "#,##0";

                    ws.Range("I5:I200").Style.NumberFormat.Format = "#,##0.00";
                    ws.Range("J5:J200").Style.NumberFormat.Format = "0.00%";
                    ws.Range("G5:G200").Style.NumberFormat.Format = "yyyy-mm-dd";

                    ws.Cell("A2").Style.Font.FontSize = 18;
                    ws.Cell("A3").Style.Font.FontSize = 18;

                    ws.Range(4, 1, 4, 38).Style.Font.FontSize = 14;


                    ws.Range(4, 1, 4, 38).Style.Fill.PatternType = XLFillPatternValues.Solid;

                    ws.Range(4, 1, 4, 38).Style.Fill.SetBackgroundColor(XLColor.DarkGray);

                    ws.Range(4, 1, 4, 38).Style.Font.SetFontColor(XLColor.White);
                    ws.Cell("A4").Value = "Client Name";
                    ws.Cell("B4").Value = "Entity Name";
                    ws.Cell("C4").Value = "Loan Number";
                    ws.Cell("D4").Value = "Name";
                    ws.Cell("E4").Value = "Open Mortgage Balance";
                    ws.Cell("F4").Value = "Open Advance Balance";
                    ws.Cell("G4").Value = "Date Deposited in Escrow";
                    ws.Cell("H4").Value = "Days Outstanding (pending)";
                    ws.Cell("I4").Value = "Appraisal";
                    ws.Cell("J4").Value = "Appraisal %";

                    int clientStartRow = 0;
                    int clientEntityStartRow = 0;
                    string rowType = "";
                    //double totalOpenMortgageBalance = 0;
                    //double totalOpenAdvanceBalance = 0;
                    int dataStartRow = 5;
                    int dataEndRow = 0;
                    int row = 5;
                    foreach (var loan in rptData)
                    {

                        if (row == dataStartRow)
                        {
                            client = loan.Client.ClientName;
                            entity = loan.Entity.EntityName;
                            clientStartRow = row;
                            clientEntityStartRow = row;
                        }
                        if (loan.Client.ClientName != client)
                        {
                            ws.Cell(row, 1).Value = client;
                            ws.Cell(row, 4).Value = "";

                            ws.Cell(row, 5).FormulaA1 = "=SUBTOTAL(9,E" + clientStartRow + ":E" + (row - 1) + ")";
                            ws.Cell(row, 6).FormulaA1 = "=SUBTOTAL(9,F" + clientStartRow + ":F" + (row - 1) + ")";

                            ws.Cell(row, 8).Value = "";
                            ws.Cell(row, 9).Value = "";
                            ws.Cell(row, 10).Value = "";
                            ws.Cell(row, 11).Value = "";
                            ws.Range(row, 1, row, 11).Style.Font.Bold = true;
                            ws.Range(row, 1, row, 11).Style.Font.FontSize = 12;

                            ws.Range(row, 1, row, 11).Style.Fill.PatternType = XLFillPatternValues.Solid;

                            ws.Range(row, 1, row, 11).Style.Fill.SetBackgroundColor(XLColor.LightGray);
                            row++;

                            clientStartRow = row;
                            clientEntityStartRow = row;
                            client = loan.Client.ClientName;
                            entity = loan.Entity.EntityName;
                        }

                        ws.Cell(row, 1).Value = loan.Client.ClientName;
                        ws.Cell(row, 2).Value = loan.Entity.EntityName;
                        ws.Cell(row, 3).Value = loan.LoanNumber;
                        ws.Cell(row, 4).Value = loan.LoanMortgagee;
                        ws.Cell(row, 5).Value = loan.OpenMortgageBalance;
                        ws.Cell(row, 6).Value = loan.OpenAdvanceBalance;
                        ws.Cell(row, 7).Value = loan.DateDepositedInEscrow;
                        ws.Cell(row, 8).Value = loan.DaysOutstandingPending;
                        ws.Cell(row, 9).Value = loan.LoanUWAppraisalAmount;
                        ws.Cell(row, 10).Value = loan.AppraisalPcnt;


                        row++;
                    }
                    //Client Total
                    ws.Cell(row, 1).Value = client;
                    ws.Cell(row, 4).Value = "";

                    ws.Cell(row, 5).FormulaA1 = "=SUBTOTAL(9,E" + clientStartRow + ":E" + (row - 1) + ")";
                    ws.Cell(row, 6).FormulaA1 = "=SUBTOTAL(9,F" + clientStartRow + ":F" + (row - 1) + ")";

                    ws.Cell(row, 8).Value = "";
                    ws.Cell(row, 9).Value = "";
                    ws.Cell(row, 10).Value = "";
                    ws.Cell(row, 11).Value = "";
                    ws.Range(row, 1, row, 11).Style.Font.Bold = true;
                    ws.Range(row, 1, row, 11).Style.Font.FontSize = 12;

                    ws.Range(row, 1, row, 11).Style.Fill.PatternType = XLFillPatternValues.Solid;

                    ws.Range(row, 1, row, 11).Style.Fill.SetBackgroundColor(XLColor.LightGray);
                    //Grand Total
                    row++;
                    ws.Cell(row, 1).Value = "Total";
                    ws.Cell(row, 4).Value = "";
                    ws.Cell(row, 5).FormulaA1 = "=SUBTOTAL(9,E" + dataStartRow + ":E" + (row - 1) + ")";
                    ws.Cell(row, 6).FormulaA1 = "=SUBTOTAL(9,F" + dataStartRow + ":F" + (row - 1) + ")";

                    ws.Cell(row, 8).Value = "";
                    ws.Range(row, 1, row, 11).Style.Font.Bold = true;
                    ws.Range(row, 1, row, 11).Style.Font.FontSize = 14;

                    ws.Range(row, 1, row, 11).Style.Fill.PatternType = XLFillPatternValues.Solid;

                    ws.Range(row, 1, row, 11).Style.Fill.SetBackgroundColor(XLColor.DarkGray);

                    ws.Range(row, 1, row, 11).Style.Font.SetFontColor(XLColor.White);


                    ws.Columns("A:P").AdjustToContents();
                    workbook.SaveAs("C:\\Users\\pdean\\Reports\\AgingReport.xlsx");
                }
            }
            catch (Exception ex)
            {
                //ErrorLabel.Content = ex.Message;
                //ErrorLabel.Foreground = new SolidColorBrush(Colors.Red);
            }

            System.Diagnostics.Process.Start("C:\\Users\\pdean\\Reports\\AgingReport.xlsx");
            return Json("");
        }
        [HttpPost]
        public JsonResult RunTrackingReport(DateTime? dtFromDate, DateTime? dtToDate)
        {
            try
            {

                DateTime fromDate = new DateTime();
                DateTime toDate = new DateTime();

                //if (!dtFromDate.SelectedDate.HasValue)
                //{
                //    ErrorLabel.Content = "Please enter a valid date";
                //    return;
                //}
                //else
                //    fromDate = dtFromDate.SelectedDate.Value;
                //if (!dtToDate.SelectedDate.HasValue)
                //{
                //    ErrorLabel.Content = "Please enter a valid date";
                //    return;
                //}
                //else
                //    toDate = dtToDate.SelectedDate.Value;
                //fromDate = dtFromDate.Value;
                //toDate = dtToDate.Value;
                var rptData = _loanRepository.GetLoans();


                using (var workbook = new XLWorkbook())
                {
                    //    ExcelWorksheet ws = pck.Workbook.Worksheets.Add("Aging Report");
                    var ws = workbook.Worksheets.Add("Tracking Report");

                    //ws.Cells["A4"].LoadFromDataTable(report.Tables[0], true);
                    ws.Cell("A2").Value = "Tracking Report";
                    ws.Cell("A3").FormulaA1 = "=Today()";
                    ws.Cell("A3").Style.NumberFormat.Format = "yyyy-mm-dd";

                    ws.Cell("A2").Style.Font.FontSize = 18;
                    ws.Cell("A3").Style.Font.FontSize = 18;

                    ws.Range("D5:D3000").Style.NumberFormat.Format = "#,##0.00";
                    ws.Range("G5:I3000").Style.NumberFormat.Format = "#,##0.00";
                    ws.Range("P5:P3000").Style.NumberFormat.Format = "#,##0.00";
                    ws.Range("R5:R3000").Style.NumberFormat.Format = "#,##0.00";
                    ws.Range("T5:Z3000").Style.NumberFormat.Format = "#,##0.00";
                    ws.Range("AH5:AH3000").Style.NumberFormat.Format = "#,##0.00";

                    ws.Range("E5:E3000").Style.NumberFormat.Format = "0%";
                    ws.Range("AG5:AG3000").Style.NumberFormat.Format = "0%";
                    ws.Range("K5:K3000").Style.NumberFormat.Format = "0.00%";
                    ws.Range("M5:O3000").Style.NumberFormat.Format = "0.00%";

                    ws.Range("S5:S3000").Style.NumberFormat.Format = "#,##0";
                    ws.Range("AA5:SAA000").Style.NumberFormat.Format = "#,##0";

                    ws.Range(4, 1, 4, 38).Style.Font.FontSize = 14;
                    ws.Range(4, 1, 4, 38).Style.Fill.PatternType = XLFillPatternValues.Solid;

                    ws.Range(4, 1, 4, 38).Style.Fill.SetBackgroundColor(XLColor.DarkGray);
                    ws.Range(4, 1, 4, 38).Style.Font.SetFontColor(XLColor.White);
                    ws.Cell("A4").Value = "Client Name";
                    ws.Cell("B4").Value = "Loan Number";
                    ws.Cell("C4").Value = "Loan Mortgagee";
                    ws.Cell("D4").Value = "Loan Mortgage Amount";
                    ws.Cell("E4").Value = "Advance Rate";
                    ws.Cell("F4").Value = "Loan Property Address";
                    ws.Cell("G4").Value = "Advance Balance";
                    ws.Cell("H4").Value = "Wire Fee";
                    ws.Cell("I4").Value = "Advance With Wire Fee";
                    ws.Cell("J4").Value = "Date Deposited In Escrow";
                    ws.Cell("K4").Value = "Loan Interest Rate";
                    ws.Cell("L4").Value = "Underwriting Fee";
                    ws.Cell("M4").Value = "Minimum Interest";
                    ws.Cell("N4").Value = "Greater of Min and Coupon Interest";
                    ws.Cell("O4").Value = "Daily Interest Rate";
                    ws.Cell("P4").Value = "Origination Discount";
                    ws.Cell("Q4").Value = "Investor Proceeds Date";
                    ws.Cell("R4").Value = "Investor Proceeds";
                    ws.Cell("S4").Value = "Days Outstanding Closed";
                    ws.Cell("T4").Value = "Interest Fee";
                    ws.Cell("U4").Value = "Origination Fee";
                    ws.Cell("V4").Value = "Underwriting Fee";
                    ws.Cell("W4").Value = "Total Fees";
                    ws.Cell("X4").Value = "Client Proceeds";
                    ws.Cell("Y4").Value = "Open Mortgage Balance";
                    ws.Cell("Z4").Value = "OpenAdvance Balance";
                    ws.Cell("AA4").Value = "Days Outstanding Pending";
                    ws.Cell("AB4").Value = "Annualized Yield";
                    ws.Cell("AC4").Value = "Entity Name";
                    ws.Cell("AD4").Value = "Client Number";
                    ws.Cell("AE4").Value = "Investor Name";
                    ws.Cell("AF4").Value = "Loan Status";
                    ws.Cell("AG4").Value = "Appraisal Pcnt";
                    ws.Cell("AH4").Value = "Appraisal";
                    ws.Cell("AI4").Value = "Loan Funding Date";
                    ws.Cell("AJ4").Value = "State";
                    ws.Cell("AK4").Value = "Dwelling Type";
                    ws.Cell("AL4").Value = "Loan Type";
                    ws.Cell("AM4").Value = "RowType";
                    ws.Cell("AN4").Value = "SortOrder";
                    ws.Cell("AO4").Value = "SortField";
                    string client = "";
                    string entity = "";
                    int clientStartRow = 0;
                    int clientEntityStartRow = 0;
                    string rowType = "";

                    int dataStartRow = 5;
                    int dataEndRow = 0;
                    int row = 5;
                    foreach (var loan in rptData)
                    {
                        if (row == dataStartRow)
                        {
                            client = loan.Client.ClientName;
                            entity = loan.Entity.EntityName;
                            clientStartRow = row;
                            clientEntityStartRow = row;
                        }
                        if (loan.Client.ClientName != client)
                        {
                            ws.Cell(row, 1).Value = client;

                            ws.Cell(row, 4).FormulaA1 = "=SUBTOTAL(9,D" + clientStartRow + ":D" + (row - 1) + ")";
                            ws.Cell(row, 8).FormulaA1 = "=SUBTOTAL(9,H" + clientStartRow + ":H" + (row - 1) + ")";
                            ws.Cell(row, 9).FormulaA1 = "=SUBTOTAL(9,I" + clientStartRow + ":I" + (row - 1) + ")";
                            ws.Cell(row, 18).FormulaA1 = "=SUBTOTAL(9,R" + clientStartRow + ":R" + (row - 1) + ")";
                            ws.Cell(row, 20).FormulaA1 = "=SUBTOTAL(9,T" + clientStartRow + ":T" + (row - 1) + ")";
                            ws.Cell(row, 21).FormulaA1 = "=SUBTOTAL(9,U" + clientStartRow + ":U" + (row - 1) + ")";
                            ws.Cell(row, 22).FormulaA1 = "=SUBTOTAL(9,V" + clientStartRow + ":V" + (row - 1) + ")";
                            ws.Cell(row, 23).FormulaA1 = "=SUBTOTAL(9,W" + clientStartRow + ":W" + (row - 1) + ")";
                            ws.Cell(row, 24).FormulaA1 = "=SUBTOTAL(9,X" + clientStartRow + ":X" + (row - 1) + ")";
                            ws.Range(row, 1, row, 38).Style.Font.Bold = true;
                            ws.Range(row, 1, row, 38).Style.Font.FontSize = 12;

                            ws.Range(row, 1, row, 38).Style.Fill.PatternType = XLFillPatternValues.Solid;

                            ws.Range(row, 1, row, 38).Style.Fill.SetBackgroundColor(XLColor.LightGray);
                            row++;

                            clientStartRow = row;
                            clientEntityStartRow = row;
                            client = loan.Client.ClientName;
                            entity = loan.Entity.EntityName;
                        }

                        ws.Cell(row, 1).Value = loan.Client.ClientName;
                        ws.Cell(row, 2).Value = loan.LoanNumber;
                        ws.Cell(row, 3).Value = loan.LoanMortgagee;
                        ws.Cell(row, 4).Value = loan.LoanMortgageAmount.Value;
                        ws.Cell(row, 5).Value = loan.LoanAdvanceRate.Value;
                        ws.Cell(row, 6).Value = loan.LoanPropertyAddress;
                        ws.Cell(row, 7).Value = loan.LoanAdvanceAmount.Value;
                        ws.Cell(row, 8).Value = loan.WireFee;
                        ws.Cell(row, 9).Value = loan.AdvanceWithWireFee;
                        ws.Cell(row, 10).Value = loan.DateDepositedInEscrow.Value;
                        ws.Cell(row, 11).Value = loan.LoanInterestRate.Value;
                        ws.Cell(row, 12).Value = loan.UnderwritingFee.Value;
                        ws.Cell(row, 13).Value = loan.MinimumInterest.Value;
                        ws.Cell(row, 14).Value = loan.GreaterofMinandCouponInterest;
                        ws.Cell(row, 15).Value = loan.DailyInterestRate.Value;
                        ws.Cell(row, 16).Value = loan.OriginationFee.Value;
                        ws.Cell(row, 17).Value = loan.InvestorProceedsDate.Value;
                        ws.Cell(row, 18).Value = loan.InvestorProceeds.Value;

                        ws.Cell(row, 19).Value = loan.DaysOutstandingClosed;
                        ws.Cell(row, 20).Value = loan.InterestFee.Value;
                        ws.Cell(row, 21).Value = loan.OriginationFee.Value;
                        ws.Cell(row, 22).Value = loan.UnderwritingFee.Value;
                        ws.Cell(row, 23).Value = loan.TotalFees.Value;
                        ws.Cell(row, 24).Value = loan.MortgageOriginatorProceeds.Value;
                        ws.Cell(row, 25).Value = loan.OpenMortgageBalance.Value;
                        ws.Cell(row, 26).Value = loan.OpenAdvanceBalance.Value;
                        ws.Cell(row, 27).Value = loan.DaysOutstandingPending;
                        ws.Cell(row, 28).Value = loan.AnnualizedYield;
                        ws.Cell(row, 29).Value = loan.Entity.EntityName;
                        ws.Cell(row, 30).Value = "";
                        ws.Cell(row, 31).Value = loan.Investor.InvestorName;
                        ws.Cell(row, 32).Value = loan.LoanStatus.LoanStatusName;
                        ws.Cell(row, 33).Value = loan.AppraisalPcnt;
                        ws.Cell(row, 34).Value = loan.LoanUWAppraisalAmount.Value;
                        ws.Cell(row, 35).Value = loan.LoanFundingDate;
                        ws.Cell(row, 36).Value = loan.State.State1;
                        ws.Cell(row, 37).Value = loan.DwellingType.DwellingType1;
                        ws.Cell(row, 38).Value = loan.LoanType.LoanTypeName;
                        row++;
                    }

                    //Client Total
                    ws.Cell(row, 1).Value = client;

                    ws.Cell(row, 4).FormulaA1 = "=SUBTOTAL(9,D" + clientStartRow + ":D" + (row - 1) + ")";
                    ws.Cell(row, 8).FormulaA1 = "=SUBTOTAL(9,H" + clientStartRow + ":H" + (row - 1) + ")";
                    ws.Cell(row, 9).FormulaA1 = "=SUBTOTAL(9,I" + clientStartRow + ":I" + (row - 1) + ")";
                    ws.Cell(row, 18).FormulaA1 = "=SUBTOTAL(9,R" + clientStartRow + ":R" + (row - 1) + ")";
                    ws.Cell(row, 20).FormulaA1 = "=SUBTOTAL(9,T" + clientStartRow + ":T" + (row - 1) + ")";
                    ws.Cell(row, 21).FormulaA1 = "=SUBTOTAL(9,U" + clientStartRow + ":U" + (row - 1) + ")";
                    ws.Cell(row, 22).FormulaA1 = "=SUBTOTAL(9,V" + clientStartRow + ":V" + (row - 1) + ")";
                    ws.Cell(row, 23).FormulaA1 = "=SUBTOTAL(9,W" + clientStartRow + ":W" + (row - 1) + ")";
                    ws.Cell(row, 24).FormulaA1 = "=SUBTOTAL(9,X" + clientStartRow + ":X" + (row - 1) + ")";

                    ws.Range(row, 1, row, 38).Style.Font.Bold = true;
                    ws.Range(row, 1, row, 38).Style.Font.FontSize = 12;

                    ws.Range(row, 1, row, 38).Style.Fill.PatternType = XLFillPatternValues.Solid;

                    ws.Range(row, 1, row, 38).Style.Fill.SetBackgroundColor(XLColor.LightGray);
                    //Grand Total
                    row++;
                    ws.Cell(row, 1).Value = "Total";
                    ws.Cell(row, 4).FormulaA1 = "=SUBTOTAL(9,D" + dataStartRow + ":D" + (row - 1) + ")";
                    ws.Cell(row, 8).FormulaA1 = "=SUBTOTAL(9,H" + dataStartRow + ":H" + (row - 1) + ")";
                    ws.Cell(row, 9).FormulaA1 = "=SUBTOTAL(9,I" + dataStartRow + ":I" + (row - 1) + ")";
                    ws.Cell(row, 18).FormulaA1 = "=SUBTOTAL(9,R" + dataStartRow + ":R" + (row - 1) + ")";
                    ws.Cell(row, 20).FormulaA1 = "=SUBTOTAL(9,T" + dataStartRow + ":T" + (row - 1) + ")";
                    ws.Cell(row, 21).FormulaA1 = "=SUBTOTAL(9,U" + dataStartRow + ":U" + (row - 1) + ")";
                    ws.Cell(row, 22).FormulaA1 = "=SUBTOTAL(9,V" + dataStartRow + ":V" + (row - 1) + ")";
                    ws.Cell(row, 23).FormulaA1 = "=SUBTOTAL(9,W" + dataStartRow + ":W" + (row - 1) + ")";
                    ws.Cell(row, 24).FormulaA1 = "=SUBTOTAL(9,X" + dataStartRow + ":X" + (row - 1) + ")";

                    ws.Cell(row, 8).Value = "";
                    ws.Range(row, 1, row, 38).Style.Font.Bold = true;
                    ws.Range(row, 1, row, 38).Style.Font.FontSize = 14;

                    ws.Range(row, 1, row, 38).Style.Fill.PatternType = XLFillPatternValues.Solid;

                    ws.Range(row, 1, row, 38).Style.Fill.SetBackgroundColor(XLColor.DarkGray);

                    ws.Range(row, 1, row, 38).Style.Font.SetFontColor(XLColor.White);
                    ws.Columns("A:AO").AdjustToContents();
                    workbook.SaveAs("C:\\Users\\pdean\\Reports\\TrackingReport.xlsx");
                }

            }
            catch (Exception ex)
            {
                //ErrorLabel.Content = ex.Message;
                //ErrorLabel.Foreground = new SolidColorBrush(Colors.Red);
            }

            System.Diagnostics.Process.Start("C:\\Users\\pdean\\Reports\\TrackingReport.xlsx");
            return Json("");
        }
        [HttpPost]
        public JsonResult RunCollectionReport(DateTime? dtFromDate, DateTime? dtToDate)
        {
            try
            {

                DateTime fromDate = new DateTime();
                DateTime toDate = new DateTime();

                //if (!dtFromDate.SelectedDate.HasValue)
                //{
                //    ErrorLabel.Content = "Please enter a valid date";
                //    return;
                //}
                //else
                //    fromDate = dtFromDate.SelectedDate.Value;
                //if (!dtToDate.SelectedDate.HasValue)
                //{
                //    ErrorLabel.Content = "Please enter a valid date";
                //    return;
                //}
                //else
                //    toDate = dtToDate.SelectedDate.Value;
                //fromDate = dtFromDate.Value;
                //toDate = dtToDate.Value;
                var rptData = _loanRepository.GetLoans();

                using (var workbook = new XLWorkbook())
                {
                    //    ExcelWorksheet ws = pck.Workbook.Worksheets.Add("Aging Report");
                    var ws = workbook.Worksheets.Add("Collection Report");

                    //ws.Cells["A4"].LoadFromDataTable(report.Tables[0], true);
                    ws.Cell("A2").Value = "Collection Report";

                    ws.Cell("A3").FormulaA1 = "=Today()";

                    ws.Cell("A3").Style.NumberFormat.Format = "yyyy-mm-dd";
                    ws.Cell("A2").Style.Font.FontSize = 18;
                    ws.Cell("A3").Style.Font.FontSize = 18;

                    ws.Range("E5:G100").Style.NumberFormat.Format = "#,##0.00";
                    ws.Range("K5:N100").Style.NumberFormat.Format = "#,##0.00";

                    ws.Range("H5:H100").Style.NumberFormat.Format = "#,##0.00%";
                    ws.Range("I5:I100").Style.NumberFormat.Format = "#,##0";
                    ws.Range("J5:J100").Style.NumberFormat.Format = "yyyy-mm-dd";


                    ws.Range(4, 1, 4, 38).Style.Font.FontSize = 14;


                    ws.Range(4, 1, 4, 38).Style.Fill.PatternType = XLFillPatternValues.Solid;

                    ws.Range(4, 1, 4, 38).Style.Fill.SetBackgroundColor(XLColor.DarkGray);

                    ws.Range(4, 1, 4, 38).Style.Font.SetFontColor(XLColor.White);

                    ws.Cell("A4").Value = "Entity Name";
                    ws.Cell("B4").Value = "Client Name";
                    ws.Cell("C4").Value = "Loan Number";
                    ws.Cell("D4").Value = "Loan Mortgagee";
                    ws.Cell("E4").Value = "Loan Mortgage Amount";
                    ws.Cell("F4").Value = "Loan Advance Amount";
                    ws.Cell("G4").Value = "Investor Proceeds";
                    ws.Cell("H4").Value = "Yield";
                    ws.Cell("I4").Value = "Days Outstanding Closed";
                    ws.Cell("J4").Value = "Investor Proceeds Date";
                    ws.Cell("K4").Value = "Interest Fee";
                    ws.Cell("L4").Value = "Origination Fee";
                    ws.Cell("M4").Value = "Underwriting Fee";
                    ws.Cell("N4").Value = "Total Fees";
                    ws.Cell("O4").Value = "SortOrder";
                    string client = "";
                    string investor = "";
                    int clientStartRow = 0;
                    int clientInvestorStartRow = 0;
                    string rowType = "";
                    //double totalOpenMortgageBalance = 0;
                    //double totalOpenAdvanceBalance = 0;
                    int dataStartRow = 5;
                    int dataEndRow = 0;
                    int row = 5;

                    foreach (var loan in rptData)
                    {
                        if (row == dataStartRow)
                        {
                            client = loan.Client.ClientName;
                            //entity = loan.Entity.EntityName;
                            clientStartRow = row;
                            //clientEntityStartRow = row;
                        }
                        if (loan.Client.ClientName != client)
                        {
                            ws.Cell(row, 1).Value = client;

                            ws.Cell(row, 5).FormulaA1 = "=SUBTOTAL(9,E" + clientStartRow + ":E" + (row - 1) + ")";
                            ws.Cell(row, 6).FormulaA1 = "=SUBTOTAL(9,F" + clientStartRow + ":F" + (row - 1) + ")";
                            ws.Cell(row, 7).FormulaA1 = "=SUBTOTAL(9,G" + clientStartRow + ":G" + (row - 1) + ")";
                            ws.Cell(row, 11).FormulaA1 = "=SUBTOTAL(9,K" + clientStartRow + ":K" + (row - 1) + ")";
                            ws.Cell(row, 12).FormulaA1 = "=SUBTOTAL(9,L" + clientStartRow + ":L" + (row - 1) + ")";
                            ws.Cell(row, 13).FormulaA1 = "=SUBTOTAL(9,M" + clientStartRow + ":M" + (row - 1) + ")";
                            ws.Cell(row, 14).FormulaA1 = "=SUBTOTAL(9,N" + clientStartRow + ":N" + (row - 1) + ")";
                            ws.Range(row, 1, row, 15).Style.Font.Bold = true;
                            ws.Range(row, 1, row, 15).Style.Font.FontSize = 12;

                            ws.Range(row, 1, row, 15).Style.Fill.PatternType = XLFillPatternValues.Solid;

                            ws.Range(row, 1, row, 15).Style.Fill.SetBackgroundColor(XLColor.LightGray);
                            row++;

                            clientStartRow = row;
                            //clientEntityStartRow = row;
                            client = loan.Client.ClientName;
                            //entity = loan.Entity.EntityName;
                        }

                        ws.Cell(row, 1).Value = loan.Entity.EntityName;
                        ws.Cell(row, 2).Value = loan.Client.ClientName;
                        ws.Cell(row, 3).Value = loan.LoanNumber;
                        ws.Cell(row, 4).Value = loan.LoanMortgagee;
                        ws.Cell(row, 5).Value = loan.LoanMortgageAmount;
                        ws.Cell(row, 6).Value = loan.LoanAdvanceAmount;
                        ws.Cell(row, 7).Value = loan.InvestorProceeds;
                        ws.Cell(row, 8).Value = loan.AnnualizedYield;
                        ws.Cell(row, 9).Value = loan.DaysOutstandingClosed;
                        ws.Cell(row, 10).Value = loan.InvestorProceedsDate;
                        ws.Cell(row, 11).Value = loan.InterestFee;
                        ws.Cell(row, 12).Value = loan.OriginationFee;
                        ws.Cell(row, 13).Value = loan.UnderwritingFee;
                        ws.Cell(row, 14).Value = loan.TotalFees;
                        ws.Cell(row, 15).Value = "";
                        row++;

                    }

                    //Client Total
                    ws.Cell(row, 1).Value = client;

                    ws.Cell(row, 5).FormulaA1 = "=SUBTOTAL(9,E" + clientStartRow + ":E" + (row - 1) + ")";
                    ws.Cell(row, 6).FormulaA1 = "=SUBTOTAL(9,F" + clientStartRow + ":F" + (row - 1) + ")";
                    ws.Cell(row, 7).FormulaA1 = "=SUBTOTAL(9,G" + clientStartRow + ":G" + (row - 1) + ")";
                    ws.Cell(row, 11).FormulaA1 = "=SUBTOTAL(9,K" + clientStartRow + ":K" + (row - 1) + ")";
                    ws.Cell(row, 12).FormulaA1 = "=SUBTOTAL(9,L" + clientStartRow + ":L" + (row - 1) + ")";
                    ws.Cell(row, 13).FormulaA1 = "=SUBTOTAL(9,M" + clientStartRow + ":M" + (row - 1) + ")";
                    ws.Cell(row, 14).FormulaA1 = "=SUBTOTAL(9,N" + clientStartRow + ":N" + (row - 1) + ")";

                    ws.Range(row, 1, row, 15).Style.Font.Bold = true;
                    ws.Range(row, 1, row, 15).Style.Font.FontSize = 12;

                    ws.Range(row, 1, row, 15).Style.Fill.PatternType = XLFillPatternValues.Solid;

                    ws.Range(row, 1, row, 15).Style.Fill.SetBackgroundColor(XLColor.LightGray);
                    //Grand Total
                    row++;
                    ws.Cell(row, 1).Value = "Total";
                    ws.Cell(row, 5).FormulaA1 = "=SUBTOTAL(9,E" + dataStartRow + ":E" + (row - 1) + ")";
                    ws.Cell(row, 6).FormulaA1 = "=SUBTOTAL(9,F" + dataStartRow + ":F" + (row - 1) + ")";
                    ws.Cell(row, 7).FormulaA1 = "=SUBTOTAL(9,G" + dataStartRow + ":G" + (row - 1) + ")";
                    ws.Cell(row, 11).FormulaA1 = "=SUBTOTAL(9,K" + dataStartRow + ":K" + (row - 1) + ")";
                    ws.Cell(row, 12).FormulaA1 = "=SUBTOTAL(9,L" + dataStartRow + ":L" + (row - 1) + ")";
                    ws.Cell(row, 13).FormulaA1 = "=SUBTOTAL(9,M" + dataStartRow + ":M" + (row - 1) + ")";
                    ws.Cell(row, 14).FormulaA1 = "=SUBTOTAL(9,N" + dataStartRow + ":N" + (row - 1) + ")";

                    ws.Cell(row, 8).Value = "";
                    ws.Range(row, 1, row, 15).Style.Font.Bold = true;
                    ws.Range(row, 1, row, 15).Style.Font.FontSize = 14;

                    ws.Range(row, 1, row, 15).Style.Fill.PatternType = XLFillPatternValues.Solid;

                    ws.Range(row, 1, row, 15).Style.Fill.SetBackgroundColor(XLColor.DarkGray);

                    ws.Range(row, 1, row, 15).Style.Font.SetFontColor(XLColor.White);
                    ws.Columns("A:P").AdjustToContents();
                    workbook.SaveAs("C:\\Users\\pdean\\Reports\\CollectionReport.xlsx");
                }
            }
            catch (Exception ex)
            {
                //ErrorLabel.Content = ex.Message;
                //ErrorLabel.Foreground = new SolidColorBrush(Colors.Red);
            }
            System.Diagnostics.Process.Start("C:\\Users\\pdean\\Reports\\CollectionReport.xlsx");
            return Json("");
        }
        [HttpPost]
        public JsonResult RunSalesReport(DateTime? dtFromDate, DateTime? dtToDate)
        {
            try
            {

                DateTime fromDate = new DateTime();
                DateTime toDate = new DateTime();

                //if (!dtFromDate.SelectedDate.HasValue)
                //{
                //    ErrorLabel.Content = "Please enter a valid date";
                //    return;
                //}
                //else
                //    fromDate = dtFromDate.SelectedDate.Value;
                //if (!dtToDate.SelectedDate.HasValue)
                //{
                //    ErrorLabel.Content = "Please enter a valid date";
                //    return;
                //}
                //else
                //    toDate = dtToDate.SelectedDate.Value;
                //fromDate = dtFromDate.Value;
                //toDate = dtToDate.Value;
                var rptData = _loanRepository.GetLoans();

                using (var workbook = new XLWorkbook())
                {
                    //    ExcelWorksheet ws = pck.Workbook.Worksheets.Add("Aging Report");
                    var ws = workbook.Worksheets.Add("Sales Report");

                    //ws.Cells["A4"].LoadFromDataTable(report.Tables[0], true);
                    ws.Cell("A2").Value = "Sales Report";

                    ws.Cell("A3").FormulaA1 = "=Today()";

                    ws.Cell("A3").Style.NumberFormat.Format = "yyyy-mm-dd";
                    ws.Cell("A2").Style.Font.FontSize = 18;
                    ws.Cell("A3").Style.Font.FontSize = 18;
                    ws.Range("F5:G200").Style.NumberFormat.Format = "#,##0.00";

                    ws.Range("H5:H200").Style.NumberFormat.Format = "yyyy-mm-dd";

                    ws.Cell("A2").Style.Font.FontSize = 18;
                    ws.Cell("A3").Style.Font.FontSize = 18;

                    ws.Range(4, 1, 4, 11).Style.Font.FontSize = 14;


                    ws.Range(4, 1, 4, 11).Style.Fill.PatternType = XLFillPatternValues.Solid;

                    ws.Range(4, 1, 4, 11).Style.Fill.SetBackgroundColor(XLColor.DarkGray);

                    ws.Range(4, 1, 4, 11).Style.Font.SetFontColor(XLColor.White);
                    ws.Cell("A4").Value = "ClientName";
                    ws.Cell("B4").Value = "Entity";
                    ws.Cell("C4").Value = "Investor";
                    ws.Cell("D4").Value = "LoanNumber";
                    ws.Cell("E4").Value = "Customer Name";
                    ws.Cell("F4").Value = "Gross Amount";
                    ws.Cell("G4").Value = "Advance";
                    ws.Cell("H4").Value = "State";
                    ws.Cell("I4").Value = "LoanDwellingType";
                    ws.Cell("J4").Value = "DateDepositedInEscrow";
                    ws.Cell("K4").Value = "LoanType";
                    ws.Cell("L4").Value = "WeekToDate";
                    ws.Cell("M4").Value = "MonthToDate";
                    ws.Cell("N4").Value = "RowType";
                    ws.Cell("O4").Value = "SortOrder";
                    ws.Cell("P4").Value = "SortField";

                    string client = "";
                    string investor = "";
                    int clientStartRow = 0;
                    int clientInvestorStartRow = 0;
                    string rowType = "";
                    //double totalOpenMortgageBalance = 0;
                    //double totalOpenAdvanceBalance = 0;
                    int row = 5;
                    int dataStartRow = 5;
                    int dataEndRow = 0;
                    foreach (var loan in rptData)
                    {
                        if (row == dataStartRow)
                        {
                            client = loan.Client.ClientName;
                            //entity = loan.Entity.EntityName;
                            clientStartRow = row;
                            //clientEntityStartRow = row;
                        }
                        if (loan.Client.ClientName != client)
                        {
                            ws.Cell(row, 1).Value = client;

                            ws.Cell(row, 6).FormulaA1 = "=SUBTOTAL(9,F" + clientStartRow + ":F" + (row - 1) + ")";
                            ws.Cell(row, 7).FormulaA1 = "=SUBTOTAL(9,G" + clientStartRow + ":G" + (row - 1) + ")";
                            ws.Range(row, 1, row, 11).Style.Font.Bold = true;
                            ws.Range(row, 1, row, 11).Style.Font.FontSize = 12;

                            ws.Range(row, 1, row, 11).Style.Fill.PatternType = XLFillPatternValues.Solid;

                            ws.Range(row, 1, row, 11).Style.Fill.SetBackgroundColor(XLColor.LightGray);
                            row++;

                            clientStartRow = row;
                            //clientEntityStartRow = row;
                            client = loan.Client.ClientName;
                            //entity = loan.Entity.EntityName;
                        }

                        ws.Cell(row, 1).Value = loan.Client.ClientName;
                        ws.Cell(row, 2).Value = loan.Entity.EntityName;
                        ws.Cell(row, 3).Value = loan.Investor.InvestorName;
                        ws.Cell(row, 4).Value = loan.LoanNumber;
                        ws.Cell(row, 5).Value = loan.LoanMortgagee;
                        ws.Cell(row, 6).Value = loan.LoanMortgageAmount;
                        ws.Cell(row, 7).Value = loan.OpenAdvanceBalance;
                        ws.Cell(row, 8).Value = loan.State.State1;
                        ws.Cell(row, 9).Value = loan.DwellingType.DwellingType1;
                        ws.Cell(row, 10).Value = loan.DateDepositedInEscrow;
                        ws.Cell(row, 11).Value = loan.LoanType.LoanTypeName;
                        ws.Cell(row, 12).Value = "";
                        ws.Cell(row, 13).Value = "";
                        ws.Cell(row, 14).Value = "";
                        ws.Cell(row, 15).Value = "";
                        ws.Cell(row, 16).Value = "";
                        row++;
                    }

                    //Client Total
                    ws.Cell(row, 1).Value = client;

                    ws.Cell(row, 6).FormulaA1 = "=SUBTOTAL(9,F" + clientStartRow + ":F" + (row - 1) + ")";
                    ws.Cell(row, 7).FormulaA1 = "=SUBTOTAL(9,G" + clientStartRow + ":G" + (row - 1) + ")";
                    ws.Range(row, 1, row, 11).Style.Font.Bold = true;
                    ws.Range(row, 1, row, 11).Style.Font.FontSize = 12;

                    ws.Range(row, 1, row, 11).Style.Fill.PatternType = XLFillPatternValues.Solid;

                    ws.Range(row, 1, row, 11).Style.Fill.SetBackgroundColor(XLColor.LightGray);
                    //Grand Total
                    row++;
                    ws.Cell(row, 1).Value = "Total";
                    ws.Cell(row, 6).FormulaA1 = "=SUBTOTAL(9,F" + dataStartRow + ":F" + (row - 1) + ")";
                    ws.Cell(row, 7).FormulaA1 = "=SUBTOTAL(9,G" + dataStartRow + ":G" + (row - 1) + ")";

                    ws.Cell(row, 8).Value = "";
                    ws.Range(row, 1, row, 11).Style.Font.Bold = true;
                    ws.Range(row, 1, row, 11).Style.Font.FontSize = 14;

                    ws.Range(row, 1, row, 11).Style.Fill.PatternType = XLFillPatternValues.Solid;

                    ws.Range(row, 1, row, 11).Style.Fill.SetBackgroundColor(XLColor.DarkGray);

                    ws.Range(row, 1, row, 11).Style.Font.SetFontColor(XLColor.White);

                    ws.Columns("A:P").AdjustToContents();
                    workbook.SaveAs("C:\\Users\\pdean\\Reports\\SalesReport.xlsx");
                }
            }
            catch (Exception ex)
            {
                //ErrorLabel.Content = ex.Message;
                //ErrorLabel.Foreground = new SolidColorBrush(Colors.Red);
            }

            System.Diagnostics.Process.Start("C:\\Users\\pdean\\Reports\\SalesReport.xlsx");
            return Json("");
        }

        public void ValidateLoan(Loan loan)
        {
            if (ModelState.IsValidField("LoanMortgageAmount") && (loan.LoanMortgageAmount == null || loan.LoanMortgageAmount <= 0))
            {
                ModelState.AddModelError("LoanMortgageAmount", "The mortgage amount must be greater than 0");
            }
            if (ModelState.IsValidField("LoanInterestRate") && (loan.LoanInterestRate == null || loan.LoanInterestRate <= 0))
            {
                ModelState.AddModelError("LoanInterestRate", "The interest rate must be greater than 0");
            }
           
            if (ModelState.IsValidField("ClientID") && loan.ClientID == null)
            {
                ModelState.AddModelError("ClientID", "You must select a client");
            }
            if (ModelState.IsValidField("InvestorID") && loan.InvestorID == null)
            {
                ModelState.AddModelError("InvestorID", "You must select a investor");
            }

        }
        public ActionResult About()
        {
            ViewBag.Message = "Your application description page.";

            return View();
        }

        public ActionResult Contact()
        {
            ViewBag.Message = "Your contact page.";

            return View();
        }
        private void SetupClientsSelectListItems()
        {
            ViewBag.ClientsSelectListItems = _loanRepository.GetClients();

        }
        private void SetupEntitiesSelectListItems()
        {
            ViewBag.EntitiesSelectListItems = _loanRepository.GetEntities();

        }
        private void SetupStatesSelectListItems()
        {
            ViewBag.StatesSelectListItems = _loanRepository.GetStates();

        }
        private void SetupLoanTypesSelectListItems()
        {
            ViewBag.LoanTypeSelectListItems = _loanRepository.GetTypes();

        }
        private void SetupInvestorsSelectListItems()
        {
            ViewBag.InvestorsSelectListItems = _loanRepository.GetInvestors();

        }
        private void SetupUsersSelectListItems()
        {
            ViewBag.UsersSelectListItems = _loanRepository.GetUsers();

        }
        private void SetupStatusSelectListItems()
        {
            ViewBag.StatusSelectListItems = _loanRepository.GetStatus();

        }
        private void SetupRoleSelectListItems()
        {
            //ViewBag.RoleSelectListItems = _loanRepository.GetRoles();

        }
        private void SetupSelectListItems()
        {
            SetupClientsSelectListItems();
            SetupEntitiesSelectListItems();
            SetupStatesSelectListItems();
            SetupLoanTypesSelectListItems();
            SetupInvestorsSelectListItems();
            SetupUsersSelectListItems();
            SetupStatusSelectListItems();
            SetupRoleSelectListItems();
        }
        public string FormatNumberCommas(string input)
        {
            double no = double.Parse(input.ToString());
            return no.ToString("#,##0");
        }
        public string FormatNumberCommas2dec(string input)
        {
            double no = double.Parse(input.ToString());
            return no.ToString("#,##0.00;(#,##0.00)");
        }

        public string FormatPcnt(string input)
        {
            double no = double.Parse(input.ToString());
            return no.ToString("0.000%");
        }
        public string GetExcelColumnName(int columnNumber)
        {
            string columnName = String.Empty;
            while (columnNumber > 0)
            {
                int modulo = (columnNumber - 1) % 26;
                columnName = Convert.ToChar('A' + modulo) + columnName;
                columnNumber = (columnNumber - modulo) / 26;

            }
            return columnName;
        }
    }
}