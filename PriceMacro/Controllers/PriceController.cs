using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Security;
using System.Text.RegularExpressions;
using System.Web;
using System.Web.Http;
using System.Web.Mvc;
using Microsoft.SharePoint.Client;
using Microsoft.Office.Interop.Excel;
using PriceMacro.Models;
using System.Configuration;
using PriceMacro.Utilities;
using System.Net.Mail;
using System.Net;
using Microsoft.Office.SharePoint.Tools;
using Microsoft.SharePoint.News.DataModel;
using Newtonsoft.Json;

namespace PriceMacro.Controllers
{
    public class PriceController : ApiController 
    {
        [System.Web.Http.HttpPost] 
        public IHttpActionResult Index(PowerAppsModel PowerAppsModel)
        {
            PowerAppsOutputModels model = new PowerAppsOutputModels();
            string json = "";
           Logger.LogInfo($"Processing started for company: {PowerAppsModel.CompanyName}");
            String Message = "";
            string pattern = @"[\\/:*?""<>|].";
            string targetFileName = "Pricing 25Jul2024.XLSM";
            string documentLibraryName = "Documents";  // Default library name
            string folderPath = "Automation Calculated Excel";  // Folder name inside the document library
            // string targetFileName = "TestingDemoTue2025-02-04T03_58_17.9317583Z.XLSM";
            string NEWtargetFileName = PowerAppsModel.CompanyName.Trim().ToUpper()+ DateTime.Now.ToString();
            string DStargetFileName = Regex.Replace(NEWtargetFileName, pattern, "_")+  ".XLSM";
            string destinationPath = ConfigurationManager.AppSettings["LocalFolderPath"]+ DStargetFileName;
            string siteUrl = "https://rentalpha.sharepoint.com/sites/Pricingsheetapp";
            string fileUrl = "/sites/Pricingsheetapp/Shared Documents/Calculation Excel Template/" + targetFileName;  // Replace with your actual file path
            string EditfileUrl = "/sites/Pricingsheetapp/Shared Documents/Automation Calculated Excel/" + DStargetFileName;
            string username = "Partner.PowerFlow@rentalpha.com"; // Replace with your username
            Logger.LogInfo("Initializing SharePoint connection"); // string password = "PFrapl$2024X";  // Replace with your password
            string myString = "PFrapl$2024X";

            SecureString secureString = new SecureString();
            foreach (char c in myString)
            {
                secureString.AppendChar(c);
            }

            secureString.MakeReadOnly();

            var context = new ClientContext(siteUrl)
            {
                Credentials = new SharePointOnlineCredentials(username, secureString)
            };
         
            try
            {
                    Logger.LogInfo($"Attempting to access file: {fileUrl}");
                    // Access the file in SharePoint
                    Microsoft.SharePoint.Client.File files = context.Web.GetFileByServerRelativeUrl(fileUrl);
                    context.Load(files);
                    context.ExecuteQuery();

                    ClientResult<Stream> data = files.OpenBinaryStream();
                    context.ExecuteQuery();

                    if (System.IO.File.Exists(destinationPath))
                    {
                        // Delete the file from local
                        Logger.LogInfo($"Deleting existing file at: {destinationPath}");
                        System.IO.File.Delete(destinationPath);
                        // Console.WriteLine("File deleted successfully!");
                    }

                    // Write the stream to a local file
                    using (FileStream fileStream = new FileStream(destinationPath, FileMode.Create))
                    {
                        Logger.LogInfo("Writing SharePoint file to local storage");
                        data.Value.CopyTo(fileStream);
                    }
                //}

                Logger.LogInfo("Opening Excel application");

                Microsoft.Office.Interop.Excel.Application excelApp = new Microsoft.Office.Interop.Excel.Application();

                Microsoft.Office.Interop.Excel.Workbook workbook = excelApp.Workbooks.Open(destinationPath, ReadOnly: false);
                // Update SingleRow sheet
                Logger.LogInfo("Updating SingleRow sheet with PowerApps data");
                Microsoft.Office.Interop.Excel.Worksheet singleRowSheet = (Microsoft.Office.Interop.Excel.Worksheet)workbook.Sheets["SingleRow"];
                // Start API Testing
                decimal gstPercentage = Convert.ToDecimal(PowerAppsModel.GSTRate);
                decimal gstDecimal = gstPercentage / 100;
                decimal EligibleGST = Convert.ToDecimal(PowerAppsModel.EligibleGST);
                decimal EligibleGSTdecimal = EligibleGST / 100;
                int FirmTerm=Convert.ToInt32(PowerAppsModel.FirmTerm);
                decimal DealTargetPNI = Convert.ToDecimal(PowerAppsModel.DealTargetPNI);
                decimal DealTargetPNIdecimal = DealTargetPNI / 100;
                Logger.LogInfo($"Setting company name: {PowerAppsModel.CompanyName}");

                ((Microsoft.Office.Interop.Excel.Range)singleRowSheet.Cells[2, 1]).Value = PowerAppsModel.CompanyName.ToString(); // String
                ((Microsoft.Office.Interop.Excel.Range)singleRowSheet.Cells[2, 2]).Value = Convert.ToInt32(PowerAppsModel.TermInMonth); // Numeric
                ((Microsoft.Office.Interop.Excel.Range)singleRowSheet.Cells[2, 3]).Value = FirmTerm; // String
                ((Microsoft.Office.Interop.Excel.Range)singleRowSheet.Cells[2, 4]).Value = PowerAppsModel.AssetCategory; // Numeric
                ((Microsoft.Office.Interop.Excel.Range)singleRowSheet.Cells[2, 5]).Value = Convert.ToInt32(PowerAppsModel.Deposit); // String (Column AA)
                ((Microsoft.Office.Interop.Excel.Range)singleRowSheet.Cells[2, 6]).Value = Convert.ToInt32(PowerAppsModel.GVR);// String
                ((Microsoft.Office.Interop.Excel.Range)singleRowSheet.Cells[2, 7]).Value = Convert.ToString(PowerAppsModel.RentalRebate); // Numeric
                ((Microsoft.Office.Interop.Excel.Range)singleRowSheet.Cells[2, 8]).Value = Convert.ToString(PowerAppsModel.Rentalcompany); // String (Column AA)
                ((Microsoft.Office.Interop.Excel.Range)singleRowSheet.Cells[2, 9]).Value = PowerAppsModel.Rating.ToString(); // String
                ((Microsoft.Office.Interop.Excel.Range)singleRowSheet.Cells[2, 11]).Value = Convert.ToDouble(PowerAppsModel.Volume); // Numeric
                ((Microsoft.Office.Interop.Excel.Range)singleRowSheet.Cells[2, 12]).Value = gstDecimal; // String (Column AA)
                ((Microsoft.Office.Interop.Excel.Range)singleRowSheet.Cells[2, 15]).Value = EligibleGSTdecimal; // String (Column AA)
                ((Microsoft.Office.Interop.Excel.Range)singleRowSheet.Cells[2, 17]).Value = Convert.ToInt32(PowerAppsModel.stringerimDays);  // String
                ((Microsoft.Office.Interop.Excel.Range)singleRowSheet.Cells[2, 18]).Value = Convert.ToDateTime(PowerAppsModel.FirmTtermrentalDate); // Numeric
                ((Microsoft.Office.Interop.Excel.Range)singleRowSheet.Cells[2, 19]).Value = PowerAppsModel.RentalFrequency.ToString(); // String (Column AA)
                ((Microsoft.Office.Interop.Excel.Range)singleRowSheet.Cells[2, 20]).Value = PowerAppsModel.RentalPaymentType.ToString();
                ((Microsoft.Office.Interop.Excel.Range)singleRowSheet.Cells[2, 26]).Value = DealTargetPNIdecimal; // String
                ((Microsoft.Office.Interop.Excel.Range)singleRowSheet.Cells[2, 27]).Value = PowerAppsModel.DealTargetPNIMacro.ToString(); // String (Column AA)
                ((Microsoft.Office.Interop.Excel.Range)singleRowSheet.Cells[2, 38]).Value = Convert.ToInt32(PowerAppsModel.FunderPVCap);// String (Column AA)
                ((Microsoft.Office.Interop.Excel.Range)singleRowSheet.Cells[2, 39]).Value = Convert.ToDouble(PowerAppsModel.FunderDiscountingRate); // String
                ((Microsoft.Office.Interop.Excel.Range)singleRowSheet.Cells[2, 40]).Value = PowerAppsModel.FunderDiscountingRatentalType.ToString(); ; // Numeric
                ((Microsoft.Office.Interop.Excel.Range)singleRowSheet.Cells[2, 43]).Value = Convert.ToDateTime(PowerAppsModel.DateOfDiscounting);  // String (Column AA)

                // End API Testing
                Logger.LogInfo("Running macro based on cell value");
                Microsoft.Office.Interop.Excel.Range cell = singleRowSheet.Range["AA2"];
                var cellValue = cell.Value;
                // Run the macro
                if (cellValue.ToString() == "CFPL")
                {
                    Logger.LogInfo("Executing CFPL macro");
                    excelApp.Run("Sheet20.MYPNIGoalSeekCFPL"); // Use the CFPL macro name}
                }
                else if(cellValue.ToString() == "RAPL")
                {
                    Logger.LogInfo("Executing RAPL macro");
                    excelApp.Run("Sheet20.MYPNIGoalSeek"); // Use the RAPL macro name}
                }
               
                workbook.Save();
                workbook.Close();
                Logger.LogInfo("Excel processing completed successfully");
                Console.WriteLine("Macro run successfully!");

                Logger.LogInfo("Preparing to upload processed file to SharePoint");
                Web web = context.Web;
                List documentLibrary = web.Lists.GetByTitle(documentLibraryName);
                Microsoft.SharePoint.Client.Folder folder = documentLibrary.RootFolder.Folders.GetByUrl(folderPath);
                context.Load(folder);
                context.ExecuteQuery();
                // Read the file content
                byte[] fileContent = System.IO.File.ReadAllBytes(destinationPath);
                string FileName = PowerAppsModel.CompanyName.Trim().ToUpper()+DateTime.Now.ToString() + ".XLSM";
               
                // Replace invalid characters with an underscore (_)
                string sanitizedFileName = Regex.Replace(FileName, pattern, "_");
                // Create a FileCreationInformation object for the file upload
                FileCreationInformation newFile = new FileCreationInformation
                {
                    Content = fileContent,
                    Url = sanitizedFileName
                };
                // Upload the file to the specific folder
                Microsoft.SharePoint.Client.File uploadFile = folder.Files.Add(newFile);
                // Execute the upload
                context.ExecuteQuery();
                Logger.LogInfo($"File uploaded successfully to SharePoint: {sanitizedFileName}");
                Console.WriteLine("File uploaded SharePoint successfully!");

                //Send Mail
                SendMail(destinationPath, PowerAppsModel);

                Logger.LogInfo("Send Mail Successfully");

                Logger.LogInfo("Open Output workSheet SingleRow sheet with PowerApps data");
                Microsoft.Office.Interop.Excel.Application excelAppp = new Microsoft.Office.Interop.Excel.Application();
                Microsoft.Office.Interop.Excel.Workbook workbookk = excelAppp.Workbooks.Open(destinationPath, ReadOnly: false);
                Microsoft.Office.Interop.Excel.Worksheet singleRowSheett = (Microsoft.Office.Interop.Excel.Worksheet)workbookk.Sheets["SingleRow"];
                Logger.LogInfo("Get Output SingleRow sheet with PowerApps data");
                model.TermInMonths = Convert.ToString(((Range)singleRowSheett.Cells[2, 2]).Value);  // B2 cell
                model.FirstTermPTPQ = ((Range)singleRowSheett.Cells[2, 3]).Value?.ToString();    // C2 cell
                model.AssetCategory = ((Range)singleRowSheett.Cells[2, 4]).Value?.ToString();     // D2 cell
                model.DepositPercentage = Convert.ToString(((Range)singleRowSheett.Cells[2, 5]).Value); // E2 cell
                model.GRVAsPercentageOfVolume = Convert.ToString(((Range)singleRowSheett.Cells[2, 6]).Value); // F2 cell
                model.RentalRebate = Convert.ToString(((Range)singleRowSheett.Cells[2, 7]).Value); // G2 cell
                model.Rating = ((Range)singleRowSheett.Cells[2, 9]).Value?.ToString();            // H2 cell
                model.Volume = Convert.ToString(((Range)singleRowSheett.Cells[2, 11]).Value);
                decimal gstPercentagee = Convert.ToDecimal(((Range)singleRowSheett.Cells[2, 12]).Value);
                decimal gstDecimall = gstPercentagee * 100;// I2 cell
                model.GSTRate = Convert.ToString(gstDecimall);    // J2 cell
                model.GSTValue = Convert.ToString(((Range)singleRowSheett.Cells[2, 13]).Value);   // K2 cell
                model.TotalColumn = Convert.ToString(((Range)singleRowSheett.Cells[2, 14]).Value); // L2 cell
                model.DebtRateQuarterly =  Convert.ToString(((Range)singleRowSheett.Cells[2, 10]).Value); // M2 cell
                model.EligibleGST = Convert.ToString(((Range)singleRowSheett.Cells[2, 15]).Value);  // N2 cell
                model.VolumeConsideredForRentalWorking = Convert.ToString(((Range)singleRowSheett.Cells[2, 16]).Value); // O2 cell
                model.InterimDays = Convert.ToString(((Range)singleRowSheett.Cells[2, 17]).Value);  // P2 cell
                model.FirmTermRentalDate = ((Range)singleRowSheett.Cells[2, 18]).Value?.ToString(); // Q2 cell
                model.RentalFrequency = ((Range)singleRowSheett.Cells[2, 19]).Value?.ToString();  // R2 cell

                // Deal Details
                model.DealTargetPNI = Convert.ToString(((Range)singleRowSheett.Cells[2, 27]).Value);  // S2 cell
                model.Slab1 = ((Range)singleRowSheett.Cells[2, 28]).Value?.ToString();  // T2 cell

                // Funder Details
                model.FunderPVCap = Convert.ToString(((Range)singleRowSheett.Cells[2, 38]).Value);  // U2 cell
                model.FunderDiscountingRate = Convert.ToString(((Range)singleRowSheett.Cells[2, 39]).Value);  // V2 cell
                model.FunderDiscountingRateType = ((Range)singleRowSheett.Cells[2, 40]).Value?.ToString();  // W2 cell
                model.QuarterlyRate = Convert.ToString(((Range)singleRowSheett.Cells[2, 41]).Value);  // X2 cell
                model.AnnualizedRate =Convert.ToString(((Range)singleRowSheett.Cells[2, 42]).Value);  // Y2 cell
                model.DateOfDiscounting = ((Range)singleRowSheett.Cells[2, 43]).Value?.ToString(); // Z2 cell

                // Client Details
                model.ClientName = ((Range)singleRowSheett.Cells[2, 47]).Value?.ToString();  // AA2 cell
                model.Tenure =  Convert.ToString(((Range)singleRowSheett.Cells[2, 48]).Value);  // AB2 cell
                model.PTPMLabel = ((Range)singleRowSheett.Cells[2, 49]).Value?.ToString();  // AC2 cell
                model.PTPMValue1 = Convert.ToString(((Range)singleRowSheett.Cells[2, 50]).Value);  // AD2 cell
                model.PTPMValue2 = Convert.ToString(((Range)singleRowSheett.Cells[2, 51]).Value);  // AE2 cell
                model.DepositValue1 = Convert.ToString(((Range)singleRowSheett.Cells[2, 52]).Value);  // AF2 cell
                model.DepositValue2 = Convert.ToString(((Range)singleRowSheett.Cells[2, 53]).Value);  // AG2 cell

                // Financial Calculations
                model.GVRValue1 = Convert.ToString(((Range)singleRowSheett.Cells[2, 56]).Value);  // AH2 cell
                model.GVRValue2 = Convert.ToString(((Range)singleRowSheett.Cells[2, 57]).Value);  // AI2 cell
                model.GSTValue1 = Convert.ToString(((Range)singleRowSheett.Cells[2, 58]).Value);  // AJ2 cell
                model.GSTValue2 = Convert.ToString(((Range)singleRowSheett.Cells[2, 59]).Value);  // AK2 cell
                model.XIRR = Convert.ToString(((Range)singleRowSheett.Cells[2, 76]).Value);  // AM2 cell
                model.PNI = Convert.ToString(((Range)singleRowSheett.Cells[2, 77]).Value);  // AN2 cell
                model.Rating6 = ((Range)singleRowSheett.Cells[2, 78]).Value?.ToString();  // AO2 cell
                var aaa = ((Range)singleRowSheett.Cells[2, 73]).Value;
                // Debt Details
                model.DebtRate = ((Range)singleRowSheett.Cells[2, 79]).Value?.ToString();  // AP2 cell
                model.DebtRateLabel = ((Range)singleRowSheett.Cells[2, 80]).Value?.ToString();  // AQ2 cell
                model.DebtRateValue =((Range)singleRowSheett.Cells[2, 81]).Value?.ToString();  // AR2 cell

                // Additional Financial Details
                model.PNI7 = Convert.ToString(((Range)singleRowSheett.Cells[2, 82]).Value);  // AS2 cell
                model.FeeOrInvestment = Convert.ToString(((Range)singleRowSheett.Cells[2, 83]).Value);  // AT2 cell
                model.FunderPVCapPercentage = Convert.ToString(((Range)singleRowSheett.Cells[2, 84]).Value);  // AU2 cell
                model.ActualFunderPV = Convert.ToString(((Range)singleRowSheett.Cells[2, 85]).Value);  // AV2 cell
                model.RNSCount = Convert.ToString(((Range)singleRowSheett.Cells[2, 86]).Value);  // AW2 cell
                model.TotalRental = Convert.ToString(((Range)singleRowSheett.Cells[2, 87]).Value);  // AX2 cell
                model.PNIAdjustedVolume = Convert.ToString(((Range)singleRowSheett.Cells[2, 88]).Value);  // AY2 cell
                model.NumberOfPayments = Convert.ToString(((Range)singleRowSheett.Cells[2, 89]).Value);  // AZ2 cell
                model.XIRRError = Convert.ToString(((Range)singleRowSheett.Cells[2, 90]).Value);  // BA2 cell
                model.TotalRentalError = Convert.ToString(((Range)singleRowSheett.Cells[2, 91]).Value);  // BB2 cell

                // Identifiers
                model.ID = PowerAppsModel.QuoteID;
                model.MacrosList = PowerAppsModel.DealTargetPNIMacro.ToString();
                model.Created = DateTime.Now.ToShortDateString();
                model.RentalPaymentType = ((Range)singleRowSheett.Cells[2, 20]).Value?.ToString();
                model.Modified = DateTime.Now.ToShortDateString();
                // json = JsonConvert.SerializeObject(model, Formatting.None);
                workbookk.Close();

            

                if (System.IO.File.Exists(destinationPath))
                {
                    Logger.LogInfo("File deleted successfully From"+ destinationPath);
                    // Delete the file from local
                    System.IO.File.Delete(destinationPath);
                 
                }
                Logger.LogInfo($"Processing completed successfully for company: {PowerAppsModel.CompanyName}");

            }
            catch (Exception ex)
            {
                Console.WriteLine("Error: " + ex.Message);
                Message = "Error: " + ex.Message;
                return Ok(Message);
            }
           return Ok(model);
        }

        private void SendMail(string destinationPath, PowerAppsModel PowerAppsModel)
        {
            // Create a new SmtpClient object
            SmtpClient smtpClient = new SmtpClient("smtp.ethereal.email", 587)
            {
                // SMTP server settings
                Credentials = new NetworkCredential("noah.goodwin@ethereal.email", "tE86s3qyH1GeqT4Jr8"),
                EnableSsl = true
            };
            System.Net.ServicePointManager.SecurityProtocol = System.Net.SecurityProtocolType.Tls12; 
            // Create the email message
            MailMessage mailMessage = new MailMessage
            {
                From = new MailAddress("noah.goodwin@ethereal.email"),
                Subject = "Pricing for  CHECK ALL DATA  ID" + PowerAppsModel.QuoteID,
                Body = "Hi Partner_RA PowerFlow,\r\nPlease find attach Pricing Sheet for CHECK ALL DATA\r\nQuote ID :" + PowerAppsModel.QuoteID + "\r\nAssets Type :" + PowerAppsModel.AssetCategory + "\r\nTenor (in Months) : " + PowerAppsModel.TermInMonth + "\r\n\n\n\nRegards\r\nPricing APP\r\n(This is automated email do not reply)\r\nFor Any Queries get in touch with vilesh.modi@capsvefinance.com\r\n "
            };

            // Add recipient email address
            mailMessage.To.Add("noah.goodwin@ethereal.email");

            System.Net.Mail.Attachment attachment = new System.Net.Mail.Attachment(destinationPath);
            mailMessage.Attachments.Add(attachment);


            // Send the email
            smtpClient.Send(mailMessage);
        }
    }
}
