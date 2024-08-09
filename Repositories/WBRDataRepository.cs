using Microsoft.EntityFrameworkCore;
using Microsoft.EntityFrameworkCore.Metadata.Internal;
using Syncfusion.Pdf;
using Syncfusion.XlsIO;
using Syncfusion.XlsIORenderer;
using Syncfusion.Pdf.Parsing;
using System.Data;
using Microsoft.Data.SqlClient;
using System.IO;
using System.Xml.Linq;
using wbrapi7_appservices.Data;
using wbrapi7_appservices.Entities;
using wbrapi7_appservices.Models;
using static Microsoft.EntityFrameworkCore.DbLoggerCategory.Database;
using static System.Net.Mime.MediaTypeNames;
using System.Runtime.Intrinsics.X86;
using wbrapi7_appservices.Services;
using Microsoft.SharePoint.Client;
using Microsoft.Online.SharePoint.TenantAdministration;
using Microsoft.AspNetCore.Mvc;
using System.Net;
using Microsoft.SharePoint.Client.Search.Query;

namespace wbrapi7_appservices.Repositories
{
    public class WBRDataRepository : IWBRDataRepository
    {
        WBRDataContext _context;

        public IConfiguration Configuration { get; }

        public WBRDataRepository(WBRDataContext context)
        {
            _context = context;
        }

        //public System.Data.SqlClient.SqlConnection SQLDatabaseConnection()
        //{
        //    try
        //    {

        //        var sSQLConn = new ConfigurationBuilder().AddJsonFile("appsettings.json").Build().GetSection("ConnectionStrings")["DefaultConnection"];

        //        System.Data.SqlClient.SqlConnection dbCon = new System.Data.SqlClient.SqlConnection(sSQLConn);

        //        return dbCon;

        //    }
        //    catch
        //    {
        //        throw;
        //    }
        //    finally
        //    {

        //    }
        //}

        public int apStatementImport(string sTicketNo)
        {

            return 0;
        }

        public IEnumerable<vciSafIncHeadStatus> GetvciSafIncHeadStatus()
        {
            return _context.vciSafIncHeadStatus.OrderBy(a => a.StatusKey);
        }

        public IEnumerable<vapJIBSharepoint> vapJIBSharepoint()
        {
            return _context.vapJIBSharepoint;
        }

        public IEnumerable<vapJIBHeader> vapJIBHeaderbyKey(int intjibHeaderKey)
        {
            return _context.vapJIBHeader.Where(a => a.JIBHeaderKey == intjibHeaderKey);

        }

        public IEnumerable<vapJIBPDFs> vapJIBPDFsbyKey(int intjibHeaderKey)
        {

            var files = _context.vapJIBPDFs.Where(a => a.PrimaryKey == intjibHeaderKey).ToList();


            // return _context.vapJIBPDFs.Where(a => a.PrimaryKey == intjibHeaderKey);
            return files;
        }


        public int GetMaxIntegerFromColumnB(IWorksheet wsSheet)
        {
            // Assuming the numbers in column B start at row 1 and there's no header you want to skip
            int maxValueInColumnB = int.MinValue;

            // Accessing the cells directly by their column and row indices
            foreach (IRange row in wsSheet.Columns[0].Cells) // Index 1 refers to column B since the collection is 0-indexed
            {
                if (row.DisplayText != null && int.TryParse(row.DisplayText.ToString(), out int cellValue))
                {
                    maxValueInColumnB = Math.Max(maxValueInColumnB, cellValue);
                }
            }

            // If maxValueInColumnB is still int.MinValue, then no numeric cells were found
            if (maxValueInColumnB == int.MinValue)
            {
                maxValueInColumnB = 0; // Or any other default value you consider appropriate
            }

            return maxValueInColumnB;
        }



        public int fnExportSheet(IWorksheet wsSheet)
        {
            var value = wsSheet.Range["B6"].DisplayText;
            var value2 = wsSheet.Range["F6"].Value2;
            int maxIntValue = GetMaxIntegerFromColumnB(wsSheet);

            //var value3 = wsSheet.Range["M6"].;
            String fullPath = Path.Combine(Directory.GetCurrentDirectory(), "Resources", "Submit_Supplier_Invoice.xml");


            XDocument xmlDocument = XDocument.Load(fullPath);
            for (int i = 1; i <= maxIntValue; i++)
            {
                String filename = "c:\\temp\\jib\\Submit_Supplier_Invoice_" + i.ToString() + ".xml";
                xmlDocument.Save(filename);
            }



            return 1;
        }




        //public async Task<int> JItest()
        //      {
        public string JIBReadFolder()
        {
            int jibHeaderKey = 0;
            string strReturnMsg = "ok";


            using (ExcelEngine excelEngine = new ExcelEngine())
            {
                IApplication application = excelEngine.Excel;
                application.DefaultVersion = ExcelVersion.Xlsx;
                application.RangeIndexerMode = ExcelRangeIndexerMode.Relative;


                //IWorkbook workbookN = application.Workbooks.Create(1);
                //IWorksheet sheetN = workbookN.Worksheets[0];

                //string strfilename = "c:\\temp\\jib\\EVX JIB Workbook 01.31.24b.xlsx";
                //string strfilename = "c:\\temp\\jib\\testSDB JIB Workbook - 02.29.2024.xlsx";

                //SharepointService spSVR = new SharepointService("WBRSAAPTicket@h2obridge.com", "7JwYq*V%w5g9m#");

                //ClientContext clientContext = spSVR.getSharepointClientContext("https://h20bridge.sharepoint.com/sites/IT");
                //Microsoft.SharePoint.Client.Web WebClient = null;
                //WebClient = clientContext.Web;
                //Folder srcFolder = WebClient.GetFolderByServerRelativeUrl("Development/JIB");
                //clientContext.Load(srcFolder, f2 => f2.Files);
                //clientContext.ExecuteQuery();
                //if (spSVR.UploadFiles("https://h20bridge.sharepoint.com/sites/AccountingGroup",
                //                 "Shared Documents/04. Accounts Payable/AP ticket attachments/AP000002",
                //                 "", "bla.pdf", btByt, false))

                //FileCollection fcol = spSVR.getSharepointFile("https://h20bridge.sharepoint.com/sites/IT", "Shared Documents/Development/JIB", "");
                //FileCollection fcol = spSVR.getSharepointFile("https://h20bridge.sharepoint.com/sites/AccountingGroup", "Shared Documents/04. Accounts Payable/AP ticket attachments/AP000002", "");
                //Stream fileStream = spSVR.getSharepointFileStream("https://h20bridge.sharepoint.com/sites/AccountingGroup", "Shared Documents/04. Accounts Payable/AP ticket attachments/AP000002", "");




                //foreach (var file in fcol)
                //{
                //    Console.WriteLine(file.Name);
                //    // Your code here to work with each 'file'
                //}

                string strSharepointLoc = "";

                //FileStream inputStream = null;
                Stream? fileStream = null;
                Microsoft.SharePoint.Client.File? file = null;
                var vLoginSharepoint = vapJIBSharepoint().FirstOrDefault();
                SharepointService? spSVR = null;

                try
                {

                    if (vLoginSharepoint == null)
                    {
                        throw new InvalidOperationException("vapJIBSharepoint login info failed to load");
                    }

                    //SharepointService spSVR = new SharepointService("WBRSAAPTicket@h2obridge.com", "7JwYq*V%w5g9m#");
                    //(fileStream, string strfilename) = spSVR.getSharepointFileStream("https://h20bridge.sharepoint.com/sites/AccountingGroup", "Shared Documents/04. Accounts Payable/AP ticket attachments/AP000002", "");
                    spSVR = new SharepointService(vLoginSharepoint.UserID, vLoginSharepoint.Password);
                    (fileStream, file) = spSVR.getSharepointFileStream(vLoginSharepoint.ServerSiteUrl, vLoginSharepoint.LibaryURL, "");



                    if (fileStream == null)
                    {
                        throw new InvalidOperationException("The file not found or failed to load. Please check the folder:" + vLoginSharepoint.ServerSiteUrl + "/" + vLoginSharepoint.LibaryURL);
                    }
                }
                catch (Exception ex)
                {
                    strReturnMsg = $"Error: {ex.Message}";
                    fileStream?.Dispose();
                    return strReturnMsg;



                }


                //if (strReturnMsg != "ok")
                //{
                //    if (fileStream != null)
                //    {
                //        fileStream.Dispose();
                //    }
                //   //strReturnMsg = "ok";
                //    return strReturnMsg;
                //}


                try
                {

                    jibHeaderKey = spapInsertJIBHeader(file.Name, "sharepoint loc");

                    if (jibHeaderKey <= 0)
                    {
                        throw new InvalidOperationException("spapInsertJIBHeader failed");
                    }



                    //var jibHeader = vapJIBHeader().FirstOrDefault(jib => jib.JIBHeaderKey == jibHeaderKey);
                    var jibHeader = vapJIBHeaderbyKey(jibHeaderKey).SingleOrDefault();

                    strSharepointLoc = vLoginSharepoint.ServerSiteUrl + "/" + vLoginSharepoint.LibaryURL + "/Archived/" + jibHeader.JIBNO;

                    if (!MoveFile(spSVR, file, vLoginSharepoint.ServerSiteUrl, vLoginSharepoint.LibaryURL + "/Archived", jibHeader.JIBNO))
                    { };


                    //  byte[] btByt = null;


                    //  fileStream.Position = 0;
                    //  using (MemoryStream ms = new MemoryStream())
                    //  {
                    //      fileStream.CopyTo(ms); 
                    //      btByt = ms.ToArray(); 
                    //  }



                    //  spSVR.UploadFiles(vLoginSharepoint.ServerSiteUrl, vLoginSharepoint.LibaryURL + "/",
                    //jibHeader.JIBNO, strfilename, btByt, false);




                    fileStream.Position = 0;
                    IWorkbook workbook = application.Workbooks.Open(fileStream);




                    //worked fronm local folder
                    //inputStream = new FileStream(strfilename, FileMode.Open, FileAccess.Read);
                    //IWorkbook workbook = application.Workbooks.Open(inputStream);
                    //end worked fronm local folder



                    //IWorksheet worksheet = workbook.Worksheets[0];
                    //IWorksheet worksheet = workbook.Worksheets["LE140 EIB Upload"];

                    List<string> eibuploadSheetNames = new List<string>();
                    foreach (IWorksheet worksheet in workbook.Worksheets)
                    {
                        if (worksheet.Name.ToLower().Contains("eib upload"))
                        {
                            eibuploadSheetNames.Add(worksheet.Name);
                        }
                    }
                    if (eibuploadSheetNames.Count == 0)
                    {
                        throw new InvalidOperationException("Eib Upload tab not found in " + file.Name);

                    }

                    foreach (string eibuploadSheetName in eibuploadSheetNames)
                    {
                        IWorksheet eibuploadSheet = workbook.Worksheets[eibuploadSheetName];



                        //    for (int i = 1; i <= eibuploadSheet.Columns.Length; i++)
                        //{
                        //        eibuploadSheet.ShowColumn(i, true);
                        //}

                        //int lastUsedRow = worksheet.UsedRange.LastRow;

                        int intlastDataRow = 0;
                        for (int i = eibuploadSheet.Rows.Length; i >= 1; i--)
                        {
                            if (eibuploadSheet.Range["B" + i].Value != null && eibuploadSheet.Range["B" + i].Value.ToString() != String.Empty)
                            {
                                intlastDataRow = i;
                                break;
                            }
                        }
                        //int intbatchSize = 20; 
                        int intstartRow = 6;

                        int intbatchSize = 100;
                        //int intstartRow = 76;
                        DataTable batchTable = null;

                        for (int intstart = intstartRow; intstart <= intlastDataRow; intstart += intbatchSize)
                        {

                            int intend = intstart + intbatchSize - 1;

                            if (intend > intlastDataRow) intend = intlastDataRow;



                            //IRange fullRange = eibuploadSheet.Range["A6:GM" + intlastDataRow.ToString()];
                            IRange batchRange = eibuploadSheet.Range["A" + intstart.ToString() + ":GM" + intend.ToString()];

                            batchTable = eibuploadSheet.ExportDataTable(batchRange, ExcelExportDataTableOptions.ComputedFormulaValues | ExcelExportDataTableOptions.ExportHiddenColumns);

                            for (int columnIndex = 1; columnIndex <= batchTable.Columns.Count; columnIndex++)
                            {

                                string excelColumnName = GetExcelColumnName(columnIndex);
                                batchTable.Columns[columnIndex - 1].ColumnName = excelColumnName;
                            }


                            bool success = InsertData(batchTable, jibHeaderKey, intstart, eibuploadSheetName);


                            batchTable.Clear();


                            if (intend == intlastDataRow) break;
                        }



                        //bool successpdf = CreatePDFs(batchTable);
                        //foreach (DataRow row in batchTable.Rows)
                        //{
                        //    if (row["A"] != null && row["A"].ToString().Length > 2)
                        //    {

                        //        bool successpdf = JIBExceltoPDF(strfilename, row["A"].ToString());

                        //        if (!successpdf)
                        //        {
                        //            bool recy = spapUpdateJIBHeader(
                        //   jibHeaderKey,
                        //   null,            // strStatus
                        //   null,            // strFilename
                        //   null,            // strSharepointLoc
                        //   row["A"].ToString() + " PDF NOT SAVED ",            // strComments
                        //   null,            // strResult
                        //   null,            //@FinishedDatetime
                        //   null            // dtSendDateTime 
                        //   );



                        //        }
                        //    }
                        //}





                        // DataTable tbTable = eibuploadSheet.ExportDataTable(fullRange, ExcelExportDataTableOptions.ComputedFormulaValues);


                    }


                    //if (!await bExportPDFs(jibHeaderKey, strfilename))
                    //{
                    //};



                    //if (!bExportPDFs(jibHeaderKey, strfilename))
                    if (!bExportPDFs(jibHeaderKey, fileStream))
                    {
                    };

                    string strPDFS = JIBPreparePDF(jibHeaderKey);



                    bool rec = spapUpdateJIBHeader(
                                    jibHeaderKey,
                                    "EIB Ready",            // strStatus
                                    null,            // strFilename
                                    strSharepointLoc,            // strSharepointLoc
                                    null,            // strComments
                                    null,            // strResult
                                    DateTime.Now,           //FinishedDatetime
                                    null            // dtSendDateTime 
                                    );


                }
                catch (Exception ex)
                {
                    strReturnMsg = $"Error: {ex.Message}";

                    bool rec = spapUpdateJIBHeader(
                                    jibHeaderKey,
                                    "Failed",            // strStatus
                                    null,            // strFilename
                                    strSharepointLoc,            // strSharepointLoc
                                    ex.Message,            // strComments
                                    null,            // strResult
                                    DateTime.Now,           //FinishedDatetime
                                    null            // dtSendDateTime 
                                    );

                }
                finally
                {
                    // Check if the inputStream was opened
                    if (fileStream != null)
                    {
                        fileStream.Dispose();
                    }
                }
            }

            return strReturnMsg;



        }







        public string JIBPreparePDF(int JIBHeaderKey)
        {
            //int jibHeaderKey = 0;
            string strReturnMsg = "ok";

            //Stream? fileStream = null;
            //Microsoft.SharePoint.Client.File? file = null;
            var vLoginSharepoint = vapJIBSharepoint().FirstOrDefault();
            SharepointService? spSVR = null;


            try
            {

                if (vLoginSharepoint == null)
                {
                    throw new InvalidOperationException("vapJIBSharepoint login info failed to load");
                }

                //(fileStream, string strfilename) = spSVR.getSharepointFileStream("https://h20bridge.sharepoint.com/sites/AccountingGroup", "Shared Documents/04. Accounts Payable/AP ticket attachments/AP000002", "");
                spSVR = new SharepointService(vLoginSharepoint.UserID, vLoginSharepoint.Password);

                var jibHeader = vapJIBHeaderbyKey(JIBHeaderKey).SingleOrDefault();

                if (jibHeader == null)
                {
                    throw new InvalidOperationException("JIB header not found");
                }
                string filenameWithoutExtension = Path.GetFileNameWithoutExtension(jibHeader.Filename);

                // Append "_INVOICES" and change the extension to ".pdf"
                string newFilename = $"{filenameWithoutExtension}_INVOICES.pdf";


                var jibFiles = vapJIBPDFsbyKey(JIBHeaderKey).ToList();
                if (!jibFiles.Any())
                {
                    throw new InvalidOperationException("JIB PDF file(s) not found");
                }

                List<Stream> pdfStreams = new List<Stream>();


                foreach (var jibFile in jibFiles)
                {
                    byte[] fileBytes = Convert.FromBase64String(jibFile.DocFileDataBase64);

                    bool individualUploadSuccess = UploadIndividualFile(spSVR,
                                                                vLoginSharepoint.ServerSiteUrl,
                                                                vLoginSharepoint.LibaryURL + "/Archived",
                                                                jibHeader.JIBNO,
                                                                jibFile.DocFileName,
                                                                fileBytes);
                    if (!individualUploadSuccess)
                    {
                        throw new InvalidOperationException($"Failed to upload individual file: {jibFile.DocFileName}");
                    }

                    MemoryStream ms = new MemoryStream(fileBytes);
                    pdfStreams.Add(ms);

                }

                using (PdfDocument finalDocument = new PdfDocument())
                {
                    PdfDocumentBase.Merge(finalDocument, pdfStreams.ToArray());

                    using (MemoryStream msMergedPdf = new MemoryStream())
                    {
                        finalDocument.Save(msMergedPdf);
                        msMergedPdf.Position = 0;

                        byte[] mergedPdfBytes = msMergedPdf.ToArray();

                        bool uploadSuccess = spSVR.UploadFiles(vLoginSharepoint.ServerSiteUrl,
                                                               vLoginSharepoint.LibaryURL + "/Archived",
                                                               jibHeader.JIBNO,
                                                               newFilename,
                                                               mergedPdfBytes);

                        if (!uploadSuccess)
                        {
                            throw new InvalidOperationException("Failed to upload merged PDF file");
                        }
                    }
                }

                return jibHeader.SharepointLoc;
            }
            catch (Exception ex)
            {
                strReturnMsg = $"Error: {ex.Message}";
                return strReturnMsg;
            }
        }

        private bool UploadIndividualFile(SharepointService spSVR, string serverSiteUrl, string libraryUrl, string jibNO, string fileName, byte[] fileContent)
        {
            try
            {

                return spSVR.UploadFiles(serverSiteUrl, libraryUrl, jibNO, fileName, fileContent);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error uploading file {fileName}: {ex.Message}");
                return false;
            }
        }


        public bool MoveFile(SharepointService spSVR, Microsoft.SharePoint.Client.File file, string ServerSiteUrl, string LibaryURL, string JIBNO)
        {
            try
            {

                Microsoft.SharePoint.Client.ClientContext clientContext;
                clientContext = spSVR.getSharepointClientContext(ServerSiteUrl);

                if (spSVR.Connect(ServerSiteUrl))
                {
                    var parentFolder = clientContext.Web.GetFolderByServerRelativeUrl(ServerSiteUrl + "/" + LibaryURL);
                    clientContext.Load(parentFolder);

                    var newFolder = parentFolder.Folders.Add(JIBNO);
                    clientContext.ExecuteQuery();

                    string strfilename = ServerSiteUrl + "/" + LibaryURL + "/" + JIBNO + "/" + file.Name;
                    file.MoveTo(strfilename, Microsoft.SharePoint.Client.MoveOperations.Overwrite);
                    clientContext.ExecuteQuery();

                }
                return true;
            }
            catch (Exception ex)
            {
                throw;
            }
        }

        public static byte[] streamToByteArray(Stream input)
        {
            MemoryStream ms = new MemoryStream();
            input.CopyTo(ms);
            return ms.ToArray();
        }
        //public async Task<bool> bExportPDFs(int jibHeaderKey, string strfilename)

        //public bool bExportPDFs(int jibHeaderKey, string strfilename)
        public bool bExportPDFs(int jibHeaderKey, Stream fileStream)
        {

            string strSheetToPdf = "";
            try
            {

                //--upload excel -----------------------------------------------------
                //---------------------------------------------------------------------

                //bool bexcelupload = JIBUploadExcel(jibHeaderKey, strfilename);

                //--- end upload excel ---------------------------------------------
                //---------------------------------------------------------------------

                IEnumerable<tapEIBSubmitSupplierInv> eibRecords = GettapEIBSubmitSupplierInvForPDF(jibHeaderKey);

                string result = JIBExceltoPDF(fileStream, eibRecords);
                if (result != "ok")
                {
                    // Handle the error
                    throw new InvalidOperationException(result);
                }

                //foreach (var record in eibRecords)
                //{
                //    strSheetToPdf = record.SheetToPdf;


                //    string[] lines = strSheetToPdf.Split(';');
                //    foreach (string strSingleSheet in lines)
                //    {


                //        string successpdf = JIBExceltoPDF(fileStream, strSingleSheet, record.EIBSubmitSupplierInvKey);

                //        if (successpdf != "ok")
                //        {
                //         bool recy = spapUpdateJIBHeader(
                //jibHeaderKey,
                //null,            // strStatus
                //null,            // strFilename
                //null,            // strSharepointLoc
                //record.SheetToPdf + " PDF NOT SAVED " + successpdf,            // strComments
                //null,            // strResult
                //null,            //FinishedDatetime
                //null            // dtSendDateTime (this should be a nullable DateTime type in your method signature)
                //);
                //        }


                //    }

                //}


                return true;

            }
            catch (Exception ex)
            {
                //  bool recy = spapUpdateJIBHeader(
                //jibHeaderKey,
                //null,            // strStatus
                //null,            // strFilename
                //null,            // strSharepointLoc
                //strSheetToPdf + " Error: " + ex.Message,            // strComments
                //null,            // strResult
                //null,             //@FinishedDatetime
                //null            // dtSendDateTime 
                //);
                throw;
                //return false;

            }


        }


        public bool JIBUploadExcel(int jibHeaderKey, string strfilename)
        {

            try
            {
                //byte[] excelBytes = File.ReadAllBytes(strfilename);
                //string base64EncodedExcel = Convert.ToBase64String(excelBytes);
                //string filenameOnly = Path.GetFileName(strfilename);

                ////bool rec = await spapAddAttachment(
                //bool rec = spapAddAttachment(
                //                   filenameOnly,
                //                   "tarJIBHeader",
                //                   jibHeaderKey,
                //                    "application/vnd.ms-excel",
                //                   base64EncodedExcel
                //                   );

                //return true;

                byte[] excelBytes = System.IO.File.ReadAllBytes(strfilename);
                string filenameOnly = Path.GetFileName(strfilename);
                //int chunkSize = 1 * 1024 * 200; // 1 * 1024 * 1024= 1 MB
                string base64EncodedFile = Convert.ToBase64String(excelBytes);

                int base64ChunkSize = 4 * 1024 * 256;
                int numberOfBase64Chunks = (int)Math.Ceiling((double)base64EncodedFile.Length / base64ChunkSize);


                for (int i = 0; i < numberOfBase64Chunks; i++)
                {
                    int startIndex = i * base64ChunkSize;
                    int length = Math.Min(base64ChunkSize, base64EncodedFile.Length - startIndex);
                    string base64Chunk = base64EncodedFile.Substring(startIndex, length);



                    // Now, call the stored procedure to save the current accumulated chunks to the database
                    bool rec = spapAddAttachment(
                        filenameOnly,
                        "tapJIBHeader",
                        jibHeaderKey,
                        "application/vnd.ms-excel",
                        base64Chunk
                    );

                    if (!rec)
                    {
                        // Handle the error
                        return false;
                    }
                }

                return true;

            }
            catch (Exception ex)
            {
                bool recy = spapUpdateJIBHeader(
             jibHeaderKey,
             null,            // strStatus
             null,            // strFilename
             null,            // strSharepointLoc
             strfilename + " Error: " + ex.Message,            // strComments
             null,            // strResult
             null,          //FinishedDatetime
             null            // dtSendDateTime 
             );
                return false;
            }
        }



        public IEnumerable<tapEIBSubmitSupplierInv> GettapEIBSubmitSupplierInvForPDF(int JIBHeaderKey)
        {
            return _context.tapEIBSubmitSupplierInv
                .Where(a => a.PrimaryKey == JIBHeaderKey)
                .Where(a => a.SheetToPdf != null && a.SheetToPdf.Length > 2)
                .Where(a => a.LinkTable == "tapJIBHeader");
            ;
        }


        static string GetExcelColumnName(int columnIndex)
        {
            string columnName = String.Empty;
            int modulo;

            while (columnIndex > 0)
            {
                modulo = (columnIndex - 1) % 26;
                columnName = Convert.ToChar('A' + modulo) + columnName;
                columnIndex = (columnIndex - modulo) / 26;
            }

            return columnName;
        }

        public bool InsertData(DataTable dtTable, int jibHeaderKey, int originalOrder, string eibuploadSheetName)
        {
            //int originalOrder = 1; // Initialize a counter outside the loop
            try

            {
                foreach (DataRow row in dtTable.Rows)
                {
                    // Create an instance of the model and populate it with data from the DataRow
                    var fromBody = new EIBSubmitSupplierInvForAdd();



                    fromBody.LinkTable = "tapJIBHeader";
                    fromBody.OriginalOrder = originalOrder++;
                    fromBody.PrimaryKey = Convert.ToInt32(jibHeaderKey);
                    fromBody.SheetName = eibuploadSheetName;
                    fromBody.SheetToPdf = row["A"] != DBNull.Value ? (string)row["A"] : null;
                    fromBody.SpreadsheetKey = Convert.ToInt32(row["B"]);
                    fromBody.Submit = row["F"] != DBNull.Value ? (string)row["F"] : null;
                    fromBody.LockedInWorkday = row["G"] != DBNull.Value ? (string)row["G"] : null;
                    fromBody.InvoiceAccountingDate = DateTime.TryParse(row["L"].ToString(), out var parsedDate) ? parsedDate : (DateTime?)null;
                    fromBody.Company = row["M"] != DBNull.Value ? (string)row["M"] : null;
                    fromBody.Currency = row["O"] != DBNull.Value ? (string)row["O"] : null;
                    fromBody.Supplier = row["P"] != DBNull.Value ? (string)row["P"] : null;
                    fromBody.InvoiceDate = DateTime.TryParse(row["Y"].ToString(), out var invDate) ? invDate : (DateTime?)null;
                    fromBody.InvoiceReceivedDate = DateTime.TryParse(row["Z"].ToString(), out var invrecDate) ? invrecDate : (DateTime?)null;
                    fromBody.ControlAmount = decimal.TryParse(row["AE"].ToString(), out var parsedUControlAmount) ? parsedUControlAmount : (decimal?)null;
                    fromBody.SuppliersInvoiceNumber = row["AN"] != DBNull.Value ? (string)row["AN"] : null;
                    fromBody.Memo = row["AU"] != DBNull.Value ? (string)row["AU"] : null;
                    fromBody.RowID = int.TryParse(row["CY"].ToString(), out var intRowid) ? intRowid : (int?)null;
                    fromBody.IntercompanyAffiliate = row["DB"] != DBNull.Value ? (string)row["DB"] : null;
                    fromBody.SpendCategory = row["DK"] != DBNull.Value ? (string)row["DK"] : null;
                    fromBody.Quantity = decimal.TryParse(row["EN"].ToString(), out var parsedQuantity) ? parsedQuantity : (decimal?)null;
                    fromBody.UnitCost = decimal.TryParse(row["EP"].ToString(), out var parsedUnitCost) ? parsedUnitCost : (decimal?)null;
                    fromBody.ExtendedAmount = decimal.TryParse(row["EQ"].ToString(), out var parsedUExtendedAmount) ? parsedUExtendedAmount : (decimal?)null;
                    fromBody.Memo2 = row["EW"] != DBNull.Value ? (string)row["EW"] : null;
                    fromBody.Project = row["EX"] != DBNull.Value ? (string)row["EX"] : null;
                    fromBody.CostCenter = (string)row["EY"];
                    fromBody.Site = (string)row["EZ"];
                    fromBody.ServiceMonth = (string)row["FA"];
                    fromBody.ServiceYear = (string)row["FB"];
                    fromBody.IntercompanyAffiliate2 = row["FC"] != DBNull.Value ? (string)row["FC"] : null;


                    //fromBody.RowID2 = int.TryParse(row["GM"].ToString(), out var intRowid2) ? intRowid2 : (int?)null;
                    fromBody.RowID2 = int.TryParse(row["GM"].ToString(), out var intRowid2) && intRowid2 != 0 ? intRowid2 : (int?)null;

                    if (!fromBody.RowID.HasValue || fromBody.RowID == 0)
                    {
                        continue;
                    }


                    bool isSuccess = spapInsertEIBSubmitSupplierInv(fromBody);
                    if (!isSuccess)
                    {

                        return false;
                    }
                }

                return true;
            }
            catch (Exception)
            {
                throw;
            }
        }

        public bool spapInsertEIBSubmitSupplierInv(EIBSubmitSupplierInvForAdd fromBody)
        {
            bool breturn = false;
            String sSQLPara = "";
            List<SqlParameter> updatePara = new List<SqlParameter>();

            updatePara.Add(new SqlParameter("@PrimaryKey", fromBody.PrimaryKey));
            updatePara.Add(new SqlParameter("@LinkTable", fromBody.LinkTable ?? (object)DBNull.Value));
            updatePara.Add(new SqlParameter("@OriginalOrder", fromBody.OriginalOrder));
            updatePara.Add(new SqlParameter("@SheetName", fromBody.SheetName ?? (object)DBNull.Value));
            updatePara.Add(new SqlParameter("@SheetToPdf", fromBody.SheetToPdf ?? (object)DBNull.Value));
            updatePara.Add(new SqlParameter("@SpreadsheetKey", fromBody.SpreadsheetKey));
            updatePara.Add(new SqlParameter("@Submit", fromBody.Submit ?? (object)DBNull.Value));
            updatePara.Add(new SqlParameter("@LockedInWorkday", fromBody.LockedInWorkday ?? (object)DBNull.Value));
            updatePara.Add(new SqlParameter("@InvoiceAccountingDate", fromBody.InvoiceAccountingDate ?? (object)DBNull.Value));
            updatePara.Add(new SqlParameter("@Company", fromBody.Company ?? (object)DBNull.Value));
            updatePara.Add(new SqlParameter("@Currency", fromBody.Currency ?? (object)DBNull.Value));
            updatePara.Add(new SqlParameter("@Supplier", fromBody.Supplier ?? (object)DBNull.Value));
            updatePara.Add(new SqlParameter("@InvoiceDate", fromBody.InvoiceDate ?? (object)DBNull.Value));
            updatePara.Add(new SqlParameter("@InvoiceReceivedDate", fromBody.InvoiceReceivedDate ?? (object)DBNull.Value));
            updatePara.Add(new SqlParameter("@ControlAmount", fromBody.ControlAmount ?? (object)DBNull.Value));
            updatePara.Add(new SqlParameter("@SuppliersInvoiceNumber", fromBody.SuppliersInvoiceNumber ?? (object)DBNull.Value));
            updatePara.Add(new SqlParameter("@Memo", fromBody.Memo ?? (object)DBNull.Value));
            updatePara.Add(new SqlParameter("@RowID", fromBody.RowID ?? (object)DBNull.Value));
            updatePara.Add(new SqlParameter("@IntercompanyAffiliate", fromBody.IntercompanyAffiliate ?? (object)DBNull.Value));
            updatePara.Add(new SqlParameter("@SpendCategory", fromBody.SpendCategory ?? (object)DBNull.Value));
            updatePara.Add(new SqlParameter("@Quantity", fromBody.Quantity ?? (object)DBNull.Value));
            updatePara.Add(new SqlParameter("@UnitCost", fromBody.UnitCost ?? (object)DBNull.Value));
            updatePara.Add(new SqlParameter("@ExtendedAmount", fromBody.ExtendedAmount ?? (object)DBNull.Value));
            updatePara.Add(new SqlParameter("@Memo2", fromBody.Memo2 ?? (object)DBNull.Value));
            updatePara.Add(new SqlParameter("@Project", fromBody.Project ?? (object)DBNull.Value));
            updatePara.Add(new SqlParameter("@CostCenter", fromBody.CostCenter ?? (object)DBNull.Value));
            updatePara.Add(new SqlParameter("@Site", fromBody.Site ?? (object)DBNull.Value));
            updatePara.Add(new SqlParameter("@ServiceMonth", fromBody.ServiceMonth ?? (object)DBNull.Value));
            updatePara.Add(new SqlParameter("@ServiceYear", fromBody.ServiceYear ?? (object)DBNull.Value));
            updatePara.Add(new SqlParameter("@IntercompanyAffiliate2", fromBody.IntercompanyAffiliate2 ?? (object)DBNull.Value));
            updatePara.Add(new SqlParameter("@RowID2", fromBody.RowID2 ?? (object)DBNull.Value));


            SqlParameter ObjectLinkKey = new SqlParameter();
            ObjectLinkKey.ParameterName = "EIBSubmitSupplierInvKey";
            ObjectLinkKey.Direction = ParameterDirection.Output;
            ObjectLinkKey.SqlDbType = SqlDbType.Int;

            updatePara.Add(ObjectLinkKey);

            sSQLPara = "@PrimaryKey, @LinkTable, @OriginalOrder, @SheetName, @SheetToPdf, @SpreadsheetKey, @Submit, @LockedInWorkday, @InvoiceAccountingDate";
            sSQLPara = sSQLPara + ",@Company, @Currency, @Supplier, @InvoiceDate, @InvoiceReceivedDate,@ControlAmount, @SuppliersInvoiceNumber";
            sSQLPara = sSQLPara + ",@Memo, @RowID, @IntercompanyAffiliate, @SpendCategory, @Quantity, @UnitCost, @ExtendedAmount, @Memo2";
            sSQLPara = sSQLPara + ",@Project, @CostCenter, @Site, @ServiceMonth, @ServiceYear, @IntercompanyAffiliate2, @RowID2, @EIBSubmitSupplierInvKey output";

            try
            {
                //_context.Database.ExecuteSqlCommand("spapInsertIntoEIBSubmitSupplierInv " + sSQLPara, updatePara);
                _context.Database.ExecuteSqlRaw("EXEC spapInsertEIBSubmitSupplierInv " + sSQLPara, updatePara.ToArray());
                //resultKey = (int)ObjectLinkKey.Value;
                breturn = true;

                return breturn;
            }
            catch (Exception)
            {
                throw;
            }

        }

        public int spapInsertJIBHeader(string strFilename, string strSharepointLoc)
        {
            String sSQLPara = "";
            List<SqlParameter> updatePara = new List<SqlParameter>();

            updatePara.Add(new SqlParameter("@Status", "Started"));
            updatePara.Add(new SqlParameter("@Filename", strFilename));
            updatePara.Add(new SqlParameter("@SharepointLoc", strSharepointLoc));
            updatePara.Add(new SqlParameter("@Comments", "new"));



            SqlParameter ObjectLinkKey = new SqlParameter();
            ObjectLinkKey.ParameterName = "JIBHeaderKey";
            ObjectLinkKey.Direction = ParameterDirection.Output;
            ObjectLinkKey.SqlDbType = SqlDbType.Int;

            updatePara.Add(ObjectLinkKey);

            sSQLPara = "@Status, @Filename, @SharepointLoc, @Comments, @JIBHeaderKey output";

            try
            {
                //_context.Database.ExecuteSqlCommand("spapInsertIntoEIBSubmitSuppliertInv " + sSQLPara, updatePara);
                _context.Database.ExecuteSqlRaw("EXEC spapInsertJIBHeader " + sSQLPara, updatePara.ToArray());
                int jibHeaderKey = (int)ObjectLinkKey.Value;
                return jibHeaderKey;

            }
            catch (Exception)
            {
                throw;
            }

        }


        public Boolean spapUpdateJIBHeader(int intJIBHeaderKey, string? strStatus, string? strFilename, string? strSharepointLoc, string? strComments, string? strResult, DateTime? dtFinishedDatetime, DateTime? dtSendDateTime)
        {
            String sSQLPara = "";
            List<SqlParameter> updatePara = new List<SqlParameter>();

            updatePara.Add(new SqlParameter("@JIBHeaderKey", intJIBHeaderKey));
            updatePara.Add(new SqlParameter("@Status", strStatus ?? (object)DBNull.Value));
            updatePara.Add(new SqlParameter("@Filename", strFilename ?? (object)DBNull.Value));
            updatePara.Add(new SqlParameter("@SharepointLoc", strSharepointLoc ?? (object)DBNull.Value));
            updatePara.Add(new SqlParameter("@Comments", strComments ?? (object)DBNull.Value));
            updatePara.Add(new SqlParameter("@Result", strResult ?? (object)DBNull.Value));
            updatePara.Add(new SqlParameter("@FinishedDatetime", dtFinishedDatetime.HasValue ? (object)dtFinishedDatetime.Value : DBNull.Value));
            updatePara.Add(new SqlParameter("@SendDateTime", dtSendDateTime.HasValue ? (object)dtSendDateTime.Value : DBNull.Value));
            updatePara.Add(new SqlParameter("@WorkdayAction", (object)DBNull.Value)); //used in api when send to wd or sandbox
            updatePara.Add(new SqlParameter("@SendBy", (object)DBNull.Value)); //used in api when send to wd or sandbox




            //SqlParameter ObjectLinkKey = new SqlParameter();
            //ObjectLinkKey.ParameterName = "JIBHeaderKey";
            //ObjectLinkKey.Direction = ParameterDirection.Output;
            //ObjectLinkKey.SqlDbType = SqlDbType.Int;

            //updatePara.Add(ObjectLinkKey);

            sSQLPara = "@JIBHeaderKey,@Status,@Filename,@SharepointLoc,@Comments,@Result,@FinishedDatetime,@SendDateTime,@WorkdayAction,@SendBy";

            try
            {
                //_context.Database.ExecuteSqlCommand("spapInsertIntoEIBSubmitSupplierInv " + sSQLPara, updatePara);
                _context.Database.ExecuteSqlRaw("EXEC spapUpdateJIBHeader " + sSQLPara, updatePara.ToArray());
                //int jibHeaderKey = (int)ObjectLinkKey.Value;
                //return jibHeaderKey;

            }
            catch (Exception)
            {
                //throw new Exception (ex.Message);
                throw;
                //return false;
            }
            return true;

        }


        public string JIBExceltoPDF(Stream fileStream, IEnumerable<tapEIBSubmitSupplierInv> eibRecords)
        {
            IWorkbook workbook = null;

            try
            {
                using (ExcelEngine excelEngine = new ExcelEngine())
                {
                    IApplication application = excelEngine.Excel;
                    application.DefaultVersion = ExcelVersion.Xlsx;

                    fileStream.Position = 0; // Reset stream position
                    workbook = application.Workbooks.Open(fileStream);

                    XlsIORenderer renderer = new XlsIORenderer();

                    foreach (var record in eibRecords)
                    {
                        string[] sheetNames = record.SheetToPdf.Split(';');
                        foreach (string sheetName in sheetNames)
                        {
                            try
                            {



                                IWorksheet wsSheet1 = workbook.Worksheets[sheetName];
                                if (wsSheet1 == null)
                                {
                                    LogPDFError(record, $"Sheet not found");
                                    continue; // Continue to the next sheetName
                                }

                                using (PdfDocument pdfDocument = renderer.ConvertToPDF(wsSheet1))
                                {
                                    using (MemoryStream pdfStream = new MemoryStream())
                                    {
                                        pdfDocument.Save(pdfStream);
                                        byte[] btByt = pdfStream.ToArray();

                                        string base64EncodedPDF = Convert.ToBase64String(btByt);
                                        bool rec = spapAddAttachment(
                                            sheetName + ".pdf",
                                            "tapEIBSubmitSupplierInv",
                                            record.EIBSubmitSupplierInvKey,
                                            "application/pdf",
                                            base64EncodedPDF
                                        );

                                        if (!rec)
                                        {
                                            LogPDFError(record, $"Failed to add attachment for sheet");
                                            continue;
                                        }
                                    }
                                }
                            }
                            catch (Exception ex)
                            {
                                LogPDFError(record, $"Error processing sheet: {ex.Message}");
                                continue;
                            }
                        }
                    }
                }
                return "ok";
            }
            catch (OutOfMemoryException ex)
            {
                return "Out of memory: " + ex.Message;
            }
            catch (Exception ex)
            {
                return ex.Message;
            }
            finally
            {
                workbook?.Close();
            }
        }


        private void LogPDFError(tapEIBSubmitSupplierInv record, string errorMessage)
        {
            spapUpdateJIBHeader(
                record.PrimaryKey,
                null, // strStatus
                null, // strFilename
                null, // strSharepointLoc
                record.SheetToPdf + " PDF NOT SAVED " + errorMessage, // strComments
                null, // strResult
                null, // FinishedDatetime
                null  // dtSendDateTime (this should be a nullable DateTime type in your method signature)
            );
        }






        public string JIBExceltoPDFprev(Stream fileStream, string strSheetToPdf, int intEIBSubmitSupplierInvKey)
        {
            IWorkbook workbook = null;

            try
            {
                //SharepointService spSVR = new SharepointService("WBRSAAPTicket@h2obridge.com", "7JwYq*V%w5g9m#");

                using (ExcelEngine excelEngine = new ExcelEngine())
                {
                    IApplication application = excelEngine.Excel;
                    application.DefaultVersion = ExcelVersion.Xlsx;

                    fileStream.Position = 0; // Reset stream position
                    workbook = application.Workbooks.Open(fileStream);

                    IWorksheet wsSheet1 = workbook.Worksheets[strSheetToPdf];
                    if (wsSheet1 == null)
                    {
                        return "Sheet not found";
                    }

                    XlsIORenderer renderer = new XlsIORenderer();
                    using (PdfDocument pdfDocument = renderer.ConvertToPDF(wsSheet1))
                    {
                        //string pdfFilePath = $"c:\\temp\\jib\\{wsSheet1.Name}.pdf";
                        using (MemoryStream pdfStream = new MemoryStream())
                        {
                            pdfDocument.Save(pdfStream);
                            byte[] btByt = pdfStream.ToArray();

                            string base64EncodedPDF = Convert.ToBase64String(btByt);
                            bool rec = spapAddAttachment(
                                strSheetToPdf + ".pdf",
                                "tapEIBSubmitSupplierInv",
                                intEIBSubmitSupplierInvKey,
                                "application/pdf",
                                base64EncodedPDF
                            );

                            //pdfStream.Dispose();
                        }
                    }
                }

                return "ok";
            }
            catch (OutOfMemoryException ex)
            {
                return "Out of memory: " + ex.Message;
            }
            catch (Exception ex)
            {
                return ex.Message;
            }
            finally
            {
                if (workbook != null)
                {
                    workbook.Close();
                }
            }
        }



        //public async Task<bool> spapAddAttachment(string? strFilename, string strLinkTableName, int intLinkPrimaryKey, string strDocfiletype, string content)
        public bool spapAddAttachment(string? strFilename, string strLinkTableName, int intLinkPrimaryKey, string strDocfiletype, string content)
        {
            String sSQLPara = "";
            List<SqlParameter> updatePara = new List<SqlParameter>();

            updatePara.Add(new SqlParameter("@DocFileName", strFilename));
            updatePara.Add(new SqlParameter("@LinkTableName", strLinkTableName));
            updatePara.Add(new SqlParameter("@LinkPrimaryKey", intLinkPrimaryKey));
            updatePara.Add(new SqlParameter("@DocFileType", strDocfiletype));
            updatePara.Add(new SqlParameter("@DocFileDataBase64", content));




            sSQLPara = "@DocFileName,@LinkTableName,@LinkPrimaryKey,@DocFileType,@DocFileDataBase64";

            try
            {
                _context.Database.ExecuteSqlRaw("EXEC spapAddAttachment " + sSQLPara, updatePara.ToArray());
                //await _context.Database.ExecuteSqlRawAsync("EXEC spapAddAttachment " + sSQLPara, updatePara.ToArray());
                return true;
                //int jibHeaderKey = (int)ObjectLinkKey.Value;
                //return jibHeaderKey;

            }
            catch (Exception)
            {
                //throw new Exception (ex.Message);
                throw;
                //return false;
            }
            return true;

        }

        public int JIBExceltoPDFOriginal()
        {
            string[] keywordsToExclude = new string[] { "summary", "Piped Allocations", "Revenue Pivot", "Journal Lines",
                "Coding", "EIB Upload", "Firefox", "JIB_Capex", "Data Audit", "payments", "volumes", "Spend Categories" };


            using (ExcelEngine excelEngine = new ExcelEngine())
            {
                IApplication application = excelEngine.Excel;
                application.DefaultVersion = ExcelVersion.Xlsx;
                FileStream excelStream = new FileStream("c:\\temp\\jib\\EVX JIB Workbook 01.31.24.xlsx", FileMode.Open, FileAccess.Read);
                IWorkbook workbook = application.Workbooks.Open(excelStream);

                IWorksheet wsSheet1 = workbook.Worksheets["LE138 EIB Upload"];

                if (wsSheet1 != null)
                {
                    //var value = specificSheet.Range["A1"].Value;


                    fnExportSheet(wsSheet1);
                }






                XlsIORenderer renderer = new XlsIORenderer();



                //Convert Excel document into PDF document 
                //PdfDocument pdfDocument = renderer.ConvertToPDF(workbook);

                //Stream stream = new FileStream("c:\\temp\\jib\\ExcelToPDF.pdf", FileMode.Create, FileAccess.ReadWrite);
                //pdfDocument.Save(stream);
                for (int i = 0; i < workbook.Worksheets.Count; i++)
                {
                    IWorksheet sheet = workbook.Worksheets[i];

                    if (keywordsToExclude.Any(keyword => sheet.Name.IndexOf(keyword, StringComparison.OrdinalIgnoreCase) >= 0))

                    {
                        continue;
                    }



                    // Convert the sheet to a PDF document
                    PdfDocument pdfDocument = renderer.ConvertToPDF(sheet);

                    // Define the PDF file path for each sheet
                    string pdfFilePath = $"c:\\temp\\jib\\{sheet.Name}.pdf";

                    // Save the PDF document
                    using (FileStream pdfStream = new FileStream(pdfFilePath, FileMode.Create, FileAccess.Write))
                    {
                        pdfDocument.Save(pdfStream);
                        pdfStream.Dispose();
                    }
                }



                excelStream.Dispose();

            }


            return 1;
        }

    }
}
