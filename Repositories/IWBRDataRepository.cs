using System.IO;
using wbrapi7_appservices.Entities;


namespace wbrapi7_appservices.Repositories
{
    public interface IWBRDataRepository
    {
        //System.Data.SqlClient.SqlConnection SQLDatabaseConnection();
        int apStatementImport(string sTicketNo);

        IEnumerable<vciSafIncHeadStatus> GetvciSafIncHeadStatus();
        IEnumerable<vapJIBSharepoint> vapJIBSharepoint();
        IEnumerable<vapJIBHeader> vapJIBHeaderbyKey(int intjibHeaderKey);
        




        string JIBExceltoPDF(Stream fileStream, string strSheetToPdf, int intEIBSubmitSupplierInvKey);
        //Task<string> JIBExceltoPDF(string strFilename, string strSheetToPdf, int intEIBSubmitSupplierInvKey);



        //Task<int> JItest();
        
        string JIBReadFolder();



    }


}
