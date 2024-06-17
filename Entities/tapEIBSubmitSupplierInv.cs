using System.ComponentModel.DataAnnotations;

namespace wbrapi7_appservices.Entities
{
    public partial class tapEIBSubmitSupplierInv
    {
        [Key]

        public int EIBSubmitSupplierInvKey { get; set; }
        public int PrimaryKey { get; set; }
        public string LinkTable { get; set; }
        public string SheetToPdf { get; set; }

       



    }
}
