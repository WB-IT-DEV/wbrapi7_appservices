using System.ComponentModel.DataAnnotations;

namespace wbrapi7_appservices.Entities
{
    public partial class vciSafIncHeadStatus
    {
        [Key]

        public int StatusKey { get; set; }
        public string StatusID { get; set; }
        public string StatusDesc { get; set; }




    }
}
