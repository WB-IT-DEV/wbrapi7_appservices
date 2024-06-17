using System.ComponentModel.DataAnnotations;

namespace wbrapi7_appservices.Entities
{
    public partial class vapJIBSharepoint
    {
        [Key]

        public int CodeDefinitionKey { get; set; }
        public string ServerSiteUrl { get; set; }
        public string LibaryURL { get; set; }
        public string UserID { get; set; }
        public string Password { get; set; }




    }
}
