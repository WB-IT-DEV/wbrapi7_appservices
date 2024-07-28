using Microsoft.EntityFrameworkCore;
using System.ComponentModel.DataAnnotations;

namespace wbrapi7_appservices.Entities
{
    [Keyless]
    public partial class vapJIBPDFs
    {
        //[Key]

        public int PrimaryKey { get; set; }
        public string? DocFileDataBase64 { get; set; }
        public string? DocFileName { get; set; }





    }
}
