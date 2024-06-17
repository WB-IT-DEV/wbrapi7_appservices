using System.ComponentModel.DataAnnotations;

namespace wbrapi7_appservices.Entities
{
    public partial class vapJIBHeader
    {
        [Key]

        public int JIBHeaderKey { get; set; }
        public string? JIBNO { get; set; }

        public DateTime? CreatedDatetime { get; set; }
        public string? Status { get; set; }
        public string? Filename { get; set; }
        public string? SharepointLoc { get; set; }
        public string? Comments { get; set; }
        public string? Result { get; set; }


        public DateTime? SendDateTime { get; set; }



    }
}
