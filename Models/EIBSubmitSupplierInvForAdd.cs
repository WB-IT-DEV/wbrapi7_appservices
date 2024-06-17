using System;

namespace wbrapi7_appservices.Models
{
    public class EIBSubmitSupplierInvForAdd
    {
        
        public string LinkTable { get; set; } 
        public int OriginalOrder { get; set; }
        public int PrimaryKey { get; set; }
        public string? SheetName { get; set; } 
        public string? SheetToPdf { get; set; } // -- A
        public int SpreadsheetKey { get; set; } // -- B
        public string? Submit { get; set; } // -- F
        public string? LockedInWorkday { get; set; } // -- G
        public DateTime? InvoiceAccountingDate { get; set; } // -- L
        public string? Company { get; set; } // -- M
        public string? Currency { get; set; } // -- O
        public string? Supplier { get; set; } // -- P
        public DateTime? InvoiceDate { get; set; } // -- Y
        public DateTime? InvoiceReceivedDate { get; set; } // -- Z

        public decimal? ControlAmount { get; set; } // -- AE
        public string? SuppliersInvoiceNumber { get; set; } // -- AN
        public string? Memo { get; set; } // -- AU
        public int? RowID { get; set; } // -- CY
        public string? IntercompanyAffiliate { get; set; } // -- DB
        public string? SpendCategory { get; set; } // -- DK
        public decimal? Quantity { get; set; } // -- EN
        public decimal? UnitCost { get; set; } // -- EP
        public decimal? ExtendedAmount { get; set; } // -- EQ
        public string? Memo2 { get; set; } // -- EW
        public string? Project { get; set; } // -- EX
        public string CostCenter { get; set; } // -- EY
        public string Site { get; set; } // -- EZ
        public string ServiceMonth { get; set; } // -- FA
        public string ServiceYear { get; set; } // -- FB
        public string? IntercompanyAffiliate2 { get; set; } // -- FC
        public int? RowID2 { get; set; } // -- GM

        // You may add constructors or methods as needed.
    }
}
