using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace JobCosting
{
    /// <summary>
    /// Struct that represents the mapping of the columns in Excel Job Costing Doc
    /// </summary>
    public struct ExcelColumn
    {
        public static string partNumber { get; private set; } = "A";
        public static string salesOrder { get; private set; } = "B";
        public static string orderQuantity { get; private set; } = "H";
        public static string expectedAmount { get; private set; } = "K";
        public static string salesRep { get; private set; } = "U";
        public static string actualCost { get; private set; } = "V";
        public static string actualRevenue { get; private set; } = "W";
        public static string difference { get; private set; } = "X";
        public static string grossMargin { get; private set; } = "Y";
        public static string unitHigh { get; private set; } = "Z";
        public static string unitMed { get; private set; } = "AA";
        public static string unitLow { get; private set; } = "AB";
        public static string unitFloor { get; private set; } = "AC";
        public static string freight { get; private set; } = "AD";
        public static string marlinFreight { get; private set; } = "AE";
        public static string miscTooling { get; private set; } = "AF";
    }
}
