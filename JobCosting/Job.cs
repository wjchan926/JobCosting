using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;
using System.Collections;

namespace JobCosting
{
    class Job
    {
        // Queried
        public string salesOrder;
        public string customerName;
        public string partNumber;
        public string salesRep;
        public decimal freight;
        public decimal marlinFreight;
        public decimal miscToolingCost;
        public decimal productCost;
        public decimal amountActualRevenue;

        // Calculated
        public double grossMargin;
        public double unitHigh;
        public double unitLow;
        public double unitFloor;
        public double calcCost;
        public double calcRevenue;
        public long orderQuantity;

        public Job()
        {

        }
 
        public Job(string salesOrder, string partNumber)
        {
            this.salesOrder = salesOrder;
            this.partNumber = partNumber;
        }
    }    
}
