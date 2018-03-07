using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace JobCosting
{
    public sealed class StockJob : SuperJob
    {
        public double expectedRevenue { get; set; } = 0;

        private StockJob() : base() { }

        public StockJob(string salesOrder, string partNumber, long orderQuantity) : base(salesOrder, partNumber, orderQuantity) { }

        /// <summary>
        /// Overriden setAmountActualCost method.  Stock jobs will never include the bad material data
        /// </summary>
        public override void setAmountActualCost()
        {            
            amountActualCost = amountActualCost + productCost * orderQuantity;          
        }

        public void setAmmountActualRevenue()
        {
            amountActualRevenue = (decimal)expectedRevenue;
        }

        public override void calculateFields()
        {
            setAmmountActualRevenue();
            base.calculateFields();
        }
    }
}
