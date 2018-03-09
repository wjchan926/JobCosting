using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace JobCosting
{
    /// <summary>
    /// Stock Job Class
    /// </summary>
    public sealed class StockJob : SuperJob
    {
        public double expectedRevenue { get; set; } = 0;

        /// <summary>
        /// Private default constructor
        /// </summary>
        private StockJob() : base() { }

        /// <summary>
        /// Overloaded constructor with 3 arguments
        /// </summary>
        /// <param name="salesOrder"></param> Sales Order to be anaylized
        /// <param name="partNumber"></param> Part Number
        /// <param name="orderQuantity"></param> Order Quantity
        public StockJob(string salesOrder, string partNumber, long orderQuantity) : base(salesOrder, partNumber, orderQuantity) { }

        /// <summary>
        /// Overriden setAmountActualCost method.  Stock jobs will never include the bad material data
        /// </summary>
        public override void setAmountActualCost()
        {            
            amountActualCost = amountActualCost + productCost * orderQuantity;          
        }

        /// <summary>
        /// Method that sets revenue for stock jobs.
        /// Calcualted from the Expected Revenue column base on produciton schedule data
        /// </summary>
        public void setAmmountActualRevenue()
        {
          amountActualRevenue = (decimal)expectedRevenue;
        }

        /// <summary>
        /// Overriden calculate fields since revenue has to be adjusted for stock jobs
        /// </summary>
        public override void calculateFields()
        {
            setAmmountActualRevenue();
            base.calculateFields();
        }
    }
}
