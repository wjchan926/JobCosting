using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;
using System.Collections;

namespace JobCosting
{
    public sealed class Job : SuperJob
    {
        /// <summary>
        /// Private defualt constructor for Job Object.  Job must be intialezed with certain fields.
        /// </summary>
        private Job() : base() { }

        /// <summary>
        /// Constructor for Job object with 2 arguments
        /// </summary>
        /// <param name="salesOrder"></param> Alpha portion of the sales order, extract from Excel
        /// <param name="partNumber"></param> Part Number, extract from Excel
        public Job(string salesOrder, string partNumber) : base(salesOrder, partNumber) { }

        /// <summary>
        /// Construct for Job object with 3 arguements
        /// </summary>
        /// <param name="salesOrder"></param> Alpha portion of the sales order, extract from Excel
        /// <param name="partNumber"></param> Part Number, extract from Excel
        /// <param name="orderQuantitiy"></param> Order quantiy, extracted from excel
        public Job(string salesOrder, string partNumber, long orderQuantity) : base(salesOrder, partNumber, orderQuantity) { }
        

    }    
}
