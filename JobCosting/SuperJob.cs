using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace JobCosting
{
    /// <summary>
    /// Abtract parent job class used as template for other types of ojbs
    /// </summary>
    public abstract class SuperJob
    {
        // Excel
        public string salesOrder { get; set; }
        public string partNumber { get; set; }
        public long orderQuantity { get; set; }

        // Queried
        public string customerName { get; set; }
        public string salesRep { get; set; }
        public decimal freight { get; set; }
        public decimal miscToolingCost { get; set; }
        public decimal productCost { get; set; }
        public decimal amountActualRevenue { get; set; }
        public decimal badCostData { get; set; }

        // Calculated
        public decimal amountActualCost { get; set; } = 0;
        public double marlinFreight { get; protected set; } = 0;
        public double grossMargin { get; protected set; } = 0;
        public double unitHigh { get; protected set; } = 0;
        public double unitMed { get; protected set; } = 0;
        public double unitLow { get; protected set; } = 0;
        public double unitFloor { get; protected set; } = 0;
        public double calcCost { get; protected set; } = 0;
        public double calcRevenue { get; protected set; } = 0;
        public double difference { get; protected set; } = 0;
        public double costToCure { get; protected set; } = 0;

        /// <summary>
        /// Default constructor
        /// </summary>
        public SuperJob() { }

        /// <summary>
        /// Contstructor with Sales Order and Part Number arguments
        /// </summary>
        /// <param name="salesOrder"></param> Sales Order to be Analyzed
        /// <param name="partNumber"></param> Part Number
        public SuperJob(string salesOrder, string partNumber)
        {
            this.salesOrder = salesOrder;
            this.partNumber = partNumber;
        }

        /// <summary>
        /// Constructor with 3 arguments
        /// </summary>
        /// <param name="salesOrder"></param> Sales Order to be Analyzed
        /// <param name="partNumber"></param> Part Number
        /// <param name="orderQuantity"></param> Order Quantity
        public SuperJob(string salesOrder, string partNumber, long orderQuantity)
        {
            this.salesOrder = salesOrder;
            this.partNumber = partNumber;
            this.orderQuantity = orderQuantity;
        }

        // Calculation methods
        /// <summary>
        /// Sets all the calculate values, calls all the calculation methods
        /// </summary>
        public virtual void calculateFields()
        {
            setAmountActualCost();
            setDifference();
            setGrossMargin();
            setUnitHigh();
            setUnitMed();
            setUnitLow();
            setUnitFloor();
            setMarlinFreight();
            setSaleRep();
            setCostToCure();
        }
        
        public virtual void setAmountActualCost()
        {
            if (amountActualCost == 0 && badCostData ==0)
            {
                salesRep = "TimeClock Not Imported";
            }
            amountActualCost = amountActualCost - badCostData + productCost * orderQuantity;
        }

        public void setDifference()
        {
            difference = (double)amountActualRevenue - (double)amountActualCost;
        }

        public void setGrossMargin()
        {
            if (amountActualRevenue == 0)
            {
                grossMargin = 0;
            }
            else
            {
                grossMargin = difference / (double)amountActualRevenue;
            }
        }

        public void setUnitHigh()
        {
            double profitMargin = .42;
            unitHigh = (double)amountActualCost / (1 - profitMargin) / (double)orderQuantity;
        }

        public void setUnitMed()
        {
            double profitMargin = .35;
            unitMed = (double)amountActualCost / (1 - profitMargin) / (double)orderQuantity;
        }

        public void setUnitLow()
        {
            double profitMargin = .3;
            unitLow = (double)amountActualCost / (1 - profitMargin) / (double)orderQuantity;
        }

        public void setUnitFloor()
        {
            double profitMargin = .25;
            unitFloor = (double)amountActualCost / (1 - profitMargin) / (double)orderQuantity;
        }

        public void setMarlinFreight()
        {
            marlinFreight = (double)freight / 1.75;
        }

        public void setSaleRep()
        {
            if(amountActualRevenue == 0)
            {
                salesRep = "No Revenue for Job";
            }
        }

        public void setCostToCure()
        {
            if (grossMargin < .42)
            {
                costToCure = -difference + unitHigh * orderQuantity;
            }
        }
    }
}
