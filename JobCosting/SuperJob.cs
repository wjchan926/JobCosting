using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace JobCosting
{
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

        public SuperJob() { }

        public SuperJob(string salesOrder, string partNumber)
        {
            this.salesOrder = salesOrder;
            this.partNumber = partNumber;
        }

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
        }

        public virtual void setAmountActualCost()
        {
            amountActualCost = amountActualCost - badCostData + productCost * orderQuantity;
        }

        public void setDifference()
        {
            difference = (double)amountActualRevenue - (double)amountActualCost;
        }

        public void setGrossMargin()
        {
            grossMargin = difference / (double)amountActualRevenue;
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
    }
}
