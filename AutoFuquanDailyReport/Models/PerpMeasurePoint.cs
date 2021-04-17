using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AutoFuquanDailyReport.Models
{
    public class PerpMeasurePoint : MeasurePoint
    {
        /// <summary>
        /// 前次累计变化值
        /// </summary>
        public decimal PreviousAccumulateY;

        public decimal PreviousAccumulateX;

        /// <summary>
        /// 本次累计变化值
        /// </summary>
        public decimal AccumulateY;

        public decimal AccumulateX;


        /// <summary>
        /// 本次变化值
        /// </summary>
        public decimal DeltaY;

        public decimal DeltaX;


        /// <summary>
        /// 本次变化速率
        /// </summary>
        public decimal RateOfChangeY;

        public decimal RateOfChangeX;

    }
}
