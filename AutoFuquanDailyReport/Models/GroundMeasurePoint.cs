using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AutoFuquanDailyReport.Models
{
    /// <summary>
    /// 地表沉降及路基横断面测点沉降只测Z
    /// </summary>
    public class GroundMeasurePoint:MeasurePoint
    {
        public string ZDirection;

        public decimal PreviousAccumulateZ;

        public decimal AccumulateZ;

        public decimal DeltaZ;

        public decimal RateOfChangeZ;
    }
}
