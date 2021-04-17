using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AutoFuquanDailyReport.Models
{
    /// <summary>
    /// 桥墩测点
    /// </summary>
    public class PierMeasurePoint : MeasurePoint
    {
        /// <summary>
        /// 本次y方向
        /// </summary>
        public string YDirection;
        //public string YDirection
        //{
        //    get
        //    {
        //        if (AccumulateY > 0)
        //            return "东";
        //        if (AccumulateY < 0)
        //            return "西";
        //        else
        //            return "/";
        //    }
        //}
        /// <summary>
        /// 本次X方向
        /// </summary>
        public string XDirection;

        public string ZDirection;

        /// <summary>
        /// 前次累计变化值
        /// </summary>
        public decimal PreviousAccumulateY;

        public decimal PreviousAccumulateX;

        public decimal PreviousAccumulateZ;

        /// <summary>
        /// 本次累计变化值
        /// </summary>
        public decimal AccumulateY;

        public decimal AccumulateX;

        public decimal AccumulateZ;

        /// <summary>
        /// 本次变化值
        /// </summary>
        public decimal DeltaY;

        public decimal DeltaX;

        public decimal DeltaZ;

        /// <summary>
        /// 本次变化速率
        /// </summary>
        public decimal RateOfChangeY;

        public decimal RateOfChangeX;

        public decimal RateOfChangeZ;
        /// <summary>
        /// 计算方向，正值输出正方向，负值输出负方向，0值输出“/”
        /// </summary>
        /// <param name="value">数值</param>
        /// <param name="positiveDirection">正方向代表的字符</param>
        /// <param name="negativeDirection">负方向代表的字符</param>
        /// <returns></returns>
        private string GetDirection(decimal value,string positiveDirection,string negativeDirection)
        {
            if(value>0)
                return positiveDirection;
            else if(value<0)
                return negativeDirection;
            else
                return "/";
        }
        /// <summary>
        /// 计算方向，正值输出正方向，负值输出负方向，0值输出“/”
        /// </summary>
        /// <param name="positiveDirection">正方向代表的字符</param>
        /// <param name="negativeDirection">负方向代表的字符</param>
        /// <returns></returns>
        public string GetYDirection(string positiveDirection, string negativeDirection)
        {
            return GetDirection(AccumulateY,positiveDirection, negativeDirection);
        }

        public string GetXDirection(string positiveDirection, string negativeDirection)
        {
            return GetDirection(AccumulateX, positiveDirection, negativeDirection);
        }
    }

}
