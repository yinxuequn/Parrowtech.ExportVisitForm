using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ExportVisitForm
{
    public class VisitReport
    {
        public VisitReport()
        {
            VisitRecords = new List<VisitRecord>();
        }

        public string VisiterCode { get; set; }
        /// <summary>
        /// 拜访人名称
        /// </summary>
        public string VisiterName { get; set; }
        /// <summary>
        /// 门店编码
        /// </summary>
        public string StoreCode { get; set; }
        /// <summary>
        /// 门店名称
        /// </summary>
        public string StoreName { get; set; }
        /// <summary>
        /// 门店地址
        /// </summary>
        public string StoreAddress { get; set; }
        /// <summary>
        /// 不合规次数
        /// </summary>
        public int IrregularCount { get; set; }
        /// <summary>
        /// 基准点不准次数
        /// </summary>
        public int ReferencePointErrorCount { get; set; }
        /// <summary>
        /// 拜访点不准次数
        /// </summary>
        public int VisitPointErrorCount { get; set; }
        /// <summary>
        /// 多次基准点不准原因
        /// </summary>
        public int VisitPointErrorReason { get; set; }

        public IList<VisitRecord> VisitRecords { get; set; }
    }
}
