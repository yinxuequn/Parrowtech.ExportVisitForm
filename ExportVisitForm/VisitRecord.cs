using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ExportVisitForm
{
    public class VisitRecord
    {
        public Guid ID { get; set; }
        /// <summary>
        /// 拜访人编码
        /// </summary>
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
        /// 拜访日期
        /// </summary>
        public DateTime Date { get; set; }
        /// <summary>
        /// 偏差情况
        /// </summary>
        public string Deviation { get; set; }
        /// <summary>
        /// 偏差原因
        /// </summary>
        public string DeviationReason { get; set; }

    }
}
