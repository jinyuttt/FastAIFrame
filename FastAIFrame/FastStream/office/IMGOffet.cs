using System;
using System.Collections.Generic;
using System.Text;

namespace FastStream.office
{

    /// <summary>
    /// Excel中插入图片
    /// </summary>
   public class IMGOffet
    {
        /// <summary>
        /// 起始单元格的x偏移量
        /// </summary>
        public int StartOffsetX;


        /// <summary>
        /// 起始单元格的y偏移量
        /// </summary>
        public int StartOffsetY;

        /// <summary>
        /// 终止单元格的x偏移量
        /// </summary>
        public int EndOffsetX;

        /// <summary>
        /// 终止单元格的y偏移量
        /// </summary>
        public int EndOffsetY;

        /// <summary>
        /// 起始单元格列序号
        /// </summary>
        public int StartCellCol;

        /// <summary>
        /// 起始单元格行序号
        /// </summary>
        public int StartCellRow;

        /// <summary>
        /// 终止单元格列序号
        /// </summary>
        public int EndCellCol;

        /// <summary>
        /// 终止单元格行序号
        /// </summary>
        public int EndCellRow;


    }
}
