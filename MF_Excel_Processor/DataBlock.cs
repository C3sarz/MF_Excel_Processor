using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MF_Excel_Processor
{
    /// <summary>
    /// Stores a data unit of the Excel file.
    /// </summary>
    public class DataBlock
    {
        /// <summary>
        /// Item name.
        /// </summary>
        public string Name;

        /// <summary>
        /// Item code.
        /// </summary>
        public long Code;

        /// <summary>
        /// Item quantity.
        /// </summary>
        public double Quantity;

        /// <summary>
        /// Item total (Gs).
        /// </summary>
        public double Total;

        /// <summary>
        /// Row of the item in the output file.
        /// </summary>
        public int DataRow;

        /// <summary>
        /// Constructor.
        /// </summary>
        /// <param name="Name">Item name.</param>
        /// <param name="Code">Item code.</param>
        /// <param name="Quantity">Item quantity.</param>
        /// <param name="Total">Item total (Gs).</param>
        /// <param name="DataRow">Row of the item in the output file.</param>
        public DataBlock(string Name, long Code, double Quantity, double Total, int DataRow)
        {
            this.Name = Name;
            this.Code = Code;
            this.Quantity = Quantity;
            this.Total = Total;
            this.DataRow = DataRow;
        }

    }
}
