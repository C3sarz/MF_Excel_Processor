using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MF_Excel_Processor
{
    /// <summary>
    /// Contains the type strings for an item in the Excel file.
    /// </summary>
    public class TypeBlock
    {
        /// <summary>
        /// Seccion
        /// </summary>
        public string section { get; private set; }

        /// <summary>
        /// Grupo
        /// </summary>
        public string group { get; private set; }

        /// <summary>
        /// Categoria
        /// </summary>
        public string category { get; private set; }

        /// <summary>
        /// Sub Categoria.
        /// </summary>
        public string subCategory { get; private set; }

        /// <summary>
        /// Constructor.
        /// </summary>
        /// <param name="section"></param>
        /// <param name="group"></param>
        /// <param name="category"></param>
        /// <param name="sub"></param>
        public TypeBlock(string section, string group, string category, string sub)
        {
            this.section = section;
            this.group = group;
            this.category = category;
            this.subCategory = sub;
        }

    }
}
