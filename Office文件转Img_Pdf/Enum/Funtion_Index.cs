using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Office_ToDLL.Enum
{
    /// <summary>
    /// Office方法选择下标
    /// </summary>
    public enum Funtion_Index
    {
        /// <summary>
        /// Excel转图片
        /// </summary>
        Excel_to_img,
        /// <summary>
        /// Excel转PDF
        /// </summary>
        Excel_to_pdf,
        /// <summary>
        /// PDF转图片
        /// </summary>
        Pdf_to_img,
        /// <summary>
        /// PDF转Word
        /// </summary>
        Pdf_to_word,
        /// <summary>
        /// Word转图片
        /// </summary>
        Word_to_img,
        /// <summary>
        /// Word转PDF
        /// </summary>
        Word_to_pdf
        
    }
}
