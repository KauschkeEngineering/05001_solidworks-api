using System;
using SolidWorks.Interop.swconst;

namespace AngelSix.SolidDna
{
    /// <summary>
    /// Document types.
    /// The type of template for a model, from <see cref="swDocTemplateTypes_e"/>
    /// </summary>
    [Flags]
    public enum DocumentTemplateTypes
    {
        /// <summary>
        /// Nothing is open
        /// </summary>
        None = swDocTemplateTypes_e.swDocTemplateTypeNONE,

        /// <summary>
        /// A part is open
        /// </summary>
        Part = swDocTemplateTypes_e.swDocTemplateTypePART,

        /// <summary>
        /// An assembly is open
        /// </summary>
        Assembly = swDocTemplateTypes_e.swDocTemplateTypeASSEMBLY,

        /// <summary>
        /// A drawing is open
        /// </summary>
        Drawing = swDocTemplateTypes_e.swDocTemplateTypeDRAWING,

        /// <summary>
        /// An in-context part is open
        /// </summary>
        InContext = swDocTemplateTypes_e.swDocTemplateTypeInContext
    }
}
