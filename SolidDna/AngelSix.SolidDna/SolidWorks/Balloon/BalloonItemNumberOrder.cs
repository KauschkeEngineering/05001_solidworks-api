using SolidWorks.Interop.swconst;
using System;

namespace AngelSix.SolidDna
{
    /// <summary>
    /// Balloon item ordering options
    /// </summary>
    [Flags]
    public enum BalloonItemNumberOrder
    {
        /// <summary>
        /// Do not change the item numbers
        /// </summary>
        DoNotChangeItemNumbers = swBalloonItemNumbersOrder_e.swBalloonItemNumbers_DoNotChangeItemNumbers,

        /// <summary>
        /// Follow the same order as the assembly
        /// </summary>
        FollowAssemblyOrder = swBalloonItemNumbersOrder_e.swBalloonItemNumbers_FollowAssemblyOrder,

        /// <summary>
        /// Order sequentially
        /// </summary>
        OrderSequentially = swBalloonItemNumbersOrder_e.swBalloonItemNumbers_OrderSequentially
    }
}
