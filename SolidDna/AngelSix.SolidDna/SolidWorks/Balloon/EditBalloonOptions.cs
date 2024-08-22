using SolidWorks.Interop.swconst;

namespace AngelSix.SolidDna
{
    /// <summary>
    /// Options for the <see cref="AutoBalloonOptions.EditBalloonOption"/>
    /// </summary>
    public enum EditBalloonOptions
    {
        /// <summary>
        /// Replaces existing balloons
        /// </summary>
        Replace = swEditBalloonOption_e.swEditBalloonOption_Replace,

        /// <summary>
        /// Re-sequences existing balloons
        /// </summary>
        Resequence = swEditBalloonOption_e.swEditBalloonOption_Resequence
    }
}
