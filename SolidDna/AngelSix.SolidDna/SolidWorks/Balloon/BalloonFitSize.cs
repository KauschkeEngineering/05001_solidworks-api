using SolidWorks.Interop.swconst;

namespace AngelSix.SolidDna
{
    /// <summary>
    /// Types of balloon and label location fits
    /// </summary>
    public enum BalloonFitSize
    {
        /// <summary>
        /// Tightest fit (not available for a label location)
        /// </summary>
        Tightest = swBalloonFit_e.swBF_Tightest,

        /// <summary>
        /// Fits a single character
        /// </summary>
        Character1 = swBalloonFit_e.swBF_1Char,

        /// <summary>
        /// Fits 2 characters
        /// </summary>
        Character2 = swBalloonFit_e.swBF_2Chars,

        /// <summary>
        /// Fits 3 characters
        /// </summary>
        Character3 = swBalloonFit_e.swBF_3Chars,

        /// <summary>
        /// Fits 4 characters
        /// </summary>
        Character4 = swBalloonFit_e.swBF_4Chars,

        /// <summary>
        /// Fits 5 characters
        /// </summary>
        Character5 = swBalloonFit_e.swBF_5Chars,

        /// <summary>
        /// Size defined by <see cref="AutoBalloonOptions.CustomSize"/>
        /// </summary>
        UserDefined = swBalloonFit_e.swBF_UserDef
    }
}
