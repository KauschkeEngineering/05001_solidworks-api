using SolidWorks.Interop.swconst;

namespace AngelSix.SolidDna
{
    /// <summary>
    /// The style of a balloon
    /// </summary>
    public enum BalloonStyle
    {
        /// <summary>
        /// Use the document layout default
        /// </summary>
        DocumentDefault = -1,

        /// <summary>
        /// Circular
        /// </summary>
        Circular = swBalloonStyle_e.swBS_Circular,

        /// <summary>
        /// Triangle
        /// </summary>
        Triangle = swBalloonStyle_e.swBS_Triangle,

        /// <summary>
        /// Hexagon
        /// </summary>
        Hexagon = swBalloonStyle_e.swBS_Hexagon,

        /// <summary>
        /// Box
        /// </summary>
        Box = swBalloonStyle_e.swBS_Box,

        /// <summary>
        /// Diamond
        /// </summary>
        Diamond = swBalloonStyle_e.swBS_Diamond,

        /// <summary>
        /// Pentagon. Can be used for label location selection Circular Spit Line
        /// </summary>
        Pentagon = swBalloonStyle_e.swBS_Pentagon,

        /// <summary>
        /// Split circle. Not valid for notes; only valid for balloons
        /// </summary>
        SplitCircle = swBalloonStyle_e.swBS_SplitCirc,

        /// <summary>
        /// Flag pentagon
        /// </summary>
        FlagPentagon = swBalloonStyle_e.swBS_FlagPentagon,

        /// <summary>
        /// Flag triangle
        /// </summary>
        FlagTriangle = swBalloonStyle_e.swBS_FlagTriangle,

        /// <summary>
        /// Underline
        /// </summary>
        Underline = swBalloonStyle_e.swBS_Underline,

        /// <summary>
        /// Square
        /// </summary>
        Square = swBalloonStyle_e.swBS_Square,

        /// <summary>
        /// Square circle
        /// </summary>
        SquareCircle = swBalloonStyle_e.swBS_SCircle,

        /// <summary>
        /// Inspection
        /// </summary>
        Inspection = swBalloonStyle_e.swBS_Inspection,

        /// <summary>
        /// Arc bracket
        /// </summary>
        ArcBracket = swBalloonStyle_e.swBS_ArcBracket,

        /// <summary>
        /// Rectangle bracket
        /// </summary>
        RectangleBracket = swBalloonStyle_e.swBS_RectBracket,

        /// <summary>
        /// Arc length symbol
        /// </summary>
        ArcLengthSymbol = swBalloonStyle_e.swBS_ArclenSym,

        /// <summary>
        /// Fixed symbol
        /// </summary>
        FixedSymbol = swBalloonStyle_e.swBS_FixedSym,

        /// <summary>
        /// Double arrow
        /// </summary>
        DoubleArrow = swBalloonStyle_e.swBS_DoubleArrow,

        /// <summary>
        /// Split square. Can be used for label location selection Square Spit Line
        /// </summary>
        SplitSquare = swBalloonStyle_e.swBS_SplitSquare,

        /// <summary>
        /// Verbose
        /// </summary>
        Verbose = swBalloonStyle_e.swBS_Verbose
    }
}
