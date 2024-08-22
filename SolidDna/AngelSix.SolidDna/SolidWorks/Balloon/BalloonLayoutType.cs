using SolidWorks.Interop.swconst;

namespace AngelSix.SolidDna
{
    /// <summary>
    /// Arrangements for automatic BOM balloons in relation to the drawing views
    /// with which they are associated
    /// </summary>
    public enum BalloonLayoutType
    {
        /// <summary>
        /// Use the document layout default
        /// </summary>
        DocumentDefault = -1,

        /// <summary>
        /// In a box around the drawing view
        /// </summary>
        Square = swBalloonLayoutType_e.swDetailingBalloonLayout_Square,

        /// <summary>
        /// In a circle around the drawing view
        /// </summary>
        Circle = swBalloonLayoutType_e.swDetailingBalloonLayout_Circle,

        /// <summary>
        /// Along the top edge of the drawing view
        /// </summary>
        Top = swBalloonLayoutType_e.swDetailingBalloonLayout_Top,

        /// <summary>
        /// Along the bottom edge of the drawing view
        /// </summary>
        Bottom = swBalloonLayoutType_e.swDetailingBalloonLayout_Bottom,

        /// <summary>
        /// Along the right edge of the drawing view
        /// </summary>
        Right = swBalloonLayoutType_e.swDetailingBalloonLayout_Right,

        /// <summary>
        /// Along the left edge of the drawing view
        /// </summary>
        Left = swBalloonLayoutType_e.swDetailingBalloonLayout_Left
    }
}
