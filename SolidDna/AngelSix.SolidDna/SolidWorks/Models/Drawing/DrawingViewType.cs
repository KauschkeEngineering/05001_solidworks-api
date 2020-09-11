using SolidWorks.Interop.swconst;

namespace AngelSix.SolidDna
{
    /// <summary>
    /// Drawing view types.
    /// <see cref="swDrawingViewTypes_e"/>
    /// </summary>
    public enum DrawingViewType
    {
        /// <summary>
        /// Unknown
        /// </summary>
        Unknown = 0,

        /// <summary>
        /// Drawing sheet
        /// </summary>
        DrawingSheet = swDrawingViewTypes_e.swDrawingSheet,

        /// <summary>
        /// Section view
        /// </summary>
        SectionView = swDrawingViewTypes_e.swDrawingSectionView,

        /// <summary>
        /// Detail view
        /// </summary>
        DetailView = swDrawingViewTypes_e.swDrawingDetailView,

        /// <summary>
        /// Projected (unfolded) view
        /// </summary>
        ProjectedView = swDrawingViewTypes_e.swDrawingProjectedView,

        /// <summary>
        /// Auxiliary view
        /// </summary>
        AuxiliaryView = swDrawingViewTypes_e.swDrawingAuxiliaryView,

        /// <summary>
        /// Standard view
        /// </summary>
        StandardView = swDrawingViewTypes_e.swDrawingStandardView,

        /// <summary>
        /// Named view
        /// </summary>
        NamedView = swDrawingViewTypes_e.swDrawingNamedView,

        /// <summary>
        /// Relative view to the model
        /// </summary>
        RelativeView = swDrawingViewTypes_e.swDrawingRelativeView,

        /// <summary>
        /// Detached view
        /// </summary>
        DetachedView = swDrawingViewTypes_e.swDrawingDetachedView,

        /// <summary>
        /// Alternate position view
        /// </summary>
        AlternatePositionView = swDrawingViewTypes_e.swDrawingAlternatePositionView
    }
}
