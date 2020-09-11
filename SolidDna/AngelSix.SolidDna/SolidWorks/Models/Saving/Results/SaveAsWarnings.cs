using SolidWorks.Interop.swconst;
using System;

namespace AngelSix.SolidDna
{
    /// <summary>
    /// Values for File, Save warnings that can be returned from the IModelDoc2 Save methods. 
    /// These warnings do not cause the File, Save operation to fail.
    /// <see cref="swFileSaveWarning_e"/>
    /// Any warnings of a model save operation. 
    /// Warnings mean the save was successful, but it had some warnings.
    /// </summary>
    /// <remarks>
    /// <para>
    ///     Warnings are bit-masked (flags) so you can get each warning via:
    ///     
    ///     <code>
    ///         warnings.GetFlags();
    ///     </code>
    /// </para>
    /// </remarks>
    [Flags]
    public enum SaveAsWarnings
    {
        /// <summary>
        /// No warnings
        /// </summary>
        None = 0,

        /// <summary>
        /// There was a problem rebuilding the model
        /// </summary>
        RebuildError = swFileSaveWarning_e.swFileSaveWarning_RebuildError,

        /// <summary>
        /// The model needs to be rebuilt
        /// </summary>
        NeedsRebuild = swFileSaveWarning_e.swFileSaveWarning_NeedsRebuild,

        /// <summary>
        /// The drawing views need updating
        /// </summary>
        ViewsNeedUpdate = swFileSaveWarning_e.swFileSaveWarning_ViewsNeedUpdate,

        /// <summary>
        /// The animator needs to be solved
        /// </summary>
        AnimatorNeedToSolve = swFileSaveWarning_e.swFileSaveWarning_AnimatorNeedToSolve,

        /// <summary>
        /// The animator feature has edits
        /// </summary>
        AnimatorFeatureEdits = swFileSaveWarning_e.swFileSaveWarning_AnimatorFeatureEdits,

        /// <summary>
        /// The eDrawing file has a bad selection
        /// </summary>
        EdrwingsBadSelection = swFileSaveWarning_e.swFileSaveWarning_EdrwingsBadSelection,

        /// <summary>
        /// The animator lights has edits
        /// </summary>
        AnimatorLightEdits = swFileSaveWarning_e.swFileSaveWarning_AnimatorLightEdits,

        /// <summary>
        /// The animator camera views have issues
        /// </summary>
        AnimatorCameraViews = swFileSaveWarning_e.swFileSaveWarning_AnimatorCameraViews,

        /// <summary>
        /// The animator section views have issues
        /// </summary>
        AnimatorSectionViews = swFileSaveWarning_e.swFileSaveWarning_AnimatorSectionViews,

        /// <summary>
        /// The file is missing OLE objects
        /// </summary>
        MissingOLEObjects = swFileSaveWarning_e.swFileSaveWarning_MissingOLEObjects,

        /// <summary>
        /// The file is using the opened view only
        /// </summary>
        OpenedViewOnly = swFileSaveWarning_e.swFileSaveWarning_OpenedViewOnly
    }
}
