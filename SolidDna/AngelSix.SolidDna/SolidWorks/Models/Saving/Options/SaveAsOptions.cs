using SolidWorks.Interop.swconst;
using System;

namespace AngelSix.SolidDna
{
    /// <summary>
    /// Save as options, from <see cref="SolidWorks.Interop.swconst.swSaveAsOptions_e"/>
    /// </summary>
    /// <remarks>
    /// <para>
    ///     Options are bit-masked (flags) so you can specify one or more:
    ///     
    ///     <code>
    ///         model.SaveAs("name.sldprt", options: SaveAsOptions.Silent | SaveAsOptions.AvoidRebuildOnSave);
    ///     </code>
    /// </para>
    /// <para>
    ///     These options only apply to saving to native SolidWorks file formats. 
    /// </para>
    /// <para>
    ///     For example, to export a SolidWorks file to a VRML file format, use 
    ///     ISldWorks::GetUserPreferenceToggle and ISldWorks::SetUserPreferenceToggle 
    ///     with swExportVrmlAllComponentsInSingleFile
    /// </para>
    /// </remarks>
    [Flags]
    public enum SaveAsOptions
    {
        /// <summary>
        /// To specify no specific options
        /// </summary>
        None = 0,

        /// <summary>
        /// Saves without any user interaction
        /// </summary>
        Silent = swSaveAsOptions_e.swSaveAsOptions_Silent,

        /// <summary>
        /// Saves the file as a copy
        /// </summary>
        Copy = swSaveAsOptions_e.swSaveAsOptions_Copy,

        /// <summary>
        /// Supports parts, assemblies and drawings; this setting indicates to 
        /// save all components (sub-assemblies and parts) in both assemblies 
        /// and drawings. If a part has an external reference, then this setting
        /// indicates to save the external reference
        /// </summary>
        SaveReferenced = swSaveAsOptions_e.swSaveAsOptions_SaveReferenced,

        /// <summary>
        /// Prevents rebuilding the model prior to saving
        /// </summary>
        AvoidRebuildOnSave = swSaveAsOptions_e.swSaveAsOptions_AvoidRebuildOnSave,

        /// <summary>
        /// Not a valid option for IPartDoc::SaveToFile2; this setting is only
        /// applicable for a drawing that has one or more sheets; this setting 
        /// updates the views on inactive sheets
        /// </summary>
        UpdateInactiveViews = swSaveAsOptions_e.swSaveAsOptions_UpdateInactiveViews,

        /// <summary>
        /// Saves eDrawings-related information into a section of the file being 
        /// saved; specifying this setting overrides the Tools, Options,
        /// System Options, General, Save eDrawings data in SolidWorks document 
        /// setting; not a valid option for IPartDoc::SaveToFile2
        /// </summary>
        OverrideSaveEmodel = swSaveAsOptions_e.swSaveAsOptions_OverrideSaveEmodel,

        /// <summary>
        /// Obsolete. Do not use
        /// </summary>
        [Obsolete]
        SaveEmodelData = swSaveAsOptions_e.swSaveAsOptions_SaveEmodelData,

        /// <summary>
        /// Saves a drawing as a detached drawing.
        /// Not a valid option for IPartDoc::SaveToFile2
        /// </summary>
        DetachedDrawing = swSaveAsOptions_e.swSaveAsOptions_DetachedDrawing,

        /// <summary>
        /// Prune a SolidWorks file's revision history to just the current file name
        /// </summary>
        IgnoreBiography = swSaveAsOptions_e.swSaveAsOptions_IgnoreBiography
    }
}
