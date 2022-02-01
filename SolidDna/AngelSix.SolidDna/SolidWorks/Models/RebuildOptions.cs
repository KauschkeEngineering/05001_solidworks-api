using SolidWorks.Interop.swconst;
using System;

namespace AngelSix.SolidDna
{
    /// <summary>
    /// Save as options, from <see cref="SolidWorks.Interop.swconst.swRebuildOptions_e"/>
    /// </summary>
    /// <remarks>
    /// <para>
    ///     Options are bit-masked (flags) so you can specify one or more:
    ///     
    ///     <code>
    ///         model.Extension.Rebuild(options: RebuildOptions.Silent | RebuildOptions.AvoidRebuildOnSave);
    ///     </code>
    /// </para>
    /// </remarks>
    [Flags]
    public enum RebuildOptions
    {
        /// <summary>
        /// No rebuild is done
        /// </summary>
        None = 0,

        /// <summary>
        /// Assembly or drawing
        /// Rebuilds geometry that has not been regenerated
        /// </summary>
        RebuildAll = swRebuildOptions_e.swRebuildAll,

        /// <summary>
        /// Assembly or drawing
        /// Forces a rebuild of all geometry
        /// </summary>
        ForceRebuildAll = swRebuildOptions_e.swForceRebuildAll,

        /// <summary>
        /// Assembly only
        /// Only rebuilds mates, which is much faster than rebuilding the geometry. 
        /// Especially useful for IComponent2::Transform2
        /// </summary>
        UpdateMates = swRebuildOptions_e.swUpdateMates,

        /// <summary>
        /// Drawing only
        /// Only rebuilds the display of the views on the current drawing sheet
        /// </summary>
        CurrentSheetDisp = swRebuildOptions_e.swCurrentSheetDisp,

        /// <summary>
        /// Drawing only
        /// Only rebuilds drawing views that are dirty when OR'd with swCurrentSheetDisp option
        /// </summary>
        UpdateDirtyOnly = swRebuildOptions_e.swUpdateDirtyOnly,
    }
}
