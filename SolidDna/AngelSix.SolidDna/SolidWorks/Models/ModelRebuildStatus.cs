using SolidWorks.Interop.swconst;
using System;

namespace AngelSix.SolidDna
{
    /// <summary>
    /// Model rebuild status, from <see cref="SolidWorks.Interop.swconst.swModelRebuildStatus_e"/>
    /// </summary>
    [Flags]
    public enum ModelRebuildStatus
    {
        /// <summary>
        /// Model is fully rebuild
        /// </summary>
        FullyRebuilt = swModelRebuildStatus_e.swModelRebuildStatus_FullyRebuilt,

        /// <summary>
        /// Model is not fully rebuild so rebuild is necessary
        /// </summary>
        NonFrozenFeatureNeedsRebuild = swModelRebuildStatus_e.swModelRebuildStatus_NonFrozenFeatureNeedsRebuild,

        /// <summary>
        /// Model is not fully rebuild so rebuild is necessary
        /// </summary>
        FrozenFeatureNeedsRebuild = swModelRebuildStatus_e.swModelRebuildStatus_FrozenFeatureNeedsRebuild,


    }
}
