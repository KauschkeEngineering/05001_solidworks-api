using SolidWorks.Interop.swconst;
using System;

namespace AngelSix.SolidDna
{
    /// <summary>
    /// States for component suppression.
    /// <see cref="swComponentSuppressionState_e"
    /// </summary>
    [Flags]
    public enum ComponentSuppressionStates
    {
        // fully suppressed - recursively suppresses the component and any child components
        Suppressed = swComponentSuppressionState_e.swComponentSuppressed,
        // lightweight - makes only the component lightweight
        Lightweight = swComponentSuppressionState_e.swComponentLightweight,
        // fully resolved - recursively resolves the component and any child components
        FullyResolved = swComponentSuppressionState_e.swComponentFullyResolved,
        // resolved - resolves only the component
        Resolved = swComponentSuppressionState_e.swComponentResolved,
        // fully lightweight - recursively makes the component and any child components lightweight
        FullyLightweight = swComponentSuppressionState_e.swComponentFullyLightweight,
        // a internal mismatch of the component id
        InternalIdMismatch = swComponentSuppressionState_e.swComponentInternalIdMismatch
    }
}
