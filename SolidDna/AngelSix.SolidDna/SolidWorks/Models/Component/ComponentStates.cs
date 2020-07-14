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

    /// <summary>
    /// States for resolving components.
    /// <see cref="swComponentResolveStatus_e"
    /// </summary>
    [Flags]
    public enum ComponentResolveStates
    {
        // components resolved okay
        ResolveOk = swComponentResolveStatus_e.swResolveOk,
        // user aborted resolving the components
        ResolveAbortedByUser = swComponentResolveStatus_e.swResolveAbortedByUser,
        // some of the components did not get resolved despite the user requesting it
        ResolveNotPerformed = swComponentResolveStatus_e.swResolveNotPerformed,
        // not used
        ResolveError = swComponentResolveStatus_e.swResolveError,
    }
}
