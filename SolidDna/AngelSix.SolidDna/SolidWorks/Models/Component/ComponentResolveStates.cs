using SolidWorks.Interop.swconst;
using System;

namespace AngelSix.SolidDna
{
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
        ResolveError = swComponentResolveStatus_e.swResolveError
    }
}
