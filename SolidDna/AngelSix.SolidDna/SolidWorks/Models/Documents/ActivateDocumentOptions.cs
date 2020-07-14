using SolidWorks.Interop.swconst;

namespace AngelSix.SolidDna
{
    /// <summary>
    /// Rebuild options during document activation.
    /// </summary>
    public enum RebuildOnActivationOptions
    {
        // prompt the user whether to rebuild the activated document
        UserDecision = swRebuildOnActivation_e.swUserDecision,
        // do not rebuild the activated document
        DontRebuildActiveDoc = swRebuildOnActivation_e.swDontRebuildActiveDoc,
        // rebuild the activated document
        RebuildActiveDoc = swRebuildOnActivation_e.swRebuildActiveDoc
    }
}