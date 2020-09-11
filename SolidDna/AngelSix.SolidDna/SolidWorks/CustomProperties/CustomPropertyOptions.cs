using SolidWorks.Interop.swconst;

namespace AngelSix.SolidDna
{
    /// <summary>
    /// Options when adding custom properties.
    /// <see cref="swCustomPropertyAddOption_e"/>
    /// </summary>
    public enum CustomPropertyAddOption
    {
        // add the custom property only if it is new
        OnlyIfNew = swCustomPropertyAddOption_e.swCustomPropertyOnlyIfNew,
        // delete an existing custom property having the same name and add the new custom property
        DeleteAndAdd = swCustomPropertyAddOption_e.swCustomPropertyDeleteAndAdd,
        // replace the value of an existing custom property having the same name
        ReplaceValue = swCustomPropertyAddOption_e.swCustomPropertyReplaceValue
    }

}
