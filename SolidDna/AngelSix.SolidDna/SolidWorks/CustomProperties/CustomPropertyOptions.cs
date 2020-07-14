using SolidWorks.Interop.swconst;

namespace AngelSix.SolidDna
{

    public enum CustomPropertyAddOption
    {
        // Add the custom property only if it is new
        OnlyIfNew = swCustomPropertyAddOption_e.swCustomPropertyOnlyIfNew,
        // Delete an existing custom property having the same name and add the new custom property
        DeleteAndAdd = swCustomPropertyAddOption_e.swCustomPropertyDeleteAndAdd,
        // Replace the value of an existing custom property having the same name
        ReplaceValue = swCustomPropertyAddOption_e.swCustomPropertyReplaceValue
    }

}
