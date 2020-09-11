using SolidWorks.Interop.swconst;

namespace AngelSix.SolidDna
{
    /// <summary>
    /// Status of each document intended for Pack and Go.
    /// The options from a <see cref="Model.PackAndGo(string)"/> call,
    /// from <see cref="swPackAndGoSaveStatus_e"/>
    /// </summary>
    public enum PackAndGoSaveStatus
    {
        /// <summary>
        /// Successfully saved
        /// </summary>
        Success = swPackAndGoSaveStatus_e.swPackAndGoSaveStatus_Succeed,

        /// <summary>
        /// User input was not correct
        /// </summary>
        UserInputNotCorrect = swPackAndGoSaveStatus_e.swPackAndGoSaveStatus_UserInputNotCorrect,

        /// <summary>
        /// File already exists
        /// </summary>
        FileAlreadyExists = swPackAndGoSaveStatus_e.swPackAndGoSaveStatus_FileAlreadyExist,

        /// <summary>
        /// Saving an empty file
        /// </summary>
        SavingEmptyFile = swPackAndGoSaveStatus_e.swPackAndGoSaveStatus_SaveToEmpty,

        /// <summary>
        /// Error when saving
        /// </summary>
        SaveError = swPackAndGoSaveStatus_e.swPackAndGoSaveStatus_SaveError
    }
}
