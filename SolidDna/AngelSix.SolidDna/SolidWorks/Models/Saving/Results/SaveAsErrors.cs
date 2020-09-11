using SolidWorks.Interop.swconst;
using System;

namespace AngelSix.SolidDna
{
    /// <summary>
    /// File save errors.
    /// <see cref="swFileSaveError_e"/>
    /// Any errors of a model save operation. 
    /// </summary>
    /// <remarks>
    /// <para>
    ///     Errors are bit-masked (flags) so you can get each error via:
    ///     
    ///     <code>
    ///         errors.GetFlags();
    ///     </code>
    /// </para>
    /// </remarks>
    [Flags]
    public enum SaveAsErrors
    {
        /// <summary>
        /// No errors
        /// </summary>
        None = 0,

        /// <summary>
        /// There was an unknown error
        /// </summary>
        GenericSaveError = swFileSaveError_e.swGenericSaveError,

        /// <summary>
        /// Failed to save as the destination file is read-only
        /// </summary>
        ReadOnlySaveError = swFileSaveError_e.swReadOnlySaveError,

        /// <summary>
        /// The filename was not provided
        /// </summary>
        FileNameEmpty = swFileSaveError_e.swFileNameEmpty,

        /// <summary>
        /// Filename cannot contain the @ symbol
        /// </summary>
        FileNameContainsAtSign = swFileSaveError_e.swFileNameContainsAtSign,

        /// <summary>
        /// The file is write-locked
        /// </summary>
        FileLockError = swFileSaveError_e.swFileLockError,

        /// <summary>
        /// The save as file type is not valid
        /// </summary>
        FileSaveFormatNotAvailable = swFileSaveError_e.swFileSaveFormatNotAvailable,

        /// <summary>
        /// Obsolete: This error is now in a warning
        /// </summary>
        [Obsolete("This error is now a warning")]
        FileSaveWithRebuildError = swFileSaveError_e.swFileSaveWithRebuildError,

        /// <summary>
        /// The file already exists, and the save is set to not override existing files
        /// </summary>
        FileSaveAsDoNotOverwrite = swFileSaveError_e.swFileSaveAsDoNotOverwrite,

        /// <summary>
        /// The file save extension does not match the SolidWorks document type
        /// </summary>
        FileSaveAsInvalidFileExtension = swFileSaveError_e.swFileSaveAsInvalidFileExtension,

        /// <summary>
        /// Save the selected bodies in a part document. Valid option for IPartDoc::SaveToFile2; 
        /// however, not a valid option for IModelDocExtension::SaveAs
        /// </summary>
        FileSaveAsNoSelection = swFileSaveError_e.swFileSaveAsNoSelection,

        /// <summary>
        /// The version of eDrawings is invalid
        /// </summary>
        FileSaveAsBadEDrawingsVersion = swFileSaveError_e.swFileSaveAsBadEDrawingsVersion,

        /// <summary>
        /// The filename is too long
        /// </summary>
        FileSaveAsNameExceedsMaxPathLength = swFileSaveError_e.swFileSaveAsNameExceedsMaxPathLength,

        /// <summary>
        /// The save as operation is not supported, or was executed is such a way that the resulting
        /// file might not be complete, possibly because SolidWorks is hidden; 
        /// </summary>
        FileSaveAsNotSupported = swFileSaveError_e.swFileSaveAsNotSupported,

        /// <summary>
        /// Saving an assembly with renamed components requires saving the references
        /// </summary>
        FileSaveRequiresSavingReferences = swFileSaveError_e.swFileSaveRequiresSavingReferences
    }
}
