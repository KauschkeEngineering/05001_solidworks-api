using SolidWorks.Interop.swconst;
using System;

namespace AngelSix.SolidDna
{
    /// <summary>
    /// File save errors.
    /// <see cref="swRenameDocumentError_e"/>
    /// Any errors of a document rename operation. 
    /// </summary>
    [Flags]
    public enum RenameDocumentError
    {
        /// <summary>
        /// Success
        /// </summary>
        None = swRenameDocumentError_e.swRenameDocumentError_None,

        /// <summary>
        /// You cannot rename the component due to an internal error
        /// </summary>
        UnspecifiedInternalError = swRenameDocumentError_e.swRenameDocumentError_UnspecifiedInternalError,

        /// <summary>
        /// You must select a valid component; your selection is invalid for renaming
        /// </summary>
        InvalidSelection = swRenameDocumentError_e.swRenameDocumentError_InvalidSelection,

        /// <summary>
        /// You cannot rename drawing documents
        /// </summary>
        InvalidForDrawings = swRenameDocumentError_e.swRenameDocumentError_InvalidForDrawings,

        /// <summary>
        /// You cannot rename the component because a model is not loaded in memory
        /// </summary>
        NoModelLoaded = swRenameDocumentError_e.swRenameDocumentError_NoModelLoaded,

        /// <summary>
        /// You must resolve the component before attempting to rename it
        /// </summary>
        ComponentNotResolved = swRenameDocumentError_e.swRenameDocumentError_ComponentNotResolved,

        /// <summary>
        /// You cannot rename the child component whose parent component is lightweight
        /// </summary>
        FileSaveFormatNotAvailable = swRenameDocumentError_e.swRenameDocumentError_LightWeightComponent,

        /// <summary>
        /// You cannot rename routing components
        /// </summary>
        RoutingComponent = swRenameDocumentError_e.swRenameDocumentError_RoutingComponent,

        /// <summary>
        /// You cannot rename the component to the specified name because a file with that name exists on disk
        /// </summary>
        FileAlreadyExists = swRenameDocumentError_e.swRenameDocumentError_FileAlreadyExists,

        /// <summary>
        /// You specified an invalid component name; the name is either too long or contains invalid characters
        /// </summary>
        InvalidCharactersInName = swRenameDocumentError_e.swRenameDocumentError_InvalidCharactersInName,

        /// <summary>
        /// You cannot rename virtual components
        /// </summary>
        InvalidVirtualComponent = swRenameDocumentError_e.swRenameDocumentError_InvalidVirtualComponent,

        /// <summary>
        /// You cannot rename the component to the specified name because that name is too long
        /// </summary>
        NameTooLong = swRenameDocumentError_e.swRenameDocumentError_NameTooLong,

        /// <summary>
        /// You cannot rename the component to the specified name because a document with that name is open
        /// </summary>
        DocumentNameInUse = swRenameDocumentError_e.swRenameDocumentError_DocumentNameInUse,

        /// <summary>
        /// You cannot rename the component to the specified name because a document with the same name has been temporarily renamed but not yet saved
        /// </summary>
        PendingNameAlreadyInUse = swRenameDocumentError_e.swRenameDocumentError_PendingNameAlreadyInUse,

        /// <summary>
        /// You cannot rename read-only documents or documents referenced by read-only documents
        /// </summary>
        ReadOnlyDocument = swRenameDocumentError_e.swRenameDocumentError_ReadOnlyDocument,

        /// <summary>
        /// You cannot rename the document, and the document was not saved
        /// </summary>
        DocumentNotSaved = swRenameDocumentError_e.swRenameDocumentError_DocumentNotSaved,

        /// <summary>
        /// You cannot rename virtual components
        /// </summary>
        VirtualComponent = swRenameDocumentError_e.swRenameDocumentError_VirtualComponent,

        /// <summary>
        /// You cannot rename the component because the SOLIDWORKS Professional PDM add-in is not loaded
        /// </summary>
        NotAllowedWithPDM = swRenameDocumentError_e.swRenameDocumentError_NotAllowedWithPDM,

        /// <summary>
        /// You cannot rename Toolbox components
        /// </summary>
        ToolboxComponent = swRenameDocumentError_e.swRenameDocumentError_ToolboxComponent,

        /// <summary>
        /// You cannot rename patterned components
        /// </summary>
        PatternedComponent = swRenameDocumentError_e.swRenameDocumentError_PatternedComponent
    }
}
