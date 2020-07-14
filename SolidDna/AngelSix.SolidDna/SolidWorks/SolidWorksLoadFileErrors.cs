using SolidWorks.Interop.swconst;

namespace AngelSix.SolidDna
{
    public enum FileLoadErrors
    {
        // another error was encountered
        GenericError = swFileLoadError_e.swGenericError,
        // unable to locate the file; the file is not loaded or the referenced file (that is, component) is suppressed
        FileNotFoundError = swFileLoadError_e.swFileNotFoundError,
        // obsolete see swFileLoadWarning_e
        IdMatchError = swFileLoadError_e.swIdMatchError,
        // obsolete see swFileLoadWarning_e
        ReadOnlyWarn = swFileLoadError_e.swReadOnlyWarn,
        // obsolete see swFileLoadWarning_e
        SharingViolationWarn = swFileLoadError_e.swSharingViolationWarn,
        // obsolete see swFileLoadWarning_e
        DrawingANSIUpdateWarn = swFileLoadError_e.swDrawingANSIUpdateWarn,
        // obsolete see swFileLoadWarning_e
        SheetScaleUpdateWarn = swFileLoadError_e.swSheetScaleUpdateWarn,
        // obsolete see swFileLoadWarning_e
        NeedsRegenWarn = swFileLoadError_e.swNeedsRegenWarn,
        // obsolete see swFileLoadWarning_e
        BasePartNotLoadedWarn = swFileLoadError_e.swBasePartNotLoadedWarn,
        // obsolete see swFileLoadWarning_e
        FileAlreadyOpenWarn = swFileLoadError_e.swFileAlreadyOpenWarn,
        // file type argument is not valid
        InvalidFileTypeError = swFileLoadError_e.swInvalidFileTypeError,
        // obsolete see swFileLoadWarning_e
        DrawingsOnlyRapidDraftWarn = swFileLoadError_e.swDrawingsOnlyRapidDraftWarn,
        // obsolete see swFileLoadWarning_e
        ViewOnlyRestrictions = swFileLoadError_e.swViewOnlyRestrictions,
        // the document was saved in a future version of SOLIDWORKS
        FutureVersionError = swFileLoadError_e.swFutureVersion,
        // obsolete see swFileLoadWarning_e
        ViewMissingReferencedConfig = swFileLoadError_e.swViewMissingReferencedConfig,
        // obsolete see swFileLoadWarning_e
        DrawingSFSymbolConvertWarn = swFileLoadError_e.swDrawingSFSymbolConvertWarn,
        // a document with the same name is already open
        FileWithSameTitleAlreadyOpenError = swFileLoadError_e.swFileWithSameTitleAlreadyOpen,
        // file encrypted by Liquid Machines
        LiquidMachineDocError = swFileLoadError_e.swLiquidMachineDoc,
        // file is open and blocked because the system memory is low, or the number of GDI handles has exceeded the allowed maximum
        LowResourcesError = swFileLoadError_e.swLowResourcesError,
        // file contains no display data
        NoDisplayDataError = swFileLoadError_e.swNoDisplayData,
        // the user attempted to open a file, and then interrupted the open-file routine to open a different file
        AddinInteruptError = swFileLoadError_e.swAddinInteruptError,
        // a document has non-critical custom property data corruption
        FileRequiresRepairError = swFileLoadError_e.swFileRequiresRepairError,
        // a document has critical data corruption
        FileCriticalDataRepairError = swFileLoadError_e.swFileCriticalDataRepairError,
        // the application was to busy to load the document
        ApplicationBusyError = swFileLoadError_e.swApplicationBusy
    }
}
