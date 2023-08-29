using DevelopmentFramework.Logging;
using Serilog;
using SolidWorks.Interop.swdocumentmgr;
using System.Linq;
using static AngelSix.SolidDna.Model;

namespace AngelSix.SolidDna.DocumentManager
{
    public sealed class Application
    {
        private static Application _instance = null;
        private static readonly SwDMClassFactory _classFactory = new SwDMClassFactory();
        private readonly SwDMApplication _documentManagerApplication;


        private Application(string licKey)
        {
            _documentManagerApplication = _classFactory.GetApplication(licKey);
        }

        public static Application GetInstance()
        {
            if (_instance == null)
            {
                _instance = new Application(Credential.GetSolidWorksLicenseAPIKey());
            }
            return _instance;
        }

        public SwDmDocumentType GetDocumentType(string filePath)
        {
            // Determine type of SOLIDWORKS file based on file extension
            if (filePath == null || string.Equals(filePath, string.Empty))
            {
                return SwDmDocumentType.swDmDocumentUnknown;
            }
            if (filePath.ToLower().EndsWith(PartDocument.FILE_EXTENSION))
            {
                return SwDmDocumentType.swDmDocumentPart;
            }
            else if (filePath.ToLower().EndsWith(AssemblyDocument.FILE_EXTENSION))
            {
                return SwDmDocumentType.swDmDocumentAssembly;
            }
            else if (filePath.ToLower().EndsWith(DrawingDocument.FILE_EXTENSION))
            {
                return SwDmDocumentType.swDmDocumentDrawing;
            }
            else
            {
                return SwDmDocumentType.swDmDocumentUnknown;
            }
        }

        public MajorSolidWorksVersions GetFileSWVersion(string filePath)
        {

            if (_documentManagerApplication != null)
            {
                var _swDoc = _documentManagerApplication.GetDocument(filePath, GetDocumentType(filePath), true, out var _retVal);
                if (_swDoc != null)
                {
                    return (MajorSolidWorksVersions)_swDoc.GetVersion();
                }
            }
            return MajorSolidWorksVersions.UNKNOWN;
        }

        public string[] GetDocumentReferences(string documentFileName)
        {
            if (_documentManagerApplication != null)
            {
                int _docType;
                switch (GetDocumentType(documentFileName))
                {
                    case SwDmDocumentType.swDmDocumentUnknown:
                        return null;
                    case SwDmDocumentType.swDmDocumentPart:
                        _docType = (int)SwDmDocumentType.swDmDocumentPart;
                        break;
                    case SwDmDocumentType.swDmDocumentAssembly:
                        _docType = (int)SwDmDocumentType.swDmDocumentAssembly;
                        break;
                    case SwDmDocumentType.swDmDocumentDrawing:
                        _docType = (int)SwDmDocumentType.swDmDocumentDrawing;
                        break;
                    default:
                        return null;
                }

                var swSearchOpt = _documentManagerApplication.GetSearchOptionObject();
                SwDmDocumentOpenError nRetVal;
                var swDoc = (SwDMDocumentClass)_documentManagerApplication.GetDocument(documentFileName, (SwDmDocumentType)_docType, true, out nRetVal);
                if (swDoc != null && nRetVal == SwDmDocumentOpenError.swDmDocumentOpenErrorNone)
                    return (string[])swDoc.GetAllExternalReferences(swSearchOpt);
            }
            return null;
        }


        public bool IsTopParentAssembly(string assemblyFilePath, string searchPath)
        {
            if (_documentManagerApplication != null)
            {
                var swSearchOpt = _documentManagerApplication.GetSearchOptionObject();
                swSearchOpt.ClearAllSearchPaths();
                swSearchOpt.AddSearchPath(searchPath);
                var filter = swSearchOpt.SearchFilters;
                swSearchOpt.SearchFilters = (int)(SwDmSearchFilters.SwDmSearchExternalReference | SwDmSearchFilters.SwDmSearchForAssembly);
                var nRetVal = SwDmDocumentOpenError.swDmDocumentOpenErrorNone;

                var swDoc = (SwDMDocumentClass)_documentManagerApplication.GetDocument(assemblyFilePath, SwDmDocumentType.swDmDocumentAssembly, true, out nRetVal);
                if (swDoc != null)
                {
                    var references = swDoc.WhereUsed(swSearchOpt);
                    if (references == null)
                    {
                        swDoc.CloseDoc();
                        return true;
                    }
                }
            }
            return false;
        }

        public object GetPreviewBitmap(string filePath, bool isDrawingSheet = false, string drawingSheetName = "")
        {
            Logger.LogDebug($"Start getting preview image of file path: {filePath} and drawing sheet name: {drawingSheetName}");
            if (_documentManagerApplication != null)
            {
                var version = (MajorSolidWorksVersions)_documentManagerApplication.GetLatestSupportedFileVersion();
                Logger.LogDebug($"Latest supported version for getting preview image: {version}");
                SwDmDocumentOpenError nRetVal;
                var solidWorksDocument = (SwDMDocument12)_documentManagerApplication.GetDocument(filePath, Application.GetInstance().GetDocumentType(filePath), true, out nRetVal);
                if (solidWorksDocument != null)
                {
                    Logger.LogDebug($"Latest supported version for getting preview image: {version}");
                    object previewImage = null;
                    if (isDrawingSheet)
                    {
                        var drawingSheets = (object[])solidWorksDocument.GetSheets();
                        if (drawingSheets != null)
                        {
                            foreach (var drawingSheet in drawingSheets)
                            {
                                if (((SwDMSheet2)drawingSheet).Name.Equals(drawingSheetName))
                                {
                                    previewImage = ((SwDMSheet2)drawingSheet).GetPreviewPNGBitmap(out var error);
                                }
                            }
                        }
                        Logger.LogDebug($"Successfully created preview image for drawing sheet");
                        return previewImage;
                    }
                    else
                    {
                        previewImage = solidWorksDocument.GetPreviewBitmap(out var error);
                        Logger.LogDebug($"Successfully created preview image");
                        return previewImage;
                    }
                }
                else
                {
                    Logger.LogWaring($"Could not get SOLIDWORKS document of document manager open error: {nRetVal}");
                }
            }
            // SwDMDocument10::GetPreviewBitmap throws an unmanaged COM exception 
            // for out-of-process C# console applications
            // Use the following code in SOLIDWORKS C# macros and add-ins  
            return null;
        }

    }
}
