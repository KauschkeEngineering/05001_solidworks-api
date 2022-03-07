using SolidWorks.Interop.swdocumentmgr;
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
                _instance = new Application(Credential.getSolidWorksLicenseAPIKey());
            }
            return _instance;
        }

        public SwDmDocumentType GetDocumentType(string filePath)
        {
            // Determine type of SOLIDWORKS file based on file extension
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

            if (_instance._documentManagerApplication != null)
            {
                var _swDoc = _instance._documentManagerApplication.GetDocument(filePath, GetDocumentType(filePath), true, out var _retVal);
                if (_swDoc != null)
                {
                    return (MajorSolidWorksVersions)_swDoc.GetVersion();
                }
            }
            return MajorSolidWorksVersions.UNKNOWN;
        }

        public string[] GetDocumentReferences(string documentFileName)
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
            return null;
        }

    }
}
