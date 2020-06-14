using System;
using SolidWorks.Interop.sldworks;
using System.Threading.Tasks;
using System.Collections.Generic;
using System.Linq;
using System.IO;
using System.Xml.Linq;
using SolidWorks.Interop.swconst;
using SolidWorks.Interop.swdocumentmgr;

namespace AngelSix.SolidDna
{
    /// <summary>
    /// Represents the current SolidWorks application
    /// </summary>
    public partial class SolidWorksApplication : SharedSolidDnaObject<SldWorks>
    {
        public enum RebuildOnActivation
        {
            UserDecision = swRebuildOnActivation_e.swUserDecision,
            DontRebuildActiveDoc = swRebuildOnActivation_e.swDontRebuildActiveDoc,
            RebuildActiveDoc = swRebuildOnActivation_e.swRebuildActiveDoc
        }


        private const string API_LICENCE_KEY = "KauschkeEngineeringServicesGmbH:swdocmgr_general-11785-02051-00064-01025-08567-34307-00007-16944-03729-52630-20621-27155-52442-46742-57351-01653-28064-37004-12470-18657-42155-23021-34092-52693-41357-38317-47381-42397-38329-51605-47525-19869-51605-42457-38285-07629-35253-00289-24632-23148-57538-23268-24676-24676-7,swdocmgr_previews-11785-02051-00064-01025-08567-34307-00007-23424-25180-19561-08976-33418-56614-15908-28672-33039-34039-15128-63516-16802-33026-22959-34092-52693-41357-38317-47381-42397-38329-51605-47525-19869-51605-42457-38285-07629-35253-00289-24632-23148-57538-23268-24676-24676-4,swdocmgr_dimxpert-11785-02051-00064-01025-08567-34307-00007-02224-60566-22735-30039-32017-34229-26945-40960-13408-24595-25589-46377-44335-06006-23917-34092-52693-41357-38317-47381-42397-38329-51605-47525-19869-51605-42457-38285-07629-35253-00289-24632-23148-57538-23268-24676-24676-4,swdocmgr_geometry-11785-02051-00064-01025-08567-34307-00007-25440-44169-22552-63994-59551-37072-09667-14340-27239-01502-21198-57986-04466-22801-22651-34092-52693-41357-38317-47381-42397-38329-51605-47525-19869-51605-42457-38285-07629-35253-00289-24632-23148-57538-23268-24676-24676-5,swdocmgr_xml-11785-02051-00064-01025-08567-34307-00007-65328-31737-49269-00552-43417-32429-25145-30724-20516-24804-09579-38046-62867-26541-24480-34092-52693-41357-38317-47381-42397-38329-51605-47525-19869-51605-42457-38285-07629-35253-00289-24632-23148-57538-23268-24676-24676-6,swdocmgr_tessellation-11785-02051-00064-01025-08567-34307-00007-48088-40139-18107-01106-37523-23509-29212-14337-40415-34512-32340-56451-08411-09207-22831-34092-52693-41357-38317-47381-42397-38329-51605-47525-19869-51605-42457-38285-07629-35253-00289-24632-23148-57538-23268-24676-24676-2";
        //Old 2015 - private const string API_LICENCE_KEY = "KauschkeEngineeringServicesGmbH:swdocmgr_general-11785-02051-00064-17409-08473-34307-00007-19656-13610-17716-24516-52581-13763-62678-28678-03781-27226-07511-14366-10812-12747-23344-34092-52693-41357-38317-47381-42397-38329-51605-47525-19869-51605-42457-38285-07629-35253-00289-25656-23152-57052-23276-24676-27746-1,swdocmgr_previews-11785-02051-00064-17409-08473-34307-00007-51024-52025-35553-51281-51480-35643-63443-14342-31683-18932-27645-24746-19372-11348-23297-34092-52693-41357-38317-47381-42397-38329-51605-47525-19869-51605-42457-38285-07629-35253-00289-25656-23152-57052-23276-24676-27746-1,swdocmgr_xml-11785-02051-00064-17409-08473-34307-00007-13784-50962-42612-58001-63569-50093-44301-40960-12153-36851-02018-22860-03363-63670-23674-34092-52693-41357-38317-47381-42397-38329-51605-47525-19869-51605-42457-38285-07629-35253-00289-25656-23152-57052-23276-24676-27746-5";

        #region Protected Members

        /// <summary>
        /// The cookie to the current instance of SolidWorks we are running inside of
        /// </summary>
        protected int mSwCookie;

        /// <summary>
        /// The file path of the current file that is loading. 
        /// Used to ignore active document changed events during opening of a file
        /// </summary>
        protected string mFileLoading;

        /// <summary>
        /// The currently active document
        /// </summary>
        protected Model mActiveModel;

        #endregion

        #region Private Members

        /// <summary>
        /// Locking object for synchronizing the disposing of SolidWorks and reloading active model info.
        /// </summary>
        private readonly object mDisposingLock = new object();

        #endregion

        #region Public Properties

        /// <summary>
        /// The currently active model
        /// </summary>
        public Model ActiveModel => mActiveModel;

        /// <summary>
        /// Various preferences for SolidWorks
        /// </summary>
        public SolidWorksPreferences Preferences { get; protected set; }

        /// <summary>
        /// Gets the current SolidWorks version information
        /// </summary>
        public SolidWorksVersion SolidWorksVersion => GetSolidWorksVersion();

        /// <summary>
        /// The SolidWorks instance cookie
        /// </summary>
        public int SolidWorksCookie => mSwCookie;

        /// <summary>
        /// The command manager
        /// </summary>
        public CommandManager CommandManager { get; }

        /// <summary>
        /// True if the application is disposing
        /// </summary>
        public bool Disposing { get; private set; }

        public bool IsVisible => BaseObject.Visible;

        #endregion

        #region Public Events

        /// <summary>
        /// Called when any information about the currently active model has changed
        /// </summary>
        public event Action<Model> ActiveModelInformationChanged = (model) => { };

        /// <summary>
        /// Called when a new file has been created
        /// </summary>
        public event Action<Model> FileCreated = (model) => { };

        /// <summary>
        /// Called when a file has been opened
        /// </summary>
        public event Action<string, Model> FileOpened = (path, model) => { };

        /// <summary>
        /// Called when the currently active file has been saved
        /// </summary>
        public event Action<string, Model> ActiveFileSaved = (path, model) => { };

        #endregion

        #region Constructor

        /// <summary>
        /// Default constructor
        /// </summary>
        public SolidWorksApplication(SldWorks solidWorks, int cookie) : base(solidWorks)
        {
            // Set preferences
            Preferences = new SolidWorksPreferences();
            var loadExternal = BaseObject.GetUserPreferenceIntegerValue((int)swUserPreferenceIntegerValue_e.swLoadExternalReferences);
            //BaseObject.SetUserPreferenceIntegerValue((int)swUserPreferenceIntegerValue_e.swLoadExternalReferences, (int)swLoadExternalReferences_e.swLoadExternalReferences_None);

            // Store cookie Id
            mSwCookie = cookie;
            //
            //   NOTE: As we are in our own AppDomain, the callback is registered in the main SolidWorks AppDomain
            //         We then pass that into our domain
            //
            // Setup callback info
            // var ok = BaseObject.SetAddinCallbackInfo2(0, this, cookie);

            // Hook into main events
            BaseObject.ActiveModelDocChangeNotify += ActiveModelChanged;
            BaseObject.FileOpenPreNotify += FileOpenPreNotify;
            BaseObject.FileOpenPostNotify += FileOpenPostNotify;
            BaseObject.FileNewNotify2 += FileNewPostNotify;
            BaseObject.OnIdleNotify += OnIdleIsLoadedNotify;

            // If we have a cookie...
            if (cookie > 0)
                // Get command manager
                CommandManager = new CommandManager(UnsafeObject.GetCommandManager(mSwCookie));

            // Get whatever the current model is on load
            ReloadActiveModelInformation();
        }

        #endregion

        private int OnIdleIsLoadedNotify()
        {
            return 0;
        }

        #region Public Callback Events

        /// <summary>
        /// Informs this class that the active model may have changed and it should be reloaded
        /// </summary>
        public void RequestActiveModelChanged()
        {
            ReloadActiveModelInformation();
        }

        #endregion

        #region Version

        /// <summary>
        /// Gets the current SolidWorks version information
        /// </summary>
        /// <returns></returns>
        protected SolidWorksVersion GetSolidWorksVersion()
        {
            // Wrap any error
            return SolidDnaErrors.Wrap(() =>
            {
                // Get version string (such as 23.2.0 for 2015 SP2.0)
                var revisionNumber = BaseObject.RevisionNumber();

                // Get revision string (such as sw2015_SP20)
                // Get build number (such as d150130.002)
                // Get the hot fix string
                BaseObject.GetBuildNumbers2(out var revisionString, out var buildNumber, out var hotfixString);

                return new SolidWorksVersion(revisionNumber, revisionString, buildNumber, hotfixString);
            },
                SolidDnaErrorTypeCode.SolidWorksApplication,
                SolidDnaErrorCode.SolidWorksApplicationVersionError,
                Localization.GetString("SolidWorksApplicationVersionError"));
        }

        #endregion

        #region SolidWorks Event Methods

        #region File New

        /// <summary>
        /// Called after a new file has been created.
        /// <see cref="ActiveModel"/> is updated to the new file before this event is called.
        /// </summary>
        /// <param name="newDocument"></param>
        /// <param name="documentType"></param>
        /// <param name="templatePath"></param>
        /// <returns></returns>
        private int FileNewPostNotify(object newDocument, int documentType, string templatePath)
        {
            // Inform listeners
            FileCreated(mActiveModel);

            // NOTE: 0 is OK, anything else is an error
            return 0;
        }

        #endregion

        #region File Open

        /// <summary>
        /// Called after a file has finished opening
        /// </summary>
        /// <param name="filename">The filename to the file being opened</param>
        /// <returns></returns>
        private int FileOpenPostNotify(string filename)
        {
            // Wrap any error
            SolidDnaErrors.Wrap(() =>
            {
                // If this is the file we were opening...
                if (string.Equals(filename, mFileLoading, StringComparison.OrdinalIgnoreCase))
                {
                    // File has been loaded, so clear loading flag
                    mFileLoading = null;

                    // And update all properties and models
                    // TODO: Check how to reload properly to get sync into both directions
                    ReloadActiveModelInformation();

                    // Inform listeners
                    FileOpened(filename, mActiveModel);
                }

            },
                SolidDnaErrorTypeCode.SolidWorksApplication,
                SolidDnaErrorCode.SolidWorksApplicationFilePostOpenError,
                Localization.GetString("SolidWorksApplicationFilePostOpenError"));

            // NOTE: 0 is OK, anything else is an error
            return 0;
        }

        /// <summary>
        /// Called before a file has started opening
        /// </summary>
        /// <param name="filename">The filename to the file being opened</param>
        /// <returns></returns>
        private int FileOpenPreNotify(string filename)
        {
            // Don't handle the ActiveModelDocChangeNotify event for file open events
            // - wait until the file is open instead

            // NOTE: We need to check if the variable already has a value because in the case of a drawing
            // we get multiple pre-events - one for the drawing, and one for each model in it,
            // we're only interested in the first

            // Wrap any error
            SolidDnaErrors.Wrap(() =>
            {
                if (mFileLoading == null)
                    mFileLoading = filename;
            },
                SolidDnaErrorTypeCode.SolidWorksApplication,
                SolidDnaErrorCode.SolidWorksApplicationFilePreOpenError,
                Localization.GetString("SolidWorksApplicationFilePreOpenError"));

            // NOTE: 0 is OK, anything else is an error
            return 0;
        }

        #endregion

        #region Model Changed

        /// <summary>
        /// Called when the active model has changed
        /// </summary>
        /// <returns></returns>
        private int ActiveModelChanged()
        {
            // Wrap any error
            SolidDnaErrors.Wrap(() =>
            {
                // If we are currently loading a file...
                if (mFileLoading != null)
                {
                    // Check the active document
                    using (var activeDoc = new Model(BaseObject.IActiveDoc2))
                    {
                        // If this is the same file that is currently being loaded, ignore this event
                        if (string.Equals(mFileLoading, activeDoc.FilePath, StringComparison.OrdinalIgnoreCase))
                            return;
                    }
                }

                // If we got here, it isn't the current document so reload the data
                ReloadActiveModelInformation();
            },
                SolidDnaErrorTypeCode.SolidWorksApplication,
                SolidDnaErrorCode.SolidWorksApplicationActiveModelChangedError,
                Localization.GetString("SolidWorksApplicationActiveModelChangedError"));

            // NOTE: 0 is OK, anything else is an error
            return 0;
        }

        #endregion

        #endregion

        #region Active Model

        /// <summary>
        /// Reloads all of the variables, data and COM objects for the newly available SolidWorks model/state
        /// </summary>
        private void ReloadActiveModelInformation()
        {
            // First clean-up any previous SW data
            CleanActiveModelData();

            // Now get the new data
            if (BaseObject.IActiveDoc2 == null)
                mActiveModel = null;
            else
                mActiveModel = new Model(BaseObject.IActiveDoc2);

            // Listen out for events
            if (mActiveModel != null)
            {
                mActiveModel.ModelSaved += ActiveModel_Saved;
                mActiveModel.ModelInformationChanged += ActiveModel_InformationChanged;
                mActiveModel.ModelClosing += ActiveModel_Closing;
            }

            // Inform listeners
            ActiveModelInformationChanged(mActiveModel);
        }

        /// <summary>
        /// Disposes of any active model-specific data ready for refreshing
        /// </summary>
        private void CleanActiveModelData()
        {
            // Active model
            mActiveModel?.Dispose();
        }

        #region Event Callbacks

        /// <summary>
        /// Called when the active model has informed us its information has changed
        /// </summary>
        private void ActiveModel_InformationChanged()
        {
            // Inform listeners
            ActiveModelInformationChanged(mActiveModel);
        }

        /// <summary>
        /// Called when the active document is closed
        /// </summary>
        private void ActiveModel_Closing()
        {
            // 
            // NOTE: There is no event to detect when all documents are closed 
            // 
            //       So, each model that is closing (not closed) wait 200ms 
            //       then check on the current number of active documents
            //       or if ActiveDoc is already set to null.
            //
            //       If ActiveDoc is null or the document count is 0 at that 
            //       moment in time, do an active model information refresh.
            //
            //       If another document opens in the meantime it won't fire
            //       but that's fine as the active doc changed event will fire
            //       in that case anyway
            //

            // Check for every file if it may have been the last one.
            Task.Run(async () =>
            {
                // Wait for it to close
                await Task.Delay(200);

                // Lock to prevent Disposing to change while this section is running.
                lock (mDisposingLock)
                {
                    if (Disposing)
                        // If we are disposing SolidWorks, there is no need to reload active model info.
                        return;
                    // TODO: Check how to reload properly to get sync into both directions
                    // Now if we have none open, reload information
                    // ActiveDoc is quickly set to null after the last document is closed
                    // GetDocumentCount takes longer to go to zero for big assemblies, but it might be a more reliable indicator.
                    if (BaseObject?.ActiveDoc == null || BaseObject?.GetDocumentCount() == 0)
                        ReloadActiveModelInformation();

                }
            });
        }

        /// <summary>
        /// Called when the currently active file has been saved
        /// </summary>
        private void ActiveModel_Saved()
        {
            // Inform listeners
            ActiveFileSaved(mActiveModel?.FilePath, mActiveModel);
        }

        #endregion

        #endregion

        #region Save Data

        /// <summary>
        /// Gets an <see cref="IExportPdfData"/> object for use with a <see cref="PdfExportData"/>
        /// object used in <see cref="Model.SaveAs(string, SaveAsVersion, SaveAsOptions, PdfExportData)"/> call
        /// </summary>
        /// <returns></returns>
        public IExportPdfData GetPdfExportData()
        {
            // NOTE: No point making our own enumerator for the export file type
            //       as right now and for many years it's only ever been
            //       1 for PDF. I do not see this ever changing
            return BaseObject.GetExportFileData((int)swExportDataFileType_e.swExportPdfData) as IExportPdfData;
        }

        #endregion

        #region Materials

        /// <summary>
        /// Gets a list of all materials in SolidWorks
        /// </summary>
        /// <param name="database">If specified, limits the results to the specified database full path</param>
        public List<Material> GetMaterials(string database = null)
        {
            // Wrap any error
            return SolidDnaErrors.Wrap(() =>
            {
                // Create an empty list
                var list = new List<Material>();

                // If we are using a specified database, use that
                if (database != null)
                    ReadMaterials(database, ref list);
                else
                {
                    // Otherwise, get all known ones
                    // Get the list of material databases (full paths to SLDMAT files)
                    var databases = (string[])BaseObject.GetMaterialDatabases();

                    // Get materials from each
                    if (databases != null)
                        foreach (var d in databases)
                            ReadMaterials(d, ref list);
                }

                // Order the list
                return list.OrderBy(f => f.DisplayName).ToList();
            },
                SolidDnaErrorTypeCode.SolidWorksApplication,
                SolidDnaErrorCode.SolidWorksApplicationGetMaterialsError,
                Localization.GetString("SolidWorksApplicationGetMaterialsError"));
        }

        /// <summary>
        /// Attempts to find the material from a SolidWorks material database file (SLDMAT)
        /// If found, returns the full information about the material
        /// </summary>
        /// <param name="database">The full path to the database</param>
        /// <param name="materialName">The material name to find</param>
        /// <returns></returns>
        public Material FindMaterial(string database, string materialName)
        {
            // Wrap any error
            return SolidDnaErrors.Wrap(() =>
            {
                // Get all materials from the database
                var materials = GetMaterials(database);

                // Return if found the material with the same name
                return materials?.FirstOrDefault(f => string.Equals(f.Name, materialName, StringComparison.InvariantCultureIgnoreCase));
            },
                SolidDnaErrorTypeCode.SolidWorksApplication,
                SolidDnaErrorCode.SolidWorksApplicationFindMaterialsError,
                Localization.GetString("SolidWorksApplicationFindMaterialsError"));
        }

        #region Private Helpers

        /// <summary>
        /// Reads the material database and adds the materials to the given list
        /// </summary>
        /// <param name="database">The database to read</param>
        /// <param name="list">The list to add materials to</param>
        private static void ReadMaterials(string database, ref List<Material> list)
        {
            // First make sure the file exists
            if (!File.Exists(database))
                throw new SolidDnaException(
                    SolidDnaErrors.CreateError(
                        SolidDnaErrorTypeCode.SolidWorksApplication,
                        SolidDnaErrorCode.SolidWorksApplicationGetMaterialsFileNotFoundError,
                        Localization.GetString("SolidWorksApplicationGetMaterialsFileNotFoundError")));

            try
            {
                // File should be an XML document, so attempt to read that
                using (var stream = File.Open(database, FileMode.Open, FileAccess.Read, FileShare.Read))
                {
                    // Try and parse the Xml
                    var xmlDoc = XDocument.Load(stream);

                    var materials = new List<Material>();

                    // Iterate all classification nodes and inside are the materials
                    xmlDoc.Root.Elements("classification")?.ToList()?.ForEach(f =>
                    {
                        // Get classification name
                        var classification = f.Attribute("name")?.Value;

                        // Iterate all materials
                        f.Elements("material").ToList().ForEach(material =>
                        {
                            // Add them to the list
                            materials.Add(new Material
                            {
                                Database = database,
                                DatabaseFileFound = true,
                                Classification = classification,
                                Name = material.Attribute("name")?.Value,
                                Description = material.Attribute("description")?.Value,
                            });
                        });
                    });

                    // If we found any materials, add them
                    if (materials.Count > 0)
                        list.AddRange(materials);
                }
            }
            catch (Exception ex)
            {
                // If we crashed for any reason during parsing, wrap in SolidDna exception
                if (!File.Exists(database))
                    throw new SolidDnaException(
                        SolidDnaErrors.CreateError(
                            SolidDnaErrorTypeCode.SolidWorksApplication,
                            SolidDnaErrorCode.SolidWorksApplicationGetMaterialsFileFormatError,
                            Localization.GetString("SolidWorksApplicationGetMaterialsFileFormatError"),
                            ex));
            }
        }

        #endregion

        #endregion

        #region Preferences

        /// <summary>
        /// Gets the specified user preference value
        /// </summary>
        /// <param name="preference">The preference to get</param>
        /// <returns></returns>
        public double GetUserPreferencesDouble(swUserPreferenceDoubleValue_e preference)
        {
            return BaseObject.GetUserPreferenceDoubleValue((int)preference);
        }

        #endregion

        #region Taskpane Methods

        /// <summary>
        /// Attempts to create 
        /// </summary>
        /// <param name="iconPath">An absolute path to an icon to use for the taskpane (ideally 37x37px)</param>
        /// <param name="toolTip">The title text to show at the top of the taskpane</param>
        public async Task<Taskpane> CreateTaskpaneAsync(string iconPath, string toolTip)
        {
            // Wrap any error creating the taskpane in a SolidDna exception
            return SolidDnaErrors.Wrap<Taskpane>(() =>
            {
                // Attempt to create the taskpane
                var comTaskpane = BaseObject.CreateTaskpaneView2(iconPath, toolTip);

                // If we fail, return null
                if (comTaskpane == null)
                    return null;

                // If we succeed, create SolidDna object
                return new Taskpane(comTaskpane);
            },
                SolidDnaErrorTypeCode.SolidWorksTaskpane,
                SolidDnaErrorCode.SolidWorksTaskpaneCreateError,
                await Localization.GetStringAsync("ErrorSolidWorksTaskpaneCreateError"));
        }

        #endregion

        #region User Interaction

        /// <summary>
        /// Pops up a message box to the user with the given message
        /// </summary>
        /// <param name="message">The message to display to the user</param>
        /// <param name="icon">The severity icon to display</param>
        /// <param name="buttons">The buttons to display</param>
        public SolidWorksMessageBoxResult ShowMessageBox(string message, SolidWorksMessageBoxIcon icon = SolidWorksMessageBoxIcon.Information, SolidWorksMessageBoxButtons buttons = SolidWorksMessageBoxButtons.Ok)
        {
            // Send message to user
            return (SolidWorksMessageBoxResult)BaseObject.SendMsgToUser2(message, (int)icon, (int)buttons);
        }

        #endregion

        #region Dispose

        /// <summary>
        /// Disposing
        /// </summary>
        public override void Dispose()
        {
            lock (mDisposingLock)
            {

                // Flag as disposing
                Disposing = true;

                // Clean active model
                ActiveModel?.Dispose();

                // Dispose command manager
                CommandManager?.Dispose();

                // NOTE: Don't dispose the application, SolidWorks does that itself
                //base.Dispose();
            }
        }

        #endregion

        private SwDmDocumentType GetDocumentType(string filePath)
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

        public object GetPreviewBitmap(string filePath, bool isDrawingSheet = false, string drawingSheetName = "")
        {
            var swClassFact = new SwDMClassFactory();
            var swDocMgr = (SwDMApplication)swClassFact.GetApplication(API_LICENCE_KEY);
            if (swDocMgr != null)
            {
                var ver = (Model.MajorSolidWorksVersions)swDocMgr.GetLatestSupportedFileVersion();
                var swDoc = (SwDMDocument12)swDocMgr.GetDocument(filePath, GetDocumentType(filePath), true, out var nRetVal);
                if (swDoc != null)
                {
                    if (isDrawingSheet)
                    {
                        var drawingSheets = (object[])swDoc.GetSheets();
                        foreach (var drawingSheet in drawingSheets)
                        {
                            if (((SwDMSheet2)drawingSheet).Name.Equals(drawingSheetName))
                                return ((SwDMSheet2)drawingSheet).GetPreviewPNGBitmap(out var error);
                        }
                    }
                    else
                        return swDoc.GetPreviewBitmap(out var error);
                }
            }
            // SwDMDocument10::GetPreviewBitmap throws an unmanaged COM exception 
            // for out-of-process C# console applications
            // Use the following code in SOLIDWORKS C# macros and add-ins  
            return null;
        }

        public List<DrawingSheet> GetDrawingSheets(string drawingFilePath)
        {
            var drawingSheets = new List<DrawingSheet>();
            var swClassFact = new SwDMClassFactory();
            var swDocMgr = (SwDMApplication)swClassFact.GetApplication(API_LICENCE_KEY);
            if (swDocMgr != null)
            {
                var swDoc = (SwDMDocument12)swDocMgr.GetDocument(drawingFilePath, SwDmDocumentType.swDmDocumentDrawing, true, out var nRetVal);
                if (swDoc != null)
                {
                    var sheets = (object[])swDoc.GetSheets();

                    foreach (var sheet in sheets)
                    {
                        drawingSheets.Add(new DrawingSheet((SwDMSheet2)sheet));
                    }
                    return drawingSheets;
                }
            }
            return drawingSheets;
        }

        public object GetDPreviewBitmap(string drawingFilePath)
        {
            return ((SwDMSheet2)mBaseObject).GetPreviewPNGBitmap(out var error);
        }

        public void CloseDocumentWithoutSaving(string documentPath)
        {
            BaseObject.QuitDoc(documentPath);
        }

        public void CloseDocument(string documentName)
        {
            BaseObject.CloseDoc(documentName);
        }

        public int OpenDocument(string fullDocumentFilePath, bool readOnly = true, bool silent = true, bool useLightWeightDefault = true, bool loadLightWeight = false, bool ignoreHiddenComponents = true, bool loadExternalReferencesInMemory = true)
        {
            Tuple<Model, int> solidWorksModel = null;

            solidWorksModel = OpenDocumentWithSpecification(fullDocumentFilePath, readOnly, silent, useLightWeightDefault, loadLightWeight, ignoreHiddenComponents, loadExternalReferencesInMemory, false);
            BaseObject.SetCurrentWorkingDirectory(fullDocumentFilePath.Replace(fullDocumentFilePath.Split('\\').Last(), ""));

            if (solidWorksModel != null)
            {
                if (solidWorksModel.Item1 != null)
                {
                    mActiveModel = solidWorksModel.Item1;
                    mActiveModel.SetUserControlable(false, false, false, true);
                }
            }

            return solidWorksModel.Item2;
        }

        //assembly/part document needs to stay open otherwise it is not fully accessable

        private Tuple<Model, int> OpenDocumentWithSpecification(string fullDocumentFilePath, bool readOnly, bool openSilent, bool useLightWeightDefault, bool loadLightWeight, bool ignoreHiddenComponents, bool loadExternalReferencesInMemory, bool selective)
        {
            var swDocSpecification = default(DocumentSpecification);
            swDocSpecification = (DocumentSpecification)BaseObject.GetOpenDocSpec(fullDocumentFilePath);
            swDocSpecification.ReadOnly = readOnly;
            swDocSpecification.Silent = openSilent;
            swDocSpecification.UseLightWeightDefault = useLightWeightDefault;
            swDocSpecification.LightWeight = loadLightWeight;
            swDocSpecification.IgnoreHiddenComponents = ignoreHiddenComponents; 
            swDocSpecification.LoadExternalReferencesInMemory = true;
            swDocSpecification.Selective = selective;

            switch (GetDocumentType(fullDocumentFilePath))
            {
                case SwDmDocumentType.swDmDocumentPart:
                    swDocSpecification.DocumentType = (int)swDocumentTypes_e.swDocPART;
                    break;
                case SwDmDocumentType.swDmDocumentAssembly:
                    swDocSpecification.DocumentType = (int)swDocumentTypes_e.swDocASSEMBLY;
                    break;
                case SwDmDocumentType.swDmDocumentDrawing:
                    swDocSpecification.DocumentType = (int)swDocumentTypes_e.swDocDRAWING;
                    break;
            }

            // creating a document invisibly by passing false to DocumentVisible method, then it is not possible to make it visible with IModelDoc2:Visible
            // if only set to true for drawing the SOLIDWORKS session will also be terminated
            // so every document nees to be se to true
            BaseObject.DocumentVisible(true, (int)swDocSpecification.DocumentType);

            // Use KeepInvisible when SOLIDWORKS is invisible and it shall activate a component and SOLIDWORKS has to be prevented from becoming visible
            // be sure to set this property back to false after the operation for which it was to true completes
            var modelData = new Tuple<Model, int>(new Model((ModelDoc2)BaseObject.OpenDoc7(swDocSpecification)), swDocSpecification.Error);
            BaseObject.DocumentVisible(false, (int)swDocSpecification.DocumentType);
            return modelData;
        }


        public Tuple<Model, int> LoadModelInMemory(string documentPath, bool readOnly = true, bool silent = true, bool useLightWeightDefault = true, bool loadLightWeight = false, bool ignoreHiddenComponents = true, bool loadExternalReferencesInMemory = true)
        {
            Tuple<Model, int> solidWorksModel = null;

            //1. Set ISldWorks::EnableBackgroundProcessing to true.
            //2. Use ISldWorks Event BackgroundProcessingStartNotify to handle the background processing start event.
            //3. Open the drawing document by calling either ISldWorks::OpenDoc6 or ISldWorks::OpenDoc7.
            //4. Set IDrawingDoc::BackgroundProcessingOption to swBackgroundProcessOption_e.swBackgroundProcessing_DeferToApplication.
            //5. Call ISldWorks::IsBackgroundProcessingCompleted repeatedly, which polls the status of the open operation, to know when background processing ends.
            //6. Use ISldWorks Event BackgroundProcessingEndNotify to handle the background processing end event.
            //7. When the open operation is finished, set ISldWorks::EnableBackgroundProcessing to false.

            switch (GetDocumentType(documentPath))
            {
                case SwDmDocumentType.swDmDocumentDrawing:
                    // set EnableBackgroundProcessing = true to more efficiently and programmatically open a drawing document that requires a lot of CPU time and no user input
                    BaseObject.EnableBackgroundProcessing = true;
                    break;
                default:
                    break;
            }

            solidWorksModel = OpenDocumentWithSpecification(documentPath, readOnly, silent, useLightWeightDefault, loadLightWeight, ignoreHiddenComponents, loadExternalReferencesInMemory, false);

            if (solidWorksModel != null)
            {
                if (solidWorksModel.Item1 != null)
                {
                    solidWorksModel.Item1.SetUserControlable(false, false, false, true);
                    if (solidWorksModel.Item1.IsDrawing == true)
                    {
                        // TODO: Add event handlers for background processing
                        solidWorksModel.Item1.Drawing.SetBackgroundProcessingOption(DrawingDocument.BackgroundProcessOptions.BackgroundProcessingDeferToApplication);
                    }
                }
            }
            // set EnableBackgroundProcessing = false when the open operation is finished
            BaseObject.EnableBackgroundProcessing = false;
            return solidWorksModel;
        }

        public string GetDefaultAssemblyTemplatePath()
        {
            return BaseObject.GetUserPreferenceStringValue((int)swUserPreferenceStringValue_e.swDefaultTemplateAssembly);
        }

        public string GetDefaultPartTemplatePath()
        {
            return BaseObject.GetUserPreferenceStringValue((int)swUserPreferenceStringValue_e.swDefaultTemplatePart);
        }

        public string GetDefaultDrawingTemplatePath()
        {
            return BaseObject.GetUserPreferenceStringValue((int)swUserPreferenceStringValue_e.swDefaultTemplateDrawing);
        }

        public bool IsTopParentAssembly(string assemblyFilePath, string searchPath)
        {

            var swClassFact = new SwDMClassFactory();
            var swDocMgr = swClassFact.GetApplication(API_LICENCE_KEY);
            var swSearchOpt = swDocMgr.GetSearchOptionObject();
            swSearchOpt.ClearAllSearchPaths();
            swSearchOpt.AddSearchPath(searchPath);
            var nRetVal = SwDmDocumentOpenError.swDmDocumentOpenErrorNone;

            var swDoc = (SwDMDocumentClass)swDocMgr.GetDocument(assemblyFilePath, SwDmDocumentType.swDmDocumentAssembly, true, out nRetVal);
            if (swDoc != null)
            {
                var references = swDoc.WhereUsed(swSearchOpt);
                if (references == null ||
                    ((string[])references).Where(referencedFile => referencedFile.ToLower().Contains(AssemblyDocument.FILE_EXTENSION)).Count() == 0)
                {
                    return true;
                }
            }
            return false;
        }

        public string[] GetDocumentReferences(string documentFileName)
        {
            var swClassFact = default(SwDMClassFactory);
            var swDocMgr = default(SwDMApplication);
            var swDoc = default(SwDMDocumentClass);
            var swSearchOpt = default(SwDMSearchOption);
            var nDocType = 0;

            switch (GetDocumentType(documentFileName))
            {
                case SwDmDocumentType.swDmDocumentUnknown:
                    nDocType = (int)SwDmDocumentType.swDmDocumentUnknown;
                    return null;
                case SwDmDocumentType.swDmDocumentPart:
                    nDocType = (int)SwDmDocumentType.swDmDocumentPart;
                    break;
                case SwDmDocumentType.swDmDocumentAssembly:
                    nDocType = (int)SwDmDocumentType.swDmDocumentAssembly;
                    break;
                case SwDmDocumentType.swDmDocumentDrawing:
                    nDocType = (int)SwDmDocumentType.swDmDocumentDrawing;
                    break;
                default:
                    nDocType = (int)SwDmDocumentType.swDmDocumentUnknown;
                    return null;
            }

            swClassFact = new SwDMClassFactory();
            swDocMgr = (SwDMApplication)swClassFact.GetApplication(API_LICENCE_KEY);
            swSearchOpt = swDocMgr.GetSearchOptionObject();
            var nRetVal = SwDmDocumentOpenError.swDmDocumentOpenErrorNone;
            swDoc = (SwDMDocumentClass)swDocMgr.GetDocument(documentFileName, (SwDmDocumentType)nDocType, true, out nRetVal);
            if (swDoc != null)
                return (string[])swDoc.GetAllExternalReferences(swSearchOpt);
            return null;
        }

        public List<Model.MajorSolidWorksVersions> GetFileSWVersion(string fileName)
        {
            var solidWorksFileVersion = new List<Model.MajorSolidWorksVersions>();
            if (BaseObject != null)
            {
                var fileVersionHistories = (string[])BaseObject.VersionHistory(fileName);
                foreach (var fileVersionHistory in fileVersionHistories)
                {
                    var majorVersion = fileVersionHistory.Split('[')[0];
                    solidWorksFileVersion.Add((Model.MajorSolidWorksVersions)int.Parse(majorVersion));
                }
                return solidWorksFileVersion;
            }
            return solidWorksFileVersion;
        }

        // TODO: Dont forget to delete the template model after retriving the desired data
        public Model GetModelDummyByTemplate(string templateFilePath)
        {
            if (BaseObject != null)
            {
                var nDocType = -1;
                if (templateFilePath.EndsWith(".asmdot"))
                {
                    nDocType = (int)SwDmDocumentType.swDmDocumentAssembly;
                }
                else if (templateFilePath.EndsWith(".prtdot"))
                {
                    nDocType = (int)SwDmDocumentType.swDmDocumentPart;
                }
                else if (templateFilePath.EndsWith(".drwdot"))
                {
                    nDocType = (int)SwDmDocumentType.swDmDocumentDrawing;
                }

                BaseObject.DocumentVisible(true, nDocType);
                var dummyModel = new Model((ModelDoc2)BaseObject.NewDocument(templateFilePath, 0, 0, 0));
                BaseObject.DocumentVisible(false, nDocType);

                dummyModel.SetUserControlable(false, false, false, true);
                return dummyModel;
            }
            return null;
        }

        public void ExitApplication()
        {
            if (BaseObject != null)
            {
                while (HasOpenDocuments())
                {
                    BaseObject.CloseAllDocuments(true);
                }
                BaseObject.ExitApp();
            }
        }

        public bool CloseAllDoucments(bool closeUnsafed)
        {
            if (BaseObject != null)
                return BaseObject.CloseAllDocuments(closeUnsafed);
            return false;
        }

        public bool MoveFile(string sourceFilePath, string destinationFilePath)
        {
            if (BaseObject != null)
                if (BaseObject.MoveDocument(sourceFilePath, destinationFilePath, null, null, 1) == 0)
                    return true;
            return false;
        }

        public void SetUserControlable(bool controlable, bool isApplicationVisible)
        {
            if (BaseObject != null)
            {
                BaseObject.CommandInProgress = !controlable;
                BaseObject.UserControl = controlable;
                BaseObject.UserControlBackground = controlable;
                BaseObject.Visible = isApplicationVisible;
                ((IFrame)BaseObject.Frame()).KeepInvisible = !isApplicationVisible;
            }
        }

        public void SetApplicationVisible(bool isApplicationVisible)
        {
            if (BaseObject != null)
            {
                BaseObject.Visible = isApplicationVisible;
                ((IFrame)BaseObject.Frame()).KeepInvisible = !isApplicationVisible;
            }
        }

        public IntPtr GetWindowHandleX64()
        {
            if (BaseObject != null)
                return new IntPtr(((IFrame)BaseObject.Frame()).GetHWndx64());
            return new IntPtr(-1);
        }

        public void ActivateDocument(string documentName)
        {
            if (BaseObject != null)
            {
                var error = 0;
                BaseObject.ActivateDoc3(documentName, false, (int)RebuildOnActivation.RebuildActiveDoc, ref error);
            }
        }

        public bool HasOpenDocuments()
        {
            if (BaseObject != null)
            {
                if (BaseObject.GetFirstDocument() != null)
                {
                    return true;
                }
            }
            return false;
        }

        public void CheckModelWindows()
        {
            var swFrame = (Frame)BaseObject.Frame();
            var modelWindows = (object[])swFrame.ModelWindows;
            foreach (var obj in modelWindows)
            {
                var swModelWindow = (ModelWindow)obj;
                //Get the model document in this model window
                var swModelDoc = (ModelDoc2)swModelWindow.ModelDoc;
                //Rebuild the document
                swModelDoc.EditRebuild3();
                swModelDoc = null;
                //Show the model window
                swFrame.ShowModelWindow(swModelWindow);
                //Get and print the model window handle
                var HWnd = swModelWindow.HWnd;
                var title = swModelWindow.Title;
            }
        }

    }
}
