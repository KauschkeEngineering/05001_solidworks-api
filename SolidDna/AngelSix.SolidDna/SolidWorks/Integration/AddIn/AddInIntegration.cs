using Dna;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Win32;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swpublished;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using DevelopmentFramework.Logging;
using System.Runtime.InteropServices.ComTypes;

namespace AngelSix.SolidDna
{
    public enum SWProgIdVersion
    {
        UNKNOWN = -1,
        SW_2016 = 24,
        SW_2017,
        SW_2018,
        SW_2019,
        SW_2020,
        SW_2021,
        SW_2022,
        SW_2023
    }

    /// <summary>
    /// Integrates into SolidWorks as an add-in and registers for callbacks provided by SolidWorks
    /// 
    /// IMPORTANT: The class that overrides <see cref="ISwAddin"/> MUST be the same class that 
    /// contains the ComRegister and ComUnregister functions due to how SolidWorks loads add-ins
    /// </summary>
    public abstract class AddInIntegration : ISwAddin
    {
        private const uint SOLIDWORKS_APP_SUPPORTED_START_YEAR = 2016;
        private const uint SOLIDWORKS_APP_SUPPORTED_END_YEAR = 2023;

        public const string SOLIDWORKS_PROCESS_NAME = "SLDWORKS";

        private static Process _solidWorksProcess = null;

        #region Protected Members

        /// <summary>
        /// Flag if we have loaded into memory (as ConnectedToSolidWorks can happen multiple times if unloaded/reloaded)
        /// </summary>
        protected static bool mLoaded = false;

        #endregion

        #region Public Properties

        /// <summary>
        /// The title displayed for this SolidWorks Add-in
        /// </summary>
        public static string SolidWorksAddInTitle { get; set; } = "AngelSix SolidDna AddIn";

        /// <summary>
        /// The description displayed for this SolidWorks Add-in
        /// </summary>
        public static string SolidWorksAddInDescription { get; set; } = "All your pixels are belong to us!";

        /// <summary>
        /// Represents the current SolidWorks application
        /// </summary>
        public static SolidWorksApplication SolidWorks { get; set; }

        /// <summary>
        /// Gets the list of all known reference assemblies in this solution

        /// <summary>
        /// If true, loads the plug-ins in their own app-domain
        /// NOTE: Must be set before connecting to SolidWorks
        /// </summary>
        public bool DetachedAppDomain { get; set; }

        public static Process SolidWorksProcess => _solidWorksProcess;

        #endregion

        #region Public Events

        /// <summary>
        /// Called once SolidWorks has loaded our add-in and is ready.
        /// Now is a good time to create taskpanes, menu bars or anything else.
        ///  
        /// NOTE: This call will be made twice, one in the default domain and one in the AppDomain as the SolidDna plug-ins
        /// </summary>
        public static event Action ConnectedToSolidWorks = () => { };

        /// <summary>
        /// Called once SolidWorks has unloaded our add-in.
        /// Now is a good time to clean up taskpanes, menu bars or anything else.
        /// 
        /// NOTE: This call will be made twice, one in the default domain and one in the AppDomain as the SolidDna plug-ins
        /// </summary>
        public static event Action DisconnectedFromSolidWorks = () => { };

        #endregion

        #region Constructor

        /// <summary>
        /// Default constructor
        /// </summary>
        /// <param name="standAlone">
        ///     If true, sets the SolidWorks Application to the active instance
        ///     (if available) so the environment can be used from a stand alone application.
        /// </param>
        public AddInIntegration(bool standAlone = false)
        {
            try
            {
                // Get the path to this actual add-in dll
                var assemblyFilePath = this.AssemblyFilePath();
                var assemblyPath = this.AssemblyPath();

                // Setup IoC
                IoC.Setup(assemblyFilePath, construction =>
                {
                    //  Add SolidDna-specific services
                    // --------------------------------

                    // Add reference to the add-in integration
                    // Which can then be fetched anywhere with
                    // IoC.AddIn
                    construction.Services.AddSingleton(this);

                    // Add localization manager
                    construction.AddLocalizationManager();

                    //  Configure any services this class wants to add
                    // ------------------------------------------------
                    ConfigureServices(construction);
                });

                // Log details
                Logger.LogDebug($"DI Setup complete");
                Logger.LogDebug($"Assembly File Path {assemblyFilePath}");
                Logger.LogDebug($"Assembly Path {assemblyPath}");

                // If we are in stand-alone mode...
                if (standAlone)
                    // Connect to active SolidWorks
                    ConnectToActiveSolidWork();
            }
            catch (Exception ex)
            {
                // Fall-back just write a static log directly
                Logger.LogException("Exception in AddInIntegration constructor: ", ex);
                //File.AppendAllText(Path.ChangeExtension(this.AssemblyFilePath(), "fatal.log.txt"), $"\r\nUnexpected error: {ex}");
            }
        }

        #endregion

        #region Public Abstract / Virtual Methods

        /// <summary>
        /// Specific application startup code when SolidWorks is connected 
        /// and before any plug-ins or listeners are informed
        /// 
        /// NOTE: This call will not be in the same AppDomain as the SolidDna plug-ins
        /// </summary>
        /// <returns></returns>
        public abstract void ApplicationStartup();

        /// <summary>
        /// Run immediately when <see cref="ConnectToSW(object, int)"/> is called
        /// to do any pre-setup such as <see cref="PlugInIntegration.UseDetachedAppDomain"/>
        /// </summary>
        public abstract void PreConnectToSolidWorks();

        /// <summary>
        /// Run before loading plug-ins.
        /// This call should be used to add plug-ins to be loaded, via <see cref="PlugInIntegration.AddPlugIn{T}"/>
        /// </summary>
        /// <returns></returns>
        public abstract void PreLoadPlugIns();

        /// <summary>
        /// The method to implement and flag with <see cref="ConfigureServiceAttribute"/>
        /// and a custom name if you want this method to be called during IoC build
        /// </summary>
        /// <param name="construction">The IoC framework construction</param>
        [ConfigureService]
        public virtual void ConfigureServices(FrameworkConstruction construction)
        {
            // Add reference to the add-in integration
            // Which can then be fetched anywhere with
            // IoC.AddIn
            construction.Services.AddSingleton(this);

            // Add localization manager
            Framework.Construction.AddLocalizationManager();
        }

        #endregion

        #region SolidWorks Add-in Callbacks

        /// <summary>
        /// Used to pass a callback message onto our plug-ins
        /// </summary>
        /// <param name="arg"></param>
        public void Callback(string arg)
        {
            // Log it
            Logger.LogDebug($"SolidWorks Callback fired {arg}");

            PlugInIntegration.OnCallback(arg);
        }

        /// <summary>
        /// Called when SolidWorks has loaded our add-in and wants us to do our connection logic
        /// </summary>
        /// <param name="thisSw">The current SolidWorks instance</param>
        /// <param name="cookie">The current SolidWorks cookie Id</param>
        /// <returns></returns>
        public bool ConnectToSW(object thisSw, int cookie)
        {
            try
            {
                // Fire event
                PreConnectToSolidWorks();

                // Setup application (allowing for AppDomain boundary setup)
                AppDomainBoundary.Setup(this.AssemblyPath(), this.AssemblyFilePath(),
                    // The type of this abstract class will be the class implementing it
                    GetType().Assembly.Location, "");

                // Log it
                Logger.LogTrace($"Fired PreConnectToSolidWorks...");

                // Get the directory path to this actual add-in dll
                var assemblyPath = this.AssemblyPath();

                // Log it
                Logger.LogDebug($"{SolidWorksAddInTitle} Connected to SolidWorks...");

                //
                //   NOTE: Do not need to create it here, as we now create it inside PlugInIntegration.Setup in it's own AppDomain
                //         If we change back to loading directly (not in an app domain) then uncomment this 
                //
                // Store a reference to the current SolidWorks instance
                // Initialize SolidWorks (SolidDNA class)
                //SolidWorks = new SolidWorksApplication((SldWorks)ThisSW, Cookie);

                // Log it
                Logger.LogDebug($"Setting AddinCallbackInfo...");

                // Setup callback info
                var ok = ((SldWorks)thisSw).SetAddinCallbackInfo2(0, this, cookie);

                // Log it
                Logger.LogDebug($"PlugInIntegration Setup...");

                // Setup plug-in application domain
                PlugInIntegration.Setup(assemblyPath, ((SldWorks)thisSw).RevisionNumber(), cookie);

                // Log it
                Logger.LogDebug($"Firing PreLoadPlugIns...");

                // If this is the first load, or we are not loading add-ins 
                // into this domain they need loading every time as they were
                // fully unloaded on disconnect
                if (!mLoaded || AppDomainBoundary.UseDetachedAppDomain)
                {
                    // Any pre-load steps
                    PreLoadPlugIns();

                    // Log it
                    Logger.LogDebug($"Configuring PlugIns...");

                    // Perform any plug-in configuration
                    PlugInIntegration.ConfigurePlugIns(assemblyPath);

                    // Now loaded so don't do it again
                    mLoaded = true;
                }

                // Log it
                Logger.LogDebug($"Firing ApplicationStartup...");

                // Call the application startup function for an entry point to the application
                ApplicationStartup();

                // Log it
                Logger.LogDebug($"Firing ConnectedToSolidWorks...");

                // Inform listeners
                ConnectedToSolidWorks();

                // Log it
                Logger.LogDebug($"PlugInIntegration ConnectedToSolidWorks...");

                // And plug-in domain listeners
                PlugInIntegration.ConnectedToSolidWorks();

                _solidWorksProcess = Process.GetCurrentProcess();
                _solidWorksProcess.PriorityClass = ProcessPriorityClass.High;

                // Return ok
                return true;
            }
            catch (Exception ex)
            {
                // Log it
                Logger.LogCritical($"Unexpected error: {ex}");

                return false;
            }
        }

        /// <summary>
        /// Called when SolidWorks is about to unload our add-in and wants us to do our disconnection logic
        /// </summary>
        /// <returns></returns>
        public bool DisconnectFromSW()
        {
            // Log it
            Logger.LogDebug($"{SolidWorksAddInTitle} Disconnected from SolidWorks...");

            // Log it
            Logger.LogDebug($"Firing DisconnectedFromSolidWorks...");

            // Inform listeners
            DisconnectedFromSolidWorks();

            // And plug-in domain listeners
            PlugInIntegration.DisconnectedFromSolidWorks();

            // Log it
            Logger.LogDebug($"Tearing down...");

            // Clean up plug-in app domain
            PlugInIntegration.Teardown();

            // Cleanup ourselves
            TearDown();

            // Unload our domain
            AppDomainBoundary.Unload();

            // Return ok
            return true;
        }

        #endregion

        #region Connected to SolidWorks Event Calls

        /// <summary>
        /// When the add-in has connected to SolidWorks
        /// </summary>
        public static void OnConnectedToSolidWorks()
        {
            // Log it
            Logger.LogDebug($"Firing ConnectedToSolidWorks event...");

            ConnectedToSolidWorks();
        }

        /// <summary>
        /// When the add-in has disconnected to SolidWorks
        /// </summary>
        public static void OnDisconnectedFromSolidWorks()
        {
            // Log it
            Logger.LogDebug($"Firing DisconnectedFromSolidWorks event...");

            DisconnectedFromSolidWorks();
        }

        #endregion

        #region Com Registration

        /// <summary>
        /// The COM registration call to add our registry entries to the SolidWorks add-in registry
        /// </summary>
        /// <param name="t"></param>
        [ComRegisterFunction()]
        protected static void ComRegister(Type t)
        {
            // Create new instance of ComRegister add-in to setup DI
            new ComRegisterAddInIntegration();

            try
            {
                // Get assembly name
                var assemblyName = t.Assembly.Location;

                // Log it
                Logger.LogInformation($"Registering {assemblyName}");

                // Get registry key path
                var keyPath = string.Format(@"SOFTWARE\SolidWorks\AddIns\{0:b}", t.GUID);

                // Create our registry folder for the add-in
                using (var rk = Microsoft.Win32.Registry.LocalMachine.CreateSubKey(keyPath))
                {
                    // Load add-in when SolidWorks opens
                    rk.SetValue(null, 1);

                    //
                    // IMPORTANT: 
                    //
                    //   In this special case, COM register won't load the wrong AngelSix.SolidDna.dll file 
                    //   as it isn't loading multiple instances and keeping them in memory
                    //            
                    //   So loading the path of the AngelSix.SolidDna.dll file that should be in the same
                    //   folder as the add-in dll right now will work fine to get the add-in path
                    //
                    var pluginPath = typeof(PlugInIntegration).CodeBaseNormalized();

                    // Force auto-discovering plug-in during COM registration
                    PlugInIntegration.AutoDiscoverPlugins = true;

                    Logger.LogInformation("Configuring plugins...");

                    // Let plug-ins configure title and descriptions
                    PlugInIntegration.ConfigurePlugIns(pluginPath);

                    // Set SolidWorks add-in title and description
                    rk.SetValue("Title", SolidWorksAddInTitle);
                    rk.SetValue("Description", SolidWorksAddInDescription);

                    Logger.LogInformation($"COM Registration successful. '{SolidWorksAddInTitle}' : '{SolidWorksAddInDescription}'");
                }
            }
            catch (Exception ex)
            {
                Logger.LogCritical($"COM Registration error. {ex}");
                throw;
            }
        }

        /// <summary>
        /// The COM unregister call to remove our custom entries we added in the COM register function
        /// </summary>
        /// <param name="t"></param>
        [ComUnregisterFunction()]
        protected static void ComUnregister(Type t)
        {
            // Get registry key path
            var keyPath = string.Format(@"SOFTWARE\SolidWorks\AddIns\{0:b}", t.GUID);

            // Remove our registry entry
            Microsoft.Win32.Registry.LocalMachine.DeleteSubKeyTree(keyPath);

        }

        #endregion

        #region Stand Alone Methods

        /// <summary>
        /// Attempts to set the SolidWorks property to the active SolidWorks instance
        /// </summary>
        /// <returns></returns>
        public bool ConnectToActiveSolidWork()
        {
            try
            {
                // Clean up old one
                TearDown();

                // Try and get the active SolidWorks instance
                SolidWorks = new SolidWorksApplication((SldWorks)Marshal.GetActiveObject("SldWorks.Application"), 0);

                // Log it
                Logger.LogDebug($"Acquired active instance SolidWorks in Stand-Alone mode");

                // Return if successful
                return SolidWorks != null;
            }
            // If we failed to get active instance...
            catch (COMException)
            {
                // Log it
                Logger.LogDebug($"Failed to get active instance of SolidWorks in Stand-Alone mode");

                // Return failure
                return false;
            }
        }

        /// <summary>
        /// Attempts to set the SolidWorks property to the active SolidWorks instance
        /// </summary>
        /// <returns></returns>
        public bool ConnectToActiveSolidWork(SWProgIdVersion progIdVersion)
        {
            try
            {
                // Clean up old one
                TearDown();

                // Try and get the active SolidWorks instance

                var sw_ProgId = string.Format("SldWorks.Application.{0}", (int)progIdVersion);
                // Log it
                Logger.LogDebug($"Aquired active instance SolidWorks in Stand-Alone mode");

                // Try and get the active SolidWorks instance
                var obj = Marshal.GetActiveObject(sw_ProgId);
                Logger.LogDebug($"Get GetActiveObject finish");
                _solidWorksProcess = Process.GetCurrentProcess();
                SolidWorks = new SolidWorksApplication((SldWorks)obj, 0);
                Logger.LogDebug($"finish with quired active instance SolidWorks in Stand-Alone mode");

                // Return if successful
                return SolidWorks != null;
            }
            // If we failed to get active instance...
            catch (COMException ex)
            {
                Logger.LogException($"COM error while connecting to active SOLIDWORKS with progid: {progIdVersion}", ex);
                return false;
            }
        }

        /// <summary>
        /// Attempts to set the SolidWorks property to the active SolidWorks instance.
        /// Remember to call <see cref="TearDown"/> once done.
        /// </summary>
        /// <returns></returns>
        public static bool ConnectToActiveSolidWorks()
        {
            // Create new blank add-in
            var addIn = new BlankAddInIntegration();

            // Return if we successfully got an instance
            return addIn.ConnectToActiveSolidWork();
        }

        /// <summary>
        /// Attempts to set the SolidWorks property to the active SolidWorks instance.
        /// Remember to call <see cref="TearDown"/> once done.
        /// </summary>
        /// <returns></returns>
        public static bool ConnectToActiveSolidWorks(SWProgIdVersion progIdVersion)
        {
            // Create new blank add-in
            var addIn = new BlankAddInIntegration();

            // Return if we successfully got an instance
            return addIn.ConnectToActiveSolidWork(progIdVersion);
        }

        public bool StartSolidWorkProcess(string solidWorksExePath)
        {
            if (!solidWorksExePath.Equals(string.Empty))
            {
                try
                {
                    var processInfo = new ProcessStartInfo()
                    {
                        FileName = solidWorksExePath,
                        Arguments = "/r", //no splash screen will be shown while loading SolidWorks application
                        CreateNoWindow = false,
                        WindowStyle = ProcessWindowStyle.Normal
                    };
                    _solidWorksProcess = Process.Start(processInfo);
                    // set the priorty to high for SOLIDWORKS to gain more CPU time
                    _solidWorksProcess.PriorityClass = ProcessPriorityClass.High;
                }
                catch (Exception ex)
                {
                    Logger.LogException($"Error while starting SOLIDWORKS process with exe path: {solidWorksExePath}", ex);
                }
            }

            return true;
        }

        public static bool StartSolidWorksProcess(string solidWorksExePath)
        {
            var addIn = new BlankAddInIntegration();
            return addIn.StartSolidWorkProcess(solidWorksExePath);
        }

        public static List<Tuple<SWProgIdVersion, string>> GetInstalledSolidWorksVersionExePaths()
        {
            var installedSolidWorksVersions = new List<Tuple<SWProgIdVersion, string>>();

            for (var version = SOLIDWORKS_APP_SUPPORTED_START_YEAR; version <= SOLIDWORKS_APP_SUPPORTED_END_YEAR; version++)
            {
                using (var baseRegistryKey = RegistryKey.OpenBaseKey(RegistryHive.LocalMachine, RegistryView.Registry64).OpenSubKey(string.Format("Software\\SolidWorks\\SOLIDWORKS {0}\\Setup", version)))
                {
                    if (baseRegistryKey != null)
                    {
                        var registryValue = baseRegistryKey.GetValue("SolidWorks Folder");
                        if (registryValue != null)
                        {
                            switch (version)
                            {
                                case 2016:
                                    installedSolidWorksVersions.Add(new Tuple<SWProgIdVersion, string>(SWProgIdVersion.SW_2016, registryValue.ToString()));
                                    break;
                                case 2017:
                                    installedSolidWorksVersions.Add(new Tuple<SWProgIdVersion, string>(SWProgIdVersion.SW_2017, registryValue.ToString()));
                                    break;
                                case 2018:
                                    installedSolidWorksVersions.Add(new Tuple<SWProgIdVersion, string>(SWProgIdVersion.SW_2018, registryValue.ToString()));
                                    break;
                                case 2019:
                                    installedSolidWorksVersions.Add(new Tuple<SWProgIdVersion, string>(SWProgIdVersion.SW_2019, registryValue.ToString()));
                                    break;
                                case 2020:
                                    installedSolidWorksVersions.Add(new Tuple<SWProgIdVersion, string>(SWProgIdVersion.SW_2020, registryValue.ToString()));
                                    break;
                                case 2021:
                                    installedSolidWorksVersions.Add(new Tuple<SWProgIdVersion, string>(SWProgIdVersion.SW_2021, registryValue.ToString()));
                                    break;
                                case 2022:
                                    installedSolidWorksVersions.Add(new Tuple<SWProgIdVersion, string>(SWProgIdVersion.SW_2022, registryValue.ToString()));
                                    break;
                                case 2023:
                                    installedSolidWorksVersions.Add(new Tuple<SWProgIdVersion, string>(SWProgIdVersion.SW_2023, registryValue.ToString()));
                                    break;
                            }
                        }
                    }
                }
            }
            return installedSolidWorksVersions;
        }

        public static Tuple<SWProgIdVersion, string> GetInstalledSolidWorksVersionExePath(SWProgIdVersion sWProgIdVersion)
        {
            var version = string.Empty;
            switch (sWProgIdVersion)
            {
                case SWProgIdVersion.SW_2016:
                    version = "2016";
                    break;
                case SWProgIdVersion.SW_2017:
                    version = "2017";
                    break;
                case SWProgIdVersion.SW_2018:
                    version = "2018";
                    break;
                case SWProgIdVersion.SW_2019:
                    version = "2019";
                    break;
                case SWProgIdVersion.SW_2020:
                    version = "2020";
                    break;
                case SWProgIdVersion.SW_2021:
                    version = "2021";
                    break;
                case SWProgIdVersion.SW_2022:
                    version = "2022";
                    break;
                case SWProgIdVersion.SW_2023:
                    version = "2023";
                    break;
            }
            if (version.Equals("") == false)
            {
                using (var baseRegistryKey = RegistryKey.OpenBaseKey(RegistryHive.LocalMachine, RegistryView.Registry64).OpenSubKey(string.Format("Software\\SolidWorks\\SOLIDWORKS {0}\\Setup", version)))
                {
                    if (baseRegistryKey != null)
                    {
                        var registryValue = baseRegistryKey.GetValue("SolidWorks Folder");
                        if (registryValue != null)
                        {
                            switch (sWProgIdVersion)
                            {
                                case SWProgIdVersion.SW_2016:
                                    return new Tuple<SWProgIdVersion, string>(SWProgIdVersion.SW_2016, registryValue.ToString());
                                case SWProgIdVersion.SW_2017:
                                    return new Tuple<SWProgIdVersion, string>(SWProgIdVersion.SW_2017, registryValue.ToString());
                                case SWProgIdVersion.SW_2018:
                                    return new Tuple<SWProgIdVersion, string>(SWProgIdVersion.SW_2018, registryValue.ToString());
                                case SWProgIdVersion.SW_2019:
                                    return new Tuple<SWProgIdVersion, string>(SWProgIdVersion.SW_2019, registryValue.ToString());
                                case SWProgIdVersion.SW_2020:
                                    return new Tuple<SWProgIdVersion, string>(SWProgIdVersion.SW_2021, registryValue.ToString());
                                case SWProgIdVersion.SW_2021:
                                    return new Tuple<SWProgIdVersion, string>(SWProgIdVersion.SW_2021, registryValue.ToString());
                                case SWProgIdVersion.SW_2022:
                                    return new Tuple<SWProgIdVersion, string>(SWProgIdVersion.SW_2022, registryValue.ToString());
                                case SWProgIdVersion.SW_2023:
                                    return new Tuple<SWProgIdVersion, string>(SWProgIdVersion.SW_2023, registryValue.ToString());
                            }
                        }
                    }
                }
            }
            return null;
        }

        public static List<SWProgIdVersion> GetRunningSolidWorksVersionProgIds()
        {
            var runningSolidWorksVersionProgIds = new List<SWProgIdVersion>();
            foreach (var progIdVersion in Enum.GetValues(typeof(SWProgIdVersion)))
            {
                try
                {
                    var solidWorksInstance = new SolidWorksApplication((SldWorks)Marshal.GetActiveObject(string.Format("SldWorks.Application.{0}", (int)progIdVersion)), 0);
                    runningSolidWorksVersionProgIds.Add((SWProgIdVersion)progIdVersion);
                    // If we have an reference...
                    if (solidWorksInstance != null)
                    {

                        // Dispose SolidWorks COM
                        solidWorksInstance?.Dispose();
                    }

                    // Set to null
                    solidWorksInstance = null;
                }
                catch (COMException ex)
                {
                    Logger.LogException($"COM error while getting running SOLIDWORKS version prog ids.", ex);
                }
            }
            return runningSolidWorksVersionProgIds;
        }

        public static bool IsSolidWorksApplicationRunning(SWProgIdVersion sWProgIdVersion)
        {
            try
            {
                var solidWorksInstance = new SolidWorksApplication((SldWorks)Marshal.GetActiveObject(string.Format("SldWorks.Application.{0}", (int)sWProgIdVersion)), 0);
                // if there is already an running app
                if (solidWorksInstance != null)
                {
                    // Dispose SolidWorks COM
                    solidWorksInstance?.Dispose();
                    solidWorksInstance = null;
                    return true;
                }
            }
            catch (COMException ex)
            {
                // if there is no SOLIDWORKS application is available this error will occur
                Logger.LogException($"COM error while checking if SOLIDWORKS application is running.", ex);
            }
            return false;
        }

        public static void LoadActiveSolidWorksProcess()
        {
            // get the current process of SolidWorks
            if (_solidWorksProcess == null)
            {
                _solidWorksProcess = Process.GetProcessesByName(SOLIDWORKS_PROCESS_NAME).FirstOrDefault();
            }
        }

        public static bool KillHangSolidWorksProcess()
        {
            // get the current process of SolidWorks
            if (_solidWorksProcess != null)
            {
                _solidWorksProcess.Kill();
                _solidWorksProcess.Dispose();
                _solidWorksProcess = null;
                return true;
            }
            return false;
        }
        #endregion

        #region Tear Down

        /// <summary>
        /// Cleans up the SolidWorks instance
        /// </summary>
        public static void TearDown()
        {
            Logger.LogDebug($"Disposing SolidWorks COM reference...");
            // Dispose SolidWorks COM
            SolidWorks?.Dispose();
            SolidWorksProcess?.Dispose();

            // Set to null
            SolidWorks = null;
            _solidWorksProcess = null;
        }
        #endregion



        public static bool GetAssemblyNameFromProcess(int processId, ref string assemblyName)
        {
            SldWorks app = GetSwAppFromProcess(processId);


            foreach (var doc in (object[])app.GetDocuments())
            {
                ModelDoc2 modelDoc = (ModelDoc2)doc;
                if (modelDoc != null)
                {
                    var model = new Model((ModelDoc2)modelDoc);
                    if (model.IsAssembly)
                    {
                        app.ActivateDoc(model.Name);
                        break;
                    }

                }
            }
            ModelDoc2 activeModelDoc = (ModelDoc2)app.ActiveDoc;
            if (activeModelDoc != null)
            {
                assemblyName = activeModelDoc.GetPathName();
                Logger.LogDebug($"Found solidworks app with process id: " + processId + " and file name: " + assemblyName);

                if (assemblyName.ToLower().EndsWith(DrawingDocument.FILE_EXTENSION))
                {
                    Logger.Log(LogLevel.WARN, "Can not use active solidworks process. It is a drawing and not a assembly");
                    return false;
                }

                return true;
            }
            else
            {
                Logger.LogDebug($"Dead SolidWorks process!");
                return false;
            }
        }

        [DllImport("ole32.dll")]
        private static extern int CreateBindCtx(uint reserved, out IBindCtx ppbc);

        private static SldWorks GetSwAppFromProcess(int processId)
        {
            var monikerName = "SolidWorks_PID_" + processId.ToString();
            IBindCtx context = null;
            IRunningObjectTable rot = null;
            IEnumMoniker monikers = null;

            try
            {
                CreateBindCtx(0, out context);

                context.GetRunningObjectTable(out rot);
                rot.EnumRunning(out monikers);

                var moniker = new IMoniker[1];
                IntPtr pNumFetched = new IntPtr();

                while (monikers.Next(1, moniker, pNumFetched) == 0)
                {
                    var curMoniker = moniker.First();

                    string name = null;


                    if (curMoniker != null)
                    {
                        try
                        {
                            curMoniker.GetDisplayName(context, null, out name);
                        }
                        catch (UnauthorizedAccessException ex)
                        {
                            Logger.LogException($"Unauthorized access error while getting SOLIDWORKS application from process with pid: {processId}", ex);
                        }
                    }

                    if (string.Equals(monikerName, name, StringComparison.CurrentCultureIgnoreCase))
                    {
                        //Logger.LogDebugSource($"Found correct com object");
                        rot.GetObject(curMoniker, out var com_instance);
                        SldWorks app = com_instance as SldWorks;
                        return app;
                    }
                }
            }
            catch (Exception ex)
            {
                Logger.LogException($"Error while getting SOLIDWORKS application from process with pid: {processId}", ex);
            }
            return null;
        }

        public static bool ConnectToSwAppFromProcessId(int processId)
        {
            SldWorks app = GetSwAppFromProcess(processId);
            if (app != null)
            {
                SolidWorks = new SolidWorksApplication(app, 0);
                _solidWorksProcess = Process.GetProcessById(processId);
            }

            if (SolidWorks != null)
            {
                return true;
            }
            else
            {
                return false;
            }
        }

        public static void ActivateBoundSolidWorks(Process solidWorksProcess, SolidWorksApplication solidWorksApplication)
        {
            _solidWorksProcess = solidWorksProcess;
            SolidWorks = solidWorksApplication;
        }
    }
}
