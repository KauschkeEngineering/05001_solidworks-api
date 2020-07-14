using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using static AngelSix.SolidDna.CustomPropertyEditor;

namespace AngelSix.SolidDna
{
    /// <summary>
    /// Represents a SolidWorks model of any type (Drawing, Part or Assembly)
    /// </summary>
    public class Model : SharedSolidDnaObject<ModelDoc2>
    {
        public enum MajorSolidWorksVersions
        {
            SOLIDWORKS_95 = 44,
            SOLIDWORKS_96 = 243,
            SOLIDWORKS_97 = 483,
            SOLIDWORKS_97Plus = 629,
            SOLIDWORKS_98 = 822,
            SOLIDWORKS_98Plus = 1008,
            SOLIDWORKS_99 = 1137,
            SOLIDWORKS_2000 = 1500,
            SOLIDWORKS_2001 = 1750,
            SOLIDWORKS_2001Plus = 1950,
            SOLIDWORKS_2003 = 2200,
            SOLIDWORKS_2004 = 2500,
            SOLIDWORKS_2005 = 2800,
            SOLIDWORKS_2006 = 3100,
            SOLIDWORKS_2007 = 3400,
            SOLIDWORKS_2008 = 3800,
            SOLIDWORKS_2009 = 4100,
            SOLIDWORKS_2010 = 4400,
            SOLIDWORKS_2011 = 4700,
            SOLIDWORKS_2012 = 5000,
            SOLIDWORKS_2013 = 6000,
            SOLIDWORKS_2014 = 7000,
            SOLIDWORKS_2015 = 8000,
            SOLIDWORKS_2016 = 9000,
            SOLIDWORKS_2017 = 10000,
            SOLIDWORKS_2018 = 11000,
            SOLIDWORKS_2019 = 12000,
            SOLIDWORKS_2020 = 13000
        }

        #region Public Properties

        /// <summary>
        /// The absolute file path of this model if it has been saved
        /// </summary>
        public string FilePath { get; protected set; }

        public string Name { get; protected set; }
        //{
        //    get
        //    {
        //       if (FilePath.Equals("") == false)
        //        {
        //            return FilePath.Split('\\')[FilePath.Split('\\').Length - 1];
        //        }
        //        return "";
        //    }
        //}

        /// <summary>
        /// Indicates if this file has been saved (so exists on disk).
        /// If not, it's a new model currently only in-memory and will not have a file path
        /// </summary>
        public bool HasBeenSaved => !string.IsNullOrEmpty(FilePath);

        /// <summary>
        /// The type of document such as a part, assembly or drawing
        /// </summary>
        public ModelType ModelType { get; protected set; }

        /// <summary>
        /// True if this model is a part
        /// </summary>
        public bool IsPart => ModelType == ModelType.Part;

        /// <summary>
        /// True if this model is an assembly
        /// </summary>
        public bool IsAssembly => ModelType == ModelType.Assembly;

        /// <summary>
        /// True if this model is a drawing
        /// </summary>
        public bool IsDrawing => ModelType == ModelType.Drawing;

        /// <summary>
        /// Contains extended information about the model
        /// </summary>
        public ModelExtension Extension { get; protected set; }

        /// <summary>
        /// Contains the current active configuration information
        /// </summary>
        public ModelConfiguration ActiveConfiguration { get; protected set; }

        /// <summary>
        /// The selection manager for this model
        /// </summary>
        public SelectionManager SelectionManager { get; protected set; }

        /// <summary>
        /// Get the number of configurations
        /// </summary>
        public int ConfigurationCount => BaseObject.GetConfigurationCount();

        /// <summary>
        /// Gets the configuration names
        /// </summary>
        public List<string> ConfigurationNames => new List<string>((string[])BaseObject.GetConfigurationNames());

        /// <summary>
        /// The mass properties of the part
        /// </summary>
        public MassProperties MassProperties => Extension.GetMassProperties();

        public bool IsVisible => BaseObject.Visible;

        public ModelView ActiveModelView => new ModelView(BaseObject.IActiveView);

        #endregion

        #region Public Events

        /// <summary>
        /// Called after the a drawing sheet was added
        /// </summary>
        public event Action<string> DrawingSheetAdded = (sheetName) => { };

        /// <summary>
        /// Called after the active drawing sheet has changed
        /// </summary>
        public event Action<string> DrawingActiveSheetChanged = (sheetName) => { };

        /// <summary>
        /// Called before the active drawing sheet changes
        /// </summary>
        public event Action<string> DrawingActiveSheetChanging = (sheetName) => { };

        /// <summary>
        /// Called after the a drawing sheet was deleted
        /// </summary>
        public event Action<string> DrawingSheetDeleted = (sheetName) => { };

        /// <summary>
        /// Called as the model is about to be closed
        /// </summary>
        public event Action ModelClosing = () => { };

        /// <summary>
        /// Called when the model is first modified since it was last saved.
        /// SOLIDWORKS marks the file as Dirty and sets <see cref="IModelDoc2.GetSaveFlag"/>
        /// </summary>
        public event Action ModelModified = () => { };

        /// <summary>
        /// Called after a model was rebuilt (any model type) or if the rollback bar position changed (for parts and assemblies).
        /// NOTE: Does not always fire on normal rebuild (Ctrl+B) on assemblies.
        /// </summary>
        public event Action ModelRebuilt = () => { };

        /// <summary>
        /// Called as the model has been saved
        /// </summary>
        public event Action ModelSaved = () => { };

        /// <summary>
        /// Called when the user cancels the save action and <see cref="ModelSaved"/> will not be fired.
        /// </summary>
        public event Action ModelSaveCanceled = () => { };

        /// <summary>
        /// Called before a model is saved with a new file name.
        /// Called before the Save As dialog is shown.
        /// Allows you to make changes that need to be included in the save. 
        /// </summary>
        public event Action<string> ModelSavingAs = (fileName) => { };

        /// <summary>
        /// Called when any of the model properties changes
        /// </summary>
        public event Action ModelInformationChanged = () => { };

        /// <summary>
        /// Called when the active configuration has changed
        /// </summary>
        public event Action ActiveConfigurationChanged = () => { };

        /// <summary>
        /// Called when the selected objects in the model have changed
        /// </summary>
        public event Action SelectionChanged = () => { };

        #endregion

        #region Constructor

        /// <summary>
        /// Default constructor
        /// </summary>
        public Model(ModelDoc2 model) : base(model)
        {
            // Update information about this model
            ReloadModelData();

            // Attach event handlers
            SetupModelEventHandlers();
        }

        #endregion

        #region Model Data

        /// <summary>
        /// Reloads all variables and data about this model
        /// </summary>
        protected void ReloadModelData()
        {
            // Clean up any previous data
            DisposeAllReferences();

            // Can't do much if there is no document
            if (BaseObject == null)
                return;

            // Get the file path
            FilePath = BaseObject.GetPathName();

            Name = BaseObject.GetTitle();

            // Get the models type
            ModelType = (ModelType)BaseObject.GetType();

            // Get the extension
            Extension = new ModelExtension(BaseObject.Extension, this);

            // Get the active configuration
            ActiveConfiguration = new ModelConfiguration(BaseObject.IGetActiveConfiguration());

            // Get the selection manager
            SelectionManager = new SelectionManager(BaseObject.ISelectionManager);

            // Set drawing access
            Drawing = IsDrawing ? new DrawingDocument((DrawingDoc)BaseObject) : null;

            // Set part access
            Part = IsPart ? new PartDocument((PartDoc)BaseObject) : null;

            // Set assembly access
            Assembly = IsAssembly ? new AssemblyDocument((AssemblyDoc)BaseObject) : null;

            // Inform listeners
            ModelInformationChanged();
        }

        /// <summary>
        /// Hooks into model-specific events for keeping track of up-to-date information
        /// </summary>
        protected void SetupModelEventHandlers()
        {
            // Based on the type of model this is...
            switch (ModelType)
            {
                // Hook into the save and destroy events to keep data fresh
                case ModelType.Assembly:
                    AsAssembly().ActiveConfigChangePostNotify += ActiveConfigChangePostNotify;
                    AsAssembly().DestroyNotify += FileDestroyedNotify;
                    AsAssembly().FileSaveAsNotify2 += FileSaveAsPreNotify;
                    AsAssembly().FileSavePostCancelNotify += FileSaveCanceled;
                    AsAssembly().FileSavePostNotify += FileSavePostNotify;
                    AsAssembly().ModifyNotify += FileModified;
                    AsAssembly().RegenPostNotify2 += AssemblyOrPartRebuilt;
                    AsAssembly().UserSelectionPostNotify += UserSelectionPostNotify;
                    AsAssembly().ClearSelectionsNotify += UserSelectionPostNotify;
                    break;
                case ModelType.Part:
                    AsPart().ActiveConfigChangePostNotify += ActiveConfigChangePostNotify;
                    AsPart().DestroyNotify += FileDestroyedNotify;
                    AsPart().FileSaveAsNotify2 += FileSaveAsPreNotify;
                    AsPart().FileSavePostCancelNotify += FileSaveCanceled;
                    AsPart().FileSavePostNotify += FileSavePostNotify;
                    AsPart().ModifyNotify += FileModified;
                    AsPart().RegenPostNotify2 += AssemblyOrPartRebuilt;
                    AsPart().UserSelectionPostNotify += UserSelectionPostNotify;
                    AsPart().ClearSelectionsNotify += UserSelectionPostNotify;
                    break;
                case ModelType.Drawing:
                    AsDrawing().ActivateSheetPostNotify += SheetActivatePostNotify;
                    AsDrawing().ActivateSheetPreNotify += SheetActivatePreNotify;
                    AsDrawing().AddItemNotify += DrawingItemAddNotify;
                    AsDrawing().DeleteItemNotify += DrawingDeleteItemNotify;
                    AsDrawing().DestroyNotify += FileDestroyedNotify;
                    AsDrawing().FileSaveAsNotify2 += FileSaveAsPreNotify;
                    AsDrawing().FileSavePostCancelNotify += FileSaveCanceled;
                    AsDrawing().FileSavePostNotify += FileSavePostNotify;
                    AsDrawing().ModifyNotify += FileModified;
                    AsDrawing().RegenPostNotify += DrawingRebuilt;
                    AsDrawing().UserSelectionPostNotify += UserSelectionPostNotify;
                    AsDrawing().ClearSelectionsNotify += UserSelectionPostNotify;
                    break;
            }
        }

        #endregion

        #region Model Event Methods

        /// <summary>
        /// Called when a part or assembly has its active configuration changed
        /// </summary>
        /// <returns></returns>
        protected int ActiveConfigChangePostNotify()
        {
            // Refresh data
            ReloadModelData();

            // Inform listeners
            ActiveConfigurationChanged();

            // NOTE: 0 is success, anything else is an error
            return 0;
        }

        /// <summary>
        /// Called after an assembly or part was rebuilt or if the rollback bar position changed.
        /// </summary>
        /// <param name="firstFeatureBelowRollbackBar"></param>
        /// <returns></returns>
        protected int AssemblyOrPartRebuilt(object firstFeatureBelowRollbackBar)
        {
            // Inform listeners
            ModelRebuilt();

            // NOTE: 0 is success, anything else is an error
            return 0;
        }

        /// <summary>
        /// Called when a drawing item is added to the feature tree
        /// </summary>
        /// <param name="entityType">Type of entity that is changed</param>
        /// <param name="itemName">Name of entity that is changed</param>
        /// <returns></returns>
        protected int DrawingItemAddNotify(int entityType, string itemName)
        {
            // Check if a sheet is added.
            // SolidWorks always activates the new sheet, but the sheet activate events aren't fired.
            if (EntityIsDrawingSheet(entityType))
            {
                SheetAddedNotify(itemName);
                SheetActivatePostNotify(itemName);
            }

            // NOTE: 0 is success, anything else is an error
            return 0;
        }

        /// <summary>
        /// Called when a drawing item is removed from the feature tree
        /// </summary>
        /// <param name="entityType">Type of entity that is changed</param>
        /// <param name="itemName">Name of entity that is changed</param>
        /// <returns></returns>
        protected int DrawingDeleteItemNotify(int entityType, string itemName)
        {
            // Check if the removed items is a sheet
            if (EntityIsDrawingSheet(entityType))
            {
                // Inform listeners
                DrawingSheetDeleted(itemName);
            }

            // NOTE: 0 is success, anything else is an error
            return 0;
        }

        /// <summary>
        /// Called after a drawing was rebuilt.
        /// </summary>
        /// <returns></returns>
        protected int DrawingRebuilt()
        {
            // Inform listeners
            ModelRebuilt();

            // NOTE: 0 is success, anything else is an error
            return 0;
        }

        /// <summary>
        /// Called when the user changes the selected objects
        /// </summary>
        /// <returns></returns>
        protected int UserSelectionPostNotify()
        {
            // Inform Listeners
            SelectionChanged();

            return 0;
        }

        /// <summary>
        /// Called when the user cancels the save action and <see cref="FileSavePostNotify"/> will not be fired.
        /// </summary>
        /// <returns></returns>
        private int FileSaveCanceled()
        {
            // Inform listeners
            ModelSaveCanceled();

            // NOTE: 0 is success, anything else is an error
            return 0;
        }

        /// <summary>
        /// Called when a model has been saved
        /// </summary>
        /// <param name="fileName">The name of the file that has been saved</param>
        /// <param name="saveType">The type of file that has been saved</param>
        /// <returns></returns>
        protected int FileSavePostNotify(int saveType, string fileName)
        {
            // Update filepath
            if (BaseObject != null)
                FilePath = BaseObject.GetPathName();
            else
                FilePath = fileName;

            // Inform listeners
            ModelSaved();

            // NOTE: 0 is success, anything else is an error
            return 0;
        }

        /// <summary>
        /// Called when a model is about to be saved with a new file name.
        /// Called before the Save As dialog is shown.
        /// </summary>
        /// <param name="fileName"></param>
        /// <returns></returns>
        private int FileSaveAsPreNotify(string fileName)
        {
            // Inform listeners
            ModelSavingAs(fileName);

            // NOTE: 0 is success, anything else is an error
            return 0;
        }

        /// <summary>
        /// Called when a model is about to be destroyed.
        /// This is a pre-notify so just clean up data, don't reload for new data yet
        /// </summary>
        /// <returns></returns>
        protected int FileDestroyedNotify()
        {
            // Inform listeners
            ModelClosing();

            // This is a pre-notify but we are going to be dead
            // so dispose ourselves (our underlying COM objects)
            Dispose();

            // NOTE: 0 is success, anything else is an error
            return 0;
        }

        /// <summary>
        /// Called when the model is first modified since it was last saved.
        /// </summary>
        /// <returns></returns>
        protected int FileModified()
        {
            // Inform listeners
            ModelModified();

            // NOTE: 0 is success, anything else is an error
            return 0;
        }

        /// <summary>
        /// Called after the active drawing sheet is changed.
        /// </summary>
        /// <param name="sheetName">Name of the sheet that is activated</param>
        /// <returns></returns>
        protected int SheetActivatePostNotify(string sheetName)
        {
            // Inform listeners
            DrawingActiveSheetChanged(sheetName);

            // NOTE: 0 is success, anything else is an error
            return 0;
        }

        /// <summary>
        /// Called before the active drawing sheet changes
        /// </summary>
        protected int SheetActivatePreNotify(string sheetName)
        {
            // Inform listeners
            DrawingActiveSheetChanging(sheetName);

            // NOTE: 0 is success, anything else is an error
            return 0;

        }

        /// <summary>
        /// Called after a sheet is added.
        /// </summary>
        /// <param name="sheetName">Name of the new sheet</param>
        /// <returns></returns>
        protected int SheetAddedNotify(string sheetName)
        {
            // Inform listeners
            DrawingSheetAdded(sheetName);

            // NOTE: 0 is success, anything else is an error
            return 0;
        }

        #endregion

        #region Specific Model

        /// <summary>
        /// Casts the current model to an assembly
        /// NOTE: Check the <see cref="ModelType"/> to confirm this model is of the correct type before casting
        /// </summary>
        /// <returns></returns>
        public AssemblyDoc AsAssembly() { return (AssemblyDoc)BaseObject; }

        /// <summary>
        /// Casts the current model to a part
        /// NOTE: Check the <see cref="ModelType"/> to confirm this model is of the correct type before casting
        /// </summary>
        /// <returns></returns>
        public PartDoc AsPart() { return (PartDoc)BaseObject; }

        /// <summary>
        /// Casts the current model to a drawing
        /// NOTE: Check the <see cref="ModelType"/> to confirm this model is of the correct type before casting
        /// </summary>
        /// <returns></returns>
        public DrawingDoc AsDrawing() { return (DrawingDoc)BaseObject; }

        /// <summary>
        /// Accesses the current model as a drawing to expose all Drawing API calls.
        /// Check <see cref="IsDrawing"/> before calling into this.
        /// </summary>
        public DrawingDocument Drawing { get; private set; }

        /// <summary>
        /// Accesses the current model as a part to expose all Part API calls.
        /// Check <see cref="IsPart"/> before calling into this.
        /// </summary>
        public PartDocument Part { get; private set; }

        /// <summary>
        /// Accesses the current model as an assembly to expose all Assembly API calls.
        /// Check <see cref="IsAssembly"/> before calling into this.
        /// </summary>
        public AssemblyDocument Assembly { get; private set; }

        public bool IsLoaded
        {
            get
            {
                if (BaseObject != null)
                    return true;
                else
                    return false;
            }
        }

        #endregion

        #region Custom Properties

        /// <summary>
        /// Sets a custom property to the given value.
        /// If a configuration is specified then the configuration-specific property is set
        /// </summary>
        /// <param name="name">The name of the property</param>
        /// <param name="value">The value of the property</param>
        /// <param name="configuration">The configuration to set the properties from, otherwise set custom property</param>
        public CustomPropertySetResult SetCustomPropertyValue(string name, string value, string configuration = null)
        {
            // Get the custom property editor
            using (var editor = Extension.CustomPropertyEditor(configuration))
            {
                // Set the property
                return editor.SetCustomPropertyValue(name, value);
            }
        }

        /// <summary>
        /// Adds a custom property.
        /// If a configuration is specified then the configuration-specific property is set
        /// </summary>
        /// <param name="name">The name of the custom property</param>
        /// <param name="type">The type of the custom property</param>
        /// <param name="option">The option how to handle existing custom property</param>
        /// <param name="value">The value of the custom property</param>
        public CustomPropertyAddResult AddCustomPropertyValue(string name, CustomPropertyTypes type, string value, CustomPropertyAddOption option, string configuration = null)
        {
            // Get the custom property editor
            using (var editor = Extension.CustomPropertyEditor(configuration))
            {
                // Add the property
                return editor.AddCustomPropertyValue(name, type, value, option);
            }
        }

        /// <summary>
        /// Checks if a custom property exists
        /// </summary>
        /// <param name="name">The name of the custom property</param>
        /// <returns></returns>
        public bool CustomPropertyExists(string name, string configuration = null)
        {
            // Get the custom property editor
            using (var editor = Extension.CustomPropertyEditor(configuration))
            {
                return editor.CustomPropertyExists(name);
            }
        }

        /// <summary>
        /// Gets a custom property by the given name
        /// </summary>
        /// <param name="name">The name of the custom property</param>
        /// <param name="configuration">The configuration to get the properties from, otherwise get custom property</param>
        ///<param name="resolved">True to get the resolved value of the property, false to get the actual text</param>
        /// <returns></returns>
        public Tuple<CustomPropertyGetResult, string> GetCustomPropertyValue(string name, string configuration = null, bool resolved = false)
        {
            // Get the custom property editor
            using (var editor = Extension.CustomPropertyEditor(configuration))
            {
                // Get the property
                return editor.GetCustomPropertyValue(name, resolved);
            }
        }

        /// <summary>
        /// Gets all of the custom properties in this model.
        /// Simply set the Value of the custom property to edit it
        /// </summary>
        /// <param name="action">The custom properties list to be worked on inside the action. NOTE: Do not store references to them outside of this action</param>
        /// <param name="configuration">Specify a configuration to get configuration-specific properties</param>
        /// <returns></returns>
        public void CustomProperties(Action<List<CustomProperty>> action, string configuration = null)
        {
            // Get the custom property editor
            using (var editor = Extension.CustomPropertyEditor(configuration))
            {
                // Get the properties
                var properties = editor.GetCustomProperties();

                // Let the action use them
                action(properties);
            }
        }

        public Type GetCustomPropertyType(string name, string configuration = null)
        {
            // Get the custom property editor
            using (var editor = Extension.CustomPropertyEditor(configuration))
            {
                // Get the properties
                return editor.GetCustomPropertyType(name);
            }
        }

        public List<string> CustomPropertyNames(string configuration = null)
        {
            if (Extension != null)
            {
                // Get the custom property editor
                using (var editor = Extension.CustomPropertyEditor(configuration))
                {
                    return editor.GetCustomPropertyNames();
                }
            }
            return new List<string>();
        }

        /// <summary>
        /// Deletes a existing custom property.
        /// If a configuration is specified then the configuration-specific property is set
        /// </summary>
        /// <param name="name">The name of the custom property</param>
        public CustomPropertyDeleteResult DeleteCustomPropertyValue(string name, string configuration = null)
        {
            // Get the custom property editor
            using (var editor = Extension.CustomPropertyEditor(configuration))
            {
                // Add the property
                return editor.DeleteCustomProperty(name);
            }
        }

        #endregion

        #region Drawings

        /// <summary>
        /// Check if an entity that was added, changed or removed is a drawing sheet.
        /// </summary>
        /// <param name="entityType">Type of the entity</param>
        /// <returns></returns>
        private static bool EntityIsDrawingSheet(int entityType)
        {
            return entityType == (int)swNotifyEntityType_e.swNotifyDrawingSheet;
        }

        #endregion

        #region Material

        /// <summary>
        /// Read the material from the model
        /// </summary>
        /// <returns></returns>
        public Material GetMaterial()
        {
            // Wrap any error
            return SolidDnaErrors.Wrap(() =>
            {
                // Get the Id's
                var idString = BaseObject.MaterialIdName;

                // Make sure we have some data
                if (idString == null || !idString.Contains("|"))
                    return null;

                // The Id string is split by pipes |
                var ids = idString.Split('|');

                // We need at least the first and second 
                // (first is database file name, second is material name)
                if (ids.Length < 2)
                    throw new ArgumentOutOfRangeException(Localization.GetString("SolidWorksModelGetMaterialIdMissingError"));

                // Extract data
                var databaseName = ids[0];
                var materialName = ids[1];

                // See if we have a database file with the same name
                var fullPath = SolidWorksEnvironment.Application.GetMaterials()?.FirstOrDefault(f => string.Equals(databaseName, Path.GetFileNameWithoutExtension(f.Database), StringComparison.InvariantCultureIgnoreCase));
                var found = fullPath != null;

                // Now we have the file, try and find the material from it
                if (found)
                {
                    var foundMaterial = SolidWorksEnvironment.Application.FindMaterial(fullPath.Database, materialName);
                    if (foundMaterial != null)
                        return foundMaterial;
                }

                // If we got here, the material was not found
                // So fill in as much information as we have
                return new Material
                {
                    Database = databaseName,
                    Name = materialName
                };
            },
                SolidDnaErrorTypeCode.SolidWorksModel,
                SolidDnaErrorCode.SolidWorksModelGetMaterialError,
                Localization.GetString("SolidWorksModelGetMaterialError"));
        }

        /// <summary>
        /// Sets the material for the model
        /// </summary>
        /// <param name="material">The material</param>
        /// <param name="configuration">The configuration to set the material on, null for the default</param>
        /// 
        public void SetMaterial(Material material, string configuration = null)
        {
            // Wrap any error
            SolidDnaErrors.Wrap(() =>
            {
                // Make sure we are a part
                if (!IsPart)
                    throw new InvalidOperationException(Localization.GetString("SolidWorksModelSetMaterialModelNotPartError"));

                // If the material is null, remove the material
                if (material == null || !material.DatabaseFileFound)
                    AsPart().SetMaterialPropertyName2(string.Empty, string.Empty, string.Empty);
                // Otherwise set the material
                else
                    AsPart().SetMaterialPropertyName2(configuration, material.Database, material.Name);
            },
                SolidDnaErrorTypeCode.SolidWorksModel,
                SolidDnaErrorCode.SolidWorksModelSetMaterialError,
                Localization.GetString("SolidWorksModelSetMaterialError"));
        }

        #endregion

        #region Selected Entities

        /// <summary>
        /// Gets all of the selected objects in the model
        /// </summary>
        /// <param name="action">The selected objects list to be worked on inside the action. NOTE: Do not store references to them outside of this action</param>
        /// <returns></returns>
        public void SelectedObjects(Action<List<SelectedObject>> action)
        {
            SelectionManager?.SelectedObjects(action);
        }

        #endregion

        #region Features

        public ModelFeature GetFirstFeature()
        {
            if (UnsafeObject != null)
                return new ModelFeature((Feature)UnsafeObject.FirstFeature());
            return null;
        }

        /// <summary>
        /// Recurses the model for all of it's features and sub-features
        /// </summary>
        /// <param name="featureAction">The callback action that is called for each feature in the model</param>
        public void Features(Action<ModelFeature, int> featureAction)
        {
            RecurseFeatures(featureAction, UnsafeObject.FirstFeature() as Feature);
        }

        #region Private Feature Helpers

        /// <summary>
        /// Recurses features and sub-features and provides a callback action to process and work with each feature
        /// </summary>
        /// <param name="featureAction">The callback action that is called for each feature in the model</param>
        /// <param name="startFeature">The feature to start at</param>
        /// <param name="featureDepth">The current depth of the sub-features based on the original calling feature</param>
        private void RecurseFeatures(Action<ModelFeature, int> featureAction, Feature startFeature = null, int featureDepth = 0)
        {
            // Get the current feature
            var currentFeature = startFeature;

            // While that feature is not null...
            while (currentFeature != null)
            {
                // Now get the first sub-feature
                var subFeature = currentFeature.GetFirstSubFeature() as Feature;

                // Get the next feature if we should
                var nextFeature = default(Feature);
                if (featureDepth == 0)
                    nextFeature = currentFeature.GetNextFeature() as Feature;

                // Create model feature
                using (var modelFeature = new ModelFeature((Feature)currentFeature))
                {
                    // Inform callback of the feature
                    featureAction(modelFeature, featureDepth);
                }

                // While we have a sub-feature...
                while (subFeature != null)
                {
                    // Get its next sub-feature
                    var nextSubFeature = subFeature.GetNextSubFeature() as Feature;

                    // Recurse all of the sub-features
                    RecurseFeatures(featureAction, subFeature, featureDepth + 1);

                    // And once back up out of the recursive dive
                    // Move to the next sub-feature and process that
                    subFeature = nextSubFeature;
                }

                // If we are at the top-level...
                if (featureDepth == 0)
                {
                    // And update the current feature reference to the next feature
                    // to carry on the loop
                    currentFeature = nextFeature;
                }
                // Otherwise...
                else
                {
                    // Remove the current feature as it is a sub-feature
                    // and is processed in the `while (subFeature != null)` loop
                    currentFeature = null;
                }
            }
        }

        private IEnumerable<Type> FindSpecificInterfacesFromDispatch(object disp)
        {
            if (disp == null)
            {
                throw new ArgumentNullException("disp");
            }

            var types = typeof(ISldWorks).Assembly.GetTypes();

            foreach (var type in types)
            {
                if (type.IsInstanceOfType(disp))
                {
                    yield return type;
                }
            }
        }

        #endregion

        #endregion

        #region Components

        /// <summary>
        /// Recurses the model for all of it's components and sub-components
        /// </summary>
        /// <param name="componentAction">The callback action that is called for each component in the model</param>
        public void Components(Action<Component, int> componentAction)
        {
            RecurseComponents(componentAction, new Component(ActiveConfiguration.UnsafeObject?.GetRootComponent3(true)));
        }

        public Component GetRootComponent()
        {
            return new Component(ActiveConfiguration.UnsafeObject?.GetRootComponent3(true));
        }

        #region Private Component Helpers

        /// <summary>
        /// Recurses components and sub-components and provides a callback action to process and work with each components
        /// </summary>
        /// <param name="componentAction">The callback action that is called for each components in the component</param>
        /// <param name="startComponent">The components to start at</param>
        /// <param name="componentDepth">The current depth of the sub-components based on the original calling components</param>
        private void RecurseComponents(Action<Component, int> componentAction, Component startComponent = null, int componentDepth = 0)
        {
            // While that component is not null...
            if (startComponent != null)
                // Inform callback of the feature
                componentAction(startComponent, componentDepth);

            // Loop each child
            foreach (Component2 childComponent in startComponent.Children)
            {
                // Get the current component
                using (var currentComponent = new Component(childComponent))
                {
                    // If we have a component
                    if (currentComponent != null)
                        // Recurse into it
                        RecurseComponents(componentAction, currentComponent, componentDepth + 1);
                }
            }
        }

        #endregion

        #endregion

        #region Saving

        /// <summary>
        /// Saves a file to the specified path, with the specified options
        /// </summary>
        /// <param name="savePath">The path of the file to save as</param>
        /// <param name="version">The version</param>
        /// <param name="options">Any save as options</param>
        /// <param name="pdfExportData">The PDF Export data if the save as type is a PDF</param>
        /// <returns></returns>
        public ModelSaveResult SaveAs(string savePath, SaveAsVersion version = SaveAsVersion.CurrentVersion, SaveAsOptions options = SaveAsOptions.None, PdfExportData pdfExportData = null)
        {
            // Start with a successful result
            var results = new ModelSaveResult();

            // Set errors and warnings to none to start with
            var errors = 0;
            var warnings = 0;

            // Wrap any error
            return SolidDnaErrors.Wrap(() =>
            {
                // Try and save the model using the SaveAs method
                BaseObject.Extension.SaveAs(savePath, (int)version, (int)options, pdfExportData?.ExportData, ref errors, ref warnings);

                // If this fails, try another way
                if (errors != 0)
                    BaseObject.SaveAs4(savePath, (int)version, (int)options, ref errors, ref warnings);

                // Add any warnings
                results.Warnings = (SaveAsWarnings)warnings;

                // Add any errors
                results.Errors = (SaveAsErrors)errors;

                // If successful, and this is not a new file 
                // (otherwise the RCW changes and SolidWorksApplication has to reload ActiveModel)...
                if (results.Successful && HasBeenSaved)
                    // Reload model data
                    ReloadModelData();

                // If we have not been saved, SolidWorks never fires any FileSave events at all
                // so request a refresh of the ActiveModel. That is the best we can do
                // as this RCW is now invalid. If this model is not active when saved then 
                // it will simply reload the active models information
                if (!HasBeenSaved)
                    SolidWorksEnvironment.Application.RequestActiveModelChanged();

                // Return result
                return results;
            },
                SolidDnaErrorTypeCode.SolidWorksModel,
                SolidDnaErrorCode.SolidWorksModelSaveAsError,
                Localization.GetString("SolidWorksModelGetMaterialError"));
        }

        #endregion

        #region Dispose

        /// <summary>
        /// Clean up all COM references for this model, its children and anything that was used by this model
        /// </summary>
        protected void DisposeAllReferences()
        {
            // Tidy up embedded SolidDNA objects
            Extension?.Dispose();
            Extension = null;

            // Release the active configuration
            ActiveConfiguration?.Dispose();
            ActiveConfiguration = null;

            // Selection manager
            SelectionManager?.Dispose();
            SelectionManager = null;
        }

        public override void Dispose()
        {
            // Clean up embedded objects
            DisposeAllReferences();

            // Dispose self
            base.Dispose();
        }

        #endregion

        public delegate int AttachedRenamedDocumentNotify(ref object target);

        public bool RebuildAndSave(AttachedRenamedDocumentNotify attachedRenamedDocumentNotify, bool updateReferences, bool rebuildNecessaryModels, bool ignoreSaveFlag = false)
        {
            var errors = 0;
            var warnings = 0;
            // rebuilds only those features that need to be rebuilt in the active configuration in the model
            if (rebuildNecessaryModels)
                BaseObject.EditRebuild3();
            // rebuilds the model in assembly and drawing documents and returns the status of the rebuild. 
            //Extension.Rebuild(swRebuildOptions_e.swRebuildAll);

            if (BaseObject.GetSaveFlag() || ignoreSaveFlag)
            {
                if (updateReferences)
                    AttachEventHandlers(attachedRenamedDocumentNotify);
                var ret = BaseObject.Save3((int)swSaveAsOptions_e.swSaveAsOptions_SaveReferenced, ref errors, ref warnings);
                if (updateReferences)
                    DetachEventHandlers(attachedRenamedDocumentNotify);
                return ret;
            }
            return true;
        }

        public bool Rename(string oldName, string newName)
        {
            return Extension.Rename(oldName, newName);
        }

        private void AttachEventHandlers(AttachedRenamedDocumentNotify attachedRenamedDocumentNotify)
        {
            if (BaseObject != null)
            {
                if (IsAssembly)
                {
                    ((AssemblyDoc)BaseObject).RenameItemNotify += RenameItemNotify;
                    ((AssemblyDoc)BaseObject).RenamedDocumentNotify += RenamedDocumentNotify;
                    if (attachedRenamedDocumentNotify != null)
                        ((AssemblyDoc)BaseObject).RenamedDocumentNotify += new DAssemblyDocEvents_RenamedDocumentNotifyEventHandler(attachedRenamedDocumentNotify);
                }
                else if (IsPart)
                {
                    ((PartDoc)BaseObject).RenameItemNotify += RenameItemNotify;
                    ((PartDoc)BaseObject).RenamedDocumentNotify += RenamedDocumentNotify;
                    if (attachedRenamedDocumentNotify != null)
                        ((PartDoc)BaseObject).RenamedDocumentNotify += new DPartDocEvents_RenamedDocumentNotifyEventHandler(attachedRenamedDocumentNotify);
                }
                else if (IsDrawing)
                {
                    ((DrawingDoc)BaseObject).RenameItemNotify += RenameItemNotify;
                }
            }
        }

        private void DetachEventHandlers(AttachedRenamedDocumentNotify attachedRenamedDocumentNotify)
        {
            if (BaseObject != null)
            {
                if (IsAssembly)
                {
                    ((AssemblyDoc)BaseObject).RenameItemNotify -= RenameItemNotify;
                    ((AssemblyDoc)BaseObject).RenamedDocumentNotify -= RenamedDocumentNotify;
                    if (attachedRenamedDocumentNotify != null)
                        ((AssemblyDoc)BaseObject).RenamedDocumentNotify -= new DAssemblyDocEvents_RenamedDocumentNotifyEventHandler(attachedRenamedDocumentNotify);
                }
                else if (IsPart)
                {
                    ((PartDoc)BaseObject).RenameItemNotify -= RenameItemNotify;
                    ((PartDoc)BaseObject).RenamedDocumentNotify -= RenamedDocumentNotify;
                    if (attachedRenamedDocumentNotify != null)
                        ((PartDoc)BaseObject).RenamedDocumentNotify -= new DPartDocEvents_RenamedDocumentNotifyEventHandler(attachedRenamedDocumentNotify);
                }
                else if (IsDrawing)
                {
                    ((DrawingDoc)BaseObject).RenameItemNotify -= RenameItemNotify;
                }
            }
        }

        //Fire notification when item is renamed
        public int RenameItemNotify(int entType, string oldName, string newName)
        {
            return 0;
        }


        //Fire notification for Rename Documents dialog
        public int RenamedDocumentNotify(ref object swObj)
        {
            var swRenamedDocumentReferences = default(RenamedDocumentReferences);
            object[] searchPaths = null;
            object[] pathNames = null;

            swRenamedDocumentReferences = (RenamedDocumentReferences)swObj;

            swRenamedDocumentReferences.UpdateWhereUsedReferences = true;
            swRenamedDocumentReferences.IncludeFileLocations = true;

            //get search paths
            searchPaths = (object[])swRenamedDocumentReferences.GetSearchPath();

            swRenamedDocumentReferences.Search();

            //get references
            pathNames = (object[])swRenamedDocumentReferences.ReferencesArray();

            swRenamedDocumentReferences.CompletionAction = (int)swRenamedDocumentFinalAction_e.swRenamedDocumentFinalAction_Ok;

            return 0;
        }

        public static Model GetModel(object component)
        {
            return Component.GetComponent(component).AsModel;
        }

        public Component GetSelectedComponentAt(int index)
        {
            var comp = SelectionManager.GetComponent(index, -1);
            return comp;
        }

        public Component GetComponentOfFeature(ModelFeature modelFeature)
        {
            if (modelFeature.FeatureType == ModelFeatureType.MateReference)
            {
                modelFeature.SelectFirst();
                return SelectionManager.GetComponent(1, -1);
            }
            return null;
        }

        public void SetVisible(bool isVisible)
        {
            if (BaseObject != null)
                BaseObject.Visible = isVisible;
        }

        public void SetUserControlable(bool isFeatureTreeEnabled, bool isGraphicsUpdateEnabled, bool isVisible, bool isLocked)
        {
            if (BaseObject != null)
            {
                ((FeatureManager)BaseObject.FeatureManager).EnableFeatureTree = isFeatureTreeEnabled;
                if (ActiveModelView != null)
                {
                    ActiveModelView.EnableGraphicsUpdate = isGraphicsUpdateEnabled;
                }
                SetVisible(isVisible);
                SetLocked(isLocked);
            }
        }

        public void SetLocked(bool isLocked)
        {
            if (BaseObject != null)
            {
                if (isLocked)
                    BaseObject.Lock();
                else
                    BaseObject.UnLock();
            }
        }

        public void ZoomModelViewToFit()
        {
            if (BaseObject != null)
            {
                BaseObject.ShowNamedView2("*Isometric", 7);
                BaseObject.ViewZoomtofit2();
            }
        }

        public void ZoomModelViewIn()
        {
            BaseObject?.ViewZoomin();
        }

        public void ZoomModelViewOut()
        {
            BaseObject?.ViewZoomout();
        }

        public bool HideAllItemTypes()
        {
            // 198 = swViewDisplayHideAllTypes  
            return BaseObject.SetUserPreferenceToggle(198, true);

        }

        public bool ShowAllItemTypes()
        {
            // 198 = swViewDisplayHideAllTypes  
            return BaseObject.SetUserPreferenceToggle(198, false);

        }

        public void HideCompleteSketch()
        {
            BaseObject.ClearSelection2(true);

            var boolstatus = BaseObject.Extension.SelectByID2("", "SKETCH", 0, 0, 0, false, 0, null, 0);
            BaseObject.BlankSketch();
        }

        public void ShowCompleteSketch()
        {
            BaseObject.ClearSelection2(true);

            var boolstatus = BaseObject.Extension.SelectByID2("", "SKETCH", 0, 0, 0, false, 0, null, 0);
            BaseObject.UnblankSketch();
        }

        private void SelectAllSketchPoints()
        {
            if (BaseObject.SketchManager.ActiveSketch != null)
            {
                if (BaseObject.SketchManager.ActiveSketch.GetSketchPoints2() != null)
                {
                    foreach (var sketchPoint in (ISketchPoint[])BaseObject.SketchManager.ActiveSketch.GetSketchPoints2())
                    {
                        sketchPoint.Select4(true, null);
                    }
                }
            }
        }

        private void SelectAllSketchSegments()
        {
            if (BaseObject.SketchManager.ActiveSketch != null)
            {
                if (BaseObject.SketchManager.ActiveSketch.GetSketchSegments() != null)
                {
                    foreach (var sketchSegment in (ISketchSegment[])BaseObject.SketchManager.ActiveSketch.GetSketchSegments())
                    {
                        sketchSegment.Select4(true, null);
                    }
                }
            }
        }
    }
}
