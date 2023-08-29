using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using System;
using System.IO;
using System.Linq;
using System.Threading.Tasks;

namespace AngelSix.SolidDna
{
    /// <summary>
    /// Represents a SolidWorks model of any type (Drawing, Part or Assembly)
    /// </summary>
    public class Component : SolidDnaObject<Component2>
    {

        #region Public Properties

        /// <summary>
        /// Get the Model from the component
        /// </summary>
        public Model AsModel => new Model(BaseObject.GetModelDoc2() as ModelDoc2);

        /// <summary>
        /// Check if the Component is Root
        /// </summary>
        public bool IsRoot => BaseObject.IsRoot();

        /// <summary>
        /// Get children from this Component
        /// </summary>
        public object[] Children => BaseObject.GetChildren() as object[];

        /// <summary>
        /// Get the name of the component whithout the instance count number at the end identicated by -xx e.g. -1
        /// </summary>
        public string Name
        {
            get
            {
                if (NameWithInstanceCount.Split('-').Length > 1)
                    return NameWithInstanceCount.Remove(NameWithInstanceCount.Length - (NameWithInstanceCount.Split('-').Last().Length + 1));
                else
                    return NameWithInstanceCount;
            }
        }

        public string NameWithInstanceCount
        {
            get => BaseObject.Name2;
            set => BaseObject.Name2 = value;
        }

        public string FilePath => BaseObject.GetPathName();

        public ComponentSuppressionStates SuppressionState
        {
            get
            {
                if (AddInIntegration.SolidWorks.SolidWorksVersion.Version <= 2018)
                    return (ComponentSuppressionStates)BaseObject.GetSuppression();
                else
                    return (ComponentSuppressionStates)BaseObject.GetSuppression2();
            }
        }

        public bool IsSuppressed => BaseObject.IsSuppressed();

        // states vary depending on suppression and hidden state
        // ConsiderSuppressed        Component       Component       IsHidden
        //       True                 Hidden       Unsuppressed        True
        //       True                 Hidden        Suppressed         True
        //       True                 Shown        Unsuppressed        False
        //       True                 Shown         Suppressed         True
        public bool IsConsiderSuppressedHidden => BaseObject.IsHidden(true);

        // states depending on hidden state
        // ConsiderSuppressed        Component       Component       IsHidden
        //       False                 Hidden       Unsuppressed        True
        //       False                 Hidden        Suppressed         True
        //       False                 Shown        Unsuppressed        False
        //       False                 Shown         Suppressed         False
        public bool IsHidden => BaseObject.IsHidden(false);

        public bool IsLoaded => BaseObject.IsLoaded();
        public bool IsVirtual => BaseObject.IsVirtual;

        public string ReferencedConfiguration
        {
            get => BaseObject.ReferencedConfiguration;
            set => BaseObject.ReferencedConfiguration = value;
        }

        public bool IsComponentVisible
        {
            get => BaseObject != null ? Convert.ToBoolean(BaseObject.Visible) : false;
            set
            {
                if (BaseObject != null)
                    BaseObject.Visible = Convert.ToInt32(value);
            }
        }

        #endregion

        #region Constructor

        /// <summary>
        /// Default constructor
        /// </summary>
        public Component(Component2 component) : base(component)
        {

        }

        #endregion

        #region ToString

        /// <summary>
        /// Returns a user-friendly string with component properties.
        /// </summary>
        /// <returns></returns>
        public override string ToString()
        {
            return $"Name: {Name}. Is root: {IsRoot}";
        }

        #endregion

        public ComponentSuppressionErrors SetSupressionState(ComponentSuppressionStates suppressionState)
        {
            lock (AddInIntegration.SolidWorks.SetSuppressionLock)
                return (ComponentSuppressionErrors)BaseObject.SetSuppression2((int)suppressionState);
        }

        public static Component GetComponent(object component)
        {
            return new Component((Component2)component);
        }

        public bool HasDrawingFileDocument()
        {
            var componentPath = new FileInfo(FilePath).Directory;
            return File.Exists(componentPath + "\\" + Name + DrawingDocument.FILE_EXTENSION);
        }

        public Component GetParentComponent()
        {
            return new Component(BaseObject.GetParent());
        }

        public async Task<ExcludeFromBOMError> StartExcludeFromBOMAsync()
        {
            return await Task.Run(() =>
            {
                return ExcludeFromBOM();
            });
        }

        public ExcludeFromBOMError ExcludeFromBOM()
        {
            if (AddInIntegration.SolidWorks.SolidWorksVersion.Version <= 2018)
            {
                BaseObject.ExcludeFromBOM = true;
                return ExcludeFromBOMError.Success;
            }
            else
                return (ExcludeFromBOMError)BaseObject.SetExcludeFromBOM2(true, (int)ModelConfigurationOptions.ThisConfiguration, null);
        }

        public async Task<ExcludeFromBOMError> StartIncludeToBOMAsync()
        {
            return await Task.Run(() =>
            {
                return IncludeToBOM();
            });
        }

        public ExcludeFromBOMError IncludeToBOM()
        {
            if (AddInIntegration.SolidWorks.SolidWorksVersion.Version <= 2018)
            {
                BaseObject.ExcludeFromBOM = false;
                return ExcludeFromBOMError.Success;
            }
            else
                return (ExcludeFromBOMError)BaseObject.SetExcludeFromBOM2(false, (int)ModelConfigurationOptions.ThisConfiguration, null);

        }

        public bool IsExcludedFromBOM()
        {
            if (AddInIntegration.SolidWorks.SolidWorksVersion.Version <= 2018)
                return BaseObject.ExcludeFromBOM;
            else
            {
                var result = (bool[])BaseObject.GetExcludeFromBOM2((int)ModelConfigurationOptions.ThisConfiguration, null);
                return result[0];
            }
        }

        #region Dispose

        public void DisposeComponent()
        {
            Dispose();
        }

        public override void Dispose()
        {
            // Clean up embedded objects
            AsModel?.Dispose();

            // Dispose self
            base.Dispose();
        }

        #endregion
    }
}
