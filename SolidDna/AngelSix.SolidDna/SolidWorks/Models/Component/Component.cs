using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using System.IO;
using System.Linq;

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

        public string NameWithInstanceCount => BaseObject.Name2;

        public string FilePath => BaseObject.GetPathName();

        public ComponentSuppressionStates SuppressionState => (ComponentSuppressionStates)BaseObject.GetSuppression();

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

        #region Dispose

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
