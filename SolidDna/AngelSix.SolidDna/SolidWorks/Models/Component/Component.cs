using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using System.Linq;

namespace AngelSix.SolidDna
{
    /// <summary>
    /// Represents a SolidWorks model of any type (Drawing, Part or Assembly)
    /// </summary>
    public class Component : SolidDnaObject<Component2>
    {
        public enum SuppressionStates
        {
            Suppressed = swComponentSuppressionState_e.swComponentSuppressed,
            Lightweight = swComponentSuppressionState_e.swComponentLightweight,
            FullyResolved = swComponentSuppressionState_e.swComponentFullyResolved,
            Resolved = swComponentSuppressionState_e.swComponentResolved,
            FullyLightweight = swComponentSuppressionState_e.swComponentFullyLightweight,
            InternalIdMismatch = swComponentSuppressionState_e.swComponentInternalIdMismatch
        }

        public enum SuppressionErrors
        {
            BadComponent = swSuppressionError_e.swSuppressionBadComponent,
            BadState = swSuppressionError_e.swSuppressionBadState,
            ChangeOk = swSuppressionError_e.swSuppressionChangeOk,
            ChangeFailed = swSuppressionError_e.swSuppressionChangeFailed
        }

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

        public SuppressionStates SuppressionState => (SuppressionStates)BaseObject.GetSuppression();

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

        public SuppressionErrors SetSupressionState(SuppressionStates suppressionState)
        {
            return (SuppressionErrors)BaseObject.SetSuppression2((int)suppressionState);
        }

        #endregion

        public static Component GetComponent(object component)
        {
            return new Component((Component2)component);
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
