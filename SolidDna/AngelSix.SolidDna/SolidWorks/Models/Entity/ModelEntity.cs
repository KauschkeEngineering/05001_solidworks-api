using SolidWorks.Interop.sldworks;

namespace AngelSix.SolidDna
{
    public class ModelEntity : SharedSolidDnaObject<Entity>
    {
        #region Protected Members

        #endregion

        #region Public Properties

        public Component Component => new Component(BaseObject.GetComponent() as Component2);

        #endregion

        #region Constructor

        /// <summary>
        /// Default constructor
        /// </summary>
        public ModelEntity(Feature model) : base(model as Entity)
        {
        }

        #endregion


    }
}
