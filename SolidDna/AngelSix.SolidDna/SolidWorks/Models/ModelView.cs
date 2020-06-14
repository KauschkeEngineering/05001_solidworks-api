using SolidWorks.Interop.sldworks;

namespace AngelSix.SolidDna
{
    public class ModelView : SolidDnaObject<IModelView>
    {
        #region Public Properties

        public double ScaleFactor => BaseObject.Scale2;

        public bool EnableGraphicsUpdate
        {
            get
            {
                if (BaseObject != null)
                    return BaseObject.EnableGraphicsUpdate;
                return false;
            }
            set
            {
                if (BaseObject != null)
                    BaseObject.EnableGraphicsUpdate = value;
            }
        }

        #endregion

        #region Constructor

        /// <summary>
        /// Default constructor
        /// </summary>
        public ModelView(IModelView modelView) : base(modelView)
        {
        }

        #endregion

        #region Dispose

        public override void Dispose()
        {
            // Clear reference to be safe
            base.Dispose();
        }

        #endregion

        public void Scale(double scaleFactor)
        {
            if (BaseObject != null)
                BaseObject.Scale2 = scaleFactor;
        }

        public void Zoom(double zoomFactor)
        {
            if (BaseObject != null)
                if (zoomFactor > 0 && zoomFactor < 2)
                    BaseObject.ZoomByFactor(zoomFactor);
        }

    }
}
