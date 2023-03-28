using SolidWorks.Interop.sldworks;
using System.IO;
using System.Web.WebSockets;

namespace AngelSix.SolidDna
{
    /// <summary>
    /// Represents a SolidWorks Derived Part feature data
    /// </summary>
    public class FeatureDerivedPartData : SolidDnaObject<IDerivedPartFeatureData>
    {
        public string PathName => BaseObject.PathName;

        #region Constructor

        /// <summary>
        /// Default constructor
        /// </summary>
        public FeatureDerivedPartData(object model) : base((IDerivedPartFeatureData)model)
        {

        }

        #endregion

        public Model GetModel(Model TopLevelModel)
        {
            //AddInIntegration.SolidWorks.ActivateDocument(new FileInfo(PathName).Name);
            //var model = BaseObject.GetModelDoc();
            //var Model1 = new AngelSix.SolidDna.Model(model);
            return AddInIntegration.SolidWorks.OpenFile(PathName);
            
        }

        private bool AccessSelections(Model TopLevelModel, Component componentToModify = null)
        {
            var status = BaseObject.AccessSelections(TopLevelModel.AsModelDoc2, componentToModify);
            return status;
        }

        private void ReleaseSelectionAccess()
        {
            BaseObject.ReleaseSelectionAccess();
        }

    }
}
