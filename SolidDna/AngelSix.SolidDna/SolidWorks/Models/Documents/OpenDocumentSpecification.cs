using SolidWorks.Interop.sldworks;

namespace AngelSix.SolidDna
{
    /// <summary>
    /// Exposes all Drawing Document calls from a <see cref="Model"/>
    /// </summary>
    public class OpenDocumentSpecification : IDocumentSpecification
    {
        private readonly DocumentSpecification mDocumentSpecification = default;
        public string FileName
        {
            get => mDocumentSpecification.FileName;
            set => mDocumentSpecification.FileName = value;
        }

        public int DocumentType
        {
            get => mDocumentSpecification.DocumentType;
            set => mDocumentSpecification.DocumentType = value;
        }

        public string ConfigurationName
        {
            get => mDocumentSpecification.ConfigurationName;
            set => mDocumentSpecification.ConfigurationName = value;
        }

        public bool Silent
        {
            get => mDocumentSpecification.Silent;
            set => mDocumentSpecification.Silent = value;
        }

        public bool ReadOnly
        {
            get => mDocumentSpecification.ReadOnly;
            set => mDocumentSpecification.ReadOnly = value;
        }

        public bool ViewOnly
        {
            get => mDocumentSpecification.ViewOnly;
            set => mDocumentSpecification.ViewOnly = value;
        }

        public bool Selective
        {
            get => mDocumentSpecification.Selective;
            set => mDocumentSpecification.Selective = value;
        }

        public bool LoadModel
        {
            get => mDocumentSpecification.LoadModel;
            set => mDocumentSpecification.LoadModel = value;
        }

        public bool AutoMissingConfig
        {
            get => mDocumentSpecification.AutoMissingConfig;
            set => mDocumentSpecification.AutoMissingConfig = value;
        }

        public object ComponentList
        {
            get => mDocumentSpecification.ComponentList;
            set => mDocumentSpecification.ComponentList = value;
        }

        public string DisplayState
        {
            get => mDocumentSpecification.DisplayState;
            set => mDocumentSpecification.DisplayState = value;
        }

        public bool LightWeight
        {
            get => mDocumentSpecification.LightWeight;
            set => mDocumentSpecification.LightWeight = value;
        }

        public bool IgnoreHiddenComponents
        {
            get => mDocumentSpecification.IgnoreHiddenComponents;
            set => mDocumentSpecification.IgnoreHiddenComponents = value;
        }

        public bool UseLightWeightDefault
        {
            get => mDocumentSpecification.UseLightWeightDefault;
            set => mDocumentSpecification.UseLightWeightDefault = value;
        }

        public int Warning
        {
            get => mDocumentSpecification.Warning;
            set => mDocumentSpecification.Warning = value;
        }

        public int Error
        {
            get => mDocumentSpecification.Error;
            set => mDocumentSpecification.Error = value;
        }

        public bool InteractiveComponentSelection
        {
            get => mDocumentSpecification.InteractiveComponentSelection;
            set => mDocumentSpecification.InteractiveComponentSelection = value;
        }

        public string SheetName
        {
            get => mDocumentSpecification.SheetName;
            set => mDocumentSpecification.SheetName = value;
        }
        public bool InteractiveAdvancedOpen
        {
            get => mDocumentSpecification.InteractiveAdvancedOpen;
            set => mDocumentSpecification.InteractiveAdvancedOpen = value;
        }

        public bool Query
        {
            get => mDocumentSpecification.Query;
            set => mDocumentSpecification.Query = value;
        }

        public bool UseSpeedPak
        {
            get => mDocumentSpecification.UseSpeedPak;
            set => mDocumentSpecification.UseSpeedPak = value;
        }

        public bool AutoRepair
        {
            get => mDocumentSpecification.AutoRepair;
            set => mDocumentSpecification.AutoRepair = value;
        }

        public bool CriticalDataRepair
        {
            get => mDocumentSpecification.CriticalDataRepair;
            set => mDocumentSpecification.CriticalDataRepair = value;
        }

        public bool LoadExternalReferencesInMemory
        {
            get => mDocumentSpecification.LoadExternalReferencesInMemory;
            set => mDocumentSpecification.LoadExternalReferencesInMemory = value;
        }

        public bool DetailingMode
        {
            get => mDocumentSpecification.DetailingMode;
            set => mDocumentSpecification.DetailingMode = value;
        }

        internal OpenDocumentSpecification(DocumentSpecification documentSpecification)
        {
            mDocumentSpecification = documentSpecification;
        }
    }
}
