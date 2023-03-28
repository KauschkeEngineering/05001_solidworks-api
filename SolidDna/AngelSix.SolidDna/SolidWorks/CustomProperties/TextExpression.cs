using static AngelSix.SolidDna.Model;
using System.ComponentModel;
using System.Xml.Serialization;

namespace AngelSix.SolidDna.SolidWorks.CustomProperties
{
    public class TextExpression
    {

        #region Constants

        public const string DEFAULT_EXTENSION = ".texpr";

        #endregion

        #region Private Members

        private string _name = string.Empty;
        private string _expression = string.Empty;

        #endregion

        #region Properties

        [XmlElement("Name")]
        public string Name
        {
            get => _name;
            set => _name = value;
        }

        [XmlElement("Expression")]
        public string Expression
        {
            get => _expression;
            set => _expression = value;
        }

        #endregion

        #region Constructor

        public TextExpression(string name, string expression)
        {
            _name = name;
            _expression = expression;
        }

        public TextExpression()
        {

        }

        #endregion

    }
}
