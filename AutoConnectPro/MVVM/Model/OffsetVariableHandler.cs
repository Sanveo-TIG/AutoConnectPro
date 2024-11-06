using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Revit.SDK.Samples.AutoConnectPro.CS;

namespace AutoConnectPro
{
    [Transaction(TransactionMode.Manual)]
    public class OffsetVariableHandler : IExternalEventHandler
    {
        Document _doc = null;
        UIDocument _uidoc = null;
        public void Execute(UIApplication uiapp)
        {
            _uidoc = uiapp.ActiveUIDocument;
            _doc = _uidoc.Document;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Document doc = uidoc.Document;
            int.TryParse(uiapp.Application.VersionNumber, out int RevitVersion);
            string offsetVariable = RevitVersion < 2020 ? "Offset" : "Middle Elevation";
            MainWindow.Instance.offsetvariable = offsetVariable;
        }
        public string GetName()
        {
            return "Revit Addin";
        }

    }
}
