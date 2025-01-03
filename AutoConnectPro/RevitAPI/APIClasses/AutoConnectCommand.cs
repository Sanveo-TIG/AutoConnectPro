using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Revit.SDK.Samples.AutoConnectPro.CS;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using TIGUtility;

namespace AutoConnectPro
{
    [Transaction(TransactionMode.Manual)]
    public class AutoConnectCommand : IExternalCommand
    {
        /// <summary>
        /// External command mainline
        /// </summary>
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            try
            {
                if(Utility.HasValidLicense("Public"))
                {
                    if(Utility.ReadPremiumLicense("AutoConnect"))
                    {
                        CustomUIApplication customUIApplication = new CustomUIApplication
                        {
                            CommandData = commandData
                        };
                        System.Windows.Window window = new MainWindow();
                        window.Show();
                        window.Closed += OnClosing;
                        MainWindow.Instance.isStaticTool = true;
                        if (ExternalApplication.ToggleConPakToolsButtonSample != null)
                            ExternalApplication.ToggleConPakToolsButtonSample.Enabled = false;
                    }
                }
                return Result.Succeeded;
            }
            catch (Exception ex)
            {
                message = ex.Message;
                return Result.Failed;
            }
        }
        public void OnClosing(object senTagProToolr, EventArgs e)
        {
            if (ExternalApplication.ToggleConPakToolsButtonSample != null)
                ExternalApplication.ToggleConPakToolsButtonSample.Enabled = true;
        }
    }
}
