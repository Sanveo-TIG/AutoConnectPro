//
// (C) Copyright 2003-2019 by Autodesk, Inc.
//
// Permission to use, copy, modify, and distribute this software in
// object code form for any purpose and without fee is hereby granted,
// provided that the above copyright notice appears in all copies and
// that both that copyright notice and the limited warranty and
// restricted rights notice below appear in all supporting
// documentation.
//
// AUTODESK PROVIDES THIS PROGRAM "AS IS" AND WITH ALL FAULTS.
// AUTODESK SPECIFICALLY DISCLAIMS ANY IMPLIED WARRANTY OF
// MERCHANTABILITY OR FITNESS FOR A PARTICULAR USE. AUTODESK, INC.
// DOES NOT WARRANT THAT THE OPERATION OF THE PROGRAM WILL BE
// UNINTERRUPTED OR ERROR FREE.
//
// Use, duplication, or disclosure by the U.S. Government is subject to
// restrictions set forth in FAR 52.227-19 (Commercial Computer
// Software - Restricted Rights) and DFAR 252.227-7013(c)(1)(ii)
// (Rights in Technical Data and Computer Software), as applicable.
//

using System;
using System.IO;
using System.Data;
using System.Text;
using System.Reflection;
using System.Diagnostics;
using System.Collections.Generic;
using Application = Autodesk.Revit.ApplicationServices.Application;
using Autodesk.Revit.UI;
using Autodesk.Revit.DB;
using Autodesk.Revit.ApplicationServices;
using System.Linq;
using System.Windows;
using Autodesk.Revit.DB.Electrical;
using TIGUtility;
using Autodesk.Revit.DB.Events;
using System.Windows.Controls;
using Autodesk.Revit.Attributes;
using System.ComponentModel;
using AutoConnectPro;
using System.Security.Cryptography;
using System.Windows.Media.Imaging;
using System.Runtime.Remoting.Contexts;
using Autodesk.Revit.UI.Selection;
using Newtonsoft.Json;
using Autodesk.Windows;
using System.Runtime.InteropServices.ComTypes;

namespace Revit.SDK.Samples.AutoConnectPro.CS
{

    /// <summary>
    /// A class inherits IExternalApplication interface and provide an entry of the sample.
    /// It create a modeless dialog to track the changes.
    /// </summary>
    [Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)]
    [Autodesk.Revit.Attributes.Regeneration(Autodesk.Revit.Attributes.RegenerationOption.Manual)]
    [Autodesk.Revit.Attributes.Journaling(Autodesk.Revit.Attributes.JournalingMode.NoCommandData)]

    public class ExternalApplication : IExternalApplication
    {

        List<Element> collection = null;
        bool iswindowClose = false;
        public bool isOffsetTool = true;
        public bool successful;
        bool _isfirst;
        public List<Element> _deleteElements = new List<Element>();
        private readonly DateTime startDate = DateTime.UtcNow;
        bool isoffsetwindowClose = false;

        #region  Class Member Variables
        /// <summary>
        /// A controlled application used to register the DocumentChanged event. Because all trigger points
        /// in this sample come from UI, the event must be registered to ControlledApplication. 
        /// If the trigger point is from API, user can register it to application 
        /// which can retrieve from ExternalCommand.
        /// </summary>
        private static ControlledApplication m_CtrlApp;

        /// <summary>
        /// data table for information windows.
        /// </summary>
        private static DataTable m_ChangesInfoTable;
        List<Element> DistanceElements = new List<Element>();

        /// <summary>
        /// The window is used to show changes' information.
        /// </summary>
        private static ChangesInformationForm m_InfoForm;
        public static System.Windows.Window window;
        #endregion

        #region Class Static Property
        /// <summary>
        /// Property to get and set private member variables of changes log information.
        /// </summary>
        public static DataTable ChangesInfoTable
        {
            get { return m_ChangesInfoTable; }
            set { m_ChangesInfoTable = value; }
        }
        public static PushButton AutoConnectButton { get; set; }
        public static PushButton ToggleConPakToolsButton { get; set; }
        public static PushButton ToggleConPakToolsButtonSample { get; set; }
        public static List<PushButton> ToggleConPakToolsButtonList { get; set; }
        /// <summary>
        /// Property to get and set private member variables of info form.
        /// </summary>
        public static ChangesInformationForm InfoForm
        {
            get { return ExternalApplication.m_InfoForm; }
            set { ExternalApplication.m_InfoForm = value; }
        }
        #endregion

        #region IExternalApplication Members
        /// <summary>
        /// Implement this method to implement the external application which should be called when 
        /// Revit starts before a file or default template is actually loaded.
        /// </summary>
        /// <param name="application">An object that is passed to the external application 
        /// which contains the controlled application.</param> 
        /// <returns>Return the status of the external application. 
        /// A result of Succeeded means that the external application successfully started. 
        /// Cancelled can be used to signify that the user cancelled the external operation at 
        /// some point.
        /// If false is returned then Revit should inform the user that the external application 
        /// failed to load and the release the internal reference.</returns>
        public Result OnStartup(UIControlledApplication application)
        {
            /* UIDocument _uidoc = null;
             Document _doc = null;*/
            // initialize member variables.
            OnButtonCreate(application);
            m_CtrlApp = application.ControlledApplication;
            application.Idling += OnIdling;
            //window = new MainWindow();

            // register the DocumentChanged event
            // m_CtrlApp.DocumentOpened += new EventHandler<Autodesk.Revit.DB.Events.DocumentOpenedEventArgs>(application_DocumentOpened);
            //m_CtrlApp.DocumentChanged += new EventHandler<Autodesk.Revit.DB.Events.DocumentChangedEventArgs>(CtrlApp_DocumentChanged);
            // show dialog

            //m_InfoForm.Width = 300;
            // m_InfoForm.Show();
            // TaskDialog.Show("ChangesInfoTable", m_ChangesInfoTable.ToString());
            // Debug.Print(m_ChangesInfoTable.ToString());

            return Result.Succeeded;
        }

        /// <summary>
        /// Implement this method to implement the external application which should be called when 
        /// Revit is about to exit,Any documents must have been closed before this method is called.
        /// </summary>
        /// <param name="application">An object that is passed to the external application 
        /// which contains the controlled application.</param>
        /// <returns>Return the status of the external application. 
        /// A result of Succeeded means that the external application successfully shutdown. 
        /// Cancelled can be used to signify that the user cancelled the external operation at 
        /// some point.
        /// If false is returned then the Revit user should be warned of the failure of the external 
        /// application to shut down correctly.</returns>
        public Result OnShutdown(UIControlledApplication application)
        {
            //m_CtrlApp.DocumentChanged += CtrlApp_DocumentChanged;
            m_InfoForm = null;
            m_ChangesInfoTable = null;
            return Result.Succeeded;
        }
        #endregion

        #region Event handler
        /// <summary>
        /// This method is the event handler, which will dump the change information to tracking dialog
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        /// 
        public static Assembly DocumentFormatAssemblyLoad(object sender, ResolveEventArgs args)
        {
            if (args.Name.Contains("resources"))
            {
                return null;
            }
            if (args.Name.Contains("TIGUtility"))
            {
                string assemblyPath = $"{Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location)}\\TIGUtility.dll";
                var assembly = Assembly.Load(assemblyPath);
                return assembly;
            }
            return null;
        }

        public static void Toggle()
        {
            try
            {
                string s = ToggleConPakToolsButton.ItemText;
                BitmapImage OffLargeImage = new BitmapImage(new Uri("pack://application:,,,/AutoConnectPro;component/Resources/Auto-Connect-32X32-red.png"));
                BitmapImage OffImage = new BitmapImage(new Uri("pack://application:,,,/AutoConnectPro;component/Resources/Auto-Connect-16X16-red.png"));

                BitmapImage OnImage = new BitmapImage(new Uri("pack://application:,,,/AutoConnectPro;component/Resources/Auto-Connect-16X16-green.png"));
                BitmapImage OnLargeImage = new BitmapImage(new Uri("pack://application:,,,/AutoConnectPro;component/Resources/Auto-Connect-32X32-green.png"));

                if (s == "AutoConnect OFF")
                {
                    ProjectParameterHandler projectParameterHandler = new ProjectParameterHandler();
                    ExternalEvent Event = ExternalEvent.Create(projectParameterHandler);
                    Event.Raise();
                    ToggleConPakToolsButton.LargeImage = OnLargeImage;
                    ToggleConPakToolsButton.Image = OnImage;

                    ToggleConPakToolsButtonSample.Enabled = false;
                }
                else
                {
                    ToggleConPakToolsButton.LargeImage = OffLargeImage;
                    ToggleConPakToolsButton.Image = OffImage;

                    ToggleConPakToolsButtonSample.Enabled = true;
                }
                ToggleConPakToolsButton.ItemText = s.Equals("AutoConnect OFF") ? "AutoConnect ON" : "AutoConnect OFF";

            }
            catch (Exception)
            {
            }

            //METHOD 1
            /*foreach (PushButton button in ToggleConPakToolsButtonList)
            {
                if (ToggleConPakToolsButton.ItemText == "AutoConnect ON")
                {
                    if (button.ItemText == "Sample Button")
                    {
                        button.Enabled = false;
                    }
                }
                else if (ToggleConPakToolsButton.ItemText == "AutoConnect OFF")
                {
                    if (button.ItemText == "Sample Button")
                    {
                        button.Enabled = true;
                    }
                }
            }*/

            //METHOD 2
            /*foreach (Autodesk.Windows.RibbonTab tab in Autodesk.Windows.ComponentManager.Ribbon.Tabs)
            {
                if (tab.Title == "Sanveo Tools")
                {
                    foreach (Autodesk.Windows.RibbonPanel panel in tab.Panels)
                    {
                        if (panel.Source.Title == "AutoConnect")
                        {
                            foreach (var button in panel.Source.Items)
                            {
                                if (button.Text == "Sample Button")
                                {
                                    if (panel.Source.Items.Any(x => x.Text == "AutoConnect OFF"))
                                    {
                                        button.IsEnabled = true;
                                    }
                                    else
                                    {
                                        button.IsEnabled = false;
                                    }
                                }
                            }
                        }
                    }
                }
            }*/
        }
        private void OnButtonCreate(UIControlledApplication application)
        {
            string executableLocation = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
            string dllLocation = Path.Combine(executableLocation, "AutoConnectPro.dll");

            PushButtonData buttondata = new PushButtonData("ModifierBtn", "AutoConnect OFF", dllLocation, "Revit.SDK.Samples.AutoConnectPro.CS.Command");
            BitmapImage pb1Image = new BitmapImage(new Uri("pack://application:,,,/AutoConnectPro;component/Resources/Auto-Connect-32X32-red.png"));
            buttondata.LargeImage = pb1Image;
            BitmapImage pb1Image2 = new BitmapImage(new Uri("pack://application:,,,/AutoConnectPro;component/Resources/Auto-Connect-16X16-red.png"));
            buttondata.Image = pb1Image2;
            buttondata.AvailabilityClassName = "Revit.SDK.Samples.AutoConnectPro.CS.Availability";

            #region Sample PushButton 
            PushButtonData buttondataSample1 = new PushButtonData("ModifierBtnCommandAutoConnect", "AutoConnect", dllLocation, "AutoConnectPro.AutoConnectCommand");
            BitmapImage pb1ImageSample11 = new BitmapImage(new Uri("pack://application:,,,/AutoConnectPro;component/Resources/Auto-Connect-32X32-green.png"));
            buttondataSample1.LargeImage = pb1ImageSample11;
            BitmapImage pb1ImageSample12 = new BitmapImage(new Uri("pack://application:,,,/AutoConnectPro;component/Resources/Auto-Connect-16X16-green.png"));
            buttondataSample1.Image = pb1ImageSample12;
            #endregion

            var ribbonPanel = RibbonPanel(application);
            if (ribbonPanel != null)
            {
                ToggleConPakToolsButton = ribbonPanel.AddItem(buttondata) as PushButton;
                ToggleConPakToolsButtonSample = ribbonPanel.AddItem(buttondataSample1) as PushButton;
                //ToggleConPakToolsButtonList = new List<PushButton> { ToggleConPakToolsButtonSample, ToggleConPakToolsButton };
                ToggleConPakToolsButtonList = new List<PushButton> { ToggleConPakToolsButton };
            }
            AppDomain.CurrentDomain.AssemblyResolve += new ResolveEventHandler(DocumentFormatAssemblyLoad);
        }
        public Autodesk.Revit.UI.RibbonPanel RibbonPanel(UIControlledApplication a)
        {
            string tab = "Sanveo Tools"; // Archcorp
            string ribbonPanelText = "AutoConnect"; // Architecture

            // Empty ribbon panel 
            Autodesk.Revit.UI.RibbonPanel ribbonPanel = null;
            // Try to create ribbon tab. 
            try
            {
                a.CreateRibbonTab(tab);
            }
            catch { }
            // Try to create ribbon panel.
            try
            {
                Autodesk.Revit.UI.RibbonPanel panel = a.CreateRibbonPanel(tab, ribbonPanelText);
            }
            catch { }
            // Search existing tab for your panel.
            List<Autodesk.Revit.UI.RibbonPanel> panels = a.GetRibbonPanels(tab);
            foreach (Autodesk.Revit.UI.RibbonPanel p in panels)
            {
                if (p.Name == ribbonPanelText)
                {
                    ribbonPanel = p;
                }
            }
            //return panel 
            return ribbonPanel;
        }

        private void OnIdling(object sender, Autodesk.Revit.UI.Events.IdlingEventArgs e)  ///////////
        {
            bool isDisabled = false;
            Autodesk.Windows.RibbonControl ribbon = Autodesk.Windows.ComponentManager.Ribbon;
            foreach (Autodesk.Windows.RibbonTab tab in ribbon.Tabs)
            {
                if (tab.Title.Equals("Sanveo Tools"))
                {
                    foreach (Autodesk.Windows.RibbonPanel panel in tab.Panels)
                    {
                        string panelName = panel.Source.Title; // Ribbon Panel Name
                        RibbonItemCollection collctn = panel.Source.Items;
                        foreach (Autodesk.Windows.RibbonItem ri in collctn)
                        {
                            if (ri != null && !string.IsNullOrEmpty(ri.AutomationName))
                            {
                                if (ri != null && !string.IsNullOrEmpty(ri.AutomationName))
                                {
                                    if (ri is Autodesk.Windows.RibbonSplitButton splitButton)
                                    {
                                        string splitButtonName = splitButton.AutomationName; // SplitButton Name
                                        RibbonItemCollection subItems = splitButton.Items;
                                        foreach (Autodesk.Windows.RibbonItem subItem in subItems)
                                        {
                                            if (subItem != null && !string.IsNullOrEmpty(subItem.AutomationName))
                                            {
                                                ///ALL TOOL NAMES
                                                string subItemName = subItem.AutomationName; // Sub-item Name
                                                if (subItemName == "AutoConnect OFF" || subItemName == "AutoConnect ON" || subItemName == "AutoConnect" ||
                                                    subItemName == "AutoUpdate OFF" || subItemName == "AutoUpdate ON" || subItemName == "AutoUpdate")
                                                    continue;
                                                if (!subItem.IsEnabled)
                                                {
                                                    isDisabled = true;
                                                    break;
                                                }
                                            }
                                        }
                                    }
                                    else
                                    {
                                        string mainItemName = ri.AutomationName;
                                        if (mainItemName == "AutoConnect OFF" || mainItemName == "AutoConnect ON" || mainItemName == "AutoConnect" ||
                                            mainItemName == "AutoUpdate OFF" || mainItemName == "AutoUpdate ON" || mainItemName == "AutoUpdate")
                                            continue;

                                        if (!ri.IsEnabled)
                                        {
                                            isDisabled = true;
                                            break;
                                        }
                                    }
                                }
                            }
                        }
                        if (isDisabled) break;
                    }
                    if (isDisabled) break;
                }
            }
            try
            {
                if (!isDisabled) //Another Tool OFF
                {
                    if (ToggleConPakToolsButton.ItemText == "AutoConnect ON")
                    {
                        List<Element> selectedElements = new List<Element>();
                        UIApplication uiApp = sender as UIApplication;
                        UIDocument uiDoc = uiApp.ActiveUIDocument;
                        Document doc = uiDoc.Document;
                        int.TryParse(uiApp.Application.VersionNumber, out int RevitVersion);
                        string offsetVariable = RevitVersion < 2020 ? "Offset" : "Middle Elevation";

                        if (doc != null && !doc.IsReadOnly)
                        {
                            // Get the current selection
                            Selection selection = uiDoc.Selection;
                            List<ElementId> selectedIds = selection.GetElementIds().ToList();
                            List<Element> SelectedElements = new List<Element>();
                            foreach (ElementId elementID in selectedIds)
                            {
                                if (doc.GetElement(elementID).Category != null)
                                {
                                    if (doc.GetElement(elementID).Category.Name == "Conduits")
                                    {
                                        SelectedElements.Add(doc.GetElement(elementID));
                                    }
                                    else if (doc.GetElement(elementID).Category.Name == "Conduit Fittings")
                                    {
                                        SelectedElements = new List<Element>();
                                        break;
                                    }
                                }
                            }
                            if (selectedIds.Count > 1)
                            {
                                if (doc.GetElement(selectedIds.FirstOrDefault()).Category != null)
                                {
                                    if (doc.GetElement(selectedIds.FirstOrDefault()).Category.Name == "Conduits")
                                    {
                                        if (window == null)
                                        {
                                            // if (SelectedElements.Any())
                                            //{
                                            //    if (SelectedElements.Count % 2 == 0)
                                            //    {
                                            //ExternalEvent.Create(new WindowOpenCloseHandler()).Raise();
                                            // var CongridDictionary1 = Utility.GroupByElements(SelectedElements);

                                            if (SelectedElements != null && SelectedElements.Count > 0)
                                            {
                                                var CongridDictionary1 = GroupStubElements(SelectedElements);
                                                Dictionary<double, List<Element>> group = new Dictionary<double, List<Element>>();
                                                if (CongridDictionary1.Count == 2)
                                                {
                                                    Dictionary<double, List<Element>> groupPrimary = GroupByElementsWithElevation(CongridDictionary1.First().Value.Select(x => x.Conduit).ToList(), "Middle Elevation");
                                                    Dictionary<double, List<Element>> groupSecondary = GroupByElementsWithElevation(CongridDictionary1.Last().Value.Select(x => x.Conduit).ToList(), "Middle Elevation");
                                                    foreach (var elem in groupPrimary)
                                                    {
                                                        foreach (var elem2 in elem.Value)
                                                        {
                                                            DistanceElements.Add(elem2);
                                                        }
                                                    }
                                                    if (groupPrimary.Count > 0 && groupSecondary.Count > 0)
                                                    {
                                                        //CHECK THE ELEMENTS HAVE BOTH SIDES FITTINGS
                                                        bool isErrorOccuredinAutoConnect = false;
                                                        List<Element> groupPrimarySelectedElements = new List<Element>();
                                                        List<Element> groupSecondarySelectedElements = new List<Element>();
                                                        groupPrimarySelectedElements = groupPrimary.Select(x => x.Value.FirstOrDefault()).ToList();
                                                        groupSecondarySelectedElements = groupSecondary.Select(x => x.Value.FirstOrDefault()).ToList();
                                                        if ((groupPrimary.Select(x => x.Value).ToList().FirstOrDefault().Count > 0 && groupSecondary.Select(x => x.Value).ToList().FirstOrDefault().Count > 0) &&
                                                            (groupPrimary.Select(x => x.Value).ToList().FirstOrDefault().Count == groupSecondary.Select(x => x.Value).ToList().FirstOrDefault().Count))
                                                        {
                                                            if (groupPrimarySelectedElements != null && groupPrimarySelectedElements.Count > 0)
                                                            {
                                                                List<Element> elementlist = new List<Element>();
                                                                foreach (ElementId id in groupPrimarySelectedElements.Select(x => x.Id))
                                                                {
                                                                    Element elem = doc.GetElement(id);
                                                                    if (elem.Category != null && elem.Category.Name == "Conduits")
                                                                    {
                                                                        elementlist.Add(elem);
                                                                    }
                                                                }
                                                                List<ElementId> FittingElem = new List<ElementId>();
                                                                for (int i = 0; i < elementlist.Count; i++)
                                                                {
                                                                    ConnectorSet connector = GetConnectorSet(elementlist[i]);
                                                                    List<ElementId> Icollect = new List<ElementId>();
                                                                    foreach (Connector connect in connector)
                                                                    {
                                                                        ConnectorSet cs1 = connect.AllRefs;
                                                                        foreach (Connector c in cs1)
                                                                        {
                                                                            Icollect.Add(c.Owner.Id);
                                                                        }
                                                                        foreach (ElementId eid in Icollect)
                                                                        {
                                                                            if (doc.GetElement(eid) != null && (doc.GetElement(eid).Category != null && doc.GetElement(eid).Category.Name == "Conduit Fittings"))
                                                                            {
                                                                                FittingElem.Add(eid);
                                                                            }
                                                                        }
                                                                    }
                                                                }
                                                                List<ElementId> FittingElements = new List<ElementId>();
                                                                FittingElements = FittingElem.Distinct().ToList();
                                                                if (FittingElements.Count == (2 * (elementlist.Count)))
                                                                {
                                                                    isErrorOccuredinAutoConnect = true;
                                                                }
                                                            }
                                                            if (groupSecondarySelectedElements != null && groupSecondarySelectedElements.Count > 0)
                                                            {
                                                                List<Element> elementlist = new List<Element>();
                                                                foreach (ElementId id in groupSecondarySelectedElements.Select(x => x.Id))
                                                                {
                                                                    Element elem = doc.GetElement(id);
                                                                    if (elem.Category != null && elem.Category.Name == "Conduits")
                                                                    {
                                                                        elementlist.Add(elem);
                                                                    }
                                                                }
                                                                List<ElementId> FittingElem = new List<ElementId>();
                                                                for (int i = 0; i < elementlist.Count; i++)
                                                                {
                                                                    ConnectorSet connector = GetConnectorSet(elementlist[i]);
                                                                    List<ElementId> Icollect = new List<ElementId>();
                                                                    foreach (Connector connect in connector)
                                                                    {
                                                                        ConnectorSet cs1 = connect.AllRefs;
                                                                        foreach (Connector c in cs1)
                                                                        {
                                                                            Icollect.Add(c.Owner.Id);
                                                                        }
                                                                        foreach (ElementId eid in Icollect)
                                                                        {
                                                                            if (doc.GetElement(eid) != null && (doc.GetElement(eid).Category != null && doc.GetElement(eid).Category.Name == "Conduit Fittings"))
                                                                            {
                                                                                FittingElem.Add(eid);
                                                                            }
                                                                        }
                                                                    }
                                                                }
                                                                List<ElementId> FittingElements = new List<ElementId>();
                                                                FittingElements = FittingElem.Distinct().ToList();
                                                                if (FittingElements.Count == (2 * (elementlist.Count)))
                                                                {
                                                                    isErrorOccuredinAutoConnect = true;
                                                                }
                                                            }
                                                        }
                                                        if (isErrorOccuredinAutoConnect)
                                                        {
                                                            BitmapImage OffLargeImage = new BitmapImage(new Uri("pack://application:,,,/AutoConnectPro;component/Resources/Auto-Connect-32X32-red.png"));
                                                            BitmapImage OffImage = new BitmapImage(new Uri("pack://application:,,,/AutoConnectPro;component/Resources/Auto-Connect-16X16-red.png"));
                                                            ToggleConPakToolsButton.ItemText = "AutoConnect OFF";
                                                            ToggleConPakToolsButton.LargeImage = OffLargeImage;
                                                            ToggleConPakToolsButton.Image = OffImage;
                                                            ToggleConPakToolsButtonSample.Enabled = true;
                                                        }
                                                        else
                                                        {
                                                            if (groupPrimary.Select(x => x.Value).ToList().FirstOrDefault().Count == groupSecondary.Select(x => x.Value).ToList().FirstOrDefault().Count)
                                                            {
                                                                window = new MainWindow();
                                                                MainWindow.Instance.firstElement = new List<Element>();
                                                                MainWindow.Instance.firstElement.AddRange(SelectedElements);
                                                                MainWindow.Instance._document = doc;
                                                                MainWindow.Instance._uiDocument = uiDoc;
                                                                MainWindow.Instance._uiApplication = uiApp;

                                                                //window = new MainWindow();
                                                                window.Show();
                                                            }
                                                            else if (groupSecondary.Count == 1)
                                                            {
                                                                List<Element> dictFirstElement = CongridDictionary1.First().Value.Select(x => x.Conduit).ToList();
                                                                List<Element> dictSecondElement = CongridDictionary1.Last().Value.Select(x => x.Conduit).ToList();
                                                                LocationCurve locCurve1 = dictFirstElement[0].Location as LocationCurve;
                                                                LocationCurve locCurve2 = dictSecondElement[0].Location as LocationCurve;
                                                                XYZ startPoint = Utility.GetXYvalue(locCurve1.Curve.GetEndPoint(0));
                                                                XYZ endPoint = Utility.GetXYvalue(locCurve2.Curve.GetEndPoint(1));
                                                                Line connectLine = Line.CreateBound(startPoint, endPoint);
                                                                if (Math.Round(locCurve1.Curve.GetEndPoint(0).Z, 4) == Math.Round(locCurve1.Curve.GetEndPoint(1).Z, 4) &&
                                                                   Math.Round(locCurve2.Curve.GetEndPoint(0).Z, 4) != Math.Round(locCurve2.Curve.GetEndPoint(1).Z, 4))
                                                                {
                                                                    window = new MainWindow();
                                                                    MainWindow.Instance.firstElement = new List<Element>();
                                                                    MainWindow.Instance.firstElement.AddRange(SelectedElements);
                                                                    MainWindow.Instance._document = doc;
                                                                    MainWindow.Instance._uiDocument = uiDoc;
                                                                    MainWindow.Instance._uiApplication = uiApp;

                                                                    //window = new MainWindow();
                                                                    window.Show();
                                                                }
                                                                else
                                                                {
                                                                    if (groupPrimary.Select(x => x.Value).ToList().FirstOrDefault().Count != groupSecondary.Select(x => x.Value).ToList().FirstOrDefault().Count)
                                                                    {
                                                                        System.Windows.MessageBox.Show("Please select equal count of conduits", "Warning-AUTOCONNECT", MessageBoxButton.OK, MessageBoxImage.Warning);
                                                                        window = new MainWindow();
                                                                        window.Close();
                                                                        ExternalApplication.window = null;
                                                                        SelectedElements.Clear();
                                                                        uiDoc.Selection.SetElementIds(new List<ElementId> { ElementId.InvalidElementId });
                                                                    }
                                                                }
                                                                /* List<Element> OrderSecondary = new List<Element>();
                                                                 Dictionary<double, Element> ordertheUpperElements = new Dictionary<double, Element>();
                                                                 Dictionary<double, Element> ordertheLowerElements = new Dictionary<double, Element>();
                                                                 ordertheUpperElements = new Dictionary<double, Element>();
                                                                 ordertheLowerElements = new Dictionary<double, Element>();
                                                                 foreach (Element conduit in groupSecondary.Select(x => x.Value).FirstOrDefault())
                                                                 {
                                                                     LocationCurve locationCurve = conduit.Location as LocationCurve;
                                                                     if (locationCurve != null)
                                                                     {
                                                                         XYZ sp = locationCurve.Curve.GetEndPoint(0);
                                                                         XYZ ep = locationCurve.Curve.GetEndPoint(1);
                                                                         double Value = (sp.Y);
                                                                         if (!ordertheUpperElements.ContainsKey(Value))
                                                                         {
                                                                             ordertheUpperElements.Add(Value, conduit);
                                                                         }
                                                                         else
                                                                         {
                                                                             ordertheLowerElements.Add(Value, conduit);
                                                                         }
                                                                     }
                                                                 }
                                                                 groupSecondary = new Dictionary<double, List<Element>>();
                                                                 ordertheUpperElements = ordertheUpperElements.OrderBy(kvp => kvp.Key).ToDictionary(kvp => kvp.Key, kvp => kvp.Value);
                                                                 ordertheLowerElements = ordertheLowerElements.OrderBy(kvp => kvp.Key).ToDictionary(kvp => kvp.Key, kvp => kvp.Value);
                                                                 OrderSecondary = new List<Element>();
                                                                 OrderSecondary.AddRange(ordertheUpperElements.Select(e => e.Value));
                                                                 groupSecondary.Add(1, OrderSecondary);
                                                                 OrderSecondary = new List<Element>();
                                                                 OrderSecondary.AddRange(ordertheLowerElements.Select(e => e.Value));
                                                                 groupSecondary.Add(2, OrderSecondary);
                                                                 if (groupPrimary.Select(x => x.Value).ToList().FirstOrDefault().Count == groupSecondary.Select(x => x.Value).ToList().FirstOrDefault().Count)
                                                                 {
                                                                     window = new MainWindow();
                                                                     MainWindow.Instance.firstElement = new List<Element>();
                                                                     MainWindow.Instance.firstElement.AddRange(SelectedElements);
                                                                     MainWindow.Instance._document = doc;
                                                                     MainWindow.Instance._uiDocument = uiDoc;
                                                                     MainWindow.Instance._uiApplication = uiApp;

                                                                     //window = new MainWindow();
                                                                     window.Show();
                                                                 }*/
                                                            }
                                                            else
                                                            {
                                                                System.Windows.MessageBox.Show("Please select equal count of conduits", "Warning", MessageBoxButton.OK, MessageBoxImage.Warning);
                                                                window = new MainWindow();
                                                                window.Close();
                                                                ExternalApplication.window = null;
                                                                SelectedElements.Clear();
                                                                uiDoc.Selection.SetElementIds(new List<ElementId> { ElementId.InvalidElementId });
                                                            }
                                                        }
                                                    }
                                                }
                                                else
                                                {
                                                    Autodesk.Revit.UI.RibbonPanel autoUpdaterPanel = null;
                                                    string tabName = "Sanveo Tools";
                                                    string panelName = "AutoUpdate";
                                                    string panelNameAC = "AutoConnect";
                                                    string executableLocation = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
                                                    string dllLocation = Path.Combine(executableLocation, "AutoUpdaterPro.dll");
                                                    List<Autodesk.Revit.UI.RibbonPanel> panels = uiApp.GetRibbonPanels(tabName);
                                                    Autodesk.Revit.UI.RibbonPanel autoUpdaterPanel01 = panels.FirstOrDefault(p => p.Name == panelName);
                                                    Autodesk.Revit.UI.RibbonPanel autoUpdaterPanel02 = panels.FirstOrDefault(p => p.Name == panelNameAC);
                                                    bool ErrorOccured = false;
                                                    if (autoUpdaterPanel01 != null)
                                                    {
                                                        IList<Autodesk.Revit.UI.RibbonItem> items = autoUpdaterPanel01.GetItems();
                                                        foreach (Autodesk.Revit.UI.RibbonItem item in items)
                                                        {
                                                            if (item is PushButton pushButton && pushButton.ItemText == "AutoUpdate ON")
                                                            {
                                                                ErrorOccured = true;
                                                            }
                                                            else if (item.ItemText == "AutoUpdate" && !item.Enabled && autoUpdaterPanel02.GetItems().OfType<PushButton>().Any(btn => btn.ItemText == "AutoConnect ON")
                                                                && autoUpdaterPanel01.GetItems().OfType<PushButton>().Any(btn => btn.ItemText == "AutoUpdate OFF"))
                                                            {
                                                                ErrorOccured = true;
                                                                BitmapImage OffLargeImage = new BitmapImage(new Uri("pack://application:,,,/AutoConnectPro;component/Resources/Auto-Connect-32X32-red.png"));
                                                                BitmapImage OffImage = new BitmapImage(new Uri("pack://application:,,,/AutoConnectPro;component/Resources/Auto-Connect-16X16-red.png"));
                                                                ToggleConPakToolsButton.ItemText = "AutoConnect OFF";
                                                                ToggleConPakToolsButton.LargeImage = OffLargeImage;
                                                                ToggleConPakToolsButton.Image = OffImage;
                                                                ToggleConPakToolsButtonSample.Enabled = true;
                                                            }
                                                        }
                                                    }
                                                    if (!ErrorOccured)
                                                    {
                                                        if (CongridDictionary1.Count == 1)
                                                        {
                                                            Utility.GroupByElevation(SelectedElements, offsetVariable, ref group);
                                                            Dictionary<double, List<Element>> groupPrimary = new Dictionary<double, List<Element>>();
                                                            Dictionary<double, List<Element>> groupSecondary = new Dictionary<double, List<Element>>();
                                                            int k = group.Count / 2;
                                                            for (int i = 0; i < group.Count(); i++)
                                                            {
                                                                if (i >= k)
                                                                {
                                                                    groupSecondary.Add(group.ElementAt(i).Key, group.ElementAt(i).Value);
                                                                }
                                                                else
                                                                {
                                                                    groupPrimary.Add(group.ElementAt(i).Key, group.ElementAt(i).Value);
                                                                }

                                                            }
                                                            /////////
                                                            if (groupPrimary.Count > 0 && groupSecondary.Count > 0)
                                                            {
                                                                //CHECK THE ELEMENTS HAVE BOTH SIDES FITTINGS
                                                                bool isErrorOccuredinAutoConnect = false;
                                                                List<Element> groupPrimarySelectedElements = new List<Element>();
                                                                List<Element> groupSecondarySelectedElements = new List<Element>();
                                                                groupPrimarySelectedElements = groupPrimary.Select(x => x.Value.FirstOrDefault()).ToList();
                                                                groupSecondarySelectedElements = groupSecondary.Select(x => x.Value.FirstOrDefault()).ToList();
                                                                //if ((groupPrimary.Select(x => x.Value).ToList().FirstOrDefault().Count > 0 && groupSecondary.Select(x => x.Value).ToList().FirstOrDefault().Count > 0) &&
                                                                //    (groupPrimary.Select(x => x.Value).ToList().FirstOrDefault().Count == groupSecondary.Select(x => x.Value).ToList().FirstOrDefault().Count))
                                                                if (groupPrimary == null && groupSecondary != null && groupPrimary.Count > 0 && groupSecondary.Count > 0 &&
                                                                   groupPrimary.Count == groupSecondary.Count)
                                                                {
                                                                    if (groupPrimary.Count == groupSecondary.Count)
                                                                    {
                                                                        for (int i = 0; i < groupPrimary.Count; i++)
                                                                        {
                                                                            List<Element> primarySortedElementspre = SortbyPlane(doc, groupPrimary.ElementAt(i).Value);
                                                                            List<Element> secondarySortedElementspre = SortbyPlane(doc, groupSecondary.ElementAt(i).Value);
                                                                            bool isNotStaright = ReverseingConduits(doc, ref primarySortedElementspre, ref secondarySortedElementspre);
                                                                            //defind the primary and secondary sets 
                                                                            double conduitlengthone = primarySortedElementspre[0].LookupParameter("Length").AsDouble();
                                                                            double conduitlengthsecond = secondarySortedElementspre[0].LookupParameter("Length").AsDouble();
                                                                            List<Element> primarySortedElements = new List<Element>();
                                                                            List<Element> secondarySortedElements = new List<Element>();
                                                                            if (conduitlengthone < conduitlengthsecond)
                                                                            {
                                                                                primarySortedElements = primarySortedElementspre;
                                                                                secondarySortedElements = secondarySortedElementspre;
                                                                            }
                                                                            else
                                                                            {
                                                                                primarySortedElements = secondarySortedElementspre;
                                                                                secondarySortedElements = primarySortedElementspre;
                                                                            }
                                                                            if (primarySortedElements.Count == secondarySortedElements.Count)
                                                                            {
                                                                                Element primaryFirst = primarySortedElements.First();
                                                                                Element secondaryFirst = secondarySortedElements.First();
                                                                                Element primaryLast = primarySortedElements.Last();
                                                                                Element secondaryLast = secondarySortedElements.Last();

                                                                                XYZ priFirstDir = ((primaryFirst.Location as LocationCurve).Curve as Line).Direction;
                                                                                XYZ priLastDir = ((primaryLast.Location as LocationCurve).Curve as Line).Direction;
                                                                                XYZ secFirstDir = ((secondaryFirst.Location as LocationCurve).Curve as Line).Direction;
                                                                                XYZ secLastDir = ((secondaryLast.Location as LocationCurve).Curve as Line).Direction;

                                                                                bool isSamDireFirst = Utility.IsSameDirectionWithRoundOff(priFirstDir, secFirstDir, 3) || Utility.IsSameDirectionWithRoundOff(priFirstDir, secLastDir, 3);
                                                                                bool isSamDireLast = Utility.IsSameDirectionWithRoundOff(priLastDir, secFirstDir, 3) || Utility.IsSameDirectionWithRoundOff(priLastDir, secLastDir, 3);
                                                                                //Same Elevations 
                                                                                bool isSamDir = !isNotStaright || isSamDireFirst && isSamDireLast;
                                                                                if (!isSamDir)
                                                                                {
                                                                                    Line priFirst = ((primaryFirst.Location as LocationCurve).Curve as Line);
                                                                                    Line priLast = ((primaryLast.Location as LocationCurve).Curve as Line);
                                                                                    Line secFirst = ((secondaryFirst.Location as LocationCurve).Curve as Line);
                                                                                    Line secLast = ((secondaryLast.Location as LocationCurve).Curve as Line);

                                                                                    XYZ firstInte = MultiConnectFindIntersectionPoint(priFirst, secFirst);
                                                                                    if (firstInte != null)
                                                                                    {
                                                                                        firstInte = MultiConnectFindIntersectionPoint(priFirst, secLast);

                                                                                        if (firstInte != null)
                                                                                        {
                                                                                            isSamDir = false;
                                                                                        }
                                                                                    }
                                                                                }
                                                                                if (!isSamDir && Math.Round(groupPrimary.ElementAt(i).Key, 4) == Math.Round(groupSecondary.ElementAt(i).Key, 4))
                                                                                {
                                                                                    //Multi connect
                                                                                    System.Windows.MessageBox.Show("Warning. \n" + "Please use Multi Connect tool for the selected group of conduits to connect", "Alert", MessageBoxButton.OK, MessageBoxImage.Warning);
                                                                                    return;
                                                                                }
                                                                                else if (isSamDir && Math.Round(groupPrimary.ElementAt(i).Key, 4) == Math.Round(groupSecondary.ElementAt(i).Key, 4))
                                                                                {
                                                                                    Line priFirst = ((primaryFirst.Location as LocationCurve).Curve as Line);
                                                                                    Line priLast = ((primaryLast.Location as LocationCurve).Curve as Line);
                                                                                    Line secFirst = ((secondaryFirst.Location as LocationCurve).Curve as Line);
                                                                                    Line secLast = ((secondaryLast.Location as LocationCurve).Curve as Line);
                                                                                    ConnectorSet firstConnectors = null;
                                                                                    ConnectorSet secondConnectors = null;
                                                                                    firstConnectors = Utility.GetConnectors(primaryFirst);
                                                                                    secondConnectors = Utility.GetConnectors(secondaryFirst);
                                                                                    Utility.GetClosestConnectors(firstConnectors, secondConnectors, out Connector ConnectorOne, out Connector ConnectorTwo);
                                                                                    Line checkline = Line.CreateBound(ConnectorOne.Origin, new XYZ(ConnectorTwo.Origin.X, ConnectorTwo.Origin.Y, ConnectorOne.Origin.Z));
                                                                                    XYZ p1 = new XYZ(Math.Round(priFirst.Direction.X, 2), Math.Round(priFirst.Direction.Y, 2), 0);
                                                                                    XYZ p2 = new XYZ(Math.Round(checkline.Direction.X, 2), Math.Round(checkline.Direction.Y, 2), 0);
                                                                                    bool isSamDirecheckline = new XYZ(Math.Abs(p1.X), Math.Abs(p1.Y), 0).IsAlmostEqualTo(new XYZ(Math.Abs(p2.X), Math.Abs(p2.Y), 0));
                                                                                    if (isSamDirecheckline)
                                                                                    {
                                                                                        //Extend
                                                                                        XYZ pickpoint = new XYZ();
                                                                                        if (ChangesInformationForm.instance.MidSaddlePt != null)
                                                                                        {
                                                                                            // ElementId midelm= ChangesInformationForm.instance.MidSaddlePt.Owner.Id;
                                                                                            var CongridDictionary = Utility.GroupByElements(ChangesInformationForm.instance.MidSaddlePt);
                                                                                            Dictionary<double, List<Element>> grPrimary = Utility.GroupByElementsWithElevation(CongridDictionary.First().Value.Select(x => x.Conduit).ToList(), offsetVariable);
                                                                                            Dictionary<double, List<Element>> grSecondary = Utility.GroupByElementsWithElevation(CongridDictionary.Last().Value.Select(x => x.Conduit).ToList(), offsetVariable);
                                                                                            if (grPrimary.Count == grSecondary.Count)
                                                                                            {
                                                                                                for (int j = 0; j < grPrimary.Count; j++)
                                                                                                {

                                                                                                    List<Element> primarySortedElem = SortbyPlane(doc, grPrimary.ElementAt(j).Value);

                                                                                                    List<Element> secondarySortedElem = SortbyPlane(doc, grSecondary.ElementAt(j).Value);
                                                                                                    ConnectorSet connectorSetOne = Utility.GetConnectors(primarySortedElem.FirstOrDefault());
                                                                                                    ConnectorSet connectorSetTwo = Utility.GetConnectors(secondarySortedElem.FirstOrDefault());
                                                                                                    foreach (Connector connector in connectorSetOne)
                                                                                                    {
                                                                                                        foreach (Connector connector2 in connectorSetTwo)
                                                                                                        {
                                                                                                            ConnectorSet cs = connector.AllRefs;
                                                                                                            ConnectorSet cs2 = connector2.AllRefs;
                                                                                                            foreach (Connector c in cs)
                                                                                                            {
                                                                                                                foreach (Connector c2 in cs2)
                                                                                                                {
                                                                                                                    if (c.Owner.Id == c2.Owner.Id)
                                                                                                                    {
                                                                                                                        //List<XYZ> StEn = Utility.GetFittingStartAndEndPoint(doc.GetElement(c.Owner.Id) as FamilyInstance);
                                                                                                                        // Line li = Line.CreateBound(StEn[0], StEn[1]);
                                                                                                                        //XYZ m = Utility.GetMidPoint(li);
                                                                                                                        // XYZ cross = li.Direction.CrossProduct(XYZ.BasisZ);
                                                                                                                        // XYZ newStart = m + cross.Multiply(1);
                                                                                                                        // XYZ newEnd = m - cross.Multiply(1);
                                                                                                                        // Line verticalLine = Line.CreateBound(newStart, newEnd);
                                                                                                                        // XYZ interSectionPoint = Utility.FindIntersectionPoint(verticalLine,li);
                                                                                                                        // MessageBox.Show(interSectionPoint.ToString());
                                                                                                                        XYZ ElbowLocationPoint = (doc.GetElement(c.Owner.Id).Location as LocationPoint).Point;
                                                                                                                        pickpoint = ElbowLocationPoint;//new XYZ(c.Origin.X, c.Origin.Y, 0);
                                                                                                                        break;
                                                                                                                    }
                                                                                                                }
                                                                                                            }
                                                                                                        }
                                                                                                    }
                                                                                                }
                                                                                            }
                                                                                        }
                                                                                        else
                                                                                        {
                                                                                            pickpoint = Utility.PickPoint(uiDoc);
                                                                                        }
                                                                                        if (pickpoint != null)
                                                                                            ThreePtSaddleExecute(uiApp, ref primarySortedElements, ref secondarySortedElements, pickpoint);

                                                                                    }
                                                                                    else
                                                                                    {
                                                                                        //Hoffset //else if
                                                                                        HoffsetExecute(uiApp, ref primarySortedElements, ref secondarySortedElements);
                                                                                        isOffsetTool = true;
                                                                                        isoffsetwindowClose = true;
                                                                                    }
                                                                                }
                                                                                else
                                                                                {
                                                                                    bool isVerticalConduits = false;
                                                                                    Line priFirst = ((primaryFirst.Location as LocationCurve).Curve as Line);
                                                                                    Line priLast = ((primaryLast.Location as LocationCurve).Curve as Line);
                                                                                    Line secFirst = ((secondaryFirst.Location as LocationCurve).Curve as Line);
                                                                                    Line secLast = ((secondaryLast.Location as LocationCurve).Curve as Line);
                                                                                    XYZ directionOne = priFirst.Direction;
                                                                                    XYZ directionTwo = secFirst.Direction;
                                                                                    isVerticalConduits = new XYZ(0, 0, Math.Abs(directionOne.Z)).IsAlmostEqualTo(XYZ.BasisZ)
                                                                                        && new XYZ(0, 0, Math.Abs(directionTwo.Z)).IsAlmostEqualTo(XYZ.BasisZ);

                                                                                    ConnectorSet firstConnectors = null;
                                                                                    ConnectorSet secondConnectors = null;
                                                                                    firstConnectors = Utility.GetConnectors(primaryFirst);
                                                                                    secondConnectors = Utility.GetConnectors(secondaryFirst);
                                                                                    Utility.GetClosestConnectors(firstConnectors, secondConnectors, out Connector ConnectorOne, out Connector ConnectorTwo);
                                                                                    Line checkline = Line.CreateBound(ConnectorOne.Origin, new XYZ(ConnectorTwo.Origin.X, ConnectorTwo.Origin.Y, ConnectorOne.Origin.Z));
                                                                                    XYZ p1 = new XYZ(Math.Round(priFirst.Direction.X, 2), Math.Round(priFirst.Direction.Y, 2), 0);
                                                                                    XYZ p2 = new XYZ(Math.Round(checkline.Direction.X, 2), Math.Round(checkline.Direction.Y, 2), 0);
                                                                                    bool isSamDirecheckline = new XYZ(Math.Abs(p1.X), Math.Abs(p1.Y), 0).IsAlmostEqualTo(new XYZ(Math.Abs(p2.X), Math.Abs(p2.Y), 0));

                                                                                    double priSlope = -Math.Round(priFirst.Direction.X, 6) / Math.Round(priFirst.Direction.Y, 6);
                                                                                    double SecSlope = -Math.Round(secFirst.Direction.X, 6) / Math.Round(secFirst.Direction.Y, 6);

                                                                                    if ((priSlope == -1 && SecSlope == 0) ||
                                                                                        Math.Round((Math.Round(priSlope, 5)) * (Math.Round(SecSlope, 5)), 4) == -1 ||
                                                                                        Math.Round((Math.Round(priSlope, 5)) * (Math.Round(SecSlope, 5)), 4).ToString() == double.NaN.ToString() && !isVerticalConduits)
                                                                                    {
                                                                                        //kick
                                                                                        KickExecute(uiApp, ref primarySortedElements, ref secondarySortedElements, i);
                                                                                    }
                                                                                    else if (isSamDirecheckline)
                                                                                    {
                                                                                        //Voffset
                                                                                        VoffsetExecute(uiApp, ref primarySortedElements, ref secondarySortedElements);
                                                                                    }
                                                                                    else
                                                                                    {

                                                                                        //Hoffset //else if
                                                                                        HoffsetExecute(uiApp, ref primarySortedElements, ref secondarySortedElements);
                                                                                        isOffsetTool = true;
                                                                                        isoffsetwindowClose = true;
                                                                                    }
                                                                                }
                                                                            }
                                                                        }
                                                                    }
                                                                }
                                                                if (isErrorOccuredinAutoConnect)
                                                                {
                                                                    BitmapImage OffLargeImage = new BitmapImage(new Uri("pack://application:,,,/AutoConnectPro;component/Resources/Auto-Connect-32X32-red.png"));
                                                                    BitmapImage OffImage = new BitmapImage(new Uri("pack://application:,,,/AutoConnectPro;component/Resources/Auto-Connect-16X16-red.png"));
                                                                    ToggleConPakToolsButton.ItemText = "AutoConnect OFF";
                                                                    ToggleConPakToolsButton.LargeImage = OffLargeImage;
                                                                    ToggleConPakToolsButton.Image = OffImage;
                                                                    ToggleConPakToolsButtonSample.Enabled = true;
                                                                }
                                                                else
                                                                {
                                                                    if (groupPrimary.Select(x => x.Value).ToList().FirstOrDefault().Count == groupSecondary.Select(x => x.Value).ToList().FirstOrDefault().Count)
                                                                    {
                                                                        window = new MainWindow();
                                                                        MainWindow.Instance.firstElement = new List<Element>();
                                                                        MainWindow.Instance.firstElement.AddRange(SelectedElements);
                                                                        MainWindow.Instance._document = doc;
                                                                        MainWindow.Instance._uiDocument = uiDoc;
                                                                        MainWindow.Instance._uiApplication = uiApp;

                                                                        //window = new MainWindow();
                                                                        window.Show();
                                                                    }
                                                                    else if (groupSecondary.Count == 1)
                                                                    {
                                                                        List<Element> dictFirstElement = CongridDictionary1.First().Value.Select(x => x.Conduit).ToList();
                                                                        List<Element> dictSecondElement = CongridDictionary1.Last().Value.Select(x => x.Conduit).ToList();
                                                                        LocationCurve locCurve1 = dictFirstElement[0].Location as LocationCurve;
                                                                        LocationCurve locCurve2 = dictSecondElement[0].Location as LocationCurve;
                                                                        XYZ startPoint = Utility.GetXYvalue(locCurve1.Curve.GetEndPoint(0));
                                                                        XYZ endPoint = Utility.GetXYvalue(locCurve2.Curve.GetEndPoint(1));
                                                                        Line connectLine = Line.CreateBound(startPoint, endPoint);
                                                                        if (Math.Round(locCurve1.Curve.GetEndPoint(0).Z, 4) == Math.Round(locCurve1.Curve.GetEndPoint(1).Z, 4) &&
                                                                           Math.Round(locCurve2.Curve.GetEndPoint(0).Z, 4) != Math.Round(locCurve2.Curve.GetEndPoint(1).Z, 4))
                                                                        {
                                                                            window = new MainWindow();
                                                                            MainWindow.Instance.firstElement = new List<Element>();
                                                                            MainWindow.Instance.firstElement.AddRange(SelectedElements);
                                                                            MainWindow.Instance._document = doc;
                                                                            MainWindow.Instance._uiDocument = uiDoc;
                                                                            MainWindow.Instance._uiApplication = uiApp;

                                                                            //window = new MainWindow();
                                                                            window.Show();
                                                                        }
                                                                    }
                                                                    else
                                                                    {
                                                                        System.Windows.MessageBox.Show("Please select equal count of conduits", "Warning", MessageBoxButton.OK, MessageBoxImage.Warning);
                                                                        window = new MainWindow();
                                                                        window.Close();
                                                                        ExternalApplication.window = null;
                                                                        SelectedElements.Clear();
                                                                        uiDoc.Selection.SetElementIds(new List<ElementId> { ElementId.InvalidElementId });
                                                                    }
                                                                }
                                                            }
                                                            else
                                                            {
                                                                System.Windows.MessageBox.Show("Please select equal count of conduits", "Warning", MessageBoxButton.OK, MessageBoxImage.Warning);
                                                                window = new MainWindow();
                                                                window.Close();
                                                                ExternalApplication.window = null;
                                                                SelectedElements.Clear();
                                                                uiDoc.Selection.SetElementIds(new List<ElementId> { ElementId.InvalidElementId });
                                                            }
                                                        }
                                                        else
                                                        {
                                                            System.Windows.MessageBox.Show("The conduits are at maximum distance or are unaligned", "Warning-AutoConnect", MessageBoxButton.OK, MessageBoxImage.Warning);
                                                            SelectedElements.Clear();
                                                            uiDoc.Selection.SetElementIds(new List<ElementId> { ElementId.InvalidElementId });
                                                        }
                                                    }
                                                }
                                            }
                                            //}
                                            //else
                                            //{
                                            //    System.Windows.MessageBox.Show("Please select the same number of conduits", "Warning", MessageBoxButton.OK, MessageBoxImage.Warning);
                                            //    ExternalApplication.window = null;
                                            //    SelectedElements.Clear();
                                            //    uiDoc.Selection.SetElementIds(new List<ElementId> { ElementId.InvalidElementId });
                                            //}
                                            // }
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
                else if (isDisabled) //Another Tool ON
                {
                    if (ToggleConPakToolsButton.ItemText == "AutoConnect ON")
                    {
                        ToggleConPakToolsButton.ItemText = "AutoConnect OFF";
                        ToggleConPakToolsButton.LargeImage = new BitmapImage(new Uri("pack://application:,,,/AutoConnectPro;component/Resources/Auto-Connect-32X32-red.png"));
                        ToggleConPakToolsButton.Image = new BitmapImage(new Uri("pack://application:,,,/AutoConnectPro;component/Resources/Auto-Connect-16X16-red.png"));
                        ToggleConPakToolsButtonSample.Enabled = true;
                    }
                }
            }
            catch (Exception)
            {
            }
        }
        private static List<Element> SortbyPlane(Document doc, List<Element> arrelements)
        {
            List<Element> conduitCollection = new List<Element>();

            //ascending conduits based on the intersection
            Dictionary<double, Conduit> dictcond = new Dictionary<double, Conduit>();
            View view = doc.ActiveView;
            XYZ vieworgin = view.Origin;
            XYZ viewdirection = view.ViewDirection;

            Line CondutitLine1 = (arrelements.First().Location as LocationCurve).Curve as Line;
            XYZ vieworgin1 = CondutitLine1.Origin;

            try
            {
                foreach (Conduit c in arrelements)
                {
                    conduitCollection.Clear();
                    Line CondutitLine = (c.Location as LocationCurve).Curve as Line;

                    SketchPlane sp = SketchPlane.Create(doc, Plane.CreateByNormalAndOrigin(CondutitLine1.Direction, vieworgin1));
                    double denominator = CondutitLine.Direction.Normalize().DotProduct(sp.GetPlane().Normal);
                    double numerator = (sp.GetPlane().Origin - CondutitLine.GetEndPoint(0)).DotProduct(sp.GetPlane().Normal);
                    double parameter = numerator / denominator;
                    XYZ intersectionPoint = CondutitLine.GetEndPoint(0) + parameter * CondutitLine.Direction;
                    double xdirection = Math.Round(CondutitLine.Direction.X, 6);
                    double ydirection = Math.Round(CondutitLine.Direction.Y, 6);
                    double zdirection = Math.Round(CondutitLine.Direction.Z, 6);

                    if (ydirection == -1 || ydirection == 1)
                    {
                        dictcond.Add(intersectionPoint.X, c);
                    }
                    else
                    {
                        dictcond.Add(intersectionPoint.Y, c);
                    }


                }
                conduitCollection = dictcond.OrderBy(x => x.Key).Select(x => x.Value as Element).ToList();
            }
            catch
            {
                conduitCollection = arrelements;
            }
            return conduitCollection;
        }
        private static Line AlignElement(Element pickedElement, XYZ refPoint, Document doc)
        {
            Line NewLine = null;

            Line firstLine = (pickedElement.Location as LocationCurve).Curve as Line;
            XYZ startPoint = firstLine.GetEndPoint(0);
            XYZ endPoint = firstLine.GetEndPoint(1);
            LocationCurve curve = pickedElement.Location as LocationCurve;
            XYZ normal = firstLine.Direction;
            XYZ cross = normal.CrossProduct(XYZ.BasisZ);
            XYZ newEndPoint = refPoint + cross.Multiply(5);
            Line boundLine = Line.CreateBound(refPoint, newEndPoint);
            XYZ interSecPoint = Utility.FindIntersectionPoint(firstLine.GetEndPoint(0), firstLine.GetEndPoint(1), boundLine.GetEndPoint(0), boundLine.GetEndPoint(1));
            interSecPoint = new XYZ(interSecPoint.X, interSecPoint.Y, startPoint.Z);
            ConnectorSet connectorSet = Utility.GetUnusedConnectors(pickedElement);
            if (connectorSet.Size == 2)
            {
                if (startPoint.DistanceTo(interSecPoint) > endPoint.DistanceTo(interSecPoint))
                {
                    NewLine = Line.CreateBound(startPoint, interSecPoint);
                }
                else
                {
                    NewLine = Line.CreateBound(interSecPoint, endPoint);
                }
            }
            else
            {
                connectorSet = Utility.GetConnectors(pickedElement);
                foreach (Connector con in connectorSet)
                {
                    if (con.IsConnected)
                    {
                        if (Utility.IsXYZTrue(con.Origin, startPoint))
                        {
                            NewLine = Line.CreateBound(con.Origin, interSecPoint);
                            break;
                        }
                        if (Utility.IsXYZTrue(con.Origin, endPoint))
                        {
                            NewLine = Line.CreateBound(interSecPoint, con.Origin);
                            break;
                        }
                    }
                }
            }
            return NewLine;
        }
        private bool ReverseingConduits(Document doc, ref List<Element> primaryElements, ref List<Element> secondaryElements)
        {
            Line priFirst = ((primaryElements.First().Location as LocationCurve).Curve as Line);
            Line prilast = ((primaryElements.Last().Location as LocationCurve).Curve as Line);
            Line secFirst = ((secondaryElements.First().Location as LocationCurve).Curve as Line);
            Line seclast = ((secondaryElements.Last().Location as LocationCurve).Curve as Line);

            XYZ firstinter = MultiConnectFindIntersectionPoint(priFirst, secFirst);
            XYZ lastinter = MultiConnectFindIntersectionPoint(prilast, seclast);
            if (firstinter == null || lastinter == null)
            {
                return false;
            }
            priFirst = AlignElement(primaryElements.First(), firstinter, doc);
            secFirst = AlignElement(secondaryElements.First(), firstinter, doc);
            prilast = AlignElement(primaryElements.Last(), lastinter, doc);
            seclast = AlignElement(secondaryElements.Last(), lastinter, doc);

            Line primFirstextentionline = Line.CreateBound(new XYZ(priFirst.GetEndPoint(0).X, priFirst.GetEndPoint(0).Y, 0), new XYZ(priFirst.GetEndPoint(1).X, priFirst.GetEndPoint(1).Y, 0));
            Line secoFirstnextentionline = Line.CreateBound(new XYZ(secFirst.GetEndPoint(0).X, secFirst.GetEndPoint(0).Y, 0), new XYZ(secFirst.GetEndPoint(1).X, secFirst.GetEndPoint(1).Y, 0));
            Line primLastextentionline = Line.CreateBound(new XYZ(prilast.GetEndPoint(0).X, prilast.GetEndPoint(0).Y, 0), new XYZ(prilast.GetEndPoint(1).X, prilast.GetEndPoint(1).Y, 0));
            Line secoLastnextentionline = Line.CreateBound(new XYZ(seclast.GetEndPoint(0).X, seclast.GetEndPoint(0).Y, 0), new XYZ(seclast.GetEndPoint(1).X, seclast.GetEndPoint(1).Y, 0));

            XYZ interpointset1 = Utility.GetIntersection(primFirstextentionline, secoLastnextentionline);
            XYZ interpointset2 = Utility.GetIntersection(secoFirstnextentionline, primLastextentionline);
            if (interpointset1 == null || interpointset2 == null)
            {
                secondaryElements.Reverse();
            }
            if (interpointset1 == null && interpointset2 == null)
            {
                primaryElements.Reverse();
            }
            return true;
        }
        public static XYZ MultiConnectFindIntersectionPoint(Line lineOne, Line lineTwo)
        {
            return MultiConnectFindIntersectionPoint(lineOne.GetEndPoint(0), lineOne.GetEndPoint(1), lineTwo.GetEndPoint(0), lineTwo.GetEndPoint(1));
        }
        public static XYZ MultiConnectFindIntersectionPoint(XYZ s1, XYZ e1, XYZ s2, XYZ e2)
        {
            s1 = Utility.XYZroundOf(s1, 5);
            e1 = Utility.XYZroundOf(e1, 5);
            s2 = Utility.XYZroundOf(s2, 5);
            e2 = Utility.XYZroundOf(e2, 5);

            double a1 = e1.Y - s1.Y;
            double b1 = s1.X - e1.X;
            double c1 = a1 * s1.X + b1 * s1.Y;

            double a2 = e2.Y - s2.Y;
            double b2 = s2.X - e2.X;
            double c2 = a2 * s2.X + b2 * s2.Y;

            double delta = a1 * b2 - a2 * b1;
            //If lines are parallel, the result will be (NaN, NaN).
            return delta == 0 || Convert.ToString(delta).Contains("E") == true ? null
                : new XYZ((b2 * c1 - b1 * c2) / delta, (a1 * c2 - a2 * c1) / delta, 0);
        }
        public static List<Element> GetElementsByOderDescending(List<Element> a_PrimaryElements)
        {
            List<Element> PrimaryElements = new List<Element>();
            XYZ PrimaryDirection = ((a_PrimaryElements.FirstOrDefault().Location as LocationCurve).Curve as Line).Direction;
            if (Math.Abs(PrimaryDirection.Z) != 1)
            {
                PrimaryElements = a_PrimaryElements.OrderByDescending(x => ((((x.Location as LocationCurve).Curve as Line).GetEndPoint(0) + ((x.Location as LocationCurve).Curve as Line).GetEndPoint(1)) / 2).Y).ToList();
                if (PrimaryDirection.Y == 1 || PrimaryDirection.Y == -1)
                {
                    PrimaryElements = a_PrimaryElements.OrderByDescending(x => ((((x.Location as LocationCurve).Curve as Line).GetEndPoint(0) + ((x.Location as LocationCurve).Curve as Line).GetEndPoint(1)) / 2).X).ToList();
                }
            }
            else
            {
                PrimaryElements = a_PrimaryElements.OrderByDescending(x => ((((x.Location as LocationCurve).Curve as Line).GetEndPoint(0) + ((x.Location as LocationCurve).Curve as Line).GetEndPoint(1)) / 2).X).ToList();
            }
            return PrimaryElements;
        }
        public void ThreePtSaddleExecute(UIApplication uiapp, ref List<Element> PrimaryElements, ref List<Element> SecondaryElements, XYZ pickpoint)
        {
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Application app = uiapp.Application;
            Document doc = uidoc.Document;
            int.TryParse(uiapp.Application.VersionNumber, out int RevitVersion);
            ElementsFilter filter = new ElementsFilter("Conduit Tags");

            try
            {

                //PrimaryElements = Elements.GetElementsByReference(PrimaryReference, doc);
                //SecondaryElements = Elemen    ts.GetElementsByReference(SecondaryReference, doc);
                LocationCurve findDirec = PrimaryElements[0].Location as LocationCurve;
                Line n = findDirec.Curve as Line;
                XYZ dire = n.Direction;
                XYZ MPtAngle = new XYZ();
                Connector ConnectOne = null;
                Connector ConnectTwo = null;
                Utility.GetClosestConnectors(PrimaryElements[0], SecondaryElements[0], out ConnectOne, out ConnectTwo);
                XYZ ax = ConnectOne.Origin;
                Line pickline = null;

                pickline = Line.CreateBound(pickpoint, pickpoint + new XYZ(dire.X + 10, dire.Y, dire.Z));
                if (dire.X == 1 || dire.X == -1)
                {
                    if (pickline.Origin.Y < ax.Y)
                    {
                        PrimaryElements = GetElementsByOderDescending(PrimaryElements);
                        SecondaryElements = GetElementsByOderDescending(SecondaryElements);
                    }
                    else
                    {
                        PrimaryElements = GetElementsByOder(PrimaryElements);
                        SecondaryElements = GetElementsByOder(SecondaryElements);
                    }
                }
                else if (dire.Y == -1 || dire.Y == 1)
                {

                    if (pickline.Origin.X < ax.X)
                    {
                        PrimaryElements = GetElementsByOderDescending(PrimaryElements);
                        SecondaryElements = GetElementsByOderDescending(SecondaryElements);
                    }
                    else
                    {
                        PrimaryElements = GetElementsByOder(PrimaryElements);
                        SecondaryElements = GetElementsByOder(SecondaryElements);
                    }
                }
                else
                {
                    if (pickline.Origin.X < ax.X)
                    {
                        if (dire.X == -1)
                        {
                            PrimaryElements = GetElementsByOderDescending(PrimaryElements);
                            SecondaryElements = GetElementsByOderDescending(SecondaryElements);
                        }
                        else
                        {
                            PrimaryElements = GetElementsByOder(PrimaryElements);
                            SecondaryElements = GetElementsByOder(SecondaryElements);
                        }
                    }
                    else
                    {
                        if (pickline.Origin.X < ax.X)
                        {

                            PrimaryElements = GetElementsByOderDescending(PrimaryElements);
                            SecondaryElements = GetElementsByOderDescending(SecondaryElements);
                        }
                        else
                        {
                            PrimaryElements = GetElementsByOder(PrimaryElements);
                            SecondaryElements = GetElementsByOder(SecondaryElements);
                        }
                    }
                }


                List<Element> thirdElements = new List<Element>();
                List<Element> forthElements = new List<Element>();
                bool isVerticalConduits = false;
                // Modify document within a transaction
                try
                {
                    using (Transaction tx = new Transaction(doc))
                    {
                        ConnectorSet PrimaryConnectors = null;
                        ConnectorSet SecondaryConnectors = null;
                        Connector ConnectorOne = null;
                        Connector ConnectorTwo = null;

                        tx.Start("Threesaddle");
                        double l_angle = Convert.ToDouble(MainWindow.Instance.angleDegree) * (Math.PI / 180);
                        double givendist = 0;
                        for (int i = 0; i < PrimaryElements.Count; i++)
                        {
                            List<XYZ> ConnectorPoints = new List<XYZ>();
                            PrimaryConnectors = Utility.GetConnectors(PrimaryElements[i]);
                            SecondaryConnectors = Utility.GetConnectors(SecondaryElements[i]);
                            Utility.GetClosestConnectors(PrimaryConnectors, SecondaryConnectors, out ConnectorOne, out ConnectorTwo);
                            foreach (Connector con in PrimaryConnectors)
                            {
                                ConnectorPoints.Add(con.Origin);
                            }
                            Element el = PrimaryElements[0];
                            LocationCurve findDirect = el.Location as LocationCurve;
                            Line ncDer = findDirect.Curve as Line;
                            XYZ dir = ncDer.Direction;
                            XYZ newenpt = new XYZ(ConnectorTwo.Origin.X, ConnectorTwo.Origin.Y, ConnectorOne.Origin.Z);
                            XYZ newenpt2 = new XYZ(ConnectorOne.Origin.X, ConnectorOne.Origin.Y, ConnectorTwo.Origin.Z);

                            Conduit newConCopy = Utility.CreateConduit(doc, PrimaryElements[i] as Conduit, ConnectorOne.Origin, newenpt);
                            Conduit newCon2Copy = Utility.CreateConduit(doc, SecondaryElements[i] as Conduit, ConnectorTwo.Origin, newenpt2);
                            string offsetVariable = RevitVersion < 2020 ? "Offset" : "Middle Elevation";
                            Parameter parameter = newConCopy.LookupParameter(offsetVariable);
                            var middle = parameter.AsDouble();
                            XYZ Pri_mid = Utility.GetMidPoint(newConCopy);
                            XYZ Sec_mid = Utility.GetMidPoint(newCon2Copy);

                            double distance = 0;
                            DistanceElements = Utility.ConduitInOrder(DistanceElements);
                            LocationCurve newcurve = DistanceElements[0].Location as LocationCurve;
                            Line ncl = newcurve.Curve as Line;
                            XYZ direc = ncl.Direction;
                            if (DistanceElements.Count() >= 2)
                            {
                                LocationCurve newcurve2 = DistanceElements[1].Location as LocationCurve;
                                XYZ start1 = Utility.GetMidPoint(DistanceElements[0]);
                                Line cross = Utility.CrossProductLine(ncl, start1, 5, true);
                                start1 = new XYZ(start1.X, start1.Y, 0);
                                XYZ start2 = Utility.FindIntersection(DistanceElements[1], cross);
                                distance = start1.DistanceTo(start2);
                            }
                            Conduit newCon = null;
                            Conduit newCon2 = null;
                            var l = Utility.GetLineFromConduit(newConCopy);
                            MPtAngle = Utility.GetMidPoint(l);

                            if (IsHorizontal(ncDer))
                            {
                                if (pickline.Origin.Y < ax.Y)
                                {
                                    if (i == 0)
                                    {
                                        newCon = Utility.CreateConduit(doc, PrimaryElements[i] as Conduit, pickpoint, pickpoint + new XYZ(direc.X + 3, direc.Y, direc.Z));
                                        newCon2 = Utility.CreateConduit(doc, SecondaryElements[i] as Conduit, pickpoint, pickpoint + new XYZ(direc.X - 3, direc.Y, direc.Z));
                                    }
                                    else
                                    {
                                        givendist += distance;
                                        newCon = Utility.CreateConduit(doc, PrimaryElements[i] as Conduit, new XYZ(pickpoint.X, pickpoint.Y - givendist, pickpoint.Z), pickpoint + new XYZ(direc.X + 3, direc.Y - givendist, direc.Z));
                                        newCon2 = Utility.CreateConduit(doc, SecondaryElements[i] as Conduit, new XYZ(pickpoint.X, pickpoint.Y - givendist, pickpoint.Z), pickpoint + new XYZ(direc.X - 3, direc.Y - givendist, direc.Z));
                                    }
                                }
                                else
                                {
                                    if (i == 0)
                                    {
                                        newCon = Utility.CreateConduit(doc, PrimaryElements[i] as Conduit, pickpoint, pickpoint + new XYZ(direc.X + 3, direc.Y, direc.Z));
                                        newCon2 = Utility.CreateConduit(doc, SecondaryElements[i] as Conduit, pickpoint, pickpoint + new XYZ(direc.X - 3, direc.Y, direc.Z));

                                    }
                                    else
                                    {
                                        givendist += distance;
                                        newCon = Utility.CreateConduit(doc, PrimaryElements[i] as Conduit, new XYZ(pickpoint.X, pickpoint.Y + givendist, pickpoint.Z), pickpoint + new XYZ(direc.X + 3, direc.Y + givendist, direc.Z));
                                        newCon2 = Utility.CreateConduit(doc, SecondaryElements[i] as Conduit, new XYZ(pickpoint.X, pickpoint.Y + givendist, pickpoint.Z), pickpoint + new XYZ(direc.X - 3, direc.Y + givendist, direc.Z));
                                    }
                                }
                            }
                            else if (IsVertical(ncDer)) //vertical
                            {
                                if (pickline.Origin.X < ax.X)
                                {
                                    if (i == 0)
                                    {
                                        newCon = Utility.CreateConduit(doc, PrimaryElements[i] as Conduit, pickpoint, pickpoint + direc.Multiply(.5));
                                        newCon2 = Utility.CreateConduit(doc, SecondaryElements[i] as Conduit, pickpoint, pickpoint + direc.Multiply(-.5));
                                    }
                                    else
                                    {

                                        givendist += distance;
                                        newCon = Utility.CreateConduit(doc, PrimaryElements[i] as Conduit, new XYZ(pickpoint.X - givendist, pickpoint.Y, pickpoint.Z), pickpoint + new XYZ(direc.X - givendist, direc.Y - .5, direc.Z));////
                                        newCon2 = Utility.CreateConduit(doc, SecondaryElements[i] as Conduit, new XYZ(pickpoint.X - givendist, pickpoint.Y, pickpoint.Z), pickpoint + new XYZ(direc.X - givendist, direc.Y - .5, direc.Z));
                                    }
                                }
                                else //right
                                {
                                    if (i == 0)
                                    {
                                        newCon = Utility.CreateConduit(doc, PrimaryElements[i] as Conduit, pickpoint, pickpoint + direc.Multiply(.5));
                                        newCon2 = Utility.CreateConduit(doc, SecondaryElements[i] as Conduit, pickpoint, pickpoint + direc.Multiply(-.5));
                                    }
                                    else
                                    {

                                        givendist += distance;
                                        newCon = Utility.CreateConduit(doc, PrimaryElements[i] as Conduit, new XYZ(pickpoint.X + givendist, pickpoint.Y, pickpoint.Z), pickpoint + new XYZ(direc.X + givendist, direc.Y + .5, direc.Z));
                                        newCon2 = Utility.CreateConduit(doc, SecondaryElements[i] as Conduit, new XYZ(pickpoint.X + givendist, pickpoint.Y, pickpoint.Z), pickpoint + new XYZ(direc.X + givendist, direc.Y - .5, direc.Z));
                                    }
                                }
                            }
                            else //angled
                            {
                                if (dir.X > 0)
                                {
                                    if (pickline.Origin.X < MPtAngle.X) //left
                                    {
                                        if (i == 0)
                                        {
                                            newCon = Utility.CreateConduit(doc, PrimaryElements[i] as Conduit, pickpoint, pickpoint + new XYZ(direc.X, direc.Y, direc.Z));
                                            newCon2 = Utility.CreateConduit(doc, SecondaryElements[i] as Conduit, pickpoint, pickpoint + new XYZ(direc.X, direc.Y, direc.Z));
                                        }
                                        else
                                        {

                                            givendist += distance;
                                            newCon = Utility.CreateConduit(doc, PrimaryElements[i] as Conduit, new XYZ(pickpoint.X + givendist, pickpoint.Y, pickpoint.Z), pickpoint + new XYZ(direc.X + givendist, direc.Y, direc.Z));
                                            newCon2 = Utility.CreateConduit(doc, SecondaryElements[i] as Conduit, new XYZ(pickpoint.X + givendist, pickpoint.Y, pickpoint.Z), pickpoint + new XYZ(direc.X + givendist, direc.Y, direc.Z));
                                        }
                                    }
                                    else //right
                                    {
                                        if (i == 0)
                                        {
                                            XYZ end = pickpoint + new XYZ(direc.X, direc.Y, direc.Z);
                                            newCon = Utility.CreateConduit(doc, PrimaryElements[i] as Conduit, pickpoint, pickpoint + new XYZ(direc.X, direc.Y, direc.Z));
                                            newCon2 = Utility.CreateConduit(doc, SecondaryElements[i] as Conduit, pickpoint, end);
                                        }
                                        else
                                        {

                                            givendist += distance;
                                            newCon = Utility.CreateConduit(doc, PrimaryElements[i] as Conduit, new XYZ(pickpoint.X + givendist, pickpoint.Y, pickpoint.Z), pickpoint + new XYZ(direc.X + givendist, direc.Y, direc.Z));////
                                            newCon2 = Utility.CreateConduit(doc, SecondaryElements[i] as Conduit, new XYZ(pickpoint.X + givendist, pickpoint.Y, pickpoint.Z), pickpoint + new XYZ(direc.X + givendist, direc.Y, direc.Z));
                                        }
                                    }
                                }
                                else
                                {
                                    if (pickline.Origin.X < MPtAngle.X) //left
                                    {
                                        if (i == 0)
                                        {
                                            newCon = Utility.CreateConduit(doc, PrimaryElements[i] as Conduit, pickpoint, pickpoint + new XYZ(direc.X, direc.Y, direc.Z));
                                            newCon2 = Utility.CreateConduit(doc, SecondaryElements[i] as Conduit, pickpoint, pickpoint + new XYZ(direc.X, direc.Y, direc.Z));
                                        }
                                        else
                                        {

                                            givendist += distance;
                                            newCon = Utility.CreateConduit(doc, PrimaryElements[i] as Conduit, new XYZ(pickpoint.X + givendist, pickpoint.Y, pickpoint.Z), pickpoint + new XYZ(direc.X + givendist, direc.Y, direc.Z));
                                            newCon2 = Utility.CreateConduit(doc, SecondaryElements[i] as Conduit, new XYZ(pickpoint.X + givendist, pickpoint.Y, pickpoint.Z), pickpoint + new XYZ(direc.X + givendist, direc.Y, direc.Z));
                                        }
                                    }
                                    else //right
                                    {
                                        if (i == 0)
                                        {
                                            XYZ end = pickpoint + new XYZ(direc.X, direc.Y, direc.Z);
                                            newCon = Utility.CreateConduit(doc, PrimaryElements[i] as Conduit, pickpoint, pickpoint + new XYZ(direc.X, direc.Y, direc.Z));
                                            newCon2 = Utility.CreateConduit(doc, SecondaryElements[i] as Conduit, pickpoint, end);
                                        }
                                        else
                                        {

                                            givendist += distance;
                                            newCon = Utility.CreateConduit(doc, PrimaryElements[i] as Conduit, new XYZ(pickpoint.X + givendist, pickpoint.Y, pickpoint.Z), pickpoint + new XYZ(direc.X + givendist, direc.Y, direc.Z));////
                                            newCon2 = Utility.CreateConduit(doc, SecondaryElements[i] as Conduit, new XYZ(pickpoint.X + givendist, pickpoint.Y, pickpoint.Z), pickpoint + new XYZ(direc.X + givendist, direc.Y, direc.Z));
                                        }
                                    }
                                }
                            }


                            Parameter param = newCon.LookupParameter("Middle Elevation");
                            Parameter param2 = newCon2.LookupParameter("Middle Elevation");

                            param.Set(middle);
                            param2.Set(middle);

                            Utility.DeleteElement(doc, newConCopy.Id);
                            Utility.DeleteElement(doc, newCon2Copy.Id);

                            if (Utility.IsXYTrue(ConnectorPoints.FirstOrDefault(), ConnectorPoints.LastOrDefault()))
                            {
                                isVerticalConduits = true;
                            }
                            Element e = doc.GetElement(newCon.Id);
                            Element e2 = doc.GetElement(newCon2.Id);
                            thirdElements.Add(e);
                            forthElements.Add(e2);
                            //RetainParameters(PrimaryElements[i], SecondaryElements[i]);
                            //RetainParameters(PrimaryElements[i], e);
                        }
                        //Rotate Elements at Once
                        Element ElementOne = PrimaryElements[0];
                        Element ElementTwo = SecondaryElements[0];
                        Utility.GetClosestConnectors(ElementOne, ElementTwo, out ConnectorOne, out ConnectorTwo);
                        LocationCurve findDirection = ElementOne.Location as LocationCurve;
                        Line nc = findDirection.Curve as Line;
                        Curve refcurve = findDirection.Curve;
                        XYZ direct = nc.Direction;
                        LocationCurve findDirection2 = ElementTwo.Location as LocationCurve;
                        Line nc2 = findDirection2.Curve as Line;
                        XYZ directDown = nc2.Direction;
                        //primary
                        LocationCurve newconcurve = thirdElements[0].Location as LocationCurve;
                        Line ncl1 = newconcurve.Curve as Line;
                        XYZ direction = ncl1.Direction;
                        XYZ axisStart = null;
                        axisStart = pickpoint;
                        XYZ axisSt = ConnectorOne.Origin;
                        XYZ axisEnd = axisStart.Add(XYZ.BasisZ.CrossProduct(direction));
                        Line axisLine = Line.CreateBound(axisStart, axisEnd);
                        //secondary
                        LocationCurve newconcurve2 = forthElements[0].Location as LocationCurve;
                        Line ncl2 = newconcurve2.Curve as Line;
                        XYZ direction2 = ncl2.Direction;
                        XYZ axisStart2 = null;
                        axisStart2 = pickpoint;
                        XYZ axisEnd2 = axisStart2.Add(XYZ.BasisZ.CrossProduct(direction2));
                        Line axisLine2 = Line.CreateBound(axisStart2, axisEnd2);
                        Line pickedline = null;

                        pickedline = Line.CreateBound(pickpoint, pickpoint + new XYZ(direction.X + 10, direction.Y, direction.Z));
                        Curve cu = pickedline as Curve;

                        double PrimaryOffset = RevitVersion < 2020 ? PrimaryElements[0].LookupParameter("Offset").AsDouble() :
                                                 PrimaryElements[0].LookupParameter("Middle Elevation").AsDouble();
                        double SecondaryOffset = RevitVersion < 2020 ? SecondaryElements[0].LookupParameter("Offset").AsDouble() :
                                                  SecondaryElements[0].LookupParameter("Middle Elevation").AsDouble();
                        if (isVerticalConduits)
                        {
                            l_angle = (Math.PI / 2) - l_angle;
                        }
                        if (PrimaryOffset > SecondaryOffset)
                        {
                            //rotate down
                            l_angle = -l_angle;
                        }
                        try
                        {
                            if (IsHorizontal(nc))
                            {
                                if (pickedline.Origin.Y < axisSt.Y)
                                {
                                    XYZ end = new XYZ(axisStart.X, axisStart.Y, axisStart.Z + 10);
                                    Line l1 = Line.CreateBound(axisStart, end);
                                    ElementTransformUtils.RotateElements(doc, thirdElements.Select(r => r.Id).ToList(), l1, l_angle);
                                }
                                else
                                {
                                    XYZ end = new XYZ(axisStart.X, axisStart.Y, axisStart.Z - 10);
                                    Line l1 = Line.CreateBound(axisStart, end);
                                    ElementTransformUtils.RotateElements(doc, thirdElements.Select(r => r.Id).ToList(), l1, l_angle);
                                }
                            }
                            else if (IsVertical(nc))
                            {
                                if (pickedline.Origin.X < axisSt.X)
                                {
                                    //left in vertical
                                    XYZ end = new XYZ(axisStart.X, axisStart.Y, axisStart.Z - 10);
                                    Line l1 = Line.CreateBound(axisStart, end);
                                    ElementTransformUtils.RotateElements(doc, thirdElements.Select(r => r.Id).ToList(), l1, -l_angle);
                                }
                                else
                                {
                                    //right in vertical
                                    XYZ end = new XYZ(axisStart.X, axisStart.Y, axisStart.Z + 10);/////////////////////
                                    Line l1 = Line.CreateBound(axisStart, end);
                                    ElementTransformUtils.RotateElements(doc, thirdElements.Select(r => r.Id).ToList(), l1, -l_angle);
                                }
                            }
                            else //angle conduit
                            {
                                if (pickedline.Origin.X < axisSt.X) //left
                                {
                                    XYZ end = new XYZ(axisStart.X, axisStart.Y, axisStart.Z - 10);
                                    Line l1 = Line.CreateBound(axisStart, end);
                                    ElementTransformUtils.RotateElements(doc, thirdElements.Select(r => r.Id).ToList(), l1, -l_angle);
                                }
                                else //right
                                {
                                    XYZ end = new XYZ(axisStart.X, axisStart.Y, axisStart.Z + 10);
                                    Line l1 = Line.CreateBound(axisStart, end);
                                    ElementTransformUtils.RotateElements(doc, thirdElements.Select(r => r.Id).ToList(), l1, -l_angle);
                                }
                            }
                            for (int i = 0; i < PrimaryElements.Count; i++)
                            {
                                Element firstElement = PrimaryElements[i];
                                Element secondElement = SecondaryElements[i];
                                Element thirdElement = thirdElements[i];
                                Element forthElement = forthElements[i];
                                Utility.AutoRetainParameters(PrimaryElements[i], SecondaryElements[i], uidoc, uiapp);
                                Utility.AutoRetainParameters(PrimaryElements[i], thirdElement, doc, uiapp);

                                if (IsVertical(refcurve))
                                {
                                    if (pickline.Origin.X < ax.X) //left
                                    {
                                        if (direct.Y == -1 && directDown.Y == 1)
                                            Utility.CreateElbowFittings(PrimaryElements[i], thirdElement, doc, uiapp);
                                        else if (direct.Y == 1 && directDown.Y == 1)
                                            Utility.CreateElbowFittings(PrimaryElements[i], thirdElement, doc, uiapp);
                                        else if (direct.Y == -1 && directDown.Y == -1)
                                        {
                                            //Utility.CreateElbowFittings(PrimaryElements[i], thirdElement, doc, uiapp);
                                            Utility.CreateElbowFittings(secondElement, thirdElement, doc, uiapp);
                                        }
                                        else if (direct.Y == 1 && directDown.Y == -1)
                                        {
                                            if (direct.X < 0 || directDown.X < 0)
                                                Utility.CreateElbowFittings(PrimaryElements[i], thirdElement, doc, uiapp);
                                            else
                                                Utility.CreateElbowFittings(secondElement, thirdElement, doc, uiapp);
                                        }
                                        else
                                            Utility.CreateElbowFittings(secondElement, thirdElement, doc, uiapp);
                                    }
                                    else
                                    {
                                        if (direct.Y == -1 && directDown.Y == 1)
                                            Utility.CreateElbowFittings(PrimaryElements[i], thirdElement, doc, uiapp);
                                        else if (direct.Y == 1 && directDown.Y == 1)
                                            Utility.CreateElbowFittings(PrimaryElements[i], thirdElement, doc, uiapp);
                                        else if (direct.Y == -1 && directDown.Y == -1)
                                        {
                                            //Utility.CreateElbowFittings(PrimaryElements[i], thirdElement, doc, uiapp);
                                            Utility.CreateElbowFittings(secondElement, thirdElement, doc, uiapp);
                                        }
                                        else if (direct.Y == 1 && directDown.Y == -1)
                                        {
                                            if (direct.X < 0 || directDown.X < 0)
                                                Utility.CreateElbowFittings(secondElement, thirdElement, doc, uiapp);
                                            else
                                                Utility.CreateElbowFittings(PrimaryElements[i], thirdElement, doc, uiapp);
                                        }
                                        else
                                            Utility.CreateElbowFittings(secondElement, thirdElement, doc, uiapp);
                                    }

                                }
                                else if (Math.Round(direct.X) == Math.Round(directDown.X) && Math.Round(direct.Y) == Math.Round(directDown.Y) && direct.X != 1 && direct.Y != 1 && direct.X != -1 && direct.Y != -1)
                                {
                                    if (direct.X > 0 && directDown.X > 0)
                                    {
                                        if (pickedline.Origin.X < MPtAngle.X) //left
                                        {
                                            //Utility.CreateElbowFittings(secondElement, thirdElement, doc, uiapp);
                                            Utility.CreateElbowFittings(PrimaryElements[i], thirdElement, doc, uiapp);
                                        }
                                        else //right
                                        {
                                            if (direct.Z == 0)
                                                Utility.CreateElbowFittings(secondElement, thirdElement, doc, uiapp);
                                            else
                                                Utility.CreateElbowFittings(PrimaryElements[i], thirdElement, doc, uiapp);
                                        }
                                    }
                                    else
                                    {
                                        if (pickedline.Origin.X < MPtAngle.X) //left
                                        {
                                            Utility.CreateElbowFittings(secondElement, thirdElement, doc, uiapp);
                                        }
                                        else //right
                                        {
                                            Utility.CreateElbowFittings(PrimaryElements[i], thirdElement, doc, uiapp);
                                        }
                                    }
                                }
                                else //horizontal
                                {
                                    if (pickedline.Origin.Y > axisSt.Y)
                                    {
                                        if (direct.X == -1 && directDown.X == 1)
                                        {
                                            if (direct.Y < 0 || directDown.Y < 0)
                                                Utility.CreateElbowFittings(PrimaryElements[i], thirdElement, doc, uiapp);
                                            else
                                                Utility.CreateElbowFittings(secondElement, thirdElement, doc, uiapp);
                                        }
                                        else if (direct.X == 1 && directDown.X == 1)
                                        {
                                            if (Math.Round(direct.Y) == 0)
                                                Utility.CreateElbowFittings(secondElement, thirdElement, doc, uiapp);
                                            else
                                                Utility.CreateElbowFittings(secondElement, thirdElement, doc, uiapp);
                                        }
                                        else if (direct.X == -1 && directDown.X == -1)
                                        {
                                            if (Math.Round(direct.Y, 2) == Math.Round(directDown.Y, 2))
                                                Utility.CreateElbowFittings(PrimaryElements[i], thirdElement, doc, uiapp);
                                            else
                                                Utility.CreateElbowFittings(secondElement, thirdElement, doc, uiapp);
                                        }
                                        else
                                            Utility.CreateElbowFittings(secondElement, thirdElement, doc, uiapp);
                                    }
                                    else
                                    {
                                        if (direct.X == -1 && directDown.X == 1)
                                        {
                                            if (direct.Y < 0 || directDown.Y < 0)
                                                Utility.CreateElbowFittings(PrimaryElements[i], thirdElement, doc, uiapp);
                                            else
                                                Utility.CreateElbowFittings(secondElement, thirdElement, doc, uiapp);
                                        }
                                        else if (direct.X == 1 && directDown.X == 1)
                                            Utility.CreateElbowFittings(secondElement, thirdElement, doc, uiapp);
                                        else if (direct.X == -1 && directDown.X == -1)
                                        {
                                            if (Math.Round(direct.Y, 2) == Math.Round(directDown.Y, 2))
                                                Utility.CreateElbowFittings(PrimaryElements[i], thirdElement, doc, uiapp);
                                            else
                                                Utility.CreateElbowFittings(secondElement, thirdElement, doc, uiapp);
                                        }
                                        else
                                            Utility.CreateElbowFittings(PrimaryElements[i], thirdElement, doc, uiapp);
                                    }
                                }
                            }

                            if (IsHorizontal(nc))
                            {
                                if (pickedline.Origin.Y < axisSt.Y)
                                {
                                    XYZ end2 = new XYZ(axisStart2.X, axisStart2.Y, axisStart2.Z + 10);
                                    Line l2 = Line.CreateBound(axisStart2, end2);
                                    ElementTransformUtils.RotateElements(doc, forthElements.Select(r => r.Id).ToList(), l2, -l_angle);
                                }

                                else
                                {
                                    XYZ end = new XYZ(axisStart2.X, axisStart2.Y, axisStart2.Z - 10);
                                    Line l1 = Line.CreateBound(axisStart2, end);
                                    ElementTransformUtils.RotateElements(doc, forthElements.Select(r => r.Id).ToList(), l1, -l_angle);
                                }
                            }
                            else if (IsVertical(nc))
                            {
                                if (pickedline.Origin.X < axisSt.X)
                                {
                                    //left in vertical
                                    XYZ end2 = new XYZ(axisStart2.X, axisStart2.Y, axisStart2.Z - 10);
                                    Line l2 = Line.CreateBound(axisStart2, end2);
                                    ElementTransformUtils.RotateElements(doc, forthElements.Select(r => r.Id).ToList(), l2, l_angle);
                                }

                                else
                                {
                                    //right in vertical
                                    XYZ end = new XYZ(axisStart2.X, axisStart2.Y, axisStart2.Z + 10);/////////////
                                    Line l1 = Line.CreateBound(axisStart2, end);
                                    ElementTransformUtils.RotateElements(doc, forthElements.Select(r => r.Id).ToList(), l1, l_angle);
                                }
                            }
                            else
                            {
                                if (pickedline.Origin.X < axisSt.X) //left
                                {
                                    XYZ end = new XYZ(axisStart2.X, axisStart2.Y, axisStart2.Z - 10);
                                    Line l1 = Line.CreateBound(axisStart2, end);
                                    ElementTransformUtils.RotateElements(doc, forthElements.Select(r => r.Id).ToList(), l1, l_angle);
                                }
                                else //right
                                {

                                    XYZ end2 = new XYZ(axisStart2.X, axisStart2.Y, axisStart2.Z - 10);
                                    Line l2 = Line.CreateBound(axisStart2, end2);
                                    ElementTransformUtils.RotateElements(doc, forthElements.Select(r => r.Id).ToList(), l2, -l_angle);
                                }
                            }



                            for (int i = 0; i < SecondaryElements.Count; i++)
                            {
                                Element firstElement = PrimaryElements[i];
                                Element secondElement = SecondaryElements[i];
                                Element thirdElement = thirdElements[i];
                                Element forthElement = forthElements[i];
                                Utility.AutoRetainParameters(PrimaryElements[i], SecondaryElements[i], uidoc, uiapp);
                                Utility.AutoRetainParameters(PrimaryElements[i], thirdElement, doc, uiapp);

                                if (IsVertical(refcurve))
                                {
                                    if (pickline.Origin.X < ax.X) //left
                                    {
                                        if (direct.Y == -1 && directDown.Y == 1)
                                            Utility.CreateElbowFittings(SecondaryElements[i], forthElement, doc, uiapp);
                                        else if (direct.Y == 1 && directDown.Y == 1)
                                            Utility.CreateElbowFittings(SecondaryElements[i], forthElement, doc, uiapp);
                                        else if (direct.Y == -1 && directDown.Y == -1)
                                        {
                                            //Utility.CreateElbowFittings(SecondaryElements[i], forthElement, doc, uiapp);
                                            Utility.CreateElbowFittings(firstElement, forthElement, doc, uiapp);
                                        }
                                        else if (direct.Y == 1 && directDown.Y == -1)
                                        {
                                            if (direct.X < 0 || directDown.X < 0)
                                                Utility.CreateElbowFittings(SecondaryElements[i], forthElement, doc, uiapp);
                                            else
                                                Utility.CreateElbowFittings(firstElement, forthElement, doc, uiapp);
                                        }
                                        else
                                            Utility.CreateElbowFittings(firstElement, forthElement, doc, uiapp);
                                    }
                                    else
                                    {
                                        if (direct.Y == -1 && directDown.Y == 1)
                                            Utility.CreateElbowFittings(SecondaryElements[i], forthElement, doc, uiapp);
                                        else if (direct.Y == 1 && directDown.Y == 1)
                                            Utility.CreateElbowFittings(SecondaryElements[i], forthElement, doc, uiapp);
                                        else if (direct.Y == -1 && directDown.Y == -1)
                                        {
                                            //Utility.CreateElbowFittings(SecondaryElements[i], forthElement, doc, uiapp);
                                            Utility.CreateElbowFittings(firstElement, forthElement, doc, uiapp);
                                        }
                                        else if (direct.Y == 1 && directDown.Y == -1)
                                        {
                                            if (direct.X < 0 || directDown.X < 0)
                                                Utility.CreateElbowFittings(firstElement, forthElement, doc, uiapp);
                                            else
                                                Utility.CreateElbowFittings(SecondaryElements[i], forthElement, doc, uiapp);
                                        }
                                        else
                                            Utility.CreateElbowFittings(firstElement, forthElement, doc, uiapp);
                                    }

                                }
                                else if (Math.Round(direct.X) == Math.Round(directDown.X) && Math.Round(direct.Y) == Math.Round(directDown.Y) && direct.X != 1 && direct.Y != 1 && direct.X != -1 && direct.Y != -1)
                                {
                                    if (direct.X > 0 && directDown.X > 0)
                                    {
                                        if (pickedline.Origin.X < MPtAngle.X) //left
                                        {
                                            //Utility.CreateElbowFittings(firstElement, forthElement, doc, uiapp);
                                            Utility.CreateElbowFittings(SecondaryElements[i], forthElement, doc, uiapp);
                                        }
                                        else //right
                                        {
                                            if (direct.Z == 0)
                                                Utility.CreateElbowFittings(firstElement, forthElement, doc, uiapp);
                                            else
                                                Utility.CreateElbowFittings(SecondaryElements[i], forthElement, doc, uiapp);
                                        }
                                    }
                                    else
                                    {
                                        if (pickedline.Origin.X < MPtAngle.X) //left
                                        {
                                            Utility.CreateElbowFittings(firstElement, forthElement, doc, uiapp);
                                        }
                                        else //right
                                        {
                                            Utility.CreateElbowFittings(SecondaryElements[i], forthElement, doc, uiapp);
                                        }
                                    }
                                }
                                else
                                {
                                    if (pickedline.Origin.Y > axisSt.Y)
                                    {
                                        if (direct.X == -1 && directDown.X == 1)
                                        {
                                            if (direct.Y < 0 || directDown.Y < 0)
                                                Utility.CreateElbowFittings(SecondaryElements[i], forthElement, doc, uiapp);
                                            else
                                                Utility.CreateElbowFittings(firstElement, forthElement, doc, uiapp);
                                        }
                                        else if (direct.X == 1 && directDown.X == 1)
                                        {
                                            if (Math.Round(direct.Y) == 0)
                                                Utility.CreateElbowFittings(firstElement, forthElement, doc, uiapp);
                                            else
                                                Utility.CreateElbowFittings(firstElement, forthElement, doc, uiapp);
                                        }
                                        else if (direct.X == -1 && directDown.X == -1)
                                        {
                                            if (Math.Round(direct.Y, 2) == Math.Round(directDown.Y, 2))
                                                Utility.CreateElbowFittings(SecondaryElements[i], forthElement, doc, uiapp);
                                            else
                                                Utility.CreateElbowFittings(firstElement, forthElement, doc, uiapp);
                                        }
                                        else
                                            Utility.CreateElbowFittings(firstElement, forthElement, doc, uiapp);
                                    }
                                    else
                                    {

                                        if (direct.X == -1 && directDown.X == 1)
                                        {
                                            if (direct.Y < 0 || directDown.Y < 0)
                                                Utility.CreateElbowFittings(SecondaryElements[i], forthElement, doc, uiapp);
                                            else
                                                Utility.CreateElbowFittings(firstElement, forthElement, doc, uiapp);
                                        }
                                        else if (direct.X == 1 && directDown.X == 1)
                                            Utility.CreateElbowFittings(firstElement, forthElement, doc, uiapp);
                                        else if (direct.X == -1 && directDown.X == -1)
                                        {
                                            if (Math.Round(direct.Y, 2) == Math.Round(directDown.Y, 2))
                                                Utility.CreateElbowFittings(SecondaryElements[i], forthElement, doc, uiapp);
                                            else
                                                Utility.CreateElbowFittings(firstElement, forthElement, doc, uiapp);
                                        }
                                        else
                                            Utility.CreateElbowFittings(SecondaryElements[i], forthElement, doc, uiapp);
                                    }
                                }
                            }
                            for (int i = 0; i < thirdElements.Count; i++)
                            {
                                try
                                {
                                    Utility.CreateElbowFittings(thirdElements[i], forthElements[i], doc, uiapp);
                                }
                                catch (Exception)
                                {
                                    try
                                    {
                                        Utility.CreateElbowFittings(forthElements[i], thirdElements[i], doc, uiapp);
                                    }
                                    catch (Exception ex)
                                    {
                                        MessageBox.Show(ex.Message);
                                        return;
                                    }
                                }
                            }


                        }
                        catch (Exception)
                        {

                            ElementTransformUtils.RotateElements(doc, thirdElements.Select(r => r.Id).ToList(), axisLine, l_angle * 2 + Math.PI);

                            for (int i = 0; i < PrimaryElements.Count; i++)
                            {
                                Element firstElement = PrimaryElements[i];
                                Element secondElement = SecondaryElements[i];
                                Element thirdElement = thirdElements[i];
                                Utility.AutoRetainParameters(PrimaryElements[i], SecondaryElements[i], uidoc, uiapp);
                                Utility.AutoRetainParameters(PrimaryElements[i], thirdElement, doc, uiapp);
                                //Utility.CreateElbowFittings(PrimaryElements[i], thirdElement, doc, uiapp);
                                //Utility.CreateElbowFittings(thirdElement, SecondaryElements[i], doc, uiapp);
                                try
                                {
                                    _deleteElements.Add(CreateElbowFittings(PrimaryElements[i], thirdElement, doc, uiapp));
                                    _deleteElements.Add(CreateElbowFittings(thirdElement, SecondaryElements[i], doc, uiapp));
                                }
                                catch
                                {
                                    try
                                    {

                                        _deleteElements.Add(CreateElbowFittings(thirdElement, PrimaryElements[i], doc, uiapp));
                                        _deleteElements.Add(CreateElbowFittings(SecondaryElements[i], thirdElement, doc, uiapp));
                                    }
                                    catch
                                    {

                                        try
                                        {

                                            _deleteElements.Add(CreateElbowFittings(SecondaryElements[i], thirdElement, doc, uiapp));
                                            _deleteElements.Add(CreateElbowFittings(thirdElement, PrimaryElements[i], doc, uiapp));
                                        }
                                        catch
                                        {
                                            try
                                            {

                                                _deleteElements.Add(CreateElbowFittings(thirdElement, SecondaryElements[i], doc, uiapp));
                                                _deleteElements.Add(CreateElbowFittings(PrimaryElements[i], thirdElement, doc, uiapp));

                                            }
                                            catch
                                            {
                                                try
                                                {

                                                    _deleteElements.Add(CreateElbowFittings(PrimaryElements[i], thirdElement, doc, uiapp));
                                                    _deleteElements.Add(CreateElbowFittings(SecondaryElements[i], thirdElement, doc, uiapp));

                                                }
                                                catch
                                                {
                                                    try
                                                    {

                                                        _deleteElements.Add(CreateElbowFittings(SecondaryElements[i], thirdElement, doc, uiapp));
                                                        _deleteElements.Add(CreateElbowFittings(PrimaryElements[i], thirdElement, doc, uiapp));
                                                    }
                                                    catch
                                                    {
                                                        try
                                                        {

                                                            _deleteElements.Add(CreateElbowFittings(thirdElement, SecondaryElements[i], doc, uiapp));
                                                            _deleteElements.Add(CreateElbowFittings(thirdElement, PrimaryElements[i], doc, uiapp));
                                                        }
                                                        catch
                                                        {
                                                            try
                                                            {
                                                                _deleteElements.Add(CreateElbowFittings(thirdElement, PrimaryElements[i], doc, uiapp));
                                                                _deleteElements.Add(CreateElbowFittings(thirdElement, SecondaryElements[i], doc, uiapp));
                                                            }
                                                            catch
                                                            {

                                                                string message = string.Format("Make sure conduits are having less overlap, if not please reduce the overlapping distance.");
                                                                System.Windows.MessageBox.Show("Warning. \n" + message, "Alert", MessageBoxButton.OK, MessageBoxImage.Warning);
                                                                return;
                                                            }


                                                        }


                                                    }



                                                }
                                            }

                                        }
                                    }


                                }

                            }
                        }

                        doc.Delete(_deleteElements.Select(x => x.Id).ToList());
                        tx.Commit();
                        doc.Regenerate();
                        successful = true;
                        _ = Utility.UserActivityLog(System.Reflection.Assembly.GetExecutingAssembly(), uiapp, "AutoConnect", startDate, "Completed", "Vertical Offset", "Public", "Connect");

                    }
                    using (Transaction tx = new Transaction(doc))
                    {
                        tx.Start("FinalSync");
                        Utility.ApplySync(PrimaryElements, uiapp);
                        tx.Commit();
                    }
                }
                catch (Exception ex)
                {
                    string message = string.Format("Make sure conduits are aligned to each other properly, if not please align primary conduit to secondary conduit. Error :{0}", ex.Message);
                    System.Windows.MessageBox.Show("Warning. \n" + message, "Alert", MessageBoxButton.OK, MessageBoxImage.Warning);
                    _ = Utility.UserActivityLog(System.Reflection.Assembly.GetExecutingAssembly(), uiapp, "AutoConnect", startDate, "Failed", "Vertical Offset", "Public", "Connect");
                    successful = false;
                }
            }
            catch (Exception exception)
            {
                System.Windows.MessageBox.Show("Warning. \n" + exception.Message, "Alert", MessageBoxButton.OK, MessageBoxImage.Warning);
                successful = false;
                _ = Utility.UserActivityLog(System.Reflection.Assembly.GetExecutingAssembly(), uiapp, "AutoConnect", startDate, "Failed", "Vertical Offset", "Public", "Connect");
            }
        }
        public static Autodesk.Revit.DB.FamilyInstance CreateElbowFittings(ConnectorSet PrimaryConnectors, ConnectorSet SecondaryConnectors, Document Doc)
        {
            Utility.GetClosestConnectors(PrimaryConnectors, SecondaryConnectors, out var ConnectorOne, out var ConnectorTwo);
            return Doc.Create.NewElbowFitting(ConnectorOne, ConnectorTwo);
        }
        public bool IsVertical(Curve curve)
        {
            // Implement your logic to check if the curve is vertical
            return Math.Abs(curve.GetEndPoint(0).X - curve.GetEndPoint(1).X) < 0.001;
        }
        public bool IsHorizontal(Curve curve)
        {
            // Implement your logic to check if the curve is horizontal
            return Math.Abs(curve.GetEndPoint(0).Y - curve.GetEndPoint(1).Y) < 0.001;
        }
        public static Autodesk.Revit.DB.FamilyInstance CreateElbowFittings(Autodesk.Revit.DB.Element One, Autodesk.Revit.DB.Element Two, Document doc, Autodesk.Revit.UI.UIApplication uiApp)
        {
            ConnectorSet connectorSet = GetConnectorSet(One);
            ConnectorSet connectorSet2 = GetConnectorSet(Two);
            Utility.AutoRetainParameters(One, Two, doc, uiApp);
            return CreateElbowFittings(connectorSet, connectorSet2, doc);
        }
        public static List<Element> GetElementsByOder(List<Element> a_PrimaryElements)
        {
            List<Element> PrimaryElements = new List<Element>();
            XYZ PrimaryDirection = ((a_PrimaryElements.LastOrDefault().Location as LocationCurve).Curve as Line).Direction;
            if (Math.Abs(PrimaryDirection.Z) != 1)
            {
                PrimaryElements = a_PrimaryElements.OrderBy(x => ((((x.Location as LocationCurve).Curve as Line).GetEndPoint(0) + ((x.Location as LocationCurve).Curve as Line).GetEndPoint(1)) / 2).Y).ToList();
                if (PrimaryDirection.Y == 1 || PrimaryDirection.Y == -1)
                {
                    PrimaryElements = a_PrimaryElements.OrderBy(x => ((((x.Location as LocationCurve).Curve as Line).GetEndPoint(0) + ((x.Location as LocationCurve).Curve as Line).GetEndPoint(1)) / 2).X).ToList();
                }
            }
            else
            {
                PrimaryElements = a_PrimaryElements.OrderBy(x => ((((x.Location as LocationCurve).Curve as Line).GetEndPoint(0) + ((x.Location as LocationCurve).Curve as Line).GetEndPoint(1)) / 2).X).ToList();
            }
            return PrimaryElements;
        }
        public void HoffsetExecute(UIApplication uiapp, ref List<Element> PrimaryElements, ref List<Element> SecondaryElements)
        {
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Application app = uiapp.Application;
            Document doc = uidoc.Document;
            int.TryParse(uiapp.Application.VersionNumber, out int RevitVersion);
            string offsetVariable = RevitVersion < 2020 ? "Offset" : "Middle Elevation";
            DateTime startDate = DateTime.UtcNow;
            ElementsFilter filter = new ElementsFilter("Conduits");
            double angle = Convert.ToDouble(MainWindow.Instance.angleDegree) * (Math.PI / 180);
            try
            {
                List<Element> thirdElements = new List<Element>();

                bool isVerticalConduits = false;
                using (Transaction tx = new Transaction(doc))
                {
                    tx.Start("Hoff");
                    Line refloineforanglecheck = null;
                    for (int i = 0; i < PrimaryElements.Count; i++)
                    {
                        Element firstElement = PrimaryElements[i];
                        Element secondElement = SecondaryElements[i];
                        Line firstLine = (firstElement.Location as LocationCurve).Curve as Line;
                        Line secondLine = (secondElement.Location as LocationCurve).Curve as Line;
                        Line newLine = GetParallelLine(firstElement, secondElement, ref isVerticalConduits, doc);
                        double elevation = firstElement.LookupParameter(offsetVariable).AsDouble();
                        XYZ newlineSeconpoint = newLine.GetEndPoint(0) + newLine.Direction.Multiply(20);
                        Conduit thirdConduit = Utility.CreateConduit(doc, firstElement as Conduit, newLine.GetEndPoint(0), newLine.GetEndPoint(1));
                        Element thirdElement = doc.GetElement(thirdConduit.Id);
                        thirdElements.Add(thirdElement);
                        if (i == 0)
                        {
                            refloineforanglecheck = newLine;
                        }

                    }
                    //Rotate Elements at Once
                    Element ElementOne = PrimaryElements[0];
                    Element ElementTwo = SecondaryElements[0];
                    Utility.GetClosestConnectors(ElementOne, ElementTwo, out Connector ConnectorOne, out Connector ConnectorTwo);
                    XYZ axisStart = ConnectorOne.Origin;
                    XYZ axisEnd = new XYZ(axisStart.X, axisStart.Y, axisStart.Z + 10);
                    Line axisLine = Line.CreateBound(axisStart, axisEnd);
                    if (isVerticalConduits)
                    {
                        axisEnd = new XYZ(ConnectorTwo.Origin.X, ConnectorTwo.Origin.Y, ConnectorOne.Origin.Z);
                        axisLine = Line.CreateBound(axisStart, axisEnd);
                        XYZ dir = axisLine.Direction;
                        dir = new XYZ(dir.X, dir.Y, 0);
                        XYZ cross = dir.CrossProduct(XYZ.BasisZ);
                        Element ele = thirdElements[0];
                        LocationCurve newconcurve = ele.Location as LocationCurve;
                        Line ncl1 = newconcurve.Curve as Line;
                        XYZ MidPoint = ncl1.GetEndPoint(0);
                        axisLine = Line.CreateBound(MidPoint, MidPoint + cross.Multiply(10));
                        ElementTransformUtils.RotateElements(doc, thirdElements.Select(r => r.Id).ToList(), axisLine, -angle);
                    }
                    else
                    {
                        ElementTransformUtils.RotateElements(doc, thirdElements.Select(r => r.Id).ToList(), axisLine, angle);

                        //find right angle
                        Conduit refcondone = SecondaryElements[0] as Conduit;
                        Line refcondoneline = (refcondone.Location as LocationCurve).Curve as Line;
                        XYZ refcondonelinedir = refcondoneline.Direction;
                        XYZ refcondonelinemidept = (refcondoneline.GetEndPoint(0) + refcondoneline.GetEndPoint(1)) / 2;
                        XYZ addedpt1 = refcondonelinemidept + refcondonelinedir.Multiply(250);
                        XYZ addedpt2 = refcondonelinemidept - refcondonelinedir.Multiply(250);
                        Line addedline = Line.CreateBound(addedpt1, addedpt2);

                        Conduit refcondtwo = thirdElements[0] as Conduit;
                        Line refcondtwoline = (refcondtwo.Location as LocationCurve).Curve as Line;
                        XYZ newlineSeconpoint = refcondtwoline.GetEndPoint(0) + refcondtwoline.Direction.Multiply(20);
                        addedpt1 = new XYZ(addedpt1.X, addedpt1.Y, newlineSeconpoint.Z);
                        addedpt2 = new XYZ(addedpt2.X, addedpt2.Y, newlineSeconpoint.Z);
                        refcondtwoline = Line.CreateBound(refcondtwoline.GetEndPoint(0), newlineSeconpoint);
                        addedline = Line.CreateBound(addedpt1, addedpt2);


                        XYZ intersectionpoint = Utility.GetIntersection(addedline, refcondtwoline);

                        if (intersectionpoint == null)
                        {
                            ElementTransformUtils.RotateElements(doc, thirdElements.Select(r => r.Id).ToList(), axisLine, -2 * angle);
                        }
                    }
                    for (int i = 0; i < PrimaryElements.Count; i++)
                    {
                        Element firstElement = PrimaryElements[i];
                        Element secondElement = SecondaryElements[i];
                        Element thirdElement = thirdElements[i];
                        Utility.CreateElbowFittings(SecondaryElements[i], thirdElement, doc, uiapp);
                        Utility.CreateElbowFittings(PrimaryElements[i], thirdElement, doc, uiapp);
                    }
                    tx.Commit();
                    successful = true;
                }
            }
            catch
            {
                try
                {
                    List<Element> thirdElements = new List<Element>();

                    bool isVerticalConduits = false;
                    using (Transaction tx = new Transaction(doc))
                    {
                        tx.Start("Hoffset");
                        for (int i = 0; i < PrimaryElements.Count; i++)
                        {
                            Element firstElement = PrimaryElements[i];
                            Element secondElement = SecondaryElements[i];
                            Line firstLine = (firstElement.Location as LocationCurve).Curve as Line;
                            Line secondLine = (secondElement.Location as LocationCurve).Curve as Line;
                            Line newLine = GetParallelLine(firstElement, secondElement, ref isVerticalConduits, doc);
                            double elevation = firstElement.LookupParameter(offsetVariable).AsDouble();
                            Conduit thirdConduit = Utility.CreateConduit(doc, firstElement as Conduit, newLine.GetEndPoint(0), newLine.GetEndPoint(1));
                            Element thirdElement = doc.GetElement(thirdConduit.Id);
                            thirdElements.Add(thirdElement);
                        }
                        //Rotate Elements at Once

                        Element ElementOne = PrimaryElements[0];
                        Element ElementTwo = SecondaryElements[0];
                        Utility.GetClosestConnectors(ElementOne, ElementTwo, out Connector ConnectorOne, out Connector ConnectorTwo);
                        XYZ axisStart = ConnectorOne.Origin;
                        XYZ axisEnd = new XYZ(axisStart.X, axisStart.Y, axisStart.Z + 10);
                        Line axisLine = Line.CreateBound(axisStart, axisEnd);
                        if (isVerticalConduits)
                        {
                            axisEnd = new XYZ(ConnectorTwo.Origin.X, ConnectorTwo.Origin.Y, ConnectorOne.Origin.Z);
                            axisLine = Line.CreateBound(axisStart, axisEnd);
                            XYZ dir = axisLine.Direction;
                            dir = new XYZ(dir.X, dir.Y, 0);
                            XYZ cross = dir.CrossProduct(XYZ.BasisZ);
                            Element ele = thirdElements[0];
                            LocationCurve newconcurve = ele.Location as LocationCurve;
                            Line ncl1 = newconcurve.Curve as Line;
                            XYZ MidPoint = ncl1.GetEndPoint(0);
                            axisLine = Line.CreateBound(MidPoint, MidPoint + cross.Multiply(10));
                            ElementTransformUtils.RotateElements(doc, thirdElements.Select(r => r.Id).ToList(), axisLine, angle);
                            Line rotatedLine = Line.CreateBound(ncl1.GetEndPoint(0) - ncl1.Direction.Multiply(10), ncl1.GetEndPoint(1) + ncl1.Direction.Multiply(10));
                        }
                        else
                        {
                            ElementTransformUtils.RotateElements(doc, thirdElements.Select(r => r.Id).ToList(), axisLine, -angle);
                        }
                        for (int i = 0; i < PrimaryElements.Count; i++)
                        {
                            Element firstElement = PrimaryElements[i];
                            Element secondElement = SecondaryElements[i];
                            Element thirdElement = thirdElements[i];
                            Utility.CreateElbowFittings(SecondaryElements[i], thirdElement, doc, uiapp);
                            Utility.CreateElbowFittings(PrimaryElements[i], thirdElement, doc, uiapp);
                        }
                        tx.Commit();
                        successful = true;
                    }
                }
                catch
                {
                    try
                    {
                        List<Element> thirdElements = new List<Element>();

                        bool isVerticalConduits = false;
                        using (Transaction tx = new Transaction(doc))
                        {
                            tx.Start("HoffsetExe");
                            for (int i = 0; i < PrimaryElements.Count; i++)
                            {
                                Element firstElement = PrimaryElements[i];
                                Element secondElement = SecondaryElements[i];
                                Line firstLine = (firstElement.Location as LocationCurve).Curve as Line;
                                Line secondLine = (secondElement.Location as LocationCurve).Curve as Line;
                                Line newLine = GetParallelLine(firstElement, secondElement, ref isVerticalConduits, doc);
                                double elevation = firstElement.LookupParameter(offsetVariable).AsDouble();
                                Conduit thirdConduit = Utility.CreateConduit(doc, firstElement as Conduit, newLine.GetEndPoint(0), newLine.GetEndPoint(1));
                                Element thirdElement = doc.GetElement(thirdConduit.Id);
                                thirdElements.Add(thirdElement);
                            }
                            //Rotate Elements at Once

                            Element ElementOne = PrimaryElements[0];
                            Element ElementTwo = SecondaryElements[0];
                            Utility.GetClosestConnectors(ElementOne, ElementTwo, out Connector ConnectorOne, out Connector ConnectorTwo);
                            XYZ axisStart = ConnectorOne.Origin;
                            XYZ axisEnd = new XYZ(axisStart.X, axisStart.Y, axisStart.Z + 10);
                            Line axisLine = Line.CreateBound(axisStart, axisEnd);
                            if (isVerticalConduits)
                            {
                                axisEnd = new XYZ(ConnectorTwo.Origin.X, ConnectorTwo.Origin.Y, ConnectorOne.Origin.Z);
                                axisLine = Line.CreateBound(axisStart, axisEnd);
                                XYZ dir = axisLine.Direction;
                                dir = new XYZ(dir.X, dir.Y, 0);
                                XYZ cross = dir.CrossProduct(XYZ.BasisZ);
                                Element ele = thirdElements[0];
                                LocationCurve newconcurve = ele.Location as LocationCurve;
                                Line ncl1 = newconcurve.Curve as Line;
                                XYZ MidPoint = ncl1.GetEndPoint(0);
                                axisLine = Line.CreateBound(MidPoint, MidPoint + cross.Multiply(10));
                                ElementTransformUtils.RotateElements(doc, thirdElements.Select(r => r.Id).ToList(), axisLine, angle);
                                Line rotatedLine = Line.CreateBound(ncl1.GetEndPoint(0) - ncl1.Direction.Multiply(10), ncl1.GetEndPoint(1) + ncl1.Direction.Multiply(10));
                            }
                            else
                            {
                                ElementTransformUtils.RotateElements(doc, thirdElements.Select(r => r.Id).ToList(), axisLine, angle);
                            }
                            for (int i = 0; i < PrimaryElements.Count; i++)
                            {
                                Element firstElement = PrimaryElements[i];
                                Element secondElement = SecondaryElements[i];
                                Element thirdElement = thirdElements[i];
                                //Utility.CreateElbowFittings(thirdElement, firstElement, doc, PrimaryElements[i], true);
                                Utility.CreateElbowFittings(SecondaryElements[i], thirdElement, doc, uiapp);
                                Utility.CreateElbowFittings(PrimaryElements[i], thirdElement, doc, uiapp);
                            }
                            tx.Commit();
                            successful = true;
                        }
                    }
                    catch (Exception ex)
                    {
                        System.Windows.MessageBox.Show("Warning. \n" + ex.Message, "Alert", MessageBoxButton.OK, MessageBoxImage.Warning);
                        successful = false;
                    }
                }
            }
        }
        public static Line GetParallelLine(Element firstElement, Element secondElement, ref bool isVerticalConduits, Document doc)
        {
            Line line = (firstElement.Location as LocationCurve).Curve as Line;
            Line line2 = (secondElement.Location as LocationCurve).Curve as Line;
            ConnectorSet connectors = Utility.GetConnectors(firstElement);
            ConnectorSet connectors2 = Utility.GetConnectors(secondElement);
            Utility.GetClosestConnectors(connectors, connectors2, out var ConnectorOne, out var ConnectorTwo);
            List<XYZ> list = new List<XYZ>();
            foreach (Connector item in connectors)
            {
                list.Add(item.Origin);
            }

            if (Utility.IsXYTrue(list.FirstOrDefault(), list.LastOrDefault()))
            {
                isVerticalConduits = true;
            }

            if (!isVerticalConduits)
            {
                XYZ direction = line2.Direction;
                XYZ xYZ = direction.CrossProduct(XYZ.BasisZ);
                Line line3 = Line.CreateBound(ConnectorTwo.Origin, ConnectorTwo.Origin + xYZ.Multiply(ConnectorOne.Origin.DistanceTo(ConnectorTwo.Origin)));
                XYZ xYZ2 = Utility.FindIntersectionPoint(line.GetEndPoint(0), line.GetEndPoint(1), line3.GetEndPoint(0), line3.GetEndPoint(1));
                XYZ endpoint = new XYZ(xYZ2.X, xYZ2.Y, ConnectorOne.Origin.Z);
                return Line.CreateBound(ConnectorOne.Origin, endpoint);
            }
            Line line4 = Line.CreateBound(ConnectorTwo.Origin, new XYZ(ConnectorOne.Origin.X, ConnectorOne.Origin.Y, ConnectorTwo.Origin.Z));
            XYZ xYZ3 = FindIntersectionPoint(line.GetEndPoint(0), line.GetEndPoint(1), line4.GetEndPoint(0), line4.GetEndPoint(1));
            if (xYZ3 == null)
            {
                xYZ3 = new XYZ(ConnectorOne.Origin.X, ConnectorOne.Origin.Y, 0);
            }
            XYZ endpoint2 = xYZ3 != null ? new XYZ(xYZ3.X, xYZ3.Y, ConnectorTwo.Origin.Z) : new XYZ(0, 0, ConnectorTwo.Origin.Z);
            return Line.CreateBound(ConnectorOne.Origin, endpoint2);
        }
        public static XYZ FindIntersectionPoint(XYZ s1, XYZ e1, XYZ s2, XYZ e2, int roundOff = 0)
        {
            if (roundOff > 0)
            {
                s1 = XYZroundOf(s1, roundOff);
                e1 = XYZroundOf(e1, roundOff);
                s2 = XYZroundOf(s2, roundOff);
                e2 = XYZroundOf(e2, roundOff);
            }

            double num = e1.Y - s1.Y;
            double num2 = s1.X - e1.X;
            double num3 = num * s1.X + num2 * s1.Y;
            double num4 = e2.Y - s2.Y;
            double num5 = s2.X - e2.X;
            double num6 = num4 * s2.X + num5 * s2.Y;
            double num7 = num * num5 - num4 * num2;
            return (num7 == 0.0) ? null : new XYZ((num5 * num3 - num2 * num6) / num7, (num * num6 - num4 * num3) / num7, 0.0);
        }
        public static XYZ XYZroundOf(XYZ xyz, int digit)
        {
            return new XYZ(Math.Round(xyz.X, digit), Math.Round(xyz.Y, digit), Math.Round(xyz.Z, digit));
        }
        public void VoffsetExecute(UIApplication uiapp, ref List<Element> PrimaryElements, ref List<Element> SecondaryElements)
        {
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Application app = uiapp.Application;
            Document doc = uidoc.Document;
            int.TryParse(uiapp.Application.VersionNumber, out int RevitVersion);
            List<Element> _deleteElements = new List<Element>();
            List<Element> SelectedElements = new List<Element>();
            ElementsFilter filter = new ElementsFilter("Conduits");

            try
            {

                //PrimaryElements = Elements.GetElementsByReference(PrimaryReference, doc);
                //SecondaryElements = Elements.GetElementsByReference(SecondaryReference, doc);
                PrimaryElements = GetElementsByOder(PrimaryElements);
                SecondaryElements = GetElementsByOder(SecondaryElements);
                List<Element> thirdElements = new List<Element>();
                bool isVerticalConduits = false;
                // Modify document within a transaction
                try
                {
                    using (Transaction tx = new Transaction(doc))
                    {
                        ConnectorSet PrimaryConnectors = null;
                        ConnectorSet SecondaryConnectors = null;
                        Connector ConnectorOne = null;
                        Connector ConnectorTwo = null;
                        tx.Start("Voff");
                        double l_angle = Convert.ToDouble(MainWindow.Instance.angleDegree) * (Math.PI / 180);
                        for (int i = 0; i < PrimaryElements.Count; i++)
                        {
                            List<XYZ> ConnectorPoints = new List<XYZ>();
                            PrimaryConnectors = Utility.GetConnectors(PrimaryElements[i]);
                            SecondaryConnectors = Utility.GetConnectors(SecondaryElements[i]);
                            Utility.GetClosestConnectors(PrimaryConnectors, SecondaryConnectors, out ConnectorOne, out ConnectorTwo);
                            foreach (Connector con in PrimaryConnectors)
                            {
                                ConnectorPoints.Add(con.Origin);
                            }
                            XYZ newenpt = new XYZ(ConnectorTwo.Origin.X, ConnectorTwo.Origin.Y, ConnectorOne.Origin.Z);
                            Conduit newCon = Utility.CreateConduit(doc, PrimaryElements[i] as Conduit, ConnectorOne.Origin, newenpt);
                            if (Utility.IsXYTrue(ConnectorPoints.FirstOrDefault(), ConnectorPoints.LastOrDefault()))
                            {
                                isVerticalConduits = true;
                            }
                            Element e = doc.GetElement(newCon.Id);
                            thirdElements.Add(e);
                            //RetainParameters(PrimaryElements[i], SecondaryElements[i]);
                            //RetainParameters(PrimaryElements[i], e);
                        }
                        //Rotate Elements at Once
                        Element ElementOne = PrimaryElements[0];
                        Element ElementTwo = SecondaryElements[0];
                        Utility.GetClosestConnectors(ElementOne, ElementTwo, out ConnectorOne, out ConnectorTwo);
                        LocationCurve newconcurve = thirdElements[0].Location as LocationCurve;
                        Line ncl1 = newconcurve.Curve as Line;
                        XYZ direction = ncl1.Direction;
                        XYZ axisStart = ConnectorOne.Origin;
                        XYZ axisEnd = axisStart.Add(XYZ.BasisZ.CrossProduct(direction));
                        Line axisLine = Line.CreateBound(axisStart, axisEnd);
                        double PrimaryOffset = RevitVersion < 2020 ? PrimaryElements[0].LookupParameter("Offset").AsDouble() :
                                                 PrimaryElements[0].LookupParameter("Middle Elevation").AsDouble();
                        double SecondaryOffset = RevitVersion < 2020 ? SecondaryElements[0].LookupParameter("Offset").AsDouble() :
                                                  SecondaryElements[0].LookupParameter("Middle Elevation").AsDouble();
                        if (isVerticalConduits)
                        {
                            l_angle = (Math.PI / 2) - l_angle;
                        }
                        if (PrimaryOffset > SecondaryOffset)
                        {
                            //rotate down
                            l_angle = -l_angle;
                        }
                        try
                        {
                            ElementTransformUtils.RotateElements(doc, thirdElements.Select(r => r.Id).ToList(), axisLine, -l_angle);
                            for (int i = 0; i < PrimaryElements.Count; i++)
                            {
                                Element firstElement = PrimaryElements[i];
                                Element secondElement = SecondaryElements[i];
                                Element thirdElement = thirdElements[i];
                                Utility.AutoRetainParameters(PrimaryElements[i], SecondaryElements[i], uidoc, uiapp);
                                Utility.AutoRetainParameters(PrimaryElements[i], thirdElement, doc, uiapp);
                                try
                                {

                                    Utility.CreateElbowFittings(SecondaryElements[i], thirdElement, doc, uiapp);
                                    Utility.CreateElbowFittings(PrimaryElements[i], thirdElement, doc, uiapp);
                                }
                                catch
                                {
                                    Utility.CreateElbowFittings(thirdElement, SecondaryElements[i], doc, uiapp);
                                    Utility.CreateElbowFittings(thirdElement, PrimaryElements[i], doc, uiapp);

                                }
                            }
                        }
                        catch (Exception)
                        {

                            ElementTransformUtils.RotateElements(doc, thirdElements.Select(r => r.Id).ToList(), axisLine, l_angle * 2 + Math.PI);

                            for (int i = 0; i < PrimaryElements.Count; i++)
                            {
                                Element firstElement = PrimaryElements[i];
                                Element secondElement = SecondaryElements[i];
                                Element thirdElement = thirdElements[i];
                                Utility.AutoRetainParameters(PrimaryElements[i], SecondaryElements[i], uidoc, uiapp);
                                Utility.AutoRetainParameters(PrimaryElements[i], thirdElement, doc, uiapp);
                                //Utility.CreateElbowFittings(PrimaryElements[i], thirdElement, doc, uiapp);
                                //Utility.CreateElbowFittings(thirdElement, SecondaryElements[i], doc, uiapp);
                                try
                                {
                                    _deleteElements.Add(CreateElbowFittings(PrimaryElements[i], thirdElement, doc, uiapp));
                                    _deleteElements.Add(CreateElbowFittings(thirdElement, SecondaryElements[i], doc, uiapp));
                                }
                                catch
                                {
                                    try
                                    {

                                        _deleteElements.Add(CreateElbowFittings(thirdElement, PrimaryElements[i], doc, uiapp));
                                        _deleteElements.Add(CreateElbowFittings(SecondaryElements[i], thirdElement, doc, uiapp));
                                    }
                                    catch
                                    {

                                        try
                                        {

                                            _deleteElements.Add(CreateElbowFittings(SecondaryElements[i], thirdElement, doc, uiapp));
                                            _deleteElements.Add(CreateElbowFittings(thirdElement, PrimaryElements[i], doc, uiapp));
                                        }
                                        catch
                                        {
                                            try
                                            {

                                                _deleteElements.Add(CreateElbowFittings(thirdElement, SecondaryElements[i], doc, uiapp));
                                                _deleteElements.Add(CreateElbowFittings(PrimaryElements[i], thirdElement, doc, uiapp));

                                            }
                                            catch
                                            {
                                                try
                                                {

                                                    _deleteElements.Add(CreateElbowFittings(PrimaryElements[i], thirdElement, doc, uiapp));
                                                    _deleteElements.Add(CreateElbowFittings(SecondaryElements[i], thirdElement, doc, uiapp));

                                                }
                                                catch
                                                {
                                                    try
                                                    {

                                                        _deleteElements.Add(CreateElbowFittings(SecondaryElements[i], thirdElement, doc, uiapp));
                                                        _deleteElements.Add(CreateElbowFittings(PrimaryElements[i], thirdElement, doc, uiapp));
                                                    }
                                                    catch
                                                    {
                                                        try
                                                        {

                                                            _deleteElements.Add(CreateElbowFittings(thirdElement, SecondaryElements[i], doc, uiapp));
                                                            _deleteElements.Add(CreateElbowFittings(thirdElement, PrimaryElements[i], doc, uiapp));
                                                        }
                                                        catch
                                                        {
                                                            try
                                                            {
                                                                _deleteElements.Add(CreateElbowFittings(thirdElement, PrimaryElements[i], doc, uiapp));
                                                                _deleteElements.Add(CreateElbowFittings(thirdElement, SecondaryElements[i], doc, uiapp));
                                                            }
                                                            catch
                                                            {

                                                                string message = string.Format("Make sure conduits are having less overlap, if not please reduce the overlapping distance.");
                                                                System.Windows.MessageBox.Show("Warning. \n" + message, "Alert", MessageBoxButton.OK, MessageBoxImage.Warning);
                                                                successful = false;
                                                                return;
                                                            }


                                                        }


                                                    }



                                                }
                                            }

                                        }
                                    }


                                }

                            }
                        }
                        doc.Regenerate();
                        tx.Commit();
                        successful = true;


                    }
                    using (Transaction tx = new Transaction(doc))
                    {
                        tx.Start("Sync");
                        Utility.ApplySync(PrimaryElements, uiapp);
                        tx.Commit();
                    }
                }
                catch (Exception ex)
                {
                    string message = string.Format("Make sure conduits are aligned to each other properly, if not please align primary conduit to secondary conduit. Error :{0}", ex.Message);
                    System.Windows.MessageBox.Show("Warning. \n" + message, "Alert", MessageBoxButton.OK, MessageBoxImage.Warning);
                    successful = false;

                }
            }
            catch (Exception exception)
            {
                System.Windows.MessageBox.Show("Warning. \n" + exception.Message, "Alert", MessageBoxButton.OK, MessageBoxImage.Warning);
                successful = false;
            }
        }
        public void KickExecute(UIApplication uiapp, ref List<Element> PrimaryElements, ref List<Element> SecondaryElements, int first)
        {
            DateTime startDate = DateTime.UtcNow;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Document doc = uidoc.Document;
            int.TryParse(uiapp.Application.VersionNumber, out int RevitVersion);
            string offsetVariable = RevitVersion < 2020 ? "Offset" : "Middle Elevation";
            try
            {
                Reference reference = null;
                if (first == 0)
                {
                    ElementsFilter filter = new ElementsFilter("Conduits");
                    //if (ChangesInformationForm.instance != null && ChangesInformationForm.instance._refConduitKick == null)
                    //{
                    reference = uidoc.Selection.PickObject(ObjectType.Element, filter, "Please select the conduit in group to define 90 near and 90 far");
                    //}
                    //ElementId refId = ChangesInformationForm.instance._refConduitKick[0];
                    if (!PrimaryElements.Any(e => e.Id == doc.GetElement(reference.ElementId).Id))
                    {
                        var temp = PrimaryElements;
                        PrimaryElements = SecondaryElements;
                        SecondaryElements = temp;

                        _isfirst = true;
                    }
                }
                if (first > 0)
                {
                    if (_isfirst)
                    {
                        var temp = PrimaryElements;
                        PrimaryElements = SecondaryElements;
                        SecondaryElements = temp;
                    }
                }


                double l_angle;
                bool isUp = PrimaryElements.FirstOrDefault().LookupParameter(offsetVariable).AsDouble() <
                    SecondaryElements.FirstOrDefault().LookupParameter(offsetVariable).AsDouble();
                if (!isUp)
                {
                    l_angle = Convert.ToDouble(MainWindow.Instance.angleDegree) * (Math.PI / 180);
                    try
                    {
                        using (Transaction tx = new Transaction(doc))
                        {
                            ConnectorSet PrimaryConnectors = null;
                            ConnectorSet SecondaryConnectors = null;
                            ConnectorSet ThirdConnectors = null;
                            tx.Start("Kick");
                            for (int i = 0; i < PrimaryElements.Count; i++)
                            {
                                double elevation = PrimaryElements[i].LookupParameter(offsetVariable).AsDouble();
                                LocationCurve lc1 = PrimaryElements[i].Location as LocationCurve;
                                Line l1 = lc1.Curve as Line;
                                LocationCurve lc2 = SecondaryElements[i].Location as LocationCurve;
                                Line l2 = lc2.Curve as Line;
                                XYZ interSecPoint = Utility.FindIntersectionPoint(l1.GetEndPoint(0), l1.GetEndPoint(1), l2.GetEndPoint(0), l2.GetEndPoint(1));
                                PrimaryConnectors = Utility.GetConnectors(PrimaryElements[i]);
                                SecondaryConnectors = Utility.GetConnectors(SecondaryElements[i]);
                                Utility.GetClosestConnectors(PrimaryConnectors, SecondaryConnectors, out Connector ConnectorOne, out Connector ConnectorTwo);
                                XYZ EndPoint = ConnectorTwo.Origin;
                                XYZ NewEndPoint = new XYZ(interSecPoint.X, interSecPoint.Y, EndPoint.Z);
                                Conduit newCon = Utility.CreateConduit(doc, PrimaryElements[i] as Conduit, EndPoint, NewEndPoint);
                                newCon.LookupParameter(offsetVariable).Set(elevation);
                                Element e = doc.GetElement(newCon.Id);
                                LocationCurve newConcurve = newCon.Location as LocationCurve;
                                Line ncl1 = newConcurve.Curve as Line;
                                XYZ ncenpt = ncl1.GetEndPoint(1);
                                XYZ direction = ncl1.Direction;
                                XYZ midPoint = ncenpt;
                                XYZ midHigh = midPoint.Add(XYZ.BasisZ.CrossProduct(direction));
                                Line axisLine = Line.CreateBound(midPoint, midHigh);
                                newConcurve.Rotate(axisLine, -l_angle);
                                ThirdConnectors = Utility.GetConnectors(e);
                                Utility.AutoRetainParameters(PrimaryElements[i], SecondaryElements[i], uidoc, uiapp);
                                Utility.AutoRetainParameters(PrimaryElements[i], e, doc, uiapp);
                                Utility.CreateElbowFittings(SecondaryElements[i], e, doc, uiapp);
                                Utility.CreateElbowFittings(PrimaryElements[i], e, doc, uiapp);
                            }
                            tx.Commit();
                            successful = true;
                        }
                    }
                    catch
                    {
                        try
                        {
                            using (Transaction tx = new Transaction(doc))
                            {
                                ConnectorSet PrimaryConnectors = null;
                                ConnectorSet SecondaryConnectors = null;
                                ConnectorSet ThirdConnectors = null;

                                tx.Start("Kickexe");
                                for (int i = 0; i < PrimaryElements.Count; i++)
                                {
                                    double elevation = PrimaryElements[i].LookupParameter(offsetVariable).AsDouble();
                                    LocationCurve lc1 = PrimaryElements[i].Location as LocationCurve;
                                    Line l1 = lc1.Curve as Line;
                                    LocationCurve lc2 = SecondaryElements[i].Location as LocationCurve;
                                    Line l2 = lc2.Curve as Line;
                                    XYZ interSecPoint = Utility.FindIntersectionPoint(l1.GetEndPoint(0), l1.GetEndPoint(1), l2.GetEndPoint(0), l2.GetEndPoint(1));
                                    PrimaryConnectors = Utility.GetConnectors(PrimaryElements[i]);
                                    SecondaryConnectors = Utility.GetConnectors(SecondaryElements[i]);
                                    Utility.GetClosestConnectors(PrimaryConnectors, SecondaryConnectors, out Connector ConnectorOne, out Connector ConnectorTwo);
                                    XYZ EndPoint = ConnectorTwo.Origin;
                                    XYZ NewEndPoint = new XYZ(interSecPoint.X, interSecPoint.Y, EndPoint.Z);
                                    Conduit newCon = Utility.CreateConduit(doc, PrimaryElements[i] as Conduit, EndPoint, NewEndPoint);
                                    newCon.LookupParameter(offsetVariable).Set(elevation);
                                    Element e = doc.GetElement(newCon.Id);
                                    LocationCurve newConcurve = newCon.Location as LocationCurve;
                                    Line ncl1 = newConcurve.Curve as Line;
                                    XYZ ncenpt = ncl1.GetEndPoint(1);
                                    XYZ direction = ncl1.Direction;
                                    XYZ midPoint = ncenpt;
                                    XYZ midHigh = midPoint.Add(XYZ.BasisZ.CrossProduct(direction));
                                    Line axisLine = Line.CreateBound(midPoint, midHigh);
                                    newConcurve.Rotate(axisLine, l_angle);
                                    ThirdConnectors = Utility.GetConnectors(e);
                                    Utility.AutoRetainParameters(PrimaryElements[i], SecondaryElements[i], uidoc, uiapp);
                                    Utility.AutoRetainParameters(PrimaryElements[i], e, doc, uiapp);
                                    Utility.CreateElbowFittings(SecondaryElements[i], e, doc, uiapp);
                                    Utility.CreateElbowFittings(PrimaryElements[i], e, doc, uiapp);
                                }
                                tx.Commit();
                                successful = true;
                            }
                        }
                        catch (Exception exception)
                        {
                            System.Windows.MessageBox.Show("Warning. \n" + exception.Message, "Alert", MessageBoxButton.OK, MessageBoxImage.Warning);
                            successful = false;
                        }
                    }
                }
                if (isUp)
                {
                    l_angle = Convert.ToDouble(MainWindow.Instance.angleDegree) * (Math.PI / 180);
                    try
                    {

                        using (Transaction tx = new Transaction(doc))
                        {
                            ConnectorSet PrimaryConnectors = null;
                            ConnectorSet SecondaryConnectors = null;
                            ConnectorSet ThirdConnectors = null;

                            tx.Start("KickExeute");
                            for (int i = 0; i < PrimaryElements.Count; i++)
                            {
                                double elevation = PrimaryElements[i].LookupParameter(offsetVariable).AsDouble();
                                LocationCurve lc1 = PrimaryElements[i].Location as LocationCurve;
                                Line l1 = lc1.Curve as Line;
                                LocationCurve lc2 = SecondaryElements[i].Location as LocationCurve;
                                Line l2 = lc2.Curve as Line;
                                XYZ interSecPoint = Utility.FindIntersectionPoint(l1.GetEndPoint(0), l1.GetEndPoint(1), l2.GetEndPoint(0), l2.GetEndPoint(1));
                                PrimaryConnectors = Utility.GetConnectors(PrimaryElements[i]);
                                SecondaryConnectors = Utility.GetConnectors(SecondaryElements[i]);
                                Utility.GetClosestConnectors(PrimaryConnectors, SecondaryConnectors, out Connector ConnectorOne, out Connector ConnectorTwo);
                                XYZ EndPoint = ConnectorTwo.Origin;
                                XYZ NewEndPoint = new XYZ(interSecPoint.X, interSecPoint.Y, EndPoint.Z);
                                Conduit newCon = Utility.CreateConduit(doc, PrimaryElements[i] as Conduit, EndPoint, NewEndPoint);
                                newCon.LookupParameter(offsetVariable).Set(elevation);
                                Element e = doc.GetElement(newCon.Id);
                                LocationCurve newConcurve = newCon.Location as LocationCurve;
                                Line ncl1 = newConcurve.Curve as Line;
                                XYZ ncenpt = ncl1.GetEndPoint(1);
                                XYZ direction = ncl1.Direction;
                                XYZ midPoint = ncenpt;
                                XYZ midHigh = midPoint.Add(XYZ.BasisZ.CrossProduct(direction));
                                Line axisLine = Line.CreateBound(midPoint, midHigh);
                                newConcurve.Rotate(axisLine, l_angle);
                                ThirdConnectors = Utility.GetConnectors(e);
                                Utility.AutoRetainParameters(PrimaryElements[i], SecondaryElements[i], uidoc, uiapp);
                                Utility.AutoRetainParameters(PrimaryElements[i], e, doc, uiapp);
                                Utility.CreateElbowFittings(SecondaryElements[i], e, doc, uiapp);
                                Utility.CreateElbowFittings(PrimaryElements[i], e, doc, uiapp);
                            }
                            tx.Commit();
                            successful = true;

                        }
                    }
                    catch
                    {
                        try
                        {
                            using (Transaction tx = new Transaction(doc))
                            {
                                ConnectorSet PrimaryConnectors = null;
                                ConnectorSet SecondaryConnectors = null;
                                ConnectorSet ThirdConnectors = null;

                                tx.Start("KickExecutetrans");
                                for (int i = 0; i < PrimaryElements.Count; i++)
                                {
                                    double elevation = PrimaryElements[i].LookupParameter(offsetVariable).AsDouble();
                                    LocationCurve lc1 = PrimaryElements[i].Location as LocationCurve;
                                    Line l1 = lc1.Curve as Line;
                                    LocationCurve lc2 = SecondaryElements[i].Location as LocationCurve;
                                    Line l2 = lc2.Curve as Line;
                                    XYZ interSecPoint = Utility.FindIntersectionPoint(l1.GetEndPoint(0), l1.GetEndPoint(1), l2.GetEndPoint(0), l2.GetEndPoint(1));
                                    PrimaryConnectors = Utility.GetConnectors(PrimaryElements[i]);
                                    SecondaryConnectors = Utility.GetConnectors(SecondaryElements[i]);
                                    Utility.GetClosestConnectors(PrimaryConnectors, SecondaryConnectors, out Connector ConnectorOne, out Connector ConnectorTwo);
                                    XYZ EndPoint = ConnectorTwo.Origin;
                                    XYZ NewEndPoint = new XYZ(interSecPoint.X, interSecPoint.Y, EndPoint.Z);
                                    Conduit newCon = Utility.CreateConduit(doc, PrimaryElements[i] as Conduit, EndPoint, NewEndPoint);
                                    newCon.LookupParameter(offsetVariable).Set(elevation);
                                    Element e = doc.GetElement(newCon.Id);
                                    LocationCurve newConcurve = newCon.Location as LocationCurve;
                                    Line ncl1 = newConcurve.Curve as Line;
                                    XYZ ncenpt = ncl1.GetEndPoint(1);
                                    XYZ direction = ncl1.Direction;
                                    XYZ midPoint = ncenpt;
                                    XYZ midHigh = midPoint.Add(XYZ.BasisZ.CrossProduct(direction));
                                    Line axisLine = Line.CreateBound(midPoint, midHigh);
                                    newConcurve.Rotate(axisLine, -l_angle);
                                    ThirdConnectors = Utility.GetConnectors(e);
                                    Utility.AutoRetainParameters(PrimaryElements[i], SecondaryElements[i], uidoc, uiapp);
                                    Utility.AutoRetainParameters(PrimaryElements[i], e, doc, uiapp);
                                    Utility.CreateElbowFittings(SecondaryElements[i], e, doc, uiapp);
                                    Utility.CreateElbowFittings(PrimaryElements[i], e, doc, uiapp);
                                }
                                tx.Commit();
                                successful = true;
                            }
                        }
                        catch (Exception exception)
                        {
                            System.Windows.MessageBox.Show("Warning. \n" + exception.Message, "Alert", MessageBoxButton.OK, MessageBoxImage.Warning);
                            successful = false;
                        }
                    }
                }
            }
            catch (Exception exception)
            {
                System.Windows.MessageBox.Show("Warning. \n" + exception.Message, "Alert", MessageBoxButton.OK, MessageBoxImage.Warning);
                successful = false;
            }
        }

        #region GROUPING DOUBLE LAYERED
        private XYZ GetConduitStartPoint(Element conduit)
        {
            LocationCurve locationCurve = conduit.Location as LocationCurve;
            if (locationCurve == null)
                throw new InvalidOperationException("Conduit does not have a valid location curve.");

            Line line = locationCurve.Curve as Line;
            if (line == null)
                throw new InvalidOperationException("Conduit location curve is not a line.");

            return line.GetEndPoint(0);
        }
        private List<List<Element>> GroupConduitsIntoPairs(List<Element> conduits, double tolerance = 0.01)
        {
            List<List<Element>> groupedConduits = new List<List<Element>>();
            HashSet<ElementId> processedIds = new HashSet<ElementId>();



            var sortedConduits = conduits.OrderBy(c => GetConduitStartPoint(c).X)
                                            .ThenBy(c => GetConduitStartPoint(c).Y)
                                            .ThenBy(c => GetConduitStartPoint(c).Z)
                                            .ToList();

            for (int i = 0; i < sortedConduits.Count - 1; i++)
            {
                if (processedIds.Contains(sortedConduits[i].Id))
                    continue;

                var firstConduit = sortedConduits[i];
                double firstZ = GetConduitStartPoint(firstConduit).Z;


                var secondConduit = sortedConduits.Skip(i + 1).FirstOrDefault(c =>
                {
                    double secondZ = GetConduitStartPoint(c).Z;
                    return Math.Abs(secondZ - firstZ) > tolerance && !processedIds.Contains(c.Id);
                });

                if (secondConduit != null)
                {
                    groupedConduits.Add(new List<Element> { firstConduit, secondConduit });
                    processedIds.Add(firstConduit.Id);
                    processedIds.Add(secondConduit.Id);
                }
                else
                {

                    groupedConduits.Add(new List<Element> { firstConduit });
                    processedIds.Add(firstConduit.Id);
                }
            }


            var remainingConduits = sortedConduits.Where(c => !processedIds.Contains(c.Id)).ToList();
            for (int i = 0; i < remainingConduits.Count - 1; i += 2)
            {
                var group = new List<Element> { remainingConduits[i] };
                if (i + 1 < remainingConduits.Count)
                {
                    group.Add(remainingConduits[i + 1]);
                }
                groupedConduits.Add(group);
            }

            return groupedConduits;
        }
        #endregion
        public static Dictionary<int, List<ConduitGrid>> GroupStubElements(List<Autodesk.Revit.DB.Element> ElementCollection, double maximumSpacing = 0.5)
        {
            List<ConduitGrid> list = new List<ConduitGrid>();
            List<Autodesk.Revit.DB.Element> list2 = new List<Autodesk.Revit.DB.Element>();
            foreach (Autodesk.Revit.DB.Element item2 in ElementCollection)
            {
                Line line = (item2.Location as LocationCurve).Curve as Line;
                XYZ DirectionOne = line.Direction;
                if (!list2.Any((Autodesk.Revit.DB.Element r) => Utility.IsSameDirection(((r.Location as LocationCurve).Curve as Line).Direction, DirectionOne)))
                {
                    list2.Add(item2);
                }
            }

            foreach (Autodesk.Revit.DB.Element item3 in list2)
            {
                Line line2 = (item3.Location as LocationCurve).Curve as Line;
                XYZ direction = line2.Direction;
                foreach (Autodesk.Revit.DB.Element item4 in ElementCollection)
                {

                    Line line3 = (item4.Location as LocationCurve).Curve as Line;
                    XYZ direction2 = line3.Direction;

                    if (!Utility.IsSameDirection(direction, direction2))
                    {
                        continue;
                    }


                    XYZ endPoint = line2.GetEndPoint(0);
                    XYZ endPoint2 = line2.GetEndPoint(1);

                    Line line4 = null;
                    XYZ xYZ = (endPoint + endPoint2) / 2.0;
                    XYZ xYZ2 = line2.Direction.CrossProduct(XYZ.BasisZ);
                    XYZ endpoint = xYZ + xYZ2.Multiply(100.0);
                    XYZ endpoint2 = xYZ - xYZ2.Multiply(100.0);
                    try
                    {
                        line4 = Line.CreateBound(endpoint, endpoint2);
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine(ex.Message);
                    }
                    try
                    {
                        XYZ xYZ3 = null;
                        try
                        {
                            xYZ3 = Utility.FindIntersection(item4, line4);
                        }
                        catch
                        {

                        }
                        if (xYZ3 != null)
                        {
                            ConduitGrid item = new ConduitGrid(item4, item3, xYZ3.DistanceTo(new XYZ(xYZ.X, xYZ.Y, 0.0)));
                            list.Add(item);
                        }
                        else
                        {
                            endPoint = line2.GetEndPoint(0);
                            endPoint2 = line3.GetEndPoint(0);
                            ConduitGrid item = new ConduitGrid(item4, item3, new XYZ(endPoint.X, endPoint.Y, 0).DistanceTo(new XYZ(endPoint2.X, endPoint2.Y, 0)));
                            list.Add(item);
                        }

                    }
                    catch (Exception ex2)
                    {
                        Console.WriteLine(ex2.Message);
                    }
                }
            }

            Dictionary<int, List<ConduitGrid>> CongridDictionary = new Dictionary<int, List<ConduitGrid>>();
            List<Autodesk.Revit.DB.Element> list3 = new List<Autodesk.Revit.DB.Element>();
            int index = 0;
            foreach (IGrouping<Autodesk.Revit.DB.ElementId, ConduitGrid> item5 in from r in list
                                                                                  group r by r.RefConduit.Id)
            {
                foreach (ConduitGrid congrid in item5.OrderBy((ConduitGrid x) => x.Distance))
                {
                    if (list3.Any((Autodesk.Revit.DB.Element r) => r.Id == congrid.Conduit.Id))
                    {
                        continue;
                    }

                    List<ConduitGrid> source = list.Where((ConduitGrid r) => r.RefConduit.Id == congrid.RefConduit.Id && !CongridDictionary.Any((KeyValuePair<int, List<ConduitGrid>> x) => x.Value.Any((ConduitGrid y) => y.Conduit.Id == r.Conduit.Id))).ToList();
                    List<ConduitGrid> IntersectingConduits = source.Where((ConduitGrid r) => GetIntersection(r.Conduit, congrid, congrid.StartPoint, maximumSpacing) != null).ToList();
                    IntersectingConduits.AddRange(source.Where((ConduitGrid r) => !IntersectingConduits.Any((ConduitGrid x) => r.Conduit.Id == x.Conduit.Id) && GetIntersection(r.Conduit, congrid, congrid.EndPoint, maximumSpacing) != null).ToList());
                    IntersectingConduits.AddRange(source.Where((ConduitGrid r) => !IntersectingConduits.Any((ConduitGrid x) => r.Conduit.Id == x.Conduit.Id) && GetIntersection(r.Conduit, congrid, congrid.MidPoint, maximumSpacing) != null).ToList());
                    if (IntersectingConduits == null || !IntersectingConduits.Any())
                    {
                        continue;
                    }

                    if (!CongridDictionary.Any((KeyValuePair<int, List<ConduitGrid>> r) => r.Value.Any((ConduitGrid x) => IntersectingConduits.Any((ConduitGrid y) => y.Conduit.Id == x.Conduit.Id))))
                    {
                        foreach (ConduitGrid cg in IntersectingConduits)
                        {
                            source = list.Where((ConduitGrid r) => r.RefConduit.Id == cg.RefConduit.Id).ToList();
                            List<ConduitGrid> ISC = source.Where((ConduitGrid r) => GetIntersection(r.Conduit, cg, cg.StartPoint, maximumSpacing) != null).ToList();
                            ISC.AddRange(source.Where((ConduitGrid r) => !ISC.Any((ConduitGrid x) => r.Conduit.Id == x.Conduit.Id) && GetIntersection(r.Conduit, cg, cg.EndPoint, maximumSpacing) != null).ToList());
                            ISC.AddRange(source.Where((ConduitGrid r) => !ISC.Any((ConduitGrid x) => r.Conduit.Id == x.Conduit.Id) && GetIntersection(r.Conduit, cg, cg.MidPoint, maximumSpacing) != null).ToList());
                            if (ISC == null)
                            {
                                continue;
                            }

                            if (!CongridDictionary.Any((KeyValuePair<int, List<ConduitGrid>> r) => r.Value.Any((ConduitGrid x) => ISC.Any((ConduitGrid y) => y.Conduit.Id == x.Conduit.Id))))
                            {
                                ISC = ISC.Where((ConduitGrid x) => !CongridDictionary.Any((KeyValuePair<int, List<ConduitGrid>> r) => r.Value.Any((ConduitGrid y) => x.Conduit.Id == y.Conduit.Id))).ToList();
                                if (!CongridDictionary.Any((KeyValuePair<int, List<ConduitGrid>> r) => r.Key == index))
                                {
                                    CongridDictionary.Add(index, ISC);
                                }

                                continue;
                            }

                            KeyValuePair<int, List<ConduitGrid>> keyValuePair = CongridDictionary.FirstOrDefault((KeyValuePair<int, List<ConduitGrid>> r) => r.Value.Any((ConduitGrid x) => ISC.Any((ConduitGrid y) => y.Conduit.Id == x.Conduit.Id)));
                            List<ConduitGrid> value = keyValuePair.Value;
                            if (value == null)
                            {
                                continue;
                            }

                            foreach (ConduitGrid conGrid3 in ISC)
                            {
                                if (!value.Any((ConduitGrid r) => r.Conduit.Id == conGrid3.Conduit.Id && r.RefConduit.Id == conGrid3.RefConduit.Id) && !CongridDictionary.Any((KeyValuePair<int, List<ConduitGrid>> r) => r.Value.Any((ConduitGrid x) => x.Conduit.Id == conGrid3.Conduit.Id)))
                                {
                                    value.Add(conGrid3);
                                }
                            }

                            CongridDictionary[keyValuePair.Key] = value;
                        }

                        if (!CongridDictionary.Any((KeyValuePair<int, List<ConduitGrid>> r) => IntersectingConduits.Any((ConduitGrid x) => r.Value.Any((ConduitGrid y) => x.Conduit.Id == y.Conduit.Id))))
                        {
                            IntersectingConduits = IntersectingConduits.Where((ConduitGrid x) => !CongridDictionary.Any((KeyValuePair<int, List<ConduitGrid>> r) => r.Value.Any((ConduitGrid y) => x.Conduit.Id == y.Conduit.Id))).ToList();
                            if (!CongridDictionary.Any((KeyValuePair<int, List<ConduitGrid>> r) => r.Key == index))
                            {
                                if (!CongridDictionary.Any((KeyValuePair<int, List<ConduitGrid>> r) => IntersectingConduits.Any((ConduitGrid x) => r.Value.Any((ConduitGrid y) => x.Conduit.Id == y.RefConduit.Id || x.RefConduit.Id == y.Conduit.Id))))
                                {
                                    CongridDictionary.Add(index, IntersectingConduits);
                                }
                                else
                                {
                                    KeyValuePair<int, List<ConduitGrid>> keyValuePair2 = CongridDictionary.FirstOrDefault((KeyValuePair<int, List<ConduitGrid>> r) => r.Value.Any((ConduitGrid x) => IntersectingConduits.Any((ConduitGrid y) => y.RefConduit.Id == x.Conduit.Id)));
                                    List<ConduitGrid> value2 = keyValuePair2.Value;
                                    if (value2 != null)
                                    {
                                        foreach (ConduitGrid conGrid2 in IntersectingConduits)
                                        {
                                            if (!value2.Any((ConduitGrid r) => r.Conduit.Id == conGrid2.Conduit.Id && r.RefConduit.Id == conGrid2.RefConduit.Id) && !CongridDictionary.Any((KeyValuePair<int, List<ConduitGrid>> r) => r.Value.Any((ConduitGrid x) => x.Conduit.Id == conGrid2.Conduit.Id)))
                                            {
                                                value2.Add(conGrid2);
                                            }
                                        }

                                        CongridDictionary[keyValuePair2.Key] = value2;
                                    }
                                }
                            }
                        }

                        int num = index;
                        index = num + 1;
                    }
                    else
                    {
                        KeyValuePair<int, List<ConduitGrid>> keyValuePair3 = CongridDictionary.FirstOrDefault((KeyValuePair<int, List<ConduitGrid>> r) => r.Value.Any((ConduitGrid x) => IntersectingConduits.Any((ConduitGrid y) => y.Conduit.Id == x.Conduit.Id)));
                        List<ConduitGrid> value3 = keyValuePair3.Value;
                        if (value3 != null)
                        {
                            foreach (ConduitGrid conGrid in IntersectingConduits)
                            {
                                if (!value3.Any((ConduitGrid r) => r.Conduit.Id == conGrid.Conduit.Id && r.RefConduit.Id == conGrid.RefConduit.Id) && !CongridDictionary.Any((KeyValuePair<int, List<ConduitGrid>> r) => r.Value.Any((ConduitGrid x) => x.Conduit.Id == conGrid.Conduit.Id)))
                                {
                                    value3.Add(conGrid);
                                }
                            }

                            CongridDictionary[keyValuePair3.Key] = value3;
                        }
                    }

                    list3.Add(congrid.Conduit);
                }
            }
            return CongridDictionary;
        }
        public static XYZ GetIntersection(Autodesk.Revit.DB.Element element, ConduitGrid conGrid, XYZ Point, double maximumSpacing = 1.0)
        {
            try
            {
                double num = ((element.GetType() == typeof(Autodesk.Revit.DB.Electrical.Conduit)) ? element.LookupParameter("Outside Diameter").AsDouble() : element.LookupParameter("Width").AsDouble());
                double num2 = ((conGrid.Conduit.GetType() == typeof(Autodesk.Revit.DB.Electrical.Conduit)) ? conGrid.Conduit.LookupParameter("Outside Diameter").AsDouble() : conGrid.Conduit.LookupParameter("Width").AsDouble());
                double num3 = num / 2.0;
                double num4 = num2 / 2.0;
                double value = num3 + num4 + maximumSpacing;
                XYZ direction = conGrid.ConduitLine.Direction;
                XYZ xYZ = direction.CrossProduct(XYZ.BasisZ);
                XYZ endpoint = Point + xYZ.Multiply(value);
                XYZ endpoint2 = Point - xYZ.Multiply(value);
                Line line = Line.CreateBound(endpoint, endpoint2);
                Line line2 = (element.Location as LocationCurve).Curve as Line;
                XYZ endPoint = line2.GetEndPoint(0);
                XYZ endPoint2 = line2.GetEndPoint(1);
                endPoint = new XYZ(endPoint.X, endPoint.Y, 0.0);
                endPoint2 = new XYZ(endPoint2.X, endPoint2.Y, 0.0);
                Line line3 = Line.CreateBound(endPoint, endPoint2);
                return GetIntersection(line3, line);
            }
            catch
            {
                Line lineOne = (element.Location as LocationCurve).Curve as Line;
                Line lineTwo = (conGrid.Conduit.Location as LocationCurve).Curve as Line;
                XYZ endPointOne = lineOne.GetEndPoint(0);
                XYZ endPointTwo = lineTwo.GetEndPoint(0);
                double num = ((element.GetType() == typeof(Autodesk.Revit.DB.Electrical.Conduit)) ? element.LookupParameter("Outside Diameter").AsDouble() : element.LookupParameter("Width").AsDouble());
                double num2 = ((conGrid.Conduit.GetType() == typeof(Autodesk.Revit.DB.Electrical.Conduit)) ? conGrid.Conduit.LookupParameter("Outside Diameter").AsDouble() : conGrid.Conduit.LookupParameter("Width").AsDouble());
                double num3 = num / 2.0;
                double num4 = num2 / 2.0;
                double value = num3 + num4 + maximumSpacing;
                double distance = new XYZ(endPointOne.X, endPointOne.Y, 0).DistanceTo(new XYZ(endPointTwo.X, endPointTwo.Y, 0));
                if (distance <= value)
                {
                    return endPointTwo;
                }
                return null;
            }

            return null;
        }
        public static XYZ GetIntersection(Line line1, Line line2)
        {
            IntersectionResultArray resultArray;
            SetComparisonResult setComparisonResult = line1.Intersect(line2, out resultArray);
            if (setComparisonResult != SetComparisonResult.Overlap)
            {
                return null;
            }

            if (resultArray == null || resultArray.Size != 1)
            {
                return XYZ.Zero;
            }

            IntersectionResult intersectionResult = resultArray.get_Item(0);
            return intersectionResult.XYZPoint;
        }
        public static Dictionary<double, List<Element>> GroupByElementsWithElevation(List<Element> ElementCollection, string offsetVariable)
        {
            Dictionary<int, List<ConduitGrid>> CongridDictionary = GroupStubElements(ElementCollection);
            Dictionary<double, List<Element>> groupedPrimaryElements = new Dictionary<double, List<Element>>();
            foreach (KeyValuePair<int, List<ConduitGrid>> kvp in CongridDictionary)
            {
                if (kvp.Value.Any())
                {
                    List<Element> groupedConduits = kvp.Value.Select(r => r.Conduit).ToList();
                    GroupByElevation(groupedConduits, offsetVariable, ref groupedPrimaryElements);
                }
            }
            return groupedPrimaryElements;
        }
        public static Dictionary<double, List<Autodesk.Revit.DB.Element>> GroupByElevation(List<Autodesk.Revit.DB.Element> a_Elements, string offSetVar, ref Dictionary<double, List<Autodesk.Revit.DB.Element>> groupedElements)
        {
            GetGroupedElementsByElevation(a_Elements, offSetVar, ref groupedElements);
            return groupedElements;
        }
        public static void GetGroupedElementsByElevation(List<Autodesk.Revit.DB.Element> a_Elements, string offSetVar, ref Dictionary<double, List<Autodesk.Revit.DB.Element>> groupedElements)
        {
            Autodesk.Revit.DB.Element element = a_Elements.OrderByDescending((Autodesk.Revit.DB.Element r) => Math.Round(r.LookupParameter(offSetVar).AsDouble(), 8)).FirstOrDefault();
            double num = element.LookupParameter(offSetVar).AsDouble();
            double num2 = element.LookupParameter("Outside Diameter").AsDouble();
            double num3 = num2 / 2.0 + 7.0 / 96.0;
            double refElevation = num - num3;
            Line elementLine = ((element.Location as LocationCurve).Curve as Line);
            if (Math.Round(elementLine.GetEndPoint(0).Z, 4) == Math.Round(elementLine.GetEndPoint(1).Z, 4))
            {
                List<Autodesk.Revit.DB.Element> list = a_Elements.Where((Autodesk.Revit.DB.Element r) => r.LookupParameter(offSetVar).AsDouble() >= refElevation).ToList();

                double middleElevation = list.FirstOrDefault().LookupParameter(offSetVar).AsDouble();
                double topElevation = 0.0;
                double bottomElevation = 0.0;
                string TopOffsetVar = string.Empty;
                string BottomOffsetVar = string.Empty;
                if (list.FirstOrDefault().LookupParameter("Top Elevation") != null)
                {
                    topElevation = list.FirstOrDefault().LookupParameter("Top Elevation").AsDouble();
                    TopOffsetVar = "Top Elevation";
                }
                else if (list.FirstOrDefault().LookupParameter("Upper End Top Elevation") != null)
                {
                    topElevation = list.FirstOrDefault().LookupParameter("Upper End Top Elevation").AsDouble();
                    TopOffsetVar = "Upper End Top Elevation";
                }
                if (list.FirstOrDefault().LookupParameter("Bottom Elevation") != null)
                {
                    bottomElevation = list.FirstOrDefault().LookupParameter("Bottom Elevation").AsDouble();
                    BottomOffsetVar = "Bottom Elevation";
                }
                else if (list.FirstOrDefault().LookupParameter("Upper End Bottom Elevation") != null)
                {
                    bottomElevation = list.FirstOrDefault().LookupParameter("Upper End Bottom Elevation").AsDouble();
                    BottomOffsetVar = "Upper End Bottom Elevation";
                }
                if (list.All((Autodesk.Revit.DB.Element r) => Math.Round(r.LookupParameter(offSetVar).AsDouble(), 2) == Math.Round(middleElevation, 5)))
                {
                    groupedElements.Add(num, list);
                }
                else if (list.All((Autodesk.Revit.DB.Element r) => Math.Round(r.LookupParameter(TopOffsetVar).AsDouble(), 2) == Math.Round(topElevation, 5)))
                {
                    groupedElements.Add(num, list);
                }
                else if (list.All((Autodesk.Revit.DB.Element r) => Math.Round(r.LookupParameter(BottomOffsetVar).AsDouble(), 2) == Math.Round(bottomElevation, 5)))
                {
                    groupedElements.Add(num, list);
                }
                else
                {
                    groupedElements.Add(num, list);
                }
                a_Elements = a_Elements.Except(list).ToList();
            }
            else
            {
                groupedElements.Add(num, a_Elements);
                a_Elements.Clear();
            }
            if (a_Elements.Count > 0)
            {
                GetGroupedElementsByElevation(a_Elements, offSetVar, ref groupedElements);
            }
        }

        /*void CtrlApp_DocumentChanged(object sender, Autodesk.Revit.DB.Events.DocumentChangedEventArgs e)
        {
            if (ToggleConPakToolsButton.ItemText == "AutoConnect ON")
            {
                // get the current document.
                Document doc = e.GetDocument();
                ICollection<ElementId> modifiedElem = e.GetModifiedElementIds();
                try
                {
                    List<Element> elementlist = new List<Element>();
                    List<ElementId> rvConduitlist = new List<ElementId>();
                    string value = string.Empty;
                    foreach (ElementId id in modifiedElem)
                    {
                        Element elem = doc.GetElement(id);
                        if (elem.Category != null && elem.Category.Name == "Conduits")
                        {
                            Parameter parameter = elem.LookupParameter("AutoUpdater BendAngle");
                            value = parameter.AsString();
                            elementlist.Add(elem);
                        }

                    }
                    ChangesInformationForm.instance.MidSaddlePt = elementlist.Distinct().ToList();
                    ChangesInformationForm.instance._elemIdone.Clear();
                    ChangesInformationForm.instance._elemIdtwo.Clear();
                    List<ElementId> FittingElem = new List<ElementId>();
                    for (int i = 0; i < elementlist.Count; i++)
                    {
                        ConnectorSet connector = GetConnectorSet(elementlist[i]);
                        List<ElementId> Icollect = new List<ElementId>();
                        foreach (Connector connect in connector)
                        {
                            ConnectorSet cs1 = connect.AllRefs;
                            foreach (Connector c in cs1)
                            {
                                Icollect.Add(c.Owner.Id);
                            }
                            foreach (ElementId eid in Icollect)
                            {
                                if (doc.GetElement(eid) != null && (doc.GetElement(eid).Category != null &&
                                    doc.GetElement(eid).Category.Name == "Conduit Fittings"))
                                {
                                    FittingElem.Add(eid);
                                }
                            }
                        }
                    }

                    List<ElementId> FittingElements = new List<ElementId>();

                    FittingElements = FittingElem.Distinct().ToList();
                    List<Element> BendElements = new List<Element>();
                    foreach (ElementId id in FittingElements)
                    {
                        BendElements.Add(doc.GetElement(id));


                    }
                    List<ElementId> Icollector = new List<ElementId>();


                    for (int i = 0; i < BendElements.Count; i++)
                    {
                        ConnectorSet connector = GetConnectorSet(BendElements[i]);
                        foreach (Connector connect in connector)
                        {
                            ConnectorSet cs1 = connect.AllRefs;
                            foreach (Connector c in cs1)
                            {
                                Icollector.Add(c.Owner.Id);
                            }
                        }
                    }

                    foreach (ElementId eid in Icollector)
                    {
                        if (doc.GetElement(eid) != null && (doc.GetElement(eid).Category != null &&
                            doc.GetElement(eid).Category.Name == "Conduits"))
                        {
                            ChangesInformationForm.instance._selectedElements.Add(eid);
                        }
                    }
                    List<Element> elementtwo = new List<Element>();
                    List<ElementId> RefID = new List<ElementId>();
                    for (int i = 0; i < BendElements.Count; i++)
                    {
                        for (int j = i + 1; j < BendElements.Count; j++)
                        {
                            Element elemOne = BendElements[i];
                            Element elemTwo = BendElements[j];
                            Parameter parameter = elemOne.LookupParameter("Angle");
                            if (parameter.AsValueString() == "90.00°" || parameter.AsValueString() == "90°")
                            {
                                ConnectorSet firstconnector = GetConnectorSet(elemOne);
                                foreach (Connector connector in firstconnector)
                                {
                                    ConnectorSet cs1 = connector.AllRefs;
                                    foreach (Connector c in cs1)
                                    {
                                        RefID.Add(c.Owner.Id);
                                    }
                                }
                            }
                            ChangesInformationForm.instance._refConduitKick.AddRange(RefID);
                            ChangesInformationForm.instance._Value = value;
                            if (elemOne != null)
                            {
                                ConnectorSet firstconnector = GetConnectorSet(elemOne);
                                ConnectorSet secondconnector = GetConnectorSet(elemTwo);
                                try
                                {
                                    List<ElementId> IDone = new List<ElementId>();
                                    foreach (Connector connector in firstconnector)
                                    {
                                        ConnectorSet cs1 = connector.AllRefs;
                                        foreach (Connector c in cs1)
                                        {
                                            IDone.Add(c.Owner.Id);
                                        }
                                        foreach (ElementId eid in IDone)
                                        {
                                            if (doc.GetElement(eid) != null && (doc.GetElement(eid).Category != null && doc.GetElement(eid).Category.Name == "Conduits"))
                                            {
                                                ChangesInformationForm.instance._elemIdone.Add(eid);
                                            }
                                        }
                                    }


                                    List<ElementId> IDtwo = new List<ElementId>();
                                    foreach (Connector connector in secondconnector)
                                    {
                                        ConnectorSet cs1 = connector.AllRefs;
                                        foreach (Connector c in cs1)
                                        {
                                            IDtwo.Add(c.Owner.Id);
                                        }
                                        foreach (ElementId eid in IDtwo)
                                        {
                                            if (doc.GetElement(eid) != null && (doc.GetElement(eid).Category != null && doc.GetElement(eid).Category.Name == "Conduits"))
                                            {
                                                ChangesInformationForm.instance._elemIdtwo.Add(eid);
                                                if (ChangesInformationForm.instance._elemIdone.Any(r => r == eid))
                                                {
                                                    ChangesInformationForm.instance._deletedIds.Add(eid);
                                                    rvConduitlist.Add(eid);

                                                }
                                            }
                                        }
                                    }
                                    ChangesInformationForm.instance._deletedIds.Add(elemOne.Id);
                                    ChangesInformationForm.instance._deletedIds.Add(elemTwo.Id);
                                }
                                catch
                                {

                                }

                            }
                        }
                    }
                    try
                    {
                        var l = rvConduitlist.Distinct();
                        ChangesInformationForm.instance._selectedElements = ChangesInformationForm.instance._selectedElements.Except(l).ToList();
                        AngleDrawHandler handler = new AngleDrawHandler();
                        ExternalEvent DrawEvent = ExternalEvent.Create(handler);
                        DrawEvent.Raise();

                    }
                    catch
                    {
                        MessageBox.Show("Error");
                    }
                }
                catch { }
            }

        }*/
        #endregion

        #region Class Methods
        /// <summary>
        /// This method is used to retrieve the changed element and add row to data table.
        /// </summary>
        /// <param name="id"></param>
        /// <param name="doc"></param>
        /// <param name="changeType"></param>
        public void AddChangeInfoRow(ElementId id, Document doc, string changeType)
        {
            // retrieve the changed element
            Element elem = doc.GetElement(id);

            MessageBox.Show(elem.Id.ToString());
            DataRow newRow = m_ChangesInfoTable.NewRow();
            if (elem != null)
            {

                Element primaryelement = null;
                ConnectorSet firstconnector = Utility.GetConnectorSet(elem);
                //ConnectorSet secondconnector = Utility.GetConnectorSet(elem);
                try
                {
                    foreach (Connector connector in firstconnector)
                    {
                        primaryelement = connector.Owner;
                        MessageBox.Show(primaryelement.Id.ToString());
                    }
                }
                catch
                {
                    MessageBox.Show("error");
                }
                Parameter parameter = elem.LookupParameter("Bend Angle");
                string value = parameter.AsString();
                // Parameter p = elem.get_Parameter(BuiltInParameter.CONDUIT_STANDARD_TYPE_PARAM);
                // Pname = p.Definition.Name;
                // string changes = string.Empty;
                List<string> list = new List<string>();
                /* Element elemtcollec = new FilteredElementCollector(doc).OfClass(typeof(Conduit)).FirstOrDefault();
                 try
                 {
                     foreach (var p in elem.GetOrderedParameters().Where(x => !x.IsReadOnly).ToList())
                     {
                         Parameter parameter1 = elem.LookupParameter(p.Definition.Name);
                         if (parameter1.AsString() != value)
                         {
                             //list.Add(p.Definition.Name);
                             //changes = parameter1.AsString();
                         }
                         else
                         {

                         }
                     }
                 }
                 catch
                 {
                     return;
                 }*/
                if (elem.Category.Name == "Center line")
                {
                    //MessageBox.Show("Category : " + elem.Category.Name + " \n" + "ID :" + id.ToString() + " in " + string.Join(",", list));

                }
                else
                {
                    if (value != null)
                        MessageBox.Show("Category : " + elem.Category.Name + " \n" + "ID : " + id.ToString() + "\n" + " Bend Angle : " + value);
                    //list.Clear();
                }
                // return value;
            }

        }

        public static ConnectorSet GetConnectorSet(Autodesk.Revit.DB.Element Ele)
        {
            ConnectorSet result = null;
            if (Ele is Autodesk.Revit.DB.FamilyInstance)
            {
                MEPModel mEPModel = ((Autodesk.Revit.DB.FamilyInstance)Ele).MEPModel;
                if (mEPModel != null && mEPModel.ConnectorManager != null)
                {
                    result = mEPModel.ConnectorManager.Connectors;
                }
            }
            else if (Ele is MEPCurve)
            {
                result = ((MEPCurve)Ele).ConnectorManager.Connectors;
            }

            return result;
        }

        /// <summary>
        /// Generate a data table with five columns for display in window
        /// </summary>
        /// <returns>The DataTable to be displayed in window</returns>
        private DataTable CreateChangeInfoTable()
        {
            // create a new dataTable
            DataTable changesInfoTable = new DataTable("ChangesInfoTable");

            // Create a "ChangeType" column. It will be "Added", "Deleted" and "Modified".
            DataColumn styleColumn = new DataColumn("ChangeType", typeof(System.String));
            styleColumn.Caption = "ChangeType";
            changesInfoTable.Columns.Add(styleColumn);

            // Create a "Id" column. It will be the Element ID
            DataColumn idColum = new DataColumn("Id", typeof(System.String));
            idColum.Caption = "Id";
            changesInfoTable.Columns.Add(idColum);

            // Create a "Name" column. It will be the Element Name
            DataColumn nameColum = new DataColumn("Name", typeof(System.String));
            nameColum.Caption = "Name";
            changesInfoTable.Columns.Add(nameColum);

            // Create a "Category" column. It will be the Category Name of the element.
            DataColumn categoryColum = new DataColumn("Category", typeof(System.String));
            categoryColum.Caption = "Category";
            changesInfoTable.Columns.Add(categoryColum);

            // Create a "Document" column. It will be the document which own the changed element.
            DataColumn docColum = new DataColumn("Document", typeof(System.String));
            docColum.Caption = "Document";
            changesInfoTable.Columns.Add(docColum);

            // return this data table 
            return changesInfoTable;
        }
        #endregion
    }

    /// <summary>
    /// This class inherits IExternalCommand interface and used to retrieve the dialog again.
    /// </summary>
    [Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)]
    [Autodesk.Revit.Attributes.Regeneration(Autodesk.Revit.Attributes.RegenerationOption.Manual)]
    public class Command : IExternalCommand
    {

        public List<ElementId> _deletedIds = new List<ElementId>();
        #region IExternalCommand Members
        /// <summary>
        /// Implement this method as an external command for Revit.
        /// </summary>
        /// <param name="commandData">An object that is passed to the external application
        /// which contains data related to the command,
        /// such as the application object and active view.</param>
        /// <param name="message">A message that can be set by the external application
        /// which will be displayed if a failure or cancellation is returned by
        /// the external command.</param>
        /// <param name="elements">A set of elements to which the external application
        /// can add elements that are to be highlighted in case of failure or cancellation.</param>
        /// <returns>Return the status of the external command.
        /// A result of Succeeded means that the API external method functioned as expected.
        /// Cancelled can be used to signify that the user cancelled the external operation 
        /// at some point. Failure should be returned if the application is unable to proceed with
        /// the operation.</returns>
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            /*if (ExternalApplication.InfoForm == null)
            {
                ExternalApplication.InfoForm = new ChangesInformationForm(ExternalApplication.ChangesInfoTable);
            }
            ExternalApplication.InfoForm.Show();*/
            ExternalApplication.Toggle();
            UIDocument uIDocument = commandData.Application.ActiveUIDocument;
            Document doc = uIDocument.Document;
            if (doc.IsReadOnly)
            {
                MessageBox.Show("doc is read Only");
            }

            return Result.Succeeded;
        }
        #endregion
    }
}





