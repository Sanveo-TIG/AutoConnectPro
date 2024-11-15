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
                BitmapImage OffLargeImage = new BitmapImage(new Uri("pack://application:,,,/AutoConnectPro;component/Resources/off 32x32.png"));
                BitmapImage OnImage = new BitmapImage(new Uri("pack://application:,,,/AutoConnectPro;component/Resources/switch-on 16x16.png"));
                BitmapImage OnLargeImage = new BitmapImage(new Uri("pack://application:,,,/AutoConnectPro;component/Resources/on 32x32.png"));
                BitmapImage OffImage = new BitmapImage(new Uri("pack://application:,,,/AutoConnectPro;component/Resources/switch-off 16x16.png"));
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
                        if (panel.Source.Title == "Auto Connect")
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
            BitmapImage pb1Image = new BitmapImage(new Uri("pack://application:,,,/AutoConnectPro;component/Resources/off 32x32.png"));
            buttondata.LargeImage = pb1Image;
            BitmapImage pb1Image2 = new BitmapImage(new Uri("pack://application:,,,/AutoConnectPro;component/Resources/switch-off 16x16.png"));
            buttondata.Image = pb1Image2;
            buttondata.AvailabilityClassName = "Revit.SDK.Samples.AutoConnectPro.CS.Availability";

            #region Sample PushButton 
            PushButtonData buttondataSample1 = new PushButtonData("ModifierBtnCommandAutoConnect", "AutoConnect", dllLocation, "AutoConnectPro.AutoConnectCommand");
            BitmapImage pb1ImageSample11 = new BitmapImage(new Uri("pack://application:,,,/AutoConnectPro;component/Resources/on 32x32.png"));
            buttondataSample1.LargeImage = pb1ImageSample11;
            BitmapImage pb1ImageSample12 = new BitmapImage(new Uri("pack://application:,,,/AutoConnectPro;component/Resources/on 32x32.png"));
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
            string ribbonPanelText = "Auto Connect"; // Architecture

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
        private void OnIdling(object sender, Autodesk.Revit.UI.Events.IdlingEventArgs e)
        {
            try
            {
                if (ToggleConPakToolsButton.ItemText == "AutoConnect ON")
                {
                    List<Element> selectedElements = new List<Element>();
                    UIApplication uiApp = sender as UIApplication;
                    UIDocument uiDoc = uiApp.ActiveUIDocument;
                    Document doc = uiDoc.Document;

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
                                                    if (isErrorOccuredinAutoConnect)
                                                    {
                                                        BitmapImage OffLargeImage = new BitmapImage(new Uri("pack://application:,,,/AutoConnectPro;component/Resources/off 32x32.png"));
                                                        BitmapImage OffImage = new BitmapImage(new Uri("pack://application:,,,/AutoConnectPro;component/Resources/switch-off 16x16.png"));
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
                                                string panelNameAC = "Auto Connect";
                                                string executableLocation = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
                                                string dllLocation = Path.Combine(executableLocation, "AutoUpdaterPro.dll");
                                                List<Autodesk.Revit.UI.RibbonPanel> panels = uiApp.GetRibbonPanels(tabName);
                                                Autodesk.Revit.UI.RibbonPanel autoUpdaterPanel01 = panels.FirstOrDefault(p => p.Name == panelName);
                                                Autodesk.Revit.UI.RibbonPanel autoUpdaterPanel02 = panels.FirstOrDefault(p => p.Name == panelNameAC);
                                                bool ErrorOccured = false;
                                                if (autoUpdaterPanel01 != null)
                                                {
                                                    IList<RibbonItem> items = autoUpdaterPanel01.GetItems();
                                                    foreach (RibbonItem item in items)
                                                    {
                                                        if (item is PushButton pushButton && pushButton.ItemText == "AutoUpdate ON")
                                                        {
                                                            ErrorOccured = true;
                                                        }
                                                        else if (item.ItemText == "AutoUpdate" && !item.Enabled && autoUpdaterPanel02.GetItems().OfType<PushButton>().Any(btn => btn.ItemText == "AutoConnect ON")
                                                            && autoUpdaterPanel01.GetItems().OfType<PushButton>().Any(btn => btn.ItemText == "AutoUpdate OFF"))
                                                        {
                                                            ErrorOccured = true;
                                                            BitmapImage OffLargeImage = new BitmapImage(new Uri("pack://application:,,,/AutoConnectPro;component/Resources/off 32x32.png"));
                                                            BitmapImage OffImage = new BitmapImage(new Uri("pack://application:,,,/AutoConnectPro;component/Resources/switch-off 16x16.png"));
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
                                                        System.Windows.MessageBox.Show("Please select two equal sets of conduits", "Warning-AutoConnect", MessageBoxButton.OK, MessageBoxImage.Warning);
                                                        SelectedElements.Clear();
                                                        uiDoc.Selection.SetElementIds(new List<ElementId> { ElementId.InvalidElementId });
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
                                    }
                                }
                            }
                        }
                    }
                }
            }
            catch (Exception)
            {
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
