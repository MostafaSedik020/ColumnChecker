﻿using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Media.Imaging;
using System.Windows.Media;

namespace ColumnChecker.Entry
{
    public class ExtApp : IExternalApplication
    {
        private UIControlledApplication uicApp;
        public Result OnShutdown(UIControlledApplication uicApp)
        {
            return Result.Succeeded;
        }

        public Result OnStartup(UIControlledApplication uicApp)
        {
            this.uicApp = uicApp;
            CreatePushButton();

            return Result.Succeeded;
        }

        private void CreatePushButton()
        {
            string tabName = "TrussT";
            string panelName = "Structure";

            try
            {
                // Create the ribbon tab if it doesn't already exist
                uicApp.CreateRibbonTab(tabName);
            }
            catch (Exception)
            {
                // Tab already exists, continue
            }

            // Get or create the ribbon panel
            RibbonPanel panel = uicApp.GetRibbonPanels(tabName)
                                       .FirstOrDefault(p => p.Name == panelName)
                                       ?? uicApp.CreateRibbonPanel(tabName, panelName);

            // Ensure the panel was created successfully
            if (panel != null)
            {
                // Get the executing assembly path

                Assembly assembly = Assembly.GetExecutingAssembly();
                string assemblyPath = assembly.Location;

                // Create push button data
                PushButtonData pbData = new PushButtonData("ColCheck_btn", "Column Checker", assemblyPath, typeof(ExtCmd).FullName);

                // Add the push button to the panel
                PushButton pb = panel.AddItem(pbData) as PushButton;
                pb.ToolTip = "Check ETABS columns with Revit columns ";

                pb.LargeImage = GetImageSource("ColumnChecker.Resources.checklist-24.png");
            }
        }
        private ImageSource GetImageSource(string imageFullName)
        {
            Stream stream = Assembly.GetExecutingAssembly().GetManifestResourceStream(imageFullName);
            PngBitmapDecoder decoder = new PngBitmapDecoder(stream, BitmapCreateOptions.PreservePixelFormat, BitmapCacheOption.Default);

            return decoder.Frames[0];
        }
    }
}
