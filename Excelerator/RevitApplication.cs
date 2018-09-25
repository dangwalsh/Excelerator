namespace Gensler.Revit.Excelerator
{
    using System;
    using System.Reflection;
    using Autodesk.Revit.UI;
    using System.Windows.Media.Imaging;

    class RevitApplication : IExternalApplication
    {
        static void AddRibbonPanel(UIControlledApplication application)
        {
            const string tabName = "Atlanta";
            application.CreateRibbonTab(tabName);
            
            var ribbonPanel = application.CreateRibbonPanel(tabName, "Tools");
            var thisAssemblyPath = Assembly.GetExecutingAssembly().Location;
            var pushButtonData = new PushButtonData("Excelerator", "Import" + Environment.NewLine + "Excel", thisAssemblyPath, "Gensler.Revit.Excelerator.RevitCommand");

            if (!(ribbonPanel.AddItem(pushButtonData) is PushButton pb1)) return;

            pb1.ToolTip = "Create a Key Schedule from Excel data";
            var pb1Image = new BitmapImage(new Uri("pack://application:,,,/Excelerator;component/Resources/excel.png"));
            pb1.LargeImage = pb1Image;
        }

        public Result OnShutdown(UIControlledApplication application)
        {
            return Result.Succeeded;
        }

        public Result OnStartup(UIControlledApplication application)
        {
            AddRibbonPanel(application);
            return Result.Succeeded;
        }
    }
}
