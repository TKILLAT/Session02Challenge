#region Namespaces
using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Security.Cryptography.X509Certificates;

#endregion

namespace Session02Challenge
{
    [Transaction(TransactionMode.Manual)]
    public class Command : IExternalCommand
    {
        public Result Execute(
          ExternalCommandData commandData,
          ref string message,
          ElementSet elements)
        {
            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Application app = uiapp.Application;
            Document doc = uidoc.Document;

            string filepathLevels = "C:\\Users\\tthomas\\OneDrive - AKF Group LLC\\Desktop\\Revit Addin Bootcamp\\Session 2 Challenge\\RAB_Session_02_Challenge_Levels.csv";
            string filepathSheets = "C:\\Users\\tthomas\\OneDrive - AKF Group LLC\\Desktop\\Revit Addin Bootcamp\\Session 2 Challenge\\RAB_Session_02_Challenge_Sheets.csv";
            string[] levelsArray = System.IO.File.ReadAllLines(filepathLevels);
            string[] sheetsArray = System.IO.File.ReadAllLines(filepathSheets);
            List<string> levels = levelsArray.ToList();
            List<string> sheets = sheetsArray.ToList();

            levels.RemoveAt(0);
            sheets.RemoveAt(0);


            Transaction t = new Transaction(doc);
            t.Start("Create Levels and Sheets");


                //create levels

                
            foreach(string rowlevels in levels)
            {
               
                string[] celllevelString = rowlevels.Split(',');

                string levelName = celllevelString[0];
                string levelElev = celllevelString[1];



                double levelelevDoub = double.Parse(levelElev);
                    

                Level myLevel = Level.Create(doc,levelelevDoub);
                myLevel.Name = levelName;

            }

        

            //collect titleblocks
            FilteredElementCollector collector = new FilteredElementCollector(doc);
            collector.OfCategory(BuiltInCategory.OST_TitleBlocks);
            ElementId tblocks = collector.FirstElementId();
          

            //create sheets
            foreach(string rowsheetsArray in sheets)
            {
                string[] cellsheetString = rowsheetsArray.Split(',');
                string sheetNumber = cellsheetString[0];
                string sheetName = cellsheetString[1];

                ViewSheet mySheet = ViewSheet.Create(doc, tblocks);
                mySheet.Name = sheetName;
                mySheet.SheetNumber = sheetNumber;


            }

            t.Commit();
            t.Dispose();

            return Result.Succeeded;
        }
    }
}
