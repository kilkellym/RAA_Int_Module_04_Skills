#region Namespaces
using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Architecture;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Diagnostics;
using System.Linq;
using System.Reflection;
using Forms = System.Windows.Forms;

#endregion

namespace RAA_Int_Module_04_Skills
{
    [Transaction(TransactionMode.Manual)]
    public class Command1 : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            // this is a variable for the Revit application
            UIApplication uiapp = commandData.Application;

            // this is a variable for the current Revit model
            Document doc = uiapp.ActiveUIDocument.Document;

            // 1. get all links
            FilteredElementCollector linkCollector = new FilteredElementCollector(doc)
                .OfClass(typeof(RevitLinkType));

            // 2. loop through links and get doc if loaded
            Document linkedDoc = null;
            RevitLinkInstance link = null;

            //foreach (RevitLinkType rvtLink in linkCollector)
            //{
            //    if (rvtLink.GetLinkedFileStatus() == LinkedFileStatus.Loaded)
            //    {
            //        link = new FilteredElementCollector(doc)
            //            .OfCategory(BuiltInCategory.OST_RvtLinks)
            //            .OfClass(typeof(RevitLinkInstance))
            //            .Where(x => x.GetTypeId() == rvtLink.Id).First() as RevitLinkInstance;

            //        linkedDoc = link.GetLinkDocument();
            //    }
            //}

            //FilteredElementCollector collectorA = new FilteredElementCollector(linkedDoc)
            //    .OfCategory(BuiltInCategory.OST_Rooms);

            //TaskDialog.Show("Test Method A", $"There are {collectorA.Count()} rooms in the linked model.");

            // Method B
            // 1. prompt user to select RVT file
            // NOTE: be sure to add reference to System.Windows.Forms
            //string revitFile = "";

            //Forms.OpenFileDialog ofd = new Forms.OpenFileDialog();
            //ofd.Title = "Select Revit file";
            //ofd.InitialDirectory = @"C:\";
            //ofd.Filter = "Revit files (*.rvt)|*.rvt";
            //ofd.RestoreDirectory = true;

            //if (ofd.ShowDialog() != Forms.DialogResult.OK)
            //    return Result.Failed;

            //revitFile = ofd.FileName;

            //// 2. Open selected file in background
            //UIDocument closedUIDoc = uiapp.OpenAndActivateDocument(revitFile);
            //Document closedDoc = closedUIDoc.Document;

            //// 3. get elements from opened file
            //FilteredElementCollector collectorB = new FilteredElementCollector(closedDoc)
            //    .OfCategory(BuiltInCategory.OST_Rooms);

            //TaskDialog.Show("Test Method B", $"There are {collectorB.Count()} rooms in the modele I just opened.");

            // 4. Make other document active then close document
            //uiapp.OpenAndActivateDocument(doc.PathName);
            //closedDoc.Close(false);

            // Method C: Get open file with specific name
            // 1. create document variable
            Document openDoc = null;

            // 2. loop through open documents and look for match
            foreach (Document curDoc in uiapp.Application.Documents)
            {
                if (curDoc.PathName.Contains("rac"))
                    openDoc = curDoc;
            }

            // create space from room
            // use LINQ to get a room
            Room curRoom = new FilteredElementCollector(openDoc)
                .OfCategory(BuiltInCategory.OST_Rooms)
                .Cast<Room>()
                .First();

            // get level from current view
            Level curLevel = doc.ActiveView.GenLevel;

            // get room data
            string roomName = curRoom.Name;
            string roomNum = curRoom.Number;
            string roomComments = curRoom.LookupParameter("Comments").AsString();

            // get room location point
            LocationPoint roomPoint = curRoom.Location as LocationPoint;

            using(Transaction t = new Transaction(doc))
            {
                t.Start("Create space");
                // create space and transfer properties
                SpatialElement newSpace = doc.Create.NewSpace
                    (curLevel, new UV(roomPoint.Point.X, roomPoint.Point.Y));
                newSpace.Name = roomName;
                newSpace.Number = roomNum;
                newSpace.LookupParameter("Comments").Set(roomComments);
                t.Commit();
            }

            // inserting groups 
            // get group type by name
            string groupName = "Group 1";

            GroupType curGroup = new FilteredElementCollector(doc)
                .OfCategory(BuiltInCategory.OST_IOSModelGroups)
                .WhereElementIsElementType()
                .Where(r => r.Name == groupName)
                .Cast<GroupType>().First();

            // insert group
            XYZ insPoint = new XYZ();

            using(Transaction t = new Transaction(doc))
            {
                t.Start("Place group");
                doc.Create.PlaceGroup(insPoint, curGroup);
                t.Commit();
            }

            // copy elements from other doc
            // create filtered element collector to get elements
            FilteredElementCollector wallCollector = new FilteredElementCollector(openDoc)
                .OfCategory(BuiltInCategory.OST_Walls)
                .WhereElementIsNotElementType();

            // get list of element Ids
            List<ElementId> elemIdList = wallCollector.Select(elem => elem.Id).ToList();

            // copy elements 
            CopyPasteOptions options = new CopyPasteOptions();
            
            using (Transaction t = new Transaction(doc))
            {
                t.Start("Copy elements");
                ElementTransformUtils.CopyElements(openDoc, elemIdList, doc, null, options);
                t.Commit();
            }

            return Result.Succeeded;
        }
        internal static PushButtonData GetButtonData()
        {
            // use this method to define the properties for this command in the Revit ribbon
            string buttonInternalName = "btnCommand1";
            string buttonTitle = "Button 1";

            ButtonDataClass myButtonData1 = new ButtonDataClass(
                buttonInternalName,
                buttonTitle,
                MethodBase.GetCurrentMethod().DeclaringType?.FullName,
                Properties.Resources.Blue_32,
                Properties.Resources.Blue_16,
                "This is a tooltip for Button 1");

            return myButtonData1.Data;
        }
    }
}
