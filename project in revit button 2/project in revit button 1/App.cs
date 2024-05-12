using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Events;
using Autodesk.Revit.UI.Selection;
using Autodesk.Revit.DB.Plumbing;
using Autodesk.Revit.DB.Mechanical;
using Autodesk.Revit.DB.Electrical;
using Autodesk.Revit.DB.Analysis;
using Autodesk.Revit.DB.Architecture;
using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;


//using Autodesk.Revit.Collections;




using System.Diagnostics;
using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Media.Imaging;

using System.Windows.Forms;
using AJC_Commands_1;
using DBLibrary;
using System.Windows.Controls.Primitives;

using System.Reflection;
using RvtApplication = Autodesk.Revit.ApplicationServices.Application;
using RvtDocument = Autodesk.Revit.DB.Document;
using OfficeOpenXml;
using Rhino5x64;
using RhinoScript4;
using Rhino.FileIO;
using Rhino.Geometry;
using Rhino.Collections;
using Rhino;
using Rhino.Input;
using RhinoCon;
using winform = System.Windows.Forms;
using System.Windows.Media;


using System.Data.SqlClient;

// use an alias because Autodesk.Revit.UI 
// uses classes which have same names:

using adWin = Autodesk.Windows;

namespace BoostYourBIM
{
    
    [Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)]
    public class PlaceView_CrateSheet : IExternalCommand
    {
        static AddInId appId = new AddInId(new Guid("F5FB1A7F-8110-4862-8820-04AE05C1239E"));
        public Autodesk.Revit.UI.Result Execute(ExternalCommandData commandData, ref string message, ElementSet elementSet)
        {
            UIDocument uidoc = commandData.Application.ActiveUIDocument;
            Autodesk.Revit.DB.Document doc = uidoc.Document;

            //try
            //{
            //    string filename = @"T:\Transfer\lopez\Book1.xlsx";
            //    using (ExcelPackage package = new ExcelPackage(new FileInfo(filename)))
            //    {
            //        ExcelWorksheet sheet = package.Workbook.Worksheets.ElementAt(0);

            //        int number = Convert.ToInt32(sheet.Cells[2, 1].Value);
            //        sheet.Cells[2, 1].Value = (number + 1); ;
            //        package.Save();
            //    }
            //}
            //catch (Exception)
            //{
            //    MessageBox.Show("Excel file not found", "");
            //}

            
            //---------------------------------------- FILTERS ------------------------------------
            IEnumerable<FamilySymbol> familyList = from elem in new FilteredElementCollector(doc).OfClass(typeof(FamilySymbol)).OfCategory(BuiltInCategory.OST_TitleBlocks)
                                                   let type = elem as FamilySymbol
                                                   select type;

            List<Element> viewElems = new List<Element>();
            FilteredElementCollector collector = new FilteredElementCollector(doc);
            viewElems.AddRange(collector.OfClass(typeof(Autodesk.Revit.DB.View)).ToElements());

            List<ElementId> lista_de_views = new List<ElementId>();
            List<ElementId> tittleblocks_list = new List<ElementId>();
            List<String> tittle_block_list_name = new List<String>();
            List<string> nombres_hojas = new List<string>();
            List<string> nameofviews = new List<string>();
            List<Element> hojitas = new List<Element>();
            List<ViewSchedule> ViewSchedule_LIST = new List<ViewSchedule>();
            List<string> gourpheader_list = new List<string>();

            Form13 form2 = new Form13();

            form2.comboBox1.Items.Add("Plans");
            form2.comboBox1.Items.Add("3D views");
            form2.comboBox1.Items.Add("Section & Elevation views");
            form2.comboBox1.Items.Add("Drafting");

            FilteredElementCollector collector2 = new FilteredElementCollector(doc);
            ICollection<Element> hojas = collector2.OfClass(typeof(Autodesk.Revit.DB.View)).ToElements();

            FilteredElementCollector fec = new FilteredElementCollector(doc);
            fec.OfClass(typeof(ViewSection));

            IEnumerable<ViewSheet> viewSheet = from elem in new FilteredElementCollector(doc)
                                                .OfClass(typeof(ViewSheet))
                                                .OfCategory(BuiltInCategory.OST_Sheets)
                                               let type = elem as ViewSheet
                                               where type.Name != null
                                               select type;


            IEnumerable<Autodesk.Revit.DB.View> view_list = from elem in new FilteredElementCollector(doc)
                                                .OfClass(typeof(Autodesk.Revit.DB.View))
                                                .OfCategory(BuiltInCategory.OST_Views)
                                                            let type = elem as Autodesk.Revit.DB.View
                                                            where type.Name != null
                                                            select type;

            List<ElementId> ids_ = new List<ElementId>();


            foreach (var sheet_ in viewSheet)
            {
                ICollection<ElementId> views_ = sheet_.GetAllPlacedViews();
                foreach (var item in views_)
                {
                    ids_.Add(item);
                }
            }


            form2.ShowDialog();

            if (form2.DialogResult == DialogResult.Cancel)
            {
                return Autodesk.Revit.UI.Result.Cancelled;
            }
            if (form2.comboBox1.SelectedItem ==  null)
            {
                TaskDialog.Show("Instruction", "You must select a type of view");
                return Autodesk.Revit.UI.Result.Cancelled;
            }

            if (form2.comboBox1.SelectedItem.ToString() == "Section & Elevation views")
            {
                var viewPlans = fec.Cast<ViewSection>().Where<ViewSection>(vp => vp.IsTemplate);
                IEnumerable<ViewSection> viewList = from elem in new FilteredElementCollector(doc)
                                                     .OfClass(typeof(ViewSection))
                                                     .OfCategory(BuiltInCategory.OST_Views)
                                                    let type = elem as ViewSection
                                                    where type.Name != null
                                                    select type;

                List<ViewSection> SecView2 = new List<ViewSection>();
                List<ViewSection> Views = new List<ViewSection>();

                foreach (var view in viewList)
                {
                    if (!view.IsTemplate)
                    {
                        Views.Add(view);

                    }
                }
                foreach (var id1 in viewList)
                {
                    foreach (var viewid in ids_)
                    {
                        if (!id1.Id.IntegerValue.Equals(viewid.IntegerValue))
                        {

                        }
                        else
                        {
                            if (!SecView2.Contains(id1))
                            {
                                SecView2.Add(id1);

                            }
                        }
                    }
                }

                for (int i = 0; i < Views.ToArray().Length; i++)
                {
                    foreach (var item in SecView2)
                    {
                        if (Views.ToArray()[i].Name == item.Name)
                        {
                            Views.RemoveAt(i);
                        }

                    }
                }
                foreach (Autodesk.Revit.DB.View i in Views) //add Name & id to project views list
                {
                    if (!i.IsTemplate)
                    {

                        lista_de_views.Add(i.Id);

                        nameofviews.Add(i.Name);

                        hojitas.Add(i);

                        nombres_hojas.Add(i.Name);

                    }

                }
            }

            if (form2.comboBox1.SelectedItem.ToString() == "3D views")
            {
                IEnumerable<View3D> viewList3d = from elem in new FilteredElementCollector(doc)
                                                .OfClass(typeof(View3D))
                                                .OfCategory(BuiltInCategory.OST_Views)
                                                 let type = elem as View3D
                                                 //where type.Name.Contains("SCHEDULE")
                                                 select type;

                List<Autodesk.Revit.DB.View> SecView2 = new List<Autodesk.Revit.DB.View>();
                List<Autodesk.Revit.DB.View> Views = new List<Autodesk.Revit.DB.View>();

                foreach (var view in viewList3d)
                {
                    if (!view.IsTemplate)
                    {
                        Views.Add(view);

                    }
                }
                foreach (var id1 in viewList3d)
                {
                    foreach (var viewid in ids_)
                    {
                        if (!id1.Id.IntegerValue.Equals(viewid.IntegerValue))
                        {

                        }
                        else
                        {
                            if (!SecView2.Contains(id1))
                            {
                                SecView2.Add(id1);

                            }
                        }
                    }
                }

                for (int i = 0; i < Views.ToArray().Length; i++)
                {
                    foreach (var item in SecView2)
                    {
                        if (Views.ToArray()[i].Name == item.Name)
                        {
                            Views.RemoveAt(i);
                        }

                    }
                }
                foreach (Autodesk.Revit.DB.View3D i in Views) //add Name & id to project views list
                {
                    if (!i.IsTemplate)
                    {
                        lista_de_views.Add(i.Id);

                        nameofviews.Add(i.Name);

                        hojitas.Add(i);

                        nombres_hojas.Add(i.Name);
                    }

                }
            }

            if (form2.comboBox1.SelectedItem.ToString() == "Plans")
            {
                IEnumerable<ViewPlan> viewListViewPlan = from elem in new FilteredElementCollector(doc)
                                                .OfClass(typeof(ViewPlan))
                                                .OfCategory(BuiltInCategory.OST_Views)
                                                         let type = elem as ViewPlan
                                                         //where type.Name.Contains("SCHEDULE")
                                                         select type;

                List<Autodesk.Revit.DB.View> SecView2 = new List<Autodesk.Revit.DB.View>();
                List<Autodesk.Revit.DB.View> Views = new List<Autodesk.Revit.DB.View>();

                foreach (var view in viewListViewPlan)
                {
                    if (!view.IsTemplate)
                    {
                        Views.Add(view);

                    }
                }
                foreach (var id1 in viewListViewPlan)
                {
                    foreach (var viewid in ids_)
                    {
                        if (!id1.Id.IntegerValue.Equals(viewid.IntegerValue))
                        {

                        }
                        else
                        {
                            if (!SecView2.Contains(id1))
                            {
                                SecView2.Add(id1);

                            }
                        }
                    }
                }
                for (int i = 0; i < Views.ToArray().Length; i++)
                {
                    foreach (var item in SecView2)
                    {
                        if (Views.ToArray()[i].Name == item.Name)
                        {
                            Views.RemoveAt(i);
                        }

                    }
                }
                foreach (ViewPlan i in Views) //add Name & id to project views list
                {
                    if (!i.IsTemplate)
                    {
                        lista_de_views.Add(i.Id);

                        nameofviews.Add(i.Name);

                        hojitas.Add(i);

                        nombres_hojas.Add(i.Name);
                    }

                }
            }
            if (form2.DialogResult == DialogResult.Cancel)
            {
                return Autodesk.Revit.UI.Result.Cancelled;
            }
            if (form2.comboBox1.SelectedItem.ToString() == "Drafting")
            {
                IEnumerable<ViewDrafting> viewdrafting_ = from elem in new FilteredElementCollector(doc)
                                                .OfClass(typeof(ViewDrafting))
                                                .OfCategory(BuiltInCategory.OST_Views)
                                                          let type = elem as ViewDrafting
                                                          //where type.Name.Contains("SCHEDULE")
                                                          select type;

                List<Autodesk.Revit.DB.View> SecView2 = new List<Autodesk.Revit.DB.View>();
                List<Autodesk.Revit.DB.View> Views = new List<Autodesk.Revit.DB.View>();

                foreach (var view in viewdrafting_)
                {
                    if (!view.IsTemplate)
                    {
                        Views.Add(view);

                    }
                }
                foreach (var id1 in viewdrafting_)
                {
                    foreach (var viewid in ids_)
                    {
                        if (!id1.Id.IntegerValue.Equals(viewid.IntegerValue))
                        {

                        }
                        else
                        {
                            if (!SecView2.Contains(id1))
                            {
                                SecView2.Add(id1);

                            }
                        }
                    }
                }
                for (int i = 0; i < Views.ToArray().Length; i++)
                {
                    foreach (var item in SecView2)
                    {
                        if (Views.ToArray()[i].Name == item.Name)
                        {
                            Views.RemoveAt(i);
                        }

                    }
                }
                foreach (ViewDrafting i in Views) //add Name & id to project views list
                {
                    if (!i.IsTemplate)
                    {
                        lista_de_views.Add(i.Id);

                        nameofviews.Add(i.Name);

                        hojitas.Add(i);

                        nombres_hojas.Add(i.Name);
                    }

                }
            }
            if (form2.DialogResult == DialogResult.Cancel)
            {
                return Autodesk.Revit.UI.Result.Cancelled;
            }
            foreach (FamilySymbol T in familyList) //add Name & id to tittle block list
            {
                ElementId HOJAS_ID = T.Id;
                tittleblocks_list.Add(HOJAS_ID);

                string nombre_hojas_each = T.Name;
                tittle_block_list_name.Add(T.FamilyName + " - " + nombre_hojas_each);
            }

            Form6 form = new Form6();

            foreach (var item in tittle_block_list_name)
            {
                form.comboBox1.Items.Add(item);
            }

            List<Element> store_Selected = new List<Element>(); //nombre hojas_copy

            foreach (var item in tittle_block_list_name)
            {
                form.comboBox1.Items.Add(item);
            }
            foreach (var item in nombres_hojas)
            {
                form.listBox1.Items.Add(item);
            }

            foreach (Autodesk.Revit.DB.View view in viewElems)
            {
                if (!view.IsTemplate && view.CanBePrinted && view.ViewType == ViewType.DrawingSheet)
                {
                    Debug.Print(view.Name);
                    
                    BrowserOrganization org = BrowserOrganization.GetCurrentBrowserOrganizationForSheets(doc);
                    IList<Parameter> param = org.GetOrderedParameters();


                    List<FolderItemInfo> folderfields = org.GetFolderItems(view.Id).ToList();

                    foreach (FolderItemInfo info in folderfields)
                    {
                        string groupheader = info.Name;

                        if (!gourpheader_list.Contains(groupheader))
                        {
                            gourpheader_list.Add(groupheader);
                        }
                        ElementId parameterId = info.ElementId;
                    }
                }
            }


            form.ShowDialog();

            if (form.DialogResult == DialogResult.Cancel)
            {
                return Autodesk.Revit.UI.Result.Cancelled;
            }

            if (nombres_hojas == null)
            {
                TaskDialog.Show("Instruction!", "no views found");
                form.Close();
            }

            if (form.Equals(false))
            {
                return Autodesk.Revit.UI.Result.Cancelled;
            }

            int index_tBox = form.comboBox1.SelectedIndex;
            ElementId choose_Tblock = tittleblocks_list.ElementAtOrDefault(index_tBox);

            if (index_tBox == -1)
            {
                choose_Tblock = tittleblocks_list.ElementAtOrDefault(0);
            }

            List<ElementId> viewport_views = new List<ElementId>();
            List<Element> viewportEle = new List<Element>();
            List<Element> ScheduleEle = new List<Element>();

            foreach (string item in form.listBox2.Items)
            {
                string nam = item;
                foreach (Element i in hojitas)
                {
                    if (i.Name == nam)
                    {
                        Element a = i;
                        viewportEle.Add(a);
                        ElementId BC = i.Id;
                        viewport_views.Add(BC);
                        string name = i.Name;

                    }
                }
            }

            List<string> fromSelected = new List<string>();

            //--------------------------------------------------------------------------------------------------------------------
            List<XYZ> Bboxcenter = new List<XYZ>();
            List<XYZ> locationInsheet = new List<XYZ>();
            IEnumerable<int> sequencePoints = Enumerable.Range(0, 2);
            List<double> count = new List<double>();
            List<Viewport> countPorts = new List<Viewport>();


            List<ElementId> sobrasId = new List<ElementId>();
            List<ElementId> sobrasId2 = new List<ElementId>();
            List<ElementId> sobrasId3 = new List<ElementId>();
            List<Element> sobras = new List<Element>();
            List<double> numHoj = new List<double>();

            double uno1 = 1;
            numHoj.Add(uno1);
            int j = 0;

          
          

            List<ViewSheet> CreatedSheets = new List<ViewSheet>();
            Parameter p = null;

            if (viewportEle.ToArray().Length == 0)
            {
                TaskDialog.Show("Fail", "You must chose a Revit view to place in a sheet");
                return Autodesk.Revit.UI.Result.Cancelled;
            }

            //if (form.radioButton2.Checked == false && form.radioButton2.Checked == false)
            //{
            //    TaskDialog.Show("Fail", "You must select a type of arrangement for the Revit views");
            //    return Autodesk.Revit.UI.Result.Cancelled;
            //}


            if (form.radioButton1.Checked == true)
            {
                //--------------------------------CREATION OF SHEET AND VIEW TO POINT 0,0,0----------------------------------
                using (Transaction t = new Transaction(doc, "Create one RevitSheet"))
                {
                    t.Start();
                    double yy1 = 0.5;
                    double xx1 = 2.2;
                    foreach (int item in numHoj)
                    {
                        try
                        {
                            if (viewportEle.ToArray().Length == 0)
                            {
                                TaskDialog.Show("Fail", "You must chose a Revit view to place in a sheet");
                                return Autodesk.Revit.UI.Result.Cancelled;
                            }

                            start1: ViewSheet sheet2 = ViewSheet.Create(doc, choose_Tblock);
                            foreach (Element v in viewportEle)
                            {
                                if (yy1 >= 0.5)
                                {
                                    if (xx1 == 2.2 && yy1 == 2.0)
                                    {

                                        goto start1;
                                    }
                                    while (yy1 >= 2.5)
                                    {
                                        yy1 = 0.5;
                                        xx1 = 2.2;
                                    }

                                    if (Viewport.CanAddViewToSheet(doc, sheet2.Id, v.Id))
                                    {
                                        Viewport viewport = Viewport.Create(doc, sheet2.Id, v.Id, new XYZ(xx1, yy1, 0));
                                        //p = viewport.LookupParameter("Detail Number");
                                        //p.SetValueString("100");
                                    }
                                    else
                                    {
                                        TaskDialog.Show("Warning", "The view is already placed on a sheet");
                                       
                                    }
                                }
                                xx1 = xx1 - 0.5;
                                while (xx1 <= 0)
                                {
                                    yy1 = yy1 + 0.5;
                                    xx1 = 2.2;
                                }
                            }
                        }
                        catch (Exception)
                        {
                            TaskDialog.Show("Tittle block", "You must chose a title block");
                            throw;
                        }
                    }
                    doc.Regenerate();
                    t.Commit();
                    TaskDialog.Show("Done","One sheet was created");
                   
                }
            }

            

            if (form.radioButton2.Checked == true)
            {
                using (Transaction t = new Transaction(doc, "CreateSheetByPlan"))
                {
                    int numero = 0;
                    t.Start();
                    foreach (var view in hojitas)
                    {
                        if (viewportEle.ToArray().Length == 0)
                        {
                            TaskDialog.Show("Fail", "You must chose a Revit view to place in a sheet");
                            return Autodesk.Revit.UI.Result.Cancelled;
                        }

                        foreach (var VP in viewportEle)
                        {
                            if (VP.Id == view.Id)
                            {
                                ViewSheet sheet2 = ViewSheet.Create(doc, choose_Tblock);
                                numero++;
                                if (Viewport.CanAddViewToSheet(doc, sheet2.Id, view.Id))
                                {
                                    Viewport viewport = Viewport.Create(doc, sheet2.Id, view.Id, new XYZ(0, 0, 0));
                                }
                            }
                        }
                    }
                    t.Commit();
                    TaskDialog.Show("Done", numero + " sheets were created");
                    
                }
            }
           
            return Autodesk.Revit.UI.Result.Succeeded;
        }
    }

    [Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)]
    public class Duplicate_0ne_sheet : IExternalCommand
    {
        static AddInId appId = new AddInId(new Guid("F5FB1A7F-8410-6562-8420-04AE05C1239E"));
        public Autodesk.Revit.UI.Result Execute(ExternalCommandData commandData, ref string message, ElementSet elementSet)
        {
            UIDocument uidoc = commandData.Application.ActiveUIDocument;
            Autodesk.Revit.DB.Document doc = uidoc.Document;

            try
            {
                string filename = @"T:\Transfer\lopez\Book1.xlsx";
                using (ExcelPackage package = new ExcelPackage(new FileInfo(filename)))
                {
                    ExcelWorksheet sheet = package.Workbook.Worksheets.ElementAt(0);

                    int column = 9;
                    int number = Convert.ToInt32(sheet.Cells[2, column].Value);
                    sheet.Cells[2, column].Value = (number + 1); ;
                    package.Save();
                }
            }
            catch (Exception)
            {
                MessageBox.Show("!", "");
            }

            //string comments = "Duplicate_0ne_sheet" + "_" + doc.Application.Username + "_" + doc.Title;
            //string filename = @"D:\Users\lopez\Desktop\Comments.txt";
            ////System.Diagnostics.Process.Start(filename);
            //StreamWriter writer = new StreamWriter(filename, true);
            ////writer.WriteLine( Environment.NewLine);
            //writer.WriteLine(DateTime.Now + " - " + comments);
            //writer.Close();



            //---------------------------------------- FILTERS ------------------------------------
            IEnumerable<FamilySymbol> familyList = from elem in new FilteredElementCollector(doc).OfClass(typeof(FamilySymbol)).OfCategory(BuiltInCategory.OST_TitleBlocks)
                                                   let type = elem as FamilySymbol
                                                   select type;

            FilteredElementCollector collector = new FilteredElementCollector(doc);
            ICollection<Element> hojas = collector.OfClass(typeof(Autodesk.Revit.DB.View)).ToElements();

            FilteredElementCollector fec = new FilteredElementCollector(doc);
            fec.OfClass(typeof(ViewSection));

            FilteredElementCollector fec2 = new FilteredElementCollector(doc);
            fec2.OfClass(typeof(ViewSheet));



            var viewPlans = fec.Cast<ViewSection>().Where<ViewSection>(vp => vp.IsTemplate);
            IEnumerable<Autodesk.Revit.DB.View> viewList = from elem in new FilteredElementCollector(doc)
                                                 .OfClass(typeof(Autodesk.Revit.DB.View))
                                                 .OfCategory(BuiltInCategory.OST_Views)
                                                           let type = elem as Autodesk.Revit.DB.View
                                                           where type.Name != null
                                                           select type;

            IEnumerable<ViewSheet> viewSheet = from elem in new FilteredElementCollector(doc)
                                                 .OfClass(typeof(ViewSheet))
                                                 .OfCategory(BuiltInCategory.OST_Sheets)
                                               let type = elem as ViewSheet
                                               where type.Name != null
                                               select type;

            IEnumerable<TextNote> textnoteslist = from elem in new FilteredElementCollector(doc)
                                               .OfClass(typeof(TextNote)).OfCategory(BuiltInCategory.OST_TextNotes)
                                                  let type = elem as TextNote
                                                  where type.Name != null
                                                  select type;




            //TaskDialog.Show("point1", "point1");

            FamilySymbol FamilySymbol = null;
            IList<Element> m_alltitleblocks = new List<Element>();
            IList<Element> ElementsOnSheet = new List<Element>();
            List<ViewSheet> ViewSchedule_LIST = new List<ViewSheet>();


            List<Viewport> ViewPorts_ = new List<Viewport>();
            List<ElementId> Ids_ = new List<ElementId>();
            List<ElementId> vtype_ = new List<ElementId>();

            ICollection<ElementId> Ids2_ = new List<ElementId>();

            List<Autodesk.Revit.DB.View> Total_viewcount_onproject = new List<Autodesk.Revit.DB.View>();
            List<Autodesk.Revit.DB.View> view_to_copy = new List<Autodesk.Revit.DB.View>();
            List<ElementId> views_already_copied = new List<ElementId>();
            List<BoundingBoxXYZ> sectionBox = new List<BoundingBoxXYZ>();
            List<XYZ> vp_centers = new List<XYZ>();
            List<ViewSchedule> schedule_existing_insheet = new List<ViewSchedule>();

            Form18 form = new Form18();

            ViewFamilyType vft = new FilteredElementCollector(doc).OfClass(typeof(ViewFamilyType)).Cast<ViewFamilyType>().FirstOrDefault<ViewFamilyType>(x => ViewFamily.Section == x.ViewFamily);

            ViewSheet activeViewSheet = doc.ActiveView as ViewSheet;

            FilteredElementCollector col_sheetele = new FilteredElementCollector(doc, activeViewSheet.Id);
            var scheduleSheetInstances = col_sheetele.OfClass(typeof(ScheduleSheetInstance)).ToElements().OfType<ScheduleSheetInstance>();

            IList<ElementId> rev_Id = activeViewSheet.GetAllRevisionIds();



            List<Viewport> viewports = new List<Viewport>();
            List<ElementId> views = new List<ElementId>();


            List<XYZ> schudule_pt = new List<XYZ>();
            foreach (var scheduleSheetInstance in scheduleSheetInstances)
            {
                if (!scheduleSheetInstance.Name.Contains("Revision"))
                {
                    //pt = scheduleSheetInstance.Point;

                    var scheduleId = scheduleSheetInstance.ScheduleId;
                    if (scheduleId == ElementId.InvalidElementId)
                        continue;
                    var viewSchedule_ = doc.GetElement(scheduleId) as ViewSchedule;
                    schedule_existing_insheet.Add(viewSchedule_);
                }
            }
            foreach (var scheduleSheetInstance in scheduleSheetInstances)
            {
                if (!scheduleSheetInstance.Name.Contains("Revision"))
                {
                    schudule_pt.Add(scheduleSheetInstance.Point);
                }
            }

            List<ElementId> text_type = new List<ElementId>();
            List<XYZ> text_pt = new List<XYZ>();
            List<TextNote> textnotessearch = new List<TextNote>();
            foreach (var item in textnoteslist)
            {
                if (item.OwnerViewId == activeViewSheet.Id)
                {
                    text_type.Add(item.TextNoteType.Id);
                    text_pt.Add(item.Coord /*as XYZ*/);
                    textnotessearch.Add(item);
                }

            }

            try
            {
                ICollection<ElementId> views_ = activeViewSheet.GetAllPlacedViews();
                foreach (var item in views_)
                {
                    views.Add(item);
                }


                IList<Viewport> viewports__ = new FilteredElementCollector(doc).OfClass(typeof(Viewport)).Cast<Viewport>()
            .Where(q => q.SheetId == activeViewSheet.Id).ToList();
                foreach (var item in viewports__)
                {
                    viewports.Add(item);
                }
            }
            catch (Exception)
            {
                TaskDialog.Show("Warning", "Active View must be a sheet");
                throw;
            }



            //TaskDialog.Show("point2", "point2");


            foreach (var VID in views)
            {
                foreach (var VP in viewports)
                {
                    if (VP.ViewId == VID)
                    {
                        XYZ center = VP.GetBoxCenter();
                        vp_centers.Add(center);

                    }
                }
            }

            // foreach (var item in viewList)
            // {
            //     foreach (var item2 in new FilteredElementCollector(doc).OfClass(typeof(Autodesk.Revit.DB.View)).Cast<Autodesk.Revit.DB.View>()
            //.Where(q => q.Id == item.Id).ToList())
            //     {
            //         Total_viewcount_onproject.Add(item2);
            //     }
            // }

            //TaskDialog.Show("point3", "point3");

            foreach (var view_onproject in viewList)
            {
                foreach (var view_ID_oncurrentpage in views)
                {
                    if (view_onproject.Id == view_ID_oncurrentpage)
                    {
                        view_to_copy.Add(view_onproject);
                        ViewType vt = view_onproject.ViewType;



                        BoundingBoxXYZ room_box = view_onproject.get_BoundingBox(null);
                        sectionBox.Add(room_box);
                        vtype_.Add(view_onproject.ViewTemplateId);

                    }
                }
            }

            //TaskDialog.Show("point4", "point4");

            foreach (Element e in new FilteredElementCollector(doc).OwnedByView(/*sheet_.Id*/activeViewSheet.Id))
            {
                if (e.Category != null && e.Category.Name == "Viewports")
                {
                    ViewPorts_.Add(e as Viewport);
                }

                ElementsOnSheet.Add(e);
            }



            foreach (Element el in ElementsOnSheet)
            {
                foreach (FamilySymbol Fs in /*m_alltitleblocks*/ familyList)
                {
                    if (el.GetTypeId().IntegerValue == Fs.Id.IntegerValue)
                    {
                        FamilySymbol = Fs;
                    }
                }
            }

            form.ShowDialog();

            //TaskDialog.Show("point5", "point5");

            using (Transaction t = new Transaction(doc, "Create RevitSheet"))
            {
                t.Start();




                foreach (var item in view_to_copy)
                {

                    ElementId view_ = item.Duplicate(ViewDuplicateOption.WithDetailing);
                    views_already_copied.Add(view_);

                }

                //TaskDialog.Show("point6", "point6");

                IEnumerable<Autodesk.Revit.DB.View> viewList_new = from elem in new FilteredElementCollector(doc)
                                                 .OfClass(typeof(Autodesk.Revit.DB.View))
                                                 .OfCategory(BuiltInCategory.OST_Views)
                                                                   let type = elem as Autodesk.Revit.DB.View
                                                                   where type.Name != null
                                                                   select type;

                ViewSheet sheet2 = ViewSheet.Create(doc, FamilySymbol.Id);
                try
                {
                    sheet2.Name = activeViewSheet.Name + "copy";
                    sheet2.SheetNumber = activeViewSheet.SheetNumber + "copy";
                }
                catch (Exception)
                {

                    MessageBox.Show("Name or Number might be already in use!", "");
                }

                try
                {
                    for (int i = 0; i < schedule_existing_insheet.ToArray().Length; i++)
                    {
                        ScheduleSheetInstance.Create(doc, sheet2.Id, schedule_existing_insheet.ToArray()[i].Id, schudule_pt.ToArray()[i]);
                    }
                }
                catch (Exception)
                {

                    MessageBox.Show("The program can not copy schedule, try hidding them before copying the sheet", "");
                }

                /*ElementId defaultTypeId = doc.GetDefaultElementTypeId(ElementTypeGroup.TextNoteType);*/
                try
                {
                    for (int i = 0; i < textnotessearch.ToArray().Length; i++)
                    {
                        TextNote.Create(doc, sheet2.Id, text_pt.ToArray()[i], textnotessearch.ToArray()[i].Text, /*defaultTypeId*/ text_type.ToArray()[i]);
                    }
                }
                catch (Exception)
                {

                    MessageBox.Show("The program can not copy textnotes, try hidding them before copying the sheet", "");
                }




                if (form.radioButton1.Checked)
                {

                    sheet2.SetAdditionalRevisionIds(rev_Id);

                }



                if (sheet2.LookupParameter("View Organization") != null)
                {
                    string parametro = activeViewSheet.LookupParameter("View Organization").AsString();
                    Parameter param = sheet2.LookupParameter("View Organization");
                    param.Set(parametro);
                }

                if (sheet2.LookupParameter("Drawing Series") != null)
                {
                    string parametro = activeViewSheet.LookupParameter("Drawing Series").AsString();
                    Parameter param2 = sheet2.LookupParameter("Drawing Series");
                    param2.Set(parametro);
                }

                int centerpt = 0;
                foreach (var view_onproject in viewList_new)
                {
                    foreach (var view_ID_oncurrentpage in views_already_copied)
                    {
                        if (view_onproject.Id == view_ID_oncurrentpage)
                        {


                            if (Viewport.CanAddViewToSheet(doc, sheet2.Id, view_onproject.Id))
                            {
                                Viewport viewport = Viewport.Create(doc, sheet2.Id, view_onproject.Id, vp_centers.ToArray()[centerpt]);
                            }
                            centerpt++;
                        }
                    }
                }
                //TaskDialog.Show("point8", "point8");

                doc.Regenerate();
                t.Commit();
                uidoc.ActiveView = sheet2;
                TaskDialog.Show("Completed", "1 Sheet copied");
            }
            return Autodesk.Revit.UI.Result.Succeeded;
        }
    }

    [Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)]
    public class CreatMultipleSheet : IExternalCommand
    {
        static AddInId appId = new AddInId(new Guid("F5FB1A7F-8220-4862-8820-04AE05C1239E"));
        public Autodesk.Revit.UI.Result Execute(ExternalCommandData commandData, ref string message, ElementSet elementSet)
        {
            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = commandData.Application.ActiveUIDocument;
            Autodesk.Revit.DB.Document doc = uidoc.Document;
            Autodesk.Revit.ApplicationServices.Application app = uiapp.Application;

            //try
            //{
            //    string filename = @"T:\Transfer\lopez\Book1.xlsx";
            //    using (ExcelPackage package = new ExcelPackage(new FileInfo(filename)))
            //    {
            //        ExcelWorksheet sheet = package.Workbook.Worksheets.ElementAt(0);

            //        int column = 7;
            //        int number = Convert.ToInt32(sheet.Cells[2, column].Value);
            //        sheet.Cells[2, column].Value = (number + 1); ;
            //        package.Save();
            //    }
            //}
            //catch (Exception)
            //{
            //    MessageBox.Show("Excel file not found", "");
            //}

            //---------------------------------------- FILTERS ------------------------------------
            IEnumerable<FamilySymbol> TittleBlock_List = from elem in new FilteredElementCollector(doc).OfClass(typeof(FamilySymbol)).OfCategory(BuiltInCategory.OST_TitleBlocks)
                                                         let type = elem as FamilySymbol select type;

            List<Element> viewElems = new List<Element>();
            FilteredElementCollector collector = new FilteredElementCollector(doc);
            viewElems.AddRange(collector.OfClass(typeof(Autodesk.Revit.DB.View)).ToElements());

            List<ElementId> tittleblocks_list = new List<ElementId>();
            List<String> tittle_block_list_name = new List<String>();
            List<string> gourpheader_list = new List<string>();

            int distance = 0;
            Form9 form = new Form9();

            foreach (FamilySymbol T in TittleBlock_List) 
            {
                ElementId HOJAS_ID = T.Id;
                tittleblocks_list.Add(HOJAS_ID);

                string nombre_hojas_each = T.Name;
                tittle_block_list_name.Add(T.FamilyName + " - " + nombre_hojas_each);
            }

            foreach (var item in tittle_block_list_name)
            {
                form.comboBox1.Items.Add(item);
            }

            List<Element> store_Selected = new List<Element>(); //nombre hojas_copy

            foreach (var item in tittle_block_list_name)
            {
                form.comboBox1.Items.Add(item);
            }

            int numero_hojas = 0;
            form.ShowDialog();

            if (form.DialogResult == DialogResult.Cancel)
            {
                return Autodesk.Revit.UI.Result.Cancelled;
            }

            if (form.DialogResult == System.Windows.Forms.DialogResult.Cancel)
            {
                return Autodesk.Revit.UI.Result.Cancelled;
            }

            if (form.Equals(false))
            {
                return Autodesk.Revit.UI.Result.Cancelled;
            }

            int index_tBox = form.comboBox1.SelectedIndex;
            ElementId choose_Tblock = tittleblocks_list.ElementAtOrDefault(index_tBox);

            numero_hojas = (int)form.numericUpDown1.Value;

            using (Transaction t = new Transaction(doc, "Create RevitSheet"))
            {
                t.Start();

                {
                    for (int i = 0; i < numero_hojas; i++)
                    {

                        ViewSheet sheet2 = ViewSheet.Create(doc, choose_Tblock);
                        try
                        {
                            //sheet2.SheetNumber = form.comboBox2.SelectedItem.ToString() + (distance + i);
                            sheet2.Name = "New sheet " + i.ToString();
                        }
                        catch
                        {
                            TaskDialog.Show("warning", "name Exists");
                        }
                    }
                }
                doc.Regenerate();
                t.Commit();

                return Autodesk.Revit.UI.Result.Succeeded;
            }
        }
    }

    [Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)]
    public class copy_schedule : IExternalCommand
    {
        static AddInId appId = new AddInId(new Guid("6C22CC72-A167-4819-AAF1-A178F6B44BAB"));
        public Autodesk.Revit.UI.Result Execute(ExternalCommandData commandData, ref string message, ElementSet elementSet)
        {
            UIDocument uidoc = commandData.Application.ActiveUIDocument;
            Autodesk.Revit.DB.Document doc = uidoc.Document;

            //try
            //{
            //    string filename = @"T:\Transfer\lopez\Book1.xlsx";
            //    using (ExcelPackage package = new ExcelPackage(new FileInfo(filename)))
            //    {
            //        ExcelWorksheet sheet = package.Workbook.Worksheets.ElementAt(0);

            //        int column = 20;
            //        int number = Convert.ToInt32(sheet.Cells[2, column].Value);
            //        sheet.Cells[2, column].Value = (number + 1); ;
            //        package.Save();
            //    }
            //}
            //catch (Exception)
            //{
            //    MessageBox.Show("Excel file not found", "");
            //}

            FilteredElementCollector collector = new FilteredElementCollector(doc).OfClass(typeof(ViewSchedule));

            IEnumerable<ViewSchedule> views = from Autodesk.Revit.DB.ViewSchedule f in collector
                                              where (f.ViewType == ViewType.Schedule /*&& f.ViewName.Equals("Unit Default")*/)
                                              select f;

            IEnumerable<ViewSheet> viewSheet = from elem in new FilteredElementCollector(doc).OfClass(typeof(ViewSheet)).OfCategory(BuiltInCategory.OST_Sheets)
                                               let type = elem as ViewSheet
                                               where type.Name != null
                                               select type;

            IEnumerable<Autodesk.Revit.DB.View> legends = from elem in new FilteredElementCollector(doc).OfClass(typeof(Autodesk.Revit.DB.View)).OfCategory(BuiltInCategory.OST_Views)
                                                          let type = elem as Autodesk.Revit.DB.View
                                                          where type.ViewType == ViewType.Legend
                                                          select type;

            Form19 form = new Form19();
            ViewSheet activeViewSheet = null;

            if (doc.ActiveView.ViewType == ViewType.DrawingSheet)
            {
                activeViewSheet = doc.ActiveView as ViewSheet;
            }
            else
            {
                TaskDialog.Show("warning", "The active view must be a view sheet");
                return Autodesk.Revit.UI.Result.Cancelled;
            }
            List<ViewSchedule> schedule_existing_insheet = new List<ViewSchedule>();
            List<ImageType> img_existing_insheet = new List<ImageType>();
            List<Autodesk.Revit.DB.View> Legend_list = new List<Autodesk.Revit.DB.View>();

            FilteredElementCollector col_sheetele = new FilteredElementCollector(doc, activeViewSheet.Id);
            var scheduleSheetInstances = col_sheetele.OfClass(typeof(ScheduleSheetInstance)).ToElements().OfType<ScheduleSheetInstance>();
            var ImageType_ = col_sheetele.OfClass(typeof(ImageType)).ToElements().OfType<ImageType>();
            XYZ schudule_pt = null;
            XYZ legend_pt = null;

            List<Viewport> viewports = new List<Viewport>();
            List<ElementId> views_list = new List<ElementId>();

            try
            {
                ICollection<ElementId> views_ = activeViewSheet.GetAllPlacedViews();
                foreach (var item in views_)
                {
                    views_list.Add(item);
                }


                IList<Viewport> viewports__ = new FilteredElementCollector(doc).OfClass(typeof(Viewport)).Cast<Viewport>()
            .Where(q => q.SheetId == activeViewSheet.Id).ToList();
                foreach (var item in viewports__)
                {

                    viewports.Add(item);
                }
            }
            catch (Exception)
            {
                TaskDialog.Show("Warning", "Active View must be a sheet");
                throw;
            }

            foreach (var VP in viewports)
            {
                foreach (var item in legends)
                {
                    if (VP.ViewId == item.Id)
                    {
                        form.listBox1.Items.Add(item.Name);
                    }
                }
            }

            foreach (var scheduleSheetInstance in scheduleSheetInstances)
            {
                if (!scheduleSheetInstance.Name.Contains("Revision"))
                {
                    var scheduleId = scheduleSheetInstance.ScheduleId;
                    if (scheduleId == ElementId.InvalidElementId)
                        continue;
                    var viewSchedule_ = doc.GetElement(scheduleId) as ViewSchedule;
                    schedule_existing_insheet.Add(viewSchedule_);
                }
            }

            foreach (var item in schedule_existing_insheet)
            {
                if (!item.Name.Contains("Revision"))
                {
                    form.listBox1.Items.Add(item.Name);
                }
            }

            foreach (var item in viewSheet)
            {
                form.listBox2.Items.Add(item.Name);
            }

            form.ShowDialog();



            Autodesk.Revit.DB.View vID = null;
            foreach (var item in legends)
            {
                if (item.Name == form.listBox1.SelectedItem.ToString())
                {
                    vID = item;
                }
            }



            if (vID != null)
            {
                foreach (var VP in viewports)
                {

                    if (VP.ViewId == vID.Id)
                    {
                        legend_pt = VP.GetBoxCenter();
                    }
                }
            }

            foreach (var scheduleSheetInstance in scheduleSheetInstances)
            {
                if (!scheduleSheetInstance.Name.Contains("Revision"))
                {
                    if (form.listBox1.SelectedItem.ToString() == scheduleSheetInstance.Name)
                    {
                        schudule_pt = scheduleSheetInstance.Point;
                    }
                }
            }

            ViewSchedule schedule_selected = null;
            foreach (string item in form.listBox1.SelectedItems)
            {
                string nam = item;
                foreach (Element i in schedule_existing_insheet)
                {
                    if (i.Name == nam)
                    {

                        schedule_selected = i as ViewSchedule;
                    }
                }
            }

            List<ViewSheet> viewsheet_selected = new List<ViewSheet>();
            foreach (string item in form.listBox2.SelectedItems)
            {

                string nam = item;
                foreach (ViewSheet i in viewSheet)
                {
                    if (i.Name == nam)
                    {
                        if (!viewsheet_selected.Contains(i as ViewSheet))
                        {
                            viewsheet_selected.Add(i as ViewSheet);
                        }
                    }
                }
            }

            using (Transaction trans = new Transaction(doc, "ViewDuplicate"))
            {
                trans.Start();

                if (schedule_selected != null)
                {
                    for (int i = 0; i < viewsheet_selected.ToArray().Length; i++)
                    {
                        try
                        {
                            ScheduleSheetInstance.Create(doc, viewsheet_selected.ToArray()[i].Id, schedule_selected.Id, schudule_pt);
                        }
                        catch (Exception)
                        {

                            TaskDialog.Show("warning", "The selected sheet template does not contain a schedule");
                        }
                    }
                    trans.Commit();
                }

                if (vID != null)
                {
                    for (int i = 0; i < viewsheet_selected.ToArray().Length; i++)
                    {
                        try
                        {
                            Viewport.Create(doc, viewsheet_selected.ToArray()[i].Id, vID.Id, legend_pt);
                        }
                        catch (Exception)
                        {

                            TaskDialog.Show("warning", "The selected sheet template does not contain a legend");
                        }
                    }
                    trans.Commit();
                }

            }
            return Autodesk.Revit.UI.Result.Succeeded;
        }
    }

    [Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)]
    public class DeleteAllViews : IExternalCommand
    {
        static AddInId appId = new AddInId(new Guid("5F22CC78-A557-4819-AAF1-A678F6B22BAB"));
        public Autodesk.Revit.UI.Result Execute(ExternalCommandData commandData, ref string message, ElementSet elementSet)
        {
            //try
            //{
            //    string filename = @"T:\Transfer\lopez\Book1.xlsx";
            //    using (ExcelPackage package = new ExcelPackage(new FileInfo(filename)))
            //    {
            //        ExcelWorksheet sheet = package.Workbook.Worksheets.ElementAt(0);

            //        int column = 12;
            //        int number = Convert.ToInt32(sheet.Cells[2, column].Value);
            //        sheet.Cells[2, column].Value = (number + 1); ;
            //        package.Save();
            //    }
            //}
            //catch (Exception)
            //{
            //    MessageBox.Show("Excel file not found", "");
            //}

            Form15 form = new Form15();
            form.ShowDialog();

            if (form.DialogResult == DialogResult.Cancel)
            {
                return Autodesk.Revit.UI.Result.Cancelled;
            }

            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Autodesk.Revit.DB.Document doc = uidoc.Document;
            Autodesk.Revit.ApplicationServices.Application app = uiapp.Application;

            try
            {
                Autodesk.Revit.DB.View viewTemplate = (from v in new FilteredElementCollector(doc).OfClass(typeof(Autodesk.Revit.DB.View)).Cast<Autodesk.Revit.DB.View>()
                                                       where !v.IsTemplate && v.Name == "Home"
                                                       select v).First();
                uidoc.ActiveView = viewTemplate;
            }
            catch (Exception)
            {
                TaskDialog.Show("Warning", "'Home' view was not found in this project, name a view Home and try again");
                return Autodesk.Revit.UI.Result.Cancelled;
            }

            FilteredElementCollector collector = new FilteredElementCollector(doc);
            ICollection<Element> collection = collector.OfClass(typeof(Autodesk.Revit.DB.View)).ToElements();

            using (Transaction t = new Transaction(doc, "Delete All Sheets and Views"))
            {
                t.Start();
                int x = 0;
                foreach (Element e in collection)
                {
                    try
                    {
                        Autodesk.Revit.DB.View view = e as Autodesk.Revit.DB.View;
                        doc.Delete(e.Id);
                        x += 1;
                    }
                    catch (Exception ex)
                    {
                    }
                }
                t.Commit();
                //doc.Regenerate();

                TaskDialog.Show("DeleteAllSheetsViews", "Views & Sheets Deleted: " + x.ToString());
            }
            return Autodesk.Revit.UI.Result.Succeeded;
        }
    }

    [Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)]
    public class ReNumbering : IExternalCommand
    {

        private Parameter getParameterForReference(Autodesk.Revit.DB.Document doc, Reference r)
        {
            Element e = doc.GetElement(r);
            Parameter p = null;
            if (e is Grid)
                p = e.LookupParameter("Name");
            else if (e is Room)
                p = e.LookupParameter("Number");
            else if (e is FamilyInstance)
                p = e.LookupParameter("Mark");
            else if (e is Viewport) // Viewport class is new to Revit 2013 API
                p = e.LookupParameter("Detail Number");
            else
            {
                TaskDialog.Show("Error", "Unsupported element");
                return null;
            }
            return p;
        }
        private void setParameterToValue(Parameter p, int i)
        {
            if (p.StorageType == StorageType.Integer)
                p.Set(i);
            else if (p.StorageType == StorageType.String)
                p.Set(i.ToString());
        }

        static AddInId appId = new AddInId(new Guid("F6FB1A7F-8410-6562-8420-06AE05C1239E"));
        public Autodesk.Revit.UI.Result Execute(ExternalCommandData commandData, ref string message, ElementSet elementSet)
        {
            UIDocument uidoc = commandData.Application.ActiveUIDocument;
            Autodesk.Revit.DB.Document doc = uidoc.Document;

            //try
            //{
            //    string filename = @"T:\Transfer\lopez\Book1.xlsx";
            //    using (ExcelPackage package = new ExcelPackage(new FileInfo(filename)))
            //    {
            //        ExcelWorksheet sheet = package.Workbook.Worksheets.ElementAt(0);

            //        int column = 10;
            //        int number = Convert.ToInt32(sheet.Cells[2, column].Value);
            //        sheet.Cells[2, column].Value = (number + 1); ;
            //        package.Save();
            //    }
            //}
            //catch (Exception)
            //{
            //    MessageBox.Show("Excel file not found", "");
            //}

            IList<Reference> refList = new List<Reference>();

            TaskDialog.Show("Grids - family instances - Viewports", "Select elements in order to be renumbered and then press ESCAPE to finish (the sequence will start with the number of the first selected element)");

            try
            {
                while (true)
                    refList.Add(uidoc.Selection.PickObject(ObjectType.Element, "Select elements in order to be renumbered. ESC when finished."));
            }
            catch
            { }

            using (Transaction t = new Transaction(doc, "Renumber"))
            {
                t.Start();
                // need to avoid encountering the error "The name entered is already in use. Enter a unique name."
                // for example, if there is already a grid 2 we can't renumber some other grid to 2
                // therefore, first renumber everny element to a temporary name then to the real one
                int ctr = 1;
                int startValue = 0;
                foreach (Reference r in refList)
                {
                    Parameter p = getParameterForReference(doc, r);



                    // get the value of the first element to use as the start value for the renumbering in the next loop
                    if (ctr == 1)
                        startValue = Convert.ToInt16(p.AsString());

                    setParameterToValue(p, ctr + 12345); // hope this # is unused (could use Failure API to make this more robust
                    ctr++;
                }

                ctr = startValue;
                foreach (Reference r in refList)
                {
                    Parameter p = getParameterForReference(doc, r);
                    setParameterToValue(p, ctr);
                    ctr++;
                }
                t.Commit();
            }
            return Autodesk.Revit.UI.Result.Succeeded;
        }
    }

    [Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)]
    public class Wall_Angle_to : IExternalCommand
    {
        static AddInId appId = new AddInId(new Guid("5F46AA78-A136-6509-AAF8-A478F3B24BAB"));
        public Autodesk.Revit.UI.Result Execute(ExternalCommandData commandData, ref string message, ElementSet elementSet)
        {
            UIDocument uidoc = commandData.Application.ActiveUIDocument;
            Autodesk.Revit.DB.Document doc = uidoc.Document;

            List<Element> ele = new List<Element>();
            List<Element> ele2 = new List<Element>();
            IList<Reference> refList = new List<Reference>();

            TaskDialog.Show("!", "Select a reference Grid to find orthogonal walls");

            Grid levelBelow = doc.GetElement(uidoc.Selection.PickObject(ObjectType.Element, "Select Grid")) as Grid;

            Autodesk.Revit.DB.Curve dircurve = levelBelow.Curve;
            Autodesk.Revit.DB.Line line = dircurve as Autodesk.Revit.DB.Line;
            XYZ dir = line.Direction;

            XYZ origin = line.Origin;
            XYZ viewdir = line.Direction;
            XYZ up = XYZ.BasisZ;
            XYZ right = up.CrossProduct(viewdir);

            foreach (Element wall in new FilteredElementCollector(doc).OfClass(typeof(Wall)))
            {
                LocationCurve lc = wall.Location as LocationCurve;
                Autodesk.Revit.DB.Transform curveTransform = lc.Curve.ComputeDerivatives(0.5, true);

                try
                {
                    XYZ origin2 = curveTransform.Origin;
                    XYZ viewdir2 = curveTransform.BasisX.Normalize();
                    XYZ viewdir2_back = curveTransform.BasisX.Normalize() * -1;

                    XYZ up2 = XYZ.BasisZ;
                    XYZ right2 = up.CrossProduct(viewdir2);
                    XYZ left2 = up.CrossProduct(viewdir2 * -1);

                    double y_onverted = Math.Round(-1 * viewdir2.X);

                    if (viewdir.IsAlmostEqualTo(right2/*, 0.3333333333*/))
                    {
                        ele.Add(wall);
                    }
                    if (viewdir.IsAlmostEqualTo(left2))
                    {
                        ele.Add(wall);
                    }
                    if (viewdir.IsAlmostEqualTo(viewdir2))
                    {
                        ele.Add(wall);
                    }
                    if (viewdir.IsAlmostEqualTo(viewdir2_back))
                    {
                        ele.Add(wall);
                    }


                }
                catch (Exception)
                {
                    return Autodesk.Revit.UI.Result.Cancelled;
                }
            }
            uidoc.Selection.SetElementIds(ele.Select(q => q.Id).ToList());
            return Autodesk.Revit.UI.Result.Succeeded;
        }
    }

    [Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)]
    public class text_upper : IExternalCommand
    {
        static AddInId appId = new AddInId(new Guid("5F88CC78-A137-6509-AAF8-A478F3B24BAB"));
        public Autodesk.Revit.UI.Result Execute(ExternalCommandData commandData, ref string message, ElementSet elementSet)
        {

            UIDocument uidoc = commandData.Application.ActiveUIDocument;
            Autodesk.Revit.DB.Document doc = uidoc.Document;



            FilteredElementCollector rooms1 = new FilteredElementCollector(doc).WhereElementIsNotElementType().OfClass(typeof(SpatialElement)).OfCategory(BuiltInCategory.OST_Rooms)
                                             ;
            ICollection<Element> room2 = rooms1.ToElements();

            IEnumerable<TextNote> textnoteslist = from elem in new FilteredElementCollector(doc)
                                               .OfClass(typeof(TextNote)).OfCategory(BuiltInCategory.OST_TextNotes)
                                                  let type = elem as TextNote
                                                  where type.Name != null
                                                  select type;

            IEnumerable<Autodesk.Revit.DB.View> view_list = from elem in new FilteredElementCollector(doc)
                                                .OfClass(typeof(Autodesk.Revit.DB.View))
                                                .OfCategory(BuiltInCategory.OST_Views)
                                                            let type = elem as Autodesk.Revit.DB.View
                                                            where type.Name != null
                                                            select type;

            IEnumerable<Autodesk.Revit.DB.Grid> Grid_list = from elem in new FilteredElementCollector(doc)
                                                .OfClass(typeof(Autodesk.Revit.DB.Grid))
                                                .OfCategory(BuiltInCategory.OST_Grids)
                                                            let type = elem as Autodesk.Revit.DB.Grid
                                                            where type.Name != null
                                                            select type;

            IEnumerable<FamilySymbol> familyList = from elem in new FilteredElementCollector(doc).OfClass(typeof(FamilySymbol)).OfCategory(BuiltInCategory.OST_TitleBlocks)
                                                   let type = elem as FamilySymbol
                                                   select type;

            FamilyInstance fi = new FilteredElementCollector(doc, doc.ActiveView.Id).OfClass(typeof(FamilyInstance)).OfCategory(BuiltInCategory.OST_TitleBlocks)
      .FirstOrDefault() as FamilyInstance;


            Form25 form = new Form25();
            form.ShowDialog();



            List<TextNote> textnotessearch = new List<TextNote>();
            foreach (var item in textnoteslist)
            {

                textnotessearch.Add(item);
            }

            //if (doc.ActiveView.ViewType == ViewType. ViewSheet)
            //{

            //}

            ViewType type_ = doc.ActiveView.ViewType;

            ViewSheet activeViewSheet = doc.ActiveView as ViewSheet;

            Autodesk.Revit.DB.View activeViewSheet_ = doc.ActiveView as Autodesk.Revit.DB.View;

            List<Viewport> viewports = new List<Viewport>();
            List<ElementId> views = new List<ElementId>();
            List<Grid> g_list = new List<Grid>();
            List<Grid> tblock_list = new List<Grid>();
            FamilySymbol FamilySymbol = null;
            IList<Element> ElementsOnSheet = new List<Element>();


            ICollection<ElementId> views_;

            try
            {
                views_ = activeViewSheet.GetAllPlacedViews();
            }
            catch (Exception)
            {
                MessageBox.Show("Warning", "You must be on a Revit sheet before running this tool");
                return Autodesk.Revit.UI.Result.Cancelled;
            }


            foreach (var item in views_)
            {

                views.Add(item);
            }
            foreach (Grid item in Grid_list)
            {
                g_list.Add(item);
            }


            IList<Viewport> viewports__ = new FilteredElementCollector(doc).OfClass(typeof(Viewport)).Cast<Viewport>()
        .Where(q => q.SheetId == activeViewSheet.Id).ToList();
            foreach (var item in viewports__)
            {

                viewports.Add(item);
            }

            foreach (Element e in new FilteredElementCollector(doc).OwnedByView(/*sheet_.Id*/activeViewSheet.Id))
            {
                //if (e.Category != null && e.Category.Name == "Viewports")
                //{
                //    ViewPorts_.Add(e as Viewport);
                //}

                ElementsOnSheet.Add(e);
            }

            using (Transaction trans = new Transaction(doc, "Capital letters"))
            {
                trans.Start();

                if (form.checkBox1.Checked)
                {
                    foreach (var text_ in textnotessearch)
                    {
                        if (text_.OwnerViewId == activeViewSheet.Id)
                        {
                            string upper = text_.Text.ToUpper();
                            text_.Text = upper;
                        }

                    }
                }

                if (form.checkBox1.Checked)
                {
                    foreach (var text_ in textnotessearch)
                    {
                        foreach (var view_id in views)
                        {
                            if (text_.OwnerViewId == view_id)
                            {
                                string upper = text_.Text.ToUpper();
                                text_.Text = upper;
                            }
                        }
                    }
                }

                if (form.checkBox2.Checked)
                {
                    foreach (var item in room2)
                    {
                        string upper = item.Name.ToUpper();
                        item.Name = upper;
                    }
                }

                if (form.checkBox3.Checked)
                {
                    foreach (var view in view_list)
                    {
                        foreach (var viewid in views)
                        {
                            if (viewid == view.Id)
                            {
                                string upper = view.Name.ToUpper();
                                view.Name = upper;
                            }
                        }
                    }
                }

                if (form.checkBox4.Checked)
                {
                    foreach (Grid item in g_list)
                    {
                        string upper = item.Name.ToUpper();
                        item.Name = upper;
                    }
                }

                if (form.checkBox6.Checked)
                {
                    string upper2 = fi.LookupParameter("Sheet Name").AsString().ToUpper();
                    Parameter param = fi.LookupParameter("Sheet Name") /*fi.get_Parameter("Sheet Name")*/;
                    param.Set(upper2);
                }

                //foreach (var item in viewports__)
                //{
                //    string upper = item.LookupParameter("View Name").ToString().ToUpper();

                //    p = item.LookupParameter("Title on Sheet");

                //    p.SetValueString("lfdkgjsdf");


                //    viewports.Add(item);
                //}

                trans.Commit();
            }


            //string appdataFolder = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
            //string folderPath = Path.Combine(appdataFolder, @"Autodesk\Revit\Addins\2019\AJC_Commands\img\solar analisys guide.pdf");
            //Process.Start($"{folderPath}");

            return Autodesk.Revit.UI.Result.Succeeded;
        }
    }

    [Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)]
    public class TotalLenght : IExternalCommand
    {
        static AddInId appId = new AddInId(new Guid("5F92CC78-A127-4809-AAF8-A478F3B24BAB"));
        public Autodesk.Revit.UI.Result Execute(ExternalCommandData commandData, ref string message, ElementSet elementSet)
        {
            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Autodesk.Revit.DB.Document doc = uidoc.Document;

            double length = 0;

            ICollection<ElementId> ids = uidoc.Selection.GetElementIds();

            //if (true)
            //{

            //}
            //TaskDialog.Show("Error", ids.ToList().Count ==);
            //return Autodesk.Revit.UI.Result.Cancelled;

            foreach (ElementId id in ids)
            {
                Element e = doc.GetElement(id);
                Parameter lengthParam = e.get_Parameter(BuiltInParameter.CURVE_ELEM_LENGTH);
                if (lengthParam == null)
                    continue;
                length += lengthParam.AsDouble();
            }
            string lengthWithUnits = UnitFormatUtils.Format(doc.GetUnits(), UnitType.UT_Length, length, false, false);
            TaskDialog.Show("Length", ids.Count + " elements = " + lengthWithUnits);

            return Autodesk.Revit.UI.Result.Succeeded;
        }
    }

    [Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)]
    public class isolate : IExternalCommand
    {
        static AddInId appId = new AddInId(new Guid("5F88CC78-A137-4809-AAF8-A478F3B24BAB"));
        public Autodesk.Revit.UI.Result Execute(ExternalCommandData commandData, ref string message, ElementSet elementSet)
        {

            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Autodesk.Revit.ApplicationServices.Application app = uiapp.Application;
            Autodesk.Revit.DB.Document doc = uidoc.Document;

            ElementId elementIdToIsolate = null;
            try
            {
                elementIdToIsolate = uidoc.Selection.PickObject(ObjectType.Element, "Select element or ESC to reset the view").ElementId;
            }
            catch
            {

            }

            OverrideGraphicSettings ogsFade = new OverrideGraphicSettings();
            ogsFade.SetSurfaceTransparency(80);
            ogsFade.SetSurfaceForegroundPatternVisible(false);
            ogsFade.SetSurfaceBackgroundPatternVisible(false);
            ogsFade.SetHalftone(true);

            OverrideGraphicSettings ogsIsolate = new OverrideGraphicSettings();
            ogsIsolate.SetSurfaceTransparency(0);
            ogsIsolate.SetSurfaceForegroundPatternVisible(true);
            ogsIsolate.SetSurfaceBackgroundPatternVisible(true);
            ogsIsolate.SetHalftone(false);

            using (Transaction t = new Transaction(doc, "Isolate with Fade"))
            {
                t.Start();
                foreach (Element e in new FilteredElementCollector(doc, doc.ActiveView.Id).WhereElementIsNotElementType())
                {
                    if (e.Id == elementIdToIsolate || elementIdToIsolate == null)
                        doc.ActiveView.SetElementOverrides(e.Id, ogsIsolate);
                    else
                        doc.ActiveView.SetElementOverrides(e.Id, ogsFade);
                }
                t.Commit();
            }

            return Autodesk.Revit.UI.Result.Succeeded;
        }
    }

    [Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)]
    public class isolate_category : IExternalCommand
    {
        static AddInId appId = new AddInId(new Guid("5F88CC78-A137-4809-AAF8-A478F3B24BAB"));
        public Autodesk.Revit.UI.Result Execute(ExternalCommandData commandData, ref string message, ElementSet elementSet)
        {

            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Autodesk.Revit.ApplicationServices.Application app = uiapp.Application;
            Autodesk.Revit.DB.Document doc = uidoc.Document;

            Reference myRef2 = uidoc.Selection.PickObject(ObjectType.Element);
            Element e2 = doc.GetElement(myRef2.ElementId);

            GeometryObject geomObj2 = e2.GetGeometryObjectFromReference(myRef2);
            Wall wall_ = e2 as Wall;

            Category cat = e2.Category;
            var catname = cat.Name;

            OverrideGraphicSettings ogsFade = new OverrideGraphicSettings();
            ogsFade.SetSurfaceTransparency(80);
            ogsFade.SetSurfaceForegroundPatternVisible(false);
            ogsFade.SetSurfaceBackgroundPatternVisible(false);
            ogsFade.SetHalftone(true);

            OverrideGraphicSettings ogsIsolate = new OverrideGraphicSettings();
            ogsIsolate.SetSurfaceTransparency(0);
            ogsIsolate.SetSurfaceForegroundPatternVisible(true);
            ogsIsolate.SetSurfaceBackgroundPatternVisible(true);
            ogsIsolate.SetHalftone(false);

            using (Transaction t = new Transaction(doc, "Isolate with Fade"))
            {
                t.Start();
                foreach (Element e in new FilteredElementCollector(doc, doc.ActiveView.Id).WhereElementIsNotElementType())
                {
                    if (e.Category != null)
                    {
                        if (e.Category.Name == cat.Name)
                            doc.ActiveView.SetElementOverrides(e.Id, ogsIsolate);
                        else
                            doc.ActiveView.SetElementOverrides(e.Id, ogsFade);
                    }

                }
                t.Commit();
            }

            return Autodesk.Revit.UI.Result.Succeeded;
        }
    }

    [Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)]
    public class Clean_view : IExternalCommand
    {
        static AddInId appId = new AddInId(new Guid("5F88CC78-A137-4809-AAF8-A478F3B24BAB"));
        public Autodesk.Revit.UI.Result Execute(ExternalCommandData commandData, ref string message, ElementSet elementSet)
        {

            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Autodesk.Revit.ApplicationServices.Application app = uiapp.Application;
            Autodesk.Revit.DB.Document doc = uidoc.Document;



            OverrideGraphicSettings ogsFade = new OverrideGraphicSettings();
            ogsFade.SetSurfaceTransparency(85);
            ogsFade.SetSurfaceForegroundPatternVisible(false);
            ogsFade.SetSurfaceBackgroundPatternVisible(false);
            ogsFade.SetHalftone(true);

            OverrideGraphicSettings ogsIsolate = new OverrideGraphicSettings();
            ogsIsolate.SetSurfaceTransparency(0);
            ogsIsolate.SetSurfaceForegroundPatternVisible(true);
            ogsIsolate.SetSurfaceBackgroundPatternVisible(true);
            ogsIsolate.SetHalftone(false);

            using (Transaction t = new Transaction(doc, "Isolate with Fade"))
            {
                t.Start();
                foreach (Element e in new FilteredElementCollector(doc, doc.ActiveView.Id).WhereElementIsNotElementType())
                {
                    doc.ActiveView.SetElementOverrides(e.Id, ogsIsolate);
                }
                t.Commit();
            }
            return Autodesk.Revit.UI.Result.Succeeded;
        }
    }

    [Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)]
    public class View_range_by_bbox : IExternalCommand
    {
        static AddInId appId = new AddInId(new Guid("5F88CC78-A137-4809-AAF8-A478F3B24BAB"));
        public Autodesk.Revit.UI.Result Execute(ExternalCommandData commandData, ref string message, ElementSet elementSet)
        {
            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Autodesk.Revit.ApplicationServices.Application app = uiapp.Application;
            Autodesk.Revit.DB.Document doc = uidoc.Document;

            ViewPlan viewPlan = doc.ActiveView as ViewPlan;

            if (viewPlan == null)
            {
                TaskDialog.Show("Error", "Active view must be a plan view.");
                return Autodesk.Revit.UI.Result.Cancelled;
            }

            Level level = viewPlan.GenLevel;

            View3D view3d = null;
            try
            {
                view3d = (from v in new FilteredElementCollector(doc).OfClass(typeof(View3D)).Cast<View3D>() where v.Name == viewPlan.Name select v).First();
            }
            catch
            {
                TaskDialog.Show("Error", "There is no 3D view named '" + viewPlan.Name + "'");
                return Autodesk.Revit.UI.Result.Cancelled;
            }

            BoundingBoxXYZ bbox = view3d.get_BoundingBox(null);

            // The coordinates of the bounding box are defined relative to a coordinate system specific to the bounding box
            // When setting the view range offsets, the values will need to be relative to the model 

            // This transform translates from coordinate system of the bounding box to the model coordinate system
            Autodesk.Revit.DB.Transform transform = bbox.Transform;
            // Transform.Origin defines the origin of the bounding box's coordinate system in the model coordinate system
            // The Z value indicates the vertical offset of the bounding box's coordinate system
            double bboxOriginZ = transform.Origin.Z;

            // BoundingBoxXYZ.Min.Z and BoundingBoxXYZ.Max.Z give the Z values for the bottom and top of the section box
            // Adding the Transform.Origin.Z converts the value to the model coordinate system
            double minZ = bbox.Min.Z + bboxOriginZ;
            double maxZ = bbox.Max.Z + bboxOriginZ;

            // Get the PlanViewRange object from the plan view
            PlanViewRange viewRange = viewPlan.GetViewRange();

            // Set all planes of the view range to use the plan view's level
            viewRange.SetLevelId(PlanViewPlane.TopClipPlane, level.Id);
            viewRange.SetLevelId(PlanViewPlane.CutPlane, level.Id);
            viewRange.SetLevelId(PlanViewPlane.BottomClipPlane, level.Id);
            viewRange.SetLevelId(PlanViewPlane.ViewDepthPlane, level.Id);

            // Set the view depth offset to the difference between the bottom of the section box and
            // the elevation of the level
            viewRange.SetOffset(PlanViewPlane.ViewDepthPlane, minZ - level.Elevation);

            // Set all other offsets to to the difference between the top of the section box and
            // the elevation of the level
            viewRange.SetOffset(PlanViewPlane.TopClipPlane, maxZ - level.Elevation);
            viewRange.SetOffset(PlanViewPlane.CutPlane, maxZ - level.Elevation);
            viewRange.SetOffset(PlanViewPlane.BottomClipPlane, maxZ - level.Elevation);

            using (Transaction t = new Transaction(doc, "Set View Range"))
            {
                t.Start();
                // Set the view range
                viewPlan.SetViewRange(viewRange);
                t.Commit();
            }
            return Autodesk.Revit.UI.Result.Succeeded;

        }
    }

    [Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)]
    public class Create_solid_rooms : IExternalCommand
    {

        public static ModelLine Makeline(Autodesk.Revit.DB.Document doc, XYZ pta, XYZ ptb)
        {
            ModelLine modelLine = null;
            double distance = pta.DistanceTo(ptb);
            if (distance < 0.01)
            {
                TaskDialog.Show("Error", "Distance" + distance);
                return modelLine;
            }

            XYZ norm = pta.CrossProduct(ptb);
            if (norm.GetLength() == 0)
            {
                XYZ aSubB = pta.Subtract(ptb);
                XYZ aSubBcrossz = aSubB.CrossProduct(XYZ.BasisZ);
                double crosslenght = aSubBcrossz.GetLength();
                if (crosslenght == 0)
                {
                    norm = XYZ.BasisY;
                }
                else
                {
                    norm = XYZ.BasisZ;
                }
            }

            Autodesk.Revit.DB.Plane plane = Autodesk.Revit.DB.Plane.CreateByNormalAndOrigin(norm, ptb);


            SketchPlane skplane = SketchPlane.Create(doc, plane);

            Autodesk.Revit.DB.Line line = Autodesk.Revit.DB.Line.CreateBound(pta, ptb);

            if (doc.IsFamilyDocument)
            {
                modelLine = doc.FamilyCreate.NewModelCurve(line, skplane) as ModelLine;
            }
            else
            {
                modelLine = doc.Create.NewModelCurve(line, skplane) as ModelLine;
            }
            if (modelLine == null)
            {
                TaskDialog.Show("Error", "Model line = null");
            }
            return modelLine;
        }

        public static bool IsZero(double a)
        {
            const double _eps = 1.0e-9;
            return _eps > Math.Abs(a);
        }

        public static bool IsEqual(double a, double b)
        {
            return IsZero(b - a);
        }

        public static int Compare(double a, double b)
        {
            return IsEqual(a, b) ? 0 : (a < b ? -1 : 1);
        }

        public static int Compare(XYZ p, XYZ q)
        {

            int diff = Compare(p.X, q.X);
            if (0 == diff)
            {
                diff = Compare(p.Y, q.Y);
                if (0 == diff)
                {
                    diff = Compare(p.Z, q.Z);
                }

            }

            return diff;
        }

        private static Wall CreateWall(FamilyInstance cube, Autodesk.Revit.DB.Curve curve, double height)
        {
            var doc = cube.Document;

            var wallTypeId = doc.GetDefaultElementTypeId(
              ElementTypeGroup.WallType);

            return Wall.Create(doc, curve.CreateReversed(),
              wallTypeId, cube.LevelId, height, 0, false,
              false);
        }


        ElementId _id_category_for_direct_shape = new ElementId(BuiltInCategory.OST_GenericModel);

        /// <summary>
        /// DirectShape parameter to populate with JSON
        /// dictionary containing all room properies
        /// </summary>
        BuiltInParameter _bip_properties = BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS;

        static AddInId appId = new AddInId(new Guid("5F88CC78-A137-4809-AAF8-A478F3B24BAB"));
        public Autodesk.Revit.UI.Result Execute(ExternalCommandData commandData, ref string message, ElementSet elementSet)
        {




            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Autodesk.Revit.ApplicationServices.Application app = uiapp.Application;
            Autodesk.Revit.DB.Document doc = uidoc.Document;


            string id_addin = uiapp.ActiveAddInId.ToString();

            IEnumerable<Room> rooms
              = new FilteredElementCollector(doc)
              .WhereElementIsNotElementType()
              .OfClass(typeof(SpatialElement))
              .Where(e => e.GetType() == typeof(Room))
              .Cast<Room>();

            using (Transaction tx = new Transaction(doc))
            {
                tx.Start("Generate Direct Shape Elements "
                  + "Representing Room Volumes");

                foreach (Room r in rooms)
                {
                    Debug.Print(r.Name);

                    GeometryElement geo = r.ClosedShell;

                    //Dictionary<string, string> param_values = GetParamValues(r);

                    //string json = FormatDictAsJson(param_values);
                    try
                    {
                        DirectShape ds = DirectShape.CreateElement(
                      doc, _id_category_for_direct_shape);

                        ds.ApplicationId = id_addin;
                        ds.ApplicationDataId = r.UniqueId;
                        ds.SetShape(geo.ToList<GeometryObject>());
                        //ds.get_Parameter(_bip_properties).Set(json);
                        ds.Name = "Room volume for " + r.Name;
                    }
                    catch (Exception)
                    {

                    }

                }
                tx.Commit();
            }

            return Autodesk.Revit.UI.Result.Succeeded;
        }
    }

    [Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)]
    public class DeleteLevel : IExternalCommand
    {
        static AddInId appId = new AddInId(new Guid("5F92CC68-A127-4809-AAF8-A478F3B24BAB"));
        public Autodesk.Revit.UI.Result Execute(ExternalCommandData commandData, ref string message, ElementSet elementSet)
        {
            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Autodesk.Revit.DB.Document doc = uidoc.Document;

            try
            {
                string filename = @"T:\Transfer\lopez\Book1.xlsx";
                using (ExcelPackage package = new ExcelPackage(new FileInfo(filename)))
                {
                    ExcelWorksheet sheet = package.Workbook.Worksheets.ElementAt(0);

                    int column = 5;
                    int number = Convert.ToInt32(sheet.Cells[2, column].Value);
                    sheet.Cells[2, column].Value = (number + 1); ;
                    package.Save();
                }
            }
            catch (Exception)
            {
                MessageBox.Show("Excel file not found", "");
            }

            //string comments = "DeleteLevel" + "_" + doc.Application.Username + "_" + doc.Title;
            //string filename = @"D:\Users\lopez\Desktop\Comments.txt";
            ////System.Diagnostics.Process.Start(filename);
            //StreamWriter writer = new StreamWriter(filename, true);
            ////writer.WriteLine( Environment.NewLine);
            //writer.WriteLine(DateTime.Now + " - " + comments);
            //writer.Close();

            MessageBox.Show("Please Select Level to be deleted.", "Steps 1", MessageBoxButtons.OK,
          MessageBoxIcon.Information);

            Level level = doc.GetElement(uidoc.Selection.PickObject(ObjectType.Element, "Select level")) as Level;

            MessageBox.Show("Please Select a new hosting level .", "Steps 2", MessageBoxButtons.OK,
          MessageBoxIcon.Information);

            Level levelBelow = doc.GetElement(uidoc.Selection.PickObject(ObjectType.Element, "Select level")) as Level;

            //Level levelBelow = new FilteredElementCollector(doc)
            //    .OfClass(typeof(Level))
            //    .Cast<Level>()
            //    .OrderBy(q => q.Elevation)
            //    .Where(q => q.Elevation <= level.Elevation)
            //    .FirstOrDefault();

            if (levelBelow == null)
            {
                TaskDialog.Show("Error", "No level below " + level.Elevation);
                return Autodesk.Revit.UI.Result.Succeeded;
            }

            List<string> paramsToAdjust = new List<string> { "Base Offset", "Sill Height", "Top Offset", "Height Offset From Level" };
            List<List<ElementId>> ids = new List<List<ElementId>>();
            GroupType gtype = null;
            List<Element> elements = new FilteredElementCollector(doc).WherePasses(new ElementLevelFilter(level.Id)).ToList();
            using (Transaction t = new Transaction(doc, "un group"))
            {
                t.Start();
                foreach (Element e in elements)
                {
                    Group gr_ = e as Group;
                    if (gr_ != null)
                    {
                        List<ElementId> ids_ = new List<ElementId>();
                        Group gr = e as Group;
                        gtype = gr.GroupType;
                        ICollection<ElementId> groups = gr.UngroupMembers();

                        foreach (var item in groups)
                        {
                            ids_.Add(item);
                        }

                        ids.Add(ids_);
                    }
                }
                t.Commit();
            }
            List<Element> elements2 = new FilteredElementCollector(doc).WherePasses(new ElementLevelFilter(level.Id)).ToList();
            using (Transaction t = new Transaction(doc, "Level Remap"))
            {
                t.Start();

                foreach (Element e in elements2)
                {
                    Parameter param = e.LookupParameter("Top Constraint");
                    if (param != null && param.ToString() != "Unconnected")
                    {
                        param.Set("Unconnected");
                    }
                    try
                    {
                        foreach (Parameter p in e.Parameters)
                        {
                            if (p.StorageType != StorageType.ElementId || p.IsReadOnly)
                                continue;
                            if (p.AsElementId() != level.Id)
                                continue;

                            double elevationDiff = level.Elevation - levelBelow.Elevation;

                            p.Set(levelBelow.Id);

                            foreach (string paramName in paramsToAdjust)
                            {
                                Parameter pToAdjust = e.LookupParameter(paramName);
                                if (pToAdjust != null && !pToAdjust.IsReadOnly)
                                    pToAdjust.Set(pToAdjust.AsDouble() + elevationDiff);
                            }
                        }
                    }
                    catch (Exception)
                    {
                        MessageBox.Show("!", "RevitLookup", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        throw;
                    }
                }
                t.Commit();
                MessageBox.Show("Delete level.", "Steps 3", MessageBoxButtons.OK,
          MessageBoxIcon.Information);
            }
            using (Transaction t = new Transaction(doc, "Level Remap"))
            {
                t.Start("Group");
                foreach (var item in ids)
                {
                    Group grpNew = doc.Create.NewGroup(item);

                    // Access the name of the previous group type 
                    // and change the new group type to previous 
                    // group type to retain the previous group 
                    // configuration

                    FilteredElementCollector coll
                      = new FilteredElementCollector(doc)
                        .OfClass(typeof(GroupType));

                    IEnumerable<GroupType> grpTypes
                      = from GroupType g in coll
                        where g.Name == gtype.Name.ToString()
                        select g;

                    grpNew.GroupType = grpTypes.First<GroupType>();
                }


                t.Commit();
            }

            return Autodesk.Revit.UI.Result.Succeeded;
        }
    }


    [Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)]
    public class remove_paint : IExternalCommand
    {


        public class MassSelectionFilter : ISelectionFilter
        {
            public bool AllowElement(Element element)
            {
                if (element.Category.Name == "Walls")
                {
                    return true;
                }
                return false;
            }

            public bool AllowReference(Reference refer, XYZ point)
            {
                return false;
            }
        }

        static AddInId appId = new AddInId(new Guid("5F88CC78-A137-6509-AAF8-A578F3B24BAB"));
        public Autodesk.Revit.UI.Result Execute(ExternalCommandData commandData, ref string message, ElementSet elementSet)
        {
            try
            {
                string filename = @"T:\Transfer\lopez\Book1.xlsx";
                using (ExcelPackage package = new ExcelPackage(new FileInfo(filename)))
                {
                    ExcelWorksheet sheet = package.Workbook.Worksheets.ElementAt(0);

                    int column = 21;
                    int number = Convert.ToInt32(sheet.Cells[2, column].Value);
                    sheet.Cells[2, column].Value = (number + 1); ;
                    package.Save();
                }
            }
            catch (Exception)
            {
                MessageBox.Show("Excel file not found", "");
            }


            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Autodesk.Revit.DB.Document doc = uidoc.Document;

            ReferenceArray ra = new ReferenceArray();
            ISelectionFilter selFilter = new MassSelectionFilter();

            ICollection<ElementId> ids = uidoc.Selection.GetElementIds();
            List<Element> walls = new List<Element>();

            foreach (var item in ids)
            {

                walls.Add(doc.GetElement(item));
            }

            //IList<Element> refList = uidoc.Selection.PickElementsByRectangle(selFilter, "Select multiple faces") as IList<Element>;
            //Reference hasPickOne = uidoc.Selection.PickObject(Autodesk.Revit.UI.Selection.ObjectType.Element, selFilter);
            ICollection<Reference> refList = uidoc.Selection.PickObjects(Autodesk.Revit.UI.Selection.ObjectType.Face, "Select ceilings to be reproduced in rhino geometry");

            //IList<Reference> refList = uidoc.Selection.PickObjects(ObjectType.Element, selFilter, "Pick elements to add to selection filter");


            using (Transaction trans = new Transaction(doc, "ViewDuplicate"))
            {
                trans.Start();

                foreach (var item_myRefWall in refList)
                {
                    Element e = doc.GetElement(item_myRefWall);


                    //Wall wall_ = e as Wall;
                    GeometryElement geometryElement = e.get_Geometry(new Options());
                    foreach (GeometryObject geometryObject in geometryElement)
                    {
                        if (geometryObject is Solid)
                        {
                            Solid solid = geometryObject as Solid;
                            foreach (Face face_ in solid.Faces)
                            {
                                doc.RemovePaint(e.Id, face_);
                            }
                        }
                    }
                }
                trans.Commit();
            }
            return Autodesk.Revit.UI.Result.Succeeded;
        }
    }

    [Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)]
    public class make_line : IExternalCommand
    {
        public static ModelLine Makeline(Autodesk.Revit.DB.Document doc, XYZ pta, XYZ ptb)
        {
            ModelLine modelLine = null;
            double distance = pta.DistanceTo(ptb);
            if (distance < 0.01)
            {
                TaskDialog.Show("Error", "Distance" + distance);
                return modelLine;
            }

            XYZ norm = pta.CrossProduct(ptb);
            if (norm.GetLength() == 0)
            {
                XYZ aSubB = pta.Subtract(ptb);
                XYZ aSubBcrossz = aSubB.CrossProduct(XYZ.BasisZ);
                double crosslenght = aSubBcrossz.GetLength();
                if (crosslenght == 0)
                {
                    norm = XYZ.BasisY;
                }
                else
                {
                    norm = XYZ.BasisZ;
                }
            }

            //Autodesk.Revit.DB.Plane plane = Autodesk.Revit.DB.Plane.CreateByNormalAndOrigin(norm, ptb);

            Autodesk.Revit.DB.Line line = Autodesk.Revit.DB.Line.CreateUnbound(ptb, XYZ.BasisZ /* XYZ.BasisZ*/);

            Autodesk.Revit.DB.Plane pl = Autodesk.Revit.DB.Plane.CreateByThreePoints(pta,
             line.Evaluate(5, false), ptb);

            SketchPlane skplane = SketchPlane.Create(doc, pl);

            Autodesk.Revit.DB.Line line2 = Autodesk.Revit.DB.Line.CreateBound(pta, ptb);

            if (doc.IsFamilyDocument)
            {
                modelLine = doc.FamilyCreate.NewModelCurve(line2, skplane) as ModelLine;
            }
            else
            {
                modelLine = doc.Create.NewModelCurve(line2, skplane) as ModelLine;
            }
            if (modelLine == null)
            {
                TaskDialog.Show("Error", "Model line = null");
            }
            return modelLine;
        }

        static AddInId appId = new AddInId(new Guid("5F56AA78-A136-6509-AAF8-A478F3B24BAB"));
        public Autodesk.Revit.UI.Result Execute(ExternalCommandData commandData, ref string message, ElementSet elementSet)
        {
            UIDocument uidoc = commandData.Application.ActiveUIDocument;
            Autodesk.Revit.DB.Document doc = uidoc.Document;

            //ObjectSnapTypes snapTypes = ObjectSnapTypes.Endpoints | ObjectSnapTypes.Intersections;
            //XYZ point = uidoc.Selection.PickPoint(snapTypes, "Select an end point or intersection");


            //  XYZ point_in_3d_1 = uidoc.Selection.PickPoint(snapTypes,
            //"Please pick a point on the plane"
            //+ " defined by the selected face");


            Reference myRef2 = uidoc.Selection.PickObject(ObjectType.PointOnElement);
            Element e2 = doc.GetElement(myRef2.ElementId);

            GeometryObject geomObj2 = e2.GetGeometryObjectFromReference(myRef2);
            //			Point p = geomObj as Point;
            XYZ p2 = myRef2.GlobalPoint;
            //string pointString2 = p2 == null ? "<reference has no globalpoint>" : string.Format("{0}", p2);
            //TaskDialog.Show("Element Info", e2.Name + "   " + pointString2);

            Reference myRef1 = uidoc.Selection.PickObject(ObjectType.PointOnElement);
            Element e1 = doc.GetElement(myRef1.ElementId);

            GeometryObject geomObj1 = e1.GetGeometryObjectFromReference(myRef1);
            //			Point p = geomObj as Point;
            XYZ p1 = myRef1.GlobalPoint;
            //string pointString1 = p1 == null ? "<reference has no globalpoint>" : string.Format("{0}", p1);
            //TaskDialog.Show("Element Info", e1.Name + "   " + pointString1);

            using (Transaction tx = new Transaction(doc))
            {
                tx.Start("Make Line");

                ModelLine ml = Makeline(doc, p1, p2);

                //Autodesk.Revit.DB.Line line = Autodesk.Revit.DB.Line.CreateUnbound(p2, XYZ.BasisZ);

                //Autodesk.Revit.DB.Plane pl = Autodesk.Revit.DB.Plane.CreateByThreePoints(p1,
                // line.Evaluate(5, false), p2);

                SketchPlane sketch = ml.SketchPlane /*SketchPlane.Create(doc, pl)*/;
                doc.ActiveView.SketchPlane = sketch;
                doc.ActiveView.ShowActiveWorkPlane();

                try
                {

                }
                catch (Exception)
                {

                    //TaskDialog.Show("!", "the project does not contain the parameter (Audit Orthogonality)assigned to walls" +
                    //    " ");
                }

                tx.Commit();
            }




            //uidoc.Selection.SetElementIds(ele.Select(q => q.Id).ToList());
            return Autodesk.Revit.UI.Result.Succeeded;
        }
    }












    [Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)]
    public class DeleteAllSheets : IExternalCommand
    {
        static AddInId appId = new AddInId(new Guid("5F22CC78-A137-4819-AAF1-A678F6B22BAB"));
        public Autodesk.Revit.UI.Result Execute(ExternalCommandData commandData, ref string message, ElementSet elementSet)
        {
            try
            {
                string filename = @"T:\Transfer\lopez\Book1.xlsx";
                using (ExcelPackage package = new ExcelPackage(new FileInfo(filename)))
                {
                    ExcelWorksheet sheet = package.Workbook.Worksheets.ElementAt(0);

                    int column = 11;
                    int number = Convert.ToInt32(sheet.Cells[2, column].Value);
                    sheet.Cells[2, column].Value = (number + 1); ;
                    package.Save();
                }
            }
            catch (Exception)
            {
                MessageBox.Show("Excel file not found", "");
            }

            Form15 form = new Form15();

            form.ShowDialog();

            if (form.DialogResult == DialogResult.Cancel)
            {
                return Autodesk.Revit.UI.Result.Cancelled;
            }

            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Autodesk.Revit.DB.Document doc = uidoc.Document;
            Autodesk.Revit.ApplicationServices.Application app = uiapp.Application;


            IEnumerable<ViewSheet> viewSheet = from elem in new FilteredElementCollector(doc)
                                                 .OfClass(typeof(ViewSheet))
                                                 .OfCategory(BuiltInCategory.OST_Sheets)
                                               let type = elem as ViewSheet
                                               where type.Name != null
                                               select type;

            IEnumerable<Autodesk.Revit.DB.Viewport> viewports = from elem in new FilteredElementCollector(doc)
                                                 .OfClass(typeof(Autodesk.Revit.DB.Viewport))
                                                 .OfCategory(BuiltInCategory.OST_Views)
                                                                let type = elem as Autodesk.Revit.DB.Viewport
                                                                where type.Name != null
                                                                select type;

            List<ViewSheet> viewSheet_list = new List<ViewSheet>();
            List<ViewSheet> viewports_ = new List<ViewSheet>();

            foreach (var item in viewSheet)
            {
                viewSheet_list.Add(item);
            }

            try
            {
                Autodesk.Revit.DB.View viewTemplate = (from v in new FilteredElementCollector(doc).OfClass(typeof(Autodesk.Revit.DB.View)).Cast<Autodesk.Revit.DB.View>()
                                                       where !v.IsTemplate && v.Name == "Home"
                                                       select v).First();


                uidoc.ActiveView = viewTemplate;
            }
            catch (Exception)
            {
                TaskDialog.Show("Warning", "DISCLAIMER view was not found in this project");
                return Autodesk.Revit.UI.Result.Cancelled;
                //throw;
            }

            using (Transaction t = new Transaction(doc, "Delete All Sheets"))
            {
                t.Start();


                try
                {
                    for (int i = 0; i < viewports_.ToArray().Length; i++)
                    {
                        doc.Delete(viewports_.ToArray()[i].Id);
                    }
                }
                catch (Exception)
                {
                    TaskDialog.Show("Warning", "Active view must not be a sheet");
                    throw;
                }

                try
                {
                    for (int i = 0; i < viewSheet_list.ToArray().Length; i++)
                    {
                        doc.Delete(viewSheet_list.ToArray()[i].Id);
                    }
                }
                catch (Exception)
                {
                    TaskDialog.Show("Warning", "Active view must not be a sheet");
                    throw;
                }




                t.Commit();


            }

            return Autodesk.Revit.UI.Result.Succeeded;
        }


    }

    [Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)]
    public class Create_Sun_Eye_view : IExternalCommand
    {
        public static List<List<XYZ>> GeTpoints(Autodesk.Revit.DB.Document doc_, List<List<XYZ>> xyz_faces, IList<CurveLoop> faceboundaries, List<List<Face>> list_faces)
        {
            if (list_faces == null)
            {
                list_faces = new List<List<Face>>();
            }

            for (int i = 0; i < list_faces.ToArray().Length; i++)
            {
                List<XYZ> puntos_ = new List<XYZ>();
                foreach (Face f in list_faces.ToArray()[i])
                {

                    faceboundaries = f.GetEdgesAsCurveLoops();//new trying to get the outline of the face instead of the edges
                    EdgeArrayArray edgeArrays = f.EdgeLoops;
                    foreach (CurveLoop edges in faceboundaries)
                    {
                        puntos_.Add(null);
                        foreach (Autodesk.Revit.DB.Curve edge in edges)
                        {
                            XYZ testPoint1 = edge.GetEndPoint(1);
                            XYZ testPoint2 = edge.GetEndPoint(0);
                            double lenght = Math.Round(edge.ApproximateLength, 0);
                            double lenght_convert = UnitUtils.Convert(lenght, DisplayUnitType.DUT_DECIMAL_FEET, DisplayUnitType.DUT_MILLIMETERS);

                            double x = Math.Round(testPoint1.X, 0);
                            double y = Math.Round(testPoint1.Y, 0);
                            double z = Math.Round(testPoint1.Z, 0);

                            ElementClassFilter filter = new ElementClassFilter(typeof(Floor));

                            XYZ newpt = new XYZ(x, y, z);

                            if (!puntos_.Contains(testPoint1))
                            {
                                puntos_.Add(testPoint1);

                            }
                        }
                    }
                    int num = f.EdgeLoops.Size;
                }
                xyz_faces.Add(puntos_);
            }
            return xyz_faces;
        }
        //private ModelLine Makeline(Autodesk.Revit.DB.Document doc, XYZ pta, XYZ ptb)
        //{
        //    ModelLine modelLine = null;
        //    double distance = pta.DistanceTo(ptb);
        //    if (distance < 0.01)
        //    {
        //        TaskDialog.Show("Error", "Distance" + distance);
        //        return modelLine;
        //    }

        //    XYZ norm = pta.CrossProduct(ptb);
        //    if (norm.GetLength() == 0)
        //    {
        //        XYZ aSubB = pta.Subtract(ptb);
        //        XYZ aSubBcrossz = aSubB.CrossProduct(XYZ.BasisZ);
        //        double crosslenght = aSubBcrossz.GetLength();
        //        if (crosslenght == 0)
        //        {
        //            norm = XYZ.BasisY;
        //        }
        //        else
        //        {
        //            norm = XYZ.BasisZ;
        //        }
        //    }

        //    Autodesk.Revit.DB.Plane plane = Autodesk.Revit.DB.Plane.CreateByNormalAndOrigin(norm, ptb);


        //    SketchPlane skplane = SketchPlane.Create(doc, plane);

        //    Autodesk.Revit.DB.Line line = Autodesk.Revit.DB.Line.CreateBound(pta, ptb);

        //    if (doc.IsFamilyDocument)
        //    {
        //        modelLine = doc.FamilyCreate.NewModelCurve(line, skplane) as ModelLine;
        //    }
        //    else
        //    {
        //        modelLine = doc.Create.NewModelCurve(line, skplane) as ModelLine;
        //    }
        //    if (modelLine == null)
        //    {
        //        TaskDialog.Show("Error", "Model line = null");
        //    }
        //    return modelLine;
        //}

        public static ModelLine Makeline(Autodesk.Revit.DB.Document doc, XYZ pta, XYZ ptb)
        {
            ModelLine modelLine = null;
            double distance = pta.DistanceTo(ptb);
            if (distance < 0.01)
            {
                TaskDialog.Show("Error", "Distance" + distance);
                return modelLine;
            }

            XYZ norm = pta.CrossProduct(ptb);
            if (norm.GetLength() == 0)
            {
                XYZ aSubB = pta.Subtract(ptb);
                XYZ aSubBcrossz = aSubB.CrossProduct(XYZ.BasisZ);
                double crosslenght = aSubBcrossz.GetLength();
                if (crosslenght == 0)
                {
                    norm = XYZ.BasisY;
                }
                else
                {
                    norm = XYZ.BasisZ;
                }
            }

            //Autodesk.Revit.DB.Plane plane = Autodesk.Revit.DB.Plane.CreateByNormalAndOrigin(norm, ptb);

            Autodesk.Revit.DB.Line line = Autodesk.Revit.DB.Line.CreateUnbound(ptb, XYZ.BasisZ /* XYZ.BasisZ*/);

            Autodesk.Revit.DB.Plane pl = Autodesk.Revit.DB.Plane.CreateByThreePoints(pta,
             line.Evaluate(5, false), ptb);

            SketchPlane skplane = SketchPlane.Create(doc, pl);

            Autodesk.Revit.DB.Line line2 = Autodesk.Revit.DB.Line.CreateBound(pta, ptb);

            if (doc.IsFamilyDocument)
            {
                modelLine = doc.FamilyCreate.NewModelCurve(line2, skplane) as ModelLine;
            }
            else
            {
                modelLine = doc.Create.NewModelCurve(line2, skplane) as ModelLine;
            }
            if (modelLine == null)
            {
                TaskDialog.Show("Error", "Model line = null");
            }
            return modelLine;
        }

        public double MapValue(double start_n, double end_n, double mapped_n_menusone, double mapped_n_one, double number_tobe_map)
        {
            return mapped_n_menusone + (mapped_n_one - mapped_n_menusone) * ((number_tobe_map - start_n) / (end_n - start_n));
        }
        public ViewOrientation3D GetCurrentViewOrientation(UIDocument doc)
        {
            XYZ UpDir = doc.ActiveView.UpDirection;
            XYZ ViewDir = doc.ActiveView.ViewDirection;
            XYZ ViewInvDir = InvCoord(ViewDir);
            XYZ eye = new XYZ(0, 0, 0);
            XYZ up = UpDir;
            XYZ forward = ViewInvDir;
            ViewOrientation3D MyNewOrientation = new ViewOrientation3D(eye, up, forward);
            return MyNewOrientation;
        }

        public XYZ InvCoord(XYZ MyCoord)
        {
            XYZ invcoord = new XYZ((Convert.ToDouble(MyCoord.X * -1)),
                (Convert.ToDouble(MyCoord.Y * -1)),
                (Convert.ToDouble(MyCoord.Z * -1)));
            return invcoord;
        }
        public XYZ CrossProduct(XYZ v1, XYZ v2)
        {
            double x, y, z;
            x = v1.Y * v2.Z - v2.Y * v1.Z;
            y = (v1.X * v2.Z - v2.X * v1.Z) * -1;
            z = v1.X * v2.Y - v2.X * v1.Y;
            var rtnvector = new XYZ(x, y, z);
            return rtnvector;
        }
        public XYZ VectorFromHorizVertAngles(double angleHorizD, double angleVertD)
        {
            double degToRadian = Math.PI * 2 / 360;
            double angleHorizR = angleHorizD * degToRadian;
            double angleVertR = angleVertD * degToRadian;
            double a = Math.Cos(angleVertR);
            double b = Math.Cos(angleHorizR);
            double c = Math.Sin(angleHorizR);
            double d = Math.Sin(angleVertR);
            return new XYZ(a * b, a * c, d);
        }

        public class Vector3D
        {
            public Vector3D(XYZ revitXyz)
            {
                XYZ = revitXyz;
            }
            public Vector3D() : this(XYZ.Zero)
            { }
            public Vector3D(double x, double y, double z)
              : this(new XYZ(x, y, z))
            { }
            public XYZ XYZ { get; private set; }
            public double X => XYZ.X;
            public double Y => XYZ.Y;
            public double Z => XYZ.Z;
            public Vector3D CrossProduct(Vector3D source)
            {
                return new Vector3D(XYZ.CrossProduct(source.XYZ));
            }
            public double GetLength()
            {
                return XYZ.GetLength();
            }
            public override string ToString()
            {
                return XYZ.ToString();
            }
            public static Vector3D BasisX => new Vector3D(
              XYZ.BasisX);
            public static Vector3D BasisY => new Vector3D(
              XYZ.BasisY);
            public static Vector3D BasisZ => new Vector3D(
              XYZ.BasisZ);
        }
        static AddInId appId = new AddInId(new Guid("8D3F5703-A09A-6ED6-864C-5720329D9677"));
        public Autodesk.Revit.UI.Result Execute(ExternalCommandData commandData, ref string message, ElementSet elementSet)
        {
            UIDocument uidoc = commandData.Application.ActiveUIDocument;
            Autodesk.Revit.DB.Document doc = uidoc.Document;
            UIApplication uiapp = commandData.Application;
            Autodesk.Revit.ApplicationServices.Application app = doc.Application;

            try
            {
                string filename = @"T:\Transfer\lopez\Book1.xlsx";
                using (ExcelPackage package = new ExcelPackage(new FileInfo(filename)))
                {
                    ExcelWorksheet sheet = package.Workbook.Worksheets.ElementAt(0);

                    int column = 2;
                    int number = Convert.ToInt32(sheet.Cells[2, column].Value);
                    sheet.Cells[2, column].Value = (number + 1); ;
                    package.Save();
                }
            }
            catch (Exception)
            {
                MessageBox.Show("Excel file not found", "");
            }

            Form26 form2 = new Form26();
            form2.ShowDialog();

            ProjectLocation plCurrent = doc.ActiveProjectLocation;
           

            Autodesk.Revit.DB.View currentView = uidoc.ActiveView;
            SunAndShadowSettings sunSettings = currentView.SunAndShadowSettings;


            // Set the initial direction of the sun at ground level (like sunrise level)
            XYZ initialDirection = XYZ.BasisY;

            // Get the altitude of the sun from the sun settings
            double altitude = sunSettings.GetFrameAltitude(
              sunSettings.ActiveFrame);

            // Create a transform along the X axis based on the altitude of the sun
            Autodesk.Revit.DB.Transform altitudeRotation = Autodesk.Revit.DB.Transform
              .CreateRotation(XYZ.BasisX, altitude);

            // Create a rotation vector for the direction of the altitude of the sun
            XYZ altitudeDirection = altitudeRotation
              .OfVector(initialDirection);

            // Get the azimuth from the sun settings of the scene
            double azimuth = sunSettings.GetFrameAzimuth(
              sunSettings.ActiveFrame);

            // Correct the value of the actual azimuth with true north

            // Get the true north angle of the project
            Element projectInfoElement
              = new FilteredElementCollector(doc)
                .OfCategory(BuiltInCategory.OST_ProjectBasePoint)
                .FirstElement();

            BuiltInParameter bipAtn
              = BuiltInParameter.BASEPOINT_ANGLETON_PARAM;

            Parameter patn = projectInfoElement.get_Parameter(
              bipAtn);

            double trueNorthAngle = patn.AsDouble();

            // Add the true north angle to the azimuth
            double actualAzimuth = 2 * Math.PI - azimuth + trueNorthAngle;

            // Create a rotation vector around the Z axis
            Autodesk.Revit.DB.Transform azimuthRotation = Autodesk.Revit.DB.Transform
              .CreateRotation(XYZ.BasisZ, actualAzimuth);

            // Finally, calculate the direction of the sun
            XYZ sunDirection = azimuthRotation.OfVector(
              altitudeDirection);




            //BuiltInParameter bipAtn = BuiltInParameter.BASEPOINT_ANGLETON_PARAM;
            //Parameter patn = projectInfoElement.get_Parameter(bipAtn);
            //double atn = patn.AsDouble();
            //foreach (ProjectLocation location in doc.ProjectLocations)
            //{
            //    ProjectPosition projectPosition
            //      = location./*get_ProjectPosition(XYZ.Zero)*/ GetProjectPosition(XYZ.Zero);
            //    double x = projectPosition.EastWest;
            //    double y = projectPosition.NorthSouth;
            //    XYZ pnp = new XYZ(x, y, 0.0);
            //    double pna = projectPosition.Angle;
            //}
            IEnumerable<ViewFamilyType> viewFamilyTypes = from elem in new FilteredElementCollector(doc).OfClass(typeof(ViewFamilyType))
                                                          let type = elem as ViewFamilyType
                                                          where type.ViewFamily == ViewFamily.ThreeDimensional
                                                          select type;




            //XYZ initialDirection = XYZ.BasisY;
            //double altitude = sunSettings.GetFrameAltitude(sunSettings.ActiveFrame);
            //Autodesk.Revit.DB.Transform altitudeRotation = Autodesk.Revit.DB.Transform.CreateRotation(XYZ.BasisX, altitude);
            //XYZ altitudeDirection = altitudeRotation.OfVector(initialDirection);
            //double azimuth = sunSettings.GetFrameAzimuth(sunSettings.ActiveFrame);
            //double actualAzimuth = 2 * Math.PI - azimuth;
            //Autodesk.Revit.DB.Transform azimuthRotation = Autodesk.Revit.DB.Transform.CreateRotation(XYZ.BasisZ, actualAzimuth);
            //double northrotation = 2 * Math.PI - atn;
            //XYZ sunDirection = azimuthRotation.OfVector(altitudeDirection);
            //Autodesk.Revit.DB.Transform tran01 = Autodesk.Revit.DB.Transform.CreateRotationAtPoint(XYZ.BasisZ, northrotation * -1, new XYZ(0, 0, 0));
            //XYZ new_p = tran01.OfVector(sunDirection);
            //sunDirection = new_p;

            XYZ UpDir = uidoc.ActiveView.UpDirection;
            Form8 form = new Form8();

            ViewOrientation3D viewOrientation3D;
            using (Transaction tr1 = new Transaction(doc))
            {
                form.ShowDialog();
                tr1.Start("Place vs in sheet");
                View3D view3D = View3D.CreateIsometric(doc, viewFamilyTypes.First().Id);
                tr1.SetName("Create view " + view3D.Name);
                view3D.Name = form.textBox1.Text;

                XYZ eye = XYZ.Zero;
                XYZ inverted_sun_location = InvCoord(sunDirection);
                
                XYZ origin_b = new XYZ(0, 0, 0);
                XYZ normal_B = new XYZ(1, 0, 0);

                
                Autodesk.Revit.DB.Transform trans3;
                Autodesk.Revit.DB.Transform trans_inverted_direction;


                Autodesk.Revit.DB.Line line2 = Autodesk.Revit.DB.Line.CreateUnbound(new XYZ(0, 0, 0), sunDirection);
                XYZ vect1 = line2.Direction * (100000 / 304.8);
                XYZ vect2 = vect1 + new XYZ(0, 0, 0);
                ModelCurve mc = Makeline(doc, line2.Origin, vect2);

                Autodesk.Revit.DB.Plane Plane_mirror = Autodesk.Revit.DB.Plane.CreateByNormalAndOrigin(/*mc.SketchPlane.GetPlane().Normal*/ new XYZ(0, 0, 1), sunDirection);
                Autodesk.Revit.DB.Plane forward_dir = Autodesk.Revit.DB.Plane.CreateByNormalAndOrigin(mc.SketchPlane.GetPlane().Normal, sunDirection);

                trans3 = Autodesk.Revit.DB.Transform.CreateReflection(Plane_mirror);
                XYZ inv_sun_mirrored = trans3.OfVector(inverted_sun_location);

                trans_inverted_direction = Autodesk.Revit.DB.Transform.CreateReflection(forward_dir);
                XYZ inverted_direction = trans_inverted_direction.OfVector(inverted_sun_location);

                Autodesk.Revit.DB.Line line3 = Autodesk.Revit.DB.Line.CreateUnbound(new XYZ(0, 0, 0), inv_sun_mirrored);
                XYZ invertedvect1 = line3.Direction * (100000 / 304.8);
                XYZ invertedvect2 = invertedvect1 + new XYZ(0, 0, 0);
                Makeline(doc, line2.Origin, invertedvect2);


                XYZ cross_product = CrossProduct(/*inv_sun,*/ /*new_p*/ inv_sun_mirrored, inverted_sun_location);
                Autodesk.Revit.DB.Line line4 = Autodesk.Revit.DB.Line.CreateUnbound(new XYZ(0, 0, 0), cross_product);
                XYZ invertedvect3 = line4.Direction * (100000 / 304.8);
                XYZ invertedvect4 = invertedvect3 + new XYZ(0, 0, 0);
                Makeline(doc, line2.Origin, invertedvect4);


                XYZ cross_product_up = CrossProduct(/*inv_sun,*/ /*new_p*/ line2.Direction, cross_product);

                //XYZ origin_c = sunDirection;
                //XYZ normal_c = cross_product;
                //Autodesk.Revit.DB.Plane Plane_mirror_c = Autodesk.Revit.DB.Plane.CreateByNormalAndOrigin(normal_c, origin_c);

                //XYZ no_z_90d = new XYZ(cross_product.X, cross_product.Y, 0);
                //XYZ no_z_orig = new XYZ(inverted_sun_location.X, inverted_sun_location.Y, 0);
                //Autodesk.Revit.DB.Line orientLine = Autodesk.Revit.DB.Line.CreateBound(sunDirection, inverted_sun_location);

                //double angle_1 = /*XYZ.BasisX*/sunDirection.AngleTo(XYZ.BasisY);
                //double angleDegrees = angle_1 * 180 / Math.PI;
                //if (no_z_90d.X < no_z_orig.X)
                //    angle_1 = 2 * Math.PI - angle_1;
                //double angleDegreesCorrected = angle_1 * 180 / Math.PI;
                //Autodesk.Revit.DB.Transform rot = Autodesk.Revit.DB.Transform.CreateRotation(orientLine.Direction, angleDegrees);
                //XYZ rotated_vec = rot.OfVector(-1 * cross_product);
                //XYZ dir = new XYZ(0, 0, 1);
                //XYZ normal = orientLine.Direction.Normalize();


                //XYZ cross = normal.CrossProduct(dir);
                //Autodesk.Revit.DB.Line line5 = Autodesk.Revit.DB.Line.CreateUnbound(new XYZ(5, 5, 5), cross_product);
                //Makeline(doc, line5.Origin, line5.Evaluate(5, false));



                XYZ startPoint = sunDirection;
                XYZ endPoint = inverted_sun_location;
                Autodesk.Revit.DB.Line geomLine = Autodesk.Revit.DB.Line.CreateBound(startPoint, endPoint);
                XYZ pntCenter = geomLine.Evaluate(0.0, true);
                Autodesk.Revit.DB.Line geomLine2 = Autodesk.Revit.DB.Line.CreateBound(doc.ActiveView.Origin, XYZ.BasisZ);
                Autodesk.Revit.DB.Plane geomPlane = Autodesk.Revit.DB.Plane.CreateByNormalAndOrigin(sunDirection, sunDirection);
                if (form.checkBox1.Checked)
                {
                    SketchPlane sketch = SketchPlane.Create(doc, geomPlane);
                    sketch.Name = view3D.Name;
                    doc.ActiveView.SketchPlane = sketch;
                    doc.ActiveView.ShowActiveWorkPlane();
                    view3D.SketchPlane = sketch;
                    view3D.ShowActiveWorkPlane();
                }


                Autodesk.Revit.DB.Transform rot2 = Autodesk.Revit.DB.Transform.CreateRotation(/*orientLine.Direction*/mc.GeometryCurve.GetEndPoint(1), -2.60);
                XYZ rotated_vec2 = rot2.OfVector(inverted_direction);

                viewOrientation3D = new ViewOrientation3D(/*eye*/ mc.GeometryCurve.GetEndPoint(0), cross_product_up/*rotated_vec2*/ * -1, /*inverted_direction*/rotated_vec2);
                view3D.SetOrientation(viewOrientation3D);
                view3D.SaveOrientationAndLock();
                tr1.Commit();
            }


            //IList<Reference> refList = new List<Reference>();
            //try
            //{
            //    while (true)
            //        refList.Add(uidoc.Selection.PickObject(ObjectType.Element, "Select elements in order to be renumbered. ESC when finished."));
            //}
            //catch
            //{ }

            //ICollection<Reference> refList = uidoc.Selection.PickObjects(Autodesk.Revit.UI.Selection.ObjectType.Face, "Select ceilings to be reproduced in rhino geometry");

            //List<ElementId> elementos = new List<ElementId>();

            //List<Reference> my_faces = new List<Reference>();
            //List<List<XYZ>> xyz_faces = new List<List<XYZ>>();
            //IList<Face> face_with_regions = new List<Face>();
            //List<List<Face>> Faces_lists_excel = new List<List<Face>>();
            //IList<CurveLoop> faceboundaries = new List<CurveLoop>();
            //List<Face> faces_picked = new List<Face>();

            //ICollection<Reference> my_faces_ = uidoc.Selection.PickObjects(Autodesk.Revit.UI.Selection.ObjectType.Face, "Select ceilings to be reproduced in rhino geometry");

            //foreach (var item in my_faces_)
            //{
               
            //    my_faces.Add(item);
            //}

            //foreach (var item in my_faces_)
            //{
            //    elementos.Add(item.ElementId);
            //}

            //foreach (var item_myRefWall in my_faces)
            //{
            //    Element e = doc.GetElement(item_myRefWall);
            //    GeometryObject geoobj = e.GetGeometryObjectFromReference(item_myRefWall);
            //    Face face = geoobj as Face;
            //    PlanarFace planarFace = face as PlanarFace;
            //    XYZ normal = planarFace.ComputeNormal(new UV(planarFace.Origin.X, planarFace.Origin.Y));


               
            //    faces_picked.Add(face);


            //    if (item_myRefWall == my_faces.ToArray().Last())
            //    {
            //        Faces_lists_excel.Add(faces_picked);

            //        //names.Add(name_of_roof);
            //    }
            //}

            //GeTpoints(doc, xyz_faces, faceboundaries, Faces_lists_excel);


            //for (int i = 0; i < xyz_faces.ToArray().Length; i++)
            //{

            //    foreach (var xyz_ in xyz_faces.ToArray()[i])
            //    {
                    

            //        if (xyz_  != null)
            //        {
            //            ElementClassFilter filter = new ElementClassFilter(typeof(Floor));
            //            XYZ dir2 = /*new XYZ(0, 0, 0)*/sunDirection - viewOrientation3D.ForwardDirection;
            //            ReferenceIntersector refIntersector = new ReferenceIntersector(filter, FindReferenceTarget.Face, doc.ActiveView as View3D);
            //            ReferenceWithContext referenceWithContext = refIntersector.FindNearest(xyz_, viewOrientation3D.ForwardDirection);

            //            if (referenceWithContext != null)
            //            {
            //                Reference reference = referenceWithContext.GetReference();
            //                XYZ intersection = reference.GlobalPoint;


            //                Autodesk.Revit.DB.Transform transun = Autodesk.Revit.DB.Transform.CreateTranslation(XYZ.BasisY);
            //                XYZ projection = transun.OfPoint(intersection * -1);

            //                Autodesk.Revit.DB.Plane Plane_mirror2 = Autodesk.Revit.DB.Plane.CreateByNormalAndOrigin(new XYZ(1, 0, 0), intersection);

            //                Autodesk.Revit.DB.Transform trans4 = Autodesk.Revit.DB.Transform.CreateReflection(Plane_mirror2);

            //                XYZ mirroredvec = trans4.OfVector(projection);
            //                XYZ cross_productofpoints = CrossProduct(mirroredvec, projection);

            //                Autodesk.Revit.DB.Plane AngulotoRotatePoint = Autodesk.Revit.DB.Plane.CreateByNormalAndOrigin(cross_productofpoints, projection);


            //                Autodesk.Revit.DB.Transform rot3 = Autodesk.Revit.DB.Transform.CreateRotation(projection, 90);
            //                XYZ rotated_vec3 = rot3.OfVector(projection);



            //                //Autodesk.Revit.DB.Transform x_ = Autodesk.Revit.DB.Transform.Identity;
            //                //x_ = x_.ScaleBasis(10);

            //                Autodesk.Revit.DB.Line line = Autodesk.Revit.DB.Line.CreateBound(intersection, rotated_vec3);


            //                //line.CreateTransformed(x_);

            //                //Autodesk.Revit.DB.Curve newCurve = curve.get_Transformed(x);

            //                using (Transaction tr2 = new Transaction(doc))
            //                {
            //                    tr2.Start("line");
            //                    Makeline(doc, xyz_, intersection /*line.GetEndPoint(0), line.GetEndPoint(1)*/ /*intersection, projection*/ /*, viewOrientation3D.ForwardDirection*/ /*, sunDirection*/ /*, p2*/);

                                

            //                    tr2.Commit();
            //                }
            //            }
            //        }
            //    }
            //}

           
            return Autodesk.Revit.UI.Result.Succeeded;
        }
    }

    [Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)]
    public class Rhino_access : IExternalCommand

    {
        #region "Public Members"
        private IRhino5x64Application _RhinoApp = null;
        private IRhinoScript _RhinoScript = null;
        #endregion

        public class RhinoScript
        {
            #region "Public Members"
            private IRhino5x64Application _RhinoApp = null;
            private IRhinoScript _RhinoScript = null;
            #endregion

            public RhinoScript(RhinoApplication rhinoapp)
            {
                //widen scope
                _RhinoApp = rhinoapp.RhinoApp;
                _RhinoScript = rhinoapp.RhinoScript;
            }

            /// <summary>
            /// Sends a Rhino Command
            /// </summary>
            /// <param name="str"></param>
            public void SendCommand(string str)
            {
                try
                {
                    _RhinoScript.Command(str);
                }
                catch { }
            }
        }

        public static string RhinoCommand(RhinoApplication RhinoApp, string Command, bool SendCommand)
        {
            try
            {
                if (SendCommand == true)
                {

                    RhinoScript m_RhinoScript = new RhinoScript(RhinoApp);
                    m_RhinoScript.SendCommand(Command);

                    return "Command Sent Success";
                }
                else { return "Set to true to send commands."; }

            }
            catch (Exception ex) { return ex.ToString(); }
        }

        class clsRhinoObject
        {
            public Object RhinoObject;
            public Rhino.DocObjects.Layer ObjectLayer;
            public string ObjectName;
            public System.Drawing.Color ObjectColor;

            public clsRhinoObject(Object _obj, Rhino.DocObjects.Layer _layer, string _name, System.Drawing.Color _color)
            {
                RhinoObject = _obj;
                ObjectLayer = _layer;
                ObjectName = _name;
                ObjectColor = _color;
            }
        }

        public static System.Drawing.Color Create_RhinoColor(int Red, int Green, int Blue)
        {
            try
            {
                return System.Drawing.Color.FromArgb(Red, Green, Blue);
            }
            catch { return System.Drawing.Color.FromArgb(0, 0, 0); }
        }

        public static Rhino.DocObjects.Layer Create_RhinoLayer(string LayerName, System.Drawing.Color LayerColor)
        {
            try
            {
                Rhino.DocObjects.Layer m_layer = new Rhino.DocObjects.Layer();
                m_layer.Name = LayerName;
                m_layer.Color = LayerColor;
                return m_layer;
            }
            catch { return null; }
        }

        public static Object Create_RhinoObject(Object RhinoGeometry, string ObjName, System.Drawing.Color ObjColor, Rhino.DocObjects.Layer ObjLayer)
        {
            try
            {
                clsRhinoObject m_rhinoobj = new clsRhinoObject(RhinoGeometry, ObjLayer, ObjName, ObjColor);
                return (Object)m_rhinoobj;
            }
            catch { return null; }
        }

        public class RhinoApplication
        {
            #region "Public Members"
            public IRhino5x64Application RhinoApp = null;
            public IRhinoScript RhinoScript = null;
            #endregion

            private string _progID = null;
            private bool _visible = true;

            public RhinoApplication(string id, bool IsVisible)
            {
                // widen scope
                _progID = id;
                _visible = IsVisible;

                //setup
                DoSetup();
            }

            /// <summary>
            /// Setup Rhino and RhinoScript
            /// </summary>
            private void DoSetup()
            {
                try
                {
                    // Rhino Program ID
                    string m_rhinoprogramID = _progID;
                    IRhino5x64Application m_RhinoApp = null;

                    // get Rhino application
                    Type type = Type.GetTypeFromProgID(m_rhinoprogramID);
                    dynamic rhino = Activator.CreateInstance(type); // Create Rhino instance
                    m_RhinoApp = rhino as IRhino5x64Application;
                    m_RhinoApp.Visible = Convert.ToInt16(_visible);  // 0 = hidden,  1 = visible

                    RhinoApp = m_RhinoApp;  // Rhino Application

                    IRhinoScript m_RhinoScript = null;
                    m_RhinoScript = m_RhinoApp.GetScriptObject() as IRhinoScript;
                    RhinoScript = m_RhinoScript;
                }
                catch { }
            }
            class clsRhinoObject
            {
                public Object RhinoObject;
                public Rhino.DocObjects.Layer ObjectLayer;
                public string ObjectName;
                public System.Drawing.Color ObjectColor;

                public clsRhinoObject(Object _obj, Rhino.DocObjects.Layer _layer, string _name, System.Drawing.Color _color)
                {
                    RhinoObject = _obj;
                    ObjectLayer = _layer;
                    ObjectName = _name;
                    ObjectColor = _color;
                }
            }

        }

        public static class SaveRhino3dmModel
        {
            /// <summary>
            /// Save Rhino Model
            /// </summary>
            /// <param name="FilePath">Location of the Rhino File</param>
            /// <param name="SaveBool">Toggle to Save the File</param>
            /// <param name="RhinoObjects">Rhino Geometry to Save</param>
            /// <param name="Units">Model units as string (ex. millimeters, centimeters, meters, feet, inches)</param>
            /// <returns name="RhinoModel">Rhino 3dm Model Object</returns>
            /// <search>case,rhino,3dm,rhynamo,save,model</search>
            public static File3dm Save_Rhino3dmModel(string FilePath, bool SaveBool, Object[] RhinoObjects, string Units)
            {
                try
                {
                    // Rhino model
                    File3dm m_model = new File3dm();

                    if (Units == "millimeters")
                    {
                        m_model.Settings.ModelUnitSystem = Rhino.UnitSystem.Millimeters;
                    }
                    else if (Units == "centimeters")
                    {
                        m_model.Settings.ModelUnitSystem = Rhino.UnitSystem.Centimeters;
                    }
                    else if (Units == "meters")
                    {
                        m_model.Settings.ModelUnitSystem = Rhino.UnitSystem.Meters;
                    }
                    else if (Units == "feet")
                    {
                        m_model.Settings.ModelUnitSystem = Rhino.UnitSystem.Feet;
                    }
                    else if (Units == "inches")
                    {
                        m_model.Settings.ModelUnitSystem = Rhino.UnitSystem.Inches;
                    }

                    // Create Default Layer
                    Rhino.DocObjects.Layer m_defaultlayer = new Rhino.DocObjects.Layer();
                    m_defaultlayer.Name = "Default";
                    m_defaultlayer.LayerIndex = 0;
                    m_defaultlayer.IsVisible = true;
                    m_defaultlayer.IsLocked = false;
                    m_defaultlayer.Color = System.Drawing.Color.FromArgb(0, 0, 0);

                    // Add default layer to model
                    m_model.Layers.Add(m_defaultlayer);
                    m_model.Polish();

                    //iterate through objects
                    for (int i = 0; i < RhinoObjects.Length; i++)
                    {
                        // Rhino object
                        clsRhinoObject m_rhinoobj = RhinoObjects[i] as clsRhinoObject;

                        // rhino object information
                        Object m_objgeo = m_rhinoobj.RhinoObject;
                        string m_objname = m_rhinoobj.ObjectName;
                        System.Drawing.Color m_objcolor = m_rhinoobj.ObjectColor;
                        Rhino.DocObjects.Layer m_objlayer = m_rhinoobj.ObjectLayer;
                        m_objlayer.LayerIndex = m_model.Layers.Count;

                        // object attributes
                        Rhino.DocObjects.ObjectAttributes objatt = new Rhino.DocObjects.ObjectAttributes();
                        objatt.Name = m_objname;
                        objatt.ObjectColor = m_objcolor;

                        if (m_objcolor != m_objlayer.Color)
                        {
                            objatt.ColorSource = Rhino.DocObjects.ObjectColorSource.ColorFromObject;
                        }
                        else
                        {
                            objatt.ColorSource = Rhino.DocObjects.ObjectColorSource.ColorFromLayer;
                        }

                        // setup layer
                        try
                        {
                            bool m_layerexists = false;
                            int m_layerindex = -1;
                            for (int j = 0; j < m_model.Layers.Count; j++)
                            {
                                Rhino.DocObjects.Layer l = m_model.Layers[j];
                                if (l.Name == m_objlayer.Name)
                                {
                                    m_layerexists = true;
                                    m_layerindex = j;
                                }
                            }

                            // if the layer exists
                            if (m_layerexists == true)
                            {
                                // set the object attribute layer index
                                objatt.LayerIndex = m_layerindex;
                            }
                            else
                            {
                                // add layer to the model and set the object attribute
                                m_model.Layers.Add(m_objlayer);
                                m_model.Polish();
                                objatt.LayerIndex = m_objlayer.LayerIndex;
                            }
                        }
                        catch { }

                        // check if point
                        if (m_objgeo is Rhino.Geometry.Point)
                        {
                            Rhino.Geometry.Point pt = (Rhino.Geometry.Point)m_objgeo;
                            double x = pt.Location.X;
                            double y = pt.Location.Y;
                            double z = pt.Location.Z;

                            Rhino.Geometry.Point3d pt3d = new Point3d(x, y, z);
                            m_model.Objects.AddPoint(pt3d, objatt);
                        }

                        // check if it is a curve
                        if (m_objgeo is Rhino.Geometry.Curve)
                        {
                            Rhino.Geometry.Curve cv = m_objgeo as Rhino.Geometry.Curve;
                            m_model.Objects.AddCurve(cv, objatt);
                        }

                        //check if surface
                        if (m_objgeo is Rhino.Geometry.NurbsSurface)
                        {
                            Rhino.Geometry.Surface srf = m_objgeo as Rhino.Geometry.Surface;
                            m_model.Objects.AddSurface(srf, objatt);
                        }

                        //check if brep
                        if (m_objgeo is Rhino.Geometry.Brep)
                        {
                            Rhino.Geometry.Brep brp = m_objgeo as Rhino.Geometry.Brep;
                            m_model.Objects.AddBrep(brp, objatt);
                        }

                        //check if mesh
                        if (m_objgeo is Rhino.Geometry.Mesh)
                        {
                            Rhino.Geometry.Mesh msh = m_objgeo as Rhino.Geometry.Mesh;
                            m_model.Objects.AddMesh(msh, objatt);
                        }
                    }

                    // write file
                    if (SaveBool == true)
                    {
                        try
                        {
                            // write model file
                            m_model.Write(FilePath, 1);
                        }
                        catch { }
                    }

                    return m_model;
                }
                catch { return null; }
            }
        }

        public string LoadFile(string filename)
        {
            string myDocuments = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
            string path = Path.Combine(myDocuments, filename);
            //Process.Start(path);


            return path;
        }

        public static List<List<Element>> GetMembersRecursive(Autodesk.Revit.DB.Document d, Group g, List<List<Element>> r, List<List<string>> strin_name, List<int> int_)
        {
            if (strin_name == null)
            {
                strin_name = new List<List<string>>();
            }
            if (r == null)
            {
                r = new List<List<Element>>();
            }
            if (int_ == null)
            {
                int_ = new List<int>();
            }
            List<string> lista_nombre_main = new List<string>();
            List<string> lista_nombre_buildings = new List<string>();
            List<string> lista_nombre_floor = new List<string>();
            List<Group> lista_de_buildings = new List<Group>();
            List<Group> lista_de_floor = new List<Group>();
            List<List<Element>> ele_list = new List<List<Element>>();
            List<Element> ceiling_groups1 = new List<Element>();

            List<Element> elems = g.GetMemberIds().Select(q => d.GetElement(q)).ToList();
            lista_nombre_main.Add(g.Name);
            lista_nombre_buildings.Add(g.Name);

            foreach (Element el in elems)
            {
                if (el.GetType() == typeof(Group))
                {
                    Group gp = el as Group;
                    lista_de_buildings.Add(gp);
                    lista_nombre_buildings.Add(el.Name);
                }
                if (el.GetType() == typeof(Ceiling))
                {
                    ceiling_groups1.Add(el);
                }
            }
            r.Add(ceiling_groups1);
            for (int i = 0; i < lista_de_buildings.ToArray().Length; i++)
            {
                Group gp2 = lista_de_buildings.ToArray()[i];

                List<Element> elems2 = gp2.GetMemberIds().Select(q => d.GetElement(q)).ToList();
                ele_list.Add(elems2);
            }

            for (int i = 0; i < ele_list.ToArray().Length; i++)
            {
                List<Element> lista1 = ele_list.ToArray()[i];
                List<Element> ceiling_groups = new List<Element>();
                foreach (var item in lista1)
                {
                    if (item.GetType() == typeof(Group))
                    {
                        Group gp4 = item as Group;
                        List<Element> elems3 = gp4.GetMemberIds().Select(q => d.GetElement(q)).ToList();
                        foreach (var item2 in elems3)
                        {
                            if (item2.GetType() == typeof(Group))
                            {
                                Group gp5 = item2 as Group;
                                List<Element> elems4 = gp5.GetMemberIds().Select(q => d.GetElement(q)).ToList();
                                foreach (var item3 in elems4)
                                {
                                    if (item3.GetType() == typeof(Ceiling))
                                    {
                                        ceiling_groups.Add(item3);
                                        lista_nombre_floor.Add(item.Name);
                                    }
                                }
                            }
                            if (item2.GetType() == typeof(Ceiling))
                            {
                                ceiling_groups.Add(item2);
                                lista_nombre_floor.Add(item.Name);
                            }
                        }

                        //lista_de_floor.Add(gp4);
                        //lista_nombre_floor.Add(item.Name);
                    }
                    if (item.GetType() == typeof(Ceiling))
                    {
                        ceiling_groups.Add(item);
                        lista_nombre_floor.Add(item.Name);

                    }
                }
                r.Add(ceiling_groups);
            }
            strin_name.Add(lista_nombre_main);
            strin_name.Add(lista_nombre_buildings);
            strin_name.Add(lista_nombre_floor);

            return r;

        }

        public static List<List<Face>> GetFaces(Autodesk.Revit.DB.Document doc_, List<List<Element>> list_elements, List<List<Face>> list_faces)
        {
            if (list_faces == null)
            {
                list_faces = new List<List<Face>>();
            }
            for (int i = 0; i < list_elements.ToArray().Length; i++)
            {
                List<Face> faces_list = new List<Face>();
                List<Element> ele = list_elements.ToArray()[i];

                foreach (var item in ele)
                {
                    Options op = new Options();
                    op.ComputeReferences = true;
                    foreach (var item2 in item.get_Geometry(op).Where(q => q is Solid).Cast<Solid>())
                    {
                        foreach (Face item3 in item2.Faces)
                        {
                            PlanarFace planarFace = item3 as PlanarFace;
                            XYZ normal = planarFace.ComputeNormal(new UV(planarFace.Origin.X, planarFace.Origin.Y));

                            if (normal.Z == 0 && normal.Y > -0.8 /*&& normal.X < 0*/)
                            {
                                Element e = doc_.GetElement(item3.Reference);
                                GeometryObject geoobj = e.GetGeometryObjectFromReference(item3.Reference);
                                Face face = geoobj as Face;
                                faces_list.Add(face);
                            }
                        }
                    }
                }
                if (faces_list.ToArray().Length > 0)
                {
                    list_faces.Add(faces_list);
                }

            }

            return list_faces;
        }

        public static List<List<XYZ>> GeTpoints(Autodesk.Revit.DB.Document doc_, List<List<XYZ>> xyz_faces, IList<CurveLoop> faceboundaries, List<List<Face>> list_faces)
        {
            if (list_faces == null)
            {
                list_faces = new List<List<Face>>();
            }

            for (int i = 0; i < list_faces.ToArray().Length; i++)
            {
                List<XYZ> puntos_ = new List<XYZ>();
                foreach (Face f in list_faces.ToArray()[i])
                {

                    faceboundaries = f.GetEdgesAsCurveLoops();//new trying to get the outline of the face instead of the edges
                    EdgeArrayArray edgeArrays = f.EdgeLoops;
                    foreach (CurveLoop edges in faceboundaries)
                    {
                        puntos_.Add(null);
                        foreach (Autodesk.Revit.DB.Curve edge in edges)
                        {
                            XYZ testPoint1 = edge.GetEndPoint(1);
                            double lenght = Math.Round(edge.ApproximateLength, 0);
                            double lenght_convert = UnitUtils.Convert(lenght, DisplayUnitType.DUT_DECIMAL_FEET, DisplayUnitType.DUT_MILLIMETERS);

                            double x = Math.Round(testPoint1.X, 0);
                            double y = Math.Round(testPoint1.Y, 0);
                            double z = Math.Round(testPoint1.Z, 0);

                            XYZ newpt = new XYZ(x, y, z);

                            if (!puntos_.Contains(testPoint1))
                            {
                                puntos_.Add(testPoint1);

                            }
                        }
                    }
                    int num = f.EdgeLoops.Size;

                }
                xyz_faces.Add(puntos_);
            }
            return xyz_faces;

        }

        static AddInId appId = new AddInId(new Guid("D031091D-29A4-4F70-8FE5-84FBD4ED0D73"));
        public Autodesk.Revit.UI.Result Execute(ExternalCommandData commandData, ref string message, ElementSet elementSet)
        {
            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Autodesk.Revit.DB.Document doc = uidoc.Document;

            try
            {
                string filename = @"T:\Transfer\lopez\Book1.xlsx";
                using (ExcelPackage package = new ExcelPackage(new FileInfo(filename)))
                {
                    ExcelWorksheet sheet = package.Workbook.Worksheets.ElementAt(0);

                    int column = 3;
                    int number = Convert.ToInt32(sheet.Cells[2, column].Value);
                    sheet.Cells[2, column].Value = (number + 1); ;
                    package.Save();
                }
            }
            catch (Exception)
            {
                MessageBox.Show("Excel file not found", "");
            }

            Form21 form3 = new Form21();
            form3.ShowDialog();

            


            if (form3.radioButton2.Checked)
            {
                Form22 form4 = new Form22();
                form4.ShowDialog();

                Form20 form2 = new Form20();

                form2.ShowDialog();
            }
            
            

            List<Object> objs = new List<Object>();
            List<List<XYZ>> xyz_faces = new List<List<XYZ>>();
            IList<Face> face_with_regions = new List<Face>();
            String info = "";
            List<List<Face>> Faces_lists_excel = new List<List<Face>>();
            //List<FaceArray> face112 = new List<FaceArray>();
            IList<CurveLoop> faceboundaries = new List<CurveLoop>();
            List<List<Element>> elemente_selected = new List<List<Element>>();
            List<List<string>> names = new List<List<string>>();
            List<int> numeros_ = new List<int>();
            XYZ pos_z = new XYZ(0, 0, 1);
            XYZ neg_z = new XYZ(0, 0, -1);

            Form7 form = new Form7();

            List<Face> faces_picked = new List<Face>();
            List<string> name_of_roof = new List<string>();

            //Group grp_Lot = doc.GetElement(uidoc.Selection.PickObject(ObjectType.Element, "Select an existing group")) as Group;
            ///this code was use to select ceilings and explore is facing direction so they can be reproduce in rhino geometry///
            try
            {
                ICollection<Reference> my_faces = uidoc.Selection.PickObjects(Autodesk.Revit.UI.Selection.ObjectType.Face, "Select ceilings to be reproduced in rhino geometry");
                foreach (var item_myRefWall in my_faces)
                {
                    Element e = doc.GetElement(item_myRefWall);
                    GeometryObject geoobj = e.GetGeometryObjectFromReference(item_myRefWall);
                    Face face = geoobj as Face;
                    PlanarFace planarFace = face as PlanarFace;
                    XYZ normal = planarFace.ComputeNormal(new UV(planarFace.Origin.X, planarFace.Origin.Y));


                    name_of_roof.Add("roof");
                    faces_picked.Add(face);


                    if (item_myRefWall == my_faces.ToArray().Last())
                    {
                        Faces_lists_excel.Add(faces_picked);

                        //names.Add(name_of_roof);
                    }

                }
            }
            catch (Exception)
            {

                MessageBox.Show("no surfaces were selected", "Warning");
            }
            
            
           

            
            Group grpExisting = doc.GetElement(uidoc.Selection.PickObject(ObjectType.Element, "Select an existing group")) as Group;
            GetMembersRecursive(doc, grpExisting, elemente_selected, names, numeros_);

            foreach (var item in elemente_selected)
            {
                foreach (var item2 in item)
                {
                    if (!form.listBox1.Items.Contains(item2.Name))
                    {
                        form.listBox1.Items.Add(item2.Name);
                    }
                }
            }

            Form23 form5 = new Form23();
            form5.ShowDialog();

            Form24 form6 = new Form24();
            form6.ShowDialog();
            

            form.ShowDialog();

            if (form.DialogResult == DialogResult.Cancel)
            {
                return Autodesk.Revit.UI.Result.Cancelled;
            }

            List<List<Element>> elemente_selected_to_bedeleted = new List<List<Element>>();
            foreach (var item in elemente_selected)
            {
                List<Element> ele_sel = new List<Element>();
                foreach (var item2 in item)
                {
                    foreach (var item3 in form.listBox2.Items)
                    {
                        if (item3.ToString() == item2.Name)
                        {
                            ele_sel.Add(item2);
                        }
                    }
                }
                elemente_selected_to_bedeleted.Add(ele_sel);
            }

            string name_of_group = names.ToArray()[0].ToArray()[0].ToString();
            names.ToArray()[1].Insert(1, "roof" + name_of_group);
            GetFaces(doc, elemente_selected_to_bedeleted, Faces_lists_excel);
            GeTpoints(doc, xyz_faces, faceboundaries, Faces_lists_excel);
            TaskDialog.Show("info faces", info);
            //string filename = Path.Combine(Path.GetTempPath(), "Book1.xlsx"); /// this line was used to automatically look for this excel file///

            string filename2 = "";
            System.Windows.Forms.OpenFileDialog openDialog = new System.Windows.Forms.OpenFileDialog();
            openDialog.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
            //openDialog.Filter = "Excel Files (*.xlsx) |*.xslx)"; // TODO: Change to .csv
            if (openDialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                filename2 = openDialog.FileName;
            }
            int numero = 0;
            using (ExcelPackage package = new ExcelPackage(new FileInfo(filename2)))
            {
                package.Workbook.Worksheets.Delete(1);
                ExcelWorksheet sheet = package.Workbook.Worksheets.Add("my_data");
                int row = 1;
                for (int i = 0; i < xyz_faces.ToArray().Length; i++)
                {
                    numero = 0;
                    foreach (var item in xyz_faces.ToArray()[i])
                    {
                        if (item == null)
                        {
                            numero += 1;
                            sheet.Cells[row, 1].Value = names.ToArray()[0][0];
                            sheet.Cells[row, 2].Value = names.ToArray()[1][i + 1];
                            sheet.Cells[row, 3].Value = numero;
                            row++;
                        }
                        else
                        {
                            sheet.Cells[row, 1].Value = Math.Round(item.X, 1);
                            sheet.Cells[row, 2].Value = Math.Round(item.Y, 1);
                            sheet.Cells[row, 3].Value = Math.Round(item.Z, 1);
                            row++;
                        }

                        if (item == xyz_faces.ToArray()[i].Last())
                        {
                            sheet.Cells[row, 1].Value = "Next";
                            sheet.Cells[row, 2].Value = ".";
                            sheet.Cells[row, 3].Value = ".";
                            row++;
                        }
                    }
                }
                package.Save();
            }
            Process.Start(filename2);
            return Autodesk.Revit.UI.Result.Succeeded;
        }
    }

    [Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)]
    public class Rhino_access_faces : IExternalCommand
    {
       
        public static List<List<Face>> GetFaces_individual(Autodesk.Revit.DB.Document doc_, List<List<Element>> list_elements, List<List<Face>> list_faces)
        {
            if (list_faces == null)
            {
                list_faces = new List<List<Face>>();
            }
            for (int i = 0; i < list_elements.ToArray().Length; i++)
            {
                List<Face> faces_list = new List<Face>();
                List<Element> ele = list_elements.ToArray()[i];

                foreach (var item in ele)
                {
                    Options op = new Options();
                    op.ComputeReferences = true;
                    foreach (var item2 in item.get_Geometry(op).Where(q => q is Solid).Cast<Solid>())
                    {
                        foreach (Face item3 in item2.Faces)
                        {
                            PlanarFace planarFace = item3 as PlanarFace;
                            XYZ normal = planarFace.ComputeNormal(new UV(planarFace.Origin.X, planarFace.Origin.Y));

                            Element e = doc_.GetElement(item3.Reference);
                            GeometryObject geoobj = e.GetGeometryObjectFromReference(item3.Reference);
                            Face face = geoobj as Face;
                            faces_list.Add(face);
                        }
                    }
                }
                if (faces_list.ToArray().Length > 0)
                {
                    list_faces.Add(faces_list);
                }

            }

            return list_faces;
        }
        public static List<List<XYZ>> GeTpoints(Autodesk.Revit.DB.Document doc_, List<List<XYZ>> xyz_faces, IList<CurveLoop> faceboundaries, List<List<Face>> list_faces)
        {
            if (list_faces == null)
            {
                list_faces = new List<List<Face>>();
            }

            for (int i = 0; i < list_faces.ToArray().Length; i++)
            {
                List<XYZ> puntos_ = new List<XYZ>();
                foreach (Face f in list_faces.ToArray()[i])
                {

                    faceboundaries = f.GetEdgesAsCurveLoops();//new trying to get the outline of the face instead of the edges
                    EdgeArrayArray edgeArrays = f.EdgeLoops;
                    foreach (CurveLoop edges in faceboundaries)
                    {
                        puntos_.Add(null);
                        foreach (Autodesk.Revit.DB.Curve edge in edges)
                        {
                            XYZ testPoint1 = edge.GetEndPoint(1);
                            XYZ testPoint2 = edge.GetEndPoint(0);
                            double lenght = Math.Round(edge.ApproximateLength, 0);
                            double lenght_convert = UnitUtils.Convert(lenght, DisplayUnitType.DUT_DECIMAL_FEET, DisplayUnitType.DUT_MILLIMETERS);

                            double x = Math.Round(testPoint1.X, 0);
                            double y = Math.Round(testPoint1.Y, 0);
                            double z = Math.Round(testPoint1.Z, 0);

                            ElementClassFilter filter = new ElementClassFilter(typeof(Floor));
                            XYZ dir = new XYZ(0, 0, 0) - testPoint1;

                            //ReferenceIntersector refIntersector = new ReferenceIntersector(filter, FindReferenceTarget.Face,doc_.ActiveView as View3D);
                            //ReferenceWithContext referenceWithContext = refIntersector.FindNearest(testPoint1 , dir);
                            
                            //Reference reference = referenceWithContext.GetReference();
                            //XYZ intersection = reference.GlobalPoint;
                            //XYZ newpt = new XYZ(x, y, z);

                            if (!puntos_.Contains(testPoint1))
                            {
                                puntos_.Add(testPoint1);
                            }
                        }
                    }
                    int num = f.EdgeLoops.Size;
                }
                xyz_faces.Add(puntos_);
            }
            return xyz_faces;
        }

        static AddInId appId = new AddInId(new Guid("D031092D-29A4-4F70-8FE5-84FBD4ED0D73"));
        public Autodesk.Revit.UI.Result Execute(ExternalCommandData commandData, ref string message, ElementSet elementSet)
        {
            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Autodesk.Revit.DB.Document doc = uidoc.Document;

            try
            {
                string filename = @"T:\Transfer\lopez\Book1.xlsx";
                using (ExcelPackage package = new ExcelPackage(new FileInfo(filename)))
                {
                    ExcelWorksheet sheet = package.Workbook.Worksheets.ElementAt(0);

                    int column = 4;
                    int number = Convert.ToInt32(sheet.Cells[2, column].Value);
                    sheet.Cells[2, column].Value = (number + 1); ;
                    package.Save();
                }
            }
            catch (Exception)
            {
                MessageBox.Show("Excel file not found", "");
            }

            //string comments = "Createsheet" + "_" + doc.Application.Username + "_" + doc.Title;
            //string filename = @"D:\Users\lopez\Desktop\Comments.txt";
            ////System.Diagnostics.Process.Start(filename);
            //StreamWriter writer = new StreamWriter(filename, true);
            ////writer.WriteLine( Environment.NewLine);
            //writer.WriteLine(DateTime.Now + " - " + comments);
            //writer.Close();

            List<Object> objs = new List<Object>();
            List<List<XYZ>> xyz_faces = new List<List<XYZ>>();
            IList<Face> face_with_regions = new List<Face>();
            List<List<Face>> Faces_lists_excel = new List<List<Face>>();
            //List<FaceArray> face112 = new List<FaceArray>();
            IList<CurveLoop> faceboundaries = new List<CurveLoop>();
            List<List<Element>> elemente_selected = new List<List<Element>>();
            List<Face> element = new List<Face>();
            List<Reference> my_faces = new List<Reference>();
            List<List<string>> names = new List<List<string>>();
            List<int> numeros_ = new List<int>();
            XYZ pos_z = new XYZ(0, 0, 1);
            XYZ neg_z = new XYZ(0, 0, -1);

            Filter_for_exporter form = new Filter_for_exporter();

            form.ShowDialog();

            if (form.DialogResult == DialogResult.Cancel)
            {
                return Autodesk.Revit.UI.Result.Cancelled;
            }

            if (form.checkBox1.Checked == false)
            {
                foreach (var item in new FilteredElementCollector(doc).OfClass(typeof(Ceiling)))
                {
                    if (item.Name.Contains(form.textBox1.Text))
                    {
                        Options op = new Options();
                        op.ComputeReferences = true;
                        foreach (var item2 in item.get_Geometry(op).Where(q => q is Solid).Cast<Solid>())
                        {
                            foreach (Face item3 in item2.Faces)
                            {
                                PlanarFace planarFace = item3 as PlanarFace;
                                XYZ normal = planarFace.ComputeNormal(new UV(planarFace.Origin.X, planarFace.Origin.Y));

                                Element e = doc.GetElement(item3.Reference);
                                GeometryObject geoobj = e.GetGeometryObjectFromReference(item3.Reference);
                                Face face = geoobj as Face;
                                element.Add(face);
                            }
                        }

                        //element.Clear();
                    }
                }
                Faces_lists_excel.Add(element);
            }
            else
            {
                //Group grp_Lot = doc.GetElement(uidoc.Selection.PickObject(ObjectType.Element, "Select an existing group")) as Group;
                ///this code was use to select ceilings and explore is facing direction so they can be reproduce in rhino geometry///
                ICollection<Reference> my_faces_ = uidoc.Selection.PickObjects(Autodesk.Revit.UI.Selection.ObjectType.Face, "Select ceilings to be reproduced in rhino geometry");

                foreach (var item in my_faces_)
                {
                    my_faces.Add(item);
                }
            }

            List<Face> faces_picked = new List<Face>();
            List<string> name_of_roof = new List<string>();

            foreach (var item_myRefWall in my_faces)
            {
                Element e = doc.GetElement(item_myRefWall);
                GeometryObject geoobj = e.GetGeometryObjectFromReference(item_myRefWall);
                Face face = geoobj as Face;
                PlanarFace planarFace = face as PlanarFace;
                XYZ normal = planarFace.ComputeNormal(new UV(planarFace.Origin.X, planarFace.Origin.Y));


                name_of_roof.Add("roof");
                faces_picked.Add(face);


                if (item_myRefWall == my_faces.ToArray().Last())
                {
                    Faces_lists_excel.Add(faces_picked);

                    //names.Add(name_of_roof);
                }

            }

            //names.ToArray()[1].Insert(1, "individual faces" );
            //GetFaces_individual(doc, elemente_selected, Faces_lists_excel);
            GeTpoints(doc, xyz_faces, faceboundaries, Faces_lists_excel);
            TaskDialog.Show("Excel writting", "Writting faces information in a temporary excel file");

            //string filename = Path.Combine(Path.GetTempPath(), "Book1.xlsx"); /// this line was used to automatically look for this excel file///

            string filename2 = "";
            System.Windows.Forms.OpenFileDialog openDialog = new System.Windows.Forms.OpenFileDialog();
            openDialog.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
            //openDialog.Filter = "Excel Files (*.xlsx) |*.xslx)"; // TODO: Change to .csv
            if (openDialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                filename2 = openDialog.FileName;
            }

            int numero = 0;
            using (ExcelPackage package = new ExcelPackage(new FileInfo(filename2)))
            {
                package.Workbook.Worksheets.Delete(1);
                ExcelWorksheet sheet = package.Workbook.Worksheets.Add("my_data");
                int row = 1;
                for (int i = 0; i < xyz_faces.ToArray().Length; i++)
                {
                    numero = 0;
                    
                    foreach (var item in xyz_faces.ToArray()[i])
                    {

                        if (item == null)
                        {
                            numero += 1;

                            sheet.Cells[row, 1].Value = form.textBox2.Text;
                            sheet.Cells[row, 2].Value = form.textBox2.Text;
                            sheet.Cells[row, 3].Value = ".";
                            row++;
                        }
                        else
                        {
                            sheet.Cells[row, 1].Value = Math.Round(item.X, 1);
                            sheet.Cells[row, 2].Value = Math.Round(item.Y, 1);
                            sheet.Cells[row, 3].Value = Math.Round(item.Z, 1);
                            row++;
                        }

                        if (item == xyz_faces.ToArray()[i].Last())
                        {
                            sheet.Cells[row, 1].Value = "Next";
                            sheet.Cells[row, 2].Value = ".";
                            sheet.Cells[row, 3].Value = ".";
                            row++;
                        }

                    }

                }

                package.Save();
            }

            Process.Start(filename2);

            return Autodesk.Revit.UI.Result.Succeeded;
        }
    }

   

    [Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)]
    public class Create_floor : IExternalCommand
    {
        static AddInId appId = new AddInId(new Guid("5F88CC78-A137-4809-AAF8-A478F3B24BAB"));
        public Autodesk.Revit.UI.Result Execute(ExternalCommandData commandData, ref string message, ElementSet elementSet)
        {
            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Autodesk.Revit.DB.Document doc = uidoc.Document;
            Autodesk.Revit.ApplicationServices.Application app = uiapp.Application;

            try
            {
                string filename = @"T:\Transfer\lopez\Book1.xlsx";
                using (ExcelPackage package = new ExcelPackage(new FileInfo(filename)))
                {
                    ExcelWorksheet sheet = package.Workbook.Worksheets.ElementAt(0);

                    int column = 8;
                    int number = Convert.ToInt32(sheet.Cells[2, column].Value);
                    sheet.Cells[2, column].Value = (number + 1); ;
                    package.Save();
                }
            }
            catch (Exception)
            {
                MessageBox.Show("Excel file not found", "");
            }

            //string comments = "Create_floor" + "_" + doc.Application.Username + "_" + doc.Title;
            //string filename = @"D:\Users\lopez\Desktop\Comments.txt";
            ////System.Diagnostics.Process.Start(filename);
            //StreamWriter writer = new StreamWriter(filename, true);
            ////writer.WriteLine( Environment.NewLine);
            //writer.WriteLine(DateTime.Now + " - " + comments);
            //writer.Close();


            int numero_de_punto = 0;
            int row_num = 0;

            List<List<XYZ>> pts_list = new List<List<XYZ>>();

            //string filename2 = Path.Combine(Path.GetTempPath(), "Book1.xlsx");
            string filename2 = "";
            System.Windows.Forms.OpenFileDialog openDialog = new System.Windows.Forms.OpenFileDialog();
            openDialog.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
            openDialog.Filter = "Excel Files (*.xlsx) |*.xslx)";

            if (openDialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                filename2 = openDialog.FileName;
            }
            else
                return Result.Failed;


            using (ExcelPackage package = new ExcelPackage(new FileInfo(filename2)))
            {
                List<List<int>> numrandom = new List<List<int>>();
                List<List<int>> num = new List<List<int>>();
                List<List<string>> nombres_ = new List<List<string>>();

                ExcelWorksheet sheet = package.Workbook.Worksheets[1];

                double x_ = 0;
                double y_ = 0;
                double z_ = 0;


                List<XYZ> pts = new List<XYZ>();

                for (int row = 1; row < 9999; row++)
                {
                    row_num++;

                    numero_de_punto++;
                    bool contains_letter = false;


                    var thisValue = sheet.Cells[row, 1].Value;

                    if (thisValue == null)
                    {
                        break;
                    }


                    for (int col = 1; col < 9999; col++)
                    {
                        thisValue = sheet.Cells[row, col].Value;



                        if (thisValue != null)
                        {
                            if (Char.IsLetter(thisValue.ToString().First()))
                            {
                                contains_letter = true;
                            }
                            else
                            {
                                contains_letter = false;
                            }
                        }

                        if (contains_letter == true)
                        {
                            List<XYZ> pt_2 = new List<XYZ>();
                            foreach (var item in pts)
                            {
                                pt_2.Add(item);
                            }
                            pts_list.Add(pt_2);
                            pts.Clear();
                            //row++;
                            break;

                        }

                        if (thisValue == null)
                        {
                            XYZ pt = new XYZ(x_, y_, z_);
                            pts.Add(pt);
                            //row++;
                            break;
                        }
                        else
                        {
                            if (col == 1)
                            {
                                x_ = Convert.ToDouble(thisValue);

                                //x_1 = x_1 * 304.8;
                            }
                            if (col == 2)
                            {
                                y_ = Convert.ToDouble(thisValue);
                                //y_1 = y_1 * 304.8;
                            }
                            if (col == 3)
                            {
                                z_ = Convert.ToDouble(thisValue);
                                //z_1 = z_1 * 304.8;
                            }
                        }
                    }

                }
            }

            using (Transaction t = new Transaction(doc, "Make Floor from Rhino"))
            {
                t.Start();

                for (int i = 0; i < pts_list.ToArray().Length; i++)
                {
                    CurveArray ca = new CurveArray();
                    XYZ prev = null;

                    foreach (var pt in pts_list[i])
                    {
                        if (prev == null)
                        {
                            prev = pts_list[i].Last();
                        }
                        Autodesk.Revit.DB.Line line = Autodesk.Revit.DB.Line.CreateBound(prev, pt);

                        ca.Append(line);
                        prev = pt;
                    }
                    doc.Create.NewFloor(ca, false);
                }
                t.Commit();

            }

           

            return Autodesk.Revit.UI.Result.Succeeded;
        }
    }

   
    [Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)]
    public class nada : IExternalCommand
    {
        static AddInId appId = new AddInId(new Guid("6C22CC72-A167-4819-AAF1-A178F6B44BAB"));
        public Autodesk.Revit.UI.Result Execute(ExternalCommandData commandData, ref string message, ElementSet elementSet)
        {

            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Autodesk.Revit.DB.Document doc = uidoc.Document;

            try
            {
                string filename = @"T:\Transfer\lopez\Book1.xlsx";
                using (ExcelPackage package = new ExcelPackage(new FileInfo(filename)))
                {
                    ExcelWorksheet sheet = package.Workbook.Worksheets.ElementAt(0);

                    int column = 13;
                    int number = Convert.ToInt32(sheet.Cells[2, column].Value);
                    sheet.Cells[2, column].Value = (number + 1); ;
                    package.Save();
                }
            }
            catch (Exception)
            {
                MessageBox.Show("Excel file not found", "");
            }

            string filename2 = "";
            System.Windows.Forms.OpenFileDialog openDialog = new System.Windows.Forms.OpenFileDialog();
            openDialog.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
            if (openDialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                filename2 = openDialog.FileName;
            }

            Point3d p1 = new Point3d(0, 0, 0);
            Rhino.Geometry.Point3d pt3d = new Point3d(10, 10, 0);

            File3dm m_modelfile = null;
            string m_name = null;
            string m_size = null;
            string m_created = null;
            string m_createdby = null;
            string m_edited = null;
            string m_editedby = null;
            string m_revision = null;
            string m_units = null;
            string m_notes = null;

            Object RhinoFile = filename2;
            if (RhinoFile is System.IO.FileInfo)
            {
                System.IO.FileInfo m_fileinfo = (System.IO.FileInfo)RhinoFile;
                m_modelfile = File3dm.Read(m_fileinfo.FullName);
                m_name = m_fileinfo.Name;
                m_size = m_fileinfo.Length.ToString();
                m_created = m_fileinfo.CreationTimeUtc.ToString();
                m_createdby = m_modelfile.CreatedBy;
                m_edited = m_fileinfo.LastWriteTimeUtc.ToString();
                m_editedby = m_modelfile.LastEditedBy;
                m_revision = m_modelfile.Revision.ToString();
                m_units = m_modelfile.Settings.ModelUnitSystem.ToString();
                m_notes = m_modelfile.Notes.Notes;
            }
            else if (RhinoFile is string)
            {
                System.IO.FileInfo m_fileinfo = new System.IO.FileInfo((string)RhinoFile);
                m_modelfile = File3dm.Read(m_fileinfo.FullName);
                m_name = m_fileinfo.Name;
                m_size = m_fileinfo.Length.ToString();
                m_created = m_fileinfo.CreationTimeUtc.ToString();
                m_createdby = m_modelfile.CreatedBy;
                m_edited = m_fileinfo.LastWriteTimeUtc.ToString();
                m_editedby = m_modelfile.LastEditedBy;
                m_revision = m_modelfile.Revision.ToString();
                m_units = m_modelfile.Settings.ModelUnitSystem.ToString();
                m_notes = m_modelfile.Notes.Notes;
            }

            File3dmLayerTable all_layers = m_modelfile.AllLayers;

            var objs_ = m_modelfile.Objects;

            

            List<Rhino.DocObjects.Layer> listadelayers = new List<Rhino.DocObjects.Layer>();
            List<Rhino.DocObjects.Layer> borrar = new List<Rhino.DocObjects.Layer>();
            List<List<Rhino.DocObjects.Layer>> multiplelistadelayers = new List<List<Rhino.DocObjects.Layer>>();

            List<XYZ> listPoitns = new List<XYZ>();

            string hola = "";

            foreach (var item in all_layers)
            {
                if (item.Name == "TopoPoints")
                {
                    File3dmObject[] m_objs = m_modelfile.Objects.FindByLayer(item.Name);
                    foreach (File3dmObject obj in m_objs)
                    {
                        GeometryBase geo = obj.Geometry;
                        if (geo is Rhino.Geometry.Point)
                        {
                            Rhino.Geometry.Point pt = geo as Rhino.Geometry.Point;


                            Point3d PLocation = pt.Location;

                            double x_end = PLocation.X / 304.8;
                            double y_end = PLocation.Y / 304.8;
                            double z_end = PLocation.Z / 304.8;



                            XYZ pt_end = new XYZ(x_end, y_end, z_end);

                            listPoitns.Add(pt_end);

                        }
                    }
                }

                //if (item.Name.Contains("Level"))
                //{
                //    if (listadelayers.Count != 0)
                //    {
                //        multiplelistadelayers.Add(listadelayers);
                //        listadelayers.RemoveRange(0, listadelayers.Count);
                //    }

                //    //listadelayers.Clear();
                //    hola = item.FullPath;

                //    continue;
                //}

                //if (item.FullPath.Contains(hola))
                //{
                //    listadelayers.Add(item);
                //}

                //foreach (var Layer in listadelayers)
                //{


                //    if (Layer.Name == "TopoPoints")
                //    {
                //        File3dmObject[] m_objs = m_modelfile.Objects.FindByLayer(Layer.Name);
                //        foreach (File3dmObject obj in m_objs)
                //        {
                //            GeometryBase geo = obj.Geometry;
                //            if (geo is Rhino.Geometry.Point)
                //            {
                //                Rhino.Geometry.Point pt = geo as Rhino.Geometry.Point;

                                
                //                Point3d PLocation = pt.Location;

                //                double x_end = PLocation.X / 304.8;
                //                double y_end = PLocation.Y / 304.8;
                //                double z_end = PLocation.Z / 304.8;

                                

                //                XYZ pt_end = new XYZ(x_end, y_end, z_end);

                //                listPoitns.Add(pt_end);
                                
                //            }
                //        }
                //    }
                //}
                
            }

            try
            {
                using (Transaction t = new Transaction(doc, "Make Floor from Rhino"))
                {
                    t.Start();

                    TopographySurface ts = TopographySurface.Create(doc, listPoitns);

                    t.Commit();
                }

            }
            catch (Exception)
            {
                //throw;
            }
            return Autodesk.Revit.UI.Result.Succeeded;
        }
    }

    [Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)]
    public class info : IExternalCommand
    {

        static AddInId appId = new AddInId(new Guid("7C82CC72-A167-4819-AAF1-A178F6B44BAB"));
        public Autodesk.Revit.UI.Result Execute(ExternalCommandData commandData, ref string message, ElementSet elementSet)
        {


            try
            {
                string filename = @"T:\Transfer\lopez\Book1.xlsx";
                using (ExcelPackage package = new ExcelPackage(new FileInfo(filename)))
                {
                    ExcelWorksheet sheet = package.Workbook.Worksheets.ElementAt(0);

                    int column = 14;
                    int number = Convert.ToInt32(sheet.Cells[2, column].Value);
                    sheet.Cells[2, column].Value = (number + 1); ;
                    package.Save();
                }
            }
            catch (Exception)
            {
                MessageBox.Show("Excel file not found", "");
            }


            Form16 form2 = new Form16();
            form2.ShowDialog();


            return Autodesk.Revit.UI.Result.Succeeded;
        }
    }

    [Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)]
    public class Reading_from_rhino : IExternalCommand
    {

        public static ModelLine Makeline(Autodesk.Revit.DB.Document doc, XYZ pta, XYZ ptb)
        {
            ModelLine modelLine = null;
            double distance = pta.DistanceTo(ptb);
            if (distance < 0.01)
            {
                TaskDialog.Show("Error", "Distance" + distance);
                return modelLine;
            }

            XYZ norm = pta.CrossProduct(ptb);
            if (norm.GetLength() == 0)
            {
                XYZ aSubB = pta.Subtract(ptb);
                XYZ aSubBcrossz = aSubB.CrossProduct(XYZ.BasisZ);
                double crosslenght = aSubBcrossz.GetLength();
                if (crosslenght == 0)
                {
                    norm = XYZ.BasisY;
                }
                else
                {
                    norm = XYZ.BasisZ;
                }
            }

            Autodesk.Revit.DB.Plane plane = Autodesk.Revit.DB.Plane.CreateByNormalAndOrigin(norm, ptb);


            SketchPlane skplane = SketchPlane.Create(doc, plane);

            Autodesk.Revit.DB.Line line = Autodesk.Revit.DB.Line.CreateBound(pta, ptb);

            if (doc.IsFamilyDocument)
            {
                modelLine = doc.FamilyCreate.NewModelCurve(line, skplane) as ModelLine;
            }
            else
            {
                modelLine = doc.Create.NewModelCurve(line, skplane) as ModelLine;
            }
            if (modelLine == null)
            {
                TaskDialog.Show("Error", "Model line = null");
            }
            return modelLine;
        }

        public static bool IsZero(double a)
        {
            const double _eps = 1.0e-9;
            return _eps > Math.Abs(a);
        }

        public static bool IsEqual(double a, double b)
        {
            return IsZero(b - a);
        }

        public static int Compare(double a, double b)
        {
            return IsEqual(a, b) ? 0 : (a < b ? -1 : 1);
        }

        public static int Compare(XYZ p, XYZ q)
        {
            
            int diff = Compare(p.X, q.X);
            if (0 == diff)
            {
                diff = Compare(p.Y, q.Y);
                if (0 == diff)
                {
                    diff = Compare(p.Z, q.Z);
                }

            }

            return diff;
        }

        private static Wall CreateWall(FamilyInstance cube, Autodesk.Revit.DB.Curve curve, double height)
        {
            var doc = cube.Document;

            var wallTypeId = doc.GetDefaultElementTypeId(
              ElementTypeGroup.WallType);

            return Wall.Create(doc, curve.CreateReversed(),
              wallTypeId, cube.LevelId, height, 0, false,
              false);
        }
        static AddInId appId = new AddInId(new Guid("5F88CC78-A137-4809-AAF8-A478F3B24BAB"));
        public Autodesk.Revit.UI.Result Execute(ExternalCommandData commandData, ref string message, ElementSet elementSet)
        {
            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Autodesk.Revit.DB.Document doc = uidoc.Document;

            try
            {
                string filename = @"T:\Transfer\lopez\Book1.xlsx";
                using (ExcelPackage package = new ExcelPackage(new FileInfo(filename)))
                {
                    ExcelWorksheet sheet = package.Workbook.Worksheets.ElementAt(0);

                    int column = 15;
                    int number = Convert.ToInt32(sheet.Cells[2, column].Value);
                    sheet.Cells[2, column].Value = (number + 1); ;
                    package.Save();
                }
            }



            catch (Exception)
            {
                MessageBox.Show("Excel file not found", "");
            }




            //string comments = "Create_floor_from_rhino" + "_" + doc.Application.Username + "_" + doc.Title;
            //string filename = @"D:\Users\lopez\Desktop\Comments.txt";
            //StreamWriter writer = new StreamWriter(filename, true);
            //writer.WriteLine(DateTime.Now + " - " + comments);
            //writer.Close();

            List<Object> objs = new List<Object>();
            List<List<XYZ>> xyz_faces = new List<List<XYZ>>();
            IList<Face> face_with_regions = new List<Face>();
            String info = "";
            List<List<Face>> Faces_lists_excel = new List<List<Face>>();
            //List<FaceArray> face112 = new List<FaceArray>();
            IList<CurveLoop> faceboundaries = new List<CurveLoop>();
            List<List<Element>> elemente_selected = new List<List<Element>>();
            List<List<string>> names = new List<List<string>>();
            List<int> numeros_ = new List<int>();
            XYZ pos_z = new XYZ(0, 0, 1);
            XYZ neg_z = new XYZ(0, 0, -1);



            string filename2 = "";
            System.Windows.Forms.OpenFileDialog openDialog = new System.Windows.Forms.OpenFileDialog();
            openDialog.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
            if (openDialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                filename2 = openDialog.FileName;
            }

            Point3d p1 = new Point3d(0, 0, 0);
            Rhino.Geometry.Point3d pt3d = new Point3d(10, 10, 0);

            File3dm m_modelfile = null;
            string m_name = null;
            string m_size = null;
            string m_created = null;
            string m_createdby = null;
            string m_edited = null;
            string m_editedby = null;
            string m_revision = null;
            string m_units = null;
            string m_notes = null;

            Object RhinoFile = filename2;
            if (RhinoFile is System.IO.FileInfo)
            {
                System.IO.FileInfo m_fileinfo = (System.IO.FileInfo)RhinoFile;
                m_modelfile = File3dm.Read(m_fileinfo.FullName);
                m_name = m_fileinfo.Name;
                m_size = m_fileinfo.Length.ToString();
                m_created = m_fileinfo.CreationTimeUtc.ToString();
                m_createdby = m_modelfile.CreatedBy;
                m_edited = m_fileinfo.LastWriteTimeUtc.ToString();
                m_editedby = m_modelfile.LastEditedBy;
                m_revision = m_modelfile.Revision.ToString();
                m_units = m_modelfile.Settings.ModelUnitSystem.ToString();
                m_notes = m_modelfile.Notes.Notes;
            }
            else if (RhinoFile is string)
            {
                System.IO.FileInfo m_fileinfo = new System.IO.FileInfo((string)RhinoFile);
                m_modelfile = File3dm.Read(m_fileinfo.FullName);
                m_name = m_fileinfo.Name;
                m_size = m_fileinfo.Length.ToString();
                m_created = m_fileinfo.CreationTimeUtc.ToString();
                m_createdby = m_modelfile.CreatedBy;
                m_edited = m_fileinfo.LastWriteTimeUtc.ToString();
                m_editedby = m_modelfile.LastEditedBy;
                m_revision = m_modelfile.Revision.ToString();
                m_units = m_modelfile.Settings.ModelUnitSystem.ToString();
                m_notes = m_modelfile.Notes.Notes;
            }

            File3dmLayerTable all_layers = m_modelfile.AllLayers;

            var objs_ = m_modelfile.Objects;

            //List<File3dmObject> m_objslist = new List<File3dmObject>();

            //foreach (Rhino.DocObjects.Layer lay in m_modelfile.AllLayers)
            //{
            //    File3dmObject[] m_objs = m_modelfile.Objects.FindByLayer(lay.Name);
            //    foreach (File3dmObject obj in m_objs)
            //    {
            //        m_objslist.Add(obj);
            //    }
            //}

            List<Rhino.DocObjects.Layer> listadelayers = new List<Rhino.DocObjects.Layer>();

            List<Rhino.DocObjects.Layer> borrar = new List<Rhino.DocObjects.Layer>();

            List<List<Rhino.DocObjects.Layer>> multiplelistadelayers = new List<List<Rhino.DocObjects.Layer>>();

            string hola = "";

            MessageBox.Show("This tool will read 3D information only if the following Rhino Layers exist; Levels, Grids, Structure, Floor, Walls, Points", "!");

            foreach (var item in all_layers)
            {
                 

                //if (item.Name.Contains( "Level"))
                //{
                //    if (listadelayers.Count != 0)
                //    {
                //        multiplelistadelayers.Add(listadelayers);
                //        listadelayers.RemoveRange(0, listadelayers.Count);
                //    }
                    
                //    //listadelayers.Clear();
                //    hola = item.FullPath;
                   
                //    continue;
                //}
                
                if (item.FullPath.Contains(hola))
                {
                    listadelayers.Add(item);
                }
            }

            List<Rhino.Geometry.Brep> rh_breps = new List<Rhino.Geometry.Brep>();
            List<Rhino.Geometry.Curve> curves_frames = new List<Rhino.Geometry.Curve>();
            List<Autodesk.Revit.DB.Curve> revit_crv = new List<Autodesk.Revit.DB.Curve>();
            List<string> m_names = new List<string>();
            List<int> m_layerindeces = new List<int>();
            List<System.Drawing.Color> m_colors = new List<System.Drawing.Color>();
            List<string> m_guids = new List<string>();

            foreach (var Layer in listadelayers)
            {
                if (Layer.Name == "Grids")
                {
                    File3dmObject[] m_objs = m_modelfile.Objects.FindByLayer(Layer.Name);
                    foreach (File3dmObject obj in m_objs)
                    {
                        GeometryBase geo = obj.Geometry;
                        if (geo is Rhino.Geometry.Curve)
                        {
                            Rhino.Geometry.Curve crv_ = geo as Rhino.Geometry.Curve;
                            curves_frames.Add(crv_);

                            Point3d end = crv_.PointAtEnd;
                            Point3d start = crv_.PointAtStart;

                            double x_end = end.X / 304.8;
                            double y_end = end.Y / 304.8;
                            double z_end = end.Z / 304.8;

                            double x_start = start.X / 304.8;
                            double y_start = start.Y / 304.8;
                            double z_start = start.Z / 304.8;

                            XYZ pt_end = new XYZ(x_end, y_end, z_end);
                            XYZ pt_start = new XYZ(x_start, y_start, z_start);

                            using (Transaction t = new Transaction(doc, "Make Floor from Rhino"))
                            {
                                t.Start();

                                Autodesk.Revit.DB.Curve curve1 = Autodesk.Revit.DB.Line.CreateBound(pt_end, pt_start);
                                Autodesk.Revit.DB.Line line = Autodesk.Revit.DB.Line.CreateBound(pt_end, pt_start);
                                Grid lineGrid = Grid.Create(doc, line);

                                revit_crv.Add(curve1);

                                t.Commit();
                            }

                            try
                            {
                                

                            }
                            catch (Exception)
                            {
                                //throw;
                            }
                        }
                    }
                }

                if (Layer.Name == "Levels")
                {
                    File3dmObject[] m_objs = m_modelfile.Objects.FindByLayer(Layer.Name);
                    foreach (File3dmObject obj in m_objs)
                    {
                        GeometryBase geo = obj.Geometry;
                        if (geo is Rhino.Geometry.Curve)
                        {
                            Rhino.Geometry.Curve crv_ = geo as Rhino.Geometry.Curve;
                            curves_frames.Add(crv_);

                            Point3d end = crv_.PointAtEnd;
                            Point3d start = crv_.PointAtStart;

                            double x_end = end.X / 304.8;
                            double y_end = end.Y / 304.8;
                            double z_end = end.Z / 304.8;

                            double x_start = start.X / 304.8;
                            double y_start = start.Y / 304.8;
                            double z_start = start.Z / 304.8;

                            XYZ pt_end = new XYZ(x_end, y_end, z_end);
                            XYZ pt_start = new XYZ(x_start, y_start, z_start);


                            using (Transaction t = new Transaction(doc, "Make ducts from Rhino"))
                            {
                                t.Start();

                                Autodesk.Revit.DB.Curve curve1 = Autodesk.Revit.DB.Line.CreateBound(pt_end, pt_start);
                                Autodesk.Revit.DB.Line line = Autodesk.Revit.DB.Line.CreateBound(pt_end, pt_start);
                                Level level_ = Level.Create(doc, z_end);
                                

                                revit_crv.Add(curve1);

                                t.Commit();
                            }

                            try
                            {
                                

                            }
                            catch (Exception)
                            {
                                //throw;
                            }
                        }
                    }
                }
                
                if (Layer.Name == "Floors")
                {
                    List<List<Rhino.Geometry.Curve>> floor_breps_Curve = new List<List<Rhino.Geometry.Curve>>();
                    List<List<Autodesk.Revit.DB.Curve>> wall_breps_Curve = new List<List<Autodesk.Revit.DB.Curve>>();
                    List<Rhino.Geometry.Surface> srflist = new List<Rhino.Geometry.Surface>();

                    File3dmObject[] m_objs = m_modelfile.Objects.FindByLayer(Layer.Name);
                    foreach (var obj in m_objs)
                    {


                        GeometryBase brep = obj.Geometry;

                        Rhino.Geometry.Brep brep_ = brep as Rhino.Geometry.Brep;
                        Rhino.Geometry.Extrusion ext = brep as Rhino.Geometry.Extrusion;
                        
                       
                        if (brep_ is Rhino.Geometry.Brep)
                        {
                            foreach (BrepFace bf in brep_.Faces)
                            {
                               
                                List<Rhino.Geometry.NurbsSurface> m_nurbslist = new List<Rhino.Geometry.NurbsSurface>();

                                Vector3d _normal = bf.NormalAt(0.5, 0.5);

                                if (_normal.Z == 1.0)
                                {
                                   
                                   

                                    Rhino.Geometry.NurbsSurface m_nurbsurface = bf.ToNurbsSurface();
                                    m_nurbslist.Add(m_nurbsurface);

                                    List<Rhino.Geometry.Curve> rh_loops = new List<Rhino.Geometry.Curve>();

                                    if (bf.Loops.Count > 0)
                                    {
                                        foreach (BrepLoop bloop in bf.Loops)
                                        {
                                            List<Rhino.Geometry.Curve> Curve_ = new List<Rhino.Geometry.Curve>();

                                            List<Rhino.Geometry.Curve> m_trims = new List<Rhino.Geometry.Curve>();
                                            foreach (BrepTrim t in bloop.Trims)
                                            {
                                                if (t.TrimType == BrepTrimType.Boundary || t.TrimType == BrepTrimType.Mated) // ignore "seams"
                                                {
                                                    Rhino.Geometry.Curve m_edgecurve = t.Edge.DuplicateCurve();
                                                    Curve_.Add(m_edgecurve);
                                                }
                                            }
                                            floor_breps_Curve.Add(Curve_);
                                        }
                                    }
                                }

                            }
                        }
                    }

                    List<List<XYZ>> floor_listas_pts = new List<List<XYZ>>();
                    if (floor_breps_Curve.ToArray().Length > 0)
                    {
                        for (int i = 0; i < floor_breps_Curve.ToArray().Length; i++)
                        {
                            List<XYZ> lista_pt = new List<XYZ>();
                            foreach (var item in floor_breps_Curve[i])
                            {
                                Point3d pt = item.PointAtStart;
                                double x = pt.X / 304.8;
                                double y = pt.Y / 304.8;
                                double z = pt.Z / 304.8;

                                XYZ pt_ = new XYZ(x, y, z);
                                lista_pt.Add(pt_);
                            }
                            floor_listas_pts.Add(lista_pt);
                        }
                    }

                    if (floor_breps_Curve.ToArray().Length > 0)
                    {
                        using (Transaction t = new Transaction(doc, "Make Floor from Rhino"))
                        {
                            t.Start();

                            for (int i = 0; i < floor_listas_pts.ToArray().Length; i++)
                            {
                                CurveArray ca = new CurveArray();
                                XYZ prev = null;

                                foreach (var pt in floor_listas_pts[i])
                                {
                                    if (prev == null)
                                    {
                                        prev = floor_listas_pts[i].Last();
                                    }
                                    Autodesk.Revit.DB.Line line = Autodesk.Revit.DB.Line.CreateBound(prev, pt);
                                    ca.Append(line);
                                    prev = pt;
                                    try
                                    {
                                      
                                    }
                                    catch (Exception)
                                    {

                                        //throw;
                                    }
                                }
                                doc.Create.NewFloor(ca, false);
                            }
                            t.Commit();
                        }
                    }
                }

                if (Layer.Name == "Walls")
                {
                    File3dmObject[] m_objs = m_modelfile.Objects.FindByLayer(Layer.Name);
                    foreach (var obj in m_objs)
                    {
                        GeometryBase brep = obj.Geometry;
                        Rhino.Geometry.Brep brep_ = brep as Rhino.Geometry.Brep;
                        Rhino.Geometry.Extrusion ext = brep as Rhino.Geometry.Extrusion;
                        
                        if (brep_ is Rhino.Geometry.Brep)
                        {
                            foreach (BrepFace bf in brep_.Faces)
                            {
                                List<Rhino.Geometry.NurbsSurface> m_nurbslist = new List<Rhino.Geometry.NurbsSurface>();
                                Vector3d _normal = bf.NormalAt(0.5, 0.5);
                                if (_normal.Z == 0.0)
                                {
                                    Rhino.Geometry.NurbsSurface m_nurbsurface = bf.ToNurbsSurface();
                                    m_nurbslist.Add(m_nurbsurface);
                                    if (bf.Loops.Count > 0)
                                    {
                                        foreach (BrepLoop bloop in bf.Loops)
                                        {
                                            List<Autodesk.Revit.DB.Curve> Curve_ = new List<Autodesk.Revit.DB.Curve>();
                                            foreach (BrepTrim t in bloop.Trims)
                                            {
                                                if (t.TrimType == BrepTrimType.Boundary || t.TrimType == BrepTrimType.Mated) // ignore "seams"
                                                {
                                                    Rhino.Geometry.Curve m_edgecurve = t.Edge.DuplicateCurve();
                                                    Point3d end = m_edgecurve.PointAtEnd;
                                                    Point3d start = m_edgecurve.PointAtStart;

                                                    double x_end = end.X / 304.8;
                                                    double y_end = end.Y / 304.8;
                                                    double z_end = end.Z / 304.8;

                                                    double x_start = start.X / 304.8;
                                                    double y_start = start.Y / 304.8;
                                                    double z_start = start.Z / 304.8;

                                                    XYZ pt_end = new XYZ(x_end, y_end, z_end);
                                                    XYZ pt_start = new XYZ(x_start, y_start, z_start);

                                                    try
                                                    {
                                                        Autodesk.Revit.DB.Curve curve1 = Autodesk.Revit.DB.Line.CreateBound(pt_end, pt_start);

                                                        Curve_.Add(curve1);
                                                    }
                                                    catch (Exception)
                                                    {
                                                    }
                                                }
                                            }

                                            using (Transaction t = new Transaction(doc, "Make Floor from Rhino"))
                                            {
                                                t.Start();
                                                Autodesk.Revit.DB.Wall.Create (doc, Curve_, true);
                                                t.Commit();
                                            }
                                        }
                                    }
                                }
                            }
                        }
                    }   
                }
            }
            return Autodesk.Revit.UI.Result.Succeeded;
        }  
    }

    [Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)]
    public class _lock : IExternalCommand
    {

        static AddInId appId = new AddInId(new Guid("5F88CC09-A137-6509-AAF8-A478F3B24BAB"));
        public Autodesk.Revit.UI.Result Execute(ExternalCommandData commandData, ref string message, ElementSet elementSet)
        {
            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Autodesk.Revit.DB.Document doc = uidoc.Document;

            try
            {
                string filename = @"T:\Transfer\lopez\Book1.xlsx";
                using (ExcelPackage package = new ExcelPackage(new FileInfo(filename)))
                {
                    ExcelWorksheet sheet = package.Workbook.Worksheets.ElementAt(0);

                    int column = 16;
                    int number = Convert.ToInt32(sheet.Cells[2, column].Value);
                    sheet.Cells[2, column].Value = (number + 1); ;
                    package.Save();
                }
            }
            catch (Exception)
            {
                MessageBox.Show("Excel file not found", "");
            }

            using (Transaction t = new Transaction(doc, "fghfgh"))
            {
                t.Start();

                try
                {

                    adWin.RibbonControl ribbon = adWin.ComponentManager.Ribbon;

                    foreach (adWin.RibbonTab tab in ribbon.Tabs)
                    {
                        string name = tab.AutomationName;

                        if (name == "Insert")
                        {
                            foreach (adWin.RibbonPanel panel in tab.Panels)
                            {
                                adWin.RibbonItemCollection items = panel.Source.Items;
                                foreach (var item in items)
                                {
                                    string name_ = item.Id;
                                    if (name_ == "ID_FILE_IMPORT")
                                    {
                                        item.IsEnabled = false;
                                    }
                                }
                            }
                        }

                    }
                    //adWin.ComponentManager.UIElementActivated += new
                    //  EventHandler<adWin.UIElementActivatedEventArgs>(
                    //    ComponentManager_UIElementActivated);
                }
                catch (Exception ex)
                {
                    //winform.MessageBox.Show(
                    //  ex.StackTrace + "\r\n" + ex.InnerException,
                    //  "Error", winform.MessageBoxButtons.OK);

                    //return Result.Failed;
                }

                t.Commit();

            }



            return Autodesk.Revit.UI.Result.Succeeded;
        }
    }

    [Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)]
    public class unlock : IExternalCommand
    {


        static AddInId appId = new AddInId(new Guid("5F88CC78-A137-6509-AAF8-A478F3B24BAB"));
        public Autodesk.Revit.UI.Result Execute(ExternalCommandData commandData, ref string message, ElementSet elementSet)
        {
            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Autodesk.Revit.DB.Document doc = uidoc.Document;

            try
            {
                string filename = @"T:\Transfer\lopez\Book1.xlsx";
                using (ExcelPackage package = new ExcelPackage(new FileInfo(filename)))
                {
                    ExcelWorksheet sheet = package.Workbook.Worksheets.ElementAt(0);

                    int column = 17;
                    int number = Convert.ToInt32(sheet.Cells[2, column].Value);
                    sheet.Cells[2, column].Value = (number + 1); ;
                    package.Save();
                }
            }
            catch (Exception)
            {
                MessageBox.Show("Excel file not found", "");
            }


            Warning form = new Warning();

            form.ShowDialog();
            String code = "0000";

            if (form.DialogResult == DialogResult.Cancel)
            {
                return Autodesk.Revit.UI.Result.Cancelled;
            }

            if (form.textBox1.Text == code)
            {
                using (Transaction t = new Transaction(doc, "fghfgh"))
                {
                    t.Start();

                    try
                    {

                        adWin.RibbonControl ribbon = adWin.ComponentManager.Ribbon;

                        foreach (adWin.RibbonTab tab in ribbon.Tabs)
                        {
                            string name = tab.AutomationName;

                            if (name == "Insert")
                            {
                                foreach (adWin.RibbonPanel panel in tab.Panels)
                                {
                                    adWin.RibbonItemCollection items = panel.Source.Items;
                                    foreach (var item in items)
                                    {
                                        string name_ = item.Id;
                                        if (name_ == "ID_FILE_IMPORT")
                                        {
                                            item.IsEnabled = true;
                                        }
                                    }
                                }
                            }

                        }
                        //adWin.ComponentManager.UIElementActivated += new
                        //  EventHandler<adWin.UIElementActivatedEventArgs>(
                        //    ComponentManager_UIElementActivated);
                    }
                    catch (Exception ex)
                    {
                        //winform.MessageBox.Show(
                        //  ex.StackTrace + "\r\n" + ex.InnerException,
                        //  "Error", winform.MessageBoxButtons.OK);

                        //return Result.Failed;
                    }

                    t.Commit();

                }
                return Autodesk.Revit.UI.Result.Succeeded;
            }
            else
            {
                return Autodesk.Revit.UI.Result.Cancelled;





            }
        }
    }

    [Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)]
    public class writting_to_rhino : IExternalCommand
    {
        public static RhinoApplication RhinoApp(bool StartRhino) // bool IsVisible)
        {
            try
            {
                if (StartRhino == true)
                {
                    // CASE Rhino Connection Class
                    RhinoCon.RhinoApplication m_RhinoCOM = new RhinoCon.RhinoApplication("Rhino5x64.Application", true);


                    return m_RhinoCOM;
                }
                else { return null; }
            }
            catch { return null; }
        }

        public static List<List<Element>> GetMembersRecursive(Autodesk.Revit.DB.Document d, Group g, List<List<Element>> r, List<List<string>> strin_name, List<int> int_)
        {
            if (strin_name == null)
            {
                strin_name = new List<List<string>>();
            }
            if (r == null)
            {
                r = new List<List<Element>>();
            }
            if (int_ == null)
            {
                int_ = new List<int>();
            }
            List<string> lista_nombre_main = new List<string>();
            List<string> lista_nombre_buildings = new List<string>();
            List<string> lista_nombre_floor = new List<string>();
            List<Group> lista_de_buildings = new List<Group>();
            List<Group> lista_de_floor = new List<Group>();
            List<List<Element>> ele_list = new List<List<Element>>();
            List<Element> ceiling_groups1 = new List<Element>();

            List<Element> elems = g.GetMemberIds().Select(q => d.GetElement(q)).ToList();
            lista_nombre_main.Add(g.Name);
            lista_nombre_buildings.Add(g.Name);

            foreach (Element el in elems)
            {
                if (el.GetType() == typeof(Group))
                {
                    Group gp = el as Group;
                    lista_de_buildings.Add(gp);
                    lista_nombre_buildings.Add(el.Name);
                }
                if (el.GetType() == typeof(Ceiling))
                {
                    ceiling_groups1.Add(el);
                }
            }
            r.Add(ceiling_groups1);
            for (int i = 0; i < lista_de_buildings.ToArray().Length; i++)
            {
                Group gp2 = lista_de_buildings.ToArray()[i];

                List<Element> elems2 = gp2.GetMemberIds().Select(q => d.GetElement(q)).ToList();
                ele_list.Add(elems2);
            }

            for (int i = 0; i < ele_list.ToArray().Length; i++)
            {
                List<Element> lista1 = ele_list.ToArray()[i];
                List<Element> ceiling_groups = new List<Element>();
                foreach (var item in lista1)
                {
                    if (item.GetType() == typeof(Group))
                    {
                        Group gp4 = item as Group;
                        List<Element> elems3 = gp4.GetMemberIds().Select(q => d.GetElement(q)).ToList();
                        foreach (var item2 in elems3)
                        {
                            if (item2.GetType() == typeof(Group))
                            {
                                Group gp5 = item2 as Group;
                                List<Element> elems4 = gp5.GetMemberIds().Select(q => d.GetElement(q)).ToList();
                                foreach (var item3 in elems4)
                                {
                                    if (item3.GetType() == typeof(Ceiling))
                                    {
                                        ceiling_groups.Add(item3);
                                        lista_nombre_floor.Add(item.Name);
                                    }
                                }
                            }
                            if (item2.GetType() == typeof(Ceiling))
                            {
                                ceiling_groups.Add(item2);
                                lista_nombre_floor.Add(item.Name);
                            }
                        }

                        //lista_de_floor.Add(gp4);
                        //lista_nombre_floor.Add(item.Name);
                    }
                    if (item.GetType() == typeof(Ceiling))
                    {
                        ceiling_groups.Add(item);
                        lista_nombre_floor.Add(item.Name);

                    }
                }
                r.Add(ceiling_groups);
            }
            strin_name.Add(lista_nombre_main);
            strin_name.Add(lista_nombre_buildings);
            strin_name.Add(lista_nombre_floor);

            return r;

        }

        public static List<List<Face>> GetFaces(Autodesk.Revit.DB.Document doc_, List<List<Element>> list_elements, List<List<Face>> list_faces)
        {
            if (list_faces == null)
            {
                list_faces = new List<List<Face>>();
            }
            for (int i = 0; i < list_elements.ToArray().Length; i++)
            {
                List<Face> faces_list = new List<Face>();
                List<Element> ele = list_elements.ToArray()[i];

                foreach (var item in ele)
                {
                    Options op = new Options();
                    op.ComputeReferences = true;
                    foreach (var item2 in item.get_Geometry(op).Where(q => q is Solid).Cast<Solid>())
                    {
                        foreach (Face item3 in item2.Faces)
                        {
                            PlanarFace planarFace = item3 as PlanarFace;
                            XYZ normal = planarFace.ComputeNormal(new UV(planarFace.Origin.X, planarFace.Origin.Y));

                            if (normal.Z == 0 && normal.Y > -0.8 /*&& normal.X < 0*/)
                            {
                                Element e = doc_.GetElement(item3.Reference);
                                GeometryObject geoobj = e.GetGeometryObjectFromReference(item3.Reference);
                                Face face = geoobj as Face;
                                faces_list.Add(face);
                            }
                        }
                    }
                }
                if (faces_list.ToArray().Length > 0)
                {
                    list_faces.Add(faces_list);
                }

            }

            return list_faces;
        }

        public static List<List<XYZ>> GeTpoints(Autodesk.Revit.DB.Document doc_, List<List<XYZ>> xyz_faces, IList<CurveLoop> faceboundaries, List<List<Face>> list_faces)
        {
            if (list_faces == null)
            {
                list_faces = new List<List<Face>>();
            }

            for (int i = 0; i < list_faces.ToArray().Length; i++)
            {
                List<XYZ> puntos_ = new List<XYZ>();
                foreach (Face f in list_faces.ToArray()[i])
                {

                    faceboundaries = f.GetEdgesAsCurveLoops();//new trying to get the outline of the face instead of the edges
                    EdgeArrayArray edgeArrays = f.EdgeLoops;
                    foreach (CurveLoop edges in faceboundaries)
                    {
                        puntos_.Add(null);
                        foreach (Autodesk.Revit.DB.Curve edge in edges)
                        {
                            XYZ testPoint1 = edge.GetEndPoint(1);
                            double lenght = Math.Round(edge.ApproximateLength, 0);
                            double lenght_convert = UnitUtils.Convert(lenght, DisplayUnitType.DUT_DECIMAL_FEET, DisplayUnitType.DUT_MILLIMETERS);

                            double x = Math.Round(testPoint1.X, 0);
                            double y = Math.Round(testPoint1.Y, 0);
                            double z = Math.Round(testPoint1.Z, 0);

                            XYZ newpt = new XYZ(x, y, z);

                            if (!puntos_.Contains(testPoint1))
                            {
                                puntos_.Add(testPoint1);

                            }
                        }
                    }
                    int num = f.EdgeLoops.Size;

                }
                xyz_faces.Add(puntos_);
            }
            return xyz_faces;

        }

        static AddInId appId = new AddInId(new Guid("5F88CC09-A137-6509-AAF8-A478F3B24BAB"));
        public Autodesk.Revit.UI.Result Execute(ExternalCommandData commandData, ref string message, ElementSet elementSet)
        {
            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Autodesk.Revit.DB.Document doc = uidoc.Document;

            try
            {
                string filename = @"T:\Transfer\lopez\Book1.xlsx";
                using (ExcelPackage package = new ExcelPackage(new FileInfo(filename)))
                {
                    ExcelWorksheet sheet = package.Workbook.Worksheets.ElementAt(0);

                    int column = 19;
                    int number = Convert.ToInt32(sheet.Cells[2, column].Value);
                    sheet.Cells[2, column].Value = (number + 1); ;
                    package.Save();
                }
            }
            catch (Exception)
            {
                MessageBox.Show("Excel file not found", "");
            }

            //string comments = "Rhino_access" + "_" + doc.Application.Username + "_" + doc.Title;
            //string filename = @"D:\Users\lopez\Desktop\Comments.txt";
            ////System.Diagnostics.Process.Start(filename);
            //StreamWriter writer = new StreamWriter(filename, true);
            ////writer.WriteLine( Environment.NewLine);
            //writer.WriteLine(DateTime.Now + " - " + comments);
            //writer.Close();

            List<Object> objs = new List<Object>();
            List<List<XYZ>> xyz_faces = new List<List<XYZ>>();
            IList<Face> face_with_regions = new List<Face>();
            String info = "";
            List<List<Face>> Faces_lists_excel = new List<List<Face>>();
            //List<FaceArray> face112 = new List<FaceArray>();
            IList<CurveLoop> faceboundaries = new List<CurveLoop>();
            List<List<Element>> elemente_selected = new List<List<Element>>();
            List<List<string>> names = new List<List<string>>();
            List<int> numeros_ = new List<int>();
            XYZ pos_z = new XYZ(0, 0, 1);
            XYZ neg_z = new XYZ(0, 0, -1);

            Form7 form = new Form7();

            string filename2 = "";
            System.Windows.Forms.OpenFileDialog openDialog = new System.Windows.Forms.OpenFileDialog();
            openDialog.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
            //openDialog.Filter = "Rhino Files (*.3dm) |*.3dm)"; // TODO: Change to .csv
            if (openDialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                filename2 = openDialog.FileName;
            }
            // this is use to open a new instance of rhino
            //RhinoCon.RhinoApplication m_RhinoCOM = RhinoApp(true);
            Point3d p1 = new Point3d(0, 0, 0);
            Rhino.Geometry.Point3d pt3d = new Point3d(10, 10, 0);
            //RhinoCon.RhinoApplication m_RhinoCOM = new RhinoCon.RhinoApplication("Rhino5x64.Application", true);
            //File3dm m_modelfile2 = m_RhinoCOM.RhinoScript.PointAdd(p1) /*as Rhino.FileIO.File3dm*/;
            File3dm m_modelfile = null;
            string m_name = null;
            string m_size = null;
            string m_created = null;
            string m_createdby = null;
            string m_edited = null;
            string m_editedby = null;
            string m_revision = null;
            string m_units = null;
            string m_notes = null;

            Object RhinoFile = filename2;
            if (RhinoFile is System.IO.FileInfo)
            {
                System.IO.FileInfo m_fileinfo = (System.IO.FileInfo)RhinoFile;
                m_modelfile = File3dm.Read(m_fileinfo.FullName);
                m_name = m_fileinfo.Name;
                m_size = m_fileinfo.Length.ToString();
                m_created = m_fileinfo.CreationTimeUtc.ToString();
                m_createdby = m_modelfile.CreatedBy;
                m_edited = m_fileinfo.LastWriteTimeUtc.ToString();
                m_editedby = m_modelfile.LastEditedBy;
                m_revision = m_modelfile.Revision.ToString();
                m_units = m_modelfile.Settings.ModelUnitSystem.ToString();
                m_notes = m_modelfile.Notes.Notes;
            }
            else if (RhinoFile is string)
            {
                System.IO.FileInfo m_fileinfo = new System.IO.FileInfo((string)RhinoFile);
                m_modelfile = File3dm.Read(m_fileinfo.FullName);
                m_name = m_fileinfo.Name;
                m_size = m_fileinfo.Length.ToString();
                m_created = m_fileinfo.CreationTimeUtc.ToString();
                m_createdby = m_modelfile.CreatedBy;
                m_edited = m_fileinfo.LastWriteTimeUtc.ToString();
                m_editedby = m_modelfile.LastEditedBy;
                m_revision = m_modelfile.Revision.ToString();
                m_units = m_modelfile.Settings.ModelUnitSystem.ToString();
                m_notes = m_modelfile.Notes.Notes;
            }

            Rhino.FileIO.File3dm df = m_modelfile as File3dm;

            RhinoCon.RhinoApplication m_RhinoCOM = new RhinoCon.RhinoApplication(filename2/*"Rhino5x64.Application"*/, true);

            

            //for (int j = 0; j <= 20; j++)

            //{

            //    m_modelfile.Objects.AddLine(new Rhino.Geometry.Line(j, 0, 5 - j, 5 + j, 0, j));

            //}
            //m_modelfile.Objects.AddPoint(p1);




            ICollection<Reference> my_faces = uidoc.Selection.PickObjects(Autodesk.Revit.UI.Selection.ObjectType.Face, "Select ceilings to be reproduced in rhino geometry");



            List<Face> faces_picked = new List<Face>();
            List<string> name_of_roof = new List<string>();

            foreach (var item_myRefWall in my_faces)
            {

                Element e = doc.GetElement(item_myRefWall);
                GeometryObject geoobj = e.GetGeometryObjectFromReference(item_myRefWall);
                Face face = geoobj as Face;
                PlanarFace planarFace = face as PlanarFace;
                XYZ normal = planarFace.ComputeNormal(new UV(planarFace.Origin.X, planarFace.Origin.Y));


                name_of_roof.Add("roof");
                faces_picked.Add(face);


                if (item_myRefWall == my_faces.ToArray().Last())
                {
                    Faces_lists_excel.Add(faces_picked);

                    //names.Add(name_of_roof);
                }

            }

            Group grpExisting = doc.GetElement(uidoc.Selection.PickObject(ObjectType.Element, "Select an existing group")) as Group;



            GetMembersRecursive(doc, grpExisting, elemente_selected, names, numeros_);

            foreach (var item in elemente_selected)
            {
                foreach (var item2 in item)
                {
                    if (!form.listBox1.Items.Contains(item2.Name))
                    {
                        form.listBox1.Items.Add(item2.Name);
                    }
                }
            }

            //form.ShowDialog();

            if (form.DialogResult == DialogResult.Cancel)
            {
                return Autodesk.Revit.UI.Result.Cancelled;
            }

            List<List<Element>> elemente_selected_to_bedeleted = new List<List<Element>>();
            List<List<string>> nombres_familias = new List<List<string>>();

            foreach (var item in elemente_selected)
            {
                List<Element> ele_sel = new List<Element>();
                foreach (var item2 in item)
                {
                    foreach (var item3 in form.listBox2.Items)
                    {
                        if (item3.ToString() == item2.Name)
                        {
                            ele_sel.Add(item2);
                        }
                    }
                }
                elemente_selected_to_bedeleted.Add(ele_sel);
            }



            string name_of_group = names.ToArray()[0].ToArray()[0].ToString();

            names.ToArray()[1].Insert(1, "roof" + name_of_group);

            GetFaces(doc, /*elemente_selected_to_bedeleted*/elemente_selected, Faces_lists_excel);

            GeTpoints(doc, xyz_faces, faceboundaries, Faces_lists_excel);

            List<List<Point3d>> listas_pts = new List<List<Point3d>>();
            ;
            List<Rhino.DocObjects.Layer> child_layer = new List<Rhino.DocObjects.Layer>();


            IList<Rhino.DocObjects.Layer> EXISTING = m_modelfile.AllLayers;
            int count_ = 0;

            foreach (var item in EXISTING)
            {
                count_++;
            }

            Rhino.DocObjects.Layer Parentlayer = new Rhino.DocObjects.Layer();
            Parentlayer.Name = names.ToArray()[0].ToArray()[0].ToString();
            Parentlayer.Index = count_ + 1;
            m_modelfile.AllLayers.Add(Parentlayer);


            List<Rhino.DocObjects.Layer> created_layers = new List<Rhino.DocObjects.Layer>();

            foreach (var item in names.ToArray()[1])
            {
                if (Parentlayer.Name != item)
                {
                    Rhino.DocObjects.Layer Childlayer = new Rhino.DocObjects.Layer();
                    //Childlayer.ParentLayerId = Parentlayer.Id;
                    Childlayer.Name = item.ToString();
                    Childlayer.Color = System.Drawing.Color.Bisque;
                    Childlayer.Index = Parentlayer.Index;



                    m_modelfile.AllLayers.Add(Childlayer);

                    created_layers.Add(Childlayer);
                }

            }

            List<List<PlaneSurface>> psrf = new List<List<PlaneSurface>>();

            for (int i = 0; i < xyz_faces.ToArray().Length; i++)
            {
                List<PlaneSurface> srflist = new List<PlaneSurface>();
                List<Point3d> lista_de_puntos = new List<Point3d>();
                //lista_de_puntos.Clear();
                int num = 0;

                foreach (var item in xyz_faces.ToArray()[i])
                {

                    if (item != null)
                    {
                        num++;
                        double x = item.X * 304.8;
                        double y = item.Y * 304.8;
                        double z = item.Z * 304.8;

                        Point3d ptstart = new Point3d(x, y, z);
                        lista_de_puntos.Add(ptstart);
                    }
                    else
                    {
                        lista_de_puntos.Clear();
                        num = 0;

                    }
                    if (num == 3)
                    {
                        Rhino.Geometry.Point3d pt0 = lista_de_puntos.ToArray()[0];
                        Rhino.Geometry.Point3d pt1 = lista_de_puntos.ToArray()[1];
                        Rhino.Geometry.Point3d pt2 = lista_de_puntos.ToArray()[2];

                        Interval int1 = new Interval(0, pt0.DistanceTo(pt1));

                        var plane = new Rhino.Geometry.Plane(pt0, pt1, pt2);
                        var plane_surface = new PlaneSurface(plane, int1, new Interval(0, pt1.DistanceTo(pt2)));
                        //m_modelfile.Objects.AddSurface(plane_surface);
                        srflist.Add(plane_surface);
                        lista_de_puntos.Clear();
                        num = 0;
                    }

                }
                psrf.Add(srflist);
            }



            for (int i = 0; i < created_layers.ToArray().Length; i++)
            {
                count_++;
                foreach (var item in psrf.ToArray()[i])
                {

                    Rhino.DocObjects.ObjectAttributes myAtt = new Rhino.DocObjects.ObjectAttributes();
                    Rhino.DocObjects.Layer layer_ = m_modelfile.AllLayers.FindIndex(count_);

                    myAtt.LayerIndex = layer_.Index;
                    m_modelfile.Objects.AddSurface(item, myAtt);

                    //int layerIndex2 = thisDoc.Layers.FindByFullPath(child_layer.ToArray()[i].FullPath,false);
                }
            }
            File3dmWriteOptions options = new File3dmWriteOptions();
            options.SaveUserData = true;
            m_modelfile.Write(filename2, 0);

            return Autodesk.Revit.UI.Result.Succeeded;
        }
    }


    [Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)]
    public class Find_dwg : IExternalCommand
    {
        static AddInId appId = new AddInId(new Guid("5F88CC78-A137-6509-AAF8-A478F3B24BAB"));
        public Autodesk.Revit.UI.Result Execute(ExternalCommandData commandData, ref string message, ElementSet elementSet)
        {
            UIDocument uidoc = commandData.Application.ActiveUIDocument;
            UIApplication uiapp = commandData.Application;
            //UIDocument uidoc = uiapp.ActiveUIDocument;
            Autodesk.Revit.DB.Document doc = uidoc.Document;

            try
            {
                string filename = @"T:\Transfer\lopez\Book1.xlsx";
                using (ExcelPackage package = new ExcelPackage(new FileInfo(filename)))
                {
                    ExcelWorksheet sheet = package.Workbook.Worksheets.ElementAt(0);

                    int column = 22;
                    int number = Convert.ToInt32(sheet.Cells[2, column].Value);
                    sheet.Cells[2, column].Value = (number + 1); ;
                    package.Save();
                }
            }
            catch (Exception)
            {
                MessageBox.Show("Excel file not found", "");
            }

            IEnumerable<Autodesk.Revit.DB.View> viewList = from elem in new FilteredElementCollector(doc)
                                               .OfClass(typeof(Autodesk.Revit.DB.View))
                                               .OfCategory(BuiltInCategory.OST_Views)
                                                           let type = elem as Autodesk.Revit.DB.View
                                                           where type.Name != null
                                                           select type;


            IEnumerable<Element> allImports = new FilteredElementCollector(doc)
                .OfClass(typeof(ImportInstance));

            // By adding the Cast, the imports are converted from the more generic 'Element' class
            // to the ImportInstance class
            // This lets us access the ImportInstance.IsLinked Property 
            IEnumerable<Element> allImportsThatAreLinked = new FilteredElementCollector(doc)
                .OfClass(typeof(ImportInstance))
                .Cast<ImportInstance>()
                .Where(q => !q.IsLinked);

            // Here the Where clause is expanded so that the results contain only the ImportInstances
            // whose 'owner view' is a drafting view
            IEnumerable<Element> allImportsThatAreLinkedAndOwnedByDraftingViews = new FilteredElementCollector(doc)
                .OfClass(typeof(ImportInstance))
                .Cast<ImportInstance>()
                .Where(q => q.IsLinked && doc.GetElement(q.OwnerViewId) is ViewDrafting);

            // Adding the Select clause changes what items the list contains
            // Instead of getting a list of imports, we now get a list of the Views that own the imports
            // https://www.dotnetperls.com/select
            // https://www.tutorialspoint.com/chash-linq-select-method
            IEnumerable<Element> viewsThatContainLinkedImports = new FilteredElementCollector(doc)
                .OfClass(typeof(ImportInstance))
                .Cast<ImportInstance>()
                .Where(q => q.IsLinked && doc.GetElement(q.OwnerViewId) is ViewDrafting)
                .Select(q => doc.GetElement(q.OwnerViewId));

            Form14 form = new Form14();

           

            List<ImportInstance> importlist = new List<ImportInstance>();
            List<CADLinkType> cadtype = new List<CADLinkType>();
            foreach (var item in allImportsThatAreLinked)
            {

                ImportInstance cadInst = doc.GetElement(item.Id) as ImportInstance;
                CADLinkType cadLinkType = doc.GetElement(cadInst.GetTypeId()) as CADLinkType;
                importlist.Add(cadInst);

                form.listBox2.Items.Add(cadLinkType.Name);
                cadtype.Add(cadLinkType);
            }

            foreach (var item in importlist)
            {
                try
                {
                    Autodesk.Revit.DB.View v_ = doc.GetElement(item.OwnerViewId) as Autodesk.Revit.DB.View;
                    form.listBox1.Items.Add(v_.Name);
                }
                catch (Exception)
                {
                }
            }
            form.ShowDialog();

            //foreach (var item in cadtype)
            //{
            //    if (item.Name == form.listBox2.SelectedItem.ToString())
            //    {
            //        BoundingBoxXYZ bbbox = item.get_BoundingBox(null);
            //        XYZ max_ = bbbox.Max;
            //        XYZ min_ = bbbox.Min;

            //        foreach (var item2 in uidoc.GetOpenUIViews())
            //        {
            //            if (item2.ViewId.Equals(doc.ActiveView.Id))
            //            {
            //                //item2.ZoomToFit();
            //                item2.ZoomAndCenterRectangle(max_, min_);
            //            }
            //        }
            //    }
            //}


            int selection = form.listBox2.SelectedIndex;

            ImportInstance select_imported = importlist.ToArray()[selection];

            List<ElementId> selections = new List<ElementId>();

            selections.Add(select_imported.Id);

            uidoc.Selection.SetElementIds(selections);
            //uidoc.Selection.SetElementIds(importlist.Select(q => q.Id).ToList());


            

            List<Autodesk.Revit.DB.View> list = new List<Autodesk.Revit.DB.View>();
            foreach (var item in viewList)
            {
                list.Add(item);
            }

           

            foreach (var item2 in viewList)
            {
                List<Element> listele = new List<Element>();

                foreach (Element e in new FilteredElementCollector(doc).OwnedByView(item2.Id))
                {
                    listele.Add(e);
                    if (select_imported.Id == e.Id)
                    {
                        uidoc.ActiveView = item2;
                    }
                }
            }

            BoundingBoxXYZ bbbox = select_imported.get_BoundingBox(null);
            XYZ max_ = bbbox.Max;
            XYZ min_ = bbbox.Min;

            foreach (var item2 in uidoc.GetOpenUIViews())
            {
                if (item2.ViewId.Equals(doc.ActiveView.Id))
                {
                    //item2.ZoomToFit();
                    item2.ZoomAndCenterRectangle(max_, min_);
                }
            }


            if (form.DialogResult == DialogResult.Cancel)
            {
               
                return Autodesk.Revit.UI.Result.Cancelled;
            }

            ////			 A view can have more than one linked import, so Distinct is used to remove duplicates
            ////			 string.Join returns a string that has the delimiter inserted between each item in the list
            //TaskDialog.Show("Views", string.Join(",", viewsThatContainLinkedImports.Select(q => q.Name).Distinct()));

            return Autodesk.Revit.UI.Result.Succeeded;
        }
    }

    [Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)]
    public class select_detailline : IExternalCommand
    {
        static AddInId appId = new AddInId(new Guid("5F88CC78-A137-6509-AAF8-A478F3B24BAB"));
        public Autodesk.Revit.UI.Result Execute(ExternalCommandData commandData, ref string message, ElementSet elementSet)
        {
            UIDocument uidoc = commandData.Application.ActiveUIDocument;
            UIApplication uiapp = commandData.Application;
            //UIDocument uidoc = uiapp.ActiveUIDocument;
            Autodesk.Revit.DB.Document doc = uidoc.Document;

            try
            {
                string filename = @"T:\Transfer\lopez\Book1.xlsx";
                using (ExcelPackage package = new ExcelPackage(new FileInfo(filename)))
                {
                    ExcelWorksheet sheet = package.Workbook.Worksheets.ElementAt(0);

                    int column = 23;
                    int number = Convert.ToInt32(sheet.Cells[2, column].Value);
                    sheet.Cells[2, column].Value = (number + 1); ;
                    package.Save();
                }
            }
            catch (Exception)
            {
                MessageBox.Show("Excel file not found", "");
            }

            Form3 form = new Form3();

            form.ShowDialog();
            
            DetailLine line = doc.GetElement(uidoc.Selection.PickObject(ObjectType.Element, "Select detail line")) as DetailLine;
            
           List< Autodesk.Revit.DB.Element> lines = new FilteredElementCollector(doc).OfClass(typeof(CurveElement)).OfCategory(BuiltInCategory.OST_Lines).Where(q => q is DetailLine).ToList(); ;

            List<DetailLine> selected = new List<DetailLine>();

            if (form.radioButton1.Checked == true)
            {
                foreach (var item in lines)
                {
                    DetailLine DL = item as DetailLine;
                    if (DL.LineStyle.Id == line.LineStyle.Id)
                    {
                        if (line.OwnerViewId == DL.OwnerViewId)
                        {
                            selected.Add(DL);
                        }
                    }
                }
            }


            if (form.radioButton2.Checked == true)
            {
                foreach (var item in lines)
                {
                    DetailLine DL = item as DetailLine;

                    if (DL.LineStyle.Id == line.LineStyle.Id)
                    {
                        selected.Add(DL);
                    }
                }

            }



            uidoc.Selection.SetElementIds(selected.Select(q => q.Id).ToList());


            return Autodesk.Revit.UI.Result.Cancelled;
        }
    }

    [Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)]
    public class RoomElevations : IExternalCommand
    {
        static AddInId appId = new AddInId(new Guid("07B86EFE-F18B-4354-AA9B-29F3E9C5F5AB"));
        public Autodesk.Revit.UI.Result Execute(ExternalCommandData commandData, ref string message, ElementSet elementSet)
        {
            UIDocument uidoc = commandData.Application.ActiveUIDocument;
            Autodesk.Revit.DB.Document doc = uidoc.Document;

            try
            {
                string filename = @"T:\Transfer\lopez\Book1.xlsx";
                using (ExcelPackage package = new ExcelPackage(new FileInfo(filename)))
                {
                    ExcelWorksheet sheet = package.Workbook.Worksheets.ElementAt(0);

                    int column = 24;
                    int number = Convert.ToInt32(sheet.Cells[2, column].Value);
                    sheet.Cells[2, column].Value = (number + 1); ;
                    package.Save();
                }
            }
            catch (Exception)
            {
                MessageBox.Show("Excel file not found", "");
            }

            //---------------------------------------- FILTERS ------------------------------------

            FilteredElementCollector levels = new FilteredElementCollector(doc).OfClass(typeof(FamilyInstance)).OfCategory(BuiltInCategory.OST_Levels);
            ICollection<Element> level1 = levels.ToElements();

            ViewFamilyType vft = new FilteredElementCollector(doc).OfClass(typeof(ViewFamilyType)).Cast<ViewFamilyType>().FirstOrDefault<ViewFamilyType>(x => ViewFamily.Section == x.ViewFamily);
            ViewFamilyType vftele = new FilteredElementCollector(doc).OfClass(typeof(ViewFamilyType)).Cast<ViewFamilyType>().FirstOrDefault<ViewFamilyType>(x => ViewFamily.Elevation == x.ViewFamily);

            FilteredElementCollector rooms1 = new FilteredElementCollector(doc).WhereElementIsNotElementType().OfClass(typeof(SpatialElement));
            ICollection<Element> room2 = rooms1.ToElements();

            //Autodesk.Revit.DB.View viewTemplate = (from v in new FilteredElementCollector(doc).OfClass(typeof(Autodesk.Revit.DB.View)).Cast<Autodesk.Revit.DB.View>()
            //                                       where v.IsTemplate && v.Name == "Architectural Section"
            //                                       select v).First();

            IEnumerable<ViewFamilyType> viewFamilyTypes = from elem in new FilteredElementCollector(doc).OfClass(typeof(ViewFamilyType))
                                                          let type = elem as ViewFamilyType
                                                          where type.ViewFamily == ViewFamily.ThreeDimensional
                                                          select type;


            ViewFamilyType ceiling_plan_view = new FilteredElementCollector(doc).OfClass(typeof(ViewFamilyType)).Cast<ViewFamilyType>().FirstOrDefault<ViewFamilyType>(x => ViewFamily.CeilingPlan == x.ViewFamily);
            ViewFamilyType floor_plan_view = new FilteredElementCollector(doc).OfClass(typeof(ViewFamilyType)).Cast<ViewFamilyType>().FirstOrDefault<ViewFamilyType>(x => ViewFamily.FloorPlan == x.ViewFamily);
            string project_param = /*"Schedule Identifier"*/"Comments";

            using (Transaction t1 = new Transaction(doc))
            {
                List<Element> roomele = new List<Element>();
                List<Element> selected_room = new List<Element>();


                

                foreach (Element j in room2)
                {
                    try
                    {
                        ParameterMap parmap = j.ParametersMap;
                        Parameter par = parmap.get_Item(project_param);
                        string value = par.AsString();

                        if (value != null)
                        {
                            Element E = j;
                            roomele.Add(E);

                        }
                    }
                    catch (Exception)
                    {
                        TaskDialog.Show("Task Cancelled", "Project does not contain parameter " + project_param + "or may be no room has an identifier");
                        //TaskDialog.Show("Warning", "Task Cancelled");
                        return Autodesk.Revit.UI.Result.Cancelled;
                    }
                    
                   
                }

                Form_for_room_Schedule form = new Form_for_room_Schedule();



                foreach (var item in roomele)
                {
                    form.listBox1.Items.Add(item.Name);
                }
                

                if (roomele.Count == 0)
                {
                    TaskDialog.Show("Task Cancelled", "no room has an identifier");
                    
                    return Autodesk.Revit.UI.Result.Cancelled;
                }

                form.ShowDialog();


                if (form.DialogResult == DialogResult.Cancel)
                {
                    return Autodesk.Revit.UI.Result.Cancelled;
                }

                foreach (var item in roomele)
                {
                    try
                    {
                        for (int i = 0; i < form.listBox1.SelectedItems.Count; i++)
                        {
                            if (form.listBox1.SelectedItems[i].ToString() == item.Name)
                            {
                                selected_room.Add(item);
                            }
                        }

                        
                    }
                    catch (Exception)
                    {
                        TaskDialog.Show("Task Cancelled", "A room must be selected");
                        
                        return Autodesk.Revit.UI.Result.Cancelled;
                        throw;
                    }

                }


                t1.Start("Create views from selected view");

                if (form.checkBox1.Checked)
                {
                    foreach (Element item in selected_room)
                    {
                        ElementId id = item.LevelId;
                        try
                        {
                            ElementId level11 = doc.ActiveView.GenLevel.Id;
                        }
                        catch
                        {
                            TaskDialog.Show("Warning", "you must be in a floor plan");
                            goto final;
                        }

                        if (item.LevelId.IntegerValue == doc.ActiveView.GenLevel.Id.IntegerValue)
                        {

                            if (item.Location is LocationPoint)
                            {
                                LocationPoint lp = item.Location as LocationPoint;
                                double pz = lp.Point.X;
                                XYZ point = lp.Point;


                                try
                                {

                                    ElevationMarker marker = ElevationMarker.CreateElevationMarker(doc, vftele.Id, point, 100);
                                    for (int i = 0; i < 4; i++)
                                    {
                                        ViewSection elevation1 = marker.CreateElevation(doc, doc.ActiveView.Id, i);
                                        elevation1.Name = form.textBox1.Text + " "+  item.Name + " " + i.ToString() ;
                                    }
                                    //TaskDialog.Show("View/s created", "Interior Elevations were created for " + form.listBox1.SelectedItems.Count + " Rooms");
                                }
                                catch
                                {
                                   

                                    TaskDialog.Show("Element info", " Name must be unique ");
                                }
                            }
                        }
                    }
                }
                if (form.checkBox2.Checked)
                {
                    foreach (Element item in selected_room)
                    {
                        ElementId id = item.LevelId;
                        string name = item.Name;
                        ElementId level11 = doc.ActiveView.GenLevel.Id;

                        BoundingBoxXYZ room_box = item.get_BoundingBox(null);
                        XYZ mas_10 = new XYZ(4, 4, 4);
                        XYZ max_p = room_box.Max + mas_10;
                        XYZ min_p = room_box.Min - mas_10;

                        View3D view3D = View3D.CreateIsometric(doc, viewFamilyTypes.First().Id);

                        try
                        {
                            view3D.Name = form.textBox1.Text + " " + item.Name + item.Name;
                        }
                        catch (Exception)
                        {

                            TaskDialog.Show("Element info", " Name must be unique ");
                        }
                        
                        BoundingBoxXYZ boundingBoxXYZ = new BoundingBoxXYZ();
                        boundingBoxXYZ.Min = min_p;
                        boundingBoxXYZ.Max = max_p;
                        view3D.SetSectionBox(boundingBoxXYZ);
                        //TaskDialog.Show("View/s created", "Isometric view was created for " + form.listBox1.SelectedItems.Count + " Rooms");
                    }
                }

                if (form.checkBox3.Checked)
                {

                    foreach (Element item in selected_room)
                    {
                        ElementId id = item.LevelId;
                        string name = item.Name;


                        BoundingBoxXYZ room_box = item.get_BoundingBox(null);
                        XYZ mas_10 = new XYZ(4, 4, 4);
                        XYZ max_p = room_box.Max + mas_10;
                        XYZ min_p = room_box.Min - mas_10;

                        BoundingBoxXYZ boundingBoxXYZ = new BoundingBoxXYZ();
                        boundingBoxXYZ.Min = min_p;
                        boundingBoxXYZ.Max = max_p;

                        ElementId level11 = doc.ActiveView.GenLevel.Id;
                        ViewPlan floorView = ViewPlan.Create(doc, floor_plan_view.Id, id);

                        try
                        {
                            floorView.Name = form.textBox1.Text + " " + item.Name + item.Name;
                        }
                        catch (Exception)
                        {

                            TaskDialog.Show("Element info", " Name must be unique ");
                        }
                        
                        floorView.CropBox = boundingBoxXYZ;
                        floorView.CropBox.Enabled = true;
                        floorView.CropBoxActive = true;
                        //TaskDialog.Show("View/s created", "ViewPlan was created for " + form.listBox1.SelectedItems.Count + " Rooms");
                    }
                }

                if (form.checkBox4.Checked)
                {

                    foreach (Element item in selected_room)
                    {
                        ElementId id = item.LevelId;
                        string name = item.Name;


                        BoundingBoxXYZ room_box = item.get_BoundingBox(null);
                        XYZ mas_10 = new XYZ(4, 4, 4);
                        XYZ max_p = room_box.Max + mas_10;
                        XYZ min_p = room_box.Min - mas_10;

                        BoundingBoxXYZ boundingBoxXYZ = new BoundingBoxXYZ();
                        boundingBoxXYZ.Min = min_p;
                        boundingBoxXYZ.Max = max_p;

                        ElementId level11 = doc.ActiveView.GenLevel.Id;
                        ViewPlan ceiling_View = ViewPlan.Create(doc, ceiling_plan_view.Id, id);

                        try
                        {
                            ceiling_View.Name = form.textBox1.Text + " " + item.Name + item.Name;
                        }
                        catch (Exception)
                        {

                            TaskDialog.Show("Element info", " Name must be unique ");
                        }
                        
                        ceiling_View.CropBox = boundingBoxXYZ;
                        ceiling_View.CropBox.Enabled = true;
                        ceiling_View.CropBoxActive = true;
                        //TaskDialog.Show("View/s created", "Ceiling was created for " + form.listBox1.SelectedItems.Count + " Rooms");
                    }
                }
                t1.Commit();

            }
            final: return Autodesk.Revit.UI.Result.Succeeded;
        }
    }


    [Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)]
    public class RevCloud : IExternalCommand
    {
        static AddInId appId = new AddInId(new Guid("5F92CC78-A137-4809-AAF8-A478F3B24BAB"));
        public Autodesk.Revit.UI.Result Execute(ExternalCommandData commandData, ref string message, ElementSet elementSet)
        {
            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Autodesk.Revit.DB.Document doc = uidoc.Document;



            // get elements from user selection
            List<Element> elements = new List<Element>();

            // string of ids of the selected elements to use for Revision description
            string description = "";
            foreach (Reference r in uidoc.Selection.PickObjects(ObjectType.Element))
            {
                elements.Add(doc.GetElement(r));
                description += r.ElementId.IntegerValue + " ";
            }

            using (Transaction t = new Transaction(doc, "Make Revision & Cloud"))
            {
                t.Start();

                // create new revision
                Revision rev = Revision.Create(doc);
                rev.RevisionDate = DateTime.Now.ToShortDateString();
                rev.Description = description;
                rev.IssuedBy = doc.Application.Username;

                // use bounding box of element and offset to create curves for revision cloud
                foreach (Element e in elements)
                {
                    BoundingBoxXYZ bbox = e.get_BoundingBox(doc.ActiveView);
                    List<Autodesk.Revit.DB.Curve> curves = new List<Autodesk.Revit.DB.Curve>();
                    double offset = 2;
                    XYZ pt1 = bbox.Min.Subtract(XYZ.BasisX.Multiply(offset)).Subtract(XYZ.BasisY.Multiply(offset));
                    XYZ pt2 = new XYZ(bbox.Min.X - offset, bbox.Max.Y + offset, 0);
                    XYZ pt3 = bbox.Max.Add(XYZ.BasisX.Multiply(offset)).Add(XYZ.BasisY.Multiply(offset)); ;
                    XYZ pt4 = new XYZ(bbox.Max.X + offset, bbox.Min.Y - offset, 0);
                    curves.Add(Autodesk.Revit.DB.Line.CreateBound(pt1, pt2));
                    curves.Add(Autodesk.Revit.DB.Line.CreateBound(pt2, pt3));
                    curves.Add(Autodesk.Revit.DB.Line.CreateBound(pt3, pt4));
                    curves.Add(Autodesk.Revit.DB.Line.CreateBound(pt4, pt1));

                    // create revision cloud
                    RevisionCloud cloud = RevisionCloud.Create(doc, doc.ActiveView, rev.Id, curves);

                    // set Comments of revision cloud
                    cloud.get_Parameter(BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS)
                        .Set(e.GetType().Name + "-" + e.Name + ": " + doc.Application.Username);

                    // tag the revision cloud
                    //IndependentTag tag = doc.Create.NewSpaceTag(doc.ActiveView, cloud, true, TagMode.TM_ADDBY_CATEGORY, TagOrientation.Horizontal, pt3);
                    //tag.TagHeadPosition = pt3.Add(new XYZ(2, 2, 0));
                }
                t.Commit();

                return Autodesk.Revit.UI.Result.Succeeded;
            }
        }
    }

    [Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)]
    public class Door_Section : IExternalCommand
    {
       
        class ProjectParameterData
        {
            public Definition Definition = null;
            public ElementBinding Binding = null;
            public string Name = null;                // Needed because accsessing the Definition later may produce an error.
            public bool IsSharedStatusKnown = false;  // Will probably always be true when the data is gathered
            public bool IsShared = false;
            public string GUID = null;
        }
       
        static List<ProjectParameterData>
          GetProjectParameterData(
            Autodesk.Revit.DB.Document doc)
        {
            // Following good SOA practices, first validate incoming parameters

            if (doc == null)
            {
                throw new ArgumentNullException("doc");
            }

            if (doc.IsFamilyDocument)
            {
                throw new Exception("doc can not be a family document.");
            }

            List<ProjectParameterData> result
              = new List<ProjectParameterData>();

            BindingMap map = doc.ParameterBindings;
            DefinitionBindingMapIterator it
              = map.ForwardIterator();
            it.Reset();
            while (it.MoveNext())
            {
                ProjectParameterData newProjectParameterData
                  = new ProjectParameterData();

                newProjectParameterData.Definition = it.Key;
                newProjectParameterData.Name = it.Key.Name;
                newProjectParameterData.Binding = it.Current
                  as ElementBinding;

                result.Add(newProjectParameterData);
            }
            return result;
        }


        static AddInId appId = new AddInId(new Guid("BFB59CE4-49D2-4C53-84D8-726E441220DD"));
        public Autodesk.Revit.UI.Result Execute(ExternalCommandData commandData, ref string message, ElementSet elementSet)
        {
            UIDocument uidoc = commandData.Application.ActiveUIDocument;
            Autodesk.Revit.DB.Document doc = uidoc.Document;


            //string name = doc.Title;
            //string path = doc.PathName;
            //string comments = "Door_Section" + "_" + doc.Application.Username + "_" + doc.Title;
            // Use the @ before the filename to avoid errors related to the \ character
            // Alternative is to use \\ instead of \
            //string filename = @"C:\Users\Alex\Desktop Comments.txt";
            //System.Diagnostics.Process.Start(filename);
            // Create a StreamWriter object to write to a file
            //StreamWriter writer = new StreamWriter(filename);
            //writer.WriteLine(DateTime.Now + " - " + comments);
            //writer.Close();
            //using (System.IO.StreamWriter writer = new System.IO.StreamWriter(filename))
            //{
            //    writer.WriteLine(DateTime.Now + " - " + comments);
            //    writer.Close();
            //}
            // open the text file


            if (doc.IsFamilyDocument)
            {
                message = "The document must be a project document.";
                return Result.Failed;
            }

            Element projectInfoElement
              = new FilteredElementCollector(doc)
                .OfCategory(BuiltInCategory.OST_ProjectInformation)
                .FirstElement();
            
            Element firstWallTypeElement
              = new FilteredElementCollector(doc)
                .OfCategory(BuiltInCategory.OST_Walls)
                .WhereElementIsElementType()
                .FirstElement();
            
            List<ProjectParameterData> projectParametersData
              = GetProjectParameterData(doc);

            

            //---------------------------------------- FILTERS ------------------------------------------------------------------------------------------------------------------------------------------
            FilteredElementCollector Doorcollector = new FilteredElementCollector(doc).OfClass(typeof(FamilyInstance)).OfCategory(BuiltInCategory.OST_Doors);
            ICollection<Element> doorfilter = Doorcollector.ToElements();

            FilteredElementCollector anyfamily = new FilteredElementCollector(doc).OfClass(typeof(FamilyInstance));
            ICollection<Element> listanyfamily = anyfamily.ToElements();

            ViewFamilyType vft = new FilteredElementCollector(doc).OfClass(typeof(ViewFamilyType)).Cast<ViewFamilyType>().FirstOrDefault<ViewFamilyType>(x => ViewFamily.Section == x.ViewFamily);
            ViewFamilyType vftele = new FilteredElementCollector(doc).OfClass(typeof(ViewFamilyType)).Cast<ViewFamilyType>().FirstOrDefault<ViewFamilyType>(x => ViewFamily.Elevation == x.ViewFamily);
            

            FilteredElementCollector fec = new FilteredElementCollector(doc);
            fec.OfClass(typeof(ViewSection));
            var viewPlans = fec.Cast<ViewSection>().Where<ViewSection>(vp => vp.IsTemplate);
            //---------------------------------------- FILTERS ------------------------------------------------------------------------------------------------------------------------------------------
            

            //---------------------------------------- LISTS ------------------------------------
            List<Element> lista_de_nombres_puertas_elements = new List<Element>();
            List<string> sobras = new List<string>();
            List<string> lista_de_nombres_puertas = new List<string>();
            List<string> lista3 = new List<string>();
            List<string> SortByMark = new List<string>();
            List<ElementId> templaId = new List<ElementId>();

            List<Element> doorEle = new List<Element>();
            //---------------------------------------- LISTS ------------------------------------
           

            //---------------------------------------- DOORS TO DROPDOWN------------------------------------
            SortByMark.Sort();
            if (Doorcollector.ToArray().Length == 0)
            {
                TaskDialog.Show("warning", "Project contain no doors");
            }

            string project_param = /*"Schedule Identifier"*/ "Comments";
;
            foreach (Element Door in Doorcollector)
            {
                try
                {
                    ParameterMap parmap = Door.ParametersMap;
                    Parameter par = parmap.get_Item(project_param);
                    string value = par.AsString();

                    lista_de_nombres_puertas.Add(value);
                }
                catch (Exception)
                {
                    TaskDialog.Show("!", "Project requires (Schedule Identifier) parameter");
                    return Autodesk.Revit.UI.Result.Cancelled;
                }
            }
            //---------------------------------------- DOORS TO DROPDOWN-----------------------------------


            //---------------------------------------- TEMPLATES -------------------------------
            if (viewPlans.ToArray().Length == 0)
            {
                TaskDialog.Show("warning", "Project does not contain templates");
                return Autodesk.Revit.UI.Result.Cancelled;
            }
            foreach (Element j in viewPlans)
            {
                lista3.Add(j.Name);
                templaId.Add(j.Id);
            }
            //---------------------------------------- WINDOW FORM----------------------------------
            
            Form_for_door_Schedule form = new Form_for_door_Schedule();
            foreach (var item in lista_de_nombres_puertas)
            {
                if (item != null )
                {
                    if (!form.dropdown1.Items.Contains(item))
                    {
                        form.dropdown1.Items.Add(item);
                    }
                }
            }
            foreach (var item in lista3)
            {
                form.comboBox2.Items.Add(item);
            }

            if (form.dropdown1.Items.Count == 0)
            {
                TaskDialog.Show("Instruction!", "no identifiers specified");
                form.Close();
                goto final;
            }

            form.ShowDialog();

            if (form.DialogResult == DialogResult.Cancel)
            {
                return Autodesk.Revit.UI.Result.Cancelled;
            }

            if (form.Equals(false))
            {
                return Autodesk.Revit.UI.Result.Cancelled;
            }

            int index_tBox = form.comboBox2.SelectedIndex;
            ElementId choose_template = null;
            if (index_tBox != -1)
            {
                choose_template = templaId.ElementAt(index_tBox);
            }
            else
            {
                TaskDialog.Show("Warning", "A template must be choosen");
                goto final;
            }

            List<double> numHoj = new List<double>();
            double uno1 = 1;
            numHoj.Add(uno1);

            ViewSection vs = null;
            List<ViewSection> viewsec = new List<ViewSection>();
            viewsec.Sort();
            //---------------------------------------- WINDOW FORM------------------------------------
            //---------------------------------------- DOOR VIEWS CREATION ------------------------------------
            using (Transaction t1 = new Transaction(doc))
            {
                foreach (Element j in Doorcollector)
                {
                    ParameterMap parmap = j.ParametersMap;
                    Parameter par = parmap.get_Item(project_param);
                    string value = par.AsString();

                    if (form.dropdown1.SelectedItem.ToString() == value)
                    {
                        doorEle.Add(j);
                    }
                }

                t1.Start("Create elevations from doors");

                for (int i = 0; i < doorEle.ToArray().Length; i++)
                {
                    Autodesk.Revit.DB.Transform trans = null;
                    if (doorEle.ToArray()[i] is FamilyInstance)
                    {
                        if (doorEle.ToArray()[i].Location is LocationPoint)
                        {
                            LocationPoint lp = doorEle.ToArray()[i].Location as LocationPoint;
                            double pz = lp.Point.X;
                            XYZ point = lp.Point;
                            double pointoffsetX = point.X + -1;
                            XYZ offsetX = new XYZ(pointoffsetX, point.Y, point.Z);
                            double pointoffsetY = point.Y + -1;
                            XYZ offsetY = new XYZ(point.X, pointoffsetY, point.Z);

                            FamilyInstance fi = doorEle.ToArray()[i] as FamilyInstance;
                            XYZ orientation = fi.FacingOrientation;
                            //orientation = new XYZ(1, 0, 0);
                            //fi.flipFacing();

                            trans = fi.GetTransform();

                            if (fi.FacingOrientation.Y == -1)
                            {
                                Autodesk.Revit.DB.Transform t = Autodesk.Revit.DB.Transform.Identity;
                                t.Origin = offsetY;
                                t.BasisX = new XYZ(0, 0, 1);
                                t.BasisY = new XYZ(1, 0, 0);
                                t.BasisZ = new XYZ(0, 1, 0);
                                BoundingBoxXYZ newbbox = new BoundingBoxXYZ();
                                newbbox.Transform = t;
                                newbbox.Min = new XYZ(0, -3, 0);
                                newbbox.Max = new XYZ(10,3,10);
                                ViewFamilyType vft1 = vft;

                                vs = ViewSection.CreateSection(doc, vft1.Id, newbbox);


                                XYZ levelPoint = new XYZ(point.X, point.Y, point.Z);
                                //FamilyInstance doorTag = doc.Create.NewFamilyInstance(levelPoint, doorTagType, vs);


                                if (vs != null)
                                {
                                    vs.ViewTemplateId = choose_template;
                                }
                                //vs.Scale = 200;
                                if (vs != null)
                                {
                                    try
                                    {
                                        vs.Name = form.textBox1.Text + " " + doorEle.ToArray()[i].Name + " " + i;
                                        viewsec.Add(vs);
                                    }
                                    catch
                                    {

                                        //vs.Name = "SCHEDULE Type -" + doorEle.ToArray()[i].LookupParameter(project_param).AsString() + "_" + i;
                                        TaskDialog.Show("Error", "Name must be unique");
                                    }
                                }
                                doc.Regenerate();
                            }

                            if (fi.FacingOrientation.Y == 1)
                            {
                                Autodesk.Revit.DB.Transform t = Autodesk.Revit.DB.Transform.Identity;
                                t.Origin = offsetY;
                                t.BasisX = new XYZ(0, 0, 1);
                                t.BasisY = new XYZ(1, 0, 0);
                                t.BasisZ = new XYZ(0, 1, 0);
                                BoundingBoxXYZ newbbox = new BoundingBoxXYZ();
                                newbbox.Transform = t;
                                newbbox.Min = new XYZ(0, -3, 0);
                                newbbox.Max = new XYZ(10, 3, 10);
                                ViewFamilyType vft1 = vft;
                                vs = ViewSection.CreateSection(doc, vft1.Id, newbbox);

                                XYZ levelPoint = new XYZ(point.X, point.Y, point.Z);
                                //FamilyInstance doorTag = doc.Create.NewFamilyInstance(levelPoint, doorTagType, vs);

                                if (vs != null)
                                {
                                    vs.ViewTemplateId = choose_template;
                                }
                                //vs.Scale = 200;
                                if (vs != null)
                                {
                                    try
                                    {
                                        vs.Name = form.textBox1.Text + " " + doorEle.ToArray()[i].Name + " " + i;
                                        viewsec.Add(vs);
                                    }
                                    catch
                                    {

                                        //vs.Name = "SCHEDULE Type -" + doorEle.ToArray()[i].LookupParameter(project_param).AsString() + "_" + i;
                                        TaskDialog.Show("Error", "Name must be unique");
                                    }
                                }
                                doc.Regenerate();
                            }

                            if (fi.FacingOrientation.X == 1)
                            {
                                Autodesk.Revit.DB.Transform t = Autodesk.Revit.DB.Transform.Identity;
                                t.Origin = offsetX;
                                t.BasisX = new XYZ(0, 1, 0);
                                t.BasisY = new XYZ(0, 0, 1);
                                t.BasisZ = new XYZ(1, 0, 0);
                                BoundingBoxXYZ newbbox = new BoundingBoxXYZ();
                                newbbox.Transform = t;
                                newbbox.Min = new XYZ(-3, 0, 0);
                                newbbox.Max = new XYZ(3, 10,10);
                                ViewFamilyType vft1 = vft;
                                vs = ViewSection.CreateSection(doc, vft1.Id, newbbox);

                                XYZ levelPoint = new XYZ(point.X, point.Y, point.Z);
                                //FamilyInstance doorTag = doc.Create.NewFamilyInstance(levelPoint, doorTagType, vs);

                                string lookp = fi.LookupParameter("Mark").AsString();
                                if (vs != null)
                                {
                                    vs.ViewTemplateId = choose_template;
                                }

                                //vs.Scale = 200;
                                if (vs != null)
                                {
                                    try
                                    {
                                        vs.Name = form.textBox1.Text + " " + doorEle.ToArray()[i].Name + "_" + i;
                                        viewsec.Add(vs);
                                    }
                                    catch
                                    {

                                        //vs.Name = "SCHEDULE Type -" + doorEle.ToArray()[i].LookupParameter(project_param).AsString() + "_" + i;
                                        TaskDialog.Show("Error", "Name must be unique");
                                    }
                                }
                                doc.Regenerate();
                            }
                            if (fi.FacingOrientation.X == -1)
                            {
                                Autodesk.Revit.DB.Transform t = Autodesk.Revit.DB.Transform.Identity;
                                t.Origin = offsetX;
                                t.BasisX = new XYZ(0, 1, 0);
                                t.BasisY = new XYZ(0, 0, 1);
                                t.BasisZ = new XYZ(1, 0, 0);
                                BoundingBoxXYZ newbbox = new BoundingBoxXYZ();
                                newbbox.Transform = t;
                                newbbox.Min = new XYZ(-3, 0, 0);
                                newbbox.Max = new XYZ(3, 10, 10);
                                ViewFamilyType vft1 = vft;
                                vs = ViewSection.CreateSection(doc, vft1.Id, newbbox);

                                XYZ levelPoint = new XYZ(point.X, point.Y, point.Z);
                                //FamilyInstance doorTag = doc.Create.NewFamilyInstance(levelPoint, doorTagType, vs);

                                if (vs != null)
                                {
                                    vs.ViewTemplateId = choose_template;
                                }

                                if (vs != null)
                                {
                                    try
                                    {
                                        vs.Name = form.textBox1.Text + " " + doorEle.ToArray()[i].Name + "_" + i;
                                        viewsec.Add(vs);
                                    }
                                    catch
                                    {

                                        //vs.Name = "SCHEDULE Type -" + doorEle.ToArray()[i].LookupParameter(project_param).AsString() + "_" + i;
                                        TaskDialog.Show("Error", "Name must be unique");
                                    }
                                }
                                //vs.Scale = 200;
                                doc.Regenerate();
                                string lookp = fi.LookupParameter("Mark").AsString();
                            }
                        }
                    }
                }
               
                t1.Commit();
                TaskDialog.Show("! ", doorEle.ToArray().Length + " - New Views in the project");

            }

            uidoc.ActiveView = vs;
            final: return Autodesk.Revit.UI.Result.Succeeded;
        }
    }

    [Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)]
    public class WindowSection : IExternalCommand
    {
        static AddInId appId = new AddInId(new Guid("131C8090-A5B2-4D65-A0AC-79FBCFD8F756"));
        public Autodesk.Revit.UI.Result Execute(ExternalCommandData commandData, ref string message, ElementSet elementSet)
        {
            UIDocument uidoc = commandData.Application.ActiveUIDocument;
            Autodesk.Revit.DB.Document doc = uidoc.Document;
            //---------------------------------------- FILTERS ------------------------------------
            FilteredElementCollector windowcollwctor = new FilteredElementCollector(doc).OfClass(typeof(FamilyInstance)).OfCategory(BuiltInCategory.OST_Windows);
            ICollection<Element> Windowfamily = windowcollwctor.ToElements();


            ViewFamilyType vft = new FilteredElementCollector(doc).OfClass(typeof(ViewFamilyType)).Cast<ViewFamilyType>().FirstOrDefault<ViewFamilyType>(x => ViewFamily.Section == x.ViewFamily);
            ViewFamilyType vftele = new FilteredElementCollector(doc).OfClass(typeof(ViewFamilyType)).Cast<ViewFamilyType>().FirstOrDefault<ViewFamilyType>(x => ViewFamily.Elevation == x.ViewFamily);
            
            FilteredElementCollector viewtemplates = new FilteredElementCollector(doc).OfClass(typeof(ViewFamilyType));
            ICollection<ElementId> viewtemplatescombobox = viewtemplates.ToElementIds();
            ICollection<Element> viewtemplatesName = viewtemplates.ToElements();
            FilteredElementCollector viewTemplate = new FilteredElementCollector(doc).OfClass(typeof(Autodesk.Revit.DB.View));
            ICollection<Element> VTcolector = viewTemplate.ToElements();

            FilteredElementCollector fec = new FilteredElementCollector(doc);
            fec.OfClass(typeof(ViewSection));
            var viewPlans = fec.Cast<ViewSection>().Where<ViewSection>(vp => vp.IsTemplate);
            
            List<string> lista_de_nombres_ventanas_elements = new List<string>();
            List<string> sobras = new List<string>();
            List<string> window_names = new List<string>();
            List<string> template_list = new List<string>();
            List<string> SortByMark = new List<string>();
            List<ElementId> templaId = new List<ElementId>();
            List<Element> windowelevations = new List<Element>();
            //---------------------------------------- FILTERS ------------------------------------

            


            SortByMark.Sort();
           
            string project_param = /*"Schedule Identifier"*/  "Comments";

            //---------------------------------------- windows ------------------------------------

            if (Windowfamily.Count == 0)
            {
                TaskDialog.Show("Task Cancelled", "Project does not Windows families");

                return Autodesk.Revit.UI.Result.Cancelled;
            }


            foreach (Element j in Windowfamily)
            {
                try
                {
                    ParameterMap parmap = j.ParametersMap;
                    Parameter par = parmap.get_Item(project_param);
                    string value = par.AsString();

                    if (value != null)
                    {
                        if (!lista_de_nombres_ventanas_elements.Contains(value))
                        {
                            lista_de_nombres_ventanas_elements.Add(value);
                        }
                    }
                    else
                    {

                    }
                }
                catch (Exception)
                {
                    TaskDialog.Show("Task Cancelled", "Project does not contain parameter " + project_param);

                    return Autodesk.Revit.UI.Result.Cancelled;
                }
            }

            if (lista_de_nombres_ventanas_elements.Count == 0 )
            {
                TaskDialog.Show("warning", "Project does not contain templates");
                return Autodesk.Revit.UI.Result.Cancelled;
            } 

            //---------------------------------------- windows ------------------------------------

            //---------------------------------------- TEMPLATES ------------------------------------
            foreach (Element j in viewPlans)
            {
                template_list.Add(j.Name);
                templaId.Add(j.Id);
            }
            //---------------------------------------- TEMPLATES ------------------------------------


            //---------------------------------------- SORT ------------------------------------
            List<string> sortby = new List<string>();
            //sortby.Add("Mark");
            sortby.Add(project_param);


            //---------------------------------------- SORT ------------------------------------
            //---------------------------------------- WINDOW ------------------------------------

            Form5 form = new Form5();

            foreach (var item in lista_de_nombres_ventanas_elements)
            {
                if (!form.comboBox1.Items.Contains(item))
                {
                    form.comboBox1.Items.Add(item);
                }
            }
            foreach (var item in template_list)
            {
                form.comboBox2.Items.Add(item);
            }


            form.ShowDialog();

            if (form.DialogResult == DialogResult.Cancel)
            {
                return Autodesk.Revit.UI.Result.Cancelled;
            }

            if (window_names == null)
                form.Close();
            if (form.Equals(false))
            {
                return Autodesk.Revit.UI.Result.Cancelled;
            }

            int index_tBox = form.comboBox2.SelectedIndex;
            ElementId choose_template = templaId.ElementAt(index_tBox);

            //int index = form.comboBox1.SelectedIndex;
            //string text = window_names.ElementAt(index);

            List<ViewSection> viewsec = new List<ViewSection>();
            viewsec.Sort();

            ViewSection vs = null;
            //---------------------------------------- WINDOW ------------------------------------
            //---------------------------------------- DOOR CREATION ------------------------------------
            using (Transaction t1 = new Transaction(doc))
            {
                foreach (Element j in Windowfamily)
                {
                    ParameterMap parmap = j.ParametersMap;
                    Parameter par = parmap.get_Item(project_param);
                    string value = par.AsString();

                    if (value == form.comboBox1.SelectedItem.ToString())
                    {
                        windowelevations.Add(j);

                    }
                }

                t1.Start("create elevations from doors");

                for (int i = 0; i < windowelevations.ToArray().Length; i++)
                {
                    Autodesk.Revit.DB.Transform trans = null;
                    if (windowelevations.ToArray()[i] is FamilyInstance)
                    {
                        if (windowelevations.ToArray()[i].Location is LocationPoint)
                        {
                            LocationPoint lp = windowelevations.ToArray()[i].Location as LocationPoint;
                            double pz = lp.Point.X;
                            XYZ point = lp.Point;
                            double pointoffsetX = point.X + -1;
                            XYZ offsetX = new XYZ(pointoffsetX, point.Y, point.Z);
                            double pointoffsetY = point.Y + -1;
                            XYZ offsetY = new XYZ(point.X, pointoffsetY, point.Z);

                            FamilyInstance fi = windowelevations.ToArray()[i] as FamilyInstance;
                            XYZ orientation = fi.FacingOrientation;

                            trans = fi.GetTransform();

                            if (fi.FacingOrientation.Y == -1)
                            {
                                Autodesk.Revit.DB.Transform t = Autodesk.Revit.DB.Transform.Identity;
                                t.Origin = offsetY;
                                t.BasisX = new XYZ(0, 0, 1);
                                t.BasisY = new XYZ(1, 0, 0);
                                t.BasisZ = new XYZ(0, 1, 0);
                                BoundingBoxXYZ newbbox = new BoundingBoxXYZ();
                                newbbox.Transform = t;
                                newbbox.Min = new XYZ(0, -3, 0);
                                newbbox.Max = new XYZ(10, 3, 10);
                                ViewFamilyType vft1 = vft;

                                vs = ViewSection.CreateSection(doc, vft1.Id, newbbox);

                                if (vs != null)
                                {
                                    vs.ViewTemplateId = choose_template;
                                }
                                //vs.Scale = 200;
                                if (vs != null)
                                {
                                    try
                                    {
                                        vs.Name = form.textBox1.Text + " " + windowelevations.ToArray()[i].Name + " " + i;
                                        viewsec.Add(vs);
                                    }
                                    catch
                                    {

                                       

                                        TaskDialog.Show("Error", "Name must be unique");
                                    }
                                }
                                doc.Regenerate();
                            }

                            if (fi.FacingOrientation.Y == 1)
                            {
                                Autodesk.Revit.DB.Transform t = Autodesk.Revit.DB.Transform.Identity;
                                t.Origin = offsetY;
                                t.BasisX = new XYZ(0, 0, 1);
                                t.BasisY = new XYZ(1, 0, 0);
                                t.BasisZ = new XYZ(0, 1, 0);
                                BoundingBoxXYZ newbbox = new BoundingBoxXYZ();
                                newbbox.Transform = t;
                                newbbox.Min = new XYZ(0, -3, 0);
                                newbbox.Max = new XYZ(10, 3, 10);
                                ViewFamilyType vft1 = vft;
                                vs = ViewSection.CreateSection(doc, vft1.Id, newbbox);

                                if (vs != null)
                                {
                                    vs.ViewTemplateId = choose_template;
                                }
                                //vs.Scale = 200;
                                if (vs != null)
                                {
                                    try
                                    {

                                        vs.Name = form.textBox1.Text + " " + windowelevations.ToArray()[i].Name + " " + i;
                                        viewsec.Add(vs);
                                    }
                                    catch
                                    {


                                        TaskDialog.Show("Error", "Name must be unique");
                                    }
                                }
                                doc.Regenerate();
                            }

                            if (fi.FacingOrientation.X == 1)
                            {
                                Autodesk.Revit.DB.Transform t = Autodesk.Revit.DB.Transform.Identity;
                                t.Origin = offsetX;
                                t.BasisX = new XYZ(0, 1, 0);
                                t.BasisY = new XYZ(0, 0, 1);
                                t.BasisZ = new XYZ(1, 0, 0);
                                BoundingBoxXYZ newbbox = new BoundingBoxXYZ();
                                newbbox.Transform = t;
                                newbbox.Min = new XYZ(-3, 0, 0);
                                newbbox.Max = new XYZ(3, 10, 10);
                                ViewFamilyType vft1 = vft;
                                vs = ViewSection.CreateSection(doc, vft1.Id, newbbox);

                                if (vs != null)
                                {
                                    vs.ViewTemplateId = choose_template;
                                }

                                //vs.Scale = 200;
                                if (vs != null)
                                {
                                    try
                                    {

                                        vs.Name = form.textBox1.Text + " " + windowelevations.ToArray()[i].Name + " " + i;
                                        viewsec.Add(vs);
                                    }
                                    catch
                                    {


                                        TaskDialog.Show("Error", "Name must be unique");
                                    }
                                }
                                doc.Regenerate();
                            }
                            if (fi.FacingOrientation.X == -1)
                            {
                                Autodesk.Revit.DB.Transform t = Autodesk.Revit.DB.Transform.Identity;
                                t.Origin = offsetX;
                                t.BasisX = new XYZ(0, 1, 0);
                                t.BasisY = new XYZ(0, 0, 1);
                                t.BasisZ = new XYZ(1, 0, 0);
                                BoundingBoxXYZ newbbox = new BoundingBoxXYZ();
                                newbbox.Transform = t;
                                newbbox.Min = new XYZ(-3, 0, 0);
                                newbbox.Max = new XYZ(3, 10, 10);
                                ViewFamilyType vft1 = vft;
                                vs = ViewSection.CreateSection(doc, vft1.Id, newbbox);

                                if (vs != null)
                                {
                                    vs.ViewTemplateId = choose_template;
                                }

                                if (vs != null)
                                {
                                    try
                                    {

                                        vs.Name = form.textBox1.Text + " " + windowelevations.ToArray()[i].Name + " " + i;
                                        viewsec.Add(vs);
                                    }
                                    catch
                                    {
                                        TaskDialog.Show("Error", "View name must be unique");
                                    }
                                }
                                //vs.Scale = 200;
                                doc.Regenerate();
                                TaskDialog.Show("! ", windowelevations.ToArray().Length + " - New Views in the project");
                            }
                        }
                    }
                }
             
                t1.Commit();
            }

            uidoc.ActiveView = vs;
            return Autodesk.Revit.UI.Result.Succeeded;
        }
    }

    [Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)]
    public class CreateSchedule : IExternalCommand
    {

        static AddInId appId = new AddInId(new Guid("BFB59CE4-49D2-4C53-84D8-726E441220DD"));
        public Autodesk.Revit.UI.Result Execute(ExternalCommandData commandData, ref string message, ElementSet elementSet)
        {
            UIDocument uidoc = commandData.Application.ActiveUIDocument;
            Autodesk.Revit.DB.Document doc = uidoc.Document;
            UIApplication uiapp = commandData.Application;
            //---------------------------------------- FILTERS ------------------------------------


            ViewFamilyType vft = new FilteredElementCollector(doc).OfClass(typeof(ViewFamilyType)).Cast<ViewFamilyType>().FirstOrDefault<ViewFamilyType>(x => ViewFamily.Section == x.ViewFamily);
            ViewFamilyType vftele = new FilteredElementCollector(doc).OfClass(typeof(ViewFamilyType)).Cast<ViewFamilyType>().FirstOrDefault<ViewFamilyType>(x => ViewFamily.Elevation == x.ViewFamily);

            FilteredElementCollector viewTemplate = new FilteredElementCollector(doc).OfClass(typeof(Autodesk.Revit.DB.View));
            ICollection<Element> VTcolector = viewTemplate.ToElements();
            
            FilteredElementCollector fec = new FilteredElementCollector(doc);
            fec.OfClass(typeof(ViewSection));
            var viewtem = fec.Cast<ViewSection>().Where<ViewSection>(vp => vp.IsTemplate);
            Form4 form = new Form4();
            ScheduleFieldId fieldid_1 = null;
            ScheduleField foundField = null;
            ScheduleDefinition definition = null;

            //------------------------Walls lists---------------------------------------------------
            List<Element> wallsEle = new List<Element>();
            List<string> _comment = new List<string>();
            List<string> _mark = new List<string>();
            List<ViewSchedule> lista_sch = new List<ViewSchedule>();
            List<ElementId> templaId = new List<ElementId>();

            string debug = null;

            string project_param = /*"Schedule Identifier"*/ "Comments";




            foreach (Element e in new FilteredElementCollector(doc).OfClass(typeof(Wall)))
            {
                try
                {
                    ParameterMap parmap = e.ParametersMap;
                    Parameter par = parmap.get_Item(project_param);
                    string value = par.AsString();

                    if (value != null)
                    {
                        _comment.Add(value);
                    }
                    
                }
                catch (Exception)
                {
                    TaskDialog.Show("!", "Project requires (Schedule Identifier) parameter");
                    return Autodesk.Revit.UI.Result.Cancelled;
                }
            }

           

            foreach (var item in _comment)
            {
                if (item != null)
                {
                    if (!form.comboBox1.Items.Contains(item))
                    {
                        form.comboBox1.Items.Add(item);
                    }
                }
            }

            foreach (var item in viewtem)
            {
                form.comboBox2.Items.Add(item.Name);
                templaId.Add(item.Id);
            }

            if (_comment.Count == 0 )
            {
                TaskDialog.Show("!", "No walls contain a identifier");
                return Autodesk.Revit.UI.Result.Cancelled;
            }

            form.ShowDialog();

            if (form.DialogResult == DialogResult.Cancel)
            {
                return Autodesk.Revit.UI.Result.Cancelled;
            }

            //int selectedindex = form.comboBox1.SelectedIndex;
            //string identifier = _comment.ElementAt(selectedindex).ToString();

            ElementId choose_template;
            try
            {
                int index_tBox = form.comboBox2.SelectedIndex;
                choose_template = templaId.ElementAt(index_tBox);
            }
            catch (Exception)
            {

                TaskDialog.Show("! ", "A template must be selected");
                TaskDialog.Show("! ", "Task Cancelled");
                return Autodesk.Revit.UI.Result.Cancelled;
            }
            

            if (_comment == null)
                form.Close();
            if (form.Equals(false))
            {
                return Autodesk.Revit.UI.Result.Cancelled;
            }

            foreach (Element e in new FilteredElementCollector(doc).OfClass(typeof(Wall)))
            {
                try
                {

                    ParameterMap parmap = e.ParametersMap;
                    Parameter par = parmap.get_Item(project_param);
                    string value = par.AsString();


                    if (form.comboBox1.SelectedItem.ToString() == value)
                    {
                        Element E = e;
                        wallsEle.Add(E);
                    }
                }
                catch (Exception)
                {
                    continue;
                }

            }

            if (form.checkBox1.Checked == true)
            {
                using (Transaction t = new Transaction(doc, "Create single-category"))
                {
                    t.Start();

                    int count = 0;
                    foreach (var item in wallsEle)
                    {


                        try
                        {
                            ViewSchedule vs_ = ViewSchedule.CreateSchedule(doc, new ElementId(BuiltInCategory.OST_Walls));

                            vs_.Name = item.Name + " " + count;
                            definition = vs_.Definition;
                        }
                        catch (Exception)
                        {
                            TaskDialog.Show("! ", "Task failed, check if view name is already in use");
                            TaskDialog.Show("! ", "Task Cancelled");
                            return Autodesk.Revit.UI.Result.Cancelled;
                        }



                        SchedulableField schedulableField = definition.GetSchedulableFields().FirstOrDefault<SchedulableField>(sf => sf.ParameterId == new ElementId(BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS));
                        SchedulableField schedulableField_2 = definition.GetSchedulableFields().FirstOrDefault<SchedulableField>(sf => sf.ParameterId == new ElementId(BuiltInParameter.WALL_BASE_CONSTRAINT));
                        SchedulableField schedulableField_3 = definition.GetSchedulableFields().FirstOrDefault<SchedulableField>(sf => sf.ParameterId == new ElementId(BuiltInParameter.WALL_ATTR_WIDTH_PARAM));
                        SchedulableField schedulableField_4 = definition.GetSchedulableFields().FirstOrDefault<SchedulableField>(sf => sf.ParameterId == new ElementId(BuiltInParameter.ALL_MODEL_MARK));

                        if (schedulableField != null)
                        {

                            definition.AddField(schedulableField);
                            definition.AddField(schedulableField_2);
                            definition.AddField(schedulableField_3);
                            definition.AddField(schedulableField_4);
                        }

                        ElementId paramId = new ElementId(BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS);

                        foreach (ScheduleFieldId fieldId in definition.GetFieldOrder())
                        {
                            foundField = definition.GetField(fieldId);
                            if (foundField.ParameterId == paramId)
                            {
                                fieldid_1 = foundField.FieldId;
                            }
                        }

                        definition.AddFilter(new ScheduleFilter(fieldid_1, ScheduleFilterType.Equal, "IDENTIFIER"));

                        doc.Regenerate();

                        count++;
                    }
                    t.Commit();
                }

            }

            ViewSection vs = null;
            double lenght = 0;
            //------------------------Creating sections for each Wall---------------------------------------------------
            LibraryGeometry libGeo = new LibraryGeometry();

            for (int i = 0; i < wallsEle.ToArray().Length; i++)
            {

                Wall wall = null;
                if (wallsEle.ToArray()[i] != null)

                    wall = wallsEle.ToArray()[i] as Wall;


                XYZ orientation1 = wall.Orientation;
                //data += orientation1.ToString() + "\n";


                BoundingBoxXYZ bb = wall.get_BoundingBox(null);
                double minZ = bb.Min.Z + 3.5;
                double maxZ = bb.Max.Z + 1;

                //double h = maxZ - minZ;
                //Level level = doc.ActiveView.GenLevel;
                //double top = 10 + level.Elevation;
                //double bottom = level.Elevation;

                LocationCurve lc = wall.Location as LocationCurve;
                Autodesk.Revit.DB.Line origWallLine = lc.Curve as Autodesk.Revit.DB.Line;


                CurtainGrid cgrid = wall.CurtainGrid;
                Options options = new Options();
                options.ComputeReferences = true;
                options.IncludeNonVisibleObjects = true;
                options.View = doc.ActiveView;

                GeometryElement geomElem = wall.get_Geometry(options);
                List<Autodesk.Revit.DB.Line> line_list = new List<Autodesk.Revit.DB.Line>();
                List<Autodesk.Revit.DB.Line> ver_line_list = new List<Autodesk.Revit.DB.Line>();


                try
                {
                    foreach (GeometryObject obj in geomElem)
                    {
                        Visibility vis = obj.Visibility;

                        string visString = vis.ToString();


                        Autodesk.Revit.DB.Line line_ = obj as Autodesk.Revit.DB.Line;
                        Solid solid = obj as Solid;

                        if (geomElem.ToArray().First() == obj)
                        {
                            lenght = line_.ApproximateLength;
                        }


                        XYZ dir = new XYZ(0, 0, 1);
                        if (line_ != null)
                        {
                            if (line_.ApproximateLength == lenght)
                            {
                                if (line_.Direction.Z == 1.0)
                                {
                                    line_list.Add(line_);
                                }
                            }
                        }

                        XYZ dir2 = new XYZ(0, 1, 0);
                        if (line_ != null)
                        {
                            if (line_.ApproximateLength == lenght)
                            {
                                if (line_.Direction.Z == 1.0)
                                {
                                    ver_line_list.Add(line_);
                                }
                            }
                        }
                    }
                }
                catch (Exception)
                {

                    TaskDialog.Show("! ", "Wall Selected must to be a curtain wall type");
                    return Autodesk.Revit.UI.Result.Cancelled;
                }

               





                Autodesk.Revit.DB.Curve offsetWallLine = origWallLine.CreateOffset(3, XYZ.BasisZ);
                //double unconnected_height = wall.LookupParameter("unconnected height");


                XYZ p = offsetWallLine.GetEndPoint(0);
                XYZ q = offsetWallLine.GetEndPoint(1);


                XYZ p2 = new XYZ(p.X, p.Y, maxZ + 1);
                XYZ q2 = new XYZ(q.X, q.Y, maxZ + 1);
                XYZ p2_ = new XYZ(p.X, p.Y, maxZ + 1);
                XYZ q2_ = new XYZ(q.X, q.Y, maxZ + 1);
                Autodesk.Revit.DB.Line line = Autodesk.Revit.DB.Line.CreateBound(p2, q2);
                Autodesk.Revit.DB.Line line2_ = Autodesk.Revit.DB.Line.CreateBound(p2_, q2_);
                Autodesk.Revit.DB.Line line2 = Autodesk.Revit.DB.Line.CreateBound(p, p2);

                Autodesk.Revit.DB.Curve curve1 = Autodesk.Revit.DB.Line.CreateBound(p, p2);
                Autodesk.Revit.DB.Curve curve2 = Autodesk.Revit.DB.Line.CreateBound(q, q2);

                Autodesk.Revit.DB.Curve line3_ver = Autodesk.Revit.DB.Line.CreateBound(q, p);
                Autodesk.Revit.DB.Curve line4_ver = Autodesk.Revit.DB.Line.CreateBound(p2, q2);

                Autodesk.Revit.DB.Line ver_line1 = Autodesk.Revit.DB.Line.CreateBound(p, p2);


                double dist = p.DistanceTo(q);
                XYZ v = p - q;
                double halfLength = v.GetLength() / 2;
                double offset = 5; // offset by 3 feet.
                XYZ min = new XYZ(-halfLength, 0, -offset);
                XYZ max = new XYZ(halfLength, 10, offset);
                XYZ midpoint = q + 0.5 * v; // q get lower midpoint. 
                XYZ walldir = v.Normalize();
                XYZ up = XYZ.BasisZ;
                XYZ viewdir = walldir.CrossProduct(up);
                Autodesk.Revit.DB.Transform t = Autodesk.Revit.DB.Transform.Identity;
                t.Origin = midpoint;
                t.BasisX = walldir;
                t.BasisY = up;
                t.BasisZ = viewdir;
                BoundingBoxXYZ sectionBox = new BoundingBoxXYZ();
                sectionBox.Transform = t /*transform*/;
                sectionBox.Min = new XYZ(-halfLength , -0.5, 0);
                sectionBox.Max = new XYZ(halfLength + 0.5 , maxZ, 10);
                ViewFamilyType vft1 = vft;




                using (Transaction tx = new Transaction(doc))
                {
                    try
                    {
                        tx.Start("Create wall Section");
                        vs = ViewSection.CreateSection(doc, vft.Id, sectionBox);
                        vs.Name = form.textBox1.Text + " " + wallsEle.ToArray()[i].Name + " " + i;
                        vs.ViewTemplateId = choose_template;
                        vs.Scale = 100;
                        doc.Regenerate();



                        tx.Commit();
                        TaskDialog.Show("View created", "A section view was created");

                    }
                    catch
                    {

                        TaskDialog.Show("! ", "Task failed, check if view name is already in use");
                        TaskDialog.Show("! ", "Task Cancelled");
                        return Autodesk.Revit.UI.Result.Cancelled;
                    }



                }

               
                List<Autodesk.Revit.DB.DetailCurve> dCurve_list = new List<DetailCurve>();
                ReferenceArray refArray = new ReferenceArray();
                uidoc.ActiveView = vs;
                using (Transaction tx = new Transaction(doc))
                {
                    tx.Start("create dimension by mullion");

                    DetailCurve dCurve1 = null;
                    
                    foreach (var Line in line_list)
                    {
                        if (!doc.IsFamilyDocument)
                        {
                            //Reference gridRef = null;
                            //gridRef = line.Reference;
                            //refArray.Append(gridRef);
                            dCurve1 = doc.Create.NewDetailCurve(doc.ActiveView, Line);
                            dCurve_list.Add(dCurve1);
                        }
                        else
                        {
                            //Reference gridRef = null;
                            //gridRef = line.Reference;
                            //refArray.Append(gridRef);
                            dCurve1 = doc.Create.NewDetailCurve(doc.ActiveView, Line);
                            dCurve_list.Add(dCurve1);
                        }
                    }
                    foreach (var curve_ in dCurve_list)
                    {
                        refArray.Append(curve_.GeometryCurve.Reference);
                    }


                    try
                    {
                        if (!doc.IsFamilyDocument)
                        {
                            doc.Create.NewDimension(
                              doc.ActiveView, line, refArray);
                        }
                        else
                        {
                            doc.FamilyCreate.NewDimension(
                              doc.ActiveView, line, refArray);
                        }
                    }
                    catch (Exception)
                    {

                        TaskDialog.Show("! ", "Task Cancelled");
                        goto finish;
                    }


                    finish: tx.Commit();
                }
                using (Transaction tx = new Transaction(doc))
                {
                    tx.Start("Overhall dimension");

                    DetailCurve dCurve1 = null;
                    DetailCurve dCurve2 = null;
                    //Reference gridRef1 = null;
                    //Reference gridRef2 = null;
                    ReferenceArray refArray2 = new ReferenceArray();

                    if (!doc.IsFamilyDocument)
                    {
                        dCurve1 = doc.Create.NewDetailCurve(doc.ActiveView, curve1);
                        dCurve2 = doc.Create.NewDetailCurve(doc.ActiveView, curve2);
                        //gridRef1 = curve1.Reference;
                    }
                    else
                    {
                        dCurve1 = doc.Create.NewDetailCurve(doc.ActiveView, curve1);
                        dCurve2 = doc.Create.NewDetailCurve(doc.ActiveView, curve2);
                        //gridRef2 = curve2.Reference;
                    }

                    refArray2.Append(/*curve1.Reference*/dCurve1.GeometryCurve.Reference /*gridRef1*/);
                    refArray2.Append(/*curve2.Reference*/dCurve2.GeometryCurve.Reference /*gridRef2*/);

                    if (!doc.IsFamilyDocument)
                    {
                        doc.Create.NewDimension(
                          doc.ActiveView, line2_, refArray2);
                    }
                    else
                    {
                        doc.FamilyCreate.NewDimension(
                          doc.ActiveView, line2_, refArray2);
                    }

                    tx.Commit();
                }
                using (Transaction tx = new Transaction(doc))
                {
                    tx.Start("ver dim");

                    DetailCurve dCurve1 = null;
                    DetailCurve dCurve2 = null;
                    //Reference gridRef1 = null;
                    //Reference gridRef2 = null;
                    ReferenceArray refArray2 = new ReferenceArray();

                    if (!doc.IsFamilyDocument)
                    {
                        dCurve1 = doc.Create.NewDetailCurve(doc.ActiveView, line3_ver);
                        dCurve2 = doc.Create.NewDetailCurve(doc.ActiveView, line4_ver);
                        //gridRef1 = curve1.Reference;
                    }
                    else
                    {
                        dCurve1 = doc.Create.NewDetailCurve(doc.ActiveView, line3_ver);
                        dCurve2 = doc.Create.NewDetailCurve(doc.ActiveView, line4_ver);
                        //gridRef2 = curve2.Reference;
                    }

                    refArray2.Append(/*curve1.Reference*/dCurve1.GeometryCurve.Reference /*gridRef1*/);
                    refArray2.Append(/*curve2.Reference*/dCurve2.GeometryCurve.Reference /*gridRef2*/);

                    if (!doc.IsFamilyDocument)
                    {
                        doc.Create.NewDimension(
                          doc.ActiveView, ver_line1, refArray2);
                    }
                    else
                    {
                        doc.FamilyCreate.NewDimension(
                          doc.ActiveView, ver_line1, refArray2);
                    }
                    tx.Commit();
                }
            }
            //foreach (Element e in wallsEle)
            //{


            //    uidoc.ActiveView = vs;
            //    List<Autodesk.Revit.DB.DetailCurve> dCurve_list = new List<DetailCurve>();
            //    ReferenceArray refArray = new ReferenceArray();

            //    using (Transaction tx = new Transaction(doc))
            //    {
            //        tx.Start("create dimension by mullion");

            //        DetailCurve dCurve1 = null;

            //        foreach (var Line in line_list)
            //        {
            //            if (!doc.IsFamilyDocument)
            //            {
            //                //Reference gridRef = null;
            //                //gridRef = line.Reference;
            //                //refArray.Append(gridRef);
            //                dCurve1 = doc.Create.NewDetailCurve(doc.ActiveView, Line);
            //                dCurve_list.Add(dCurve1);
            //            }
            //            else
            //            {
            //                //Reference gridRef = null;
            //                //gridRef = line.Reference;
            //                //refArray.Append(gridRef);
            //                dCurve1 = doc.Create.NewDetailCurve(doc.ActiveView, Line);
            //                dCurve_list.Add(dCurve1);
            //            }
            //        }
            //        foreach (var curve_ in dCurve_list)
            //        {
            //            refArray.Append(curve_.GeometryCurve.Reference);
            //        }



            //        if (!doc.IsFamilyDocument)
            //        {
            //            doc.Create.NewDimension(
            //              doc.ActiveView, line, refArray);
            //        }
            //        else
            //        {
            //            doc.FamilyCreate.NewDimension(
            //              doc.ActiveView, line, refArray);
            //        }

            //        tx.Commit();
            //    }
            //    using (Transaction tx = new Transaction(doc))
            //    {
            //        tx.Start("Overhall dimension");

            //        DetailCurve dCurve1 = null;
            //        DetailCurve dCurve2 = null;
            //        //Reference gridRef1 = null;
            //        //Reference gridRef2 = null;
            //        ReferenceArray refArray2 = new ReferenceArray();

            //        if (!doc.IsFamilyDocument)
            //        {
            //            dCurve1 = doc.Create.NewDetailCurve(doc.ActiveView, curve1);
            //            dCurve2 = doc.Create.NewDetailCurve(doc.ActiveView, curve2);
            //            //gridRef1 = curve1.Reference;
            //        }
            //        else
            //        {
            //            dCurve1 = doc.Create.NewDetailCurve(doc.ActiveView, curve1);
            //            dCurve2 = doc.Create.NewDetailCurve(doc.ActiveView, curve2);
            //            //gridRef2 = curve2.Reference;
            //        }

            //        refArray2.Append(/*curve1.Reference*/dCurve1.GeometryCurve.Reference /*gridRef1*/);
            //        refArray2.Append(/*curve2.Reference*/dCurve2.GeometryCurve.Reference /*gridRef2*/);

            //        if (!doc.IsFamilyDocument)
            //        {
            //            doc.Create.NewDimension(
            //              doc.ActiveView, line2_, refArray2);
            //        }
            //        else
            //        {
            //            doc.FamilyCreate.NewDimension(
            //              doc.ActiveView, line2_, refArray2);
            //        }

            //        tx.Commit();
            //    }
            //    using (Transaction tx = new Transaction(doc))
            //    {
            //        tx.Start("ver dim");

            //        DetailCurve dCurve1 = null;
            //        DetailCurve dCurve2 = null;
            //        //Reference gridRef1 = null;
            //        //Reference gridRef2 = null;
            //        ReferenceArray refArray2 = new ReferenceArray();

            //        if (!doc.IsFamilyDocument)
            //        {
            //            dCurve1 = doc.Create.NewDetailCurve(doc.ActiveView, line3_ver);
            //            dCurve2 = doc.Create.NewDetailCurve(doc.ActiveView, line4_ver);
            //            //gridRef1 = curve1.Reference;
            //        }
            //        else
            //        {
            //            dCurve1 = doc.Create.NewDetailCurve(doc.ActiveView, line3_ver);
            //            dCurve2 = doc.Create.NewDetailCurve(doc.ActiveView, line4_ver);
            //            //gridRef2 = curve2.Reference;
            //        }

            //        refArray2.Append(/*curve1.Reference*/dCurve1.GeometryCurve.Reference /*gridRef1*/);
            //        refArray2.Append(/*curve2.Reference*/dCurve2.GeometryCurve.Reference /*gridRef2*/);

            //        if (!doc.IsFamilyDocument)
            //        {
            //            doc.Create.NewDimension(
            //              doc.ActiveView, ver_line1, refArray2);
            //        }
            //        else
            //        {
            //            doc.FamilyCreate.NewDimension(
            //              doc.ActiveView, ver_line1, refArray2);
            //        }
            //        tx.Commit();
            //    }
            //}
            return Autodesk.Revit.UI.Result.Succeeded;
        }
    }

    [Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)]
    public class room_data : IExternalCommand
    {
        public void WriteTextFile(string line)
        {
            string filename2 = "";
            System.Windows.Forms.OpenFileDialog openDialog = new System.Windows.Forms.OpenFileDialog();
            openDialog.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
            //openDialog.Filter = "Excel Files (*.xlsx) |*.xslx)"; // TODO: Change to .csv
            if (openDialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                filename2 = openDialog.FileName;
            }

            //string tempPath = Path.GetTempPath();

            //string myDocs = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
            //string file = "test.txt";

            //string filename = Path.Combine(tempPath, file);

            if (File.Exists(filename2))
            {
                File.Delete(filename2);
            }
            using (StreamWriter writer = new StreamWriter(filename2))
            {
                writer.WriteLine(line);
                
            }
        }

        static AddInId appId = new AddInId(new Guid("5F44AA78-A137-6509-AAF8-A478F3B24BAB"));
        public Autodesk.Revit.UI.Result Execute(ExternalCommandData commandData, ref string message, ElementSet elementSet)
        {
            UIDocument uidoc = commandData.Application.ActiveUIDocument;
            Autodesk.Revit.DB.Document doc = uidoc.Document;

            FilteredElementCollector rooms1 = new FilteredElementCollector(doc).WhereElementIsNotElementType().OfClass(typeof(SpatialElement));
            ICollection<Element> room2 = rooms1.ToElements();
            SpatialElementBoundaryOptions opt = new SpatialElementBoundaryOptions();
            opt.SpatialElementBoundaryLocation = SpatialElementBoundaryLocation.Center;
            //Level level = doc.GetElement(uidoc.Selection.PickObject(ObjectType.Element, "Select level")) as Level;
            List<List<XYZ>> lista_nombres = new List<List<XYZ>>();
            List<string> names = new List<string>();

            string filename2 = "";
            System.Windows.Forms.OpenFileDialog openDialog = new System.Windows.Forms.OpenFileDialog();
            openDialog.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
            //openDialog.Filter = "Excel Files (*.xlsx) |*.xslx)"; // TODO: Change to .csv
            if (openDialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                filename2 = openDialog.FileName;
            }

            List<Room> rooms = new List<Room>();

            foreach (Room r in new FilteredElementCollector(doc).OfClass(typeof(SpatialElement)).OfCategory(BuiltInCategory.OST_Rooms).Cast<Room>().Where(q => q.Area > 0))
            {
                rooms.Add(r);
            }

            List<XYZ> ptlist = new List<XYZ>();
            


            using (Transaction tx = new Transaction(doc))
            {
                tx.Start("ver dim");

                using (StreamWriter writer = new StreamWriter(filename2))
                {
                    foreach (var r in rooms)
                    {

                        GeometryElement geo = r.ClosedShell;

                        Options op = new Options();
                        op.ComputeReferences = true;


                        foreach (var item in r.get_Geometry(op).Where(q => q is Solid).Cast<Solid>())
                        {
                            foreach (Face item3 in item.Faces)
                            {
                                //PlanarFace planarFace = item3 as PlanarFace;
                                //XYZ normal = planarFace.ComputeNormal(new UV(planarFace.Origin.X, planarFace.Origin.Y));


                                Element e = doc.GetElement(item3.Reference);
                                GeometryObject geoobj = e.GetGeometryObjectFromReference(item3.Reference);
                                Face face = geoobj as Face;


                                foreach (var edges in face.GetEdgesAsCurveLoops() /*face.GetEdgesAsCurveLoops()*/)
                                {

                                    foreach (Autodesk.Revit.DB.Curve edge in edges)
                                    {
                                        XYZ testPoint1 = edge.GetEndPoint(1);
                                        XYZ testPoint2 = edge.GetEndPoint(0);
                                        double lenght = Math.Round(edge.ApproximateLength, 0);
                                        double lenght_convert = UnitUtils.Convert(lenght, DisplayUnitType.DUT_DECIMAL_FEET, DisplayUnitType.DUT_MILLIMETERS);

                                        double x = Math.Round(testPoint1.X, 0);
                                        double y = Math.Round(testPoint1.Y, 0);
                                        double z = Math.Round(testPoint1.Z, 0);

                                        ElementClassFilter filter = new ElementClassFilter(typeof(Floor));
                                        XYZ dir = new XYZ(0, 0, 0) - testPoint1;



                                        string x0 = edge.GetEndPoint(0).X.ToString();
                                        string y0 = edge.GetEndPoint(0).Y.ToString();
                                        string z0 = edge.GetEndPoint(0).Z.ToString();

                                        string x1 = edge.GetEndPoint(1).X.ToString();
                                        string y1 = edge.GetEndPoint(1).Y.ToString();
                                        string z1 = edge.GetEndPoint(1).Z.ToString();
                                        writer.WriteLine(x0 + "," + y0 + "," + z0 + "," + x1 + "," + y1 + "," + z1);

                                    }
                                }
                            }
                        }

                        //IList<IList<BoundarySegment>> loops = r.GetBoundarySegments(opt);

                        //for (int i = 0; i < loops.Count ; i++)
                        //{
                        //    foreach (var item in loops.ToArray()[i])
                        //    {
                               
                                

                        //        Autodesk.Revit.DB.Curve cr = item.GetCurve();
                        //        //ptlist.Add(cr.GetEndPoint(0));
                        //        //ptlist.Add(cr.GetEndPoint(1));

                        //        string x0 = cr.GetEndPoint(0).X.ToString();
                        //        string y0 = cr.GetEndPoint(0).Y.ToString();
                        //        string z0 = cr.GetEndPoint(0).Z.ToString();

                        //        string x1 = cr.GetEndPoint(1).X.ToString();
                        //        string y1 = cr.GetEndPoint(1).Y.ToString();
                        //        string z1 = cr.GetEndPoint(1).Z.ToString();

                        //        string h_z1 = geo.GetBoundingBox().Max.Z.ToString();

                        //        writer.WriteLine(x0 + "," + y0 + "," + z0 + "," + x1 + "," + y1 + "," + z1 );



                        //    }
                        //}
                    }
                }
                //Autodesk.Revit.DB.Curve curve1 = Autodesk.Revit.DB.Line.CreateBound(cr.GetEndPoint(0), cr.GetEndPoint(1)) as Autodesk.Revit.DB.Curve;
                //DetailLine line = doc.Create.NewDetailCurve(doc.ActiveView, curve1) as DetailLine;
                tx.Commit();
            }
            return Autodesk.Revit.UI.Result.Succeeded;
        }
    }

    [Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)]
    public class Wall_Bounding_room : IExternalCommand
    {
        static AddInId appId = new AddInId(new Guid("5F46AA78-A136-6509-AAF8-A478F3B24BAB"));
        public Autodesk.Revit.UI.Result Execute(ExternalCommandData commandData, ref string message, ElementSet elementSet)
        {
            UIDocument uidoc = commandData.Application.ActiveUIDocument;
            Autodesk.Revit.DB.Document doc = uidoc.Document;

            List<Element> ele = new List<Element>();

            foreach (Element e in new FilteredElementCollector(doc).OfClass(typeof(Wall)))
            {
                try
                {
                    ParameterMap parmap = e.ParametersMap;
                    Parameter par = parmap.get_Item("Room Bounding");
                    string value = par.AsString();

                    Parameter value2 = e.LookupParameter("Room Bounding");
                    string value3 = value2.Definition.Name;

                    var value_ = value2.AsValueString();

                  
                    if (value_ == "No")
                    {
                        ele.Add(e);
                    }
                }
                catch (Exception)
                {
                    //    TaskDialog.Show("!", "Project requires (Schedule Identifier) parameter");
                    return Autodesk.Revit.UI.Result.Cancelled;
                }
            }


            TaskDialog.Show("!", ele.Count.ToString() + " Wall were selected" );
            uidoc.Selection.SetElementIds(ele.Select(q => q.Id).ToList());

            return Autodesk.Revit.UI.Result.Succeeded;
        }
    }

    [Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)]
    public class revision_on_project : IExternalCommand
    {
        private Autodesk.Revit.DB.Form makeForm(Autodesk.Revit.DB.Document doc, XYZ pt1, XYZ pt2, XYZ pt3)
        {
            Autodesk.Revit.DB.Form form = null;

            XYZ u = pt2.Subtract(pt1);
            XYZ v = pt3.Subtract(pt1);
            double area = u.CrossProduct(v).GetLength() / 2;
            if (area < 10)
                return null;

            ReferenceArray ra = new ReferenceArray();
            ra.Append(MakeCuveByPoints(doc, pt1, pt2).GeometryCurve.Reference);
            ra.Append(MakeCuveByPoints(doc, pt2, pt3).GeometryCurve.Reference);
            ra.Append(MakeCuveByPoints(doc, pt3, pt1).GeometryCurve.Reference);

            form = doc.FamilyCreate.NewFormByCap(true, ra);

            return form;
        }

        private CurveByPoints MakeCuveByPoints(Autodesk.Revit.DB.Document doc, XYZ ptA, XYZ ptB)
        {
            ReferencePointArray rpa = new ReferencePointArray();
            rpa.Append(doc.FamilyCreate.NewReferencePoint(ptA));
            rpa.Append(doc.FamilyCreate.NewReferencePoint(ptB));
            return doc.FamilyCreate.NewCurveByPoints(rpa);
        }

        private bool isXYZinList(XYZ point, IList<XYZ> myList)
        {
            foreach  (XYZ xyz in myList)
            {
                if (point.IsAlmostEqualTo(xyz))
                {
                    return true;
                }
            }
            return false;
        }

        private IList<XYZ> sortpoints(IList<XYZ> input)
        {
            IList<XYZ> ret = new List<XYZ>();
            foreach (var xyz in input)
            {
                if (ret.Count == 0)
                {
                    ret.Add(xyz);
                    continue;
                }

                XYZ nearestXYZ = null;
                double nearestDistance = Double.PositiveInfinity;

                foreach (XYZ two in input)
                {
                    if (isXYZinList(two, ret))
                    {
                        continue;
                    }
                    double thisDist = two.DistanceTo(ret.Last());
                    if (thisDist < 0.01)
                    {
                        continue;
                    }
                    if (thisDist < nearestDistance)
                    {
                        nearestDistance = thisDist;
                        nearestXYZ = two;
                    }
                }
                ret.Add(nearestXYZ);
                    
            }
            return ret;
        }

        public static ModelLine Makeline(Autodesk.Revit.DB.Document doc, XYZ pta, XYZ ptb)
        {
            ModelLine modelLine = null;
            double distance = pta.DistanceTo(ptb);
            if (distance < 0.01)
            {
                TaskDialog.Show("Error", "Distance" + distance);
                return modelLine;
            }

            XYZ norm = pta.CrossProduct(ptb);
            if (norm.GetLength() == 0)
            {
                XYZ aSubB = pta.Subtract(ptb);
                XYZ aSubBcrossz = aSubB.CrossProduct(XYZ.BasisZ);
                double crosslenght = aSubBcrossz.GetLength();
                if (crosslenght == 0)
                {
                    norm = XYZ.BasisY;
                }
                else
                {
                    norm = XYZ.BasisZ;
                }
            }

            Autodesk.Revit.DB.Plane plane = Autodesk.Revit.DB.Plane.CreateByNormalAndOrigin(norm, ptb);


            SketchPlane skplane = SketchPlane.Create(doc, plane);

            Autodesk.Revit.DB.Line line = Autodesk.Revit.DB.Line.CreateBound(pta, ptb);

            if (doc.IsFamilyDocument)
            {
                modelLine = doc.FamilyCreate.NewModelCurve(line, skplane) as ModelLine;
            }
            else
            {
                modelLine = doc.Create.NewModelCurve(line, skplane) as ModelLine;
            }
            if (modelLine == null)
            {
                TaskDialog.Show("Error", "Model line = null");
            }
            return modelLine;
        }


        static AddInId appId = new AddInId(new Guid("5F44BB78-A137-6509-AAF8-A478F3B24BAB"));
        public Autodesk.Revit.UI.Result Execute(ExternalCommandData commandData, ref string message, ElementSet elementSet)
        {

            
            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Autodesk.Revit.DB.Document doc = uidoc.Document;
            Autodesk.Revit.ApplicationServices.Application app = uiapp.Application;


            //FilteredElementCollector revfilter = new FilteredElementCollector(doc).WhereElementIsNotElementType().OfClass(typeof(Revision));
            //ICollection<Element> revisions = revfilter.ToElements();

            //ViewSheet activeViewSheet = doc.ActiveView as ViewSheet;

            //FilteredElementCollector col_sheetele = new FilteredElementCollector(doc, activeViewSheet.Id);
            //var scheduleSheetInstances = col_sheetele.OfClass(typeof(ScheduleSheetInstance)).ToElements().OfType<ScheduleSheetInstance>();

            //IList<ElementId> rev_Id = activeViewSheet.GetAllRevisionIds();

            //// get list of all revisions in the document
            //ICollection<ElementId> elementIds = new FilteredElementCollector(doc)
            //    .OfClass(typeof(Revision))
            //    .ToElementIds();

            //// Remove the first revision from the list
            //// Revit must have one revision in the document, so we can't delete them all
            //elementIds.Remove(elementIds.First());

            //using (Transaction t = new Transaction(doc, "Delete Revisions"))
            //{
            //    t.Start();
            //    doc.Delete(/*elementIds*/  rev_Id);
            //    t.Commit();
            //}

            Form27 form = new Form27();
            form.ShowDialog();

            string st = "";

            IList<XYZ> points = new List<XYZ>();

            if (form.radioButton1.Checked)
            {
                try
                {
                    MessageBox.Show("Please select a topography", "!");

                    TopographySurface ts = doc.GetElement(uidoc.Selection.PickObject(ObjectType.Element, "select topo")) as TopographySurface;
                    //Floor floor = doc.GetElement(uidoc.Selection.PickObject(ObjectType.Element, "select floor")) as Floor;
                    //IEnumerable<XYZ> points = ts.GetPoints().Where(q => ts.IsBoundaryPoint(q));

                    points = ts.GetBoundaryPoints();

                    using (Transaction t = new Transaction(doc, "Create Topo Boundary Lines"))
                    {
                        t.Start();
                        XYZ prev = null;
                        foreach (XYZ point in points)
                        {
                            XYZ pt1 = null;
                            XYZ pt2 = null;

                            if (prev == null)
                            {
                                pt1 = points.First();
                                pt2 = points.Last();
                            }
                            else
                            {
                                pt1 = prev;
                                pt2 = point;
                            }

                            Makeline(doc, pt1, pt2);

                            //Autodesk.Revit.DB.Line line = Autodesk.Revit.DB.Line.CreateBound(pt1, pt2);
                            //XYZ v = pt1 - pt2;
                            //double dxy = Math.Abs(v.X) + Math.Abs(v.Y);
                            //XYZ w = (dxy > 0.0001) ? XYZ.BasisZ : XYZ.BasisY;
                            //XYZ norm = v.CrossProduct(w).Normalize();
                            //Autodesk.Revit.DB.Plane plane = Autodesk.Revit.DB.Plane.CreateByNormalAndOrigin(norm, pt2);
                            //SketchPlane skplane = SketchPlane.Create(doc, plane);
                            ////SketchPlane skplane = doc.Create.NewSketchPlane(app.Create.NewPlane(norm, pt2));
                            //ModelCurve mc = doc.Create.NewModelCurve(line, skplane);

                            prev = point;
                        }
                        t.Commit();
                    }




                    IList<XYZ> ptboundary = sortpoints(ts.GetBoundaryPoints());
                    List<XYZ> pts = new List<XYZ>();

                    //foreach (var item in sortpoints(ts.GetBoundaryPoints()))
                    //{
                    //    ptboundary.Add(item);
                    //}
                    foreach (var item in ts.GetPoints())
                    {
                        pts.Add(item);
                    }

                    Floor floor = null;

                    using (Transaction T = new Transaction(doc, "Edit Floor"))
                    {
                        T.Start();

                        CurveArray ca = new CurveArray();
                        XYZ prev = null;
                        foreach (var pt in ptboundary)
                        {
                            if (prev == null)
                            {
                                prev = ptboundary.Last();
                            }

                            Autodesk.Revit.DB.Line line = Autodesk.Revit.DB.Line.CreateBound(prev, pt);
                            ca.Append(line);
                            prev = pt;
                        }
                        floor = doc.Create.NewFloor(ca, false);


                        T.Commit();
                    }
                    using (Transaction T = new Transaction(doc, "Edit Floor"))
                    {
                        T.Start();

                        SlabShapeEditor ed = floor.SlabShapeEditor;
                        foreach (var xyz in pts)
                        {
                            ed.DrawPoint(xyz);

                        }
                        T.Commit();
                    }
                }
                catch (Exception)
                {

                    MessageBox.Show("Points couldn't generate geometry ", "Cancelled");
                    return Autodesk.Revit.UI.Result.Cancelled;
                }


            }
            if (form.radioButton2.Checked)
            {
                try
                {
                    MessageBox.Show("Please select a topography", "!");

                    TopographySurface ts = doc.GetElement(uidoc.Selection.PickObject(ObjectType.Element, "select topo")) as TopographySurface;

                    MessageBox.Show("Please select an existing floor", "!");
                    Floor floor = doc.GetElement(uidoc.Selection.PickObject(ObjectType.Element, "select floor")) as Floor;

                    IList<XYZ> ptboundary = sortpoints(ts.GetBoundaryPoints());
                    List<XYZ> pts = new List<XYZ>();


                    foreach (var item in ts.GetPoints())
                    {
                        pts.Add(item);
                    }


                    using (Transaction T = new Transaction(doc, "Edit Floor"))
                    {
                        T.Start();

                        SlabShapeEditor ed = floor.SlabShapeEditor;
                        foreach (var xyz in pts)
                        {
                            ed.DrawPoint(xyz);

                        }
                        T.Commit();
                    }
                }
                catch (Exception)
                {

                    MessageBox.Show("Points couldn't generate geometry ", "Cancelled");
                    return Autodesk.Revit.UI.Result.Cancelled;
                }

                //TopographySurface topo = doc.GetElement(uidoc.Selection.PickObject(ObjectType.Element)) as TopographySurface;

                //Autodesk.Revit.DB.Mesh mesh = topo.get_Geometry(new Options()).First(q => q is Autodesk.Revit.DB.Mesh) as Autodesk.Revit.DB.Mesh;

                //Autodesk.Revit.DB.Document famDoc = app.NewFamilyDocument(@"C:\ProgramData\Autodesk\RVT 2019\Family Templates\English\Conceptual Mass\Metric Mass.rft");

                //using (Transaction t = new Transaction(famDoc, "Create massing surfaces"))
                //{
                //    t.Start();

                //    for (int i = 0; i < mesh.NumTriangles; i++)
                //    {
                //        MeshTriangle mt = mesh.get_Triangle(i);
                //        makeForm(famDoc, mt.get_Vertex(0), mt.get_Vertex(1), mt.get_Vertex(2));

                //        if (i > 0 && i % 100 == 0)
                //        {
                //            TaskDialog td = new TaskDialog("Form Counter");
                //            td.CommonButtons = TaskDialogCommonButtons.Yes | TaskDialogCommonButtons.No;
                //            td.MainInstruction = i + " out of " + mesh.NumTriangles + " triangles processed. Do you want to continue?";
                //            if (td.Show() == TaskDialogResult.No)
                //                break;
                //        }
                //    }
                //    t.Commit();
                //}
                //famDoc.LoadFamily(doc);

            }
            return Autodesk.Revit.UI.Result.Succeeded;


        }
    }

  

    [Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)]
    public class make_line_from_surface_normal : IExternalCommand
    {
        public static ModelLine Makeline(Autodesk.Revit.DB.Document doc, XYZ pta, XYZ ptb, bool click)
        {

            Autodesk.Revit.DB.Line line2 = Autodesk.Revit.DB.Line.CreateBound(pta, ptb);

            ModelLine modelLine = null;
            double distance = pta.DistanceTo(ptb);
            if (distance < 0.01)
            {
                TaskDialog.Show("Error", "Distance" + distance);
                return modelLine;
            }

            XYZ norm = pta.CrossProduct(ptb);
            if (norm.GetLength() == 0)
            {
                XYZ aSubB = pta.Subtract(ptb);
                XYZ aSubBcrossz = null;
                if (click)
                {
                    aSubBcrossz = aSubB.CrossProduct(XYZ.BasisX);
                   
                }

                if (click == false)
                {
                    aSubBcrossz = aSubB.CrossProduct(XYZ.BasisZ);
                }

                
                double crosslenght = aSubBcrossz.GetLength();
                if (crosslenght == 0)
                {
                    norm = XYZ.BasisY;
                }
                else
                {
                    norm = XYZ.BasisZ;
                }
            }
            Autodesk.Revit.DB.Line line = null;

            //Autodesk.Revit.DB.Plane plane = Autodesk.Revit.DB.Plane.CreateByNormalAndOrigin(norm, ptb);
            if (click)
            {
                line = Autodesk.Revit.DB.Line.CreateUnbound(ptb, XYZ.BasisY);

                if (line2.Direction.X == 1 || line2.Direction.X == -1)
                {
                    line = Autodesk.Revit.DB.Line.CreateUnbound(ptb, XYZ.BasisY );
                }
              

                if (line2.Direction.Y == 1 || line2.Direction.Y == -1)
                {
                    line = Autodesk.Revit.DB.Line.CreateUnbound(ptb, XYZ.BasisX);
                }
                
            }
            else
            {
                line = Autodesk.Revit.DB.Line.CreateUnbound(ptb, XYZ.BasisZ);
            }


            Autodesk.Revit.DB.Plane pl = Autodesk.Revit.DB.Plane.CreateByThreePoints(pta,
          line.Evaluate(5, false), ptb);
            
            SketchPlane skplane = SketchPlane.Create(doc, pl);
            
            if (doc.IsFamilyDocument)
            {
                modelLine = doc.FamilyCreate.NewModelCurve(line2, skplane) as ModelLine;
            }
            else
            {
                modelLine = doc.Create.NewModelCurve(line2, skplane) as ModelLine;
            }
            if (modelLine == null)
            {
                TaskDialog.Show("Error", "Model line = null");
            }
            return modelLine;
        }

        static AddInId appId = new AddInId(new Guid("5F56AA78-A136-6509-AAF8-A478F3B24BAB"));
        public Autodesk.Revit.UI.Result Execute(ExternalCommandData commandData, ref string message, ElementSet elementSet)
        {
            UIDocument uidoc = commandData.Application.ActiveUIDocument;
            Autodesk.Revit.DB.Document doc = uidoc.Document;

            //ObjectSnapTypes snapTypes = ObjectSnapTypes.Endpoints | ObjectSnapTypes.Intersections;
            //XYZ point = uidoc.Selection.PickPoint(snapTypes, "Select an end point or intersection");

            var pipeTypes = new FilteredElementCollector(doc).OfClass(typeof(SpotDimension)).OfType<SpotDimension>().ToList();

            Form1 form = new Form1();
            

            Reference r = uidoc.Selection.PickObject(ObjectType.Face, "Please pick a point on a " + "face for family instance insertion");

            Element e = doc.GetElement(r.ElementId);
            GeometryObject obj
              = e.GetGeometryObjectFromReference(r);

            XYZ p = r.GlobalPoint;
            XYZ v = null;

            form.ShowDialog();
            

            try
            {
                PlanarFace face = obj as PlanarFace;
                v = face.FaceNormal;
                if (v.IsZeroLength())
                {
                    v = face.FaceNormal.CrossProduct(XYZ.BasisX);
                }
            }
            catch (Exception)
            {
            }
            try
            {
                CylindricalFace face = obj as CylindricalFace;

               
               v = face.Axis.CrossProduct(XYZ.BasisZ);
                if (v.IsZeroLength())
                {
                    v = face.Axis.CrossProduct(XYZ.BasisX);
                }
            }
            catch (Exception)
            {
            }

            Autodesk.Revit.DB.Line line = Autodesk.Revit.DB.Line.CreateUnbound(p, v);
           
            
            double movenumber;
            Double.TryParse(form.textBox1.Text, out movenumber);


            double rotnumber;
            Double.TryParse(form.textBox2.Text, out rotnumber);

            XYZ vect1 = line.Direction * (movenumber / 304.8);

            XYZ vect2 = vect1 + p;

            Autodesk.Revit.DB.Line line2 = Autodesk.Revit.DB.Line.CreateUnbound(p, p.CrossProduct(line.Evaluate(5,false)));

            using (Transaction tx = new Transaction(doc))
            {
                tx.Start("Make Line on face from input");

                ModelLine ml = Makeline(doc, p, vect2, form.radioButton1.Checked);

                //GeometryCreationUtilities.CreateLoftGeometry()
                
                XYZ cc = new XYZ(p.X, p.Y + 10, p.Z );
                Autodesk.Revit.DB.Line axis = Autodesk.Revit.DB.Line.CreateBound(p, cc);
                //ModelLine ml2 = Makeline(doc, p, cc);
                
                ElementTransformUtils.RotateElement(doc, ml.Id, axis, rotnumber);
                
                
                //XYZ bend = p2.Add(new XYZ(2, 2, 0));
                //XYZ end = p2.Add(new XYZ(3, 2, 0));
                //doc.Create.NewSpotElevation(doc.ActiveView, myRef2, p2, bend, end, p2, true);
                
                tx.Commit();
            }
            //uidoc.Selection.SetElementIds(ele.Select(q => q.Id).ToList());
            return Autodesk.Revit.UI.Result.Succeeded;
        }
    }

    [Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)]
    public class make_pipe_by_line : IExternalCommand
    {
        private XYZ intersect(XYZ point, XYZ direction, Autodesk.Revit.DB.Curve curve)
        {
            Autodesk.Revit.DB.Line unbound = Autodesk.Revit.DB.Line.CreateUnbound(new XYZ(point.X, point.Y, curve.GetEndPoint(0).Z), direction);
            IntersectionResultArray ira = null;
            unbound.Intersect(curve, out ira);
            if (ira == null)
            {
                TaskDialog td = new TaskDialog("Error");
                td.MainInstruction = "no intersection";
                td.MainContent = point.ToString() + Environment.NewLine + direction.ToString();
                td.Show();

                return null;
            }
            IntersectionResult ir = ira.get_Item(0);
            return ir.XYZPoint;
        }
        public static ModelLine Makeline(Autodesk.Revit.DB.Document doc, XYZ pta, XYZ ptb)
        {
            ModelLine modelLine = null;
            double distance = pta.DistanceTo(ptb);
            if (distance < 0.01)
            {
                TaskDialog.Show("Error", "Distance" + distance);
                return modelLine;
            }

            XYZ norm = pta.CrossProduct(ptb);
            if (norm.GetLength() == 0)
            {
                XYZ aSubB = pta.Subtract(ptb);
                XYZ aSubBcrossz = aSubB.CrossProduct(XYZ.BasisZ);
                double crosslenght = aSubBcrossz.GetLength();
                if (crosslenght == 0)
                {
                    norm = XYZ.BasisY;
                }
                else
                {
                    norm = XYZ.BasisZ;
                }
            }

            Autodesk.Revit.DB.Plane plane = Autodesk.Revit.DB.Plane.CreateByNormalAndOrigin(norm, ptb);


            SketchPlane skplane = SketchPlane.Create(doc, plane);

            Autodesk.Revit.DB.Line line = Autodesk.Revit.DB.Line.CreateBound(pta, ptb);

            if (doc.IsFamilyDocument)
            {
                modelLine = doc.FamilyCreate.NewModelCurve(line, skplane) as ModelLine;
            }
            else
            {
                modelLine = doc.Create.NewModelCurve(line, skplane) as ModelLine;
            }
            if (modelLine == null)
            {
                TaskDialog.Show("Error", "Model line = null");
            }
            return modelLine;
        }

        static AddInId appId = new AddInId(new Guid("5F56AA78-A136-6509-AAF8-A478F3B24BAB"));
        public Autodesk.Revit.UI.Result Execute(ExternalCommandData commandData, ref string message, ElementSet elementSet)
        {
            UIDocument uidoc = commandData.Application.ActiveUIDocument;
            Autodesk.Revit.DB.Document doc = uidoc.Document;

            List<Autodesk.Revit.DB.Curve> crvs = new List<Autodesk.Revit.DB.Curve>();

            ICollection<Reference> my_lines = uidoc.Selection.PickObjects(Autodesk.Revit.UI.Selection.ObjectType.Element, "Select lines");
            foreach (var item_myRefWall in my_lines)
            {
                Element ele = doc.GetElement(item_myRefWall);
                GeometryObject geoobj = ele.GetGeometryObjectFromReference(item_myRefWall);
                Face face = geoobj as Face;
                LocationCurve locationCurve2 = ele.Location as LocationCurve;

                crvs.Add(locationCurve2.Curve);
            }
            
            var mepSystemTypes = new FilteredElementCollector(doc).OfClass(typeof(PipingSystemType)).OfType<PipingSystemType>().ToList();
            
            var domesticHotWaterSystemType = mepSystemTypes.FirstOrDefault(st => st.SystemClassification ==MEPSystemClassification.DomesticHotWater);

            if (domesticHotWaterSystemType == null)
            {
                message = "Could not found Domestic Hot Water System Type";
                return Result.Failed;
            }
            
            var pipeTypes = new FilteredElementCollector(doc).OfClass(typeof(PipeType)).OfType<PipeType>().ToList();
            
            var firstPipeType = pipeTypes.FirstOrDefault();

            if (firstPipeType == null)
            {
                message = "Could not found Pipe Type";
                return Result.Failed;
            }
            
            FilteredElementCollector collector2 = new FilteredElementCollector(doc);
            collector2.OfClass(typeof(Level));
            Level lv = collector2.First() as Level;

            if (lv == null)
            {
                message = "Wrong Active View";
                return Result.Failed;
            }

            using (Transaction t = new Transaction(doc, "Create  pipe by line"))
            {
                t.Start();


                foreach (var crv in crvs)
                {
                    var pipe = Pipe.Create(doc, domesticHotWaterSystemType.Id, firstPipeType.Id,
                      lv.Id, crv.GetEndPoint(0), crv.GetEndPoint(1));
                }
                
                t.Commit();
            }
            return Autodesk.Revit.UI.Result.Succeeded;
        }
    }

    [Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)]
    public class make_duck_by_line : IExternalCommand
    {
        private XYZ intersect(XYZ point, XYZ direction, Autodesk.Revit.DB.Curve curve)
        {
            Autodesk.Revit.DB.Line unbound = Autodesk.Revit.DB.Line.CreateUnbound(new XYZ(point.X, point.Y, curve.GetEndPoint(0).Z), direction);
            IntersectionResultArray ira = null;
            unbound.Intersect(curve, out ira);
            if (ira == null)
            {
                TaskDialog td = new TaskDialog("Error");
                td.MainInstruction = "no intersection";
                td.MainContent = point.ToString() + Environment.NewLine + direction.ToString();
                td.Show();

                return null;
            }
            IntersectionResult ir = ira.get_Item(0);
            return ir.XYZPoint;
        }
        public static ModelLine Makeline(Autodesk.Revit.DB.Document doc, XYZ pta, XYZ ptb)
        {
            ModelLine modelLine = null;
            double distance = pta.DistanceTo(ptb);
            if (distance < 0.01)
            {
                TaskDialog.Show("Error", "Distance" + distance);
                return modelLine;
            }

            XYZ norm = pta.CrossProduct(ptb);
            if (norm.GetLength() == 0)
            {
                XYZ aSubB = pta.Subtract(ptb);
                XYZ aSubBcrossz = aSubB.CrossProduct(XYZ.BasisZ);
                double crosslenght = aSubBcrossz.GetLength();
                if (crosslenght == 0)
                {
                    norm = XYZ.BasisY;
                }
                else
                {
                    norm = XYZ.BasisZ;
                }
            }

            Autodesk.Revit.DB.Plane plane = Autodesk.Revit.DB.Plane.CreateByNormalAndOrigin(norm, ptb);


            SketchPlane skplane = SketchPlane.Create(doc, plane);

            Autodesk.Revit.DB.Line line = Autodesk.Revit.DB.Line.CreateBound(pta, ptb);

            if (doc.IsFamilyDocument)
            {
                modelLine = doc.FamilyCreate.NewModelCurve(line, skplane) as ModelLine;
            }
            else
            {
                modelLine = doc.Create.NewModelCurve(line, skplane) as ModelLine;
            }
            if (modelLine == null)
            {
                TaskDialog.Show("Error", "Model line = null");
            }
            return modelLine;
        }

        static AddInId appId = new AddInId(new Guid("8F56AA78-A136-6509-AAF8-A478F3B24BAB"));
        public Autodesk.Revit.UI.Result Execute(ExternalCommandData commandData, ref string message, ElementSet elementSet)
        {

            UIDocument uidoc = commandData.Application.ActiveUIDocument;
            Autodesk.Revit.DB.Document doc = uidoc.Document;
            //Reference myRef = uidoc.Selection.PickObject(ObjectType.Element);
            //Element e = doc.GetElement(myRef);
            //LocationCurve locationCurve1 = e.Location as LocationCurve;
            //Autodesk.Revit.DB.Curve gridCurve = locationCurve1.Curve;
            //XYZ end1 = gridCurve.GetEndPoint(1);


            List<Autodesk.Revit.DB.Curve> crvs = new List<Autodesk.Revit.DB.Curve>();

            ICollection<Reference> my_lines = uidoc.Selection.PickObjects(Autodesk.Revit.UI.Selection.ObjectType.Element, "Select lines");
            foreach (var item_myRefWall in my_lines)
            {
                Element ele = doc.GetElement(item_myRefWall);
                GeometryObject geoobj = ele.GetGeometryObjectFromReference(item_myRefWall);
                Face face = geoobj as Face;
                LocationCurve locationCurve2 = ele.Location as LocationCurve;

                crvs.Add(locationCurve2.Curve);
            }


            var mepSystemTypes = new FilteredElementCollector(doc).OfClass(typeof(MEPSystemType))
            .OfType<MEPSystemType>().ToList();
            
            var ductTypes =
              new FilteredElementCollector(doc).OfClass(typeof(DuctType)).OfType<DuctType>().First();

           

            FilteredElementCollector collector2 = new FilteredElementCollector(doc);
            collector2.OfClass(typeof(Level));
            Level lv = collector2.First() as Level;

            if (lv == null)
            {
                message = "Wrong Active View";
                return Result.Failed;
            }

            var domesticHotWaterSystemType =
              mepSystemTypes.FirstOrDefault(
                st => st.SystemClassification ==
                  MEPSystemClassification.SupplyAir);

            if (domesticHotWaterSystemType == null)
            {
                message = "Could not found Domestic Hot Water System Type";
                return Result.Failed;
            }

            using (Transaction t = new Transaction(doc, "Create  pipe by line"))
            {
                t.Start();

                foreach (var crv in crvs)
                {
                    Duct.Create(doc, domesticHotWaterSystemType.Id, ductTypes.Id,
                      lv.Id, crv.GetEndPoint(0), crv.GetEndPoint(1));
                }
                

                t.Commit();
            }

            return Autodesk.Revit.UI.Result.Succeeded;
        }
    }

    [Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)]
    public class line_point_plane : IExternalCommand
    {
        public List<TessellatedShapeBuilderResult> Build_Tessellate2(Face faces, Autodesk.Revit.DB.Document doc)
        {
            Autodesk.Revit.DB.Mesh mesh = faces.Triangulate();
            List<XYZ> vert = new List<XYZ>();

            foreach (XYZ ij in mesh.Vertices)
            {
                XYZ vertices = ij;
                vert.Add(vertices);
            }

            TessellatedShapeBuilder builder = new TessellatedShapeBuilder();

            builder.OpenConnectedFaceSet(false);

            //Filter for Title Blocks in active document
            FilteredElementCollector materials = new FilteredElementCollector(doc)
            .OfClass(typeof(Autodesk.Revit.DB.Material))
            .OfCategory(BuiltInCategory.OST_Materials);

            ElementId materialId = materials.First().Id;

            builder.AddFace(new TessellatedFace(vert, materialId));

            builder.CloseConnectedFaceSet();
            builder.Target = TessellatedShapeBuilderTarget.AnyGeometry;
            builder.Fallback = TessellatedShapeBuilderFallback.Mesh;

            builder.Build();

            TessellatedShapeBuilderResult result3 = builder.GetBuildResult();
            List<TessellatedShapeBuilderResult> res = new List<TessellatedShapeBuilderResult>();

            if (result3.Outcome.ToString() == "Sheet")
            {
                res.Add(result3);
            }

            return res;
        }
        static AddInId appId = new AddInId(new Guid("5F92CC78-A137-4809-AAF8-A478F3B24BAB"));
        public Autodesk.Revit.UI.Result Execute(ExternalCommandData commandData, ref string message, ElementSet elementSet)
        {

            UIDocument uidoc = commandData.Application.ActiveUIDocument;
            Autodesk.Revit.DB.Document doc = uidoc.Document;
            List<CurveLoop> cls = new List<CurveLoop>();
            

            Reference myRef = uidoc.Selection.PickObject(ObjectType.Element);
            Element e = doc.GetElement(myRef);
           
            Reference myRef2 = uidoc.Selection.PickObject(ObjectType.PointOnElement);
            Element e2 = doc.GetElement(myRef2.ElementId);
            GeometryObject geomObj2 = e2.GetGeometryObjectFromReference(myRef2);
            XYZ p2 = myRef2.GlobalPoint;

            LocationCurve locationCurve1 = e.Location as LocationCurve;
            Autodesk.Revit.DB.Curve gridCurve = locationCurve1.Curve;

            

            using (Transaction tx = new Transaction(doc))
            {
                tx.Start("Make plane by line and point");


                
                Autodesk.Revit.DB.Plane pl = Autodesk.Revit.DB.Plane.CreateByThreePoints(gridCurve.GetEndPoint(0), p2
                , gridCurve.GetEndPoint(1));

                SketchPlane sketch = SketchPlane.Create(doc, pl);
                doc.ActiveView.SketchPlane = sketch;
                doc.ActiveView.ShowActiveWorkPlane();

                
                tx.Commit();
            }

            return Autodesk.Revit.UI.Result.Succeeded;
            
        }
    }

    [Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)]
    public class Closest_point_2Lines : IExternalCommand
    {
        public static ModelLine Makeline(Autodesk.Revit.DB.Document doc, XYZ pta, XYZ ptb)
        {
            ModelLine modelLine = null;
            double distance = pta.DistanceTo(ptb);
            if (distance < 0.01)
            {
                TaskDialog.Show("Error", "Distance" + distance);
                return modelLine;
            }

            XYZ norm = pta.CrossProduct(ptb);
            if (norm.GetLength() == 0)
            {
                XYZ aSubB = pta.Subtract(ptb);
                XYZ aSubBcrossz = aSubB.CrossProduct(XYZ.BasisZ);
                double crosslenght = aSubBcrossz.GetLength();
                if (crosslenght == 0)
                {
                    norm = XYZ.BasisY;
                }
                else
                {
                    norm = XYZ.BasisZ;
                }
            }

            //Autodesk.Revit.DB.Plane plane = Autodesk.Revit.DB.Plane.CreateByNormalAndOrigin(norm, ptb);

            Autodesk.Revit.DB.Line line = Autodesk.Revit.DB.Line.CreateUnbound(ptb, XYZ.BasisZ /* XYZ.BasisZ*/);

            Autodesk.Revit.DB.Plane pl = Autodesk.Revit.DB.Plane.CreateByThreePoints(pta,
             line.Evaluate(5, false), ptb);

            SketchPlane skplane = SketchPlane.Create(doc, pl);

            Autodesk.Revit.DB.Line line2 = Autodesk.Revit.DB.Line.CreateBound(pta, ptb);

            if (doc.IsFamilyDocument)
            {
                modelLine = doc.FamilyCreate.NewModelCurve(line2, skplane) as ModelLine;
            }
            else
            {
                modelLine = doc.Create.NewModelCurve(line2, skplane) as ModelLine;
            }
            if (modelLine == null)
            {
                TaskDialog.Show("Error", "Model line = null");
            }
            return modelLine;
        }

        static AddInId appId = new AddInId(new Guid("5F92AA78-A137-4809-AAF8-A478F3B24BAB"));
        public Autodesk.Revit.UI.Result Execute(ExternalCommandData commandData, ref string message, ElementSet elementSet)
        {

            UIDocument uidoc = commandData.Application.ActiveUIDocument;
            Autodesk.Revit.DB.Document doc = uidoc.Document;
            List<CurveLoop> cls = new List<CurveLoop>();


            Reference myRef = uidoc.Selection.PickObject(ObjectType.Element);
            Element e = doc.GetElement(myRef);

            Reference myRef2 = uidoc.Selection.PickObject(ObjectType.Element);
            Element e2 = doc.GetElement(myRef2);

            //GeometryObject geomObj2 = e2.GetGeometryObjectFromReference(myRef2);
            //XYZ p2 = myRef2.GlobalPoint;

            LocationCurve locationCurve1 = e.Location as LocationCurve;
            Autodesk.Revit.DB.Curve line1 = locationCurve1.Curve;

            LocationCurve locationCurve2 = e2.Location as LocationCurve;
            Autodesk.Revit.DB.Curve line2 = locationCurve2.Curve;


            IntersectionResult intres = line2.Project(line1.GetEndPoint(0));
            IntersectionResult intres2 = line2.Project(line1.GetEndPoint(1));

            IList<ClosestPointsPairBetweenTwoCurves> closestPoints = new List<ClosestPointsPairBetweenTwoCurves>();

            line1.ComputeClosestPoints(line2, true, false, false, out closestPoints);
            XYZ closestPoint1 = closestPoints.FirstOrDefault().XYZPointOnFirstCurve;

            line2.ComputeClosestPoints(line1, true, false, false, out closestPoints);
            XYZ closestPoint2 = closestPoints.FirstOrDefault().XYZPointOnFirstCurve;


            using (Transaction tx = new Transaction(doc))
            {
                tx.Start("Make plane by line and point");


                ModelLine ml = Makeline(doc, closestPoint1, closestPoint2);

                //ModelLine ml1 = Makeline(doc, line1.GetEndPoint(0), intres.XYZPoint);
                //ModelLine ml2 = Makeline(doc, line1.GetEndPoint(1), intres2.XYZPoint);




                tx.Commit();
            }

            return Autodesk.Revit.UI.Result.Succeeded;

        }
    }

    [Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)]
    public class loft : IExternalCommand
    {
        public List<TessellatedShapeBuilderResult> Build_Tessellate2(Face faces, Autodesk.Revit.DB.Document doc)
        {
            Autodesk.Revit.DB.Mesh mesh = faces.Triangulate();
            List<XYZ> vert = new List<XYZ>();

            foreach (XYZ ij in mesh.Vertices)
            {
                XYZ vertices = ij;
                vert.Add(vertices);
            }

            TessellatedShapeBuilder builder = new TessellatedShapeBuilder();

            builder.OpenConnectedFaceSet(false);

            //Filter for Title Blocks in active document
            FilteredElementCollector materials = new FilteredElementCollector(doc)
            .OfClass(typeof(Autodesk.Revit.DB.Material))
            .OfCategory(BuiltInCategory.OST_Materials);

            ElementId materialId = materials.First().Id;

            builder.AddFace(new TessellatedFace(vert, materialId));

            builder.CloseConnectedFaceSet();
            builder.Target = TessellatedShapeBuilderTarget.AnyGeometry;
            builder.Fallback = TessellatedShapeBuilderFallback.Mesh;

            builder.Build();

            TessellatedShapeBuilderResult result3 = builder.GetBuildResult();
            List<TessellatedShapeBuilderResult> res = new List<TessellatedShapeBuilderResult>();

            if (result3.Outcome.ToString() == "Sheet")
            {
                res.Add(result3);
            }

            return res;
        }
        static AddInId appId = new AddInId(new Guid("6F92CC78-A137-4809-AAF8-A478F3B24BAB"));
        public Autodesk.Revit.UI.Result Execute(ExternalCommandData commandData, ref string message, ElementSet elementSet)
        {

            UIDocument uidoc = commandData.Application.ActiveUIDocument;
            Autodesk.Revit.DB.Document doc = uidoc.Document;
            List<CurveLoop> cls = new List<CurveLoop>();
            

            Reference myRef = uidoc.Selection.PickObject(ObjectType.Element);
            Element e = doc.GetElement(myRef);


            Reference myRef3 = uidoc.Selection.PickObject(ObjectType.Element);
            Element e3 = doc.GetElement(myRef3);

            

            LocationCurve locationCurve1 = e.Location as LocationCurve;

            Autodesk.Revit.DB.Curve gridCurve = locationCurve1.Curve;

            LocationCurve locationCurve3 = e3.Location as LocationCurve;



            CurveLoop bottomLoop = new CurveLoop();
            CurveLoop topLoop = new CurveLoop();
            bottomLoop.Append(locationCurve1.Curve);
            topLoop.Append(locationCurve3.Curve);
            cls.Add(bottomLoop);
            cls.Add(topLoop);


            using (Transaction tx = new Transaction(doc))
            {
                tx.Start("Make loft");
                
                
                SolidOptions options = new SolidOptions(ElementId.InvalidElementId, ElementId.InvalidElementId);
                Solid mySolid = GeometryCreationUtilities.CreateLoftGeometry(cls, options);

            
                List<Solid> sol = new List<Solid>();

                sol.Add(mySolid);

                foreach (Face f in mySolid.Faces)
                {
                    //Build_Tessellate2(f, doc);
                    ElementId categoryId = new ElementId(BuiltInCategory.OST_GenericModel);
                    DirectShape ds = DirectShape.CreateElement(doc, categoryId);


                    ds.ApplicationId = System.Reflection.Assembly.GetExecutingAssembly().GetType().GUID.ToString();
                    ds.ApplicationDataId = Guid.NewGuid().ToString();


                    IList<GeometryObject> list = new List<GeometryObject>();
                    list.Add(mySolid);


                    // Create a direct shape.

                    DirectShape ds2 = DirectShape.CreateElement(doc,
                      new ElementId(BuiltInCategory.OST_GenericModel));

                    ds2.SetShape(list);
                    



                    //foreach (TessellatedShapeBuilderResult t1 in Build_Tessellate2(f, doc))
                    //{
                    //    ds.SetShape(t1.GetGeometricalObjects());

                    //    ds.Name = "Single_Surface";
                    //}
                }
                foreach (Solid s in sol)
                {

                }

                tx.Commit();
            }

            return Autodesk.Revit.UI.Result.Succeeded;

        }
    }

    [Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)]
    public class rhino_points_to_revit_topo : IExternalCommand
    {
       
        static AddInId appId = new AddInId(new Guid("6F92CC78-A137-4809-AAF8-A478F3B24BAB"));
        public Autodesk.Revit.UI.Result Execute(ExternalCommandData commandData, ref string message, ElementSet elementSet)
        {
            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Autodesk.Revit.DB.Document doc = uidoc.Document;

            try
            {
                string filename = @"T:\Transfer\lopez\Book1.xlsx";
                using (ExcelPackage package = new ExcelPackage(new FileInfo(filename)))
                {
                    ExcelWorksheet sheet = package.Workbook.Worksheets.ElementAt(0);

                    int column = 15;
                    int number = Convert.ToInt32(sheet.Cells[2, column].Value);
                    sheet.Cells[2, column].Value = (number + 1); ;
                    package.Save();
                }
            }



            catch (Exception)
            {
                MessageBox.Show("Excel file not found", "");
            }

            List<Object> objs = new List<Object>();
            List<List<XYZ>> xyz_faces = new List<List<XYZ>>();
            IList<Face> face_with_regions = new List<Face>();
            String info = "";
            List<List<Face>> Faces_lists_excel = new List<List<Face>>();
            IList<CurveLoop> faceboundaries = new List<CurveLoop>();
            List<List<Element>> elemente_selected = new List<List<Element>>();
            List<List<string>> names = new List<List<string>>();
            List<int> numeros_ = new List<int>();
            XYZ pos_z = new XYZ(0, 0, 1);
            XYZ neg_z = new XYZ(0, 0, -1);



            string filename2 = "";
            System.Windows.Forms.OpenFileDialog openDialog = new System.Windows.Forms.OpenFileDialog();
            openDialog.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
            if (openDialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                filename2 = openDialog.FileName;
            }

            Point3d p1 = new Point3d(0, 0, 0);
            Rhino.Geometry.Point3d pt3d = new Point3d(10, 10, 0);

            File3dm m_modelfile = null;
            string m_name = null;
            string m_size = null;
            string m_created = null;
            string m_createdby = null;
            string m_edited = null;
            string m_editedby = null;
            string m_revision = null;
            string m_units = null;
            string m_notes = null;

            Object RhinoFile = filename2;
            if (RhinoFile is System.IO.FileInfo)
            {
                System.IO.FileInfo m_fileinfo = (System.IO.FileInfo)RhinoFile;
                m_modelfile = File3dm.Read(m_fileinfo.FullName);
                m_name = m_fileinfo.Name;
                m_size = m_fileinfo.Length.ToString();
                m_created = m_fileinfo.CreationTimeUtc.ToString();
                m_createdby = m_modelfile.CreatedBy;
                m_edited = m_fileinfo.LastWriteTimeUtc.ToString();
                m_editedby = m_modelfile.LastEditedBy;
                m_revision = m_modelfile.Revision.ToString();
                m_units = m_modelfile.Settings.ModelUnitSystem.ToString();
                m_notes = m_modelfile.Notes.Notes;
            }
            else if (RhinoFile is string)
            {
                System.IO.FileInfo m_fileinfo = new System.IO.FileInfo((string)RhinoFile);
                m_modelfile = File3dm.Read(m_fileinfo.FullName);
                m_name = m_fileinfo.Name;
                m_size = m_fileinfo.Length.ToString();
                m_created = m_fileinfo.CreationTimeUtc.ToString();
                m_createdby = m_modelfile.CreatedBy;
                m_edited = m_fileinfo.LastWriteTimeUtc.ToString();
                m_editedby = m_modelfile.LastEditedBy;
                m_revision = m_modelfile.Revision.ToString();
                m_units = m_modelfile.Settings.ModelUnitSystem.ToString();
                m_notes = m_modelfile.Notes.Notes;
            }

            File3dmLayerTable all_layers = m_modelfile.AllLayers;

            var objs_ = m_modelfile.Objects;
            List<Rhino.DocObjects.Layer> listadelayers = new List<Rhino.DocObjects.Layer>();
            List<Rhino.DocObjects.Layer> borrar = new List<Rhino.DocObjects.Layer>();
            List<List<Rhino.DocObjects.Layer>> multiplelistadelayers = new List<List<Rhino.DocObjects.Layer>>();
            string hola = "";

            MessageBox.Show("This tool will read 3D information only if the following Rhino Layers exist; Levels, Grids, Structure, Floor, Walls, Points", "!");

            foreach (var item in all_layers)
            {

                if (item.FullPath.Contains(hola))
                {
                    listadelayers.Add(item);
                }
            }
            List<Rhino.Geometry.Brep> rh_breps = new List<Rhino.Geometry.Brep>();
            List<Rhino.Geometry.Curve> curves_frames = new List<Rhino.Geometry.Curve>();
            List<Autodesk.Revit.DB.Curve> revit_crv = new List<Autodesk.Revit.DB.Curve>();
            List<string> m_names = new List<string>();
            List<int> m_layerindeces = new List<int>();
            List<System.Drawing.Color> m_colors = new List<System.Drawing.Color>();
            List<string> m_guids = new List<string>();

            foreach (var Layer in listadelayers)
            {
                if (Layer.Name == "Points")
                {
                    File3dmObject[] m_objs = m_modelfile.Objects.FindByLayer(Layer.Name);


                    List<Point3d> Points = new List<Point3d>();
                    List<XYZ> Revit_Points = new List<XYZ>();
                    List<XYZ> Revit_Points_notsameloc = new List<XYZ>();

                    foreach (var obj in m_objs)
                    {
                        GeometryBase brep = obj.Geometry;
                        Rhino.Geometry.Point pt = brep as Rhino.Geometry.Point;
                        Points.Add(pt.Location);
                    }
                    foreach (var item in Points)
                    {
                        double x = 0;
                        double y = 0;
                        double z = 0;

                        x = item.X / 304.8;
                        y = item.Y / 304.8;
                        z = item.Z / 304.8;

                        XYZ newpt = new XYZ(x, y, z);

                        Revit_Points.Add(newpt);
                    }

                    foreach (var item in Revit_Points)
                    {
                        Revit_Points_notsameloc.Add(item);
                    }

                    for (int i = 0; i < Revit_Points.ToArray().Length; i++)
                    {
                        for (int ij = 0; ij < Revit_Points.ToArray().Length; ij++)
                        {
                            if (Revit_Points.ToArray()[i].X == Revit_Points.ToArray()[ij].X &&
                                Revit_Points.ToArray()[i].Y == Revit_Points.ToArray()[ij].Y &&
                                Revit_Points.ToArray()[i].Z != Revit_Points.ToArray()[ij].Z)
                            {
                                Revit_Points_notsameloc.Remove(Revit_Points.ToArray()[i]);
                            }
                            if (i != ij && Revit_Points.ToArray()[i].X == Revit_Points.ToArray()[ij].X &&
                                Revit_Points.ToArray()[i].Y == Revit_Points.ToArray()[ij].Y &&
                                Revit_Points.ToArray()[i].Z == Revit_Points.ToArray()[ij].Z)
                            {
                                Revit_Points_notsameloc.Remove(Revit_Points.ToArray()[i]);
                            }
                        }
                    }
                    using (Transaction t = new Transaction(doc, "Make topograpgy from Rhino"))
                    {
                        t.Start();

                        TopographySurface ts = TopographySurface.Create(doc, Revit_Points_notsameloc);

                        t.Commit();
                    }
                }
            }
            return Autodesk.Revit.UI.Result.Succeeded;
        }
    }
    
    [Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)]
    public class rhino_lns_to_revit_lns : IExternalCommand
    {
      

        public static ModelLine Makeline(Autodesk.Revit.DB.Document doc, XYZ pta, XYZ ptb)
        {
            ModelLine modelLine = null;
            double distance = pta.DistanceTo(ptb);
            if (distance < 0.01)
            {
                TaskDialog.Show("Error", "Distance" + distance);
                return modelLine;
            }

            XYZ norm = pta.CrossProduct(ptb);
            if (norm.GetLength() == 0)
            {
                XYZ aSubB = pta.Subtract(ptb);
                XYZ aSubBcrossz = aSubB.CrossProduct(XYZ.BasisZ);
                double crosslenght = aSubBcrossz.GetLength();
                if (crosslenght == 0)
                {
                    norm = XYZ.BasisY;
                }
                else
                {
                    norm = XYZ.BasisZ;
                }
            }

            Autodesk.Revit.DB.Plane plane = Autodesk.Revit.DB.Plane.CreateByNormalAndOrigin(norm, ptb);


            SketchPlane skplane = SketchPlane.Create(doc, plane);

            Autodesk.Revit.DB.Line line = Autodesk.Revit.DB.Line.CreateBound(pta, ptb);

            if (doc.IsFamilyDocument)
            {
                modelLine = doc.FamilyCreate.NewModelCurve(line, skplane) as ModelLine;
            }
            else
            {
                modelLine = doc.Create.NewModelCurve(line, skplane) as ModelLine;
            }
            if (modelLine == null)
            {
                TaskDialog.Show("Error", "Model line = null");
            }
            return modelLine;
        }

        public static bool IsZero(double a)
        {
            const double _eps = 1.0e-9;
            return _eps > Math.Abs(a);
        }

        public static bool IsEqual(double a, double b)
        {
            return IsZero(b - a);
        }

        public static int Compare(double a, double b)
        {
            return IsEqual(a, b) ? 0 : (a < b ? -1 : 1);
        }

        public static int Compare(XYZ p, XYZ q)
        {

            int diff = Compare(p.X, q.X);
            if (0 == diff)
            {
                diff = Compare(p.Y, q.Y);
                if (0 == diff)
                {
                    diff = Compare(p.Z, q.Z);
                }

            }

            return diff;
        }

        private static Wall CreateWall(FamilyInstance cube, Autodesk.Revit.DB.Curve curve, double height)
        {
            var doc = cube.Document;

            var wallTypeId = doc.GetDefaultElementTypeId(
              ElementTypeGroup.WallType);

            return Wall.Create(doc, curve.CreateReversed(),
              wallTypeId, cube.LevelId, height, 0, false,
              false);
        }
        static AddInId appId = new AddInId(new Guid("5F88CC78-A137-4809-AAF8-A478F3B24BAB"));
        public Autodesk.Revit.UI.Result Execute(ExternalCommandData commandData, ref string message, ElementSet elementSet)
        {
            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Autodesk.Revit.DB.Document doc = uidoc.Document;

            try
            {
                string filename = @"T:\Transfer\lopez\Book1.xlsx";
                using (ExcelPackage package = new ExcelPackage(new FileInfo(filename)))
                {
                    ExcelWorksheet sheet = package.Workbook.Worksheets.ElementAt(0);

                    int column = 15;
                    int number = Convert.ToInt32(sheet.Cells[2, column].Value);
                    sheet.Cells[2, column].Value = (number + 1); ;
                    package.Save();
                }
            }



            catch (Exception)
            {
                MessageBox.Show("Excel file not found", "");
            }

            List<Object> objs = new List<Object>();
            List<List<XYZ>> xyz_faces = new List<List<XYZ>>();
            IList<Face> face_with_regions = new List<Face>();
            String info = "";
            List<List<Face>> Faces_lists_excel = new List<List<Face>>();
            //List<FaceArray> face112 = new List<FaceArray>();
            IList<CurveLoop> faceboundaries = new List<CurveLoop>();
            List<List<Element>> elemente_selected = new List<List<Element>>();
            List<List<string>> names = new List<List<string>>();
            List<int> numeros_ = new List<int>();
            XYZ pos_z = new XYZ(0, 0, 1);
            XYZ neg_z = new XYZ(0, 0, -1);
            
            string filename2 = "";
            System.Windows.Forms.OpenFileDialog openDialog = new System.Windows.Forms.OpenFileDialog();
            openDialog.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
            if (openDialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                filename2 = openDialog.FileName;
            }

            Point3d p1 = new Point3d(0, 0, 0);
            Rhino.Geometry.Point3d pt3d = new Point3d(10, 10, 0);

            File3dm m_modelfile = null;
            string m_name = null;
            string m_size = null;
            string m_created = null;
            string m_createdby = null;
            string m_edited = null;
            string m_editedby = null;
            string m_revision = null;
            string m_units = null;
            string m_notes = null;

            Object RhinoFile = filename2;
            if (RhinoFile is System.IO.FileInfo)
            {
                System.IO.FileInfo m_fileinfo = (System.IO.FileInfo)RhinoFile;
                m_modelfile = File3dm.Read(m_fileinfo.FullName);
                m_name = m_fileinfo.Name;
                m_size = m_fileinfo.Length.ToString();
                m_created = m_fileinfo.CreationTimeUtc.ToString();
                m_createdby = m_modelfile.CreatedBy;
                m_edited = m_fileinfo.LastWriteTimeUtc.ToString();
                m_editedby = m_modelfile.LastEditedBy;
                m_revision = m_modelfile.Revision.ToString();
                m_units = m_modelfile.Settings.ModelUnitSystem.ToString();
                m_notes = m_modelfile.Notes.Notes;
            }
            else if (RhinoFile is string)
            {
                System.IO.FileInfo m_fileinfo = new System.IO.FileInfo((string)RhinoFile);
                m_modelfile = File3dm.Read(m_fileinfo.FullName);
                m_name = m_fileinfo.Name;
                m_size = m_fileinfo.Length.ToString();
                m_created = m_fileinfo.CreationTimeUtc.ToString();
                m_createdby = m_modelfile.CreatedBy;
                m_edited = m_fileinfo.LastWriteTimeUtc.ToString();
                m_editedby = m_modelfile.LastEditedBy;
                m_revision = m_modelfile.Revision.ToString();
                m_units = m_modelfile.Settings.ModelUnitSystem.ToString();
                m_notes = m_modelfile.Notes.Notes;
            }

            File3dmLayerTable all_layers = m_modelfile.AllLayers;
            var objs_ = m_modelfile.Objects;
            List<Rhino.DocObjects.Layer> listadelayers = new List<Rhino.DocObjects.Layer>();
            List<Rhino.DocObjects.Layer> borrar = new List<Rhino.DocObjects.Layer>();
            List<List<Rhino.DocObjects.Layer>> multiplelistadelayers = new List<List<Rhino.DocObjects.Layer>>();
            string hola = "";

            MessageBox.Show("This tool will read 3D information only if the following Rhino Layers exist; Levels, Grids, Structure, Floor, Walls, Points", "!");

            foreach (var item in all_layers)
            {

                if (item.FullPath.Contains(hola))
                {
                    listadelayers.Add(item);
                }
            }

            List<Rhino.Geometry.Brep> rh_breps = new List<Rhino.Geometry.Brep>();
            List<XYZ> pts_ = new List<XYZ>();
            
            List<Autodesk.Revit.DB.Curve> revit_crv = new List<Autodesk.Revit.DB.Curve>();
            List<string> m_names = new List<string>();
            List<int> m_layerindeces = new List<int>();
            List<System.Drawing.Color> m_colors = new List<System.Drawing.Color>();
            List<string> m_guids = new List<string>();
            List<double> weith = new List<double>();

            

            foreach (var Layer in listadelayers)
            {
                

                if (Layer.Name == "Lines")
                {


                    File3dmObject[] m_objs = m_modelfile.Objects.FindByLayer(Layer.Name);

                    using (Transaction t = new Transaction(doc, "Lines from Rhino"))
                    {
                        t.Start();

                        foreach (File3dmObject obj in m_objs)
                        {
                            GeometryBase geo = obj.Geometry;
                            Type TYPE = geo.GetType();

                            if (TYPE.Name == "PolyCurve")
                            {
                                Rhino.Geometry.PolyCurve POLY = geo as Rhino.Geometry.PolyCurve;
                            }


                            
                            if (TYPE.Name == "ArcCurve")
                            {
                                Rhino.Geometry.ArcCurve arc = geo as Rhino.Geometry.ArcCurve;

                                double x_end = arc.Arc.Plane.Normal.X / 304.8;
                                double y_end = arc.Arc.Plane.Normal.Y / 304.8;
                                double z_end = arc.Arc.Plane.Normal.Z / 304.8;

                                XYZ normal = new XYZ(x_end, y_end, z_end);

                                double x_ = arc.Arc.Plane.Origin.X / 304.8;
                                double y_= arc.Arc.Plane.Origin.Y / 304.8;
                                double z_ = arc.Arc.Plane.Origin.Z / 304.8;
                                XYZ origin = new XYZ(x_, y_, z_);

                                double endAngle = 2 * Math.PI;        // this arc will be a circle

                                Autodesk.Revit.DB.Plane pl = Autodesk.Revit.DB.Plane.CreateByNormalAndOrigin(normal, origin );

                                SketchPlane skplane = SketchPlane.Create(doc, pl);
                                doc.Create.NewModelCurve(Autodesk.Revit.DB.Arc.Create(pl, arc.Radius / 304.8, arc.Arc.StartAngle, arc.Arc.EndAngle), skplane);

                            }

                            if (TYPE.Name == "PolylineCurve")
                            {
                                Rhino.Geometry.PolylineCurve PolylineCurve = geo as Rhino.Geometry.PolylineCurve;
                                Rhino.Geometry.Polyline Polyline = PolylineCurve.ToPolyline();

                                for (int i = 0; i < Polyline.SegmentCount; i++)
                                {
                                    Point3d end = Polyline.SegmentAt(i).PointAt(0);
                                    Point3d start = Polyline.SegmentAt(i).PointAt(1); 
                                    

                                    double x_end = end.X / 304.8;
                                    double y_end = end.Y / 304.8;
                                    double z_end = end.Z / 304.8;

                                    double x_start = start.X / 304.8;
                                    double y_start = start.Y / 304.8;
                                    double z_start = start.Z / 304.8;

                                    XYZ pt_end = new XYZ(x_end, y_end, z_end);
                                    XYZ pt_start = new XYZ(x_start, y_start, z_start);

                                    Autodesk.Revit.DB.Curve curve1 = Autodesk.Revit.DB.Line.CreateBound(pt_end, pt_start);
                                    Autodesk.Revit.DB.Line line = Autodesk.Revit.DB.Line.CreateBound(pt_end, pt_start);

                                    Makeline(doc, line.GetEndPoint(0), line.GetEndPoint(1));
                                    revit_crv.Add(curve1);
                                }


                            }
                            
                            if (TYPE.Name == "NurbsCurve")
                            {
                                Rhino.Geometry.NurbsCurve crv_ = geo as Rhino.Geometry.NurbsCurve;
                                Rhino.Geometry.Plane pl2 ;
                                crv_.TryGetPlane(out pl2);
                                
                                double x_end = pl2.Normal.X / 304.8;
                                double y_end = pl2.Normal.Y / 304.8;
                                double z_end = pl2.Normal.Z / 304.8;

                                XYZ normal = new XYZ(x_end, y_end, z_end);

                                double x_2 = pl2.Origin.X / 304.8;
                                double y_2 = pl2.Origin.Y / 304.8;
                                double z_2 = pl2.Origin.Z / 304.8;
                                XYZ origin = new XYZ(x_2, y_2, z_2);

                                     // this arc will be a circle

                                Autodesk.Revit.DB.Plane pl3 = Autodesk.Revit.DB.Plane.CreateByNormalAndOrigin(normal, origin);

                                var pts = crv_.Points;
                                foreach (var item in pts)
                                {
                                    double x_ = item.X / 304.8;
                                    double y_ = item.Y / 304.8;
                                    double z_ = item.Z / 304.8;

                                    XYZ pt = new XYZ(x_, y_, z_);
                                    weith.Add(item.Weight);
                                    pts_.Add(pt);
                                }

                                Autodesk.Revit.DB.Curve curve1 = Autodesk.Revit.DB.Line.CreateBound(pts_.First(),pts_.Last()) ;
                                
                                XYZ norm = pts_.First().CrossProduct(curve1.Evaluate(5, false));
                                if (norm.GetLength() == 0)
                                {
                                    XYZ aSubB = pts_.First().Subtract(pts_.Last());
                                    XYZ aSubBcrossz = aSubB.CrossProduct(XYZ.BasisZ);
                                    double crosslenght = aSubBcrossz.GetLength();
                                    if (crosslenght == 0)
                                    {
                                        norm = XYZ.BasisY;
                                    }
                                    else
                                    {
                                        norm = XYZ.BasisZ;
                                    }
                                }
                                //Autodesk.Revit.DB.Plane pl = Autodesk.Revit.DB.Plane.CreateByNormalAndOrigin(norm, pts_.First());
                                Autodesk.Revit.DB.Plane pl = Autodesk.Revit.DB.Plane.CreateByThreePoints(pts_.First(), pts_.ElementAt(2), pts_.Last());
                                SketchPlane skplane = SketchPlane.Create(doc, pl);

                                Autodesk.Revit.DB.Curve nspline = Autodesk.Revit.DB.NurbSpline.CreateCurve(pts_, weith);
                                doc.Create.NewModelCurve(nspline, skplane);

                                pts_.Clear();
                                weith.Clear();
                            }

                            if (TYPE.Name == "Curve")
                            {
                                Rhino.Geometry.Curve crv_ = geo as Rhino.Geometry.Curve;


                                Point3d end = crv_.PointAtEnd;
                                Point3d start = crv_.PointAtStart;

                                double x_end = end.X / 304.8;
                                double y_end = end.Y / 304.8;
                                double z_end = end.Z / 304.8;

                                double x_start = start.X / 304.8;
                                double y_start = start.Y / 304.8;
                                double z_start = start.Z / 304.8;

                                XYZ pt_end = new XYZ(x_end, y_end, z_end);
                                XYZ pt_start = new XYZ(x_start, y_start, z_start);

                                Autodesk.Revit.DB.Curve curve1 = Autodesk.Revit.DB.Line.CreateBound(pt_end, pt_start);
                                Autodesk.Revit.DB.Line line = Autodesk.Revit.DB.Line.CreateBound(pt_end, pt_start);

                                Makeline(doc, line.GetEndPoint(0), line.GetEndPoint(1));
                                revit_crv.Add(curve1);
                            }
                        }
                        t.Commit();
                    }
                }
            }
            return Autodesk.Revit.UI.Result.Succeeded;
        }
    }

  

    [Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)]
    public class Create_flex_ducts_from_line : IExternalCommand
    {


        static AddInId appId = new AddInId(new Guid("5F88CC78-A137-4809-AAF8-A478F3B24BAB"));
        public Autodesk.Revit.UI.Result Execute(ExternalCommandData commandData, ref string message, ElementSet elementSet)
        {
            

            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Autodesk.Revit.ApplicationServices.Application app = uiapp.Application;
            Autodesk.Revit.DB.Document doc = uidoc.Document;


            List<Autodesk.Revit.DB.XYZ> crvs = new List<Autodesk.Revit.DB.XYZ>();
            Autodesk.Revit.DB.NurbSpline line = null;

            ICollection<Reference> my_lines = uidoc.Selection.PickObjects(Autodesk.Revit.UI.Selection.ObjectType.Element, "Select lines");
            foreach (var item_myRefWall in my_lines)
            {
                Element ele = doc.GetElement(item_myRefWall);
                GeometryObject geoobj = ele.GetGeometryObjectFromReference(item_myRefWall);
                Face face = geoobj as Face;
                LocationCurve locationCurve2 = ele.Location as LocationCurve;
                


                if (locationCurve2.Curve.GetType().Name == "NurbSpline")
                {
                    line = locationCurve2.Curve as Autodesk.Revit.DB.NurbSpline;
                    
                }
            }

            var ductTypes =new FilteredElementCollector(doc).OfClass(typeof(FlexDuctType)).OfType<FlexDuctType>().First();

            using (Transaction tx = new Transaction(doc))
            {
                tx.Start("flex"
                  + "flex");
                FlexDuct flexDuct = doc.Create.NewFlexDuct(line.CtrlPoints, ductTypes);

                tx.Commit();
            }

            

            return Autodesk.Revit.UI.Result.Succeeded;
        }
    }

    [Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)]
    public class rhino_lns_to_revit_structure : IExternalCommand
    {

        public static ModelLine Makeline(Autodesk.Revit.DB.Document doc, XYZ pta, XYZ ptb)
        {
            ModelLine modelLine = null;
            double distance = pta.DistanceTo(ptb);
            if (distance < 0.01)
            {
                TaskDialog.Show("Error", "Distance" + distance);
                return modelLine;
            }

            XYZ norm = pta.CrossProduct(ptb);
            if (norm.GetLength() == 0)
            {
                XYZ aSubB = pta.Subtract(ptb);
                XYZ aSubBcrossz = aSubB.CrossProduct(XYZ.BasisZ);
                double crosslenght = aSubBcrossz.GetLength();
                if (crosslenght == 0)
                {
                    norm = XYZ.BasisY;
                }
                else
                {
                    norm = XYZ.BasisZ;
                }
            }

            Autodesk.Revit.DB.Plane plane = Autodesk.Revit.DB.Plane.CreateByNormalAndOrigin(norm, ptb);


            SketchPlane skplane = SketchPlane.Create(doc, plane);

            Autodesk.Revit.DB.Line line = Autodesk.Revit.DB.Line.CreateBound(pta, ptb);

            if (doc.IsFamilyDocument)
            {
                modelLine = doc.FamilyCreate.NewModelCurve(line, skplane) as ModelLine;
            }
            else
            {
                modelLine = doc.Create.NewModelCurve(line, skplane) as ModelLine;
            }
            if (modelLine == null)
            {
                TaskDialog.Show("Error", "Model line = null");
            }
            return modelLine;
        }

        public static bool IsZero(double a)
        {
            const double _eps = 1.0e-9;
            return _eps > Math.Abs(a);
        }

        public static bool IsEqual(double a, double b)
        {
            return IsZero(b - a);
        }

        public static int Compare(double a, double b)
        {
            return IsEqual(a, b) ? 0 : (a < b ? -1 : 1);
        }

        public static int Compare(XYZ p, XYZ q)
        {

            int diff = Compare(p.X, q.X);
            if (0 == diff)
            {
                diff = Compare(p.Y, q.Y);
                if (0 == diff)
                {
                    diff = Compare(p.Z, q.Z);
                }

            }

            return diff;
        }

        private static Wall CreateWall(FamilyInstance cube, Autodesk.Revit.DB.Curve curve, double height)
        {
            var doc = cube.Document;

            var wallTypeId = doc.GetDefaultElementTypeId(
              ElementTypeGroup.WallType);

            return Wall.Create(doc, curve.CreateReversed(),
              wallTypeId, cube.LevelId, height, 0, false,
              false);
        }
        static AddInId appId = new AddInId(new Guid("5F88CC78-A137-4809-AAF8-A478F3B24BAB"));
        public Autodesk.Revit.UI.Result Execute(ExternalCommandData commandData, ref string message, ElementSet elementSet)
        {
            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Autodesk.Revit.DB.Document doc = uidoc.Document;

            try
            {
                string filename = @"T:\Transfer\lopez\Book1.xlsx";
                using (ExcelPackage package = new ExcelPackage(new FileInfo(filename)))
                {
                    ExcelWorksheet sheet = package.Workbook.Worksheets.ElementAt(0);

                    int column = 15;
                    int number = Convert.ToInt32(sheet.Cells[2, column].Value);
                    sheet.Cells[2, column].Value = (number + 1); ;
                    package.Save();
                }
            }

            catch (Exception)
            {
                MessageBox.Show("Excel file not found", "");
            }

            List<Object> objs = new List<Object>();
            List<List<XYZ>> xyz_faces = new List<List<XYZ>>();
            IList<Face> face_with_regions = new List<Face>();
            String info = "";
            List<List<Face>> Faces_lists_excel = new List<List<Face>>();
            //List<FaceArray> face112 = new List<FaceArray>();
            IList<CurveLoop> faceboundaries = new List<CurveLoop>();
            List<List<Element>> elemente_selected = new List<List<Element>>();
            List<List<string>> names = new List<List<string>>();
            List<int> numeros_ = new List<int>();
            XYZ pos_z = new XYZ(0, 0, 1);
            XYZ neg_z = new XYZ(0, 0, -1);

            string filename2 = "";
            System.Windows.Forms.OpenFileDialog openDialog = new System.Windows.Forms.OpenFileDialog();
            openDialog.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
            if (openDialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                filename2 = openDialog.FileName;
            }

            Point3d p1 = new Point3d(0, 0, 0);
            Rhino.Geometry.Point3d pt3d = new Point3d(10, 10, 0);

            File3dm m_modelfile = null;
            string m_name = null;
            string m_size = null;
            string m_created = null;
            string m_createdby = null;
            string m_edited = null;
            string m_editedby = null;
            string m_revision = null;
            string m_units = null;
            string m_notes = null;

            Object RhinoFile = filename2;
            if (RhinoFile is System.IO.FileInfo)
            {
                System.IO.FileInfo m_fileinfo = (System.IO.FileInfo)RhinoFile;
                m_modelfile = File3dm.Read(m_fileinfo.FullName);
                m_name = m_fileinfo.Name;
                m_size = m_fileinfo.Length.ToString();
                m_created = m_fileinfo.CreationTimeUtc.ToString();
                m_createdby = m_modelfile.CreatedBy;
                m_edited = m_fileinfo.LastWriteTimeUtc.ToString();
                m_editedby = m_modelfile.LastEditedBy;
                m_revision = m_modelfile.Revision.ToString();
                m_units = m_modelfile.Settings.ModelUnitSystem.ToString();
                m_notes = m_modelfile.Notes.Notes;
            }
            else if (RhinoFile is string)
            {
                System.IO.FileInfo m_fileinfo = new System.IO.FileInfo((string)RhinoFile);
                m_modelfile = File3dm.Read(m_fileinfo.FullName);
                m_name = m_fileinfo.Name;
                m_size = m_fileinfo.Length.ToString();
                m_created = m_fileinfo.CreationTimeUtc.ToString();
                m_createdby = m_modelfile.CreatedBy;
                m_edited = m_fileinfo.LastWriteTimeUtc.ToString();
                m_editedby = m_modelfile.LastEditedBy;
                m_revision = m_modelfile.Revision.ToString();
                m_units = m_modelfile.Settings.ModelUnitSystem.ToString();
                m_notes = m_modelfile.Notes.Notes;
            }

            File3dmLayerTable all_layers = m_modelfile.AllLayers;
            var objs_ = m_modelfile.Objects;

            List<Rhino.DocObjects.Layer> listadelayers = new List<Rhino.DocObjects.Layer>();
            List<Rhino.DocObjects.Layer> borrar = new List<Rhino.DocObjects.Layer>();
            List<List<Rhino.DocObjects.Layer>> multiplelistadelayers = new List<List<Rhino.DocObjects.Layer>>();
            string hola = "";
            MessageBox.Show("This tool will read 3D information only if the following Rhino Layers exist; Levels, Grids, Structure, Floor, Walls, Points", "!");

            foreach (var item in all_layers)
            {
                if (item.FullPath.Contains(hola))
                {
                    listadelayers.Add(item);
                }
            }

            List<Rhino.Geometry.Brep> rh_breps = new List<Rhino.Geometry.Brep>();
            List<Rhino.Geometry.Curve> curves_frames = new List<Rhino.Geometry.Curve>();
            List<Autodesk.Revit.DB.Curve> revit_crv = new List<Autodesk.Revit.DB.Curve>();
            List<string> m_names = new List<string>();
            List<int> m_layerindeces = new List<int>();
            List<System.Drawing.Color> m_colors = new List<System.Drawing.Color>();
            List<string> m_guids = new List<string>();

            foreach (var Layer in listadelayers)
            {
                FamilySymbol beamSymbol = null;

                if (Layer.Name == "Structure")
                {
                    Level level = new FilteredElementCollector(doc).OfClass(typeof(Level)).Cast<Level>().OrderBy(q => q.Elevation).First();
                    List<FamilyInstance> famsimbol = new List<FamilyInstance>();

                    foreach (var item in new FilteredElementCollector(doc).OfClass(typeof(FamilySymbol)).Cast<FamilySymbol>())
                    {
                        if (item.IsActive == true && item.Family.FamilyCategory.Name == "Structural Framing")
                        {
                            beamSymbol = item;
                        }
                    }

                    using (Transaction t = new Transaction(doc, "Make structural frame from Rhino"))
                    {
                        t.Start();
                        File3dmObject[] m_objs = m_modelfile.Objects.FindByLayer(Layer.Name);
                        foreach (File3dmObject obj in m_objs)
                        {
                            GeometryBase geo = obj.Geometry;
                            if (geo is Rhino.Geometry.Curve)
                            {
                                Rhino.Geometry.Curve crv_ = geo as Rhino.Geometry.Curve;
                                curves_frames.Add(crv_);

                                Point3d end = crv_.PointAtEnd;
                                Point3d start = crv_.PointAtStart;

                                double x_end = end.X / 304.8;
                                double y_end = end.Y / 304.8;
                                double z_end = end.Z / 304.8;

                                double x_start = start.X / 304.8;
                                double y_start = start.Y / 304.8;
                                double z_start = start.Z / 304.8;

                                XYZ pt_end = new XYZ(x_end, y_end, z_end);
                                XYZ pt_start = new XYZ(x_start, y_start, z_start);

                                Autodesk.Revit.DB.Curve curve1 = Autodesk.Revit.DB.Line.CreateBound(pt_end, pt_start);
                                Autodesk.Revit.DB.Line line = Autodesk.Revit.DB.Line.CreateBound(pt_end, pt_start);

                                foreach (FamilyInstance fi_ in new FilteredElementCollector(doc).OfClass(typeof(FamilyInstance)).OfCategory(BuiltInCategory.OST_StructuralFraming).Cast<FamilyInstance>())
                                {
                                    famsimbol.Add(fi_);
                                }

                                FamilyInstance fi = doc.Create.NewFamilyInstance(curve1, beamSymbol, level, Autodesk.Revit.DB.Structure.StructuralType.Beam);

                                //try
                                //{
                                //    using (Transaction t = new Transaction(doc, "Make structural frame from Rhino"))
                                //    {
                                //        t.Start();

                                //        Autodesk.Revit.DB.Curve curve1 = Autodesk.Revit.DB.Line.CreateBound(pt_end, pt_start);
                                //        Autodesk.Revit.DB.Line line = Autodesk.Revit.DB.Line.CreateBound(pt_end, pt_start);


                                //        foreach (FamilyInstance fi_ in new FilteredElementCollector(doc).OfClass(typeof(FamilyInstance)).OfCategory(BuiltInCategory.OST_StructuralFraming).Cast<FamilyInstance>())
                                //        {
                                //            famsimbol.Add(fi_);
                                //        }

                                //        FamilyInstance fi = doc.Create.NewFamilyInstance(curve1, beamSymbol, level, Autodesk.Revit.DB.Structure.StructuralType.Beam);

                                //        t.Commit();
                                //    }

                                //}
                                //catch (Exception)
                                //{
                                //    //throw;
                                //}
                            }
                        }

                        t.Commit();
                    }
                }
            }
            return Autodesk.Revit.UI.Result.Succeeded;
        }
    }

    [Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)]
    public class duct_elbo : IExternalCommand
    {

        public static bool IsZero(double a)
        {
            const double _eps = 1.0e-9;
            return _eps > Math.Abs(a);
        }

        public static bool IsEqual(double a, double b)
        {
            return IsZero(b - a);
        }

        public static int Compare(double a, double b)
        {
            return IsEqual(a, b) ? 0 : (a < b ? -1 : 1);
        }

        public static int Compare(XYZ p, XYZ q)
        {

            int diff = Compare(p.X, q.X);
            if (0 == diff)
            {
                diff = Compare(p.Y, q.Y);
                if (0 == diff)
                {
                    diff = Compare(p.Z, q.Z);
                }

            }

            return diff;
        }

        static FilteredElementCollector GetConnectorElements(Autodesk.Revit.DB.Document doc,bool include_wires)
        {
            // what categories of family instances
            // are we interested in?

             BuiltInCategory[] bics = new BuiltInCategory[] {
                //BuiltInCategory.OST_CableTray,
                BuiltInCategory.OST_CableTrayFitting,
                //BuiltInCategory.OST_Conduit,
                BuiltInCategory.OST_ConduitFitting,
                //BuiltInCategory.OST_DuctCurves,
                BuiltInCategory.OST_DuctFitting,
                BuiltInCategory.OST_DuctTerminal,
                BuiltInCategory.OST_ElectricalEquipment,
                BuiltInCategory.OST_ElectricalFixtures,
                BuiltInCategory.OST_LightingDevices,
                BuiltInCategory.OST_LightingFixtures,
                BuiltInCategory.OST_MechanicalEquipment,
             //BuiltInCategory.OST_PipeCurves,
                BuiltInCategory.OST_PipeFitting,
                BuiltInCategory.OST_PlumbingFixtures,
                BuiltInCategory.OST_SpecialityEquipment,
                BuiltInCategory.OST_Sprinklers,
                //BuiltInCategory.OST_Wire,
            };

            IList<ElementFilter> a = new List<ElementFilter>(bics.Count());

            foreach (BuiltInCategory bic in bics)
            {
                a.Add(new ElementCategoryFilter(bic));
            }

            LogicalOrFilter categoryFilter = new LogicalOrFilter(a);

            LogicalAndFilter familyInstanceFilter = new LogicalAndFilter(categoryFilter, new ElementClassFilter(
                  typeof(FamilyInstance)));

            IList<ElementFilter> b = new List<ElementFilter>(6);

            b.Add(new ElementClassFilter(typeof(CableTray)));
            b.Add(new ElementClassFilter(typeof(Conduit)));
            b.Add(new ElementClassFilter(typeof(Duct)));
            b.Add(new ElementClassFilter(typeof(Pipe)));

            if (include_wires)
            {
                b.Add(new ElementClassFilter(typeof(Wire)));
            }

            b.Add(familyInstanceFilter);

            LogicalOrFilter classFilter
              = new LogicalOrFilter(b);

            FilteredElementCollector collector
              = new FilteredElementCollector(doc);

            collector.WherePasses(classFilter);

            return collector;
        }

        static ConnectorSet GetConnectors(Element e)
        {
            ConnectorSet connectors = null;

            if (e is FamilyInstance)
            {
                MEPModel m = ((FamilyInstance)e).MEPModel;
               
                if (null != m && null != m.ConnectorManager)
                {

                    connectors = m.ConnectorManager.Connectors;
                }
            }
            else if (e is Wire)
            {
                connectors = ((Wire)e) .ConnectorManager.Connectors;
            }
            else
            {
                Debug.Assert(
                  e.GetType().IsSubclassOf(typeof(MEPCurve)),
                  "expected all candidate connector provider "
                  + "elements to be either family instances or "
                  + "derived from MEPCurve");

                if (e is MEPCurve)
                {
                    connectors = ((MEPCurve)e).ConnectorManager.Connectors;
                }
            }

           
            return connectors;
        }

        static AddInId appId = new AddInId(new Guid("5F88CC78-A137-4809-AAF8-A478F3B24BAB"));
        public Autodesk.Revit.UI.Result Execute(ExternalCommandData commandData, ref string message, ElementSet elementSet)
        {


            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Autodesk.Revit.ApplicationServices.Application app = uiapp.Application;
            Autodesk.Revit.DB.Document doc = uidoc.Document;


            List<Autodesk.Revit.DB.XYZ> crvs = new List<Autodesk.Revit.DB.XYZ>();
            Autodesk.Revit.DB.NurbSpline line = null;

            //ICollection<Reference> my_lines = uidoc.Selection.PickObjects(Autodesk.Revit.UI.Selection.ObjectType.Element, "Select lines");
            //foreach (var item_myRefWall in my_lines)
            //{
            //    Element ele = doc.GetElement(item_myRefWall);
            //    GeometryObject geoobj = ele.GetGeometryObjectFromReference(item_myRefWall);
            //    Face face = geoobj as Face;
            //    LocationCurve locationCurve2 = ele.Location as LocationCurve;
            
            //    if (locationCurve2.Curve.GetType().Name == "NurbSpline")
            //    {
            //        line = locationCurve2.Curve as Autodesk.Revit.DB.NurbSpline;

            //    }
            //}

            Reference myRef2 = uidoc.Selection.PickObject(ObjectType.Element);
            Element e2 = doc.GetElement(myRef2.ElementId);
            ConnectorSet connector = GetConnectors(e2);

            List<XYZ> pts = new List<XYZ>();

            foreach (Connector item in connector)
            {
                pts.Add(item.Origin);
            }
            
            Reference myRef1 = uidoc.Selection.PickObject(ObjectType.Element);
            Element e1 = doc.GetElement(myRef1.ElementId);
            ConnectorSet connector2 = GetConnectors(e1);

            foreach (Connector item in connector2)
            {
                pts.Add(item.Origin);
            }

            List<double> distances = new List<double>();
            
            foreach (var item in pts)
            {

                foreach (var item2 in pts)
                {

                    int hola = Compare(item, item2);

                    if (Compare(item, item2) != 0) 
                    {
                        distances.Add(item.DistanceTo(item2));
                    }
                }
            }

            distances.Sort();

            foreach (Connector item in connector)
            {
                foreach (Connector item2 in connector2)
                {
                    if (item.Origin.DistanceTo(item2.Origin) == distances.First())
                    {
                        //var ductTypes = new FilteredElementCollector(doc)
                        //.OfClass(typeof(MechanicalFitting)).OfType<MechanicalFitting>().First();
                        using (Transaction tx = new Transaction(doc))
                        {
                            tx.Start("flex"
                              + "flex");
                            FamilyInstance flexDuct = doc.Create.NewTransitionFitting(item, item2);
                            tx.Commit();
                        }

                        continue;
                    }
                }
            }

            return Autodesk.Revit.UI.Result.Succeeded;
        }
    }

    [Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)]
    public class Create_Sun_Eye_study : IExternalCommand
    {
        public static List<List<XYZ>> GeTpoints(Autodesk.Revit.DB.Document doc_, List<List<XYZ>> xyz_faces, IList<CurveLoop> faceboundaries, List<List<Face>> list_faces)
        {
            if (list_faces == null)
            {
                list_faces = new List<List<Face>>();
            }

            for (int i = 0; i < list_faces.ToArray().Length; i++)
            {
                List<XYZ> puntos_ = new List<XYZ>();
                foreach (Face f in list_faces.ToArray()[i])
                {

                    faceboundaries = f.GetEdgesAsCurveLoops();//new trying to get the outline of the face instead of the edges
                    EdgeArrayArray edgeArrays = f.EdgeLoops;
                    foreach (CurveLoop edges in faceboundaries)
                    {
                        puntos_.Add(null);
                        foreach (Autodesk.Revit.DB.Curve edge in edges)
                        {
                            XYZ testPoint1 = edge.GetEndPoint(1);
                            XYZ testPoint2 = edge.GetEndPoint(0);
                            double lenght = Math.Round(edge.ApproximateLength, 0);
                            double lenght_convert = UnitUtils.Convert(lenght, DisplayUnitType.DUT_DECIMAL_FEET, DisplayUnitType.DUT_MILLIMETERS);

                            double x = Math.Round(testPoint1.X, 0);
                            double y = Math.Round(testPoint1.Y, 0);
                            double z = Math.Round(testPoint1.Z, 0);

                            ElementClassFilter filter = new ElementClassFilter(typeof(Floor));

                            XYZ newpt = new XYZ(x, y, z);

                            if (!puntos_.Contains(testPoint1))
                            {
                                puntos_.Add(testPoint1);

                            }
                        }
                    }
                    int num = f.EdgeLoops.Size;
                }
                xyz_faces.Add(puntos_);
            }
            return xyz_faces;
        }
        private ModelLine Makeline(Autodesk.Revit.DB.Document doc, XYZ pta, XYZ ptb)
        {
            ModelLine modelLine = null;
            double distance = pta.DistanceTo(ptb);
            if (distance < 0.01)
            {
                TaskDialog.Show("Error", "Distance" + distance);
                return modelLine;
            }

            XYZ norm = pta.CrossProduct(ptb);
            if (norm.GetLength() == 0)
            {
                XYZ aSubB = pta.Subtract(ptb);
                XYZ aSubBcrossz = aSubB.CrossProduct(XYZ.BasisZ);
                double crosslenght = aSubBcrossz.GetLength();
                if (crosslenght == 0)
                {
                    norm = XYZ.BasisY;
                }
                else
                {
                    norm = XYZ.BasisZ;
                }
            }

            Autodesk.Revit.DB.Plane plane = Autodesk.Revit.DB.Plane.CreateByNormalAndOrigin(norm, ptb);


            SketchPlane skplane = SketchPlane.Create(doc, plane);

            Autodesk.Revit.DB.Line line = Autodesk.Revit.DB.Line.CreateBound(pta, ptb);

            if (doc.IsFamilyDocument)
            {
                modelLine = doc.FamilyCreate.NewModelCurve(line, skplane) as ModelLine;
            }
            else
            {
                modelLine = doc.Create.NewModelCurve(line, skplane) as ModelLine;
            }
            if (modelLine == null)
            {
                TaskDialog.Show("Error", "Model line = null");
            }
            return modelLine;
        }

        public double MapValue(double start_n, double end_n, double mapped_n_menusone, double mapped_n_one, double number_tobe_map)
        {
            return mapped_n_menusone + (mapped_n_one - mapped_n_menusone) * ((number_tobe_map - start_n) / (end_n - start_n));
        }
        public ViewOrientation3D GetCurrentViewOrientation(UIDocument doc)
        {
            XYZ UpDir = doc.ActiveView.UpDirection;
            XYZ ViewDir = doc.ActiveView.ViewDirection;
            XYZ ViewInvDir = InvCoord(ViewDir);
            XYZ eye = new XYZ(0, 0, 0);
            XYZ up = UpDir;
            XYZ forward = ViewInvDir;
            ViewOrientation3D MyNewOrientation = new ViewOrientation3D(eye, up, forward);
            return MyNewOrientation;
        }

        public XYZ InvCoord(XYZ MyCoord)
        {
            XYZ invcoord = new XYZ((Convert.ToDouble(MyCoord.X * -1)),
                (Convert.ToDouble(MyCoord.Y * -1)),
                (Convert.ToDouble(MyCoord.Z * -1)));
            return invcoord;
        }
        public XYZ CrossProduct(XYZ v1, XYZ v2)
        {
            double x, y, z;
            x = v1.Y * v2.Z - v2.Y * v1.Z;
            y = (v1.X * v2.Z - v2.X * v1.Z) * -1;
            z = v1.X * v2.Y - v2.X * v1.Y;
            var rtnvector = new XYZ(x, y, z);
            return rtnvector;
        }
        public XYZ VectorFromHorizVertAngles(double angleHorizD, double angleVertD)
        {
            double degToRadian = Math.PI * 2 / 360;
            double angleHorizR = angleHorizD * degToRadian;
            double angleVertR = angleVertD * degToRadian;
            double a = Math.Cos(angleVertR);
            double b = Math.Cos(angleHorizR);
            double c = Math.Sin(angleHorizR);
            double d = Math.Sin(angleVertR);
            return new XYZ(a * b, a * c, d);
        }

        public class Vector3D
        {
            public Vector3D(XYZ revitXyz)
            {
                XYZ = revitXyz;
            }
            public Vector3D() : this(XYZ.Zero)
            { }
            public Vector3D(double x, double y, double z)
              : this(new XYZ(x, y, z))
            { }
            public XYZ XYZ { get; private set; }
            public double X => XYZ.X;
            public double Y => XYZ.Y;
            public double Z => XYZ.Z;
            public Vector3D CrossProduct(Vector3D source)
            {
                return new Vector3D(XYZ.CrossProduct(source.XYZ));
            }
            public double GetLength()
            {
                return XYZ.GetLength();
            }
            public override string ToString()
            {
                return XYZ.ToString();
            }
            public static Vector3D BasisX => new Vector3D(
              XYZ.BasisX);
            public static Vector3D BasisY => new Vector3D(
              XYZ.BasisY);
            public static Vector3D BasisZ => new Vector3D(
              XYZ.BasisZ);
        }
        static AddInId appId = new AddInId(new Guid("8D3F5703-A09A-6ED6-864C-5720329D9677"));
        public Autodesk.Revit.UI.Result Execute(ExternalCommandData commandData, ref string message, ElementSet elementSet)
        {
            UIDocument uidoc = commandData.Application.ActiveUIDocument;
            Autodesk.Revit.DB.Document doc = uidoc.Document;
            UIApplication uiapp = commandData.Application;
            Autodesk.Revit.ApplicationServices.Application app = doc.Application;

            try
            {
                string filename = @"T:\Transfer\lopez\Book1.xlsx";
                using (ExcelPackage package = new ExcelPackage(new FileInfo(filename)))
                {
                    ExcelWorksheet sheet = package.Workbook.Worksheets.ElementAt(0);

                    int column = 2;
                    int number = Convert.ToInt32(sheet.Cells[2, column].Value);
                    sheet.Cells[2, column].Value = (number + 1); ;
                    package.Save();
                }
            }
            catch (Exception)
            {
                MessageBox.Show("Excel file not found", "");
            }

            SunAndShadowSettings sunSettings = uidoc.ActiveView.SunAndShadowSettings; // get current settings from view
            Autodesk.Revit.DB.View currentView = uidoc.ActiveView;

            using (Transaction tx = new Transaction(doc))
            {
                tx.Start("_Sun study"+ "_Sun study");

                

                // modify Sun and Shadow Settings
                //DateTime sunrise = sunSettings.GetSunrise(DateTime.SpecifyKind(new DateTime(2019, 6, 20), DateTimeKind.Local)); // sunrise on April 20, 2011
                //DateTime sunset = sunSettings.GetSunset(DateTime.SpecifyKind(new DateTime(2019, 6, 22), DateTimeKind.Local)); // sunset on April 22, 2011
                sunSettings.SunAndShadowType = SunAndShadowType.OneDayStudy;
                

                sunSettings.StartDateAndTime = DateTime.SpecifyKind(new DateTime(2019, 6, 20,9,0,0), DateTimeKind.Local); /*sunrise.AddHours(2); // start 2 hours after sunrise on April 20, 2011*/
                sunSettings.EndDateAndTime = DateTime.SpecifyKind(new DateTime(2019, 6, 20, 15, 15, 0), DateTimeKind.Local); /*sunset.AddHours(-2); // end 2 hours before sunset on April 22, 2011*/
                if (sunSettings.IsTimeIntervalValid(SunStudyTimeInterval.Minutes15)) // check that this interval is valid for this SunAndShadowType
                    sunSettings.TimeInterval = SunStudyTimeInterval.Minutes15;

                // check for validity of start and end times
                if (!(sunSettings.IsAfterStartDateAndTime(sunSettings.EndDateAndTime)
                    && sunSettings.IsBeforeEndDateAndTime(sunSettings.StartDateAndTime)))
                    TaskDialog.Show("Error", "Start and End dates are invalid");

                
                

                uidoc.ActiveView.get_Parameter(BuiltInParameter.VIEW_GRAPH_SUN_PATH).Set(1); // turn on display of the sun path
                int time121 = sunSettings.ActiveFrameTime.Hour;
                TimeSpan time122= sunSettings.ActiveFrameTime.TimeOfDay;

                for (int i = 0; i < sunSettings.NumberOfFrames; i++)
                {
                    sunSettings.ActiveFrame = i;
                   

                    DateTime time9 = currentView.SunAndShadowSettings.GetFrameTime(1);
                    DateTime time8 = currentView.SunAndShadowSettings.GetFrameTime(i);
                    //time8.AddHours(1);
                    time8.AddMinutes(15);
                    time8.Add(new TimeSpan(0, 15, 0));
                    time9.AddMinutes(15);
                    time9.Add(new TimeSpan(0, 15, 0));




                    //Form26 form2 = new Form26();
                    //form2.ShowDialog();


                    ProjectLocation plCurrent = doc.ActiveProjectLocation;
                    Element projectInfoElement = new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_ProjectBasePoint).FirstElement();

                    BuiltInParameter bipAtn = BuiltInParameter.BASEPOINT_ANGLETON_PARAM;
                    Parameter patn = projectInfoElement.get_Parameter(bipAtn);
                    double atn = patn.AsDouble();
                    foreach (ProjectLocation location in doc.ProjectLocations)
                    {
                        ProjectPosition projectPosition
                          = location./*get_ProjectPosition(XYZ.Zero)*/ GetProjectPosition(XYZ.Zero);
                        double x = projectPosition.EastWest;
                        double y = projectPosition.NorthSouth;
                        XYZ pnp = new XYZ(x, y, 0.0);
                        double pna = projectPosition.Angle;
                    }
                    IEnumerable<ViewFamilyType> viewFamilyTypes = from elem in new FilteredElementCollector(doc).OfClass(typeof(ViewFamilyType))
                                                                  let type = elem as ViewFamilyType
                                                                  where type.ViewFamily == ViewFamily.ThreeDimensional
                                                                  select type;

                    Element ele = SunAndShadowSettings.GetActiveSunAndShadowSettings(doc);


                    //Autodesk.Revit.DB.View currentView = uidoc.ActiveView;
                    //SunAndShadowSettings sunSettings = currentView.SunAndShadowSettings;

                    

                    XYZ initialDirection = XYZ.BasisY;
                    double altitude = sunSettings.GetFrameAltitude(sunSettings.ActiveFrame);
                    Autodesk.Revit.DB.Transform altitudeRotation = Autodesk.Revit.DB.Transform.CreateRotation(XYZ.BasisX, altitude);
                    XYZ altitudeDirection = altitudeRotation.OfVector(initialDirection);
                    double azimuth = sunSettings.GetFrameAzimuth(sunSettings.ActiveFrame);
                    double actualAzimuth = 2 * Math.PI - azimuth;
                    Autodesk.Revit.DB.Transform azimuthRotation = Autodesk.Revit.DB.Transform.CreateRotation(XYZ.BasisZ, actualAzimuth);
                    double northrotation = 2 * Math.PI - atn;
                    XYZ sunDirection = azimuthRotation.OfVector(altitudeDirection);
                    Autodesk.Revit.DB.Transform tran01 = Autodesk.Revit.DB.Transform.CreateRotationAtPoint(XYZ.BasisZ, northrotation * -1, new XYZ(0, 0, 0));
                    XYZ new_p = tran01.OfVector(sunDirection);
                    sunDirection = new_p;

                    XYZ UpDir = uidoc.ActiveView.UpDirection;
                    //Form8 form = new Form8();

                    ViewOrientation3D viewOrientation3D;


                    View3D view3D = View3D.CreateIsometric(doc, viewFamilyTypes.First().Id);
                    //tr1.SetName("Create view " + view3D.Name);
                    view3D.Name = sunSettings.ActiveFrame.ToString()+ " - " + i.ToString(); /*form.textBox1.Text;*/

                    XYZ eye = XYZ.Zero;
                    XYZ inverted_sun_location = InvCoord(sunDirection);

                    XYZ origin_b = new XYZ(0, 0, 0);
                    XYZ normal_B = new XYZ(1, 0, 0);

                    Autodesk.Revit.DB.Plane Plane_mirror = Autodesk.Revit.DB.Plane.CreateByNormalAndOrigin(normal_B, sunDirection);
                    Autodesk.Revit.DB.Transform trans3;

                    //Autodesk.Revit.DB.Line line2 = Autodesk.Revit.DB.Line.CreateUnbound(new XYZ(5,5,5), sunDirection);
                    //Makeline(doc, line2.Origin, line2.Evaluate(5,false));


                    trans3 = Autodesk.Revit.DB.Transform.CreateReflection(Plane_mirror);
                    XYZ inv_sun_mirrored = trans3.OfVector(inverted_sun_location);

                    //Autodesk.Revit.DB.Line line3 = Autodesk.Revit.DB.Line.CreateUnbound(new XYZ(5, 5, 5), inv_sun_mirrored);
                    //Makeline(doc, line3.Origin, line3.Evaluate(5, false));


                    XYZ cross_product = CrossProduct(/*inv_sun,*/ /*new_p*/ inv_sun_mirrored, inverted_sun_location);
                    //Autodesk.Revit.DB.Line line4 = Autodesk.Revit.DB.Line.CreateUnbound(new XYZ(5, 5, 5), cross_product);
                    //Makeline(doc, line4.Origin, line4.Evaluate(5, false));


                    XYZ origin_c = sunDirection;
                    XYZ normal_c = cross_product;
                    Autodesk.Revit.DB.Plane Plane_mirror_c = Autodesk.Revit.DB.Plane.CreateByNormalAndOrigin(normal_c, origin_c);
                    XYZ no_z_90d = new XYZ(cross_product.X, cross_product.Y, 0);
                    XYZ no_z_orig = new XYZ(inverted_sun_location.X, inverted_sun_location.Y, 0);
                    Autodesk.Revit.DB.Line orientLine = Autodesk.Revit.DB.Line.CreateBound(sunDirection, inverted_sun_location) /*app.Create.NewLineBound(pickPoint, orientPoint)*/;
                    double angle_1 = /*XYZ.BasisX*/sunDirection.AngleTo(XYZ.BasisY);
                    double angleDegrees = angle_1 * 180 / Math.PI;
                    if (no_z_90d.X < no_z_orig.X)
                        angle_1 = 2 * Math.PI - angle_1;
                    double angleDegreesCorrected = angle_1 * 180 / Math.PI;
                    Autodesk.Revit.DB.Transform rot = Autodesk.Revit.DB.Transform.CreateRotation(orientLine.Direction, angleDegrees);
                    XYZ rotated_vec = rot.OfVector(-1 * cross_product);
                    XYZ dir = new XYZ(0, 0, 1);
                    XYZ normal = orientLine.Direction.Normalize();


                    XYZ cross = normal.CrossProduct(dir);
                    //Autodesk.Revit.DB.Line line5 = Autodesk.Revit.DB.Line.CreateUnbound(new XYZ(5, 5, 5), cross_product);
                    //Makeline(doc, line5.Origin, line5.Evaluate(5, false));



                    XYZ startPoint = sunDirection;
                    XYZ endPoint = inverted_sun_location;
                    Autodesk.Revit.DB.Line geomLine = Autodesk.Revit.DB.Line.CreateBound(startPoint, endPoint);
                    XYZ pntCenter = geomLine.Evaluate(0.0, true);

                    Autodesk.Revit.DB.Line geomLine2 = Autodesk.Revit.DB.Line.CreateBound(doc.ActiveView.Origin, XYZ.BasisZ);
                    Autodesk.Revit.DB.Plane geomPlane = Autodesk.Revit.DB.Plane.CreateByNormalAndOrigin(sunDirection, sunDirection);

                    //if (form.checkBox1.Checked)
                    //{
                    //    SketchPlane sketch = SketchPlane.Create(doc, geomPlane);
                    //    sketch.Name = view3D.Name;


                    //    doc.ActiveView.SketchPlane = sketch;
                    //    doc.ActiveView.ShowActiveWorkPlane();

                    //    view3D.SketchPlane = sketch;

                    //    view3D.ShowActiveWorkPlane();
                    //}
                    Autodesk.Revit.DB.Transform rot2 = Autodesk.Revit.DB.Transform.CreateRotation(orientLine.Direction, -1.590);
                    XYZ rotated_vec2 = rot2.OfVector(cross);
                    viewOrientation3D = new ViewOrientation3D(eye, /*rotated_vec*/  /*cross_product*/  rotated_vec2, inverted_sun_location);
                    view3D.SetOrientation(viewOrientation3D);
                    view3D.SaveOrientationAndLock();
                    //tr1.Commit();
                    //using (Transaction tr1 = new Transaction(doc))
                    //{
                    //    //form.ShowDialog();
                    //    tr1.Start("Place vs in sheet");

                    //}
                }

                tx.Commit();
            }

            return Autodesk.Revit.UI.Result.Succeeded;
        }
    }
    
    [Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)]
    public class tag_element : IExternalCommand
    {

        public static FilteredElementCollector GetElementsOfType(Autodesk.Revit.DB.Document doc, Type type, BuiltInCategory bic)
        {
            FilteredElementCollector collector
              = new FilteredElementCollector(doc);

            collector.OfCategory(bic);
            collector.OfClass(type);

            return collector;
        }

        static FilteredElementCollector GetFamilySymbols(Autodesk.Revit.DB.Document doc,BuiltInCategory bic)
        {
            return GetElementsOfType(doc,typeof(FamilySymbol), bic);
        }

        static FamilySymbol GetFirstFamilySymbol(Autodesk.Revit.DB.Document doc, BuiltInCategory bic)
        {
            FamilySymbol s = GetFamilySymbols(doc, bic)
              .FirstElement() as FamilySymbol;

            Debug.Assert(null != s, string.Format(
              "expected at least one {0} symbol in project",
              bic.ToString()));

            return s;
        }

        static bool GetBottomAndTopLevels(Autodesk.Revit.DB.Document doc, ref Level levelBottom, ref Level levelTop)
        {
            FilteredElementCollector levels = GetElementsOfType(doc, typeof(Level), BuiltInCategory.OST_Levels);

            foreach (Element e in levels)
            {
                if (null == levelBottom)
                {
                    levelBottom = e as Level;
                }
                else if (null == levelTop)
                {
                    levelTop = e as Level;
                }
                else
                {
                    break;
                }
            }

            if (levelTop.Elevation < levelBottom.Elevation)
            {
                Level tmp = levelTop;
                levelTop = levelBottom;
                levelBottom = tmp;
            }
            return null != levelBottom && null != levelTop;
        }

        public static Autodesk.Revit.DB.Line /*ModelLine*/ Makeline(Autodesk.Revit.DB.Document doc, XYZ pta, XYZ ptb)
        {
            /*ModelLine*/
            Autodesk.Revit.DB.Line modelLine = null;
            double distance = pta.DistanceTo(ptb);
            if (distance < 0.01)
            {
                TaskDialog.Show("Error", "Distance" + distance);
                return modelLine;
            }

            XYZ norm = pta.CrossProduct(ptb);
            if (norm.GetLength() == 0)
            {
                XYZ aSubB = pta.Subtract(ptb);
                XYZ aSubBcrossz = aSubB.CrossProduct(XYZ.BasisZ);
                double crosslenght = aSubBcrossz.GetLength();
                if (crosslenght == 0)
                {
                    norm = XYZ.BasisY;
                }
                else
                {
                    norm = XYZ.BasisZ;
                }
            }

            //Autodesk.Revit.DB.Plane plane = Autodesk.Revit.DB.Plane.CreateByNormalAndOrigin(norm, ptb);

            Autodesk.Revit.DB.Line line = Autodesk.Revit.DB.Line.CreateUnbound(ptb, XYZ.BasisZ /* XYZ.BasisZ*/);

            Autodesk.Revit.DB.Plane pl = Autodesk.Revit.DB.Plane.CreateByThreePoints(pta,
             line.Evaluate(5, false), ptb);

            SketchPlane skplane = SketchPlane.Create(doc, pl);

            Autodesk.Revit.DB.Line line2 = Autodesk.Revit.DB.Line.CreateBound(pta, ptb);

            modelLine = line2;

            //if (doc.IsFamilyDocument)
            //{
            //    modelLine = doc.FamilyCreate.NewModelCurve(line2, skplane) as ModelLine;
            //}
            //else
            //{
            //    modelLine = doc.Create.NewModelCurve(line2, skplane) as ModelLine;
            //}
            //if (modelLine == null)
            //{
            //    TaskDialog.Show("Error", "Model line = null");
            //}
            return modelLine;
        }
        public Autodesk.Revit.DB.Curve FindDuctCurve(Duct duct)
        {
            //The wind pipe curve
            IList<XYZ> list = new List<XYZ>();
            ConnectorSetIterator csi = duct.ConnectorManager.Connectors.ForwardIterator();
            while (csi.MoveNext())
            {
                Connector conn = csi.Current as Connector;
                list.Add(conn.Origin);
            }
            Autodesk.Revit.DB.Curve curve = Autodesk.Revit.DB.Line.CreateBound(list.ElementAt(0), 
                list.ElementAt(1)) as Autodesk.Revit.DB.Curve;
            //curve.MakeUnbound();
            return curve;
        }
        private IndependentTag CreateIndependentTag(Autodesk.Revit.DB.Document doc, Element e2)
        {
            Autodesk.Revit.DB.View view = doc.ActiveView;
            IndependentTag newTag;
            IndependentTag newTag2;
            IndependentTag newTag3;
            IndependentTag newTag4;
            TagMode tagMode = TagMode.TM_ADDBY_CATEGORY;
            double width = 0; ;
            Reference ref_ = null;
            Duct duct_ = e2 as Duct;

           
            try
            {
                width  = duct_.Width * 304.8;
                ref_ =new Reference(e2);
            }
            catch (Exception)
            {
                //TaskDialog.Show(e2.Name, e2.Id.ToString());
            }
            try
            {
                width = duct_.Diameter * 304.8;
                ref_ = new Reference(e2);
            }
            catch (Exception)
            {
                //    TaskDialog.Show(e2.Name, e2.Id.ToString());
            }



            //double lenght = Math.Round(edge.ApproximateLength, 0);
            //double lenght_convert = UnitUtils.Convert(lenght, DisplayUnitType.DUT_DECIMAL_FEET, DisplayUnitType.DUT_MILLIMETERS);

            Autodesk.Revit.DB.BoundingBoxXYZ bbox = e2.get_BoundingBox(null);

            XYZ pt0 = new XYZ(bbox.Min.X, bbox.Min.Y, bbox.Min.Z);
            XYZ pt1 = new XYZ(bbox.Max.X, bbox.Min.Y, bbox.Min.Z);
            XYZ pt2 = new XYZ(bbox.Max.X, bbox.Max.Y, bbox.Min.Z);
            XYZ pt3 = new XYZ(bbox.Min.X, bbox.Max.Y, bbox.Min.Z);
            Autodesk.Revit.DB.Line edge0 = Autodesk.Revit.DB.Line.CreateBound(pt0, pt1);
            Autodesk.Revit.DB.Line edge1 = Autodesk.Revit.DB.Line.CreateBound(pt1, pt2);
            Autodesk.Revit.DB.Line edge2 = Autodesk.Revit.DB.Line.CreateBound(pt2, pt3);
            Autodesk.Revit.DB.Line edge3 = Autodesk.Revit.DB.Line.CreateBound(pt3, pt0);
            List<Autodesk.Revit.DB.Curve> edges = new List<Autodesk.Revit.DB.Curve>();
            edges.Add(edge0);
            edges.Add(edge1);
            edges.Add(edge2);
            edges.Add(edge3);

            List<Double> shortlines = new List<double>();
            foreach (var cr in edges)
            {
                shortlines.Add(cr.Length);
            }
            //double height = bbox.Max.Z - bbox.Min.Z;
            //CurveLoop baseLoop = CurveLoop.Create(edges);
            //List<CurveLoop> loopList = new List<CurveLoop>();
            //loopList.Add(baseLoop);
            //Solid preTransformBox = GeometryCreationUtilities
            //  .CreateExtrusionGeometry(loopList, XYZ.BasisZ, height);
            //Solid transformBox = SolidUtils.CreateTransformed(
            //  preTransformBox, bbox.Transform);
            
            XYZ centrept = (new XYZ(bbox.Min.X, bbox.Min.Y, bbox.Min.Z) + new XYZ(bbox.Max.X, bbox.Max.Y, bbox.Min.Z)) / 2;

            /*ModelLine*/
            Autodesk.Revit.DB.Line   ml3 = null;
            /*ModelLine*/
            Autodesk.Revit.DB.Line ml4 = null;
            /*ModelLine*/
            Autodesk.Revit.DB.Line ml5 = null;
            /*ModelLine*/
            Autodesk.Revit.DB.Line  ml6 = null;
            /*ModelLine*/
            Autodesk.Revit.DB.Line ml7 = null;
            /*ModelLine*/
            Autodesk.Revit.DB.Line ml8 = null;

            shortlines.Sort();
            //shortlines.Reverse();
            List<Autodesk.Revit.DB.Curve> edgessorted = edges.OrderBy(x => x.Length).ToList();


            if (width > 350.0)
            {
                XYZ lined_dir_cpt_1 = centrept - edgessorted.ToArray()[0].Evaluate(0.5, true);
                XYZ lined_dir_cpt_2 = centrept - edgessorted.ToArray()[1].Evaluate(0.5, true);

                XYZ vec_moved1 = lined_dir_cpt_1.Normalize() * (250 / 304.8);
                XYZ vec_moved2 = lined_dir_cpt_2.Normalize() * (250 / 304.8);

                XYZ vect_edge1_pt1 = vec_moved1 + edgessorted.ToArray()[0].GetEndPoint(0);
                XYZ vect_edge1_pt2 = vec_moved1 + edgessorted.ToArray()[0].GetEndPoint(1);

                XYZ vect_edge2_pt1 = vec_moved2 + edgessorted.ToArray()[1].GetEndPoint(0);
                XYZ vect_edge2_pt2 = vec_moved2 + edgessorted.ToArray()[1].GetEndPoint(1);
           
                XYZ ln_dir_cpt_edge1_pt1 = edgessorted.ToArray()[0].Evaluate(0.5, true) - edgessorted.ToArray()[0].GetEndPoint(0);
                XYZ ln_dir_cpt_edge1_pt2 = edgessorted.ToArray()[0].Evaluate(0.5, true) - edgessorted.ToArray()[0].GetEndPoint(1);
                XYZ ln_dir_cpt_edge2_pt1 = edgessorted.ToArray()[1].Evaluate(0.5, true) - edgessorted.ToArray()[1].GetEndPoint(0);
                XYZ ln_dir_cpt_edge2_pt2 = edgessorted.ToArray()[1].Evaluate(0.5, true) - edgessorted.ToArray()[1].GetEndPoint(1);

                XYZ vec_moved3 = ln_dir_cpt_edge1_pt1.Normalize() * (150 / 304.8);
                XYZ vec_moved4 = ln_dir_cpt_edge1_pt2.Normalize() * (150 / 304.8);

                XYZ vec_moved5 = ln_dir_cpt_edge2_pt1.Normalize() * (150 / 304.8);
                XYZ vec_moved6 = ln_dir_cpt_edge2_pt2.Normalize() * (150 / 304.8);

                XYZ vect_edge1_pt3 = vec_moved3 + vect_edge1_pt1;
                XYZ vect_edge1_pt4 = vec_moved4 + vect_edge1_pt2;

                XYZ vect_edge2_pt5 = vec_moved5 + vect_edge2_pt1;
                XYZ vect_edge2_pt6 = vec_moved6 + vect_edge2_pt2;

                ml6 = Makeline(doc, edgessorted.ToArray()[0].GetEndPoint(0), vect_edge1_pt3);
                ml5 = Makeline(doc, edgessorted.ToArray()[0].GetEndPoint(1), vect_edge1_pt4);
                ml7 = Makeline(doc, edgessorted.ToArray()[1].GetEndPoint(0), vect_edge2_pt5);
                ml8 = Makeline(doc, edgessorted.ToArray()[1].GetEndPoint(1), vect_edge2_pt6);
                ml4 = Makeline(doc, centrept, edgessorted[1].GetEndPoint(0));
            }
            else
            {
                ml4 = Makeline(doc, edgessorted[0].Evaluate(0.5, true), edgessorted[1].Evaluate(0.5, true));
            }

            ml3 = Makeline(doc, edgessorted[2].Evaluate(0.5, true), edgessorted[3].Evaluate(0.5, true));
            var line = ml3;
            var dir = line.Direction;
            
            if (dir.X == -1 || dir.X == 1)
            {
                newTag = IndependentTag.Create(doc, view.Id, ref_, true, tagMode, TagOrientation.Vertical, centrept);
                newTag2 = IndependentTag.Create(doc, view.Id, ref_, true, tagMode, TagOrientation.Vertical, centrept);
                newTag3 = IndependentTag.Create(doc, view.Id, ref_, true, tagMode, TagOrientation.Vertical, ml4/*.GeometryCurve*/.Evaluate(0.7, false));
                newTag4 = IndependentTag.Create(doc, view.Id, ref_, true, tagMode, TagOrientation.Vertical, ml4/*.GeometryCurve*/.Evaluate(0.7, false));
            }
            else
            {
                newTag = IndependentTag.Create(doc, view.Id, ref_, true, tagMode, TagOrientation.Horizontal, centrept);
                newTag2 = IndependentTag.Create(doc, view.Id, ref_, true, tagMode, TagOrientation.Horizontal, centrept);
                newTag3 = IndependentTag.Create(doc, view.Id, ref_, true, tagMode, TagOrientation.Horizontal, ml4/*.GeometryCurve*/.Evaluate(0.7, false));
                newTag4 = IndependentTag.Create(doc, view.Id, ref_, true, tagMode, TagOrientation.Horizontal, ml4/*.GeometryCurve*/.Evaluate(0.7, false));

            }

            if (null == newTag)
            {
                throw new Exception("Create IndependentTag Failed.");
            }

            if (width > 350.0)
            {
                newTag.TagHeadPosition = ml6/*.GeometryCurve*/.Evaluate(1, true);
                newTag2.TagHeadPosition = centrept;
                newTag3.TagHeadPosition = ml5/*.GeometryCurve*/.Evaluate(1, true);
                newTag4.TagHeadPosition = ml8/*.GeometryCurve*/.Evaluate(1, true);

                newTag.HasLeader = false;
                newTag2.HasLeader = false;
                newTag3.HasLeader = false;
                newTag3.HasLeader = false;
            }
            else
            {
                newTag.TagHeadPosition = ml4/*.GeometryCurve*/.Evaluate(0.25, true);
                newTag2.TagHeadPosition = centrept;
                newTag3.TagHeadPosition = ml4/*.GeometryCurve*/.Evaluate(0.7, true);
                newTag4.TagHeadPosition = ml4/*.GeometryCurve*/.Evaluate(1, true);

                newTag.HasLeader = false;
                newTag2.HasLeader = false;
                newTag3.HasLeader = false;
                newTag3.HasLeader = false;
            }

            
            
            //FamilySymbol ductTagType = GetFirstFamilySymbol(doc, BuiltInCategory.OST_DuctTags);
            List<FamilySymbol> tags = new List<FamilySymbol>();
            IEnumerable<FamilySymbol> taglist = from elem in new FilteredElementCollector(doc).OfClass(typeof(FamilySymbol)).OfCategory(BuiltInCategory.OST_DuctTags)
                                               let type = elem as FamilySymbol where type.Name != null select type;
            tags = taglist.ToList();
            
            foreach (var tag in tags)
            {
                

                if (tag.Name == "STANDARD")
                {
                    newTag.ChangeTypeId(tag.Id);
                }
                if (tag.Name == "DUCT SIZE")
                {
                    newTag2.ChangeTypeId(tag.Id);
                }
                if (tag.Name == "DUCT LENGTH")
                {
                    newTag3.ChangeTypeId(tag.Id);
                }
            }
            //ductTagType = ductTagType.Duplicate( "New door tag type") as FamilySymbol;
            //var name = ductTagType.FamilyName;
            //var loc = ductTagType.Location;

            //newTag.ChangeTypeId(ductTagType.Id);
            return newTag;
        }

        static AddInId appId = new AddInId(new Guid("5F56AA78-A136-6509-AAF8-A478F3B24BAB"));
        public Autodesk.Revit.UI.Result Execute(ExternalCommandData commandData, ref string message, ElementSet elementSet)
        {
            UIDocument uidoc = commandData.Application.ActiveUIDocument;
            Autodesk.Revit.DB.Document doc = uidoc.Document;

            foreach (Element e in new FilteredElementCollector(doc, doc.ActiveView.Id).WhereElementIsNotElementType())
            {
                Category cat = e.Category;

                using (Transaction tx = new Transaction(doc))
                {
                    tx.Start("Make Line");

                    if (e.Category != null)
                    {
                        try
                        {
                            if (e.Category.Name == "Ducts")
                                CreateIndependentTag(doc, e);
                        }
                        catch (Exception)
                        {

                        }
                        
                       
                    }

                    tx.Commit();
                }
            }
            //FilteredElementCollector allElementsInView = new FilteredElementCollector(doc, doc.ActiveView.Id);
            //IList elementsInView = (IList)allElementsInView.ToElements();
            
            //Reference myRef2 = uidoc.Selection.PickObject(ObjectType.Element);
            //Element e2 = doc.GetElement(myRef2.ElementId);
            //GeometryObject geomObj2 = e2.GetGeometryObjectFromReference(myRef2);
            //Wall wall_ = e2 as Wall;
            
            //uidoc.Selection.SetElementIds(ele.Select(q => q.Id).ToList());
            return Autodesk.Revit.UI.Result.Succeeded;
        }
    }

   

   

    [Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)]
    public class create_sheet_excel : IExternalCommand
    {
        public static void DataMapping(List<string> keyData, List<List<string>> valueData)
        {
            List<Dictionary<string, string>> items = new List<Dictionary<string, string>>();

            string prompt = "Key/Value";
            prompt += Environment.NewLine;

            foreach (List<string> list in valueData)
            {
                for (int key = 0, value = 0; key < keyData.Count && value < list.Count; key++, value++)
                {
                    Dictionary<string, string> newItem = new Dictionary<string, string>();
                    string k = keyData[key];
                    string v = list[value];
                    newItem.Add(k, v);
                    items.Add(newItem);
                }
            }

            foreach (Dictionary<string, string> item in items)
            {
                foreach (KeyValuePair<string, string> kvp in item)
                {
                    if ((kvp.Key == "Count") && (kvp.Value == ""))
                        items.Remove(item);

                    prompt += "Key: " + kvp.Key + ",Value: " + kvp.Value;
                    prompt += Environment.NewLine;
                }
            }
            Autodesk.Revit.UI.TaskDialog.Show("Revit", prompt);
        }

        static AddInId appId = new AddInId(new Guid("5F88CC78-A137-4809-AAF8-A478F3B24BAB"));
        public Autodesk.Revit.UI.Result Execute(ExternalCommandData commandData, ref string message, ElementSet elementSet)
        {

            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Autodesk.Revit.ApplicationServices.Application app = uiapp.Application;
            Autodesk.Revit.DB.Document doc = uidoc.Document;

          

            string filename_excel= "";
            System.Windows.Forms.OpenFileDialog openDialog = new System.Windows.Forms.OpenFileDialog();
            openDialog.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
            //openDialog.Filter = "Excel Files (*.xlsx) |*.xslx)"; // TODO: Change to .csv
            if (openDialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                filename_excel = openDialog.FileName;
            }

            FilteredElementCollector collector = new FilteredElementCollector(doc);
            IList<Element> collection = collector.OfClass(typeof(ViewSchedule)).ToElements();
            List<List<string>> scheduleData = new List<List<string>>();
            String prompt = "ScheduleData :";
            prompt += Environment.NewLine;

            using (ExcelPackage package = new ExcelPackage(new FileInfo(filename_excel)))
            {

                package.Workbook.Worksheets.Delete(1);
                ExcelWorksheet sheet = package.Workbook.Worksheets.Add("my_data");

                foreach (Element e in collection)
                {
                    ViewSchedule viewSchedule = e as ViewSchedule;

                    TableData table = viewSchedule.GetTableData();

                    TableSectionData section = table.GetSectionData(SectionType.Body);


                    int nRows = section.NumberOfRows;
                    int nColumns = section.NumberOfColumns;

                    if (nRows > 1)
                    {
                        //valueData.Add(viewSchedule.Name);
                        for (int i = 0; i < nRows; i++)
                        {
                            List<string> rowData = new List<string>();

                            for (int j = 0; j < nColumns; j++)
                            {

                                try
                                {
                                    //rowData.Add(viewSchedule.GetCellText(SectionType.Header, i, j));
                                    

                                    rowData.Add(viewSchedule.GetCellText(SectionType.Body, i, j));

                                    int newnum1 = i++;
                                    int newnum2 = j++;


                                    sheet.Cells[i+1, j+1].Value = viewSchedule.GetCellText(SectionType.Body, i, j).ToString();
                                }
                                catch (Exception)
                                {
                                }
                            }
                            scheduleData.Add(rowData);
                        }
                        List<string> columnData = scheduleData[0];
                        scheduleData.RemoveAt(0);
                        DataMapping(columnData, scheduleData);
                    }
                }
                package.Save();
            }

            Process.Start(filename_excel);

            //if (openDialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            //{
            //    filename2 = openDialog.FileName;
            //}
            //List<string> names = new List<string>();
            //List <string> number  = new List<string>();
            //List<string> Drawing_Series = new List<string>();
            //List<string> View_Organization = new List<string>();

            //using (ExcelPackage package = new ExcelPackage(new FileInfo(filename2)))
            //{
            //    string data = "";
            //    ExcelWorksheet sheet = package.Workbook.Worksheets[1];
            //    for (int col = 1; col < 9999; col++)
            //    {
            //        var thisValue = sheet.Cells[1, col].Value;
            //        if (thisValue == null)
            //        {
            //            break;
            //        }
            //        if (thisValue as string == "Sheet Name") 
            //        {
            //            for (int row = 2; row < 9999; row++)
            //            {
            //                var name_list = sheet.Cells[row , col].Value;
            //                if (name_list != null)
            //                {
            //                    names.Add(name_list as string);
            //                    //if (Char.IsLetter(thisValue.ToString().First()))
            //                    //{
            //                    //    names.Add(thisValue as string);
            //                    //}
            //                    //else
            //                    //{
            //                    //}
            //                }
            //                else
            //                {
            //                    break;
            //                }
            //            }
            //            //break;
            //        }
            //        if (thisValue as string == "Sheet Number")
            //        {
            //            for (int row = 2; row < 9999; row++)
            //            {
            //                var name_list = sheet.Cells[row, col].Value;
            //                if (name_list != null)
            //                {
            //                    number.Add(name_list.ToString());
            //                }
            //                else
            //                {
            //                    break;
            //                }
            //            }
            //        }
            //        if (thisValue as string == "Drawing Series")
            //        {
            //            for (int row = 2; row < 9999; row++)
            //            {
            //                var value = sheet.Cells[row, col].Value;
            //                if (value != null)
            //                {
            //                    Drawing_Series.Add(value as string);
            //                }
            //                else
            //                {
            //                    break;
            //                }
            //            }
            //        }
            //        if (thisValue as string == "View Organization")
            //        {
            //            for (int row = 2; row < 9999; row++)
            //            {
            //                var value = sheet.Cells[row, col].Value;
            //                if (value != null)
            //                {
            //                    View_Organization.Add(value as string);
            //                }
            //                else
            //                {
            //                    break;
            //                }
            //            }
            //        }
            //    }
            //}
            //IEnumerable<FamilySymbol> familyList = from elem in new FilteredElementCollector(doc).OfClass(typeof(FamilySymbol)).OfCategory(BuiltInCategory.OST_TitleBlocks)
            //                                       let type = elem as FamilySymbol
            //                                       select type;
            //using (Transaction t = new Transaction(doc, "read excel"))
            //{
            //    t.Start();

            //    for (int i = 0; i < names.ToArray().Length; i++)
            //    {
            //        ViewSheet sheet2 = ViewSheet.Create(doc, familyList.First().Id);
            //        try
            //        {
            //            sheet2.Name = names.ToArray()[i] ;
            //            sheet2.SheetNumber = number.ToArray()[i] ;
            //        }
            //        catch (Exception)
            //        {
            //            MessageBox.Show("Name - "+ names.ToArray()[i] + " - might be already in use!", "");
            //        }
            //        try
            //        {
            //            sheet2.SheetNumber = number.ToArray()[i];
            //        }
            //        catch (Exception)
            //        {
            //            MessageBox.Show("Number - " + number.ToArray()[i] + " - might be already in use!", "");
            //        }
            //        try
            //        {
            //            sheet2.LookupParameter("Drawing Series").Set(Drawing_Series.ToArray()[i]);
            //        }
            //        catch (Exception)
            //        {
            //            MessageBox.Show("Drawing Series parameter does not exist", "");
            //        }
            //        try
            //        {
            //            sheet2.LookupParameter("View Organization").Set(View_Organization.ToArray()[i]);
            //        }
            //        catch (Exception)
            //        {
            //            MessageBox.Show("View Organization parameter does not exist", "");
            //        }
            //    }
            //    t.Commit();
            //}

            return Autodesk.Revit.UI.Result.Succeeded;
        }
    }

    [Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)]
    public class rhino_lns_to_csv : IExternalCommand
    {


        public static ModelLine Makeline(Autodesk.Revit.DB.Document doc, XYZ pta, XYZ ptb)
        {
            ModelLine modelLine = null;
            double distance = pta.DistanceTo(ptb);
            if (distance < 0.01)
            {
                TaskDialog.Show("Error", "Distance" + distance);
                return modelLine;
            }

            XYZ norm = pta.CrossProduct(ptb);
            if (norm.GetLength() == 0)
            {
                XYZ aSubB = pta.Subtract(ptb);
                XYZ aSubBcrossz = aSubB.CrossProduct(XYZ.BasisZ);
                double crosslenght = aSubBcrossz.GetLength();
                if (crosslenght == 0)
                {
                    norm = XYZ.BasisY;
                }
                else
                {
                    norm = XYZ.BasisZ;
                }
            }

            Autodesk.Revit.DB.Plane plane = Autodesk.Revit.DB.Plane.CreateByNormalAndOrigin(norm, ptb);


            SketchPlane skplane = SketchPlane.Create(doc, plane);

            Autodesk.Revit.DB.Line line = Autodesk.Revit.DB.Line.CreateBound(pta, ptb);

            if (doc.IsFamilyDocument)
            {
                modelLine = doc.FamilyCreate.NewModelCurve(line, skplane) as ModelLine;
            }
            else
            {
                modelLine = doc.Create.NewModelCurve(line, skplane) as ModelLine;
            }
            if (modelLine == null)
            {
                TaskDialog.Show("Error", "Model line = null");
            }
            return modelLine;
        }

        public static bool IsZero(double a)
        {
            const double _eps = 1.0e-9;
            return _eps > Math.Abs(a);
        }

        public static bool IsEqual(double a, double b)
        {
            return IsZero(b - a);
        }

        public static int Compare(double a, double b)
        {
            return IsEqual(a, b) ? 0 : (a < b ? -1 : 1);
        }

        public static int Compare(XYZ p, XYZ q)
        {

            int diff = Compare(p.X, q.X);
            if (0 == diff)
            {
                diff = Compare(p.Y, q.Y);
                if (0 == diff)
                {
                    diff = Compare(p.Z, q.Z);
                }

            }

            return diff;
        }

        private static Wall CreateWall(FamilyInstance cube, Autodesk.Revit.DB.Curve curve, double height)
        {
            var doc = cube.Document;

            var wallTypeId = doc.GetDefaultElementTypeId(
              ElementTypeGroup.WallType);

            return Wall.Create(doc, curve.CreateReversed(),
              wallTypeId, cube.LevelId, height, 0, false,
              false);
        }
        static AddInId appId = new AddInId(new Guid("5F88CC78-A137-4809-AAF8-A478F3B24BAB"));
        public Autodesk.Revit.UI.Result Execute(ExternalCommandData commandData, ref string message, ElementSet elementSet)
        {
            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Autodesk.Revit.DB.Document doc = uidoc.Document;

            try
            {
                string filename = @"T:\Transfer\lopez\Book1.xlsx";
                using (ExcelPackage package = new ExcelPackage(new FileInfo(filename)))
                {
                    ExcelWorksheet sheet = package.Workbook.Worksheets.ElementAt(0);

                    int column = 15;
                    int number = Convert.ToInt32(sheet.Cells[2, column].Value);
                    sheet.Cells[2, column].Value = (number + 1); ;
                    package.Save();
                }
            }



            catch (Exception)
            {
                MessageBox.Show("Excel file not found", "");
            }

            List<Object> objs = new List<Object>();
            List<List<XYZ>> xyz_faces = new List<List<XYZ>>();
            IList<Face> face_with_regions = new List<Face>();
            String info = "";
            List<List<Face>> Faces_lists_excel = new List<List<Face>>();
            //List<FaceArray> face112 = new List<FaceArray>();
            IList<CurveLoop> faceboundaries = new List<CurveLoop>();
            List<List<Element>> elemente_selected = new List<List<Element>>();
            List<List<string>> names = new List<List<string>>();
            List<int> numeros_ = new List<int>();
            XYZ pos_z = new XYZ(0, 0, 1);
            XYZ neg_z = new XYZ(0, 0, -1);

            string filename2 = "";
            System.Windows.Forms.OpenFileDialog openDialog = new System.Windows.Forms.OpenFileDialog();
            openDialog.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
            if (openDialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                filename2 = openDialog.FileName;
            }

            Point3d p1 = new Point3d(0, 0, 0);
            Rhino.Geometry.Point3d pt3d = new Point3d(10, 10, 0);

            File3dm m_modelfile = null;
            string m_name = null;
            string m_size = null;
            string m_created = null;
            string m_createdby = null;
            string m_edited = null;
            string m_editedby = null;
            string m_revision = null;
            string m_units = null;
            string m_notes = null;

            Object RhinoFile = filename2;
            if (RhinoFile is System.IO.FileInfo)
            {
                System.IO.FileInfo m_fileinfo = (System.IO.FileInfo)RhinoFile;
                m_modelfile = File3dm.Read(m_fileinfo.FullName);
                m_name = m_fileinfo.Name;
                m_size = m_fileinfo.Length.ToString();
                m_created = m_fileinfo.CreationTimeUtc.ToString();
                m_createdby = m_modelfile.CreatedBy;
                m_edited = m_fileinfo.LastWriteTimeUtc.ToString();
                m_editedby = m_modelfile.LastEditedBy;
                m_revision = m_modelfile.Revision.ToString();
                m_units = m_modelfile.Settings.ModelUnitSystem.ToString();
                m_notes = m_modelfile.Notes.Notes;
            }
            else if (RhinoFile is string)
            {
                System.IO.FileInfo m_fileinfo = new System.IO.FileInfo((string)RhinoFile);
                m_modelfile = File3dm.Read(m_fileinfo.FullName);
                m_name = m_fileinfo.Name;
                m_size = m_fileinfo.Length.ToString();
                m_created = m_fileinfo.CreationTimeUtc.ToString();
                m_createdby = m_modelfile.CreatedBy;
                m_edited = m_fileinfo.LastWriteTimeUtc.ToString();
                m_editedby = m_modelfile.LastEditedBy;
                m_revision = m_modelfile.Revision.ToString();
                m_units = m_modelfile.Settings.ModelUnitSystem.ToString();
                m_notes = m_modelfile.Notes.Notes;
            }

            File3dmLayerTable all_layers = m_modelfile.AllLayers;
            var objs_ = m_modelfile.Objects;
            List<Rhino.DocObjects.Layer> listadelayers = new List<Rhino.DocObjects.Layer>();
            List<Rhino.DocObjects.Layer> borrar = new List<Rhino.DocObjects.Layer>();
            List<List<Rhino.DocObjects.Layer>> multiplelistadelayers = new List<List<Rhino.DocObjects.Layer>>();
            string hola = "";

            MessageBox.Show("This tool will read 3D information only if the following Rhino Layers exist; Levels, Grids, Structure, Floor, Walls, Points", "!");

            foreach (var item in all_layers)
            {

                if (item.FullPath.Contains(hola))
                {
                    listadelayers.Add(item);
                }
            }

            List<Rhino.Geometry.Brep> rh_breps = new List<Rhino.Geometry.Brep>();
            List<XYZ> pts_ = new List<XYZ>();

            List<Autodesk.Revit.DB.Curve> revit_crv = new List<Autodesk.Revit.DB.Curve>();
            List<string> m_names = new List<string>();
            List<int> m_layerindeces = new List<int>();
            List<System.Drawing.Color> m_colors = new List<System.Drawing.Color>();
            List<string> m_guids = new List<string>();
            List<double> weith = new List<double>();



            foreach (var Layer in listadelayers)
            {


                if (Layer.Name == "Lines")
                {


                    File3dmObject[] m_objs = m_modelfile.Objects.FindByLayer(Layer.Name);

                    using (Transaction t = new Transaction(doc, "Lines from Rhino"))
                    {
                        t.Start();

                        foreach (File3dmObject obj in m_objs)
                        {
                            GeometryBase geo = obj.Geometry;
                            Type TYPE = geo.GetType();

                            if (TYPE.Name == "PolyCurve")
                            {
                                Rhino.Geometry.PolyCurve POLY = geo as Rhino.Geometry.PolyCurve;
                            }



                            if (TYPE.Name == "ArcCurve")
                            {
                                Rhino.Geometry.ArcCurve arc = geo as Rhino.Geometry.ArcCurve;

                                double x_end = arc.Arc.Plane.Normal.X / 304.8;
                                double y_end = arc.Arc.Plane.Normal.Y / 304.8;
                                double z_end = arc.Arc.Plane.Normal.Z / 304.8;

                                XYZ normal = new XYZ(x_end, y_end, z_end);

                                double x_ = arc.Arc.Plane.Origin.X / 304.8;
                                double y_ = arc.Arc.Plane.Origin.Y / 304.8;
                                double z_ = arc.Arc.Plane.Origin.Z / 304.8;
                                XYZ origin = new XYZ(x_, y_, z_);

                                double endAngle = 2 * Math.PI;        // this arc will be a circle

                                Autodesk.Revit.DB.Plane pl = Autodesk.Revit.DB.Plane.CreateByNormalAndOrigin(normal, origin);

                                SketchPlane skplane = SketchPlane.Create(doc, pl);
                                doc.Create.NewModelCurve(Autodesk.Revit.DB.Arc.Create(pl, arc.Radius / 304.8, arc.Arc.StartAngle, arc.Arc.EndAngle), skplane);

                            }

                            if (TYPE.Name == "PolylineCurve")
                            {
                                Rhino.Geometry.PolylineCurve PolylineCurve = geo as Rhino.Geometry.PolylineCurve;
                                Rhino.Geometry.Polyline Polyline = PolylineCurve.ToPolyline();

                                for (int i = 0; i < Polyline.SegmentCount; i++)
                                {
                                    Point3d end = Polyline.SegmentAt(i).PointAt(0);
                                    Point3d start = Polyline.SegmentAt(i).PointAt(1);


                                    double x_end = end.X / 304.8;
                                    double y_end = end.Y / 304.8;
                                    double z_end = end.Z / 304.8;

                                    double x_start = start.X / 304.8;
                                    double y_start = start.Y / 304.8;
                                    double z_start = start.Z / 304.8;

                                    XYZ pt_end = new XYZ(x_end, y_end, z_end);
                                    XYZ pt_start = new XYZ(x_start, y_start, z_start);

                                    Autodesk.Revit.DB.Curve curve1 = Autodesk.Revit.DB.Line.CreateBound(pt_end, pt_start);
                                    Autodesk.Revit.DB.Line line = Autodesk.Revit.DB.Line.CreateBound(pt_end, pt_start);

                                    Makeline(doc, line.GetEndPoint(0), line.GetEndPoint(1));
                                    revit_crv.Add(curve1);
                                }


                            }

                            if (TYPE.Name == "NurbsCurve")
                            {
                                Rhino.Geometry.NurbsCurve crv_ = geo as Rhino.Geometry.NurbsCurve;
                                Rhino.Geometry.Plane pl2;
                                crv_.TryGetPlane(out pl2);

                                double x_end = pl2.Normal.X / 304.8;
                                double y_end = pl2.Normal.Y / 304.8;
                                double z_end = pl2.Normal.Z / 304.8;

                                XYZ normal = new XYZ(x_end, y_end, z_end);

                                double x_2 = pl2.Origin.X / 304.8;
                                double y_2 = pl2.Origin.Y / 304.8;
                                double z_2 = pl2.Origin.Z / 304.8;
                                XYZ origin = new XYZ(x_2, y_2, z_2);

                                // this arc will be a circle

                                Autodesk.Revit.DB.Plane pl3 = Autodesk.Revit.DB.Plane.CreateByNormalAndOrigin(normal, origin);

                                var pts = crv_.Points;
                                foreach (var item in pts)
                                {
                                    double x_ = item.X / 304.8;
                                    double y_ = item.Y / 304.8;
                                    double z_ = item.Z / 304.8;

                                    XYZ pt = new XYZ(x_, y_, z_);
                                    weith.Add(item.Weight);
                                    pts_.Add(pt);
                                }

                                Autodesk.Revit.DB.Curve curve1 = Autodesk.Revit.DB.Line.CreateBound(pts_.First(), pts_.Last());

                                XYZ norm = pts_.First().CrossProduct(curve1.Evaluate(5, false));
                                if (norm.GetLength() == 0)
                                {
                                    XYZ aSubB = pts_.First().Subtract(pts_.Last());
                                    XYZ aSubBcrossz = aSubB.CrossProduct(XYZ.BasisZ);
                                    double crosslenght = aSubBcrossz.GetLength();
                                    if (crosslenght == 0)
                                    {
                                        norm = XYZ.BasisY;
                                    }
                                    else
                                    {
                                        norm = XYZ.BasisZ;
                                    }
                                }
                                //Autodesk.Revit.DB.Plane pl = Autodesk.Revit.DB.Plane.CreateByNormalAndOrigin(norm, pts_.First());
                                Autodesk.Revit.DB.Plane pl = Autodesk.Revit.DB.Plane.CreateByThreePoints(pts_.First(), pts_.ElementAt(2), pts_.Last());
                                SketchPlane skplane = SketchPlane.Create(doc, pl);

                                Autodesk.Revit.DB.Curve nspline = Autodesk.Revit.DB.NurbSpline.CreateCurve(pts_, weith);
                                doc.Create.NewModelCurve(nspline, skplane);

                                pts_.Clear();
                                weith.Clear();
                            }

                            if (TYPE.Name == "Curve")
                            {
                                Rhino.Geometry.Curve crv_ = geo as Rhino.Geometry.Curve;


                                Point3d end = crv_.PointAtEnd;
                                Point3d start = crv_.PointAtStart;

                                double x_end = end.X / 304.8;
                                double y_end = end.Y / 304.8;
                                double z_end = end.Z / 304.8;

                                double x_start = start.X / 304.8;
                                double y_start = start.Y / 304.8;
                                double z_start = start.Z / 304.8;

                                XYZ pt_end = new XYZ(x_end, y_end, z_end);
                                XYZ pt_start = new XYZ(x_start, y_start, z_start);

                                Autodesk.Revit.DB.Curve curve1 = Autodesk.Revit.DB.Line.CreateBound(pt_end, pt_start);
                                Autodesk.Revit.DB.Line line = Autodesk.Revit.DB.Line.CreateBound(pt_end, pt_start);

                                Makeline(doc, line.GetEndPoint(0), line.GetEndPoint(1));
                                revit_crv.Add(curve1);
                            }
                        }
                        t.Commit();
                    }
                }
            }
            return Autodesk.Revit.UI.Result.Succeeded;
        }
    }

    [Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)]
    public class wall_data : IExternalCommand
    {

        static AddInId appId = new AddInId(new Guid("5F88CC78-A137-4809-AAF8-A478F3B24BAC"));
        public Autodesk.Revit.UI.Result Execute(ExternalCommandData commandData, ref string message, ElementSet elementSet)
        {
            UIDocument uidoc = commandData.Application.ActiveUIDocument;
            Autodesk.Revit.DB.Document doc = uidoc.Document;

            FilteredElementCollector rooms1 = new FilteredElementCollector(doc).WhereElementIsNotElementType().OfClass(typeof(SpatialElement));
            ICollection<Element> room2 = rooms1.ToElements();
            SpatialElementBoundaryOptions opt = new SpatialElementBoundaryOptions();
            opt.SpatialElementBoundaryLocation = SpatialElementBoundaryLocation.Center;
            //Level level = doc.GetElement(uidoc.Selection.PickObject(ObjectType.Element, "Select level")) as Level;
            List<List<XYZ>> lista_nombres = new List<List<XYZ>>();
            List<string> names = new List<string>();

           
            List<Room> rooms = new List<Room>();
           
            List<XYZ> ptlist = new List<XYZ>();
            

            string filename3 = "";
            System.Windows.Forms.OpenFileDialog openDialog2 = new System.Windows.Forms.OpenFileDialog();
            openDialog2.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
            //openDialog.Filter = "Excel Files (*.xlsx) |*.xslx)"; // TODO: Change to .csv
            if (openDialog2.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                filename3 = openDialog2.FileName;
            }
            using (Transaction tx = new Transaction(doc))
            {
                tx.Start("ver dim");

                using (StreamWriter writer2 = new StreamWriter(filename3))
                {
                    foreach (var item in new FilteredElementCollector(doc).OfClass(typeof(Wall)))
                    {
                        Wall w = item as Wall;

                        Options op = new Options();
                        op.ComputeReferences = true;

                       

                        foreach (var item2 in /*item*/w.get_Geometry(op).Where(q => q is Solid).Cast<Solid>())
                        {

                            Solid solid = item2 as Solid;
                            foreach (Face item3 in solid.Faces)
                            {
                                PlanarFace planarFace = item3 as PlanarFace;
                                XYZ normal = planarFace.ComputeNormal(new UV(planarFace.Origin.X, planarFace.Origin.Y));


                                //Element e = doc.GetElement(item3.Reference);
                                //GeometryObject geoobj = e.GetGeometryObjectFromReference(item3.Reference);
                                //Face face = geoobj as Face;
                                

                                foreach (var edges in planarFace.GetEdgesAsCurveLoops() /*face.GetEdgesAsCurveLoops()*/)
                                {

                                    foreach (Autodesk.Revit.DB.Curve edge in edges)
                                    {
                                        XYZ testPoint1 = edge.GetEndPoint(1);
                                        XYZ testPoint2 = edge.GetEndPoint(0);
                                        double lenght = Math.Round(edge.ApproximateLength, 0);
                                        double lenght_convert = UnitUtils.Convert(lenght, DisplayUnitType.DUT_DECIMAL_FEET, DisplayUnitType.DUT_MILLIMETERS);

                                        double x = Math.Round(testPoint1.X, 0);
                                        double y = Math.Round(testPoint1.Y, 0);
                                        double z = Math.Round(testPoint1.Z, 0);

                                        ElementClassFilter filter = new ElementClassFilter(typeof(Floor));
                                        XYZ dir = new XYZ(0, 0, 0) - testPoint1;

                                        string x0 = edge.GetEndPoint(0).X.ToString();
                                        string y0 = edge.GetEndPoint(0).Y.ToString();
                                        string z0 = edge.GetEndPoint(0).Z.ToString();

                                        string x1 = edge.GetEndPoint(1).X.ToString();
                                        string y1 = edge.GetEndPoint(1).Y.ToString();
                                        string z1 = edge.GetEndPoint(1).Z.ToString();
                                        writer2.WriteLine(x0 + "," + y0 + "," + z0 + "," + x1 + "," + y1 + "," + z1 );

                                    }
                                }
                            }
                        }
                    }
                }
                tx.Commit();
            }
            return Autodesk.Revit.UI.Result.Succeeded;
        }
    }

    [Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)]
    public class floor_data : IExternalCommand
    {
        static AddInId appId = new AddInId(new Guid("5F88CC78-A137-4809-AAF8-A478F3B24BBE"));
        public Autodesk.Revit.UI.Result Execute(ExternalCommandData commandData, ref string message, ElementSet elementSet)
        {
            UIDocument uidoc = commandData.Application.ActiveUIDocument;
            Autodesk.Revit.DB.Document doc = uidoc.Document;

            FilteredElementCollector rooms1 = new FilteredElementCollector(doc).WhereElementIsNotElementType().OfClass(typeof(SpatialElement));
            ICollection<Element> room2 = rooms1.ToElements();
            SpatialElementBoundaryOptions opt = new SpatialElementBoundaryOptions();
            opt.SpatialElementBoundaryLocation = SpatialElementBoundaryLocation.Center;
            //Level level = doc.GetElement(uidoc.Selection.PickObject(ObjectType.Element, "Select level")) as Level;
            List<List<XYZ>> lista_nombres = new List<List<XYZ>>();
            List<string> names = new List<string>();


            List<Room> rooms = new List<Room>();

            List<XYZ> ptlist = new List<XYZ>();


            string filename3 = "";
            System.Windows.Forms.OpenFileDialog openDialog2 = new System.Windows.Forms.OpenFileDialog();
            openDialog2.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
            //openDialog.Filter = "Excel Files (*.xlsx) |*.xslx)"; // TODO: Change to .csv
            if (openDialog2.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                filename3 = openDialog2.FileName;
            }
            using (Transaction tx = new Transaction(doc))
            {
                tx.Start("ver dim");

                using (StreamWriter writer2 = new StreamWriter(filename3))
                {
                    foreach (var item in new FilteredElementCollector(doc).OfClass(typeof(Floor)))
                    {
                        Options op = new Options();
                        op.ComputeReferences = true;
                        foreach (var item2 in item.get_Geometry(op).Where(q => q is Solid).Cast<Solid>())
                        {
                            foreach (Face item3 in item2.Faces)
                            {
                                PlanarFace planarFace = item3 as PlanarFace;
                                XYZ normal = planarFace.ComputeNormal(new UV(planarFace.Origin.X, planarFace.Origin.Y));


                                //Element e = doc.GetElement(item3.Reference);
                                //GeometryObject geoobj = e.GetGeometryObjectFromReference(item3.Reference);
                                //Face face = geoobj as Face;


                                foreach (var edges in planarFace.GetEdgesAsCurveLoops() /*face.GetEdgesAsCurveLoops()*/)
                                {

                                    foreach (Autodesk.Revit.DB.Curve edge in edges)
                                    {
                                        XYZ testPoint1 = edge.GetEndPoint(1);
                                        XYZ testPoint2 = edge.GetEndPoint(0);
                                        double lenght = Math.Round(edge.ApproximateLength, 0);
                                        double lenght_convert = UnitUtils.Convert(lenght, DisplayUnitType.DUT_DECIMAL_FEET, DisplayUnitType.DUT_MILLIMETERS);

                                        double x = Math.Round(testPoint1.X, 0);
                                        double y = Math.Round(testPoint1.Y, 0);
                                        double z = Math.Round(testPoint1.Z, 0);

                                        ElementClassFilter filter = new ElementClassFilter(typeof(Floor));
                                        XYZ dir = new XYZ(0, 0, 0) - testPoint1;

                                        string x0 = edge.GetEndPoint(0).X.ToString();
                                        string y0 = edge.GetEndPoint(0).Y.ToString();
                                        string z0 = edge.GetEndPoint(0).Z.ToString();

                                        string x1 = edge.GetEndPoint(1).X.ToString();
                                        string y1 = edge.GetEndPoint(1).Y.ToString();
                                        string z1 = edge.GetEndPoint(1).Z.ToString();
                                        writer2.WriteLine(x0 + "," + y0 + "," + z0 + "," + x1 + "," + y1 + "," + z1);

                                    }
                                }
                            }
                        }
                    }
                }
                tx.Commit();
            }
            return Autodesk.Revit.UI.Result.Succeeded;
        }
    }

    [Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)]
    public class roof_data : IExternalCommand
    {
        static AddInId appId = new AddInId(new Guid("5F88CC78-A137-4809-AAF8-A478F3B24BBE"));
        public Autodesk.Revit.UI.Result Execute(ExternalCommandData commandData, ref string message, ElementSet elementSet)
        {
            UIDocument uidoc = commandData.Application.ActiveUIDocument;
            Autodesk.Revit.DB.Document doc = uidoc.Document;

            FilteredElementCollector rooms1 = new FilteredElementCollector(doc).WhereElementIsNotElementType().OfClass(typeof(SpatialElement));
            ICollection<Element> room2 = rooms1.ToElements();
            SpatialElementBoundaryOptions opt = new SpatialElementBoundaryOptions();
            opt.SpatialElementBoundaryLocation = SpatialElementBoundaryLocation.Center;
            //Level level = doc.GetElement(uidoc.Selection.PickObject(ObjectType.Element, "Select level")) as Level;
            List<List<XYZ>> lista_nombres = new List<List<XYZ>>();
            List<string> names = new List<string>();


            List<Room> rooms = new List<Room>();

            List<XYZ> ptlist = new List<XYZ>();


            string filename3 = "";
            System.Windows.Forms.OpenFileDialog openDialog2 = new System.Windows.Forms.OpenFileDialog();
            openDialog2.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
            //openDialog.Filter = "Excel Files (*.xlsx) |*.xslx)"; // TODO: Change to .csv
            if (openDialog2.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                filename3 = openDialog2.FileName;
            }
            using (Transaction tx = new Transaction(doc))
            {
                tx.Start("ver dim");

                using (StreamWriter writer2 = new StreamWriter(filename3))
                {
                    foreach (var item in new FilteredElementCollector(doc).OfClass(typeof(Ceiling)))
                    {
                        Options op = new Options();
                        op.ComputeReferences = true;
                        foreach (var item2 in item.get_Geometry(op).Where(q => q is Solid).Cast<Solid>())
                        {
                            foreach (Face item3 in item2.Faces)
                            {
                                PlanarFace planarFace = item3 as PlanarFace;
                                XYZ normal = planarFace.ComputeNormal(new UV(planarFace.Origin.X, planarFace.Origin.Y));


                                //Element e = doc.GetElement(item3.Reference);
                                //GeometryObject geoobj = e.GetGeometryObjectFromReference(item3.Reference);
                                //Face face = geoobj as Face;


                                foreach (var edges in planarFace.GetEdgesAsCurveLoops() /*face.GetEdgesAsCurveLoops()*/)
                                {

                                    foreach (Autodesk.Revit.DB.Curve edge in edges)
                                    {
                                        XYZ testPoint1 = edge.GetEndPoint(1);
                                        XYZ testPoint2 = edge.GetEndPoint(0);
                                        double lenght = Math.Round(edge.ApproximateLength, 0);
                                        double lenght_convert = UnitUtils.Convert(lenght, DisplayUnitType.DUT_DECIMAL_FEET, DisplayUnitType.DUT_MILLIMETERS);

                                        double x = Math.Round(testPoint1.X, 0);
                                        double y = Math.Round(testPoint1.Y, 0);
                                        double z = Math.Round(testPoint1.Z, 0);

                                        ElementClassFilter filter = new ElementClassFilter(typeof(Floor));
                                        XYZ dir = new XYZ(0, 0, 0) - testPoint1;

                                        string x0 = edge.GetEndPoint(0).X.ToString();
                                        string y0 = edge.GetEndPoint(0).Y.ToString();
                                        string z0 = edge.GetEndPoint(0).Z.ToString();

                                        string x1 = edge.GetEndPoint(1).X.ToString();
                                        string y1 = edge.GetEndPoint(1).Y.ToString();
                                        string z1 = edge.GetEndPoint(1).Z.ToString();
                                        writer2.WriteLine(x0 + "," + y0 + "," + z0 + "," + x1 + "," + y1 + "," + z1);

                                    }
                                }
                            }
                        }
                    }
                }
                tx.Commit();
            }
            return Autodesk.Revit.UI.Result.Succeeded;
        }
    }

    [Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)]
    public class duct_data : IExternalCommand
    {

        public static FilteredElementCollector GetElementsOfType(Autodesk.Revit.DB.Document doc, Type type, BuiltInCategory bic)
        {
            FilteredElementCollector collector
              = new FilteredElementCollector(doc);

            collector.OfCategory(bic);
            collector.OfClass(type);

            return collector;
        }

        public static bool IsZero(double a)
        {
            const double _eps = 1.0e-9;
            return _eps > Math.Abs(a);
        }

        public static bool IsEqual(double a, double b)
        {
            return IsZero(b - a);
        }

        public static int Compare(double a, double b)
        {
            return IsEqual(a, b) ? 0 : (a < b ? -1 : 1);
        }

        public static int Compare(XYZ p, XYZ q)
        {

            int diff = Compare(p.X, q.X);
            if (0 == diff)
            {
                diff = Compare(p.Y, q.Y);
                if (0 == diff)
                {
                    diff = Compare(p.Z, q.Z);
                }

            }

            return diff;
        }

        public static ModelLine Makeline(Autodesk.Revit.DB.Document doc, XYZ pta, XYZ ptb)
        {
            ModelLine modelLine = null;
            double distance = pta.DistanceTo(ptb);
            if (distance < 0.01)
            {
                TaskDialog.Show("Error", "Distance" + distance);
                return modelLine;
            }

            XYZ norm = pta.CrossProduct(ptb);
            if (norm.GetLength() == 0)
            {
                XYZ aSubB = pta.Subtract(ptb);
                XYZ aSubBcrossz = aSubB.CrossProduct(XYZ.BasisZ);
                double crosslenght = aSubBcrossz.GetLength();
                if (crosslenght == 0)
                {
                    norm = XYZ.BasisY;
                }
                else
                {
                    norm = XYZ.BasisZ;
                }
            }

            //Autodesk.Revit.DB.Plane plane = Autodesk.Revit.DB.Plane.CreateByNormalAndOrigin(norm, ptb);

            Autodesk.Revit.DB.Line line = Autodesk.Revit.DB.Line.CreateUnbound(ptb, XYZ.BasisZ /* XYZ.BasisZ*/);

            Autodesk.Revit.DB.Plane pl = Autodesk.Revit.DB.Plane.CreateByThreePoints(pta,
             line.Evaluate(5, false), ptb);

            SketchPlane skplane = SketchPlane.Create(doc, pl);

            Autodesk.Revit.DB.Line line2 = Autodesk.Revit.DB.Line.CreateBound(pta, ptb);

            if (doc.IsFamilyDocument)
            {
                modelLine = doc.FamilyCreate.NewModelCurve(line2, skplane) as ModelLine;
            }
            else
            {
                modelLine = doc.Create.NewModelCurve(line2, skplane) as ModelLine;
            }
            if (modelLine == null)
            {
                TaskDialog.Show("Error", "Model line = null");
            }
            return modelLine;
        }
        
        static AddInId appId = new AddInId(new Guid("5F10CC78-A137-4809-BAF9-A478F3B24BAB"));
        public Autodesk.Revit.UI.Result Execute(ExternalCommandData commandData, ref string message, ElementSet elementSet)
        {
            UIDocument uidoc = commandData.Application.ActiveUIDocument;
            Autodesk.Revit.DB.Document doc = uidoc.Document;

            string filename3 = "";
            System.Windows.Forms.OpenFileDialog openDialog2 = new System.Windows.Forms.OpenFileDialog();
            openDialog2.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
            //openDialog.Filter = "Excel Files (*.xlsx) |*.xslx)"; // TODO: Change to .csv
            if (openDialog2.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                filename3 = openDialog2.FileName;
            }

            using (Transaction tx = new Transaction(doc))
            {
                tx.Start("ver dim");

                using (StreamWriter writer2 = new StreamWriter(filename3))
                {
                    foreach (var item in new FilteredElementCollector(doc).OfClass(typeof(Duct)))
                    {
                        Options op = new Options();
                        op.ComputeReferences = true;
                        foreach (var item2 in item.get_Geometry(op).Where(q => q is Solid).Cast<Solid>())
                        {
                            foreach (Face item3 in item2.Faces)
                            {
                                //PlanarFace planarFace = item3 as PlanarFace;
                                //XYZ normal = planarFace.ComputeNormal(new UV(planarFace.Origin.X, planarFace.Origin.Y));


                                Element e = doc.GetElement(item3.Reference);
                                GeometryObject geoobj = e.GetGeometryObjectFromReference(item3.Reference);
                                Face face = geoobj as Face;


                                foreach (var edges in face.GetEdgesAsCurveLoops() /*face.GetEdgesAsCurveLoops()*/)
                                {

                                    foreach (Autodesk.Revit.DB.Curve edge in edges)
                                    {
                                        XYZ testPoint1 = edge.GetEndPoint(1);
                                        XYZ testPoint2 = edge.GetEndPoint(0);
                                        double lenght = Math.Round(edge.ApproximateLength, 0);
                                        double lenght_convert = UnitUtils.Convert(lenght, DisplayUnitType.DUT_DECIMAL_FEET, DisplayUnitType.DUT_MILLIMETERS);

                                        double x = Math.Round(testPoint1.X, 0);
                                        double y = Math.Round(testPoint1.Y, 0);
                                        double z = Math.Round(testPoint1.Z, 0);

                                        ElementClassFilter filter = new ElementClassFilter(typeof(Floor));
                                        XYZ dir = new XYZ(0, 0, 0) - testPoint1;



                                        string x0 = edge.GetEndPoint(0).X.ToString();
                                        string y0 = edge.GetEndPoint(0).Y.ToString();
                                        string z0 = edge.GetEndPoint(0).Z.ToString();

                                        string x1 = edge.GetEndPoint(1).X.ToString();
                                        string y1 = edge.GetEndPoint(1).Y.ToString();
                                        string z1 = edge.GetEndPoint(1).Z.ToString();
                                        writer2.WriteLine(x0 + "," + y0 + "," + z0 + "," + x1 + "," + y1 + "," + z1);

                                    }
                                }
                            }
                        }
                    }
                }
                tx.Commit();
            }

            //foreach (Element e in new FilteredElementCollector(doc, doc.ActiveView.Id).WhereElementIsNotElementType())
            //{
            //    Category cat = e.Category;
            //    Duct duct_ ;

            //    using (Transaction tx = new Transaction(doc))
            //    {
            //        tx.Start("Make Line");

            //        if (e.Category != null)
            //        {
            //            try
            //            {
            //                if (e.Category.Name == "Ducts")
            //                {
            //                    duct_ = e as Duct;
            //                    Autodesk.Revit.DB.BoundingBoxXYZ bbox = e.get_BoundingBox(null);

            //                    XYZ pt0 = new XYZ(bbox.Min.X, bbox.Min.Y, bbox.Min.Z);
            //                    XYZ pt1 = new XYZ(bbox.Max.X, bbox.Min.Y, bbox.Min.Z);
            //                    XYZ pt2 = new XYZ(bbox.Max.X, bbox.Max.Y, bbox.Min.Z);
            //                    XYZ pt3 = new XYZ(bbox.Min.X, bbox.Max.Y, bbox.Min.Z);
            //                    Autodesk.Revit.DB.Line edge0 = Autodesk.Revit.DB.Line.CreateBound(pt0, pt1);
            //                    Autodesk.Revit.DB.Line edge1 = Autodesk.Revit.DB.Line.CreateBound(pt1, pt2);
            //                    Autodesk.Revit.DB.Line edge2 = Autodesk.Revit.DB.Line.CreateBound(pt2, pt3);
            //                    Autodesk.Revit.DB.Line edge3 = Autodesk.Revit.DB.Line.CreateBound(pt3, pt0);
            //                    List<Autodesk.Revit.DB.Curve> edges = new List<Autodesk.Revit.DB.Curve>();

            //                    Makeline(doc, pt0, pt1);
            //                    Makeline(doc, pt1, pt2);
            //                    Makeline(doc, pt2, pt3);
            //                    Makeline(doc, pt3, pt0);
            //                } 
            //            }
            //            catch (Exception)
            //            {

            //            }
            //        }
            //        tx.Commit();
            //    }
            //}

            return Autodesk.Revit.UI.Result.Succeeded;
        }
    }

    [Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)]
    public class Curtain_wall_data : IExternalCommand
    {
        List<List<XYZ>> GetElementSolids(Element e)
        {
            List<List<XYZ>> ptmesh = new List<List<XYZ>>();

            Options opt = new Options();
            opt.ComputeReferences = true;
            opt.DetailLevel = ViewDetailLevel.Fine;

            // Get geometry element of the selected element
            Autodesk.Revit.DB.GeometryElement geoElement = e.get_Geometry(opt);

            // Get geometry object
            foreach (GeometryObject geoObject in geoElement)
            {
                // Get the geometry instance which contains the geometry information
                Autodesk.Revit.DB.GeometryInstance instance = geoObject as Autodesk.Revit.DB.GeometryInstance;
                if (null != instance)
                {
                    foreach (GeometryObject instObj in instance.SymbolGeometry)
                    {
                        Solid solid = instObj as Solid;
                        if (null == solid || 0 == solid.Faces.Size || 0 == solid.Edges.Size)
                        {
                            continue;
                        }

                        Autodesk.Revit.DB.Transform instTransform = instance.Transform;
                        // Get the faces and edges from solid, and transform the formed points

                        
                        foreach (Face face in solid.Faces)
                        {
                            

                            Autodesk.Revit.DB.Mesh mesh = face.Triangulate();
                            foreach (XYZ ii in mesh.Vertices)
                            {

                                XYZ point = ii;
                                XYZ transformedPoint = instTransform.OfPoint(point);
                               
                            }
                        }
                        

                        foreach (Edge edge in solid.Edges)
                        {
                            List<XYZ> pt = new List<XYZ>();
                            foreach (XYZ ii in edge.Tessellate())
                            {
                                XYZ point = ii;
                                XYZ transformedPoint = instTransform.OfPoint(point);
                                pt.Add(transformedPoint);
                            }
                            ptmesh.Add(pt);
                        }
                        
                        
                    }
                }
            }
            return ptmesh;
        }

      
        void GetCurtainWallPanelGeometry(Autodesk.Revit.DB.Document doc,ElementId curtainWallId, List<List<XYZ>> XYZ_)
        {
            // First, find solid geometry from panel ids.
            // Note that the panel which contains a basic
            // wall has NO geometry!

            Wall wall = doc.GetElement(curtainWallId) as Wall;
            var grid = wall.CurtainGrid;

            foreach (ElementId id in grid.GetPanelIds())
            {
                Element e = doc.GetElement(id);

                XYZ_.AddRange(GetElementSolids(e));
            }

            // Secondly, find corresponding panel wall
            // for the curtain wall and retrieve the actual
            // geometry from that.

            FilteredElementCollector cwPanels
              = new FilteredElementCollector(doc)
                .OfCategory(BuiltInCategory.OST_CurtainWallPanels)
                .OfClass(typeof(Wall));

            foreach (Wall cwp in cwPanels)
            {
                // Find panel wall belonging to this curtain wall
                // and retrieve its geometry

                if (cwp.StackedWallOwnerId == curtainWallId)
                {
                    XYZ_.AddRange(GetElementSolids(cwp));
                }
            }
        }
       

        #region list_wall_geom
        void list_wall_geom(Wall w/*, Autodesk.Revit.ApplicationServices.Application app*/)
        {
            string s = "";

            CurtainGrid cgrid = w.CurtainGrid;

            Options options
              = /*app.Create.NewGeometryOptions()*/ new Options() ;

            options.ComputeReferences = true;
            options.IncludeNonVisibleObjects = true;

            GeometryElement geomElem
              = w.get_Geometry(options);

            foreach (GeometryObject obj in geomElem)
            {
                Visibility vis = obj.Visibility;

                string visString = vis.ToString();

              
                Solid solid = obj as Solid;

              
                if (solid != null)
                {
                    int faceCount = solid.Faces.Size;

                    s += "Faces: " + faceCount + "\n";

                    foreach (Face face in solid.Faces)
                    {
                        s += "Face area (" + visString + "): "
                          + face.Area + "\n";
                    }
                }
               
            }
            TaskDialog.Show("revit", s);
        }
        #endregion // list_wall_geom

        public static Element SelectSingleElementOfType(UIDocument uidoc, Type t,string description,bool acceptDerivedClass)
        {
            Element e = GetSingleSelectedElement(uidoc);

            if (!HasRequestedType(e, t, acceptDerivedClass))
            {
                e = SelectSingleElement(uidoc, description);
            }
            return HasRequestedType(e, t, acceptDerivedClass)
              ? e
              : null;
        }

        public static Element SelectSingleElement(UIDocument uidoc,string description)
        {
            if (ViewType.Internal == uidoc.ActiveView.ViewType)
            {
                TaskDialog.Show("Error",
                  "Cannot pick element in this view: "
                  + uidoc.ActiveView.Name);

                return null;
            }
            try
            {
                Reference r = uidoc.Selection.PickObject(
                  ObjectType.Element,
                  "Please select " + description);

                // 'Autodesk.Revit.DB.Reference.Element' is
                // obsolete: Property will be removed. Use
                // Document.GetElement(Reference) instead.
                //return null == r ? null : r.Element; // 2011

                return uidoc.Document.GetElement(r); // 2012
            }
            catch (Autodesk.Revit.Exceptions.OperationCanceledException)
            {
                return null;
            }
        }

        public static Element GetSingleSelectedElement(UIDocument uidoc)
        {
            ICollection<ElementId> ids
              = uidoc.Selection.GetElementIds();

            Element e = null;

            if (1 == ids.Count)
            {
                foreach (ElementId id in ids)
                {
                    e = uidoc.Document.GetElement(id);
                }
            }
            return e;
        }

        static bool HasRequestedType(Element e, Type t,bool acceptDerivedClass)
        {
            bool rc = null != e;

            if (rc)
            {
                Type t2 = e.GetType();

                rc = t2.Equals(t);

                if (!rc && acceptDerivedClass)
                {
                    rc = t2.IsSubclassOf(t);
                }
            }
            return rc;
        }

        static AddInId appId = new AddInId(new Guid("6F10CC78-A137-4806-BAF9-A468F3B24BAB"));
        public Autodesk.Revit.UI.Result Execute(ExternalCommandData commandData, ref string message, ElementSet elementSet)
        {
            UIDocument uidoc = commandData.Application.ActiveUIDocument;
            Autodesk.Revit.DB.Document doc = uidoc.Document;

            FilteredElementCollector rooms1 = new FilteredElementCollector(doc).WhereElementIsNotElementType().OfClass(typeof(SpatialElement));
            ICollection<Element> room2 = rooms1.ToElements();
            SpatialElementBoundaryOptions opt = new SpatialElementBoundaryOptions();
            opt.SpatialElementBoundaryLocation = SpatialElementBoundaryLocation.Center;
            //Level level = doc.GetElement(uidoc.Selection.PickObject(ObjectType.Element, "Select level")) as Level;
            List<List<XYZ>> lista_nombres = new List<List<XYZ>>();
            List<string> names = new List<string>();

            List<GeometryElement> geolist = new List<GeometryElement>();
            List<Room> rooms = new List<Room>();

            List<XYZ> ptlist = new List<XYZ>();
            List<Solid> solids = new List<Solid>();

            string filename3 = "";
            System.Windows.Forms.OpenFileDialog openDialog2 = new System.Windows.Forms.OpenFileDialog();
            openDialog2.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
            //openDialog.Filter = "Excel Files (*.xlsx) |*.xslx)"; // TODO: Change to .csv
            if (openDialog2.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                filename3 = openDialog2.FileName;
            }


            UIApplication uiapp = commandData.Application;

            Autodesk.Revit.ApplicationServices.Application app = uiapp.Application;

            //Wall wall = SelectSingleElementOfType(uidoc, typeof(Wall), "a curtain wall", false) as Wall;

            List<List<XYZ>> xyz_ = new List<List<XYZ>>();

            foreach (Element e in new FilteredElementCollector(doc, doc.ActiveView.Id).WhereElementIsNotElementType())
            {
                Category cat = e.Category;

                using (Transaction tx = new Transaction(doc))
                {
                    tx.Start("Curtain Walls");

                    if (e.Category != null)
                    {
                        try
                        {
                            if (e.Category.Name == "Walls")
                            {
                                
                                Wall wall = e as Wall;
                               
                                var grid = wall.CurtainGrid;

                                if (grid != null)
                                {
                                    GetCurtainWallPanelGeometry(doc, wall.Id, xyz_);

                                    using (StreamWriter writer2 = new StreamWriter(filename3))
                                    {
                                        if (null == wall)
                                        {
                                            message = "Please select a single " + "curtain wall element.";

                                            return Result.Failed;
                                        }
                                        else
                                        {

                                            for (int i = 0; i < xyz_.ToArray().Length; i++)
                                            {
                                                string x0 = xyz_.ToArray()[i][0].X.ToString();
                                                string y0 = xyz_.ToArray()[i][0].Y.ToString();
                                                string z0 = xyz_.ToArray()[i][0].Z.ToString();

                                                string x1 = xyz_.ToArray()[i][1].X.ToString();
                                                string y1 = xyz_.ToArray()[i][1].Y.ToString();
                                                string z1 = xyz_.ToArray()[i][1].Z.ToString();

                                                writer2.WriteLine(x0 + "," + y0 + "," + z0 + "," + x1 + "," + y1 + "," + z1);
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

                    tx.Commit();
                }
            }

            
            //using (StreamWriter writer2 = new StreamWriter(filename3))
            //{
            //    if (null == wall)
            //    {
            //        message = "Please select a single " + "curtain wall element.";

            //        return Result.Failed;
            //    }
            //    else
            //    {
            //        using (Transaction t = new Transaction(doc))
            //        {
            //            t.Start("panel faces");


            //            foreach (var face in faces_)
            //            {
            //                PlanarFace planarFace = face as PlanarFace;
            //                //XYZ normal = planarFace.ComputeNormal(new UV(planarFace.Origin.X, planarFace.Origin.Y));

            //                //Element e = doc.GetElement(item3.Reference);
            //                //GeometryObject geoobj = e.GetGeometryObjectFromReference(item3.Reference);
            //                //Face face = geoobj as Face;
            //                foreach (var edges in planarFace.GetEdgesAsCurveLoops() /*face.GetEdgesAsCurveLoops()*/)
            //                {
            //                    foreach (Autodesk.Revit.DB.Curve edge in edges)
            //                    {
            //                        XYZ testPoint1 = edge.GetEndPoint(1);
            //                        XYZ testPoint2 = edge.GetEndPoint(0);
            //                        double lenght = Math.Round(edge.ApproximateLength, 0);
            //                        double lenght_convert = UnitUtils.Convert(lenght, DisplayUnitType.DUT_DECIMAL_FEET, DisplayUnitType.DUT_MILLIMETERS);

            //                        double x = Math.Round(testPoint1.X, 0);
            //                        double y = Math.Round(testPoint1.Y, 0);
            //                        double z = Math.Round(testPoint1.Z, 0);

            //                        ElementClassFilter filter = new ElementClassFilter(typeof(Floor));
            //                        XYZ dir = new XYZ(0, 0, 0) - testPoint1;

            //                        string x0 = edge.GetEndPoint(0).X.ToString();
            //                        string y0 = edge.GetEndPoint(0).Y.ToString();
            //                        string z0 = edge.GetEndPoint(0).Z.ToString();

            //                        string x1 = edge.GetEndPoint(1).X.ToString();
            //                        string y1 = edge.GetEndPoint(1).Y.ToString();
            //                        string z1 = edge.GetEndPoint(1).Z.ToString();
            //                        writer2.WriteLine(x0 + "," + y0 + "," + z0 + "," + x1 + "," + y1 + "," + z1);

            //                    }
            //                }
            //            }

            //            t.Commit();
            //        }
            //    }
            //} 
            return Autodesk.Revit.UI.Result.Succeeded;
        }
    }

    [Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)]
    public class Stair_data : IExternalCommand
    {
        List<List<XYZ>> GetElementSolids(Element e)
        {
            List<List<XYZ>> ptmesh = new List<List<XYZ>>();

            Options opt = new Options();
            opt.ComputeReferences = true;
            opt.DetailLevel = ViewDetailLevel.Fine;

            // Get geometry element of the selected element
            Autodesk.Revit.DB.GeometryElement geoElement = e.get_Geometry(opt);

            // Get geometry object
            foreach (GeometryObject geoObject in geoElement)
            {
                if (null != geoObject) //+
                {
                    Solid solid = geoObject as Solid;
                    try
                    {
                        if (null == solid || 0 == solid.Faces.Size || 0 == solid.Edges.Size)
                        {
                            //continue;
                        }
                        foreach (Face face in solid.Faces)
                        {
                            foreach (Edge edge in solid.Edges)
                            {
                                List<XYZ> pt = new List<XYZ>();
                                foreach (XYZ ii in edge.Tessellate())
                                {
                                    pt.Add(ii);
                                }
                                ptmesh.Add(pt);
                            }
                        }
                    }
                    catch (Exception)
                    {

                        //throw;
                    }
                    
                    
                }

                if (geoObject is Autodesk.Revit.DB.GeometryInstance)
                {
                    // Get the geometry instance which contains the geometry information
                    Autodesk.Revit.DB.GeometryInstance instance = geoObject as Autodesk.Revit.DB.GeometryInstance;
                    if (null != instance)
                    {
                        foreach (GeometryObject instObj in instance.SymbolGeometry)
                        {
                            Solid solid = instObj as Solid;
                            if (null == solid || 0 == solid.Faces.Size || 0 == solid.Edges.Size)
                            {
                                continue;
                            }

                            Autodesk.Revit.DB.Transform instTransform = instance.Transform;
                            // Get the faces and edges from solid, and transform the formed points


                            foreach (Face face in solid.Faces)
                            {


                                Autodesk.Revit.DB.Mesh mesh = face.Triangulate();
                                foreach (XYZ ii in mesh.Vertices)
                                {

                                    XYZ point = ii;
                                    XYZ transformedPoint = instTransform.OfPoint(point);

                                }
                            }


                            foreach (Edge edge in solid.Edges)
                            {
                                List<XYZ> pt = new List<XYZ>();
                                foreach (XYZ ii in edge.Tessellate())
                                {
                                    XYZ point = ii;
                                    XYZ transformedPoint = instTransform.OfPoint(point);
                                    pt.Add(transformedPoint);
                                }
                                ptmesh.Add(pt);
                            }


                        }
                    }
                }

                
            }
            return ptmesh;
        }

        static AddInId appId = new AddInId(new Guid("ESB2C3B9-S146-4FED-AFA2-498B5DB74E44"));
        public Autodesk.Revit.UI.Result Execute(ExternalCommandData commandData, ref string message, ElementSet elementSet)
        {
            UIDocument uidoc = commandData.Application.ActiveUIDocument;
            Autodesk.Revit.DB.Document doc = uidoc.Document;
            

            List<List<XYZ>> xyz_ = new List<List<XYZ>>();


            string filename3 = "";
            System.Windows.Forms.OpenFileDialog openDialog2 = new System.Windows.Forms.OpenFileDialog();
            openDialog2.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
            //openDialog.Filter = "Excel Files (*.xlsx) |*.xslx)"; // TODO: Change to .csv
            if (openDialog2.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                filename3 = openDialog2.FileName;
            }
            using (Transaction tx = new Transaction(doc))
            {
                tx.Start("ver dim");

                using (StreamWriter writer2 = new StreamWriter(filename3))
                {


                    foreach (var e in new FilteredElementCollector(doc).OfClass(typeof(Stairs)))
                    {
                        //Stairs w = item as Stairs;

                        Options op = new Options();
                        op.ComputeReferences = true;

                        xyz_ = GetElementSolids(e);

                        if (null == e)
                        {
                            message = "Please select a single " + "curtain wall element.";

                            return Result.Failed;
                        }
                        else
                        {

                            for (int i = 0; i < xyz_.ToArray().Length; i++)
                            {
                                string x0 = xyz_.ToArray()[i][0].X.ToString();
                                string y0 = xyz_.ToArray()[i][0].Y.ToString();
                                string z0 = xyz_.ToArray()[i][0].Z.ToString();

                                string x1 = xyz_.ToArray()[i][1].X.ToString();
                                string y1 = xyz_.ToArray()[i][1].Y.ToString();
                                string z1 = xyz_.ToArray()[i][1].Z.ToString();

                                writer2.WriteLine(x0 + "," + y0 + "," + z0 + "," + x1 + "," + y1 + "," + z1);
                            }
                        }
                    }
                }
                tx.Commit();
            }
            return Autodesk.Revit.UI.Result.Succeeded;
        }
    }

    [Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)]
    public class Section_data : IExternalCommand
    {
        /// <summary>
        /// A class to count and report the 
        /// number of objects encountered.
        /// </summary>
        class JtObjCounter : Dictionary<string, int>
        {
            /// <summary>
            /// Count a new occurence of an object
            /// </summary>
            public void Increment(object obj)
            {
                string key = null == obj
                  ? "null"
                  : obj.GetType().Name;

                if (!ContainsKey(key))
                {
                    Add(key, 0);
                }
                ++this[key];
            }

            /// <summary>
            /// Report the number of objects encountered.
            /// </summary>
            public void Print()
            {
                List<string> keys = new List<string>(Keys);
                keys.Sort();
                foreach (string key in keys)
                {
                    Debug.Print("{0,5} {1}", this[key], key);
                }
            }
        }

        /// <summary>
        /// Maximum distance for line to be 
        /// considered to lie in plane
        /// </summary>
        const double _eps = 1.0e-6;

        /// <summary>
        /// User instructions for running this external command
        /// </summary>
        const string _instructions = "Please launch this "
          + "command in a section view with fine level of "
          + "detail and far bound clipping set to 'Clip with line'";

        /// <summary>
        /// Predicate returning true if the given line 
        /// lies in the given plane
        /// </summary>
        static bool IsLineInPlane(Autodesk.Revit.DB.Line line, Autodesk.Revit.DB.Plane plane)
        {
            XYZ p0 = line.GetEndPoint(0);
            XYZ p1 = line.GetEndPoint(1);
            UV uv0, uv1;
            double d0, d1;

            plane.Project(p0, out uv0, out d0);
            plane.Project(p1, out uv1, out d1);

            Debug.Assert(0 <= d0,
              "expected non-negative distance");
            Debug.Assert(0 <= d1,
              "expected non-negative distance");

            return (_eps > d0) && (_eps > d1);
        }

        static void GetCurvesInPlane(List<Autodesk.Revit.DB.Curve> curves, JtObjCounter geoCounter, Autodesk.Revit.DB.Plane plane, GeometryElement geo)
        {
            geoCounter.Increment(geo);

            if (null != geo)
            {
                foreach (GeometryObject obj in geo)
                {
                    geoCounter.Increment(obj);

                    Solid sol = obj as Solid;

                    if (null != sol)
                    {
                        EdgeArray edges = sol.Edges;

                        foreach (Edge edge in edges)
                        {
                            Autodesk.Revit.DB.Curve curve = edge.AsCurve();

                            Debug.Assert(curve is Autodesk.Revit.DB.Line,
                              "we currently only support lines here");

                            geoCounter.Increment(curve);

                            if (IsLineInPlane(curve as Autodesk.Revit.DB.Line, plane))
                            {
                                curves.Add(curve);
                            }
                        }
                    }
                    else
                    {
                        GeometryInstance inst = obj as GeometryInstance;

                        if (null != inst)
                        {
                            GetCurvesInPlane(curves, geoCounter,
                              plane, inst.GetInstanceGeometry());
                        }
                        else
                        {
                            Debug.Assert(false,
                              "unsupported geometry object "
                              + obj.GetType().Name);
                        }
                    }
                }
            }
        }

        static AddInId appId = new AddInId(new Guid("EEB2C3B9-D146-4FED-AFA2-498B5DB84E44"));
        public Autodesk.Revit.UI.Result Execute(ExternalCommandData commandData, ref string message, ElementSet elementSet)
        {

            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Autodesk.Revit.ApplicationServices. Application app = uiapp.Application;
            Autodesk.Revit.DB.Document doc = uidoc.Document;
            Autodesk.Revit.DB.View section_view = commandData.View;
            Parameter p = section_view.get_Parameter(
              BuiltInParameter.VIEWER_BOUND_FAR_CLIPPING);

            if (ViewType.Section != section_view.ViewType
              || ViewDetailLevel.Fine != section_view.DetailLevel
              || 1 != p.AsInteger())
            {
                message = _instructions;
                return Result.Failed;
            }

            FilteredElementCollector a = new FilteredElementCollector(doc, section_view.Id);

            Options opt = new Options()
            {
                ComputeReferences = false,
                IncludeNonVisibleObjects = false,
                View = section_view
            };

            SketchPlane plane1 = section_view.SketchPlane; // this is null

            Autodesk.Revit.DB.Plane plane2 = Autodesk.Revit.DB.Plane.CreateByNormalAndOrigin(
              section_view.ViewDirection,
              section_view.Origin);

            JtObjCounter geoCounter = new JtObjCounter();

            List<Autodesk.Revit.DB.Curve> curves = new List<Autodesk.Revit.DB.Curve>();

            foreach (Element e in a)
            {
                geoCounter.Increment(e);

                GeometryElement geo = e.get_Geometry(opt);

                GetCurvesInPlane(curves, geoCounter, plane2, geo);
            }

            Debug.Print("Objects analysed:");geoCounter.Print();

            Debug.Print("{0} cut geometry lines found in section plane.",curves.Count);

            using (Transaction tx = new Transaction(doc))
            {
                tx.Start("Create Section Cut Model Curves");

                SketchPlane plane3 = SketchPlane.Create(doc, plane2);

                foreach (Autodesk.Revit.DB.Curve c in curves)
                {
                    doc.Create.NewModelCurve(c, plane3);
                }

                tx.Commit();
            }
        


            return Autodesk.Revit.UI.Result.Succeeded;
        }
    }

    [Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)]
    public class Modelline_data : IExternalCommand
    {
        static AddInId appId = new AddInId(new Guid("03D171A0-077E-472F-9AC3-A24BB38FE7B7"));
        public Autodesk.Revit.UI.Result Execute(ExternalCommandData commandData, ref string message, ElementSet elementSet)
        {
            UIDocument uidoc = commandData.Application.ActiveUIDocument;
            Autodesk.Revit.DB.Document doc = uidoc.Document;

           
            List<Room> rooms = new List<Room>();

            List<XYZ> ptlist = new List<XYZ>();


            string filename3 = "";
            System.Windows.Forms.OpenFileDialog openDialog2 = new System.Windows.Forms.OpenFileDialog();
            openDialog2.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
            //openDialog.Filter = "Excel Files (*.xlsx) |*.xslx)"; // TODO: Change to .csv
            if (openDialog2.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                filename3 = openDialog2.FileName;
            }
            using (Transaction tx = new Transaction(doc))
            {
                tx.Start("ver dim");

                using (StreamWriter writer2 = new StreamWriter(filename3))
                {
                    foreach (var item in new FilteredElementCollector(doc).OfClass(typeof(CurveElement)))
                    {
                        try
                        {
                            LocationCurve Line = item.Location as LocationCurve;
                            Autodesk.Revit.DB.Curve cr = Line.Curve;


                            XYZ pt1 = cr.GetEndPoint(0);
                            XYZ pt2 = cr.GetEndPoint(1);

                            string x0 = pt1.X.ToString();
                            string y0 = pt1.Y.ToString();
                            string z0 = pt1.Z.ToString();

                            string x1 = pt2.X.ToString();
                            string y1 = pt2.Y.ToString();
                            string z1 = pt2.Z.ToString();
                            writer2.WriteLine(x0 + "," + y0 + "," + z0 + "," + x1 + "," + y1 + "," + z1);

                        }
                        catch (Exception)
                        {

                        }
                    }
                }
                tx.Commit();
            }
            return Autodesk.Revit.UI.Result.Succeeded;
        }
    }

    [Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)]
    public class View_by_Cat : IExternalCommand
    {
        static AddInId appId = new AddInId(new Guid("5F92CC78-A237-4809-AAF8-A478F3B24BAB"));
        public Autodesk.Revit.UI.Result Execute(ExternalCommandData commandData, ref string message, ElementSet elementSet)
        {
            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Autodesk.Revit.DB.Document doc = uidoc.Document;

            string appdataFolder = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
            string folderPath = Path.Combine(appdataFolder, @"FBX");

            IEnumerable<ViewFamilyType> viewFamilyTypes = from elem in new FilteredElementCollector(doc).OfClass(typeof(ViewFamilyType))
                                                          let type = elem as ViewFamilyType
                                                          where type.ViewFamily == ViewFamily.ThreeDimensional
                                                          select type;
            List<string> catlist = new List<string>();


            FilteredElementCollector col = new FilteredElementCollector(doc, doc.ActiveView.Id);

            
            foreach (Element e in col)
            {
                if (e.Category != null)
                {
                    Category cat = e.Category;
                    if (!catlist.Contains(cat.Name))
                    {
                        if (cat.CategoryType == CategoryType.Model)
                        {
                            catlist.Add(cat.Name);
                        }
                    }
                    
                }
            }
            
            using (Transaction ttNew = new Transaction(doc, "creating 3Dviews"))
            {
                ttNew.Start();

                for (int i = 0; i < catlist.ToArray().Length; i++)
                    {
                    View3D view3D = View3D.CreateIsometric(doc, viewFamilyTypes.First().Id);
                    view3D.Name = catlist.ToArray()[i];
                    var direction = new XYZ(-1, 1, -1);
                    view3D.SetOrientation(new ViewOrientation3D(direction, new XYZ(0, 1, 1), new XYZ(0, 1, -1)));

                    //ElementClassFilter filter = new ElementClassFilter(typeof(FamilyInstance));
                   
                    //FilteredElementCollector collector = new FilteredElementCollector(doc, view3D.Id);
                    //collector.WherePasses(filter);

                    FilteredElementCollector collector = new FilteredElementCollector(doc, view3D.Id)/*.WhereElementIsElementType()*/;
                    //FilteredElementCollector collector= new FilteredElementCollector(doc).WhereElementIsNotElementType();

                    var hideIds = new List<ElementId>();

                    foreach (var e in collector)
                    {
                        if (e.Category != null)
                        {
                            Category cat = e.Category;
                            if (cat.Name != view3D.Name && cat.Name != "Cameras")
                            {
                                if (cat.CategoryType == CategoryType.Model)
                                {
                                    try
                                    {
                                        view3D.HideCategoryTemporary(cat.Id);
                                    }
                                    catch (Exception)
                                    {

                                        
                                    }
                                    //hideIds.Add(e.Id);
                                }
                            }

                        }
                    }
                    

                }
                ttNew.Commit();
            }

            FBXExportOptions FBXOP = new FBXExportOptions();

            List<View3D> all3DViews = new FilteredElementCollector(doc).OfClass(typeof(View3D)).Cast<View3D>().ToList();
            for (int i = 0; i < all3DViews.ToArray().Length; i++)
            {
                ViewSet viewSet = new ViewSet();
                viewSet.Insert(all3DViews.ToArray()[i]);
                doc.Export(folderPath, "nameFBXExportedFile" + i.ToString(), viewSet, new FBXExportOptions());
            }

            //foreach (var _3DVIEW in all3DViews)
            //{
            //    ViewSet viewSet = new ViewSet();
            //    viewSet.Insert(_3DVIEW); 
            //    doc.Export(folderPath, all3DViews.Count "nameFBXExportedFile", viewSet, new FBXExportOptions());
            //}
            

            return Autodesk.Revit.UI.Result.Succeeded;
        }
    }

    class ribbonUI : IExternalApplication
    {

        public Autodesk.Revit.UI.Result OnStartup(UIControlledApplication application)
        {

            string appdataFolder = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
            string folderPath = Path.Combine(appdataFolder, @"Autodesk\Revit\Addins\2019\AJC_Commands\img");
            string dll = Assembly.GetExecutingAssembly().Location;

            string myRibbon_1 = " Alex Custom Tools ";
            application.CreateRibbonTab(myRibbon_1);



            RibbonPanel panel_1_a = application.CreateRibbonPanel(myRibbon_1, "Views/Sheets Tools");
            RibbonPanel panel_2_a = application.CreateRibbonPanel(myRibbon_1, "Modify");
            RibbonPanel panel_3_a = application.CreateRibbonPanel(myRibbon_1, "Analisys");
            //RibbonPanel panel_4_a = application.CreateRibbonPanel(myRibbon_1, "Legend views");
            RibbonPanel panel_5_a = application.CreateRibbonPanel(myRibbon_1, "Version");
            RibbonPanel panel_6_a = application.CreateRibbonPanel(myRibbon_1, "Rhino");
            RibbonPanel panel_7_a = application.CreateRibbonPanel(myRibbon_1, "Lines/planes");

            //-----------------------------------Views/Sheets Tools tab-----------------------------------------------------

            PushButtonData b1 = new PushButtonData("ButtonNameA", "Create Sheet/s views", dll, "BoostYourBIM.PlaceView_CrateSheet");
            b1.LargeImage = new BitmapImage(new Uri(Path.Combine(folderPath, "createsheet_32.png"), UriKind.Absolute));
            PushButtonData b2 = new PushButtonData("Create Multiple Sheets", "Create Multiple Empty Sheets", dll, "BoostYourBIM.CreatMultipleSheet");
            b2.LargeImage = new BitmapImage(new Uri(Path.Combine(folderPath, "createsheet_32.png"), UriKind.Absolute));
            PushButtonData b3 = new PushButtonData("Duplicate one Sheet", "Duplicate one Sheet", dll, "BoostYourBIM.Duplicate_0ne_sheet");
            b3.LargeImage = new BitmapImage(new Uri(Path.Combine(folderPath, "storeselect.png"), UriKind.Absolute));
            PushButtonData b4 = new PushButtonData("Copy Schedule", "Schedule/legend placement", dll, "BoostYourBIM.copy_schedule");
            b4.LargeImage = new BitmapImage(new Uri(Path.Combine(folderPath, "Form.png"), UriKind.Absolute));


           
            //b8.ToolTip = " Creates multiple views from a room ";
            //b8.LongDescription = "Enter (TO BE SCHEDULE) in (Comments) field to create views from rooms";

            SplitButtonData sb1 = new SplitButtonData("Sheet creating option", "Options to create Sheets sets");
            SplitButton sb = panel_1_a.AddItem(sb1) as SplitButton;
            sb.IsSynchronizedWithCurrentItem = false;
            sb.ItemText = "hola";

            sb.AddPushButton(b1);
            sb.AddPushButton(b2);
            sb.AddPushButton(b3);
            sb.AddPushButton(b4);

            PushButton a_1 = (PushButton)panel_1_a.AddItem(new PushButtonData("Delete All Views", "Delete All Views", dll, "BoostYourBIM.DeleteAllViews"));
            a_1.LargeImage = new BitmapImage(new Uri(Path.Combine(folderPath, "Erase.png"), UriKind.Absolute));
            a_1.ToolTip = "This tool will delete all views and sheets in the project living only one (the view named Home)";
            a_1.LongDescription = "...";

            //-----------------------------------Modify tab-----------------------------------------------------

            PushButton a_2_1 = (PushButton)panel_2_a.AddItem(new PushButtonData("ReNumbering", "ReNumbering", dll, "BoostYourBIM.ReNumbering"));
            a_2_1.LargeImage = new BitmapImage(new Uri(Path.Combine(folderPath, "rename.png"), UriKind.Absolute));
            a_2_1.ToolTip = "Renumbers a secuence of revit elements (Viewports, Doors/room number, Grids) giving that the parameter does not contain a text character";
            a_2_1.LongDescription = "...";

            PushButton a_2 = (PushButton)panel_2_a.AddItem(new PushButtonData("Delete Level ", "Delete Level ", dll, "BoostYourBIM.DeleteLevel"));
            a_2.LargeImage = new BitmapImage(new Uri(Path.Combine(folderPath, "deletelevel_32.png"), UriKind.Absolute));
            a_2.ToolTip = "Reallocated hosted element from one level to another so elements are not lose when level is deleted";

            PushButton a_3 = (PushButton)panel_2_a.AddItem(new PushButtonData("Remove paint", "Remove paint", dll, "BoostYourBIM.remove_paint"));
            a_3.LargeImage = new BitmapImage(new Uri(Path.Combine(folderPath, "remove_paint.png"), UriKind.Absolute));
            a_3.ToolTip = "Removes paint from selected set of walls";

            PushButton a_12 = (PushButton)panel_2_a.AddItem(new PushButtonData("Text to Uppercase", "Text to Uppercase", dll, "BoostYourBIM.text_upper"));
            a_12.LargeImage = new BitmapImage(new Uri(Path.Combine(folderPath, "upper_.png"), UriKind.Absolute));
            a_12.ToolTip = "";

            PushButton a_3_1 = (PushButton)panel_3_a.AddItem(new PushButtonData("Wall Angle", "Wall Angle", dll, "BoostYourBIM.Wall_Angle_to"));
            a_3_1.LargeImage = new BitmapImage(new Uri(Path.Combine(folderPath, "iconfinder_Angle_131818.png"), UriKind.Absolute));
            a_3_1.ToolTip = "...";

            PushButton a_8 = (PushButton)panel_2_a.AddItem(new PushButtonData("Detail Line select", "Detail Line select", dll, "BoostYourBIM.select_detailline"));
            a_8.LargeImage = new BitmapImage(new Uri(Path.Combine(folderPath, "Selection.png"), UriKind.Absolute));
            a_8.ToolTip = "After selecting a Detail line the tool will select all intances of the type in the view or the project";

            PushButton a_13 = (PushButton)panel_2_a.AddItem(new PushButtonData("Total Lenght", "Total Lenght", dll, "BoostYourBIM.TotalLenght"));
            a_13.LargeImage = new BitmapImage(new Uri(Path.Combine(folderPath, "TotalLenght.png"), UriKind.Absolute));
            a_13.ToolTip = "Retrives the total lenght of line base elements ";


            PushButton a_20 = (PushButton)panel_2_a.AddItem(new PushButtonData("Solid_rooms", "Solid_rooms", dll, "BoostYourBIM.Create_solid_rooms"));
            a_20.LargeImage = new BitmapImage(new Uri(Path.Combine(folderPath, "solid icon.png"), UriKind.Absolute));
            a_20.ToolTip = "";

            PushButton a_21 = (PushButton)panel_2_a.AddItem(new PushButtonData("Isolate", "Isolate", dll, "BoostYourBIM.isolate"));
            a_21.LargeImage = new BitmapImage(new Uri(Path.Combine(folderPath, "solid icon.png"), UriKind.Absolute));
            a_21.ToolTip = "";

            PushButton a_22 = (PushButton)panel_2_a.AddItem(new PushButtonData("Isolate category", "Isolate category", dll, "BoostYourBIM.isolate_category"));
            a_22.LargeImage = new BitmapImage(new Uri(Path.Combine(folderPath, "solid icon.png"), UriKind.Absolute));
            a_22.ToolTip = "";

            PushButton a_23 = (PushButton)panel_2_a.AddItem(new PushButtonData("Clean view", "Clean view", dll, "BoostYourBIM.Clean_view"));
            a_23.LargeImage = new BitmapImage(new Uri(Path.Combine(folderPath, "solid icon.png"), UriKind.Absolute));
            a_23.ToolTip = "";

            PushButton a_24 = (PushButton)panel_2_a.AddItem(new PushButtonData("View Range by bbox", "View Range by bbox", dll, "BoostYourBIM.View_range_by_bbox"));
            a_24.LargeImage = new BitmapImage(new Uri(Path.Combine(folderPath, "solid icon.png"), UriKind.Absolute));
            a_24.ToolTip = "";

            PushButtonData D_1 = new PushButtonData("Line", "Line", dll, "BoostYourBIM.make_line");
            D_1.LargeImage = new BitmapImage(new Uri(Path.Combine(folderPath, "line.png"), UriKind.Absolute));

            PushButton a_16 = (PushButton)panel_3_a.AddItem(new PushButtonData("Select NB wall", "non bounding wall", dll, "BoostYourBIM.Wall_Bounding_room"));
            a_16.LargeImage = new BitmapImage(new Uri(Path.Combine(folderPath, "nonroombounding.png"), UriKind.Absolute));
            a_16.ToolTip = "...";






            PushButtonData D_2 = new PushButtonData("Plane", "Plane", dll, "BoostYourBIM.line_point_plane");
            D_2.LargeImage = new BitmapImage(new Uri(Path.Combine(folderPath, "line.png"), UriKind.Absolute));
            PushButtonData D_3 = new PushButtonData("Line distance", "Line from surface", dll, "BoostYourBIM.make_line_from_surface_normal");
            D_3.LargeImage = new BitmapImage(new Uri(Path.Combine(folderPath, "line.png"), UriKind.Absolute));
            PushButtonData D_4 = new PushButtonData("Pipe from line", "Pipe from line", dll, "BoostYourBIM.make_pipe_by_line");
            D_4.LargeImage = new BitmapImage(new Uri(Path.Combine(folderPath, "line.png"), UriKind.Absolute));
            PushButtonData D_5 = new PushButtonData("Loft geometry", "Loft geometry", dll, "BoostYourBIM.loft");
            D_5.LargeImage = new BitmapImage(new Uri(Path.Combine(folderPath, "line.png"), UriKind.Absolute));
            PushButtonData D_6 = new PushButtonData("Duct from line", "Duct geometry", dll, "BoostYourBIM.make_duck_by_line");
            D_6.LargeImage = new BitmapImage(new Uri(Path.Combine(folderPath, "line.png"), UriKind.Absolute));
            PushButtonData D_7 = new PushButtonData("Create flex ducts", "Create flex ducts", dll, "BoostYourBIM.Create_flex_ducts_from_line");
            D_7.LargeImage = new BitmapImage(new Uri(Path.Combine(folderPath, "line.png"), UriKind.Absolute));
            PushButtonData D_8 = new PushButtonData("Create ducts fitting", "Create ducts fitting", dll, "BoostYourBIM.duct_elbo");
            D_8.LargeImage = new BitmapImage(new Uri(Path.Combine(folderPath, "line.png"), UriKind.Absolute));
            PushButtonData D_9 = new PushButtonData("Closest_point", "Closest_point", dll, "BoostYourBIM.Closest_point_2Lines");
            D_9.LargeImage = new BitmapImage(new Uri(Path.Combine(folderPath, "line.png"), UriKind.Absolute));
            //PushButtonData D_9 = new PushButtonData("Create tag", "Create tag", dll, "BoostYourBIM.tag_element");
            //D_9.LargeImage = new BitmapImage(new Uri(Path.Combine(folderPath, "Benmax.JPG"), UriKind.Absolute));


            SplitButtonData sb4 = new SplitButtonData("Sheet creating option", "Options to create Sheets sets");
            SplitButton sb_4 = panel_7_a.AddItem(sb4) as SplitButton;
            sb_4.IsSynchronizedWithCurrentItem = false;
            sb_4.ItemText = "hola";
            sb_4.IsSynchronizedWithCurrentItem = false;
            sb_4.AddPushButton(D_1);
            sb_4.AddPushButton(D_2);
            sb_4.AddPushButton(D_3);
            sb_4.AddPushButton(D_4);
            sb_4.AddPushButton(D_5);
            sb_4.AddPushButton(D_6);
            sb_4.AddPushButton(D_7);
            sb_4.AddPushButton(D_8);
            sb_4.AddPushButton(D_9);





            PushButton a_14 = (PushButton)panel_3_a.AddItem(new PushButtonData("DWG location", "DWG location", dll, "BoostYourBIM.Find_dwg"));
            a_14.LargeImage = new BitmapImage(new Uri(Path.Combine(folderPath, "finddwg.png"), UriKind.Absolute));

            //PushButton a_20 = (PushButton)panel_3_a.AddItem(new PushButtonData("View by Cat", "View by Cat", dll, "BoostYourBIM.View_by_Cat"));
            //a_20.LargeImage = new BitmapImage(new Uri(Path.Combine(folderPath, "finddwg.png"), UriKind.Absolute));

            a_14.ToolTip = "Finds the view that host a particular DWG";
            PushButton a_4 = (PushButton)panel_3_a.AddItem(new PushButtonData("Suneye View", "Suneye View", dll, "BoostYourBIM.Create_Sun_Eye_view"));
            a_4.LargeImage = new BitmapImage(new Uri(Path.Combine(folderPath, "sun_eyes_view_2.png"), UriKind.Absolute));
            a_4.ToolTip = "Create a isometric Revit view from sun position towards project origin";
            a_4.LongDescription = "...";

            PushButton a_19 = (PushButton)panel_3_a.AddItem(new PushButtonData("Suneye View set", "Suneye View set", dll, "BoostYourBIM.Create_Sun_Eye_study"));
            a_19.LargeImage = new BitmapImage(new Uri(Path.Combine(folderPath, "sun_eyes_view_2.png"), UriKind.Absolute));
            a_19.ToolTip = "Create a isometric Revit view from sun position towards project origin";
            a_19.LongDescription = "...";

            PushButton a_7 = (PushButton)panel_5_a.AddItem(new PushButtonData("Version", "Version", dll, "BoostYourBIM.info"));
            a_7.LargeImage = new BitmapImage(new Uri(Path.Combine(folderPath, "Help.bmp"), UriKind.Absolute));
            a_7.ToolTip = "...";


            PushButton a_18 = (PushButton)panel_3_a.AddItem(new PushButtonData("Floor from topo", "Floor from topo", dll, "BoostYourBIM.revision_on_project"));
            a_18.LargeImage = new BitmapImage(new Uri(Path.Combine(folderPath, "topo.png"), UriKind.Absolute));
            a_18.ToolTip = "...";




            PushButtonData C_11 = new PushButtonData("Family Geo Search", "Family Geo Search", dll, "BoostYourBIM.Rhino_access_faces");
            C_11.LargeImage = new BitmapImage(new Uri(Path.Combine(folderPath, "rhinoexport_32.png"), UriKind.Absolute));

            PushButtonData C_12 = new PushButtonData("Rhino to Revit", "Rhino to Revit", dll, "BoostYourBIM.Reading_from_rhino");
            C_12.LargeImage = new BitmapImage(new Uri(Path.Combine(folderPath, "rhinoexport_32_copy.png"), UriKind.Absolute));

            PushButtonData C_13 = new PushButtonData("Family Geo Exporter", "Rhino Geo Exporter", dll, "BoostYourBIM.Rhino_access");
            C_13.LargeImage = new BitmapImage(new Uri(Path.Combine(folderPath, "rhinoexport_32.png"), UriKind.Absolute));

            PushButtonData C_14 = new PushButtonData("Rhino pnts to Revit topo", "Rhino pnts to Revit topo", dll, "BoostYourBIM.rhino_points_to_revit_topo");
            C_14.LargeImage = new BitmapImage(new Uri(Path.Combine(folderPath, "rhinoexport_32_copy.png"), UriKind.Absolute));

            PushButtonData C_15 = new PushButtonData("Rhino lns to Revit lines", "Rhino lns to Revit lines", dll, "BoostYourBIM.rhino_lns_to_revit_lns");
            C_15.LargeImage = new BitmapImage(new Uri(Path.Combine(folderPath, "rhinoexport_32.png"), UriKind.Absolute));

            PushButtonData C_16 = new PushButtonData("Rhino lns to Frame", "Rhino lns to Frame", dll, "BoostYourBIM.rhino_lns_to_revit_structure");
            C_16.LargeImage = new BitmapImage(new Uri(Path.Combine(folderPath, "rhinoexport_32.png"), UriKind.Absolute));

            PushButtonData C_17 = new PushButtonData("Create_solid_rooms", "Create_solid_rooms", dll, "BoostYourBIM.Create_solid_rooms");
            C_17.LargeImage = new BitmapImage(new Uri(Path.Combine(folderPath, "rhinoexport_32.png"), UriKind.Absolute));

            SplitButtonData sb3 = new SplitButtonData("Sheet creating option", "Options to create Sheets sets");
            SplitButton sb_3 = panel_6_a.AddItem(sb3) as SplitButton;
            sb_3.IsSynchronizedWithCurrentItem = false;
            sb_3.ItemText = "hola";
            sb_3.IsSynchronizedWithCurrentItem = false;
            sb_3.AddPushButton(C_11);
            sb_3.AddPushButton(C_12);
            sb_3.AddPushButton(C_13);
            sb_3.AddPushButton(C_14);
            sb_3.AddPushButton(C_15);
            sb_3.AddPushButton(C_16);
            sb_3.AddPushButton(C_17);


            PushButton a_9 = (PushButton)panel_5_a.AddItem(new PushButtonData("unlock", "unlock", dll, "BoostYourBIM.unlock"));
            a_9.LargeImage = new BitmapImage(new Uri(Path.Combine(folderPath, "lock.png"), UriKind.Absolute));
            a_9.ToolTip = "...";
            PushButton a_10 = (PushButton)panel_5_a.AddItem(new PushButtonData("lock", "lock", dll, "BoostYourBIM._lock"));
            a_10.LargeImage = new BitmapImage(new Uri(Path.Combine(folderPath, "unlock.png"), UriKind.Absolute));
            a_10.ToolTip = "...";

            PushButtonData b_8 = new PushButtonData("Door Views", "Door Views", dll, "BoostYourBIM.Door_Section");
            b_8.LargeImage = new BitmapImage(new Uri(Path.Combine(folderPath, "doorSchedule.png"), UriKind.Absolute));
            b_8.ToolTip = "a section view is created of elements containing an identifier in the parameter field (Schedule Identifier)";

            PushButtonData b_9 = new PushButtonData("Room Views", "Room Views", dll, "BoostYourBIM.RoomElevations");
            b_9.LargeImage = new BitmapImage(new Uri(Path.Combine(folderPath, "roomSchedule_32.png"), UriKind.Absolute));
            b_9.ToolTip = "Section, floor/ceiling plans, Int elevation and Isometric views can be created of one room containing an identifier in the parameter field (Schedule Identifier)";

            PushButtonData b_10 = new PushButtonData("Wall Views", "Wall Views", dll, "BoostYourBIM.CreateSchedule");
            b_10.LargeImage = new BitmapImage(new Uri(Path.Combine(folderPath, "cutingwall_32.png"), UriKind.Absolute));
            b_10.ToolTip = "a section view is created of elements containing an identifier in the parameter field (Schedule Identifier)";

            PushButtonData b_11 = new PushButtonData("Window Views", "Window Views", dll, "BoostYourBIM.WindowSection");
            b_11.LargeImage = new BitmapImage(new Uri(Path.Combine(folderPath, "windowSchedule_32.png"), UriKind.Absolute));
            b_11.ToolTip = "a section view is created of elements containing an identifier in the parameter field (Schedule Identifier)";




            //SplitButtonData sb5 = new SplitButtonData("Sheet creating option", "Options to create Sheets sets");
            //SplitButton sb_5 = panel_4_a.AddItem(sb3) as SplitButton;
            //sb_5.IsSynchronizedWithCurrentItem = false;
            //sb_5.ItemText = "hola";
            //sb_5.IsSynchronizedWithCurrentItem = false;
            //sb_5.AddPushButton(b_8);
            //sb_5.AddPushButton(b_9);
            //sb_5.AddPushButton(b_10);
            //sb_5.AddPushButton(b_11);



            ////PushButton b_12 = (PushButton)panel_4_a.AddItem(new PushButtonData("re111", "ewrwe11", dll, "BoostYourBIM.revision_on_project"));
            ////b_12.LargeImage = new BitmapImage(new Uri(Path.Combine(folderPath, "windowSchedule_32.png"), UriKind.Absolute));
            ////b_12.ToolTip = "";



            //PushButtonData b6 = new PushButtonData("Modify Views/Sheet", "Modify Views/Sheet", dll, "BoostYourBIM.nada");
            //b6.LargeImage = new BitmapImage(new Uri(Path.Combine(folderPath, "createsheet_32.png"), UriKind.Absolute));


            //PushButtonData b8 = new PushButtonData("Read from excel", "Read from excel", dll, "BoostYourBIM.room_data"); 
            //b8.LargeImage = new BitmapImage(new Uri(Path.Combine(folderPath, "createsheet_32.png"), UriKind.Absolute));
            //PushButtonData b9 = new PushButtonData("wall data to TXT", "wall data to TXT", dll, "BoostYourBIM.wall_data");
            //b9.LargeImage = new BitmapImage(new Uri(Path.Combine(folderPath, "createsheet_32.png"), UriKind.Absolute));
            //PushButtonData b10 = new PushButtonData("floor data to TXT", "floor data to TXT", dll, "BoostYourBIM.floor_data");
            //b10.LargeImage = new BitmapImage(new Uri(Path.Combine(folderPath, "createsheet_32.png"), UriKind.Absolute));
            //PushButtonData b11 = new PushButtonData("Roor data to TXT", "Roor data to TXT", dll, "BoostYourBIM.roof_data");
            //b11.LargeImage = new BitmapImage(new Uri(Path.Combine(folderPath, "createsheet_32.png"), UriKind.Absolute));
            //PushButtonData b12 = new PushButtonData("Ducts data to TXT", "Ducts data to TXT", dll, "BoostYourBIM.duct_data");
            //b12.LargeImage = new BitmapImage(new Uri(Path.Combine(folderPath, "createsheet_32.png"), UriKind.Absolute));
            //PushButtonData b13 = new PushButtonData("Curtain data to TXT", "Curtain data to TXT", dll, "BoostYourBIM.Curtain_wall_data");
            //b13.LargeImage = new BitmapImage(new Uri(Path.Combine(folderPath, "createsheet_32.png"), UriKind.Absolute));
            //PushButtonData b14 = new PushButtonData("Stair data to TXT", "Stair data to TXT", dll, "BoostYourBIM.Stair_data");
            //b14.LargeImage = new BitmapImage(new Uri(Path.Combine(folderPath, "createsheet_32.png"), UriKind.Absolute));
            //PushButtonData b15 = new PushButtonData("Section data to TXT", "Section data to TXT", dll, "BoostYourBIM.Modelline_data");
            //b15.LargeImage = new BitmapImage(new Uri(Path.Combine(folderPath, "createsheet_32.png"), UriKind.Absolute));



            //PushButtonData b5 = new PushButtonData("Delete All Views", "Delete All Views", dll, "BoostYourBIM.DeleteAllViews");
            //b5.LargeImage = new BitmapImage(new Uri(Path.Combine(folderPath, "Erase.png"), UriKind.Absolute));
            //SplitButtonData sb2 = new SplitButtonData("holaa", "hola");
            //SplitButton sb_2 = panel_1_a.AddItem(sb2) as SplitButton;
            //sb_2.ItemText = "asdsa";
            //sb_2.IsSynchronizedWithCurrentItem = false;
            //sb_2.AddPushButton(b5);

            //sb.AddPushButton(b6);
            //sb.AddPushButton(b5);
            //sb.AddPushButton(b7);
            //sb.AddPushButton(b8);
            //sb.AddPushButton(b9);
            //sb.AddPushButton(b10);
            //sb.AddPushButton(b11);
            //sb.AddPushButton(b12);
            //sb.AddPushButton(b13);
            //sb.AddPushButton(b14);
            //sb.AddPushButton(b15);


            //PushButtonData b5 = new PushButtonData("Delete all Sheets", "Delete all Sheets", dll, "BoostYourBIM.DeleteAllSheets");
            //b5.LargeImage = new BitmapImage(new Uri(Path.Combine(folderPath, "Erase.png"), UriKind.Absolute));

            //adWin.ComponentManager.UIElementActivated += new EventHandler<adWin.UIElementActivatedEventArgs>(ComponentManager_UIElementActivated);

            try
            {
                foreach (Autodesk.Windows.RibbonTab tab in Autodesk.Windows.ComponentManager.Ribbon.Tabs)
                {
                    if (tab.Title == "Insert")
                    {
                        tab.IsVisible = false;
                    }
                }

                adWin.RibbonControl ribbon = adWin.ComponentManager.Ribbon;

                //ImageSource imgbg = new BitmapImage(new Uri(Path.Combine(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location),
                //    "gradient2.png"), UriKind.Relative));

                //// define an image brush

                //ImageBrush picBrush = new ImageBrush();
                //picBrush.ImageSource = imgbg;
                //picBrush.AlignmentX = AlignmentX.Left;
                //picBrush.AlignmentY = AlignmentY.Top;
                //picBrush.Stretch = Stretch.None;
                //picBrush.TileMode = TileMode.FlipXY;

                //define a linear brush from top to bottom

                LinearGradientBrush gradientBrush = new LinearGradientBrush();

                gradientBrush.StartPoint = new System.Windows.Point(0, 0);

                gradientBrush.EndPoint = new System.Windows.Point(0, 1);

                gradientBrush.GradientStops.Add(new GradientStop(Colors.White, 0.0));

                gradientBrush.GradientStops.Add(new GradientStop(Colors.Orange, 0.95));

                // change the tab header font

                //ribbon.FontFamily = new System.Windows.Media.FontFamily( "Bauhaus 93");
                ribbon.Opacity = 70;
                ribbon.FontSize = 10;
                //ribbon.Background = picBrush;
                //iterate through the tabs and their panels

                foreach (adWin.RibbonTab tab in ribbon.Tabs)
                {
                    
                    string name = tab.AutomationName;

                    if (name == "Insert")
                    {
                        foreach (adWin.RibbonPanel panel in tab.Panels)
                        {
                            
                            adWin.RibbonItemCollection items = panel.Source.Items;
                            foreach (var item in items)
                            {
                                string name_ = item.Id;
                                if (name_ == "ID_FILE_IMPORT")
                                {
                                    item.IsEnabled = false;
                                }
                            }
                        }
                    }
                    foreach (adWin.RibbonPanel panel in tab.Panels)
                    {
                        string name1 = panel.AutomationName;
                        panel.CustomPanelTitleBarBackground = gradientBrush;
                        /*panel.CustomPanelBackground = picBrush;*/ // use your picture
                    }
                }
                RibbonItemEventArgs jk = new RibbonItemEventArgs();
                List<RibbonPanel> ribbons = jk.Application.GetRibbonPanels();

            }
            catch (Exception ex)
            {
                winform.MessageBox.Show(
                  ex.StackTrace + "\r\n" + ex.InnerException,
                  "Error", winform.MessageBoxButtons.OK);

                return Result.Failed;
            }

            return Autodesk.Revit.UI.Result.Succeeded;
        }

        public Autodesk.Revit.UI.Result OnShutdown(UIControlledApplication application)
        {
            return Autodesk.Revit.UI.Result.Succeeded;
        }

    }
}

