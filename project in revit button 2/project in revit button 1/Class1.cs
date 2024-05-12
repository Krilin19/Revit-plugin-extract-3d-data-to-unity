//using Autodesk.Revit.DB;
//using Autodesk.Revit.UI;
//using Autodesk.Revit.UI.Events;
//using Autodesk.Revit.UI.Selection;
//using System.Diagnostics;
//using System;
//using System.Collections;
//using System.Collections.Generic;
//using System.IO;
//using System.Linq;
//using System.Text;
//using System.Threading.Tasks;
//using System.Windows.Media.Imaging;
//using Autodesk.Revit.ApplicationServices;
//using Autodesk.Revit.Attributes;
//using System.Windows.Forms;
//using AJC_Commands_1;
//using DBLibrary;
//using System.Windows.Controls.Primitives;
//using Autodesk.Revit.DB.Analysis;
//using Autodesk.Revit.DB.Architecture;
//using System.Reflection;
//using RvtApplication = Autodesk.Revit.ApplicationServices.Application;
//using RvtDocument = Autodesk.Revit.DB.Document;
//using OfficeOpenXml;
//using Rhino5x64;
//using RhinoScript4;
//using Rhino.FileIO;
//using Rhino.Geometry;
//using Rhino;
//using Rhino.Input;
//using RhinoCon;
//using winform = System.Windows.Forms;
//using System.Windows.Media;

//namespace AJC_Commands_1
//{
//    class Class1
//    {
        
//        var launchRhinoCommand = @"C:\Program Files\Rhinoceros 5 (64-bit)\System\Rhino.exe";
//        Process.Start($"{launchRhinoCommand}");
//        RhinoApplication m_RhinoCOM = new RhinoApplication("Rhino5x64.Application", true);
//        //File3dm m_model = new File3dm();
//        Rhino.FileIO.File3dm file1 = m_RhinoCOM.RhinoApp as Rhino.FileIO.File3dm;
//        Point3d p1 = new Point3d(0, 0, 0);
//        Rhino.Geometry.Point3d pt3d = new Point3d(10, 10, 0);
//        System.Drawing.Color rhino_col = Create_RhinoColor(10, 10, 10);
//        Rhino.DocObjects.Layer layer_ = Create_RhinoLayer("hola", rhino_col);
//        Rhino.Geometry.Line lineee = new Rhino.Geometry.Line(p1, pt3d);
//        Object line__ = Create_RhinoObject(lineee, "hola", rhino_col, layer_);
//        objs.Add(line__);
//        IList<Rhino.DocObjects.Layer> LAYERS = file1.AllLayers;
//        foreach (var item in LAYERS) 
//        {
//            TaskDialog.Show("model", "no model");
//        }
//        file1.Objects.AddLine(lineee);
///*string filename3 = LoadFile("REVIT.3dm");*/ // We're loading the 3dm file in Rhino
//                                              Rhino.FileIO.File3dm model = Rhino.FileIO.File3dm.Read(filename3);
//        SaveRhino3dmModel.Save_Rhino3dmModel(filename3, true, objs.ToArray(), "millimeters");
//                                              Rhino.FileIO.File3dm model = Rhino.FileIO.File3dm.Read(@"‪D:\lopez\Desktop\REVIT");
//                                              if (model == null)
//                                              {
//                                                  TaskDialog.Show("model", "no model");
//                                              }
//IList<Rhino.DocObjects.Layer> LAYERS = model.AllLayers;

//foreach (var item in LAYERS)
//{
//    TaskDialog.Show("model", "no model");
//}
//foreach (var item in objs)
//{
//    Rhino.DocObjects.ObjectAttributes objatt = new Rhino.DocObjects.ObjectAttributes();
//objatt.Name = "holaaa";
//    objatt.ObjectColor = rhino_col;
//    objatt.ColorSource = Rhino.DocObjects.ObjectColorSource.ColorFromObject;
//}
//Rhino.Geometry.Point point_ = new Rhino.Geometry.Point(p1);
//Rhino.Geometry.Point point_2 = new Rhino.Geometry.Point(p2);
//RhinoCommand(m_RhinoCOM, "Line", true);
//RhinoCommand(m_RhinoCOM, "Save", true);
//Rhino.Geometry.Line line = new Rhino.Geometry.Line(p2, p1);
//Rhino.RhinoApp.RunScript("_-Line 0,0,0 10,10,10", false);
//m_RhinoCOM.RhinoApp.Visible = 1;
//void SetTopOffset(Wall wall, double dOffsetInches)
//{
//    // convert user-defined offset value to feet from inches prior to setting
//    double dOffsetFeet = UnitUtils.Convert(dOffsetInches,
//                                            DisplayUnitType.DUT_DECIMAL_INCHES,
//                                            DisplayUnitType.DUT_DECIMAL_FEET);
//    Parameter paramTopOffset = wall.get_Parameter(BuiltInParameter.WALL_TOP_OFFSET);
//    paramTopOffset.Set(dOffsetFeet);
//}




