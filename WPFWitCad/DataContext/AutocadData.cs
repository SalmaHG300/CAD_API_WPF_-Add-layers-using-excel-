using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Runtime;
using System;
using System.Collections.Generic;
using WPFWitCad.Model;
using WPFWitCad.View;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.Colors;
using Excel = Microsoft.Office.Interop.Excel;
using Application = Autodesk.AutoCAD.ApplicationServices.Application;
using Microsoft.Office.Interop.Excel;
using Line = Autodesk.AutoCAD.DatabaseServices.Line;
using Arc = Autodesk.AutoCAD.DatabaseServices.Arc;
using System.Windows.Forms;

namespace WPFWitCad.DataContext
{
    public class AutocadData
    {
        public static List<CadLayerObj> GetCadLayers()
        {
            // excel path defined 
            //string excelFilePath = "D:\\iti\\api cad\\excel sheet\\ExcelCADLayers.xlsx";

            //excel file choosen  by user 
            string excelFilePath = GetExcelFilePath();
            if (string.IsNullOrEmpty(excelFilePath))
            {
                // User canceled the selection or an error occurred
                Application.ShowAlertDialog("Excel file selection canceled or failed.");
                return new List<CadLayerObj>();
            }

            // AutoCAD document
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            Editor ed = doc.Editor;

            // Open Excel application
            _Application excelApp = new Excel.Application();
            Workbook workbook = excelApp.Workbooks.Open(excelFilePath);
            Worksheet worksheet = (Worksheet)workbook.Sheets[1];
            Range range = worksheet.UsedRange;

            // Loop through Excel data and create AutoCAD layers
            for (int row = 2; row <= range.Rows.Count; row++)
            {
                string layerName = Convert.ToString((range.Cells[row, 1] as Range).Value);
                int colorIndex = Convert.ToInt32((range.Cells[row, 2] as Range).Value);
                string linetype = Convert.ToString((range.Cells[row, 3] as Range).Value);

                // Check if the layer name is empty
                if (string.IsNullOrEmpty(layerName))
                {
                    break; // Exit the loop if the layer name is empty
                }

                else
                {
                    // Create AutoCAD layer
                    using (Transaction tr = db.TransactionManager.StartTransaction())
                    {
                        try
                        {
                            LayerTable layerTable = (LayerTable)tr.GetObject(db.LayerTableId, OpenMode.ForWrite);

                            if (!layerTable.Has(layerName))
                            {
                                LayerTableRecord layer = new LayerTableRecord();
                                layer.Name = layerName;
                                layer.Color = Color.FromColorIndex(ColorMethod.ByAci, (short)colorIndex);
                                // Get the line type ID
                                ObjectId linetypeId = GetLinetypeId(db, linetype, tr);
                                if (linetypeId.IsValid)
                                {
                                    layer.LinetypeObjectId = linetypeId;
                                }
                                else
                                {
                                    Application.ShowAlertDialog("Invalid line type: " + linetype);
                                    continue; // Skip this layer and move to the next one
                                }
                                layerTable.Add(layer);
                                tr.AddNewlyCreatedDBObject(layer, true);
                            }
                            tr.Commit();
                        }

                        catch (System.Exception ex)
                        {
                            Application.ShowAlertDialog("Error importing layers: " + ex.Message);
                            tr.Abort();
                        }
                    }
                }
            }
            Application.ShowAlertDialog("Layers imported successfully.");

            // Release Excel resources
            workbook.Close();
            excelApp.Quit();

            List<CadLayerObj> CadlayerObjs = new List<CadLayerObj>();
            using (Transaction Trans = db.TransactionManager.StartTransaction())
            {
                try
                {
                    LayerTable Lts = Trans.GetObject(db.LayerTableId, OpenMode.ForWrite) as LayerTable;

                    foreach (var objectId in Lts)
                    {
                        LayerTableRecord LtRecord = Trans.GetObject(objectId, OpenMode.ForRead) as LayerTableRecord;

                        CadLayerObj cadLayerObj = new CadLayerObj();

                        cadLayerObj.Name = LtRecord.Name;

                        cadLayerObj.Color = LtRecord.Color;

                        CadlayerObjs.Add(cadLayerObj);
                    }

                    Trans.Commit();
                }
                catch (System.Exception Ex)
                {

                    ed.WriteMessage(Ex.Message);

                    Trans.Abort();
                }
            }
            return CadlayerObjs;
        }


        private static ObjectId GetLinetypeId(Database db, string linetypeName, Transaction tr)
        {
            LinetypeTable linetypeTable = db.LinetypeTableId.GetObject(OpenMode.ForWrite) as LinetypeTable;

            if (linetypeTable.Has(linetypeName))
            {
                return linetypeTable[linetypeName];
            }
            else
            {
                using (tr = db.TransactionManager.StartTransaction())
                {
                    // Open the Linetype table for read
                    LinetypeTable acLineTypTbl = tr.GetObject(db.LinetypeTableId, OpenMode.ForRead) as LinetypeTable;

                    if (acLineTypTbl.Has(linetypeName) == false)
                    {
                        // Load the Center Linetype
                        db.LoadLineTypeFile(linetypeName, "acad.lin");
                    }
                    // Save the changes and dispose of the transaction
                    tr.Commit();
                    return acLineTypTbl[linetypeName];
                }
            }
        }

        [CommandMethod("ITIPluign")]

        public void OpenWindow()
        {
            MainWindow mainWindow = new MainWindow();

            Application.ShowModalWindow(mainWindow);
        }
        public static void Getkeywords(String LayerName)
        {
            Document Doc = Application.DocumentManager.MdiActiveDocument;
            Database Db = Doc.Database;
            Editor ed = Doc.Editor;
            PromptKeywordOptions pke = new PromptKeywordOptions("choose an element");
            pke.Keywords.Add("circle");
            pke.Keywords.Add("Line");
            pke.Keywords.Add("ARC");
            pke.Keywords.Add("PolyLine");
            var Resultpke = ed.GetKeywords(pke);
            switch (Resultpke.StringResult)
            {
                case "circle":
                    CreateCircle(Db, ed, LayerName);
                    break;
                case "Line":
                    CreateLine(Db, ed, LayerName);
                    break;
                case "ARC":
                    CreateARC(Db, ed, LayerName);
                    break;
                case "PolyLine":
                    CreatePolyline(Db, ed, LayerName);
                    break;
                default:
                    ed.WriteMessage("Invalid option selected.");
                    break;
            }
        }

        ///////Circle Creation////////  
        public static void CreateCircle(Database Db, Editor Ed, String LayerName)
        {
            PromptDoubleOptions pdbtOpts = new PromptDoubleOptions("Enter Circlr radius");
            PromptDoubleResult promptResult = Ed.GetDouble(pdbtOpts);
            double radius = promptResult.Value;
            PromptPointOptions p1 = new PromptPointOptions("Enter Center of circle");
            PromptPointResult presult = Ed.GetPoint(p1);
            Point3d point3D1 = presult.Value;
            using (Transaction transaction = Db.TransactionManager.StartTransaction())
            {
                try
                {
                    BlockTableRecord btr = transaction.GetObject(Db.CurrentSpaceId, OpenMode.ForWrite) as BlockTableRecord;
                    Circle circle = new Circle();
                    circle.Radius = radius;
                    circle.Center = point3D1;
                    circle.Layer = LayerName;

                    btr.AppendEntity(circle);
                    transaction.AddNewlyCreatedDBObject(circle, true);
                    transaction.Commit();

                }
                catch (System.Exception ex)
                {
                    Ed.WriteMessage(ex.Message);
                    transaction.Abort();
                }
            }
        }

        ///////Line Creation////////  
        public static void CreateLine(Database Db, Editor Ed, String LayerName)
        {
            PromptPointOptions p1 = new PromptPointOptions("Enter first point");
            PromptPointResult presult = Ed.GetPoint(p1);
            Point3d point3D1 = presult.Value;
            p1 = new PromptPointOptions("Enter seconf point");
            presult = Ed.GetPoint(p1);
            Point3d point3D2 = presult.Value;
            using (Transaction tra = Db.TransactionManager.StartTransaction())
            {
                try
                {
                    var btr = tra.GetObject(Db.CurrentSpaceId, OpenMode.ForWrite) as BlockTableRecord;
                    Line myline = new Line(point3D1, point3D2);
                    myline.Layer = LayerName;
                    btr.AppendEntity(myline);
                    tra.AddNewlyCreatedDBObject(myline, true);
                    tra.Commit();
                }
                catch (System.Exception ex)
                {
                    Ed.WriteMessage(ex.Message);
                    tra.Abort();
                }
            }
        }

        ///////ARC Creation////////  
        public static void CreateARC(Database Db, Editor Ed, String LayerName)
        {
            PromptPointOptions p1 = new PromptPointOptions("Enter Center of ARC");
            PromptPointResult presult = Ed.GetPoint(p1);
            Point3d point3D1 = presult.Value;

            PromptDoubleOptions pdbtOpts = new PromptDoubleOptions("Enter ARC radius");
            PromptDoubleResult promptResult = Ed.GetDouble(pdbtOpts);
            double radius = promptResult.Value;

            PromptDoubleOptions pdbtOpt = new PromptDoubleOptions("Enter ARC Start Angle");
            PromptDoubleResult prmtResult = Ed.GetDouble(pdbtOpt);
            double startangle = prmtResult.Value;

            PromptDoubleOptions pdtOpt = new PromptDoubleOptions("Enter ARC End Angle");
            PromptDoubleResult pResult = Ed.GetDouble(pdtOpt);
            double endangle = pResult.Value;

            using (Transaction tra = Db.TransactionManager.StartTransaction())
            {
                try
                {
                    // Open the Block table for read
                    BlockTable BlkTbl = tra.GetObject(Db.BlockTableId, OpenMode.ForRead) as BlockTable;
                    BlockTableRecord btr = tra.GetObject(BlkTbl[BlockTableRecord.ModelSpace], OpenMode.ForWrite) as BlockTableRecord;
                    Arc arc = new Arc(point3D1, Vector3d.ZAxis, radius, startangle, endangle);
                    arc.SetDatabaseDefaults();
                    arc.Layer = LayerName;
                    btr.AppendEntity(arc);
                    tra.AddNewlyCreatedDBObject(arc, true);
                    tra.Commit();
                }
                catch (System.Exception ex)
                {
                    Ed.WriteMessage(ex.Message);
                }
            }
        }

        ///////Polyline Creation////////        
        public static void CreatePolyline(Database Db, Editor Ed, String LayerName)
        {
            PromptPointOptions p1 = new PromptPointOptions("Enter start point of Polyline");
            PromptPointResult presult = Ed.GetPoint(p1);
            Point3d startPoint = presult.Value;

            PromptIntegerOptions vertexCountOpts = new PromptIntegerOptions("Enter number of vertices for the Polyline");
            PromptIntegerResult vertexCountResult = Ed.GetInteger(vertexCountOpts);
            int vertexCount = vertexCountResult.Value;

            Polyline polyline = new Polyline();
            polyline.SetDatabaseDefaults();

            using (Transaction tra = Db.TransactionManager.StartTransaction())
            {
                try
                {
                    // Open the Block table for read
                    BlockTable BlkTbl = tra.GetObject(Db.BlockTableId, OpenMode.ForRead) as BlockTable;
                    BlockTableRecord btr = tra.GetObject(BlkTbl[BlockTableRecord.ModelSpace], OpenMode.ForWrite) as BlockTableRecord;

                    polyline.AddVertexAt(0, new Point2d(startPoint.X, startPoint.Y), 0, 0, 0);

                    for (int i = 1; i < vertexCount; i++)
                    {
                        PromptPointOptions vertexOpts = new PromptPointOptions($"Enter vertex {i + 1} of {vertexCount}");
                        PromptPointResult vertexResult = Ed.GetPoint(vertexOpts);
                        Point3d vertexPoint = vertexResult.Value;
                        polyline.AddVertexAt(i, new Point2d(vertexPoint.X, vertexPoint.Y), 0, 0, 0);
                    }
                    polyline.Layer = LayerName;
                    btr.AppendEntity(polyline);
                    tra.AddNewlyCreatedDBObject(polyline, true);
                    tra.Commit();
                }
                catch (System.Exception ex)
                {
                    Ed.WriteMessage(ex.Message);
                }
            }
        }

        private static string GetExcelFilePath()
        {
            // Create an OpenFileDialog instance
            OpenFileDialog openFileDialog = new OpenFileDialog
            {
                Title = "Select Excel File",
                Filter = "Excel Files|*.xls;*.xlsx|All Files|*.*",
                Multiselect = false
            };

            // Show the dialog and get the result
            DialogResult result = openFileDialog.ShowDialog();

            // Check if the user clicked OK
            if (result == DialogResult.OK)
            {
                return openFileDialog.FileName; // Return the selected file path
            }
            else
            {
                return null; // Return null if the user canceled the selection
            }
        }




    }
}
 