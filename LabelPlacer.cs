using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using Autodesk.Revit.DB.Architecture;
using System.Reflection;
using System.Windows.Media.Imaging;
using System.Drawing;
using System.Windows.Interop;
using System.Windows;
using System.IO;

namespace LIMA
{
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    public class LIMA : IExternalApplication
    {
        static void AddRibbonPanel(UIControlledApplication a, string i_sObject)
        {
            const string RIBBON_TAB = "LIMA";
            string RIBBON_PANEL = i_sObject + " Label Placer";

            try
            {
                a.CreateRibbonTab(RIBBON_TAB);
            }
            catch (Exception) { }

            RibbonPanel panel = null;
            List<RibbonPanel> panelS = a.GetRibbonPanels(RIBBON_TAB);


            foreach(RibbonPanel pnl in panelS)
            if (pnl.Name == RIBBON_PANEL)
                {
                    panel = pnl;
                    break;
                }

            if (panel == null)
                panel = a.CreateRibbonPanel(RIBBON_TAB, RIBBON_PANEL);

            string path = Assembly.GetExecutingAssembly().Location;

            PushButtonData data = new PushButtonData(RIBBON_TAB, RIBBON_PANEL, path, "LIMA." + i_sObject + "LabelPlacer");
            //PushButtonData data = new PushButtonData(RIBBON_TAB, RIBBON_PANEL, path, "LIMA.LabelPlacerBase");

            Bitmap bitmap = new Bitmap(Properties.Resources.LIMA_Fekete);

            data.Image = BitmapToBitmapSource(bitmap);
            data.LargeImage = BitmapToBitmapSource(bitmap);

            PushButton button = panel.AddItem(data) as PushButton;
        }

        public Result OnStartup(UIControlledApplication a)
        {
            AddRibbonPanel(a, "Pipe");
            AddRibbonPanel(a, "Duct");
            AddRibbonPanel(a, "Wire");

            return Result.Succeeded;
        }

        [System.Runtime.InteropServices.DllImport("gdi32.dll")]
        public static extern bool DeleteObject(IntPtr hObject);

        static BitmapSource BitmapToBitmapSource(Bitmap bitmap)
        {
            IntPtr hBitmap = bitmap.GetHbitmap();

            BitmapSource retval;

            try
            {
                retval = Imaging.CreateBitmapSourceFromHBitmap(
                  hBitmap, IntPtr.Zero, Int32Rect.Empty,
                  BitmapSizeOptions.FromEmptyOptions());
            }
            finally
            {
                DeleteObject(hBitmap);
            }
            return retval;
        }

        public Result OnShutdown(UIControlledApplication application)
        {
            return Result.Succeeded;
        }
    }

    [Transaction(TransactionMode.Manual)]
    public class LabelPlacerBase : IExternalCommand
    {
        public virtual string GetCat() { return "Pipes"; }
        bool isHorizontal(Element i_element)
        {
            if ((getEndPoint(i_element, 0).Z - getEndPoint(i_element, 1).Z) > 0.01)
                return false;
            return true;
        }

        XYZ midPoint(XYZ firsPoint, XYZ secondPoint)
        {
            return new XYZ(
                (firsPoint.X + secondPoint.X) / 2,
                (firsPoint.Y + secondPoint.Y) / 2,
                (firsPoint.Z + secondPoint.Z) / 2
            );
        }

        XYZ getEndPoint(Element i_element, int i_idx)
        {
            LocationCurve lc = i_element.Location as LocationCurve;

            if (lc != null)
            {
                Curve c = lc.Curve;

                return c.GetEndPoint(i_idx);
            }
            else throw new Exception();
        }

        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            var sel = commandData.Application.ActiveUIDocument.Selection;
            var doc = commandData.Application.ActiveUIDocument.Document;
            var viewId = doc.ActiveView.Id;
            Exception lastErr = null;

            using (Transaction _t = new Transaction(doc, "Place labels"))
            {
                _t.Start();

                foreach (var e in sel.GetElementIds())
                {
                    var element = doc.GetElement(e);

                    try
                    {
                        if (element.Category.Name == GetCat()
                           //&&	element.Name.Contains("Pipe") 
                           && isHorizontal(element)
                          )
                        {
                            var r = new Reference(element);
                            var tagOr = TagOrientation.Horizontal;
                            var tagMode = TagMode.TM_ADDBY_CATEGORY;

                            var xyz = midPoint(getEndPoint(element, 0), getEndPoint(element, 1));

                            IndependentTag.Create(doc, viewId, r, false, tagMode, tagOr, xyz);
                        }
                    }
                    catch (Exception ex)
                    {
                        lastErr = ex;
                        continue;
                    }
                }

                if (lastErr == null)
                    _t.Commit();
                else
                {
                    TaskDialog.Show("Last error message:", lastErr.Message);
                    //_t.RollBack();
                }
            }

            return Result.Succeeded;
        }
    }

    [Transaction(TransactionMode.Manual)]
    public class PipeLabelPlacer : LabelPlacerBase { public override string GetCat() { return "Pipes"; } }

    [Transaction(TransactionMode.Manual)]
    public class DuctLabelPlacer : LabelPlacerBase { public override string GetCat() { return "Ducts"; } }

    [Transaction(TransactionMode.Manual)]
    public class WireLabelPlacer : LabelPlacerBase { public override string GetCat() { return "Wires"; } }
}

