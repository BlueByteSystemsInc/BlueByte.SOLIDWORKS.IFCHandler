using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;

namespace BlueByte.SOLIDWORKS.IFCHandler.Sandbox
{
    public class SolidworksInterface : IDisposable
    {
        private bool isConnected;
        private SldWorks swApp;
        public bool IsConnected { get { return isConnected; } }
        public SolidworksInterface(SldWorks solidworksApp = null)
        {
            swApp = solidworksApp;
            if (swApp == null)
            {
                isConnected = ConnectToSW();
                //Console.WriteLine("SW Interface Loaded - Connected: " + isConnected);
            }
            else
            {
                isConnected = true;
            }
        }
        private bool ConnectToSW()
        {
            bool returnVal = false;
            try
            {
                swApp = (SldWorks)Marshal.GetActiveObject("SldWorks.Application");
                Process[] proc = Process.GetProcessesByName("SLDWORKS");
                if (proc.Count() > 0)
                {
                    swApp.Visible = true;
                    returnVal = true;
                }
            }
            catch (Exception e)
            {
                returnVal = false;
            }
            return returnVal;
        }

        public bool VerifyIFC_ComponentData(Component2 swComponent, double x, double y, double z)
        {
            double start = DateTime.Now.TimeOfDay.TotalMilliseconds;
            bool bRet = false;

            ModelDoc2 swCompModel = default;
            AssemblyDoc swAssyDoc = default;
            PartDoc swPartDoc = default;

            double[] dBoundingBox = default;
            MathUtility swMathUtil = default;
            MathPoint[] boundingBoxCornerPoints = default;
            MathTransform swTransform = default;

            double cornerPointTolerance = 5e-3; // bounding box is inaccurate and needs a tolerance range for comparison: [5 mm]

            bool cornerPointFound = false;

            try
            {
                //Console.WriteLine($"Checking ({x}, {y}, {z})");
                swCompModel = swComponent.GetModelDoc2() as ModelDoc2;

                switch (swCompModel.GetType())
                {
                    case (int)swDocumentTypes_e.swDocPART:
                        {
                            swPartDoc = swCompModel as PartDoc;
                            dBoundingBox = swPartDoc.GetPartBox(true) as double[];
                        }
                        break;

                    case (int)swDocumentTypes_e.swDocASSEMBLY:
                        {
                            swAssyDoc = swCompModel as AssemblyDoc;
                            dBoundingBox = swAssyDoc.GetBox((int)swBoundingBoxOptions_e.swBoundingBoxIncludeRefPlanes) as double[];
                        }
                        break;
                }

                swMathUtil = swApp.GetMathUtility() as MathUtility;
                swTransform = swComponent.Transform2;

                boundingBoxCornerPoints = new MathPoint[]
                {
                    swMathUtil.CreatePoint(new double[] { dBoundingBox[0], dBoundingBox[1], dBoundingBox[2]}) as MathPoint,
                    swMathUtil.CreatePoint(new double[] { dBoundingBox[0], dBoundingBox[1], dBoundingBox[5]}) as MathPoint,
                    swMathUtil.CreatePoint(new double[] { dBoundingBox[0], dBoundingBox[4], dBoundingBox[2]}) as MathPoint,
                    swMathUtil.CreatePoint(new double[] { dBoundingBox[0], dBoundingBox[4], dBoundingBox[5]}) as MathPoint,
                    swMathUtil.CreatePoint(new double[] { dBoundingBox[3], dBoundingBox[1], dBoundingBox[2]}) as MathPoint,
                    swMathUtil.CreatePoint(new double[] { dBoundingBox[3], dBoundingBox[1], dBoundingBox[5]}) as MathPoint,
                    swMathUtil.CreatePoint(new double[] { dBoundingBox[3], dBoundingBox[4], dBoundingBox[2]}) as MathPoint,
                    swMathUtil.CreatePoint(new double[] { dBoundingBox[3], dBoundingBox[4], dBoundingBox[5]}) as MathPoint,
                };

                bool foundx = false;
                bool foundy = false;
                bool foundz = false;

                for(int i=0; i<boundingBoxCornerPoints.Length; i++)
                {
                    MathPoint cornerPoint = boundingBoxCornerPoints[i];
                    cornerPoint = cornerPoint.MultiplyTransform(swTransform) as MathPoint;
                    double[] arrayData = cornerPoint.ArrayData as double[];
                    var cpX = Math.Round(arrayData[0], 6);
                    var cpY = Math.Round(arrayData[1], 6);
                    var cpZ = Math.Round(arrayData[2], 6);
                    if(!foundx)
                    {
                        var dx = Math.Abs(cpX - x);
                        //Console.WriteLine($"dx: {dx}");
                        if (dx <= cornerPointTolerance)
                            foundx = true;
                    }
                    if (!foundy)
                    {
                        var dy = Math.Abs(cpY - y);
                        //Console.WriteLine($"dy: {dy}");
                        if (dy <= cornerPointTolerance)
                            foundy = true;
                    }
                    if (!foundz)
                    {
                        var dz = Math.Abs(cpZ - z);
                        //Console.WriteLine($"dz: {dz}");
                        if (dz <= cornerPointTolerance)
                            foundz = true;
                    }

                    if(foundx && foundy && foundz)
                    {
                        cornerPointFound = true;
                        continue;
                    }
                }

                bRet = cornerPointFound;
            }
            catch(Exception)
            {

            }
            finally
            {
                double end = DateTime.Now.TimeOfDay.TotalMilliseconds;
                Console.WriteLine($"{nameof(VerifyIFC_ComponentData)}\t Duration: {(end - start)} milliseconds");
            }

            return bRet;
        }

        public void Dispose()
        {
            if(swApp != null)
                Marshal.FinalReleaseComObject(swApp);
            //Console.WriteLine("Disposed");
        }

        /// Functions within this region here are for troubleshooting and testing only
        #region - Not for production -

        public Component2 GetSelectedComponent()
        {
            Component2 swComponent = default;
            ModelDoc2 swModel = default;
            SelectionMgr swSelMan = default;

            try
            {
                swModel = swApp.ActiveDoc as ModelDoc2;
                swSelMan = swModel.SelectionManager as SelectionMgr;

                var selectionType = swSelMan.GetSelectedObjectType3(1, 0);

                if (selectionType == (int)swSelectType_e.swSelCOMPONENTS)
                {
                    swComponent = swSelMan.GetSelectedObject6(1, 0) as Component2;
                    Console.WriteLine("Selection: " + swComponent.Name2);
                }
                else
                {
                    Console.WriteLine("No component selected in Feature Tree");
                }
            }
            catch(Exception)
            {

            }

            return swComponent;

        }

        #endregion

    }
}
