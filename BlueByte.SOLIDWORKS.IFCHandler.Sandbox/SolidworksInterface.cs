﻿using System;
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
        private SldWorks sw_App;
        public ref SldWorks SolidworksApp { get { return ref this.sw_App; } }
        public bool IsConnected { get { return isConnected; } }
        public SolidworksInterface(SldWorks solidworksApp = null)
        {
            sw_App = solidworksApp;
            if (sw_App == null)
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
                sw_App = (SldWorks)Marshal.GetActiveObject("SldWorks.Application");
                Process[] proc = Process.GetProcessesByName("SLDWORKS");
                if (proc.Count() > 0)
                {
                    sw_App.Visible = true;
                    returnVal = true;
                }
            }
            catch (Exception e)
            {
                returnVal = false;
            }
            return returnVal;
        }
        /// <summary>
        /// Compares the entered coordinates with those of the Component2 object relative to the assembly it is active within
        /// </summary>
        /// <param name="swComponent">Component2 object from Top Level Solidworks Assembly</param>
        /// <param name="x">x component of IFC bounding box corner</param>
        /// <param name="y">y component of IFC bounding box corner</param>
        /// <param name="z">z component of IFC bounding box corner</param>
        /// <returns>true if match is likely -or- false if match is unlikely</returns>
        public bool VerifyIFC_ComponentData(ref SldWorks swApp, Component2 swComponent, double x, double y, double z)
        {
            bool bRet = false;

            ModelDoc2 swCompModel = default;
            AssemblyDoc swAssyDoc = default;
            PartDoc swPartDoc = default;

            double[] dBoundingBox = default;
            double[][] boundingBoxCornerPoints = default;
            MathUtility swMathUtil = default;
            MathTransform swTransform = default;
            MathPoint mp = default;

            double cornerPointTolerance = 5e-3; // bounding box is inaccurate and needs a tolerance range for comparison: [5 mm]

            bool cornerPointFound = false;

            try
            {
                Console.WriteLine($"Checking ({x}, {y}, {z})");
                swCompModel = swComponent.GetModelDoc2() as ModelDoc2;

                swMathUtil = swApp.GetMathUtility() as MathUtility;
                swTransform = swComponent.Transform2;

                // Check document type because built-in methods are different for both assy and part
                switch (swCompModel.GetType())
                {
                    case (int)swDocumentTypes_e.swDocPART:
                        {
                            swPartDoc = swCompModel as PartDoc;
                            dBoundingBox = swPartDoc.GetPartBox(true) as double[];
                            //dBoundingBox = GetAccurateBoundingBox(ref swPartDoc, ref swMathUtil, ref swTransform); // This is not more accurate than built-in method
                        }
                        break;

                    case (int)swDocumentTypes_e.swDocASSEMBLY:
                        {
                            swAssyDoc = swCompModel as AssemblyDoc;
                            dBoundingBox = swAssyDoc.GetBox((int)swBoundingBoxOptions_e.swBoundingBoxIncludeRefPlanes) as double[];
                        }
                        break;
                }

                // Turn 2 corner points into 8 total points representing the corners of the bounding box
                boundingBoxCornerPoints = new double[][]
                {
                    new double[] { dBoundingBox[0], dBoundingBox[1], dBoundingBox[2]},
                    new double[] { dBoundingBox[0], dBoundingBox[1], dBoundingBox[5]},
                    new double[] { dBoundingBox[0], dBoundingBox[4], dBoundingBox[2]},
                    new double[] { dBoundingBox[0], dBoundingBox[4], dBoundingBox[5]},
                    new double[] { dBoundingBox[3], dBoundingBox[1], dBoundingBox[2]},
                    new double[] { dBoundingBox[3], dBoundingBox[1], dBoundingBox[5]},
                    new double[] { dBoundingBox[3], dBoundingBox[4], dBoundingBox[2]},
                    new double[] { dBoundingBox[3], dBoundingBox[4], dBoundingBox[5]},
                };

                // Transform the bounding box points from part space to assembly space
                for(int i = 0; i<boundingBoxCornerPoints.Length; i++)
                {
                    mp = ((MathPoint)swMathUtil.CreatePoint(boundingBoxCornerPoints[i])).MultiplyTransform(swTransform) as MathPoint;
                    boundingBoxCornerPoints[i] = mp.ArrayData as double[];
                }
                //DrawBoundingBoxPoints(boundingBoxCornerPoints); // Un-comment this to see a draw the points in a 3D sketch

                // Modify bounding box planes to align with assembly X, Y, & Z axes
                dBoundingBox = GetBoundingBoxRelativeToAssyPlanes(boundingBoxCornerPoints);
                
                // Again, turn the 2 corner points returned into 8 representing the corners of the bounding box
                boundingBoxCornerPoints = new double[][]
                {
                    new double[] { dBoundingBox[0], dBoundingBox[1], dBoundingBox[2]},
                    new double[] { dBoundingBox[0], dBoundingBox[1], dBoundingBox[5]},
                    new double[] { dBoundingBox[0], dBoundingBox[4], dBoundingBox[2]},
                    new double[] { dBoundingBox[0], dBoundingBox[4], dBoundingBox[5]},
                    new double[] { dBoundingBox[3], dBoundingBox[1], dBoundingBox[2]},
                    new double[] { dBoundingBox[3], dBoundingBox[1], dBoundingBox[5]},
                    new double[] { dBoundingBox[3], dBoundingBox[4], dBoundingBox[2]},
                    new double[] { dBoundingBox[3], dBoundingBox[4], dBoundingBox[5]},
                };
                //DrawBoundingBoxPoints(boundingBoxCornerPoints);  // Un-comment this to see a 3D sketch in the assembly to show this.


                // Calculate the distance between calculated X, Y and Z inputs individually. Each should be less than the cornerPointTolerance constant. If so, return true.
                for(int i=0; i<boundingBoxCornerPoints.Length; i++)
                {
                    double[] cornerPoint = boundingBoxCornerPoints[i];
                    
                    var cpX = Math.Round(cornerPoint[0], 6);
                    var cpY = Math.Round(cornerPoint[1], 6);
                    var cpZ = Math.Round(cornerPoint[2], 6);
                    
                    double dist = Math.Sqrt(Math.Pow(cpX - x, 2) + Math.Pow(cpY - y, 2) + Math.Pow(cpZ - z, 2));
                    Console.WriteLine($"\n         ({cpX}, {cpY}, {cpZ})  =>  Delta=({Math.Abs(cpX - x)}, {Math.Abs(cpY - y)}, {Math.Abs(cpZ - z)})  TOL: {cornerPointTolerance}\n");
                    if (Math.Abs(cpX - x) < cornerPointTolerance &&
                        Math.Abs(cpY - y) < cornerPointTolerance &&
                        Math.Abs(cpZ - z) < cornerPointTolerance)
                    {
                        cornerPointFound = true;
                        break;
                    }
                }

                bRet = cornerPointFound;
            }
            catch(Exception)
            {

            }
            finally
            {
                if (swCompModel != null)
                    Marshal.FinalReleaseComObject(swCompModel);
                if (swAssyDoc != null)
                    Marshal.FinalReleaseComObject(swAssyDoc);
                if (swPartDoc != null)
                    Marshal.FinalReleaseComObject(swPartDoc);
                if (swMathUtil != null)
                    Marshal.FinalReleaseComObject(swMathUtil);
                if (swTransform != null)
                    Marshal.FinalReleaseComObject(swTransform);
                if (mp != null)
                    Marshal.FinalReleaseComObject(mp);
            }

            return bRet;
        }

        private double[] GetBoundingBoxRelativeToAssyPlanes(double[][] points)
        {
            double[] dBox = new double[6];
            double minX = 0;
            double minY = 0;
            double minZ = 0;
            double maxX = 0;
            double maxY = 0;
            double maxZ = 0;

            try
            {
                for (int i = 0; i < points.Length; i++)
                {
                    double[] point = points[i];
                    double x = point[0];
                    double y = point[1];
                    double z = point[2];

                    
                    if (i == 0 || x > maxX)
                        maxX = x;
                    
                    if (i == 0 || x < minX)
                        minX = x;
                    
                    if (i == 0 || y > maxY)
                        maxY = y;
                    
                    if (i == 0 || y < minY)
                        minY = y;
                    
                    if (i == 0 || z > maxZ)
                        maxZ = z;
                    
                    if (i == 0 || z < minZ)
                        minZ = z;
                }

                dBox[0] = minX;
                dBox[1] = minY;
                dBox[2] = minZ;
                dBox[3] = maxX;
                dBox[4] = maxY;
                dBox[5] = maxZ;
            }
            catch (Exception)
            {

            }
            return dBox;
        }

        public void Dispose()
        {
            if(sw_App != null)
                Marshal.FinalReleaseComObject(sw_App);
            //Console.WriteLine("Disposed");
        }

        /// Functions within this region here are for troubleshooting and testing only
        #region - Not for production -
        private double[] GetAccurateBoundingBox(ref PartDoc swPart, ref MathUtility swMathUtil, ref MathTransform swTransform)
        {
            double[] dBox = new double[6];
            double minX = 0;
            double minY = 0;
            double minZ = 0;
            double maxX = 0;
            double maxY = 0;
            double maxZ = 0;
            MathVector[] normalVectors = new MathVector[]
                {
                    swMathUtil.CreateVector(new double[] { 1, 0, 0}) as MathVector,
                    swMathUtil.CreateVector(new double[] { -1, 0, 0}) as MathVector,
                    swMathUtil.CreateVector(new double[] { 0, 1, 0}) as MathVector,
                    swMathUtil.CreateVector(new double[] { 0, -1, 0}) as MathVector,
                    swMathUtil.CreateVector(new double[] { 0, 0, 1}) as MathVector,
                    swMathUtil.CreateVector(new double[] { 0, 0, -1}) as MathVector,
                };

            for (int i = 0; i < normalVectors.Length - 1; i++)
            {
                //normalVectors[i] = normalVectors[i].MultiplyTransform(swTransform.Inverse()) as MathVector;
            }

            try
            {
                var bodies = swPart.GetBodies2((int)swBodyType_e.swSolidBody, true) as object[];
                if (bodies != null)
                {
                    MathPoint mp = default;
                    for (int i = 0; i < bodies.Length; i++)
                    {
                        var body = bodies[i] as Body2;
                        double x = 0;
                        double y = 0;
                        double z = 0;

                        double[] nv = normalVectors[0].ArrayData as double[];
                        body.GetExtremePoint(nv[0], nv[1], nv[2], out x, out y, out z);
                        //mp = swMathUtil.CreatePoint(new double[] { x, y, z }) as MathPoint;
                        //mp = mp.MultiplyTransform(swTransform.Inverse()) as MathPoint;
                        //DrawBoundingBoxPoints(new MathPoint[] { mp });
                        //double[] arrayData = mp.ArrayData as double[];
                        //x = arrayData[0];
                        //y = arrayData[1];
                        //z = arrayData[2];
                        if (i == 0 || x > maxX)
                            maxX = x;

                        nv = normalVectors[1].ArrayData as double[];
                        body.GetExtremePoint(nv[0], nv[1], nv[2], out x, out y, out z);
                        //mp = swMathUtil.CreatePoint(new double[] { x, y, z }) as MathPoint;
                        //mp = mp.MultiplyTransform(swTransform.Inverse()) as MathPoint;
                        //DrawBoundingBoxPoints(new MathPoint[] { mp });
                        //arrayData = mp.ArrayData as double[];
                        //x = arrayData[0];
                        //y = arrayData[1];
                        //z = arrayData[2];
                        if (i == 0 || x < minX)
                            minX = x;

                        nv = normalVectors[2].ArrayData as double[];
                        body.GetExtremePoint(nv[0], nv[1], nv[2], out x, out y, out z);
                        //mp = swMathUtil.CreatePoint(new double[] { x, y, z }) as MathPoint;
                        //mp = mp.MultiplyTransform(swTransform.Inverse()) as MathPoint;
                        //DrawBoundingBoxPoints(new MathPoint[] { mp });
                        //arrayData = mp.ArrayData as double[];
                        //x = arrayData[0];
                        //y = arrayData[1];
                        //z = arrayData[2];
                        if (i == 0 || y > maxY)
                            maxY = y;

                        nv = normalVectors[3].ArrayData as double[];
                        body.GetExtremePoint(nv[0], nv[1], nv[2], out x, out y, out z);
                        //mp = swMathUtil.CreatePoint(new double[] { x, y, z }) as MathPoint;
                        //mp = mp.MultiplyTransform(swTransform.Inverse()) as MathPoint;
                        //DrawBoundingBoxPoints(new MathPoint[] { mp });
                        //arrayData = mp.ArrayData as double[];
                        //x = arrayData[0];
                        //y = arrayData[1];
                        //z = arrayData[2];
                        if (i == 0 || y < minY)
                            minY = y;

                        nv = normalVectors[4].ArrayData as double[];
                        body.GetExtremePoint(nv[0], nv[1], nv[2], out x, out y, out z);
                        //mp = swMathUtil.CreatePoint(new double[] { x, y, z }) as MathPoint;
                        //mp = mp.MultiplyTransform(swTransform.Inverse()) as MathPoint;
                        //DrawBoundingBoxPoints(new MathPoint[] { mp });
                        //arrayData = mp.ArrayData as double[];
                        //x = arrayData[0];
                        //y = arrayData[1];
                        //z = arrayData[2];
                        if (i == 0 || z > maxZ)
                            maxZ = z;

                        nv = normalVectors[5].ArrayData as double[];
                        body.GetExtremePoint(nv[0], nv[1], nv[2], out x, out y, out z);
                        //mp = swMathUtil.CreatePoint(new double[] { x, y, z }) as MathPoint;
                        //mp = mp.MultiplyTransform(swTransform.Inverse()) as MathPoint;
                        //DrawBoundingBoxPoints(new MathPoint[] { mp });
                        //arrayData = mp.ArrayData as double[];
                        //x = arrayData[0];
                        //y = arrayData[1];
                        //z = arrayData[2];
                        if (i == 0 || z < minZ)
                            minZ = z;

                    }

                    dBox[0] = minX;
                    dBox[1] = minY;
                    dBox[2] = minZ;
                    dBox[3] = maxX;
                    dBox[4] = maxY;
                    dBox[5] = maxZ;
                }
            }
            catch (Exception)
            {

            }
            return dBox;
        }
        public Component2 GetSelectedComponent()
        {
            Component2 swComponent = default;
            ModelDoc2 swModel = default;
            SelectionMgr swSelMan = default;

            try
            {
                swModel = sw_App.ActiveDoc as ModelDoc2;
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

        private void DrawBoundingBoxPoints(MathPoint[] cornerPoints)
        {
            ModelDoc2 swModel = sw_App.ActiveDoc as ModelDoc2;
            SketchManager swSketchMan = swModel.SketchManager as SketchManager;

            swModel.ClearSelection2(true);

            swSketchMan.Insert3DSketch(true);
            swSketchMan.AddToDB = true;

            for(int i=0; i<cornerPoints.Length; i++)
            {
                double[] arrayPt1 = cornerPoints[i].ArrayData as double[];
                swSketchMan.CreatePoint(arrayPt1[0], arrayPt1[1], arrayPt1[2]);
            }

            swSketchMan.AddToDB = false;
            swSketchMan.Insert3DSketch(true);
        }

        #endregion

    }
}
