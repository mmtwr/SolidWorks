using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using System.Runtime.InteropServices;
using System;
using System.Windows.Forms;
using System.Collections.Generic;

namespace SwTEst2
{
    class SolidWorksMacro2
    {
        public SldWorks swApp;

        public class Coordinate
        {
            private double x;
            public double X
            {
                get { return x; }
                set { x = value; }
            }
            private double y;
            public double Y
            {
                get { return y; }
                set { y = value; }
            }
            public Coordinate(double x1, double y1)
            {
                this.X = x1;
                this.Y = y1;
            }
        }

        public class Functions
        {
            public double[,] GetPositionMatrix(Component2 component)
            {
                double[] buffArray = (double[])component.Transform.ArrayData;
                double[,] matrix = new double[4, 4];
                matrix[0, 0] = buffArray[0];
                matrix[0, 1] = buffArray[1];
                matrix[0, 2] = buffArray[2];
                matrix[0, 3] = buffArray[9] * 1000;
                matrix[1, 0] = buffArray[3];
                matrix[1, 1] = buffArray[4];
                matrix[1, 2] = buffArray[5];
                matrix[1, 3] = buffArray[10] * 1000;
                matrix[2, 0] = buffArray[6];
                matrix[2, 1] = buffArray[7];
                matrix[2, 2] = buffArray[8];
                matrix[2, 3] = buffArray[11] * 1000;
                matrix[3, 0] = buffArray[15];
                matrix[3, 1] = buffArray[14];
                matrix[3, 2] = buffArray[13];
                matrix[3, 3] = buffArray[12];
                return matrix;
            }
            public Vertex GetVertexInGlobalCoordSystem(Vertex local, double[,] matrix)
            {
                Vertex global = new Vertex();
                if (matrix.GetLength(0) == 4 && matrix.GetLength(1) == 4)
                {
                    for (int i = 0; i < 4; i++)
                    {
                        global.X = local.X * matrix[0, 0] + local.Y * matrix[0, 1] + local.Z * matrix[0, 2] + 1 * matrix[0, 3];
                        global.Y = local.X * matrix[1, 0] + local.Y * matrix[1, 1] + local.Z * matrix[1, 2] + 1 * matrix[1, 3];
                        global.Z = local.X * matrix[2, 0] + local.Y * matrix[2, 1] + local.Z * matrix[2, 2] + 1 * matrix[2, 3];
                    }
                }
                return global;
            }
            public List<Coordinate> GetCoordinateFromVertex(Vertex firstVertex, Vertex secondVertex)
            {
                List<Coordinate> result = new List<Coordinate>();
                if (firstVertex.X == secondVertex.X)
                {
                    result.Add(new Coordinate(firstVertex.Y, firstVertex.Z));
                    result.Add(new Coordinate(secondVertex.Y, secondVertex.Z));
                }
                else
                {
                    if (firstVertex.Y == secondVertex.Y)
                    {
                        result.Add(new Coordinate(firstVertex.X, firstVertex.Z));
                        result.Add(new Coordinate(secondVertex.X, secondVertex.Z));
                    }
                    else
                    {
                        if (firstVertex.Z == secondVertex.Z)
                        {
                            result.Add(new Coordinate(firstVertex.X, firstVertex.Y));
                            result.Add(new Coordinate(secondVertex.X, secondVertex.Y));
                        }
                    }
                }
                return result;
            }
            public void GetPlateSize(List<Coordinate> coords, ref double width, ref double height)
            {
                double minX = int.MaxValue;
                double maxX = int.MinValue;
                double minY = int.MaxValue;
                double maxY = int.MinValue;
                foreach (Coordinate item in coords)
                {
                    if (minX > item.X) minX = item.X;
                    if (maxX < item.X) maxX = item.X;
                    if (minY > item.Y) minY = item.Y;
                    if (maxY < item.Y) maxY = item.Y;
                }
                width = maxX - minX;
                height = maxY - minY;
            }
            public string GetSurfaceNameForMate(Component2 component)
            {
                double[,] matrix = new double[4, 4];
                matrix = GetPositionMatrix(component);
                string feName = "Сверху"; //Спереди
                Feature rfPlane = component.FeatureByName(feName);
                object boxObject = null;
                rfPlane.GetBox(ref boxObject);
                double[] box = (double[])boxObject;
                List<Vertex> boxCoordinate = new List<Vertex>();
                boxCoordinate.Add(new Vertex(box[0], box[1], box[2]));
                boxCoordinate.Add(new Vertex(box[3], box[4], box[5]));
                List<Vertex> globalCoordinate = new List<Vertex>();
                for (int i = 0; i < boxCoordinate.Count; i++)
                {
                    globalCoordinate.Add(GetVertexInGlobalCoordSystem(boxCoordinate[i], matrix));
                }
                if (boxCoordinate[0].X == boxCoordinate[1].X) return "Спереди";
                if (boxCoordinate[0].Y == boxCoordinate[1].Y) return "Справа";
                if (boxCoordinate[0].Z == boxCoordinate[1].Z) return "Сверху";
                return null;
            }
        }

        public class Vertex
        {
            private double x;
            public double X
            {
                get { return x; }
                set { x = value; }
            }
            private double y;
            public double Y
            {
                get { return y; }
                set { y = value; }
            }
            private double z;
            public double Z
            {
                get { return z; }
                set { z = value; }
            }
            public Vertex()
            {
                this.X = 0; this.Y = 0; this.Z = 0;
            }

            public Vertex(double x1, double y1, double z1)
            {
                this.X = x1;
                this.Y = y1;
                this.Z = z1;
            }
        }

        public void Macro2()
        {
            swApp = (SldWorks)Activator.CreateInstance(Type.GetTypeFromProgID("SldWorks.Application"));
            ModelDoc2 swDoc = ((ModelDoc2)(swApp.ActiveDoc));
            AssemblyDoc swAssembly = ((AssemblyDoc)(swDoc));
            int longwarnings = 0;
            IModelDoc2 Part = (IModelDoc2)swApp.ActiveDoc;
            ISelectionMgr SelectionManager = (ISelectionMgr)Part.SelectionManager;
            int selectedObjectCount = SelectionManager.GetSelectedObjectCount();
            if (selectedObjectCount != 3) MessageBox.Show("Неверное количество поверхностей");
            else
            {
                List<object> selectedElements = new List<object>();
                bool isFace2 = false;
                for (int i = 0; i < selectedObjectCount; i++)
                {
                    selectedElements.Add(SelectionManager.GetSelectedObject6(i + 1, -1));
                    if (selectedElements[i] is IFace2)
                    {
                        isFace2 = true;
                    }
                }
                if (!isFace2)
                {
                    MessageBox.Show("Выбранные элементы не являются поверхностями");
                }
                else
                {
                    Functions function = new Functions();
                    List<Face2> faces = new List<Face2>();
                    List<Surface> surfaces = new List<Surface>();
                    List<double[,]> matrix = new List<double[,]>();
                    for (int i = 0; i < selectedObjectCount; i++)
                    {
                        faces.Add((Face2)selectedElements[i]);
                        surfaces.Add((Surface)faces[i].GetSurface()); matrix.Add(function.GetPositionMatrix((Component2)SelectionManager.GetSelectedObjectsComponent(i + 1)));
                    }
                    if (!(surfaces[0].IsPlane() && surfaces[1].IsPlane() && surfaces[2].IsPlane()))
                    {
                        MessageBox.Show("Не все выбранные элементы не являются типом Plane");
                    }
                    else
                    {
                        List<Vertex> boxCoordinate = new List<Vertex>();
                        for (int i = 0; i < selectedObjectCount; i++)
                        {
                            double[] bufferArray = (double[])faces[i].GetBox();//(double[])swAssembly.GetBox(2);
                            for (int index = 0; index < bufferArray.Length; index++)
                            {
                                bufferArray[index] = bufferArray[index] * 1000;
                            }
                            boxCoordinate.Add(new Vertex(bufferArray[0], bufferArray[1], bufferArray[2]));
                            boxCoordinate.Add(new Vertex(bufferArray[3], bufferArray[4], bufferArray[5]));
                        }
                        List<Vertex> globalCoordinate = new List<Vertex>();
                        int j = 0;
                        for (int i = 0; i < selectedObjectCount * 2; i += 2)
                        {
                            globalCoordinate.Add(function.GetVertexInGlobalCoordSystem(boxCoordinate[i], matrix[j]));
                            globalCoordinate.Add(function.GetVertexInGlobalCoordSystem(boxCoordinate[i + 1], matrix[j]));
                            j++;
                        }
                        matrix = null;
                        boxCoordinate = null;
                        List<Coordinate> coord = new List<Coordinate>();
                        for (int i = 0; i < selectedObjectCount * 2; i += 2)
                        {
                            coord.AddRange(function.GetCoordinateFromVertex(globalCoordinate[i], globalCoordinate[i + 1]));
                        }
                        double width = 0, height = 0;
                        function.GetPlateSize(coord, ref width, ref height);
                        string partName = @"E:\_Study\3 курс 2 сем\ОАК\Units\Korpus.SLDPRT";
                        int longstatus = 0;
                        swDoc = ((ModelDoc2)(swApp.ActiveDoc));
                        object preDefinition = swApp.OpenDoc6(partName, (int)swDocumentTypes_e.swDocPART, (int)swOpenDocOptions_e.swOpenDocOptions_Silent, "", ref longstatus, ref longwarnings);
                        Component2 part = swAssembly.AddComponent5(partName, (int)swAddComponentConfigOptions_e.swAddComponentConfigOptions_CurrentSelectedConfig, "", false, "", 0, 0, 0);
                        swApp.CloseDoc(partName);
                        string feName = "Сверху";
                        Feature rfPlane = part.FeatureByName(feName);
                        rfPlane.Select(true);
                        Entity faceToSelect = SelectPlaneSurface(faces, surfaces);
                        faceToSelect.Select(true);
                        IMate2 planeMate = swAssembly.AddMate3((int)swMateType_e.swMateCOINCIDENT, (int)swMateAlign_e.swMateAlignANTI_ALIGNED, false, 0, 0, 0, 0, 0, 0, 0, 0, false, out longwarnings);
                        EquationMgr equationManager = (EquationMgr)(((IModelDoc2)part.GetModelDoc()).GetEquationMgr());
                        //string findSubstring = "PlateWidth";
                        //equationManager.set_Equation(1, findSubstring);
                        ////ChangeParameterSize(width, equationManager, findSubstring);
                        //findSubstring = "PlateHeight";
                        ////ChangeParameterSize(height, equationManager, findSubstring);
                        //equationManager.set_Equation(1, findSubstring);
                        double B = width;
                        double L = height;
                        ModelDoc2 plate_modelDoc = (ModelDoc2)part.GetModelDoc2();
                        EquationMgr eqMgr_pl = plate_modelDoc.GetEquationMgr();

                        for (int i = 0; i < eqMgr_pl.GetCount(); i++)
                        {
                            if (eqMgr_pl.get_Equation(i).Contains("\"B\"="))
                            {
                                string str_eqB = eqMgr_pl.get_Equation(i);
                                str_eqB = str_eqB.Substring(0, str_eqB.LastIndexOf('=') + 1) + B.ToString();
                                eqMgr_pl.set_Equation(i, str_eqB);
                            }
                            if (eqMgr_pl.get_Equation(i).Contains("\"L\"="))
                            {
                                string str_eqL = eqMgr_pl.get_Equation(i);
                                str_eqL = str_eqL.Substring(0, str_eqL.LastIndexOf('=') + 1) + L.ToString();
                                eqMgr_pl.set_Equation(i, str_eqL);
                            }
                        }
                        eqMgr_pl.EvaluateAll();
                        plate_modelDoc.EditRebuild3();
                        DeselectAllSurfaces(SelectionManager);
                        swDoc.Rebuild(1);
                        feName = function.GetSurfaceNameForMate(part);//"Справа"; 
                        rfPlane = (Feature)swAssembly.FeatureByName(feName);
                        rfPlane.Select(true);
                        feName = "Спереди";
                        rfPlane = part.FeatureByName(feName);
                        rfPlane.Select(true);
                        planeMate = swAssembly.AddMate3((int)swMateType_e.swMatePARALLEL, (int)swMateAlign_e.swAlignNONE, false, 0, 0, 0, 0, 0, 0, 0, 0, false, out longwarnings);
                        DeselectAllSurfaces(SelectionManager);
                    }

                }
            }

        }

        public Entity SelectPlaneSurface(List<Face2> faces, List<Surface> surfaces)
        {
            Entity faceToSelect = null;
            for (int i = 0; i < surfaces.Count; i++)
            {
                if (surfaces[i].IsPlane())
                {
                    faceToSelect = (Entity)faces[i];
                    break;
                }
            }
            return faceToSelect;
        }

        private void DeselectAllSurfaces(ISelectionMgr SelectionManager)
        {
            int selectedObjectCount = SelectionManager.GetSelectedObjectCount();
            for (int i = 0; i < selectedObjectCount; i++)
            {
                SelectionManager.DeSelect(1);
            }
        }
    }
}
