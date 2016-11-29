using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using System.Runtime.InteropServices;
using System;
using System.Windows.Forms;
using System.Collections.Generic;


namespace Macro1.csproj
{
    public class SolidWorksMacro
    {
        public SldWorks swApp;

        public void Macro1()
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
                List<object> se = new List<object>();
                bool isFace2 = false;
                for (int i = 0; i < selectedObjectCount; i++)
                {
                    se.Add(SelectionManager.GetSelectedObject6(i + 1, -1));
                    if (se[i] is IFace2)
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
                    List<Face2> faces = new List<Face2>();
                    List<Surface> surfaces = new List<Surface>();
                    List<Face2> faces1 = new List<Face2>();
                    List<Surface> surfaces1 = new List<Surface>();
                    for (int i = 0; i < selectedObjectCount; i++)
                    {
                        faces.Add((Face2)se[i]);
                        surfaces.Add((Surface)faces[i].GetSurface());
                    }
                    for (int i = 0; i < selectedObjectCount; i++)
                    {
                        faces1.Add((Face2)se[i]);
                        surfaces1.Add((Surface)faces1[i].GetSurface());
                    }

                    //////////////////////////////////////////////
                    //вставка первого пальца
                    string AdPrtName = @"E:\_Study\3 курс 2 сем\ОАК\Units\Locator1.SLDPRt";
                    int longstatus = 0;
                    swDoc = ((ModelDoc2)(swApp.ActiveDoc));
                    object Preop;
                    Preop = swApp.OpenDoc6(AdPrtName, (int)swDocumentTypes_e.swDocPART,
                        (int)swOpenDocOptions_e.swOpenDocOptions_Silent, "", ref longstatus, ref longwarnings);
                    Component2 Np;
                    Np = swAssembly.AddComponent5(AdPrtName,
                        (int)swAddComponentConfigOptions_e.swAddComponentConfigOptions_CurrentSelectedConfig, "", false, "", 0, 0, 0);
                    swApp.CloseDoc(AdPrtName);

                    //сопряжение поверхностей
                    string feName = "Спереди"; //Спереди
                    Feature rfPlane = Np.FeatureByName(feName);

                    //отменить выделение всех поверхностей
                    DeselectAllSurfaces(SelectionManager);

                    //выбор первого выбранного цилиндра
                    double cylinderDiameter = 0;
                    Entity faceToSelect = GetFirstSelectedCylinder(faces, surfaces, ref cylinderDiameter);
                    //выбор цилиндра с минимальным диаметром - палец
                    GetCylinderWithMinDiameterInFirstFinger(Np);
                    IMate2 planeMate = swAssembly.AddMate3((int)swMateType_e.swMateCONCENTRIC, -1, false, 0, 0, 0, 0, 0, 0.5, 0.5, 0.5, false, out longwarnings);
                    DeselectAllSurfaces(SelectionManager);
                    swDoc.Rebuild(1);

                    EquationMgr equationManager = (EquationMgr)(((IModelDoc2)Np.GetModelDoc()).GetEquationMgr());
                    string findSubstring = "FirstFingerDiameter";
                    ChooseDiameterInFinger(cylinderDiameter, equationManager, findSubstring);
                    swDoc.Rebuild(1);
                    MovingComponent(Np, new double[] { 0, 0, 0.045 });

                    ////////////////////////////////////////////////
                    //вставка второго пальца
                    AdPrtName = @"E:\_Study\3 курс 2 сем\ОАК\Units\Locator1.SLDPRt";
                    longstatus = 0;
                    swDoc = ((ModelDoc2)(swApp.ActiveDoc));
                    Preop = swApp.OpenDoc6(AdPrtName, (int)swDocumentTypes_e.swDocPART,
                        (int)swOpenDocOptions_e.swOpenDocOptions_Silent, "", ref longstatus, ref longwarnings);
                    Np = swAssembly.AddComponent5(AdPrtName,
                        (int)swAddComponentConfigOptions_e.swAddComponentConfigOptions_CurrentSelectedConfig, "", false, "", 0, 0, 0);
                    swApp.CloseDoc(AdPrtName);

                    //сопряжение поверхностей
                    feName = "Спереди"; //Спереди
                    rfPlane = Np.FeatureByName(feName);
                    rfPlane.Select(true);

                    //отменить выделение всех поверхностей
                    DeselectAllSurfaces(SelectionManager);

                    //выбор второго выбранного цилиндра
                    cylinderDiameter = 0;
                    faceToSelect = GetSecondSelectedCylinder(faces, surfaces, ref cylinderDiameter);
                    //выбор цилиндра с минимальным диаметром - палец
                    GetCylinderWithMinDiameterInFirstFinger(Np);
                    planeMate = swAssembly.AddMate3((int)swMateType_e.swMateCONCENTRIC, -1, false, 0, 0, 0, 0, 0, 0.5, 0.5, 0.5, false, out longwarnings);
                    DeselectAllSurfaces(SelectionManager);
                    swDoc.Rebuild(1);

                    equationManager = (EquationMgr)(((IModelDoc2)Np.GetModelDoc()).GetEquationMgr());
                    findSubstring = "FirstFingerDiameter";
                    ChooseDiameterInFinger(cylinderDiameter, equationManager, findSubstring);
                    swDoc.Rebuild(1);
                    MovingComponent(Np, new double[] { 0, 0, 0.1 });

                    ////////////////////////////////////////////Вставка пластины1
                    string AdPrtName3 = @"E:\_Study\3 курс 2 сем\ОАК\Units\Locator7.SLDPRt";
                    object Preop3;
                    swDoc = ((ModelDoc2)(swApp.ActiveDoc));
                    Preop3 = swApp.OpenDoc6(AdPrtName3, (int)swDocumentTypes_e.swDocPART, (int)swOpenDocOptions_e.swOpenDocOptions_Silent, "", ref longstatus, ref longwarnings);
                    Component2 Np1;
                    Np1 = swAssembly.AddComponent5(AdPrtName3, (int)swAddComponentConfigOptions_e.swAddComponentConfigOptions_CurrentSelectedConfig, "", false, "", 0, 0, 0);
                    swApp.CloseDoc(AdPrtName3);
                    MovingComponent(Np1, new double[] { 0.095, 0.095, 0 });

                    feName = "Сверху";
                    rfPlane = Np1.FeatureByName(feName);
                    rfPlane.Select(true);
                    faceToSelect = SelectPlaneSurface(faces, surfaces);
                    faceToSelect.Select(true);
                    planeMate = swAssembly.AddMate3((int)swMateType_e.swMateCOINCIDENT, -1, false, 0, 0, 0, 0, 0, 0, 0, 0, false, out longwarnings);

                    DeselectAllSurfaces(SelectionManager);
                    swDoc.Rebuild(1);


                    ////////////////////////////////////////////Вставка пластины2
                    AdPrtName3 = @"E:\_Study\3 курс 2 сем\ОАК\Units\Locator7.SLDPRt";
                    swDoc = ((ModelDoc2)(swApp.ActiveDoc));
                    Preop3 = swApp.OpenDoc6(AdPrtName3, (int)swDocumentTypes_e.swDocPART, (int)swOpenDocOptions_e.swOpenDocOptions_Silent, "", ref longstatus, ref longwarnings);
                    Component2 Np2 = swAssembly.AddComponent5(AdPrtName3, (int)swAddComponentConfigOptions_e.swAddComponentConfigOptions_CurrentSelectedConfig, "", false, "", 0, 0, 0);
                    swApp.CloseDoc(AdPrtName3);
                    MovingComponent(Np2, new double[] { -0.095, 0.095, 0 });

                    feName = "Сверху";
                    rfPlane = Np2.FeatureByName(feName);
                    rfPlane.Select(true);
                    faceToSelect = SelectPlaneSurface(faces, surfaces);
                    faceToSelect.Select(true);
                    planeMate = swAssembly.AddMate3((int)swMateType_e.swMateCOINCIDENT, -1, false, 0, 0, 0, 0, 0, 0, 0, 0, false, out longwarnings);

                    DeselectAllSurfaces(SelectionManager);
                    swDoc.Rebuild(1);


                    ////////////////////////////////////////////Вставка пластины3
                    AdPrtName3 = @"E:\_Study\3 курс 2 сем\ОАК\Units\Locator7.SLDPRt";
                    swDoc = ((ModelDoc2)(swApp.ActiveDoc));
                    Preop3 = swApp.OpenDoc6(AdPrtName3, (int)swDocumentTypes_e.swDocPART,
                        (int)swOpenDocOptions_e.swOpenDocOptions_Silent, "", ref longstatus, ref longwarnings);
                    Component2 Np3 = swAssembly.AddComponent5(AdPrtName3,
                        (int)swAddComponentConfigOptions_e.swAddComponentConfigOptions_CurrentSelectedConfig, "", false, "", 0, 0, 0);
                    swApp.CloseDoc(AdPrtName3);
                    MovingComponent(Np3, new double[] { 0, -0.134, 0 });

                    feName = "Сверху";
                    rfPlane = Np3.FeatureByName(feName);
                    rfPlane.Select(true);
                    faceToSelect = SelectPlaneSurface(faces, surfaces);
                    faceToSelect.Select(true);
                    planeMate = swAssembly.AddMate3((int)swMateType_e.swMateCOINCIDENT, -1, false, 0, 0, 0, 0, 0, 0, 0, 0, false, out longwarnings);

                    DeselectAllSurfaces(SelectionManager);
                    swDoc.Rebuild(1);
                }
            }
        }

        private Entity GetSecondSelectedCylinder(List<Face2> faces, List<Surface> surfaces, ref double cylinderDiameter)
        {
            Entity faceToSelect = null;
            for (int i = 0; i < surfaces.Count; i++)
            {
                if (surfaces[i].IsCylinder())
                {
                    faceToSelect = (Entity)faces[i];
                    double[] param = (double[])((Surface)faces[i].GetSurface()).CylinderParams;
                    cylinderDiameter = param[6] * 1000;
                }
            }
            faceToSelect.Select(true);
            return faceToSelect;
        }

        private void ChooseDiameterInFinger(double cylinderDiameter, EquationMgr equationManager, string findSubstring)
        {
            int equationCount = equationManager.GetCount();
            string equation = string.Empty;
            int index = 0;
            for (int i = 0; i < equationCount; i++)
            {
                equation = equationManager.get_Equation(index);
                if (equation.Contains(findSubstring))
                {
                    break;
                }
                else
                {
                    index++;
                }
            }

            string[] substringArray = equation.Split('=');
            string[] parameterArray = cylinderDiameter.ToString().Split(',');
            substringArray[1] = parameterArray[0];
            if (parameterArray.Length == 2) substringArray[1] += "." + parameterArray[1];
            equation = substringArray[0] + "=" + substringArray[1];
            equationManager.set_Equation(index, equation);
            equationManager.EvaluateAll();
        }

        private void GetCylinderWithMinDiameterInFirstFinger(Component2 Np)
        {
            Body body = (Body)Np.GetBody();
            int count = body.GetFaceCount();
            Face2 faceToSelect2 = (Face2)body.GetFirstFace();
            Face2 buffFace = (Face2)body.GetFirstFace();
            if (!((Surface)faceToSelect2.GetSurface()).IsCylinder())
            {
                while (count > 1)
                {
                    faceToSelect2 = (Face2)faceToSelect2.GetNextFace();
                    if (((Surface)faceToSelect2.GetSurface()).IsCylinder())
                    {
                        if (((Surface)buffFace.GetSurface()).IsCylinder())
                        {
                            double[] param1 = (double[])((Surface)buffFace.GetSurface()).CylinderParams;
                            double[] param2 = (double[])((Surface)faceToSelect2.GetSurface()).CylinderParams;
                            if (param1[6] > param2[6])
                            {
                                buffFace = faceToSelect2;
                            }
                        }
                        else
                        {
                            buffFace = faceToSelect2;
                        }
                    }
                    count--;
                }
            }

            ((Entity)buffFace).Select(true);
        }

        private Entity GetFirstSelectedCylinder(List<Face2> faces, List<Surface> surfaces, ref double cylinderDiameter)
        {
            Entity faceToSelect = null;
            for (int i = 0; i < surfaces.Count; i++)
            {
                if (surfaces[i].IsCylinder())
                {
                    faceToSelect = (Entity)faces[i];
                    double[] param = (double[])((Surface)faces[i].GetSurface()).CylinderParams;
                    cylinderDiameter = param[6] * 2000;
                    break;
                }
            }
            faceToSelect.Select(true);
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

        public Entity SelectPlaneSurface1(List<Face2> faces1, List<Surface> surfaces1)
        {
            Entity faceToSelect1 = null;
            for (int i = 0; i < surfaces1.Count; i++)
            {
                if (surfaces1[i].IsPlane())
                {
                    faceToSelect1 = (Entity)faces1[i + 1];
                    break;
                }
            }
            return faceToSelect1;
        }

        private void MovingComponent(Component2 component, double[] Point)
        {
            MathTransform Tnew = (MathTransform)((MathUtility)swApp.GetMathUtility()).CreateTransform(component.Transform.ArrayData);
            object X = new object();
            object Y = new object();
            object Z = new object();
            object P = new object();
            double S = 0;
            component.Transform.GetData2(ref X, ref Y, ref Z, ref P, ref S);
            //double[] _mass = (double[])((MathVector)_P).ArrayData;
            ((MathVector)P).ArrayData = Point;
            Tnew.SetData(X, Y, Z, P, S);
            component.Transform = Tnew;
            ((ModelDoc2)swApp.ActiveDoc).EditRebuild3();
        }
    }
}