using System;
using System.Collections.Generic;
using System.Reflection;
using System.Windows.Media.Imaging;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.Creation;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System.IO;
using System.Windows.Controls;
using System.Drawing;



namespace ML1Project
{
    public class MainClass : IExternalApplication
    {
        public Result OnStartup(UIControlledApplication application)
        {
            //Ribbon tab
            application.CreateRibbonTab("NEW RIBBON");

            //Ribbon panel
            RibbonPanel ribbonPanel = application.CreateRibbonPanel("NEW RIBBON","Fundaciones");

            //Create a button to trigger some command and add it to the Ribbon Panel
            string thisAssemblyPath = Assembly.GetExecutingAssembly().Location;
            PushButtonData buttonData = new PushButtonData("Boton1", "Coordenadas\npilotes", thisAssemblyPath, "ML1Project.GetCoordinates");

            PushButton pushButton = ribbonPanel.AddItem(buttonData) as PushButton;

            //optionally, you can add other properties to the button, like for example a Tool-tip
            pushButton.ToolTip = "Obtiene las coordenadas del centro de los pilotes de acuerdo al Survey Point.";

            //Bitmap Icon
            Uri urlImage = new Uri("coordinates.png", UriKind.Relative);
            //pushButton.LargeImage = new BitmapImage(urlImage);
            pushButton.LargeImage = LoadEmbeddedImage("coordinates.png");


            return Result.Succeeded;


        }
        public Result OnShutdown(UIControlledApplication application)
        {
            return Result.Succeeded;
        }

        public static BitmapImage LoadEmbeddedImage(string resourceName)
        {
            // Get the assembly where the image is embedded
            Assembly assembly = Assembly.GetExecutingAssembly();

            // Build the full name of the embedded resource
            string fullResourceName = assembly.GetName().Name + "." + resourceName;

            // Load the embedded resource stream
            using (Stream stream = assembly.GetManifestResourceStream(fullResourceName))
            {
                if (stream != null)
                {
                    // Create a BitmapImage from the stream
                    BitmapImage bitmapImage = new BitmapImage();
                    bitmapImage.BeginInit();
                    bitmapImage.StreamSource = stream;
                    bitmapImage.CacheOption = BitmapCacheOption.OnLoad;
                    bitmapImage.EndInit();
                    bitmapImage.Freeze(); // Freeze the BitmapImage to make it immutable
                    return bitmapImage;
                }
                else
                {
                    throw new Exception("Resource not found: " + fullResourceName);
                }
            }
        }

    }

    [Transaction(TransactionMode.Manual)]
    public class GetCoordinates : IExternalCommand
    {

        //Función que verifica si el documento se encuentra en la vista 3D
        public Boolean VerificarVista3D(Autodesk.Revit.DB.Document doc)
        {
            if (doc.ActiveView is Autodesk.Revit.DB.View3D)
            {
                return true;
            }
            else
            {
                return false;
            }

        }

        //Función que obtiene todos los elementos que sean una instancia familias
        public List<FamilyInstance> GetFamilyInstances(ICollection<Element> elementsInView)
        {

            //Inicializa la lista 
            List<FamilyInstance> familyInstances = new List<FamilyInstance>();

            //Recorre la lista
            foreach (Element element in elementsInView)
            {
                if (element is FamilyInstance familyInstance)
                {
                    familyInstances.Add(familyInstance);
                }
            }
            return familyInstances;
        }

        //Función que filtra las instancias de familias las cuales cumplan con el requisito de la categoria y el nombre
        public List<FamilyInstance> GetPileInstances(List<FamilyInstance> familyInstances)
        {
            List<FamilyInstance> pileInstances = new List<FamilyInstance>();

            foreach (FamilyInstance familyInstance in familyInstances)
            {
                if (familyInstance.Category.Name == "Structural Foundations" && (familyInstance.Symbol.Family.Name.Contains("PILE") || familyInstance.Symbol.Family.Name.Contains("PILOTE")))
                {

                    //Extrae los nombres de los parámetros de la instancia de la familia
                    ParameterSet parameterSet = familyInstance.Parameters;

                    //inicializa la lista para almacenar los nombres de los parámetros
                    List<String> paramNombres = new List<String>();
                    foreach (Parameter parameter in parameterSet)
                    {
                        paramNombres.Add(parameter.Definition.Name);
                    }

                    //Verifica que contenga los parámetros "Coord_E" y "Coord_N"
                    if (paramNombres.Contains("Coord_E") && paramNombres.Contains("Coord_N"))
                    {
                        pileInstances.Add(familyInstance);
                    }
                }
            }
            return pileInstances;
        }

        public XYZ ToSharedCoordinates(XYZ xyz, Autodesk.Revit.DB.Document document)
        {
#if V_2024 || V_2022
            //Desire units 
            var lenghtUnits = document.GetUnits().GetFormatOptions(SpecTypeId.Length).GetUnitTypeId();
            //Trans matrix survey point
            var surveyTransform = document.ActiveProjectLocation.GetTotalTransform();
            //return
            return this.ToInternal(surveyTransform.Inverse.OfPoint(xyz), lenghtUnits);
#elif V_2020
            //Desire units 
            var lenghtUnits = document.GetUnits().GetFormatOptions(UnitType.UT_Length).DisplayUnits;
            //Trans matrix survey point
            var surveyTransform = document.ActiveProjectLocation.GetTotalTransform();
            //return
            return this.ToInternal(surveyTransform.Inverse.OfPoint(xyz), lenghtUnits);


#endif

        }

#if V_2024 || V_2022
        public XYZ ToInternal(XYZ xyz, ForgeTypeId internalType)
        {
            var x = UnitUtils.ConvertFromInternalUnits(xyz.X, internalType);
            var y = UnitUtils.ConvertFromInternalUnits(xyz.Y, internalType);
            var z = UnitUtils.ConvertFromInternalUnits(xyz.Z, internalType);

            return new XYZ(x, y, z);
        }
#elif V_2020
        public XYZ ToInternal(XYZ xyz, DisplayUnitType internalType)
        {
            var x = UnitUtils.ConvertFromInternalUnits(xyz.X, internalType);
            var y = UnitUtils.ConvertFromInternalUnits(xyz.Y, internalType);
            var z = UnitUtils.ConvertFromInternalUnits(xyz.Z, internalType);

            return new XYZ(x, y, z);
        }
#endif

        public void updateParameterValue(XYZ surveyCoordinates, FamilyInstance familyInstance, Autodesk.Revit.DB.Document document)
        {
            //Extrae los nombres de los parámetros de la instancia de la familia
            ParameterSet parameterSet = familyInstance.Parameters;

            //ITera sobre los parámetros
            foreach (Parameter parameter in parameterSet)
            {
                if(parameter.Definition.Name == "Coord_N" || parameter.Definition.Name == "Coord_E")
                {
                    using (Transaction trans = new Transaction(document))
                    {
                        trans.Start("Agregar valor al parámetro de familia");
                        if (parameter.Definition.Name == "Coord_N")
                        {
                            familyInstance.get_Parameter(parameter.Definition).Set(surveyCoordinates.Y);
                        }
                        else
                        {   
                            familyInstance.get_Parameter(parameter.Definition).Set(surveyCoordinates.X);
                        }
                        
                        trans.Commit();
                    }
                }
            }

        }
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            //OBtiene el documento de revit
            var doc = commandData.Application.ActiveUIDocument.Document;


            //Verifica si la vista está en 3D
            Boolean boolVista3D = VerificarVista3D(doc);
            if (boolVista3D == false)
            {
                TaskDialog.Show("Error", "La vista debe estar en 3D.");
                return Result.Failed;
            }

            //Obtiene los elementos visibles en la vista actual
            FilteredElementCollector collector = new FilteredElementCollector(doc, doc.ActiveView.Id);
            ICollection<Element> elementsInView = collector.ToElements();

            //Obtiene los elementos que sean Familias
            List<FamilyInstance> familyInstances = GetFamilyInstances(elementsInView);

            //Obtiene las instancias de Familias que su categoría sea "Structural Foundations" y que su nombre de familia contenga "PILE" o "PILOTE". Adicionalmente, que "Coord_E" y "Coord_N" como parámetros
            List<FamilyInstance> pilesInstances = GetPileInstances(familyInstances);
            if (pilesInstances.Count == 0)
            {
                TaskDialog.Show("Error", "No hay familias que cumplan con los requisitos: \n\n- Family category: Structural Foundations. \n\n- Family name: contains PILE or PILOTE.");
                return Result.Failed;
            }

            foreach (FamilyInstance familyInstance in pilesInstances)
            {

                //Ubicación del pilote dentro de las coordenadas locales de Revit
                LocationPoint location = familyInstance.Location as LocationPoint;
                XYZ relativePosition = location.Point;

                // Coordenadas  con respecto al survey point
                XYZ newCoordinates = ToSharedCoordinates(relativePosition, doc);

                //Actualizar coordenada
                updateParameterValue(newCoordinates, familyInstance, doc);

            }


            TaskDialog.Show("Éxito", "Se asignaron todas las coordenadas a los pilotes.");
            return Result.Succeeded;
        }
    }
}