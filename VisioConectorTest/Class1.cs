using Microsoft.Office.Interop.Visio;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.XPath;
using Visio = Microsoft.Office.Interop.Visio;
using SRL.Mappings.Enginering;
using SRL.Mappings.Enginering.PhysicalLevel;
using SRL.ResourceShape;
using Cake.Engine;
namespace visioprueba
{
    public class Processor
    {
                 
        public static List<PhysicalModel> GetVisioShapesFromFile(string fipath)
        {

            List<PhysicalModel> pagesModels = new List<PhysicalModel>();
            Dictionary<int, PhysicalComponent> allComponents = new Dictionary<int, PhysicalComponent>();
            List<Shape> relations = new List<Shape>();
            // Creamos y abrimos documento de visio
            if (System.IO.File.Exists(fipath))
            {
                Visio.Document visioDoc = new Visio.Application().Documents.Open(fipath);

                foreach (Visio.Page page in visioDoc.Pages)
                {
                    PhysicalModel physicalModel = new PhysicalModel(page.Name, page.NameU, page.ID.ToString());    // añadir codigo identidficador de la pagina                                                                                                                                 
            
                    GetNodesAndRelations(page,ref physicalModel, ref allComponents, ref relations);
                   
                        foreach (Visio.Shape relation in relations) {
                        PhysicalComponent shapePC;
                        PhysicalComponent physicalComponentFrom = null;
                        PhysicalComponent physicalComponentTo = null;

                        string aux = string.Empty; 

                        shapePC = new PhysicalComponent(relation.Name, relation.Text, relation.ID.ToString(), physicalModel);
                       
                        Visio.Shape transitionFrom = relation.Connects.Item16[1].ToSheet;
                        Visio.Shape transitionTo = relation.Connects.Item16[2].ToSheet;
                        
                        if (transitionFrom != null && transitionTo != null) {
                            bool existsFrom = allComponents.TryGetValue(transitionFrom.ID, out physicalComponentFrom);
                            bool existsTo = allComponents.TryGetValue(transitionTo.ID, out physicalComponentTo);
                            if (existsFrom && existsTo) {
                                 physicalModel.ADD_ConnectionBetweenPhysicalComponents(relation.Name, physicalComponentFrom, physicalComponentTo,  ref aux);
                                /* por qué usar este y no ADD_ConnectionBetweenPhysicalComponents */
                            }
                        }

                    }
                    Dictionary<string, Relationship> allRelationships = physicalModel.GET_AllConnectionBetweenPhysicalComponentsAsSRL();
                    Console.Write(allRelationships.Keys.ToList().Count);
                    pagesModels.Add(physicalModel);
                    break;
                }
             }
             return pagesModels;
        }

        #region Lectura
        private static void  GetNodesAndRelations(Visio.Page visPage, ref PhysicalModel physicalModel, ref Dictionary<int,PhysicalComponent> nodes, ref List<Shape> relationship) {

            if (visPage != null && visPage.Shapes.Count > 0) {
                foreach (Visio.Shape shape in visPage.Shapes) {
                    // validamos que el objeto es un nodo, cuando valor de OneD=0
                    // si es relacion es OneD=-1
                    /* https://docs.microsoft.com/es-es/office/vba/api/overview/visio/object-model */
                    if (shape.OneD == 0 && shape.Type != (short)Visio.VisShapeTypes.visTypeForeignObject) { // info embebida
                        Console.WriteLine(shape.Name + ": " + shape.ID.ToString());                                                                                   //se insertan los nodos en la colección
                        PhysicalComponent shapePC;
                        shapePC = new PhysicalComponent(shape.Name, shape.Text, shape.ID.ToString(), physicalModel);
                        shapePC.ADD_Metadata("lastModificationDate", typeof(string), DateTime.Now); // Le podemos añadir propiedades
                        
                        // shapePC.ADD_Metadata("type", typeof(string),); // Le podemos añadir propiedades
                        // Transformaciones de interoperabilidad, herramienta que genera la indezacion, INT de interoperabilidad
                        shapePC.TYPE_SourceTool = Cake.Engine.Enums.Grammaticals.INT_Visio;
                        shapePC.TYPE_InSource = shape.NameU; // NameU es la tipologia que le pone a visio en sus componentes
                        // shape.TYPE_LibararyPath_InSource = "" // por si queremos usar una librería ya existente
                        nodes.Add(shape.ID, shapePC);

                    } else if(shape.OneD == -1 || shape.Connects.Count > 0) {
                        relationship.Add(shape);
                    }
                 }
            }
        }
        #endregion Lectura

        #region Escritura

        public static void ParseToVisio(string fipath, List<PhysicalModel> pagesModels)
        {
            if (System.IO.File.Exists(fipath))
            {
                // Abrimos doucmento
                Visio.Document visioDoc = new Visio.Application().Documents.Open(fipath);
                Dictionary<string, Relationship> allRelationships;
                Dictionary<Int32, Shape> shapesProcessed = new Dictionary<int, Shape>();
                foreach (PhysicalModel page in pagesModels)
                {
                    Visio.Page newpage = visioDoc.Pages.Add();
                    allRelationships = page.GET_AllConnectionBetweenPhysicalComponentsAsSRL();
                    Console.Write(allRelationships.Keys.ToList().Count);
                    /* Otra opcion puede ser   public Dictionary<string, MappeableElement> GET_MyContainers_MEs(MappeableElement root_ME); */
                    /* Por cada relación iteramos y parseamos */
                    foreach (var item in allRelationships)
                    {
                        Console.WriteLine("HOLA");
                        Relationship rel = item.Value;
                        string relationID = item.Key;
                        Console.WriteLine(item.Key);
                        SRL.ResourceShape.Artifact artifactTo = rel.To;
                        SRL.ResourceShape.Artifact artifactFrom = rel.From;
                        Console.WriteLine("+++++"+artifactTo.Name + ": " + artifactTo.Identifier);
                        Visio.Shape sourceShape, targetShape;
                        bool keyExists = shapesProcessed.ContainsKey(Int32.Parse(artifactTo.Identifier));
                        if (keyExists)
                        {
                            sourceShape = shapesProcessed[Int32.Parse(artifactTo.Identifier)];
                        }
                        else
                        {
                            sourceShape = CreateState(newpage, artifactTo.Name, artifactTo.Description);
                            shapesProcessed.Add(Int32.Parse(artifactTo.Identifier), sourceShape);
                        }
                         keyExists = shapesProcessed.ContainsKey(Int32.Parse(artifactFrom.Identifier));
                        if (keyExists)
                        {
                            targetShape = shapesProcessed[Int32.Parse(artifactFrom.Identifier)];
                        }
                        else
                        {
                            targetShape = CreateState(newpage, artifactFrom.Name, artifactFrom.Description);
                            shapesProcessed.Add(Int32.Parse(artifactFrom.Identifier), targetShape);
                        }
                       
                        /* Se escribe en visio*/
                        Visio.Shape transition1 = CreateTransition(newpage, sourceShape, targetShape);

                    }
                }
            }
        }

        private static Visio.Shape CreateState(Visio.Page page, string type, string text)
        {
            Visio.Shape result = null;
            if (page != null) {
                // se crea el objeto, accediendo al namespace Masters donde están los tipos de los objetos en las paletas de Visio
                string typeFormatted = "";
                if (type == "Estado inicial") {
                    typeFormatted = "Initial state";
                }
                else if (type == "Estado final"){
                    typeFormatted = "Final state";
                }
                else if (type == "Estado") {
                    typeFormatted = "State";
                }
                else {
                    typeFormatted = type;
                }
                result = page.Drop(page.Application.ActiveDocument.Masters.ItemU[typeFormatted], 0, 0);
            }
            result.Text = text;
          
            return result;
        }
        private static Visio.Shape CreateTransition(Visio.Page page, Visio.Shape sourceShape, Visio.Shape targetShape)
        {
            Visio.Shape transition = null;
            if (page != null && sourceShape != null && targetShape != null)
            {
                // se crea el shape de relacion para luego asignar source y target
                transition = page.Drop(page.Application.ConnectorToolDataObject, 0, 0);
                // se asigna source
                transition.get_CellsU("BeginX").GlueTo(sourceShape.get_CellsU("PinX"));
                // se asigna el target
                transition.get_CellsU("EndX").GlueTo(targetShape.get_CellsU("PinX"));


            }
            return transition;
        }

        #endregion Escritura

        public static List<PhysicalModel> GetVisioShapesFromFileOld(string fipath)
        {

            List<PhysicalModel> pagesModels = new List<PhysicalModel>();
            Dictionary<int, PhysicalComponent> allComponents = new Dictionary<int, PhysicalComponent>();
            List<Shape> relations = new List<Shape>();
            // Creamos y abrimos documento de visio
            if (System.IO.File.Exists(fipath))
            {
                Visio.Document visioDoc = new Visio.Application().Documents.Open(fipath);


                foreach (Visio.Page page in visioDoc.Pages)

                {
                    PhysicalModel physicalModel = new PhysicalModel(page.Name, page.NameU, page.ID.ToString());    // añadir codigo identidficador de la pagina
                    // se recorren los objetos
                    foreach (Visio.Shape shape in page.Shapes)
                    {


                        if (string.IsNullOrWhiteSpace(shape.Text) == false) // Si tiene contenido
                        {
                            PhysicalComponent shapePC;
                            shapePC = new PhysicalComponent(shape.Name, shape.Text, shape.ID.ToString(), physicalModel);
                            shapePC.ADD_Metadata("lastModificationDate", typeof(string), DateTime.Now); // Le podemos añadir propiedades
                            /* result.Add(shape.Text);
                             Console.Write(result);*/
                            // Transformaciones de interoperabilidad, herramienta que genera la indezacion, INT de interoperabilidad
                            shapePC.TYPE_SourceTool = Cake.Engine.Enums.Grammaticals.INT_Visio;
                            shapePC.TYPE_InSource = shape.NameU; // NameU es la tipologia que le pone a visio en sus componentes
                                                                 // shape.TYPE_LibararyPath_InSource = "" // por si queremos usar una librería ya existente
                            allComponents.Add(shape.ID, shapePC);
                        }
                        else
                        { // si es una relación
                            if (shape.Connects.Count > 0)
                            {

                                relations.Add(shape);
                            }

                        }
                    }
                    foreach (Visio.Shape relation in relations)
                    {
                        PhysicalComponent shapePC;
                        PhysicalComponent physicalComponentFrom = null;
                        PhysicalComponent physicalComponentTo = null;

                        string aux = string.Empty;

                        shapePC = new PhysicalComponent(relation.Name, relation.Text, relation.ID.ToString(), physicalModel);
                        Visio.Shape transitionFrom = relation.Connects.Item16[1].ToSheet;
                        Visio.Shape transitionTo = relation.Connects.Item16[2].ToSheet;

                        if (transitionFrom != null && transitionTo != null)
                        {
                            bool existsFrom = allComponents.TryGetValue(transitionFrom.ID, out physicalComponentFrom);
                            bool existsTo = allComponents.TryGetValue(transitionTo.ID, out physicalComponentTo);


                            if (existsFrom && existsTo)
                            {
                                physicalModel.ADD_Relationship(relation.Name, physicalComponentFrom, true, physicalComponentTo, true, nameof(Cake.Engine.Enums.Grammaticals.Association), (int)Cake.Engine.Enums.Grammaticals.Association, true, ref aux);
                                /* por qué usar este y no ADD_ConnectionBetweenPhysicalComponents */
                            }
                        }

                    }
                    //GetNodesAndRelations(page, ref listanodes, ref relaciones);
                    pagesModels.Add(physicalModel);
                    break;
                }


            }
            return pagesModels;
        }


    }
}
