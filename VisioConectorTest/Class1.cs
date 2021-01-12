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
            Dictionary<string, PhysicalComponent> allComponents = new Dictionary<string, PhysicalComponent>();
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
                        // PhysicalComponent shapePC;
                        PhysicalComponent physicalComponentFrom = null;
                        PhysicalComponent physicalComponentTo = null;

                        string aux = string.Empty; 

                       // shapePC = new PhysicalComponent(relation.Name, relation.Text, relation.ID.ToString(), physicalModel);
                       
                        Visio.Shape transitionFrom = relation.Connects.Item16[1].ToSheet;
                        Visio.Shape transitionTo = relation.Connects.Item16[2].ToSheet;
                        
                        if (transitionFrom != null && transitionTo != null) {
                            bool existsFrom = allComponents.TryGetValue(page.ID + ":" + transitionFrom.ID, out physicalComponentFrom);
                            bool existsTo = allComponents.TryGetValue(page.ID + ":" + transitionTo.ID, out physicalComponentTo);
                            if (existsFrom && existsTo) {
                                 physicalModel.ADD_ConnectionBetweenPhysicalComponents(relation.Text, physicalComponentFrom, physicalComponentTo, ref aux);
                                
                                /* por qué usar este y no ADD_ConnectionBetweenPhysicalComponents */
                            }
                        }
                    }
                    Dictionary<string, Relationship> allRelationships = physicalModel.GET_AllConnectionBetweenPhysicalComponentsAsSRL();
                    pagesModels.Add(physicalModel);
                    
                }
             }
             return pagesModels;
        }

        #region Lectura
        private static void  GetNodesAndRelations(Visio.Page visPage, ref PhysicalModel physicalModel, ref Dictionary<string,PhysicalComponent> nodes, ref List<Shape> relationship) {

            if (visPage != null && visPage.Shapes.Count > 0) {
                foreach (Visio.Shape shape in visPage.Shapes) {
                    // validamos que el objeto es un nodo, cuando valor de OneD=0
                    // si es relacion es OneD=-1
                    /* https://docs.microsoft.com/es-es/office/vba/api/overview/visio/object-model */
                    if (shape.OneD == 0 && shape.Type != (short)Visio.VisShapeTypes.visTypeForeignObject) { // info embebida
                        //se insertan los nodos en la colección
                        PhysicalComponent shapePC;
                        shapePC = new PhysicalComponent(shape.Name, shape.Text, shape.ID.ToString(), physicalModel);

                        if (shape.Name == "State" || shape.Name == "Estado") { // Si es state guardamos el título
                            Visio.Shape titleShape = visPage.Shapes.ItemFromID[shape.ID + 1];
                            shapePC.ADD_Metadata("SimpleStateTitle", typeof(string), titleShape.Text);
                           
                        }
                        // Transformaciones de interoperabilidad, herramienta que genera la indezacion, INT de interoperabilidad
                        shapePC.TYPE_SourceTool = Cake.Engine.Enums.Grammaticals.INT_Visio;
                        shapePC.TYPE_InSource = shape.NameU; // NameU es la tipologia que le pone a visio en sus componentes
                                                             // shape.TYPE_LibararyPath_InSource = "" // por si queremos usar una librería ya existente
                       
                        nodes.Add(visPage.ID + ":" + shape.ID, shapePC);

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
            Visio.Document visioDoc;
            if (!System.IO.File.Exists(fipath)) {
                string path = System.IO.Directory.GetCurrentDirectory();
                string result = path.Replace(@"visioConector\bin\Debug", "default.vsdx");
                visioDoc = new Visio.Application().Documents.Add(result);

            }
            else
            {
                // Abrimos doucmento
                visioDoc = new Visio.Application().Documents.Open(fipath);
            }
                Dictionary<string, Relationship> allRelationships;
                Dictionary<Int32, Shape> shapesProcessed = new Dictionary<int, Shape>();
                foreach (PhysicalModel page in pagesModels)
                {
                    Visio.Page newpage = visioDoc.Pages.Add();
                    allRelationships = page.GET_AllConnectionBetweenPhysicalComponentsAsSRL();

                    /* Otra opcion puede ser   public Dictionary<string, MappeableElement> GET_MyContainers_MEs(MappeableElement root_ME); */
                    /* Por cada relación iteramos y parseamos */
                    Int32 counter = 0; // PARA LA POSICION
                    foreach (var item in allRelationships)   {
                        Relationship rel = item.Value;
                                               
                        SRL.ResourceShape.Artifact artifactTo = rel.To;
                        SRL.ResourceShape.Artifact artifactFrom = rel.From;
                        Visio.Shape sourceShape, targetShape;

                        // procesamos la relación cpgemos shapeTo
                        bool keyExists = shapesProcessed.ContainsKey(Int32.Parse(artifactTo.Identifier));
                        if (keyExists) // Si ya se ha procesado la shape
                        {
                            sourceShape = shapesProcessed[Int32.Parse(artifactTo.Identifier)];
                        } else // Si aún no la hemos procesado
                        {
                            if (artifactTo.Name == "State" || artifactTo.Name == "Estado") {
                                sourceShape = CreateState(newpage, artifactTo.Name, artifactTo.Description, null, counter);

                            } else
                            {
                                sourceShape = CreateState(newpage, artifactTo.Name, artifactTo.Description, null, counter);
                            }
                            counter++;
                            shapesProcessed.Add(Int32.Parse(artifactTo.Identifier), sourceShape);
                        }



                         keyExists = shapesProcessed.ContainsKey(Int32.Parse(artifactFrom.Identifier));
                        if (keyExists)
                        {
                            targetShape = shapesProcessed[Int32.Parse(artifactFrom.Identifier)];
                        }
                        else
                        {
                            if (artifactFrom.Name == "State" || artifactFrom.Name == "Estado")
                            {
                               // MetaData existKey = artifactFrom.GetMetaDataByKey("SimpleStateTitle");
                                targetShape = CreateState(newpage, artifactFrom.Name, artifactFrom.Description, "", counter);

                            }
                            else
                            {
                                targetShape = CreateState(newpage, artifactFrom.Name, artifactFrom.Description, null, counter) ;
                               
                            }
                        shapesProcessed.Add(Int32.Parse(artifactFrom.Identifier), targetShape);
                        counter++;
                        }
                    /* Se escribe en visio*/
                        Visio.Shape transition1 = CreateTransition(newpage, sourceShape, targetShape, rel.Name);
                        newpage.SetTheme("Office Theme");
                        newpage.Layout();
                      

                    }
                    visioDoc.Application.ActiveDocument.SaveAs(fipath);
            }

            
        }

        private static Visio.Shape CreateState(Visio.Page page, string type, string text, string title, float position)
        {
            Visio.Shape result = null;
            if (page != null) {
                // se crea el objeto, accediendo al namespace Masters donde están los tipos de los objetos en las paletas de Visio
                string typeFormatted = "";
                if (type == "Estado inicial")
                {
                    typeFormatted = "Initial state";
                }
                else if (type == "Estado final")
                {
                    typeFormatted = "Final state";
                }
                else if (type == "Estado" || type == "State")
                {   
                    typeFormatted = "State";
                    
                }
                else if ( type == "Final state" || type == "Initial state")
                {
                    typeFormatted = type;
                }
                else {
                    typeFormatted = "State";
                }

                result = page.Drop(page.Application.ActiveDocument.Masters.ItemU[typeFormatted], position, position);
                result.Text = text;
                if (title != null) { 
                    Visio.Shape titleShape = page.Shapes.ItemFromID[result.ID + 1];
                    titleShape.Text = title;
                }

            }
           
          
            return result;
        }
        private static Visio.Shape CreateTransition(Visio.Page page, Visio.Shape sourceShape, Visio.Shape targetShape, string text)
        {
            Visio.Shape transition = null;
            if (page != null && sourceShape != null && targetShape != null)
            {
                // se crea el shape de relacion para luego asignar source y target
                transition = page.Drop(page.Application.ConnectorToolDataObject, 0, 0);
                // se asigna source
                transition.get_CellsU("BeginX").GlueTo(targetShape.get_CellsU("PinX"));
                // se asigna el target
                transition.get_CellsU("EndX").GlueTo(sourceShape.get_CellsU("PinX"));
                transition.CellsSRC[0, 0, 0].FormulaU =("13"); // intento de la flechita 
                // sourceShape.AutoConnect(targetShape, Visio.VisAutoConnectDir.visAutoConnectDirRight);
                //Assuming 'No theme' is set for the page, no arrow will 
                //be shown so change theme to see connector arrow
                
                transition.Text = text;
            

            }
            return transition;
        }

        #endregion Escritura

    }
}
