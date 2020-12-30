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
                            if (shape.Connects.Count > 0) {
                               /* Console.WriteLine("Número de relaciones" + shape.Connects.Count);
                                Visio.Shape transitionFrom = shape.Connects.Item16[1].ToSheet;
                                Visio.Shape transitionTo = shape.Connects.Item16[2].ToSheet;
                                shapePC.ADD_Metadata("from", typeof(int), transitionFrom.ID);
                               

                                shapePC.ADD_Metadata("to", typeof(int), transitionTo.ID);
                                Dictionary<string, Type> aux = shapePC.GET_Metadata();*/
                             
                                //shapePC.ADD_Metadata("from", typeof(int), DateTime.Now); // Le podemos añadir propiedades

                                relations.Add(shape);
                            }
                             
                         }
                     }
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
                                 physicalModel.ADD_Relationship(relation.Name, physicalComponentFrom, true, physicalComponentTo, true, nameof(Cake.Engine.Enums.Grammaticals.Association), (int)Cake.Engine.Enums.Grammaticals.Association, true, ref aux);
                            }
                        }

                    }

                    //GetNodesAndRelations(page, ref listanodes, ref relaciones);
                    pagesModels.Add(physicalModel);
                    break;
                }
                
                




                // Aqui escribe /*
              /*  Visio.Page newpage = visioDoc.Pages.Add();
                Visio.Shape sourceShape = CreateState(newpage, "First ");
                Visio.Shape targetShape = CreateState(newpage, "Second ");

                Visio.Shape transition1 = CreateTransition(newpage, sourceShape, targetShape); */


            }
            return pagesModels;
        }

        #region Lectura
        private static void  GetNodesAndRelations(Visio.Page visPage, ref Dictionary<int,Shape> nodes, ref Dictionary<int,Shape> relationship) {
          
      
            if (visPage != null && visPage.Shapes.Count > 0) {
                foreach (Visio.Shape shape in visPage.Shapes) {
                    // validamos que el objeto es un nodo, cuando valor de OneD=0
                    // si es relacion es OneD=-1
                    /* https://docs.microsoft.com/es-es/office/vba/api/overview/visio/object-model */
                    if (shape.OneD == 0 && shape.Type != (short)Visio.VisShapeTypes.visTypeForeignObject) { // info embebida
                        //se insertan los nodos en la colección
                        nodes.Add(shape.ID, shape);
                        // Owner, dentro de qué elmento estoy contenido
                        
                    } else if(shape.OneD == -1) {
                        relationship.Add(shape.ID, shape);
                    }
                 }
            }
        }
        #endregion Lectura

        #region Escritura
        private static Visio.Shape CreateState(Visio.Page page, string name)
        {
            Visio.Shape result = null;
            if (page != null) {
                // se crea el objeto, accediendo al namespace Masters donde están los tipos de los objetos en las paletas de Visio
                result = page.Drop(page.Application.ActiveDocument.Masters.ItemU["State"], 0, 0);
            }
            result.Text = name;
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

    }
}
