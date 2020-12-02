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

namespace visioprueba
{
    public class Processor
    {
        public static List<string> GetVisioShapesFromFile(string fipath)
        {
            // lista de las shapes
            List<string> result;
            result = new List<string>();


            Dictionary<int, Shape> listanodes = new Dictionary<int, Shape>();
            Dictionary<int, Shape> relaciones = new Dictionary<int, Shape>();
            // Creamos y abrimos documento de visio
            if (System.IO.File.Exists(fipath))
            {
                Visio.Document visioDoc = new Visio.Application().Documents.Open(fipath);

                //recorremos las páginas del doc
                List<Shape> listaNodos = new List<Shape>();
                Double xOut;
                Double yOut;
                foreach (Visio.Page page in visioDoc.Pages)
                {
                    // se recorren los objetos
                    foreach (Visio.Shape shape in page.Shapes)
                     {
                         if (string.IsNullOrWhiteSpace(shape.Text) == false) // Si tiene contenido
                         {
                            result.Add(shape.Text);
                            Console.Write(result);
                         }
                         else
                         { // si es una relación
                            if (shape.Connects.Count > 0) {
                                Visio.Shape from = shape.Connects.Item16[1].ToSheet;
                            }
                             
                         }
                     }
                    GetNodesAndRelations(page, ref listanodes, ref relaciones);
                    break;
                }




                // Aqui escribe /*
             /* Visio.Page newpage = visioDoc.Pages.Add();
                Visio.Shape sourceShape = CreateState(newpage, "First ");
                Visio.Shape targetShape = CreateState(newpage, "Second ");

                Visio.Shape transition1 = CreateTransition(newpage, sourceShape, targetShape); */


            }
            return result;
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
