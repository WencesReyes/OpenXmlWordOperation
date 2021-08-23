using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Logging;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System.IO;
using OpenXMLWordOperations.Models;

namespace OpenXMLWordOperations.Controllers
{
    [ApiController]
    [Route("[controller]")]
    public class WordController : ControllerBase
    {
        /// <summary>
        /// Codígo base para agregar rows a una tabla en word.
        /// </summary>
        /// <returns></returns>
        [HttpGet("AddingRowsInTable")]
        public IActionResult AddingRowsInTable()
        {
            var path = @"C:\Users\wence\Downloads\Cuestionario.docx";
            byte[] byteArray = System.IO.File.ReadAllBytes(path);

            var cuestionarios = Mock.GetCuestionariosMock();

            using (var stream = new MemoryStream())
            {
                stream.Write(byteArray, 0, byteArray.Length);
                using (var wordDoc = WordprocessingDocument.Open(stream, true))
                {
                    Body body = wordDoc.MainDocumentPart.Document.Body;
                    IEnumerable<Table> table = body.Descendants<Table>();
                    var bookmark = body.Descendants<BookmarkStart>().FirstOrDefault(bm => bm.Name.Equals("Miembros"));
                    var digForTable = bookmark.Parent;

                    while (!(digForTable is Table))
                    {
                        digForTable = digForTable.Parent;
                    }

                    //get rows
                    var rows = digForTable.Descendants<TableRow>().ToList();
                    //remember you have a header, so keep row 1, clone row 2 (our template for dynamic entry)
                    var myRow = (TableRow)rows.Last().Clone();
                    //remove it after cloning.
                    rows.Last().Remove();

                    foreach (var cuestionario in cuestionarios)
                    {
                        //clone our "reference row"
                        var rowToInsert = (TableRow)myRow.Clone();
                        //get list of cells
                        var listOfCellsInRow = rowToInsert.Descendants<TableCell>().ToList();
                        //just replace every bit of text in cells

                        listOfCellsInRow[0].Descendants<Text>().FirstOrDefault().Text = cuestionario.Nombre;
                        listOfCellsInRow[1].Descendants<Text>().FirstOrDefault().Text = cuestionario.Apellido;
                        listOfCellsInRow[2].Descendants<Text>().FirstOrDefault().Text = cuestionario.NumDocumento.ToString();
                        listOfCellsInRow[3].Descendants<Text>().FirstOrDefault().Text = cuestionario.Nacionalidad;
                        listOfCellsInRow[4].Descendants<Text>().FirstOrDefault().Text = cuestionario.Cargo;
                        listOfCellsInRow[5].Descendants<Text>().FirstOrDefault().Text = cuestionario.Correo;
                        listOfCellsInRow[6].Descendants<Text>().FirstOrDefault().Text = cuestionario.Telefono;
                        listOfCellsInRow[7].Descendants<Text>().FirstOrDefault().Text = cuestionario.Sexo;
                        listOfCellsInRow[8].Descendants<Text>().FirstOrDefault().Text = cuestionario.Domicilio;
                        listOfCellsInRow[9].Descendants<Text>().FirstOrDefault().Text = cuestionario.DOB.ToShortDateString();

                        digForTable.Descendants<TableRow>().Last().InsertAfterSelf(rowToInsert);
                    }

                    wordDoc.MainDocumentPart.Document.Save(); // won't update the original file
                }
                // Save the file with the new name
                stream.Position = 0;
                string filename = $"NuevoCuestionario.docx";
                return File(stream.ToArray(), "application/vnd.openxmlformats-officedocument.wordprocessingml.document", filename);
            }
        }

        /// <summary>
        /// Reemplaze un string por otra en todo el word.
        /// </summary>
        /// <returns></returns>
        [HttpGet("ReplaceText")]
        public IActionResult ReplaceText()
        {
            var path = @"D:\Usuario\Downloads\Carta de Reconocimiento (con Representación) para Compañías.docx";
            byte[] byteArray = System.IO.File.ReadAllBytes(path);
            var myObj = new { FullName = "Wenceslao Reyes Espinoza", DOB = "27/08/1997" };
            using (var stream = new MemoryStream())
            {
                stream.Write(byteArray, 0, byteArray.Length);
                using (var wordDoc = WordprocessingDocument.Open(stream, true))
                {
                    Body body = wordDoc.MainDocumentPart.Document.Body;
                    var old = body.InnerXml.ToString();
                    body.InnerXml = body.InnerXml.Replace("_PROVEEDOR", myObj.FullName);
                    wordDoc.MainDocumentPart.Document.Save(); // won't update the original file
                }
                // Save the file with the new name
                stream.Position = 0;
                string filename = $"{myObj.FullName}.docx";
                return File(stream.ToArray(), "application/vnd.openxmlformats-officedocument.wordprocessingml.document", filename);
            }
        }

        /// <summary>
        /// Codígo base para insertar una 'X' en un checkbox
        /// </summary>
        /// <returns></returns>
        [HttpGet("ModifyCheckbox")]
        public IActionResult ModifyCheckbox()
        {
            var path = @"C:\Users\wence\Downloads\Cuestionario.docx";
            byte[] byteArray = System.IO.File.ReadAllBytes(path);

            var myObj = new { FullName = "Wenceslao Reyes Espinoza", DOB = "27/08/1997" };
            using (var stream = new MemoryStream())
            {
                stream.Write(byteArray, 0, byteArray.Length);
                using (var wordDoc = WordprocessingDocument.Open(stream, true))
                {
                    Body body = wordDoc.MainDocumentPart.Document.Body;
                    var bookmark = body.Descendants<BookmarkStart>().FirstOrDefault(bm => bm.Name.Equals(true ? "SI_5_1" : "NO_5_1"));

                    var paragraphProperties = new ParagraphProperties { Justification = new Justification { Val = JustificationValues.Center } };
                    bookmark.Parent.Append(paragraphProperties);

                    var run = new Run();
                    var runProperties = new RunProperties();
                    runProperties.Bold = new Bold();

                    run.Append(runProperties);
                    run.Append(new Text("X"));

                    bookmark.Parent.Append(run);

                    wordDoc.MainDocumentPart.Document.Save(); // won't update the original file
                }
                // Save the file with the new name
                stream.Position = 0;
                string filename = "CheckboxCuestionario.docx";
                return File(stream.ToArray(), "application/vnd.openxmlformats-officedocument.wordprocessingml.document", filename);
            }
        }

        /// <summary>
        /// Codígo base para insertar un texto en un bookmark
        /// </summary>
        /// <returns></returns>
        [HttpGet("InsertTextBookMark")]
        public IActionResult InsertTextBookMark()
        {
            var path = @"C:\Users\wence\Downloads\Cuestionario.docx";
            byte[] byteArray = System.IO.File.ReadAllBytes(path);

            using (var stream = new MemoryStream())
            {
                stream.Write(byteArray, 0, byteArray.Length);
                using (var wordDoc = WordprocessingDocument.Open(stream, true))
                {
                    Body body = wordDoc.MainDocumentPart.Document.Body;
                    var bookmark = body.Descendants<BookmarkStart>().FirstOrDefault(bm => bm.Name.Equals("TELEFONO_INFO_GENERAL_PAIS"));

                    var paragraphProperties = new ParagraphProperties { Justification = new Justification { Val = JustificationValues.Left } };
                    bookmark.Parent.Append(paragraphProperties);

                    var run = new Run();
                    var runProperties = new RunProperties();
                    runProperties.Bold = new Bold();

                    run.Append(runProperties);
                    run.Append(new Text("3121589645"));

                    bookmark.Parent.Append(run);

                    wordDoc.MainDocumentPart.Document.Save(); // won't update the original file
                }
                // Save the file with the new name
                stream.Position = 0;
                string filename = "TextBookMarkCuestionario.docx";
                return File(stream.ToArray(), "application/vnd.openxmlformats-officedocument.wordprocessingml.document", filename);
            }
        }


        /// <summary>
        /// Codígo base para insertar un texto en frente de otro texto con un bookmark
        /// </summary>
        /// <returns></returns>
        [HttpGet("InsertTextAfterAnotherTextBookMark")]
        public IActionResult InsertTextAfterAnotherTextBookMark()
        {
            var path = @"C:\Users\wence\Downloads\Cuestionario.docx";
            byte[] byteArray = System.IO.File.ReadAllBytes(path);

            using (var stream = new MemoryStream())
            {
                stream.Write(byteArray, 0, byteArray.Length);
                using (var wordDoc = WordprocessingDocument.Open(stream, true))
                {
                    Body body = wordDoc.MainDocumentPart.Document.Body;
                    var bookmark = body.Descendants<BookmarkStart>().FirstOrDefault(bm => bm.Name.Equals("DOMICILIO_GENERAL_PAIS"));

                    var paragraphProperties = new ParagraphProperties { Justification = new Justification { Val = JustificationValues.Left } };
                    bookmark.Parent.Append(paragraphProperties);

                    var run = new Run();
                    var runProperties = new RunProperties();
                    runProperties.Bold = new Bold();

                    run.Append(runProperties);
                    run.Append(new Text("Chapulin #1435"));

                    bookmark.Parent.Append(run);

                    wordDoc.MainDocumentPart.Document.Save(); // won't update the original file
                }
                // Save the file with the new name
                stream.Position = 0;
                string filename = "TextAfterAnotherTextBookMarkCuestionario.docx";
                return File(stream.ToArray(), "application/vnd.openxmlformats-officedocument.wordprocessingml.document", filename);
            }
        }

        /// <summary>
        /// Codígo base para insertar rows en una tabla interna
        /// </summary>
        /// <returns></returns>
        [HttpGet("AddingRowsInInnerTable")]
        public IActionResult AddingRowsInInnerTable()
        {
            var path = @"C:\Users\wence\Downloads\Cuestionario.docx";
            byte[] byteArray = System.IO.File.ReadAllBytes(path);

            var cuestionarios = Mock.GetCuestionariosMock();

            using (var stream = new MemoryStream())
            {
                stream.Write(byteArray, 0, byteArray.Length);
                using (var wordDoc = WordprocessingDocument.Open(stream, true))
                {
                    Body body = wordDoc.MainDocumentPart.Document.Body;
                    IEnumerable<Table> table = body.Descendants<Table>();
                    var bookmark = body.Descendants<BookmarkStart>().FirstOrDefault(bm => bm.Name.Equals("PERSONAS_CONTACTO_PRIMER_CELDA"));
                    var digForRow = bookmark.Parent;

                    while (!(digForRow is TableRow))
                    {
                        digForRow = digForRow.Parent;
                    }

                    //remember you have a header, so keep row 1, clone row 2 (our template for dynamic entry)
                    var myRow = (TableRow)digForRow.Clone();

                    foreach (var cuestionario in cuestionarios)
                    {
                        //clone our "reference row"
                        var rowToInsert = (TableRow)myRow.Clone();
                        //get list of cells
                        var listOfCellsInRow = rowToInsert.Descendants<TableCell>().ToList();
                        //just replace every bit of text in cells

                        listOfCellsInRow[0].Descendants<Text>().FirstOrDefault().Text = cuestionario.Nombre;
                        listOfCellsInRow[1].Descendants<Text>().FirstOrDefault().Text = cuestionario.Apellido;
                        listOfCellsInRow[2].Descendants<Text>().FirstOrDefault().Text = cuestionario.NumDocumento.ToString();
                        listOfCellsInRow[3].Descendants<Text>().FirstOrDefault().Text = cuestionario.Nacionalidad;
                        listOfCellsInRow[4].Descendants<Text>().FirstOrDefault().Text = cuestionario.Cargo;
                        listOfCellsInRow[5].Descendants<Text>().FirstOrDefault().Text = cuestionario.Correo;
                        listOfCellsInRow[6].Descendants<Text>().FirstOrDefault().Text = cuestionario.Telefono;
                        listOfCellsInRow[7].Descendants<Text>().FirstOrDefault().Text = cuestionario.Sexo;
                        listOfCellsInRow[8].Descendants<Text>().FirstOrDefault().Text = cuestionario.Domicilio;
                        listOfCellsInRow[9].Descendants<Text>().FirstOrDefault().Text = cuestionario.DOB.ToShortDateString();

                        digForRow.InsertAfterSelf(rowToInsert);
                    }

                    digForRow.Remove();
                    wordDoc.MainDocumentPart.Document.Save(); // won't update the original file
                }
                // Save the file with the new name
                stream.Position = 0;
                string filename = $"EditInnerTable.docx";
                return File(stream.ToArray(), "application/vnd.openxmlformats-officedocument.wordprocessingml.document", filename);
            }
        }
    }
}
