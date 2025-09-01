using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

using BottomBorder = DocumentFormat.OpenXml.Wordprocessing.BottomBorder;
using Break = DocumentFormat.OpenXml.Wordprocessing.Break;
using Document = DocumentFormat.OpenXml.Wordprocessing.Document;
using LeftBorder = DocumentFormat.OpenXml.Wordprocessing.LeftBorder;
using Paragraph = DocumentFormat.OpenXml.Wordprocessing.Paragraph;
using ParagraphProperties = DocumentFormat.OpenXml.Wordprocessing.ParagraphProperties;
using RightBorder = DocumentFormat.OpenXml.Wordprocessing.RightBorder;
using Run = DocumentFormat.OpenXml.Wordprocessing.Run;
using RunProperties = DocumentFormat.OpenXml.Wordprocessing.RunProperties;
using Table = DocumentFormat.OpenXml.Wordprocessing.Table;
using TableCell = DocumentFormat.OpenXml.Wordprocessing.TableCell;
using TableCellProperties = DocumentFormat.OpenXml.Wordprocessing.TableCellProperties;
using TableProperties = DocumentFormat.OpenXml.Wordprocessing.TableProperties;
using TableRow = DocumentFormat.OpenXml.Wordprocessing.TableRow;
using TableStyle = DocumentFormat.OpenXml.Wordprocessing.TableStyle;
using Text = DocumentFormat.OpenXml.Wordprocessing.Text;
using TopBorder = DocumentFormat.OpenXml.Wordprocessing.TopBorder;

namespace ChequeBancarioOpenXML
{
    public class ChequeBancario
    {
        public string? NumeroCheque { get; set; }
        public string? TitularCuenta { get; set; }
        public string? DireccionTitular { get; set; }
        public string? Beneficiario { get; set; }
        public decimal Importe { get; set; }
        public string? ImporteEnPalabras { get; set; }
        public string? LugarEmision { get; set; }
        public DateTime FechaEmision { get; set; }
        public string? NombreBanco { get; set; }
        public string? Sucursal { get; set; }
        public string? NumeroCuenta { get; set; }
        public string? ReferenciaTransferencia { get; set; }

        public ChequeBancario()
        {
            FechaEmision = DateTime.Now;
        }

        public void GenerarCheque(string rutaDocumento)
        {
            // Eliminar el archivo si existe para evitar problemas de acceso
            if (File.Exists(rutaDocumento))
            {
                try
                {
                    File.Delete(rutaDocumento);
                }
                catch (IOException)
                {
                    throw new Exception("No se puede acceder al archivo. Puede que esté en uso por otro programa.");
                }
            }

            try
            {
                using (WordprocessingDocument wordDocument = WordprocessingDocument.Create(
                    rutaDocumento, WordprocessingDocumentType.Document))
                {
                    // Agregar declaración XML con encoding UTF-8
                    wordDocument.AddMainDocumentPart();
                    wordDocument.MainDocumentPart!.Document = new Document();

                    // Configurar encoding UTF-8 explícitamente
                    wordDocument.MainDocumentPart.Document.Save();

                    var body = new Body();

                    // Configurar márgenes y tamaño de página
                    SectionProperties sectionProps = new SectionProperties();
                    PageMargin pageMargin = new PageMargin()
                    {
                        Top = 720,
                        Right = 720,
                        Bottom = 720,
                        Left = 720,
                        Header = 360,
                        Footer = 360
                    };
                    sectionProps.Append(pageMargin);

                    // Agregar contenido al cheque
                    body.Append(CrearEncabezado());
                    body.Append(CrearSeccionDatosCuenta());
                    body.Append(CrearSeccionImporte());
                    body.Append(CrearSeccionBeneficiario());
                    body.Append(CrearSeccionFirmas());

                    // Añadir la sección al cuerpo
                    body.Append(sectionProps);
                    wordDocument.MainDocumentPart.Document.Append(body);

                    // Guardar con configuración UTF-8
                    wordDocument.MainDocumentPart.Document.Save();
                }
            }
            catch (Exception ex)
            {
                throw new Exception($"Error al crear el documento: {ex.Message}");
            }
        }

        private Paragraph CrearEncabezado()
        {
            Paragraph paragraph = new Paragraph();
            ParagraphProperties paragraphProperties = new ParagraphProperties();
            paragraphProperties.Append(new Justification() { Val = JustificationValues.Center });

            Run run = new Run();
            RunProperties runProperties = new RunProperties();
            runProperties.Append(new Bold());
            runProperties.Append(new FontSize() { Val = "28" });

            // Configurar fuente compatible con Unicode
            runProperties.Append(new RunFonts()
            {
                Ascii = "Arial",
                HighAnsi = "Arial",
                ComplexScript = "Arial"
            });

            run.RunProperties = runProperties;
            run.Append(new Text(NombreBanco!));

            paragraph.Append(paragraphProperties);
            paragraph.Append(run);
            paragraph.Append(new Run(new Break()));

            Run runSucursal = new Run();
            RunProperties runPropsSucursal = new RunProperties();
            runPropsSucursal.Append(new RunFonts()
            {
                Ascii = "Arial",
                HighAnsi = "Arial",
                ComplexScript = "Arial"
            });
            runSucursal.RunProperties = runPropsSucursal;
            runSucursal.Append(new Text($"Sucursal: {Sucursal}"));

            paragraph.Append(runSucursal);

            return paragraph;
        }

        private Table CrearSeccionDatosCuenta()
        {
            Table table = new Table();
            TableProperties tableProperties = new TableProperties();
            TableStyle tableStyle = new TableStyle() { Val = "TablaNormal" };
            TableWidth tableWidth = new TableWidth() { Width = "5000", Type = TableWidthUnitValues.Pct };
            TableBorders tableBorders = new TableBorders();
            tableBorders.Append(new TopBorder() { Val = BorderValues.Single, Size = 4 });
            tableBorders.Append(new BottomBorder() { Val = BorderValues.Single, Size = 4 });
            tableBorders.Append(new LeftBorder() { Val = BorderValues.Single, Size = 4 });
            tableBorders.Append(new RightBorder() { Val = BorderValues.Single, Size = 4 });

            tableProperties.Append(tableStyle, tableWidth, tableBorders);
            table.Append(tableProperties);

            // Fila 1: Número de cuenta
            TableRow row1 = new TableRow();
            row1.Append(CrearCeldaConTexto("Número de Cuenta:", true));
            row1.Append(CrearCeldaConTexto(NumeroCuenta!, false, 3));
            table.Append(row1);

            // Fila 2: Titular
            TableRow row2 = new TableRow();
            row2.Append(CrearCeldaConTexto("Titular:", true));
            row2.Append(CrearCeldaConTexto(TitularCuenta!, false, 3));
            table.Append(row2);

            // Fila 3: Dirección
            TableRow row3 = new TableRow();
            row3.Append(CrearCeldaConTexto("Dirección:", true));
            row3.Append(CrearCeldaConTexto(DireccionTitular!, false, 3));
            table.Append(row3);

            // Fila 4: Referencia de transferencia
            TableRow row4 = new TableRow();
            row4.Append(CrearCeldaConTexto("Ref. Transferencia:", true));
            row4.Append(CrearCeldaConTexto(ReferenciaTransferencia ?? "N/A", false, 3));
            table.Append(row4);

            return table;
        }

        private Table CrearSeccionImporte()
        {
            Table table = new Table();
            TableProperties tableProperties = new TableProperties();
            TableWidth tableWidth = new TableWidth() { Width = "5000", Type = TableWidthUnitValues.Pct };
            TableBorders tableBorders = new TableBorders();
            tableBorders.Append(new TopBorder() { Val = BorderValues.Single, Size = 4 });
            tableBorders.Append(new BottomBorder() { Val = BorderValues.Single, Size = 4 });

            tableProperties.Append(tableWidth, tableBorders);
            table.Append(tableProperties);

            // Fila 1: Importe en números
            TableRow row1 = new TableRow();
            row1.Append(CrearCeldaConTexto("Importe:", true));
            row1.Append(CrearCeldaConTexto($"{Importe:C}", false, 3));
            table.Append(row1);

            // Fila 2: Importe en palabras
            TableRow row2 = new TableRow();
            row2.Append(CrearCeldaConTexto("En palabras:", true));
            row2.Append(CrearCeldaConTexto(ImporteEnPalabras + " EUROS", false, 3));
            table.Append(row2);

            return table;
        }

        private Table CrearSeccionBeneficiario()
        {
            Table table = new Table();
            TableProperties tableProperties = new TableProperties();
            TableWidth tableWidth = new TableWidth() { Width = "5000", Type = TableWidthUnitValues.Pct };
            TableBorders tableBorders = new TableBorders();
            tableBorders.Append(new TopBorder() { Val = BorderValues.Single, Size = 4 });
            tableBorders.Append(new BottomBorder() { Val = BorderValues.Single, Size = 4 });

            tableProperties.Append(tableWidth, tableBorders);
            table.Append(tableProperties);

            // Fila 1: Beneficiario
            TableRow row1 = new TableRow();
            row1.Append(CrearCeldaConTexto("Páguese a:", true));
            row1.Append(CrearCeldaConTexto(Beneficiario!, false, 3));
            table.Append(row1);

            return table;
        }

        private Table CrearSeccionFirmas()
        {
            Table table = new Table();
            TableProperties tableProperties = new TableProperties();
            TableWidth tableWidth = new TableWidth() { Width = "5000", Type = TableWidthUnitValues.Pct };

            tableProperties.Append(tableWidth);
            table.Append(tableProperties);

            // Fila con lugar, fecha y firma
            TableRow row = new TableRow();

            // Lugar y fecha
            TableCell lugarFechaCell = new TableCell();
            Paragraph lugarFechaParagraph = new Paragraph();
            lugarFechaParagraph.Append(new Run(new Text($"{LugarEmision}, {FechaEmision:dd/MM/yyyy}")));
            lugarFechaCell.Append(lugarFechaParagraph);

            // Firma (espacio en blanco)
            TableCell firmaCell = new TableCell();
            Paragraph firmaParagraph = new Paragraph();
            firmaParagraph.Append(new Run(new Text("Firma:")));
            firmaParagraph.Append(new Run(new Break()));

            // Línea para firma
            ParagraphProperties paragraphProperties = new ParagraphProperties();
            paragraphProperties.Append(new BottomBorder()
            {
                Val = BorderValues.Single,
                Size = 2,
                Color = "000000",
                Space = 0
            });
            firmaParagraph.PrependChild(paragraphProperties);

            firmaCell.Append(firmaParagraph);

            row.Append(lugarFechaCell);
            row.Append(firmaCell);
            table.Append(row);

            return table;
        }

        private static TableCell CrearCeldaConTexto(string texto, bool negrita = false, int anchoColumna = 1)
        {
            TableCell cell = new TableCell();
            TableCellProperties cellProperties = new TableCellProperties();
            cellProperties.Append(new TableCellWidth() { Type = TableWidthUnitValues.Auto });
            cell.TableCellProperties = cellProperties;

            Paragraph paragraph = new Paragraph();
            Run run = new Run();

            // Configurar fuente compatible con Unicode
            RunProperties runProperties = new RunProperties();
            runProperties.Append(new RunFonts()
            {
                Ascii = "Arial",
                HighAnsi = "Arial",
                ComplexScript = "Arial"
            });

            if (negrita)
            {
                runProperties.Append(new Bold());
            }

            run.RunProperties = runProperties;
            run.Append(new Text(texto));
            paragraph.Append(run);
            cell.Append(paragraph);

            return cell;
        }
    }
}