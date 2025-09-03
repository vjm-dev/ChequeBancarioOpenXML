using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

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
                BarraProgreso progreso = new BarraProgreso();

                using (WordprocessingDocument wordDocument = WordprocessingDocument.Create(
                    rutaDocumento, WordprocessingDocumentType.Document))
                {
                    progreso.Report(0.1, "Generando documento...");
                    Thread.Sleep(100); // Pequeña pausa para visualización

                    // Agregar la parte principal del documento
                    MainDocumentPart mainPart = wordDocument.AddMainDocumentPart();
                    mainPart.Document = new Document(
                        new Body(
                            new SectionProperties(
                                new PageMargin()
                                {
                                    Top = 720,
                                    Right = 720,
                                    Bottom = 720,
                                    Left = 720,
                                    Header = 360,
                                    Footer = 360
                                }
                            )
                        )
                    );

                    progreso.Report(0.3, "Agregando encabezado...");
                    Thread.Sleep(100);
                    mainPart.Document.Body!.Append(CrearEncabezado());

                    progreso.Report(0.4, "Agregando datos de cuenta...");
                    Thread.Sleep(100);
                    mainPart.Document.Body.Append(CrearSeccionDatosCuenta());

                    progreso.Report(0.5, "Agregando información de importe...");
                    Thread.Sleep(100);
                    mainPart.Document.Body.Append(CrearSeccionImporte());

                    progreso.Report(0.6, "Agregando información de beneficiario...");
                    Thread.Sleep(100);
                    mainPart.Document.Body.Append(CrearSeccionBeneficiario());

                    progreso.Report(0.7, "Agregando sección de firmas...");
                    Thread.Sleep(100);
                    mainPart.Document.Body.Append(CrearSeccionFirmas());

                    progreso.Report(0.9, "Guardando documento...");
                    Thread.Sleep(200);
                    mainPart.Document.Save();

                    progreso.Report(1.0, "Completado");
                    Thread.Sleep(300); // Pausa final para mostrar el 100%
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
            paragraph.Append(paragraphProperties);

            paragraph.Append(CrearRunConTexto(NombreBanco!, true, "28"));
            paragraph.Append(new Run(new Break()));
            paragraph.Append(CrearRunConTexto($"Sucursal: {Sucursal}", false, "20"));

            return paragraph;
        }

        private Table CrearSeccionDatosCuenta()
        {
            Table table = new Table();
            TableProperties tableProperties = new TableProperties();
            TableWidth tableWidth = new TableWidth() { Width = "5000", Type = TableWidthUnitValues.Pct };
            TableBorders tableBorders = new TableBorders();
            tableBorders.Append(new TopBorder() { Val = BorderValues.Single, Size = 4 });
            tableBorders.Append(new BottomBorder() { Val = BorderValues.Single, Size = 4 });
            tableBorders.Append(new LeftBorder() { Val = BorderValues.Single, Size = 4 });
            tableBorders.Append(new RightBorder() { Val = BorderValues.Single, Size = 4 });

            tableProperties.Append(tableWidth, tableBorders);
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
            lugarFechaParagraph.Append(CrearRunConTexto($"{LugarEmision}, {FechaEmision:dd/MM/yyyy}"));
            lugarFechaCell.Append(lugarFechaParagraph);

            // Firma (espacio en blanco)
            TableCell firmaCell = new TableCell();
            Paragraph firmaParagraph = new Paragraph();
            firmaParagraph.Append(CrearRunConTexto("Firma:"));
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

        private TableCell CrearCeldaConTexto(string texto, bool negrita = false, int anchoColumna = 1)
        {
            TableCell cell = new TableCell();
            TableCellProperties cellProperties = new TableCellProperties();
            cellProperties.Append(new TableCellWidth() { Type = TableWidthUnitValues.Auto });
            cell.TableCellProperties = cellProperties;

            Paragraph paragraph = new Paragraph();
            paragraph.Append(CrearRunConTexto(texto, negrita, "20"));

            cell.Append(paragraph);
            return cell;
        }

        private Run CrearRunConTexto(string texto, bool negrita = false, string tamañoFuente = "20")
        {
            Run run = new Run();
            RunProperties runProperties = new RunProperties();

            // Configurar fuente compatible con Unicode
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

            if (!string.IsNullOrEmpty(tamañoFuente))
            {
                runProperties.Append(new FontSize() { Val = tamañoFuente });
            }

            run.RunProperties = runProperties;

            // Usar texto con espacio preservado para mantener caracteres especiales
            run.Append(new Text(Utilidades.SanitizarTexto(texto))
            {
                Space = SpaceProcessingModeValues.Preserve
            });

            return run;
        }
    }
}