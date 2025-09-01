using System.Globalization;

namespace ChequeBancarioOpenXML
{
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                Console.WriteLine("=== GENERADOR DE CHEQUES BANCARIOS ===\n");

                // Configurar la codificación de la consola para soportar caracteres especiales
                Console.OutputEncoding = System.Text.Encoding.UTF8;

                /*
                ChequeBancario chequeTest = new ChequeBancario
                {
                    NumeroCheque = "000123456",
                    TitularCuenta = "John Doe",
                    DireccionTitular = "Somestreet Str. 123",
                    Beneficiario = "Jane Doe",
                    Importe = 1245.50m,
                    LugarEmision = "London",
                    FechaEmision = DateTime.Now,
                    NombreBanco = "London Bank",
                    Sucursal = "UK London Centre",
                    NumeroCuenta = "ES12 3456 7890 1234 5678 9012",
                    ReferenciaTransferencia = "Bank loan"
                };
                */

                // Solicitar datos por consola
                ChequeBancario cheque = new ChequeBancario();

                Console.Write("Número de Cheque: ");
                cheque.NumeroCheque = Console.ReadLine();

                Console.Write("Titular de la Cuenta: ");
                cheque.TitularCuenta = Console.ReadLine();

                Console.Write("Dirección del Titular: ");
                cheque.DireccionTitular = Console.ReadLine();

                Console.Write("Beneficiario: ");
                cheque.Beneficiario = Console.ReadLine();

                Console.Write("Referencia de Transferencia: ");
                cheque.ReferenciaTransferencia = Console.ReadLine();

                Console.Write("Importe (ej: 1245,50): ");
                decimal importe;
                while (!decimal.TryParse(Console.ReadLine(), NumberStyles.Any, CultureInfo.CurrentCulture, out importe))
                {
                    Console.Write("Por favor, ingrese un importe válido: ");
                }
                cheque.Importe = importe;

                Console.Write("Lugar de Emisión: ");
                cheque.LugarEmision = Console.ReadLine();

                Console.Write("Fecha de Emisión (dd/mm/aaaa) [Enter para hoy]: ");
                string fechaInput = Console.ReadLine()!;
                if (!string.IsNullOrEmpty(fechaInput))
                {
                    DateTime fechaEmision;
                    while (!DateTime.TryParseExact(fechaInput, "dd/MM/yyyy", CultureInfo.InvariantCulture, DateTimeStyles.None, out fechaEmision))
                    {
                        Console.Write("Formato incorrecto. Use dd/mm/aaaa: ");
                        fechaInput = Console.ReadLine()!;
                    }
                    cheque.FechaEmision = fechaEmision;
                }

                Console.Write("Nombre del Banco: ");
                cheque.NombreBanco = Console.ReadLine();

                Console.Write("Sucursal: ");
                cheque.Sucursal = Console.ReadLine();

                Console.Write("Número de Cuenta: ");
                cheque.NumeroCuenta = Console.ReadLine();

                // Convertir importe a palabras
                cheque.ImporteEnPalabras = ConvertidorImporte.Convertir(cheque.Importe);

                // Generar documento
                string rutaDocumento = Path.Combine(Directory.GetCurrentDirectory(), $"Cheque_{cheque.NumeroCheque}.docx");

                try
                {
                    cheque.GenerarCheque(rutaDocumento);
                    Console.WriteLine($"\nCheque generado exitosamente en: {rutaDocumento}");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"\nError al generar el cheque: {ex.Message}");
                    Console.WriteLine("Presione cualquier tecla para salir...");
                    Console.ReadKey();
                    return;
                }

                Console.WriteLine("Presione cualquier tecla para salir...");
                Console.ReadKey();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error inesperado: {ex.Message}");
                Console.WriteLine("Presione cualquier tecla para salir...");
                Console.ReadKey();
            }
        }
    }
}