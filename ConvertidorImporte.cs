namespace ChequeBancarioOpenXML
{

    // Convertir importes numéricos a palabras
    public static class ConvertidorImporte
    {
        private static readonly string[] unidades = [
            "cero", "uno", "dos", "tres", "cuatro", "cinco", "seis", "siete", "ocho", "nueve"
        ];

        private static readonly string[] decenas = [
            "", "diez", "veinte", "treinta", "cuarenta", "cincuenta",
            "sesenta", "setenta", "ochenta", "noventa"
        ];

        private static readonly string[] especiales = [
            "", "once", "doce", "trece", "catorce", "quince", "dieciséis",
            "diecisiete", "dieciocho", "diecinueve"
        ];

        public static string Convertir(decimal importe)
        {
            if (importe == 0)
                return "cero";

            int parteEntera = (int)Math.Floor(importe);
            int centimos = (int)Math.Round((importe - parteEntera) * 100);

            string resultado = ConvertirNumero(parteEntera);

            if (centimos > 0)
            {
                resultado += " con " + ConvertirNumero(centimos);
            }

            return resultado;
        }

        private static string ConvertirNumero(int numero)
        {
            if (numero < 10)
                return unidades[numero];

            if (numero < 20)
                return especiales[numero - 10];

            if (numero < 100)
            {
                int decena = numero / 10;
                int unidad = numero % 10;

                if (unidad == 0)
                    return decenas[decena];

                return decenas[decena] + " y " + unidades[unidad];
            }

            if (numero < 1000)
            {
                int centena = numero / 100;
                int resto = numero % 100;

                string resultado = centena == 1 ? "cien" : unidades[centena] + "cientos";

                if (resto > 0)
                    resultado += " " + ConvertirNumero(resto);

                return resultado;
            }

            if (numero < 1000000)
            {
                int millar = numero / 1000;
                int resto = numero % 1000;

                string resultado = millar == 1 ? "mil" : ConvertirNumero(millar) + " mil";

                if (resto > 0)
                    resultado += " " + ConvertirNumero(resto);

                return resultado;
            }

            return "Número demasiado grande";
        }
    }
}
