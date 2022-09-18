using System;
using Aspose.Cells; //Importacion de API

namespace Simulacion
{
    class NumerosAleatorios
    {
        public static void Main(string[] args)
        {
            double aux=0;
            int i=0;
            double mediaEstadistica=0;

            //Creamos una lista para comparar los numeros del archivo que leeremos
            List<double> listaNumeros = new List<double>();

            // Abrimos el archivo, creando un objeto de tipo Workbook para poder obtener su contenido
            Workbook archivo = new Workbook("Numeros Pseudoaleatorios.xlsx");
            
            // Obtenemos el numero total de hojas del libro a trabajar
            WorksheetCollection hojasTotales = archivo.Worksheets;

            // Vamos recorriendo cada hoja del libro para trabajar en la pagina en la que nos encontramos
            for (int hoja = 0; hoja < hojasTotales.Count; hoja++)
            {
                // Vamos obteniendo la pagina en la que nos encontramos del libro, para obtener el tamaño y el contenido de sus celdas
                Worksheet pagina = hojasTotales[hoja];

                // Guardamos el numero de columnas totales para recorrerlas y obtener su dato.
                int filas = pagina.Cells.MaxDataRow;

                // Nos movemos para obtener el dato de cada fila para compararlo.
                for (  i = 0; i <=filas; i++)
                {
                    // Obtenemos el valor de la celda y la guardamos para compararlo con la lista e ir guardando los datos.
                    aux=(pagina.Cells[i,2].DoubleValue)/(filas-1);
                    //Agregamos los numeros generados a una lista
                    listaNumeros.Add(aux);
                    //Vamos sumando los valores para sacar la media estadistica de los datos
                    mediaEstadistica+=aux;
                }
                //Comparamos los numeros de la lista para verificar si tienen relacion -Si tienen relacion detenemos el programa
                for(int j=0;j<listaNumeros.Count;j++)
                {
                    aux=listaNumeros[j];
                    if(j!=(listaNumeros.Count-1))
                    {
                        if(aux==listaNumeros[j+1])
                        {
                            Console.Write("Error los numeros tienen relacion");
                            return;
                        } 
                    }        
                }
                mediaEstadistica/=filas;
                Console.WriteLine("Los "+(i)+" numeros generados son aleatorios. No tienen relacion");
                Console.WriteLine("La media estadistica es igual "+mediaEstadistica);
            }
        }
    }
}

