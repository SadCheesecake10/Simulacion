using System;
using Aspose.Cells; //Importacion de API

namespace Simulacion
{
    class NumerosAleatorios
    {
        public static void Main(string[] args)
        {
            double aux=0;
            int i=0, contador=0;
            double anchoClase=0.0109;
            StreamWriter simulacion = new StreamWriter("intervalos.txt");

            //Creamos una lista para comparar los numeros del archivo que leeremos
            List<double> listaNumeros = new List<double>();
            List<List<double>> intervalos= new List<List<double>>();
            List<double> intervalo= new List<double>();
            List<int> tamanio= new List<int>();

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
                int columnas = pagina.Cells.MaxDataColumn;

                // Nos movemos para obtener el dato de cada fila para compararlo.
                for (  i = 0; i <=filas; i++)
                {
                    // Obtenemos el valor de la celda y la guardamos para compararlo con la lista e ir guardando los datos.
                    aux=(pagina.Cells[i,2].DoubleValue)/(filas-1);
                    //Agregamos los numeros generados a una lista
                    listaNumeros.Add(aux);
                }
                //Comparamos los numeros de la lista para verificar si tienen relacion -Si tienen relacion detenemos el programa
                for(int j=0;j<listaNumeros.Count;j++)
                {
                    aux=listaNumeros[j];
                    if(j!=(listaNumeros.Count-1))
                    {
                        if(aux==listaNumeros[j+1])
                        {
                            Console.WriteLine("Error los numeros tienen relacion");
                            return;
                        } 
                    }       
                }
                while(anchoClase<=1)
                {
                    contador=1;
                    for(int j=0;j<listaNumeros.Count;j++)
                    {
                        if(listaNumeros[j]>=anchoClase && listaNumeros[j]<(anchoClase+0.0109))
                        {
                            intervalo.Add(listaNumeros[j]);
                            contador++;
                        }
                    }
                    intervalos.Add(intervalo);
                    intervalo=new List<double>();
                    tamanio.Add(contador);
                    anchoClase+=0.0109;
                }
                foreach(List<double> lista in intervalos)
                {
                    simulacion.WriteLine();
                    simulacion.Write("Intervalo: ");
                    foreach(double numero in lista)
                    {
                        simulacion.Write(numero+" ");
                    }
                    simulacion.WriteLine();
                    simulacion.WriteLine("Tamaño: "+tamanio[contador]);
                    simulacion.WriteLine();
                }
                simulacion.WriteLine("Los "+(i)+" numeros generados son aleatorios. No tienen relacion");
                for(int j=0;j<listaNumeros.Count;j++)
                {
                    if(listaNumeros[j]<=0.5 && j!=listaNumeros.Count-1)
                    {
                        Console.Write("0, ");
                    }
                    else if(listaNumeros[j]>0.5 && j!=listaNumeros.Count-1)
                    {
                        Console.Write("1, ");
                    }
                    else if(listaNumeros[j]<=0.5 && j==listaNumeros.Count-1)
                    {
                        Console.Write("0.");
                    }
                    else if(listaNumeros[j]>0.5 && j==listaNumeros.Count-1)
                    {
                        Console.Write("1.");
                    }
                }
            }
        }
    }
}

