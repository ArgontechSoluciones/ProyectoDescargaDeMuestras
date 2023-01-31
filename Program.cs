// See https://aka.ms/new-console-template for more information
using System;
using System.Data;
using System.Data.OleDb;
using System.IO;




namespace ArchivosRofex
{
    internal class Program
    {

        static string sArchivo = @"C:\Argontech\ExtraerMuestras\CUIT\CUITS.txt";
        static string sCarpetaOriginal = @"C:\netcontent-services\files\documents\legajo\";
        static string connectionString = @"Provider=SQLOLEDB.1;Password=Arg0ntech;Persist Security Info=True;User ID=SA;Initial Catalog=netcontent;Data Source=WIN-013JOGL9DVD\NETCONTENT";
        static string sCarpetaNueva = @"C:\Argontech\ExtraerMuestras\Export\";
        static void Main(string[] args)
        {
            Console.WriteLine("Hello World");
            Console.ReadLine();
            File.OpenRead(sArchivo);

            foreach (string sLinea in File.ReadLines(sArchivo))
            {
                if(sLinea.Length>3)
                    DescargarMuestras(new OleDbConnection(connectionString),sLinea);
                
            }


            Console.ReadLine();
        }

        static void DescargarMuestras(OleDbConnection connection, string sLinea)
        {
            using (connection)
                {
                    
                    
                    OleDbCommand command = new OleDbCommand("SELECT LegajoId FROM dbo.NC_LegajoCampos WHERE Valor = " + "'" + sLinea + "'");
                    Console.WriteLine("Descargando muestras del CUIT: " + sLinea);
                    
                    command.Connection = connection;


                    
                    try
                    {
                        connection.Open();
                        OleDbDataReader reader = command.ExecuteReader();

                        while (reader.Read())
                        {
                            MoverCarpeta(reader[0].ToString(), sLinea);
                            //Console.WriteLine(reader.GetInt32(0) + ", " + reader.GetString(1));
                        }
                        
                        reader.Close();
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine(ex.Message);
                    }
                    
                }
                
            
        }
        
        static void MoverCarpeta(string id, string c)
        {
            Console.WriteLine("Moviendo Legajo: " + id);
            string source = sCarpetaOriginal + id;
            DirectoryInfo d = new DirectoryInfo(source);
            Directory.CreateDirectory(sCarpetaNueva + c + @"\");

            foreach (var f in d.GetFiles("*.pdf"))
            {
                Console.WriteLine("Archivo: " + f.FullName);
                if (f.Name.Substring(f.Name.Length - 10).IndexOf("Pag") == -1) {
                    File.Copy(f.FullName, sCarpetaNueva + c + @"\" + f.Name);
                }
            }
            return;
        }

        static void InsertarAprobadores(string datos, OleDbConnection connection)
        {
            string[] dato = datos.Split(";");
            using (connection)
            {
                OleDbCommand command = new OleDbCommand("INSERT INTO[dbo].[ARG_Aprobadores] VALUES (?,?,?,?,?,?,?,?,?,?,?)");
                command.Parameters.Add("@Proveedor", OleDbType.VarChar, 255).Value = dato[0];
                command.Parameters.Add("@CUITEmisor", OleDbType.BigInt).Value = Int64.Parse(dato[1]);
                if (dato[2] == "Si")
                {
                    dato[2] = "1";
                }
                else
                {
                    dato[2] = "0";
                }
                command.Parameters.AddWithValue("@AprobacionPorCUIT", Int64.Parse(dato[2]));
                
                if (dato[3] == @"N/A")
                {
                    command.Parameters.Add("@CUITReceptor", OleDbType.BigInt).Value = null;
                }
                else 
                {
                    command.Parameters.Add("@CUITReceptor", OleDbType.BigInt).Value = Int64.Parse(dato[3]);
                }
                
                command.Parameters.Add("@Concepto", OleDbType.VarChar, 255).Value = dato[4];
                command.Parameters.Add("@ValidadorPrincipal", OleDbType.VarChar, 255).Value = dato[5];
                command.Parameters.Add("@ValidadorPrincipalCorreo", OleDbType.VarChar, 255).Value = dato[6];
                command.Parameters.Add("@ValidadorSuplente", OleDbType.VarChar, 255).Value = dato[7];
                command.Parameters.Add("@ValidadorSuplenteCorreo", OleDbType.VarChar, 255).Value = dato[8];
                command.Parameters.Add("@Autorizador", OleDbType.VarChar, 255).Value = dato[9];
                command.Parameters.Add("@AutorizadorCorreo", OleDbType.VarChar, 255).Value = dato[10];

                command.Connection = connection;

                // Open the connection and execute the insert command.
                try
                {
                    connection.Open();
                    command.ExecuteNonQuery();

                }
                catch (Exception ex)
                {
                    Console.WriteLine(ex.Message);
                }

            }
        }
    }






}