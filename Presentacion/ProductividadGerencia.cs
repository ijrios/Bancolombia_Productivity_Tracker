using Presentacion.CoreNotificaciones.Common;
using Presentacion.Generales;
using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Transversal;
using Excel = Microsoft.Office.Interop.Excel;
using Outlook = Microsoft.Office.Interop.Outlook;


namespace Presentacion.ProductividadGerencia
{
    public class GenerarProductividadGerencia
    {
        public const int IdAplicacion = IdAplicaciones.ProductividadGerencia;
        public const string nombreAplicacion = "PRODGCIA";
        Aplicacion app = new Aplicacion(IdAplicacion);
        StackTrace stack = new StackTrace();
        Respuesta DatosRespuesta = new Respuesta();
        Respuesta DatosRespuestaPlanta = new Respuesta();
        Respuesta DatosRespuestaDuo = new Respuesta();
        Respuesta DatosRespuestaTris = new Respuesta();
        Respuesta DatosRespuestaQuattuor = new Respuesta();
        Respuesta DatosRespuestaQuinque = new Respuesta();
        Respuesta DatosRespuestaSix = new Respuesta();
        Respuesta DatosRespuestaHabiles = new Respuesta();
        Respuesta DatosRespuestaUsuarios = new Respuesta();
        Respuesta DatosRespuestaUsuariosDuo = new Respuesta();

        public void Ejecutar()
        {
            Aplicaciones.listaAplicaciones.Add(app);

            while (!VPGlobal.Cerrar)
            {
                app = AMP.ConsultarAplicacion(app);


                if (app.EnEjecucion)
                {
                    if (VPGlobal.Cerrar)
                    {
                        AMP.InsertarReporteMonApli(IdAplicacion, EstadosAplicacion.Finalizando);
                        app.AgregarLog(app.NombreAplicacion, EstadosAplicacion.Finalizando);
                        app.EnEjecucion = false;
                        return;
                    }

                    //Reporte Monitor
                    AMP.InsertarReporteMonApli(IdAplicacion, EstadosAplicacion.Procesando);
                    app.AgregarLog(app.NombreAplicacion, EstadosAplicacion.Procesando, app.Intervalo.ToString());

                    //*************************************************************************************
                    //Desarrollo de app
                    DateTime Hoy = DateTime.Now;
                    //DateTime Hoy = Hoyy.AddDays(-1);
                    DateTime Ayer = Hoy.AddDays(-1);
                    DateTime HaceDosDias = Hoy.AddDays(-2);
                    DateTime HaceTresDias = Hoy.AddDays(-3);
                    DateTime HaceCuatroDias = Hoy.AddDays(-4);
                    var DiaActualSemana = Hoy.DayOfWeek;
                    var DiaAyer = Hoy.DayOfWeek;
                    var DiaHaceDosDias = Hoy.DayOfWeek;
                    var DiaHaceTresDias = Hoy.DayOfWeek;
                    var DiaHaceCuatroDias = Hoy.DayOfWeek;

                    //*************************************************************************************
                    DatosRespuestaHabiles = Negocio.ProductividadGerencia.GenerarProductividadGerencia.ConsultarDiasNoHabiles();
                    //************************************************************************************
                    Console.WriteLine(DatosRespuestaHabiles.Dt.Rows.Count);
                    //*************************************************************************************
                    //****************   PARAMETROS DE FUNCIONAMIENTO     *********************************
                    int diaLimiteInferior = 0;
                    int diaLimiteSuperior = 0;
                    string diaSemana = "";
                    string fecha_habil;
                    List<string> fechas_habiles = new List<string>();
                    string dia_habil;
                    //*************************************************************************************

                    Console.WriteLine(DatosRespuestaHabiles.Dt.Rows.Count);

                    if (!DatosRespuestaHabiles.HayFallos)
                    {
                        if (DatosRespuestaHabiles.Dt.Rows.Count > 0)
                        {
                            foreach (DataRow row in DatosRespuestaHabiles.Dt.Rows)
                            {
                                fecha_habil = row["Fecha"].ToString();
                                dia_habil = row["Día"].ToString();
                                Console.WriteLine(fecha_habil);
                                fechas_habiles.Add(fecha_habil);
                            }
                        }
                    }
                    else
                    {
                        app.EnEjecucion = false;
                    }

                    if (DiaActualSemana.ToString() == "Monday")
                    {
                        //Dia viernes festivo
                        if (fechas_habiles.Contains(HaceTresDias.ToString("dd/MM/yyyy")) && fechas_habiles.Contains(HaceCuatroDias.ToString("dd/MM/yyyy")))
                        {
                            diaLimiteInferior = -5;
                            diaLimiteSuperior = -1;
                            diaSemana = "Monday";

                        }
                        else if (fechas_habiles.Contains(HaceTresDias.ToString("dd/MM/yyyy")) && !fechas_habiles.Contains(HaceCuatroDias.ToString("dd/MM/yyyy")))
                        {
                            diaLimiteInferior = -4;
                            diaLimiteSuperior = -1;
                            diaSemana = "Monday";
                        }
                        else if (!fechas_habiles.Contains(HaceTresDias.ToString("dd/MM/yyyy")) && !fechas_habiles.Contains(HaceCuatroDias.ToString("dd/MM/yyyy")))
                        {
                            diaLimiteInferior = -3;
                            diaLimiteSuperior = -1;
                            diaSemana = "Monday";
                        }
                    }
                    else if (DiaActualSemana.ToString() == "Tuesday")
                    {
                        //Dia Lunes festivo
                        if (fechas_habiles.Contains(Ayer.ToString("dd/MM/yyyy")))
                        {
                            diaLimiteInferior = -4;
                            diaLimiteSuperior = -1;
                            diaSemana = "Tuesday";
                        }
                        else
                        {
                            diaLimiteInferior = 0;
                            diaLimiteSuperior = -1;
                            diaSemana = "Monday";
                        }
                    }
                    else if (DiaActualSemana.ToString() == "Wednesday")
                    {
                        //Dia Martes festivo
                        if (fechas_habiles.Contains(Ayer.ToString("dd/MM/yyyy")))
                        {
                            diaLimiteInferior = -2;
                            diaLimiteSuperior = -1;
                            diaSemana = "Wednesday";
                        }
                        else
                        {
                            diaLimiteInferior = 0;
                            diaLimiteSuperior = -1;
                            diaSemana = "Monday";
                        }
                    }
                    else if (DiaActualSemana.ToString() == "Thursday")
                    {
                        //Dia Miercoles festivo
                        if (fechas_habiles.Contains(Ayer.ToString("dd/MM/yyyy")))
                        {
                            diaLimiteInferior = -2;
                            diaLimiteSuperior = -1;
                            diaSemana = "Thursday";
                        }
                        else
                        {
                            diaLimiteInferior = 0;
                            diaLimiteSuperior = -1;
                            diaSemana = "Monday";
                        }
                    }
                    else if (DiaActualSemana.ToString() == "Friday")
                    {
                        //Dia Jueves festivo
                        if (fechas_habiles.Contains(Ayer.ToString("dd/MM/yyyy")))
                        {
                            diaLimiteInferior = -2;
                            diaLimiteSuperior = -1;
                            diaSemana = "Friday";
                        }
                        else
                        {
                            diaLimiteInferior = -1;
                            diaLimiteSuperior = -1;
                            diaSemana = "Monday";
                        }
                    }
                    string Fecha = Hoy.AddDays(diaLimiteSuperior).ToString("dd/MM/yyyy");
                    string FechaTransacciones = Hoy.AddDays(diaLimiteSuperior).ToString("yyyyMMdd");
                    string FechaTransaccionesLunes = Hoy.AddDays(diaLimiteInferior).ToString("yyyyMMdd");
                    string FechaLunes = Hoy.AddDays(diaLimiteInferior).ToString("dd/MM/yyyy");
                    DateTime FechaTransversal = Hoy.AddDays(diaLimiteSuperior);
                    string FechaTransversalLunes = Hoy.AddDays(diaLimiteInferior).ToString("dd/MM/yyyy");
                    string FechaModificaciones = Hoy.AddDays(diaLimiteSuperior).ToString("yyyyMMdd");
                    string FechaModificaiconesLunes = Hoy.AddDays(diaLimiteInferior).ToString("yyyyMMdd");
                    string FechaValidaciones = Hoy.AddDays(diaLimiteSuperior).ToString("yyyy-MM-dd");
                    string FechaValidacionesLunes = Hoy.AddDays(diaLimiteInferior).ToString("yyyy-MM-dd");
                    string FechaAprobacion = Hoy.AddDays(diaLimiteSuperior).ToString("yyMMdd");
                    string FechaAprobacionLunes = Hoy.AddDays(diaLimiteInferior).ToString("yyMMdd");
                    string Completa = Hoy.ToString("yyyyMMdd");
                    string CompletaAyerSemana = Hoy.AddDays(diaLimiteSuperior).ToString("yyyyMMdd");
                    string CompletaAyerFindeSemana = Hoy.AddDays(diaLimiteInferior).ToString("yyyyMMdd");
                    string CompletaMiercoles = Hoy.AddDays(-7).ToString("yyyyMMdd");
                    string Hora = Hoy.ToString("HHmm");
                    string[] gerencia = { "Cartera", "Compraventa", "Trade" };
                    int[] ok = { 0, 0, 0};
                    int[] okDuo = { 0, 0, 0 };

                    /*
                    //*********************************************************************************************************************************
                    // Si se requiere cambiar de proceso solo se se debe modificar la comuna CodBinario en la tabla [CORE_PLATAFORMA].[seg].[USUARIOS]
                    
                    [Compraventa]-[Compraventa aprobacion]-[compraventa linea]-[compraventa transacciones]-[compraventa transversales]
                    Nota: En compraventa el codigo quedo con 5 bits 
                    Nota: En cartera el codigo quedo con 3 bits 
                    Nota: En trade el codigo quedo con 2 bits
                    
                    Linea y Compraventa	10100
                    Compraventa y Aprobracion compraventa	11000
                    Transversales	00001
                    Transacciones y Compraventa	10010
                    Compraventa y transversales = 10001
                    Transacciones, compraventa y aprobacion compraventa	11010
                    Cartera	010
                    Cartera y aprobacio cartera 011
                    Trade	01

                    //*********************************************************************************************************************************
                    */

                    //*************************************************************************************
                    DatosRespuestaUsuarios = Negocio.ProductividadGerencia.GenerarProductividadGerencia.Usuarios();
                    Console.WriteLine(DatosRespuestaUsuarios.Dt.Rows.Count);
                    string NombreUsuario;
                    string UsuarioRed;
                    string UsuarioCIB;
                    string CodGrupo;
                    int UsuarioCedula;
                    string UsuarioLinea;
                    int posicionUnus = 0;
                    int posicionDuo = 1;
                    int posicionTris = 2;
                    int posicionQuattuor = 3;
                    int posicionQuinque = 4;
                    int state = 0;
                    List<string> compraventa = new List<string>();
                    List<string> compraventa_aprobacion = new List<string>();
                    List<string> cartera_aprobacion = new List<string>();
                    List<string> cartera = new List<string>();
                    List<string> trade = new List<string>();
                    Dictionary<string, string> usuarios_transacciones = new Dictionary<string, string>();
                    Dictionary<string, string> usuarios_linea = new Dictionary<string, string>();
                    Dictionary<string, string> usuarios_transversales = new Dictionary<string, string>();
                    List<Planta> Plantae = new List<Planta>();
                    //**************************************************************************************

                    if (!DatosRespuestaUsuarios.HayFallos)
                    {
                        if (DatosRespuestaUsuarios.Dt.Rows.Count > 0)
                        {
                            Console.WriteLine(DatosRespuestaUsuarios.Dt.Rows.Count);


                            foreach (DataRow row in DatosRespuestaUsuarios.Dt.Rows)
                            {

                                NombreUsuario = row["Nombre"].ToString();
                                UsuarioRed = row["Usuario"].ToString();
                                UsuarioCIB = row["UsuarioCIB"].ToString();
                                CodGrupo = row["CodBinario"].ToString();
                                UsuarioLinea = row["UsuarioLinea"].ToString();
                                int? UsuarioCedulaNullable = null;
                                try
                                {
                                    UsuarioCedulaNullable = Convert.ToInt32(row["Cedula"].ToString().Trim());
                                }
                                catch (Exception)
                                {

                                }

                                // Asignamos UsuarioCedula el valor de UsuarioCedulaNullable o 0 si es null
                                UsuarioCedula = UsuarioCedulaNullable ?? 0;


                                if (CodGrupo.Equals("01"))
                                {
                                    trade.Add(UsuarioRed);
                                }
                            }

                            foreach (DataRow row in DatosRespuestaUsuarios.Dt.Rows)
                            {

                                NombreUsuario = row["Nombre"].ToString();
                                UsuarioRed = row["Usuario"].ToString();
                                UsuarioCIB = row["UsuarioCIB"].ToString();
                                CodGrupo = row["CodBinario"].ToString();
                                UsuarioLinea = row["UsuarioLinea"].ToString();
                                int? UsuarioCedulaNullable = null;
                                try
                                {
                                    UsuarioCedulaNullable = Convert.ToInt32(row["Cedula"].ToString().Trim());
                                }
                                catch (Exception)
                                {

                                }

                                // Asignamos UsuarioCedula el valor de UsuarioCedulaNullable o 0 si es null
                                UsuarioCedula = UsuarioCedulaNullable ?? 0;

                                if (CodGrupo.Length == 3)
                                {
                                    if (CodGrupo[posicionDuo] == '1')
                                    {
                                        cartera.Add(UsuarioRed);
                                    }
                                }
                            }

                            foreach (DataRow row in DatosRespuestaUsuarios.Dt.Rows)
                            {

                                NombreUsuario = row["Nombre"].ToString();
                                UsuarioRed = row["Usuario"].ToString();
                                UsuarioCIB = row["UsuarioCIB"].ToString();
                                CodGrupo = row["CodBinario"].ToString();
                                UsuarioLinea = row["UsuarioLinea"].ToString();
                                int? UsuarioCedulaNullable = null;
                                try
                                {
                                    UsuarioCedulaNullable = Convert.ToInt32(row["Cedula"].ToString().Trim());
                                }
                                catch (Exception)
                                {

                                }

                                // Asignamos UsuarioCedula el valor de UsuarioCedulaNullable o 0 si es null
                                UsuarioCedula = UsuarioCedulaNullable ?? 0;

                                if (CodGrupo.Length == 3)
                                {
                                    if (CodGrupo[posicionTris] == '1')
                                    {
                                        cartera_aprobacion.Add(UsuarioRed);
                                    }
                                }

                            }

                            foreach (DataRow row in DatosRespuestaUsuarios.Dt.Rows)
                            {

                                NombreUsuario = row["Nombre"].ToString();
                                UsuarioRed = row["Usuario"].ToString();
                                UsuarioCIB = row["UsuarioCIB"].ToString();
                                CodGrupo = row["CodBinario"].ToString();
                                UsuarioLinea = row["UsuarioLinea"].ToString();
                                int? UsuarioCedulaNullable = null;
                                try
                                {
                                    UsuarioCedulaNullable = Convert.ToInt32(row["Cedula"].ToString().Trim());
                                }
                                catch (Exception)
                                {

                                }

                                // Asignamos UsuarioCedula el valor de UsuarioCedulaNullable o 0 si es null
                                UsuarioCedula = UsuarioCedulaNullable ?? 0;

                                if (CodGrupo.Length == 5)
                                {
                                    if (CodGrupo[posicionUnus] == '1')
                                    {
                                        compraventa.Add(UsuarioRed);
                                    }
                                }

                            }

                            foreach (DataRow row in DatosRespuestaUsuarios.Dt.Rows)
                            {

                                NombreUsuario = row["Nombre"].ToString();
                                UsuarioRed = row["Usuario"].ToString();
                                UsuarioCIB = row["UsuarioCIB"].ToString();
                                CodGrupo = row["CodBinario"].ToString();
                                UsuarioLinea = row["UsuarioLinea"].ToString();
                                int? UsuarioCedulaNullable = null;
                                try
                                {
                                    UsuarioCedulaNullable = Convert.ToInt32(row["Cedula"].ToString().Trim());
                                }
                                catch (Exception)
                                {

                                }

                                // Asignamos UsuarioCedula el valor de UsuarioCedulaNullable o 0 si es null
                                UsuarioCedula = UsuarioCedulaNullable ?? 0;

                                if (CodGrupo.Length == 5)
                                {
                                    if (CodGrupo[posicionDuo] == '1')
                                    {
                                        compraventa_aprobacion.Add(UsuarioRed);
                                    }
                                }
                            }
                            foreach (DataRow row in DatosRespuestaUsuarios.Dt.Rows)
                            {

                                NombreUsuario = row["Nombre"].ToString();
                                UsuarioRed = row["Usuario"].ToString();
                                UsuarioCIB = row["UsuarioCIB"].ToString();
                                CodGrupo = row["CodBinario"].ToString();
                                UsuarioLinea = row["UsuarioLinea"].ToString();
                                int? UsuarioCedulaNullable = null;
                                try
                                {
                                    UsuarioCedulaNullable = Convert.ToInt32(row["Cedula"].ToString().Trim());
                                }
                                catch (Exception)
                                {

                                }

                                // Asignamos UsuarioCedula el valor de UsuarioCedulaNullable o 0 si es null
                                UsuarioCedula = UsuarioCedulaNullable ?? 0;

                                if (CodGrupo.Length == 5)
                                {
                                    if (CodGrupo[posicionTris] == '1')
                                    {
                                        usuarios_linea.Add(UsuarioRed, UsuarioLinea);
                                    }
                                }
                            }

                            foreach (DataRow row in DatosRespuestaUsuarios.Dt.Rows)
                            {

                                NombreUsuario = row["Nombre"].ToString();
                                UsuarioRed = row["Usuario"].ToString();
                                UsuarioCIB = row["UsuarioCIB"].ToString();
                                CodGrupo = row["CodBinario"].ToString();
                                UsuarioLinea = row["UsuarioLinea"].ToString();
                                int? UsuarioCedulaNullable = null;
                                try
                                {
                                    UsuarioCedulaNullable = Convert.ToInt32(row["Cedula"].ToString().Trim());
                                }
                                catch (Exception)
                                {

                                }

                                // Asignamos UsuarioCedula el valor de UsuarioCedulaNullable o 0 si es null
                                UsuarioCedula = UsuarioCedulaNullable ?? 0;

                                if (CodGrupo.Length == 5)
                                {
                                    if (CodGrupo[posicionQuattuor] == '1')
                                    {
                                        usuarios_transacciones.Add(UsuarioCIB, UsuarioRed);
                                    }
                                }
                            }

                            foreach (DataRow row in DatosRespuestaUsuarios.Dt.Rows)
                            {

                                NombreUsuario = row["Nombre"].ToString();
                                UsuarioRed = row["Usuario"].ToString();
                                UsuarioCIB = row["UsuarioCIB"].ToString();
                                CodGrupo = row["CodBinario"].ToString();
                                UsuarioLinea = row["UsuarioLinea"].ToString();
                                int? UsuarioCedulaNullable = null;
                                try
                                {
                                    UsuarioCedulaNullable = Convert.ToInt32(row["Cedula"].ToString().Trim());
                                }
                                catch (Exception)
                                {

                                }

                                // Asignamos UsuarioCedula el valor de UsuarioCedulaNullable o 0 si es null
                                UsuarioCedula = UsuarioCedulaNullable ?? 0;

                                if (CodGrupo.Length == 5)
                                {
                                    if (CodGrupo[posicionQuinque] == '1')
                                    {
                                        try
                                        {
                                            if (!usuarios_transversales.ContainsKey(UsuarioCIB))
                                            {
                                                usuarios_transversales.Add(UsuarioCIB, UsuarioRed);
                                            }
                                            else
                                            {
                                                //
                                                throw new Exception("El usuario ya existe en la lista de usuarios transversales.");
                                            }
                                        }
                                        catch (Exception ex)
                                        {
                                            Console.WriteLine("Error: " + ex.Message);
                                        }
                                    }
                                }
                            }
                            foreach (DataRow row in DatosRespuestaUsuarios.Dt.Rows)
                            {
                                NombreUsuario = row["Nombre"].ToString();
                                UsuarioRed = row["Usuario"].ToString();
                                UsuarioCIB = row["UsuarioCIB"].ToString();
                                CodGrupo = row["CodBinario"].ToString();
                                int? UsuarioCedulaNullable = null;
                                try
                                {
                                    UsuarioCedulaNullable = Convert.ToInt32(row["Cedula"].ToString().Trim());
                                }
                                catch (Exception)
                                {

                                }

                                // Asignamos UsuarioCedula el valor de UsuarioCedulaNullable o 0 si es null
                                UsuarioCedula = UsuarioCedulaNullable ?? 0;


                                if (CodGrupo.Length == 2)
                                {
                                    Planta empleado = new Planta
                                    {
                                        Nombre = NombreUsuario,
                                        Cedula = UsuarioCedula,
                                        Departamento = "trade",
                                        Usuario = UsuarioRed
                                    };

                                    Plantae.Add(empleado);
                                }
                                else if (CodGrupo.Length == 3)
                                {
                                    Planta empleado = new Planta
                                    {
                                        Nombre = NombreUsuario,
                                        Cedula = UsuarioCedula,
                                        Departamento = "cartera",
                                        Usuario = UsuarioRed
                                    };

                                    Plantae.Add(empleado);
                                }
                                else if (CodGrupo.Length == 5)
                                {
                                    Planta empleado = new Planta
                                    {
                                        Nombre = NombreUsuario,
                                        Cedula = UsuarioCedula,
                                        Departamento = "compraventa",
                                        Usuario = UsuarioRed
                                    };

                                    Plantae.Add(empleado);
                                }


                            }
                            Planta.Empleados_all.AddRange(Plantae);
                        }
                    }
                    foreach (var elemento in usuarios_linea)
                    {
                        Console.WriteLine(elemento);
                    }

                    //Correos Excelencia Operacional *******************************************
                    string archiCorreo = null;
                    bool banderaCorreo = false;
                    bool banderaCorreoDuo = false;

                    if (DiaActualSemana.ToString() == "Wednesday")
                    {
                        //Inferior
                        archiCorreo = @"\\sbmdebns03\VP SERV CORP\VP_SERV_CLIE\DIR_SERV_MERC_CAP_CCIO_INT\GCIA_SERV_COM_INT\Operación\Envio y recepción\Productividad\" + "Correo" + CompletaMiercoles + ".txt";
                       
                        if (File.Exists(archiCorreo))
                        {
                            File.Delete(archiCorreo);
                            banderaCorreo = true;
                        }
                        else
                        {
                            banderaCorreo = true;
                        }
                    }
                    else
                    {
                       
                        banderaCorreo = false;
                    }



                   Respuesta CorreosReprocesos = new Respuesta();
                   string archiCorreoPrincipal = @"\\sbmdebns03\VP SERV CORP\VP_SERV_CLIE\DIR_SERV_MERC_CAP_CCIO_INT\GCIA_SERV_COM_INT\Operación\Envio y recepción\Productividad\" + "Correo" + Completa + ".txt";

                    if (!File.Exists(archiCorreoPrincipal))
                    {
                        banderaCorreoDuo = true;
                    }
                    else
                    {
                        banderaCorreoDuo = false;
                    }


                    if (banderaCorreo == true && banderaCorreoDuo == true)
                    {
                        try
                        {

                            DateTime fechaActual = DateTime.Now;
                            string diaActual = fechaActual.DayOfWeek.ToString();
                            int limiteinf = 0;
                            int limitesup = 0;

                            string fechaInicial = null;
                            string fechaFinal = null;
                            string fechaInicialCorreo = null;
                            string fechaFinalCorreo = null;
                            string fechaInicialUsuarios = null;
                            string fechaFinalUsuario = null;
                            DateTime fechaInicialUsuariosDateTime = new DateTime();
                            DateTime fechaFinalUsuariosDateTime = new DateTime();

                            if (diaActual == "Wednesday")
                            {
                                limiteinf = -9;
                                limitesup = -5;
                                fechaInicial = fechaActual.AddDays(limiteinf).ToString("MMdd");
                                fechaFinal = fechaActual.AddDays(limitesup).ToString("MMdd");
                                fechaInicialUsuarios = fechaActual.AddDays(limiteinf).ToString("yyyy-MM-dd");
                                fechaFinalUsuario = fechaActual.AddDays(limitesup).ToString("yyyy-MM-dd");
                                fechaInicialCorreo = fechaActual.AddDays(-2).ToString("yyyyMMdd");
                                fechaFinalCorreo = fechaActual.AddDays(0).ToString("yyyyMMdd");
                                fechaInicialUsuariosDateTime = DateTime.ParseExact(fechaInicialUsuarios, "yyyy-MM-dd", System.Globalization.CultureInfo.InvariantCulture);
                                fechaFinalUsuariosDateTime = DateTime.ParseExact(fechaFinalUsuario, "yyyy-MM-dd", System.Globalization.CultureInfo.InvariantCulture);
                            }
                           

                            CorreosReprocesos = Negocio.ProductividadGerencia.GenerarProductividadGerencia.ReprocesosOmniaFechaAdmin(fechaInicialUsuariosDateTime, fechaFinalUsuariosDateTime, 0);
                            List<ErrorUsuario> erroresPorArea = new List<ErrorUsuario>();

                            if (!CorreosReprocesos.HayFallos && CorreosReprocesos.Dt.Rows.Count > 0)
                            {
                                var erroresAgrupados = CorreosReprocesos.Dt.AsEnumerable()
                                    .GroupBy(row => row["UsuarioResponsable"].ToString())
                                    .Select(g => new ErrorUsuario
                                    {
                                        Usuario = g.Key,
                                        Errores = g.Count().ToString()
                                    })
                                    .OrderByDescending(es => Convert.ToInt32(es.Errores))
                                    .ToList();

                                erroresPorArea.AddRange(erroresAgrupados);
                            }
                            string nombreusu = null;
                            string[] usuarios_tp = {
            "scjarami",
            "ytdelgad",
            "yulope",
            "melres",
            "mialozan",
            "ncpabon",
            "anmagira",
            "dcsantam",
            "ympabon",
            "erigarci",
            "matgir",
            "vecorrea",
            "machav",
            "jmorale",
            "erocamp",
            "jhatorre",
            "loymena",
            "jmgarcia",
            "maclondo",
            "jcaldero",
            "jsceball",
            "naaherre",
            "dacjimen",
            "omzuluag",
            "erjmonto",
            "edtapias",
            "kycasas",
            "facanoga",
            "nazapata",
            "lfcastri",
            "astorres",
            "luecruz",
            "leydive"
        };
                            foreach (var item in erroresPorArea)
                            {
                                nombreusu = Negocio.ProductividadGerencia.GenerarProductividadGerencia.ConsultarGeneral(item.Usuario);
                                string Correos_OK = @"\\sbmdebns03\VP SERV CORP\VP_SERV_CLIE\DIR_SERV_MERC_CAP_CCIO_INT\GCIA_SERV_COM_INT\Operación\Envio y recepción\Productividad\" + "Correo" + Completa + ".txt";

                               
                                if (item.Errores == "1")
                                {
                                    if (usuarios_tp.Contains(item.Usuario))
                                    {
                                        EnviarNotificacionExcelencia(nombreusu,"Leidy.GuevaraSanta@teleperformance.com; Luis.MazoPino@teleperformance.com; Maria.ZuluagaGomez@teleperformance.com", item.Errores + " error nuevo reportado");

                                        using (StreamWriter writer = new StreamWriter(Correos_OK))
                                        {
                                            string linea = $"Archivo enviado";
                                            writer.WriteLine(linea);
                                            writer.Flush();

                                        }
                                    }
                                    else
                                    {
                                        EnviarNotificacionExcelencia(nombreusu, item.Usuario, item.Errores + " error nuevo reportado");

                                        using (StreamWriter writer = new StreamWriter(Correos_OK))
                                        {
                                            string linea = $"Archivo enviado";
                                            writer.WriteLine(linea);
                                            writer.Flush();

                                        }
                                    }
                                    
                                }
                                else
                                {
                                    if (usuarios_tp.Contains(item.Usuario))
                                    {
                                        EnviarNotificacionExcelencia(nombreusu, "Leidy.GuevaraSanta@teleperformance.com; Luis.MazoPino@teleperformance.com; Maria.ZuluagaGomez@teleperformance.com", item.Errores + " errores nuevos reportados");

                                        using (StreamWriter writer = new StreamWriter(Correos_OK))
                                        {
                                            string linea = $"Archivo enviado";
                                            writer.WriteLine(linea);
                                            writer.Flush();

                                        }
                                    }
                                    else
                                    {
                                        EnviarNotificacionExcelencia(nombreusu, item.Usuario, item.Errores + " errores nuevos reportados");

                                        using (StreamWriter writer = new StreamWriter(Correos_OK))
                                        {
                                            string linea = $"Archivo enviado";
                                            writer.WriteLine(linea);
                                            writer.Flush();

                                        }
                                    }
                                  
                                }
                              
                            }
                          

                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine("Correos no Enviados"+ex);
                        }
                       
                    }
                    else
                    {
                        Console.WriteLine("Correos Enviados");
                    }


                    List<string> FechaDuo;
                    string periodos = "";
                    string dia = "";
                    bool bandera = true;
                    int lineaDuo = 0;
                    string archi = null;
                    string archiDuo = null;
                    string archiTris = null;
                    bool banderaSecundaria = false;

                    if (DiaActualSemana.ToString() == diaSemana)
                    {
                        //Inferior
                        archi = @"\\sbmdebns03\VP SERV CORP\VP_SERV_CLIE\DIR_SERV_MERC_CAP_CCIO_INT\GCIA_SERV_COM_INT\Operación\Envio y recepción\Productividad\" + "Cartera" + CompletaAyerFindeSemana + ".txt";
                        archiDuo = @"\\sbmdebns03\VP SERV CORP\VP_SERV_CLIE\DIR_SERV_MERC_CAP_CCIO_INT\GCIA_SERV_COM_INT\Operación\Envio y recepción\Productividad\" + "Compraventa" + CompletaAyerFindeSemana + ".txt";
                        archiTris = @"\\sbmdebns03\VP SERV CORP\VP_SERV_CLIE\DIR_SERV_MERC_CAP_CCIO_INT\GCIA_SERV_COM_INT\Operación\Envio y recepción\Productividad\" + "Trade" + CompletaAyerFindeSemana + ".txt";

                        if (File.Exists(archi) && File.Exists(archi) && File.Exists(archi))
                        {
                            File.Delete(archi);
                            File.Delete(archiDuo);
                            File.Delete(archiTris);
                            banderaSecundaria = true;
                        }
                        else
                        {
                            banderaSecundaria = true;
                        }


                    }
                    else
                    {
                        //Superior
                        archi = @"\\sbmdebns03\VP SERV CORP\VP_SERV_CLIE\DIR_SERV_MERC_CAP_CCIO_INT\GCIA_SERV_COM_INT\Operación\Envio y recepción\Productividad\" + "Cartera" + CompletaAyerSemana + ".txt";
                        archiDuo = @"\\sbmdebns03\VP SERV CORP\VP_SERV_CLIE\DIR_SERV_MERC_CAP_CCIO_INT\GCIA_SERV_COM_INT\Operación\Envio y recepción\Productividad\" + "Compraventa" + CompletaAyerSemana + ".txt";
                        archiTris = @"\\sbmdebns03\VP SERV CORP\VP_SERV_CLIE\DIR_SERV_MERC_CAP_CCIO_INT\GCIA_SERV_COM_INT\Operación\Envio y recepción\Productividad\" + "Trade" + CompletaAyerSemana + ".txt";

                        if (File.Exists(archi) && File.Exists(archi) && File.Exists(archi))
                        {
                            File.Delete(archi);
                            File.Delete(archiDuo);
                            File.Delete(archiTris);
                            banderaSecundaria = true;
                        }
                        else
                        {
                            banderaSecundaria = true;
                        }

                    }


                    string archivo = @"\\sbmdebns03\VP SERV CORP\VP_SERV_CLIE\DIR_SERV_MERC_CAP_CCIO_INT\GCIA_SERV_COM_INT\Operación\Envio y recepción\Productividad\" + "Cartera" + Completa + ".txt";
                    string archivoDuo = @"\\sbmdebns03\VP SERV CORP\VP_SERV_CLIE\DIR_SERV_MERC_CAP_CCIO_INT\GCIA_SERV_COM_INT\Operación\Envio y recepción\Productividad\" + "Compraventa" + Completa + ".txt";
                    string archivoTris = @"\\sbmdebns03\VP SERV CORP\VP_SERV_CLIE\DIR_SERV_MERC_CAP_CCIO_INT\GCIA_SERV_COM_INT\Operación\Envio y recepción\Productividad\" + "Trade" + Completa + ".txt";



                    while (bandera == true && banderaSecundaria == true)
                    {
                        
                        if (File.Exists(archivo) && File.Exists(archivoDuo) && File.Exists(archivoTris))
                        {
                            app.EnEjecucion = false;
                            bandera = false;
                            banderaSecundaria = false;
                        }

                        //Validar archivo de la linea 

                        string filePathLinea = @"\\sbmdebns03\VP SERV CORP\VP_SERV_CLIE\DIR_SERV_MERC_CAP_CCIO_INT\GCIA_SERV_COM_INT\Operación\Envio y recepción\Productividad\Linea.xlsx";

                        Excel.Application excelAppLinea = new Excel.Application();
                        Excel.Workbook workbookLinea = excelAppLinea.Workbooks.Open(filePathLinea);

                        try
                        {
                            Excel.Worksheet worksheet = (Excel.Worksheet)workbookLinea.Worksheets[1]; // Obtener la primera hoja
                            Excel.Range range = worksheet.UsedRange;
                            int rowCount = range.Rows.Count;
                            string fechaLinea = null;

                            for (int row = 4; row <= rowCount; row++)
                            {
                                object cellValue = (range.Cells[row, 1] as Excel.Range)?.Value;

                                if (cellValue is DateTime fechaDateTime)
                                {
                                    fechaLinea = fechaDateTime.ToString("dd/MM/yyyy", CultureInfo.InvariantCulture);
                                    
                                }
                                else if (cellValue != null)
                                {

                                    fechaLinea = cellValue.ToString();
                                }

                                if (DiaActualSemana.ToString() == diaSemana)
                                {
                                    if (fechaLinea == FechaLunes)
                                    {
                                        lineaDuo = 1;
                                    }
                                }
                                else
                                {
                                    if (fechaLinea == Fecha)
                                    {
                                        lineaDuo = 1;
                                    }
                                }

                            }
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine("Error al leer el archivo Excel: " + ex.Message);
                        }
                        finally
                        {
                            // Cerrar y liberar recursos
                            workbookLinea.Close();
                            excelAppLinea.Quit();
                        }


                        if (gerencia.Contains("Cartera") && ok.ElementAt(0).Equals(0) && !File.Exists(archivo))
                        {
                            // Inicialización de la primera instancia de Excel Application
                            Excel.Application myexcelApplication = new Excel.Application();

                            // Inicialización de la segunda instancia de Excel Application
                            Excel.Application myexcelApplication_duo = new Excel.Application();

                            // Inicialización de la segunda instancia de Excel Application
                            Excel.Application myexcelApplication_tris = new Excel.Application();


                            Excel.Workbook myexcelWorkbook = myexcelApplication.Workbooks.Add();
                            Excel.Worksheet myexcelWorksheet = (Excel.Worksheet)myexcelWorkbook.Sheets.Add();

                            EmpleadoDuo nuevoEmpleado;
                            List<EmpleadoDuo> empleadosDuoList = new List<EmpleadoDuo>();
                            List<PuntajeDia> puntajeList = new List<PuntajeDia>();
                            int estado = 0;

                            Console.WriteLine("AdentroCartera");
                            if (DiaActualSemana.ToString() == diaSemana)
                            {
                                DatosRespuesta = Negocio.ProductividadGerencia.GenerarProductividadGerencia.ConsultarProductividad(FechaAprobacionLunes);
                                DatosRespuestaDuo = Negocio.ProductividadGerencia.GenerarProductividadGerencia.ConsultarProductividadGlobal(-1, FechaLunes, 0);
                                Console.WriteLine("Datos consultados Lunes");
                                periodos = Hoy.AddDays(diaLimiteInferior).ToString("yyyyMM");
                                dia = Hoy.AddDays(diaLimiteInferior).ToString("dd");
                            }
                            else
                            {
                                DatosRespuesta = Negocio.ProductividadGerencia.GenerarProductividadGerencia.ConsultarProductividad(FechaAprobacion);
                                DatosRespuestaDuo = Negocio.ProductividadGerencia.GenerarProductividadGerencia.ConsultarProductividadGlobal(-1, Fecha, 0);
                                Console.WriteLine("Datos consultados");
                                periodos = Hoy.AddDays(diaLimiteSuperior).ToString("yyyyMM");
                                dia = Hoy.AddDays(diaLimiteSuperior).ToString("dd");
                            }

                            Console.WriteLine(DatosRespuesta.Dt.Rows.Count);
                            Console.WriteLine(DatosRespuestaDuo.Dt.Rows.Count);

                            if (!DatosRespuesta.HayFallos && !DatosRespuestaDuo.HayFallos)
                            {
                                if (DatosRespuesta.Dt.Rows.Count > 0 && DatosRespuestaDuo.Dt.Rows.Count > 0)
                                {
                                    Console.WriteLine("Aqui en cartera");
                                    int rowunus = 2;
                                    ////////////////    PRODUCTIVIDAD GLOBAL  /////////////////////////////////////////////////////////
                                    if (estado == 0)
                                    {
                                        foreach (DataRow row in DatosRespuestaDuo.Dt.Rows)
                                        {
                                            foreach (var item in row.ItemArray)
                                            {
                                                Console.Write(item.ToString() + "\t");
                                            }
                                            Console.WriteLine();

                                            string usuario = row["Usuario"].ToString().ToLower();
                                            double total = Convert.ToDouble(row["Total"]);

                                            if (usuario != "zona de totales")
                                            {
                                                if (cartera.Contains(usuario))
                                                {

                                                    // Agregar los datos al nuevo empleado
                                                    nuevoEmpleado = new EmpleadoDuo();
                                                    nuevoEmpleado.Gundi = 70452; // Este es el código para el resto de usuarios
                                                    nuevoEmpleado.Cedula = 0;
                                                    nuevoEmpleado.Usuario = usuario;
                                                    nuevoEmpleado.Puntaje = total;
                                                    nuevoEmpleado.Periodo = periodos;

                                                    // Agregar el nuevo empleado a la lista
                                                    empleadosDuoList.Add(nuevoEmpleado);
                                                }
                                            }

                                        }
                                        EmpleadoDuo.Empleo.AddRange(empleadosDuoList);
                                        estado = 1;

                                    }

                                    //////////////////// PRODUCTIVIDAD APROBACION    ////////////////////////////////////////////////
                                    if (estado == 1)
                                    {
                                        foreach (DataRow row in DatosRespuesta.Dt.Rows)
                                        {
                                            foreach (var item in row.ItemArray)
                                            {
                                                Console.Write(item.ToString() + "\t");
                                            }
                                            Console.WriteLine();

                                            string usuario = row["USUARIORED"].ToString().ToLower();
                                            double puntos_totales = 0;

                                            // Calcular puntos totales
                                            for (int column = 7; column <= 19; column++)
                                            {
                                                puntos_totales += Convert.ToDouble(row["H" + column]);
                                            }


                                            // Verificar si el usuario está en la lista de aprobación de cartera
                                            if (cartera_aprobacion.Contains(usuario))
                                            {
                                                nuevoEmpleado = new EmpleadoDuo();
                                                nuevoEmpleado.Gundi = 41021;
                                                // Agregar los datos al nuevo empleado
                                                nuevoEmpleado.Cedula = 0;
                                                nuevoEmpleado.Usuario = usuario;
                                                nuevoEmpleado.Puntaje = puntos_totales;
                                                nuevoEmpleado.Periodo = periodos;


                                                empleadosDuoList.Add(nuevoEmpleado);
                                            }

                                        }
                                        EmpleadoDuo.Empleo.AddRange(empleadosDuoList);
                                        estado = 2;
                                    }
                                    if (estado == 2)
                                    {
                                        foreach (EmpleadoDuo planta in EmpleadoDuo.Empleo)
                                        {
                                            foreach (Planta plantita in Plantae)
                                            {
                                                if (planta.usuario.Contains(plantita.usuario))
                                                {
                                                    planta.cedula = plantita.cedula;
                                                    planta.departamento = plantita.departamento;

                                                }

                                            }

                                        }
                                        estado = 3;
                                    }

                                    if (estado == 3)
                                    {

                                        myexcelWorksheet.Cells[1, 1].Value = "gundi";
                                        myexcelWorksheet.Cells[1, 2].Value = "Cedula";
                                        myexcelWorksheet.Cells[1, 3].Value = "Usuario Red";
                                        myexcelWorksheet.Cells[1, 4].Value = "Periodo";


                                        // Agregar los días del mes como headers
                                        for (int i = 0; i < 31; i++)
                                        {
                                            myexcelWorksheet.Cells[1, i + 5].Value = "Dia " + (i + 1);
                                        }

                                        foreach (var empleado in EmpleadoDuo.Empleo)
                                        {
                                            if (empleado.departamento == "cartera")
                                            {
                                                myexcelWorksheet.Cells[rowunus, 1].Value = empleado.Gundi;
                                                myexcelWorksheet.Cells[rowunus, 2].Value = empleado.Cedula;
                                                myexcelWorksheet.Cells[rowunus, 3].Value = empleado.Usuario;
                                                myexcelWorksheet.Cells[rowunus, 4].Value = empleado.periodo;

                                                double puntajeTotal = empleado.Puntaje;

                                                int columnIndex = 5; // Comenzando desde la columna de "Dia 1"
                                                if (dia == "01" || dia == "1")
                                                {
                                                    int die = Convert.ToInt32(dia);
                                                    myexcelWorksheet.Cells[rowunus, columnIndex * die].Value = puntajeTotal;
                                                }
                                                else
                                                {
                                                    int die = Convert.ToInt32(dia);
                                                    myexcelWorksheet.Cells[rowunus, columnIndex + (die - 1)].Value = puntajeTotal;
                                                }
                                                rowunus++;
                                            }

                                        }
                                        string ProductividadCartera = @"\\sbmdebpmici01v\Files\mici\auto\" + Completa + Hora + "-1648" + ".csv";
                                        //string ProductividadCartera = @"\\sbmdebpmici01v\Files\mici\auto\" + Completa + Hora + "-1648" + ".cvs";
                                        //myexcelApplication.ActiveWorkbook.SaveAs(@"D:\abc.xls", Excel.XlFileFormat.xlWorkbookNormal);
                                        myexcelApplication.ActiveWorkbook.SaveAs(ProductividadCartera, Excel.XlFileFormat.xlCSV);
                                        Console.WriteLine("Archivo generado");

                                        myexcelWorkbook.Close();
                                        myexcelApplication.Quit();
                                        ///// Archivo Plano 
                                        string archivoconfirmacion = @"\\sbmdebns03\VP SERV CORP\VP_SERV_CLIE\DIR_SERV_MERC_CAP_CCIO_INT\GCIA_SERV_COM_INT\Operación\Envio y recepción\Productividad\" + "Cartera" + Completa + ".txt";

                                        using (StreamWriter writer = new StreamWriter(archivoconfirmacion))
                                        {
                                            string linea = $"Archivo enviado";
                                            writer.WriteLine(linea);
                                            writer.Flush();

                                        }


                                    }

                                }

                            }

                        }

                        else if (gerencia.Contains("Compraventa") && ok.ElementAt(1).Equals(0) && lineaDuo==0 && !File.Exists(archivoDuo))
                        {
                            // Inicialización de la primera instancia de Excel Application
                            Excel.Application myexcelApplication = new Excel.Application();

                            // Inicialización de la segunda instancia de Excel Application
                            Excel.Application myexcelApplication_duo = new Excel.Application();

                            // Inicialización de la segunda instancia de Excel Application
                            Excel.Application myexcelApplication_tris = new Excel.Application();



                            Excel.Workbook myexcelWorkbook = myexcelApplication.Workbooks.Add();
                            Excel.Worksheet myexcelWorksheet = (Excel.Worksheet)myexcelWorkbook.Sheets.Add();

                            // Diccionario para almacenar el nombre de usuario y la cantidad total de transacciones
                            Dictionary<string, int> transaccionesPorUsuario = new Dictionary<string, int>();
                            Dictionary<string, double> usuarioTotales = new Dictionary<string, double>();
                            List<EmpleadoDuo> empleadosDuoList = new List<EmpleadoDuo>();
                            List<PuntajeDia> puntajeList = new List<PuntajeDia>();
                            Dictionary<string, int> modificaciones = new Dictionary<string, int>();
                            Dictionary<string, int> validaciones = new Dictionary<string, int>();
                            Dictionary<string, int> validacionesANY = new Dictionary<string, int>();
                            Dictionary<string, int> rechazosANY = new Dictionary<string, int>();
                            EmpleadoDuo nuevoEmpleado;
                            PuntajeDia puntajes;

                            string fechanueva = "";
                            int rowunus = 2;
                            int estado = 0;
                            FechaDuo = new List<string>();
                            List<string> Transacciones_cib = new List<string>();
                            puntajes = new PuntajeDia();
                            if (DiaActualSemana.ToString() == diaSemana)
                            {

                                DatosRespuesta = Negocio.ProductividadGerencia.GenerarProductividadGerencia.ConsultarProductividad(FechaAprobacionLunes);
                                DatosRespuestaDuo = Negocio.ProductividadGerencia.GenerarProductividadGerencia.ConsultarProductividadGlobal(-1, FechaLunes, 0);
                                DatosRespuestaTris = Negocio.ProductividadGerencia.GenerarProductividadGerencia.ConsultarProductividadTransacciones(FechaTransaccionesLunes);
                                DatosRespuestaQuattuor = Negocio.ProductividadGerencia.GenerarProductividadGerencia.ConsultarProductividadRechazos(DateTime.Parse(FechaTransversalLunes));
                                DatosRespuestaQuinque = Negocio.ProductividadGerencia.GenerarProductividadGerencia.ModificacionUsuario(FechaModificaiconesLunes, FechaModificaiconesLunes);
                                DatosRespuestaSix = Negocio.ProductividadGerencia.GenerarProductividadGerencia.ValidacionUsuario(FechaValidacionesLunes, FechaValidacionesLunes);
                                Console.WriteLine("Datos consultados Lunes");
                                periodos = Hoy.AddDays(diaLimiteInferior).ToString("yyyyMM");
                                dia = Hoy.AddDays(diaLimiteInferior).ToString("dd");
                                fechanueva = Hoy.AddDays(diaLimiteInferior).ToString("dd/MM/yyyy");

                            }
                            else
                            {
                                DatosRespuesta = Negocio.ProductividadGerencia.GenerarProductividadGerencia.ConsultarProductividad(FechaAprobacion);
                                DatosRespuestaDuo = Negocio.ProductividadGerencia.GenerarProductividadGerencia.ConsultarProductividadGlobal(-1, Fecha, 0);
                                DatosRespuestaTris = Negocio.ProductividadGerencia.GenerarProductividadGerencia.ConsultarProductividadTransacciones(FechaTransacciones);
                                DatosRespuestaQuattuor = Negocio.ProductividadGerencia.GenerarProductividadGerencia.ConsultarProductividadRechazos(FechaTransversal);
                                DatosRespuestaQuinque = Negocio.ProductividadGerencia.GenerarProductividadGerencia.ModificacionUsuario(FechaModificaciones, FechaModificaciones);
                                DatosRespuestaSix = Negocio.ProductividadGerencia.GenerarProductividadGerencia.ValidacionUsuario(FechaValidaciones, FechaValidaciones);
                                periodos = Hoy.AddDays(diaLimiteSuperior).ToString("yyyyMM");
                                dia = Hoy.AddDays(diaLimiteSuperior).ToString("dd");
                                Console.WriteLine("Datos consultados");
                                fechanueva = Hoy.AddDays(diaLimiteSuperior).ToString("dd/MM/yyyy");
                            }

                            Console.WriteLine(DatosRespuesta.Dt.Rows.Count);
                            Console.WriteLine(DatosRespuestaDuo.Dt.Rows.Count);
                            Console.WriteLine(DatosRespuestaTris.Dt.Rows.Count);
                            Console.WriteLine(DatosRespuestaQuattuor.Dt.Rows.Count);
                            Console.WriteLine(DatosRespuestaQuinque.Dt.Rows.Count);
                            Console.WriteLine(DatosRespuestaSix.Dt.Rows.Count);

                            if (!DatosRespuesta.HayFallos && !DatosRespuestaDuo.HayFallos)
                            {
                                if (DatosRespuesta.Dt.Rows.Count > 0 && DatosRespuestaDuo.Dt.Rows.Count > 0)
                                {
                                    Console.WriteLine("Aqui en compraventa");

                                    //////   PRODUCTIVIDAD GLOBAL /////////////////////////////////////////////////////////////////////////
                                    if (estado == 0)
                                    {
                                        empleadosDuoList.Clear();
                                        puntajeList.Clear();


                                        foreach (DataRow row in DatosRespuestaDuo.Dt.Rows)
                                        {
                                            foreach (var item in row.ItemArray)
                                            {
                                                Console.Write(item.ToString() + "\t");
                                            }
                                            Console.WriteLine();

                                            string usuario = row["Usuario"].ToString().ToLower();
                                            double total = Convert.ToDouble(row["Total"]);
                                            double uno = Convert.ToDouble(row["7"]);
                                            double dos = Convert.ToDouble(row["8"]);
                                            double tres = Convert.ToDouble(row["9"]);
                                            double cuatro = Convert.ToDouble(row["10"]);
                                            double cinco = Convert.ToDouble(row["11"]);
                                            double seis = Convert.ToDouble(row["12"]);
                                            double siete = Convert.ToDouble(row["13"]);
                                            double ocho = Convert.ToDouble(row["14"]);
                                            double nueve = Convert.ToDouble(row["15"]);
                                            double diez = Convert.ToDouble(row["16"]);
                                            double once = Convert.ToDouble(row["17"]);
                                            double doce = Convert.ToDouble(row["18"]);
                                            double trece = Convert.ToDouble(row["19"]);
                                            double catorce = Convert.ToDouble(row["20"]);
                                            double quince = Convert.ToDouble(row["21"]);

                                            if (usuario != "zona de totales")
                                            {
                                                if (compraventa.Contains(usuario))
                                                {

                                                    nuevoEmpleado = new EmpleadoDuo();
                                                    puntajes = new PuntajeDia();
                                                    // Asignar el código correspondiente al usuario
                                                    if (compraventa_aprobacion.Contains(usuario))
                                                    {
                                                        nuevoEmpleado.Gundi = 45210;
                                                    }
                                                    else
                                                    {
                                                        nuevoEmpleado.Gundi = 18046; // Código para otros usuarios seleccionados
                                                    }
                                                    nuevoEmpleado.Cedula = 0;
                                                    nuevoEmpleado.Usuario = usuario;
                                                    nuevoEmpleado.Puntaje = total;
                                                    nuevoEmpleado.Periodo = periodos;

                                                    puntajes.Usuario = usuario;
                                                    puntajes.Puntaje_total = total;
                                                    puntajes.Puntaje_7 = uno;
                                                    puntajes.Puntaje_8 = dos;
                                                    puntajes.Puntaje_9 = tres;
                                                    puntajes.Puntaje_10 = cuatro;
                                                    puntajes.Puntaje_11 = cinco;
                                                    puntajes.Puntaje_12 = seis;
                                                    puntajes.Puntaje_13 = siete;
                                                    puntajes.Puntaje_14 = ocho;
                                                    puntajes.Puntaje_15 = nueve;
                                                    puntajes.Puntaje_16 = diez;
                                                    puntajes.Puntaje_17 = once;
                                                    puntajes.Puntaje_18 = doce;
                                                    puntajes.Puntaje_19 = trece;
                                                    puntajes.Puntaje_20 = catorce;
                                                    puntajes.Puntaje_21 = quince;


                                                    puntajeList.Add(puntajes);
                                                    empleadosDuoList.Add(nuevoEmpleado);

                                                }
                                            }

                                        }
                                        PuntajeDia.productivo.AddRange(puntajeList);
                                        EmpleadoDuo.Empleo.AddRange(empleadosDuoList);
                                        estado = 1;
                                    }

                                    ////////   PRODUCTIVIDAD APROBRACION   /////////////////////////////////////////////////////////////////////////
                                    if (estado == 1)
                                    {
                                        empleadosDuoList.Clear();
                                        puntajeList.Clear();
                                        bool Puntaje_totalAgregado = false;
                                        foreach (DataRow row in DatosRespuesta.Dt.Rows)
                                        {
                                            foreach (var item in row.ItemArray)
                                            {
                                                Console.Write(item.ToString() + "\t");
                                            }
                                            Console.WriteLine();

                                            string usuario = row["USUARIORED"].ToString().ToLower();
                                            double puntos_totales = 0;

                                            // Calcular puntos totales
                                            for (int column = 7; column <= 19; column++)
                                            {
                                                puntos_totales += Convert.ToDouble(row["H" + column]);
                                            }

                                            // Verificar si el nombre de usuario está en la lista de usuarios a procesar
                                            if (compraventa_aprobacion.Contains(usuario))
                                            {
                                                nuevoEmpleado = new EmpleadoDuo();
                                                nuevoEmpleado.Gundi = 15613;
                                                // Agregar los datos al nuevo empleado
                                                nuevoEmpleado.Cedula = 0;
                                                nuevoEmpleado.Usuario = usuario;
                                                nuevoEmpleado.Puntaje = puntos_totales;
                                                nuevoEmpleado.Periodo = periodos;

                                                foreach (PuntajeDia imp in PuntajeDia.productivo)
                                                {
                                                    if (compraventa_aprobacion.Contains(imp.Usuario))
                                                    {
                                                        if (!Puntaje_totalAgregado)
                                                        {
                                                            imp.Puntaje_total = imp.Puntaje_total + puntos_totales;
                                                            Puntaje_totalAgregado = true;
                                                        }

                                                        imp.Puntaje_7 += Convert.ToDouble(row[1]);
                                                        imp.Puntaje_8 += Convert.ToDouble(row[2]);
                                                        imp.Puntaje_9 += Convert.ToDouble(row[3]);
                                                        imp.Puntaje_10 += Convert.ToDouble(row[4]);
                                                        imp.Puntaje_11 += Convert.ToDouble(row[5]);
                                                        imp.Puntaje_12 += Convert.ToDouble(row[6]);
                                                        imp.Puntaje_13 += Convert.ToDouble(row[7]);
                                                        imp.Puntaje_14 += Convert.ToDouble(row[8]);
                                                        imp.Puntaje_15 += Convert.ToDouble(row[9]);
                                                        imp.Puntaje_16 += Convert.ToDouble(row[10]);
                                                        imp.Puntaje_17 += Convert.ToDouble(row[11]);
                                                        imp.Puntaje_18 += Convert.ToDouble(row[12]);
                                                        imp.Puntaje_19 += Convert.ToDouble(row[13]);
                                                    }

                                                }

                                                empleadosDuoList.Add(nuevoEmpleado);
                                            }

                                        }

                                        EmpleadoDuo.Empleo.AddRange(empleadosDuoList);
                                        estado = 2;
                                    }

                                    ////////  PRODUCTIVIDAD DE LA LINEA  /////////////////////////////////////////////////////////////////////////
                                    if (estado == 2)
                                    {
                                        empleadosDuoList.Clear();
                                        puntajeList.Clear();
                                        string filePath = @"\\sbmdebns03\VP SERV CORP\VP_SERV_CLIE\DIR_SERV_MERC_CAP_CCIO_INT\GCIA_SERV_COM_INT\Operación\Envio y recepción\Productividad\Linea.xlsx";

                                        Excel.Application excelApp = new Excel.Application();
                                        Excel.Workbook workbook = excelApp.Workbooks.Open(filePath);

                                        try
                                        {
                                            Excel.Worksheet worksheet = (Excel.Worksheet)workbook.Worksheets[1]; // Obtener la primera hoja
                                            Excel.Range range = worksheet.UsedRange;
                                            int rowCount = range.Rows.Count;

                                            for (int row = 4; row <= rowCount; row++)
                                            {
                                                string fecha = Convert.ToString(range.Cells[row, 1].Value); // Obtener la fecha de la primera columna
                                                string splitSkill = Convert.ToString(range.Cells[row, 2].Value); // Obtener el split/skill de la segunda columna
                                                string identificacion = Convert.ToString(range.Cells[row, 3].Value); // Obtener la identificación de conexión de la tercera columna
                                                string nombreAgente = Convert.ToString(range.Cells[row, 4].Value); // Obtener el nombre del agente de la cuarta columna
                                                int llamadasACD = Convert.ToInt32(range.Cells[row, 5].Value ?? 0); // Obtener el número de llamadas ACD de la quinta columna
                                                int llamadasAbandonadas = Convert.ToInt32(range.Cells[row, 6].Value ?? 0); // Obtener el número de llamadas abandonadas de la sexta columna
                                                double tiempoPromedioACD = Convert.ToDouble(range.Cells[row, 7].Value ?? 0); // Obtener el tiempo promedio ACD de la séptima columna
                                                double tiempoACD = Convert.ToDouble(range.Cells[row, 8].Value ?? 0); // Obtener el tiempo ACD de la octava columna
                                                int ACWOutCalls = Convert.ToInt32(range.Cells[row, 9].Value ?? 0); // Obtener ACWOUTCALLS de la novena columna
                                                double ACWOutTime = Convert.ToDouble(range.Cells[row, 10].Value ?? 0); // Obtener ACWOUTTIME de la décima columna
                                                int AUXOutCalls = Convert.ToInt32(range.Cells[row, 11].Value ?? 0); // Obtener AUXOUTCALLS de la undécima columna
                                                double AUXOutTime = Convert.ToDouble(range.Cells[row, 12].Value ?? 0); // Obtener AUXOUTTIME de la duodécima columna
                                                int extnOutCalls = Convert.ToInt32(range.Cells[row, 13].Value ?? 0); // Obtener Extn Out Calls de la decimotercera columna
                                                double extnOutTime = Convert.ToDouble(range.Cells[row, 14].Value ?? 0); // Obtener Extn Out Time de la decimocuarta columna
                                                double tiempoACW = Convert.ToDouble(range.Cells[row, 15].Value ?? 0); // Obtener el tiempo ACW de la decimoquinta columna
                                                double tiempoPromedioACW = Convert.ToDouble(range.Cells[row, 16].Value ?? 0); // Obtener el tiempo promedio ACW de la decimosexta columna
                                                double tiempoAUX = Convert.ToDouble(range.Cells[row, 17].Value ?? 0); // Obtener el tiempo AUX de la decimoséptima columna
                                                double tiempoLogueados = Convert.ToDouble(range.Cells[row, 18].Value ?? 0); // Obtener el tiempo logueados de la decimoctava columna
                                                double promedioConvHold = Convert.ToDouble(range.Cells[row, 19].Value ?? 0); // Obtener el promedio Conv + Hold de la decimonovena columna
                                                double tiempoHoldACDCalls = Convert.ToDouble(range.Cells[row, 20].Value ?? 0); // Obtener el tiempo de Hold/Acd Calls de la vigésima columna
                                                double tiempoHold = Convert.ToDouble(range.Cells[row, 21].Value ?? 0); // Obtener el tiempo Hold de la vigesimoprimera columna
                                                double porcentajeOcupacion = Convert.ToDouble(range.Cells[row, 22].Value ?? 0); // Obtener el porcentaje de ocupación de la vigesimosegunda columna

                                                double minutos = 60;
                                                double tiempoACD_1 = tiempoACD / minutos;
                                                double tiempo_llamadas_salida = AUXOutTime / minutos;
                                                double puntos_totales = Math.Round(tiempoACD_1 + tiempo_llamadas_salida, 2);    //MICI 
                                                double puntos_totales_totales = Math.Round((puntos_totales / 5.3), 2);

                                                string usuario = string.Empty; // Variable para almacenar el nombre de usuario

                                                // Buscamos al usuario que está creado en el diccionario de usuarios
                                                foreach (var kvp in usuarios_linea)
                                                {
                                                    if (kvp.Value == nombreAgente)
                                                    {
                                                        usuario = kvp.Key; // Obtenemos el nombre de usuario en minúsculas
                                                        break; // Salimos del bucle una vez encontrado el usuario
                                                    }

                                                }
                                                // Verificamos si se encontró el usuario
                                                if (string.IsNullOrEmpty(usuario))
                                                {
                                                    Console.WriteLine("Usuario no encontrado");
                                                }


                                                // Verificar si el nombre de usuario no está vacío
                                                if (!string.IsNullOrEmpty(usuario))
                                                {
                                                    // Agregar el nombre de usuario y el total al diccionario
                                                    if (!usuarioTotales.ContainsKey(usuario))
                                                        usuarioTotales.Add(usuario, puntos_totales);


                                                    else
                                                        usuarioTotales[usuario] += puntos_totales;
                                                }


                                            }
                                        }
                                        catch (Exception ex)
                                        {
                                            Console.WriteLine("Error al leer el archivo Excel: " + ex.Message);
                                        }
                                        finally
                                        {
                                            // Cerrar y liberar recursos
                                            workbook.Close();
                                            excelApp.Quit();
                                        }

                                        Console.ReadLine();


                                        foreach (var kvp in usuarios_linea)
                                        {
                                            string usuario = kvp.Key; // Obtenemos el nombre de usuario

                                            // Verificamos si el nombre de usuario está en el diccionario de totales
                                            if (usuarioTotales.ContainsKey(usuario))
                                            {
                                                double puntos_totales = usuarioTotales[usuario]; // Obtenemos los puntos totales para este usuario

                                                nuevoEmpleado = new EmpleadoDuo();
                                                nuevoEmpleado.Gundi = 56491;
                                                // Agregar los datos al nuevo empleado
                                                nuevoEmpleado.Cedula = 0;
                                                nuevoEmpleado.Usuario = usuario;
                                                nuevoEmpleado.Puntaje = Math.Round(puntos_totales, 2);
                                                nuevoEmpleado.Periodo = periodos;
                                                empleadosDuoList.Add(nuevoEmpleado);

                                                foreach (PuntajeDia item in PuntajeDia.productivo)
                                                {
                                                    if (item.Usuario.Contains(usuario))
                                                    {
                                                        item.Linea = Math.Round(puntos_totales, 2);
                                                    }
                                                }

                                            }


                                        }

                                        EmpleadoDuo.Empleo.AddRange(empleadosDuoList);
                                        estado = 3;
                                    }


                                    /////////   PRODUCTIVIDAD TRANSACCIONES CONTABLES  /////////////////////////////////////////////////////////////////////////
                                    if (estado == 3)
                                    {

                                        empleadosDuoList.Clear();
                                        foreach (DataRow row in DatosRespuestaTris.Dt.Rows)
                                        {
                                            foreach (var item in row.ItemArray)
                                            {
                                                Console.Write(item.ToString() + "\t");
                                            }
                                            Console.WriteLine();

                                            string fecha = row["MFECCAP"].ToString();
                                            string transacciones = row["MUSEAPR"].ToString();

                                            // Agrega los valores a las listas
                                            FechaDuo.Add(fecha);
                                            Transacciones_cib.Add(transacciones);

                                            // Obtiene el nombre de usuario de la transacción
                                            string[] partesTransaccion = transacciones.Split(new char[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);
                                            string usuario = partesTransaccion[0]; // Suponiendo que el usuario es la primera parte de la transacción

                                            // Verifica si el nombre de usuario existe en el diccionario de usuarios
                                            if (usuarios_transacciones.ContainsKey(usuario))
                                            {
                                                // Verifica si el nombre de usuario ya tiene una entrada en el diccionario de transacciones
                                                if (transaccionesPorUsuario.ContainsKey(usuarios_transacciones[usuario]))
                                                {
                                                    // Si ya tiene una entrada, incrementa la cantidad de transacciones para ese usuario
                                                    transaccionesPorUsuario[usuarios_transacciones[usuario]]++;
                                                }
                                                else
                                                {
                                                    // Si no tiene una entrada, agrega una nueva entrada para ese usuario con una transacción
                                                    transaccionesPorUsuario[usuarios_transacciones[usuario]] = 1;
                                                }
                                            }

                                        }

                                        foreach (var kvp in transaccionesPorUsuario)
                                        {
                                            nuevoEmpleado = new EmpleadoDuo();
                                            nuevoEmpleado.Cedula = 0;
                                            nuevoEmpleado.Gundi = 56496;
                                            nuevoEmpleado.Usuario = kvp.Key; // Usuario
                                            nuevoEmpleado.Puntaje = kvp.Value; // Cantidad de transacciones
                                            nuevoEmpleado.Periodo = periodos;


                                            // Agregar el empleado a la lista
                                            empleadosDuoList.Add(nuevoEmpleado);
                                            //Planta.Planta.Empleados_all.AddRange(empleadosDuoList);

                                            foreach (PuntajeDia item in PuntajeDia.productivo)
                                            {
                                                if (item.Usuario.Contains(kvp.Key))
                                                {
                                                    item.Contabilidad = kvp.Value;
                                                }
                                            }

                                        }
                                        EmpleadoDuo.Empleo.AddRange(empleadosDuoList);

                                        estado = 4;
                                    }


                                    //////   PRODUCTIVIDAD TRANSVERSALES   /////////////////////////////////////////////////////////////////////////
                                    if (estado == 4)
                                    {
                                        empleadosDuoList.Clear();
                                        int estado_transversal = 0;
                                        if (estado_transversal == 0)
                                        {
                                            empleadosDuoList.Clear();
                                            ////////    RECHAZOS ANY  /////////////////////////////////////////////////////////////////////////
                                            foreach (DataRow row in DatosRespuestaQuattuor.Dt.Rows)
                                            {
                                                string loginPlataforma = row["Usuario"].ToString().ToLower();
                                                int cantidadTransacciones = Convert.ToInt32(row["Cant Rechazos"]);

                                                if (loginPlataforma != "zona de totales")
                                                {
                                                    puntajes = new PuntajeDia();
                                                    nuevoEmpleado = new EmpleadoDuo();
                                                    rechazosANY[loginPlataforma] = cantidadTransacciones;
                                                    nuevoEmpleado.Gundi = 56287; // Este es el código para el resto de usuarios
                                                    nuevoEmpleado.Cedula = 0;
                                                    nuevoEmpleado.Usuario = loginPlataforma;
                                                    nuevoEmpleado.Puntaje = cantidadTransacciones;
                                                    nuevoEmpleado.Periodo = periodos;
                                                    // Agregar el nuevo empleado a la lista
                                                    empleadosDuoList.Add(nuevoEmpleado);

                                                    bool encontrado = false;
                                                    foreach (PuntajeDia item in PuntajeDia.productivo)
                                                    {
                                                        if (item.Usuario.Contains(loginPlataforma))
                                                        {
                                                            item.RechazoAny = cantidadTransacciones;
                                                            encontrado = true;
                                                            break;
                                                        }
                                                    }

                                                    if (!encontrado)
                                                    {
                                                        puntajes.Usuario = loginPlataforma;
                                                        puntajes.RechazoAny = cantidadTransacciones;
                                                        puntajeList.Add(puntajes);
                                                    }
                                                }
                                            }
                                            EmpleadoDuo.Empleo.AddRange(empleadosDuoList);
                                            PuntajeDia.productivo.AddRange(puntajeList);
                                            empleadosDuoList.Clear();

                                            foreach (DataRow row in DatosRespuestaQuattuor.Dt.Rows)
                                            {
                                                foreach (var item in row.ItemArray)
                                                {
                                                    Console.Write(item.ToString() + "\t");
                                                }
                                                Console.WriteLine();
                                                Console.WriteLine("Holiwi");
                                                string loginPlataforma = row["Usuario"].ToString().ToLower();
                                                int cantidadTransacciones_duo = Convert.ToInt32(row["Cant Validación"]);
                                                Console.WriteLine(loginPlataforma);
                                                Console.WriteLine();

                                                if (loginPlataforma != "zona de totales")
                                                {


                                                    nuevoEmpleado = new EmpleadoDuo();
                                                    validacionesANY[loginPlataforma] = cantidadTransacciones_duo;
                                                    nuevoEmpleado.Gundi = 56286; // Este es el código para el resto de usuarios
                                                    nuevoEmpleado.Cedula = 0;
                                                    nuevoEmpleado.Usuario = loginPlataforma;
                                                    nuevoEmpleado.Puntaje = cantidadTransacciones_duo;
                                                    nuevoEmpleado.Periodo = periodos;
                                                    // Agregar el nuevo empleado a la lista
                                                    empleadosDuoList.Add(nuevoEmpleado);

                                                    foreach (PuntajeDia item in PuntajeDia.productivo)
                                                    {
                                                        if (item.Usuario.Contains(loginPlataforma))
                                                        {
                                                            item.ValidacionAny = cantidadTransacciones_duo;
                                                        }
                                                    }
                                                }
                                            }
                                            EmpleadoDuo.Empleo.AddRange(empleadosDuoList);
                                            estado_transversal = 1;
                                        }

                                        ///////////    MODIFICACIONES   /////////////////////////////////////////////////////////////////////////
                                        if (estado_transversal == 1)
                                        {
                                            puntajeList.Clear();
                                            if (DatosRespuestaQuinque.Dt.Rows.Count <= 0)
                                            {
                                                Console.WriteLine("No hay datos");
                                            }
                                            else
                                            {
                                                foreach (DataRow row in DatosRespuestaQuinque.Dt.Rows)
                                                {
                                                    foreach (var item in row.ItemArray)
                                                    {
                                                        Console.Write(item.ToString() + "\t");
                                                    }
                                                    Console.WriteLine();

                                                    string loginPlataforma = row["USUARIO"].ToString();
                                                    int cantidadTransacciones = Convert.ToInt32(row["Cantidad"]);
                                                    //Console.WriteLine(kvp.Key);
                                                    Console.WriteLine(loginPlataforma);
                                                    Console.WriteLine("entrando");
                                                    if (usuarios_transversales.ContainsKey(loginPlataforma.Trim()))
                                                    {
                                                        string nombreUsuario = usuarios_transversales[loginPlataforma.Trim()];

                                                        modificaciones[nombreUsuario] = cantidadTransacciones;
                                                        nuevoEmpleado = new EmpleadoDuo();
                                                        nuevoEmpleado.Gundi = 56291; // Este es el código para el resto de usuarios
                                                        nuevoEmpleado.Cedula = 0;
                                                        nuevoEmpleado.Usuario = nombreUsuario;
                                                        nuevoEmpleado.Puntaje = cantidadTransacciones;
                                                        nuevoEmpleado.Periodo = periodos;
                                                        // Agregar el nuevo empleado a la lista
                                                        empleadosDuoList.Add(nuevoEmpleado);

                                                        foreach (PuntajeDia item in PuntajeDia.productivo)
                                                        {
                                                            if (item.Usuario.Contains(nombreUsuario))
                                                            {
                                                                item.Modificaciones = cantidadTransacciones;
                                                                Console.WriteLine(item.modificaciones);
                                                            }
                                                        }

                                                    }

                                                }
                                                EmpleadoDuo.Empleo.AddRange(empleadosDuoList);
                                            }
                                            estado_transversal = 2;
                                        }
                                        ///////////////    VALIDACIONES DE COINCIDENCIA /////////////////////////////////////////////////////////////////////////
                                        if (estado_transversal == 2)
                                        {
                                            empleadosDuoList.Clear();
                                            puntajeList.Clear();
                                            if (DatosRespuestaSix.Dt.Rows.Count <= 0)
                                            {
                                                Console.WriteLine("No hay datos");
                                            }
                                            else
                                            {
                                                foreach (DataRow row in DatosRespuestaSix.Dt.Rows)
                                                {
                                                    foreach (var item in row.ItemArray)
                                                    {
                                                        Console.Write(item.ToString() + "\t");
                                                    }
                                                    Console.WriteLine();

                                                    string loginPlataforma = row["USUARIO"].ToString();
                                                    int cantidadTransacciones = Convert.ToInt32(row["CANTIDAD"]);
                                                    // Crear un nuevo objeto Empleados_duo
                                                    if (usuarios_transversales.ContainsKey(loginPlataforma.Trim()))
                                                    {
                                                        string nombreUsuario = usuarios_transversales[loginPlataforma.Trim()];

                                                        validaciones[nombreUsuario] = cantidadTransacciones;
                                                        nuevoEmpleado = new EmpleadoDuo();
                                                        // Agregar los datos al nuevo empleado
                                                        nuevoEmpleado.Gundi = 56290; // Este es el código para el resto de usuarios
                                                        nuevoEmpleado.Cedula = 0;
                                                        nuevoEmpleado.Usuario = nombreUsuario;
                                                        nuevoEmpleado.Puntaje = cantidadTransacciones;
                                                        nuevoEmpleado.Periodo = periodos;
                                                        // Agregar el nuevo empleado a la lista
                                                        empleadosDuoList.Add(nuevoEmpleado);

                                                        foreach (PuntajeDia item in PuntajeDia.productivo)
                                                        {
                                                            if (item.Usuario.Contains(nombreUsuario))
                                                            {
                                                                item.Validacion_coincidencia = cantidadTransacciones;
                                                            }
                                                        }

                                                    }

                                                }
                                                EmpleadoDuo.Empleo.AddRange(empleadosDuoList);
                                            }
                                        }

                                        estado = 5;
                                    }


                                    ////////////// Consolidado final //////////////////////////////////
                                    if (estado == 5)
                                    {
                                        foreach (EmpleadoDuo planta in EmpleadoDuo.Empleo)
                                        {
                                            foreach (Planta plantita in Plantae)
                                            {
                                                if (planta.usuario.Contains(plantita.usuario))
                                                {
                                                    planta.cedula = plantita.cedula;
                                                    planta.departamento = plantita.departamento;

                                                }

                                            }

                                        }
                                        estado = 6;
                                    }

                                    if (estado == 6)
                                    {
                                        foreach (var empleado in EmpleadoDuo.Empleo)
                                        {
                                            if (empleado.departamento == "compraventa")
                                            {

                                                myexcelWorksheet.Cells[1, 1].Value = "gundi";
                                                myexcelWorksheet.Cells[1, 2].Value = "Cedula";
                                                myexcelWorksheet.Cells[1, 3].Value = "Usuario Red";
                                                myexcelWorksheet.Cells[1, 4].Value = "Periodo";


                                                // Agregar los días del mes como headers
                                                for (int i = 0; i < 31; i++)
                                                {
                                                    myexcelWorksheet.Cells[1, i + 5].Value = "Dia " + (i + 1);
                                                }

                                                myexcelWorksheet.Cells[rowunus, 1].Value = empleado.Gundi;
                                                myexcelWorksheet.Cells[rowunus, 2].Value = empleado.Cedula;
                                                myexcelWorksheet.Cells[rowunus, 3].Value = empleado.Usuario;
                                                myexcelWorksheet.Cells[rowunus, 4].Value = empleado.periodo;

                                                double puntajeTotal = empleado.Puntaje;

                                                int columnIndex = 5; // Comenzando desde la columna de "Dia 1"
                                                if (dia == "01" || dia == "1")
                                                {
                                                    int die = Convert.ToInt32(dia);
                                                    myexcelWorksheet.Cells[rowunus, columnIndex * die].Value = puntajeTotal;
                                                }
                                                else
                                                {
                                                    int die = Convert.ToInt32(dia);
                                                    myexcelWorksheet.Cells[rowunus, columnIndex + (die - 1)].Value = puntajeTotal;
                                                }
                                                rowunus++;
                                            }

                                        }
                                        string ProductividadCompraventa = @"\\sbmdebpmici01v\Files\mici\auto\" + Completa + Hora + "-331" + ".csv";
                                        //string ProductividadCartera = @"\\sbmdebpmici01v\Files\mici\auto\" + Completa + Hora + "-1648" + ".cvs";
                                        //myexcelApplication.ActiveWorkbook.SaveAs(@"D:\abc.xls", Excel.XlFileFormat.xlWorkbookNormal);
                                        myexcelApplication.ActiveWorkbook.SaveAs(ProductividadCompraventa, Excel.XlFileFormat.xlCSV);
                                        myexcelWorkbook.Close();
                                        myexcelApplication.Quit();

                                        ///// Archivo Plano 
                                        string archivoconfirmacion = @"\\sbmdebns03\VP SERV CORP\VP_SERV_CLIE\DIR_SERV_MERC_CAP_CCIO_INT\GCIA_SERV_COM_INT\Operación\Envio y recepción\Productividad\" + "Compraventa" + Completa + ".txt";
                                       
                                        using (StreamWriter writer = new StreamWriter(archivoconfirmacion))
                                        {
                                            string linea = $"Archivo enviado";
                                            writer.WriteLine(linea);
                                            writer.Flush();

                                        }

                                        estado = 7;
                                    }

                                    //GENERACION DE INFORME PARA POWER BI
                                    if (estado == 7)
                                    {
                                        empleadosDuoList.Clear();
                                        puntajeList.Clear();


                                        // Imprimir los datos antes de escribir en el archivo Excel
                                        Console.WriteLine("Datos de PuntajeDia.productivo:");
                                        foreach (PuntajeDia puntaje in PuntajeDia.productivo)
                                        {
                                            Console.WriteLine($"Usuario: {puntaje.Usuario}, Puntaje total: {puntaje.Puntaje_total}, Contabilidad: {puntaje.Contabilidad}");
                                        }

                                        try
                                        {
                                            string filePath = @"\\sbmdebns03\VP SERV CORP\VP_SERV_CLIE\DIR_SERV_MERC_CAP_CCIO_INT\GCIA_SERV_COM_INT\Operación\Envio y recepción\Productividad\ProductividadMICI.xlsx";

                                            Excel.Application excelAppDuo = new Excel.Application();
                                            Excel.Workbook workbookDuo;

                                            excelAppDuo.Visible = false; // No mostrar Excel durante el proceso
                                            excelAppDuo.DisplayAlerts = false; // Deshabilitar alertas

                                            if (File.Exists(filePath))
                                            {
                                                workbookDuo = excelAppDuo.Workbooks.Open(filePath);
                                            }
                                            else
                                            {
                                                workbookDuo = excelAppDuo.Workbooks.Add();
                                                workbookDuo.SaveAs(filePath);
                                            }

                                            // Obtener la hoja de trabajo "DB Principal" o crearla si no existe
                                            Microsoft.Office.Interop.Excel.Worksheet worksheetDuo = null;
                                            bool worksheetExists = false;
                                            foreach (Excel.Worksheet sheet in workbookDuo.Sheets)
                                            {
                                                if (sheet.Name == "DB Principal")
                                                {
                                                    worksheetDuo = sheet;
                                                    worksheetExists = true;
                                                    break;
                                                }
                                            }

                                            if (!worksheetExists)
                                            {
                                                Console.WriteLine("Alguien ha editado el archivo Excel. Corrija el nombre de la hoja, debe llamarse -DB Principal-");
                                            }
                                            else
                                            {
                                                // Encontrar la última fila utilizada en la hoja de cálculo
                                                int lastRow = worksheetDuo.Cells.SpecialCells(Microsoft.Office.Interop.Excel.XlCellType.xlCellTypeLastCell).Row;

                                                Excel.Range cellX = worksheetDuo.Columns[2];
                                                cellX.NumberFormat = "@";
                                                // Escribir los datos en la siguiente fila disponible
                                                foreach (PuntajeDia puntaje in PuntajeDia.productivo)
                                                {


                                                    worksheetDuo.Cells[lastRow + 1, 1].Value = dia;
                                                    worksheetDuo.Cells[lastRow + 1, 2].Value = fechanueva;
                                                    worksheetDuo.Cells[lastRow + 1, 4].Value = puntaje.Usuario;
                                                    worksheetDuo.Cells[lastRow + 1, 7].Value = puntaje.ValidacionAny;
                                                    worksheetDuo.Cells[lastRow + 1, 8].Value = puntaje.RechazoAny;
                                                    worksheetDuo.Cells[lastRow + 1, 9].Value = puntaje.Modificaciones;
                                                    worksheetDuo.Cells[lastRow + 1, 10].Value = puntaje.Validacion_coincidencia;
                                                    worksheetDuo.Cells[lastRow + 1, 11].Value = puntaje.Puntaje_total;
                                                    worksheetDuo.Cells[lastRow + 1, 12].Value = puntaje.Contabilidad;
                                                    worksheetDuo.Cells[lastRow + 1, 13].Value = puntaje.Linea;
                                                    worksheetDuo.Cells[lastRow + 1, 14].Value = puntaje.Puntaje_total;
                                                    worksheetDuo.Cells[lastRow + 1, 15].Value = puntaje.Puntaje_7;
                                                    worksheetDuo.Cells[lastRow + 1, 16].Value = puntaje.Puntaje_8;
                                                    worksheetDuo.Cells[lastRow + 1, 17].Value = puntaje.Puntaje_9;
                                                    worksheetDuo.Cells[lastRow + 1, 18].Value = puntaje.Puntaje_10;
                                                    worksheetDuo.Cells[lastRow + 1, 19].Value = puntaje.Puntaje_11;
                                                    worksheetDuo.Cells[lastRow + 1, 20].Value = puntaje.Puntaje_12;
                                                    worksheetDuo.Cells[lastRow + 1, 21].Value = puntaje.Puntaje_13;
                                                    worksheetDuo.Cells[lastRow + 1, 22].Value = puntaje.Puntaje_14;
                                                    worksheetDuo.Cells[lastRow + 1, 23].Value = puntaje.Puntaje_15;
                                                    worksheetDuo.Cells[lastRow + 1, 24].Value = puntaje.Puntaje_16;
                                                    worksheetDuo.Cells[lastRow + 1, 25].Value = puntaje.Puntaje_17;
                                                    worksheetDuo.Cells[lastRow + 1, 26].Value = puntaje.Puntaje_18;
                                                    worksheetDuo.Cells[lastRow + 1, 27].Value = puntaje.Puntaje_19;
                                                    worksheetDuo.Cells[lastRow + 1, 28].Value = puntaje.Puntaje_20;
                                                    worksheetDuo.Cells[lastRow + 1, 29].Value = puntaje.Puntaje_21;

                                                    lastRow++; // Avanzar a la siguiente fila

                                                }

                                                // Guardar el libro de trabajo
                                                workbookDuo.Save();

                                                Console.WriteLine("Los datos se han guardado en el archivo Excel correctamente.");
                                            }

                                            // Cerrar Excel y liberar recursos
                                            workbookDuo.Close();
                                            excelAppDuo.Quit();
                                        }
                                        catch (Exception ex)
                                        {
                                            Console.WriteLine(ex.Message);
                                        }
                                        estado = 8;
                                    }

                                    //POWER BI - PARTE 2

                                    if (estado == 8)
                                    {
                                        List<MICI> Mici_LIST = new List<MICI>();
                                        string filePath = @"\\sbmdebns03\VP SERV CORP\VP_SERV_CLIE\DIR_SERV_MERC_CAP_CCIO_INT\GCIA_SERV_COM_INT\Operación\Envio y recepción\Productividad\TableroGlobal.xls";
                                        // Inicializar una instancia de Excel
                                        Excel.Application ExcelAppTrIS = new Excel.Application();
                                        Excel.Workbook workbookTris = ExcelAppTrIS.Workbooks.Open(filePath);
                                        Excel.Worksheet worksheet = null;

                                        try
                                        {
                                            // Obtener la hoja de trabajo por nombre
                                            foreach (Excel.Worksheet sheet in workbookTris.Sheets)
                                            {
                                                if (sheet.Name == "TableroGlobal")
                                                {
                                                    worksheet = sheet;
                                                    break;
                                                }
                                            }

                                            if (worksheet == null)
                                            {
                                                throw new InvalidOperationException("La hoja TableroGlobal no existe en el archivo Excel.");
                                            }

                                            Excel.Range usedRange = worksheet.UsedRange;
                                            int rowCount = usedRange.Rows.Count;

                                            // Leer los datos desde la hoja de trabajo
                                            for (int row = 5; row <= rowCount; row++) // Comenzamos desde la fila 5
                                            {
                                                MICI mici = new MICI();

                                                mici.Numero = (usedRange.Cells[row, 1] as Excel.Range)?.Value?.ToString();
                                                mici.Cedula = Convert.ToInt32((usedRange.Cells[row, 2] as Excel.Range)?.Value);
                                                mici.Rol = (usedRange.Cells[row, 3] as Excel.Range)?.Value?.ToString();
                                                mici.Nombre = (usedRange.Cells[row, 4] as Excel.Range)?.Value?.ToString();
                                                mici.SemanasComprimidas = (usedRange.Cells[row, 5] as Excel.Range)?.Value?.ToString();
                                                mici.CursoVirtual = (usedRange.Cells[row, 6] as Excel.Range)?.Value?.ToString();
                                                mici.TrabajoVisible = (usedRange.Cells[row, 7] as Excel.Range)?.Value?.ToString();
                                                mici.Core = (usedRange.Cells[row, 8] as Excel.Range)?.Value?.ToString();
                                                mici.PlanesTrabajo = (usedRange.Cells[row, 9] as Excel.Range)?.Value?.ToString();
                                                mici.AsignacionAdministrativa = (usedRange.Cells[row, 10] as Excel.Range)?.Value?.ToString();
                                                mici.FallasInterrupciones = (usedRange.Cells[row, 11] as Excel.Range)?.Value?.ToString();
                                                mici.Ausentes = (usedRange.Cells[row, 12] as Excel.Range)?.Value?.ToString();
                                                mici.Ocupacion = (usedRange.Cells[row, 13] as Excel.Range)?.Value?.ToString();
                                                mici.FuenteDisponible = (usedRange.Cells[row, 14] as Excel.Range)?.Value?.ToString();
                                                mici.Productividad = (usedRange.Cells[row, 15] as Excel.Range)?.Value?.ToString();
                                                mici.Adherencia = (usedRange.Cells[row, 16] as Excel.Range)?.Value?.ToString();
                                                mici.RiesgoPsicosocial = (usedRange.Cells[row, 17] as Excel.Range)?.Value?.ToString();
                                                mici.Area = (usedRange.Cells[row, 18] as Excel.Range)?.Value?.ToString();
                                                mici.PlanesAccion = (usedRange.Cells[row, 19] as Excel.Range)?.Value?.ToString();

                                                Mici_LIST.Add(mici);
                                            }
                                            MICI.Mici.AddRange(Mici_LIST);

                                            // Imprimir los datos leídos para verificar
                                            foreach (MICI mici in MICI.Mici)
                                            {
                                                Console.WriteLine($"Nombre={mici.Nombre}, Numero={mici.Numero}, Cedula={mici.Cedula}, Rol={mici.Rol}, SemanasComprimidas={mici.SemanasComprimidas}");
                                            }
                                        }
                                        finally
                                        {
                                            // Cerrar y liberar recursos
                                            workbookTris.Close();
                                            ExcelAppTrIS.Quit();

                                        }
                                        estado = 9;

                                        foreach (MICI MICI in MICI.Mici)
                                        {
                                            foreach (Planta plantita in Planta.Empleados_all)
                                            {
                                                if (MICI.Cedula.Equals(plantita.cedula))
                                                {
                                                    MICI.Usuario = plantita.usuario;

                                                }

                                            }

                                        }


                                    }

                                    if (estado == 9)
                                    {

                                        try
                                        {
                                            string filePath = @"\\sbmdebns03\VP SERV CORP\VP_SERV_CLIE\DIR_SERV_MERC_CAP_CCIO_INT\GCIA_SERV_COM_INT\Operación\Envio y recepción\Productividad\ProductividadMICI.xlsx";

                                            Excel.Application excelAppDuo = new Excel.Application();
                                            Excel.Workbook workbookDuo;

                                            excelAppDuo.Visible = false; // No mostrar Excel durante el proceso
                                            excelAppDuo.DisplayAlerts = false; // Deshabilitar alertas

                                            if (File.Exists(filePath))
                                            {
                                                workbookDuo = excelAppDuo.Workbooks.Open(filePath);
                                            }
                                            else
                                            {
                                                workbookDuo = excelAppDuo.Workbooks.Add();
                                                workbookDuo.SaveAs(filePath);
                                            }

                                            // Obtener la hoja de trabajo "DB Principal" o crearla si no existe
                                            Microsoft.Office.Interop.Excel.Worksheet worksheetDuo = null;
                                            bool worksheetExists = false;
                                            foreach (Excel.Worksheet sheet in workbookDuo.Sheets)
                                            {
                                                if (sheet.Name == "MICI")
                                                {
                                                    worksheetDuo = sheet;
                                                    worksheetExists = true;
                                                    break;
                                                }
                                            }

                                            if (!worksheetExists)
                                            {
                                                Console.WriteLine("Alguien ha editado el archivo Excel. Corrija el nombre de la hoja, debe llamarse -DB Principal-");
                                            }
                                            else
                                            {
                                                // Encontrar la última fila utilizada en la hoja de cálculo
                                                int lastRow = worksheetDuo.Cells.SpecialCells(Microsoft.Office.Interop.Excel.XlCellType.xlCellTypeLastCell).Row;
                                                Excel.Range cellX = worksheetDuo.Columns[1];
                                                cellX.NumberFormat = "@";
                                                //  Escribir los datos en la siguiente fila disponible
                                                foreach (MICI mici in MICI.Mici)
                                                {
                                                    if (!mici.Cedula.Equals(0))
                                                    {

                                                        //Console.WriteLine(mici.Usuario);
                                                        worksheetDuo.Cells[lastRow + 1, 1].Value = fechanueva;
                                                        worksheetDuo.Cells[lastRow + 1, 2].Value = mici.Usuario;
                                                        worksheetDuo.Cells[lastRow + 1, 3].Value = mici.Cedula;
                                                        worksheetDuo.Cells[lastRow + 1, 4].Value = mici.Numero;
                                                        worksheetDuo.Cells[lastRow + 1, 5].Value = mici.Rol;
                                                        worksheetDuo.Cells[lastRow + 1, 6].Value = mici.Nombre;
                                                        worksheetDuo.Cells[lastRow + 1, 7].Value = mici.SemanasComprimidas;
                                                        worksheetDuo.Cells[lastRow + 1, 8].Value = mici.CursoVirtual;
                                                        worksheetDuo.Cells[lastRow + 1, 9].Value = mici.TrabajoVisible;
                                                        worksheetDuo.Cells[lastRow + 1, 10].Value = mici.Core;
                                                        worksheetDuo.Cells[lastRow + 1, 11].Value = mici.PlanesTrabajo;
                                                        worksheetDuo.Cells[lastRow + 1, 12].Value = mici.AsignacionAdministrativa;
                                                        worksheetDuo.Cells[lastRow + 1, 13].Value = mici.FallasInterrupciones;
                                                        worksheetDuo.Cells[lastRow + 1, 14].Value = mici.Ausentes;
                                                        worksheetDuo.Cells[lastRow + 1, 15].Value = mici.Ocupacion;
                                                        worksheetDuo.Cells[lastRow + 1, 16].Value = mici.FuenteDisponible;
                                                        worksheetDuo.Cells[lastRow + 1, 17].Value = mici.Productividad;
                                                        worksheetDuo.Cells[lastRow + 1, 18].Value = mici.Adherencia;
                                                        worksheetDuo.Cells[lastRow + 1, 19].Value = mici.RiesgoPsicosocial;
                                                        worksheetDuo.Cells[lastRow + 1, 20].Value = mici.Area;
                                                        worksheetDuo.Cells[lastRow + 1, 21].Value = mici.PlanesAccion;

                                                        lastRow++; // Avanzar a la siguiente fila
                                                    }



                                                }

                                                // Guardar el libro de trabajo
                                                workbookDuo.Save();


                                                Console.WriteLine("Los datos se han guardado en el archivo Excel correctamente.");
                                            }

                                            // Cerrar Excel y liberar recursos
                                            workbookDuo.Close();
                                            excelAppDuo.Quit();
                                        }
                                        catch (Exception ex)
                                        {
                                            Console.WriteLine(ex.Message);
                                        }
                                    }

                                }

                            }

                        }
                        else if (gerencia.Contains("Trade") && ok.ElementAt(2).Equals(0) && !File.Exists(archivoTris))
                        {


                            // Inicialización de la primera instancia de Excel Application
                            Excel.Application myexcelApplication = new Excel.Application();

                            // Inicialización de la segunda instancia de Excel Application
                            Excel.Application myexcelApplication_duo = new Excel.Application();

                            // Inicialización de la segunda instancia de Excel Application
                            Excel.Application myexcelApplication_tris = new Excel.Application();

                            Excel.Workbook myexcelWorkbook = myexcelApplication.Workbooks.Add();
                            Excel.Worksheet myexcelWorksheet = (Excel.Worksheet)myexcelWorkbook.Sheets.Add();
                            EmpleadoDuo nuevoEmpleado;
                            PuntajeDia puntajes;
                            List<EmpleadoDuo> empleadosDuoList = new List<EmpleadoDuo>();
                            List<PuntajeDia> puntajeList = new List<PuntajeDia>();
                            int rowunus = 2;
                            int estado = 0;

                            if (DiaActualSemana.ToString() == diaSemana)
                            {
                                DatosRespuestaDuo = Negocio.ProductividadGerencia.GenerarProductividadGerencia.ConsultarProductividadGlobal(-1, FechaLunes, 0);
                                periodos = Hoy.AddDays(diaLimiteInferior).ToString("yyyyMM");
                                dia = Hoy.AddDays(diaLimiteInferior).ToString("dd");
                                Console.WriteLine("Datos consultados Lunes");

                            }
                            else
                            {
                                DatosRespuestaDuo = Negocio.ProductividadGerencia.GenerarProductividadGerencia.ConsultarProductividadGlobal(-1, Fecha, 0);
                                periodos = Hoy.AddDays(diaLimiteSuperior).ToString("yyyyMM");
                                dia = Hoy.AddDays(diaLimiteSuperior).ToString("dd");
                                Console.WriteLine("Datos consultados");
                            }

                            Console.WriteLine(DatosRespuestaDuo.Dt.Rows.Count);

                            if (!DatosRespuestaDuo.HayFallos)
                            {
                                if (DatosRespuestaDuo.Dt.Rows.Count > 0)
                                {
                                    Console.WriteLine("Aqui");
                                    if (estado == 0)
                                    {
                                        foreach (DataRow row in DatosRespuestaDuo.Dt.Rows)
                                        {
                                            foreach (var item in row.ItemArray)
                                            {
                                                Console.Write(item.ToString() + "\t");
                                            }
                                            Console.WriteLine();

                                            string usuario = row["Usuario"].ToString().ToLower();
                                            double total = Convert.ToDouble(row["Total"]);

                                            if (usuario != "zona de totales")
                                            {
                                                if (trade.Contains(usuario))
                                                {
                                                    // Agregar los datos al nuevo empleado
                                                    nuevoEmpleado = new EmpleadoDuo();
                                                    nuevoEmpleado.Gundi = 18431; // Este es el código para el resto de usuarios
                                                    nuevoEmpleado.Cedula = 0;
                                                    nuevoEmpleado.Usuario = usuario;
                                                    nuevoEmpleado.Puntaje = total;
                                                    nuevoEmpleado.Periodo = periodos;

                                                    // Agregar el nuevo empleado a la lista
                                                    empleadosDuoList.Add(nuevoEmpleado);
                                                }
                                            }
                                        }
                                        EmpleadoDuo.Empleo.AddRange(empleadosDuoList);
                                        estado = 1;
                                    }

                                    if (estado == 1)
                                    {
                                        foreach (EmpleadoDuo planta in EmpleadoDuo.Empleo)
                                        {
                                            foreach (Planta plantita in Plantae)
                                            {
                                                if (planta.usuario.Contains(plantita.usuario))
                                                {
                                                    planta.cedula = plantita.cedula;
                                                    planta.departamento = plantita.departamento;

                                                }

                                            }

                                        }
                                        estado = 2;
                                    }

                                    if (estado == 2)
                                    {
                                        foreach (var empleado in EmpleadoDuo.Empleo)
                                        {
                                            myexcelWorksheet.Cells[1, 1].Value = "gundi";
                                            myexcelWorksheet.Cells[1, 2].Value = "Cedula";
                                            myexcelWorksheet.Cells[1, 3].Value = "Usuario Red";
                                            myexcelWorksheet.Cells[1, 4].Value = "Periodo";


                                            // Agregar los días del mes como headers
                                            for (int i = 0; i < 31; i++)
                                            {
                                                myexcelWorksheet.Cells[1, i + 5].Value = "Dia " + (i + 1);
                                            }

                                            if (empleado.departamento == "trade")
                                            {
                                                myexcelWorksheet.Cells[rowunus, 1].Value = empleado.Gundi;
                                                myexcelWorksheet.Cells[rowunus, 2].Value = empleado.Cedula;
                                                myexcelWorksheet.Cells[rowunus, 3].Value = empleado.Usuario;
                                                myexcelWorksheet.Cells[rowunus, 4].Value = empleado.periodo;

                                                double puntajeTotal = empleado.Puntaje;

                                                int columnIndex = 5; // Comenzando desde la columna de "Dia 1"
                                                if (dia == "01" || dia == "1")
                                                {
                                                    int die = Convert.ToInt32(dia);
                                                    myexcelWorksheet.Cells[rowunus, columnIndex * die].Value = puntajeTotal;
                                                }
                                                else
                                                {
                                                    int die = Convert.ToInt32(dia);
                                                    myexcelWorksheet.Cells[rowunus, columnIndex + (die - 1)].Value = puntajeTotal;
                                                }
                                                rowunus++;
                                            }

                                        }

                                        string ProductividadCartera = @"\\sbmdebpmici01v\Files\mici\auto\" + Completa + Hora + "-221" + ".csv";
                                        //myexcelApplication.ActiveWorkbook.SaveAs(@"D:\abc.xls", Excel.XlFileFormat.xlWorkbookNormal);
                                        myexcelApplication.ActiveWorkbook.SaveAs(ProductividadCartera, Excel.XlFileFormat.xlCSV);
                                        myexcelWorkbook.Close();
                                        myexcelApplication.Quit();

                                        ///// Archivo Plano 
                                        string archivoconfirmacion = @"\\sbmdebns03\VP SERV CORP\VP_SERV_CLIE\DIR_SERV_MERC_CAP_CCIO_INT\GCIA_SERV_COM_INT\Operación\Envio y recepción\Productividad\" + "Trade" + Completa + ".txt";

                                        using (StreamWriter writer = new StreamWriter(archivoconfirmacion))
                                        {
                                            string linea = $"Archivo enviado";
                                            writer.WriteLine(linea);
                                            writer.Flush();

                                        }

                                    }

                                }

                            }

                        }


                    } //aqui 


                    #region EN ESPERA
                    //Reporte Monitor
                    AMP.InsertarReporteMonApli(IdAplicacion, EstadosAplicacion.EnEspera);
                    app.AgregarLog(app.NombreAplicacion, EstadosAplicacion.EnEspera, app.Intervalo.ToString());
                    AMP.PausaSys(app.Intervalo * 60);
                    #endregion

                }
                else
                {
                    #region APP INACTIVA
                    app.AgregarLog(app.NombreAplicacion, EstadosAplicacion.AppInactiva, app.Intervalo.ToString());
                    AMP.InsertarReporteMonApli(IdAplicacion, EstadosAplicacion.EnEspera);
                    AMP.PausaSys(app.Intervalo * 60);
                    #endregion
                }
            }

            AMP.InsertarReporteMonApli(IdAplicacion, EstadosAplicacion.Finalizando);
            app.AgregarLog(app.NombreAplicacion, EstadosAplicacion.Finalizando, app.Intervalo.ToString());
            app.EnEjecucion = false;

        }


        private static Respuesta EnviarNotificacionExcelencia(string nombre, string correo, string cuerpo)
        {
            #region Core notificaciones
            CoreNotGenerales objCoreNot = new CoreNotGenerales();
            clsCoreNotificaciones datosCore = objCoreNot.ConsultarInfoCoreNot
                                                    (168, "App", -1);
            #endregion Core notificaciones

            if (datosCore.IdMaestro <= 0)
            {
                //detener y mostrar error
                string error = "La app no esta parametrizada en el core de notificaciones";
                return new Respuesta { HayFallos = true, Mensajes = new List<string> { error } };
            }

            datosCore.Nit = "0";
            datosCore.UsuarioCrea = "joarios";
            datosCore.NombreCliente = nombre;
            datosCore.Destinatarios = correo;
            datosCore.DestinatariosCC = " ";
            //datosCore.Destinatarios = "andalzat@bancolombia.com.co";
            datosCore.Cuerpo = datosCore.Cuerpo.Replace("@notificacion", cuerpo);
            datosCore.CodEstadoNotificacion = 2;



            Respuesta DatosRespuesta = objCoreNot.InsertarDatosCoreNot(datosCore);

            if (DatosRespuesta.HayFallos)
            {
                return DatosRespuesta;
            }
            else if (DatosRespuesta.Dt.Rows.Count > 0)
            {
                string idCore = DatosRespuesta.Dt.Rows[0]["Column1"].ToString();

                return new Respuesta { HayFallos = false, Mensajes = new List<string> { idCore } };
            }
            else
            {
                return new Respuesta { HayFallos = true, Mensajes = new List<string> { "El core no devolvio un id valido" } };
            }
        }



    }

    public class ErrorUsuario
    {
        public string Usuario { get; set; }
        public string Errores { get; set; }
    }

}
