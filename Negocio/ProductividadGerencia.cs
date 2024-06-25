using DocumentFormat.OpenXml.Office2010.Excel;
using Negocio.Generales;
using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Transversal;


namespace Negocio.ProductividadGerencia
{
    public class GenerarProductividadGerencia
    {
        public static Respuesta Usuarios()
        {
            Respuesta DatosRespuesta = new Respuesta();
            ParametrosND DatosEnvio = new ParametrosND();
            DatosEnvio.SP = SPs.UsuariosProductividades;
            DatosEnvio.CadenaBD = CadenasBD.GESTION_COMERCIO;

            try
            {
                DatosRespuesta = Datos.Conexion.EjecutarSP(DatosEnvio);
            }
            catch (Exception ex)
            {
                DatosRespuesta.HayFallos = true;
                DatosRespuesta.Mensajes.Add("Negocio - Error en [" + DatosEnvio.SP + "] - Mensaje: " + ex.Message);
            }
            return DatosRespuesta;
        }

        public static Respuesta ConsultarDiasNoHabiles()
        {
            Respuesta DatosRespuesta = new Respuesta();
            ParametrosND DatosEnvio = new ParametrosND();
            DatosEnvio.SP = SPs.ConsultarDiaHabiles;
            DatosEnvio.CadenaBD = CadenasBD.GESTION_COMERCIO;
            try
            {
                DatosRespuesta = Datos.Conexion.EjecutarSP(DatosEnvio);

            }
            catch (Exception ex)
            {
                DatosRespuesta.HayFallos = true;

                DatosRespuesta.Mensajes.Add("[" + (ex.StackTrace.Substring(ex.StackTrace.LastIndexOf('\n') + 7, (ex.StackTrace.LastIndexOf('(') - ex.StackTrace.LastIndexOf('\n') - 7))) + ": linea" + ex.StackTrace.Substring(ex.StackTrace.LastIndexOf(' ')) + "](Excep: " + ex.Message + ")");
            }
            return DatosRespuesta;
        }



        public static Respuesta ConsultarProductividadGlobal(int codeunus, string Fecha, int codeduo)
        {
            Respuesta DatosRespuesta = new Respuesta();
            ParametrosND DatosEnvio = new ParametrosND();
            DatosEnvio.SP = SPs.ConsultarProductividadGlobal;
            DatosEnvio.CadenaBD = CadenasBD.GESTION_COMERCIO;
            try
            {
                DatosEnvio.NomParametros.Add("@CodArea");
                DatosEnvio.Datos.Add(codeunus);
                DatosEnvio.NomParametros.Add("@Fecha");
                DatosEnvio.Datos.Add(Fecha);
                DatosEnvio.NomParametros.Add("@SwTP");
                DatosEnvio.Datos.Add(codeduo);

                DatosRespuesta = Datos.Conexion.EjecutarSP(DatosEnvio);
            }
            catch (Exception ex)
            {
                DatosRespuesta.HayFallos = true;
                StackTrace stack = new StackTrace();
                DatosRespuesta.Mensajes.Add("[ERROR][" + stack.ToString().Split('\n')[0] + "][EXSYS][" + ex.Message + "]");
            }
            return DatosRespuesta;

        }

        public static Respuesta ConsultarProductividad(string Fecha)
        {
            Respuesta DatosRespuesta = new Respuesta();
            ParametrosND DatosEnvio = new ParametrosND();
            DatosEnvio.SP = SPs.ConsultarProductividadGlobalApro;
            DatosEnvio.CadenaBD = CadenasBD.GESTION_COMERCIO;
            DatosEnvio.Nacional = true;

            try
            {
                DatosEnvio.NomParametros.Add("@Fecha");
                DatosEnvio.Datos.Add(Int64.Parse(Fecha));

                DatosRespuesta = Datos.Conexion.EjecutarSP(DatosEnvio);

            }
            catch (Exception ex)
            {
                DatosRespuesta.HayFallos = true;
                DatosRespuesta.Mensajes.Add("[" + (ex.StackTrace.Substring(ex.StackTrace.LastIndexOf('\n') + 7, (ex.StackTrace.LastIndexOf('(') - ex.StackTrace.LastIndexOf('\n') - 7))) + ": linea" + ex.StackTrace.Substring(ex.StackTrace.LastIndexOf(' ')) + "](Excep: " + ex.Message + ")");
            }
            return DatosRespuesta;
        }

        public static Respuesta ConsultarPlanta()
        {
            Respuesta DatosRespuesta = new Respuesta();
            ParametrosND DatosEnvio = new ParametrosND();
            DatosEnvio.SP = SPs.ConsultarPlanta;
            DatosEnvio.CadenaBD = CadenasBD.CORE_PLATAFORMA;
            try
            {
                DatosRespuesta = Datos.Conexion.EjecutarSP(DatosEnvio);

            }
            catch (Exception ex)
            {
                DatosRespuesta.HayFallos = true;

                DatosRespuesta.Mensajes.Add("[" + (ex.StackTrace.Substring(ex.StackTrace.LastIndexOf('\n') + 7, (ex.StackTrace.LastIndexOf('(') - ex.StackTrace.LastIndexOf('\n') - 7))) + ": linea" + ex.StackTrace.Substring(ex.StackTrace.LastIndexOf(' ')) + "](Excep: " + ex.Message + ")");
            }
            return DatosRespuesta;
        }

        public static Respuesta ConsultarProductividadTransacciones(string Fecha)
        {
            Respuesta DatosRespuesta = new Respuesta();
            ParametrosND DatosEnvio = new ParametrosND();
            DatosEnvio.SP = SPs.ConsultarTransaccionesContables;
            DatosEnvio.CadenaBD = CadenasBD.GESTION_COMERCIO;
            DatosEnvio.Nacional = true;

            try
            {
                DatosEnvio.NomParametros.Add("@Fecha");
                DatosEnvio.Datos.Add(Int64.Parse(Fecha));

                DatosRespuesta = Datos.Conexion.EjecutarSP(DatosEnvio);

            }
            catch (Exception ex)
            {
                DatosRespuesta.HayFallos = true;
                DatosRespuesta.Mensajes.Add("[" + (ex.StackTrace.Substring(ex.StackTrace.LastIndexOf('\n') + 7, (ex.StackTrace.LastIndexOf('(') - ex.StackTrace.LastIndexOf('\n') - 7))) + ": linea" + ex.StackTrace.Substring(ex.StackTrace.LastIndexOf(' ')) + "](Excep: " + ex.Message + ")");
            }
            return DatosRespuesta;
        }

        public static Respuesta ConsultarProductividadRechazos(DateTime Fecha)
        {
            Respuesta DatosRespuesta = new Respuesta();
            ParametrosND DatosEnvio = new ParametrosND();
            DatosEnvio.SP = SPs.ConsultarRechazosRemesas;
            DatosEnvio.CadenaBD = CadenasBD.GESTION_COMERCIO;

            try
            {
                DatosEnvio.NomParametros.Add("@FechaInicial");
                DatosEnvio.Datos.Add((Fecha));
                DatosEnvio.NomParametros.Add("@FechaFinal");
                DatosEnvio.Datos.Add((Fecha));

                DatosRespuesta = Datos.Conexion.EjecutarSP(DatosEnvio);

            }
            catch (Exception ex)
            {
                DatosRespuesta.HayFallos = true;
                DatosRespuesta.Mensajes.Add("[" + (ex.StackTrace.Substring(ex.StackTrace.LastIndexOf('\n') + 7, (ex.StackTrace.LastIndexOf('(') - ex.StackTrace.LastIndexOf('\n') - 7))) + ": linea" + ex.StackTrace.Substring(ex.StackTrace.LastIndexOf(' ')) + "](Excep: " + ex.Message + ")");
            }
            return DatosRespuesta;
        }

        public static Respuesta ModificacionUsuario(string FechaInicial, string FechaFinal)
        {
            Respuesta DatosRespuesta = new Respuesta();
            ParametrosND DatosEnvio = new ParametrosND();
            DatosEnvio.SP = SPs.ConsultarModificacionesRemesas;
            DatosEnvio.CadenaBD = CadenasBD.GESTION_COMERCIO;
            DatosEnvio.Nacional = true;

            try
            {
                DatosEnvio.NomParametros.Add("@FechaInicial");
                DatosEnvio.Datos.Add(FechaInicial);

                DatosEnvio.NomParametros.Add("@FechaFinal");
                DatosEnvio.Datos.Add(FechaFinal);

                DatosRespuesta = Datos.Conexion.EjecutarSP(DatosEnvio);
            }
            catch (Exception ex)
            {
                DatosRespuesta.HayFallos = true;
                DatosRespuesta.Mensajes.Add("[" + (ex.StackTrace.Substring(ex.StackTrace.LastIndexOf('\n') + 7, (ex.StackTrace.LastIndexOf('(') - ex.StackTrace.LastIndexOf('\n') - 7))) + ": linea" + ex.StackTrace.Substring(ex.StackTrace.LastIndexOf(' ')) + "](Excep: " + ex.Message + ")");
            }
            return DatosRespuesta;
        }


        public static Respuesta ReprocesosInsert(
          DateTime fechaOp,
          DateTime fechaReg,
          long consecutivo,
          long nit,
          int moneda,
          decimal valor,
          string cliente,
          int prodEvento,
          string responsable,
          string responsableDos,
          string usuarioReproceso,
          int area,
          string tipoError,
          string causa,
          string descripcion,
          int queja,
          int perdida,
          int impacto,
          DateTime fechaperdida,
          int segmento)
        {
            Respuesta DatosRespuesta = new Respuesta();
            ParametrosND DatosEnvio = new ParametrosND();
            DatosEnvio.SP = SPs.ReprocesosInsert;
            DatosEnvio.CadenaBD = CadenasBD.GESTION_COMERCIO;

            try
            {
                DatosEnvio.NomParametros.Add("@FechaOp");
                DatosEnvio.Datos.Add(fechaOp.Date);
                DatosEnvio.NomParametros.Add("@FechaReg");
                DatosEnvio.Datos.Add(fechaReg.Date);
                DatosEnvio.NomParametros.Add("@Consecutivo");
                if (consecutivo.ToString().Length > 15)
                {
                    consecutivo = Convert.ToInt64(consecutivo.ToString().Substring(0, 15));
                }
                DatosEnvio.Datos.Add(consecutivo);
                DatosEnvio.NomParametros.Add("@NitCliente");
                DatosEnvio.Datos.Add(nit);
                DatosEnvio.NomParametros.Add("@CodMoneda");
                DatosEnvio.Datos.Add(moneda);
                DatosEnvio.NomParametros.Add("@Monto");
                DatosEnvio.Datos.Add(valor);
                DatosEnvio.NomParametros.Add("@NombreCliente");
                DatosEnvio.Datos.Add(cliente);
                DatosEnvio.NomParametros.Add("@IdProductoEvento");
                DatosEnvio.Datos.Add(prodEvento);
                DatosEnvio.NomParametros.Add("@UsuarioResponsable");
                DatosEnvio.Datos.Add(responsable);
                DatosEnvio.NomParametros.Add("@UsuarioResponsableDos");
                DatosEnvio.Datos.Add(responsableDos);
                DatosEnvio.NomParametros.Add("@UsuarioCrea");
                DatosEnvio.Datos.Add(usuarioReproceso);
                DatosEnvio.NomParametros.Add("@CodArea");
                DatosEnvio.Datos.Add(area);
                DatosEnvio.NomParametros.Add("@TipoError");
                DatosEnvio.Datos.Add(tipoError);
                DatosEnvio.NomParametros.Add("@Causa");
                DatosEnvio.Datos.Add(causa);
                DatosEnvio.NomParametros.Add("@Descripcion");
                DatosEnvio.Datos.Add(descripcion);
                DatosEnvio.NomParametros.Add("@QuejaCliente");
                DatosEnvio.Datos.Add(queja);
                DatosEnvio.NomParametros.Add("@Perdida");
                DatosEnvio.Datos.Add(perdida);
                DatosEnvio.NomParametros.Add("@Impacto");
                DatosEnvio.Datos.Add(impacto);
                DatosEnvio.NomParametros.Add("@FechaPerdida");
                DatosEnvio.Datos.Add(fechaperdida.Date);
                DatosEnvio.NomParametros.Add("@IdSegmento");
                DatosEnvio.Datos.Add(segmento);

                DatosRespuesta = Datos.Conexion.EjecutarSP(DatosEnvio);

            }
            catch (Exception ex)
            {
                DatosRespuesta.HayFallos = true;
                DatosRespuesta.Mensajes.Add("[" + (ex.StackTrace.Substring(ex.StackTrace.LastIndexOf('\n') + 7, (ex.StackTrace.LastIndexOf('(') - ex.StackTrace.LastIndexOf('\n') - 7))) + ": linea" + ex.StackTrace.Substring(ex.StackTrace.LastIndexOf(' ')) + "](Excep: " + ex.Message + ")");
            }
            if (DatosRespuesta.HayFallos == true)
            {
                return DatosRespuesta;
            }
            else
            {
                return DatosRespuesta;
            }


        }

        public static Respuesta ConsultarInfoNroDocAny(string Nit)
        {
            Respuesta DatosRespuesta = new Respuesta();
            ParametrosND DatosEnvio = new ParametrosND();

            DatosEnvio.SP = SPs.ConsultarInfoNroDocAny;
            DatosEnvio.CadenaBD = CadenasBD.GESTION_COMERCIO;

            try
            {
                DatosEnvio.NomParametros.Add("@Nit");
                DatosEnvio.Datos.Add(Int64.Parse(Nit));

                DatosRespuesta = Datos.Conexion.EjecutarSP(DatosEnvio);
            }
            catch (Exception ex)
            {
                DatosRespuesta.HayFallos = true;
                DatosRespuesta.Mensajes.Add("[" + (ex.StackTrace.Substring(ex.StackTrace.LastIndexOf('\n') + 7, (ex.StackTrace.LastIndexOf('(') - ex.StackTrace.LastIndexOf('\n') - 7))) + ": linea" + ex.StackTrace.Substring(ex.StackTrace.LastIndexOf(' ')) + "](Excep: " + ex.Message + ")");
            }
            return DatosRespuesta;

        }

        public static string ConsecutivoBuqueda(string consecutivoMagnus, string fecha)
        {
            try
            {
                Respuesta DatosRespuesta = new Respuesta();
                Respuesta DatosRespuestaDuo = new Respuesta();

                if (consecutivoMagnus == "" || consecutivoMagnus.Length == 11 || consecutivoMagnus.Length == 8 || consecutivoMagnus.Length == 3)
                {
                    return "0";
                }
                else
                {

                    if (consecutivoMagnus.Length == 15)
                    {
                        string fechamain = consecutivoMagnus.Substring(0, 8);
                        string code = consecutivoMagnus.Substring(8, 3);
                        string consecutivo = consecutivoMagnus.Substring(11);
                        DatosRespuesta = ConsecutivoReproceso(consecutivo, fechamain, code);

                        if (!DatosRespuesta.HayFallos)
                        {
                            if (DatosRespuesta.Dt.Rows.Count > 0)
                            {
                                foreach (DataRow item2 in DatosRespuesta.Dt.Rows)
                                {
                                    try
                                    {
                                        string producto = item2["PRODUCTO"].ToString().Trim();
                                        string evento = item2["EVENTO"].ToString().Trim();

                                        string productoEvento = $"[{producto}-{evento}]";

                                        DatosRespuestaDuo = ProductoEvento();

                                        foreach (DataRow item in DatosRespuestaDuo.Dt.Rows)
                                        {
                                            string valor = item["Valor"].ToString().Trim();
                                            string codigo = item["Codigo"].ToString().Trim();

                                            if (valor.Contains(productoEvento))
                                            {
                                                return codigo;
                                            }

                                        }


                                    }
                                    catch (Exception ex)
                                    {

                                    }
                                }

                            }
                            else
                            {

                            }
                        }
                        else
                        {

                        }
                    }
                    else if(consecutivoMagnus.Length == 4)
                    {
                        
                        string code = "974";
                        string consecutivo = consecutivoMagnus;
                        DatosRespuesta = ConsecutivoReproceso(consecutivo, fecha, code);


                        if (!DatosRespuesta.HayFallos)
                        {
                            if (DatosRespuesta.Dt.Rows.Count > 0)
                            {
                                foreach (DataRow item2 in DatosRespuesta.Dt.Rows)
                                {
                                    try
                                    {
                                        string producto = item2["PRODUCTO"].ToString().Trim();
                                        string evento = item2["EVENTO"].ToString().Trim();

                                        string productoEvento = $"[{producto}-{evento}]";

                                        DatosRespuestaDuo = ProductoEvento();

                                        foreach (DataRow item in DatosRespuestaDuo.Dt.Rows)
                                        {
                                            string valor = item["Valor"].ToString().Trim();
                                            string codigo = item["Codigo"].ToString().Trim();

                                            if (valor.Contains(productoEvento))
                                            {
                                                return codigo;
                                            }

                                        }

                                    }
                                    catch (Exception ex)
                                    {

                                    }
                                }
                               
                            }
                            else
                            {
                               
                            }
                        }
                        else
                        {
                           
                        }

                    }
                    else if (consecutivoMagnus.Length == 26)
                    {
                        string ultimos15 = consecutivoMagnus.Substring(consecutivoMagnus.Length - 15);
                        fecha = consecutivoMagnus.Substring(0, 8);
                        string code = "974";
                        string consecutivo = consecutivoMagnus.Substring(11);
                        DatosRespuesta = ConsecutivoReproceso(consecutivo, fecha, code);


                        if (!DatosRespuesta.HayFallos)
                        {
                            if (DatosRespuesta.Dt.Rows.Count > 0)
                            {
                                foreach (DataRow item2 in DatosRespuesta.Dt.Rows)
                                {
                                    try
                                    {
                                        string producto = item2["PRODUCTO"].ToString().Trim();
                                        string evento = item2["EVENTO"].ToString().Trim();

                                        string productoEvento = $"[{producto}-{evento}]";

                                        DatosRespuestaDuo = ProductoEvento();

                                        foreach (DataRow item in DatosRespuestaDuo.Dt.Rows)
                                        {
                                            string valor = item["Valor"].ToString().Trim();
                                            string codigo = item["Codigo"].ToString().Trim();

                                            if (valor.Contains(productoEvento))
                                            {
                                                return codigo;
                                            }

                                        }

                                    }
                                    catch (Exception ex)
                                    {

                                    }
                                }

                            }
                            else
                            {

                            }
                        }
                        else
                        {

                        }

                    }
                    else
                    {
                        return "0";
                    }

                }



            }
            catch (Exception ex)
            {
               
            }
            return "0";
        }

        public static Respuesta ConsecutivoReproceso(string consecutivo, string fecha, string code)
        {
            string ConsecutivoReproceso = "ConsecutivoReproceso";
            Respuesta DatosRespuesta = new Respuesta();
            ParametrosND DatosEnvio = new ParametrosND();
            DatosEnvio.SP = ConsecutivoReproceso;
            DatosEnvio.CadenaBD = CadenasBD.GESTION_COMERCIO;
            DatosEnvio.Nacional = true;

            try
            {
                DatosEnvio.NomParametros.Add("@consecutivo");
                DatosEnvio.Datos.Add(consecutivo);

                DatosEnvio.NomParametros.Add("@fecha");
                DatosEnvio.Datos.Add(fecha);

                DatosEnvio.NomParametros.Add("@code");
                DatosEnvio.Datos.Add(code);


                DatosRespuesta = Datos.Conexion.EjecutarSP(DatosEnvio);
            }
            catch (Exception ex)
            {
                DatosRespuesta.HayFallos = true;
                DatosRespuesta.Mensajes.Add("[" + (ex.StackTrace.Substring(ex.StackTrace.LastIndexOf('\n') + 7, (ex.StackTrace.LastIndexOf('(') - ex.StackTrace.LastIndexOf('\n') - 7))) + ": linea" + ex.StackTrace.Substring(ex.StackTrace.LastIndexOf(' ')) + "](Excep: " + ex.Message + ")");
            }
            return DatosRespuesta;
        }

        public static Respuesta ReprocesosOmniaFechaAdmin(DateTime fechaini, DateTime fechafin, int area)
        {
            string Reproceso = "pro.MostrarReprocesos";
            Respuesta DatosRespuesta = new Respuesta();
            ParametrosND DatosEnvio = new ParametrosND();
            DatosEnvio.SP = Reproceso;
            DatosEnvio.CadenaBD = CadenasBD.GESTION_COMERCIO;

            try
            {
                DatosEnvio.NomParametros.Add("@FechaInicio");
                DatosEnvio.Datos.Add(fechaini);
                DatosEnvio.NomParametros.Add("@FechaFin");
                DatosEnvio.Datos.Add(fechafin);
                if (area != 0)
                {
                    DatosEnvio.NomParametros.Add("@FiltroArea");
                    DatosEnvio.Datos.Add(area);
                }
                else
                {
                    DatosEnvio.NomParametros.Add("@FiltroArea");
                    DatosEnvio.Datos.Add(DBNull.Value);
                }
                DatosRespuesta = Datos.Conexion.EjecutarSP(DatosEnvio);
            }
            catch (Exception ex)
            {
                DatosRespuesta.HayFallos = true;
                DatosRespuesta.Mensajes.Add("[" + (ex.StackTrace.Substring(ex.StackTrace.LastIndexOf('\n') + 7, (ex.StackTrace.LastIndexOf('(') - ex.StackTrace.LastIndexOf('\n') - 7))) + ": linea" + ex.StackTrace.Substring(ex.StackTrace.LastIndexOf(' ')) + "](Excep: " + ex.Message + ")");
            }

            return DatosRespuesta;
        }

        public static string ConsultarGeneral(string usuario)
        {
            Respuesta DatosRespuesta = new Respuesta();

            DatosRespuesta = ConsultarGeneral(
            "",
           usuario,
            "",
            "",
            "",
            "-1",
            "-1",
            "-1"
               );
            string nombre = null;

            if (DatosRespuesta.HayFallos)
            {
              
            }
            else
            {
                if (DatosRespuesta.Dt.Rows.Count > 0)
                {
                    nombre = DatosRespuesta.Dt.Rows[0]["Nombre"].ToString();
                    return nombre;
                }
                else
                {
                    nombre = DatosRespuesta.Dt.Rows[0]["Nombre"].ToString();
                    return nombre;
                }
                //tabladinamica(DatosRespuesta.Dt);
            }
            return nombre;
        }

        public static Respuesta ConsultarGeneral(
           string Nombre,
           string UsuarioWin,
           string IdPerfil,
           string HoraEntrada,
           string HoraSalida,
           string CodEstado,
           string CodGrupoForm,
           string CodGrupoOper
                       )
        {
            Respuesta DatosRespuesta = new Respuesta();
            ParametrosND DatosEnvio = new ParametrosND();
            DatosEnvio.SP = "seg.ConsultarUsuarioGen";
            DatosEnvio.CadenaBD = CadenasBD.CORE_PLATAFORMA;

            try
            {
                DatosEnvio.NomParametros.Add("@Nombre");
                DatosEnvio.Datos.Add(Nombre);

                DatosEnvio.NomParametros.Add("@UsuarioWin");
                DatosEnvio.Datos.Add(UsuarioWin);

                DatosEnvio.NomParametros.Add("@IdPerfil");
                DatosEnvio.Datos.Add(IdPerfil == "" ? -1 : int.Parse(IdPerfil));

                DatosEnvio.NomParametros.Add("@HoraEntrada");
                DatosEnvio.Datos.Add(HoraEntrada == "" ? -1 : int.Parse(HoraEntrada));

                DatosEnvio.NomParametros.Add("@HoraSalida");
                DatosEnvio.Datos.Add(HoraSalida == "" ? -1 : int.Parse(HoraSalida));

                DatosEnvio.NomParametros.Add("@CodEstado");
                DatosEnvio.Datos.Add(int.Parse(CodEstado));

                DatosEnvio.NomParametros.Add("@CodGrupoForm");
                DatosEnvio.Datos.Add(int.Parse(CodGrupoForm));

                DatosEnvio.NomParametros.Add("@CodGrupoOper");
                DatosEnvio.Datos.Add(int.Parse(CodGrupoOper));

                DatosRespuesta = Datos.Conexion.EjecutarSP(DatosEnvio);

            }
            catch (Exception ex)
            {
                DatosRespuesta.HayFallos = true;
                DatosRespuesta.Mensajes.Add("[" + (ex.StackTrace.Substring(ex.StackTrace.LastIndexOf('\n') + 7, (ex.StackTrace.LastIndexOf('(') - ex.StackTrace.LastIndexOf('\n') - 7))) + ": linea" + ex.StackTrace.Substring(ex.StackTrace.LastIndexOf(' ')) + "](Excep: " + ex.Message + ")");
            }
            return DatosRespuesta;
        }


        public static Respuesta ProductoEvento()
        {
            Respuesta DatosRespuesta = new Respuesta();
            ParametrosND DatosEnvio = new ParametrosND();

            DatosEnvio.SP = SPs.ComboProductoEvento;
            DatosEnvio.CadenaBD = CadenasBD.GESTION_COMERCIO;
            DatosRespuesta = Datos.Conexion.EjecutarSP(DatosEnvio);

            return DatosRespuesta;
        }

        public static string[] CargarInfoNroDoc(string nit)
        {
            Respuesta DatosRespuesta = new Respuesta();
            string segmento = null;
            string nombre = null;

            DatosRespuesta = ConsultarInfoNroDocAny(nit);
            int swExisteInfoNroDocAny = 0;
            if (!DatosRespuesta.HayFallos)
            {
                if (DatosRespuesta.Dt.Rows.Count > 0)
                {
                    try
                    {
                        nombre = DatosRespuesta.Dt.Rows[0]["NombreCliente"]?.ToString();
                        segmento = DatosRespuesta.Dt.Rows[0]["CodSubSegmento"]?.ToString();
                        return new string[] { nombre, segmento };

                        swExisteInfoNroDocAny = 1;
                    }
                    catch (Exception ex)
                    {
                      
                    }
                }
            }
            else
            {
              

            }

            //si no existe información en any se consulta en Nacional.
            if (swExisteInfoNroDocAny == 0)
            {
               DatosRespuesta = ConsultarInformacionNroDoc(nit.PadLeft(15,'0'));
                if (!DatosRespuesta.HayFallos)
                {
                    if (DatosRespuesta.Dt.Rows.Count > 0)
                    {
                        nombre = DatosRespuesta.Dt.Rows[0]["NOMBRE"]?.ToString();
                        segmento = DatosRespuesta.Dt.Rows[0]["COD_SEGMENTO"]?.ToString();
                        return new string[] { nombre, segmento };
                    }
                    else
                    {
                       

                    }
                }
                else
                {
                 


                }

            }

            return new string[] { nombre, segmento };

        }
        public static Respuesta ConsultarInformacionNroDoc(string NumDoc)
        {
            Respuesta DatosRespuesta = new Respuesta();
            ParametrosND DatosEnvio = new ParametrosND();
            DatosEnvio.SP = SPs.ConsultarInformacionNumDoc;
            DatosEnvio.CadenaBD = CadenasBD.GESTION_COMERCIO;
            DatosEnvio.Nacional = true;
            try
            {
                DatosEnvio.NomParametros.Add("@NumDoc");
                DatosEnvio.Datos.Add(NumDoc);

                DatosRespuesta = Datos.Conexion.EjecutarSP(DatosEnvio);

            }
            catch (Exception ex)
            {
                DatosRespuesta.HayFallos = true;
                DatosRespuesta.Mensajes.Add("[" + (ex.StackTrace.Substring(ex.StackTrace.LastIndexOf('\n') + 7, (ex.StackTrace.LastIndexOf('(') - ex.StackTrace.LastIndexOf('\n') - 7))) + ": linea" + ex.StackTrace.Substring(ex.StackTrace.LastIndexOf(' ')) + "](Excep: " + ex.Message + ")");
            }
            return DatosRespuesta;
        }

        public static Respuesta CausasInsert(
      int codArea,
      string tipo,
      string causa)
        {
            Respuesta DatosRespuesta = new Respuesta();
            ParametrosND DatosEnvio = new ParametrosND();
            DatosEnvio.SP = SPs.CausasInsert;
            DatosEnvio.CadenaBD = CadenasBD.GESTION_COMERCIO;

            try
            {
                DatosEnvio.NomParametros.Add("@CodArea");
                DatosEnvio.Datos.Add(codArea);
                DatosEnvio.NomParametros.Add("@TipoError");
                DatosEnvio.Datos.Add(tipo);
                DatosEnvio.NomParametros.Add("@Causas");
                DatosEnvio.Datos.Add(causa);

                DatosRespuesta = Datos.Conexion.EjecutarSP(DatosEnvio);

            }
            catch (Exception ex)
            {
                DatosRespuesta.HayFallos = true;
                DatosRespuesta.Mensajes.Add("[" + (ex.StackTrace.Substring(ex.StackTrace.LastIndexOf('\n') + 7, (ex.StackTrace.LastIndexOf('(') - ex.StackTrace.LastIndexOf('\n') - 7))) + ": linea" + ex.StackTrace.Substring(ex.StackTrace.LastIndexOf(' ')) + "](Excep: " + ex.Message + ")");
            }

            return DatosRespuesta;
        }






        public static Respuesta ValidacionUsuario(string FechaInicial, string FechaFinal)
        {
            Respuesta DatosRespuesta = new Respuesta();
            ParametrosND DatosEnvio = new ParametrosND();
            DatosEnvio.SP = SPs.ConsultarValidacionesRemesas;
            DatosEnvio.CadenaBD = CadenasBD.GESTION_COMERCIO;
            DatosEnvio.Nacional = true;

            try
            {
                DatosEnvio.NomParametros.Add("@FechaInicial");
                DatosEnvio.Datos.Add(FechaInicial);

                DatosEnvio.NomParametros.Add("@FechaFinal");
                DatosEnvio.Datos.Add(FechaFinal);

                DatosRespuesta = Datos.Conexion.EjecutarSP(DatosEnvio);
            }
            catch (Exception ex)
            {
                DatosRespuesta.HayFallos = true;
                DatosRespuesta.Mensajes.Add("[" + (ex.StackTrace.Substring(ex.StackTrace.LastIndexOf('\n') + 7, (ex.StackTrace.LastIndexOf('(') - ex.StackTrace.LastIndexOf('\n') - 7))) + ": linea" + ex.StackTrace.Substring(ex.StackTrace.LastIndexOf(' ')) + "](Excep: " + ex.Message + ")");
            }
            return DatosRespuesta;
        }

    }
}
