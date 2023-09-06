/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */

import com.ws_enre.ElectrodependienteService;
import java.io.IOException;
import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.PreparedStatement;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.util.logging.FileHandler;
import java.util.logging.Formatter;
import java.util.logging.Handler;
import java.util.logging.Level;
import java.util.logging.Logger;
import java.util.logging.SimpleFormatter;

public class NewMain {
    
     
    /************************************************
     * Adecuar: Ambiente, base
     ************************************************/
    
    //Ambiente de la base de datos. Valores: DEV, QA, PRO
    private static final String Ambiente = "PRO";
    
    //Driver y string de conexion a diferentes ambientes
    private static final String DriverClass = "oracle.jdbc.driver.OracleDriver";
    private static final String ConnDev = "jdbc:oracle:thin:NEXUS_ENRE/NEXUS_ENRE@NEXDBDE01.PRO.EDENOR:1521:GISDEV01";
    private static final String ConnQA = "jdbc:oracle:thin:NEXUS_ENRE/NEXUS_ENRE@TDBS5.PRO.EDENOR:1529:GISQA01";
    private static final String ConnPro = "jdbc:oracle:thin:NEXUS_ENRE/cami0net4@TCLH1.TRO.EDENOR:1528/GISPR01";
    
     //Conn
    private static Connection Connection = null;
    ResultSet rs=null;
        //Logger
    private static final Logger Log = Logger.getLogger(NewMain.class.getName());
    private static final String LogFileName = "batchLog.txt";
    
    //    private static final String RepositorioDir = "C:\\Users\\mrrodriguez\\jhc\\ww\\Procesos Java Lectura Excel\\Excels_de_entrada\\estructuraCallCenter\\carpeta_en_la_red";
    
    //private static final String LogsDir = "C:\\Users\\mrrodriguez\\Documents\\Temas\\Evolutivo_llamadas_salientes\\55067_Reportes_llamadas\\reportes_llamadas\\estructura\\logs";
    

//    //DIRECTORIOS PARA PRODUCCION
    private static final String base= "/ias/Envio_Elect_ENRE/";    
//      private static final String base= "/home/omigliarini/";    
  //    private static final String AProcesarDir = base  + "estructura/a_procesar";
  //    private static final String ProcesadosDir = base + "estructura/procesados";
  //    private static final String ErrorDir = base  + "estructura/error";
      private static final String LogsDir = base  + "logs";
    
    //Literales
    private static final String NoSeVuelveAProcesarMsj = "El archivo ya fue procesado exitosamente. No se volvera a procesar";
    private static final String NoSeVuelveAProcesarMsjErr = "El archivo ya fue procesado con ERROR. No se volvera a procesar";
    private static final String NoSeEncontraronArchivosMsj = "No se encontraron archivos en el directorio";
    private static final String FinalizadoConErrorMsj = "---------  Proceso Finalizado Con ERROR     ----------";
    private static final String FinalizadoConWarningsMsj = "---------  Proceso Finalizado Con WARNINGS  ----------";
    private static final String DepurarAProcMsj = "Depuracion: se borro el archivo {0} del directorio a_procesar";
    private static final String DepurarProcMsj = "Depuracion: se borro el archivo {0} del directorio procesados";
    private static final String DepurarErrorMsj = "Depuracion: se borro el archivo {0} del directorio error";
    private static final String DepurarMsj = "Se borraron {0} archivos en la depuracion.";
    private static final String MoverErrorMsj = "Se mueve el archivo {0} a la carpeta 'error'";
    private static final String ObtenerCredencialesErrMsj = "Error al obtener usuario y contraseña";
    private static final String FueCopiadoOkMsj = "el archivo {0} fue copiado desde el repositorio remoto";
    private static final String SeInsertaronFilasMsj = "Se insertaron {0} filas.";
    private static final String FTPErrorAlBajarMsj = "Error al intentar bajar el archivo desde servidor FTP";
    private static final String IniciandoMsj = "---------  Comenzando Proceso Batch         ----------";
    private static final String FinalizandoMsj = "---------  Proceso Finalizado Correctamente ----------";
    //

    /**
     * @param args the command line arguments
     */
    public static void main(String[] args) {
        
        try {
            
            inicializarLog();
            
            Log.log(Level.INFO, IniciandoMsj);
            
            setearConexion();
            
            enviar();
            
            Log.log(Level.INFO, FinalizandoMsj);
            
        } catch (SQLException se){
            Log.log(Level.SEVERE, "SQL Exception:");
            while (se != null) {
                Log.log(Level.SEVERE, "State  : {0}", se.getSQLState());
                Log.log(Level.SEVERE, "Message: {0}", se.getMessage());
                Log.log(Level.SEVERE, "Error Code  : {0}", se.getErrorCode());
                se = se.getNextException();
            }
            
        } catch (IOException e) {
            Log.log(Level.SEVERE, e.getMessage());            
            Log.info(FinalizadoConErrorMsj);
            
        } catch (NullPointerException e) {
            Log.log(Level.SEVERE, e.toString());            
            Log.info(FinalizadoConErrorMsj);
            
        } catch (Exception e) {
            if (e.getMessage() != null){
                switch (e.getMessage()) {
                    case NoSeVuelveAProcesarMsj:
                        Log.log(Level.WARNING, e.getMessage());
                        Log.info(FinalizadoConWarningsMsj);
                        break;
                    case NoSeVuelveAProcesarMsjErr:
                        Log.log(Level.WARNING, e.getMessage());
                        Log.info(FinalizadoConWarningsMsj);
                        break;
                    default:                        
                        Log.log(Level.SEVERE, e.getMessage());
                        Log.info(FinalizadoConErrorMsj);
                } 
            }else {
                Log.log(Level.SEVERE, e.toString());
                Log.info(FinalizadoConErrorMsj);
            }
            
        } finally {
            try {
                closeConnection();
            } catch (SQLException ex) {
                Log.log(Level.SEVERE, ex.getMessage());
                Log.info(FinalizadoConErrorMsj);
            }
        }
    }
    
    private static void inicializarLog() throws IOException{
            Handler fileHandler;
            Formatter simpleFormatter;
            fileHandler  = new FileHandler(LogsDir + "/" + LogFileName, true);
            simpleFormatter = new SimpleFormatter();
            Log.addHandler(fileHandler);
            fileHandler.setLevel(Level.ALL);
            Log.setLevel(Level.ALL);
            fileHandler.setFormatter(simpleFormatter);
    }
     
    private static void setearConexion() throws SQLException, Exception{
        String conStr;
        Log.log(Level.INFO, "Conectando al ambiente {0}", Ambiente);
        switch (Ambiente) {
            case "DEV":
                conStr = ConnDev;
                break;
            case "QA":
                conStr = ConnQA;
                break;
            case "PRO":
                conStr = ConnPro;
                break;
            default:
                conStr = "";
        }

        try {
            Class.forName(DriverClass).newInstance();
            Connection = DriverManager.getConnection(conStr);
        } catch (SQLException se) {
            throw se;
        }
    }
    public static void enviar() throws SQLException, Exception {
        String texto;
        String response;
        String Rec, Doc;
        ResultSet rs = null;
        PreparedStatement ps = null;
           
       try {
            String sql = "SELECT    EMPRESA\r\n"
            		+ "       || '|'\r\n"
            		+ "       || ID_SISTEMA\r\n"
            		+ "       || '|'\r\n"
            		+ "       || ID_USUARIO\r\n"
            		+ "       || '|'\r\n"
            		+ "       || NRO_DOCUMENTO\r\n"
            		+ "       || '|'\r\n"
            		+ "       || RPAD (\r\n"
            		+ "             NVL (TO_CHAR (FECHA_DOCUMENTO, 'YYYY/MM/DD HH24:MI:SS'), ' '),\r\n"
            		+ "             19)\r\n"
            		+ "       || '|'\r\n"
            		+ "       || SISTEMA\r\n"
            		+ "       || '|'\r\n"
            		+ "       || NRO_RECLAMO\r\n"
            		+ "       || '|'\r\n"
            		+ "       || RPAD (NVL (TO_CHAR (FECHA_RECLAMO, 'YYYY/MM/DD HH24:MI:SS'), ' '),\r\n"
            		+ "                19)\r\n"
            		+ "       || '|'\r\n"
            		+ "       || MOTIVO\r\n"
            		+ "       || '|'\r\n"
            		+ "       || FAE\r\n"
            		+ "       || '|'\r\n"
            		+ "       || AUTONOMIA_FAE\r\n"
            		+ "       || '|'\r\n"
            		+ "       || MEDIDOR\r\n"
            		+ "       || '|'\r\n"
            		+ "       || NOMBRE\r\n"
            		+ "       || '|'\r\n"
            		+ "       || CALLE\r\n"
            		+ "       || '|'\r\n"
            		+ "       || NUMERO\r\n"
            		+ "       || '|'\r\n"
            		+ "       || PISO\r\n"
            		+ "       || '|'\r\n"
            		+ "       || DPTO\r\n"
            		+ "       || '|'\r\n"
            		+ "       || COD_POSTAL\r\n"
            		+ "       || '|'\r\n"
            		+ "       || LOCALIDAD\r\n"
            		+ "       || '|'\r\n"
            		+ "       || PARTIDO\r\n"
            		+ "       || '|'\r\n"
            		+ "       || TELEFONO\r\n"
            		+ "       || '|'\r\n"
            		+ "       || CELULAR\r\n"
            		+ "       || '|'\r\n"
            		+ "       || CENTRO\r\n"
            		+ "       || '|'\r\n"
            		+ "       || ALIMENT\r\n"
            		+ "       || '|'\r\n"
            		+ "       || ROUND (coord_x, 9)\r\n"
            		+ "       || '|'\r\n"
            		+ "       || ROUND (coord_Y, 9)\r\n"
            		+ "       || '|'\r\n"
            		+ "       || NVL (TO_CHAR (FECHA_CIERRE_DOC, 'YYYY/MM/DD HH24:MI:SS'), '')\r\n"
            		+ "          AS VALOR,\r\n"
            		+ "       NRO_RECLAMO,\r\n"
            		+ "       NRO_DOCUMENTO\r\n"
            		+ "  FROM NEXUS_GIS.WSENRE_ELECTRO_AFECTADOS ws, NEXUS_GIS.OMS_DOCUMENT d\r\n"
            		+ " WHERE     ENVIADO_ANEXOII IS NULL\r\n"
            		+ "       AND F_ENVIADO_ANEXOII IS NULL\r\n"
            		+ "       AND d.LAST_STATE_ID < 5\r\n"
            		+ "       AND d.NAME = ws.NRO_DOCUMENTO";
            //System.out.println("sql de busca calle: "+sql);
            ps = Connection.prepareStatement(sql);
            rs = ps.executeQuery();
            
        Log.log(Level.INFO, "Comenzando Envío de Documentos al ENRE");   
        try {
            ElectrodependienteService wservicio = new ElectrodependienteService();
             while(rs.next()) {
                
                texto = rs.getString("VALOR");
                Rec= rs.getString("NRO_RECLAMO");
                Doc= rs.getString("NRO_DOCUMENTO");
                response=wservicio.getDominio().enviarreclamo(texto);
    
                System.out.println("Retorna:  "+response);
                ActualizaloG(texto,Rec, Doc,response); 
                Actualizar(Rec, Doc);
                 
             }
             
           } catch (SQLException e) { 
               Log.log(Level.SEVERE, e.toString());
               Log.info("Proceso Finalizado Con ERROR al enviar Documentos al ENRE");
            }  
        } catch (SQLException e) {
            Log.log(Level.SEVERE, e.toString());
            Log.info("Proceso Finalizado Con ERROR al enviar Documentos al ENRE");
        }finally{
           Connection.close();
       }
    
    } 
    public static void Actualizar(String reclamo, String documento) throws SQLException, Exception {
        String texto;
        String response;
        ResultSet rs = null;
        PreparedStatement ps = null;
           
       try {
              String sql = "UPDATE NEXUS_GIS.WSENRE_ELECTRO_AFECTADOS"
                    + " SET ENVIADO_ANEXOII=?,F_ENVIADO_ANEXOII=sysdate "
                    + " WHERE NRO_RECLAMO = ? and NRO_DOCUMENTO = ?";

            ps = Connection.prepareStatement(sql);
            ps.setString(1, "SI");
            ps.setString(2, reclamo);
            ps.setString(3, documento);                    
            ps.executeUpdate();
            ps.close();
            Connection.commit();           
         
        
        } catch (SQLException e) {
            Log.log(Level.SEVERE, e.toString());
            Log.info("Proceso Finalizado Con ERROR al Actualizar Documentos al ENRE");
        }
    
    }
     public static void ActualizaloG(String req,String reclamo, String documento,String Respuesta) throws SQLException, Exception {
        String texto;
        String response;
        ResultSet rs = null;
        PreparedStatement ps = null;
           
       try {
              String sql = "INSERT INTO NEXUS_GIS.WSENRE_ELECTRO_LOG(INFO_ENVIADA,NRO_RECLAMO,NRO_DOCUMENTO,RESPUESTA_ENRE,FECHA_RTA) "
                    + " VALUES (?,?,?,?,SYSDATE)";

            ps = Connection.prepareStatement(sql);
            ps.setString(1, req);
            ps.setString(2, reclamo);
            ps.setString(3, documento);
            ps.setString(4, Respuesta);
            ps.executeUpdate();
            ps.close();
            Connection.commit();           
         
        
        } catch (SQLException e) {
            Log.log(Level.SEVERE, e.toString());
            Log.info("Proceso Finalizado Con ERROR al Insertar Documentos en el LOG tabla ENRE");
        }
    
    }
    
       private static void closeConnection() throws SQLException {
        if (Connection != null) {
            Connection.close();
        }
    }
    
}   
   
    
     
