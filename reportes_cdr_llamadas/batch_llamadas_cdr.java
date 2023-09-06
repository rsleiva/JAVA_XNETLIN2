import static java.nio.file.StandardCopyOption.REPLACE_EXISTING;

import java.io.BufferedOutputStream;
import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.nio.file.DirectoryStream;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.nio.file.attribute.BasicFileAttributes;
import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.PreparedStatement;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.text.DateFormat;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Calendar;
import java.util.Date;
import java.util.HashMap;
import java.util.List;
import java.util.Locale;
import java.util.logging.FileHandler;
import java.util.logging.Formatter;
import java.util.logging.Handler;
import java.util.logging.Level;
import java.util.logging.Logger;
import java.util.logging.SimpleFormatter;

import org.apache.commons.net.ftp.FTP;
import org.apache.commons.net.ftp.FTPClient;
import org.apache.commons.net.ftp.FTPFile;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;

import com.csvreader.CsvReader;

public class batch_llamadas_cdr {
        
    /************************************************
     * Adecuar: Ambiente, base
     ************************************************/
    
    //Ambiente de la base de datos. Valores: DEV, QA, PRO
    private static final String Ambiente = "PRO";
    
    //user y pass se pueblan desde la base
    private static String FTPServer = ""; 
    private static int FTPPort = 0;
    private static String FTPUser = "";
    private static String FTPPass = "";
    //
    private static final String FTPWorkingDir = "./Reportes/";
    
    

   //DIRECTORIOS PARA PRODUCCION
  private static final String base= "/ias/reportes_cdr_llamadas/";      
  private static final String AProcesarDir = base  + "estructura/a_procesar";
  private static final String ProcesadosDir = base + "estructura/procesados";
  private static final String ErrorDir = base  + "estructura/error";
  private static final String LogsDir = base  + "estructura/logs";
    
    //Logger
    private static final Logger Log = Logger.getLogger(batch_llamadas_cdr.class.getName());
    private static final String LogFileName = "batchLog.txt";
  
    //Driver y string de conexion a diferentes ambientes
    private static final String DriverClass = "oracle.jdbc.driver.OracleDriver";
    private static final String ConnDev = "jdbc:oracle:thin:NEXUS_ENRE/NEXUS_ENRE@tdbs6.tro.edenor:1521/GISDEV01";
    private static final String ConnQA = "jdbc:oracle:thin:NEXUS_ENRE/NEXUS_ENRE@TDBS5.PRO.EDENOR:1529/GISQA01";
    private static final String ConnPro = "jdbc:oracle:thin:NEXUS_ENRE/cami0net4@TCLH1.TRO.EDENOR:1528/GISPR01";
    
    //FTP Connection    
    private static final String FTPFilePattern = "^InformeCDR_\\d{4}-\\d{2}-\\d{2}.csv?";
    //Patron de archivo para regEx    
    private static final String Pattern_1 = "InformeCDR_????-??-??.{csv}";    
    //Conn
    private static Connection Connection = null;
    //fecha del reporte obtenida del nombre de archivo. Formato: ddmmaaaa
    private static String FechaReporte = "";
    //
    private static String nombreArchivo = "";
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
    private static final List<HashMap<String, String>> RowList = new ArrayList<>();
    
    private static final HashMap<String, String> excelRow = new HashMap<>();
    //Cantidad de filas que conforman la cabecera
    //private static final Integer CantFilasCabecera = 7;
    private static final Integer CantFilasCabecera = 11;
    //Para parsear el nombre de archivo Reporte_MinPlan_????????_??_hs
    //informeEstadistico_2014-11-19.xls
    private static final Integer FeDiaIni = 19;
    private static final Integer FeMesIni = 16;
    private static final Integer FeAnioIni = 11;
    private static final Integer FeAnioFin = 15;
    //private static final Integer FeTandaIni = 25;
    //private static final Integer FeTandaFin = 27;
    //para el metodo depurar. Indica los archivos a borrar segun fecha.
    private static final Integer DiasHaciaAtras = 7;
    //Tipos de datos de las celdas excel segun apache poi.
    private static final int CELL_TYPE_BLANK = 3;
    private static final int CELL_TYPE_BOOLEAN = 4;
    private static final int CELL_TYPE_ERROR = 5;
    private static final int CELL_TYPE_FORMULA = 2;
    private static final int CELL_TYPE_NUMERIC = 0;
    private static final int CELL_TYPE_STRING = 1;
    
    
    
////////////////////////////////////////////////////////////////////////////////
/////////////////************       M A I N      ********************///////////
////////////////////////////////////////////////////////////////////////////////
    public static void main(String[] args) throws IOException {

        try {
            
            inicializarLog();
            
            Log.log(Level.INFO, IniciandoMsj);
            
            depurarDirectorios();
            
            setearConexion();
            
            obtenerCredenciales();
            
            traerArchivo();
            
            parsearDocumento();
            
            insertar();
            
            moveFromTo(AProcesarDir, ProcesadosDir, nombreArchivo);
            
            Log.log(Level.INFO, FinalizandoMsj);
            
        } catch (SQLException se){
            Log.log(Level.SEVERE, "SQL Exception:");
            while (se != null) {
                Log.log(Level.SEVERE, "State  : {0}", se.getSQLState());
                Log.log(Level.SEVERE, "Message: {0}", se.getMessage());
                Log.log(Level.SEVERE, "Error Code  : {0}", se.getErrorCode());
                se = se.getNextException();
            }
            if (nombreArchivo.matches(FTPFilePattern)){
                moveFromTo(AProcesarDir, ErrorDir, nombreArchivo);
                Log.log(Level.SEVERE, MoverErrorMsj, nombreArchivo);
            }
            Log.info(FinalizadoConErrorMsj);
            
        } catch (IOException e) {
            Log.log(Level.SEVERE, e.getMessage());
            if (nombreArchivo.matches(FTPFilePattern)){
                moveFromTo(AProcesarDir, ErrorDir, nombreArchivo);
                Log.log(Level.SEVERE, MoverErrorMsj, nombreArchivo);
            }
            Log.info(FinalizadoConErrorMsj);
            
        } catch (NullPointerException e) {
            Log.log(Level.SEVERE, e.toString());
            if (nombreArchivo.matches(FTPFilePattern)){
                moveFromTo(AProcesarDir, ErrorDir, nombreArchivo);
                Log.log(Level.SEVERE, MoverErrorMsj, nombreArchivo);
            }
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
                        if (nombreArchivo.matches(FTPFilePattern)){
                            moveFromTo(AProcesarDir, ErrorDir, nombreArchivo);
                            Log.log(Level.SEVERE, MoverErrorMsj, nombreArchivo);
                        }
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
    
    
////////////////////////////////////////////////////////////////////////////////
/////////////////**********   M E T O D O S    *************////////////////////
////////////////////////////////////////////////////////////////////////////////
    //obtiene la fecha del reporte parseando nombre de archivo
    public static String obtieneFechaReporte(String fileName){
        String feD = fileName.substring(FeDiaIni,FeDiaIni+2);
        String feM = fileName.substring(FeMesIni,FeMesIni+2);
        String feA = fileName.substring(FeAnioIni,FeAnioFin);
        return feD+feM+feA;
    }
    
        //Borra archivos viejos segun el valor de la constante diasHaciaAtras
    public static Integer depurarDirectorios() throws ParseException, IOException{
        Integer count = 0;
        Path archivosAP = Paths.get(AProcesarDir);
        Path archivosP = Paths.get(ProcesadosDir);
        Path archivosE = Paths.get(ErrorDir);
        
        Date date = new Date();
        Calendar c = Calendar.getInstance();
        c.setTime(date);
        c.add(Calendar.DATE, - DiasHaciaAtras);
        Date unaSemanaAtras = c.getTime();
        
        try (DirectoryStream<Path> aProc = Files.newDirectoryStream(archivosAP, Pattern_1);
             DirectoryStream<Path> proc = Files.newDirectoryStream(archivosP, Pattern_1); 
             DirectoryStream<Path> error = Files.newDirectoryStream(archivosE, Pattern_1)) {
            
            
            for (Path elem : aProc) {
                    Files.delete(elem);
                    Log.log(Level.INFO, DepurarAProcMsj , elem.getFileName().toString());
                    count++;
            }
            
            for (Path elem : proc) {
                String str = obtieneFechaReporte(elem.getFileName().toString());
                DateFormat format = new SimpleDateFormat("ddMMyyyy", Locale.ENGLISH);
                Date fechaDeArchivo = format.parse(str);
                
                if (fechaDeArchivo.before(unaSemanaAtras)){
                    Files.delete(elem);
                    Log.log(Level.INFO, DepurarProcMsj, elem.getFileName().toString());
                    count++;
                }
            }

            for (Path elem : error) {
                String str = obtieneFechaReporte(elem.getFileName().toString());
                DateFormat format = new SimpleDateFormat("ddMMyyyy", Locale.ENGLISH);
                Date fechaDeArchivo = format.parse(str);
                
                if (fechaDeArchivo.before(unaSemanaAtras)){
                    Files.delete(elem);
                    Log.log(Level.INFO, DepurarErrorMsj, elem.getFileName().toString());
                    count++;
                }
            }

        } catch (IOException | ParseException e) {
            throw e;
        } finally {
            Log.log(Level.INFO, DepurarMsj, count);
        }
        
        return count;
    }
    
    public static FTPFile lastModifiedFTPFile(FTPFile[] files) throws Exception {

        FTPFile choice = null;
        try {
            choice = files[0];
            if (choice == null){
                throw new Exception (NoSeEncontraronArchivosMsj);
            }
        } catch (Exception e){
            //Log.log(Level.WARNING, e.getMessage());
            throw e;
        }
        
        Integer feMasReciente = 0;
        for (FTPFile file : files) {
            if (file.getName().matches(FTPFilePattern)) {
                String nombre = file.getName();
                
               // System.out.println(file.getName().substring(FeDiaIni,FeDiaIni+2));
                String feD = file.getName().substring(FeDiaIni,FeDiaIni+2);
               // System.out.println(nombre);
                String feM = file.getName().substring(FeMesIni,FeMesIni+2);
                String feA = file.getName().substring(FeAnioIni,FeAnioFin);
               // String feTanda = file.getName().substring(FeTandaIni,FeTandaFin);
                //Integer fecha = Integer.parseInt(feA + feM + feD + feTanda);
                Integer fecha = Integer.parseInt(feA + feM + feD );
                if (fecha > feMasReciente) {
                    choice = file;
                    feMasReciente = fecha;
                }
            }
        }
        
        return choice;
    }

    private static void traerArchivo() throws Exception {
        String server = FTPServer;
        int port = FTPPort;
        String user = FTPUser;
        String pass = FTPPass;
        
        FTPClient ftpClient = new FTPClient();
        
        try {

            ftpClient.connect(server, port);
            ftpClient.login(user, pass);
            ftpClient.enterLocalPassiveMode();
            ftpClient.setFileType(FTP.BINARY_FILE_TYPE);
            

            ftpClient.changeWorkingDirectory(FTPWorkingDir);

            FTPFile[] files = ftpClient.listFiles();
            FTPFile lastModif = lastModifiedFTPFile(files);
            
            switch (fueProcesadoAnteriormente(lastModif.getName())) {
                case "OK":
                    throw new Exception(NoSeVuelveAProcesarMsj);
                case "ERROR":
                    throw new Exception(NoSeVuelveAProcesarMsjErr);
            }

            String remoteFile1 = lastModif.getName();
            //remoteFile1 = remoteFile1.substring(0,FeDiaIni+2)+".xls";
            File downloadFile1 = new File(AProcesarDir + "/" + lastModif.getName());
            boolean success;
            try (OutputStream outputStream1 = new BufferedOutputStream(new FileOutputStream(downloadFile1))) {
                success = ftpClient.retrieveFile(remoteFile1, outputStream1);
            }

            if (success) {
                Log.log(Level.INFO, FueCopiadoOkMsj,lastModif.getName());
            } else {
                throw new Exception(FTPErrorAlBajarMsj);
            }
            

        } catch (IOException ex) {
            throw ex;
        } catch (Exception ex) {
            throw ex;
        } finally {
            try {
                if (ftpClient.isConnected()) {
                    ftpClient.logout();
                    ftpClient.disconnect();
                }
            } catch (IOException ex) {
                throw ex;
            }
        }
    }
    
    
    //Toma una celda de tipo numerico y la formatea
    private static String formatNumeric(Cell cell, String formato) throws Exception{
        
        if (cell.getCellType() != CELL_TYPE_NUMERIC){
            throw new Exception ("formatNumeric: el tipo de datos debe ser Numeric");
        }
        
        SimpleDateFormat mesAnio = new SimpleDateFormat("MM/yyyy");
        SimpleDateFormat diaMesAnio = new SimpleDateFormat("dd/MM/yyyy");
        String cellAux = "";
        Date fe = null;
        switch (formato) {
            case "NumEntero":
                Double numEnt = cell.getNumericCellValue();
                cellAux = Integer.toString(numEnt.intValue());
                break;
            case "NumDecimal":
                Double numDec = cell.getNumericCellValue();
                cellAux = Double.toString(numDec);
                break;
            case "MesAnio":
                fe = cell.getDateCellValue();
                cellAux = (fe != null) ? mesAnio.format(fe) : ""; 
                break;
            case "DiaMesAnio":
                fe = cell.getDateCellValue();
                cellAux = (fe != null) ? diaMesAnio.format(fe) : ""; 
                break;
            case "Porcentaje":
                Double numPor = cell.getNumericCellValue()*100;
                cellAux = Double.toString(numPor.intValue());
                cellAux += "%";
                break;
            default:
                throw new Exception("excelTypeToStr: " + formato + " tipo no soportado");
        }
        return cellAux;
    }
    
    
    //Toma el valor de una celda Excel y lo transforma en String.
    private static String excelTypeToStr(Cell cell, String formato) throws Exception {
        String cellAux = "";
        
        switch (cell.getCellType()) {
            case CELL_TYPE_BLANK: 
                break;
            case CELL_TYPE_BOOLEAN: 
                break;
            case CELL_TYPE_ERROR: 
                break;
            case CELL_TYPE_FORMULA: 
                throw new Exception ("No se permiten fórmulas en el documento");
            case CELL_TYPE_NUMERIC: 
                cellAux = formatNumeric(cell, formato);
                break;
            case CELL_TYPE_STRING: 
                cellAux = cell.getStringCellValue();
                break;
        }
        
        return cellAux;

    }

    private static boolean esFilaVacia(Row row) throws Exception {
        boolean esFilaVacia = true;
        for (Cell campo : row) {
            if (campo.getColumnIndex() == 0) {
                continue;
            }
            if (campo.getNumericCellValue() > 0) {
                esFilaVacia = false;
                break;
            }
        }
        return esFilaVacia;
    }

    private static void parsearDocumento()
            throws FileNotFoundException, IOException, InvalidFormatException,
            Exception {
        
        Path masReciente = obtenerUltimoModificado(AProcesarDir);
        //nombreArchivo = "InformeCDR_2015-05-21.csv";//masReciente.getFileName().toString();
        nombreArchivo = masReciente.getFileName().toString();
        Path baseP = Paths.get(AProcesarDir);
        Path archAProcesar = baseP.resolve(nombreArchivo);
        
        String archAProcStr = archAProcesar.toString();
        
        //FileInputStream file = new FileInputStream(new File(archAProcStr));

        //** inicio lecutra del archivo
     /*   BufferedReader bf = null;
        try {
            bf = new BufferedReader(new FileReader(archAProcStr));
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        }
            String line = null;
        try {
            while ((line = bf.readLine())!=null) {
                StringTokenizer tokens = new StringTokenizer(line, ";");
                while(tokens.hasMoreTokens()){
                    System.out.println(tokens.nextToken());
                }
            }
        } catch (IOException e) {
            e.printStackTrace();
        }
        //*fin lectura
        */
        try {
           
            CsvReader cvsReader = null;
            cvsReader = new CsvReader(archAProcStr.toString());  
            cvsReader.readHeaders();            
            String description=null;
           
            while (cvsReader.readRecord()){
                 
             // System.out.println(num+"|"+cvsReader.getRawRecord().length());
                
              String linea = cvsReader.getRawRecord();
              /*if (linea.indexOf(";;") != -1){
                  continue;
              }*/
                //System.out.println("Luego del Insert|"+linea);
                String campo[] = linea.split(";",-1);
                 //System.out.println("a");
                  //  System.out.println(num+"|"+linea.split(";").length );
                HashMap<String, String> excelRow = new HashMap<>();
                            
                excelRow.put("IDCAMPANIA", campo[0]);
                excelRow.put("IDLLAMADO", campo[1]);
                excelRow.put("TIPOCAMPANIA", campo[2]);
                excelRow.put("PEDIDO_DOCUMENTO", campo[3]);
                excelRow.put("CUENTA", campo[4]);
                excelRow.put("FECHA_LLAMADA", campo[5]);
                excelRow.put("HORA_LLAMADA", campo[6]);
                excelRow.put("TELEFONO", campo[7]);
                excelRow.put("ESTADO_LLAMADA", campo[8]);
                /*System.out.println("b");
                System.out.println("leng de campo[9]:"+campo[9].length());
                //if (campo[9]==null){
                    campo[9]="NULL";
                    System.out.println(linea);
                    System.out.println("c");
                //}
                    */
                excelRow.put("DURACION", campo[9]);
               
                if (linea.split(";").length==10) {
                    description=null;
                }else{
                    description=campo[10];
                }
                excelRow.put("DESCRIPCION_ESTADO", description); 
                excelRow.put("TIPODESTINO", campo[11]);
               // num=num+1;
                RowList.add(excelRow);
                
                //if (num == 100) {
               // System.out.println("d");
                 //  insertar(linea);
                  // System.out.println(num+"|"+excelRow);
                 // num=0;
                  // RowList.clear();
                //}
                //System.out.println(num);
                
            }
             cvsReader.close();
             //System.out.println(RowList);
            
            //System.out.println("fin");

       } catch (FileNotFoundException e) {
           
           System.out.println("FileNotFoundException|"+excelRow);
            e.printStackTrace();
       } catch (IOException e) {
           System.out.println("IOException|"+excelRow);
            e.printStackTrace();
        }
     

    }

    private static void insertar() throws SQLException {
        Integer rt = 0;
        Integer count = 0;
        //String description=null;
        FechaReporte = obtieneFechaReporte(nombreArchivo);
        
          /*String campo[] = linea.split(";");
                
               
                if (campo[9]==null){
                    campo[9]="NULL";
                }
               
                
                if (linea.split(";").length==10) {
                    description=null;
                }else{
                    description=campo[10];
                }*/
               
               // num=num+1;
        for (HashMap<String, String> row : RowList) {
      
            String IDCAMPANIA =        (String) row.get("IDCAMPANIA");//campo[0];//
            String IDLLAMADO =         (String) row.get("IDLLAMADO");//campo[1];//
            String TIPOCAMPANIA =      (String) row.get("TIPOCAMPANIA");//campo[2];//
            String PEDIDO_DOCUMENTO =  (String) row.get("PEDIDO_DOCUMENTO");//campo[3];//
            String CUENTA =            (String) row.get("CUENTA");//campo[4];
            String FECHA_LLAMADA =     (String) row.get("FECHA_LLAMADA");//campo[5];//
            String HORA_LLAMADA =      (String) row.get("HORA_LLAMADA");//campo[6];//
            String TELEFONO =          (String) row.get("TELEFONO");//campo[7];//
            String ESTADO_LLAMADA =    (String) row.get("ESTADO_LLAMADA");//campo[8];//
            String DURACION =          (String) row.get("DURACION");//campo[9];//
            String DESCRIPCION_ESTADO =(String) row.get("DESCRIPCION_ESTADO");//description;// 
            String TIPO_DESTINO =      (String) row.get("TIPODESTINO");//TIPO_DESTINO;// 
            
            try {
                String sql =
                        "INSERT INTO NEXUS_GIS.LLAM_CDR_DEV_SONDEOS VALUES "
                        + "(?,?,?,?,?,?,?,?,?,?,?, SYSDATE,?)";

                PreparedStatement ps = Connection.prepareStatement(sql);
                ps.setString(1, IDCAMPANIA);
                ps.setString(2, IDLLAMADO);
                ps.setString(3, TIPOCAMPANIA);
                ps.setString(4, PEDIDO_DOCUMENTO);
                ps.setString(5, CUENTA);
                ps.setString(6, FECHA_LLAMADA);
                ps.setString(7, HORA_LLAMADA);
                ps.setString(8, TELEFONO);
                ps.setString(9, ESTADO_LLAMADA);
                ps.setString(10, DURACION);
                ps.setString(11, DESCRIPCION_ESTADO);                
                ps.setString(12, TIPO_DESTINO);  

                rt = ps.executeUpdate();
                ps.close();
                
                count += rt;
            } catch (SQLException ex) {
                throw ex;
            }
        }
        Log.log(Level.INFO, SeInsertaronFilasMsj, count);
    }

    private static void closeConnection() throws SQLException {
        if (Connection != null) {
            Connection.close();
        }
    }

    private static void moveFromTo(String oriDir, String destDir, String name) throws IOException {
        Path oriBase = Paths.get(oriDir);
        Path destBase = Paths.get(destDir);
        Path fullPathOri = oriBase.resolve(name);
        Path fullPathDest = destBase.resolve(name);
        Files.move(fullPathOri, fullPathDest, REPLACE_EXISTING);
    }

    private static String fueProcesadoAnteriormente(String arch) throws IOException {
        String nomArch = arch;
        Path archivosP = Paths.get(ProcesadosDir);
        Path archivosE = Paths.get(ErrorDir);
        String ret = "";
        try (DirectoryStream<Path> proc = Files.newDirectoryStream(archivosP);
                DirectoryStream<Path> error = Files.newDirectoryStream(archivosE)) {

            for (Path elem : proc) {
                if (elem.getFileName().toString().equals(nomArch)) {
                    ret = "OK";
                }
            }

            for (Path elem : error) {
                if (elem.getFileName().toString().equals(nomArch)) {
                    ret = "ERROR";
                }
            }

        } catch (IOException e) {
            throw e;
        }
        return ret;
    }

    private static Path obtenerUltimoModificado(String dir) throws ParseException, IOException {
        Path ultimoArch = null;
        Date d2 = null;
        Date d1 = null;
        SimpleDateFormat df = new SimpleDateFormat("yyyy-MM-dd hh:mm:ss");

        Path basePath = Paths.get(dir);
        try (DirectoryStream<Path> stream_1 = Files.newDirectoryStream(basePath, Pattern_1)) {
           
            for (Path path : stream_1) {
                String feModif = Files.readAttributes(path, BasicFileAttributes.class).lastModifiedTime().toString();
                feModif = feModif.substring(0, 19);
                feModif = feModif.replaceAll("T", " ");

                try {
                    d1 = df.parse(feModif);
                } catch (ParseException ex) {
                    throw ex;
                }

                if (d2 == null) {
                    d2 = d1;
                    ultimoArch = path;
                }

                if (d1.after(d2)) {
                    d2 = d1;
                    ultimoArch = path;
                }
            }

        } catch (IOException ex) {
            throw ex;
        }
        return ultimoArch;

    }
    
    private static void inicializarLog() throws IOException{
            Handler fileHandler  = null;
            Formatter simpleFormatter = null;
            fileHandler  = new FileHandler(LogsDir + "/" + LogFileName, true);
            simpleFormatter = new SimpleFormatter();
            Log.addHandler(fileHandler);
            fileHandler.setLevel(Level.ALL);
            Log.setLevel(Level.ALL);
            fileHandler.setFormatter(simpleFormatter);
    }
    
    private static void setearConexion() throws SQLException, Exception{
        String conStr = "";
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
        } catch (Exception e) {
            throw e;
        }
    }
    
    public static void obtenerCredenciales() throws SQLException, Exception {
        
        ResultSet rs = null;
        PreparedStatement ps = null;
        
        try {
            String sql = "SELECT TIPO, VALOR "
                       + "FROM NEXUS_GIS.WS_DATA_BATCH "
                       + "WHERE ORIGEN = 'FTP1'";

            ps = Connection.prepareStatement(sql);
            rs = ps.executeQuery();
            
            while (rs.next()) {
                 
                switch (rs.getString("TIPO")) {
                    case "SERVER":
                        FTPServer = (String) rs.getString("VALOR");
                        break;
                    case "PORT":
                        FTPPort = (int) rs.getInt("VALOR");
                        break;
                    case "USUARIO":
                        FTPUser = (String) rs.getString("VALOR");
                        break;
                    case "PASS":
                        FTPPass = (String) rs.getString("VALOR");
                        break;
                    default:
                        throw new Exception(ObtenerCredencialesErrMsj);
                }
            }
            
            if (FTPServer.isEmpty() 
                || FTPPort == 0
                || FTPUser.isEmpty()
                || FTPPass.isEmpty()){
                throw new Exception(ObtenerCredencialesErrMsj);
            }
            
        } catch (SQLException ex) {
            throw ex;
        } catch (Exception e) {
            throw e;
        } finally {
            if (ps!=null) {
                ps.close();
            }
            if (rs!=null){
                rs.close();
            }
        }
    }
}
