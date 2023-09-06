import java.io.*;
import java.nio.file.DirectoryStream;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import static java.nio.file.StandardCopyOption.REPLACE_EXISTING;
import java.nio.file.attribute.BasicFileAttributes;
import java.sql.*;
import java.text.DateFormat;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.*;
import java.util.logging.Formatter;
import java.util.logging.*;
import org.apache.commons.net.ftp.FTP;
import org.apache.commons.net.ftp.FTPClient;
import org.apache.commons.net.ftp.FTPFile;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.*;

public class CallCenterEnre {
    
     /************************************************
     * Adecuar: Ambiente, base
     ************************************************/
    
    //Ambiente de la base de datos. Valores: DEV, QA, PRO
    private static final String Ambiente = "PRO";
    
    

//    DIRECTORIOS RELATIVOS PARA PRODUCCION
    private static final String base= "/ias/CallCenter_ENRE/";    
    private static final String AProcesarDir = base  + "estructura/a_procesar";
    private static final String ProcesadosDir = base + "estructura/procesados";
    private static final String ErrorDir = base  + "estructura/error";
    private static final String LogsDir = base  + "estructura/logs";
    
    //Logger
    private static final Logger Log = Logger.getLogger(CallCenterEnre.class.getName());
    private static final String LogFileName = "batchLog.txt";
    //Patrones de archivo
    private static final String Pattern_1 = "Reporte_ENRE_????????_??_hs.{xls,xlsx}";
    private static final String FTPFilePattern = "^Reporte_ENRE_\\d{8}_\\d{2}_hs.xlsx?";
    //Driver y string de conexion a diferentes ambientes
    private static final String DriverClass = "oracle.jdbc.driver.OracleDriver";
    private static final String ConnDev = "jdbc:oracle:thin:NEXUS_ENRE/NEXUS_ENRE@tdbs6.tro.edenor:1521/GISDEV01";
    private static final String ConnQA = "jdbc:oracle:thin:NEXUS_ENRE/NEXUS_ENRE@TDBS5.PRO.EDENOR:1529/GISQA01";
    private static final String ConnPro = "jdbc:oracle:thin:NEXUS_ENRE/cami0net4@TCLH1.TRO.EDENOR:1528/GISPR01";
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
    //Cantidad de filas que conforman la cabecera
    private static final Integer CantFilasCabecera = 3;
    //Para parsear el nombre de archivo Reporte_ENRE_????????_??_hs
    private static final Integer FeDiaIni = 13;
    private static final Integer FeMesIni = 15;
    private static final Integer FeAnioIni = 17;
    private static final Integer FeAnioFin = 21;
    private static final Integer FeTandaIni = 22;
    private static final Integer FeTandaFin = 24;
    //para el metodo depurar. Indica los archivos a borrar segun fecha.
    private static final Integer DiasHaciaAtras = 7;
    //Tipos de datos de las celdas excel segun apache poi.
    private static final int CELL_TYPE_BLANK = 3;
    private static final int CELL_TYPE_BOOLEAN = 4;
    private static final int CELL_TYPE_ERROR = 5;
    private static final int CELL_TYPE_FORMULA = 2;
    private static final int CELL_TYPE_NUMERIC = 0;
    private static final int CELL_TYPE_STRING = 1;
    
    //
    private static Connection Connection = null;
    private static String FechaReporte = "";
    private static String nombreArchivo = "";
    private static final List<HashMap<String, String>> RowList = new ArrayList<>();
    //port, server, user y pass se pueblan desde la base
    private static String FTPServer = ""; 
    private static int FTPPort = 0;
    private static String FTPUser = "";
    private static String FTPPass = "";
    private static final String FTPWorkingDir = "./upload/";
    
    
    
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
        String feD = fileName.substring(FeDiaIni,FeMesIni);
        String feM = fileName.substring(FeMesIni,FeAnioIni);
        String feA = fileName.substring(FeAnioIni,FeAnioFin);
        return feD+feM+feA;
    }
    
    //Borra archivos viejos segun el valor de la constante DiasHaciaAtras
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
    
    //Parsea los nombres de archivo y obtiene el mas reciente
    public static FTPFile lastModifiedFTPFile(FTPFile[] files) throws Exception {

        FTPFile choice = null;
        try {
            choice = files[0];
            if (choice == null){
                throw new Exception (NoSeEncontraronArchivosMsj);
            }
        } catch (Exception e){
            throw e;
        }
        
        Integer feMasReciente = 0;
        for (FTPFile file : files) {
            if (file.getName().matches(FTPFilePattern)) {
                String feD = file.getName().substring(FeDiaIni,FeMesIni);
                String feM = file.getName().substring(FeMesIni,FeAnioIni);
                String feA = file.getName().substring(FeAnioIni,FeAnioFin);
                String feTanda = file.getName().substring(FeTandaIni,FeTandaFin);
                Integer fecha = Integer.parseInt(feA + feM + feD + feTanda);
                if (fecha > feMasReciente) {
                    choice = file;
                    feMasReciente = fecha;
                }
            }
        }
        
        return choice;
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
                       + "WHERE ORIGEN = 'FTP'";

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
        SimpleDateFormat segundos = new SimpleDateFormat("s,S");
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
            case "Segundos":
                fe = cell.getDateCellValue();
                cellAux = (fe != null) ? segundos.format(fe) : ""; 
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
        
        
        for (Cell campo : row) {                       //Siempre traen datos  
            if (campo.getColumnIndex() == 0            //Hora
                    || campo.getColumnIndex() == 5) {  //Nivel de servicio
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
        nombreArchivo = masReciente.getFileName().toString();
        Path baseP = Paths.get(AProcesarDir);
        Path archAProcesar = baseP.resolve(nombreArchivo);
        
        String archAProcStr = archAProcesar.toString();
        FileInputStream file = new FileInputStream(new File(archAProcStr));

        Workbook workbook = WorkbookFactory.create(file);
        Sheet sheet = workbook.getSheetAt(0);

        //Para que considere las filas sin datos.
        Iterator<Row> rowIterator = sheet.rowIterator();
        int numOfCols = rowIterator.next().getLastCellNum();
        for (Row row : sheet) {
            for (int cn = 0; cn < numOfCols; cn++) {
                row.getCell(cn, Row.CREATE_NULL_AS_BLANK);
            }
        }
        //Fin para que considere las filas sin datos.
        
        //Para discriminar cuando viene el promedio espera en segundos
        boolean traePromedioEspera = false;
        traePromedioEspera = (numOfCols == 8) ? false : true;
        if (numOfCols == 8) {
            traePromedioEspera = false;
        } else if (numOfCols == 9){
            traePromedioEspera = true;
        } else {
            throw new Exception ("Error en el formato del excel");
        }
        //Fin:Para discriminar cuando viene el promedio espera en segundos

        BasicFileAttributes attrs = Files.readAttributes(archAProcesar, BasicFileAttributes.class);
        String fechaModificado = attrs.lastModifiedTime().toString();
        String feMod = fechaModificado.substring(0, 19).replaceAll("T", " ");

        Integer contador = 0;
        rowIterator = sheet.rowIterator();
        for (Row row : sheet) {
            contador++;
            HashMap<String, String> excelRow = new HashMap<>();
            if (row.getRowNum() < CantFilasCabecera) {
                continue;
            }

            if (esFilaVacia(row)) {
                break;
            }

            Iterator<Cell> cellIterator = row.cellIterator();
            Cell cell = cellIterator.next();

            excelRow.put("HORA", excelTypeToStr(cell, "NumEntero"));

            cell = cellIterator.next();
            excelRow.put("RECIBIDAS", excelTypeToStr(cell, "NumEntero"));

            cell = cellIterator.next();
            excelRow.put("ABANDONADAS", excelTypeToStr(cell, "NumEntero"));

            cell = cellIterator.next();
            excelRow.put("ATENDIDAS", excelTypeToStr(cell, "NumEntero"));
            
            if (traePromedioEspera){
                cell = cellIterator.next();
                excelRow.put("PROMEDIO_ESPERA_SEGS", excelTypeToStr(cell, "Segundos"));
            }
            
            cell = cellIterator.next();
            excelRow.put("NIVEL_DE_SERVICIO", excelTypeToStr(cell, "Porcentaje"));

            cell = cellIterator.next();
            excelRow.put("AGENTES_LOGUEADOS", excelTypeToStr(cell, "NumEntero"));

            cell = cellIterator.next();
            excelRow.put("COMERCIAL", excelTypeToStr(cell, "NumEntero"));

            cell = cellIterator.next();
            excelRow.put("TECNICO", excelTypeToStr(cell, "NumEntero"));

            excelRow.put("ULT_MODIF_ARCHIVO", feMod);

            excelRow.put("FECHA_REPORTE", FechaReporte);
            
            RowList.add(excelRow);
        }

    }

    private static void insertar() throws SQLException {
        Integer rt = 0;
        Integer count = 0;
        FechaReporte = obtieneFechaReporte(nombreArchivo);
        
        for (HashMap<String, String> row : RowList) {
            String hora                 = (String)row.get("HORA");
            String recibidas            = (String)row.get("RECIBIDAS");
            String abandonadas          = (String)row.get("ABANDONADAS");
            String atendidas            = (String)row.get("ATENDIDAS");
            String promedio_espera_segs = (String)row.get("PROMEDIO_ESPERA_SEGS");
            String nivel_de_servicio    = (String)row.get("NIVEL_DE_SERVICIO");
            String agentes_logueados    = (String)row.get("AGENTES_LOGUEADOS");
            String comercial            = (String)row.get("COMERCIAL");
            String tecnico              = (String)row.get("TECNICO");
            String ult_modif_archivo    = (String)row.get("ULT_MODIF_ARCHIVO");
            
            try {
                String sql =
                        "INSERT INTO NEXUS_GIS.WS_CALL_CENTER_ENRE VALUES "
                        + "(?, ?, ?, ?, ?, ?, ?, ?, ?, SYSDATE, TO_DATE(?, 'YYYY-MM-DD HH24:MI:SS'), TO_DATE(?, 'DDMMYYYY'))";

                PreparedStatement ps = Connection.prepareStatement(sql);
                ps.setString(1, hora);
                ps.setString(2, recibidas);
                ps.setString(3, abandonadas);
                ps.setString(4, atendidas);
                ps.setString(5, promedio_espera_segs);
                ps.setString(6, nivel_de_servicio);
                ps.setString(7, agentes_logueados);
                ps.setString(8, comercial);
                ps.setString(9, tecnico);
                ps.setString(10, ult_modif_archivo);
                ps.setString(11, FechaReporte);

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

}
