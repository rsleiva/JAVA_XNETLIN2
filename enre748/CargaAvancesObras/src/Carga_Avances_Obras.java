import java.io.*;
import java.net.MalformedURLException;
import java.net.UnknownHostException;
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
import java.util.*;
import java.util.Date;
import java.util.logging.Formatter;
import java.util.logging.*;
import javax.activation.DataHandler;
import javax.activation.DataSource;
import javax.activation.FileDataSource;
import javax.mail.Address;
import javax.mail.BodyPart;
import javax.mail.Message;
import javax.mail.MessagingException;
import javax.mail.SendFailedException;
import javax.mail.Session;
import javax.mail.Transport;
import javax.mail.internet.InternetAddress;
import javax.mail.internet.MimeBodyPart;
import javax.mail.internet.MimeMessage;
import javax.mail.internet.MimeMultipart;
import jcifs.smb.NtlmPasswordAuthentication;
import jcifs.smb.SmbException;
import jcifs.smb.SmbFile;
import jcifs.smb.SmbFileInputStream;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

/**
 *
 * @author mrrodriguez
 */
public class Carga_Avances_Obras {
    
  /************************************************
     * Adecuar: Ambiente, NetworkDomain,  
     * NetworkFolder, base 
     ************************************************/
    
    //Ambiente de la base de datos. Valores: DEV, QA, PRO
    private static final String Ambiente = "PRO";
    //Conexion a la carpeta de red
    private static final String NetworkDomain = "PRO";   
    private static final String NetworkFolder = "smb://SRVCENFSPS001.PRO.EDENOR/DatosCEN/Iox/GCASIS/DesSistemasTecnicos/WS_ENRE/";
    
    //user y pass se pueblan desde la base
    private static String NetworkUser = ""; 
    private static String NetworkPass = "";
    
  /* 
    private static final String AProcesarDir = "C:\\parte_obras\\estructura\\a_procesar";
    private static final String ProcesadosDir = "C:\\parte_obras\\estructura\\procesados";
    private static final String ErrorDir = "C:\\parte_obras\\estructura\\error";
    private static final String LogsDir = "C:\\parte_obras\\estructura\\logs";
   
  */

    //Se utiliza para simular carpeta de la red
    //private static final String LocalDebugFolder = "C:\\Users\\mrrodriguez\\Documents\\Temas\\Evolutivo_Web_Services\\58022_WSENRE_EQuipos_de_trabajo\\Excels_de_entrada\\esGrupTrabEnre\\repositorioLocal";
    private static final String LocalDebugFolder = "";
    
    //DIRECTORIOS RELATIVOS PARA PRODUCCION    
 

    private static final String base= "/ias/enre748/CargaAvancesObras/";    
    private static final String AProcesarDir = base  + "estructura/a_procesar";
    private static final String ProcesadosDir = base + "estructura/procesados";
    private static final String ErrorDir = base  + "estructura/error";
    private static final String LogsDir = base  + "estructura/logs";

    
    //Logger
    private static final Logger Log = Logger.getLogger(Carga_Avances_Obras.class.getName());
    private static final String LogFileName = "batchLog.txt";
    //Patron de archivo
    private static final String Pattern_1 = "03-Carga de Avances_??-??-????.{xls,xlsx}";
    private static final String RegExPattern = "^03-Carga de Avances_\\d{2}\\-\\d{2}\\-\\d{4}.xlsx?";
    //Driver y string de conexion a diferentes ambientes
    private static final String DriverClass = "oracle.jdbc.driver.OracleDriver";
    private static final String ConnDev = "jdbc:oracle:thin:NEXUS_ENRE/NEXUS_ENRE@tdbs6.tro.edenor:1521:GISDEV01";
    private static final String ConnQA = "jdbc:oracle:thin:NEXUS_ENRE/NEXUS_ENRE@TDBS5.PRO.EDENOR:1529:GISQA01";
    private static final String ConnPro = "jdbc:oracle:thin:NEXUS_ENRE/cami0net4@ltronxgisbdpr01.pro.edenor:1528:GISPR01";
    
    //Nombre del libro a leer
    private static final String nombreDeLibro = "Carga_de_Avances";
    
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
    private static final String ObtenerCredencialesErrMsj = "Error al obtener usuario y contraseÃƒÂ±a";
    private static final String FueCopiadoOkMsj = "el archivo {0} fue copiado desde el repositorio remoto";
    private static final String SeInsertaronFilasMsj = "Se insertaron {0} filas.";
    private static final String IniciandoMsj = "---------  Comenzando Proceso Batch         ----------";
    private static final String FinalizandoMsj = "---------  Proceso Finalizado Correctamente ----------";
    //Indices parsear el nombre de archivo 03-Carga de Avances_??-??-????
    private static final Integer FeDiaIni = 20;
    private static final Integer FeDiaFin = 22;
    private static final Integer FeMesIni = 23;
    private static final Integer FeMesFin = 25;
    private static final Integer FeAnioIni = 26;
    private static final Integer FeAnioFin = 30;
    //Definido para el metodo depurar. Indica los archivos a borrar segun fecha.
    private static final Integer DiasHaciaAtras = 60;
    
    //Tipos de datos de las celdas excel segun apache poi.
    private static final int CELL_TYPE_BLANK = 3;
    private static final int CELL_TYPE_BOOLEAN = 4;
    private static final int CELL_TYPE_ERROR = 5;
    private static final int CELL_TYPE_FORMULA = 2;
    private static final int CELL_TYPE_NUMERIC = 0;
    private static final int CELL_TYPE_STRING = 1;
    
    private static final Integer CantFilasCabecera = 1;
    private static final List<HashMap<String, String>> RowList = new ArrayList<>();
    private static String nombreArchivo = "";
    private static Connection connection = null;
    private static String msgCancela= "";
    private static Vector dir_no_encontradas= new Vector();
   private static String FileName = null;
    private static String FlagProcess = null;
    
////////////////////////////////////////////////////////////////////////////////
/////////////////************       M A I N      ********************///////////
////////////////////////////////////////////////////////////////////////////////
    @SuppressWarnings("ThrowableResultIgnored")
    public static void main(String[] args) throws IOException, Exception {

        try {
            inicializarLog();            
            Log.log(Level.INFO, IniciandoMsj);            
            depurarDirectorios();            
            Log.log(Level.INFO, "1 Finaliza OK funcion depurarDirectorios");            
            setearConexion();            
            Log.log(Level.INFO, "2 Finaliza OK funcion setearConexion");            
            obtenerCredenciales();            
            Log.log(Level.INFO, "3 Finaliza OK funcion obtenerCredenciales");            
            traerArchivo();
          //  traerArchivoDebug();
            if ( FlagProcess != null){
            parsearDocumento();
            }
            if ( FlagProcess != null){
            insertar();
             } 
            
            if ( FlagProcess != null){
            moveFromTo(AProcesarDir, ProcesadosDir, nombreArchivo);
             }            

            Log.log(Level.INFO, FinalizandoMsj);            
        } catch (SQLException se){
            Log.log(Level.SEVERE, "SQL Exception:");
            while (se != null) {
                Log.log(Level.SEVERE, "State  : {0}", se.getSQLState());
                Log.log(Level.SEVERE, "Message: {0}", se.getMessage());
                Log.log(Level.SEVERE, "Error Code  : {0}", se.getErrorCode());
                se = se.getNextException();
            }
            if (nombreArchivo.matches(RegExPattern)){
                moveFromTo(AProcesarDir, ErrorDir, nombreArchivo);
                Log.log(Level.INFO, MoverErrorMsj, nombreArchivo);
            }
            Log.info(FinalizadoConErrorMsj);
        } catch (IOException e) {
            Log.log(Level.SEVERE, e.getMessage());
            if (nombreArchivo.matches(RegExPattern)){
                moveFromTo(AProcesarDir, ErrorDir, nombreArchivo);
                Log.log(Level.INFO, MoverErrorMsj, nombreArchivo);
            }
            Log.info(FinalizadoConErrorMsj);
        } catch (NullPointerException e) {
            Log.log(Level.SEVERE, e.toString());
            if (nombreArchivo.matches(RegExPattern)){
                moveFromTo(AProcesarDir, ErrorDir, nombreArchivo);
                Log.log(Level.INFO, MoverErrorMsj, nombreArchivo);
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
                        if (nombreArchivo.matches(RegExPattern)){
                            moveFromTo(AProcesarDir, ErrorDir, nombreArchivo);
                            Log.log(Level.INFO, MoverErrorMsj, nombreArchivo);
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
    
    public static void obtenerCredenciales() throws SQLException, Exception {
        
        ResultSet rs = null;
        PreparedStatement ps = null;
        
        try {
            String sql = "SELECT TIPO, VALOR "
                       + "FROM NEXUS_GIS.WS_DATA_BATCH "
                       + "WHERE ORIGEN = 'IOX'";

            ps = connection.prepareStatement(sql);
            rs = ps.executeQuery();

            while (rs.next()) {
                switch (rs.getString("TIPO")) {
                    case "USUARIO":
                        NetworkUser = (String) rs.getString("VALOR");
                        break;
                    case "PASS":
                        NetworkPass = (String) rs.getString("VALOR");
                        break;
                    default:
                        throw new Exception(ObtenerCredencialesErrMsj);
                }
            }
            
            if (NetworkUser.isEmpty() || NetworkPass.isEmpty()){
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
    
    
    public static void traerArchivo() throws MalformedURLException, SmbException, UnknownHostException, IOException, Exception{
        String domain = NetworkDomain;
        String user = NetworkUser;
        String pass = NetworkPass;
        String path = NetworkFolder;
        NtlmPasswordAuthentication auth = new NtlmPasswordAuthentication(domain,user, pass);
        SmbFile repositorio = new SmbFile(path,auth);
        SmbFile[] archivos = repositorio.listFiles();
        
        SmbFile masReciente = traerUltimo(archivos);
        if ( masReciente == null){
             
           throw new Exception("No se encontraron archivos para procesar del tipo: Avances" ); 
                     
             
         }
         
         FlagProcess = "procesar";
        
         FileName=masReciente.getName();
         
 /*        
         String a = fueProcesadoAnteriormente(masReciente.getName());
       

         
         
        if ( a == "OK"){
//        throw new Exception(NoSeVuelveAProcesarMsj);  
          Log.log(Level.INFO, "El Archivo " + FileName + " Fue procesado anteriormente con exito, no se lo volvera a procesar.");
        //  FlagNoProc(Patron);
            FlagProcess = null;
        } 
        
       if ( a == "ERROR"){
               Log.log(Level.INFO, "El Archivo " + FileName + " Fue procesado anteriormente con error, se intentara reprocesarlo.");
             } 
       
 */
       
       
   if ( FlagProcess != null){    
        
        FileOutputStream os;
        try (SmbFileInputStream in = new SmbFileInputStream(masReciente)) {
            Path dir = Paths.get(AProcesarDir);
            Path fullPath = dir.resolve(masReciente.getName());
            os = new FileOutputStream(fullPath.toString());
            byte[] b = new byte[8192];
            int n;
            while ((n = in.read(b)) > 0) {
                os.write(b, 0, n);
            }
        }
        os.close();
        Log.log(Level.INFO, FueCopiadoOkMsj,masReciente.getName());
    }
    }
    
    //Se usa para prueba. Copia el archivo desde una carpeta local
    public static void traerArchivoDebug() throws MalformedURLException, SmbException, UnknownHostException, IOException, Exception{
        String path = LocalDebugFolder;
        File repositorio = new File(path);
        File[] archivos = repositorio.listFiles();
        
        File masReciente = traerUltimoDebug(archivos);
        
        switch (fueProcesadoAnteriormente(masReciente.getName())) {
                case "OK":
                    throw new Exception(NoSeVuelveAProcesarMsj);
                case "ERROR":
                    throw new Exception(NoSeVuelveAProcesarMsjErr);
        }
        
        FileOutputStream os;
        try (FileInputStream in = new FileInputStream(masReciente)) {
            Path dir = Paths.get(AProcesarDir);
            Path fullPath = dir.resolve(masReciente.getName());
            os = new FileOutputStream(fullPath.toString());
            byte[] b = new byte[8192];
            int n;
            while ((n = in.read(b)) > 0) {
                os.write(b, 0, n);
            }
        }
        os.close();
        Log.log(Level.INFO, FueCopiadoOkMsj,masReciente.getName());
    }
    
    //se usa para prueba
    public static File traerUltimoDebug(File[] files) throws Exception {
        File choice = null;
        try {
            choice = files[0];
            if (choice == null){
                throw new Exception (NoSeEncontraronArchivosMsj);
            }
        } catch (Exception e){
            throw e;
        }
        
        Integer feMasReciente = 0;
        for (File file : files) {
            if (file.getName().matches(RegExPattern)) {
                String feD = file.getName().substring(FeDiaIni,FeDiaFin);
                String feM = file.getName().substring(FeMesIni,FeMesFin);
                String feA = file.getName().substring(FeAnioIni,FeAnioFin);
                Integer fecha = Integer.parseInt(feA + feM + feD);
                if (fecha > feMasReciente) {
                    choice = file;
                    feMasReciente = fecha;
                }
            }
        }
        
        return choice;
    }
    
    public static SmbFile traerUltimo(SmbFile[] files) throws Exception {
        SmbFile choice = null;
        try {
            choice = files[0];
            if (choice == null){
                throw new Exception (NoSeEncontraronArchivosMsj);
            }
        } catch (Exception e){
            throw e;
        }
        choice = null;
        Integer feMasReciente = 0;
        for (SmbFile file : files) {
            if (file.getName().matches(RegExPattern)) {
                String feD = file.getName().substring(FeDiaIni,FeDiaFin);
                String feM = file.getName().substring(FeMesIni,FeMesFin);
                String feA = file.getName().substring(FeAnioIni,FeAnioFin);
                Integer fecha = Integer.parseInt(feA + feM + feD);
                if (fecha > feMasReciente) {
                    choice = file;
                    feMasReciente = fecha;
                }
            }
        }
        
        return choice;
    }
    
    
    //obtiene la fecha del reporte parseando nombre de archivo
    public static String obtieneFechaReporte(String fileName){
        String feD = fileName.substring(FeDiaIni,FeDiaFin);
        String feM = fileName.substring(FeMesIni,FeMesFin);
        String feA = fileName.substring(FeAnioIni,FeAnioFin);       
        return feA+feM+feD;
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
                DateFormat format = new SimpleDateFormat("yyMMdd", Locale.ENGLISH);
                Date fechaDeArchivo = format.parse(str);
                
                if (fechaDeArchivo.before(unaSemanaAtras)){
                    Files.delete(elem);
                    Log.log(Level.INFO, DepurarProcMsj, elem.getFileName().toString());
                    count++;
                }
            }

            for (Path elem : error) {
                String str = obtieneFechaReporte(elem.getFileName().toString());
                DateFormat format = new SimpleDateFormat("yyMMdd", Locale.ENGLISH);
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
            connection = DriverManager.getConnection(conStr);
        } catch (SQLException | ClassNotFoundException | InstantiationException | IllegalAccessException se) {
            throw se;
        }

    }

    //Toma una celda de tipo numerico y la formatea
    private static String formatNumeric(Cell cell, String formato) throws Exception{
        
        if (cell.getCellType() != CELL_TYPE_NUMERIC
                && cell.getCellType() != CELL_TYPE_FORMULA) {
            throw new Exception("formatNumeric: el tipo de datos debe ser Numeric");
        } else if (cell.getCellType() == CELL_TYPE_FORMULA
                && cell.getCachedFormulaResultType() != CELL_TYPE_NUMERIC) {
            throw new Exception("formatNumeric: el tipo de datos del rdo de la fÃƒÂ³rmula debe ser Numeric");
        }
        
        SimpleDateFormat mesAnio = new SimpleDateFormat("MM/yyyy");
        SimpleDateFormat diaMesAnio = new SimpleDateFormat("dd/MM/yyyy");
        String cellAux = "";
        Date fe;
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
    private static String excelTypeToStr(Cell cell) throws Exception {
        String cellAux = "";
        
        switch (cell.getCellType()) {
            case CELL_TYPE_BLANK:
                break;
            case CELL_TYPE_BOOLEAN:
                break;
            case CELL_TYPE_ERROR:
                break;
            case CELL_TYPE_FORMULA:
                switch (cell.getCachedFormulaResultType()) {
                    case Cell.CELL_TYPE_STRING:
                        return cell.getRichStringCellValue().getString();
                    case Cell.CELL_TYPE_NUMERIC:
                        return Double.toString(Math.round(cell.getNumericCellValue()));
                    default:
                        break;
                }
            case CELL_TYPE_NUMERIC:
                // Double numEnt = cell.getNumericCellValue();                
                //cellAux = cell.getDateCellValue().toString();                
                Double numEnt = cell.getNumericCellValue();
                cellAux = Double.toString(numEnt);                
                // cellAux = formatNumeric(cell, formato);                
                break;
            case CELL_TYPE_STRING:
                cellAux = cell.getStringCellValue();
                break;
         
        }
        
        return cellAux;
    }
    
    private static boolean esFilaAOmitir(Integer index, String ext) throws Exception{
        
        if(!ext.toUpperCase().equals("XLSX") && !ext.toUpperCase().equals("XLS")){
            throw new Exception("[ERROR:esFilaAOmitir]: ExtensiÃƒÂ³n de archivo no es la permitida (XLS, XLSX)");
        }
        
        Integer[] filasAOmitir = {0, 1, 2, 3, 4};
        
        boolean existe = false;
        for (Integer elem: filasAOmitir){
            if (Objects.equals(index, elem)) {
                existe = true;
                break;
            } else {
                existe = false;
            }
        }
        return existe;
    }
    
    
    private static boolean esFilaTurno(Integer index, String ext) throws java.lang.Exception{
        
        if(!ext.toUpperCase().equals("XLSX") && !ext.toUpperCase().equals("XLS")){
            throw new Exception("[ERROR:esFilaAOmitir]: ExtensiÃƒÂ³n de archivo no es la permitida (XLS, XLSX)");
        }
        
        Integer[] filasTurno = {4, 16, 28};
        
        boolean existe = false;
        for (Integer elem: filasTurno){
            if (Objects.equals(index, elem)) {
                existe = true;
                break;
            } else {
                existe = false;
            }
        }
        return existe;
    }
    
    private static void parsearDocumento()
            throws FileNotFoundException, IOException, InvalidFormatException,
            Exception {
        
        Path masReciente = obtenerUltimoModificado(AProcesarDir);
        
        //debug
        nombreArchivo = masReciente.getFileName().toString();
        
        Path baseP = Paths.get(AProcesarDir);
        Path archAProcesar = baseP.resolve(nombreArchivo);
        
        String archAProcStr = archAProcesar.toString();
        FileInputStream file = new FileInputStream(new File(archAProcStr));

        Workbook workbook = WorkbookFactory.create(file);
        Sheet sheet = workbook.getSheet((workbook.getSheetName(0)));//.getSheet(nombreDeLibro);
        
        BasicFileAttributes attrs = Files.readAttributes(archAProcesar, BasicFileAttributes.class);
        String fechaModificado = attrs.lastModifiedTime().toString();
        String feMod = fechaModificado.substring(0, 19).replaceAll("T", " ");
        
        
        String extension = "";
        int i = archAProcesar.getFileName().toString().lastIndexOf('.');
        if (i >= 0) {
            extension = archAProcesar.getFileName().toString().substring(i+1);
        }
        
        try {
        //Para que considere las filas sin datos.             
        // Iterator<Row> rowIterator = sheet.rowIterator();      

        Integer contador = 0;
        //rowIterator = sheet.rowIterator();
        
        SimpleDateFormat d = new SimpleDateFormat("dd-M-yyyy");
       
        for (Row row : sheet) {
            
            contador++;
            HashMap<String, String> excelRow = new HashMap<>();
            if (row.getRowNum() < CantFilasCabecera) {
                continue;
            }

            if (esFilaVacia(row)) {
                break;
            }
            
            String valor1 = excelTypeToStr(row.getCell(0));
            if (valor1.equals("")|| (valor1.equals(null))|| (valor1.toLowerCase().equals("proyecto"))){
                 throw new Exception ("Error de formato en el archivo " + FileName + " la primera fila o columna estan en blanco.");
                 
             }

            Iterator<Cell> cellIterator = row.cellIterator();
            

            
            
            Cell cell = cellIterator.next();
            if (!cell.toString().equals("-")){               
                excelRow.put("Proyecto", cell.toString());
            }else{
                excelRow.put("Proyecto",excelTypeToStr(cell));
            }
            cell = cellIterator.next();
            if (!cell.toString().equals("-")){               
                excelRow.put("Obra", cell.toString());
            }else{
                excelRow.put("Obra",excelTypeToStr(cell));
            }            
            cell = cellIterator.next();
            if (!cell.toString().equals("-")){               
                excelRow.put("Fechavance", cell.toString());
            }else{
                excelRow.put("Fechavance",excelTypeToStr(cell));
            }            
            cell = cellIterator.next();
           
            if (!cell.toString().equals("-")){            
                excelRow.put("FechaServReal", cell.toString());
            }else{
                 excelRow.put("FechaServReal", excelTypeToStr(cell));
            }
            cell = cellIterator.next();            
            excelRow.put("PorcentAvance", excelTypeToStr(cell));
            cell = cellIterator.next();
            excelRow.put("Cables_BT", excelTypeToStr(cell));
            cell = cellIterator.next();
            excelRow.put("Lineas_BT", excelTypeToStr(cell));
            cell = cellIterator.next();
            excelRow.put("Cables_MT", excelTypeToStr(cell));
            cell = cellIterator.next();
            excelRow.put("Lineas_MT", excelTypeToStr(cell));
            cell = cellIterator.next();
            excelRow.put("PotenTrafoMTBT", excelTypeToStr(cell));
            cell = cellIterator.next();
            excelRow.put("TransAereosMTBT", excelTypeToStr(cell));
            cell = cellIterator.next();
            excelRow.put("TransNOAereosMTBT", excelTypeToStr(cell));
            cell = cellIterator.next();
            excelRow.put("TelemandoRedMT", excelTypeToStr(cell));
            cell = cellIterator.next();
            excelRow.put("ExpanCT", excelTypeToStr(cell));
            cell = cellIterator.next();
            excelRow.put("RedAT", excelTypeToStr(cell));
            cell = cellIterator.next();
            excelRow.put("PotenciaAT", excelTypeToStr(cell));
            cell = cellIterator.next();
            excelRow.put("EstadoObra", excelTypeToStr(cell));
            cell = cellIterator.next();
            excelRow.put("Observacion", excelTypeToStr(cell)); 
            
            RowList.add(excelRow);
        }
         
        } catch (NullPointerException e){
            throw new NullPointerException ("read: Error al parsear el excel - " + e);
        } catch (Exception e){
            throw new Exception ("Error al procesar el archivo excel de avances - " + e);
        }

    }
    
     public static ResultSet fechainsert()throws Exception {        
        ResultSet resultado = null;
       
        try{
            Connection conexion = connection;
            String q = "select sysdate from dual";

            String query = q;
            //Creo una sentencia a partir de la conexiÃƒÂ³n
            Statement sentencia = conexion.createStatement();            
            try {
                /*Hace uso del metodo excuteQuery ocupando la sentencia pasada al metodo como
                parametro*/
                resultado = sentencia.executeQuery(query);
                //sentencia.close();
                //conexion.close();
                }catch(SQLException e){
                    System.out.println("NO SE EJECUTO QUERY!!");
                }
        }catch(SQLException e) {
            System.out.println("NO CONECTO4!!");  
        }
        return resultado;
    }
    
    private static void insertar() throws SQLException, Exception {
        Integer rt ;
        Integer count = 0;
        Integer contador = 2;
        ResultSet fecha = fechainsert();
        fecha.next();
        String f_insert = fecha.getString("SYSDATE");
        f_insert = f_insert.substring(0, 19);
        
        for (HashMap<String, String> row : RowList) {
            String Proyecto         = (String)row.get("Proyecto");
            String Obra             = (String)row.get("Obra");
            String Fechavance       = (String)row.get("Fechavance");
            String FechaServReal    = (String)row.get("FechaServReal");
            String PorcentAvance    = (String)row.get("PorcentAvance");
            //como viene en porcentaje lo multiplico
            
            if (PorcentAvance.isEmpty()){
                PorcentAvance ="0";
            }else {
                Double dato = Double.parseDouble(PorcentAvance);
                dato=dato*100;            
                PorcentAvance = Double.toString(dato);
                if (PorcentAvance.length()>5){
                    PorcentAvance= PorcentAvance.substring(0,5);
                }
            }
            
            String Cables_BT        = (String)row.get("Cables_BT");
            String Lineas_BT        = (String)row.get("Lineas_BT");
            String Cables_MT        = (String)row.get("Cables_MT");           
            String Lineas_MT        = (String)row.get("Lineas_MT");           
            String PotenTrafoMTBT   = (String)row.get("PotenTrafoMTBT");           
            String TransAereosMTBT  = (String)row.get("TransAereosMTBT");           
            String TransNOAereosMTBT= (String)row.get("TransNOAereosMTBT");           
            String TelemandoRedMT   = (String)row.get("TelemandoRedMT");           
            String ExpanCT          = (String)row.get("ExpanCT");           
            String RedAT            = (String)row.get("RedAT");           
            String PotenciaAT       = (String)row.get("PotenciaAT");           
            String EstadoObra       = (String)row.get("EstadoObra");           
            String Observacion      = (String)row.get("Observacion");           
            
            try {
                String sql =
                        "INSERT INTO NEXUS_GIS.WSENRE_CARGA_AVANCE VALUES (  "
                        + "?, "    //Proyecto
                        + "?, "    //Obra
                        + "?, "    //Fechavance
                        + "?, "    //FechaServReal
                        + "?, "    //PorcentAvance
                        + "?, "    //Cables_BT
                        + "?, "    //Lineas_BT
                        + "?, "    //Cables_MT
                        + "?, "    //Lineas_MT
                        + "?, "    //PotenTrafoMTBT
                        + "?, "    //TransAereosMTBT
                        + "?, "    //TransNOAereosMTBT
                        + "?, "    //TelemandoRedMT
                        + "?, "    //ExpanCT
                        + "?, "    //RedAT
                        + "?, "    //PotenciaAT
                        + "?, "    //EstadoObra
                        + "?, "    //Observacion                        
                        + "to_date('"+f_insert+"','yyyy/mm/dd hh24:mi:ss'))";

                try (PreparedStatement ps = connection.prepareStatement(sql)) {
                    ps.setString(1, Proyecto);
                    ps.setString(2, Obra);
                    ps.setString(3, Fechavance);
                    ps.setString(4, FechaServReal);
                    ps.setString(5, PorcentAvance);
                    ps.setString(6, Cables_BT);
                    ps.setString(7, Lineas_BT);
                    ps.setString(8, Cables_MT);
                    ps.setString(9, Lineas_MT);
                    ps.setString(10, PotenTrafoMTBT);
                    ps.setString(11, TransAereosMTBT);
                    ps.setString(12, TransNOAereosMTBT);
                    ps.setString(13, TelemandoRedMT);
                    ps.setString(14, ExpanCT);
                    ps.setString(15, RedAT);
                    ps.setString(16, PotenciaAT);                                        
                    ps.setString(17, EstadoObra);
                    ps.setString(18, Observacion);  
                    rt = ps.executeUpdate();
                    ps.close();
                }
                count += rt;
                contador = contador + 1; 
            } catch (SQLException ex) {
                                throw new Exception  ("Ocurrio un error al intentar insertar los datos en la base, posible error en el Archivo"+ FileName + " Fila: " + contador + " Error interno :" + ex);

            }
        }
        Log.log(Level.INFO, SeInsertaronFilasMsj, count);
    }

    private static void closeConnection() throws SQLException {
        if (connection != null) {
            connection.close();
        }
    }

    private static void moveFromTo(String oriDir, String destDir, String name) throws IOException, Exception {
        
        ResultSet fecha = fechainsert();
        fecha.next();
        String f_insert = fecha.getString("SYSDATE");
        f_insert = f_insert.substring(0, 19);
        f_insert = f_insert.replace(':','_').replace('-','_').replace(" ", "_");
        
       
       
        
 
        Path oriBase = Paths.get(oriDir);
        Path destBase = Paths.get(destDir);
        Path fullPathOri = oriBase.resolve(name);
        name = f_insert +"_" + name; //borrar cuando no sea necesario renombrar el archivo al mover
       // System.out.println(name);
        Path fullPathDest = destBase.resolve(name);
        Files.move(fullPathOri, fullPathDest, REPLACE_EXISTING);
    }

private static String fueProcesadoAnteriormente(String arch) throws IOException, Exception {
        String nomArch = arch;
        Path archivosP = Paths.get(ProcesadosDir);
        Path archivosE = Paths.get(ErrorDir);
        String ret = null;
        int retlegth = 0;
        String filenameorig = null;
        String FechaActual = null;
        String FechaDelArchivo = null;
        String ArchProcesado = null;
        
        
        
        
        ResultSet fecha = fechainsert();
        fecha.next();
        String f_insert = fecha.getString("SYSDATE");
        f_insert = f_insert.substring(0, 10);
        
        
        
        
        try (DirectoryStream<Path> proc = Files.newDirectoryStream(archivosP);
                DirectoryStream<Path> error = Files.newDirectoryStream(archivosE)) {


            for (Path elem : error) {
                    retlegth=elem.getFileName().toString().length();
                    filenameorig = elem.getFileName().toString().substring(20,retlegth);
                if (filenameorig.equals(nomArch)) {
                    ret = "ERROR";
                }
            }
            
                        for (Path elem : proc) {
                    retlegth=elem.getFileName().toString().length();
                    filenameorig = elem.getFileName().toString().substring(20,retlegth);
                    FechaDelArchivo = elem.getFileName().toString().substring(0,10);
                    if (filenameorig.equals(nomArch)){
                        ret = "OK";
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
            Handler fileHandler;
            Formatter simpleFormatter;
            fileHandler  = new FileHandler(LogsDir + "/" + LogFileName, true);
            simpleFormatter = new SimpleFormatter();
            Log.addHandler(fileHandler);
            fileHandler.setLevel(Level.ALL);
            Log.setLevel(Level.ALL);
            fileHandler.setFormatter(simpleFormatter);
    }
    
        private static boolean esFilaVacia(Row row) throws Exception {
           boolean esFilaVacia = true;

           for (Cell campo : row) {
               if (campo.getColumnIndex() == 0) {
                   continue;
               }

               if (!excelTypeToStr(campo).trim().isEmpty()) {
                   esFilaVacia = false;
                   break;
               }
           }

           return esFilaVacia;
       }
    
private static void generar() throws Exception{

                @SuppressWarnings("UseOfObsoleteCollectionType")
		Hashtable<String,String>   hst_Mail= new Hashtable<>();

		String HTML_Estucture;
                @SuppressWarnings("UnusedAssignment")
		boolean enviarMail = false;
		int i = 0;
		
		try{		    
			i=0;
			enviarMail = true;		
         
			hst_Mail.put("mailHost"  , "mail.edenor");
			hst_Mail.put("DE"        , "ITSM_llamadas_salientes@edenor.com");
			hst_Mail.put("PARA"      , "MRRODRIGUEZ@edenor.com");			
			hst_Mail.put("ASUNTO"    , "NOTIFICACION: Poblado de Obras");


			HTML_Estucture  = "<html>";
			HTML_Estucture += "<head>";
			HTML_Estucture += "<style id='Mail_Styles'>";
			HTML_Estucture += "<!--table";
			HTML_Estucture += ".xl1530982";
			HTML_Estucture += "	{padding:0px; 	mso-ignore:padding; 	color:black; 	font-size:11.0pt; 	font-weight:400; 	font-style:normal; 	text-decoration:none; 	font-family:Calibri, sans-serif; 	mso-font-charset:0; 	mso-number-format:General; 	text-align:general; 	vertical-align:bottom; 	mso-background-source:auto; 	mso-pattern:auto; 	white-space:nowrap;}";
			HTML_Estucture += ".xl6530982";
			HTML_Estucture += "	{padding:0px; 	mso-ignore:padding; 	color:black; 	font-size:11.0pt; 	font-weight:700; 	font-style:normal; 	text-decoration:none; 	font-family:Calibri, sans-serif; 	mso-font-charset:0; 	mso-number-format:General; 	text-align:general; 	vertical-align:middle; 	border-top:1.0pt solid windowtext; 	border-right:.5pt solid windowtext; 	border-bottom:1.0pt solid windowtext; 	border-left:1.0pt solid windowtext; 	background:#FFC000; 	mso-pattern:black none; 	white-space:nowrap;}";
			HTML_Estucture += ".xl6630982";
			HTML_Estucture += "	{padding:5px; 	color:black; 	font-size:11.0pt; 	font-weight:700; 	font-style:normal; 	text-decoration:none; 	font-family:Calibri, sans-serif; 	mso-font-charset:0; 	mso-number-format:General; 	text-align:center; 	vertical-align:middle; 	border-top:1.0pt solid windowtext; 	border-right:.5pt solid windowtext; 	border-bottom:1.0pt solid windowtext; 	border-left:.5pt solid windowtext; 	background:#FFC000; 	mso-pattern:black none; 	white-space:nowrap;}";
			HTML_Estucture += ".xl6730982";
			HTML_Estucture += "	{padding:0px; 	mso-ignore:padding; 	color:black; 	font-size:11.0pt; 	font-weight:700; 	font-style:normal; 	text-decoration:none; 	font-family:Calibri, sans-serif; 	mso-font-charset:0; 	mso-number-format:General; 	text-align:right; 	vertical-align:middle; 	border-top:1.0pt solid windowtext; 	border-right:1.0pt solid windowtext; 	border-bottom:1.0pt solid windowtext; 	border-left:.5pt solid windowtext; 	background:#FFC000; 	mso-pattern:black none; 	white-space:nowrap;}";
			HTML_Estucture += ".xl6830982";
			HTML_Estucture += "	{padding:5px; 	mso-ignore:padding; 	color:black; 	font-size:11.0pt; 	font-weight:400; 	font-style:normal; 	text-decoration:none; 	font-family:Calibri, sans-serif; 	mso-font-charset:0; 	mso-number-format:General; 	text-align:general; 	vertical-align:middle; 	border-top:1.0pt solid windowtext; 	border-right:.5pt solid windowtext; 	border-bottom:.5pt solid windowtext; 	border-left:.5pt solid windowtext; 	mso-background-source:auto; 	mso-pattern:auto; 	white-space:nowrap;}";
			HTML_Estucture += ".xl6930982";
			HTML_Estucture += "	{padding:0px; 	mso-ignore:padding; 	color:black; 	font-size:11.0pt; 	font-weight:400; 	font-style:normal; 	text-decoration:none; 	font-family:Calibri, sans-serif; 	mso-font-charset:0; 	mso-number-format:'General Date'; 	text-align:general; 	vertical-align:middle; 	border-top:1.0pt solid windowtext; 	border-right:.5pt solid windowtext; 	border-bottom:.5pt solid windowtext; 	border-left:.5pt solid windowtext; 	mso-background-source:auto; 	mso-pattern:auto; 	white-space:nowrap;}";
			HTML_Estucture += ".xl7030982";
			HTML_Estucture += "	{padding:0px; 	mso-ignore:padding; 	color:black; 	font-size:11.0pt; 	font-weight:400; 	font-style:normal; 	text-decoration:none; 	font-family:Calibri, sans-serif; 	mso-font-charset:0; 	mso-number-format:'Short Time'; 	text-align:general; 	vertical-align:middle; 	border-top:1.0pt solid windowtext; 	border-right:1.0pt solid windowtext; 	border-bottom:.5pt solid windowtext; 	border-left:.5pt solid windowtext; 	mso-background-source:auto; 	mso-pattern:auto; 	white-space:nowrap;}";
			HTML_Estucture += ".xl7130982";
			HTML_Estucture += "	{padding:0px; 	mso-ignore:padding; 	color:black; 	font-size:11.0pt; 	font-weight:400; 	font-style:normal; 	text-decoration:none; 	font-family:Calibri, sans-serif; 	mso-font-charset:0; 	mso-number-format:General; 	text-align:general; 	vertical-align:middle; 	border:.5pt solid windowtext; 	mso-background-source:auto; 	mso-pattern:auto; 	white-space:nowrap;}";
			HTML_Estucture += ".xl7230982";
			HTML_Estucture += "	{padding:0px; 	mso-ignore:padding; 	color:black; 	font-size:11.0pt; 	font-weight:400; 	font-style:normal; 	text-decoration:none; 	font-family:Calibri, sans-serif; 	mso-font-charset:0; 	mso-number-format:'General Date'; 	text-align:general; 	vertical-align:middle; 	border:.5pt solid windowtext; 	mso-background-source:auto; 	mso-pattern:auto; 	white-space:nowrap;}";
			HTML_Estucture += ".xl7330982";
			HTML_Estucture += "	{padding:0px; 	mso-ignore:padding; 	color:black; 	font-size:11.0pt; 	font-weight:400; 	font-style:normal; 	text-decoration:none; 	font-family:Calibri, sans-serif; 	mso-font-charset:0; 	mso-number-format:'Short Time'; 	text-align:general; 	vertical-align:middle; 	border-top:.5pt solid windowtext; 	border-right:1.0pt solid windowtext; 	border-bottom:.5pt solid windowtext; 	border-left:.5pt solid windowtext; 	mso-background-source:auto; 	mso-pattern:auto; 	white-space:nowrap;}";
			HTML_Estucture += ".xl7430982";
			HTML_Estucture += "	{padding:0px; 	mso-ignore:padding; 	color:black; 	font-size:11.0pt; 	font-weight:400; 	font-style:normal; 	text-decoration:none; 	font-family:Calibri, sans-serif; 	mso-font-charset:0; 	mso-number-format:General; 	text-align:general; 	vertical-align:middle; 	border-top:.5pt solid windowtext; 	border-right:.5pt solid windowtext; 	border-bottom:1.0pt solid windowtext; 	border-left:.5pt solid windowtext; 	mso-background-source:auto; 	mso-pattern:auto; 	white-space:nowrap;}";
			HTML_Estucture += ".xl7530982";
			HTML_Estucture += "	{padding:0px; 	mso-ignore:padding; 	color:black; 	font-size:11.0pt; 	font-weight:400; 	font-style:normal; 	text-decoration:none; 	font-family:Calibri, sans-serif; 	mso-font-charset:0; 	mso-number-format:'General Date'; 	text-align:general; 	vertical-align:middle; 	border-top:.5pt solid windowtext; 	border-right:.5pt solid windowtext; 	border-bottom:1.0pt solid windowtext; 	border-left:.5pt solid windowtext; 	mso-background-source:auto; 	mso-pattern:auto; 	white-space:nowrap;}";
			HTML_Estucture += ".xl7630982";
			HTML_Estucture += "	{padding:0px; 	mso-ignore:padding; 	color:black; 	font-size:11.0pt; 	font-weight:400; 	font-style:normal; 	text-decoration:none; 	font-family:Calibri, sans-serif; 	mso-font-charset:0; 	mso-number-format:'Short Time'; 	text-align:general; 	vertical-align:middle; 	border-top:.5pt solid windowtext; 	border-right:1.0pt solid windowtext; 	border-bottom:1.0pt solid windowtext; 	border-left:.5pt solid windowtext; 	mso-background-source:auto; 	mso-pattern:auto; 	white-space:nowrap;}";
			HTML_Estucture += ".xl7730982";
			HTML_Estucture += "	{padding:0px; 	mso-ignore:padding; 	color:black; 	font-size:12.0pt; 	font-weight:400; 	font-style:normal; 	text-decoration:none; 	font-family:'Times New Roman', serif; 	mso-font-charset:0; 	mso-number-format:General; 	text-align:general; 	vertical-align:bottom; 	mso-background-source:auto; 	mso-pattern:auto; 	white-space:nowrap;}";
			HTML_Estucture += ".xl7830982";
			HTML_Estucture += "	{padding:5px; 	mso-ignore:padding; 	color:black; 	font-size:11.0pt; 	font-weight:400; 	font-style:normal; 	text-decoration:none; 	font-family:Calibri, sans-serif; 	mso-font-charset:0; 	mso-number-format:General; 	text-align:general; 	vertical-align:middle; 	border-top:1.0pt solid windowtext; 	border-right:.5pt solid windowtext; 	border-bottom:.5pt solid windowtext; 	border-left:1.0pt solid windowtext; 	background:#C5D9F1; 	mso-pattern:black none; 	white-space:nowrap;}";
			HTML_Estucture += ".xl7930982";
			HTML_Estucture += "	{padding:0px; 	mso-ignore:padding; 	color:black; 	font-size:11.0pt; 	font-weight:400; 	font-style:normal; 	text-decoration:none; 	font-family:Calibri, sans-serif; 	mso-font-charset:0; 	mso-number-format:General; 	text-align:general; 	vertical-align:middle; 	border-top:.5pt solid windowtext; 	border-right:.5pt solid windowtext; 	border-bottom:.5pt solid windowtext; 	border-left:1.0pt solid windowtext; 	background:#C5D9F1; 	mso-pattern:black none; 	white-space:nowrap;}";
			HTML_Estucture += ".xl8030982";
			HTML_Estucture += "	{padding:0px; 	mso-ignore:padding; 	color:black; 	font-size:11.0pt; 	font-weight:400; 	font-style:normal; 	text-decoration:none; 	font-family:Calibri, sans-serif; 	mso-font-charset:0; 	mso-number-format:General; 	text-align:general; 	vertical-align:middle; 	border-top:.5pt solid windowtext; 	border-right:.5pt solid windowtext; 	border-bottom:1.0pt solid windowtext; 	border-left:1.0pt solid windowtext; 	background:#C5D9F1; 	mso-pattern:black none; 	white-space:nowrap;}";
			HTML_Estucture += ".padded";
			HTML_Estucture += "	{padding:5px; 	color:black; 	font-size:10.0pt; 	font-weight:400; 	font-style:normal; 	text-decoration:none; 	font-family:Calibri, sans-serif; 	mso-font-charset:0; 	mso-number-format:General; 	text-align:general; 	vertical-align:middle; 	border-top:.5pt solid windowtext; 	border-right:.5pt solid windowtext; 	border-bottom:1.0pt solid windowtext; 	border-left:1.0pt solid windowtext; 	background:#C5D9F1; 	mso-pattern:black none; 	white-space:nowrap;}";
			HTML_Estucture += "-->";
			HTML_Estucture += "</style>";
			HTML_Estucture += "</head>";
			HTML_Estucture += "<body>";
			HTML_Estucture += "<div id='Mail_body' align=left>";
			HTML_Estucture += "<table border=0 cellpadding=0 cellspacing=0 width=494 style='border-collapse: collapse;table-layout:auto;width:371pt'>";
			HTML_Estucture += " <tr height=21 style='height:15.75pt'>";
                        HTML_Estucture += " </tr>";
            

                String sinDatosMsj = "IMPORTANTE: El presente es un mail automÃƒÂ¡tico informando el poblado de obras de manera exitosa. Ambiente: "+Ambiente+ ", Archivo procesado: "+ nombreArchivo;
    
                HTML_Estucture += "</br></br>";
                HTML_Estucture += " <tr height=21 style='height:15.75pt'>";
                HTML_Estucture += "  <td height=21 class=xl7730982 colspan=9 width=494 style='height:15.75pt;width:600pt'>" + sinDatosMsj + "</td>";
                HTML_Estucture += " </tr>";
                HTML_Estucture += "</table>";                
                HTML_Estucture += " <br>";
                HTML_Estucture += " <hr>";
                HTML_Estucture += " <br>";                
                HTML_Estucture += "<table border=0 cellpadding=5px cellspacing=0 width=494 style='border-collapse: collapse;table-layout:auto;width:371pt'>";
                HTML_Estucture += " <tr height=20 style='height:15.0pt'>";
                HTML_Estucture += "  <td height=20 style='height:15.0pt' align=left valign=top colspan=5 >";
                HTML_Estucture += "  <span style='mso-ignore:vglayout;position:absolute;z-index:1;margin-left:3px;margin-top:3px;width:71px;height:15px'>";
                HTML_Estucture += "  <img width=71 height=15 src='Logo_edenor.gif'  alt='DescripciÃƒÆ’Ã‚Â³n: logo de Edenor' ></span>";
                HTML_Estucture += "  </td>";
                HTML_Estucture += " </tr>";
                HTML_Estucture += "</table>";
                HTML_Estucture += "</div>";
                HTML_Estucture += "</body>";
                HTML_Estucture += "</html>";
      
	        hst_Mail.put("CUERPO"    , HTML_Estucture);			  
			if (enviarMail) {
				mailSender(hst_Mail);			
			}			
			
		}catch(Exception e){
			System.out.println("Error en generar()"+e);
			msgCancela= e.toString();
			hst_Mail.put("CUERPO"    , msgCancela);
			mailSender(hst_Mail);
			throw e;
		}finally{
			try{
				System.out.println("final ");
			}catch(Exception e1){}
		}
}
//------------------------------------------------------------------------------------------------------

public static void mailSender(@SuppressWarnings("UseOfObsoleteCollectionType") Hashtable<String,String> hst_values_mail) throws Exception {
	try {
		
		Properties properties = new Properties();
		properties.put("mail.smtp.host",hst_values_mail.get("mailHost"));
		properties.put("mail.from"     ,hst_values_mail.get("DE"));		
		properties.put("mail.debug"    , "true");

		
		Session session = Session.getInstance(properties, null);
		MimeMessage msg = new MimeMessage(session);

		msg.setFrom(new InternetAddress(hst_values_mail.get("DE")));
		msg.setFrom(InternetAddress.getLocalAddress(session));
		msg.setSubject(hst_values_mail.get("ASUNTO"));
		msg.setSentDate(new java.util.Date());
		
		InternetAddress[] paraArray;
		paraArray= InternetAddress.parse(hst_values_mail.get("PARA"));
		msg.setRecipients(Message.RecipientType.TO,paraArray);
		
		InternetAddress[] ccArray= null;
		if(hst_values_mail.get("CC") != null){
			ccArray= InternetAddress.parse(hst_values_mail.get("CC"));
			msg.setRecipients(Message.RecipientType.CC,ccArray);
		}		
		
		InternetAddress[] bccArray= null;
		if(hst_values_mail.get("CCO") != null){
			bccArray= InternetAddress.parse(hst_values_mail.get("CCO"));
			msg.setRecipients(Message.RecipientType.BCC,bccArray);
		}
		
		MimeMultipart multiParte = new MimeMultipart();		
		
		BodyPart adjunto = new MimeBodyPart();
		adjunto.setDataHandler(new DataHandler(new FileDataSource("Logo_edenor.gif")));
		adjunto.setFileName("Logo_edenor.gif");
		multiParte.addBodyPart(adjunto);
        
		BodyPart texto = new MimeBodyPart();
		texto.setDataHandler(new DataHandler(new HTMLDataSource(hst_values_mail.get("CUERPO"))));
		multiParte.addBodyPart(texto);

		msg.setContent(multiParte);
	
  	        int i,j,k, total;					
		total = paraArray.length;
		if (ccArray!=null) 
		   total+=ccArray.length;
		if (bccArray!=null) 
		   total+=bccArray.length;   

	   InternetAddress[] address= new InternetAddress[total];
		
		for(i=0;i<paraArray.length;i++)
			address[i]= paraArray[i];		
        if (ccArray!=null)			
			for(j=0;j<ccArray.length;j++){
				address[i]= ccArray[j];
				i++; 
			}
		if (bccArray!=null)		
			for(k=0;k<bccArray.length;k++){
				address[i]= bccArray[k];
				i++; 
			}									
							
		Transport transporte = session.getTransport(address[0]);
		transporte.connect();
		transporte.sendMessage(msg,address);

	}catch(SendFailedException e){
		Address[] listaInval= e.getInvalidAddresses();
            for (Address listaInval1 : listaInval) {
                dir_no_encontradas.add(listaInval1.toString());
                System.out.println("No encontrada: " + listaInval1.toString());
            }
	}catch(MessagingException e){
		System.out.println("Exception (mailSender) : "+e);
		throw e;
	}
}

static class HTMLDataSource implements DataSource {
        private final String html;
 
        public HTMLDataSource(String htmlString) {
            html = htmlString;
        }
 
        @Override
        public InputStream getInputStream() throws IOException {
            if (html == null) throw new IOException("Null HTML");
            return new ByteArrayInputStream(html.getBytes());
        }
 
        @Override
        public OutputStream getOutputStream() throws IOException {
            throw new IOException("Este DataHandler no puede crear HTML");
        }
 
        @Override
        public String getContentType() {
            return "text/html";
        }
 
        @Override
        public String getName() {
            return "text/html dataSource para solo enviar e-mail";
        }
    }

}
