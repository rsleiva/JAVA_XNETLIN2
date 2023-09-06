import static java.nio.file.StandardCopyOption.REPLACE_EXISTING;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.net.MalformedURLException;
import java.net.UnknownHostException;
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
import java.sql.Timestamp;
import java.text.DateFormat;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.time.LocalDate;
import java.util.ArrayList;
import java.util.Calendar;
import java.util.Date;
import java.util.HashMap;
import java.util.Iterator;
import java.util.List;
import java.util.Locale;
import java.util.logging.FileHandler;
import java.util.logging.Formatter;
import java.util.logging.Handler;
import java.util.logging.Level;
import java.util.logging.Logger;
import java.util.logging.SimpleFormatter;

import org.apache.poi.hssf.usermodel.HSSFDateUtil;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

import jcifs.smb.NtlmPasswordAuthentication;
import jcifs.smb.SmbException;
import jcifs.smb.SmbFile;
import jcifs.smb.SmbFileInputStream;

import java.io.FileInputStream;
import java.io.IOException;
import java.util.Properties;

public class ParteObras {
    
    /************************************************
     * Adecuar: Ambiente, NetworkDomain,  
     * NetworkFolder, base 
     ************************************************/
    
//Ambiente de la base de datos. Valores: DEV, QA, PRO
    private static String Ambiente = "PRO";
    //Conexion a la carpeta de red
    private static String NetworkDomain = "PRO";   
    //user y pass se pueblan desde la base
    private static String NetworkFolder = "smb://SRVCENFSPS001.PRO.EDENOR/DatosCEN/Iox/GCASIS/DesSistemasTecnicos/WS_ENRE/";
    private static String NetworkUser = ""; 
    private static String NetworkPass = "";
    private static Timestamp fechaProceso = new Timestamp(new Date().getTime());
    //private static java.sql.Date fechaProceso = new java.sql.Date(new Date().getTime());
    
//    private static final String AProcesarDir = "C:\\STSWorkspace\\Partes_Obras\\estructura\\a_procesar";
//    private static final String ProcesadosDir = "C:\\STSWorkspace\\Partes_Obras\\estructura\\procesados";
//    private static final String ErrorDir = "C:\\STSWorkspace\\Partes_Obras\\estructura\\error";
//    private static final String LogsDir = "C:\\STSWorkspace\\Partes_Obras\\estructura\\logs";
    //Se utiliza para simular carpeta de la red    
    //private static final String LocalDebugFolder = "";
    
    //DIRECTORIOS RELATIVOS PARA PRODUCCION    
   private static String base= "/ias/enre748/PartesObras/";    
    private static String AProcesarDir = base  + "estructura/a_procesar";
    private static String ProcesadosDir = base + "estructura/procesados";
    private static String ErrorDir = base  + "estructura/error";
    private static String LogsDir = base  + "estructura/logs";
  
    //Logger
    private static final Logger Log = Logger.getLogger(ParteObras.class.getName());
    private static final String LogFileName = "batchLog.txt";
    //Patron de archivo    
    private static final String Pattern_1 = "02-Parte de Obras_??-??-????.{xls,xlsx}";    
    private static final String RegExPattern = "^02-Parte de Obras_\\d{2}\\-\\d{2}\\-\\d{4}.xlsx?";
    //Driver y string de conexion a diferentes ambientes
    private static String DriverClass = "oracle.jdbc.driver.OracleDriver";
    //private static final String ConnDev = "jdbc:oracle:thin:NEXUS_ENRE/NEXUS_ENRE@tdbs6.tro.edenor:1521:GISDEV01";
    //private static final String ConnQA = "jdbc:oracle:thin:NEXUS_ENRE/NEXUS_ENRE@TDBS5.PRO.EDENOR:1529/GISQA01";
    //private static final String ConnPro = "jdbc:oracle:thin:NEXUS_ENRE/cami0net4@ltronxgisbdpr01.pro.edenor:1528:GISPR01";
    
    //Nombre del libro a leer
    private static final String nombreDeLibro = "Parte_de_Obras";
    
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
    private static final String IniciandoMsj = "---------  Comenzando Proceso Batch         ----------";
    private static final String FinalizandoMsj = "---------  Proceso Finalizado Correctamente ----------";
    private static final String archivoFechaHoyErr = "El archivo encontrado no posee la fecha de hoy, no se procesaran datos";
    //Indices parsear el nombre de archivo 02-Parte de Obras_??-??-????
    private static final Integer FeDiaIni = 18;
    private static final Integer FeDiaFin = 20;
    private static final Integer FeMesIni = 21;
    private static final Integer FeMesFin = 23;
    private static final Integer FeAnioIni = 24;
    private static final Integer FeAnioFin = 28;
    //Definido para el metodo depurar. Indica los archivos a borrar segun fecha.
    private static final Integer DiasHaciaAtras = 30;

    //Tipos de datos de las celdas excel segun apache poi.
    private static final int CELL_TYPE_BLANK = 3;
    private static final int CELL_TYPE_BOOLEAN = 4;
    private static final int CELL_TYPE_ERROR = 5;
    private static final int CELL_TYPE_FORMULA = 2;
    private static final int CELL_TYPE_NUMERIC = 0;
    private static final int CELL_TYPE_STRING = 1;
    private static final Integer CantFilasCabecera = 1;
    
    //
    private static final List<HashMap<String, String>> RowList = new ArrayList<>();
    private static String nombreArchivo = "";
    private static Connection connection = null;

    private static Properties properties;
    
    
////////////////////////////////////////////////////////////////////////////////
/////////////////************       M A I N      ********************///////////
////////////////////////////////////////////////////////////////////////////////
    @SuppressWarnings("ThrowableResultIgnored")
    public static void main(String[] args) throws IOException {

        try {
            findPropiedades();
            if (base==null) {
                Log.log(Level.SEVERE, "FAIL PROPERTY");
            } else {
                inicializarLog();            
                Log.log(Level.INFO, IniciandoMsj);            
                depurarDirectorios();            
                setearConexion();            
                obtenerCredenciales();            
                traerArchivo();
                //traerArchivoDebug();            
                parsearDocumento();            
                //insertar();            
                moveFromTo(AProcesarDir, ProcesadosDir, nombreArchivo);            
                Log.log(Level.INFO, FinalizandoMsj);            
            }
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
                Log.log(Level.SEVERE, MoverErrorMsj, nombreArchivo);
            }
            Log.info(FinalizadoConErrorMsj);
        } catch (IOException e) {
            Log.log(Level.SEVERE, e.getMessage());
            if (nombreArchivo.matches(RegExPattern)){
                moveFromTo(AProcesarDir, ErrorDir, nombreArchivo);
                Log.log(Level.SEVERE, MoverErrorMsj, nombreArchivo);
            }
            Log.info(FinalizadoConErrorMsj);
        } catch (NullPointerException e) {
            Log.log(Level.SEVERE, e.toString());
            if (nombreArchivo.matches(RegExPattern)){
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
                    case archivoFechaHoyErr:
                    	Log.log(Level.INFO, e.getMessage());
                    	Log.info("Finaliza la ejecuci�n");
                    	break;
                    default:
                        if (nombreArchivo.matches(RegExPattern)){
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
    
    public static void findPropiedades(){
        properties= new Properties();
        try (FileInputStream input = new FileInputStream("./config.properties")) {
            properties.load(input);
            Ambiente=properties.getProperty("Ambiente");
            NetworkFolder=properties.getProperty("NetworkFolder");
            NetworkDomain=properties.getProperty("NetworkDomain");
            DriverClass=properties.getProperty("DriverClass");
            base=properties.getProperty("base");
            AProcesarDir = base  + "estructura/a_procesar";
            ProcesadosDir = base + "estructura/procesados";
            ErrorDir = base  + "estructura/error";
            LogsDir = base  + "estructura/logs";
        } catch (IOException e) {
            e.printStackTrace();
            base=null;
        }
    }    
    
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
        
//        switch (fueProcesadoAnteriormente(masReciente.getName())) {
//                case "OK":
//                    throw new Exception(NoSeVuelveAProcesarMsj);
//                case "ERROR":
//                    throw new Exception(NoSeVuelveAProcesarMsjErr);
//        }
        
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
        
        LocalDate aux = LocalDate.now();
        Integer fechaHoy = Integer.parseInt(String.format("%d%02d%02d", aux.getYear(), aux.getMonthValue(), aux.getDayOfMonth()));
        if(feMasReciente.compareTo(fechaHoy) != 0 ) {
        	throw new Exception (archivoFechaHoyErr);
        }
        
        return choice;
    }
    
    
    //obtiene la fecha del reporte parseando nombre de archivo
    public static String obtieneFechaReporte(String fileName){
        String feD = fileName.substring(FeDiaIni,FeDiaFin);
        String feM = fileName.substring(FeMesIni,FeMesFin);
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
    

    private static void setearConexion() throws SQLException, Exception{
        String conStr;
        Log.log(Level.INFO, "Conectando al ambiente {0}", Ambiente);
        switch (Ambiente) {
            case "DEV":
                conStr = properties.getProperty("oracle_cnn_dev");
                break;
            case "QA":
                conStr = properties.getProperty("oracle_cnn_qaa");
                break;
            case "PRO":
                conStr = properties.getProperty("oracle_cnn_pro");
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
            throw new Exception("formatNumeric: el tipo de datos del rdo de la fórmula debe ser Numeric");
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
        String dato ="";
        Date fe;
        SimpleDateFormat mesAnio = new SimpleDateFormat("dd/MM/yyyy");
        switch (cell.getCellType()) {
            case CELL_TYPE_BLANK:
                System.out.println("Entro por CELL_TYPE_BLANK: " );  
                break;
            case CELL_TYPE_BOOLEAN:
                break;
            case CELL_TYPE_ERROR:
                break;
            case CELL_TYPE_FORMULA:
                switch (cell.getCachedFormulaResultType()) {
                    case Cell.CELL_TYPE_STRING:
                        dato = cell.getStringCellValue();
                        if (dato == null){
                            cellAux ="";
                        }else{
                            cellAux=dato;
                        }
                          System.out.println("Entro por CELL_TYPE_FORMULA + CELL_TYPE_STRING: " );  
                        //return cell.getRichStringCellValue().getString();
                        break;
                    case Cell.CELL_TYPE_NUMERIC:
                        System.out.println("celda NUMERO: "+cell.getNumericCellValue() );
                        return Double.toString(Math.round(cell.getNumericCellValue()));
                        
                    default:
                        break;
                }
            case CELL_TYPE_NUMERIC:
                if( HSSFDateUtil.isCellDateFormatted(cell) ){
                    fe = cell.getDateCellValue();
                    cellAux = (fe != null) ? mesAnio.format(fe) : ""; 
                     System.out.println("Entro por CELL_TYPE_NUMERIC if: " );  
                    break;
                }else{
                    Double numEnt = cell.getNumericCellValue();
                    cellAux = Double.toString(numEnt);	
                    Double.toString(cell.getNumericCellValue());
                     System.out.println("Entro por CELL_TYPE_NUMERIC else: " );  
		    break;
                }
            case CELL_TYPE_STRING:
                ///cellAux = cell.getStringCellValue();
                dato = cell.getStringCellValue();
                if (dato == null){
                            cellAux ="";
                        }else{
                            cellAux=dato;
                        }
                System.out.println("Entro por string: " );         
                        //return cell.getRichStringCellValue().getString();
                        break;
                
               
        }
        System.out.println("celda1: "+cellAux);
        return cellAux;
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
        Sheet sheet = workbook.getSheet(nombreDeLibro);
        
        BasicFileAttributes attrs = Files.readAttributes(archAProcesar, BasicFileAttributes.class);
        String fechaModificado = attrs.lastModifiedTime().toString();
        String feMod = fechaModificado.substring(0, 19).replaceAll("T", " ");
        
        
        String extension = "";
        int i = archAProcesar.getFileName().toString().lastIndexOf('.');
        if (i >= 0) {
            extension = archAProcesar.getFileName().toString().substring(i+1);
        }
        //Para que considere las filas sin datos.             
        Integer contador = 0;
        //****************** INICIO **********************************
        Iterator<Row> rowIterator = sheet.iterator();
        Integer rt ;
        Integer count = 0;
        String Fecha = null;
        String Proyecto = null;
        String Obra = null;
        String FechaCorteestm = null;
        String FechaPstaServ = null;
        String Inspector = null;
        String Tel_Inspector = null;
        String Partido = null;
        String Localidad = null;
        String Barrio = null;
        String Calle = null;
        String Numero = null;
        String Coordenadax = null;
        String Coordenaday = null;
        Row row;
        HashMap<String, String> excelRow = new HashMap<>();
        // se recorre cada fila hasta el final
        while (rowIterator.hasNext()) {
            row = rowIterator.next();
            if (row.getRowNum()>0){
                
                
                // row = rowIterator.next();
                //se obtiene las celdas por fila
                Iterator<Cell> cellIterator = row.cellIterator();
                Cell cell;
                int rowcolumn;
                
                while (cellIterator.hasNext()) {
                    // se obtiene la celda en específico y se la imprime
                    
                    cell = cellIterator.next();
                    rowcolumn= cell.getColumnIndex();
                    //System.out.print(rowcolumn);
                    if (rowcolumn ==0) {
                        Fecha  = cell.getStringCellValue();
                    }
                    if (rowcolumn ==1) {
                        Proyecto  = cell.getStringCellValue();
                        
                    }
                    if (rowcolumn ==2) {
                        Obra  = cell.getStringCellValue();
                        
                    }
                    if (rowcolumn ==3) {
                        FechaCorteestm  = cell.getStringCellValue();
                        
                    }
                    if (rowcolumn ==4) {
                        FechaPstaServ  = cell.getStringCellValue();
                        
                    }
                    if (rowcolumn ==5) {
                        Inspector  = cell.getStringCellValue();
                        
                    }
                    if (rowcolumn ==6) {
                        Tel_Inspector  = cell.getStringCellValue();
                        
                    }
                    if (rowcolumn ==7) {
                        Partido  = cell.getStringCellValue();
                        
                    }
                    if (rowcolumn ==8) {
                        Localidad  = cell.getStringCellValue();
                        
                    }
                    if (rowcolumn ==9) {
                        Barrio  = cell.getStringCellValue();
                        
                    }
                    if (rowcolumn ==10) {
                        Calle  = cell.getStringCellValue();
                        
                    }
                    if (rowcolumn ==11) {
                        Numero  = cell.getStringCellValue();
                        
                    }
                    if (rowcolumn ==12) {
                        Coordenadax  = cell.getStringCellValue();
                        
                    }
                    if (rowcolumn ==13) {
                        Coordenaday  = cell.getStringCellValue();
                    }
                    System.out.print(cell.getStringCellValue()+" | ");
                    
                }
                //se recorre cada celda
                System.out.println();
                //RowList.add(excelRow);
                
                //************
                //*************
                
                // System.out.println(RowList.get(i).get("Obra"));
                
                
                
                String sql =
                        "INSERT INTO NEXUS_GIS.WSENRE_PARTE_OBRAS VALUES (  "
                        + "?, "    //Fecha
                        + "?, "    //Proyecto
                        + "?, "    //Obra
                        + "?, "    //FechaCorteestm
                        + "?, "    //FechaPstaServ
                        + "?, "    //Inspector
                        + "?, "    //Tel_Inspector
                        + "?, "    //Partido
                        + "?, "    //Localidad
                        + "?, "    //Barrio
                        + "?, "    //Calle
                        + "?, "    //Numero
                        + "?, "    //Coordenadax
                        + "?, "    //Coordenaday
                        + "?, 'Parte_de_obras')";
                
                try (PreparedStatement ps = connection.prepareStatement(sql)) {
                    ps.setString(1, Fecha);
                    ps.setString(2, Proyecto);
                    ps.setString(3, Obra);
                    ps.setString(4, FechaCorteestm);
                    ps.setString(5, FechaPstaServ);
                    ps.setString(6, Inspector);
                    ps.setString(7, Tel_Inspector);
                    ps.setString(8,Partido);
                    ps.setString(9,Localidad);
                    ps.setString(10,Barrio);
                    ps.setString(11,Calle);
                    ps.setString(12,Numero);
                    ps.setString(13,Coordenadax);
                    ps.setString(14,Coordenaday);
                    ps.setTimestamp(15, fechaProceso);
                    rt = ps.executeUpdate();
                }
                
                  count += rt;
                System.out.println("Empiezo a imprimir: "+i);
                System.out.println();
            }
            Log.log(Level.INFO, SeInsertaronFilasMsj, count);
            
            
            
            //**************
            //************
            
        }
         
    
    }
 
    //@SuppressWarnings("empty-statement")
    /*private static void insertar() throws SQLException, Exception {
        Integer rt ;
        Integer count = 0;
        
       for (int i = 0; i < RowList.size(); ++i){
        //for (count to RowList.size()) {
            //System.out.println(RowList.get(i).get("Fecha"));
            System.out.println(RowList.get(i).get("Obra"));
            String Fecha            = RowList.get(i).get("Fecha");
            String Proyecto         = RowList.get(i).get("Proyecto");
            String Obra             = RowList.get(i).get("Obra");
            String FechaCorteestm   = RowList.get(i).get("FechaCorteestm");
            String FechaPstaServ   = RowList.get(i).get("FechaPstaServ");           
            String Inspector        = RowList.get(i).get("Inspector");
            String Tel_Inspector    = RowList.get(i).get("Tel_Inspector");           
            String Partido          = RowList.get(i).get("Partido");
            String Localidad        = RowList.get(i).get("Localidad");
            String Barrio           = RowList.get(i).get("Barrio");
            String Calle            = RowList.get(i).get("Calle");
            String Numero           = RowList.get(i).get("Numero");
            String Coordenadax      =RowList.get(i).get("Coordenadax");
            String Coordenaday      = RowList.get(i).get("Coordenaday");
       /*
             String sql =
            "INSERT INTO NEXUS_GIS.WSENRE_PARTE_OBRAS VALUES (  "
            + "?, "    //Fecha
            + "?, "    //Proyecto
            + "?, "    //Obra
            + "?, "    //FechaCorteestm
            + "?, "    //FechaPstaServ
            + "?, "    //Inspector
            + "?, "    //Tel_Inspector
            + "?, "    //Partido
            + "?, "    //Localidad
            + "?, "    //Barrio
            + "?, "    //Calle
            + "?, "    //Numero
            + "?, "    //Coordenadax
            + "?, "    //Coordenaday
            + "SYSDATE, 'Parte_de_obras')";
            
            try (PreparedStatement ps = connection.prepareStatement(sql)) {
            ps.setString(1, Fecha);
            ps.setString(2, Proyecto);
            ps.setString(3, Obra);
            ps.setString(4, FechaCorteestm);
            ps.setString(5, FechaPstaServ);
            ps.setString(6, Inspector);
            ps.setString(7, Tel_Inspector);
            ps.setString(8,Partido);
            ps.setString(9,Localidad);
            ps.setString(10,Barrio);
            ps.setString(11,Calle);
            ps.setString(12,Numero);
            ps.setString(13,Coordenadax);
            ps.setString(14,Coordenaday);
            rt = ps.executeUpdate();
            }
            */
            //  count += rt;
    /*        System.out.println("Empiezo a imprimir: "+i);
           System.out.println();
        }
        Log.log(Level.INFO, SeInsertaronFilasMsj, count);
    }*/

    private static void closeConnection() throws SQLException {
        if (connection != null) {
            connection.close();
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
}
