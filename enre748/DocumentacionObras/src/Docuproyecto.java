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
import java.text.NumberFormat;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.*;
import java.util.Date;
import java.util.logging.Formatter;
import java.util.logging.*;
import jcifs.smb.NtlmPasswordAuthentication;
import jcifs.smb.SmbException;
import jcifs.smb.SmbFile;
import jcifs.smb.SmbFileInputStream;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

public class Docuproyecto {
    
    /************************************************
     * Adecuar: Ambiente, NetworkDomain,  
     * NetworkFolder, base 
     ************************************************/
    
    //Ambiente de la base de datos. Valores: DEV, QA, PRO
    private static final String Ambiente = "PRO";
    //Conexion a la carpeta de red
    private static final String NetworkDomain = "PRO";    
    //user y pass se pueblan desde la base
    private static final String NetworkFolder = "smb://SRVCENFSPS001.PRO.EDENOR/DatosCEN/Iox/GCASIS/DesSistemasTecnicos/WS_ENRE/";
    private static String NetworkUser = ""; 
    private static String NetworkPass = "";
    
      /*     
    private static final String AProcesarDir = "C:\\Temp\\parte_obras\\estructura\\a_procesar";
    private static final String ProcesadosDir = "C:\\Temp\\parte_obras\\estructura\\procesados";
    private static final String ErrorDir = "C:\\Temp\\parte_obras\\estructura\\error";
    private static final String LogsDir = "C:\\Temp\\parte_obras\\estructura\\logs";
    */
   
    
    //Se utiliza para simular carpeta de la red
    //private static final String LocalDebugFolder = "";
    
    //DIRECTORIOS RELATIVOS PARA PRODUCCION    
  
  
    private static final String base= "/ias/enre748/DocumentacionObras/";    
    private static final String AProcesarDir = base  + "estructura/a_procesar";
    private static final String ProcesadosDir = base + "estructura/procesados";
    private static final String ErrorDir = base  + "estructura/error";
    private static final String LogsDir = base  + "estructura/logs";
 
    //Logger
    private static final Logger Log = Logger.getLogger(Docuproyecto.class.getName());
    private static final String LogFileName = "batchLog.txt";
    //Patron de archivo
    //private static final String Pattern_1 = "01-DocumentaciÃ³n_??-??-????.{xls,xlsx}";
    private static final String Pattern_1 = "01-Documentación_??-??-????.{xls,xlsx}";
    //private static final String RegExPattern = "^01-DocumentaciÃ³n_\\d{2}\\-\\d{2}\\-\\d{4}.xlsx?";
    private static final String RegExPattern = "^01-Documentación_\\d{2}\\-\\d{2}\\-\\d{4}.xlsx?";
    //Driver y string de conexion a diferentes ambientes
    private static final String DriverClass = "oracle.jdbc.driver.OracleDriver";
    private static final String ConnDev = "jdbc:oracle:thin:NEXUS_ENRE/NEXUS_ENRE@tdbs6.tro.edenor:1521:GISDEV01";
    private static final String ConnQA = "jdbc:oracle:thin:NEXUS_ENRE/NEXUS_ENRE@TDBS5.PRO.EDENOR:1529:GISQA01";
    private static final String ConnPro = "jdbc:oracle:thin:NEXUS_ENRE/cami0net4@ltronxgisbdpr01.pro.edenor:1528:GISPR01";
    
    //Nombre del libro a leer
    private static final String nombreDeLibro = "Documentacion";
    
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
    private static final String ObtenerCredencialesErrMsj = "Error al obtener usuario y contraseÃ±a";
    private static final String FueCopiadoOkMsj = "el archivo {0} fue copiado desde el repositorio remoto";
    private static final String SeInsertaronFilasMsj = "Se insertaron {0} filas.";
    private static final String IniciandoMsj = "---------  Comenzando Proceso Batch         ----------";
    private static final String FinalizandoMsj = "---------  Proceso Finalizado Correctamente ----------";
    //Indices parsear el nombre de archivo 01-DocumentaciÃ³n_??-??-????
    private static final Integer FeDiaIni = 17;
    private static final Integer FeDiaFin = 19;
    private static final Integer FeMesIni = 20;
    private static final Integer FeMesFin = 22;
    private static final Integer FeAnioIni = 23;
    private static final Integer FeAnioFin = 27;
    //Definido para el metodo depurar. Indica los archivos a borrar segun fecha.
    private static final Integer DiasHaciaAtras = 30;
    //Ultima fila que va a leer
    private static final Integer UltimaFila = 39;
    //Tipos de datos de las celdas excel segun apache poi.
    private static final int CELL_TYPE_BLANK = 3;
    private static final int CELL_TYPE_BOOLEAN = 4;
    private static final int CELL_TYPE_ERROR = 5;
    private static final int CELL_TYPE_FORMULA = 2;
    private static final int CELL_TYPE_NUMERIC = 0;
    private static final int CELL_TYPE_STRING = 1;
    
    //
    private static final List<HashMap<String, String>> RowList = new ArrayList<>();
    private static String nombreArchivo = "";
    private static Connection connection = null;
    private static String FileName = null;
    private static String FlagProcess = null;
    
    
////////////////////////////////////////////////////////////////////////////////
/////////////////************       M A I N      ********************///////////
////////////////////////////////////////////////////////////////////////////////
    
    public static void main(String[] args) throws IOException, Exception {

        try {
            
            inicializarLog();
            Log.log(Level.INFO, IniciandoMsj);
            depurarDirectorios();
            setearConexion();
            obtenerCredenciales();
            traerArchivo();
            
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
             
           throw new Exception("No se encontraron archivos para procesar del tipo: DocumentaciÃ³n" ); 
                     
             
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
            throw new Exception("formatNumeric: el tipo de datos del rdo de la fÃ³rmula debe ser Numeric");
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
    private static String excelTypeToStr(Cell cell, String formato) throws Exception {
        String cellAux = "";
      try {  
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
              // cellAux = formatNumeric(cell, formato);
                 Double numEnt = cell.getNumericCellValue();
                cellAux = Double.toString(numEnt);
                break;
            case CELL_TYPE_STRING:
                cellAux = cell.getStringCellValue();
                
                break;
        }
        
        
        
                            } catch (NullPointerException e){
                            cellAux = ""; 
                            return cellAux;
                    }
      return cellAux;
      }
    
    
    private static void parsearDocumento()
            throws FileNotFoundException, IOException, InvalidFormatException,
            Exception {
     
        try {  
        
        Path masReciente = obtenerUltimoModificado(AProcesarDir);

        nombreArchivo = masReciente.getFileName().toString();
        
        Path baseP = Paths.get(AProcesarDir);
        Path archAProcesar = baseP.resolve(nombreArchivo);
        
        String archAProcStr = archAProcesar.toString();
        FileInputStream file = new FileInputStream(new File(archAProcStr));

        Workbook workbook = WorkbookFactory.create(file);
        Sheet sheet = workbook.getSheet(nombreDeLibro);
        int numFilas = sheet.getPhysicalNumberOfRows();//sheet.getLastRowNum();
        BasicFileAttributes attrs = Files.readAttributes(archAProcesar, BasicFileAttributes.class);
        String fechaModificado = attrs.lastModifiedTime().toString();
        String feMod = fechaModificado.substring(0, 19).replaceAll("T", " ");
        
        String extension = "";
        int i = archAProcesar.getFileName().toString().lastIndexOf('.');
        if (i >= 0) {
            extension = archAProcesar.getFileName().toString().substring(i+1);
        }
        
      
            for(Integer j=1; j <= numFilas; j++){
                //System.out.println("Fila: "+j);
                try{
                    //System.out.println(j+" : "+sheet.getRow(j).getCell(0).getStringCellValue().equalsIgnoreCase(null));
                    if (sheet.getRow(j).getCell(0).getStringCellValue().equalsIgnoreCase(null)){
                        //Aqui estoy verificando que se pueda leer la celda, el valor que viene es falso
                        break;
                    }
                
                    } catch (NullPointerException e){
                            //throw new NullPointerException ("read: Error al parsear el excel - " + e);
                            //cuando no se pueda leer la celda porque esta vacia se va por excepcion y salgo del bucle
                            break;
                    } catch (Exception e){
                            //throw new Exception ("read: Error al parsear el excel - " + e);
                            //cuando no se pueda leer la celda porque esta vacia se va por excepcion y salgo del bucle                            
                            break;
                    }
                


                
                Cell Proyecto = sheet.getRow(j).getCell(0); 
                Cell Obra = sheet.getRow(j).getCell(1);
                Cell Denominacion = sheet.getRow(j).getCell(2);
                Cell Mant_Correctiv = sheet.getRow(j).getCell(3);
                Cell Mot_Mant_Corre = sheet.getRow(j).getCell(4);
                Cell FechaInicioEs = sheet.getRow(j).getCell(5);
                Cell FechaFinEs = sheet.getRow(j).getCell(6);                
                Cell Clasificacion = sheet.getRow(j).getCell(7);
                Cell Niveldetension = sheet.getRow(j).getCell(8);
                Cell CablesBT = sheet.getRow(j).getCell(9);
                Cell LineasBT = sheet.getRow(j).getCell(10);
                Cell CablesMT = sheet.getRow(j).getCell(11);
                Cell LineasMT = sheet.getRow(j).getCell(12);
                Cell PotTrafoMTBT = sheet.getRow(j).getCell(13);
                Cell TransAereoMTBT = sheet.getRow(j).getCell(14);
                Cell TransNoAerMTBT = sheet.getRow(j).getCell(15);
                Cell TelemandoMT = sheet.getRow(j).getCell(16);
                Cell ExpansionCT = sheet.getRow(j).getCell(17);
                Cell RedAT = sheet.getRow(j).getCell(18);
                Cell PotenciaAT = sheet.getRow(j).getCell(19);
                Cell OtrosEquipam = sheet.getRow(j).getCell(20);
                Cell Partido = sheet.getRow(j).getCell(21);
                Cell Localidad = sheet.getRow(j).getCell(22);
                Cell Barrio = sheet.getRow(j).getCell(23);
                Cell ZonaTecnica = sheet.getRow(j).getCell(24);
                Cell Calle = sheet.getRow(j).getCell(25);
                Cell Numero = sheet.getRow(j).getCell(26);
                Cell FechaPtaServ = sheet.getRow(j).getCell(27);
                Cell FechaBaja = sheet.getRow(j).getCell(28);
                Cell MontoTotal = sheet.getRow(j).getCell(29);
                Cell Coordenadax = sheet.getRow(j).getCell(30);
                Cell Coordenaday = sheet.getRow(j).getCell(31);
                Cell ExInterrupcion = sheet.getRow(j).getCell(32);
                Cell Observaciones = sheet.getRow(j).getCell(33);
               

                HashMap<String, String> excelRow = new HashMap<>();
                
                 if (Proyecto.toString().equals("")|| (Proyecto.toString().equals(null))|| (Proyecto.toString().toLowerCase().equals("proyecto"))){
                 throw new Exception ("Error de formato en el archivo " + FileName + " la primera fila o columna estan en blanco.");
                 
             }
                excelRow.put("Proyecto", excelTypeToStr(Proyecto, "String"));
                

                
                
                //System.out.println("Proyecto");
                excelRow.put("Obra", excelTypeToStr(Obra, "String"));
                //System.out.println("Obra");
                excelRow.put("DenominaciÃ³n", excelTypeToStr(Denominacion, "String"));
                //System.out.println("DenominaciÃ³n    ");
                excelRow.put("Mantenimiento Correctivo", excelTypeToStr(Mant_Correctiv, "String"));
                //System.out.println("Mantenimiento Correctivo");
                excelRow.put("Motivo Mantenimiento Correctivo", excelTypeToStr(Mot_Mant_Corre, "String"));
                //System.out.println("Motivo Mantenimiento Correctivo");
                   if (!FechaInicioEs.toString().equals("-")){                        
                        excelRow.put("Fecha Inicio Estimada", FechaInicioEs.toString());
                         //System.out.println("Fecha Inicio Estimada a");
                    }else{
                        excelRow.put("Fecha Inicio Estimada", excelTypeToStr(FechaInicioEs, "String"));
                         //System.out.println("Fecha Inicio Estimada B");
                    }
                   if (!FechaFinEs.toString().equals("-")){                        
                        excelRow.put("Fecha Fin Estimada", FechaFinEs.toString());
                        //System.out.println("Fecha Fin Estimada a");
                    }else{
                        excelRow.put("Fecha Fin Estimada", excelTypeToStr(FechaFinEs, "String"));
                        //System.out.println("Fecha Fin Estimada B");
                    }                
                excelRow.put("ClasificaciÃ³n", excelTypeToStr(Clasificacion, "String"));
                //System.out.println("ClasificaciÃ³n");
                excelRow.put("Nivel de tensiÃ³n", excelTypeToStr(Niveldetension, "String"));
                //System.out.println("Nivel de tensiÃ³n");
                excelRow.put("Cables BT*", excelTypeToStr(CablesBT, "String"));
                //System.out.println("Cables BT*");
                excelRow.put("LÃ­neas BT*", excelTypeToStr(LineasBT, "String"));
                //System.out.println("LÃ­neas BT*");
                excelRow.put("Cables MT*", excelTypeToStr(CablesMT, "String"));
                //System.out.println("Cables MT*");
                excelRow.put("LÃ­neas MT*", excelTypeToStr(LineasMT, "String"));
                //System.out.println("LÃ­neas MT*");
                excelRow.put("Potencia Trafo MT/BT", excelTypeToStr(PotTrafoMTBT, "String"));
                //System.out.println("Potencia Trafo MT/BT");
                excelRow.put("Transformador AÃ©reo MT/BT*", excelTypeToStr(TransAereoMTBT, "String"));
                //System.out.println("Transformador AÃ©reo MT/BT*");
                excelRow.put("Transformador No AÃ©reo MT/BT*", excelTypeToStr(TransNoAerMTBT, "String"));
                //System.out.println("Transformador No AÃ©reo MT/BT*");
                excelRow.put("Telemando MT*", excelTypeToStr(TelemandoMT, "String"));
                //System.out.println("Telemando MT*");
                excelRow.put("ExpansiÃ³n CT*", excelTypeToStr(ExpansionCT, "String"));
                //System.out.println("ExpansiÃ³n CT*");
                excelRow.put("Red AT*", excelTypeToStr(RedAT, "String"));
                //System.out.println("Red AT*");
                excelRow.put("Potencia AT*", excelTypeToStr(PotenciaAT, "String"));
                //System.out.println("Potencia AT*");
                excelRow.put("Otros Equipamientos", excelTypeToStr(OtrosEquipam, "String"));
                //System.out.println("Otros Equipamientos");
                excelRow.put("Partido", excelTypeToStr(Partido, "String"));
                //System.out.println("Partido");
                excelRow.put("Localidad", excelTypeToStr(Localidad, "String"));
                //System.out.println("Localidad");
                excelRow.put("Barrio", excelTypeToStr(Barrio, "String"));
                //System.out.println("Barrio");
                excelRow.put("Zona TÃ©cnica", excelTypeToStr(ZonaTecnica, "String"));
                //System.out.println("Zona TÃ©cnica");
                excelRow.put("Calle", excelTypeToStr(Calle, "String"));
                //System.out.println("Calle");
                excelRow.put("NÃºmero", excelTypeToStr(Numero, "String"));
                //System.out.println("NÃºmero");
                if (!FechaPtaServ.toString().equals("")){                        
                    excelRow.put("Fecha Puesta Servicio", FechaPtaServ.toString());
                    //System.out.println("Fecha Puesta Servicio a");
                }else{
                    excelRow.put("Fecha Puesta Servicio", excelTypeToStr(FechaPtaServ, "String"));
                    //System.out.println("Fecha Puesta Servicio B");
                }
                if (!FechaBaja.toString().equals("")){                    
                    excelRow.put("Fecha Baja", FechaBaja.toString());
                    //System.out.println("Fecha Baja A");
                }else{
                    excelRow.put("Fecha Baja", excelTypeToStr(FechaBaja, "String"));
                    //System.out.println("Fecha Baja B");
                }
                //System.out.println("MontoTotal ANTES: "+MontoTotal.toString().trim());
                /*if (!MontoTotal.toString().trim().equals("")){
                     
                    Locale locale = new Locale("es","AR"); // elegimos Argentina
                    NumberFormat nf = NumberFormat.getCurrencyInstance(locale);
                        
                    //excelRow.put("Monto Total", nf.format(MontoTotal.getNumericCellValue()));
                    excelRow.put("Monto Total", excelTypeToStr(MontoTotal, "String"));
                    System.out.println("Monto Total A");
                    
                }else{*/
                     
                    excelRow.put("Monto Total", excelTypeToStr(MontoTotal, "String"));
                    //System.out.println("Monto Total B");
                //}                
                excelRow.put("Coordenada (x)", excelTypeToStr(Coordenadax, "String"));
                //System.out.println("Coordenada (x)");
                excelRow.put("Coordenada (y)", excelTypeToStr(Coordenaday, "String"));
                //System.out.println("Coordenada (y)");
                excelRow.put("ExclusiÃ³n InterrupciÃ³n", excelTypeToStr(ExInterrupcion, "String"));
                //System.out.println("ExclusiÃ³n InterrupciÃ³n");
                excelRow.put("Observaciones", excelTypeToStr(Observaciones, "String"));
                //System.out.println("Observaciones");
                
                RowList.add(excelRow);
                
               //System.out.println("***** FIN *****Fila: "+j);
               
            }
//            System.out.println("***** FIN *****Fila: ");
            
        } catch (NullPointerException e){
            
             throw new Exception ("No se pudo procesar el archivo " + FileName + " compruebe que el nombre del libro sea el correcto. Codigo Interno:"  +  e.toString());
            
        } catch (Exception e){
            throw new Exception ("read: Error al parsear el excel - " + e);
        }

    }
    
    public static ResultSet fechainsert()throws Exception {        
        ResultSet resultado = null;
       
        try{
            Connection conexion = connection;
            String q = "select sysdate from dual";

            String query = q;
            //Creo una sentencia a partir de la conexiÃ³n
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
        ResultSet fecha = fechainsert();
        fecha.next();
        String f_insert = fecha.getString("SYSDATE");
        f_insert = f_insert.substring(0, 19);
        Integer contador =2;
        
        
      
        
        
        for (HashMap<String, String> row : RowList) {

            String Proyecto                             = (String)row.get("Proyecto");
            String Obra                                 = (String)row.get("Obra");
            String Denominacion                         = (String)row.get("DenominaciÃ³n");
            String Mant_Correctiv                       = (String)row.get("Mantenimiento Correctivo");
            String Mot_Mant_Corre                       = (String)row.get("Motivo Mantenimiento Correctivo");
            String FechaInicioEs                        = (String)row.get("Fecha Inicio Estimada");
            String FechaFinEs                           = (String)row.get("Fecha Fin Estimada");            
            String Clasificacion                        = (String)row.get("ClasificaciÃ³n");
            String Niveldetension                       = (String)row.get("Nivel de tensiÃ³n"); 
            String CablesBT                             = (String)row.get("Cables BT*");
            String LineasBT                             = (String)row.get("LÃ­neas BT*");
            String CablesMT                             = (String)row.get("Cables MT*");
            String LineasMT                             = (String)row.get("LÃ­neas MT*");
            String PotTrafoMTBT                         = (String)row.get("Potencia Trafo MT/BT");
            String TransAereoMTBT                       = (String)row.get("Transformador AÃ©reo MT/BT*");
            String TransNoAerMTBT                       = (String)row.get("Transformador No AÃ©reo MT/BT*");
            String TelemandoMT                          = (String)row.get("Telemando MT*");
            String ExpansionCT                          = (String)row.get("ExpansiÃ³n CT*");
            String RedAT                                = (String)row.get("Red AT*");
            String PotenciaAT                           = (String)row.get("Potencia AT*");
            String OtrosEquipam                         = (String)row.get("Otros Equipamientos");
            String Partido                              = (String)row.get("Partido");
            if (Partido.length()>29){                
                Partido= Partido.substring(0, 29);                
            }
            String Localidad                            = (String)row.get("Localidad");
             if (Localidad.length()>29){                
                Localidad= Localidad.substring(0, 29);                
            }
            String Barrio                               = (String)row.get("Barrio");
            String ZonaTecnica                          = (String)row.get("Zona TÃ©cnica");
            String Calle                                = (String)row.get("Calle");
            if (Calle.length()>29){                
                Calle= Calle.substring(0, 29);                
            }
            String Numero                               = (String)row.get("NÃºmero");            
            Integer pos = Numero.indexOf(".");            
            if (pos!=-1){                
                Numero = Numero.substring(0,pos);                
            }
            
            String FechaPtaServ                         = (String)row.get("Fecha Puesta Servicio");
            String FechaBaja                            = (String)row.get("Fecha Baja");
            String MontoTotal                           = (String)row.get("Monto Total");
            
            if (!MontoTotal.isEmpty()){
               // MontoTotal= MontoTotal.substring(1);
                MontoTotal = MontoTotal.replace(".", "");
                MontoTotal = MontoTotal.replace(",", ".");
            }
            
            //parsemonto = MontoTotal.replace(",", "");
            
            String Coordenadax                          = (String)row.get("Coordenada (x)");
            String Coordenaday                          = (String)row.get("Coordenada (y)");
            Coordenaday= Coordenaday.replace("E", "");
            String ExInterrupcion                       = (String)row.get("ExclusiÃ³n InterrupciÃ³n");
            String Observaciones                        = (String)row.get("Observaciones");
            if (Observaciones.trim().length()>19){                
                Observaciones= Observaciones.substring(0, 19);                
            }
            
            try {
                String sql =
                        "INSERT INTO NEXUS_GIS.WSENRE_DOC_OBRA VALUES (  "
                        + "?, "    //Proyecto      
                        + "?, "    //Obra          
                        + "?, "    //Denominacion  
                        + "?, "    //Mant_Correctiv
                        + "?, "    //Mot_Mant_Corre
                        + "?, "    //FechaInicioEs 
                        + "?, "    //FechaFinEs                           
                        + "?, "    //Clasificacion 
                        + "?, "    //Niveldetension                        
                        + "?, "    //CablesBT 
                        + "?, "    //LineasBT
                        + "?, "    //CablesMT 
                        + "?, "    //LineasMT      
                        + "?, "    //PotTrafoMTBT  
                        + "?, "    //TransAereoMTBT
                        + "?, "    //TransNoAerMTBT
                        + "?, "    //TelemandoMT   
                        + "?, "    //ExpansionCT   
                        + "?, "    //RedAT         
                        + "?, "    //PotenciaAT    
                        + "?, "    //OtrosEquipam  
                        + "?, "    //Partido       
                        + "?, "    //Localidad     
                        + "?, "    //Barrio        
                        + "?, "    //ZonaTecnica   
                        + "?, "    //Calle         
                        + "?, "    //Numero        
                        + "?, "    //FechaPtaServ  
                        + "?, "    //FechaBaja     
                        + "?, "    //MontoTotal    
                        + "?, "    //Coordenadax   
                        + "?, "    //Coordenaday   
                        + "?, "    //ExInterrupcion
                        + "?, "    //Observaciones                         
                        + "to_date('"+f_insert+"','yyyy/mm/dd hh24:mi:ss'))";

                try (PreparedStatement ps = connection.prepareStatement(sql)) {
                    ps.setString(1, Proyecto);          
                    ps.setString(2, Obra);              
                    ps.setString(3, Denominacion);      
                    ps.setString(4, Mant_Correctiv);    
                    ps.setString(5, Mot_Mant_Corre);    
                    ps.setString(6, FechaInicioEs);     
                    ps.setString(7, FechaFinEs);                            
                    ps.setString(8,Clasificacion);      
                    ps.setString(9,Niveldetension);                    
                    ps.setString(10,CablesBT);          
                    ps.setString(11,LineasBT);          
                    ps.setString(12,CablesMT);          
                    ps.setString(13,LineasMT);          
                    ps.setString(14,PotTrafoMTBT);      
                    ps.setString(15,TransAereoMTBT);    
                    ps.setString(16,TransNoAerMTBT);    
                    ps.setString(17,TelemandoMT);       
                    ps.setString(18,ExpansionCT);       
                    ps.setString(19,RedAT);             
                    ps.setString(20,PotenciaAT);        
                    ps.setString(21,OtrosEquipam);      
                    ps.setString(22,Partido);           
                    ps.setString(23,Localidad);         
                    ps.setString(24,Barrio);            
                    ps.setString(25,ZonaTecnica);       
                    ps.setString(26,Calle);             
                    ps.setString(27,Numero);            
                    ps.setString(28,FechaPtaServ);      
                    ps.setString(29,FechaBaja);         
                    ps.setString(30,MontoTotal);        
                    ps.setString(31,Coordenadax);       
                    ps.setString(32,Coordenaday);       
                    ps.setString(33,ExInterrupcion);    
                    ps.setString(34,Observaciones);
                    
                   
                    rt = ps.executeUpdate();
                    ps.close();
                }
                contador = contador + 1;
                count += rt;
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
        name = f_insert +"_" + name; 
        //System.out.println(name);
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
    

}
