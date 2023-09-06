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
import jcifs.smb.NtlmPasswordAuthentication;
import jcifs.smb.SmbException;
import jcifs.smb.SmbFile;
import jcifs.smb.SmbFileInputStream;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

public class EquiposDeTrabajoEnre {
    
    /************************************************
     * Adecuar: Ambiente, NetworkDomain,  
     * NetworkFolder, base 
     ************************************************/
    
    //Ambiente de la base de datos. Valores: DEV, QA, PRO
    private static final String Ambiente = "PRO";
    //Conexion a la carpeta de red
    private static final String NetworkDomain = "PRO";
    private static final String NetworkFolder = "smb://192.168.190.150/DatosCEN/Iox/DIRDYC/Requerimientos_ENRE/";
    //user y pass se pueblan desde la base
    private static String NetworkUser = ""; 
    private static String NetworkPass = "";
    
    
  //  private static final String AProcesarDir = "C:\\Users\\mrrodriguez\\Documents\\Temas\\Evolutivo_Web_Services\\58022_WSENRE_EQuipos_de_trabajo\\Excels_de_entrada\\esGrupTrabEnre\\a_procesar";
  //  private static final String ProcesadosDir = "C:\\Users\\mrrodriguez\\Documents\\Temas\\Evolutivo_Web_Services\\58022_WSENRE_EQuipos_de_trabajo\\Excels_de_entrada\\esGrupTrabEnre\\procesados";
  //  private static final String ErrorDir = "C:\\Users\\mrrodriguez\\Documents\\Temas\\Evolutivo_Web_Services\\58022_WSENRE_EQuipos_de_trabajo\\Excels_de_entrada\\esGrupTrabEnre\\error";
  //  private static final String LogsDir = "C:\\Users\\mrrodriguez\\Documents\\Temas\\Evolutivo_Web_Services\\58022_WSENRE_EQuipos_de_trabajo\\Excels_de_entrada\\esGrupTrabEnre\\logs";
    
    //Se utiliza para simular carpeta de la red
    //private static final String LocalDebugFolder = "C:\\Users\\mrrodriguez\\Documents\\Temas\\Evolutivo_Web_Services\\58022_WSENRE_EQuipos_de_trabajo\\Excels_de_entrada\\esGrupTrabEnre\\repositorioLocal";
    private static final String LocalDebugFolder = "";
    
    //DIRECTORIOS RELATIVOS PARA PRODUCCION
    private static final String base= "/ias/EquipTrab_ENRE/";    
    private static final String AProcesarDir = base  + "estructura/a_procesar";
    private static final String ProcesadosDir = base + "estructura/procesados";
    private static final String ErrorDir = base  + "estructura/error";
    private static final String LogsDir = base  + "estructura/logs";
    
    
    //Logger
    private static final Logger Log = Logger.getLogger(EquiposDeTrabajoEnre.class.getName());
    private static final String LogFileName = "batchLog.txt";
    //Patron de archivo
    private static final String Pattern_1 = "Estado_de_Equipos_de_Trabajo_??-??-????.{xls,xlsx}";
    private static final String RegExPattern = "^Estado_de_Equipos_de_Trabajo_\\d{2}\\-\\d{2}\\-\\d{4}.xlsx?";
    //Driver y string de conexion a diferentes ambientes
    private static final String DriverClass = "oracle.jdbc.driver.OracleDriver";
    private static final String ConnDev = "jdbc:oracle:thin:NEXUS_ENRE/NEXUS_ENRE@tdbs6.tro.edenor:1521:GISDEV01";
    private static final String ConnQA = "jdbc:oracle:thin:NEXUS_ENRE/NEXUS_ENRE@TDBS5.PRO.EDENOR:1529:GISQA01";
    private static final String ConnPro = "jdbc:oracle:thin:NEXUS_ENRE/cami0net4@TCLH1.TRO.EDENOR:1528:GISPR01";
    
    //Nombre del libro a leer
    private static final String nombreDeLibro = "Informe ENRE";
    
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
    //Indices parsear el nombre de archivo Estado_de_Equipos_de_Trabajo_??-??-????
    private static final Integer FeDiaIni = 29;
    private static final Integer FeDiaFin = 31;
    private static final Integer FeMesIni = 32;
    private static final Integer FeMesFin = 34;
    private static final Integer FeAnioIni = 35;
    private static final Integer FeAnioFin = 39;
    //Definido para el metodo depurar. Indica los archivos a borrar segun fecha.
    private static final Integer DiasHaciaAtras = 7;
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
    
    
////////////////////////////////////////////////////////////////////////////////
/////////////////************       M A I N      ********************///////////
////////////////////////////////////////////////////////////////////////////////
    @SuppressWarnings("ThrowableResultIgnored")
    public static void main(String[] args) throws IOException {

        try {
            
            inicializarLog();
            
            Log.log(Level.INFO, IniciandoMsj);
            
            depurarDirectorios();
            
            setearConexion();
            
            obtenerCredenciales();
            
            traerArchivo();
            //traerArchivoDebug();
            
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
        
        switch (fueProcesadoAnteriormente(masReciente.getName())) {
                case "OK":
                    throw new Exception(NoSeVuelveAProcesarMsj);
                case "ERROR":
                    throw new Exception(NoSeVuelveAProcesarMsjErr);
        }
        
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
                switch (cell.getCachedFormulaResultType()) {
                    case Cell.CELL_TYPE_STRING:
                        return cell.getRichStringCellValue().getString();
                    case Cell.CELL_TYPE_NUMERIC:
                        return Double.toString(Math.round(cell.getNumericCellValue()));
                    default:
                        break;
                }
            case CELL_TYPE_NUMERIC:
                cellAux = formatNumeric(cell, formato);
                break;
            case CELL_TYPE_STRING:
                cellAux = cell.getStringCellValue();
                break;
        }
        
        return cellAux;
    }
    
    private static boolean esFilaAOmitir(Integer index, String ext) throws Exception{
        
        if(!ext.toUpperCase().equals("XLSX") && !ext.toUpperCase().equals("XLS")){
            throw new Exception("[ERROR:esFilaAOmitir]: Extensión de archivo no es la permitida (XLS, XLSX)");
        }
        
        Integer[] filasAOmitir = {0, 1, 2, 3, 5, 6, 15, 17, 18, 27, 29, 30};
        
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
            throw new Exception("[ERROR:esFilaAOmitir]: Extensión de archivo no es la permitida (XLS, XLSX)");
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
        //nombreArchivo = masReciente.getFileName().toString();
        nombreArchivo = "Estado_de_Equipos_de_Trabajo_05-02-2016.xlsx";
        
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
        
        Cell turno = null;
        Cell situacion ;
        Cell situacionAnterior = null;
        Cell tipo ;
        Cell tipoAnterior = null;
        
        try {
            for(Integer j=0; j < UltimaFila; j++){

                if (esFilaAOmitir(j, extension))
                    continue;

                if (esFilaTurno(j, extension)){
                    turno = sheet.getRow(j).getCell(1);
                    continue;
                }
                
                //Tratamiento para filas combinadas
                situacion = sheet.getRow(j).getCell(1);
                if (situacion == null || situacion.toString().isEmpty()){
                    situacion = situacionAnterior;
                } else {
                    situacionAnterior = situacion;
                }
                
                tipo = sheet.getRow(j).getCell(2);
                if (tipo == null || tipo.toString().isEmpty()){
                    tipo = tipoAnterior;
                } else {
                    tipoAnterior = tipo;
                }

                //Obtengo datos del resto de las filas
                Cell conformacion = sheet.getRow(j).getCell(3);
                
                Cell cantModPropio = sheet.getRow(j).getCell(4);
                
                Cell cantTecPropio = sheet.getRow(j).getCell(5);
                
                Cell totalPersPropio = sheet.getRow(j).getCell(6);
                
                Cell cantModContrat = sheet.getRow(j).getCell(7);
                
                Cell cantTecContrat = sheet.getRow(j).getCell(8);
                
                Cell totalContrat = sheet.getRow(j).getCell(9);
                
                Cell totalRecursos = sheet.getRow(j).getCell(10);
                
                Cell eqLocFallas = sheet.getRow(j).getCell(11);
                
                Cell gruposElec = sheet.getRow(j).getCell(12);

                HashMap<String, String> excelRow = new HashMap<>();
                excelRow.put("TURNO", excelTypeToStr(turno, "String"));
                excelRow.put("SITUACION", excelTypeToStr(situacion, "String"));
                excelRow.put("TIPO", excelTypeToStr(tipo, "String"));
                excelRow.put("CONFORMACION", excelTypeToStr(conformacion, "NumEntero"));
                excelRow.put("CANT_MOD_PROPIO", excelTypeToStr(cantModPropio, "NumEntero"));
                excelRow.put("CANT_TECNICOS_PROPIO", excelTypeToStr(cantTecPropio, "NumEntero"));
                excelRow.put("TOTAL_PERS_PROPIO", excelTypeToStr(totalPersPropio, "NumEntero"));
                excelRow.put("CANT_MOD_CONTRAT", excelTypeToStr(cantModContrat, "NumEntero"));
                excelRow.put("CANT_TECNICOS_CONTRAT", excelTypeToStr(cantTecContrat, "NumEntero"));
                excelRow.put("TOTAL_CONTRATISTA", excelTypeToStr(totalContrat, "NumEntero"));
                excelRow.put("TOTAL_RECURSOS", excelTypeToStr(totalRecursos, "NumEntero"));
                excelRow.put("EQUIPOS_LOC_FALLAS", excelTypeToStr(eqLocFallas, "NumEntero"));
                excelRow.put("GRUPOS_ELECTROGENOS", excelTypeToStr(gruposElec, "NumEntero"));

                excelRow.put("ULT_MODIF_ARCHIVO", feMod);

                RowList.add(excelRow);
            }
        } catch (NullPointerException e){
            throw new NullPointerException ("read: Error al parsear el excel - " + e);
        } catch (Exception e){
            throw new Exception ("read: Error al parsear el excel - " + e);
        }

    }

    private static void insertar() throws SQLException {
        Integer rt ;
        Integer count = 0;
        for (HashMap<String, String> row : RowList) {
            String turno               = (String)row.get("TURNO");
            String situacion           = (String)row.get("SITUACION");
            String tipo                = (String)row.get("TIPO");
            String conformacion        = (String)row.get("CONFORMACION");
            String cantModPropio       = (String)row.get("CANT_MOD_PROPIO");
            String cantTecnicosPropio  = (String)row.get("CANT_TECNICOS_PROPIO");
            String totalPersPropio     = (String)row.get("TOTAL_PERS_PROPIO");
            String cantModContrat      = (String)row.get("CANT_MOD_CONTRAT");
            String cantTecnicosContrat = (String)row.get("CANT_TECNICOS_CONTRAT");
            String totalContratista    = (String)row.get("TOTAL_CONTRATISTA");
            String totalRecursos       = (String)row.get("TOTAL_RECURSOS");
            String equiposLocFallas    = (String)row.get("EQUIPOS_LOC_FALLAS");
            String gruposElectrogenos  = (String)row.get("GRUPOS_ELECTROGENOS");
            
            String ult_modif_archivo = (String)row.get("ULT_MODIF_ARCHIVO");
            
            try {
                String sql =
                        "INSERT INTO NEXUS_GIS.WSENRE_EQUIPOTRABAJO VALUES (  "
                        + "?, "    //turno
                        + "?, "    //situacion
                        + "?, "    //tipo
                        + "?, "    //conformacion
                        + "?, "    //cantModPropio
                        + "?, "    //cantTecnicosPropio
                        + "?, "    //totalPersPropio
                        + "?, "    //cantModContrat
                        + "?, "    //cantTecnicosContrat
                        + "?, "    //totalContratista
                        + "?, "    //totalRecursos
                        + "?, "    //equiposLocFallas
                        + "?, "    //gruposElectrogenos
                        + "SYSDATE, "
                        + "NULL,"
                        + "TO_DATE(?, 'YYYY-MM-DD HH24:MI:SS'))";

                try (PreparedStatement ps = connection.prepareStatement(sql)) {
                    ps.setString(1, turno);
                    ps.setString(2, situacion);
                    ps.setString(3, tipo);
                    ps.setString(4, conformacion);
                    ps.setString(5, cantModPropio);
                    ps.setString(6, cantTecnicosPropio);
                    ps.setString(7, totalPersPropio);
                    ps.setString(8, cantModContrat);
                    ps.setString(9, cantTecnicosContrat);
                    ps.setString(10, totalContratista);
                    ps.setString(11, totalRecursos);
                    ps.setString(12, equiposLocFallas);
                    ps.setString(13, gruposElectrogenos);
                    
                    ps.setString(14, ult_modif_archivo);
                    
                    rt = ps.executeUpdate();
                }
                
                count += rt;
            } catch (SQLException ex) {
                throw ex;
            }
        }
        Log.log(Level.INFO, SeInsertaronFilasMsj, count);
    }

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
    

}
