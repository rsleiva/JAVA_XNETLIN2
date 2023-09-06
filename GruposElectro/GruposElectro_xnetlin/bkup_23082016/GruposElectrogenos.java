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
import java.util.Date;
import java.util.*;
import java.util.logging.Formatter;
import java.util.logging.*;
import jcifs.smb.NtlmPasswordAuthentication;
import jcifs.smb.SmbException;
import jcifs.smb.SmbFile;
import jcifs.smb.SmbFileInputStream;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.*;

public class GruposElectrogenos {
    
    /************************************************
     * Adecuar: Ambiente, NetworkDomain,
     * NetworkFolder, base 
     ************************************************/
    
    //Ambiente de la base de datos. Valores: DEV, QA, PRO
    private static final String Ambiente = "PRO";
    //Conexion a la carpeta de red
    private static final String NetworkDomain = "PRO";
    private static final String NetworkFolder = "smb://SRVCENFSPS001.PRO.EDENOR/DatosCEN/Iox/DIRDYC/Requerimientos_ENRE/";
    //user y pass se pueblan desde la base
    private static String NetworkUser = ""; 
    private static String NetworkPass = "";
    
    //DIRECTORIOS RELATIVOS PARA PRODUCCION
    private static final String base= "/ias/GruposElectro/";    
//	private static final String base= "/home/omigliarini/GruposElectro/";    
    private static final String AProcesarDir = base  + "estructura/a_procesar";
    private static final String ProcesadosDir = base + "estructura/procesados";
    private static final String ErrorDir = base  + "estructura/error";
    private static final String LogsDir = base  + "estructura/logs";
    
    //
    private static final String RegExPattern = "^Grupos_Electr�genos_\\d{2}\\-\\d{2}\\-\\d{4}.xlsx?";
    //Logger
    private static final Logger Log = Logger.getLogger(GruposElectrogenos.class.getName());
    private static final String LogFileName = "batchLog.txt";
    //Patron de archivo
    private static final String Pattern_1 = "Grupos_Electr�genos_??-??-????.{xls,xlsx}";
    //Driver y string de conexion a diferentes ambientes
    private static final String DriverClass = "oracle.jdbc.driver.OracleDriver";
    private static final String ConnDev = "jdbc:oracle:thin:NEXUS_ENRE/NEXUS_ENRE@tdbs6.tro.edenor:1521:GISDEV01";
    private static final String ConnQA = "jdbc:oracle:thin:NEXUS_ENRE/NEXUS_ENRE@TDBS5.PRO.EDENOR:1529:GISQA01";
    private static final String ConnPro = "jdbc:oracle:thin:NEXUS_ENRE/cami0net4@TCLH1.TRO.EDENOR:1528:GISPR01";
    //Conn
    private static Connection Connection = null;
    //
    private static String nombreArchivo = "";
    //Literales
    private static final String NoSeVuelveAProcesarMsj = "El archivo ya fue procesado exitosamente. No se volvera a procesar";
    private static final String NoSeVuelveAProcesarMsjErr = "El archivo ya fue procesado con ERROR. No se volvera a procesar";
    private static final String NoSeEncontraronArchivos = "No se encontraron archivos en el directorio";
    private static final String FinalizadoConErrorMsj = "------------- Proceso finalizado con ERROR ----------";
    private static final String FinalizadoConWarningsMsj = "--------- Proceso finalizado con WARNINGS ---------";
    private static final String DepurarAProcMsj = "Depuracion: se borro el archivo {0} del directorio a_procesar";
    private static final String DepurarProcMsj = "Depuracion: se borro el archivo {0} del directorio procesados";
    private static final String DepurarErrorMsj = "Depuracion: se borro el archivo {0} del directorio error";
    private static final String DepurarMsj = "Se borraron {0} archivos en la depuracion.";
    private static final String MoverErrorMsj = "Se mueve el archivo {0} a la carpeta 'error'";
    private static final String ObtenerCredencialesErrMsj = "Error al obtener usuario y contraseña";
    private static final String FueCopiadoOk = "el archivo {0} fue copiado desde el repositorio remoto";
    private static final String IniciandoMsj = "--------- COMENZANDO PROCESO BATCH ---------";
    private static final String FinalizandoMsj = "--------- Proceso finalizado correctamente ---------";
    //Buffers
    private static final List<HashMap<String, String>> rowList = new ArrayList<>();
    //Indices parsear el nombre de archivo Grupos_Electrógenos_??-??-????
    private static final Integer FeDiaIni = 20;
    private static final Integer FeDiaFin = 22;
    private static final Integer FeMesIni = 23;
    private static final Integer FeMesFin = 25;
    private static final Integer FeAnioIni = 26;
    private static final Integer FeAnioFin = 30;
    //Definido para el metodo depurar. Indica los archivos a borrar segun fecha.
    private static final Integer DiasHaciaAtras = 7;
    //Ultima fila que va a leer
    private static final Integer CantFilasCabecera = 1;
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

            ps = Connection.prepareStatement(sql);
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
    
    public static void traerArchivo() throws MalformedURLException, 
            SmbException, UnknownHostException, IOException, Exception{
        
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
        
        SmbFileInputStream in = new SmbFileInputStream(masReciente);
        Path dir = Paths.get(AProcesarDir);
        Path fullPath = dir.resolve(masReciente.getName());
        FileOutputStream os = new FileOutputStream(fullPath.toString());
        
        byte[] b = new byte[8192];
        int n;
        while ((n = in.read(b)) > 0) {
            os.write(b, 0, n);
        }
        
        Log.log(Level.INFO, FueCopiadoOk,masReciente.getName());
    }
    

    public static SmbFile traerUltimo(SmbFile[] files) throws Exception {

        SmbFile choice = null;
        try {
            choice = files[0];
            if (choice == null){
                throw new Exception (NoSeEncontraronArchivos);
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
    
    //Toma una celda de tipo numerico y la formatea
    private static String formatNumeric(Cell cell, String formato) throws Exception{
        
        if (cell.getCellType() != 0){ //CELL_TYPE_NUMERIC
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
            case CELL_TYPE_BLANK: //CELL_TYPE_BLANK
                break;
            case CELL_TYPE_BOOLEAN: //CELL_TYPE_BOOLEAN
                break;
            case CELL_TYPE_ERROR: //CELL_TYPE_ERROR
                break;
            case CELL_TYPE_FORMULA: //CELL_TYPE_FORMULA
                throw new Exception ("No se permiten fórmulas en el documento");
            case CELL_TYPE_NUMERIC: //CELL_TYPE_NUMERIC
                cellAux = formatNumeric(cell, formato);
                break;
            case CELL_TYPE_STRING: //CELL_TYPE_STRING
                cellAux = cell.getStringCellValue();
                break;
        }
        
        return cellAux;
    }
    
    private static String excelTypeToStr(Cell cell) throws Exception {
        String cellAux = "";
        
        switch (cell.getCellType()) {
            case CELL_TYPE_BLANK: //CELL_TYPE_BLANK
                cellAux = "";
                break;
            case CELL_TYPE_BOOLEAN: //CELL_TYPE_BOOLEAN
                cellAux = "";
                break;
            case CELL_TYPE_ERROR: //CELL_TYPE_ERROR
                cellAux = "";
                break;
            case CELL_TYPE_FORMULA: //CELL_TYPE_FORMULA
                throw new Exception ("No se permiten fórmulas en el documento");
            case CELL_TYPE_NUMERIC: //CELL_TYPE_NUMERIC
                Double numDec = cell.getNumericCellValue();
                cellAux = Double.toString(numDec.intValue());
                break;
            case CELL_TYPE_STRING: //CELL_TYPE_STRING
                cellAux = cell.getStringCellValue();
                break;
        }
        
        return cellAux;

    }
    
    //parsea el archivo excel
    private static void parsearDocumento()
            throws FileNotFoundException, IOException, InvalidFormatException,
            Exception {
        
        Path masReciente = obtenerUltimoModificado(AProcesarDir);
        nombreArchivo = masReciente.getFileName().toString();
        Path baseP = Paths.get(AProcesarDir);
        Path archAProcesar = baseP.resolve(nombreArchivo);
        
        Workbook workbook;
        String archAProcStr = archAProcesar.toString();
        FileInputStream file = new FileInputStream(new File(archAProcStr));

        workbook = WorkbookFactory.create(file);
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

            excelRow.put("MES_ANIO", excelTypeToStr(cell, "MesAnio"));

            cell = cellIterator.next();
            excelRow.put("AREA_OPERA", excelTypeToStr(cell, "String"));

            cell = cellIterator.next();
            excelRow.put("ALIMENTADOR", excelTypeToStr(cell, "NumEntero"));

            cell = cellIterator.next();
            excelRow.put("CT_CONEC", excelTypeToStr(cell, "NumEntero"));
            
            cell = cellIterator.next();
            excelRow.put("DIRECCION", excelTypeToStr(cell, "String"));
            
            cell = cellIterator.next();
            excelRow.put("LOCALIDAD", excelTypeToStr(cell, "String"));

            cell = cellIterator.next();
            excelRow.put("PARTIDO", excelTypeToStr(cell, "String"));
            
            cell = cellIterator.next();
            excelRow.put("PROPIETARIO", excelTypeToStr(cell, "String"));

            cell = cellIterator.next();
            excelRow.put("EMPRESA", excelTypeToStr(cell, "String"));

            cell = cellIterator.next();
            excelRow.put("IDENTI", excelTypeToStr(cell, "String"));

            cell = cellIterator.next();
            excelRow.put("MARCA", excelTypeToStr(cell, "String"));

            cell = cellIterator.next();
            excelRow.put("MODELO", excelTypeToStr(cell, "String"));

            cell = cellIterator.next();
            excelRow.put("POTENCIA", excelTypeToStr(cell, "String"));

            cell = cellIterator.next();
            excelRow.put("CONSUMO_COMBUS", excelTypeToStr(cell, "String"));

            cell = cellIterator.next();
            String emisiones = excelTypeToStr(cell, "String");
            cell = cellIterator.next();
            emisiones += " " + excelTypeToStr(cell, "String");
            cell = cellIterator.next();
            emisiones += " " + excelTypeToStr(cell, "String");
            cell = cellIterator.next();
            emisiones += " " + excelTypeToStr(cell, "String");
            excelRow.put("EMISIONES_GAS", emisiones);

            cell = cellIterator.next();
            String nivelesSonoros = excelTypeToStr(cell, "String");
            cell = cellIterator.next();
            nivelesSonoros += " " + excelTypeToStr(cell, "String");
            cell = cellIterator.next();
            nivelesSonoros += " " + excelTypeToStr(cell, "String");
            excelRow.put("NIVELES_SONOROS", nivelesSonoros);

            cell = cellIterator.next();
            excelRow.put("FRANJA_HR", excelTypeToStr(cell, "String"));

            cell = cellIterator.next();
            excelRow.put("HR_MENS_FUNC", excelTypeToStr(cell, "NumEntero"));

            cell = cellIterator.next();
            excelRow.put("FECHA_COL", excelTypeToStr(cell, "DiaMesAnio"));

            cell = cellIterator.next();
            excelRow.put("FECHA_RETIRO", excelTypeToStr(cell, "DiaMesAnio"));

            cell = cellIterator.next();
            excelRow.put("ESTADO", excelTypeToStr(cell, "String"));

            cell = cellIterator.next();
            excelRow.put("INSPEC_INI", excelTypeToStr(cell, "DiaMesAnio"));

            cell = cellIterator.next();
            excelRow.put("FECHA_REVIS_SEM_1", excelTypeToStr(cell, "String"));

            cell = cellIterator.next();
            excelRow.put("FECHA_REVIS_SEM_2", excelTypeToStr(cell, "String"));

            cell = cellIterator.next();
            excelRow.put("FECHA_REVIS_SEM_3", excelTypeToStr(cell, "String"));

            cell = cellIterator.next();
            excelRow.put("FECHA_REVIS_SEM_4", excelTypeToStr(cell, "String"));

            cell = cellIterator.next();
            excelRow.put("FECHA_REVIS_SEM_5", excelTypeToStr(cell, "String"));

            cell = cellIterator.next();
            excelRow.put("CANT_RECL_RECIBIDOS", excelTypeToStr(cell, "String"));

            cell = cellIterator.next();
            excelRow.put("X", excelTypeToStr(cell, "NumDecimal"));

            cell = cellIterator.next();
            excelRow.put("Y", excelTypeToStr(cell, "NumDecimal"));

            cell = cellIterator.next();
            excelRow.put("OBSERVACIONES", excelTypeToStr(cell, "String"));

            excelRow.put("ULT_MODIF_ARCHIVO", feMod);

            rowList.add(excelRow);
        }

    }

    //realiza la insercion en la base de lo parseado en el excel
    private static void insertar() throws SQLException {
        Integer rt = 0;
        Integer count = 0;
        for (HashMap<String, String> row : rowList) {
            String mes_anio = (String) row.get("MES_ANIO");
            String area_opera = (String) row.get("AREA_OPERA");
            String alimentador = (String) row.get("ALIMENTADOR");
            String ct_conec = (String) row.get("CT_CONEC");
            String direccion = (String) row.get("DIRECCION");
            String localidad = (String) row.get("LOCALIDAD");
            String partido = (String) row.get("PARTIDO");
            String propietario = (String) row.get("PROPIETARIO");
            String empresa = (String) row.get("EMPRESA");
            String identi = (String) row.get("IDENTI");
            String marca = (String) row.get("MARCA");
            String modelo = (String) row.get("MODELO");
            String potencia = (String) row.get("POTENCIA");
            String consumo_combus = (String) row.get("CONSUMO_COMBUS");
            String emisiones_gas = (String) row.get("EMISIONES_GAS");
            String niveles_sonoros = (String) row.get("NIVELES_SONOROS");
            String franja_hr = (String) row.get("FRANJA_HR");
            String hr_mens_func = (String) row.get("HR_MENS_FUNC");
            String fecha_col = (String) row.get("FECHA_COL");
            String fecha_retiro = (String) row.get("FECHA_RETIRO");
            String estado = (String) row.get("ESTADO");
            String inspec_ini = (String) row.get("INSPEC_INI");
            String fecha_revis_sem_1 = (String) row.get("FECHA_REVIS_SEM_1");
            String fecha_revis_sem_2 = (String) row.get("FECHA_REVIS_SEM_2");
            String fecha_revis_sem_3 = (String) row.get("FECHA_REVIS_SEM_3");
            String fecha_revis_sem_4 = (String) row.get("FECHA_REVIS_SEM_4");
            String fecha_revis_sem_5 = (String) row.get("FECHA_REVIS_SEM_5");
            String cant_recl_recibidos = (String) row.get("CANT_RECL_RECIBIDOS");
            String x = (String) row.get("X");
            String y = (String) row.get("Y");
            String observaciones = (String) row.get("OBSERVACIONES");
            String ult_modif_archivo = (String) row.get("ULT_MODIF_ARCHIVO");

            
            try {
                String sql = "INSERT INTO NEXUS_GIS.WSENRE_GR_ELECTROGENOS VALUES "
                           + "(?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, "
                           + " ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, "
                           + " SYSDATE, TO_DATE(?, 'YYYY-MM-DD HH24:MI:SS'), ?)";

                PreparedStatement ps = Connection.prepareStatement(sql);
                ps.setString(1, mes_anio); 
                ps.setString(2, area_opera); 
                ps.setString(3, alimentador); 
                ps.setString(4, ct_conec); 
                ps.setString(5, direccion); 
                ps.setString(6, localidad); 
                ps.setString(7, partido);
                ps.setString(8, empresa); 
                ps.setString(9, identi); 
                ps.setString(10, marca); 
                ps.setString(11, modelo); 
                ps.setString(12, potencia); 
                ps.setString(13, consumo_combus); 
                ps.setString(14, emisiones_gas); 
                ps.setString(15, niveles_sonoros); 
                ps.setString(16, franja_hr); 
                ps.setString(17, hr_mens_func); 
                ps.setString(18, fecha_col); 
                ps.setString(19, fecha_retiro); 
                ps.setString(20, estado); 
                ps.setString(21, inspec_ini); 
                ps.setString(22, fecha_revis_sem_1); 
                ps.setString(23, fecha_revis_sem_2); 
                ps.setString(24, fecha_revis_sem_3); 
                ps.setString(25, fecha_revis_sem_4); 
                ps.setString(26, fecha_revis_sem_5); 
                ps.setString(27, cant_recl_recibidos); 
                ps.setString(28, x); 
                ps.setString(29, y); 
                ps.setString(30, observaciones); 
                ps.setString(31, ult_modif_archivo);
                ps.setString(32, propietario);

                rt = ps.executeUpdate();
                count += rt;
            } catch (SQLException ex) {
                throw ex;
            }
        }
        Log.log(Level.INFO, "Se insertaron {0} filas.", count);
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
    

}
