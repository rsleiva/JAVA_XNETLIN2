
import java.awt.Color;
import java.io.BufferedReader;
import java.io.ByteArrayInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.io.Reader;
import java.sql.Clob;
import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.sql.Statement;
import java.util.ArrayList;
import java.util.Date;
import java.util.Enumeration;
import java.util.HashMap;
import java.util.List;
import java.util.Properties;
import java.util.PropertyResourceBundle;
import java.util.logging.Level;
import java.util.logging.Logger;
import javax.activation.DataHandler;
import javax.activation.DataSource;
import javax.activation.FileDataSource;
import javax.mail.Address;
import javax.mail.BodyPart;
import javax.mail.Message;
import javax.mail.SendFailedException;
import javax.mail.Session;
import javax.mail.Transport;
import javax.mail.internet.InternetAddress;
import javax.mail.internet.MimeBodyPart;
import javax.mail.internet.MimeMessage;
import javax.mail.internet.MimeMultipart;
import oracle.jdbc.OracleCallableStatement;
import oracle.jdbc.OracleTypes;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xssf.usermodel.extensions.XSSFCellBorder;

/**
 *
 * @author mrrodriguez
 */
public class Rep_Electrodependientes {
    private static String workDir= System.getProperty("user.dir");//PROD
    private static String ARCHIVO_PROP= workDir + "/config.properties";  //PROD
    //private static String ARCHIVO_PROP= "/config.properties";  
    private static Connection con = null;
    private static String fecha = null;
     private static String destinatarios;
    private static String mailto; //TEST
// private static String mailto ="gmeyer@edenor.com,PMAZZA@edenor.com,centrodeinformacion@edenor.com, despacho_bt@edenor.com,ITSM_Desarrollos_propios@edenor.com";
   //private static String xlsfile ="H:/NetBeansProjects/Rep_Electrodependientes/Rep_Electrodependientes/rep_electrodep.xlsx";  
   private static String xlsfile ="/ias/rep_electrodep/rep_electrodep.xlsx"; //Produccion
    
    //private ResultSet rs;
    private static List<HashMap<String, String>> tablaForz = null;
    private static HashMap<String, String> ultimaForz = null;
    private static List<HashMap<String, String>> tablaProgra = null;
    private static HashMap<String, String> ultimaProgra = null;
    private static List<HashMap<String, String>> tablaPuntuales = null;
    private static HashMap<String, String> ultimaPuntuales = null;
    static ResultSet rs1;
    private static String Mensaje = null;
    
//------------------------------------------------//
//----------------------MAIN----------------------//
//------------------------------------------------//
    
    public static void main(String[] args) {

        try {
            
            Lee_propiedades();
            OracleConnect();
           
            rs1 = Electrodependientes(con); 
            if(rs1.isBeforeFirst())
{ 
            creaExcel();
            
            
} else   
{
  Mensaje = "Este es un mail que reporta los Clientes Electrodependientes .";  
}
            
            
            enviaMail();
        } catch (Exception ex) {
            System.out.println(ex);
        }

    }
//------------------------------------------------//
//------------------------------------------------//
    
    
     static public void Lee_propiedades() throws Exception {
         
                        
                        String V_key, V_value;
                        
                         Properties prop = new Properties();
                          InputStream input = null;
                         InputStream is = 
        Rep_Electrodependientes.class.getClassLoader().getResourceAsStream("config.properties");
         	//	InputStream is = ClassLoader.getSystemResourceAsStream(ARCHIVO_PROP);
                      //  InputStream is = ClassLoader.getSystemResourceAsStream("config.properties");
                        
			PropertyResourceBundle parametros= new PropertyResourceBundle(is);
         
     			for (Enumeration <String> lista = parametros.getKeys(); lista.hasMoreElements();) {
				V_key   =  lista.nextElement();
				V_value = ((String)parametros.handleGetObject(V_key)).trim();
				if (V_key.equals("DESTINATARIOS"))
					destinatarios=V_value;
                                        mailto=destinatarios;
		}	
         
     }
    
    
    public static void OracleConnect() throws Exception {
        final String driverClass = "oracle.jdbc.driver.OracleDriver";
        try {
            Class.forName(driverClass).newInstance();
        } catch (ClassNotFoundException | InstantiationException | IllegalAccessException e) {
            System.out.println("Error al cargar el driver: " + driverClass + " -Error: " + e);
            throw e;
        }

        try {
            con = DriverManager
                    .getConnection("jdbc:oracle:thin:@ltronxgisbdpr03.pro.edenor:1528:gispr03", "SVC_ORA_GIS", "jv506uzy");
        } catch (SQLException e) {
            switch (e.getErrorCode()) {
                case 1017:    //USUARIO O CLAVE IVALIDO; LOGON DENIED
                    System.out.println("Usuario de la base de datos incorrecto");
                    break;
                case 28000:    //CUENTA BLOQUEADA
                    System.out.println("Cuenta bloqueada");
                    break;
                case 2391:		//EXCEDIDO EN CONEXIONES
                    System.out.println("LÃ­mite de conexiones excedido, intentar mÃ¡s tarde");
                    break;
                case 28001:		//CLAVE EXPIRADA
                    System.out.println("Clave de la base de datos expirada");
                    break;
                default:
                    System.out.println("Error al conectarse: " + e);
            }
            throw e;
        }
    }
    
    static ResultSet Electrodependientes (Connection conexion) throws Exception {        
        Connection con = conexion;
            //Connection con = DbConnection.getConnection(request);
        ResultSet rs = null;
        Statement sentencia;
        
        try {
             String query_variable = 

       "SELECT SPRCLIENTS.FSCLIENTID AS CUENTA, SPRCLIENTS.FULLNAME AS RAZON_SOCIAL, SPRCLIENTS.TELEPHONENUMBER AS TELEFONO, "
      + "(SELECT SPRLOG.EVENTDATE FROM NEXUS_GIS.SPRLOG WHERE SPRLOG.LOGID = SPRCLIENTS.LOGIDFROM ) AS F_ALTA, "
      + "( SELECT SMSTREETS.STREETNAME FROM NEXUS_GIS.SMSTREETS WHERE SMSTREETS.STREETANTIQ = 0 AND SMSTREETS.STREETDELETED = 0 AND SMSTREETS.STREETID = SPRCLIENTS.STREETID ) AS CALLE,"
      + " SPRCLIENTS.STREETNUMBER AS NRO,SPRCLIENTS.STREETOTHER PISO_DPTO, "
      + "( SELECT SMSTREETS.STREETNAME FROM NEXUS_GIS.SMSTREETS WHERE SMSTREETS.STREETANTIQ = 0 AND SMSTREETS.STREETDELETED = 0 AND SMSTREETS.STREETID = SPRCLIENTS.STREETID1 ) AS ENTE_CALLE_1,"
      + " ( SELECT SMSTREETS.STREETNAME FROM NEXUS_GIS.SMSTREETS WHERE SMSTREETS.STREETANTIQ = 0 AND SMSTREETS.STREETDELETED = 0 AND SMSTREETS.STREETID = SPRCLIENTS.STREETID2 ) AS ENTE_CALLE_2, "
      + "( SELECT AMAREAS.AREANAME FROM NEXUS_GIS.AMAREAS  WHERE AMAREAS.AREAID = SPRCLIENTS.LEVELONEAREAID ) LOCALIDAD, "
      + "( SELECT AMAREAS.AREANAME FROM NEXUS_GIS.AMAREAS  WHERE AMAREAS.AREAID = SPRCLIENTS.LEVELTWOAREAID ) PARTIDO, ( SELECT LLAM_SECTORES.REGION||' - '||LLAM_SECTORES.SECTOR FROM NEXUS_GIS.LLAM_SECTORES WHERE LLAM_SECTORES.LOCA_ID = SPRCLIENTS.LEVELONEAREAID ) AS REGION, "
      + "( SELECT SPROBJECTS.X FROM NEXUS_GIS.SPROBJECTS WHERE SPROBJECTS.LOGIDTO = 0 AND SPROBJECTS.SPRID = 190 AND SPROBJECTS.OBJECTID = (SELECT max(SPRLINKS.OBJECTID) FROM NEXUS_GIS.SPRLINKS WHERE SPRLINKS.LINKID = 407 AND SPRLINKS.LOGIDTO = 0 AND SPRLINKS.LINKVALUE = RPAD(SPRCLIENTS.FSCLIENTID,30))) AS X, ( SELECT SPROBJECTS.Y FROM NEXUS_GIS.SPROBJECTS WHERE SPROBJECTS.LOGIDTO = 0 AND SPROBJECTS.SPRID = 190 AND SPROBJECTS.OBJECTID = (SELECT max(SPRLINKS.OBJECTID) FROM NEXUS_GIS.SPRLINKS WHERE SPRLINKS.LINKID = 407 AND SPRLINKS.LOGIDTO = 0 AND SPRLINKS.LINKVALUE = RPAD(SPRCLIENTS.FSCLIENTID,30))) AS Y, SPRCLIENTS.CUSTATT25||' - '||SPRCLIENTS.METERID AS MEDIDOR, "
      + "( SELECT CLIENTES_CCYB.CT FROM NEXUS_CCYB.CLIENTES_CCYB WHERE CLIENTES_CCYB.CUENTA = SPRCLIENTS.FSCLIENTID ) AS CT, "
      + "( SELECT CLIENTES_CCYB.ALIMENTADOR FROM NEXUS_CCYB.CLIENTES_CCYB WHERE CLIENTES_CCYB.CUENTA = SPRCLIENTS.FSCLIENTID ) AS ALIMENTADOR, "
      + "( SELECT CLIENTES_CCYB.SSEE FROM NEXUS_CCYB.CLIENTES_CCYB WHERE CLIENTES_CCYB.CUENTA = SPRCLIENTS.FSCLIENTID ) AS SSEE,"
      + " SYSDATE AS ACTUALIZADO FROM NEXUS_GIS.SPRCLIENTS "
      + "WHERE SPRCLIENTS.CUSTATT16 = '1A' and SPRCLIENTS.LOGIDTO = 0 and SPRCLIENTS.CUSTATT21 = '12521' ";            

          System.out.println(query_variable);
          
            sentencia = con.createStatement();
            rs = sentencia.executeQuery(query_variable);
            return rs;
            
        } catch (SQLException e) {
            throw e;            
        }
    }
    
    static ResultSet Fechainforme(Connection conexion) throws Exception {        
        Connection con = conexion;
            //Connection con = DbConnection.getConnection(request);
        ResultSet rs = null;
        Statement sentencia;
        
        try {
             String query_variable = "select sysdate as fecha from dual";
         
            sentencia = con.createStatement();
            rs = sentencia.executeQuery(query_variable);
            return rs;
            
        } catch (SQLException e) {
            throw e;            
        }
    }

    private static void creaExcel() throws FileNotFoundException, IOException {
        
       
        Workbook wb = new XSSFWorkbook();
        
        Sheet shForzados = wb.createSheet("Electrodependientes");
        pueblaForzados(shForzados, wb);
        
        
        //try (FileOutputStream fileOut = new FileOutputStream("/ias/SMS_Informe_Diario/doc/rep_electrodep.xlsx")) {
        try (FileOutputStream fileOut = new FileOutputStream(xlsfile)) {
            wb.write(fileOut);
        }
    }

    private static void enviaMail() throws Exception {

        HashMap<String, String> hst_Mail = new HashMap<>();
         ResultSet rs;
         String Dia = null;
         String Dias = null;
         String hs = null;
         String HTML_Estucture;
         rs = Fechainforme(con) ;
         
          while (rs.next()) {
                Dias = rs.getString("FECHA");                
          }
          System.out.println(Dias);
          Dia = Dias.substring(0, 10);
          System.out.println(Dia);
          hs = Dias.substring(11, 16);
        
        if (Mensaje != null){

            HTML_Estucture = Mensaje +  " ( " + hs+ " Hs. )";
            
        } else {

            HTML_Estucture = "Este es un mail automÃ¡tico para adjuntar la base clientes electrodependientes, actualizada al dÃ­a de la fecha y en formato Excel.";
         
        }
          
         boolean enviarMail;
        try {
            enviarMail = true;
            hst_Mail.put("mailHost", "mail.edenor");
             hst_Mail.put("DE"        , "ITSM_Desarrollos_propios@edenor.com");
            hst_Mail.put("PARA"      , mailto);
            hst_Mail.put("ASUNTO", "Base de clientes electrodependientes");
            hst_Mail.put("CUERPO", HTML_Estucture);
            if (enviarMail) {
                mailSender(hst_Mail);
            }

        } catch (Exception e) {
            System.out.println("Error en generar()" + e);
            String msgCancela = e.toString();
            hst_Mail.put("CUERPO", msgCancela);
            mailSender(hst_Mail);
            throw e;
        } finally {
            try {
                System.out.println("final ");
            } catch (Exception e1) {
                System.out.println(e1);
            }
        }
    }

    private static void mailSender(HashMap<String, String> hst_values_mail) throws Exception {
        try {

            Properties properties = new Properties();
            properties.put("mail.smtp.host", hst_values_mail.get("mailHost"));
            properties.put("mail.from", hst_values_mail.get("DE"));
            properties.put("mail.debug", "true");

            Session session = Session.getInstance(properties, null);
            MimeMessage msg = new MimeMessage(session);

            msg.setFrom(new InternetAddress(hst_values_mail.get("DE")));
            msg.setFrom(InternetAddress.getLocalAddress(session));
            msg.setSubject(hst_values_mail.get("ASUNTO"));
            msg.setSentDate(new java.util.Date());

            InternetAddress[] paraArray ;
            paraArray = InternetAddress.parse(hst_values_mail.get("PARA"));
            msg.setRecipients(Message.RecipientType.TO, paraArray);

            InternetAddress[] ccArray = null;
            if (hst_values_mail.get("CC") != null) {
                ccArray = InternetAddress.parse(hst_values_mail.get("CC"));
                msg.setRecipients(Message.RecipientType.CC, ccArray);
            }

            InternetAddress[] bccArray = null;
            if (hst_values_mail.get("CCO") != null) {
                bccArray = InternetAddress.parse(hst_values_mail.get("CCO"));
                msg.setRecipients(Message.RecipientType.BCC, bccArray);
            }

            MimeMultipart multiParte = new MimeMultipart();

            BodyPart adjunto = new MimeBodyPart();
            adjunto.setDataHandler(new DataHandler(new FileDataSource(xlsfile)));
           
            if (Mensaje == null){
            adjunto.setFileName("rep_electrodep.xlsx");  
            multiParte.addBodyPart(adjunto);
            }
   


            BodyPart texto = new MimeBodyPart();
            texto.setDataHandler(new DataHandler(new HTMLDataSource(hst_values_mail.get("CUERPO"))));
            multiParte.addBodyPart(texto);

            msg.setContent(multiParte);

            int i, j, k, total;
            total = paraArray.length;
            if (ccArray != null) {
                total += ccArray.length;
            }
            if (bccArray != null) {
                total += bccArray.length;
            }

            InternetAddress[] address = new InternetAddress[total];

            for (i = 0; i < paraArray.length; i++) {
                address[i] = paraArray[i];
            }
            if (ccArray != null) {
                for (j = 0; j < ccArray.length; j++) {
                    address[i] = ccArray[j];
                    i++;
                }
            }
            if (bccArray != null) {
                for (k = 0; k < bccArray.length; k++) {
                    address[i] = bccArray[k];
                    i++;
                }
            }

            Transport transporte = session.getTransport(address[0]);
            transporte.connect();
            transporte.sendMessage(msg, address);

        } catch (SendFailedException e) {
            Address[] listaInval = e.getInvalidAddresses();
            for (Address listaInval1 : listaInval) {
                System.out.println("No encontrada: " + listaInval1.toString());
            }
        }
    }

    static class HTMLDataSource implements DataSource {

        private final String html;

        public HTMLDataSource(String htmlString) {
            html = htmlString;
        }

        @Override
        public InputStream getInputStream() throws IOException {
            if (html == null) {
                throw new IOException("Null HTML");
            }
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

    private static CellStyle creaEstilosCabe(Workbook wb) {
        XSSFCellStyle style = (XSSFCellStyle) wb.createCellStyle();

        XSSFFont negrita = (XSSFFont) wb.createFont();
        negrita.setBoldweight(XSSFFont.BOLDWEIGHT_BOLD);
        negrita.setFontHeightInPoints((short) 9);
        style.setFont(negrita);

        XSSFColor myColor = new XSSFColor(Color.decode("#C5D9F1"));
        style.setFillForegroundColor(myColor);

        style.setAlignment(HorizontalAlignment.RIGHT);

        style.setFillPattern(CellStyle.SOLID_FOREGROUND);
        style.setAlignment(CellStyle.ALIGN_CENTER);
        style.setVerticalAlignment(CellStyle.VERTICAL_CENTER);
        style.setWrapText(true);

        XSSFColor negro = new XSSFColor(Color.decode("#000000"));

        style.setBorderBottom(CellStyle.BORDER_MEDIUM);
        style.setBorderTop(CellStyle.BORDER_MEDIUM);
        style.setBorderRight(CellStyle.BORDER_MEDIUM);
        style.setBorderLeft(CellStyle.BORDER_MEDIUM);

        style.setBorderColor(XSSFCellBorder.BorderSide.LEFT, negro);
        style.setBorderColor(XSSFCellBorder.BorderSide.TOP, negro);
        style.setBorderColor(XSSFCellBorder.BorderSide.RIGHT, negro);
        style.setBorderColor(XSSFCellBorder.BorderSide.BOTTOM, negro);

        return style;
    }

    private static CellStyle creaEstilosDatos(Workbook wb, String alignment) {
        XSSFCellStyle style = (XSSFCellStyle) wb.createCellStyle();

        switch (alignment) {
            case "right":
                style.setAlignment(HorizontalAlignment.RIGHT);
                break;
            case "center":
                style.setAlignment(HorizontalAlignment.CENTER);
                break;
            case "left":
                style.setAlignment(HorizontalAlignment.LEFT);
                break;
            default:
                break;
        }

        style.setVerticalAlignment(CellStyle.VERTICAL_CENTER);
        style.setWrapText(true);

        XSSFColor negro = new XSSFColor(Color.decode("#000000"));

        style.setBorderBottom(CellStyle.BORDER_THIN);
        style.setBorderTop(CellStyle.BORDER_THIN);
        style.setBorderRight(CellStyle.BORDER_THIN);
        style.setBorderLeft(CellStyle.BORDER_THIN);

        style.setBorderColor(XSSFCellBorder.BorderSide.LEFT, negro);
        style.setBorderColor(XSSFCellBorder.BorderSide.TOP, negro);
        style.setBorderColor(XSSFCellBorder.BorderSide.RIGHT, negro);
        style.setBorderColor(XSSFCellBorder.BorderSide.BOTTOM, negro);

        return style;
    }

    private static CellStyle creaEstilosTotales(Workbook wb, String alignment, String color) {

        XSSFCellStyle style = (XSSFCellStyle) wb.createCellStyle();

        XSSFFont negrita = (XSSFFont) wb.createFont();
        negrita.setBoldweight(XSSFFont.BOLDWEIGHT_BOLD);
        style.setFont(negrita);
        
        String codigo;
        switch (color) {
            case "grisado":
                codigo = "#B8BDBF";
                break;
            case "celeste":
                codigo = "#C5D9F1";
                break;
            default:
                codigo = "#C5D9F1";
                break;
        }
        
        XSSFColor myColor = new XSSFColor(Color.decode(codigo));
        style.setFillForegroundColor(myColor);
        style.setFillPattern(CellStyle.SOLID_FOREGROUND);

        switch (alignment) {
            case "right":
                style.setAlignment(HorizontalAlignment.RIGHT);
                break;
            case "center":
                style.setAlignment(HorizontalAlignment.CENTER);
                break;
            case "left":
                style.setAlignment(HorizontalAlignment.LEFT);
                break;
        }

        style.setVerticalAlignment(CellStyle.VERTICAL_CENTER);
        style.setWrapText(true);

        XSSFColor negro = new XSSFColor(Color.decode("#000000"));

        style.setBorderBottom(CellStyle.BORDER_MEDIUM);
        style.setBorderTop(CellStyle.BORDER_MEDIUM);
        style.setBorderRight(CellStyle.BORDER_MEDIUM);
        style.setBorderLeft(CellStyle.BORDER_MEDIUM);

        style.setBorderColor(XSSFCellBorder.BorderSide.LEFT, negro);
        style.setBorderColor(XSSFCellBorder.BorderSide.TOP, negro);
        style.setBorderColor(XSSFCellBorder.BorderSide.RIGHT, negro);
        style.setBorderColor(XSSFCellBorder.BorderSide.BOTTOM, negro);

        return style;
    }

    private static void pueblaForzados(Sheet sh, Workbook wb) {
  //      ResultSet rs = null;
        Integer rowCount = 0;
        //Cabecera
        Row row = sh.createRow(rowCount++);
        
      
        CellStyle cabecera = creaEstilosCabe(wb);
        CellStyle datosLeft = creaEstilosDatos(wb, "left");
        CellStyle datosRight = creaEstilosDatos(wb, "right");
        CellStyle datosCenter = creaEstilosDatos(wb, "center");

        CellStyle totalesCenter = creaEstilosTotales(wb, "center", "celeste");
        CellStyle totalesRight = creaEstilosTotales(wb, "right", "celeste");
        CellStyle totalesGrisado = creaEstilosTotales(wb, "right", "grisado");


        Cell cell = row.createCell(0);
        cell.setCellValue("CUENTA");
        cell.setCellStyle(cabecera);
        sh.setColumnWidth(0, 4000);
        
        cell = row.createCell(1);
        cell.setCellValue("RAZON_SOCIAL");
        cell.setCellStyle(cabecera);
        sh.setColumnWidth(1, 4000);
        

        cell = row.createCell(2);
        cell.setCellValue("TELEFONO");
        cell.setCellStyle(cabecera);
        sh.setColumnWidth(2, 4500);

        cell = row.createCell(3);
        cell.setCellValue("F_ALTA");
        cell.setCellStyle(cabecera);
        sh.setColumnWidth(3, 4500);

        cell = row.createCell(4);
        cell.setCellValue("CALLE");
        cell.setCellStyle(cabecera);
        sh.setColumnWidth(4, 4000);

        cell = row.createCell(5);
        cell.setCellValue("NRO");
        cell.setCellStyle(cabecera);
        sh.setColumnWidth(5, 5000);

        cell = row.createCell(6);
        cell.setCellValue("PISO_DPTO");
        cell.setCellStyle(cabecera);
        sh.setColumnWidth(6, 5000);

        cell = row.createCell(7);
        cell.setCellValue("ENTE_CALLE_1");
        cell.setCellStyle(cabecera);
        sh.setColumnWidth(7, 5000);

        cell = row.createCell(8);
        cell.setCellValue("ENTE_CALLE_2");
        cell.setCellStyle(cabecera);
        sh.setColumnWidth(8, 3000);
        
        cell = row.createCell(9);
        cell.setCellValue("LOCALIDAD");
        cell.setCellStyle(cabecera);
        sh.setColumnWidth(9, 5000);
        
        cell = row.createCell(10);
        cell.setCellValue("PARTIDO");
        cell.setCellStyle(cabecera);
        sh.setColumnWidth(10, 5000);

        cell = row.createCell(11);
        cell.setCellValue("REGION");
        cell.setCellStyle(cabecera);
        sh.setColumnWidth(11, 5000);
        
        
        cell = row.createCell(12);
        cell.setCellValue("X");
        cell.setCellStyle(cabecera);
        sh.setColumnWidth(12, 5000);
        
        
        cell = row.createCell(13);
        cell.setCellValue("Y");
        cell.setCellStyle(cabecera);
        sh.setColumnWidth(13, 5000);
        
        cell = row.createCell(14);
        cell.setCellValue("MEDIDOR");
        cell.setCellStyle(cabecera);
        sh.setColumnWidth(14, 5000);
        
                
        cell = row.createCell(15);
        cell.setCellValue("CT");
        cell.setCellStyle(cabecera);
        sh.setColumnWidth(15, 5000);
       
                        
        cell = row.createCell(16);
        cell.setCellValue("ALIMENTADOR");
        cell.setCellStyle(cabecera);
        sh.setColumnWidth(16, 5000);
               
                        
        cell = row.createCell(17);
        cell.setCellValue("SSEE");
        cell.setCellStyle(cabecera);
        sh.setColumnWidth(17, 5000);
        
        cell = row.createCell(18);
        cell.setCellValue("ACTUALIZADO");
        cell.setCellStyle(cabecera);
        sh.setColumnWidth(18, 5000);
        
        row = sh.createRow(rowCount++);
        
/*
        try {
                      
            rs = Electrodependientes(con);
        } catch (Exception ex) {
            Logger.getLogger(Rep_Electrodependientes.class.getName()).log(Level.SEVERE, null, ex);
        }
        
  */      
        
        try {
            while (rs1.next()) {
                row = sh.createRow(rowCount);
                row.createCell(0).setCellValue((String) rs1.getString("CUENTA"));
                row.getCell(0).setCellStyle(datosLeft);
                
                row.createCell(1).setCellValue((String) rs1.getString("RAZON_SOCIAL"));
                row.getCell(1).setCellStyle(datosLeft);

                row.createCell(2).setCellValue((String) rs1.getString("TELEFONO"));
                row.getCell(2).setCellStyle(datosLeft);

                row.createCell(3).setCellValue((String) rs1.getString("F_ALTA"));
                row.getCell(3).setCellStyle(datosLeft);

                row.createCell(4).setCellValue((String) rs1.getString("CALLE"));
                row.getCell(4).setCellStyle(datosLeft);

                row.createCell(5).setCellValue((String) rs1.getString("NRO"));
                row.getCell(5).setCellStyle(datosLeft);

                row.createCell(6).setCellValue((String) rs1.getString("PISO_DPTO"));
                row.getCell(6).setCellStyle(datosLeft);

                row.createCell(7).setCellValue((String) rs1.getString("ENTE_CALLE_1"));
               row.getCell(7).setCellStyle(datosLeft);
               
               row.createCell(8).setCellValue((String) rs1.getString("ENTE_CALLE_2"));
               row.getCell(8).setCellStyle(datosLeft);
               
               row.createCell(9).setCellValue((String) rs1.getString("LOCALIDAD"));
               row.getCell(9).setCellStyle(datosLeft);
             
               row.createCell(10).setCellValue((String) rs1.getString("PARTIDO"));
               row.getCell(10).setCellStyle(datosLeft);

             
               row.createCell(11).setCellValue((String) rs1.getString("REGION"));
               row.getCell(11).setCellStyle(datosLeft);  

             
               row.createCell(12).setCellValue((String) rs1.getString("X"));
               row.getCell(12).setCellStyle(datosLeft);                 

             
               row.createCell(13).setCellValue((String) rs1.getString("Y"));
               row.getCell(13).setCellStyle(datosLeft);               

             
               row.createCell(14).setCellValue((String) rs1.getString("MEDIDOR"));
               row.getCell(14).setCellStyle(datosLeft);                              

             
               row.createCell(15).setCellValue((String) rs1.getString("CT"));
               row.getCell(15).setCellStyle(datosLeft);                  

             
               row.createCell(16).setCellValue((String) rs1.getString("ALIMENTADOR"));
               row.getCell(16).setCellStyle(datosLeft);  

             
               row.createCell(17).setCellValue((String) rs1.getString("SSEE"));
               row.getCell(17).setCellStyle(datosLeft);  
 
               String Actualizado = rs1.getString("ACTUALIZADO");
               Actualizado = Actualizado.substring(0, Actualizado.length() - 2);
               row.createCell(18).setCellValue((String) Actualizado);
               row.getCell(18).setCellStyle(datosLeft);               
               
               
               
               rowCount++;
                
            }
        } catch (SQLException ex) {
            Logger.getLogger(Rep_Electrodependientes.class.getName()).log(Level.SEVERE, null, ex);
        }
        
        
        
        

    }

}
