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
import java.sql.SQLException;
import java.sql.Types;
import java.time.LocalDate;
import java.util.ArrayList;
import java.util.Date;
import java.util.HashMap;
import java.util.List;
import java.util.Properties;

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

import oracle.jdbc.OracleCallableStatement;
/**
 *
 * @author mrrodriguez    03/05/2018
 */
public class mail_llamadas_anual {
    
    private static Connection con = null;
    private static List<HashMap<String, String>> tablaForz = null;
    private static HashMap<String, String> ultimaForz = null;
    
    
//------------------------------------------------//
//----------------------MAIN----------------------//
//------------------------------------------------//
    static public void main(String[] args) {

        try {
            System.out.println("Inicio proceso: OracleConnect " );            
            OracleConnect();
            System.out.println("Fin proceso: OracleConnect " );
            System.out.println("Inicio proceso: obtieneDatos " );
            obtieneDatos(con);
            System.out.println("Fin proceso: obtieneDatos " );
            System.out.println("Inicio proceso: creaExcel " );
            creaExcel();
            System.out.println("Fin proceso: creaExcel " );
            System.out.println("Inicio proceso: enviaMail " );
            enviaMail();
             System.out.println("Fin proceso: enviaMail " );
        } catch (Exception ex) {
            System.out.println(ex);
        }

    }
//------------------------------------------------//
//------------------------------------------------//
    
    
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
                    .getConnection("jdbc:oracle:thin:@NEXGISPR02.PRO.EDENOR:1528:gispr01s", "SVC_ORA_GIS", "jv506uzy");
        } catch (SQLException e) {
            switch (e.getErrorCode()) {
                case 1017:    //USUARIO O CLAVE IVALIDO; LOGON DENIED
                    System.out.println("Usuario de la base de datos incorrecto");
                    break;
                case 28000:    //CUENTA BLOQUEADA
                    System.out.println("Cuenta bloqueada");
                    break;
                case 2391:		//EXCEDIDO EN CONEXIONES
                    System.out.println("Límite de conexiones excedido, intentar más tarde");
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

    private static void obtieneDatos(Connection con) throws SQLException {

        String retorno;
        String plForzados = "declare\n" +
"\n" +
"v_mes number;\n" +
"v_anio number;\n" +
"v_camp_min number;\n" +
"v_camp_max number;\n" +
"v_ivr_edenor number;\n" +
"v_cant_forzados number;\n" +
"v_cant_programados number;\n" +
"v_prog_cat number;\n" +
"v_total_personalizada number:=0;\n" +
"v_total_program_sondeos number:=0;\n" +
"v_total_forzados_sondeos number:=0;\n" +
"v_total_ivr_edenor number:=0;\n" +
"v_fecha_sondeos Varchar2(20 byte);\n" +
"v_month Varchar2(20 byte);\n" +
"v_texto_salida VARCHAR2(30000):=null; \n" +
"v_totales_salida VARCHAR2(30000):=null; \n" +
"\n" +
"\n" +
"begin\n" +
"\n" +
"-- capturo el mes en curso\n" +
"select extract (month from sysdate), extract (year from sysdate) into v_mes, v_anio from dual;\n" +
"\n" +
"-- Corro desde Enero hasta el mes en curso, para sacar las campañas de cada mes\n" +
"for i in 1 .. (v_mes-1) loop\n" +
"    \n" +
"--verifico que cumpla con el formato de fecha sondeas\n" +
"if length (i)=1 then\n" +
" v_fecha_sondeos:= to_char(v_anio)||'-0'||to_char(i);\n" +
" v_month:='0'||to_char(i);\n" +
"else\n" +
" v_fecha_sondeos:= to_char(v_anio)||to_char(i);\n" +
" v_month:=to_char(i);\n" +
"end if;\n" +
"\n" +
"--extraigo el mes como nombre\n" +
"select to_char(to_date(v_month,'mm'), 'Month','nls_date_language=spanish') as mes into v_month from dual;\n" +
"\n" +
"-- capturo el rango de campañas del mes\n" +
"select min(nro_camp), max(nro_camp) into v_camp_min,v_camp_max\n" +
"from NEXUS_GIS.OMS_CUST_SITUATION_INTERF where extract (month from fe_ultim_update) =i and extract (year from fe_ultim_update)=extract (year from sysdate);\n" +
"\n" +
"--sumo las llamadas efectivas que se realizaron para el mes\n" +
"select sum(cantidad) into v_ivr_edenor from (\n" +
"select count(*) as cantidad from NEXUS_GIS.OMS_CUST_SITUATION_INTERF p where P.NRO_CAMP between v_camp_min and v_camp_max  and last_action =3\n" +
"union all\n" +
"select count(*) as cantidad from NEXUS_GIS.OMS_RELLAMADAS p where P.NRO_CAMP between v_camp_min and v_camp_max  \n" +
");\n" +
"\n" +
"--Cantidad de llamados forzados MT sondeos\n" +
"select count(*) into v_cant_forzados from NEXUS_GIS.LLAM_CDR_DEV_SONDEOS where tipocampania = 'Corte Forzado' and estado_llamada ='ANSWER' and substr(fecha_llamada,0,7) =v_fecha_sondeos;\n" +
"\n" +
"--Cantidad de llamados Programados MT sondeos\n" +
"select count(*) into v_cant_programados from NEXUS_GIS.LLAM_CDR_DEV_SONDEOS where tipocampania = 'Corte Programado' and estado_llamada ='ANSWER' and substr(fecha_llamada,0,7) =v_fecha_sondeos;\n" +
"\n" +
"--Cantidad de llamados personalizados \n" +
"select count(*) into v_prog_cat from NEXUS_GIS.LLAM_REG_DEV_T2T3 where dev_llamada ='Contactado' and extract (month from fecha_creacion) =i and extract (year from fecha_creacion)=extract (year from sysdate);\n" +
"\n" +
"--Totales\n" +
"v_total_personalizada:=v_total_personalizada+v_prog_cat;\n" +
"v_total_program_sondeos:= v_total_program_sondeos + v_cant_programados;\n" +
"v_total_forzados_sondeos := v_total_forzados_sondeos + v_cant_forzados;\n" +
"v_total_ivr_edenor:= v_total_ivr_edenor+v_ivr_edenor;\n" +
"\n" +
"v_texto_salida:= v_texto_salida||chr(59)||trim(v_month)||chr(59)||v_prog_cat||chr(59)||v_cant_programados||chr(59)||v_cant_forzados||chr(59)||v_ivr_edenor||'X';\n" +
"\n" +
"end loop;\n" +
"v_totales_salida:= to_char(v_total_personalizada)||chr(59)||to_char(v_total_program_sondeos)||chr(59)||to_char(v_total_forzados_sondeos)||chr(59)||to_char(v_total_ivr_edenor);\n" +
"?:=v_texto_salida||'W'||v_totales_salida;\n" +
"?:=v_totales_salida;\n" +
"\n" +
"end;";
       
        OracleCallableStatement cs;

        //Captura de los datos de llamadas
        cs = (OracleCallableStatement) con.prepareCall(plForzados);
        cs.registerOutParameter(1, Types.VARCHAR);
        cs.registerOutParameter(2, Types.VARCHAR);
        cs.execute();

        retorno = (String)cs.getObject(1);        
        String strarray[] = retorno.split("W");       
        tablaForz = strToHashForzados(strarray[0]);
        ultimaForz = strToHashUltLineaForzados(strarray[1]);

        

    }

    private static void creaExcel() throws FileNotFoundException, IOException {
    	
    	String year = String.valueOf(LocalDate.now().getYear());
        Workbook wb = new XSSFWorkbook();

        Sheet shForzados = wb.createSheet(year);
        pueblaForzados(shForzados, wb, year);

        try (FileOutputStream fileOut = new FileOutputStream("./doc/rep_anual_llam_salientes.xlsx")) {
            wb.write(fileOut);
        }
    }

    private static void enviaMail() throws Exception {

        HashMap<String, String> hst_Mail = new HashMap<>();

        String HTML_Estucture = "Este es un mail automático que exporta el resumen de llamadas salientes Efectivas al Excel adjunto.";
        boolean enviarMail;
        try {
            enviarMail = true;
            hst_Mail.put("mailHost", "mail.edenor");
            hst_Mail.put("DE", "ITSM_Llamadas_salientes@edenor.com");            
            hst_Mail.put("PARA", "grabinovich@edenor.com,pmulet@edenor.com,hgonzalez@edenor.com,ITSM_Llamadas_salientes@edenor.com");
            hst_Mail.put("CC", "ITSM_Desarrollos_propios@edenor.com");
            hst_Mail.put("ASUNTO", "Informacion Resumen Llamadas Salientes Efectivas Consolidado Mensual");
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
            adjunto.setDataHandler(new DataHandler(new FileDataSource("./doc/rep_anual_llam_salientes.xlsx")));
            adjunto.setFileName("rep_anual_llam_salientes.xlsx");
            multiParte.addBodyPart(adjunto);

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
        } catch (Exception e) {
            System.out.println("Exception (mailSender) : " + e);
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

    private static void pueblaForzados(Sheet sh, Workbook wb, String year) {

        Integer rowCount = 0;
        //Cabecera
        Row row = sh.createRow(rowCount++);
        sh.addMergedRegion(new CellRangeAddress(0, 0, 0, 4));//LLAMADAS 2018
        sh.addMergedRegion(new CellRangeAddress(1, 3, 0, 0));//MES
        sh.addMergedRegion(new CellRangeAddress(1, 1, 1, 3));//MEDIA TENSION
        sh.addMergedRegion(new CellRangeAddress(1, 1, 4, 4));//BAJA TENSION
        sh.addMergedRegion(new CellRangeAddress(2, 2, 1, 2));//CORTES PROGRAMADOS
        sh.addMergedRegion(new CellRangeAddress(2, 2, 3, 3));//CORTES FORZADOS
        sh.addMergedRegion(new CellRangeAddress(2, 3, 4, 4));//IVR EDENOR
        sh.addMergedRegion(new CellRangeAddress(3, 3, 1, 1));//PERSONALIZADAS
        sh.addMergedRegion(new CellRangeAddress(3, 3, 2, 2));//IVR SONDEOS
        sh.addMergedRegion(new CellRangeAddress(3, 3, 3, 3));//IVR SONDEOS
        
        

        CellStyle cabecera = creaEstilosCabe(wb);
        CellStyle datosLeft = creaEstilosDatos(wb, "left");
        CellStyle datosRight = creaEstilosDatos(wb, "right");
        CellStyle datosCenter = creaEstilosDatos(wb, "center");

        CellStyle totalesCenter = creaEstilosTotales(wb, "center", "celeste");
        CellStyle totalesRight = creaEstilosTotales(wb, "right", "celeste");
        CellStyle totalesGrisado = creaEstilosTotales(wb, "right", "grisado");
        
        //primer linea titulo
        Cell cell = row.createCell(0);
        cell.setCellValue("LLAMADAS " + year);
        cell.setCellStyle(cabecera);
        for (int i = 1; i < 5; i++) {
            cell = row.createCell(i);
            cell.setCellStyle(cabecera);
        }
        //segunda linea titulo
        row = sh.createRow(rowCount++);

        cell = row.createCell(0);
        cell.setCellValue("MES");
        cell.setCellStyle(cabecera);

        cell = row.createCell(1);
        cell.setCellValue("MEDIA TENSION");
        cell.setCellStyle(cabecera);
         
        cell = row.createCell(2);        
        cell.setCellStyle(cabecera);

        cell = row.createCell(3);       
        cell.setCellStyle(cabecera);

        cell = row.createCell(4);
        cell.setCellValue("BAJA TENSION");
        cell.setCellStyle(cabecera);
        
        //tercer linea titulo
        row = sh.createRow(rowCount++);
        sh.setDefaultRowHeight((short) 25.5);
        cell = row.createCell(0);    
        cell.setCellStyle(cabecera);
        
        cell = row.createCell(1);
        cell.setCellValue("CORTES PROGRAMADOS");
        cell.setCellStyle(cabecera);
        
       
        cell = row.createCell(2);
        cell.setCellStyle(cabecera);       

        cell = row.createCell(3);
        cell.setCellValue("CORTES FORZADOS");
        cell.setCellStyle(cabecera);       

        cell = row.createCell(4);
        cell.setCellValue("IVR EDENOR");
        cell.setCellStyle(cabecera);


        row = sh.createRow(rowCount++);

        cell = row.createCell(0);
        cell.setCellStyle(cabecera);

        cell = row.createCell(1);
        cell.setCellValue("PERSONALIZADAS");
        cell.setCellStyle(cabecera);

        cell = row.createCell(2);
        cell.setCellValue("IVR SONDEOS");
        cell.setCellStyle(cabecera);
       
       
        cell = row.createCell(3);
        cell.setCellValue("IVR SONDEOS");
        cell.setCellStyle(cabecera);


        cell = row.createCell(4);        
        cell.setCellStyle(cabecera);

        //Cuerpo
        for (HashMap fila : tablaForz) {
            row = sh.createRow(rowCount);
            row.createCell(0).setCellValue((String) fila.get("MES"));
            row.getCell(0).setCellStyle(datosLeft);

            row.createCell(1).setCellValue((String) fila.get("PERSONALIZADAS"));
            row.getCell(1).setCellStyle(datosRight);

            row.createCell(2).setCellValue((String) fila.get("PROG IVR SONDEOS"));
            row.getCell(2).setCellStyle(datosRight);

            row.createCell(3).setCellValue((String) fila.get("FORZ IVR SONDEOS"));
            row.getCell(3).setCellStyle(datosRight);

            row.createCell(4).setCellValue((String) fila.get("IVR EDENOR"));
            row.getCell(4).setCellStyle(datosRight);

            rowCount++;
        }

        ///Totales        
        row = sh.createRow(rowCount);
        row.setHeight((short) 800);
        row.createCell(0).setCellValue((String) "TOTAL");
        row.getCell(0).setCellStyle(totalesCenter);

      

        row.createCell(1).setCellValue((String) ultimaForz.get("TOT_PERSONALIZA"));
        row.getCell(1).setCellStyle(totalesRight);

        row.createCell(2).setCellValue((String) ultimaForz.get("TOT_PROG_IVRSONDEOS"));
        row.getCell(2).setCellStyle(totalesRight);

        row.createCell(3).setCellValue((String) ultimaForz.get("TOT_FORZ_IVRSONDEOS"));
        row.getCell(3).setCellStyle(totalesRight);

        row.createCell(4).setCellValue((String) ultimaForz.get("TOT_IVREDENOR"));
        row.getCell(4).setCellStyle(totalesRight);

       
    }

    

    private static String clobToStr(Clob clb) {
        StringBuilder sb = new StringBuilder();

        try {
            final Reader reader = clb.getCharacterStream();
            try (BufferedReader br = new BufferedReader(reader)) {
                int b;
                while (-1 != (b = br.read())) {
                    sb.append((char) b);
                }
            }
        } catch (SQLException e) {
            System.out.println("RedElectrica::clobToStr: SQL. No se pudo convertir CLOB a String");
        } catch (IOException e) {
            System.out.println("RedElectrica::clobToStr: IO. No se pudo convertir CLOB a String");
        }

        return sb.toString();
    }

    static List<HashMap<String, String>> strToHashForzados(String reporte) {

        List<HashMap<String, String>> tabla = new ArrayList<>();
        reporte = reporte.substring(0, reporte.length() - 1);

        String[] lineas = reporte.split("X");

        for (String linea : lineas) {

            String[] campo = linea.split(";");
            HashMap<String, String> fila = new HashMap<>();
            fila.put("MES", campo[1]);
            fila.put("PERSONALIZADAS", campo[2]);
            fila.put("PROG IVR SONDEOS", campo[3]);
            fila.put("FORZ IVR SONDEOS", campo[4]);
            fila.put("IVR EDENOR", campo[5]);           
            tabla.add(fila);
        }

        return tabla;
    }

    static HashMap<String, String> strToHashUltLineaForzados(String reporte) {

        HashMap<String, String> fila = null;
        String[] lineas = reporte.split(";", -1);
        for (int i = 0; i < lineas.length; i++) {
            fila = new HashMap<>();
            fila.put("TOT_PERSONALIZA", lineas[i++]);
            fila.put("TOT_PROG_IVRSONDEOS", lineas[i++]);
            fila.put("TOT_FORZ_IVRSONDEOS", lineas[i++]);
            fila.put("TOT_IVREDENOR", lineas[i++]);
        }
        return fila;
    }

}

