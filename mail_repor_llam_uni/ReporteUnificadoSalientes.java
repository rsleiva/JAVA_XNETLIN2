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
import java.util.ArrayList;
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

public class ReporteUnificadoSalientes {

    private static Connection con = null;
    private static List<HashMap<String, String>> tablaForz = null;
    private static HashMap<String, String> ultimaForz = null;
    private static List<HashMap<String, String>> tablaProgra = null;
    private static HashMap<String, String> ultimaProgra = null;
    private static List<HashMap<String, String>> tablaPuntuales = null;
    private static HashMap<String, String> ultimaPuntuales = null;
    
//------------------------------------------------//
//----------------------MAIN----------------------//
//------------------------------------------------//
    static public void main(String[] args) {

        try {
            OracleConnect();
            obtieneDatos(con);
            creaExcel();
            enviaMail();
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
                    .getConnection("jdbc:oracle:thin:@ltronxgisbdpr03.pro.edenor:1528:GISPR03", "SVC_ORA_GIS", "jv506uzy");
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

    private static void obtieneDatos(Connection con) throws SQLException {

        String retorno;
        String plForzados = "{?=call NEXUS_GIS.LLAM_PROCESOS.LLAM_REPORT_FORZADOS_MT(?)}";
        String plProgramados = "{?=call NEXUS_GIS.LLAM_PROCESOS.LLAM_REPORT_PROGRAMADOS_MT(?)}";
        String plPuntuales = "{?=call NEXUS_GIS.LLAM_PROCESOS.LLAM_REPORT_PUNTUAL_BT(?)}";
        OracleCallableStatement cs;

        //Forzados
        cs = (OracleCallableStatement) con.prepareCall(plForzados);
        cs.registerOutParameter(1, OracleTypes.CLOB);
        cs.registerOutParameter(2, OracleTypes.CLOB);
        cs.execute();

        retorno =clobToStr(cs.getCLOB(1));
        String strarray[] = retorno.split("T");
        strarray[1] = strarray[1].replace(strarray[1], "T" + strarray[1]);
        tablaForz = strToHashForzados(strarray[0]);
        ultimaForz = strToHashUltLineaForzados(strarray[1]);

        System.out.println("luego del llamado a forzados");
        //Programados
        cs = (OracleCallableStatement) con.prepareCall(plProgramados);
        cs.registerOutParameter(1, OracleTypes.CLOB);
        cs.registerOutParameter(2, OracleTypes.CLOB);
        cs.execute();

        retorno = clobToStr(cs.getCLOB(1));
        strarray = retorno.split("To");
        strarray[1] = strarray[1].replace(strarray[1], "To" + strarray[1]);
        tablaProgra = strToHashProgramados(strarray[0]);
        ultimaProgra = strToHashUltLineaProgramados(strarray[1]);

        System.out.println("luego del llamado a programados");
        //Puntuales
        cs = (OracleCallableStatement) con.prepareCall(plPuntuales);
        cs.registerOutParameter(1, OracleTypes.CLOB);
        cs.registerOutParameter(2, OracleTypes.CLOB);
        cs.execute();

        retorno=clobToStr(cs.getCLOB(1));
        strarray = retorno.split("T");
        strarray[1] = strarray[1].replace(strarray[1], "T" + strarray[1]);
        tablaPuntuales = strToHashPuntuales(strarray[0]);
        ultimaPuntuales = strToHashUltLineaPuntuales(strarray[1]);
        
        System.out.println("luego del llamado a puntuales");

    }

    private static void creaExcel() throws FileNotFoundException, IOException {
        Workbook wb = new XSSFWorkbook();

        Sheet shForzados = wb.createSheet("Documentos Forzados");
        Sheet shProgramados = wb.createSheet("Documentos Programados");
        Sheet shPuntuales = wb.createSheet("Documentos Puntuales");

        pueblaForzados(shForzados, wb);
        pueblaProgramados(shProgramados, wb);
        pueblaPuntuales(shPuntuales, wb);

        try (FileOutputStream fileOut = new FileOutputStream("./doc/rep_unificado_llam_salientes.xlsx")) {
//        try (FileOutputStream fileOut = new FileOutputStream("H://NetBeansProjects//ReporteUnificadoSalientes//ReporteUnificadoSalientes//rep_unificado_llam_salientes.xlsx")) {
            wb.write(fileOut);
        }
    }

    private static void enviaMail() throws Exception {

        HashMap<String, String> hst_Mail = new HashMap<>();

        String HTML_Estucture = "Este es un mail automatico que exporta el resumen de llamadas salientes al Excel adjunto.";
        boolean enviarMail;
        try {
            enviarMail = true;
            hst_Mail.put("mailHost", "mail.edenor");
            hst_Mail.put("DE", "centrodeinformacion@edenor.com");
            hst_Mail.put("PARA"      , "OSUAREZ@edenor.com,PPEREZ@edenor.com,Okovalow@edenor.com,PMAZZA@edenor.com,HGONZALEZ@edenor.com,ELAFUENTE@edenor.com,RPUCCAR@edenor.com,MSCALELLA@edenor.com,grabinovich@edenor.com,itsm_desarrollos_propios@edenor.com,dteodosio@edenor.com,PMULET@edenor.com,gmeyer@edenor.com,LD_NEXUS_PRODUCCION@edenor.com");
            //hst_Mail.put("PARA"      , "vdiciurcio@edenor.com,rsleiva@edenor.com");
            hst_Mail.put("ASUNTO", "Informacion Resumen Llamadas Salientes");
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
            adjunto.setDataHandler(new DataHandler(new FileDataSource("./doc/rep_unificado_llam_salientes.xlsx")));
            //adjunto.setDataHandler(new DataHandler(new FileDataSource("H://NetBeansProjects//ReporteUnificadoSalientes//ReporteUnificadoSalientes//rep_unificado_llam_salientes.xlsx")));
            adjunto.setFileName("rep_unificado_llam_salientes.xlsx");
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

    private static void pueblaForzados(Sheet sh, Workbook wb) {

        Integer rowCount = 0;
        //Cabecera
        Row row = sh.createRow(rowCount++);
        sh.addMergedRegion(new CellRangeAddress(0, 0, 0, 10));
        sh.addMergedRegion(new CellRangeAddress(1, 4, 0, 0));
        sh.addMergedRegion(new CellRangeAddress(1, 4, 1, 1));
        sh.addMergedRegion(new CellRangeAddress(1, 1, 2, 4));
        sh.addMergedRegion(new CellRangeAddress(1, 4, 5, 5));
        sh.addMergedRegion(new CellRangeAddress(1, 4, 6, 6));
        sh.addMergedRegion(new CellRangeAddress(1, 1, 7, 10));
        sh.addMergedRegion(new CellRangeAddress(2, 4, 7, 7));
        sh.addMergedRegion(new CellRangeAddress(2, 4, 8, 8));
        sh.addMergedRegion(new CellRangeAddress(2, 4, 9, 9));
        sh.addMergedRegion(new CellRangeAddress(2, 4, 10, 10));

        CellStyle cabecera = creaEstilosCabe(wb);
        CellStyle datosLeft = creaEstilosDatos(wb, "left");
        CellStyle datosRight = creaEstilosDatos(wb, "right");
        CellStyle datosCenter = creaEstilosDatos(wb, "center");

        CellStyle totalesCenter = creaEstilosTotales(wb, "center", "celeste");
        CellStyle totalesRight = creaEstilosTotales(wb, "right", "celeste");
        CellStyle totalesGrisado = creaEstilosTotales(wb, "right", "grisado");

        Cell cell = row.createCell(0);
        cell.setCellValue("FORZADOS MT");
        cell.setCellStyle(cabecera);
        for (int i = 1; i < 11; i++) {
            cell = row.createCell(i);
            cell.setCellStyle(cabecera);
        }
        row = sh.createRow(rowCount++);

        cell = row.createCell(0);
        cell.setCellValue("N° de campaña");
        cell.setCellStyle(cabecera);

        cell = row.createCell(1);
        cell.setCellValue("Fecha");
        cell.setCellStyle(cabecera);

        cell = row.createCell(2);
        cell.setCellValue("(*)");
        cell.setCellStyle(cabecera);

        cell = row.createCell(3);
        cell.setCellStyle(cabecera);

        cell = row.createCell(4);
        cell.setCellStyle(cabecera);

        cell = row.createCell(5);
        cell.setCellValue("Clientes Afectados Por Cortes");
        cell.setCellStyle(cabecera);
        sh.setColumnWidth(5, 5000);

        cell = row.createCell(6);
        cell.setCellValue("Cantidad De Llamados");
        cell.setCellStyle(cabecera);
        sh.setColumnWidth(6, 5000);

        cell = row.createCell(7);
        cell.setCellValue("COMUNICACIÓN");
        cell.setCellStyle(cabecera);

        cell = row.createCell(8);
        cell.setCellStyle(cabecera);

        cell = row.createCell(9);
        cell.setCellStyle(cabecera);

        cell = row.createCell(10);
        cell.setCellStyle(cabecera);

        row = sh.createRow(rowCount++);

        cell = row.createCell(0);
        cell.setCellStyle(cabecera);

        cell = row.createCell(1);
        cell.setCellStyle(cabecera);

        cell = row.createCell(2);
        cell.setCellValue("6");
        cell.setCellStyle(cabecera);

        cell = row.createCell(3);
        cell.setCellValue("12");
        cell.setCellStyle(cabecera);

        cell = row.createCell(4);
        cell.setCellValue("18");
        cell.setCellStyle(cabecera);

        cell = row.createCell(5);
        cell.setCellStyle(cabecera);

        cell = row.createCell(6);
        cell.setCellStyle(cabecera);

        cell = row.createCell(7);
        cell.setCellValue("Exitosa");
        sh.setColumnWidth(7, 4000);
        cell.setCellStyle(cabecera);

        cell = row.createCell(8);
        cell.setCellValue("No Contesta");
        sh.setColumnWidth(8, 4000);
        cell.setCellStyle(cabecera);

        cell = row.createCell(9);
        cell.setCellValue("Nro Teléfono Erroneo");
        sh.setColumnWidth(9, 4000);
        cell.setCellStyle(cabecera);

        cell = row.createCell(10);
        cell.setCellValue("Suspendidos");
        sh.setColumnWidth(10, 4000);
        cell.setCellStyle(cabecera);

        row = sh.createRow(rowCount++);

        cell = row.createCell(0);
        cell.setCellStyle(cabecera);

        cell = row.createCell(1);
        cell.setCellStyle(cabecera);

        cell = row.createCell(2);
        cell.setCellValue("A");
        cell.setCellStyle(cabecera);

        cell = row.createCell(3);
        cell.setCellValue("A");
        cell.setCellStyle(cabecera);

        cell = row.createCell(4);
        cell.setCellValue("A");
        cell.setCellStyle(cabecera);

        cell = row.createCell(5);
        cell.setCellStyle(cabecera);

        cell = row.createCell(6);
        cell.setCellStyle(cabecera);

        cell = row.createCell(7);
        cell.setCellStyle(cabecera);

        cell = row.createCell(8);
        cell.setCellStyle(cabecera);

        cell = row.createCell(9);
        cell.setCellStyle(cabecera);

        cell = row.createCell(10);
        cell.setCellStyle(cabecera);

        row = sh.createRow(rowCount++);

        cell = row.createCell(0);
        cell.setCellStyle(cabecera);

        cell = row.createCell(1);
        cell.setCellStyle(cabecera);

        cell = row.createCell(2);
        cell.setCellValue("12");
        cell.setCellStyle(cabecera);

        cell = row.createCell(3);
        cell.setCellValue("18");
        cell.setCellStyle(cabecera);

        cell = row.createCell(4);
        cell.setCellValue("24");
        cell.setCellStyle(cabecera);

        cell = row.createCell(5);
        cell.setCellStyle(cabecera);

        cell = row.createCell(6);
        cell.setCellStyle(cabecera);

        cell = row.createCell(7);
        cell.setCellStyle(cabecera);

        cell = row.createCell(8);
        cell.setCellStyle(cabecera);

        cell = row.createCell(9);
        cell.setCellStyle(cabecera);

        cell = row.createCell(10);
        cell.setCellStyle(cabecera);
        
        //Cuerpo
        for (HashMap fila : tablaForz) {
            row = sh.createRow(rowCount);
            row.createCell(0).setCellValue((String) fila.get("CAMPANIA"));
            row.getCell(0).setCellStyle(datosLeft);

            row.createCell(1).setCellValue((String) fila.get("FECHA"));
            row.getCell(1).setCellStyle(datosCenter);

            row.createCell(2).setCellValue((String) fila.get("TURNO_6A12"));
            row.getCell(2).setCellStyle(datosCenter);

            row.createCell(3).setCellValue((String) fila.get("TURNO_12A18"));
            row.getCell(3).setCellStyle(datosCenter);

            row.createCell(4).setCellValue((String) fila.get("TURNO_18A24"));
            row.getCell(4).setCellStyle(datosCenter);

            row.createCell(5).setCellValue((String) fila.get("CLI_AFECTADOS"));
            row.getCell(5).setCellStyle(datosRight);

            row.createCell(6).setCellValue((String) fila.get("CANT_LLAMADAS"));
            row.getCell(6).setCellStyle(datosRight);

            row.createCell(7).setCellValue((String) fila.get("EXITOSA"));
            row.getCell(7).setCellStyle(datosRight);

            row.createCell(8).setCellValue((String) fila.get("NO_CONTESTA"));
            row.getCell(8).setCellStyle(datosRight);

            row.createCell(9).setCellValue((String) fila.get("TEL_INCORRECTO"));
            row.getCell(9).setCellStyle(datosRight);

            row.createCell(10).setCellValue((String) fila.get("SUSPENDIDOS"));
            row.getCell(10).setCellStyle(datosRight);

            rowCount++;
        }

        ///Totales
        sh.addMergedRegion(new CellRangeAddress(rowCount, rowCount, 1, 4));
        row = sh.createRow(rowCount);
        row.setHeight((short) 800);
        row.createCell(0).setCellValue((String) ultimaForz.get("TOT_SEM_LITERAL"));
        row.getCell(0).setCellStyle(totalesCenter);

        cell = row.createCell(1);
        cell.setCellStyle(totalesGrisado);

        cell = row.createCell(2);
        cell.setCellStyle(totalesGrisado);

        cell = row.createCell(3);
        cell.setCellStyle(totalesGrisado);

        cell = row.createCell(4);
        cell.setCellStyle(totalesGrisado);

        row.createCell(5).setCellValue((String) ultimaForz.get("TOT_CLI_AFECTADOS"));
        row.getCell(5).setCellStyle(totalesRight);

        row.createCell(6).setCellValue((String) ultimaForz.get("TOT_CANT_LLAMADOS"));
        row.getCell(6).setCellStyle(totalesRight);

        row.createCell(7).setCellValue((String) ultimaForz.get("TOT_EXITOSA"));
        row.getCell(7).setCellStyle(totalesRight);

        row.createCell(8).setCellValue((String) ultimaForz.get("TOT_NO_CONTESTA"));
        row.getCell(8).setCellStyle(totalesRight);

        row.createCell(9).setCellValue((String) ultimaForz.get("TOT_TEL_ERROR"));
        row.getCell(9).setCellStyle(totalesRight);

        row.createCell(10).setCellValue((String) ultimaForz.get("TOT_SUSPEN"));
        row.getCell(10).setCellStyle(totalesRight);

        for (int i = 0; i < 5; i++) {
            sh.autoSizeColumn(i);
        }

    }

    private static void pueblaProgramados(Sheet sh, Workbook wb) {

        Integer rowCount = 0;

        //Cabecera
        Row row = sh.createRow(rowCount++);
        sh.addMergedRegion(new CellRangeAddress(0, 0, 0, 16));
        sh.addMergedRegion(new CellRangeAddress(1, 4, 0, 0));
        sh.addMergedRegion(new CellRangeAddress(1, 4, 1, 1));
        sh.addMergedRegion(new CellRangeAddress(1, 1, 2, 4));
        sh.addMergedRegion(new CellRangeAddress(1, 4, 5, 5));
        sh.addMergedRegion(new CellRangeAddress(1, 4, 6, 6));
        sh.addMergedRegion(new CellRangeAddress(1, 4, 7, 7));
        sh.addMergedRegion(new CellRangeAddress(1, 1, 8, 10));
        sh.addMergedRegion(new CellRangeAddress(1, 4, 11, 11));
        sh.addMergedRegion(new CellRangeAddress(1, 4, 12, 12));
        sh.addMergedRegion(new CellRangeAddress(1, 1, 13, 15));

        sh.addMergedRegion(new CellRangeAddress(2, 4, 8, 8));
        sh.addMergedRegion(new CellRangeAddress(2, 4, 9, 9));
        sh.addMergedRegion(new CellRangeAddress(2, 4, 10, 10));

        sh.addMergedRegion(new CellRangeAddress(2, 4, 13, 13));
        sh.addMergedRegion(new CellRangeAddress(2, 4, 14, 14));
        sh.addMergedRegion(new CellRangeAddress(2, 4, 15, 15));
         sh.addMergedRegion(new CellRangeAddress(1, 4, 16, 16));

        //////////////////
        CellStyle cabecera = creaEstilosCabe(wb);
        CellStyle datosLeft = creaEstilosDatos(wb, "left");
        CellStyle datosRight = creaEstilosDatos(wb, "right");
        CellStyle datosCenter = creaEstilosDatos(wb, "center");

        CellStyle totalesRight = creaEstilosTotales(wb, "right", "celeste");
        CellStyle totalesCenter = creaEstilosTotales(wb, "center", "celeste");
        CellStyle totalesGrisado = creaEstilosTotales(wb, "center", "grisado");

        Cell cell = row.createCell(0);
        cell.setCellValue("PROGRAMADOS MT");
        cell.setCellStyle(cabecera);
        for (int i = 1; i <= 16; i++) {
            cell = row.createCell(i);
            cell.setCellStyle(cabecera);
        }
        row = sh.createRow(rowCount++);

        cell = row.createCell(0);
        cell.setCellValue("N° de campaña");
        cell.setCellStyle(cabecera);

        cell = row.createCell(1);
        cell.setCellValue("Fecha");
        cell.setCellStyle(cabecera);

        cell = row.createCell(2);
        cell.setCellValue("(*)");
        cell.setCellStyle(cabecera);

        cell = row.createCell(3);
        cell.setCellStyle(cabecera);

        cell = row.createCell(4);
        cell.setCellStyle(cabecera);

        cell = row.createCell(5);
        cell.setCellValue("Total Clientes Afectados");
        cell.setCellStyle(cabecera);
        sh.setColumnWidth(5, 3000);

        cell = row.createCell(6);
        cell.setCellValue("Clientes T1 Afectados por Cortes");
        cell.setCellStyle(cabecera);
        sh.setColumnWidth(6, 3000);

        cell = row.createCell(7);
        cell.setCellValue("Cantidad De Llamados");
        cell.setCellStyle(cabecera);
        sh.setColumnWidth(7, 3000);

        cell = row.createCell(8);
        cell.setCellValue("IVR");
        cell.setCellStyle(cabecera);

        cell = row.createCell(9);
        cell.setCellStyle(cabecera);

        cell = row.createCell(10);
        cell.setCellStyle(cabecera);

        cell = row.createCell(11);
        cell.setCellValue("Clientes T2/T3 Afectados por Cortes");
        cell.setCellStyle(cabecera);
        sh.setColumnWidth(11, 3000);

        cell = row.createCell(12);
        cell.setCellValue("Cantidad De Llamados");
        cell.setCellStyle(cabecera);
        sh.setColumnWidth(12, 3000);

        cell = row.createCell(13);
        cell.setCellValue("CAT");
        cell.setCellStyle(cabecera);

        cell = row.createCell(14);
        cell.setCellStyle(cabecera);

        cell = row.createCell(15);
        cell.setCellStyle(cabecera);
        
        cell = row.createCell(16);
        cell.setCellValue("Fecha Inicio Corte");
        cell.setCellStyle(cabecera);

        row = sh.createRow(rowCount++);

        cell = row.createCell(0);
        cell.setCellStyle(cabecera);

        cell = row.createCell(1);
        cell.setCellStyle(cabecera);

        cell = row.createCell(2);
        cell.setCellValue("6");
        cell.setCellStyle(cabecera);

        cell = row.createCell(3);
        cell.setCellValue("12");
        cell.setCellStyle(cabecera);

        cell = row.createCell(4);
        cell.setCellValue("18");
        cell.setCellStyle(cabecera);

        cell = row.createCell(5);
        cell.setCellStyle(cabecera);

        cell = row.createCell(6);
        cell.setCellStyle(cabecera);

        cell = row.createCell(7);
        cell.setCellStyle(cabecera);
        
        cell = row.createCell(8);
        cell.setCellValue("Exitosa");
        cell.setCellStyle(cabecera);
        sh.setColumnWidth(8, 3000);
        
        cell = row.createCell(9);
        cell.setCellValue("No Contesta");
        cell.setCellStyle(cabecera);
        sh.setColumnWidth(8, 3000);

        cell = row.createCell(10);
        cell.setCellValue("Nro Telefono Erroneo");
        cell.setCellStyle(cabecera);
        sh.setColumnWidth(9, 3000);

        
        cell = row.createCell(11);
        cell.setCellStyle(cabecera);

        cell = row.createCell(12);
        cell.setCellStyle(cabecera);

        cell = row.createCell(13);
        cell.setCellValue("Exitosa");
        cell.setCellStyle(cabecera);

        cell = row.createCell(14);
        cell.setCellValue("No contesta");
        cell.setCellStyle(cabecera);

        cell = row.createCell(15);
        cell.setCellValue("Nro telefono erroneo");
        cell.setCellStyle(cabecera);

        row = sh.createRow(rowCount++);

        cell = row.createCell(0);
        cell.setCellStyle(cabecera);

        cell = row.createCell(1);
        cell.setCellStyle(cabecera);

        cell = row.createCell(2);
        cell.setCellValue("A");
        cell.setCellStyle(cabecera);

        cell = row.createCell(3);
        cell.setCellValue("A");
        cell.setCellStyle(cabecera);

        cell = row.createCell(4);
        cell.setCellValue("A");
        cell.setCellStyle(cabecera);

        cell = row.createCell(5);
        cell.setCellStyle(cabecera);

        cell = row.createCell(6);
        cell.setCellStyle(cabecera);

        cell = row.createCell(7);
        cell.setCellStyle(cabecera);

        cell = row.createCell(8);
        cell.setCellStyle(cabecera);

        cell = row.createCell(9);
        cell.setCellStyle(cabecera);

        cell = row.createCell(10);
        cell.setCellStyle(cabecera);

        cell = row.createCell(11);
        cell.setCellStyle(cabecera);

        cell = row.createCell(12);
        cell.setCellStyle(cabecera);

        cell = row.createCell(13);
        cell.setCellStyle(cabecera);

        cell = row.createCell(14);
        cell.setCellStyle(cabecera);

        cell = row.createCell(15);
        cell.setCellStyle(cabecera);
        
        cell = row.createCell(16);
        cell.setCellStyle(cabecera);

        row = sh.createRow(rowCount++);

        cell = row.createCell(0);
        cell.setCellStyle(cabecera);

        cell = row.createCell(1);
        cell.setCellStyle(cabecera);

        cell = row.createCell(2);
        cell.setCellValue("12");
        cell.setCellStyle(cabecera);

        cell = row.createCell(3);
        cell.setCellValue("18");
        cell.setCellStyle(cabecera);

        cell = row.createCell(4);
        cell.setCellValue("24");
        cell.setCellStyle(cabecera);

        cell = row.createCell(5);
        cell.setCellStyle(cabecera);

        cell = row.createCell(6);
        cell.setCellStyle(cabecera);

        cell = row.createCell(7);
        cell.setCellStyle(cabecera);

        cell = row.createCell(8);
        cell.setCellStyle(cabecera);

        cell = row.createCell(9);
        cell.setCellStyle(cabecera);

        cell = row.createCell(10);
        cell.setCellStyle(cabecera);

        cell = row.createCell(11);
        cell.setCellStyle(cabecera);

        cell = row.createCell(12);
        cell.setCellStyle(cabecera);

        cell = row.createCell(13);
        cell.setCellStyle(cabecera);

        cell = row.createCell(14);
        cell.setCellStyle(cabecera);

        cell = row.createCell(15);
        cell.setCellStyle(cabecera);
        
         cell = row.createCell(16);
        cell.setCellStyle(cabecera);
        
        //Cuerpo
        for (HashMap fila : tablaProgra) {
            row = sh.createRow(rowCount);
            row.createCell(0).setCellValue((String) fila.get("CAMPANIA"));
            row.getCell(0).setCellStyle(datosLeft);

            row.createCell(1).setCellValue((String) fila.get("FECHA"));
            row.getCell(1).setCellStyle(datosCenter);

            row.createCell(2).setCellValue((String) fila.get("TURNO_6A12"));
            row.getCell(2).setCellStyle(datosCenter);

            row.createCell(3).setCellValue((String) fila.get("TURNO_12A18"));
            row.getCell(3).setCellStyle(datosCenter);

            row.createCell(4).setCellValue((String) fila.get("TURNO_18A24"));
            row.getCell(4).setCellStyle(datosCenter);

            row.createCell(5).setCellValue((String) fila.get("TOT_CLI_AFECTADOS"));
            row.getCell(5).setCellStyle(datosRight);

            row.createCell(6).setCellValue((String) fila.get("CLI_T1_AFECTADOS"));
            row.getCell(6).setCellStyle(datosRight);

            row.createCell(7).setCellValue((String) fila.get("CANT_LLAMADAS"));
            row.getCell(7).setCellStyle(datosRight);

            row.createCell(8).setCellValue((String) fila.get("EXITOSA"));
            row.getCell(8).setCellStyle(datosRight);

            row.createCell(9).setCellValue((String) fila.get("NO_CONTESTA"));
            row.getCell(9).setCellStyle(datosRight);

            row.createCell(10).setCellValue((String) fila.get("TEL_INCORRECTO"));
            row.getCell(10).setCellStyle(datosRight);

            row.createCell(11).setCellValue((String) fila.get("CLIENTES_T2_T3_AFECT"));
            row.getCell(11).setCellStyle(datosRight);

            row.createCell(12).setCellValue((String) fila.get("CANT_LLAMA_CAT"));
            row.getCell(12).setCellStyle(datosRight);

            row.createCell(13).setCellValue((String) fila.get("EXITOSA_CAT"));
            row.getCell(13).setCellStyle(datosRight);

            row.createCell(14).setCellValue((String) fila.get("NO_CONTESTA_CAT"));
            row.getCell(14).setCellStyle(datosRight);

            row.createCell(15).setCellValue((String) fila.get("TEL_INCORRECTO_CAT"));
            row.getCell(15).setCellStyle(datosRight);
            
            row.createCell(16).setCellValue((String) fila.get("F_INICIO"));
            row.getCell(16).setCellStyle(datosCenter);
            rowCount++;
        }
        
        
        //Ultima fila
        sh.addMergedRegion(new CellRangeAddress(rowCount, rowCount, 1, 4));
        row = sh.createRow(rowCount);
        row.setHeight((short) 800);
        row.createCell(0).setCellValue((String) ultimaProgra.get("TOT_SEM_LITERAL"));
        row.getCell(0).setCellStyle(totalesCenter);

        cell = row.createCell(1);
        cell.setCellStyle(totalesGrisado);

        cell = row.createCell(2);
        cell.setCellStyle(totalesGrisado);

        cell = row.createCell(3);
        cell.setCellStyle(totalesGrisado);

        cell = row.createCell(4);
        cell.setCellStyle(totalesGrisado);

        row.createCell(5).setCellValue((String) ultimaProgra.get("TOT_TOT_CLI_AFECTADOS"));
        row.getCell(5).setCellStyle(totalesRight);

        row.createCell(6).setCellValue((String) ultimaProgra.get("TOT_CLI_T1_AFECTADOS"));
        row.getCell(6).setCellStyle(totalesRight);

        row.createCell(7).setCellValue((String) ultimaProgra.get("TOT_CANT_LLAMADAS"));
        row.getCell(7).setCellStyle(totalesRight);

        row.createCell(8).setCellValue((String) ultimaProgra.get("TOT_EXITOSA"));
        row.getCell(8).setCellStyle(totalesRight);

        row.createCell(9).setCellValue((String) ultimaProgra.get("TOT_NO_CONTESTA"));
        row.getCell(9).setCellStyle(totalesRight);

        row.createCell(10).setCellValue((String) ultimaProgra.get("TOT_TEL_INCORRECTO"));
        row.getCell(10).setCellStyle(totalesRight);

        row.createCell(11).setCellValue((String) ultimaProgra.get("TOT_CLIENTES_T2_T3_AFECT"));
        row.getCell(11).setCellStyle(totalesRight);

        row.createCell(12).setCellValue((String) ultimaProgra.get("TOT_CANT_LLAMA_CAT"));
        row.getCell(12).setCellStyle(totalesRight);

        row.createCell(13).setCellValue((String) ultimaProgra.get("TOT_EXITOSA_CAT"));
        row.getCell(13).setCellStyle(totalesRight);

        row.createCell(14).setCellValue((String) ultimaProgra.get("TOT_NO_CONTESTA_CAT"));
        row.getCell(14).setCellStyle(totalesRight);

        row.createCell(15).setCellValue((String) ultimaProgra.get("TOT_TEL_INCORRECTO_CAT"));
        row.getCell(15).setCellStyle(totalesRight);
        
        cell = row.createCell(16);
        cell.setCellStyle(totalesGrisado);

        for (int i = 0; i < 5; i++) {
            sh.autoSizeColumn(i);
        }

    }

    private static void pueblaPuntuales(Sheet sh, Workbook wb) {

        Integer rowCount = 0;

        //Cabecera
        Row row = sh.createRow(rowCount++);
        sh.addMergedRegion(new CellRangeAddress(0, 0, 0, 19));

        sh.addMergedRegion(new CellRangeAddress(1, 3, 0, 0));
        sh.addMergedRegion(new CellRangeAddress(1, 3, 1, 1));
        sh.addMergedRegion(new CellRangeAddress(1, 3, 2, 2));
        sh.addMergedRegion(new CellRangeAddress(1, 1, 3, 10));
        sh.addMergedRegion(new CellRangeAddress(1, 3, 11, 11));
        sh.addMergedRegion(new CellRangeAddress(1, 1, 12, 17));
        sh.addMergedRegion(new CellRangeAddress(1, 3, 18, 18));
        sh.addMergedRegion(new CellRangeAddress(1, 3, 19, 19));

        sh.addMergedRegion(new CellRangeAddress(2, 3, 3, 3));
        sh.addMergedRegion(new CellRangeAddress(2, 3, 4, 4));
        sh.addMergedRegion(new CellRangeAddress(2, 2, 5, 7));
        sh.addMergedRegion(new CellRangeAddress(2, 3, 8, 8));
        sh.addMergedRegion(new CellRangeAddress(2, 3, 9, 9));
        sh.addMergedRegion(new CellRangeAddress(2, 3, 10, 10));
        sh.addMergedRegion(new CellRangeAddress(2, 3, 12, 12));
        sh.addMergedRegion(new CellRangeAddress(2, 3, 13, 13));
        sh.addMergedRegion(new CellRangeAddress(2, 2, 14, 16));
        sh.addMergedRegion(new CellRangeAddress(2, 3, 17, 17));
        
        CellStyle cabecera = creaEstilosCabe(wb);
        CellStyle datosLeft = creaEstilosDatos(wb, "left");
        CellStyle datosRight = creaEstilosDatos(wb, "right");
        CellStyle datosCenter = creaEstilosDatos(wb, "center");

        CellStyle totalesRight = creaEstilosTotales(wb, "right", "celeste");
        CellStyle totalesCenter = creaEstilosTotales(wb, "center", "celeste");
        CellStyle totalesGrisado = creaEstilosTotales(wb, "center", "grisado");

        Cell cell = row.createCell(0);
        cell.setCellValue("PUNTUALES BT");
        cell.setCellStyle(cabecera);

        for (int i = 1; i < 20; i++) {
            cell = row.createCell(i);
            cell.setCellStyle(cabecera);
        }
        
        row = sh.createRow(rowCount++);

        cell = row.createCell(0);
        cell.setCellValue("Nro de campaña");
        cell.setCellStyle(cabecera);

        cell = row.createCell(1);
        cell.setCellValue("Fecha y hora");
        cell.setCellStyle(cabecera);

        cell = row.createCell(2);
        cell.setCellValue("Cant. llam. (1)");
        cell.setCellStyle(cabecera);

        cell = row.createCell(3);
        cell.setCellValue("IVR");
        cell.setCellStyle(cabecera);

        cell = row.createCell(4);
        cell.setCellStyle(cabecera);

        cell = row.createCell(5);
        cell.setCellStyle(cabecera);

        cell = row.createCell(6);
        cell.setCellStyle(cabecera);

        cell = row.createCell(7);
        cell.setCellStyle(cabecera);

        cell = row.createCell(8);
        cell.setCellStyle(cabecera);

        cell = row.createCell(9);
        cell.setCellStyle(cabecera);

        cell = row.createCell(10);
        cell.setCellStyle(cabecera);

        cell = row.createCell(11);
        cell.setCellValue("Con Reitera.");
        cell.setCellStyle(cabecera);
        sh.setColumnWidth(11, 3000);

        cell = row.createCell(12);
        cell.setCellValue("CAT");
        cell.setCellStyle(cabecera);

        cell = row.createCell(13);
        cell.setCellStyle(cabecera);

        cell = row.createCell(14);
        cell.setCellStyle(cabecera);

        cell = row.createCell(15);
        cell.setCellStyle(cabecera);

        cell = row.createCell(16);
        cell.setCellStyle(cabecera);

        cell = row.createCell(17);
        cell.setCellStyle(cabecera);

        cell = row.createCell(18);
        cell.setCellValue("Total Cerrados (7)");
        cell.setCellStyle(cabecera);

        cell = row.createCell(19);
        cell.setCellValue("% Cerrados (8)");
        cell.setCellStyle(cabecera);

        row = sh.createRow(rowCount++);

        cell = row.createCell(0);
        cell.setCellStyle(cabecera);

        cell = row.createCell(1);
        cell.setCellStyle(cabecera);

        cell = row.createCell(2);
        cell.setCellStyle(cabecera);

        cell = row.createCell(3);
        cell.setCellValue("Con Luz (2)");
        cell.setCellStyle(cabecera);

        cell = row.createCell(4);
        cell.setCellValue("Sin Luz");
        cell.setCellStyle(cabecera);

        cell = row.createCell(5);
        cell.setCellValue("Sin Contactar");
        cell.setCellStyle(cabecera);

        cell = row.createCell(6);
        cell.setCellStyle(cabecera);

        cell = row.createCell(7);
        cell.setCellStyle(cabecera);

        cell = row.createCell(8);
        cell.setCellValue("Con Reclamo. Ant.");
        cell.setCellStyle(cabecera);
        sh.setColumnWidth(8, 3000);

        cell = row.createCell(9);
        cell.setCellValue("Sin Reitera. (3)");
        cell.setCellStyle(cabecera);
        sh.setColumnWidth(9, 3000);

        cell = row.createCell(10);
        cell.setCellValue("% Reitera. Prox.48 hr. (4)");
        cell.setCellStyle(cabecera);
        sh.setColumnWidth(10, 3000);

        cell = row.createCell(11);
        cell.setCellStyle(cabecera);

        cell = row.createCell(12);
        cell.setCellValue("Con Luz (5)");
        cell.setCellStyle(cabecera);

        cell = row.createCell(13);
        cell.setCellValue("Sin Luz");
        cell.setCellStyle(cabecera);

        cell = row.createCell(14);
        cell.setCellValue("Sin Contactar");
        cell.setCellStyle(cabecera);

        cell = row.createCell(15);
        cell.setCellStyle(cabecera);

        cell = row.createCell(16);
        cell.setCellStyle(cabecera);

        cell = row.createCell(17);
        cell.setCellValue("No gestionado por el CAT");
        cell.setCellStyle(cabecera);

        cell = row.createCell(18);
        cell.setCellStyle(cabecera);

        cell = row.createCell(19);
        cell.setCellStyle(cabecera);

        row = sh.createRow(rowCount++);

        cell = row.createCell(0);
        cell.setCellStyle(cabecera);

        cell = row.createCell(1);
        cell.setCellStyle(cabecera);

        cell = row.createCell(2);
        cell.setCellStyle(cabecera);

        cell = row.createCell(3);
        cell.setCellStyle(cabecera);

        cell = row.createCell(4);
        cell.setCellStyle(cabecera);

        cell = row.createCell(5);
        cell.setCellValue("No Contac.");
        cell.setCellStyle(cabecera);

        cell = row.createCell(6);
        cell.setCellValue("Nro Tel Erroneo");
        cell.setCellStyle(cabecera);

        cell = row.createCell(7);
        cell.setCellValue("Total");
        cell.setCellStyle(cabecera);

        cell = row.createCell(8);
        cell.setCellStyle(cabecera);

        cell = row.createCell(9);
        cell.setCellStyle(cabecera);

        cell = row.createCell(10);
        cell.setCellStyle(cabecera);

        cell = row.createCell(11);
        cell.setCellStyle(cabecera);

        cell = row.createCell(12);
        cell.setCellStyle(cabecera);

        cell = row.createCell(13);
        cell.setCellStyle(cabecera);

        cell = row.createCell(14);
        cell.setCellValue("No Contac.");
        cell.setCellStyle(cabecera);

        cell = row.createCell(15);
        cell.setCellValue("Nro Tel Erroneo");
        cell.setCellStyle(cabecera);

        cell = row.createCell(16);
        cell.setCellValue("Total (6)");
        cell.setCellStyle(cabecera);

        cell = row.createCell(17);
        cell.setCellStyle(cabecera);

        cell = row.createCell(18);
        cell.setCellStyle(cabecera);

        cell = row.createCell(19);
        cell.setCellStyle(cabecera);
        
        //Cuerpo
        for (HashMap fila : tablaPuntuales) {
            row = sh.createRow(rowCount);
            row.createCell(0).setCellValue((String) fila.get("CAMPANIA"));
            row.getCell(0).setCellStyle(datosLeft);

            row.createCell(1).setCellValue((String) fila.get("FECHA"));
            row.getCell(1).setCellStyle(datosCenter);

            row.createCell(2).setCellValue((String) fila.get("CANT_LLAMADOS"));
            row.getCell(2).setCellStyle(datosRight);

            row.createCell(3).setCellValue((String) fila.get("CON_LUZ_IVR"));
            row.getCell(3).setCellStyle(datosRight);

            row.createCell(4).setCellValue((String) fila.get("SIN_LUZ_IVR"));
            row.getCell(4).setCellStyle(datosRight);

            row.createCell(5).setCellValue((String) fila.get("NOCONTAC_IVR"));
            row.getCell(5).setCellStyle(datosRight);

            row.createCell(6).setCellValue((String) fila.get("ERRORTEL_IVR"));
            row.getCell(6).setCellStyle(datosRight);

            row.createCell(7).setCellValue((String) fila.get("NOCONTACTOT_IVR"));
            row.getCell(7).setCellStyle(datosRight);

            row.createCell(8).setCellValue((String) fila.get("CONRECLANT_IVR"));
            row.getCell(8).setCellStyle(datosRight);

            row.createCell(9).setCellValue((String) fila.get("SINREITE_IVR"));
            row.getCell(9).setCellStyle(datosRight);

            row.createCell(10).setCellValue((String) fila.get("PORC_REITERA"));
            row.getCell(10).setCellStyle(datosRight);

            row.createCell(11).setCellValue((String) fila.get("CON_REITERA"));
            row.getCell(11).setCellStyle(datosRight);

            row.createCell(12).setCellValue((String) fila.get("CONLUZ_CALL"));
            row.getCell(12).setCellStyle(datosRight);

            row.createCell(13).setCellValue((String) fila.get("SINLUZ_CALL"));
            row.getCell(13).setCellStyle(datosRight);

            row.createCell(14).setCellValue((String) fila.get("NOCONTAC_CALL"));
            row.getCell(14).setCellStyle(datosRight);

            row.createCell(15).setCellValue((String) fila.get("ERRORTEL_CALL"));
            row.getCell(15).setCellStyle(datosRight);

            row.createCell(16).setCellValue((String) fila.get("NOCONTACTOT_CALL"));
            row.getCell(16).setCellStyle(datosRight);

            row.createCell(17).setCellValue((String) fila.get("NOGEST_CALL"));
            row.getCell(17).setCellStyle(datosRight);

            row.createCell(18).setCellValue((String) fila.get("TOTAL_CLOSE"));
            row.getCell(18).setCellStyle(datosRight);

            row.createCell(19).setCellValue((String) fila.get("POR_TOTAL_CERR"));
            row.getCell(19).setCellStyle(datosRight);

            rowCount++;
        }
        
        //Totales
        row = sh.createRow(rowCount++);
        row.setHeight((short) 800);
        row.createCell(0).setCellValue((String) ultimaPuntuales.get("TOT_SEM_LITERAL"));
        row.getCell(0).setCellStyle(totalesCenter);

        cell = row.createCell(1);
        cell.setCellStyle(totalesGrisado);

        row.createCell(2).setCellValue((String) ultimaPuntuales.get("TOT_CANT_LLAMADOS"));
        row.getCell(2).setCellStyle(totalesCenter);

        row.createCell(3).setCellValue((String) ultimaPuntuales.get("TOT_CON_LUZIVR"));
        row.getCell(3).setCellStyle(totalesCenter);

        row.createCell(4).setCellValue((String) ultimaPuntuales.get("TOT_SIN_LUZIVR"));
        row.getCell(4).setCellStyle(totalesCenter);

        row.createCell(5).setCellValue((String) ultimaPuntuales.get("TOT_NOCONTACIVR"));
        row.getCell(5).setCellStyle(totalesCenter);

        row.createCell(6).setCellValue((String) ultimaPuntuales.get("TOT_ERRORTELIVR"));
        row.getCell(6).setCellStyle(totalesCenter);

        row.createCell(7).setCellValue((String) ultimaPuntuales.get("TOT_NOCONTACTOTIVR"));
        row.getCell(7).setCellStyle(totalesCenter);

        row.createCell(8).setCellValue((String) ultimaPuntuales.get("TOT_CONRECLANTIRV"));
        row.getCell(8).setCellStyle(totalesRight);

        row.createCell(9).setCellValue((String) ultimaPuntuales.get("TOT_SINREITEIVR"));
        row.getCell(9).setCellStyle(totalesRight);

        row.createCell(10).setCellValue((String) ultimaPuntuales.get("TOT_PORC_REITERA"));
        row.getCell(10).setCellStyle(totalesRight);

        row.createCell(11).setCellValue((String) ultimaPuntuales.get("TOT_CON_REITERA"));
        row.getCell(11).setCellStyle(totalesRight);

        row.createCell(12).setCellValue((String) ultimaPuntuales.get("TOT_CON_LUZCALL"));
        row.getCell(12).setCellStyle(totalesRight);

        row.createCell(13).setCellValue((String) ultimaPuntuales.get("TOT_SIN_LUZCALL"));
        row.getCell(13).setCellStyle(totalesRight);

        row.createCell(14).setCellValue((String) ultimaPuntuales.get("TOT_NO_CONTACALL"));
        row.getCell(14).setCellStyle(totalesRight);

        row.createCell(15).setCellValue((String) ultimaPuntuales.get("TOT_ERROR_TELCALL"));
        row.getCell(15).setCellStyle(totalesRight);

        row.createCell(16).setCellValue((String) ultimaPuntuales.get("TOT_NOCONTACTOTCALL"));
        row.getCell(16).setCellStyle(totalesRight);

        row.createCell(17).setCellValue((String) ultimaPuntuales.get("TOT_NO_GESTCALL"));
        row.getCell(17).setCellStyle(totalesRight);

        row.createCell(18).setCellValue((String) ultimaPuntuales.get("TOT_CLOSE"));
        row.getCell(18).setCellStyle(totalesRight);

        row.createCell(19).setCellValue((String) ultimaPuntuales.get("TOT_POR_TOTALCLOSE"));
        row.getCell(19).setCellStyle(totalesRight);
        
        for (int i = 0; i < 5; i++) {
            sh.autoSizeColumn(i);
        }
        
        //NOTAS
        String msj1 = "1. La cantidad de Llamados aplica solo para documentos con un solo reclamo";
        String msj2 = "4. El porcentaje aplica sobre la cantidad de clientes que reclaman en las proximas 48 hs posteriores del cierre de (3+6). ";
        String msj3 = "7. Resultado de 2+3+5+6. ";
        String msj4 = "8. Resultado de 7/1. ";
       
        sh.createRow(rowCount++);
        sh.createRow(rowCount++);
        
        sh.addMergedRegion(new CellRangeAddress(rowCount, rowCount, 0, 19));
        row = sh.createRow(rowCount++);
        row.createCell(0).setCellValue(msj1);
        
        sh.addMergedRegion(new CellRangeAddress(rowCount, rowCount, 0, 19));
        row = sh.createRow(rowCount++);
        row.createCell(0).setCellValue(msj2);
        
        sh.addMergedRegion(new CellRangeAddress(rowCount, rowCount, 0, 19));
        row = sh.createRow(rowCount++);
        row.createCell(0).setCellValue(msj3);
        
        sh.addMergedRegion(new CellRangeAddress(rowCount, rowCount, 0, 19));
        row = sh.createRow(rowCount++);
        row.createCell(0).setCellValue(msj4);

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
            fila.put("CAMPANIA", campo[0]);
            fila.put("FECHA", campo[1]);
            fila.put("TURNO_6A12", campo[2]);
            fila.put("TURNO_12A18", campo[3]);
            fila.put("TURNO_18A24", campo[4]);
            fila.put("CLI_AFECTADOS", campo[5]);
            fila.put("CANT_LLAMADAS", campo[6]);
            fila.put("EXITOSA", campo[7]);
            fila.put("NO_CONTESTA", campo[8]);
            fila.put("TEL_INCORRECTO", campo[9]);
            fila.put("SUSPENDIDOS", campo[10]);
            tabla.add(fila);
        }

        return tabla;
    }

    static HashMap<String, String> strToHashUltLineaForzados(String reporte) {

        HashMap<String, String> fila = null;
        String[] lineas = reporte.split(";", -1);
        for (int i = 0; i < lineas.length; i++) {
            fila = new HashMap<>();
            fila.put("TOT_SEM_LITERAL", lineas[i++]);
            i += 4; //siempre vienen vacios
            fila.put("TOT_CLI_AFECTADOS", lineas[i++]);
            fila.put("TOT_CANT_LLAMADOS", lineas[i++]);
            fila.put("TOT_EXITOSA", lineas[i++]);
            fila.put("TOT_NO_CONTESTA", lineas[i++]);
            fila.put("TOT_TEL_ERROR", lineas[i++]);
            fila.put("TOT_SUSPEN", lineas[i]);
        }
        return fila;
    }

    static List<HashMap<String, String>> strToHashProgramados(String reporte) {

        List<HashMap<String, String>> tabla = new ArrayList<>();
        reporte = reporte.replace("|", "X");

        String[] lineas = reporte.split("X");
        for (String linea : lineas) {
            String[] campo = linea.split(";");

            HashMap<String, String> fila = new HashMap<>();
            fila.put("CAMPANIA", campo[0]);
            fila.put("FECHA", campo[1]);
            fila.put("TURNO_6A12", campo[2]);
            fila.put("TURNO_12A18", campo[3]);
            fila.put("TURNO_18A24", campo[4]);
            fila.put("TOT_CLI_AFECTADOS", campo[5]);
            fila.put("CLI_T1_AFECTADOS", campo[6]);
            fila.put("CANT_LLAMADAS", campo[7]);
            fila.put("EXITOSA", campo[8]);
            fila.put("NO_CONTESTA", campo[9]);
            fila.put("TEL_INCORRECTO", campo[10]);
            fila.put("CLIENTES_T2_T3_AFECT", campo[11]);
            fila.put("CANT_LLAMA_CAT", campo[12]);
            fila.put("EXITOSA_CAT", campo[13]);
            fila.put("NO_CONTESTA_CAT", campo[14]);
            fila.put("TEL_INCORRECTO_CAT", campo[15]);
            fila.put("F_INICIO", campo[16]);
            
            tabla.add(fila);
        }
        return tabla;
    }

    static HashMap<String, String> strToHashUltLineaProgramados(String reporte) {

        HashMap<String, String> fila = null;
        String[] lineas = reporte.split(";", -1);
        for (int i = 0; i < lineas.length; i++) {
            fila = new HashMap<>();
            fila.put("TOT_SEM_LITERAL", lineas[i++]);
            i += 4; //siempre vienen vacios
            fila.put("TOT_TOT_CLI_AFECTADOS", lineas[i++]);
            fila.put("TOT_CLI_T1_AFECTADOS", lineas[i++]);
            fila.put("TOT_CANT_LLAMADAS", lineas[i++]);
            fila.put("TOT_EXITOSA", lineas[i++]);
            fila.put("TOT_NO_CONTESTA", lineas[i++]);
            fila.put("TOT_TEL_INCORRECTO", lineas[i++]);
            fila.put("TOT_CLIENTES_T2_T3_AFECT", lineas[i++]);
            fila.put("TOT_CANT_LLAMA_CAT", lineas[i++]);
            fila.put("TOT_EXITOSA_CAT", lineas[i++]);
            fila.put("TOT_NO_CONTESTA_CAT", lineas[i++]);
            fila.put("TOT_TEL_INCORRECTO_CAT", lineas[i]);
        }
        return fila;
    }
    
    static List<HashMap<String, String>> strToHashPuntuales(String reporte) {

        List<HashMap<String, String>> tabla = new ArrayList<>();
        reporte = reporte.substring(0, reporte.length() - 2);

        String[] lineas = reporte.split("n");
        for (String linea : lineas) {
            String[] campo = linea.split(";");
            String total;
            total = campo[19].substring(0, campo[19].length() - 1);
            int prim = total.indexOf(".");
            if (prim > 0) {
                total = total.substring(0, prim);
            }
            HashMap<String, String> fila = new HashMap<>();
            fila.put("CAMPANIA", campo[0]);
            fila.put("FECHA", campo[1]);
            fila.put("CANT_LLAMADOS", campo[2]);
            fila.put("CON_LUZ_IVR", campo[3]);
            fila.put("SIN_LUZ_IVR", campo[4]);
            fila.put("NOCONTAC_IVR", campo[5]);
            fila.put("ERRORTEL_IVR", campo[6]);
            fila.put("NOCONTACTOT_IVR", campo[7]);
            fila.put("CONRECLANT_IVR", campo[8]);
            fila.put("SINREITE_IVR", campo[9]);
            fila.put("PORC_REITERA", campo[10]);
            fila.put("CON_REITERA", campo[11]);
            fila.put("CONLUZ_CALL", campo[12]);
            fila.put("SINLUZ_CALL", campo[13]);
            fila.put("NOCONTAC_CALL", campo[14]);
            fila.put("ERRORTEL_CALL", campo[15]);
            fila.put("NOCONTACTOT_CALL", campo[16]);
            fila.put("NOGEST_CALL", campo[17]);
            fila.put("TOTAL_CLOSE", campo[18]);
            fila.put("POR_TOTAL_CERR", total);
            tabla.add(fila);
        }
        return tabla;
    }

    static HashMap<String, String> strToHashUltLineaPuntuales(String reporte) {

        HashMap<String, String> fila = null;
        String[] lineas = reporte.split(";", -1);
        for (int i = 0; i < lineas.length; i++) {
            fila = new HashMap<>();
            fila.put("TOT_SEM_LITERAL", lineas[i++]);
            i++; //FECHA
            fila.put("TOT_CANT_LLAMADOS", lineas[i++]);
            fila.put("TOT_CON_LUZIVR", lineas[i++]);
            fila.put("TOT_SIN_LUZIVR", lineas[i++]);
            fila.put("TOT_NOCONTACIVR", lineas[i++]);
            fila.put("TOT_ERRORTELIVR", lineas[i++]);
            fila.put("TOT_NOCONTACTOTIVR", lineas[i++]);
            fila.put("TOT_CONRECLANTIRV", lineas[i++]);
            fila.put("TOT_SINREITEIVR", lineas[i++]);
            fila.put("TOT_PORC_REITERA", lineas[i++]);
            fila.put("TOT_CON_REITERA", lineas[i++]);
            fila.put("TOT_CON_LUZCALL", lineas[i++]);
            fila.put("TOT_SIN_LUZCALL", lineas[i++]);
            fila.put("TOT_NO_CONTACALL", lineas[i++]);
            fila.put("TOT_ERROR_TELCALL", lineas[i++]);
            fila.put("TOT_NOCONTACTOTCALL", lineas[i++]);
            fila.put("TOT_NO_GESTCALL", lineas[i++]);
            fila.put("TOT_CLOSE", lineas[i++]);
            fila.put("TOT_POR_TOTALCLOSE", lineas[i]);

        }
        return fila;
    }

}
