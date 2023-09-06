import java.awt.Color;
import java.io.BufferedReader;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
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
import java.util.Hashtable;
import java.util.Iterator;
import java.util.List;
import java.util.Map;
import java.util.Properties;
import java.util.PropertyResourceBundle;
import java.util.Vector;

import javax.activation.DataHandler;
import javax.activation.FileDataSource;
import javax.mail.Address;
import javax.mail.Authenticator;
import javax.mail.BodyPart;
import javax.mail.Message.RecipientType;
import javax.mail.MessagingException;
import javax.mail.SendFailedException;
import javax.mail.Session;
import javax.mail.Transport;
import javax.mail.internet.InternetAddress;
import javax.mail.internet.MimeBodyPart;
import javax.mail.internet.MimeMessage;
import javax.mail.internet.MimeMultipart;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.FontUnderline;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xssf.usermodel.extensions.XSSFCellBorder.BorderSide;

import java.io.FileInputStream;
import java.io.IOException;
import java.util.Properties;

final class ConsolidadoBTMT {
	private static String destinatarios;
	String msgCancela = "";
	Vector dir_no_encontradas = new Vector();
	Connection con_1;
	static ResultSet rs1;
	static ResultSet rs2;
	static ResultSet rs3;
	static ResultSet rs4;
	static String[] arraybarrio = new String[40];
	String tablahtml = "";
	static int[] cant_clientes = new int[40];
	static int[] afectbt = new int[40];
	static int[] cli_afectbt = new int[40];
	static int[] cli_afectmt = new int[40];
	static int[] afectmt = new int[40];
	String Dia = null;
	String Fecha = null;
	String Hora = null;
	List<HashMap<String, String>> tabla2 = null;
	static String fuente_html_fin = "</table> <br><table class='tlogo'> <tr>  <td>  <img width=450 height=145 src='logo.png'  alt='Descripci�n: logo de Edenor' >  </td> </tr></table></body> </html>";
	final String driverClass = "oracle.jdbc.driver.OracleDriver";
	private Iterator<HashMap<String, String>> it;
        
        private static Properties properties;
        private static String excel;

	public void Oracle_connect() throws Exception {
            try {
                Class.forName(properties.getProperty("driverClass")).newInstance();         //"oracle.jdbc.driver.OracleDriver"
            } catch (InstantiationException | IllegalAccessException | ClassNotFoundException var2) {
                System.out.println("Error al cargar el driver: oracle.jdbc.driver.OracleDriver -Error: " + var2);
                throw var2;
            }

            try {
                this.con_1 = DriverManager.getConnection(
                        properties.getProperty("oracle_cnn_pro"),             //jdbc:oracle:thin:@ltronxgisbdpr03.pro.edenor:1528:GISPR03
                        properties.getProperty("oracle_cnn_usu"),             //SVC_ORA_GIS
                        properties.getProperty("oracle_cnn_pss"));            //jv506uzy
            } catch (SQLException var3) {
                switch (var3.getErrorCode()) {
                case 1017:
                        System.out.println("Usuario de la base de datos incorrecto");
                        break;
                case 2391:
                        System.out.println("LÃ�\u00admite de conexiones excedido, intentar mas tarde");
                        break;
                case 28000:
                        System.out.println("Cuenta bloqueada");
                        break;
                case 28001:
                        System.out.println("Clave de la base de datos expirada");
                        break;
                default:
                        System.out.println("Error al conectarse: " + var3);
                }
                throw var3;
            }
	}

	static ResultSet Fechainforme(Connection conexion) throws Exception {
		Connection con = conexion;
		ResultSet rs = null;

		try {
			String query_variable = "select to_char (sysdate, 'dd/mm/yyyy HH24:MI:SS') as FECHA from dual";
			Statement sentencia = con.createStatement();
			rs = sentencia.executeQuery(query_variable);
			return rs;
		} catch (SQLException var5) {
			throw var5;
		}
	}

	private static CellStyle creaEstilostitulo(Workbook wb) {
		XSSFCellStyle style = (XSSFCellStyle) wb.createCellStyle();
		XSSFColor blanco = new XSSFColor(Color.decode("#FFFFFF"));
		XSSFFont negrita = (XSSFFont) wb.createFont();
		negrita.setBold(true);
		style.setFont(negrita);
		negrita.setItalic(true);
		negrita.setUnderline(FontUnderline.SINGLE);
		style.setBorderLeft((short) 0);
		style.setBorderRight((short) 0);
		style.setBorderBottom((short) 0);
		style.setBorderTop((short) 0);
		style.setBorderColor(BorderSide.LEFT, blanco);
		style.setBorderColor(BorderSide.TOP, blanco);
		style.setBorderColor(BorderSide.RIGHT, blanco);
		style.setBorderColor(BorderSide.BOTTOM, blanco);
		return style;
	}

	private static CellStyle creaEstilosnoborder(Workbook wb) {
		XSSFCellStyle noborder = (XSSFCellStyle) wb.createCellStyle();
		XSSFColor blanco = new XSSFColor(Color.decode("#FFFFFF"));
		noborder.setBorderLeft((short) 0);
		noborder.setBorderRight((short) 0);
		noborder.setBorderBottom((short) 0);
		noborder.setBorderTop((short) 0);
		noborder.setBorderColor(BorderSide.LEFT, blanco);
		noborder.setBorderColor(BorderSide.TOP, blanco);
		noborder.setBorderColor(BorderSide.RIGHT, blanco);
		noborder.setBorderColor(BorderSide.BOTTOM, blanco);
		noborder.setBorderRight((short) 1);
		noborder.setRightBorderColor(IndexedColors.WHITE.getIndex());
		noborder.setBorderBottom((short) 1);
		noborder.setBottomBorderColor(IndexedColors.WHITE.getIndex());
		noborder.setBorderLeft((short) 1);
		noborder.setLeftBorderColor(IndexedColors.WHITE.getIndex());
		noborder.setBorderTop((short) 1);
		noborder.setTopBorderColor(IndexedColors.WHITE.getIndex());
		noborder.setFillForegroundColor((short) 9);
		noborder.setFillPattern(FillPatternType.SOLID_FOREGROUND);
		return noborder;
	}

	private static CellStyle creaEstilosCabe(Workbook wb) {
		XSSFCellStyle style = (XSSFCellStyle) wb.createCellStyle();
		XSSFFont negrita = (XSSFFont) wb.createFont();
		negrita.setBoldweight((short) 700);
		negrita.setFontHeightInPoints((short) 9);
		style.setFont(negrita);
		XSSFColor myColor = new XSSFColor(Color.decode("#C5D9F1"));
		style.setFillForegroundColor(myColor);
		style.setAlignment(HorizontalAlignment.RIGHT);
		style.setFillPattern((short) 1);
		style.setAlignment((short) 2);
		style.setVerticalAlignment((short) 1);
		style.setWrapText(true);
		XSSFColor negro = new XSSFColor(Color.decode("#000000"));
		style.setBorderBottom((short) 2);
		style.setBorderTop((short) 2);
		style.setBorderRight((short) 2);
		style.setBorderLeft((short) 2);
		style.setBorderColor(BorderSide.LEFT, negro);
		style.setBorderColor(BorderSide.TOP, negro);
		style.setBorderColor(BorderSide.RIGHT, negro);
		style.setBorderColor(BorderSide.BOTTOM, negro);
		return style;
	}

	private static CellStyle creaEstilosCabegreen(Workbook wb) {
		XSSFCellStyle style = (XSSFCellStyle) wb.createCellStyle();
		XSSFFont negrita = (XSSFFont) wb.createFont();
		negrita.setBoldweight((short) 700);
		negrita.setFontHeightInPoints((short) 9);
		style.setFont(negrita);
		Font font = wb.createFont();
		font.setColor((short) 17);
		style.setFont(font);
		XSSFColor myColor = new XSSFColor(Color.decode("#C5D9F1"));
		style.setFillForegroundColor(myColor);
		style.setAlignment(HorizontalAlignment.RIGHT);
		style.setFillPattern((short) 1);
		style.setAlignment((short) 2);
		style.setVerticalAlignment((short) 1);
		style.setWrapText(true);
		XSSFColor negro = new XSSFColor(Color.decode("#000000"));
		style.setBorderBottom((short) 2);
		style.setBorderTop((short) 2);
		style.setBorderRight((short) 2);
		style.setBorderLeft((short) 2);
		style.setBorderColor(BorderSide.LEFT, negro);
		style.setBorderColor(BorderSide.TOP, negro);
		style.setBorderColor(BorderSide.RIGHT, negro);
		style.setBorderColor(BorderSide.BOTTOM, negro);
		return style;
	}

	private static CellStyle creaEstilosCabered(Workbook wb) {
		XSSFCellStyle style = (XSSFCellStyle) wb.createCellStyle();
		XSSFFont negrita = (XSSFFont) wb.createFont();
		negrita.setBoldweight((short) 700);
		negrita.setFontHeightInPoints((short) 9);
		style.setFont(negrita);
		Font font = wb.createFont();
		font.setColor((short) 10);
		style.setFont(font);
		XSSFColor myColor = new XSSFColor(Color.decode("#C5D9F1"));
		style.setFillForegroundColor(myColor);
		style.setAlignment(HorizontalAlignment.RIGHT);
		style.setFillPattern((short) 1);
		style.setAlignment((short) 2);
		style.setVerticalAlignment((short) 1);
		style.setWrapText(true);
		XSSFColor negro = new XSSFColor(Color.decode("#000000"));
		style.setBorderBottom((short) 2);
		style.setBorderTop((short) 2);
		style.setBorderRight((short) 2);
		style.setBorderLeft((short) 2);
		style.setBorderColor(BorderSide.LEFT, negro);
		style.setBorderColor(BorderSide.TOP, negro);
		style.setBorderColor(BorderSide.RIGHT, negro);
		style.setBorderColor(BorderSide.BOTTOM, negro);
		return style;
	}

	private static CellStyle creaEstilosCabeyellow(Workbook wb) {
		XSSFCellStyle style = (XSSFCellStyle) wb.createCellStyle();
		XSSFFont negrita = (XSSFFont) wb.createFont();
		negrita.setBoldweight((short) 700);
		negrita.setFontHeightInPoints((short) 9);
		style.setFont(negrita);
		Font font = wb.createFont();
		font.setColor((short) 13);
		style.setFont(font);
		XSSFColor myColor = new XSSFColor(Color.decode("#C5D9F1"));
		style.setFillForegroundColor(myColor);
		style.setAlignment(HorizontalAlignment.RIGHT);
		style.setFillPattern((short) 1);
		style.setAlignment((short) 2);
		style.setVerticalAlignment((short) 1);
		style.setWrapText(true);
		XSSFColor negro = new XSSFColor(Color.decode("#000000"));
		style.setBorderBottom((short) 2);
		style.setBorderTop((short) 2);
		style.setBorderRight((short) 2);
		style.setBorderLeft((short) 2);
		style.setBorderColor(BorderSide.LEFT, negro);
		style.setBorderColor(BorderSide.TOP, negro);
		style.setBorderColor(BorderSide.RIGHT, negro);
		style.setBorderColor(BorderSide.BOTTOM, negro);
		return style;
	}

	private static CellStyle creaEstilosDatos(Workbook wb, String alignment) {
		XSSFCellStyle style = (XSSFCellStyle) wb.createCellStyle();
		byte var4 = -1;
		switch (alignment.hashCode()) {
		case -1364013995:
			if (alignment.equals("center")) {
				var4 = 1;
			}
			break;
		case 3317767:
			if (alignment.equals("left")) {
				var4 = 2;
			}
			break;
		case 108511772:
			if (alignment.equals("right")) {
				var4 = 0;
			}
		}

		switch (var4) {
		case 0:
			style.setAlignment(HorizontalAlignment.RIGHT);
			break;
		case 1:
			style.setAlignment(HorizontalAlignment.CENTER);
			break;
		case 2:
			style.setAlignment(HorizontalAlignment.LEFT);
		}

		style.setVerticalAlignment((short) 1);
		style.setWrapText(true);
		XSSFColor negro = new XSSFColor(Color.decode("#000000"));
		style.setBorderBottom((short) 1);
		style.setBorderTop((short) 1);
		style.setBorderRight((short) 1);
		style.setBorderLeft((short) 1);
		style.setBorderColor(BorderSide.LEFT, negro);
		style.setBorderColor(BorderSide.TOP, negro);
		style.setBorderColor(BorderSide.RIGHT, negro);
		style.setBorderColor(BorderSide.BOTTOM, negro);
		return style;
	}

	private static CellStyle creaEstilosbordesgrueso(Workbook wb) {
		XSSFCellStyle style = (XSSFCellStyle) wb.createCellStyle();
		style.setAlignment(HorizontalAlignment.LEFT);
		style.setVerticalAlignment((short) 1);
		style.setWrapText(true);
		XSSFColor negro = new XSSFColor(Color.decode("#000000"));
		style.setBorderBottom((short) 1);
		style.setBorderTop((short) 1);
		style.setBorderRight((short) 2);
		style.setBorderLeft((short) 2);
		style.setBorderColor(BorderSide.LEFT, negro);
		style.setBorderColor(BorderSide.TOP, negro);
		style.setBorderColor(BorderSide.RIGHT, negro);
		style.setBorderColor(BorderSide.BOTTOM, negro);
		return style;
	}

	private static CellStyle creaEstilosColorRed(Workbook wb) {
		XSSFCellStyle style = (XSSFCellStyle) wb.createCellStyle();
		style.setAlignment(HorizontalAlignment.CENTER);
		XSSFFont negrita = (XSSFFont) wb.createFont();
		negrita.setBoldweight((short) 700);
		negrita.setFontHeightInPoints((short) 9);
		style.setFont(negrita);
		XSSFColor myColor = new XSSFColor(Color.decode("#FF0000"));
		style.setFillForegroundColor(myColor);
		style.setAlignment(HorizontalAlignment.RIGHT);
		style.setFillPattern((short) 1);
		style.setAlignment((short) 2);
		style.setVerticalAlignment((short) 1);
		style.setWrapText(true);
		XSSFColor negro = new XSSFColor(Color.decode("#000000"));
		style.setBorderBottom((short) 2);
		style.setBorderTop((short) 2);
		style.setBorderRight((short) 2);
		style.setBorderLeft((short) 2);
		style.setBorderColor(BorderSide.LEFT, negro);
		style.setBorderColor(BorderSide.TOP, negro);
		style.setBorderColor(BorderSide.RIGHT, negro);
		style.setBorderColor(BorderSide.BOTTOM, negro);
		return style;
	}

	private static CellStyle creaEstilosColorYellow(Workbook wb) {
		XSSFCellStyle style = (XSSFCellStyle) wb.createCellStyle();
		style.setAlignment(HorizontalAlignment.CENTER);
		XSSFFont negrita = (XSSFFont) wb.createFont();
		negrita.setBoldweight((short) 700);
		negrita.setFontHeightInPoints((short) 9);
		style.setFont(negrita);
		XSSFColor myColor = new XSSFColor(Color.decode("#FFFF00"));
		style.setFillForegroundColor(myColor);
		style.setAlignment(HorizontalAlignment.RIGHT);
		style.setFillPattern((short) 1);
		style.setAlignment((short) 2);
		style.setVerticalAlignment((short) 1);
		style.setWrapText(true);
		XSSFColor negro = new XSSFColor(Color.decode("#000000"));
		style.setBorderBottom((short) 2);
		style.setBorderTop((short) 2);
		style.setBorderRight((short) 2);
		style.setBorderLeft((short) 2);
		style.setBorderColor(BorderSide.LEFT, negro);
		style.setBorderColor(BorderSide.TOP, negro);
		style.setBorderColor(BorderSide.RIGHT, negro);
		style.setBorderColor(BorderSide.BOTTOM, negro);
		return style;
	}

	private static CellStyle creaEstilosColorGreen(Workbook wb) {
		XSSFCellStyle style = (XSSFCellStyle) wb.createCellStyle();
		style.setAlignment(HorizontalAlignment.CENTER);
		XSSFFont negrita = (XSSFFont) wb.createFont();
		negrita.setBoldweight((short) 700);
		negrita.setFontHeightInPoints((short) 9);
		style.setFont(negrita);
		XSSFColor myColor = new XSSFColor(Color.decode("#00FF00"));
		style.setFillForegroundColor(myColor);
		style.setAlignment(HorizontalAlignment.RIGHT);
		style.setFillPattern((short) 1);
		style.setAlignment((short) 2);
		style.setVerticalAlignment((short) 1);
		style.setWrapText(true);
		XSSFColor negro = new XSSFColor(Color.decode("#000000"));
		style.setBorderBottom((short) 2);
		style.setBorderTop((short) 2);
		style.setBorderRight((short) 2);
		style.setBorderLeft((short) 2);
		style.setBorderColor(BorderSide.LEFT, negro);
		style.setBorderColor(BorderSide.TOP, negro);
		style.setBorderColor(BorderSide.RIGHT, negro);
		style.setBorderColor(BorderSide.BOTTOM, negro);
		return style;
	}

	private static CellStyle creaEstilosTotales(Workbook wb, String alignment, String color) {
		XSSFCellStyle style = (XSSFCellStyle) wb.createCellStyle();
		XSSFFont negrita = (XSSFFont) wb.createFont();
		negrita.setBoldweight((short) 700);
		style.setFont(negrita);
		byte var7 = -1;
		switch (color.hashCode()) {
		case 287813943:
			if (color.equals("grisado")) {
				var7 = 0;
			}
			break;
		case 662958825:
			if (color.equals("celeste")) {
				var7 = 1;
			}
		}

		String codigo;
		switch (var7) {
		case 0:
			codigo = "#B8BDBF";
			break;
		case 1:
			codigo = "#C5D9F1";
			break;
		default:
			codigo = "#C5D9F1";
		}

		XSSFColor myColor = new XSSFColor(Color.decode(codigo));
		style.setFillForegroundColor(myColor);
		style.setFillPattern((short) 1);
		byte var8 = -1;
		switch (alignment.hashCode()) {
		case -1364013995:
			if (alignment.equals("center")) {
				var8 = 1;
			}
			break;
		case 3317767:
			if (alignment.equals("left")) {
				var8 = 2;
			}
			break;
		case 108511772:
			if (alignment.equals("right")) {
				var8 = 0;
			}
		}

		switch (var8) {
		case 0:
			style.setAlignment(HorizontalAlignment.RIGHT);
			break;
		case 1:
			style.setAlignment(HorizontalAlignment.CENTER);
			break;
		case 2:
			style.setAlignment(HorizontalAlignment.LEFT);
		}

		style.setVerticalAlignment((short) 1);
		style.setWrapText(true);
		XSSFColor negro = new XSSFColor(Color.decode("#000000"));
		style.setBorderBottom((short) 2);
		style.setBorderTop((short) 2);
		style.setBorderRight((short) 2);
		style.setBorderLeft((short) 2);
		style.setBorderColor(BorderSide.LEFT, negro);
		style.setBorderColor(BorderSide.TOP, negro);
		style.setBorderColor(BorderSide.RIGHT, negro);
		style.setBorderColor(BorderSide.BOTTOM, negro);
		return style;
	}

	private List<HashMap<String, String>> strToHash2(String reporte) {
		List<HashMap<String, String>> tabla = new ArrayList();
		String[] lineas = reporte.split(",");
		String[] var5 = lineas;
		int var6 = lineas.length;

		for (int var7 = 0; var7 < var6; ++var7) {
			String linea = var5[var7];
			HashMap<String, String> fila = new HashMap();
			String[] campo = linea.split(";");
			fila.put("REGION", campo[0]);
			fila.put("ZONA", campo[1]);
			fila.put("BT", campo[2]);
			fila.put("MT", campo[3]);
			fila.put("TOTAL", campo[4]);
			tabla.add(fila);
		}

		return tabla;
	}

	public ResultSet get_clients() throws Exception {
		Object var2 = null;

		ResultSet var5;
		try {
			String plsql1 = "SELECT * from NEXUS_GIS.CANT_CLIENTES";
			System.out.println("Cantidad de clientes es " + plsql1);
			Statement stmt = this.con_1.createStatement(1004, 1007);
			ResultSet rs1 = stmt.executeQuery(plsql1);
			var5 = rs1;
		} catch (Exception var14) {
			this.con_1.close();
			System.out.println("Error en raster_bt()" + var14);
			throw var14;
		} finally {
			try {
				System.out.println("Final One");
			} catch (Exception var13) {
			}

		}

		return var5;
	}

	public ResultSet raster_mtat_loc() throws Exception {
		Object var2 = null;

		ResultSet var5;
		try {
			this.Oracle_connect();
			String plsql1 = "SELECT zona, COUNT (zona) AS CANT_CORTES, SUM (cli_actual) AS CLI_AFECT\r\n" + 
					"    FROM ( (SELECT DISTINCT (tdet.NRO_DOCUMENTO),\r\n" + 
					"                            --        SEC.REGION,\r\n" + 
					"                            sec.PARTIDO zona,\r\n" + 
					"                            (SELECT SUM (tdet1.cant_afectaciones)\r\n" + 
					"                               FROM NEXUS_GIS.TABLA_ENREMTAT_DET tdet1\r\n" + 
					"                              WHERE tdet1.NRO_DOCUMENTO = tdet.NRO_DOCUMENTO)\r\n" + 
					"                               cli_actual\r\n" + 
					"              FROM NEXUS_GIS.TABLA_ENREMTAT_DET tdet,\r\n" + 
					"                   nexus_gis.oms_document     doc,\r\n" + 
					"                   nexus_gis.llam_sectores    sec,\r\n" + 
					"                   NEXUS_GIS.OMS_ADDRESS      a,\r\n" + 
					"                   NEXUS_GIS.AMAREAS          am\r\n" + 
					"             WHERE     tdet.estado NOT IN ('Cerrado', 'Cancelado')\r\n" + 
					"                   AND tdet.fecha_documento IS NOT NULL\r\n" + 
					"                   AND tdet.fecha_documento = (SELECT MIN (fecha_documento)\r\n" + 
					"                                                 FROM NEXUS_GIS.TABLA_ENREMTAT_DET\r\n" + 
					"                                                WHERE nro_documento = tdet.NRO_DOCUMENTO)\r\n" + 
					"                   AND tdet.zona IS NOT NULL\r\n" + 
					"                   AND tdet.nro_documento = DOC.NAME\r\n" + 
					"                   AND DOC.OPERATIVE_AREA_ID = AM.AREAID\r\n" + 
					"                   AND A.ID = DOC.ADDRESS_ID\r\n" + 
					"                   AND SEC.PARTIDO != 'CAPITAL FEDERAL'\r\n" + 
					"                   AND A.LARGE_AREA_ID = SEC.area_id\r\n" + 
					"                   AND A.MEDIUM_AREA_ID = SEC.part_ID\r\n" + 
					"                   AND A.SMALL_AREA_ID = SEC.LOCA_ID)\r\n" + 
					"          UNION\r\n" + 
					"          (SELECT DISTINCT (tdet.NRO_DOCUMENTO),\r\n" + 
					"                           --        SEC.REGION,\r\n" + 
					"                           sec.LOCALIDAD zona,\r\n" + 
					"                           (SELECT SUM (tdet1.cant_afectaciones)\r\n" + 
					"                              FROM NEXUS_GIS.TABLA_ENREMTAT_DET tdet1\r\n" + 
					"                             WHERE tdet1.NRO_DOCUMENTO = tdet.NRO_DOCUMENTO)\r\n" + 
					"                              cli_actual\r\n" + 
					"             FROM NEXUS_GIS.TABLA_ENREMTAT_DET tdet,\r\n" + 
					"                  nexus_gis.oms_document     doc,\r\n" + 
					"                  nexus_gis.llam_sectores    sec,\r\n" + 
					"                  NEXUS_GIS.OMS_ADDRESS      a,\r\n" + 
					"                  NEXUS_GIS.AMAREAS          am\r\n" + 
					"            WHERE     tdet.estado NOT IN ('Cerrado', 'Cancelado')\r\n" + 
					"                  AND tdet.fecha_documento IS NOT NULL\r\n" + 
					"                  AND tdet.fecha_documento = (SELECT MIN (fecha_documento)\r\n" + 
					"                                                FROM NEXUS_GIS.TABLA_ENREMTAT_DET\r\n" + 
					"                                               WHERE nro_documento = tdet.NRO_DOCUMENTO)\r\n" + 
					"                  AND tdet.zona IS NOT NULL\r\n" + 
					"                  AND tdet.nro_documento = DOC.NAME\r\n" + 
					"                  AND DOC.OPERATIVE_AREA_ID = AM.AREAID\r\n" + 
					"                  AND A.ID = DOC.ADDRESS_ID\r\n" + 
					"                  AND SEC.PARTIDO = 'CAPITAL FEDERAL'\r\n" + 
					"                  AND A.LARGE_AREA_ID = SEC.area_id\r\n" + 
					"                  AND A.MEDIUM_AREA_ID = SEC.part_ID\r\n" + 
					"                  AND A.SMALL_AREA_ID = SEC.LOCA_ID))\r\n" + 
					"GROUP BY zona";
			System.out.println("Raster MT AT LOC es " + plsql1);
			Statement stmt = this.con_1.createStatement(1004, 1007);
			ResultSet rs1 = stmt.executeQuery(plsql1);
			var5 = rs1;
		} catch (Exception var14) {
			this.con_1.close();
			System.out.println("Error en raster_bt()" + var14);
			throw var14;
		} finally {
			try {
				System.out.println("Final One");
			} catch (Exception var13) {
			}

		}

		return var5;
	}

	public ResultSet raster_tablahtml() throws Exception {
		Object var2 = null;

		ResultSet var5;
		try {
			this.Oracle_connect();
			String plsql1 = "select REGION, zona,sum(CLIENTES_BT) CLIENTES_BT,sum (CLIENTES_MTA) CLIENTES_MTA from (\n(SELECT sec.REGION as REGION, sec.SECTOR zona , sum(TE.CANTIDAD_EST_USU_AFECTADOS) CLIENTES_BT, 0 CLIENTES_MTA\n  FROM nexus_gis.tabla_enre te, nexus_gis.amareas am, nexus_gis.amareas amz, NEXUS_GIS.OMS_DOCUMENT doc, nexus_gis.llam_sectores sec,NEXUS_GIS.OMS_ADDRESS a       \n WHERE AM.AREANAME = te.partido \n   AND AM.AREATYPEID = 9 \n   AND am.superarea = amz.areaid \n   AND amz.areatypeid = 8 \n   AND TE.ESTADO not in ('Cerrado','Cancelado')\n   and DOC.NAME = TE.NRO_DOCUMENTO \n   and A.ID = DOC.ADDRESS_ID\n   and A.LARGE_AREA_ID = SEC.area_id\n   and A.MEDIUM_AREA_ID = SEC.part_ID\n   and A.SMALL_AREA_ID = SEC.LOCA_ID   \n GROUP BY sec.REGION, sec.SECTOR)\nunion\n(select REGION, zona, 0 CLIENTES_BT,sum (cli_actual) as CLIENTES_MTA from ( \n select distinct(tdet.NRO_DOCUMENTO),\n        SEC.REGION,SEC.SECTOR zona,\n        (select sum(tdet1.cant_afectaciones) from NEXUS_GIS.TABLA_ENREMTAT_DET tdet1 where tdet1.NRO_DOCUMENTO=tdet.NRO_DOCUMENTO ) cli_actual,\n        tdet.fecha_documento\n   from NEXUS_GIS.TABLA_ENREMTAT_DET tdet,\n        nexus_gis.oms_document doc,\n        nexus_gis.llam_sectores sec,\n        NEXUS_GIS.OMS_ADDRESS a,\n        NEXUS_GIS.AMAREAS am\n  where tdet.estado not in  ('Cerrado','Cancelado')\n    and tdet.fecha_documento is not null\n    and tdet.fecha_documento = (select min(fecha_documento) from NEXUS_GIS.TABLA_ENREMTAT_DET where nro_documento=tdet.NRO_DOCUMENTO)\n    and tdet.zona is not null\n    and tdet.nro_documento= DOC.NAME\n    and DOC.OPERATIVE_AREA_ID = AM.AREAID\n    and A.ID = DOC.ADDRESS_ID\n    and A.LARGE_AREA_ID = SEC.area_id\n    and A.MEDIUM_AREA_ID = SEC.part_ID\n    and A.SMALL_AREA_ID = SEC.LOCA_ID) group by REGION, zona)\n) group by region, zona\norder by 1,2 asc";
			System.out.println("Raster raster_tablahtml es " + plsql1);
			Statement stmt = this.con_1.createStatement(1004, 1007);
//         Statement stmt = this.con_1.createStatement(ResultSet.TYPE_SCROLL_SENSITIVE, ResultSet.CONCUR_READ_ONLY);
			ResultSet rs1 = stmt.executeQuery(plsql1);
			var5 = rs1;
		} catch (Exception var14) {
			this.con_1.close();
			System.out.println("Error en raster_bt()" + var14);
			throw var14;
		} finally {
			try {
				System.out.println("Final One");
			} catch (Exception var13) {
			}

		}

		return var5;
	}

	public ResultSet raster_bt() throws Exception {
		Object var2 = null;

		ResultSet var5;
		try {
			String plsql1 = "SELECT SEC.PARTIDO, trim(count(*)) cortes, trim(sum(TE.CANTIDAD_EST_USU_AFECTADOS)) Clientes_afect  \n  FROM nexus_gis.tabla_enre te, nexus_gis.amareas am, nexus_gis.amareas amz, NEXUS_GIS.OMS_DOCUMENT doc, nexus_gis.llam_sectores sec,NEXUS_GIS.OMS_ADDRESS a       \n WHERE AM.AREANAME = te.partido \n   AND AM.AREATYPEID = 9 \n   AND am.superarea = amz.areaid \n   AND amz.areatypeid = 8 \n   AND TE.ESTADO not in ('Cerrado','Cancelado')\n   and DOC.NAME = TE.NRO_DOCUMENTO \n   and A.ID = DOC.ADDRESS_ID\n   and A.LARGE_AREA_ID = SEC.area_id\n   and A.MEDIUM_AREA_ID = SEC.part_ID\n   and A.SMALL_AREA_ID = SEC.LOCA_ID   \n   and SEC.PARTIDO != 'CAPITAL FEDERAL'\n GROUP BY  SEC.PARTIDO\n UNION  \nSELECT SEC.LOCALIDAD PARTIDO, trim(count(*)) cortes, trim(sum(TE.CANTIDAD_EST_USU_AFECTADOS)) Clientes_afect  \n  FROM nexus_gis.tabla_enre te, nexus_gis.amareas am, nexus_gis.amareas amz, NEXUS_GIS.OMS_DOCUMENT doc, nexus_gis.llam_sectores sec,NEXUS_GIS.OMS_ADDRESS a       \n WHERE AM.AREANAME = te.partido \n   AND AM.AREATYPEID = 9 \n   AND am.superarea = amz.areaid \n   AND amz.areatypeid = 8 \n   AND TE.ESTADO not in ('Cerrado','Cancelado')\n   and DOC.NAME = TE.NRO_DOCUMENTO \n   and A.ID = DOC.ADDRESS_ID\n   and A.LARGE_AREA_ID = SEC.area_id\n   and A.MEDIUM_AREA_ID = SEC.part_ID\n   and A.SMALL_AREA_ID = SEC.LOCA_ID\n   and SEC.PARTIDO = 'CAPITAL FEDERAL'   \n GROUP BY SEC.LOCALIDAD \n order by 1 asc";
			System.out.println("Raster bt es: " + plsql1);
			Statement stmt = this.con_1.createStatement(1004, 1007);
			ResultSet rs1 = stmt.executeQuery(plsql1);
			var5 = rs1;
		} catch (Exception var14) {
			this.con_1.close();
			System.out.println("Error en raster_bt()" + var14);
			throw var14;
		} finally {
			try {
				System.out.println("Final One");
			} catch (Exception var13) {
			}

		}

		return var5;
	}

	private void creaExcel() throws FileNotFoundException, IOException, SQLException, Exception {
            Workbook wb = new XSSFWorkbook();
            Sheet shForzados = wb.createSheet("ConsolidadoBTMT");
            this.pueblaExel(shForzados, wb);
            FileOutputStream fileOut = new FileOutputStream(excel);
            Throwable var4 = null;

            try {
                wb.write(fileOut);
            } catch (Throwable var13) {
                var4 = var13;
                throw var13;
            } finally {
                if (fileOut != null) {
                    if (var4 != null) {
                        try {
                                fileOut.close();
                        } catch (Throwable var12) {
                                var4.addSuppressed(var12);
                        }
                    } else {
                        fileOut.close();
                    }
                }
            }
	}

	private void pueblaExel(Sheet sh, Workbook wb) throws SQLException, Exception {
		Integer rowCount = 0;
		int i = 0;
		int tot_clientes = 0;
		int tot_cortesbt = 0;
		int tot_afectbt = 0;
		int tot_cortesmt = 0;
		int tot_afectmt = 0;
		int tot_cliafect = 0;
		// int porc_total = false;
		rowCount++;
		Row row = sh.createRow(rowCount);
		CellStyle bordesgruesos = creaEstilosbordesgrueso(wb);
		CellStyle cabecerayellow = creaEstilosCabeyellow(wb);
		CellStyle cabecerared = creaEstilosCabered(wb);
		CellStyle cabeceragreen = creaEstilosCabegreen(wb);
		CellStyle cabecera = creaEstilosCabe(wb);
		CellStyle datosLeft = creaEstilosDatos(wb, "left");
		CellStyle datosRight = creaEstilosDatos(wb, "right");
		CellStyle datosCenter = creaEstilosDatos(wb, "center");
		CellStyle datosred = creaEstilosColorRed(wb);
		CellStyle sinborde = creaEstilosnoborder(wb);
		CellStyle titulo = creaEstilostitulo(wb);
		CellStyle totalesCenter = creaEstilosTotales(wb, "center", "celeste");
		CellStyle totalesRight = creaEstilosTotales(wb, "right", "celeste");
		CellStyle totalesGrisado = creaEstilosTotales(wb, "right", "grisado");
		Cell cell = row.createCell(0);

		for (int z = 0; z < 300; ++z) {
			row = sh.createRow(z);

			for (int j = 0; j < 300; ++j) {
				cell = row.createCell(j);
				cell.setCellStyle(sinborde);
				if (z == 7 && j == 2) {
					cell.setCellValue("Afectaciones y clientes sin servicio e indices de incidencia - " + this.Hora + " hs");
					cell.setCellStyle(titulo);
				}
			}
		}

		rowCount = 10;
		row = sh.getRow(rowCount);
		cell = row.createCell(1);
		cell.setCellValue("Partido/Barrio");
		cell.setCellStyle(cabecera);
		sh.setColumnWidth(0, 6000);
		cell = row.createCell(2);
		cell.setCellValue("Total Usuarios x Partido /Barrio");
		cell.setCellStyle(cabecera);
		sh.setColumnWidth(1, 4000);
		cell = row.createCell(3);
		cell.setCellValue("Afectaciones BT");
		cell.setCellStyle(cabecera);
		sh.setColumnWidth(2, 4500);
		cell = row.createCell(4);
		cell.setCellValue("Clientes afectados BT");
		cell.setCellStyle(cabecera);
		sh.setColumnWidth(3, 4500);
		cell = row.createCell(5);
		cell.setCellValue("Afectaciones MT");
		cell.setCellStyle(cabecera);
		sh.setColumnWidth(4, 4000);
		cell = row.createCell(6);
		cell.setCellValue("Clientes afectados MT");
		cell.setCellStyle(cabecera);
		sh.setColumnWidth(5, 5000);
		cell = row.createCell(7);
		cell.setCellValue("Consolidado Clientes Afectados");
		cell.setCellStyle(cabecera);
		sh.setColumnWidth(6, 5000);
		cell = row.createCell(8);
		cell.setCellValue("%Clientes afectados sobre total partido/barrio");
		cell.setCellStyle(cabecera);
		sh.setColumnWidth(7, 5000);
		ResultSet r = this.raster_bt();

		for (ResultSet r1 = this.raster_mtat_loc(); rs1.next(); ++i) {
                    arraybarrio[i] = rs1.getString(1);
                    cant_clientes[i] = rs1.getInt(2);
                    r.beforeFirst();

                    String barrio1;
                    while (r.next()) {
                        barrio1 = r.getString(1).replace(" ", "");
                        if (barrio1.equals(arraybarrio[i].replace(" ", ""))) {
                            afectbt[i] = r.getInt(2);
                            cli_afectbt[i] = r.getInt(3);
                        }
                    }

                    r1.beforeFirst();

                    while (r1.next()) {
                            barrio1 = r1.getString(1).replace(" ", "");
                            if (barrio1.equals(arraybarrio[i].replace(" ", ""))) {
                                    afectmt[i] = r1.getInt(2);
                                    cli_afectmt[i] = r1.getInt(3);
                            }
                    }
		}

		int suma;
		for (suma = 0; suma < 40; ++suma) {
			System.out.println(arraybarrio[suma] + "||" + cant_clientes[suma] + "||" + afectbt[suma] + "||"
					+ cli_afectbt[suma] + "||" + afectmt[suma] + "||" + cli_afectmt[suma]);
		}

		rowCount = rowCount + 1;

		for (suma = 0; suma < 35; ++suma) {
			System.out.println(suma);
			System.out.println(rowCount);
			row = sh.getRow(rowCount);
			row.createCell(1).setCellValue(arraybarrio[suma]);
			row.getCell(1).setCellStyle(bordesgruesos);
			row.createCell(2).setCellValue((double) cant_clientes[suma]);
			tot_clientes += cant_clientes[suma];
			row.getCell(2).setCellStyle(datosCenter);
			row.createCell(3).setCellValue((double) afectbt[suma]);
			tot_cortesbt += afectbt[suma];
			row.getCell(3).setCellStyle(datosCenter);
			row.createCell(4).setCellValue((double) cli_afectbt[suma]);
			tot_afectbt += cli_afectbt[suma];
			row.getCell(4).setCellStyle(datosCenter);
			row.createCell(5).setCellValue((double) afectmt[suma]);
			tot_cortesmt += afectmt[suma];
			row.getCell(5).setCellStyle(datosCenter);
			row.createCell(6).setCellValue((double) cli_afectmt[suma]);
			tot_afectmt += cli_afectmt[suma];
			row.getCell(6).setCellStyle(datosCenter);
			row.createCell(7).setCellValue((double) (cli_afectbt[suma] + cli_afectmt[suma]));
			tot_cliafect = tot_cliafect + cli_afectbt[suma] + cli_afectmt[suma];
			row.getCell(7).setCellStyle(datosCenter);
			int suma2 = (cli_afectbt[suma] + cli_afectmt[suma]) * 100;
			float porcentaje1 = (float) suma2 / (float) cant_clientes[suma];
			row.createCell(8).setCellValue(String.format("%.2f", porcentaje1) + "%");
			this.setbgcolor(sh, wb, row, 8, porcentaje1, cant_clientes[suma]);
			rowCount = rowCount + 1;
		}

		row = sh.getRow(rowCount);
		cell = row.createCell(1);
		cell.setCellValue("TOTALES");
		cell.setCellStyle(cabecera);
		sh.setColumnWidth(0, 6000);
		cell = row.createCell(2);
		cell.setCellValue((double) tot_clientes);
		cell.setCellStyle(cabecera);
		sh.setColumnWidth(1, 4000);
		cell = row.createCell(3);
		cell.setCellValue((double) tot_cortesbt);
		cell.setCellStyle(cabecera);
		sh.setColumnWidth(2, 4500);
		cell = row.createCell(4);
		cell.setCellValue((double) tot_afectbt);
		cell.setCellStyle(cabecera);
		sh.setColumnWidth(3, 4500);
		cell = row.createCell(5);
		cell.setCellValue((double) tot_cortesmt);
		cell.setCellStyle(cabecera);
		sh.setColumnWidth(4, 4000);
		cell = row.createCell(6);
		cell.setCellValue((double) tot_afectmt);
		cell.setCellStyle(cabecera);
		sh.setColumnWidth(5, 5000);
		cell = row.createCell(7);
		cell.setCellValue((double) (tot_afectmt + tot_afectbt));
		cell.setCellStyle(cabecera);
		sh.setColumnWidth(6, 5000);
		cell = row.createCell(8);
		cell.setCellValue((double) ((tot_afectmt + tot_afectbt) * 100 / tot_clientes));
		suma = (tot_afectmt + tot_afectbt) * 100;
		float porcentaje1 = (float) suma / (float) tot_clientes;
		cell.setCellValue(String.format("%.2f", porcentaje1) + "%");
		if ((double) porcentaje1 < 1.5D) {
			cell.setCellStyle(cabeceragreen);
		} else if (porcentaje1 > 4.0F) {
			cell.setCellStyle(cabecerared);
		} else {
			cell.setCellStyle(cabecerayellow);
		}

		sh.setColumnWidth(7, 5000);
	}

	private void setbgcolor(Sheet sh, Workbook wb, Row row, int numcel, float porcentaje1, int cant_total_usuarios) {
		CellStyle cabecera = creaEstilosCabe(wb);
		CellStyle datosred = creaEstilosColorRed(wb);
		CellStyle datosgreen = creaEstilosColorGreen(wb);
		CellStyle datosyellow = creaEstilosColorYellow(wb);
		if (cant_total_usuarios < 10000) {
			if (porcentaje1 < 20.0F) {
				row.getCell(numcel).setCellStyle(datosgreen);
			} else if (porcentaje1 > 50.0F) {
				row.getCell(numcel).setCellStyle(datosred);
			} else {
				row.getCell(numcel).setCellStyle(datosyellow);
			}
		} else if (cant_total_usuarios > 50000) {
			if (porcentaje1 < 10.0F) {
				row.getCell(numcel).setCellStyle(datosgreen);
			} else if (porcentaje1 > 30.0F) {
				row.getCell(numcel).setCellStyle(datosred);
			} else {
				row.getCell(numcel).setCellStyle(datosyellow);
			}
		} else if (porcentaje1 < 15.0F) {
			row.getCell(numcel).setCellStyle(datosgreen);
		} else if (porcentaje1 > 40.0F) {
			row.getCell(numcel).setCellStyle(datosred);
		} else {
			row.getCell(numcel).setCellStyle(datosyellow);
		}

	}

	public ConsolidadoBTMT() throws Exception {
		try {
			this.Oracle_connect();

			for (ResultSet rs = Fechainforme(this.con_1); rs.next(); this.Dia = rs.getString("FECHA")) {
			}

			this.Hora = this.Dia.substring(11, 13);
			String dia1 = this.Dia.substring(0, 5);
			this.Dia = dia1;
			int i = 0;
			int j = 0;
			int totbt = 0;
			int totmt = 0;
			rs2 = this.raster_tablahtml();

			// Armo lista base
			Map<Integer, String> listaPartidos = new HashMap<>();
			listaPartidos.put(0, "R1;NORTE;0;0;0,");
			listaPartidos.put(1, "R1;OLIVOS;0;0;0,");
			listaPartidos.put(2, "R1;SAN MARTIN;0;0;0,");
			listaPartidos.put(3, "R2;MATANZA;0;0;0,");
			listaPartidos.put(4, "R2;MERLO;0;0;0,");
			listaPartidos.put(5, "R2;MORON;0;0;0,");
			listaPartidos.put(6, "R3;MORENO;0;0;0,");
			listaPartidos.put(7, "R3;PILAR;0;0;0,");
			listaPartidos.put(8, "R3;SAN MIGUEL;0;0;0,");
			listaPartidos.put(9, "R3;TIGRE;0;0;0,");

			// Armo lista y totales con lo que me llega de la DB
			Map<String, String> registros = new HashMap<>();
			while (rs2.next()) {
				registros.put(rs2.getString(2).trim(), rs2.getString(1) + ";" + rs2.getString(2).trim() + ";" + rs2.getString(3) + ";"
						+ rs2.getString(4) + ";" + (rs2.getInt(3) + rs2.getInt(4)) + ",");
				totbt += rs2.getInt(3);
				totmt += rs2.getInt(4);
			}

			// Recorro y si existe registro en la DB, uso esos datos
			for (i = 0; i < 10; i++) {
				String res = listaPartidos.get(i);
				String partido = res.split(";")[1];

				if (registros.containsKey(partido)) {
					res = registros.get(partido);
				}

				this.tablahtml += res;

			}

//         while(true) {
			this.tablahtml = this.tablahtml + "TOTAL; ;" + totbt + ";" + totmt + ";" + (totbt + totmt) + ",";
			this.tabla2 = this.strToHash2(this.tablahtml);
			rs1 = this.get_clients();
			this.creaExcel();
			this.generar(this.tabla2, this.msgCancela, true);
			return;
//         }
		} catch (Exception var8) {
			System.out.println("Error en Mail_Sender() = " + var8);
			var8.printStackTrace();
			throw var8;
		}
	}

	private void generar(List<HashMap<String, String>> tabla2, String encabezado, boolean hayDatos) throws Exception {
		Hashtable<String, String> hst_Mail = new Hashtable();
		boolean enviarMail = false;
		boolean i = false;

		try {
			i = false;
			enviarMail = true;
			hst_Mail.put("mailHost", properties.getProperty("mail_host"));              //mail.edenor
			hst_Mail.put("DE", properties.getProperty("mail_from"));                    //centrodeinformacion@edenor.com
			hst_Mail.put("PARA", properties.getProperty("mail_to"));
			hst_Mail.put("ASUNTO", String.format(properties.getProperty("mail_subject"),this.Dia,this.Hora));              //"Informacion Enre - " + this.Dia + "-" + this.Hora + ":00 Hs.");
			String HTML_Estucture = "<html>";
			HTML_Estucture = HTML_Estucture + "<head>";
			HTML_Estucture = HTML_Estucture + "<style id='Mail_Styles'>";
			HTML_Estucture = HTML_Estucture + "<!--table";
			HTML_Estucture = HTML_Estucture + ".xl1530982";
			HTML_Estucture = HTML_Estucture
					+ "\t{padding:0px; \tmso-ignore:padding; \tcolor:black; \tfont-size:11.0pt; \tfont-weight:400; \tfont-style:normal; \ttext-decoration:none; \tfont-family:Calibri, sans-serif; \tmso-font-charset:0; \tmso-number-format:General; \ttext-align:general; \tvertical-align:bottom; \tmso-background-source:auto; \tmso-pattern:auto; \twhite-space:nowrap;}";
			HTML_Estucture = HTML_Estucture + ".xl6530982";
			HTML_Estucture = HTML_Estucture
					+ "\t{padding:0px; \tmso-ignore:padding; \tcolor:black; \tfont-size:11.0pt; \tfont-weight:700; \tfont-style:normal; \ttext-decoration:none; \tfont-family:Calibri, sans-serif; \tmso-font-charset:0; \tmso-number-format:General; \ttext-align:general; \tvertical-align:middle; \tborder-top:1.0pt solid windowtext; \tborder-right:.5pt solid windowtext; \tborder-bottom:1.0pt solid windowtext; \tborder-left:1.0pt solid windowtext; \tbackground:#FFC000; \tmso-pattern:black none; \twhite-space:nowrap;}";
			HTML_Estucture = HTML_Estucture + ".xl6630982";
			HTML_Estucture = HTML_Estucture
					+ "\t{padding:5px; \tcolor:black; \tfont-size:10pt; \tfont-weight:700; \tfont-style:normal; \ttext-decoration:none; \tfont-family:Calibri, sans-serif; \tmso-font-charset:0; \tmso-number-format:General; \ttext-align:center; \tvertical-align:middle; \tborder-top:1.0pt solid windowtext; \tborder-right:.5pt solid windowtext; \tborder-bottom:1.0pt solid windowtext; \tborder-left:.5pt solid windowtext; \tbackground:#C5D9F1; \tmso-pattern:black none; \twhite-space:nowrap;}";
			HTML_Estucture = HTML_Estucture + ".xl6730982";
			HTML_Estucture = HTML_Estucture
					+ "\t{padding:0px; \tmso-ignore:padding; \tcolor:black; \tfont-size:11.0pt; \tfont-weight:700; \tfont-style:normal; \ttext-decoration:none; \tfont-family:Calibri, sans-serif; \tmso-font-charset:0; \tmso-number-format:General; \ttext-align:right; \tvertical-align:middle; \tborder-top:1.0pt solid windowtext; \tborder-right:1.0pt solid windowtext; \tborder-bottom:1.0pt solid windowtext; \tborder-left:.5pt solid windowtext; \tbackground:#FFC000; \tmso-pattern:black none; \twhite-space:nowrap;}";
			HTML_Estucture = HTML_Estucture + ".xl6830982";
			HTML_Estucture = HTML_Estucture
					+ "\t{padding:5px; \tmso-ignore:padding; \tcolor:black; \tfont-size:10.0pt; \tfont-weight:400; \tfont-style:normal; \ttext-decoration:none; \tfont-family:Calibri, sans-serif; \tmso-font-charset:0; \tmso-number-format:General; \ttext-align:general; \tvertical-align:middle; \tborder-top:1.0pt solid windowtext; \tborder-right:.5pt solid windowtext; \tborder-bottom:.5pt solid windowtext; \tborder-left:.5pt solid windowtext; \tmso-background-source:auto; \tmso-pattern:auto; \twhite-space:nowrap;}";
			HTML_Estucture = HTML_Estucture + ".xl6930982";
			HTML_Estucture = HTML_Estucture
					+ "\t{padding:0px; \tmso-ignore:padding; \tcolor:black; \tfont-size:11.0pt; \tfont-weight:400; \tfont-style:normal; \ttext-decoration:none; \tfont-family:Calibri, sans-serif; \tmso-font-charset:0; \tmso-number-format:'General Date'; \ttext-align:general; \tvertical-align:middle; \tborder-top:1.0pt solid windowtext; \tborder-right:.5pt solid windowtext; \tborder-bottom:.5pt solid windowtext; \tborder-left:.5pt solid windowtext; \tmso-background-source:auto; \tmso-pattern:auto; \twhite-space:nowrap;}";
			HTML_Estucture = HTML_Estucture + ".xl7030982";
			HTML_Estucture = HTML_Estucture
					+ "\t{padding:0px; \tmso-ignore:padding; \tcolor:black; \tfont-size:11.0pt; \tfont-weight:400; \tfont-style:normal; \ttext-decoration:none; \tfont-family:Calibri, sans-serif; \tmso-font-charset:0; \tmso-number-format:'Short Time'; \ttext-align:general; \tvertical-align:middle; \tborder-top:1.0pt solid windowtext; \tborder-right:1.0pt solid windowtext; \tborder-bottom:.5pt solid windowtext; \tborder-left:.5pt solid windowtext; \tmso-background-source:auto; \tmso-pattern:auto; \twhite-space:nowrap;}";
			HTML_Estucture = HTML_Estucture + ".xl7130982";
			HTML_Estucture = HTML_Estucture
					+ "\t{padding:0px; \tmso-ignore:padding; \tcolor:black; \tfont-size:11.0pt; \tfont-weight:400; \tfont-style:normal; \ttext-decoration:none; \tfont-family:Calibri, sans-serif; \tmso-font-charset:0; \tmso-number-format:General; \ttext-align:general; \tvertical-align:middle; \tborder:.5pt solid windowtext; \tmso-background-source:auto; \tmso-pattern:auto; \twhite-space:nowrap;}";
			HTML_Estucture = HTML_Estucture + ".xl7230982";
			HTML_Estucture = HTML_Estucture
					+ "\t{padding:0px; \tmso-ignore:padding; \tcolor:black; \tfont-size:11.0pt; \tfont-weight:400; \tfont-style:normal; \ttext-decoration:none; \tfont-family:Calibri, sans-serif; \tmso-font-charset:0; \tmso-number-format:'General Date'; \ttext-align:general; \tvertical-align:middle; \tborder:.5pt solid windowtext; \tmso-background-source:auto; \tmso-pattern:auto; \twhite-space:nowrap;}";
			HTML_Estucture = HTML_Estucture + ".xl7330982";
			HTML_Estucture = HTML_Estucture
					+ "\t{padding:0px; \tmso-ignore:padding; \tcolor:black; \tfont-size:11.0pt; \tfont-weight:400; \tfont-style:normal; \ttext-decoration:none; \tfont-family:Calibri, sans-serif; \tmso-font-charset:0; \tmso-number-format:'Short Time'; \ttext-align:general; \tvertical-align:middle; \tborder-top:.5pt solid windowtext; \tborder-right:1.0pt solid windowtext; \tborder-bottom:.5pt solid windowtext; \tborder-left:.5pt solid windowtext; \tmso-background-source:auto; \tmso-pattern:auto; \twhite-space:nowrap;}";
			HTML_Estucture = HTML_Estucture + ".xl7430982";
			HTML_Estucture = HTML_Estucture
					+ "\t{padding:0px; \tmso-ignore:padding; \tcolor:black; \tfont-size:11.0pt; \tfont-weight:400; \tfont-style:normal; \ttext-decoration:none; \tfont-family:Calibri, sans-serif; \tmso-font-charset:0; \tmso-number-format:General; \ttext-align:general; \tvertical-align:middle; \tborder-top:.5pt solid windowtext; \tborder-right:.5pt solid windowtext; \tborder-bottom:1.0pt solid windowtext; \tborder-left:.5pt solid windowtext; \tmso-background-source:auto; \tmso-pattern:auto; \twhite-space:nowrap;}";
			HTML_Estucture = HTML_Estucture + ".xl7530982";
			HTML_Estucture = HTML_Estucture
					+ "\t{padding:0px; \tmso-ignore:padding; \tcolor:black; \tfont-size:11.0pt; \tfont-weight:400; \tfont-style:normal; \ttext-decoration:none; \tfont-family:Calibri, sans-serif; \tmso-font-charset:0; \tmso-number-format:'General Date'; \ttext-align:general; \tvertical-align:middle; \tborder-top:.5pt solid windowtext; \tborder-right:.5pt solid windowtext; \tborder-bottom:1.0pt solid windowtext; \tborder-left:.5pt solid windowtext; \tmso-background-source:auto; \tmso-pattern:auto; \twhite-space:nowrap;}";
			HTML_Estucture = HTML_Estucture + ".xl7630982";
			HTML_Estucture = HTML_Estucture
					+ "\t{padding:0px; \tmso-ignore:padding; \tcolor:black; \tfont-size:11.0pt; \tfont-weight:400; \tfont-style:normal; \ttext-decoration:none; \tfont-family:Calibri, sans-serif; \tmso-font-charset:0; \tmso-number-format:'Short Time'; \ttext-align:general; \tvertical-align:middle; \tborder-top:.5pt solid windowtext; \tborder-right:1.0pt solid windowtext; \tborder-bottom:1.0pt solid windowtext; \tborder-left:.5pt solid windowtext; \tmso-background-source:auto; \tmso-pattern:auto; \twhite-space:nowrap;}";
			HTML_Estucture = HTML_Estucture + ".xl7730982";
			HTML_Estucture = HTML_Estucture
					+ "\t{padding:0px; \tmso-ignore:padding; \tcolor:black; \tfont-size:12.0pt; \tfont-weight:400; \tfont-style:normal; \ttext-decoration:none; \tfont-family:'Times New Roman', serif; \tmso-font-charset:0; \tmso-number-format:General; \ttext-align:general; \tvertical-align:bottom; \tmso-background-source:auto; \tmso-pattern:auto; \twhite-space:nowrap;}";
			HTML_Estucture = HTML_Estucture + ".xl7830982";
			HTML_Estucture = HTML_Estucture
					+ "\t{padding:5px; \tmso-ignore:padding; \tcolor:black; \tfont-size:11.0pt; \tfont-weight:400; \tfont-style:normal; \ttext-decoration:none; \tfont-family:Calibri, sans-serif; \tmso-font-charset:0; \tmso-number-format:General; \ttext-align:general; \tvertical-align:middle; \tborder-top:1.0pt solid windowtext; \tborder-right:.5pt solid windowtext; \tborder-bottom:.5pt solid windowtext; \tborder-left:1.0pt solid windowtext; \tbackground:#C5D9F1; \tmso-pattern:black none; \twhite-space:nowrap;}";
			HTML_Estucture = HTML_Estucture + ".xl7930982";
			HTML_Estucture = HTML_Estucture
					+ "\t{padding:0px; \tmso-ignore:padding; \tcolor:black; \tfont-size:11.0pt; \tfont-weight:400; \tfont-style:normal; \ttext-decoration:none; \tfont-family:Calibri, sans-serif; \tmso-font-charset:0; \tmso-number-format:General; \ttext-align:general; \tvertical-align:middle; \tborder-top:.5pt solid windowtext; \tborder-right:.5pt solid windowtext; \tborder-bottom:.5pt solid windowtext; \tborder-left:1.0pt solid windowtext; \tbackground:#C5D9F1; \tmso-pattern:black none; \twhite-space:nowrap;}";
			HTML_Estucture = HTML_Estucture + ".xl8030982";
			HTML_Estucture = HTML_Estucture
					+ "\t{padding:0px; \tmso-ignore:padding; \tcolor:black; \tfont-size:11.0pt; \tfont-weight:400; \tfont-style:normal; \ttext-decoration:none; \tfont-family:Calibri, sans-serif; \tmso-font-charset:0; \tmso-number-format:General; \ttext-align:general; \tvertical-align:middle; \tborder-top:.5pt solid windowtext; \tborder-right:.5pt solid windowtext; \tborder-bottom:1.0pt solid windowtext; \tborder-left:1.0pt solid windowtext; \tbackground:#C5D9F1; \tmso-pattern:black none; \twhite-space:nowrap;}";
			HTML_Estucture = HTML_Estucture + ".padded";
			HTML_Estucture = HTML_Estucture
					+ "\t{padding:5px; \tcolor:black; \tfont-size:10.0pt; \tfont-weight:400; \tfont-style:normal; \ttext-decoration:none; \tfont-family:Calibri, sans-serif; \tmso-font-charset:0; \tmso-number-format:General; \ttext-align:general; \tvertical-align:middle; \tborder-top:.5pt solid windowtext; \tborder-right:.5pt solid windowtext; \tborder-bottom:1.0pt solid windowtext; \tborder-left:1.0pt solid windowtext; \tbackground:#C5D9F1; \tmso-pattern:black none; \twhite-space:nowrap;}";
			HTML_Estucture = HTML_Estucture + "-->";
			HTML_Estucture = HTML_Estucture + "</style>";
			HTML_Estucture = HTML_Estucture + "</head>";
			HTML_Estucture = HTML_Estucture + "<body>";
			HTML_Estucture = HTML_Estucture + "<div id='Mail_body' align=left>";
			HTML_Estucture = HTML_Estucture
					+ "<table border=0 cellpadding=0 cellspacing=0 width=494 style='border-collapse: collapse;table-layout:auto;width:371pt'>";
			HTML_Estucture = HTML_Estucture + " <tr height=21 style='height:15.75pt'>";
			HTML_Estucture = HTML_Estucture
					+ "  <td height=21 class=xl7730982 colspan=9 width=494 style='height:15.75pt;width:371pt'>" + encabezado
					+ "</td>";
			HTML_Estucture = HTML_Estucture + " </tr>";
			HTML_Estucture = HTML_Estucture
					+ "<table border=0 cellpadding=0 cellspacing=0 style='border-collapse: collapse;table-layout:auto;width:150pt'>";
			HTML_Estucture = HTML_Estucture + " <tr height=27 style='mso-height-source:userset;height:20.25pt'>";
			HTML_Estucture = HTML_Estucture + "  <td height=27 width=90 class=xl6630982 style='height:20.25pt'>REGION</td>";
			HTML_Estucture = HTML_Estucture + "  <td height=27 width=90 class=xl6630982 style='height:20.25pt'>ZONA</td>";
			HTML_Estucture = HTML_Estucture + "  <td height=27 width=90 class=xl6630982 style='height:20.25pt'>BT</td>";
			HTML_Estucture = HTML_Estucture + "  <td height=27 width=90 class=xl6630982 style='height:20.25pt'>MT</td>";
			HTML_Estucture = HTML_Estucture + "  <td class=xl6630982 style='border-left:none'>TOTAL</td>";
			HTML_Estucture = HTML_Estucture + " </tr>";
			this.it = tabla2.iterator();

			while (true) {
                            if (!this.it.hasNext()) {
                                    HTML_Estucture = HTML_Estucture + "</table>";
                                    HTML_Estucture = HTML_Estucture + fuente_html_fin;
                                    System.out.println("enviado: " + HTML_Estucture);
                                    hst_Mail.put("CUERPO", HTML_Estucture);
                                    if (enviarMail) {
                                            this.mailSender(hst_Mail);
                                    }
                                    break;
                            }

                            HashMap fila = (HashMap) this.it.next();
                            if (fila.get("REGION").equals("TOTAL")) {
                                    HTML_Estucture = HTML_Estucture
                                                    + " <td colspan='2' height=27 width=90 class=xl7830982 align=center  style='height:40.50pt; font-weight: bold'>"
                                                    + fila.get("REGION") + "</td>";
                            } else {
                                    HTML_Estucture = HTML_Estucture
                                                    + " <td height=27 width=90 class=xl7830982 align=center  style='height:20.25pt'>" + fila.get("REGION")
                                                    + "</td>";
                                    HTML_Estucture = HTML_Estucture
                                                    + " <td height=27 width=90 class=xl7830982 align=center  style='height:20.25pt'>" + fila.get("ZONA")
                                                    + "</td>";
                            }

                            HTML_Estucture = HTML_Estucture
                                            + " <td height=27 width=90 class=xl6830982 align=center style='border-left:none'>" + fila.get("BT")
                                            + "</td>";
                            HTML_Estucture = HTML_Estucture
                                            + " <td height=27 width=90 class=xl6830982 align=center style='border-left:none'>" + fila.get("MT")
                                            + "</td>";
                            HTML_Estucture = HTML_Estucture
                                            + " <td height=27 width=90 class=xl6830982 align=center style='border-left:none'>" + fila.get("TOTAL")
                                            + "</td>";
                            HTML_Estucture = HTML_Estucture + " </tr>";
			}
		} catch (Exception var16) {
			System.out.println("Error en generar()" + var16);
			this.msgCancela = var16.toString();
			hst_Mail.put("CUERPO", this.msgCancela);
			this.mailSender(hst_Mail);
			throw var16;
		} finally {
			try {
				System.out.println("final ");
			} catch (Exception var15) {
			}

		}

	}

	public void mailSender(Hashtable<String, String> hst_values_mail) throws Exception {
		try {
			Properties properties = new Properties();
			properties.put("mail.smtp.host", hst_values_mail.get("mailHost"));
			properties.put("mail.from", hst_values_mail.get("DE"));
			properties.put("mail.debug", "true");
			Session session = Session.getInstance(properties, (Authenticator) null);
			MimeMessage msg = new MimeMessage(session);
			msg.setFrom(new InternetAddress((String) hst_values_mail.get("DE")));
			msg.setFrom(InternetAddress.getLocalAddress(session));
			msg.setSubject((String) hst_values_mail.get("ASUNTO"));
			msg.setSentDate(new Date());
			InternetAddress[] paraArray = InternetAddress.parse((String) hst_values_mail.get("PARA"));
			msg.setRecipients(RecipientType.TO, paraArray);
			InternetAddress[] ccArray = null;
			if (hst_values_mail.get("CC") != null) {
				ccArray = InternetAddress.parse((String) hst_values_mail.get("CC"));
				msg.setRecipients(RecipientType.CC, ccArray);
			}

			InternetAddress[] bccArray = null;
			if (hst_values_mail.get("CCO") != null) {
				bccArray = InternetAddress.parse((String) hst_values_mail.get("CCO"));
				msg.setRecipients(RecipientType.BCC, bccArray);
			}

			MimeMultipart multiParte = new MimeMultipart();
			BodyPart logo = new MimeBodyPart();
			logo.setDataHandler(new DataHandler(new FileDataSource("./logo.png")));
			logo.setFileName("logo.png");
			multiParte.addBodyPart(logo);
			BodyPart adjunto = new MimeBodyPart();
			adjunto.setDataHandler(new DataHandler(new FileDataSource(excel)));
			adjunto.setFileName("ConsolidadoBTMT.xlsx");
			multiParte.addBodyPart(adjunto);
			BodyPart texto = new MimeBodyPart();
			texto.setDataHandler(new DataHandler(new ConsolidadoBTMT_HTMLDataSource((String) hst_values_mail.get("CUERPO"))));
			multiParte.addBodyPart(texto);
			msg.setContent(multiParte);
			int total = paraArray.length;
			if (ccArray != null) {
				total += ccArray.length;
			}

			if (bccArray != null) {
				total += bccArray.length;
			}

			InternetAddress[] address = new InternetAddress[total];

			int i;
			for (i = 0; i < paraArray.length; ++i) {
				address[i] = paraArray[i];
			}

			if (ccArray != null) {
				for (int j = 0; j < ccArray.length; ++j) {
					address[i] = ccArray[j];
					++i;
				}
			}

			if (bccArray != null) {
				for (int k = 0; k < bccArray.length; ++k) {
					address[i] = bccArray[k];
					++i;
				}
			}

			Transport transporte = session.getTransport(address[0]);
			transporte.connect();
			transporte.sendMessage(msg, address);
		} catch (SendFailedException var18) {
			Address[] listaInval = var18.getInvalidAddresses();
			Address[] var4 = listaInval;
			int var5 = listaInval.length;

			for (int var6 = 0; var6 < var5; ++var6) {
				Address listaInval1 = var4[var6];
				this.dir_no_encontradas.add(listaInval1.toString());
				System.out.println("No encontrada: " + listaInval1.toString());
			}
		} catch (MessagingException var19) {
			System.out.println("Exception (mailSender) : " + var19);
			throw var19;
		}

	}

	public String Minute_To_Hour(long L_Minute) throws Exception {
		try {
			long L_Hour = 0L;
			long L_Aux = 0L;

			try {
				L_Hour = L_Minute / 60L;
				L_Aux = L_Minute % 60L;
			} catch (NumberFormatException var9) {
				throw new Exception("Error en el String de Hora " + var9);
			}

			String S_out;
			if (L_Hour < 10L) {
				S_out = "0" + String.valueOf(L_Hour);
			} else {
				S_out = String.valueOf(L_Hour);
			}

			S_out = S_out + ":";
			if (L_Aux < 10L) {
				S_out = S_out + "0" + L_Aux;
			} else {
				S_out = S_out + String.valueOf(L_Aux);
			}

			return S_out;
		} catch (Exception var10) {
			System.out.println("Error en Minute_To_Hour() = " + var10);
			throw var10;
		}
	}

	private static String clobToStr(Clob clb) {
		StringBuilder sb = new StringBuilder();

		try {
			Reader reader = clb.getCharacterStream();
			BufferedReader br = new BufferedReader(reader);
			Throwable var4 = null;

			try {
				int b;
				try {
					while (-1 != (b = br.read())) {
						sb.append((char) b);
					}
				} catch (Throwable var15) {
					var4 = var15;
					throw var15;
				}
			} finally {
				if (br != null) {
					if (var4 != null) {
						try {
							br.close();
						} catch (Throwable var14) {
							var4.addSuppressed(var14);
						}
					} else {
						br.close();
					}
				}

			}
		} catch (SQLException var17) {
			System.out.println("RedElectrica::clobToStr: SQL. No se pudo convertir CLOB a String");
		} catch (IOException var18) {
			System.out.println("RedElectrica::clobToStr: IO. No se pudo convertir CLOB a String");
		}

		return sb.toString();
	}

	public static void main(String[] args) {
            try {
                findPropiedades();
                if (excel!=null) {
                    new ConsolidadoBTMT();
                    System.out.println("\nProcedimiento MAIL terminado exitosamente");
                } else {
                    System.out.println("\nError en PROPIERTIES");
                }
                System.exit(0);
            } catch (Exception var2) {
                System.out.println(var2);
                System.exit(1);
            }

	}
        
        public static void findPropiedades(){
            properties= new Properties();
            try (FileInputStream input = new FileInputStream("./config.properties")) {
                properties.load(input);
                excel=properties.getProperty("excel");                          //"/ias/ConsolidadoBTMT/ConsolidadoBTMT.xlsx"
                destinatarios= properties.getProperty("mail_to");
            } catch (IOException e) {
                e.printStackTrace();
                excel=null;
                destinatarios=null;
            }
            
            System.out.println(String.format(properties.getProperty("mail_subject"),5,7));
        }        
        
}