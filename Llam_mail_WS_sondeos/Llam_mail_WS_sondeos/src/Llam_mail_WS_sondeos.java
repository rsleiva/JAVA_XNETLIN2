import org.apache.commons.codec.binary.*;
import java.io.BufferedReader;
import java.io.BufferedWriter;
import java.io.ByteArrayInputStream;
import java.io.File;
import java.io.FileReader;
import java.io.FileWriter;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.PreparedStatement;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.sql.Statement;
import java.util.Properties;
import java.util.Vector;
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
import java.util.Hashtable;
import java.util.logging.FileHandler;
import java.util.logging.Formatter;
import java.util.logging.Handler;
import java.util.logging.Level;
import java.util.logging.Logger;
import java.util.logging.SimpleFormatter;

/**
 *
 * @author mrrodriguez
 */
public final class Llam_mail_WS_sondeos {

	@SuppressWarnings("UseOfObsoleteCollectionType")
	Vector dir_no_encontradas = new Vector();
	Connection con_1;
	ResultSet rs = null;
	String user;
	String pass;
	String url;
	String msgCancela = "";
	String callfunction = null;
	String Destinatarios = null;
	String campostr[];
	// Logger
	private static final Logger Log = Logger.getLogger(Llam_mail_WS_sondeos.class.getName());
	private static final String LogFileName = "batchLog.txt";
	private static final String FileReqName = "Request.txt";
	// DEV
	/*
	 * String cadena =
	 * "C:\\Users\\mrrodriguez\\Documents\\NetBeansProjects\\Llam_mail_WS_sondeos\\src\\base_pro.properties";
	 * String cadenaConstants =
	 * "C:\\Users\\mrrodriguez\\Documents\\NetBeansProjects\\Llam_mail_WS_sondeos\\src\\constants.txt";
	 * String ruta =
	 * "C:\\Users\\mrrodriguez\\Documents\\NetBeansProjects\\Llam_mail_WS_sondeos\\Request.txt";
	 * private static final String LogsDir =
	 * "C:\\Users\\mrrodriguez\\Documents\\NetBeansProjects\\Llam_mail_WS_sondeos\\logs";
	 */

	// DIRECTORIOS RELATIVOS PARA PRODUCCION
        //private static final String base= "/ias/Llam_mail_WS_sondeos/"; 
        private static final String base= "H:/NetBeansProjects/Llam_mail_WS_sondeos/Llam_mail_WS_sondeos/";
//	private static final String base = "C:/STSWorkspace/Llam_mail_WS_sondeos/";
	String cadena = base + "base_pro.properties";
	String cadenaConstants = base + "constants.txt";
	String ruta = base + "Request.txt";
	private static final String LogsDir = base + "logs";

	final String pathbase = "connectionUrl";
	final String userbase = "DBUsuarioOPE";
	final String passbase = "DBClaveOPE";
	final String separador = "/";
	final String separadorConst = "=";

	private static String driverClass = null;
	private static String buscarErrorDNS = null;
	private static String buscarErrorUSER = null;
	private static String buscarErrorEstado = null;
	private static String buscarErrorFound = null;
	private static String buscarErrorGateway = null;
	private static String buscarErrorHost = null;
	private static String buscarDOBLEINV = null;
	private static String IniciandoMsj = null;
	private static String FinalizandoMsj = null;
	private static String FinalizadoConErrorMsj = null;
	private static String ErrorMsj = null;
	private static String ErrorDriverMsj = null;
	private static String ErrorUserBDMsj = null;
	private static String ErrorCtaBDMsj = null;
	private static String ErrorlimConexionBDMsj = null;
	private static String ErrorpassBDMsj = null;
	private static String ErrorconexionBDMsj = null;
	private static String InicioDNSMsj = null;
	private static String FinalizandoDNSMsj = null;
	private static String InicioUSERMsj = null;
	private static String FinalizandoUSERMsj = null;
	private static String InicioEstadoMsj = null;
	private static String FinalizandoEstadoMsj = null;
	private static String InicioFoundMsj = null;
	private static String FinalizandoFoundMsj = null;
	private static String InicioGateMsj = null;
	private static String FinalizandoGateMsj = null;
	private static String InicioHostMsj = null;
	private static String FinalizandoHostMsj = null;
	private static String InicioDOBLEINVMsj = null;
	private static String FinDOBLEINVMsj = null;
	private static String ErrorInformadoMsj = null;
	private static String IniciowriteMsj = null;
	private static String FinalizandowriteMsj = null;
	private static String conexionBDMsj = null;
	private static String SinDatosDNSMsj = null;
	private static String SinDatosUSERMsj = null;
	private static String SinDatosEstadoMsj = null;
	private static String SinDatosFoundMsj = null;
	private static String SinDatosGateMsj = null;
	private static String SinDatosHostMsj = null;
	private static String SinDatosDOBLEINVMsj = null;
	private static String ErrorRasterMsj = null;
	private static String FinalizarasterMsj = null;
	private static String ErrormailMsj = null;
	private static String IniciomessageMsj = null;
	private static String Msj = null;
	private static String Asunto = null;
	private static String FinalizandogenerarMsj = null;
	private static String ErrorgenerarMsj = null;
	private static String FinalizandogenOKMsj = null;
	private static String ErrorNoMsj = null;
	private static String ErrormailsendMsj = null;
	private static String FinalizandomailMsj = null;
	private static String ErrormailendMsj = null;
	private static String Envia = null;
	private static String Lista1 = null;
	private static String Lista2 = null;
	private static String ColRequest = null;
	private static String ColPD = null;
	private static String ColCT = null;
	private static String ColResponse = null;
	private static String sqlDNS = null;
	private static String sqlUSER = null;
	private static String sqlESTADO = null;
	private static String sqlFOUND = null;
	private static String sqlGATE = null;
	private static String sqlHOST = null;
	private static String DOBLEINVMsj = null;
	private static String sqlFAIL = null;
	private static String sqlHORA = null;
	private static final String FAILMsj = "Este es un Mail automatico:   El informe CDR del dia anterior presento casos con devolucion FAIL superiores al 15%. Favor verificar su causa.";
	private static String hora = null, minuto = null;
	double total, totalfail;
	boolean flag = false;

	private static void inicializarLog() throws IOException {
		Handler fileHandler;
		Formatter simpleFormatter;
		fileHandler = new FileHandler(LogsDir + "/" + LogFileName, true);
		simpleFormatter = new SimpleFormatter();
		Log.addHandler(fileHandler);
		fileHandler.setLevel(Level.ALL);
		Log.setLevel(Level.ALL);
		fileHandler.setFormatter(simpleFormatter);
	}

	@SuppressWarnings("CallToPrintStackTrace")
	public void propetiesread() throws Exception {
		FileReader fr = null;
		@SuppressWarnings("UnusedAssignment")
		BufferedReader br = null;

		try {
			fr = new FileReader(cadena);
			br = new BufferedReader(fr);
			String linea;

			while ((linea = br.readLine()) != null) {
				String campos[] = linea.split(separador, 2);
				String str = campos[0];
				switch (str) {
				case pathbase:
					url = campos[1];
					break;
				case userbase:
					user = campos[1];
					break;
				case passbase:
					pass = campos[1];
					break;
				}
			}
		} catch (Exception e) {
			e.printStackTrace();
		} finally {
			if (null != fr) {
				fr.close();
			}
		}
	}

	@SuppressWarnings("CallToPrintStackTrace")
	public void readConstants() throws Exception {
		FileReader fr = null;
		@SuppressWarnings("UnusedAssignment")
		BufferedReader br = null;

		try {
			fr = new FileReader(cadenaConstants);
			br = new BufferedReader(fr);
			String linea;

			while ((linea = br.readLine()) != null) {
				String campos[] = linea.split(separador);
				String str = campos[0];
				switch (str) {
				case "1":
					campostr = campos[1].split(separadorConst);
					Asunto = campostr[1];
					break;
				case "2":
					campostr = campos[1].split(separadorConst);
					ColRequest = campostr[1];
					break;
				case "3":
					campostr = campos[1].split(separadorConst);
					ColPD = campostr[1];
					break;
				case "4":
					campostr = campos[1].split(separadorConst);
					ColCT = campostr[1];
					break;
				case "5":
					campostr = campos[1].split(separadorConst);
					ColResponse = campostr[1];
					break;
				case "6":
					campostr = campos[1].split(separadorConst);
					conexionBDMsj = campostr[1];
					break;
				case "7":
					campostr = campos[1].split(separadorConst);
					ErrorMsj = campostr[1];
					break;
				case "8":
					campostr = campos[1].split(separadorConst);
					ErrorDriverMsj = campostr[1];
					break;
				case "9":
					campostr = campos[1].split(separadorConst);
					ErrorUserBDMsj = campostr[1];
					break;
				case "10":
					campostr = campos[1].split(separadorConst);
					ErrorCtaBDMsj = campostr[1];
					break;
				case "11":
					campostr = campos[1].split(separadorConst);
					ErrorlimConexionBDMsj = campostr[1];
					break;
				case "12":
					campostr = campos[1].split(separadorConst);
					ErrorpassBDMsj = campostr[1];
					break;
				case "13":
					campostr = campos[1].split(separadorConst);
					ErrorconexionBDMsj = campostr[1];
					break;
				case "14":
					campostr = campos[1].split(separadorConst);
					ErrorRasterMsj = campostr[1];
					break;
				case "15":
					campostr = campos[1].split(separadorConst);
					ErrormailMsj = campostr[1];
					break;
				case "16":
					campostr = campos[1].split(separadorConst);
					ErrorgenerarMsj = campostr[1];
					break;
				case "17":
					campostr = campos[1].split(separadorConst);
					ErrorNoMsj = campostr[1];
					break;
				case "18":
					campostr = campos[1].split(separadorConst);
					ErrormailsendMsj = campostr[1];
					break;
				case "19":
					campostr = campos[1].split(separadorConst);
					Envia = campostr[1];
					break;
				case "20":
					campostr = campos[1].split(separadorConst);
					ErrorInformadoMsj = campostr[1];
					break;
				case "21":
					campostr = campos[1].split(separadorConst);
					ErrormailendMsj = campostr[1];
					break;
				case "22":
					campostr = campos[1].split(separadorConst);
					FinalizandoMsj = campostr[1];
					break;
				case "23":
					campostr = campos[1].split(separadorConst);
					FinalizadoConErrorMsj = campostr[1];
					break;
				case "24":
					campostr = campos[1].split(separadorConst);
					FinalizandoDNSMsj = campostr[1];
					break;
				case "25":
					campostr = campos[1].split(separadorConst);
					FinalizandoUSERMsj = campostr[1];
					break;
				case "26":
					campostr = campos[1].split(separadorConst);
					FinalizandoEstadoMsj = campostr[1];
					break;
				case "27":
					campostr = campos[1].split(separadorConst);
					FinalizandoFoundMsj = campostr[1];
					break;
				case "28":
					campostr = campos[1].split(separadorConst);
					FinalizandoGateMsj = campostr[1];
					break;
				case "29":
					campostr = campos[1].split(separadorConst);
					FinalizandoHostMsj = campostr[1];
					break;
				case "30":
					campostr = campos[1].split(separadorConst);
					FinalizandowriteMsj = campostr[1];
					break;
				case "31":
					campostr = campos[1].split(separadorConst);
					FinalizarasterMsj = campostr[1];
					break;
				case "32":
					campostr = campos[1].split(separadorConst);
					FinalizandogenerarMsj = campostr[1];
					break;
				case "33":
					campostr = campos[1].split(separadorConst);
					FinalizandogenOKMsj = campostr[1];
					break;
				case "34":
					campostr = campos[1].split(separadorConst);
					FinalizandomailMsj = campostr[1];
					break;
				case "35":
					campostr = campos[1].split(separadorConst);
					IniciandoMsj = campostr[1];
					break;
				case "36":
					campostr = campos[1].split(separadorConst);
					InicioDNSMsj = campostr[1];
					break;
				case "37":
					campostr = campos[1].split(separadorConst);
					InicioUSERMsj = campostr[1];
					break;
				case "38":
					campostr = campos[1].split(separadorConst);
					InicioEstadoMsj = campostr[1];
					break;
				case "39":
					campostr = campos[1].split(separadorConst);
					InicioFoundMsj = campostr[1];
					break;
				case "40":
					campostr = campos[1].split(separadorConst);
					InicioGateMsj = campostr[1];
					break;
				case "41":
					campostr = campos[1].split(separadorConst);
					InicioHostMsj = campostr[1];
					break;
				case "42":
					campostr = campos[1].split(separadorConst);
					IniciowriteMsj = campostr[1];
					break;
				case "43":
					campostr = campos[1].split(separadorConst);
					IniciomessageMsj = campostr[1];
					break;
				case "44":
					campostr = campos[1].split(separadorConst);
					Lista1 = campostr[1];
					break;
				case "45":
					campostr = campos[1].split(separadorConst);
					Lista2 = campostr[1];
					break;
				case "46":
					campostr = campos[1].split(separadorConst);
					Msj = campostr[1];
					break;
				case "47":
					campostr = campos[1].split(separadorConst);
					SinDatosDNSMsj = campostr[1];
					break;
				case "48":
					campostr = campos[1].split(separadorConst);
					SinDatosUSERMsj = campostr[1];
					break;
				case "49":
					campostr = campos[1].split(separadorConst);
					SinDatosEstadoMsj = campostr[1];
					break;
				case "50":
					campostr = campos[1].split(separadorConst);
					SinDatosFoundMsj = campostr[1];
					break;
				case "51":
					campostr = campos[1].split(separadorConst);
					SinDatosGateMsj = campostr[1];
					break;
				case "52":
					campostr = campos[1].split(separadorConst);
					SinDatosHostMsj = campostr[1];
					break;
				case "53":
					campostr = campos[1].split(separadorConst);
					driverClass = campostr[1];
					break;
				case "54":
					campostr = campos[1].split(separadorConst);
					buscarErrorDNS = campostr[1];
					break;
				case "55":
					campostr = campos[1].split(separadorConst);
					buscarErrorUSER = campostr[1];
					break;
				case "56":
					campostr = campos[1].split(separadorConst);
					buscarErrorEstado = campostr[1];
					break;
				case "57":
					campostr = campos[1].split(separadorConst);
					buscarErrorFound = campostr[1];
					break;
				case "58":
					campostr = campos[1].split(separadorConst);
					buscarErrorGateway = campostr[1];
					break;
				case "59":
					campostr = campos[1].split(separadorConst);
					buscarErrorHost = campostr[1];
					break;
				case "60":
					campostr = campos[1].split(separadorConst);
					sqlDNS = campostr[1];
					break;
				case "61":
					campostr = campos[1].split(separadorConst);
					sqlUSER = campostr[1];
					break;
				case "62":
					campostr = campos[1].split(separadorConst);
					sqlESTADO = campostr[1];
					break;
				case "63":
					campostr = campos[1].split(separadorConst);
					sqlFOUND = campostr[1];
					break;
				case "64":
					campostr = campos[1].split(separadorConst);
					sqlGATE = campostr[1];
					break;
				case "65":
					campostr = campos[1].split(separadorConst);
					sqlHOST = campostr[1];
					break;
				case "68":
					campostr = campos[1].split(separadorConst);
					buscarDOBLEINV = campostr[1];
					break;
				case "69":
					campostr = campos[1].split(separadorConst);
					SinDatosDOBLEINVMsj = campostr[1];
					break;
				case "70":
					campostr = campos[1].split(separadorConst);
					InicioDOBLEINVMsj = campostr[1];
					break;
				case "71":
					campostr = campos[1].split(separadorConst);
					FinDOBLEINVMsj = campostr[1];
					break;
				case "72":
					campostr = campos[1].split(separadorConst);
					DOBLEINVMsj = campostr[1];
					break;
				case "73":
					campostr = campos[1].split(separadorConst);
					sqlFAIL = campostr[1];
					break;
				case "74":
					campostr = campos[1].split(separadorConst);
					sqlHORA = campostr[1];
					break;
				}
			}
		} catch (Exception e) {
			e.printStackTrace();
		} finally {
			if (null != fr) {
				fr.close();
			}
		}
	}

	public static String decoBase64Apache(String s_encoded) throws Exception {
		try {
			byte[] b_encoded = s_encoded.getBytes();
			byte[] b_decoded = Base64.decodeBase64(b_encoded);
			return new String(b_decoded);
		} catch (Exception e) {
			throw e;
		}
	}

	public void Oracle_connect() throws Exception {
		try {
			Class.forName("oracle.jdbc.driver.OracleDriver").newInstance();
		} catch (ClassNotFoundException | InstantiationException | IllegalAccessException e) {
			Log.log(Level.INFO, ErrorDriverMsj + driverClass + ErrorMsj, e);
			throw e;
		}
		try {
			pass = decoBase64Apache(pass);
			con_1 = DriverManager.getConnection(url, user, pass);
		} catch (SQLException e) {
			switch (e.getErrorCode()) {
			case 1017:// USUARIO O CLAVE IVALIDO; LOGON DENIED
				Log.log(Level.INFO, ErrorUserBDMsj);
				break;
			case 28000:// CUENTA BLOQUEADA
				Log.log(Level.INFO, ErrorCtaBDMsj);
				break;
			case 2391:// EXCEDIDO EN CONEXIONES
				Log.log(Level.INFO, ErrorlimConexionBDMsj);
				break;
			case 28001: // CLAVE EXPIRADA
				Log.log(Level.INFO, ErrorpassBDMsj);
				break;
			default:
				Log.log(Level.INFO, ErrorconexionBDMsj, e);
			}
			throw e;
		}
	}

	ResultSet buscarErrorDNS() throws Exception {
		Statement sentencia;
		try {
			sentencia = con_1.createStatement(ResultSet.TYPE_SCROLL_INSENSITIVE, ResultSet.CONCUR_READ_ONLY);
			Log.log(Level.INFO, InicioDNSMsj);
			rs = sentencia.executeQuery(sqlDNS);
			Log.log(Level.INFO, FinalizandoDNSMsj);
			return rs;
		} catch (Exception e) {
			throw e;
		}
	}

	ResultSet buscarErrorUSER() throws Exception {
		Statement sentencia;
		try {
			sentencia = con_1.createStatement(ResultSet.TYPE_SCROLL_INSENSITIVE, ResultSet.CONCUR_READ_ONLY);
			Log.log(Level.INFO, InicioUSERMsj);
			rs = sentencia.executeQuery(sqlUSER);
			Log.log(Level.INFO, FinalizandoUSERMsj);
			return rs;
		} catch (Exception e) {
			throw e;
		}
	}

	ResultSet buscarErrorEstado() throws Exception {
		Statement sentencia;
		try {
			sentencia = con_1.createStatement(ResultSet.TYPE_SCROLL_INSENSITIVE, ResultSet.CONCUR_READ_ONLY);
			Log.log(Level.INFO, InicioEstadoMsj);
			rs = sentencia.executeQuery(sqlESTADO);
			Log.log(Level.INFO, FinalizandoEstadoMsj);
			return rs;
		} catch (Exception e) {
			throw e;
		}
	}

	ResultSet buscarErrorFound() throws Exception {
		Statement sentencia;
		try {
			sentencia = con_1.createStatement(ResultSet.TYPE_SCROLL_INSENSITIVE, ResultSet.CONCUR_READ_ONLY);
			Log.log(Level.INFO, InicioFoundMsj);
			rs = sentencia.executeQuery(sqlFOUND);
			Log.log(Level.INFO, FinalizandoFoundMsj);
			return rs;
		} catch (Exception e) {
			throw e;
		}
	}

	ResultSet buscarErrorGateway() throws Exception {
		Statement sentencia;
		try {
			sentencia = con_1.createStatement(ResultSet.TYPE_SCROLL_INSENSITIVE, ResultSet.CONCUR_READ_ONLY);
			Log.log(Level.INFO, InicioGateMsj);
			rs = sentencia.executeQuery(sqlGATE);
			Log.log(Level.INFO, FinalizandoGateMsj);
			return rs;
		} catch (Exception e) {
			throw e;
		}
	}

	ResultSet buscarErrorHost() throws Exception {
		Statement sentencia;
		try {
			sentencia = con_1.createStatement(ResultSet.TYPE_SCROLL_INSENSITIVE, ResultSet.CONCUR_READ_ONLY);
			Log.log(Level.INFO, InicioHostMsj);
			rs = sentencia.executeQuery(sqlHOST);
			Log.log(Level.INFO, FinalizandoHostMsj);
			return rs;
		} catch (Exception e) {
			throw e;
		}
	}

	ResultSet buscarDOBLEINV() throws Exception {
		Statement sentencia;
		try {
			sentencia = con_1.createStatement(ResultSet.TYPE_SCROLL_INSENSITIVE, ResultSet.CONCUR_READ_ONLY);
			Log.log(Level.INFO, InicioDOBLEINVMsj);
			rs = sentencia.executeQuery(
					"select ct, pedido_or_doc,count(*) from NEXUS_GIS.LLAM_MT_LOG where trunc(fecha_invws)= trunc(sysdate) and ((sysdate - fecha_invws) * 24 * 60)<=60 and operacion!='suspension' group by ct, pedido_or_doc having count(*)>1 order by 3 desc");
			Log.log(Level.INFO, FinDOBLEINVMsj);
			return rs;
		} catch (Exception e) {
			throw e;
		}
	}

	ResultSet buscarCDRFAIL() throws Exception {
		Statement sentencia;
		try {
			sentencia = con_1.createStatement(ResultSet.TYPE_SCROLL_INSENSITIVE, ResultSet.CONCUR_READ_ONLY);
			Log.log(Level.INFO, "Inicio la extraccion de datos: buscarCDRFAIL");
			rs = sentencia.executeQuery(
					"select estado_llamada, count(*) as cantidad from NEXUS_GIS.LLAM_CDR_DEV_SONDEOS where trunc(fecha_insert) = trunc(sysdate) group by estado_llamada");
			Log.log(Level.INFO, "Finaliza la extraccion de datos: buscarCDRFAIL");
			return rs;
		} catch (Exception e) {
			throw e;
		}
	}

	ResultSet buscarhora() throws Exception {
		Statement sentencia;
		try {
			sentencia = con_1.createStatement(ResultSet.TYPE_SCROLL_INSENSITIVE, ResultSet.CONCUR_READ_ONLY);
			Log.log(Level.INFO, "Inicio la extraccion de datos: buscarhora");
			rs = sentencia.executeQuery(sqlHORA);
			Log.log(Level.INFO, "Finaliza la extraccion de datos: buscarhora");
			return rs;
		} catch (Exception e) {
			throw e;
		}
	}

	public void ErrorInformado(String ped_or_doc, String ct) throws Exception {
		try {
			con_1.setAutoCommit(false);

			String sql = "UPDATE NEXUS_GIS.LLAM_WSALARMA_LOG" + " SET INFORMADO='SI',FECHA_INFORMADO=SYSDATE "
					+ " WHERE PEDIDO_OR_DOC = ? AND CT = ?";

			try (PreparedStatement ps = con_1.prepareStatement(sql)) {
				ps.setString(1, ped_or_doc);
				ps.setString(2, ct);
				ps.executeUpdate();
			}
			con_1.commit();
		} catch (Exception e) {
			Log.log(Level.INFO, ErrorInformadoMsj, e);
			con_1.rollback();
			con_1.setAutoCommit(true);
			throw e;
		} finally {
			con_1.setAutoCommit(true);
		}
	}

	public void filewrite() throws Exception {

		File archivo = new File(ruta);
		BufferedWriter bw;
		if (archivo.exists()) {
			Log.log(Level.INFO, IniciowriteMsj);
			bw = new BufferedWriter(new FileWriter(archivo));
			rs.beforeFirst();
			while (rs.next()) {
				String linea = rs.getString(ColRequest);
				String ped_or_doc = rs.getString(ColPD);
				String ct = rs.getString(ColCT);
				bw.write(linea + "\n");
				ErrorInformado(ped_or_doc, ct);
			}
			Log.log(Level.INFO, FinalizandowriteMsj);
		} else {
			bw = new BufferedWriter(new FileWriter(archivo));
			bw.write("");
		}
		bw.close();

	}

	public void filewriteDOBLEWRITE() throws Exception {

		File archivo = new File(ruta);
		BufferedWriter bw;
		if (archivo.exists()) {
			Log.log(Level.INFO, IniciowriteMsj);
			bw = new BufferedWriter(new FileWriter(archivo));
			rs.beforeFirst();
			while (rs.next()) {
				String ped_or_doc = rs.getString(ColPD);
				String ct = rs.getString(ColCT);
				bw.write(ped_or_doc + ";" + ct + "\n");
			}
			Log.log(Level.INFO, FinalizandowriteMsj);
		} else {
			bw = new BufferedWriter(new FileWriter(archivo));
			bw.write("");
		}
		bw.close();

	}

	public void raster() throws Exception {
		try {
			boolean hayDatos;
			readConstants();
			inicializarLog();
			Log.log(Level.INFO, IniciandoMsj);
			propetiesread();
			Oracle_connect();
			Log.log(Level.INFO, conexionBDMsj);
			buscarErrorDNS();
			hayDatos = rs.first();
			if (hayDatos) {
				String error = rs.getString(ColResponse);
				rs.isBeforeFirst();
				callfunction = buscarErrorDNS;
				generar(error, callfunction);
				rs = null;
			} else {
				Log.log(Level.INFO, SinDatosDNSMsj);
			}
			buscarErrorUSER();
			hayDatos = rs.first();
			if (hayDatos) {
				String error = rs.getString(ColResponse);
				rs.isBeforeFirst();
				callfunction = buscarErrorUSER;
				generar(error, callfunction);
				rs = null;
			} else {
				Log.log(Level.INFO, SinDatosUSERMsj);
			}
			buscarErrorEstado();
			hayDatos = rs.first();
			if (hayDatos) {
				String error = rs.getString(ColResponse);
				rs.isBeforeFirst();
				callfunction = buscarErrorEstado;
				generar(error, callfunction);
				rs = null;
			} else {
				Log.log(Level.INFO, SinDatosEstadoMsj);
			}
			buscarErrorFound();
			hayDatos = rs.first();
			if (hayDatos) {
				String error = rs.getString(ColResponse);
				rs.isBeforeFirst();
				callfunction = buscarErrorFound;
				generar(error, callfunction);
				rs = null;
			} else {
				Log.log(Level.INFO, SinDatosFoundMsj);
			}
			buscarErrorGateway();
			hayDatos = rs.first();
			if (hayDatos) {
				String error = rs.getString(ColResponse);
				rs.isBeforeFirst();
				callfunction = buscarErrorGateway;
				generar(error, callfunction);
				rs = null;
			} else {
				Log.log(Level.INFO, SinDatosGateMsj);
			}
			buscarErrorHost();
			hayDatos = rs.first();
			if (hayDatos) {
				String error = rs.getString(ColResponse);
				rs.isBeforeFirst();
				callfunction = buscarErrorHost;
				generar(error, callfunction);
				rs = null;
			} else {
				Log.log(Level.INFO, SinDatosHostMsj);
			}
			buscarDOBLEINV();
			hayDatos = rs.first();
			if (hayDatos) {
				String error = rs.getString(ColPD);
				rs.isBeforeFirst();
				callfunction = buscarDOBLEINV;
				generar(error, callfunction);
				rs = null;
			} else {
				Log.log(Level.INFO, SinDatosDOBLEINVMsj);
			}
			buscarhora();
			while (rs.next()) {
				hora = rs.getString("HORA");
				minuto = rs.getString("MINUTO");
			}
			if ((hora.equals("12"))) {
				if ((Integer.parseInt(minuto) >= 0) && (Integer.parseInt(minuto) <= 11)) {
					buscarCDRFAIL();
					while (rs.next()) {
						String estado_llamada = rs.getString("ESTADO_LLAMADA");
						if (estado_llamada.trim().equals("FAIL")) {
							totalfail = Double.parseDouble(rs.getString("CANTIDAD"));
						}
						total = total + Double.parseDouble(rs.getString("CANTIDAD"));
					}
					total = (totalfail / total) * 100;
					if (total > 15) {
						Msj = FAILMsj;
						callfunction = buscarErrorUSER;
						generar(Msj, callfunction);
					}
				}
			}
		} catch (Exception e) {
			e.printStackTrace();
			con_1.close();
			Log.log(Level.INFO, ErrorRasterMsj, e);
			throw e;
		} finally {
			try {
				Log.log(Level.INFO, FinalizarasterMsj);
				con_1.close();
			} catch (Exception e1) {
			}
		}
	}

//------------------------------------------------------------------------------------------------------
	public Llam_mail_WS_sondeos() throws Exception {
		try {
			raster();
		} catch (Exception e) {
			Log.log(Level.INFO, ErrormailMsj, e);
			throw e;
		}
	}

//------------------------------------------------------------------------------------------------------
	@SuppressWarnings("IndexOfReplaceableByContains")
	private void generar(String error, String callfunction) throws Exception {

		@SuppressWarnings("UseOfObsoleteCollectionType")
		Hashtable<String, String> hst_Mail = new Hashtable<>();

		String HTML_Estucture;
		@SuppressWarnings("UnusedAssignment")
		boolean enviarMail = false;
		try {
			enviarMail = true;
			Log.log(Level.INFO, IniciomessageMsj);

			if (callfunction.equals(buscarErrorHost) || (callfunction.equals(buscarDOBLEINV))) {
				Destinatarios = Lista1;
				if (callfunction.equals(buscarDOBLEINV)) {
					error = DOBLEINVMsj;
				}
			} else {
				Destinatarios = Lista2;
			}

			hst_Mail.put("mailHost", "mail.edenor");
			hst_Mail.put("DE", Envia);
			hst_Mail.put("PARA", Destinatarios);
			// hst_Mail.put("PARA" , "mrrodriguez@edenor.com");
			hst_Mail.put("ASUNTO", Asunto);

			HTML_Estucture = "<html>";
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
			HTML_Estucture += "</br></br>";
			HTML_Estucture += " <tr height=21 style='height:15.75pt'>";
			HTML_Estucture += "  <td height=21 class=xl7730982 colspan=9 width=494 style='height:15.75pt;width:600pt'>" + Msj
					+ "</td>";
			HTML_Estucture += " </tr>";
			HTML_Estucture += "</br></br>";
			HTML_Estucture += "</br></br>";
			HTML_Estucture += " <tr height=21 style='height:15.75pt'>";
			HTML_Estucture += "  <td height=21 class=xl7730982 colspan=9 width=494 style='height:15.75pt;width:600pt'></td>";
			HTML_Estucture += " </tr>";
			HTML_Estucture += "</br></br>";
			HTML_Estucture += "</br></br>";

			if (Msj.indexOf("CDR") > -1) {
			} else {
				HTML_Estucture += " <tr height=21 style='height:15.75pt'>";
				HTML_Estucture += "  <td height=21 class=xl7730982 colspan=9 width=494 style='height:15.75pt;width:600pt'>"
						+ error + "</td>";
				HTML_Estucture += " </tr>";
				HTML_Estucture += "</br></br>";
				HTML_Estucture += "</br></br>";
				HTML_Estucture += " <tr height=21 style='height:15.75pt'>";
				HTML_Estucture += "  <td height=21 class=xl7730982 colspan=9 width=494 style='height:15.75pt;width:600pt'></td>";
				HTML_Estucture += " </tr>";
				HTML_Estucture += "</br></br>";
				HTML_Estucture += "</br></br>";
				HTML_Estucture += " <tr height=21 style='height:15.75pt'>";
				HTML_Estucture += "  <td height=21 class=xl7730982 colspan=9 width=494 style='height:15.75pt;width:600pt'>En el adjunto encuentran los Request que se intentó hacer en la última hora    con esta causa</td>";
				HTML_Estucture += " </tr>";
			}
			HTML_Estucture += "</table>";
			HTML_Estucture += "</div>";
			HTML_Estucture += "</body>";
			HTML_Estucture += "</html>";
			Log.log(Level.INFO, FinalizandogenerarMsj);
			if (Msj.indexOf("CDR") > -1) {
				flag = false;
			} else {
				if (callfunction.equals(buscarDOBLEINV)) {
					filewriteDOBLEWRITE();
				} else {
					filewrite();
				}
				flag = true;
			}

			hst_Mail.put("CUERPO", HTML_Estucture);
			if (enviarMail) {
				mailSender(hst_Mail, flag);
			}

		} catch (Exception e) {
			Log.log(Level.INFO, ErrorgenerarMsj, e);
			msgCancela = e.toString();
			hst_Mail.put("CUERPO", msgCancela);
			mailSender(hst_Mail, flag);
			throw e;
		} finally {
			try {
				Log.log(Level.INFO, FinalizandogenOKMsj);
			} catch (Exception e1) {
			}
		}
	}

//------------------------------------------------------------------------------------------------------
	public void mailSender(@SuppressWarnings("UseOfObsoleteCollectionType") Hashtable<String, String> hst_values_mail,
			boolean flag) throws Exception {
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

			InternetAddress[] paraArray;
			paraArray = InternetAddress.parse(hst_values_mail.get("PARA"));
			System.out.println("Destinatarios: " + paraArray.toString());
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
			if (flag) {
				BodyPart adjunto = new MimeBodyPart();
				adjunto.setDataHandler(new DataHandler(new FileDataSource(FileReqName)));
				adjunto.setFileName(FileReqName);
				multiParte.addBodyPart(adjunto);
			}

			BodyPart texto = new MimeBodyPart();
			texto.setDataHandler(new DataHandler(new HTMLDataSource(hst_values_mail.get("CUERPO"))));
			multiParte.addBodyPart(texto);

			msg.setContent(multiParte);

			int i, j, k, total;
			total = paraArray.length;
			if (ccArray != null)
				total += ccArray.length;
			if (bccArray != null)
				total += bccArray.length;

			InternetAddress[] address = new InternetAddress[total];

			for (i = 0; i < paraArray.length; i++)
				address[i] = paraArray[i];
			if (ccArray != null)
				for (j = 0; j < ccArray.length; j++) {
					address[i] = ccArray[j];
					i++;
				}
			if (bccArray != null)
				for (k = 0; k < bccArray.length; k++) {
					address[i] = bccArray[k];
					i++;
				}

			Transport transporte = session.getTransport(address[0]);
			transporte.connect();
			transporte.sendMessage(msg, address);

		} catch (SendFailedException e) {
			Address[] listaInval = e.getInvalidAddresses();
			for (Address listaInval1 : listaInval) {
				dir_no_encontradas.add(listaInval1.toString());
				Log.log(Level.INFO, ErrorNoMsj, listaInval1.toString());
			}
		} catch (MessagingException e) {
			Log.log(Level.INFO, ErrormailsendMsj, e);
			throw e;
		}
	}

//------------------------------------------------------------------------------------------------------
	static class HTMLDataSource implements DataSource {
		private final String html;

		public HTMLDataSource(String htmlString) {
			html = htmlString;
		}

		@Override
		public InputStream getInputStream() throws IOException {
			if (html == null)
				throw new IOException("Null HTML");
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

//------------------------------------------------------------------------------------------------------
	public static void main(String[] args) {
		try {
			Llam_mail_WS_sondeos g = new Llam_mail_WS_sondeos();
			Log.log(Level.INFO, FinalizandomailMsj);
			Log.log(Level.INFO, FinalizandoMsj);
			System.exit(0);
		} catch (Exception e) {
			Log.log(Level.INFO, ErrormailendMsj, e);
			Log.log(Level.INFO, FinalizadoConErrorMsj);
			System.exit(1);
		}
	}
}
