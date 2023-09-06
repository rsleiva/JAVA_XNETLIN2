import java.io.BufferedReader;
import java.io.ByteArrayInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.io.Reader;
import java.sql.CallableStatement;
import java.sql.Clob;
import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.PreparedStatement;
import java.sql.SQLException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.Hashtable;
import java.util.Iterator;
import java.util.List;
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

//import oracle.jdbc.driver.OracleTypes;
import oracle.jdbc.OracleTypes;

final class Mail_Sender {

    String msgCancela= "";
    Vector dir_no_encontradas= new Vector();	
    Connection con_1;
    
    final String driverClass="oracle.jdbc.driver.OracleDriver";
    private Iterator<HashMap<String, String>> it;
	
	public void Oracle_connect() throws Exception{
	try{
		Class.forName(driverClass).newInstance();
	}catch (ClassNotFoundException | InstantiationException | IllegalAccessException e){
	      System.out.println("Error al cargar el driver: "+driverClass+" -Error: "+e);
		  throw e;
	}
	try{
            
            //con_1= DriverManager.getConnection( "jdbc:oracle:thin:NEXUS_ENRE/NEXUS_ENRE@tdbs6.tro.edenor:1521:GISDEV01");
            con_1= DriverManager.getConnection( "jdbc:oracle:thin:@NEXGISPR02.PRO.EDENOR:1528/gispr01s", "SVC_ORA_GIS", "jv506uzy" );
	}catch(SQLException e) {
		switch(e.getErrorCode()) {
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
	
    private List<HashMap<String, String>> strToHash(String reporte){
         
        List<HashMap<String, String>> tabla = new ArrayList<>();
        reporte = reporte.replace(",", "X");        
        String[] lineas = reporte.split("X");
        
        for (String linea : lineas){
             String[] campo = linea.split(";"); 
	
            HashMap<String, String> fila = new HashMap<>(); 
            fila.put("REGION",campo[0]);
            fila.put("ZONA",campo[1]); 
            fila.put("SUBEST",campo[2]);    
            fila.put("CLI_ACTU",campo[3]);    
            fila.put("CT_ACTU",campo[4]);    
            fila.put("DOCUMENTO",campo[5]);    
            fila.put("ALIM",campo[6]);    
            fila.put("HORA_INICIO",campo[7]);    
            fila.put("CLI_INIC",campo[8]);
            fila.put("CT_INIC",campo[9]);
            fila.put("TIPO",campo[10]);
            tabla.add(fila);        
	}
        return tabla;
    }
    
    private List<HashMap<String, String>> strToHash2(String reporte){
        
        List<HashMap<String, String>> tabla = new ArrayList<>();
         
        HashMap<String, String> fila;
        String[] lineas = reporte.split(",");
        for (String linea : lineas) {
            fila = new HashMap<>();
            String[] campo = linea.split(";");
            fila.put("REGION",campo[0]);
            fila.put("ZONA",campo[1]);
            fila.put("CLIENTES_FS_FORZ",campo[2]);
            fila.put("CLIENTES_FS_PROGR",campo[3]);
            fila.put("CLIENTES_FS_ACTUALES",campo[4]);
            tabla.add(fila);
        }
        return tabla;
    }
    
    public void raster() throws Exception{

        String  plsql;
		PreparedStatement ps_1=null;
        
	try{
	    Oracle_connect();		
			
            plsql ="{?=call NEXUS_GIS.INFORMACION_ENREMTAT.REPORT_WS_MT(?)}";
             
            CallableStatement cs = con_1.prepareCall(plsql);
            cs.registerOutParameter(1, OracleTypes.CLOB);
            cs.registerOutParameter(2, OracleTypes.CLOB);
            cs.execute();
            
            plsql=clobToStr(cs.getClob(1));
           // plsql = "SIXIMPORTANTE: El presente es un mail automático que informa situación de documentos MTAT informados al ENRE<br>Hora Extracto: 24/02/2016 12:28:21<br>XR1;CABA;VIDAL;558;2;D-16-02-057877;11445;23/02/2016 20:38:51;4755;15;F,R1;CABA;SAAVEDRA;12;2;D-16-02-058358;13752/13747;24/02/2016 03:48:28;662;12;F,R1;CABA;COLEGIALES/J. NEWBERY;1;1;D-16-02-059728;23505/4661;24/02/2016 11:58:04;2;2;F,R1;OLIVOS;BOULOGNE;681;6;D-16-02-056019;5614;24/02/2016 09:31:58;1688;14;P,R1;SAN MARTIN;ROTONDA;348;1;D-16-02-057622;26222/26227;23/02/2016 18:25:39;3929;16;F,R1;SAN MARTIN;CASEROS;1;1;D-16-02-058334;6915;24/02/2016 01:51:31;1749;9;F,R2;MATANZA;PANTANOSA;167;8;D-16-02-055919;36518;24/02/2016 09:13:13;3713;52;P,R2;MATANZA;MATANZA;1;1;D-16-02-059076;6515;24/02/2016 10:02:25;1;1;P,R2;MERLO;L. HERAS;277;1;D-16-02-058985;26921;24/02/2016 09:37:02;277;1;P,R2;MORON;S/N/CATONAS/NOGUES/S.MIGUEL/SAN MIGUEL/MORON/MUNIZ/MUÑIZ/PASO DEL REY;95504;554;D-16-02-059477;26523/26517/26522/26516/16527/26512/16525/26514/26525/16523/16513/16511/16522/25822/6724A/26528/16512/6722/25628/26527/16524B/16528/16518/15931/16526/26526/26511/16516/16517/16515/16524A/15932/26515/26513/26524/15937/26521/16521/16514;24/02/2016 11:33:21;88221;554;F,R2;MORON;HURLINGHAM/CASTELAR;34911;158;D-16-02-059495;16633/16632/16638/16634/16626/16624/16628/16623/16636/16635/16621/6111/16627/16631/16622/16625;24/02/2016 11:33:38;34911;158;F,R2;MORON;TORTUGUITAS/NOGUES/TORTUGUITA/DEL VISO/MORON/DELVISO/BENAVIDEZ;68360;427;D-16-02-059496;15944/25217/15948/15941/15911/25215/15917/25222/15912/15915/15943/25113B/25243/25216B/25112/25113A/25116/25242/15946/25244/15924/25114/25218/15913/25124/25211/15914/15918/25216/25213/25117/25111/5911/25225/25214/25115/15916/6704;24/02/2016 11:34:11;68360;427;F,R3;PILAR;TORTUGUITAS/FORD/NOGUES/TORTUGUITA/DEL VISO/S.MIGUEL/DELVISO/MUNIZ;78871;498;D-16-02-059535;25227/25217/15945/25223/25224/15911/15936/25127/25222/15948/15917/15923/15927/15934/25123/25128/15921/25116/16518/15142/15928/15925/15933/15931/25122/15924/15932/26511/25126/25218/25124/25225/25111/25216/25117/15141/15935/15937/15916/25125/25226;24/02/2016 11:34:17;76564;498;F,XR1;CABA;571;0;571,R1;OLIVOS;0;681;681,R1;SAN MARTIN;349;0;349,R2;MATANZA;0;168;168,R2;MERLO;0;277;277,R2;MORON;198775;0;198775,R3;MORENO;0;0;0,R3;PILAR;78871;0;78871,R3;SAN MIGUEL;0;0;0,R3;TIGRE;0;0;0,TOTAL; ;278566;1126;279692";
            
            String strarray[] = plsql.split("X");
            String v_text = strarray[0];
            String encabezado =strarray[1];
            String v_tabla = strarray[2];
            String v_tabla2 = strarray[3];            
           
            @SuppressWarnings("UnusedAssignment")
            boolean hayDatos= false;
            hayDatos = (v_text.equals("SI"));
            
            List<HashMap<String, String>> tabla = null;
            List<HashMap<String, String>> tabla2 = null;
            
            if(hayDatos){
                tabla = strToHash(v_tabla);
                tabla2 = strToHash2(v_tabla2);
            }

            generar(tabla, tabla2, encabezado, hayDatos);
            con_1.close();
        
	}catch(Exception e){
             con_1.close();
	     System.out.println("Error en raster()"+e);
   	     throw e;
	}finally{
	     try{
		System.out.println("Final One");
		}catch(Exception e1){}
		}
	}	
	
//------------------------------------------------------------------------------------------------------
	public Mail_Sender() throws Exception{
		try{
			raster();
		}catch(Exception e){
			System.out.println("Error en Mail_Sender() = "+e );
			throw e;
		}
	}
//------------------------------------------------------------------------------------------------------
	private void generar(List<HashMap<String, String>> tabla, 
                         List<HashMap<String, String>> tabla2, 
                         String encabezado, boolean hayDatos) throws Exception{

		Hashtable<String,String>   hst_Mail= new Hashtable<>();

		String HTML_Estucture;
                @SuppressWarnings("UnusedAssignment")
		boolean enviarMail = false;
		int i = 0;
		
		try{		    
			i=0;
			enviarMail = true;		
         
			hst_Mail.put("mailHost"  , "mail.edenor");
			hst_Mail.put("DE"        , "centrodeinformacion@edenor.com");			
                        //hst_Mail.put("DE"        , "MRRODRIGUEZ@edenor.com");			
                        //hst_Mail.put("PARA"      , "MRRODRIGUEZ@edenor.com");
                        //hst_Mail.put("PARA"      , "centrodeinformacion@edenor.com,MRRODRIGUEZ@edenor.com");
			hst_Mail.put("PARA"      , "centrodeinformacion@edenor.com,EMAGGI@edenor.com,PPEREZ@edenor.com,PMAZZA@edenor.com,ELAFUENTE@edenor.com,MRRODRIGUEZ@edenor.com");			
			hst_Mail.put("ASUNTO"    , "Información sobre situación de documentos MTAT informados al ENRE");


			HTML_Estucture  = "<html>";
			HTML_Estucture += "<head>";
			HTML_Estucture += "<style id='Mail_Styles'>";
			HTML_Estucture += "<!--table";
			HTML_Estucture += ".xl1530982";
			HTML_Estucture += "	{padding:0px; 	mso-ignore:padding; 	color:black; 	font-size:11.0pt; 	font-weight:400; 	font-style:normal; 	text-decoration:none; 	font-family:Calibri, sans-serif; 	mso-font-charset:0; 	mso-number-format:General; 	text-align:general; 	vertical-align:bottom; 	mso-background-source:auto; 	mso-pattern:auto; 	white-space:nowrap;}";
			HTML_Estucture += ".xl6530982";
			HTML_Estucture += "	{padding:0px; 	mso-ignore:padding; 	color:black; 	font-size:11.0pt; 	font-weight:700; 	font-style:normal; 	text-decoration:none; 	font-family:Calibri, sans-serif; 	mso-font-charset:0; 	mso-number-format:General; 	text-align:general; 	vertical-align:middle; 	border-top:1.0pt solid windowtext; 	border-right:.5pt solid windowtext; 	border-bottom:1.0pt solid windowtext; 	border-left:1.0pt solid windowtext; 	background:#FFC000; 	mso-pattern:black none; 	white-space:nowrap;}";
			HTML_Estucture += ".xl6630982";
			HTML_Estucture += "	{padding:5px; 	color:black; 	font-size:10pt; 	font-weight:700; 	font-style:normal; 	text-decoration:none; 	font-family:Calibri, sans-serif; 	mso-font-charset:0; 	mso-number-format:General; 	text-align:center; 	vertical-align:middle; 	border-top:1.0pt solid windowtext; 	border-right:.5pt solid windowtext; 	border-bottom:1.0pt solid windowtext; 	border-left:.5pt solid windowtext; 	background:#C5D9F1; 	mso-pattern:black none; 	white-space:nowrap;}";
			HTML_Estucture += ".xl6730982";
			HTML_Estucture += "	{padding:0px; 	mso-ignore:padding; 	color:black; 	font-size:11.0pt; 	font-weight:700; 	font-style:normal; 	text-decoration:none; 	font-family:Calibri, sans-serif; 	mso-font-charset:0; 	mso-number-format:General; 	text-align:right; 	vertical-align:middle; 	border-top:1.0pt solid windowtext; 	border-right:1.0pt solid windowtext; 	border-bottom:1.0pt solid windowtext; 	border-left:.5pt solid windowtext; 	background:#FFC000; 	mso-pattern:black none; 	white-space:nowrap;}";
			HTML_Estucture += ".xl6830982";
			HTML_Estucture += "	{padding:5px; 	mso-ignore:padding; 	color:black; 	font-size:10.0pt; 	font-weight:400; 	font-style:normal; 	text-decoration:none; 	font-family:Calibri, sans-serif; 	mso-font-charset:0; 	mso-number-format:General; 	text-align:general; 	vertical-align:middle; 	border-top:1.0pt solid windowtext; 	border-right:.5pt solid windowtext; 	border-bottom:.5pt solid windowtext; 	border-left:.5pt solid windowtext; 	mso-background-source:auto; 	mso-pattern:auto; 	white-space:nowrap;}";
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
            HTML_Estucture += "  <td height=21 class=xl7730982 colspan=9 width=494 style='height:15.75pt;width:371pt'>" + encabezado + "</td>";
            HTML_Estucture += " </tr>";
            
            if (!hayDatos){
                String sinDatosMsj = "NO HAY DATOS REPORTADOS... ";
                HTML_Estucture += "</br></br>";
                HTML_Estucture += " <tr height=21 style='height:15.75pt'>";
                HTML_Estucture += "  <td height=21 class=xl7730982 colspan=9 width=494 style='height:15.75pt;width:371pt'>" + sinDatosMsj + "</td>";
                HTML_Estucture += " </tr>";
                HTML_Estucture += "</table>";
                
                HTML_Estucture += " <br>";
                HTML_Estucture += " <hr>";
                HTML_Estucture += " <br>";
                
                HTML_Estucture += "<table border=0 cellpadding=5px cellspacing=0 width=494 style='border-collapse: collapse;table-layout:auto;width:371pt'>";
                HTML_Estucture += " <tr height=20 style='height:15.0pt'>";
                HTML_Estucture += "  <td height=20 style='height:15.0pt' align=left valign=top colspan=5 >";
                HTML_Estucture += "  <span style='mso-ignore:vglayout;position:absolute;z-index:1;margin-left:3px;margin-top:3px;width:71px;height:15px'>";
                HTML_Estucture += "  <img width=71 height=15 src='Logo_edenor.gif'  alt='DescripciÃ³n: logo de Edenor' ></span>";
                HTML_Estucture += "  </td>";
                HTML_Estucture += " </tr>";
                HTML_Estucture += "</table>";
                HTML_Estucture += "</div>";
                HTML_Estucture += "</body>";
                HTML_Estucture += "</html>";
            } else {
            
                HTML_Estucture += " <tr height=21 style='height:15.75pt'>";
                HTML_Estucture += "  <td height=21 class=xl1530982 style='height:15.75pt' colspan=5 ></td>";
                HTML_Estucture += " </tr>";
                HTML_Estucture += " <tr height=27 style='mso-height-source:userset;height:20.25pt'>";
                
                HTML_Estucture += "  <td height=27 class=xl6630982 style='height:20.25pt'>REGION</td>";
                HTML_Estucture += "  <td height=27 class=xl6630982 style='height:20.25pt'>ZONA</td>";
                HTML_Estucture += "  <td class=xl6630982 style='border-left:none'>SUBEST</td>";
                HTML_Estucture += "  <td class=xl6630982 style='border-left:none'>CLI_ACTU</td>";
                HTML_Estucture += "  <td class=xl6630982 style='border-left:none'>CT_ACTU</td>";
                HTML_Estucture += "  <td class=xl6630982 style='border-left:none'>DOCUMENTO</td>";
                HTML_Estucture += "  <td class=xl6630982 style='border-left:none'>ALIM</td>";
                HTML_Estucture += "  <td class=xl6630982 style='border-left:none'>HORA INICIO</td>";
                HTML_Estucture += "  <td class=xl6630982 style='border-left:none'>CLI_INIC</td>";
                HTML_Estucture += "  <td class=xl6630982 style='border-left:none'>CT_INIC</td>";
                HTML_Estucture += "  <td class=xl6630982 style='border-left:none'>TIPO</td>";
                            for (HashMap fila : tabla) {
                                HTML_Estucture += " <tr height=27 style='mso-height-source:userset;height:20.25pt'>";
                                HTML_Estucture += " <td height=27 class=xl7830982 style='height:20.25pt'>"+fila.get("REGION")+"</td>";
                                HTML_Estucture += " <td height=27 class=xl7830982 style='height:20.25pt'>"+fila.get("ZONA")+"</td>";
                                HTML_Estucture += "<td class=xl6830982 align=right style='border-left:none'>"+fila.get("SUBEST")+"</td>";
                                HTML_Estucture += "  <td class=xl6830982 align=right style='border-left:none'>"+fila.get("CLI_ACTU")+"</td>";
                                HTML_Estucture += "  <td class=xl6830982 align=right style='border-left:none'>"+fila.get("CT_ACTU")+"</td>";
                                HTML_Estucture += "  <td class=xl6830982 align=right style='border-left:none'>"+fila.get("DOCUMENTO")+"</td>";
                                HTML_Estucture += "  <td class=xl6830982 align=right style='border-left:none'>"+fila.get("ALIM")+"</td>";
                                HTML_Estucture += "  <td class=xl6830982 align=right style='border-left:none'>"+fila.get("HORA_INICIO")+"</td>";
                                HTML_Estucture += "  <td class=xl6830982 align=right style='border-left:none'>"+fila.get("CLI_INIC")+"</td>";
                                HTML_Estucture += "  <td class=xl6830982 align=right style='border-left:none'>"+fila.get("CT_INIC")+"</td>";
                                HTML_Estucture += "  <td class=xl6830982 align=right style='border-left:none'>"+fila.get("TIPO")+"</td>";
                                HTML_Estucture += " </tr>";
                            }

                HTML_Estucture += " <tr height=20 style='height:15.0pt'>";
                HTML_Estucture += "  <td height=20 class=xl1530982 style='height:15.0pt' colspan=5 ></td>";
                HTML_Estucture += " </tr>";
                
                HTML_Estucture += "</table>";
                
                /////la 2da tabla///
                HTML_Estucture += "<table border=0 cellpadding=0 cellspacing=0 style='border-collapse: collapse;table-layout:auto;width:150pt'>";
                HTML_Estucture += " <tr height=21 style='height:15.75pt'>";
                HTML_Estucture += "  <td height=21 class=xl1530982 style='height:15.75pt' colspan=2 ></td>";
                HTML_Estucture += " </tr>";
                HTML_Estucture += " <tr height=27 style='mso-height-source:userset;height:20.25pt'>";
                HTML_Estucture += "  <td height=27 class=xl6630982 style='height:20.25pt'>REGION</td>";
                HTML_Estucture += "  <td height=27 class=xl6630982 style='height:20.25pt'>ZONA</td>";
                HTML_Estucture += "  <td height=27 class=xl6630982 style='height:20.25pt'>CLIENTES FS FORZ.</td>";
                HTML_Estucture += "  <td height=27 class=xl6630982 style='height:20.25pt'>CLIENTES FS PROGR.</td>";
                HTML_Estucture += "  <td class=xl6630982 style='border-left:none'>CLIENTES FS ACTUALES</td>";
                HTML_Estucture += " </tr>";
                
                it = tabla2.iterator();
                while (it.hasNext()){
                    HashMap fila = (HashMap)it.next();
                    HTML_Estucture += " <tr height=27 style='mso-height-source:userset;height:20.25pt'>";
                    	
                    if (fila.get("REGION").equals("TOTAL")){
                        HTML_Estucture += " <td height=27 class=xl7830982 style='height:20.25pt; font-weight: bold'>"+fila.get("REGION")+"</td>";
                        HTML_Estucture += " <td height=27 class=xl7830982 style='height:20.25pt; font-weight: bold'>"+fila.get("ZONA")+"</td>";
                    } else {
                        HTML_Estucture += " <td height=27 class=xl7830982 style='height:20.25pt'>"+fila.get("REGION")+"</td>"; 
                        HTML_Estucture += " <td height=27 class=xl7830982 style='height:20.25pt'>"+fila.get("ZONA")+"</td>"; 
                    }
                    
                    HTML_Estucture += " <td class=xl6830982 align=right style='border-left:none'>"+fila.get("CLIENTES_FS_FORZ")+"</td>";
                    HTML_Estucture += " <td class=xl6830982 align=right style='border-left:none'>"+fila.get("CLIENTES_FS_PROGR")+"</td>";
                    HTML_Estucture += " <td class=xl6830982 align=right style='border-left:none'>"+fila.get("CLIENTES_FS_ACTUALES")+"</td>";
                    HTML_Estucture += " </tr>";
                }
                 HTML_Estucture += "</table>";
                /////fin 2da tabla///
                 
                HTML_Estucture += " <br>";
                HTML_Estucture += " <hr>";
                HTML_Estucture += " <br>";
                
                HTML_Estucture += "<table border=0 cellpadding=5px cellspacing=0 width=494 style='border-collapse: collapse;table-layout:auto;width:371pt'>";
                HTML_Estucture += " <tr height=20 style='height:15.0pt'>";
                HTML_Estucture += "  <td height=20 style='height:15.0pt' align=left valign=top colspan=5 >";
                HTML_Estucture += "  <span style='mso-ignore:vglayout;position:absolute;z-index:1;margin-left:3px;margin-top:3px;width:71px;height:15px'>";
                HTML_Estucture += "  <img width=71 height=15 src='Logo_edenor.gif'  alt='DescripciÃ³n: logo de Edenor' ></span>";
                HTML_Estucture += "  </td>";
                HTML_Estucture += " </tr>";
                HTML_Estucture += "</table>";
                
                HTML_Estucture += "</div>";
                HTML_Estucture += "</body>";
                HTML_Estucture += "</html>";
            }

            
            System.out.println("enviado: " +  HTML_Estucture);
            
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

public void mailSender(Hashtable<String,String> hst_values_mail) throws Exception {
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

//------------------------------------------------------------------------------------------------------

public String Minute_To_Hour(long L_Minute) throws Exception{
		try	{
			long L_Hour    = 0l;
			long L_Aux     = 0l;
			String S_out;
			try {
			   L_Hour   = L_Minute / 60l;
			   L_Aux    = L_Minute % 60l;
	      	}
	        catch (NumberFormatException e){
			    throw new Exception("Error en el String de Hora "+ e);
	        }
			if (L_Hour < 10)
			   S_out = "0" + String.valueOf(L_Hour);
			else
			   S_out = String.valueOf(L_Hour);
			S_out += ":";
			if (L_Aux < 10)
			   S_out += "0" + String.valueOf(L_Aux);
			else
			   S_out += String.valueOf(L_Aux);
						
			return  S_out;

		}catch (Exception e){
			System.out.println("Error en Minute_To_Hour() = "+e);
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

private static String clobToStr(Clob clb){
        StringBuilder sb = new StringBuilder();
        
        try{
            final Reader         reader = clb.getCharacterStream();
            try (BufferedReader br = new BufferedReader(reader)) {
                int b;
                while(-1 != (b = br.read())){
                    sb.append((char)b);
                }
            }
        }
        catch (SQLException e){
            System.out.println("RedElectrica::clobToStr: SQL. No se pudo convertir CLOB a String");
        }
        catch (IOException e){
            System.out.println("RedElectrica::clobToStr: IO. No se pudo convertir CLOB a String");
        }
        
        return sb.toString();
    }

//------------------------------------------------------------------------------------------------------
	public static void main(String[] args) {
		//String in_arg;			
			try{				
				Mail_Sender g= new Mail_Sender();
				System.out.println("\nProcedimiento MAIL terminado exitosamente");
				System.exit(0);
			}catch(Exception e){
				System.out.println(e);
				System.exit(1);
			}
		
	}
}
