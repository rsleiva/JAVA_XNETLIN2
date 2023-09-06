import java.io.ByteArrayInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.io.FileInputStream;
import java.sql.CallableStatement;
import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.PreparedStatement;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.sql.Types;

import java.util.Hashtable;
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

/**
 *
 * @author mrrodriguez
 */
final class Mail_llam_fallidos {

    String msgCancela= "";
    Vector dir_no_encontradas= new Vector();	
	
    Connection con_1;
    final String driverClass="oracle.jdbc.driver.OracleDriver";
	
	public void Oracle_connect() throws Exception{
	try{
	      Class.forName(driverClass).newInstance();
	}catch (ClassNotFoundException | InstantiationException | IllegalAccessException e){
	      System.out.println("Error al cargar el driver: "+driverClass+" -Error: "+e);
		  throw e;
	}
	try{
		String rootPath = Thread.currentThread().getContextClassLoader().getResource("").getPath();
		String propertiesPath = rootPath + "Mail_llam_fallidos.properties";
		
		Properties appProps = new Properties();
		appProps.load(new FileInputStream(propertiesPath));
		
		String dbUrl = appProps.getProperty("DB_URL");
		String dbUser = appProps.getProperty("DB_USER");
		String dbPassword = appProps.getProperty("DB_PASSWORD");

		con_1= DriverManager.getConnection( dbUrl, dbUser, dbPassword );
                // con_1= DriverManager.getConnection( "jdbc:oracle:thin:@tdbs6.tro.edenor:1521:GISDEV01", "NEXUS_GIS", "nexus_gis" );
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
	
 	public void raster() throws Exception{

        String  plsql;
		PreparedStatement ps_1=null;
       			
		try{
		    Oracle_connect();		
			
            plsql ="{?=call NEXUS_GIS.LLAM_PROCESOS.LLAM_MAIL_FALLIDOS(?)}";
            CallableStatement cs = con_1.prepareCall(plsql);
            cs.registerOutParameter(1, Types.VARCHAR);
            cs.registerOutParameter(2, Types.VARCHAR);           
            cs.execute();
            String v_text = (String)cs.getObject(1);
           
            boolean hayDatos =false;
            if (v_text!= null){
                hayDatos = (v_text.equals("MAIL"));
            }            

            if(hayDatos){
                generar();            
            }else {
                System.out.println("No hay datos");
            }
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

 	public void fn_ter() throws Exception{

        ResultSet rs = null;
        PreparedStatement ps = null;
        
        try {
             Oracle_connect();	
             
            String sql = "SELECT NEXUS_GIS.FN_TER FROM DUAL";

            ps = con_1.prepareStatement(sql);
            rs = ps.executeQuery();
 
           con_1.close();
            
        } catch (SQLException ex) {
            con_1.close();
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
	
	
	
//------------------------------------------------------------------------------------------------------
	public Mail_llam_fallidos() throws Exception{
		try{
			raster();
			fn_ter(); 
		}catch(Exception e){
			System.out.println("Error en Mail_Sender() = "+e );
			throw e;
		}
	}
//------------------------------------------------------------------------------------------------------
	private void generar() throws Exception{

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
                       //hst_Mail.put("PARA"      , "ghidalgo@externo.edenor.com,MRRODRIGUEZ@externo.edenor.com");
                       // hst_Mail.put("PARA"      , "backoffice@celinteractive.com.ar");
                        // hst_Mail.put("PARA"      , "MRRODRIGUEZ@edenor.com");
												
			hst_Mail.put("PARA"      , "EmergenciasoFaltadeSuministro@edenor.com,SUBGERENCIA_CAT@edenor.com,ITSM_Desarrollos_propios@edenor.com");			
			hst_Mail.put("ASUNTO"    , "NOTIFICACION: Existen casos de Llamadas Salientes en el procedimiento Fallidos para su gestion  ");


			HTML_Estucture  = "<html>";
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
            

                String sinDatosMsj = "IMPORTANTE: El presente es un mail automático que informa casos de Llamadas Salientes en el procedimiento Fallidos disponibles para su gestión";
    
                HTML_Estucture += "</br></br>";
                HTML_Estucture += " <tr height=21 style='height:15.75pt'>";
                HTML_Estucture += "  <td height=21 class=xl7730982 colspan=9 width=494 style='height:15.75pt;width:600pt'>" + sinDatosMsj + "</td>";
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

//------------------------------------------------------------------------------------------------------
	public static void main(String[] args) {
		String in_arg;			
			try{				
				Mail_llam_fallidos g= new Mail_llam_fallidos();
				System.out.println("\nProcedimiento MAIL terminado exitosamente");
				System.exit(0);
			}catch(Exception e){
				System.out.println(e);
				System.exit(1);
			}
		
	}
}

