import java.io.*;
import java.sql.*;
import java.util.*;
import javax.activation.DataHandler;
import javax.activation.DataSource;
import javax.activation.FileDataSource;
import javax.mail.*;
import javax.mail.internet.InternetAddress;
import javax.mail.internet.MimeBodyPart;
import javax.mail.internet.MimeMessage;
import javax.mail.internet.MimeMultipart;
import oracle.jdbc.OracleCallableStatement;
import oracle.jdbc.OracleTypes;

final class SalientesForzados {

	String msgCancela= "";
    Vector dir_no_encontradas= new Vector();	

    //Estructura HTML
    private final String fuente_html_ppio = 
          "<html>"
        + "<head>"
        + "<meta charset='UTF-8'> "
        + "<style>"
        + ".tdatos "
        + "	{border: 2px solid black;"
        + "	border-collapse:collapse;"
        + "	table-layout:fixed }"
        + ".tlogo "
        + "	{border: 0px}"
        + ".tdatos tr, .tdatos td "
        + "	{border: 1px solid black }"
        + ".resaltados"
        + "	{font-weight: bold;"
        + "	text-align:center;"
        + "	font-size: 10pt;"
        + "	border: 2px solid black;"
        + "	vertical-align:middle;"
        + "	padding: 4px;"
        + "	background:#C5D9F1 }"
        + ".datos "
        + "	{ font-size: 10pt; }"
        + "	"
        + ".datosHr "
        + "	{font-size: 8pt;"
        + "	border: 2px solid black;"
        + "	vertical-align:middle;"
        + "	text-align:center;"
        + "	padding: 4px;"
        + "	width: 20px }"
        + ".turnos "
        + "	{font-size: 8pt;"
        + "	border: 1px solid black;"
        + "	vertical-align:middle;"
        + "	text-align:center;"
        + "	padding: 2px;"
        + "	background:#C5D9F1;"
        + "	width: 20px }"
        + ".turnosInferior "
        + "	{font-size: 8pt;"
        + " border-bottom: 2px solid black;"
        + "	vertical-align:middle;"
        + "	text-align:center;"
        + "	padding: 2px;"
        + "	background:#C5D9F1;"
        + "	width: 20px }"
        + ".noAplica {"
        + "	background: #919199 }"
        + "</style>"
        + "</head>"
        + "<body>"
        + "<table class='tdatos'>"
        + " <tr >"
        + "  <td colspan=11 class='resaltados'>FORZADOS MT</td>"
        + " </tr>"
        + " <tr >"
        + "  <td rowspan=4 class='resaltados'> N° de Campaña</td>"
        + "  <td rowspan=4 class='resaltados'> Fecha</td>"
        + "  <td colspan=3 class='resaltados'> (*)</td>"
        + "  <td rowspan=4 width=100 class='resaltados'> "
        + "    Clientes Afectados por Cortes</td>"
        + "  <td rowspan=4 class='resaltados'> Cantidad de Llamados</td>"
        + "  <td colspan=4 class='resaltados'> COMUNICACIÓN</td>"
        + " </tr>"
        + " <tr>"
        + "  <td class='turnos'>6</td>"
        + "  <td class='turnos'>12</td>"
        + "  <td class='turnos'>18</td>"
        + "  <td rowspan=3 class='resaltados'>Exitosa</td>"
        + "  <td rowspan=3 class='resaltados'>No Contesta</td>"
        + "  <td rowspan=3 class='resaltados'>N° Teléfono Erróneo</td>"
        + "  <td rowspan=3 class='resaltados'>Suspendidos</td>"
        + " </tr>"
        + " <tr>"
        + "  <td class='turnos'>A</td>"
        + "  <td class='turnos'>A</td>"
        + "  <td class='turnos'>A</td>"
        + " </tr>"
        + " <tr>"
        + "  <td class='turnosInferior'>12</td>"
        + "  <td class='turnosInferior'>18</td>"
        + "  <td class='turnosInferior'>24</td>"
        + " </tr>";
    
    private final String fuente_html_fin = 
         "</table> "
        + "<br>"    
        + "<table class='tlogo'>"
        + " <tr>"
        + "  <td>"
        + "  <img width=71 height=15 src='logo.png'  alt='Descripción: logo de Edenor' >"
        + "  </td>"
        + " </tr>"
        + "</table>"
        + "</body> </html>";
	
	Connection con_1;
	final String driverClass="oracle.jdbc.driver.OracleDriver";
	
	public void Oracle_connect() throws Exception{
	try{
		Class.forName(driverClass).newInstance();
	}catch (ClassNotFoundException e){
	      System.out.println("Error al cargar el driver: "+driverClass+" -Error: "+e);
		  throw e;
	}   catch (InstantiationException e) {
            System.out.println("Error al cargar el driver: "+driverClass+" -Error: "+e);
            throw e;
            } catch (IllegalAccessException e) {
                System.out.println("Error al cargar el driver: "+driverClass+" -Error: "+e);
                throw e;
            }
	try{
		//con_1= DriverManager.getConnection( "jdbc:oracle:thin:@tclh1.tro.edenor:1528:gispr01", "NEX_GIS03", "termostato01" );
                con_1= DriverManager.getConnection( "jdbc:oracle:thin:@nexgispr02.pro.edenor:1528:gispr01s", "SVC_ORA_GIS", "jv506uzy" );
	}catch(SQLException e) {
		switch(e.getErrorCode()) {
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
	
    private List<HashMap<String, String>> strToHash(String reporte){
        
        List<HashMap<String, String>> tabla = new ArrayList<HashMap<String, String>>();        
        reporte = reporte.substring(0, reporte.length() - 1);
        
        String[] lineas = reporte.split("X");
            
        for (String linea : lineas){
            
            String[] campo = linea.split(";");            
            HashMap<String, String> fila = new HashMap<String, String>(); 
            if (!campo[0].equals("")){            
            fila.put("CAMPANIA",campo[0]);          
            fila.put("FECHA",campo[1]);
            fila.put("TURNO_6A12",campo[2]);
            fila.put("TURNO_12A18",campo[3]);
            fila.put("TURNO_18A24",campo[4]);
            fila.put("CLI_AFECTADOS",campo[5]);
            fila.put("CANT_LLAMADAS",campo[6]);
            fila.put("EXITOSA",campo[7]);
            fila.put("NO_CONTESTA",campo[8]);
            fila.put("TEL_INCORRECTO",campo[9]);
            fila.put("SUSPENDIDOS",campo[10]);
            tabla.add(fila);
            }else{
            fila.put("CAMPANIA","0");          
            fila.put("FECHA","0");
            fila.put("TURNO_6A12","0");
            fila.put("TURNO_12A18","0");
            fila.put("TURNO_18A24","0");
            fila.put("CLI_AFECTADOS","0");
            fila.put("CANT_LLAMADAS","0");
            fila.put("EXITOSA","0");
            fila.put("NO_CONTESTA","0");
            fila.put("TEL_INCORRECTO","0");
            fila.put("SUSPENDIDOS","0");
            tabla.add(fila);
            
            }
          }
        
        return tabla;
    }
    
    private HashMap<String, String> strToHashUltLinea(String reporte){
        
        HashMap<String, String> fila = null;
        String[] lineas = reporte.split(";",-1);
        for (int i=0; i < lineas.length; i++){
            fila = new HashMap<String, String>(); 
            fila.put("TOT_SEM_LITERAL",lineas[i++]);
            i+=4; //siempre vienen vacios
            fila.put("TOT_CLI_AFECTADOS",lineas[i++]);
            fila.put("TOT_CANT_LLAMADOS",lineas[i++]);
            fila.put("TOT_EXITOSA",lineas[i++]);
            fila.put("TOT_NO_CONTESTA",lineas[i++]);
            fila.put("TOT_TEL_ERROR",lineas[i++]);
            fila.put("TOT_SUSPEN",lineas[i]);
        }
        return fila;
    }
    
	public void raster() throws Exception{

        String  plsql;
	plsql ="{?=call NEXUS_GIS.LLAM_PROCESOS.LLAM_REPORT_FORZADOS_MT(?)}"; 
	try{
	    Oracle_connect();					
            OracleCallableStatement cs = (OracleCallableStatement)con_1.prepareCall(plsql);
            cs.registerOutParameter(1, OracleTypes.CLOB);
            cs.registerOutParameter(2, OracleTypes.CLOB);
            cs.execute();
            
            List<HashMap<String, String>> tabla;
            HashMap<String, String> ultima;            
            plsql=clobToStr(cs.getCLOB(1));            
            String strarray[] = plsql.split("T");            
            strarray[1]= strarray[1].replace(strarray[1],"T"+strarray[1]);
            
            tabla = strToHash(strarray[0]);            
            ultima = strToHashUltLinea(strarray[1]);
            generar(tabla, ultima);
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
	public SalientesForzados() throws Exception{
		try{
			raster();
		}catch(Exception e){
			System.out.println("Error en Mail_Sender() = "+e );
			throw e;
		}
	}
//------------------------------------------------------------------------------------------------------
	private void generar(List<HashMap<String, String>> tabla, HashMap<String, String> ultima) throws Exception{

		Hashtable<String,String>   hst_Mail= new Hashtable<String,String>();
	
		//long fileSize = 0, SecondsOfDay =0 , SecondsOfFile =0;
		String HTML_Estucture;
                @SuppressWarnings("UnusedAssignment")
		boolean enviarMail = false;
		int i = 0;
		
		try{		    
			i=0;
			enviarMail = true;		
         
			hst_Mail.put("mailHost"  , "mail.edenor");
			hst_Mail.put("DE"        , "ITSM_Llamadas_salientes@edenor.com");
                       // hst_Mail.put("PARA"      , "EMAGGI@edenor.com,OSUAREZ@edenor.com,PPEREZ@edenor.com,Okovalow@edenor.com,PMAZZA@edenor.com,ELAFUENTE@edenor.com,JCALMA@edenor.com,MRRODRIGUEZ@edenor.com");
                       hst_Mail.put("PARA"      , "Okovalow@edenor.com,CIGLESIAS@edenor.com,MRRODRIGUEZ@edenor.com");
                       //hst_Mail.put("PARA"      , "MRRODRIGUEZ@edenor.com");
            
			hst_Mail.put("ASUNTO"    , "Información Resumen Llamadas Salientes de Documentos Forzados ");

			HTML_Estucture  = fuente_html_ppio;
                    if (tabla!= null){
                    for (HashMap fila : tabla) {
                        HTML_Estucture += " <tr class='datosHr'>";
                        HTML_Estucture += "  <td >"+fila.get("CAMPANIA")+"</td>";
                        HTML_Estucture += "  <td>"+fila.get("FECHA")+"</td>";
                        HTML_Estucture += "  <td>"+fila.get("TURNO_6A12")+"</td>";
                        HTML_Estucture += "  <td>"+fila.get("TURNO_12A18")+"</td>";
                        HTML_Estucture += "  <td>"+fila.get("TURNO_18A24")+"</td>";
                        HTML_Estucture += "  <td>"+fila.get("CLI_AFECTADOS")+"</td>";
                        HTML_Estucture += "  <td>"+fila.get("CANT_LLAMADAS")+"</td>";
                        HTML_Estucture += "  <td>"+fila.get("EXITOSA")+"</td>";
                        HTML_Estucture += "  <td>"+fila.get("NO_CONTESTA")+"</td>";
                        HTML_Estucture += "  <td>"+fila.get("TEL_INCORRECTO")+"</td>";
                        HTML_Estucture += "  <td>"+fila.get("SUSPENDIDOS")+"</td>";
                        HTML_Estucture += " </tr>";
                    }
            
            HTML_Estucture +=  " <tr> ";
            HTML_Estucture +=  "  <td rowspan=3 class='resaltados'> "+ultima.get("TOT_SEM_LITERAL")+"</td>";
            HTML_Estucture +=  "  <td rowspan=3 class='noAplica'>&nbsp;</td>";
            HTML_Estucture +=  "  <td colspan=3 rowspan=3 class='noAplica'>&nbsp;</td>";
            HTML_Estucture +=  "  <td rowspan=3 class='resaltados'> "+ultima.get("TOT_CLI_AFECTADOS")+"</td>";
            HTML_Estucture +=  "  <td rowspan=3 class='resaltados'>"+ultima.get("TOT_CANT_LLAMADOS")+"</td>";
            HTML_Estucture +=  "  <td rowspan=3 class='resaltados'>"+ultima.get("TOT_EXITOSA")+"</td>";
            HTML_Estucture +=  "  <td rowspan=3 class='resaltados'>"+ultima.get("TOT_NO_CONTESTA")+"</td>";
            HTML_Estucture +=  "  <td rowspan=3 class='resaltados'>"+ultima.get("TOT_TEL_ERROR")+"</td>";
            HTML_Estucture +=  "  <td rowspan=3 class='resaltados'>"+ultima.get("TOT_SUSPEN")+"</td>";
            HTML_Estucture +=  " </tr>";
            
                    }else{
                        
                        HTML_Estucture += " <tr class='datosHr'>";
                        HTML_Estucture += "  <td >NO HAY DATOS A LA FECHA</td>";
                        HTML_Estucture += " </tr>";
                    
                    }
            
			HTML_Estucture += fuente_html_fin;


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
		adjunto.setDataHandler(new DataHandler(new FileDataSource("logo.png")));
		adjunto.setFileName("logo.png");
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
private static String clobToStr(Clob clb){
        StringBuilder sb = new StringBuilder();
        
        try{
            final Reader         reader = clb.getCharacterStream();
            final BufferedReader br     = new BufferedReader(reader);
            int b;
            while(-1 != (b = br.read())){
                sb.append((char)b);
            }
            br.close();
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
				SalientesForzados g= new SalientesForzados();
				System.out.println("\nProcedimiento MAIL terminado exitosamente");
				System.exit(0);
			}catch(Exception e){
				System.out.println(e);
				System.exit(1);
			}
		
	}
}

