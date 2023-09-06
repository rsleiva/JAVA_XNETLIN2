/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */

/**
 *
 * @author oviglione
 */

import java.io.BufferedReader;
import java.io.ByteArrayInputStream;
import java.io.File;
import java.io.FileReader;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.text.DateFormat;
import java.util.Date;
import java.util.Hashtable;
import java.util.Locale;
import java.util.Properties;
import java.util.Vector;

import javax.activation.DataHandler;
import javax.activation.DataSource;
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
import oracle.net.aso.o;


public class Chequeo_de_Errores {
    
    private static Vector dir_no_encontradas= new Vector();
    static String CadenaError=null;
    private static boolean ErrorFound = false;
    //static String opentd="<td height=21 class=xl7730982 colspan=9 width=494 style='height:15.75pt;width:600pt'>";
    static String opentd="<br>";
    static String closetd = "</br>";
    static String logfileAnt = null;
    static String NoErrores = null  ;
    static String Consulta =  "<br></br><br></br><br> Ante cualquier consulta contactar a <a>ITSM_Desarrollos_propios@edenor.com</a></br>";
    static String Destinatarios = null;
    static String DestinatariosNOErrores = "ITSM_Desarrollos_propios@edenor.com";
    static String Destinatarios_ERROR = "ITSM_Desarrollos_propios@edenor.com,htestagrossa@edenor.com,vggolidgur@edenor.com,MFORESTELLO@edenor.com,pcanales@edenor.com,sfrancica@edenor.com,SMAZZEO@edenor.com,nespiasse@edenor.com,istephan@edenor.com,srossello@edenor.com,ggarcia@edenor.com";
   

 // Variables para desarrollar :
   
    /* 
    static String LogPartes = "C:\\parte_obras\\estructura\\logs\\LogsObras\\batchLog.txt";
    static String Logavances = "C:\\parte_obras\\estructura\\logs\\LogsAvances\\batchLog.txt";
    static String Logdocs = "C:\\parte_obras\\estructura\\logs\\LogsDoc\\batchLog.txt";
   
     */
 

    
// Variables en el servidor :

    private static final String base= "/ias/enre748/"; 
    static String LogPartes = base +  "PartesObras/estructura/logs/batchLog.txt";
    static String Logavances = base +  "CargaAvancesObras/estructura/logs/batchLog.txt";
    static String Logdocs = base +  "DocumentacionObras/estructura/logs/batchLog.txt";



     public static void main(String [] arg) throws Exception{
    
         
   LeerLogs (LogPartes);
//   LeerLogs (Logavances);
//   LeerLogs (Logdocs);
   PrepararMail();
  
     }

  public static void LeerLogs (String LogFile) {
      File archivo = null;
      FileReader fr = null;
     int found;
     String LineaAnt=null;
      BufferedReader br = null;
      Date today;
      String dateOut;
       DateFormat dateFormatter;
       Locale currentLocale = new Locale ("en", "US") ;
       String fecha = null;
       String Mes = null;
       String Dia = null;
       String Anio = null;
       String TMes = null;
       String TDia = null;
       String TA = null;
       Date Date = null;
       int found1;
       int found2;
       int found3;
       int found4;
       int found5;
       int found6;
       int found7;
       String proceso = null;
       String Fechahora;
       Date FechayHora = new Date(); 

       
       if (LogFile == LogPartes){       
           proceso = "Parte de obras"; 
       }   
       else if (LogFile == Logavances){
           proceso = "Avances"; 
       } 
              else if (LogFile == Logdocs){
           proceso = "Documentacion"; 
       } 

 

    NoErrores = "<br> Este es un e-mail automatico.</br><br></br><br><strong>No se encontraron errores ni advertencias en el procesamiento de ningun archivo de obras el dia de la fecha. (" + FechayHora + ") </strong></br> <br></br>" ;
       
    String Newprocesslog = "<br></br><br> <u> Se encontraron los siguientes errores o advertencias en el proceso de " + proceso + " : </u> </br>" ;
    String Newprocesslog1 = "<br> Este es un e-mail automatico.</br><br>Se reportan los errores encontrados en los archivos de procesos de obras a la fecha:" + FechayHora + "</br><br></br>" ;
      
      
       
    dateFormatter = DateFormat.getDateInstance(DateFormat.DEFAULT, currentLocale);

      try {
        today = new Date();
//      System.out.println(today);
        dateOut = dateFormatter.format(today);
//        System.out.println(dateOut);
        String Datenow =dateOut.toString();

                String[] partsDateParts = Datenow.split(",");
                String part1 = partsDateParts[0]; // 123
                String part2 = partsDateParts[1]; // 654321
                Integer largop1 = part1.length();

       TMes=Datenow.substring(0,3).toLowerCase();

       TDia=part1.substring(largop1 -2 ,largop1);
       TDia=TDia.trim();

       TA= part2.trim() ;
//       System.out.println(TAÃ±o);
          
          
         // Apertura del fichero y creacion de BufferedReader para poder
         // hacer una lectura comoda (disponer del metodo readLine()).
         archivo = new File (LogFile);
         fr = new FileReader (archivo);
         br = new BufferedReader(fr);

         // Lectura del fichero
         String linea;
 
         while((linea=br.readLine())!=null){
       
      found =  linea.indexOf("SEVERE:");
      found1 = linea.indexOf("GRAVE :");
      found2= linea.indexOf("GRAVE:");
      found3 =  linea.indexOf("SEVERE :");
      found4 =  linea.indexOf("WARNING:");
      found5 =  linea.indexOf("WARNING :");
      found6 =  linea.indexOf("ADVERTENCIA:");
      found7 =  linea.indexOf("ADVERTENCIA :");
      System.out.println(linea);
      if ((found != -1)||(found1 != -1)||(found2 != -1)||(found3 != -1)||(found4 != -1)||(found5 != -1)||(found6 != -1)||(found7 != -1)) { 

       fecha = LineaAnt.substring(0,13);
       Mes=fecha.substring(0,3).toLowerCase();
       Dia=fecha.substring(4,6);
       Anio=fecha.substring(8,12);

       if ((Anio.equals(TA)) && (Mes.equals(TMes))&& (Integer.valueOf(Dia)== Integer.valueOf(TDia)) ) {   
           
//        System.out.println(LineaAnt);
//        System.out.println(linea);
        
//       System.out.println("Antes del if");
        
        if (CadenaError == null) {     
            
//            System.out.println("entro porque la cadena es nula");
            
                CadenaError =Newprocesslog1 + Newprocesslog + opentd + LineaAnt + closetd + opentd + linea + closetd;
                logfileAnt = LogFile ;
                
                
         }else if (logfileAnt == LogFile) {
            
//            System.out.println("entro porque NO cambio de archivo");
            
            
                CadenaError = CadenaError + opentd + LineaAnt + closetd + opentd + linea + closetd;
                
        }else if ((logfileAnt != LogFile)&& (logfileAnt != null)) {
            
//                        System.out.println("entro porque SI cambio de archivo");
                        
                        logfileAnt = LogFile ;
            
                CadenaError = CadenaError + Newprocesslog + opentd + LineaAnt + closetd + opentd + linea + closetd;
        }                

        ErrorFound = true;

       } 
       
      }
          
      LineaAnt=linea;
      found=0;
      
         
         }
      logfileAnt = LogFile ;

       

         
  //      System.out.println(part1); 
  //      System.out.println(part2); 

      }catch(Exception e){
         e.printStackTrace();
      }finally{
         // En el finally cerramos el fichero, para asegurarnos
         // que se cierra tanto si todo va bien como si salta 
         // una excepcion.
         try{                    
            if( null != fr ){   
               fr.close();     
            }                  
         }catch (Exception e2){ 
            e2.printStackTrace();
         }
      }     
   }
  
  
      private static void PrepararMail() throws Exception {
      
    try  {
          
                if (ErrorFound == true)  {
                    
                    Destinatarios = Destinatarios_ERROR;                            
                    //Destinatarios = "oviglione@edenor.com";
                    generar(CadenaError + Consulta);

                }else {
                
                    Destinatarios = DestinatariosNOErrores;
                    //Destinatarios = "oviglione@edenor.com";
                    generar (NoErrores + Consulta);
                    }
                     
          
            }catch (Exception e2){ 
            e2.printStackTrace();
            }
      
      
      }

  
  private static void generar(String Mensaje) throws Exception{

                @SuppressWarnings("UseOfObsoleteCollectionType")
		Hashtable<String,String>   hst_Mail= new Hashtable<>();

		String HTML_Estucture;
                @SuppressWarnings("UnusedAssignment")
                      

		int i = 0;
                boolean enviarMail = true;
                
  
		
		try{		    
			i=0;
         
			hst_Mail.put("mailHost"  , "mail.edenor");
			hst_Mail.put("DE"        , "ITSM_Desarrollos_propios@edenor.com");
			hst_Mail.put("PARA"      , Destinatarios);			
			hst_Mail.put("ASUNTO"    , "Procesamiento de archivos de obras.");


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
            

                String sinDatosMsj = "IMPORTANTE: El presente es un mail automÃ¡tico informando el poblado de obras de manera exitosa. Ambiente: PROD, Archivo procesado: ";
    
                HTML_Estucture += "</br></br>";
                HTML_Estucture += " <tr height=21 style='height:15.75pt'>";
                HTML_Estucture +=  Mensaje ;
                HTML_Estucture += " </tr>";
                HTML_Estucture += "</table>";                
                HTML_Estucture += " <br>";
                HTML_Estucture += " <hr>";
                HTML_Estucture += " <br>";                
                HTML_Estucture += "<table border=0 cellpadding=5px cellspacing=0 width=494 style='border-collapse: collapse;table-layout:auto;width:371pt'>";
                HTML_Estucture += " <tr height=20 style='height:15.0pt'>";
                HTML_Estucture += "  <td height=20 style='height:15.0pt' align=left valign=top colspan=5 >";
               // HTML_Estucture += "  <span style='mso-ignore:vglayout;position:absolute;z-index:1;margin-left:3px;margin-top:3px;width:71px;height:15px'>";
               // HTML_Estucture += "  <img width=71 height=15 src='Logo_edenor.gif'  alt='DescripciÃƒÂ³n: logo de Edenor' ></span>";
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
			String msgCancela= e.toString();
			hst_Mail.put("CUERPO"    , msgCancela);
			mailSender(hst_Mail);
			throw e;
		}finally{
			try{
				System.out.println("Proceso Finalizado");
			}catch(Exception e1){}
		}
}
  
  
  
  public static void mailSender(@SuppressWarnings("UseOfObsoleteCollectionType") Hashtable<String,String> hst_values_mail) throws Exception {
	try {
		
		Properties properties = new Properties();
		properties.put("mail.smtp.host",hst_values_mail.get("mailHost"));
		properties.put("mail.from"     ,hst_values_mail.get("DE"));		
		properties.put("mail.debug"    , "false");

		
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
		
//		BodyPart adjunto = new MimeBodyPart();
//		adjunto.setDataHandler(new DataHandler(new FileDataSource("Logo_edenor.gif")));
//		adjunto.setFileName("Logo_edenor.gif");
//		multiParte.addBodyPart(adjunto);
        
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


}

