package com.edenor.conexionoracle;

import java.sql.DriverManager;
import java.sql.Connection;
import java.sql.Statement;
import java.sql.ResultSet;

/**
 *
 * @author rsleiva
 */
public class Oracle {
    
    String nombre_servidor;
    String numero_puerto;
    String sid;
    String url;
    String usuario;
    String password;

    public Oracle() {
        //nombre del servidor
        nombre_servidor = "nexgispr02.pro.edenor";
        //numero del puerto
        numero_puerto = "1528";
        //SID
        sid = "gispr01S";
        //URL "jdbc:oracle:thin:@nombreServidor:numeroPuerto:SID"
        url = "jdbc:oracle:thin:@" + nombre_servidor + ":" + numero_puerto + ":" + sid;

        //Nombre usuario y password
        usuario = "rsleiva";
        password = "G1s_klg4fys";
    }
    
    public void Conectar(){
        try
        {
            //Se carga el driver JDBC
            DriverManager.registerDriver( new oracle.jdbc.driver.OracleDriver() );
             
 
            //Obtiene la conexion
            Connection conexion = DriverManager.getConnection( url, usuario, password );
             
            //Para realiza una consulta
            Statement sentencia = conexion.createStatement();
            ResultSet resultado = sentencia.executeQuery( "SELECT * FROM NEXUS_GIS.SMS_LOG WHERE ROWNUM<100" );
             
            //Se recorre el resultado obtenido
            while ( resultado.next() )
            {
                //Se imprime el resultado colocando
                //Para obtener la primer columna se coloca el numero 1 y para la segunda columna 2 el numero 2
                System.out.println ( resultado.getInt( 1 ) + "\t" + resultado.getString( 2 ) );
            }
             
            //Cerramos la sentencia
            sentencia.close();
        }catch( Exception e ){
            e.printStackTrace();
        }
    }    
}
