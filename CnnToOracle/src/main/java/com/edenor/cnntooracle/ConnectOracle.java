package com.edenor.cnntooracle;

import java.sql.DriverManager;
import java.sql.Connection;
import java.sql.Statement;
import java.sql.ResultSet;
import java.util.ArrayList;
import java.util.List;

/**
 *
 * @author rsleiva
 */
public class ConnectOracle {
    
    private final String nombre_servidor;
    private final String numero_puerto;
    private final String sid;
    private final String url;
    private final String usuario;
    private final String password;
    private List<MyEntity> lista;

    public ConnectOracle() {
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
    
    public void passToList(){
        try {
            DriverManager.registerDriver( new oracle.jdbc.driver.OracleDriver() );
            Connection conexion = DriverManager.getConnection( url, usuario, password );
            Statement sentencia = conexion.createStatement();
            ResultSet resultado = sentencia.executeQuery( "SELECT IDTRANSACCION FROM NEXUS_GIS.SMS_LOG WHERE ROWNUM<1000000" );
            
            lista = new ArrayList<>();
            while (resultado.next()) {
              int id = resultado.getInt("IDTRANSACCION");
              MyEntity entity = new MyEntity(id);
              lista.add(entity);
            }            
            
        } catch (Exception e) {
            lista= null;
        }
    }
    
    public void analizaDatos(){
        try {
            DriverManager.registerDriver( new oracle.jdbc.driver.OracleDriver() );
            Connection conexion = DriverManager.getConnection( url, usuario, password );
            Statement sentencia = conexion.createStatement();
            ResultSet resultado = sentencia.executeQuery( 
                "SELECT IDTRANSACCION FROM NEXUS_GIS.SMS_LOG WHERE IDTRANSACCION BETWEEN 1347758 AND 3000000 ORDER BY IDTRANSACCION" );
            
            int num=1347758;
            int dif=0;
            int id;
            
            while (resultado.next()) {
                if (num+1<Integer.MAX_VALUE){
                    num++;
                    id = resultado.getInt("IDTRANSACCION");
                    if (num<id) {
                        dif+=id-num;
                        System.out.println(String.format("Desde %d hasta %d hay %d ids libres.",num,id-1,id-num));
                    }
                    num=id;
                } else {
                    System.out.println("Se genera desbordamiento de Integer");
                    break;
                }
            }   
            
            System.out.println(String.format("La cantidad totales de ids libres hasta el desobrdamiento es %d: ",dif));
            System.out.println(String.format("Tomando el maximo mensual de SMS (18.000), podemos estimar que con esta cantidad de ids disponibles tenemos para adelante %d meses. ",Integer.valueOf(dif/18000)));
            
        } catch (Exception e) {
            System.out.println("Error de conexion");
        }
    }
    
}
