

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

/**
 *
 * @author rsleiva
 */
public class OracleToSharepoint {

    /**
     * @param args the command line arguments
     */
    public static void main(String[] args) {
        
        file_into_project fip= new file_into_project();

        String sourceFilePath = fip.getFileIntoProject();
        String destinationFilePath = "https://edenor.sharepoint.com/sites/GO365_DesarrollosPropios641/Documentos%20compartidos/Forms/AllItems.aspx?id=%2Fsites%2FGO365%5FDesarrollosPropios641%2FDocumentos%20compartidos%2F202306&viewid=43c482a7%2D6ed6%2D4b62%2Db7df%2D2a2e6e32d7c3/destino.txt"; // Ruta de destino en SharePoint

        try {
            // Leer el archivo de origen
            File sourceFile = new File(sourceFilePath);
            FileInputStream fileInputStream = new FileInputStream(sourceFile);
            byte[] buffer = new byte[(int) sourceFile.length()];
            fileInputStream.read(buffer);
            fileInputStream.close();

            // Guardar el archivo en SharePoint
            File destinationFile = new File(destinationFilePath);
            FileOutputStream fileOutputStream = new FileOutputStream(destinationFile);
            fileOutputStream.write(buffer);
            fileOutputStream.close();

            System.out.println("Archivo guardado en SharePoint correctamente.");
        } catch (IOException e) {
            System.out.println("Error al realizar la operación: " + e.getMessage());
        }        
        
    }
    
}
