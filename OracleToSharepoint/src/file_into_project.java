

import java.io.File;
import java.io.IOException;
import java.io.InputStream;
import java.net.URL;
import java.util.Properties;

/**
 *
 * @author rsleiva
 */
public class file_into_project {

    public String getFileIntoProject(){
        // Obtener el ClassLoader del proyecto
        ClassLoader classLoader = file_into_project.class.getClassLoader();

        // Cargar el archivo properties
        Properties properties = new Properties();
        try {
            //busca el archivo properties en la base raiz del projecto, no en src.
            InputStream input = classLoader.getResourceAsStream("config.properties");
            properties.load(input);
        } catch (IOException e) {
            System.out.println("No se pudo cargar el archivo properties.");
            e.printStackTrace();
            return null;
        }

        // Obtener el nombre del archivo desde el archivo properties
        String fileName = properties.getProperty("nombreArchivo");

        // Buscar el archivo dentro del proyecto
        URL resource = classLoader.getResource(fileName);

        if (resource != null) {
            // Si se encuentra el archivo, obtener la ruta de acceso
            File file = new File(resource.getFile());
            String filePath = file.getAbsolutePath();
            System.out.println("Ruta del archivo encontrado: " + filePath);
            return filePath;
        } else {
            System.out.println("El archivo no se encontró en el proyecto.");
            return null;
        }
    }
}
