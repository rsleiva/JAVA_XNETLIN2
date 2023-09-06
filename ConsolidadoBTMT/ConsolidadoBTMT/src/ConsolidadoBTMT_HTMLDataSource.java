import java.io.ByteArrayInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import javax.activation.DataSource;

public class ConsolidadoBTMT_HTMLDataSource implements DataSource  {
	private final String html;

  public ConsolidadoBTMT_HTMLDataSource(String htmlString) {
     this.html = htmlString;
  }

  public InputStream getInputStream() throws IOException {
     if (this.html == null) {
        throw new IOException("Null HTML");
     } else {
        return new ByteArrayInputStream(this.html.getBytes());
     }
  }

  public OutputStream getOutputStream() throws IOException {
     throw new IOException("Este DataHandler no puede crear HTML");
  }

  public String getContentType() {
     return "text/html";
  }

  public String getName() {
     return "text/html dataSource para solo enviar e-mail";
  }
}
