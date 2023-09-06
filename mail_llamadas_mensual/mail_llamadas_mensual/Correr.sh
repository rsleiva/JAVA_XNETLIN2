cd /ias/mail_llamadas_mensual
/usr/java/jre1.8.0_131/bin/java -Duser.timezone=GMT-3:00 -classpath .:/ias/mail_llamadas_mensual/mail_llamadas_anual.class:./jar/javax.mail.jar:./jar/poi-ooxml-schemas-3.7-20101029.jar:./jar/commons-net-3.3.jar:./jar/dom4j-1.6.1.jar:./jar/ojdbc7.jar:./jar/org.apache.poi-poi-3.8.jar:./jar/poi-ooxml-3.7-20101029.jar:./jar/xmlbeans-2.3.0.jar  mail_llamadas_anual
