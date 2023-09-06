cd /ias/mail_reporte_mtat
/usr/java/jre1.8.0_131/bin/java -Duser.timezone=GMT-3:00 -classpath .:/ias/mail_reporte_mtat/Mail_Sender.class:./jar/smtp.jar:./jar/mailapi.jar:./jar/activation.jar:./jar/javax.mail.jar:./jar/ojdbc14.jar  Mail_Sender

