package com.sriwin.outlook.email.reader;

import org.apache.commons.lang3.StringUtils;
import org.jsoup.Jsoup;

import javax.mail.Folder;
import javax.mail.Message;
import javax.mail.MessagingException;
import javax.mail.Part;
import javax.mail.Session;
import javax.mail.Store;
import javax.mail.internet.MimeMultipart;
import java.util.Properties;

public class ReadEmailTest {
  public static void main(String[] args) {
    try {
      Properties properties = new Properties();
      properties.setProperty("mail.imap.ssl.enable", "true");
      properties.put("mail.imap.starttls.enable", "true");
      properties.put("mail.debug", "false");

      Session mailSession = Session.getInstance(properties);
      mailSession.setDebug(true);

      Store mailStore = mailSession.getStore("imap");
      mailStore.connect("outlook.office365.com", "abc@xyz.com", "abc123xyz");

      //create the folder object and open it
      Folder folder = mailStore.getFolder("INBOX");
      folder.open(Folder.READ_ONLY);

      Message[] messages = folder.getMessages();
      for (int i = 0; i < messages.length; i++) {
        Message message = messages[i];
        System.out.println("Email Number " + (i + 1));
        System.out.println("From: " + message.getFrom()[0]);
        System.out.println("Subject: " + message.getSubject());
        System.out.println("Text: " + getEmailBody(message));
      }

      folder.close(false);
      mailStore.close();

    } catch (Exception e) {
      e.printStackTrace();
    }
  }

  private static String getEmailBody(Message message) {
    String emailBody = "";
    try {
      if (message.isMimeType("text/plain")) {
        emailBody = getFinalBody(emailBody, message.getContent().toString());
      } else if (message.isMimeType("multipart/*")) {
        MimeMultipart mimeMultipart = (MimeMultipart) message.getContent();
        emailBody = getFinalBody(emailBody, getTextFromMimeMultipart(mimeMultipart));
      }
    } catch (Exception e) {
      e.printStackTrace();
    }
    return emailBody;
  }

  private static String getTextFromMimeMultipart(MimeMultipart mimeMultipart) throws MessagingException {
    String emailBody = "";
    for (int i = 0; i < mimeMultipart.getCount(); i++) {
      Part part = mimeMultipart.getBodyPart(i);
      emailBody = getFinalBody(emailBody, getTextFromBodyPart(part));
    }
    return emailBody;
  }

  private static String getTextFromBodyPart(Part part) {
    String result = "";
    try {
      if (part.isMimeType("text/plain")) {
        result = (String) part.getContent();
      } else if (part.isMimeType("text/html")) {
        String html = (String) part.getContent();
        result = Jsoup.parse(html).text();
      } else if (part.getContent() instanceof MimeMultipart) {
        result = getTextFromMimeMultipart((MimeMultipart) part.getContent());
      }
    } catch (Exception e) {
      e.printStackTrace();
    }
    return result;
  }

  private static String getFinalBody(String oldData, String newData) {
    if (StringUtils.isBlank(oldData) && StringUtils.isBlank(newData)) {
      return "";
    }
    if (!StringUtils.isBlank(oldData) && StringUtils.isBlank(newData)) {
      return oldData;
    }
    if (StringUtils.isBlank(oldData) && !StringUtils.isBlank(newData)) {
      return newData;
    }
    return "";
  }
}