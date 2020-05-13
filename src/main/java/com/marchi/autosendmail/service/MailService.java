package com.marchi.autosendmail.service;

import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.core.io.FileSystemResource;
import org.springframework.mail.javamail.JavaMailSender;
import org.springframework.mail.javamail.MimeMessageHelper;
import org.springframework.stereotype.Service;

import javax.mail.Session;
import javax.mail.internet.MimeMessage;
import javax.mail.internet.MimeUtility;
import java.io.File;
import java.util.Properties;

@Service
public class MailService {
    Logger logger= LoggerFactory.getLogger(this.getClass());

    @Autowired
    JavaMailSender jms;

    public void sendAttachmentsMail(String subject,String text,String sendFrom,String sendTo, String attachmentPath) {



        String [] fileArray={attachmentPath};

        MimeMessage message=jms.createMimeMessage();
        try {


            MimeMessageHelper helper=new MimeMessageHelper(message,true);
            helper.setFrom(sendFrom);
            helper.setTo(sendTo);
            helper.setSubject(subject);
            helper.setText(text,true);
            //验证文件数据是否为空
            if(null != fileArray){
                FileSystemResource file=null;
                for (int i = 0; i < fileArray.length; i++) {
                    //添加附件
                    file=new FileSystemResource(fileArray[i]);
                    helper.addAttachment(MimeUtility.encodeText(fileArray[i].substring(fileArray[i].lastIndexOf(File.separator))), file);
                }
            }
            jms.send(message);
            logger.info("发送邮件到{}成功！",sendTo);
            //System.out.println("带附件的邮件发送成功");
        }catch (Exception e){
            logger.info("{}发送邮件失败:"+e.getMessage(),sendTo);
            //System.out.println("发送带附件的邮件失败");
        }
    }

    public void sendAttachmentsMail(String subject,String text,String sendFrom,String sendTo, String[] attachmentPath) {



        String [] fileArray=attachmentPath;

        MimeMessage message=jms.createMimeMessage();
        try {


            MimeMessageHelper helper=new MimeMessageHelper(message,true);
            helper.setFrom(sendFrom);
            helper.setTo(sendTo);
            helper.setSubject(subject);
            helper.setText(text,true);
            //验证文件数据是否为空
            if(null != fileArray){
                FileSystemResource file=null;
                for (int i = 0; i < fileArray.length; i++) {
                    //添加附件
                    file=new FileSystemResource(fileArray[i]);
                    helper.addAttachment(fileArray[i].substring(fileArray[i].lastIndexOf(File.separator)), file);
                }
            }
            jms.send(message);
            logger.info("发送邮件到{}成功！",sendTo);
            //System.out.println("带附件的邮件发送成功");
        }catch (Exception e){
            logger.info("{}发送邮件失败:"+e.getMessage(),sendTo);
            //System.out.println("发送带附件的邮件失败");
        }
    }
}
