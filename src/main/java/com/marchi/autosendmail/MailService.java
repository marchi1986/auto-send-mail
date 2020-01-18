package com.marchi.autosendmail;

import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.core.io.FileSystemResource;
import org.springframework.mail.javamail.JavaMailSender;
import org.springframework.mail.javamail.MimeMessageHelper;
import org.springframework.stereotype.Service;

import javax.mail.internet.MimeMessage;
import java.io.File;

@Service
public class MailService {
    @Autowired
    JavaMailSender jms;

    public void sendAttachmentsMail(String subject,String text,String attachmentPath) {



        String [] fileArray={attachmentPath};

        MimeMessage message=jms.createMimeMessage();
        try {
            MimeMessageHelper helper=new MimeMessageHelper(message,true);
            helper.setFrom("58844411@qq.com");
            helper.setTo("albee_liang@skechers.cn");
            helper.setSubject(subject);
            helper.setText(text);
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
            System.out.println("带附件的邮件发送成功");
        }catch (Exception e){
            e.printStackTrace();
            System.out.println("发送带附件的邮件失败");
        }
    }
}
