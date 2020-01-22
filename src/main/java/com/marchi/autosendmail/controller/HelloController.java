package com.marchi.autosendmail.controller;

import com.marchi.autosendmail.MailService;
import com.marchi.autosendmail.utils.ExcelUtils;
import com.marchi.autosendmail.utils.ExportWordUtils;
import com.marchi.autosendmail.utils.ReadExcelUtils;
import org.apache.commons.lang3.StringUtils;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.beans.factory.annotation.Value;
import org.springframework.stereotype.Controller;
import org.springframework.ui.Model;
import org.springframework.ui.ModelMap;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.bind.annotation.ResponseBody;
import org.springframework.web.multipart.MultipartFile;

import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import java.io.File;
import java.io.IOException;
import java.math.BigDecimal;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

@Controller
public class HelloController {

    Logger logger= LoggerFactory.getLogger(this.getClass());

    @Autowired
    MailService mailService;

    @Value("${userName}")
    private String userName;

    @Value("${bookTitle}")
    private String bookTitle;

    @RequestMapping("/")
    public String index(ModelMap map) {
        // 加入一个属性，用来在模板中读取
        map.addAttribute("name", userName);
        map.addAttribute("bookTitle", bookTitle);
        // return模板文件的名称，对应src/main/resources/templates/welcome.html
        return "welcome";
    }


    @GetMapping("/index_file")
    public String index(Model modelMap){
        modelMap.addAttribute("msg","文件上传下载");
        return "upload";
    }

    private static final String filePath="D:/downloads/";

    @RequestMapping(value="/upload")
    public String upload(@RequestParam("file") MultipartFile file,ModelMap map){
        if(file.isEmpty()){
            return "文件为空";
        }
        try {
            //获取文件名
            String fileName = file.getOriginalFilename();
            logger.info("上传的文件名称为》》》,{}",fileName);
            //设置文件存储路径
            String path=filePath+fileName;
            File dest = new File(path);
            //检测是否存在目录
            if(!dest.getParentFile().exists()){
                //新建文件夹
                dest.getParentFile().mkdirs();
            }
            //文件写入

            file.transferTo(dest);
            map.addAttribute("msg","文件上传成功");
            return "upload";
        } catch (IOException e) {
            logger.error(e.getMessage(),e);
        }
        map.addAttribute("msg","文件上传失败");
        return  "upload";
    }

    @RequestMapping("file/upload")
    public String pubggupload(@RequestParam("file")MultipartFile file,ModelMap map, HttpServletRequest request, HttpServletResponse response) throws Exception {

        List<ArrayList<String>> readExcel = new ArrayList<>();
        try {

            readExcel = new ReadExcelUtils().readExcel(file);
            for(int i=0;i<readExcel.size();i++){

                    ArrayList<String> record= readExcel.get(i);
                    Map<String,Object> params = new HashMap<>();

                    String staffNo=record.get(2);
                    if(StringUtils.isEmpty(staffNo)){
                        continue;
                    }

                    params.put("staffNo",staffNo);
                    params.put("username",record.get(1));
                    params.put("dept",record.get(0));
                    params.put("rank",record.get(4));
                    params.put("position",record.get(3));
                    params.put("ceoBonus",record.get(6));
                    params.put("bonus",record.get(5));
                    params.put("total",record.get(7));
                    String email=record.get(8);
                    String password=record.get(9);
                    String languare=record.get(10);
                    String sendFrom=record.get(11);
                    String templatePath="";
                    if("EN".equals(languare)){
                        if(new BigDecimal(record.get(6)).compareTo(BigDecimal.ZERO)!=0){
                            templatePath="static/2019奖金沟通信(EN-CEO).docx";
                        }else{
                            templatePath="static/2019奖金沟通信(EN).docx";
                        }
                    }else{
                        if(new BigDecimal(record.get(6)).compareTo(BigDecimal.ZERO)!=0){
                            templatePath="static/2019奖金沟通信(CN-CEO).docx";
                        }else{
                            templatePath="static/2019奖金沟通信(CN).docx";
                        }
                    }



                    //这里是我说的一行代码
                    ExportWordUtils.exportWord(templatePath,"d:/downloads",password,staffNo+".docx",params,request,response);
                    mailService.sendAttachmentsMail("2019年年度奖金沟通信",emailContext(),sendFrom,email,"d:/downloads"+File.separator+staffNo+".docx");
            }
            file.transferTo(new File("D:/downloads/"+file.getOriginalFilename()));
            map.addAttribute("name", userName);
            map.addAttribute("bookTitle", bookTitle);

        } catch (IOException e) {
            // TODO Auto-generated catch block
            e.printStackTrace();
        }

        return  "success";

    }

    private String emailContext(){



        StringBuffer sb=new StringBuffer();
        sb.append("<html>");
        sb.append("<body lang=ZH-CN link=\"#0563C1\" vlink=\"#954F72\" style='text-justify-trim:punctuation'>");
        sb.append("<div class=WordSection1>");
        sb.append("<p class=MsoPlainText><span style='font-size:11.0pt;font-family:\"微软雅黑\",sans-serif'>亲爱的同事，<span lang=EN-US><o:p></o:p></span></span></p>");
        sb.append("<p class=MsoPlainText><span style='font-size:11.0pt;font-family:\"微软雅黑\",sans-serif'>您好！请查阅附件，为您<span lang=EN-US>2019</span>年年度奖金沟通信！密码是您身份证号码后<span lang=EN-US>6</span>位。<span lang=EN-US style='color:#1F497D'><o:p></o:p></span></span></p>");
        sb.append("<p class=MsoPlainText><span style='font-size:11.0pt;font-family:\"微软雅黑\",sans-serif'>感谢您对公司做出的贡献！祝您和您的家人阖家幸福，新年快乐！<span lang=EN-US><o:p></o:p></span></span></p>");
        sb.append("<p class=MsoPlainText><span style='font-size:11.0pt;font-family:\"微软雅黑\",sans-serif'>如有任何问题，请咨询人力资源及行政部。<span lang=EN-US><o:p></o:p></span></span></p>");
        sb.append("<p class=MsoNormal><span lang=EN-US style='font-size:12.0pt;color:#1F497D'><o:p>&nbsp;</o:p></span></p>");
        //sb.append("<p class=MsoNormal><span lang=EN-US style='color:#1F497D'><o:p>&nbsp;</o:p></span></p>");
        sb.append("<p class=Default style='line-height:150%'><span lang=EN-US style='font-size:12.0pt;font-family:\"Calibri\";line-height:150%'>Dear Colleague,<o:p></o:p></span></p>");
        sb.append("<p class=Default style='line-height:150%'><span lang=EN-US style='font-size:12.0pt;font-family:\"Calibri\";line-height:150%'>Please refer to the attachment for your Y2019 bonus communication letter. Password for the file is the last six numbers of your Identity No. <o:p></o:p></span></p>");
        sb.append("<p class=Default style='line-height:150%'><span lang=EN-US style='font-size:12.0pt;font-family:\"Calibri\";line-height:150%'>Highly appreciated your effort &amp; contribution to the company. Wish you a prosperous Y2020 and happy Spring Festival with your family.<o:p></o:p></span></p>");
        sb.append("<p class=Default style='line-height:150%'><span lang=EN-US style='font-size:12.0pt;font-family:\"Calibri\";line-height:150%'>Should you have any questions, please contact HR.<o:p></o:p></span></p>");
        sb.append("<p class=Default style='line-height:150%'><span lang=EN-US style='font-size:10.5pt;font-family:\"Calibri\";line-height:150%;color:#1F497D'><o:p>&nbsp;</o:p></span></p>");
        sb.append("<p class=MsoNormal><b><span style='font-family:\"微软雅黑\",sans-serif;color:black'>人力资源及行政部</span></b><span lang=EN-US style='font-size:10.0pt;font-family:\"微软雅黑\",sans-serif;color:black'><o:p></o:p></span></p>");
        sb.append("<p class=MsoNormal><b><span lang=EN-US style='font-family:\"Calibri Light\",sans-serif;color:black'>HR &amp; Admin. Department</span></b><span lang=EN-US style='font-size:10.0pt;font-family:\"Calibri Light\",sans-serif;color:black'><o:p></o:p></span></p>");
        sb.append("</div>");
        sb.append("</body>");
        sb.append("</html>");
        return sb.toString();
    }


}
