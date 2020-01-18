package com.marchi.autosendmail.controller;

import com.marchi.autosendmail.MailService;
import com.marchi.autosendmail.utils.ExcelUtils;
import com.marchi.autosendmail.utils.ExportWordUtils;
import com.marchi.autosendmail.utils.ReadExcelUtils;
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
                    params.put("staffNo",record.get(0));
                    params.put("username",record.get(1));
                    params.put("dept",record.get(2));
                    params.put("rank",record.get(3));
                    params.put("position","normal");
                    params.put("doublePay",record.get(4));
                    params.put("bonus",record.get(5));
                    params.put("total",record.get(6));



                    //这里是我说的一行代码
                    ExportWordUtils.exportWord("static/2019奖金沟通信.docx","d:/downloads",record.get(0)+".docx",params,request,response);
                    mailService.sendAttachmentsMail("test","test","d:/downloads"+File.separator+record.get(0)+".docx");
            }
            map.addAttribute("name", userName);
            map.addAttribute("bookTitle", bookTitle);

        } catch (IOException e) {
            // TODO Auto-generated catch block
            e.printStackTrace();
        }

        return  "welcome";

    }


}
