package com.marchi.autosendmail.service;

import com.alibaba.excel.EasyExcel;
import com.marchi.autosendmail.listener.StoreInfoEasyExcelListener;
import com.marchi.autosendmail.model.StoreInfo;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Service;

import java.util.List;

@Service
public class AutoSendStoreQRCodeService {

    Logger LOGGER = LoggerFactory.getLogger(this.getClass());

    @Autowired
    MailService mailService;


    public void send(){
        //String path="C:\\Users\\marchi_ma\\Desktop\\门店信息-final.xls";
        String path="C:\\Users\\marchi_ma\\Desktop\\test.xls";
       //List<Object> objects = ExcelUtil.readLessThan1000Row(filePath);
    }

    public void read() {
        // 有个很重要的点 DemoDataListener 不能被spring管理，要每次读取excel都要new,然后里面用到spring可以构造方法传进去
        // 写法1：
        //String fileName = "C:\\Users\\marchi_ma\\Desktop\\test.xlsx";
        String fileName = "C:\\Users\\marchi_ma\\Desktop\\门店信息-final.xls";
        // 这里 需要指定读用哪个class去读，然后读取第一个sheet 文件流会自动关闭
        StoreInfoEasyExcelListener storeInfoEasyExcelListener=new StoreInfoEasyExcelListener(mailService);
        EasyExcel.read(fileName, StoreInfo.class, storeInfoEasyExcelListener).sheet().doRead();
        LOGGER.info(storeInfoEasyExcelListener.list.toString());

        // 写法2：
        /*
        fileName = TestFileUtil.getPath() + "demo" + File.separator + "demo.xlsx";
        ExcelReader excelReader = EasyExcel.read(fileName, DemoData.class, new DemoDataListener()).build();
        ReadSheet readSheet = EasyExcel.readSheet(0).build();
        excelReader.read(readSheet);
        // 这里千万别忘记关闭，读的时候会创建临时文件，到时磁盘会崩的
        excelReader.finish();

         */
    }

}
