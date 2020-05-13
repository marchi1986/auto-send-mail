package com.marchi.autosendmail.listener;

import com.alibaba.excel.EasyExcel;
import com.alibaba.excel.context.AnalysisContext;
import com.alibaba.excel.event.AnalysisEventListener;
import com.alibaba.fastjson.JSON;
import com.marchi.autosendmail.model.StoreInfo;
import com.marchi.autosendmail.service.MailService;
import org.apache.commons.io.FileUtils;
import org.apache.commons.io.filefilter.FileFilterUtils;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Component;

import java.io.File;
import java.net.URLEncoder;
import java.util.ArrayList;
import java.util.Collection;
import java.util.List;

@Component
public class StoreInfoEasyExcelListener extends AnalysisEventListener<StoreInfo> {


    public  StoreInfoEasyExcelListener(MailService mailService){
        this.mailService=mailService;
    }

    Logger LOGGER = LoggerFactory.getLogger(this.getClass());

    private MailService mailService;

    /**
     * 每隔5条存储数据库，实际使用中可以3000条，然后清理list ，方便内存回收
     */
    private static final int BATCH_COUNT = 5;
    public List<StoreInfo> list = new ArrayList<StoreInfo>();

    /**
     * 这个每一条数据解析都会来调用
     *
     * @param storeInfo
     *            one row value. Is is same as {@link AnalysisContext#readRowHolder()}
     * @param analysisContext
     */
    @Override
    public void invoke(StoreInfo storeInfo, AnalysisContext analysisContext) {

        LOGGER.info("解析到一条数据:{}", JSON.toJSONString(storeInfo));
        list.add(storeInfo);
        // 达到BATCH_COUNT了，需要去存储一次数据库，防止数据几万条数据在内存，容易OOM
        if (list.size() >= BATCH_COUNT) {
            saveData();
            // 存储完成清理 list
            list.clear();
        }
    }

    /**
     * 所有数据解析完成了 都会来调用
     *
     * @param context
     */
    @Override
    public void doAfterAllAnalysed(AnalysisContext context) {
        // 这里也要保存数据，确保最后遗留的数据也存储到数据库
        saveData();
        LOGGER.info("所有数据解析完成！");
    }

    /**
     * 加上存储数据库
     */
    private void saveData() {
        LOGGER.info("{}条数据，开始存储数据库！", list.size());
        for(StoreInfo storeInfo:list){
            // 写法1
            String excelfileName = "F:\\StoreQRCode\\"+storeInfo.getMobile()+".xlsx";
            // 这里 需要指定写用哪个class去写，然后写到第一个sheet，名字为模板 然后文件流会自动关闭
            // 如果这里想使用03 则 传入excelType参数即可
            List<StoreInfo> storeInfos=new ArrayList<StoreInfo>();
            storeInfos.add(storeInfo);
            EasyExcel.write(excelfileName, StoreInfo.class).sheet("模板").doWrite(storeInfos);
            String[] extensions = new String[]{"java", "xml", "txt","png"};
            Collection<File> files=FileUtils.listFiles(new File("F:\\二维码"), extensions,true);
            String pngFileName="";
            try {
            for(File file:files){
                String qrfileName=file.getName();
                //LOGGER.info("QR File Name  {}",qrfileName);
                String[] arry=qrfileName.split("_");
                if(arry.length>0){
                    if(arry[1].equals(storeInfo.getStoreName())){
                        LOGGER.info("QR Code IMAGE {}",file.getAbsolutePath());
                        String newQrCodeFileName=storeInfo.getMerchantStoreCode()+"_QRCODE";
                        File newQrCodeFile=new File("F:\\"+newQrCodeFileName+".png");
                        FileUtils.copyFile(file,newQrCodeFile);
                        pngFileName=newQrCodeFile.getAbsolutePath();
                    }
                }
            }
                String sendTo=storeInfo.getMerchantStoreCode()+"@skechers.cn";
                String[] atts = new String[]{excelfileName, pngFileName,"F:\\Notice.docx"};
                mailService.sendAttachmentsMail("顾客非到店收款二维码", getContent(storeInfo.getStoreName()), "Finance_China@skechers.cn", sendTo, atts);
            }catch (Exception e){
                e.printStackTrace();
            }

            // 写法2
            /*
            fileName = TestFileUtil.getPath() + "simpleWrite" + System.currentTimeMillis() + ".xlsx";
            // 这里 需要指定写用哪个class去写
            ExcelWriter excelWriter = EasyExcel.write(fileName, DemoData.class).build();
            WriteSheet writeSheet = EasyExcel.writerSheet("模板").build();
            excelWriter.write(data(), writeSheet);
            // 千万别忘记finish 会帮忙关闭流
            excelWriter.finish();

             */
        }
        LOGGER.info("存储数据库成功！");
    }

    public String getContent(String name){
        StringBuffer sb=new StringBuffer();
        sb.append("<body lang=ZH-CN link=\"#0563C1\" vlink=\"#954F72\" style='text-justify-trim:punctuation'>");
        sb.append("<div class=WordSection1>");
        sb.append("<p class=MsoNormal style='line-height:22.0pt;mso-line-height-rule:exactly'>");
        sb.append("<span lang=EN-US style='font-size:12.0pt;font-family:华文细黑'>Dear</span>");
        sb.append("<span style='font-size:12.0pt;font-family:华文细黑'>").append(name).append(" <span lang=EN-US><o:p></o:p></span></span>");
        sb.append("</p>");
        sb.append("<p class=MsoNormal style='line-height:22.0pt;mso-line-height-rule:exactly'>");
        sb.append("<span lang=EN-US style='font-family:\"Calibri\",sans-serif'><o:p>&nbsp;</o:p></span>");
        sb.append("</p>");
        sb.append("<p class=MsoNormal style='text-indent:31.1pt;line-height:22.0pt;mso-line-height-rule:exactly'>");
        sb.append("<span style='font-size:12.0pt;font-family:华文细黑'>现顾客非到店情况下收款方案操作指引如下：<span lang=EN-US><o:p></o:p></span></span>");
        sb.append("</p>");
        sb.append("<p class=MsoListParagraph style='margin-left:57.75pt;text-indent:-21.75pt;line-height:22.0pt;mso-line-height-rule:exactly;mso-list:l0 level1 lfo1'>");
        sb.append("<![if !supportLists]><span lang=EN-US style='font-size:12.0pt;font-family:华文细黑'><span style='mso-list:Ignore'>1.<span style='font:7.0pt \"Times New Roman\"'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; </span></span></span><![endif]><span style='font-size:12.0pt;font-family:华文细黑'>非到店顾客通过微信<span lang=EN-US>/</span>支付宝扫描店铺二维码，输入<u>购买商品的交易金额</u>付款给我司（不包含运费）<span lang=EN-US><o:p></o:p></span></span></p><p class=MsoNormal style='text-indent:36.0pt;line-height:22.0pt;mso-line-height-rule:exactly'><u><span style='font-size:12.0pt;font-family:华文细黑;color:red'>注意事项：（<span lang=EN-US>1</span>）店铺员工不可使用个人支付宝、微信等接收客户款项；<span lang=EN-US><o:p></o:p></span></span></u></p><p class=MsoNormal style='text-indent:36.0pt;line-height:22.0pt;mso-line-height-rule:exactly'><span lang=EN-US style='font-size:12.0pt;font-family:华文细黑;color:red'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</span><u><span style='font-size:12.0pt;font-family:华文细黑;color:red'>（<span lang=EN-US>2</span>）只接受公司销售包邮或顾客邮费到付</span></u><span lang=EN-US style='font-size:12.0pt;font-family:华文细黑;color:red'><o:p></o:p></span></p><p class=MsoListParagraph style='margin-left:57.75pt;text-indent:-21.75pt;line-height:22.0pt;mso-line-height-rule:exactly;mso-list:l0 level1 lfo2'>");
        sb.append("<![if !supportLists]><span lang=EN-US style='font-size:12.0pt;font-family:华文细黑'><span style='mso-list:Ignore'>2.<span style='font:7.0pt \"Times New Roman\"'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; </span></span></span><![endif]><span style='font-size:12.0pt;font-family:华文细黑'>店员查询到账情况操作指引如下：<span lang=EN-US><o:p></o:p></span></span></p><p class=MsoListParagraph style='margin-left:96.0pt;text-indent:-36.0pt;line-height:22.0pt;mso-line-height-rule:exactly;mso-list:l1 level1 lfo2'><![if !supportLists]><span lang=EN-US style='font-size:12.0pt;font-family:华文细黑;color:black'><span style='mso-list:Ignore'>A．<span style='font:7.0pt \"Times New Roman\"'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; </span></span></span><![endif]><span style='font-size:12.0pt;font-family:华文细黑;color:black'>店员微信关注【收钱吧】公众号，按指引下载收钱吧<span lang=EN-US>APP</span>：进入【收钱吧】公众号<span lang=EN-US> - </span>开通合作 – 下载<span lang=EN-US>App<o:p></o:p></span></span></p><p class=MsoListParagraph style='margin-left:96.0pt;text-indent:-36.0pt;line-height:22.0pt;mso-line-height-rule:exactly;mso-list:l1 level1 lfo2'><![if !supportLists]><span lang=EN-US style='font-size:12.0pt;font-family:华文细黑;color:black'><span style='mso-list:Ignore'>B．<span style='font:7.0pt \"Times New Roman\"'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; </span></span></span><![endif]><span style='font-size:12.0pt;font-family:华文细黑;color:black'>店员通过收钱吧<span lang=EN-US>APP</span>确认顾客支付情况，（收钱吧<span lang=EN-US>App</span>登陆账号管家卡绑定手机号，详见附件。密码均为<span lang=EN-US>a12345678</span>）<span lang=EN-US><o:p></o:p></span></span></p><p class=MsoListParagraph style='margin-left:57.75pt;text-indent:-21.75pt;line-height:22.0pt;mso-line-height-rule:exactly;mso-list:l0 level1 lfo1'>");
        sb.append("<![if !supportLists]><span lang=EN-US style='font-size:12.0pt;font-family:华文细黑;color:black'><span style='mso-list:Ignore'>3.<span style='font:7.0pt \"Times New Roman\"'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; </span></span></span><![endif]><span style='font-size:12.0pt;font-family:华文细黑;color:black'>伯俊开单：确认款项到账后的当天在伯俊系统开单，付款方式选择【非到店二维码收款】。<span lang=EN-US><o:p></o:p></span></span></p><p class=MsoListParagraph style='margin-left:57.75pt;text-indent:-21.75pt;line-height:22.0pt;mso-line-height-rule:exactly;mso-list:l0 level1 lfo1'>");
        sb.append("<![if !supportLists]><span lang=EN-US style='font-size:12.0pt;font-family:华文细黑;color:black'><span style='mso-list:Ignore'>4.<span style='font:7.0pt \"Times New Roman\"'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; </span></span></span><![endif]><span style='font-size:12.0pt;font-family:华文细黑;color:black'>伯销售退款操作流程：<span lang=EN-US><o:p></o:p></span></span></p><p class=MsoListParagraph style='margin-left:57.75pt;text-indent:0cm;line-height:22.0pt;mso-line-height-rule:exactly'><span lang=EN-US style='font-size:12.0pt;font-family:华文细黑;color:black'>a.&nbsp;&nbsp;&nbsp;&nbsp; </span><span style='font-size:12.0pt;font-family:华文细黑;color:black'>门店提交申请 <span lang=EN-US><o:p></o:p></span></span></p><p class=MsoListParagraph style='margin-left:57.75pt;text-indent:0cm;line-height:22.0pt;mso-line-height-rule:exactly'><span lang=EN-US style='font-size:12.0pt;font-family:华文细黑;color:black'>b.&nbsp;&nbsp;&nbsp;&nbsp; </span><span style='font-size:12.0pt;font-family:华文细黑;color:black'>营运领导审批 <span lang=EN-US><o:p></o:p></span></span></p><p class=MsoListParagraph style='margin-left:57.75pt;text-indent:0cm;line-height:22.0pt;mso-line-height-rule:exactly'><span lang=EN-US style='font-size:12.0pt;font-family:华文细黑;color:black'>c.&nbsp;&nbsp;&nbsp;&nbsp; </span><span style='font-size:12.0pt;font-family:华文细黑;color:black'>财务部确认后在后台退款。<span lang=EN-US><o:p></o:p></span></span></p><p class=MsoNormal><span lang=EN-US><o:p>&nbsp;</o:p></span></p><p class=MsoNormal><span lang=EN-US><o:p>&nbsp;</o:p></span></p>");
        sb.append("</div>");
        sb.append("</body>");
        return sb.toString();
    }

    public static void main(String[] args){
        String[] extensions = new String[]{"java", "xml", "txt","png"};
        Collection<File> files=FileUtils.listFiles(new File("F:\\二维码"), extensions,true);
        for(File file:files){
            try{

            }catch (Exception e){

            }

        }
    }
}
