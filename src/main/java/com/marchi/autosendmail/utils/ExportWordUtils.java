package com.marchi.autosendmail.utils;

import cn.afterturn.easypoi.word.WordExportUtil;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.openxml4j.opc.PackageAccess;
import org.apache.poi.poifs.crypt.EncryptionInfo;
import org.apache.poi.poifs.crypt.EncryptionMode;
import org.apache.poi.poifs.crypt.Encryptor;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.springframework.util.Assert;

import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import java.io.File;
import java.io.FileOutputStream;
import java.io.OutputStream;
import java.net.URLEncoder;
import java.util.Map;

public class ExportWordUtils {

    public static void enctypt(String path) {
        try {
//Add password protection and encrypt the file
            final POIFSFileSystem fs = new POIFSFileSystem();
            final EncryptionInfo info = new EncryptionInfo(EncryptionMode.standard);
//final EncryptionInfo info = new EncryptionInfo(EncryptionMode.agile, CipherAlgorithm.aes256, HashAlgorithm.sha256, -1, -1, null);
            final Encryptor enc = info.getEncryptor();

//set the password
            enc.confirmPassword("abcdef");

//encrypt the file
            final OPCPackage opc = OPCPackage.open(path, PackageAccess.READ_WRITE);
            final OutputStream os = enc.getDataStream(fs);
            opc.save(os);
            opc.close();

//save the file back to the filesystem
            final FileOutputStream fos = new FileOutputStream(path);
            fs.writeFilesystem(fos);
            fos.close();

        } catch (Exception e) {
            e.printStackTrace();
        }

    }

    /**
     * 导出word
     * <p>第一步生成替换后的word文件，只支持docx</p>
     * <p>第二步下载生成的文件</p>
     * <p>第三步删除生成的临时文件</p>
     * 模版变量中变量格式：{{foo}}
     * @param templatePath word模板地址
     * @param temDir 生成临时文件存放地址
     * @param fileName 文件名
     * @param params 替换的参数
     * @param request HttpServletRequest
     * @param response HttpServletResponse
     */
    public static void exportWord(String templatePath, String temDir, String fileName, Map<String, Object> params, HttpServletRequest request, HttpServletResponse response) {
        Assert.notNull(templatePath,"模板路径不能为空");
        Assert.notNull(temDir,"临时文件路径不能为空");
        Assert.notNull(fileName,"导出文件名不能为空");
        Assert.isTrue(fileName.endsWith(".docx"),"word导出请使用docx格式");
        if (!temDir.endsWith("/")){
            temDir = temDir + File.separator;
        }
        File dir = new File(temDir);
        if (!dir.exists()) {
            dir.mkdirs();
        }
        try {
            String userAgent = request.getHeader("user-agent").toLowerCase();
            if (userAgent.contains("msie") || userAgent.contains("like gecko")) {
                fileName = URLEncoder.encode(fileName, "UTF-8");
            } else {
                fileName = new String(fileName.getBytes("utf-8"), "ISO-8859-1");
            }
            XWPFDocument doc = WordExportUtil.exportWord07(templatePath, params);

            String tmpPath = temDir + fileName;
            FileOutputStream fos = new FileOutputStream(tmpPath);
            doc.write(fos);
            enctypt(tmpPath);



        } catch (Exception e) {
            e.printStackTrace();
        } finally {
            //delFileWord(temDir,fileName);//这一步看具体需求，要不要删
        }
    }
    /**
     * 删除零时生成的文件
     */
    public static void delFileWord(String filePath, String fileName){
        File file =new File(filePath+fileName);
        File file1 =new File(filePath);
        file.delete();
        file1.delete();
    }
}
