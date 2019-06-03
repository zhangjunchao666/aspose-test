package com.hndfsj.aspose.controller;

import com.aspose.words.*;
import common.ExportWordUtils;
import common.SystemStartUpListener;
import org.springframework.beans.factory.annotation.Value;
import org.springframework.context.annotation.PropertySource;
import org.springframework.core.io.ClassPathResource;
import org.springframework.stereotype.Controller;
import org.springframework.util.ResourceUtils;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.ResponseBody;

import javax.servlet.ServletContext;
import javax.servlet.ServletContextEvent;
import javax.servlet.annotation.WebListener;
import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import javax.swing.filechooser.FileSystemView;
import java.io.File;
import java.io.InputStream;
import java.io.OutputStream;
import java.net.URL;
import java.net.URLEncoder;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;

/**
 * <pre>
 * TODO：
 * </pre>
 *
 * @author zhangjunchao
 * @date 2019/6/2
 */

@Controller
public class WordController {


    /**
     * 获取license证书,消除word水印问题
     * @return
     */
    public static boolean getLicense(HttpServletRequest request) {
        boolean result = false;
        try {
            ClassPathResource pathResource = new ClassPathResource("word/license.xml");
            License aposeLic = new License();
            aposeLic.setLicense(pathResource.getInputStream());
            result = true;
        } catch (Exception e) {
            e.printStackTrace();
        }
        return result;
    }


    @RequestMapping("/down/word")
    @ResponseBody
    public void DownLoadDocs(HttpServletRequest request, HttpServletResponse response)   throws Exception {
        //判断:无License证书,return
        if (!getLicense(request)) {
            return;
        }
        //模板获取的路径是：编译过后的Tomcat下的路径
        // String path_index = request.getSession().getServletContext().getRealPath("/") + "WEB-INF/classes/export_template" + "/faultReportModel.doc";
          String path_index = request.getSession().getServletContext().getContextPath()+ "/resources/WEB-INF/classes/export_template" + "/faultReportModel.doc";




        //创建Document对象，读取Word模板
        //Document document_result = new Document("D:/faultReportModel.doc");
        Document document_result = new Document(path_index);


        // 获取书签
        Range range = document_result.getRange();

        range.replace("tasknum", "123", true, true);
        range.replace("reportor", "张俊超", true, true);
        range.replace("reporttime", LocalDateTime.now().format(DateTimeFormatter.ofPattern("yyyy-MM-dd")), true, true);
        range.replace("faulttime", "2019-5-1", true, true);
        range.replace("equip", "摄像头", true, true);
        range.replace("location", "新政管理处", true, true);


        //操作书签
        DocumentBuilder builder = new DocumentBuilder(document_result);
        //移动到书签位置
        builder.moveToBookmark("faultdescrib");
        String faultdescrib = "摄像头黑屏";
        //写操作
        builder.write(faultdescrib);

        /**
         *取得输出流,保存到流里,返回前端
         * @author zhangjunchao
         * @date 2019/6/2
         */
        OutputStream os = response.getOutputStream();
        /**
         *保存到电脑指定位置,这里指定桌面
         * @author zhangjunchao
         * @date 2019/6/2
         */
        //导出的表格默认路径为桌面
//        FileSystemView fsv = FileSystemView.getFileSystemView();
//        File oss = fsv.getHomeDirectory();
        //System.out.println(com.getPath());  C:\Users\lenovo\Desktop

        //定义fileType
       /* String fileType = "doc";
        if (fileType == null) {
            fileType = "doc";
        }*/

        String fileType = "pdf";
        if (fileType == null) {
            fileType = "pdf";
        }

        OutputStream oss= response.getOutputStream();
        if (fileType.equals("pdf")) {

            response.setContentType("application/pdf");

            response.setHeader("Content-disposition", "attachment;filename=" + URLEncoder.encode(("故障维修报告"+".pdf").replace(" ", ""), "UTF-8"));
            //保存为PDF文档：SaveFormat.PDF
           // document_result.save(oss.getPath()+"\\a.pdf", SaveFormat.PDF);
            document_result.save(oss, SaveFormat.PDF);
        } else {
            response.setContentType("application/vnd.ms-word");
            response.setHeader("Content-disposition", "attachment;filename=" + URLEncoder.encode(("故障维修报告" + ".doc").replace(" ", ""), "UTF-8"));
            //保存为doc word文档：SaveFormat.DOC
           // document_result.save(oss.getPath()+"\\b.doc", SaveFormat.DOC);
            document_result.save(oss, SaveFormat.DOC);
        }
        /**
         * 用流的情况下,记得关闭
         */
        os.close();

    }


}
