package com.hndfsj.aspose.controller;


import common.ExportWordUtils;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.RequestMapping;

import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import java.util.HashMap;
import java.util.Map;

/**
 * TODO easypoi生成word文档
 * @author zhangjunchao
 * @date 2019/6/3
 */
@Controller
public class TestWord {

    @RequestMapping("/demo/export")
    public void export(HttpServletRequest request, HttpServletResponse response){
        Map<String,Object> params = new HashMap<>();
        params.put("a","20190602110");
        params.put("b","张俊超");
        params.put("c","2019-5-30");
        params.put("d","2019-5-30");
        params.put("e","摄像头");
        params.put("f","新政管理处");
        //这里是我说的一行代码
        ExportWordUtils.exportWord("word/export.docx","D:/test","aaa.docx",params,request,response);
    }





}
