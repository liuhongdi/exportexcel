package com.exportexcel.demo.util;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.springframework.web.context.request.RequestContextHolder;
import org.springframework.web.context.request.ServletRequestAttributes;

import javax.servlet.http.HttpServletResponse;
import java.io.File;
import java.io.FileOutputStream;
import java.io.OutputStream;

public class ExcelUtil {

    //save excel to a file
    public static void  saveExcelFile( HSSFWorkbook wb, String filepath){

        File file = new File(filepath);
        if (file.exists()) {
            file.delete();
        }
        try {
            wb.write(new FileOutputStream(file));
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    //download a execl file
    public static void downExecelFile(HSSFWorkbook wb,String filename) {
        ServletRequestAttributes servletRequestAttributes = (ServletRequestAttributes) RequestContextHolder.getRequestAttributes();
        HttpServletResponse response = servletRequestAttributes.getResponse();
        try {
        // 输出Excel文件
        OutputStream output = response.getOutputStream();
        response.reset();
        //设置文件头
        response.setHeader("Content-Disposition",
                "attchement;filename=" + new String(filename.getBytes("utf-8"),
                        "ISO8859-1"));
        response.setContentType("application/msexcel");
        wb.write(output);
        wb.close();
        } catch (Exception e) {
            e.printStackTrace();
        }
    }
}
