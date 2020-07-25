package com.exportexcel.demo.controller;

import com.exportexcel.demo.mapper.GoodsMapper;
import com.exportexcel.demo.pojo.Goods;
import com.exportexcel.demo.util.ExcelUtil;
import org.apache.poi.hssf.usermodel.*;

import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.ss.util.CellRangeAddress;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RestController;

import javax.annotation.Resource;
import java.util.List;


@RestController
@RequestMapping("/home")
public class HomeController {

    @Resource
    private GoodsMapper goodsMapper;

    @GetMapping("/exportexcel")
    public void exportExcel() {

        HSSFWorkbook wb = new HSSFWorkbook();
        // 根据页面index 获取sheet页
        HSSFSheet sheet = wb.createSheet("商品基本信息");
        createTitle(wb,sheet);

        //设置列宽度
         List<Goods> goodsList = goodsMapper.selectAllGoods();
         //表内容从第二行开始
         int i=1;
         for (Goods goodsOne : goodsList) {
             HSSFRow row = sheet.createRow(i + 1);
             //创建HSSFCell对象 设置单元格的值
             row.createCell(0).setCellValue(goodsOne.getGoodsId());
             row.createCell(1).setCellValue(goodsOne.getGoodsName());
             row.createCell(2).setCellValue(goodsOne.getSubject());
             row.createCell(3).setCellValue(goodsOne.getStock());
             i++;
         }

         //保存成文件
        ExcelUtil.saveExcelFile(wb,"/data/file/html/商品信息.xls");
        //下载文件
        ExcelUtil.downExecelFile(wb,"商品信息.xls");
     }


    //创建标题和表头
    private void createTitle(HSSFWorkbook workbook,HSSFSheet sheet){
        //用CellRangeAddress合并单元格，作为标题
        //参数1：起始行 参数2：终止行 参数3：起始列 参数4：终止列
        CellRangeAddress region1 = new CellRangeAddress(0, 0, 0, 3);
        sheet.addMergedRegion(region1);

        HSSFRow row_title = sheet.createRow(0);
        row_title.setHeightInPoints(26);
        HSSFCell cell_title;
        cell_title = row_title.createCell(0);
        cell_title.setCellValue("商品信息表");

        //设置标题样式:居中加粗
        HSSFCellStyle style_title = workbook.createCellStyle();
        HSSFFont font_title = workbook.createFont();
        font_title.setBold(true);
        font_title.setFontHeightInPoints((short) 20);
        font_title.setColor(HSSFColor.HSSFColorPredefined.BLUE.getIndex());
        style_title.setFont(font_title);
        style_title.setAlignment(HorizontalAlignment.CENTER);
        style_title.setVerticalAlignment(VerticalAlignment.CENTER);
        cell_title.setCellStyle(style_title);


        //以下为表头
        HSSFRow row = sheet.createRow(1);
        //设置行高
        row.setHeightInPoints(18);
        //设置列宽度
        sheet.setColumnWidth(0,10*256);
        sheet.setColumnWidth(1,30*256);
        sheet.setColumnWidth(2,30*256);
        sheet.setColumnWidth(3,10*256);
        //设置样式:居中加粗
        HSSFCellStyle style = workbook.createCellStyle();
        HSSFFont font = workbook.createFont();
        font.setBold(true);
        font.setColor(HSSFColor.HSSFColorPredefined.GREEN.getIndex());
        style.setFont(font);
        style.setAlignment(HorizontalAlignment.CENTER);
        style.setVerticalAlignment(VerticalAlignment.CENTER);
        //添加4个字段
        HSSFCell cell;
        cell = row.createCell(0);
        cell.setCellValue("id");
        cell.setCellStyle(style);

        cell = row.createCell(1);
        cell.setCellValue("商品名称");
        cell.setCellStyle(style);

        cell = row.createCell(2);
        cell.setCellValue("商品描述");
        cell.setCellStyle(style);

        cell = row.createCell(3);
        cell.setCellValue("库存数量");
        cell.setCellStyle(style);
    }
}