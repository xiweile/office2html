package com.office2html;

import com.alibaba.fastjson.JSON;
import com.office2html.bean.dto.*;
import com.office2html.test.ExcelResolver;

import java.io.IOException;

public class  Demo {

    public static void main(String[] args) throws  Exception{
        //excelToJsonTest();
        //excelTest();
        wordTest();
        //pdfTest();
    }
    public static void excelToJsonTest() throws IOException {
        ExcelResolver resolver = new ExcelResolver();
       ExcelResultDto dto = resolver.toObject("C:\\Users\\Administrator\\Desktop\\表单模板1.xlsx");
        // ExcelResultDto dto = resolver.toObject("F:\\foresee\\测试指引及测试反馈报告表2.xlsx");
        //ExcelResultDto dto = resolver.toObject("E:\\work\\数字纳服\\智能填单\\修正后的样例1.xlsx");

        //转json,当前活动sheet

        String activeSheetJson = JSON.toJSONString(dto.getSheetList().get(dto.getActiveSheetIndex()).getTable());
        System.out.println(activeSheetJson);

        //转html
      /*  ExcelSheetDto sheetDto=  JSON.parseObject(activeSheetJson, ExcelSheetDto.class);
        String html = sheetDto.getTable().toString();
        System.out.println(html);*/
    }

    public static void excelTest() throws IOException {
        Excel2Html excel2Html = new Excel2Html();
        //ExcelHtmlResultDto dto =(ExcelHtmlResultDto) excel2Html.doc2Html("F:\\foresee\\测试指引及测试反馈报告表.xls");
        //ExcelHtmlResultDto dto =(ExcelHtmlResultDto) excel2Html.doc2Html("F:\\foresee\\测试指引及测试反馈报告表2.xlsx");
        ExcelHtmlResultDto dto =(ExcelHtmlResultDto) excel2Html.doc2Html("C:\\Users\\Administrator\\Desktop\\表单模板1.xlsx");

        System.out.println(dto.getSheetList().get(0).getHtml());
    }

    public static void wordTest() throws IOException {
        Word2Html word2Html = new Word2Html();
        //ExcelHtmlResultDto dto =(ExcelHtmlResultDto) excel2Html.doc2Html("F:\\foresee\\测试指引及测试反馈报告表.xls");
        //WordHtmlDto dto =(WordHtmlDto) word2Html.doc2Html("F:\\foresee\\可办理公积金提取（转移）业务网点.doc");
        WordHtmlDto dto =(WordHtmlDto) word2Html.doc2Html("E:\\work\\数字纳服\\数据API功能设计(1)(1).docx");

        System.out.println(dto.getHtml() );
    }

    public static void pdfTest() throws IOException {
        PDF2Html pdf2Html = new PDF2Html();
        //ExcelHtmlResultDto dto =(ExcelHtmlResultDto) excel2Html.doc2Html("F:\\foresee\\测试指引及测试反馈报告表.xls");
        PDFHtmlResultDto dto =(PDFHtmlResultDto) pdf2Html.doc2Html("F:\\foresee\\炎黄培训内容和安排.pdf");

        System.out.println(dto.getHtml() );
    }
}
