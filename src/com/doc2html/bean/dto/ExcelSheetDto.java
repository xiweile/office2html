package com.doc2html.bean.dto;

import com.doc2html.bean.HtmlTable;
import lombok.Getter;
import lombok.Setter;

@Setter
@Getter
public class ExcelSheetDto {

    private String sheetName;
    private Integer sheetIndex;
    private HtmlTable table;
}
