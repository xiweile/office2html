package com.office2html.bean.dto;

import lombok.Getter;
import lombok.Setter;

import java.util.List;

@Setter
@Getter
public class ExcelResultDto extends DocHtmlDto {

	private List<ExcelSheetDto> sheetList;
	private Integer activeSheetIndex;
}
