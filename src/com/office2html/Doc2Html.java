package com.office2html;

import java.io.IOException;

import com.office2html.bean.dto.DocHtmlDto;

public interface Doc2Html {

	public DocHtmlDto doc2Html(String filePath) throws IOException;
}
