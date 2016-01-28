package com.wavegis.kiwi;
import java.awt.Color;
import java.awt.Rectangle;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.List;

import org.apache.poi.xslf.usermodel.SlideLayout;
import org.apache.poi.xslf.usermodel.XMLSlideShow;
import org.apache.poi.xslf.usermodel.XSLFSlide;
import org.apache.poi.xslf.usermodel.XSLFTable;
import org.apache.poi.xslf.usermodel.XSLFTableCell;
import org.apache.poi.xslf.usermodel.XSLFTableRow;

public class D_addTable extends C_addPictureAndBG {

	private int rowCounter = 0; // 理想行數為7行 故超過7行就新增一投影片
	/**
	 * 新增一個只有標題跟表格的投影片
	 * */
	public XMLSlideShow addTableSlid(XMLSlideShow targetPPT, String title,
			List<String> columnNames, List<List<String>> datas) {
	
		// 產生一個只有標題的新投影片
		XSLFSlide slide = targetPPT.createSlide(targetPPT.getSlideMasters()
				.get(0).getLayout(SlideLayout.TITLE_ONLY));
		// 設定標題
		slide.getPlaceholder(0).setText(title);
		// 新增table
		XSLFTable table = slide.createTable();
		// 設定位置
		table.setAnchor(new Rectangle(50, 100, 450, 300));

		// 測試結果 : 理想表格寬:610 理想表格高 : 350
		Double columnWidth = 610D / new Double(columnNames.size());
		// 先增設表格的header
		XSLFTableRow headRow = table.addRow();
		int columnCounter = 0;
		for (String name : columnNames) {
			XSLFTableCell cell = headRow.addCell();
			table.setColumnWidth(columnCounter++, columnWidth);
			cell.addNewTextParagraph().addNewTextRun().setText(name);
			cell.setFillColor(new Color(150, 205, 205));
		}
		
		for (int i = rowCounter ; i < datas.size() ; i++ ) {
			XSLFTableRow row = table.addRow();
			row.setHeight(55);
			//把該筆資料的內容加入表格
			for (String dataDetail : datas.get(i)){
				XSLFTableCell cell = row.addCell();
				cell.addNewTextParagraph().addNewTextRun()
						.setText(dataDetail);
			}
			rowCounter++ ;
			if (rowCounter % 7 == 0) break ;//資料每7行後跳出迴圈  
		}
		if (rowCounter % 7 == 0) {
			//資料每7行後跳出迴圈  並新增一個新投影片	
			addTableSlid(targetPPT, title, columnNames, datas);
		}
		return targetPPT;
	}
}
