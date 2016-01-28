package com.wavegis.kiwi;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

import org.apache.poi.xslf.usermodel.XMLSlideShow;

public class PPTMaker extends E_addPictureAndDescription {
	/**
	 * PPT投影片實作包二版
	 * (新增圖片+敘述投影片)
	 * <pre>
	 * 發行日:20151101 
	 * 開發者:kiwi
	 * 用途:簡化PPT的製作流程
	 * 目前可用功能:<br>
	 * makeADataPPTFile<br>
	 * makeAPictureAndDescriptionPPTFile
	 * */
	public PPTMaker() {
	}
	/**
	 * 產生並輸出一個PPT的File以供輸出or另外使用
	 * 
	 * <pre>
	 * @param fileName 檔案名稱
	 * @param pptTitle 投影片標題(可null)
	 * @param pptSubTitle 副標題(可null)、
	 * @param columnNames 表格header、
	 * @param datas 資料(以List存List形式 - 每一筆資料都存成List(String)，再放入一個List中)、
	 * @param backgroundImagePath 背景圖(可null)
	 * @param imageType 背景圖格式("JPEG" or "PNG")
	 * @param organizationName 機構單位(可null)
	 * @throws IOException
	 * @throws FileNotFoundException
	 * */
	public File makeADataPPTFile(String fileName, String pptTitle,
			String pptSubTitle, String backgroundImagePath, String imageType,
			List<String> columnNames, List<List<String>> datas, String organizationName)
			throws FileNotFoundException, IOException {

		XMLSlideShow ppt = this.newPPT();
		// 新增標題投影片
		if (pptTitle != null && pptSubTitle != null)
			this.addTitleSlide(ppt, pptTitle, pptSubTitle);
		// 設定背景
		if (backgroundImagePath != null && imageType.equals("JPEG")) {
			this.addJPGBackGround(backgroundImagePath, ppt);
		} else if (backgroundImagePath != null && imageType.equals("PNG")) {
			this.addPNGBackGround(backgroundImagePath, ppt);
		}
		// 產生資料投影片
		if (columnNames != null && datas != null)
			this.addTableSlid(ppt, pptTitle, columnNames, datas);
		// 產生結尾投影片
		if (organizationName != null)
			this.addEndSlide(ppt, organizationName);

		return this.outputPPTFile(ppt, fileName);
	}
	/**產生一個內容為圖片+描述框的PPT檔
	 * <pre>
	 * @param fileName 檔案名稱
	 * @param pptTitle 投影片標題(可null)
	 * @param pptSubTitle 副標題(可null)、
	 * @param backgroundImagePath 背景圖(可null)
	 * @param BGimageType 背景圖格式("JPEG" or "PNG")
	 * @param imagePath 圖片(JPEG格式)
	 * @param description 右側說明文字
	 * @param organizationName 機構單位(可null)
	 * @throws IOException 
	 * @throws FileNotFoundException 
	 * */
	public File makeAPictureAndDescriptionPPTFile(String fileName, String pptTitle,
			String pptSubTitle, String backgroundImagePath, String BGimageType, String imagePath,
			String description, String organizationName) throws FileNotFoundException, IOException{
		XMLSlideShow ppt = this.newPPT();
		// 新增標題投影片
		if (pptTitle != null && pptSubTitle != null)
			this.addTitleSlide(ppt, pptTitle, pptSubTitle);
		// 設定背景
		if (backgroundImagePath != null && BGimageType.equals("JPEG")) {
			this.addJPGBackGround(backgroundImagePath, ppt);
		} else if (backgroundImagePath != null && BGimageType.equals("PNG")) {
			this.addPNGBackGround(backgroundImagePath, ppt);
		}
		// 產生資料投影片
		this.addJPGPictureAndDescriptionSlide(pptTitle, imagePath, description, ppt);
		// 產生結尾投影片
		if (organizationName != null)
			this.addEndSlide(ppt, organizationName);

		return this.outputPPTFile(ppt, fileName);
	}

	public static void main(String[] args) {
		String desktopPath = "C:\\Users\\Kiwi\\Desktop";
		PPTMaker pptMaker = new PPTMaker() ;
		try {
			pptMaker.makeAPictureAndDescriptionPPTFile(desktopPath+"\\makePictureTest", "PictureTest", "subtitleTest", desktopPath+"\\image2.jpg", "JPEG", desktopPath+"\\image.jpg", "圖文測試description", "奇異鳥公司");
		} catch (FileNotFoundException e) {
			e.printStackTrace();
		} catch (IOException e) {
			e.printStackTrace();
		}
	}
}
