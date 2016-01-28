package com.wavegis.kiwi;

import java.awt.Rectangle;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;

import org.apache.poi.sl.usermodel.PictureData;
import org.apache.poi.util.IOUtils;
import org.apache.poi.xslf.usermodel.SlideLayout;
import org.apache.poi.xslf.usermodel.XMLSlideShow;
import org.apache.poi.xslf.usermodel.XSLFPictureData;
import org.apache.poi.xslf.usermodel.XSLFPictureShape;
import org.apache.poi.xslf.usermodel.XSLFSlide;
import org.apache.poi.xslf.usermodel.XSLFTextBox;

public class E_addPictureAndDescription extends D_addTable{
	/**
	 * 新增一個JPG圖片+敘述的投影片
	 * */
	public XMLSlideShow addJPGPictureAndDescriptionSlide(String title ,String picturePath, String descriptionText,
			XMLSlideShow targetPPT) throws FileNotFoundException, IOException {
		// 產生一個只有標題的新投影片
		XSLFSlide slide = targetPPT.createSlide(targetPPT.getSlideMasters()
				.get(0).getLayout(SlideLayout.TITLE_ONLY));
		// 設定標題
		slide.getPlaceholder(0).setText(title);
		// 把圖片轉成byte模式
		byte[] pictureByte = IOUtils.toByteArray(new FileInputStream(new File(
				picturePath)));
		// 把變成byte的圖片碎片們加到PPT裡面(存成picture data)
		XSLFPictureData pictureData = targetPPT.addPicture(pictureByte,
				PictureData.PictureType.JPEG);
		// 把這個picture data 加入投影片中
		XSLFPictureShape pictureShape = slide.createPicture(pictureData);
		// 設定圖片大小
		pictureShape.setAnchor(new Rectangle(50,120,350 ,400));
		//新增訊息文字框
		XSLFTextBox description = slide.createTextBox() ;
		//設定文字框大小
		description.setAnchor(new Rectangle(400, 120, 280, 400));
		//設定文字框內容
		description.setText(descriptionText);

		return targetPPT;
	}
}
