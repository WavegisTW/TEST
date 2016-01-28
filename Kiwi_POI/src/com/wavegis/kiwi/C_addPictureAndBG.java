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


public class C_addPictureAndBG extends B_newASlid {

	/**
	 * 新增一個單純只有JPG圖片(沒有標題)的投影片
	 * */
	public XMLSlideShow addJPGPictureSlide(String picturePath,
			XMLSlideShow targetPPT) throws FileNotFoundException, IOException {
		// 新增投影片
		XSLFSlide slide = targetPPT.createSlide();
		// 把圖片轉成byte模式
		byte[] pictureByte = IOUtils.toByteArray(new FileInputStream(new File(
				picturePath)));
		// 把變成byte的圖片碎片們加到PPT裡面(存成picture data)
		XSLFPictureData pictureData = targetPPT.addPicture(pictureByte,
				PictureData.PictureType.JPEG);
		// 把這個picture data 加入投影片中
		XSLFPictureShape pictureShape = slide.createPicture(pictureData);
		// 設定圖片大小
		pictureShape.setAnchor(new Rectangle(targetPPT.getPageSize()));

		return targetPPT;
	}
	/**
	 * 設定投影片的背景圖片(JPG)
	 * */
	public XMLSlideShow addJPGBackGround(String picturePath, XMLSlideShow targetPPT)
			throws FileNotFoundException, IOException {

		// 把圖片轉成byte模式
		byte[] pictureByte = IOUtils.toByteArray(new FileInputStream(new File(
				picturePath)));
		// 把變成byte的圖片碎片們加到PPT裡面(存成picture data)
		XSLFPictureData pictureData = targetPPT.addPicture(pictureByte,
				PictureData.PictureType.JPEG);
		//把picture data 存進slideMaster中(若存在slide中會變圖片,存在master中則會變背景)
		 XSLFPictureShape BGShape = targetPPT.getSlideMasters().get(0).createPicture(pictureData);
		// 設定背景圖片大小
		BGShape.setAnchor(new Rectangle(targetPPT.getPageSize()));
		
		return targetPPT;
	}
	/**
	 * 新增一個單純只有PNG圖片(沒有標題)的投影片
	 * */
	public XMLSlideShow addPNGPictureSlide(String picturePath,
			XMLSlideShow targetPPT) throws FileNotFoundException, IOException {
		// 新增投影片
		XSLFSlide slide = targetPPT.createSlide();
		// 把圖片轉成byte模式
		byte[] pictureByte = IOUtils.toByteArray(new FileInputStream(new File(
				picturePath)));
		// 把變成byte的圖片碎片們加到PPT裡面(存成picture data)
		XSLFPictureData pictureData = targetPPT.addPicture(pictureByte,
				PictureData.PictureType.PNG);
		// 把這個picture data 加入投影片中
		XSLFPictureShape pictureShape = slide.createPicture(pictureData);
		// 設定圖片大小
		pictureShape.setAnchor(new Rectangle(targetPPT.getPageSize()));

		return targetPPT;
	}
	/**
	 * 設定投影片的背景圖片(PNG)
	 * */
	public XMLSlideShow addPNGBackGround(String picturePath, XMLSlideShow targetPPT)
			throws FileNotFoundException, IOException {

		// 把圖片轉成byte模式
		byte[] pictureByte = IOUtils.toByteArray(new FileInputStream(new File(
				picturePath)));
		// 把變成byte的圖片碎片們加到PPT裡面(存成picture data)
		XSLFPictureData pictureData = targetPPT.addPicture(pictureByte,
				PictureData.PictureType.PNG);
		//把picture data 存進slideMaster中(若存在slide中會變圖片,存在master中則會變背景)
		 XSLFPictureShape BGShape = targetPPT.getSlideMasters().get(0).createPicture(pictureData);
		// 設定背景圖片大小
		BGShape.setAnchor(new Rectangle(targetPPT.getPageSize()));
		
		return targetPPT;
	}
}
