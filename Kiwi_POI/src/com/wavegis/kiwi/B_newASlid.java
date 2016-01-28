package com.wavegis.kiwi;
import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.xslf.usermodel.SlideLayout;
import org.apache.poi.xslf.usermodel.XMLSlideShow;
import org.apache.poi.xslf.usermodel.XSLFSlide;
import org.apache.poi.xslf.usermodel.XSLFSlideLayout;
import org.apache.poi.xslf.usermodel.XSLFSlideMaster;
import org.apache.poi.xslf.usermodel.XSLFTextShape;


public class B_newASlid extends A_newAndOutputPPT{

	/**
	 * 增加一個標題投影片
	 * */
	public XMLSlideShow addTitleSlide(XMLSlideShow targetPPT ,String title, String subTitle){
		// 取出ppt的slideMaster 
		XSLFSlideMaster slideMaster = targetPPT.getSlideMasters().get(0);
		//取出slidMaster中的各種layout
		XSLFSlideLayout slideLayout = slideMaster.getLayout(SlideLayout.TITLE);
		//讓PPT新增一個這種layout的投影片
		XSLFSlide slide01 = targetPPT.createSlide(slideLayout);

		// 設定此頁面的文字
		XSLFTextShape titleText = slide01.getPlaceholder(0);
		XSLFTextShape subtitleText = slide01.getPlaceholder(1);
		titleText.setText(title);
		subtitleText.setText(subTitle);
		
		return targetPPT ;
	}	
	/**
	 * 增加一個純標題投影片(標題在最上面)
	 * */
	public XMLSlideShow addOnlyTitleSlide(XMLSlideShow targetPPT ,String title){
		// 取出ppt的slideMaster 
		XSLFSlideMaster slideMaster = targetPPT.getSlideMasters().get(0);
		//取出slidMaster中的各種layout
		XSLFSlideLayout slideLayout = slideMaster.getLayout(SlideLayout.TITLE_ONLY);
		//讓PPT新增一個這種layout的投影片
		XSLFSlide slide01 = targetPPT.createSlide(slideLayout);

		// 設定此頁面的文字
		XSLFTextShape titleText = slide01.getPlaceholder(0);
		titleText.setText(title);

		return targetPPT ;
	}
	/**
	 * 增加一個最普遍的投影片(標題在上面下面是文字)
	 * */
	public XMLSlideShow addTitleAndContentSlide(XMLSlideShow targetPPT ,String title, String content){
		// 取出ppt的slideMaster 
		XSLFSlideMaster slideMaster = targetPPT.getSlideMasters().get(0);
		//取出slidMaster中的各種layout
		XSLFSlideLayout slideLayout = slideMaster.getLayout(SlideLayout.TITLE_AND_CONTENT);
		//讓PPT新增一個這種layout的投影片
		XSLFSlide slide01 = targetPPT.createSlide(slideLayout);

		// 設定此頁面的文字
		XSLFTextShape titleText = slide01.getPlaceholder(0);
		XSLFTextShape subtitleText = slide01.getPlaceholder(1);
		titleText.setText(title);
		subtitleText.setText(content);
		
		return targetPPT ;
	}
	/**
	 * 增加一個結尾投影片the End
	 * */
	public XMLSlideShow addEndSlide(XMLSlideShow targetPPT ,String organizationName){
		// 取出ppt的slideMaster 
		XSLFSlideMaster slideMaster = targetPPT.getSlideMasters().get(0);
		//取出slidMaster中的各種layout
		XSLFSlideLayout slideLayout = slideMaster.getLayout(SlideLayout.SECTION_HEADER);
		//讓PPT新增一個這種layout的投影片
		XSLFSlide slide01 = targetPPT.createSlide(slideLayout);

		// 設定此頁面的文字
		XSLFTextShape titleText = slide01.getPlaceholder(0);
		titleText.setText("~THE END~");
		XSLFTextShape titleText2 = slide01.getPlaceholder(1);
		titleText2.setText(organizationName);

		return targetPPT ;
	}
}
