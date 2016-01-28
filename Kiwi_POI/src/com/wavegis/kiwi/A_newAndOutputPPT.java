package com.wavegis.kiwi;
import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.xslf.usermodel.XMLSlideShow;

public class A_newAndOutputPPT {

	public XMLSlideShow newPPT() {
		return new XMLSlideShow();
	}

	public void outputPPT(XMLSlideShow ppt, String filePath) {
		// 新增一個file -> ppt寫進file中
		File file = new File(filePath);
		try {
			FileOutputStream output = new FileOutputStream(file);
			ppt.write(output);
			output.close();
		} catch (FileNotFoundException e) {
			e.printStackTrace();
		} catch (IOException e) {
			e.printStackTrace();
		}
	}
	
	public File outputPPTFile(XMLSlideShow ppt, String fileName) {
		// 新增一個file -> ppt寫進file中
		File file = new File(fileName+".pptx");
		try {
			FileOutputStream output = new FileOutputStream(file);
			ppt.write(output);
			output.close();
		} catch (FileNotFoundException e) {
			e.printStackTrace();
		} catch (IOException e) {
			e.printStackTrace();
		}
		return file ;
	}

}
