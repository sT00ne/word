package word;

import java.awt.image.BufferedImage;
import java.io.ByteArrayInputStream;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;

import javax.imageio.ImageIO;

import org.apache.poi.hwpf.HWPFDocument;
import org.apache.poi.hwpf.extractor.WordExtractor;
import org.apache.poi.util.Units;
import org.apache.poi.xwpf.extractor.XWPFWordExtractor;
import org.apache.poi.xwpf.usermodel.BreakType;
import org.apache.poi.xwpf.usermodel.Document;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFPictureData;
import org.apache.poi.xwpf.usermodel.XWPFRun;

public class ReadWord {
	public static void main(String[] args) throws Exception {
		/** Blank Document **/
		// XWPFDocument document= new XWPFDocument();
		// //Write the Document in file system
		// FileOutputStream out = new FileOutputStream(new
		// File("createdocument.docx"));
		// document.write(out);
		// out.close();
		// System.out.println("createdocument.docx written successully");

		/** Add Picture **/
		// XWPFDocument document = new XWPFDocument();
		// FileInputStream fis = new FileInputStream("C:\\Koala.jpg");
		// XWPFParagraph p = document.createParagraph();
		// XWPFRun r = p.createRun();
		// r.setText("123");
		// r.addBreak();
		// r.addPicture(fis, Document.PICTURE_TYPE_JPEG, "Koala.jpg",
		// Units.toEMU(200), Units.toEMU(200));
		// r.addBreak(BreakType.PAGE);
		// FileOutputStream out = new FileOutputStream("images.docx");
		// document.write(out);
		// out.close();
		// document.close();
		// System.out.println("success");

		/** Read Document **/
		// File file = null;
		// XWPFWordExtractor extractor = null;
		// try {
		// file = new File("c:\\Test.docx");
		// FileInputStream fis = new FileInputStream(file.getAbsolutePath());
		// XWPFDocument document = new XWPFDocument(fis);
		// extractor = new XWPFWordExtractor(document);
		// String fileData = extractor.getText();
		// System.out.println(fileData);
		// } catch (Exception exep) {
		// exep.printStackTrace();
		// }

		/** Get Picture **/
		// File file = null;
		// file = new File("c:\\Test.docx");
		// FileInputStream fis = new FileInputStream(file.getAbsolutePath());
		// XWPFDocument docx = new XWPFDocument(fis);
		// List<XWPFPictureData> piclist = docx.getAllPictures();
		// Iterator<XWPFPictureData> iterator = piclist.iterator();
		// int i = 0;
		// while (iterator.hasNext()) {
		// XWPFPictureData pic = iterator.next();
		// byte[] bytepic = pic.getData();
		// BufferedImage imag = ImageIO.read(new ByteArrayInputStream(bytepic));
		// ImageIO.write(imag, "jpg", new File("C:/imagefromword" + i +
		// ".jpg"));
		// i++;
		// }
		// System.out.println("success");

		File file = null;
		file = new File("E:\\word\\Test.docx");
		FileInputStream fis = new FileInputStream(file.getAbsolutePath());
		XWPFDocument docx = new XWPFDocument(fis);
		// Text
		XWPFWordExtractor extractor = new XWPFWordExtractor(docx);
		String fileText = extractor.getText();
		System.out.println("Extract Text Success");
		XWPFDocument textDocument = new XWPFDocument();
		XWPFParagraph tp = textDocument.createParagraph();
		XWPFRun tr = tp.createRun();
		tr.setText(fileText);
		tr.addBreak(BreakType.PAGE);
		FileOutputStream tout = new FileOutputStream("E://word/text.docx");
		textDocument.write(tout);
		tout.close();
		textDocument.close();
		System.out.println("Export Text Success");
		// Picture
		List<XWPFPictureData> piclist = docx.getAllPictures();
		Iterator<XWPFPictureData> iterator = piclist.iterator();
		int i = 0;
		//Get picture and save
		while (iterator.hasNext()) {
			XWPFPictureData pic = iterator.next();
			byte[] bytepic = pic.getData();
			BufferedImage imag = ImageIO.read(new ByteArrayInputStream(bytepic));
			ImageIO.write(imag, "jpg", new File("E://word/temp/imagefromword" + i + ".jpg"));
			i++;
		}
		System.out.println("Extract Picture Success");
		File filePath = new File("E://word/temp");
		File[] files = filePath.listFiles(new ImageFileFilter());
		List<String> fileName = new ArrayList<String>();
		for (File f : files) {
			fileName.add(f.getName());
			// System.out.println("file: " + f.getName());
		}
		XWPFDocument document = new XWPFDocument();
		XWPFParagraph p = document.createParagraph();
		XWPFRun r = p.createRun();
		for (String n : fileName) {
			FileInputStream pic = new FileInputStream("E://word/temp/" + n);
			r.addPicture(pic, Document.PICTURE_TYPE_JPEG, n, Units.toEMU(200), Units.toEMU(200));
		}
		r.addBreak(BreakType.PAGE);
		FileOutputStream out = new FileOutputStream("E://word/images.docx");
		document.write(out);
		out.close();
		document.close();
		System.out.println("Export Images Success");
	}
}
