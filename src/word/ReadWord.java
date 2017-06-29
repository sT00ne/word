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

import org.apache.poi.util.Units;
import org.apache.poi.xwpf.extractor.XWPFWordExtractor;
import org.apache.poi.xwpf.usermodel.*;

public class ReadWord {
    public static void main(String[] args) throws Exception {
        finalDemo();
    }

    /**
     * create blank document
     *
     * @throws Exception
     */
    public static void blankDocument() throws Exception {
        /* Blank Document */
        XWPFDocument document = new XWPFDocument();
        //Write the Document in file system
        FileOutputStream out = new FileOutputStream(new File("createdocument.docx"));
        document.write(out);
        out.close();
        System.out.println("createdocument.docx written successully");
    }

    /**
     * add image
     *
     * @throws Exception
     */
    public static void addImage() throws Exception {
        /* Add Picture */
        XWPFDocument document = new XWPFDocument();
        FileInputStream fis = new FileInputStream("C:\\Koala.jpg");
        XWPFParagraph p = document.createParagraph();
        XWPFRun r = p.createRun();
        r.setText("123");
        r.addBreak();
        r.addPicture(fis, Document.PICTURE_TYPE_JPEG, "Koala.jpg",
                Units.toEMU(200), Units.toEMU(200));
        r.addBreak(BreakType.PAGE);
        FileOutputStream out = new FileOutputStream("images.docx");
        document.write(out);
        out.close();
        document.close();
        System.out.println("success");
    }

    /**
     * read docx
     *
     * @throws Exception
     */
    public static void readDocument() throws Exception {
        /* Read Document */
        File file = null;
        XWPFWordExtractor extractor = null;
        try {
            file = new File("D:\\word\\Test.docx");
            FileInputStream fis = new FileInputStream(file.getAbsolutePath());
            XWPFDocument document = new XWPFDocument(fis);

            List<XWPFParagraph> el = document.getParagraphs();
            for (XWPFParagraph a : el) {
                List<IRunElement> text = a.getIRuns();
                String phtext = a.getParagraphText();
                ParagraphAlignment align = a.getAlignment();
            }

            extractor = new XWPFWordExtractor(document);
            String fileData = extractor.getText();
            System.out.println(fileData);
        } catch (Exception exep) {
            exep.printStackTrace();
        }
    }

    /**
     * get image and save
     *
     * @throws Exception
     */
    public static void getImage() throws Exception {
        /* Get Picture */
        File file = null;
        file = new File("c:\\Test.docx");
        FileInputStream fis = new FileInputStream(file.getAbsolutePath());
        XWPFDocument docx = new XWPFDocument(fis);
        List<XWPFPictureData> piclist = docx.getAllPictures();
        Iterator<XWPFPictureData> iterator = piclist.iterator();
        int i = 0;
        while (iterator.hasNext()) {
            XWPFPictureData pic = iterator.next();
            byte[] bytepic = pic.getData();
            BufferedImage imag = ImageIO.read(new ByteArrayInputStream(bytepic));
            ImageIO.write(imag, "jpg",
                    new File("C:/imagefromword" + i + ".jpg"));
            i++;
        }
        System.out.println("success");
    }

    /**
     * extract images and text and save to 2 docx without format
     *
     * @throws Exception
     */
    public static void simpleDemo() throws Exception {
        File file = null;
        file = new File("D:\\word\\Test.docx");
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
        FileOutputStream tout = new FileOutputStream("D://word/text.docx");
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
            ImageIO.write(imag, "jpg", new File("D://word/temp/imagefromword" + i + ".jpg"));
            i++;
        }
        System.out.println("Extract Picture Success");
        File filePath = new File("D://word/temp");
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
            FileInputStream pic = new FileInputStream("D://word/temp/" + n);
            r.addPicture(pic, Document.PICTURE_TYPE_JPEG, n,
                    Units.toEMU(200), Units.toEMU(200));
        }

        r.addBreak(BreakType.PAGE);
        FileOutputStream out = new FileOutputStream("D://word/images.docx");
        document.write(out);
        out.close();
        document.close();
        System.out.println("Export Images Success");
    }

    /**
     * extract images and text and save to 2 docx with format
     * @throws Exception
     */
    public static void finalDemo() throws Exception {
        File file = null;
        file = new File("D:\\word\\Test.docx");
        FileInputStream fis = new FileInputStream(file.getAbsolutePath());
        FileInputStream fis_pic = new FileInputStream(file.getAbsolutePath());
        XWPFDocument document = new XWPFDocument(fis);
        XWPFDocument document_pic = new XWPFDocument(fis_pic);
        XWPFParagraph imageParagraph = null;
        Integer paraLenth = 0;
        List<Integer> textPos = new ArrayList<>();
        List<Integer> picPos = new ArrayList<>();
        List<IBodyElement> documentElements = document.getBodyElements();
        //get image index
        for (IBodyElement documentElement : documentElements) {
            if (documentElement instanceof XWPFParagraph) {
                imageParagraph = (XWPFParagraph) documentElement;
                if (imageParagraph != null && imageParagraph.getCTP() != null && imageParagraph.getCTP().toString().trim().indexOf("pic:cNvPr") != -1) {
                    picPos.add(paraLenth);
                } else {
                    textPos.add(paraLenth);
                }
            }
            paraLenth++;
        }
        //add image index
        Integer i = 0;
        for (Integer pos : picPos) {
            document.removeBodyElement(pos - i);
            i++;
        }
        //add text index
        i = 0;
        for (Integer tpos : textPos) {
            document_pic.removeBodyElement(tpos - i);
            i++;
        }
        System.out.println(i);
        FileOutputStream fos = new FileOutputStream("D:\\word\\text.docx");
        document.write(fos);
        fos.close();
        FileOutputStream picfos = new FileOutputStream("D:\\word\\image.docx");
        document_pic.write(picfos);
        picfos.close();
    }

    public static XWPFDocument replaceImage(XWPFDocument document, String imageOldName, String imagePathNew, int newImageWidth, int newImageHeight) throws Exception {
        try {
            System.out.print("replaceImage: old=" + imageOldName + ", new=" + imagePathNew);

            int imageParagraphPos = -1;
            XWPFParagraph imageParagraph = null;

            List<IBodyElement> documentElements = document.getBodyElements();
            for (IBodyElement documentElement : documentElements) {
                imageParagraphPos++;
                if (documentElement instanceof XWPFParagraph) {
                    imageParagraph = (XWPFParagraph) documentElement;
                    if (imageParagraph != null && imageParagraph.getCTP() != null && imageParagraph.getCTP().toString().trim().indexOf(imageOldName) != -1) {
                        document.removeBodyElement(document.getBodyElements().size() - 1);
                        break;
                    }
                }
            }

            if (imageParagraph == null) {
                throw new Exception("Unable to replace image data due to the exception:\n"
                        + "'" + imageOldName + "' not found in in document.");
            }
            ParagraphAlignment oldImageAlignment = imageParagraph.getAlignment();

            // remove old image
            document.removeBodyElement(imageParagraphPos);

            // now add new image

            // BELOW LINE WILL CREATE AN IMAGE
            // PARAGRAPH AT THE END OF THE DOCUMENT.
            // REMOVE THIS IMAGE PARAGRAPH AFTER
            // SETTING THE NEW IMAGE AT THE OLD IMAGE POSITION
            XWPFParagraph newImageParagraph = document.createParagraph();
            XWPFRun newImageRun = newImageParagraph.createRun();
            //newImageRun.setText(newImageText);
            newImageParagraph.setAlignment(oldImageAlignment);
            try (FileInputStream is = new FileInputStream(imagePathNew)) {
                newImageRun.addPicture(is, XWPFDocument.PICTURE_TYPE_JPEG, imagePathNew,
                        Units.toEMU(newImageWidth), Units.toEMU(newImageHeight));
            }

            // set new image at the old image position
            document.setParagraph(newImageParagraph, imageParagraphPos);

            // NOW REMOVE REDUNDANT IMAGE FORM THE END OF DOCUMENT
            document.removeBodyElement(document.getBodyElements().size() - 1);

            return document;
        } catch (Exception e) {
            throw new Exception("Unable to replace image '" + imageOldName + "' due to the exception:\n" + e);
        }
    }

}
