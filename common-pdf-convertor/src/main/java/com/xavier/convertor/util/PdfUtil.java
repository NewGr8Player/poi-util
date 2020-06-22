package com.xavier.convertor.util;

import com.lowagie.text.Font;
import com.lowagie.text.pdf.BaseFont;
import fr.opensagres.poi.xwpf.converter.pdf.PdfConverter;
import fr.opensagres.poi.xwpf.converter.pdf.PdfOptions;
import org.apache.commons.io.FileUtils;
import org.apache.poi.xwpf.usermodel.XWPFDocument;

import java.awt.*;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.OutputStream;

public class PdfUtil {

    // TODO NO test NO COMMENT
    public static void convertToPdf(File inputFile,File pdfFile) {
        try {

            // 1) Load docx with POI XWPFDocument
            XWPFDocument document = new XWPFDocument(new FileInputStream(inputFile));

            // 2) Convert POI XWPFDocument 2 PDF with iText
            FileUtils.touch(pdfFile);
            OutputStream out = new FileOutputStream(pdfFile);

            BaseFont chinese = BaseFont.createFont("/DEFAULT.TTF", BaseFont.IDENTITY_H,BaseFont.NOT_EMBEDDED);

            PdfOptions options = PdfOptions.create()
                    .fontProvider((String familyName, String encoding, float size, int style, Color color) -> {
                        try {
                            Font font = new Font(chinese, size, style, Color.black);
                            if(familyName != null){
                                font.setFamily(familyName);
                            }
                            return font;
                        } catch (Exception e) {
                            e.printStackTrace();
                        }
                        return null;
                    })
                    ;
            PdfConverter.getInstance().convert(document, out, options);
        } catch (Throwable e) {
            e.printStackTrace();
        }
    }
}
