package mn.digz.unicodeconverter.powerpoint;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.nio.file.FileSystems;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.List;
import mn.digz.unicodeconverter.lang.MongolianLanguageUtil;
import org.apache.poi.hslf.usermodel.HSLFSlide;
import org.apache.poi.hslf.usermodel.HSLFSlideShow;
import org.apache.poi.hslf.usermodel.HSLFTextParagraph;
import org.apache.poi.xslf.usermodel.XMLSlideShow;
import org.apache.poi.xslf.usermodel.XSLFPictureData;
import org.apache.poi.xslf.usermodel.XSLFPictureShape;
import org.apache.poi.xslf.usermodel.XSLFShape;
import org.apache.poi.xslf.usermodel.XSLFSlide;
import org.apache.poi.xslf.usermodel.XSLFTextShape;

/**
 *
 * @author MethoD
 */
public class PowerPointConverter {

    public static void convert(String input, String output) throws Exception {
        if(input != null && output != null) {
            Path inputPath = FileSystems.getDefault().getPath(input);
            String contentType = Files.probeContentType(inputPath);

            //SlideShow slideshow = null;
            if(contentType.equals("application/vnd.ms-powerpoint")) { // powerpoint 2003
                HSLFSlideShow slideshow = new HSLFSlideShow(new FileInputStream(inputPath.toFile()));

                for (HSLFSlide slide : slideshow.getSlides()) {
                    for (List<HSLFTextParagraph> textParagraphs : slide.getTextParagraphs()) {
                        HSLFTextParagraph.setText(textParagraphs, MongolianLanguageUtil.toUnicode(HSLFTextParagraph.getText(textParagraphs)));
                    }
                }

                FileOutputStream fos = new FileOutputStream(FileSystems.getDefault().getPath(output + ".ppt").toFile());
                try {
                    slideshow.write(fos);
                } catch (Exception e) {
                } finally {
                    try {
                        fos.close();
                    } catch (Exception e) {
                    }
                }
            } else if(contentType.equals("application/vnd.openxmlformats-officedocument.presentationml.presentation")) { // powerpoint 2007+
                XMLSlideShow slideshow = new XMLSlideShow(new FileInputStream(inputPath.toFile()));

                for (XSLFSlide slide : slideshow.getSlides()) {
                    for(XSLFShape shape : slide){
                        if(shape instanceof XSLFTextShape) {
                            XSLFTextShape txShape = (XSLFTextShape)shape;
                            txShape.setText(MongolianLanguageUtil.toUnicode(txShape.getText()));
                        }
                    }
                }

                FileOutputStream fos = new FileOutputStream(FileSystems.getDefault().getPath(output + ".ppt").toFile());
                try {
                    slideshow.write(fos);
                } catch (Exception e) {
                } finally {
                    try {
                        fos.close();
                    } catch (Exception e) {
                    }
                }
            }
        }
    }
}
