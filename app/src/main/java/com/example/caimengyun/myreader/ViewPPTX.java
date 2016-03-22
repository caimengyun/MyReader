package com.example.caimengyun.myreader;

import android.app.Activity;
import android.os.Bundle;
import android.util.Xml;
import android.webkit.WebSettings;
import android.webkit.WebView;
import android.widget.Toast;

import org.apache.poi.hslf.HSLFSlideShow;
import org.apache.poi.hslf.model.Slide;
import org.apache.poi.hslf.model.TextRun;
import org.apache.poi.hslf.usermodel.PictureData;
import org.apache.poi.hslf.usermodel.SlideShow;
import org.apache.poi.hwpf.usermodel.Picture;
import org.xmlpull.v1.XmlPullParser;
import org.xmlpull.v1.XmlPullParserException;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.List;
import java.util.zip.ZipEntry;
import java.util.zip.ZipException;
import java.util.zip.ZipFile;


/**
 * Created by caimengyun on 16-3-10.
 */
public class ViewPPTX extends Activity {

    private String LOG_TAG = "ViewPPTX_";
    private String filenameString;
    private String savePath;
    private String fileName;
    private String name;
    private String htmlPath;
    private String picturePath;
    private FileOutputStream fileOutput;
    private int numPicture = 0;
    private List<Picture> pictures;
    private WebView webView;
    private StringBuffer lsb = new StringBuffer();

    @Override
    protected void onCreate(Bundle savedInstanceState) {
        super.onCreate(savedInstanceState);
        setContentView(R.layout.activity_viewword);
        Bundle bundle = this.getIntent().getExtras();
        filenameString = bundle.getString("filePath");

        view();
        Toast.makeText(this, htmlPath, Toast.LENGTH_LONG).show();
        webView = (WebView) this.findViewById(R.id.viewWord);
        WebSettings webSettings = webView.getSettings();
        webSettings.setJavaScriptEnabled(true);
        webSettings.setLayoutAlgorithm(WebSettings.LayoutAlgorithm.SINGLE_COLUMN);
        //webView.setInitialScale(50);
        webSettings.setSupportZoom(true);
        webSettings.setBuiltInZoomControls(true);
        webView.loadUrl("file://" + htmlPath);
    }

    private void view() {
        if (filenameString.endsWith(".ppt")) {
            makeDirOrFile();
            readPpt();
            //read();
        } else {
            makeDirOrFile();
            readPptx();
        }
    }

    /**
     * 创建文件数据存放的文件夹
     */
    private void makeDirOrFile() {
        try {
            File tempFile = new File(filenameString.trim());//转换为字符串，并由绝对路径还原为文件
            fileName = tempFile.getName().toLowerCase();//获取文件名
            name = fileName.substring(0, fileName.indexOf("."));//获取文件名前缀
            fileName = fileName + name;//设置新建文件的命名，避免文件冲突
            savePath = tempFile.getParentFile().getAbsolutePath();//获取文件路径
            if (!(new File(savePath + File.separator + fileName).exists()))
                new File(savePath + File.separator + fileName).mkdirs();
            htmlPath = savePath + File.separator + fileName + File.separator + name + ".html";//html文件的路径
        } catch (Exception e) {
            System.out.println(LOG_TAG + "ViewPPTX makeDirOrFile Exception");
        }
    }

    /**
     * 创建图片文件
     */
    private void makePicFile() {
        try {
            picturePath = savePath + File.separator + fileName + File.separator + name + numPicture + ".png";
        } catch (Exception e) {
            System.out.println(LOG_TAG + "ViewWord makePicFile Exception");
        }
    }

    private void readPpt() {
        try {
            File file = new File(htmlPath);
            fileOutput = new FileOutputStream(file);
            lsb.append("<html xmlns:o=\'urn:schemas-microsoft-com:office:office\' xmlns:x=\'urn:schemas-microsoft-com:office:excel\' xmlns=\'http://www.w3.org/TR/REC-html40\'>");
            lsb.append("<head><meta http-equiv=Content-Type content=\'text/html; charset=utf-8\'><meta name=ProgId content=Excel.Sheet>");

            FileInputStream fileInputStream = new FileInputStream(filenameString);
            SlideShow ppt = new SlideShow(new HSLFSlideShow(fileInputStream));
            PictureData[] pictures = ppt.getPictureData();
            Slide[] slides = ppt.getSlides();//提取文本信息

            for (int k = 0; k < slides.length; k++) {//遍历每一张ppt
                lsb.append("<table>");//
                lsb.append("<tr height=\"25\">");
                lsb.append("<td style=\"font-weight:400;font-size:160%;\"align=\"center\"valign=\"bottom\"colspan=\"10\" rowspan=\"0\">");
                lsb.append("第" + k + "张PPT：");
                lsb.append(slides[k].getTitle());
                lsb.append("</tb></tr></table>");
                TextRun textRun[] = slides[k].getTextRuns();
                lsb.append("<table style=\"border: #333333; border-style: solid; border-top-width: 2px;border-right-width: 2px; border-bottom-width: 2px; border-left-width: 2px\">");
                for (int i = 0; i < textRun.length; i++) {
                    lsb.append("<tr height=\"25\">");
                    lsb.append("<td style=\"font-weight:400;font-size:160%;\"align=\"center\"valign=\"bottom\"colspan=\"10\" rowspan=\"0\">");
                    String text = textRun[i].getText();
                    if (text == null) {
                        lsb.append("<br>" + "null" + "</br>");
                    } else {
                        lsb.append(text);
                    }
                    lsb.append("</tb></tr>");
                    if (numPicture < pictures.length) {
                        lsb.append("<tr height=\"25\">");
                        lsb.append("<td>");
                        makePicFile();
                        File pic_file = new File(picturePath);
                        PictureData pic_data = pictures[numPicture];
                        FileOutputStream outputPicture = new FileOutputStream(pic_file);
                        outputPicture.write(pic_data.getData());
                        outputPicture.close();
                        String imageString = "<img src=\"" + pic_file.getAbsolutePath() + "\"style=\"width:200px;height:200px;>";
                        lsb.append(imageString + "</tb></tr>");
                        numPicture++;
                    }
                }
                lsb.append("</table>");
            }
            fileOutput.write(lsb.toString().getBytes());
        } catch (Exception e) {
            System.out.println(LOG_TAG + "read");
        }
    }

    /**
     * @return
     */
    private void readPptx() {
        List<String> list = new ArrayList<String>();
        ZipFile xlsxFile = null;
        try {
            File file = new File(htmlPath);
            fileOutput = new FileOutputStream(file);
            lsb.append("<html xmlns:o=\'urn:schemas-microsoft-com:office:office\' xmlns:x=\'urn:schemas-microsoft-com:office:excel\' xmlns=\'http://www.w3.org/TR/REC-html40\'>");
            lsb.append("<head><meta http-equiv=Content-Type content=\'text/html; charset=utf-8\'><meta name=ProgId content=Excel.Sheet>");

            xlsxFile = new ZipFile(new File(filenameString));
            ZipEntry sharedStringXML = xlsxFile.getEntry("[Content_Types].xml");
            InputStream inputStream = xlsxFile.getInputStream(sharedStringXML);
            XmlPullParser xmlParser = Xml.newPullParser();
            xmlParser.setInput(inputStream, "utf-8");
            int evtType = xmlParser.getEventType();
            while (evtType != XmlPullParser.END_DOCUMENT) {
                switch (evtType) {
                    case XmlPullParser.START_TAG:
                        String tag = xmlParser.getName();
                        if (tag.equalsIgnoreCase("Override")) {
                            String s = xmlParser.getAttributeValue(null, "PartName");
                            if (s.lastIndexOf("/ppt/slides/slide") == 0) {
                                list.add(s);
                            }
                        }
                        break;
                    case XmlPullParser.END_TAG:
                        break;
                    default:
                        break;
                }
                evtType = xmlParser.next();
            }
        } catch (ZipException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        } catch (XmlPullParserException e) {
            e.printStackTrace();
        }
        try {
            for (int i = 1; i < (list.size() + 1); i++) {
                lsb.append("<table>");//
                lsb.append("<tr height=\"25\">");
                lsb.append("<td style=\"font-weight:400;font-size:160%;\"align=\"center\"valign=\"bottom\"colspan=\"10\" rowspan=\"0\">");
                lsb.append("第" + i + "张PPT：");
                lsb.append("</tb></tr></table>");
                lsb.append("<table style=\"border: #333333; border-style: solid; border-top-width: 2px;border-right-width: 2px; border-bottom-width: 2px; border-left-width: 2px\">");
                ZipEntry sharedStringXML = xlsxFile.getEntry("ppt/slides/slide" + i + ".xml");
                InputStream inputStream = xlsxFile.getInputStream(sharedStringXML);
                XmlPullParser xmlParser = Xml.newPullParser();
                xmlParser.setInput(inputStream, "utf-8");
                int evtType = xmlParser.getEventType();
                while (evtType != XmlPullParser.END_DOCUMENT) {
                    switch (evtType) {
                        case XmlPullParser.START_TAG:
                            String tag = xmlParser.getName();
                            lsb.append("<tr height=\"25\">");
                            lsb.append("<td style=\"font-weight:400;font-size:160%;\"align=\"center\"valign=\"bottom\"colspan=\"10\" rowspan=\"0\">");
                            if (tag.equalsIgnoreCase("t")) {
                                lsb.append(xmlParser.nextText());
                                lsb.append("</tb></tr>");
                            } else if (tag.equalsIgnoreCase("cNvPr")) {
                                //xmlParser.
                            }
                            break;
                        case XmlPullParser.END_TAG:
                            break;
                        default:
                            break;
                    }
                    evtType = xmlParser.next();
                }
                lsb.append("</table>");
            }
            fileOutput.write(lsb.toString().getBytes());
        } catch (ZipException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        } catch (XmlPullParserException e) {
            e.printStackTrace();
        }
    }

}
