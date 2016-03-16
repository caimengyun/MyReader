package com.example.caimengyun.myreader;


import android.app.Activity;
import android.os.Bundle;
import android.util.Xml;
import android.webkit.WebSettings;
import android.webkit.WebView;
import android.widget.ImageView;
import android.widget.TextView;
import android.widget.Toast;

import org.apache.poi.hslf.HSLFSlideShow;
import org.apache.poi.hslf.model.Slide;
import org.apache.poi.hslf.model.TextRun;
import org.apache.poi.hslf.usermodel.PictureData;
import org.apache.poi.hslf.usermodel.RichTextRun;
import org.apache.poi.hslf.usermodel.SlideShow;
import org.apache.poi.hwpf.usermodel.Picture;
import org.apache.poi.ss.usermodel.Color;
import org.xmlpull.v1.XmlPullParser;
import org.xmlpull.v1.XmlPullParserException;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
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

    //private TextView testView;
    //private ImageView imageView;
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
        //imageView = (ImageView) findViewById(R.id.viewpptxImg);
        //testView = (TextView) findViewById(R.id.viewpptx);
        //testView.setText(readPPTX());
        Bundle bundle = this.getIntent().getExtras();
        filenameString = bundle.getString("filePath");
        view();
        Toast.makeText(this, htmlPath, Toast.LENGTH_LONG).show();
        webView = (WebView) this.findViewById(R.id.viewWord);
        WebSettings webSettings = webView.getSettings();
        webSettings.setJavaScriptEnabled(true);
        webView.loadUrl("file://" + htmlPath);
    }

    private void view() {
        if (filenameString.endsWith(".PPT")) {
            makeDirOrFile();
            readPpt();
        } else {
            //readPptx();
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
            FileInputStream is = new FileInputStream(filenameString);
            SlideShow ppt = new SlideShow(is);
            is.close();
            Dimension pgsize = ppt.getPageSize();
            org.apache.poi.hslf.model.Slide[] slide = ppt.getSlides();
            for (int i = 0; i < slide.length; i++) {
                System.out.print("第" + i + "页。");

                TextRun[] truns = slide[i].getTextRuns();
                for (int k = 0; k < truns.length; k++) {
                    RichTextRun[] rtruns = truns[k].getRichTextRuns();
                    for (int l = 0; l < rtruns.length; l++) {
                        int index = rtruns[l].getFontIndex();
                        String name = rtruns[l].getFontName();
                        rtruns[l].setFontIndex(1);
                        rtruns[l].setFontName("宋体");
                    }
                }
                BufferedImage img = new BufferedImage(pgsize.width,
                        pgsize.height, BufferedImage.TYPE_INT_RGB);

                Graphics2D graphics = img.createGraphics();
                graphics.setPaint(Color.BLUE);
                graphics.fill(new Rectangle2D.Float(0, 0, pgsize.width,
                        pgsize.height));
                slide[i].draw(graphics);

// 这里设置图片的存放路径和图片的格式(jpeg,png,bmp等等),注意生成文件路径
                File path = new File("F:/images");
                if (!path.exists()) {
                    path.mkdir();
                }
                FileOutputStream out = new FileOutputStream(path + "/pict_"
                        + (i + 1) + ".jpeg");
                javax.imageio.ImageIO.write(img, "jpeg", out);
                out.close();

            }
            System.out.println("success!!");
            return true;
        } catch (FileNotFoundException e) {
            System.out.println(e);
// System.out.println("Can't find the image!");
        } catch (IOException e) {
        }


        try {
            HSLFSlideShow hslfSlideShow = new HSLFSlideShow(filenameString);
            SlideShow slideShow = new SlideShow(hslfSlideShow);
            // 获取PPT文件中的图片数据
            PictureData[] pictures = slideShow.getPictureData();
            // 循环读取图片数据
            for (int i = 0; i < pictures.length; i++) {
                PictureData pic_data = pictures[i];
                // 设置格式
                switch (pic_data.getType()) {
                    case Picture.JPEG:
                        lsb.append(".jpg");
                        break;
                    case Picture.PNG:
                        lsb.append(".png");
                        break;
                    default:
                        fileName.append(".data");
                }
                // 输出文件
                fileOutput.write(pic_data.getData());
                fileOutput.close();
            }
        } catch (Exception e) {
            System.out.println(LOG_TAG + "readPpt");
        }
    }


//    /**
//     * @return
//     */
//    public String readPptx() {
//        List<String> list = new ArrayList<String>();
//        String river = "";
//        ZipFile xlsxFile = null;
//        try {
//            xlsxFile = new ZipFile(new File(filenameString));
//        } catch (ZipException e1) {
//            e1.printStackTrace();
//        } catch (IOException e1) {
//            e1.printStackTrace();
//        }
//        try {
//            ZipEntry sharedStringXML = xlsxFile.getEntry("[Content_Types].xml");
//            InputStream inputStream = xlsxFile.getInputStream(sharedStringXML);
//            XmlPullParser xmlParser = Xml.newPullParser();
//            xmlParser.setInput(inputStream, "utf-8");
//            int evtType = xmlParser.getEventType();
//            while (evtType != XmlPullParser.END_DOCUMENT) {
//                switch (evtType) {
//                    case XmlPullParser.START_TAG:
//                        String tag = xmlParser.getName();
//                        if (tag.equalsIgnoreCase("Override")) {
//                            String s = xmlParser
//                                    .getAttributeValue(null, "PartName");
//                            if (s.lastIndexOf("/ppt/slides/slide") == 0) {
//                                list.add(s);
//                            }
//                        }
//                        break;
//                    case XmlPullParser.END_TAG:
//                        break;
//                    default:
//                        break;
//                }
//                evtType = xmlParser.next();
//            }
//        } catch (ZipException e) {
//            e.printStackTrace();
//        } catch (IOException e) {
//            e.printStackTrace();
//        } catch (XmlPullParserException e) {
//            e.printStackTrace();
//        }
//        for (int i = 1; i < (list.size() + 1); i++) {
//            river += "��" + i + "��:" + "\n";
//            try {
//                ZipEntry sharedStringXML = xlsxFile.getEntry("ppt/slides/slide" + i + ".xml");
//                InputStream inputStream = xlsxFile.getInputStream(sharedStringXML);
//                XmlPullParser xmlParser = Xml.newPullParser();
//                xmlParser.setInput(inputStream, "utf-8");
//                int evtType = xmlParser.getEventType();
//                while (evtType != XmlPullParser.END_DOCUMENT) {
//                    switch (evtType) {
//                        case XmlPullParser.START_TAG:
//                            String tag = xmlParser.getName();
//                            if (tag.equalsIgnoreCase("t")) {
//                                river += xmlParser.nextText() + "\n";
//                            } else if (tag.equalsIgnoreCase("cNvPr")) {
//
//                                //  img.setImageResource();
//                            }
//                            break;
//                        case XmlPullParser.END_TAG:
//                            break;
//                        default:
//                            break;
//                    }
//                    evtType = xmlParser.next();
//                }
//            } catch (ZipException e) {
//                e.printStackTrace();
//            } catch (IOException e) {
//                e.printStackTrace();
//            } catch (XmlPullParserException e) {
//                e.printStackTrace();
//            }
//        }
//        if (river == null) {
//            river = "wrong";
//        }
//        return river;
//    }
}
