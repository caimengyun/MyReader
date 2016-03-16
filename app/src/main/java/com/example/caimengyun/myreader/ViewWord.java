package com.example.caimengyun.myreader;

import android.app.Activity;
import android.graphics.Bitmap;
import android.graphics.BitmapFactory;
import android.os.Bundle;
import android.util.Log;
import android.util.Xml;
import android.webkit.WebSettings;
import android.webkit.WebView;
import android.widget.Toast;

import org.apache.poi.hwpf.HWPFDocument;
import org.apache.poi.hwpf.usermodel.CharacterRun;
import org.apache.poi.hwpf.usermodel.Paragraph;
import org.apache.poi.hwpf.usermodel.Picture;
import org.apache.poi.hwpf.usermodel.Range;
import org.apache.poi.hwpf.usermodel.TableIterator;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.xmlpull.v1.XmlPullParser;
import org.xmlpull.v1.XmlPullParserException;

import java.io.ByteArrayOutputStream;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.List;
import java.util.zip.ZipEntry;
import java.util.zip.ZipException;
import java.util.zip.ZipFile;

/**
 * Created by caimengyun on 16-3-10.
 */
public class ViewWord extends Activity {

    private String LOG_TAG = "ViewWord_";
    private String filenameString;
    private String savePath;
    private String fileName;
    private String name;
    private String htmlPath;
    private String picturePath;
    private FileOutputStream fileOutput;
    private Range range;
    private int numPicture = 0;
    private int screenWidth;
    private List<Picture> pictures;
    private TableIterator tableIterator;
    private WebView webView;

    @Override
    public void onCreate(Bundle savedInstanceState) {
        super.onCreate(savedInstanceState);
        setContentView(R.layout.activity_viewword);
        Bundle bundle = this.getIntent().getExtras();
        filenameString = bundle.getString("filePath");
        screenWidth = this.getWindowManager().getDefaultDisplay().getWidth() - 10;
        //WebView查看HTML，设置WebView的一些缩放功能点
        //Toast.makeText(this, htmlPath, Toast.LENGTH_LONG).show();
        webView = (WebView) this.findViewById(R.id.viewWord);
        WebSettings webSettings = webView.getSettings();
        webSettings.setJavaScriptEnabled(true);
        viewWord();
    }

    /**
     *  识别doc或者docx
     */
    private void viewWord() {
        if (filenameString.endsWith(".html")) {
            Log.i(LOG_TAG ,filenameString);
            webView.loadUrl("file://" + filenameString);
        } else {
            if (filenameString.endsWith(".doc")) {
                makeDirOrFile();
                getFileRange();
                docToHTML();
            }
            if (filenameString.endsWith(".docx")) {
                makeDirOrFile();
                docxToHTML();
            }
            webView.loadUrl("file://" + htmlPath);
        }
    }

    /**
     * 解析doc文件
     */
    private void getFileRange() {
        try {
            FileInputStream fileInputStream = new FileInputStream(filenameString);//获得文件的数据流
            POIFSFileSystem poifsFileSystem = new POIFSFileSystem(fileInputStream);//poi系统流文件对象
            HWPFDocument hwpfDocument = new HWPFDocument(poifsFileSystem);//读取word文档对象
            range = hwpfDocument.getRange();//获取文档读取范围
            pictures = hwpfDocument.getPicturesTable().getAllPictures();
            tableIterator = new TableIterator(range);//获取所有表格的迭代器
        } catch (Exception e) {
            System.out.println("ViewWord getFileRange Exception");
        }
    }

    /**
     * 创建文件数据存放的文件夹
     */
    private void makeDirOrFile() {
        try {
            //Log.i(LOG_TAG + "makeDirOrFile","here");
            File tempFile = new File(filenameString.trim());//转换为字符串，并由绝对路径还原为文件
            fileName = tempFile.getName().toLowerCase();//获取文件名
            name = fileName.substring(0, fileName.indexOf("."));//获取文件名前缀
            fileName = fileName + name;//设置新建文件的命名，避免文件冲突
            savePath = tempFile.getParentFile().getAbsolutePath();//获取文件路径
            if (!(new File(savePath + File.separator + fileName).exists()))
                new File(savePath + File.separator + fileName).mkdirs();
            htmlPath = savePath + File.separator + fileName + File.separator + name + ".html";//html文件的路径
        } catch (Exception e) {
            System.out.println("ViewWord makeDirOrFile Exception");
        }
    }

    /**
     * 创建图片文件
     */
    private void makePicFile() {
        try {
            picturePath = savePath + File.separator + fileName + File.separator + name + numPicture + ".png";
        } catch (Exception e) {
            System.out.println("ViewWord makePicFile Exception");
        }
    }

    /**
     * 读取doc
     */
    private void docToHTML() {
        try {
            File file = new File(htmlPath);
            fileOutput = new FileOutputStream(file);
            String head = "<html><body>";
            String end = "</body></html>";
            String tagBegin = "<p>";
            String tagEnd = "</p>";

            fileOutput.write(head.getBytes());//输出文件头
            int numParagraphs = range.numParagraphs();//获取段落数，包括表格内的段落
            for (int num_Par = 0; num_Par < numParagraphs; num_Par++) {//遍历每一段段落
                Paragraph par = range.getParagraph(num_Par);

                if (par.isInTable()) {//判断段落是否在表格中
                    num_Par = viewTable(num_Par);
                } else {
                    fileOutput.write(tagBegin.getBytes());
                    writeParagraphContent(par);
                    fileOutput.write(tagEnd.getBytes());
                }
            }
            fileOutput.write(end.getBytes());//输出文件尾
            fileOutput.close();
        } catch (Exception e) {
            System.out.println("ViewWord docToHTML Exception");
        }
    }

    /**
     * 读取doc表格
     */
    private int viewTable(int num_Par) {
        int temp = num_Par;//段落数
        try {
            if (tableIterator.hasNext()) {
                String tagBegin = "<p>";
                String tagEnd = "</p>";
                String tableBegin = "<table style=\"border-collapse:collapse\" border=1 bordercolor=\"black\">";
                String tableEnd = "</table>";
                String rowBegin = "<tr>";
                String rowEnd = "</tr>";
                String colBegin = "<td>";
                String colEnd = "</td>";

                org.apache.poi.hwpf.usermodel.Table table = tableIterator.next();
                fileOutput.write(tableBegin.getBytes());//输出表格头

                int rows = table.numRows();//获取行数

                for (int num_rows = 0; num_rows < rows; num_rows++) {//遍历表格行
                    fileOutput.write(rowBegin.getBytes());//输出行头
                    org.apache.poi.hwpf.usermodel.TableRow row = table.getRow(num_rows);
                    int cols = row.numCells();//获取行中的单元格
                    int rowNumParagraphs = row.numParagraphs();
                    int colsNumParagraphs = 0;

                    for (int num_cols = 0; num_cols < cols; num_cols++) {//遍历行中的单元格
                        fileOutput.write(colBegin.getBytes());//输出列头
                        org.apache.poi.hwpf.usermodel.TableCell cell = row.getCell(num_cols);
                        int max = temp + cell.numParagraphs();//行数+单元格的段落数
                        colsNumParagraphs = colsNumParagraphs + cell.numParagraphs();//此行中单元格段落总数

                        for (int num = temp; num < max; num++) {//遍历此行此单元格中所有段落
                            Paragraph par1 = range.getParagraph(num);
                            fileOutput.write(tagBegin.getBytes());
                            writeParagraphContent(par1);
                            fileOutput.write(tagEnd.getBytes());
                            temp++;
                        }
                        fileOutput.write(colEnd.getBytes());//输出列尾
                    }
                    int max1 = temp + rowNumParagraphs;
                    for (int num_extra = temp + colsNumParagraphs; num_extra < max1; num_extra++) {
                        Paragraph par2 = range.getParagraph(num_extra);
                        fileOutput.write(tagBegin.getBytes());
                        writeParagraphContent(par2);
                        fileOutput.write(tagEnd.getBytes());
                        temp++;
                    }//不知道干吗
                    fileOutput.write(rowEnd.getBytes());//输出行尾
                }
                fileOutput.write(tableEnd.getBytes());//输出表格尾
            }
        } catch (Exception e) {
            System.out.println("ViewWord viewTable Exception");
        }
        return temp;
    }

    /**
     * 读取doc段落
     */
    private void writeParagraphContent(Paragraph par) {
        int pnumCharacterRuns = par.numCharacterRuns();

        for (int num_pnum = 0; num_pnum < pnumCharacterRuns; num_pnum++) {
            CharacterRun run = par.getCharacterRun(num_pnum);

            if (run.getPicOffset() == 0 || run.getPicOffset() >= 1000) {
                if (numPicture < pictures.size()) {
                    viewPicture();
                }
            } else {
                try {
                    String text = run.text();
                    if (text.length() >= 2 && pnumCharacterRuns < 2) {
                        fileOutput.write(text.getBytes());
                    } else {
                        int size = run.getFontSize();
                        int color = run.getColor();
                        String fontSizeBegin = "<font size=\"" + decideSize(size) + "\">";
                        String fontColorBegin = "<font color=\"" + decideColor(color) + "\">";
                        String fontEnd = "</font>";
                        String boldBegin = "<b>";
                        String boldEnd = "</b>";
                        String islaBegin = "<i>";
                        String islaEnd = "</i>";

                        fileOutput.write(fontSizeBegin.getBytes());
                        fileOutput.write(fontColorBegin.getBytes());

                        if (run.isBold()) {//黑体字
                            fileOutput.write(boldBegin.getBytes());
                        }
                        if (run.isItalic()) {//斜体
                            fileOutput.write(islaBegin.getBytes());
                        }
                        fileOutput.write(text.getBytes());

                        if (run.isBold()) {
                            fileOutput.write(boldEnd.getBytes());
                        }
                        if (run.isItalic()) {
                            fileOutput.write(islaEnd.getBytes());
                        }
                        fileOutput.write(fontEnd.getBytes());
                        fileOutput.write(fontEnd.getBytes());
                    }
                } catch (Exception e) {
                    System.out.println("ViewWord writeParagraphContent Exception");
                }
            }
        }
    }

    /**
     * 读取doc图片
     */
    private void viewPicture() {
        Picture picture = (Picture) pictures.get(numPicture);
        byte[] pictureBytes = picture.getContent();
        Bitmap bitmap = BitmapFactory.decodeByteArray(pictureBytes, 0, pictureBytes.length);
        makePicFile();
        File file = new File(picturePath);
        numPicture++;
        try {
            FileOutputStream outputPicture = new FileOutputStream(file);
            outputPicture.write(pictureBytes);
            outputPicture.close();
        } catch (Exception e) {
            System.out.println("outputPicture Exception");
        }
        String imageString = "<img src=\"" + file.getAbsolutePath() + "\"";
        if (bitmap.getWidth() > screenWidth) {
            imageString = imageString + " " + "width=\"" + screenWidth + "\"";
        }
        imageString = imageString + ">";
        try {
            fileOutput.write(imageString.getBytes());
        } catch (Exception e) {
            System.out.println("ViewWord viewPicture Exception");
        }
    }

    /**
     * @param size
     * @return
     */
    public int decideSize(int size) {

        if (size >= 1 && size <= 8) {
            return 1;
        }
        if (size >= 9 && size <= 11) {
            return 2;
        }
        if (size >= 12 && size <= 14) {
            return 3;
        }
        if (size >= 15 && size <= 19) {
            return 4;
        }
        if (size >= 20 && size <= 29) {
            return 5;
        }
        if (size >= 30 && size <= 39) {
            return 6;
        }
        if (size >= 40) {
            return 7;
        }
        return 3;
    }

    /**
     * @param a
     * @return
     */
    private String decideColor(int a) {
        int color = a;
        switch (color) {
            case 1:
                return "#000000";
            case 2:
                return "#0000FF";
            case 3:
            case 4:
                return "#00FF00";
            case 5:
            case 6:
                return "#FF0000";
            case 7:
                return "#FFFF00";
            case 8:
                return "#FFFFFF";
            case 9:
                return "#CCCCCC";
            case 10:
            case 11:
                return "#00FF00";
            case 12:
                return "#080808";
            case 13:
            case 14:
                return "#FFFF00";
            case 15:
                return "#CCCCCC";
            case 16:
                return "#080808";
            default:
                return "#000000";
        }
    }

    /**
     * 读取docx
     */
    private void docxToHTML() {
        String river = "";

        try {
            File file = new File(htmlPath);
            fileOutput = new FileOutputStream(file);
            String head = "<!DOCTYPE><html><meta charset=\"utf-8\"><body>";
            String end = "</body></html>";
            String tagBegin = "<p>";
            String tagEnd = "</p>";
            String tableBegin = "<table style=\"border-collapse:collapse\" border=1 bordercolor=\"black\">";
            String tableEnd = "</table>";
            String rowBegin = "<tr>";
            String rowEnd = "</tr>";
            String colBegin = "<td>";
            String colEnd = "</td>";

            fileOutput.write(head.getBytes());//输出文件头
            ZipFile xlsxFile = new ZipFile(new File(filenameString));
            ZipEntry sharedStringXML = xlsxFile.getEntry("word/document.xml");
            InputStream inputStream = xlsxFile.getInputStream(sharedStringXML);
            XmlPullParser xmlParser = Xml.newPullParser();
            xmlParser.setInput(inputStream, "utf-8");//设置数据源编码
            int evtType = xmlParser.getEventType();//获取事件类型

            boolean isTable = false;
            boolean isSize = false;
            boolean isColor = false;
            boolean isCenter = false;
            boolean isRight = false;
            boolean isItalic = false;
            boolean isUnderline = false;
            boolean isBold = false;
            boolean isR = false;
            int pictureIndex = 1;
            while (evtType != XmlPullParser.END_DOCUMENT) {
                switch (evtType) {
                    case XmlPullParser.START_TAG://开始标签
                        String tag = xmlParser.getName();
                        if (tag.equalsIgnoreCase("r")) {
                            isR = true;
                        }
                        if (tag.equalsIgnoreCase("u")) {//判断下划线
                            isUnderline = true;
                        }
                        if (tag.equalsIgnoreCase("jc")) {//判断对齐方式
                            String align = xmlParser.getAttributeValue(0);
                            if (align.equals("center")) {
                                fileOutput.write("<center>".getBytes());
                                isCenter = true;
                            }
                            if (align.equals("right")) {
                                fileOutput.write("<div align=\"right\">".getBytes());
                                isRight = true;
                            }
                        }
                        if (tag.equalsIgnoreCase("color")) {//判断颜色
                            String color = xmlParser.getAttributeValue(0);
                            fileOutput.write(("<font color=" + color + ">").getBytes());
                            isColor = true;
                        }
                        if (tag.equalsIgnoreCase("sz")) {//判断大小
                            if (isR == true) {
                                int size = decideSize(Integer.valueOf(xmlParser.getAttributeValue(0)));
                                fileOutput.write(("<font size=" + size + ">").getBytes());
                                isSize = true;
                            }
                        }
                        //表格开始
                        if (tag.equalsIgnoreCase("tbl")) {//判断表格
                            fileOutput.write(tableBegin.getBytes());
                            isTable = true;
                        }
                        if (tag.equalsIgnoreCase("tr")) {//判断行
                            fileOutput.write(rowBegin.getBytes());
                        }
                        if (tag.equalsIgnoreCase("tc")) {//判断列
                            fileOutput.write(colBegin.getBytes());
                        }
                        if (tag.equalsIgnoreCase("pic")) {//判断图片
                            Log.i(LOG_TAG + "PIC","here");
                            String entryName_jpeg = "word/media/image" + pictureIndex + ".jpeg";
                            String entryName_png = "word/media/image" + pictureIndex + ".png";
                            String entryName_gif = "word/media/image" + pictureIndex + ".gif";
                            String entryName_wmf = "word/media/image" + pictureIndex + ".wmf";
                            String entryName_jpg = "word/media/image" + pictureIndex + ".jpg";
                            String entryName_bmp = "word/media/image" + pictureIndex + ".bmp";
                            String entryName_tif = "word/media/image" + pictureIndex + ".tif";
                            String entryName_emf = "word/media/image" + pictureIndex + ".emf";
                            ZipEntry sharePicture = xlsxFile.getEntry(entryName_jpeg);
                            if (sharePicture == null) {
                                sharePicture = xlsxFile.getEntry(entryName_png);
                                Log.i(LOG_TAG + "PICpng","null");
                            }
                            if (sharePicture == null) {
                                sharePicture = xlsxFile.getEntry(entryName_gif);
                                Log.i(LOG_TAG + "PICgif","null");
                            }
                            if (sharePicture == null) {
                                sharePicture = xlsxFile.getEntry(entryName_wmf);
                                Log.i(LOG_TAG + "PICwnf","null");
                            }
                            if (sharePicture == null) {
                                sharePicture = xlsxFile.getEntry(entryName_jpg);
                                Log.i(LOG_TAG + "PICjpg","null");
                            }
                            if (sharePicture == null) {
                                sharePicture = xlsxFile.getEntry(entryName_bmp);
                                Log.i(LOG_TAG + "PICbmp","null");
                            }
                            if (sharePicture == null) {
                                sharePicture = xlsxFile.getEntry(entryName_tif);
                                Log.i(LOG_TAG + "PICtif","null");
                            }
                            if (sharePicture == null) {
                                sharePicture = xlsxFile.getEntry(entryName_emf);
                                Log.i(LOG_TAG + "PICemf","null");
                            }
                            if (sharePicture == null) {
                                Log.i(LOG_TAG + "ALLPIC","null");
                            }
                            if (sharePicture != null) {
                                Log.i(LOG_TAG + "PIC", "here2");
                                InputStream pictIS = xlsxFile.getInputStream(sharePicture);
                                Log.i(LOG_TAG + "PIC", "here2.1");
                                ByteArrayOutputStream pOut = new ByteArrayOutputStream();
                                Log.i(LOG_TAG + "PIC", "here2.2");
                                byte[] bt = null;
                                Log.i(LOG_TAG + "PIC", "here2.3");
                                byte[] b = new byte[1000];
                                int len = 0;
                                Log.i(LOG_TAG + "PIC", "here3");
                                while ((len = pictIS.read(b)) != -1) {
                                    pOut.write(b, 0, len);
                                }
                                pictIS.close();
                                pOut.close();
                                bt = pOut.toByteArray();
                                if (pictIS != null)
                                    pictIS.close();
                                if (pOut != null)
                                    pOut.close();
                                Log.i(LOG_TAG + "PIC", "here1");
                                writeDOCXPicture(bt);
                                pictureIndex++; // 转换一张后 索引+1
                            }
                        }
                        if (tag.equalsIgnoreCase("b")) {//判断加粗
                            isBold = true;
                        }
                        if (tag.equalsIgnoreCase("p")) {//判断p标签
                            if (isTable == false) {//如果是在表格内就忽略
                                fileOutput.write(tagBegin.getBytes());
                            }
                        }
                        if (tag.equalsIgnoreCase("i")) {//判断斜体
                            isItalic = true;
                        }
                        if (tag.equalsIgnoreCase("t")) {//判断到值标签，将以上判断的标签写入
                            if (isBold == true) {
                                fileOutput.write("<b>".getBytes());
                            }
                            if (isUnderline == true) {
                                fileOutput.write("<u>".getBytes());
                            }
                            if (isItalic == true) {
                                fileOutput.write("<i>".getBytes());
                            }
                            river = xmlParser.nextText();
                            fileOutput.write(river.getBytes());//写入数值
                            if (isItalic == true) {//判断到值标签，将以上的标签结尾写入，并将状态改为false
                                fileOutput.write("</i>".getBytes());
                                isItalic = false;
                            }
                            if (isUnderline == true) {
                                fileOutput.write("</u>".getBytes());
                                isUnderline = false;
                            }
                            if (isBold == true) {
                                fileOutput.write("</b>".getBytes());
                                isBold = false;
                            }
                            if (isSize == true) {
                                fileOutput.write("</font>".getBytes());
                                isSize = false;
                            }
                            if (isColor == true) {
                                fileOutput.write("</font>".getBytes());
                                isColor = false;
                            }
                            if (isCenter == true) {
                                fileOutput.write("</center>".getBytes());
                                isCenter = false;
                            }
                            if (isRight == true) {
                                fileOutput.write("</div>".getBytes());
                                isRight = false;
                            }
                        }
                        break;
                    case XmlPullParser.END_TAG:
                        String tag2 = xmlParser.getName();
                        if (tag2.equalsIgnoreCase("tbl")) {//判断表格结束
                            fileOutput.write(tableEnd.getBytes());
                            isTable = false;
                        }
                        if (tag2.equalsIgnoreCase("tr")) {//判断行结束
                            fileOutput.write(rowEnd.getBytes());
                        }
                        if (tag2.equalsIgnoreCase("tc")) {//判断列结束
                            fileOutput.write(colEnd.getBytes());
                        }
                        if (tag2.equalsIgnoreCase("p")) {
                            if (isTable == false) {
                                fileOutput.write(tagEnd.getBytes());
                            }
                        }
                        if (tag2.equalsIgnoreCase("r")) {
                            isR = false;
                        }
                        break;
                    default:
                        break;
                }
                evtType = xmlParser.next();
            }
            fileOutput.write(end.getBytes());
        } catch (ZipException e) {
            System.out.println("ViewWord docxToHTML Exception Zip");
        } catch (IOException e) {
            System.out.println("ViewWord docxToHTML Exception IOE");
        } catch (XmlPullParserException e) {
            System.out.println("ViewWord docxToHTML Exception XML");
        }
        if (river == null) {
            river = "wrong";
        }
    }

    /**
     * 读取docx图片
     * @param pictureBytes
     */
    public void writeDOCXPicture(byte[] pictureBytes) {
        Log.i(LOG_TAG + "writeDOCX", "here");
        Bitmap bitmap = BitmapFactory.decodeByteArray(pictureBytes, 0, pictureBytes.length);
        makePicFile();
        File myPicture = new File(picturePath);
        numPicture++;
        Log.i(LOG_TAG + "makeDirOrFile", "here" + numPicture);
        try {
            FileOutputStream outputPicture = new FileOutputStream(myPicture);
            outputPicture.write(pictureBytes);
            outputPicture.close();
        } catch (Exception e) {
            System.out.println("outputPicture Exception");
        }
        String imageString = "<img src=\"" + this.picturePath + "\"";
        if (bitmap.getWidth() > this.screenWidth) {
            imageString = imageString + " " + "width=\"" + this.screenWidth + "\"";
        }
        imageString = imageString + ">";
        try {
            fileOutput.write(imageString.getBytes());
        } catch (Exception e) {
            System.out.println("output Exception");
        }
    }
}
