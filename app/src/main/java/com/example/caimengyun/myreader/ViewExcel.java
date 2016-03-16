package com.example.caimengyun.myreader;

import android.app.Activity;
import android.os.Bundle;
import android.util.Log;
import android.util.Xml;
import android.webkit.WebSettings;
import android.webkit.WebView;
import android.widget.Toast;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFClientAnchor;
import org.apache.poi.hssf.usermodel.HSSFDateUtil;
import org.apache.poi.hssf.usermodel.HSSFPalette;
import org.apache.poi.hssf.usermodel.HSSFPicture;
import org.apache.poi.hssf.usermodel.HSSFPictureData;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFShape;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.hwpf.usermodel.Picture;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.xmlpull.v1.XmlPullParser;
import org.xmlpull.v1.XmlPullParserException;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.math.BigDecimal;
import java.text.DecimalFormat;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;
import java.util.zip.ZipEntry;
import java.util.zip.ZipException;
import java.util.zip.ZipFile;

/**
 * Created by caimengyun on 16-3-10.
 */
public class ViewExcel extends Activity {

    private String LOG_TAG = "ViewExcel_";
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
    public void onCreate(Bundle savedInstanceState) {
        super.onCreate(savedInstanceState);
        setContentView(R.layout.activity_viewword);
        Bundle bundle = this.getIntent().getExtras();
        filenameString = bundle.getString("filePath");
        viewExcel();
        //WebView查看HTML，设置WebView的一些缩放功能点
        Toast.makeText(this, htmlPath, Toast.LENGTH_LONG).show();
        webView = (WebView) this.findViewById(R.id.viewWord);
        WebSettings webSettings = webView.getSettings();
        webSettings.setJavaScriptEnabled(true);
        webView.loadUrl("file://" + htmlPath);
    }

    /**
     *
     */
    private void viewExcel() {
        if (filenameString.endsWith(".xls")) {
            try {
                makeDirOrFile();
                readXLS();
            } catch (Exception e) {
                System.out.println(LOG_TAG + "viewExcel");
            }
        }
        if (filenameString.endsWith(".xlsx")) {
            try {
                makeDirOrFile();
                //readXLSX();
                //readXlsx();
            } catch (Exception e) {
                System.out.println(LOG_TAG + "viewExcel");
            }

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
            System.out.println(LOG_TAG + "ViewExcel makeDirOrFile Exception");
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

    private StringBuffer readXLS() throws Exception {

        //List<PicInfo> picInfos = new ArrayList<PicInfo>();
        File file = new File(htmlPath);
        fileOutput = new FileOutputStream(file);

        lsb.append("<html xmlns:o=\'urn:schemas-microsoft-com:office:office\' xmlns:x=\'urn:schemas-microsoft-com:office:excel\' xmlns=\'http://www.w3.org/TR/REC-html40\'>");
        lsb.append("<head><meta http-equiv=Content-Type content=\'text/html; charset=utf-8\'><meta name=ProgId content=Excel.Sheet>");
        HSSFSheet sheet = null;
        String excelFileName = filenameString;
        try {
            FileInputStream fis = new FileInputStream(excelFileName);
            HSSFWorkbook workbook = (HSSFWorkbook) WorkbookFactory.create(fis); // 获整个Excel

            for (int sheetIndex = 0; sheetIndex < workbook.getNumberOfSheets(); sheetIndex++) {
                if (workbook.getSheetAt(sheetIndex) != null) {
                    sheet = workbook.getSheetAt(sheetIndex);// 获得不为空的这个sheet

//                    if (picInfos.size() > 0 && picInfos != null) {
//                        picInfos.clear();
//                    }
//                    if (sheet.getDrawingPatriarch() != null) {
//                        List<HSSFShape> shapes = sheet.getDrawingPatriarch().getChildren();
//                        for (HSSFShape shape : shapes) {
//                            HSSFClientAnchor anchor = (HSSFClientAnchor) shape.getAnchor();
//                            if (shape instanceof HSSFPicture) {
//                                HSSFPicture pic = (HSSFPicture) shape;
//                                pic.getPictureData();
//                                PicInfo info = new PicInfo();
//
//                                int row = anchor.getRow1(); // 得到锚点的起始行号
//
//                                int col = anchor.getCol1();// 得到锚点的起始列号
//
//                                info.setColNum(col);
//                                info.setRowNum(row);
//                                info.setSheetName(sheet.getSheetName());
//                                HSSFPictureData picData = pic.getPictureData();
//                                try {
//                                    info = savePic(info, sheet.getSheetName(), picData);
//                                } catch (Exception e) {
//                                    e.printStackTrace();
//                                }
//                                picInfos.add(info);
//                            }
//                        }
//
//                    }
                    int firstRowNum = sheet.getFirstRowNum(); // 第一行
                    int lastRowNum = sheet.getLastRowNum(); // 最后一行
                    // 构造Table
                    lsb.append("<table width=\"100%\" style=\"border:1px solid #000;border-width:1px 0 0 1px;margin:2px 0 2px 0;border-collapse:collapse;\">");
                    int[] widthList = new int[100];//用于记录本行各列表格的宽度，用于空行或者空格

                    for (int rowNum = firstRowNum; rowNum <= lastRowNum; rowNum++) {
                        if (sheet.getRow(rowNum) != null) {// 如果行不为空，
                            HSSFRow row = sheet.getRow(rowNum);
                            //short firstCellNum = row.getFirstCellNum(); // 该行的第一个单元格
                            short firstCellNum = 0;
                            short lastCellNum = row.getLastCellNum(); // 该行的最后一个单元格
                            int height = (int) (row.getHeight() / 15.625); // 行的高度
                            lsb.append("<tr height=\"" + height + "\" style=\"border:1px solid #000;border-width:0 1px 1px 0;margin:2px 0 2px 0;\">");
                            int width = 50;

                            for (short cellNum = firstCellNum; cellNum <= lastCellNum; cellNum++) { // 循环该行的每一个单元格
                                HSSFCell cell = row.getCell(cellNum);
                                StringBuffer tdStyle = new StringBuffer("<td style=\"border:1px solid #000; border-width:1px 1px 1px 1px;margin:2px 0 2px 0; ");
                                if (null != cell) {
                                    HSSFCellStyle cellStyle = cell.getCellStyle();
                                    short boldWeight = cellStyle.getFont(workbook).getBoldweight(); // 字体粗细
                                    short fontHeight = (short) (cellStyle.getFont(workbook).getFontHeight() / 2); // 字体大小
                                    HSSFPalette palette = workbook.getCustomPalette();
                                    HSSFColor hColor = palette.getColor(cellStyle.getFillForegroundColor());
                                    HSSFColor hColor2 = palette.getColor(cellStyle.getFont(workbook).getColor());
                                    String bgColor = convertToStardColorOfXls(hColor);
                                    String fontColor = convertToStardColorOfXls(hColor2);

                                    if (bgColor != null && !"".equals(bgColor.trim())) {
                                        tdStyle.append(" background-color:" + bgColor + "; ");
                                    }
                                    if (fontColor != null && !"".equals(fontColor.trim())) {
                                        tdStyle.append(" color:" + fontColor + "; ");
                                    }

                                    tdStyle.append(" font-weight:" + boldWeight + "; ");
                                    tdStyle.append(" font-size: " + fontHeight + "%;");
                                    lsb.append(tdStyle + "\"");

                                    width = (int) (sheet.getColumnWidth(cellNum) / 35.7);
                                    widthList[cellNum] = width;
                                    int cellReginCol = getMergerCellRegionColOfXls(sheet, rowNum, cellNum);
                                    int cellReginRow = getMergerCellRegionRowOfXls(sheet, rowNum, cellNum);
                                    String align = convertAlignToHtml(cellStyle.getAlignment());
                                    String vAlign = convertVerticalAlignToHtml(cellStyle.getVerticalAlignment());
                                    lsb.append(" align=\"" + align + "\" valign=\"" + vAlign + "\" width=\"" + width + "\" ");
                                    lsb.append(" colspan=\"" + cellReginCol + "\" rowspan=\"" + cellReginRow + "\"");
                                    lsb.append(">" + getCellValueOfXls(cell) + "</td>");
                                } else {
                                    lsb.append(tdStyle + "\"");
                                    if (widthList[cellNum] == 0) {
                                        lsb.append(" width=\"" + width + "\" ");//第一行
                                        widthList[cellNum] = width;
                                    } else {
                                        lsb.append(" width=\"" + widthList[cellNum] + "\" ");//非第一行
                                    }
                                    lsb.append(">" + "</td>");
                                }

//                              boolean flag = false;
//                                if (picInfos.size() > 0 && picInfos != null) {
//                                    for (PicInfo picInfo : picInfos) {
//                                        // 找到图片对应的单元格
//                                        if (picInfo.getSheetName().equals(sheet.getSheetName()) && picInfo.getRowNum() == rowNum && picInfo.getColNum() == cellNum) {
//                                            flag = true;
//                                            String imagePath = "<img src=\"" + picInfo.getPicPath() + "\"" + "/>";
//                                            lsb.append("<td>");
//                                            lsb.append(">" + imagePath + "</td>");
//                                        }
//                                    }
//                                }
                            }
                            lsb.append("</tr>");
                        }
                    }

                }
            }
            fileOutput.write(lsb.toString().getBytes());
            fis.close();
        } catch (FileNotFoundException e) {
            throw new Exception("文件 " + excelFileName + " 没有找到!");
        } catch (IOException e) {
            throw new Exception("文件 " + excelFileName + " 处理错误(" + e.getMessage() + ")!");
        }
        return null;
    }

    public void readXLSX() {
        try {
            File file = new File(this.htmlPath);
            fileOutput = new FileOutputStream(file);
            String head = "<!DOCTYPE HTML PUBLIC \"-//W3C//DTD HTML 4.01 Transitional//EN\"\"http://www.w3.org/TR/html4/loose.dtd\"><html><meta charset=\"utf-8\"><head></head><body>";//ͷļ,utf-8,Ȼ
            String tableBegin = "<table style=\"border-collapse:collapse\" border=1 bordercolor=\"black\">";
            String tableEnd = "</table>";
            String rowBegin = "<tr>";
            String rowEnd = "</tr>";
            String colBegin = "<td>";
            String colEnd = "</td>";
            String end = "</body></html>";

            fileOutput.write(head.getBytes());
            fileOutput.write(tableBegin.getBytes());
            String str = "";
            String v = null;
            boolean flag = false;
            List<String> list = new ArrayList<String>();
            try {
                ZipFile xlsxFile = new ZipFile(new File(filenameString));
                ZipEntry sharedStringXML = xlsxFile.getEntry("xl/sharedStrings.xml");
                InputStream inputStream = xlsxFile.getInputStream(sharedStringXML);
                XmlPullParser xmlParser = Xml.newPullParser();
                xmlParser.setInput(inputStream, "utf-8");
                int evtType = xmlParser.getEventType();
                while (evtType != XmlPullParser.END_DOCUMENT) {
                    switch (evtType) {
                        case XmlPullParser.START_TAG:
                            String tag = xmlParser.getName();
                            if (tag.equalsIgnoreCase("t")) {
                                list.add(xmlParser.nextText());
                            }
                            break;
                        case XmlPullParser.END_TAG:
                            break;
                        default:
                            break;
                    }
                    evtType = xmlParser.next();
                }
                ZipEntry sheetXML = xlsxFile.getEntry("xl/worksheets/sheet1.xml");
                InputStream inputStreamSheet = xlsxFile.getInputStream(sheetXML);
                XmlPullParser xmlParserSheet = Xml.newPullParser();
                xmlParserSheet.setInput(inputStreamSheet, "utf-8");
                int evtTypeSheet = xmlParserSheet.getEventType();
                fileOutput.write(rowBegin.getBytes());
                int i = -1;
                while (evtTypeSheet != XmlPullParser.END_DOCUMENT) {
                    switch (evtTypeSheet) {
                        case XmlPullParser.START_TAG:
                            String tag = xmlParserSheet.getName();
                            if (tag.equalsIgnoreCase("row")) {
                            } else {
                                if (tag.equalsIgnoreCase("c")) {
                                    String t = xmlParserSheet.getAttributeValue(null, "t");
                                    if (t != null) {
                                        flag = true;
                                    } else {
                                        fileOutput.write(colBegin.getBytes());
                                        fileOutput.write(colEnd.getBytes());
                                        flag = false;
                                    }
                                } else {
                                    if (tag.equalsIgnoreCase("v")) {
                                        v = xmlParserSheet.nextText();
                                        fileOutput.write(colBegin.getBytes());
                                        if (v != null) {
                                            if (flag) {
                                                str = list.get(Integer.parseInt(v));
                                            } else {
                                                str = v;
                                            }
                                            fileOutput.write(str.getBytes());
                                            fileOutput.write(colEnd.getBytes());
                                        }
                                    }
                                }
                            }
                            break;
                        case XmlPullParser.END_TAG:
                            if (xmlParserSheet.getName().equalsIgnoreCase("row") && v != null) {
                                if (i == 1) {
                                    fileOutput.write(rowEnd.getBytes());
                                    fileOutput.write(rowBegin.getBytes());
                                    i = 1;
                                } else {
                                    fileOutput.write(rowBegin.getBytes());
                                }
                            }
                            break;
                    }
                    evtTypeSheet = xmlParserSheet.next();
                }
                System.out.println(str);
            } catch (ZipException e) {
                e.printStackTrace();
            } catch (IOException e) {
                System.out.println(LOG_TAG + "ViewWord docToHTML Exception IO");
            } catch (XmlPullParserException e) {
                System.out.println(LOG_TAG + "ViewWord docToHTML Exception Xml");
            }
            if (str == null) {
                str = "wrong";
            }
            fileOutput.write(rowEnd.getBytes());
            fileOutput.write(tableEnd.getBytes());
            fileOutput.write(end.getBytes());
        } catch (Exception e) {
            System.out.println("readAndWrite Exception");
        }
    }

//    public void readXlsx() throws Exception {
//        File file = new File(htmlPath);
//        fileOutput = new FileOutputStream(file);
//
//        lsb.append("<html xmlns:o=\'urn:schemas-microsoft-com:office:office\' xmlns:x=\'urn:schemas-microsoft-com:office:excel\' xmlns=\'http://www.w3.org/TR/REC-html40\'>");
//        lsb.append("<head><meta http-equiv=Content-Type content=\'text/html; charset=utf-8\'><meta name=ProgId content=Excel.Sheet>");
//        XSSFSheet sheet = null;
//        String excelFileName = filenameString;
//        try {
//            FileInputStream fis = new FileInputStream(excelFileName);
//            XSSFWorkbook workbook = (XSSFWorkbook) WorkbookFactory.create(fis); // 获整个Excel
//
//            for (int sheetIndex = 0; sheetIndex < workbook.getNumberOfSheets(); sheetIndex++) {
//                if (workbook.getSheetAt(sheetIndex) != null) {
//                    sheet = workbook.getSheetAt(sheetIndex);// 获得不为空的这个sheet
//                    int firstRowNum = sheet.getFirstRowNum(); // 第一行
//                    int lastRowNum = sheet.getLastRowNum(); // 最后一行
//                    // 构造Table
//                    lsb.append("<table width=\"100%\" style=\"border:1px solid #000;border-width:1px 0 0 1px;margin:2px 0 2px 0;border-collapse:collapse;\">");
//                    int[] widthList = new int[100];//用于记录本行各列表格的宽度，用于空行或者空格
//
//                    for (int rowNum = firstRowNum; rowNum <= lastRowNum; rowNum++) {
//                        if (sheet.getRow(rowNum) != null) {// 如果行不为空，
//                            XSSFRow row = sheet.getRow(rowNum);
//                            //short firstCellNum = row.getFirstCellNum(); // 该行的第一个单元格
//                            short firstCellNum = 0;
//                            short lastCellNum = row.getLastCellNum(); // 该行的最后一个单元格
//                            int height = (int) (row.getHeight() / 15.625); // 行的高度
//                            lsb.append("<tr height=\"" + height + "\" style=\"border:1px solid #000;border-width:0 1px 1px 0;margin:2px 0 2px 0;\">");
//                            int width = 50;
//
//                            for (short cellNum = firstCellNum; cellNum <= lastCellNum; cellNum++) { // 循环该行的每一个单元格
//                                XSSFCell cell = row.getCell(cellNum);
//                                StringBuffer tdStyle = new StringBuffer("<td style=\"border:1px solid #000; border-width:1px 1px 1px 1px;margin:2px 0 2px 0; ");
//                                if (null != cell) {
//                                    XSSFCellStyle cellStyle = cell.getCellStyle();
//                                    short boldWeight = cellStyle.getFont(workbook).getBoldweight(); // 字体粗细
//                                    short fontHeight = (short) (cellStyle.getFont(workbook).getFontHeight() / 2); // 字体大小
//                                    XSSFPalette palette = workbook.getCustomPalette();
//                                    XSSFColor hColor = palette.getColor(cellStyle.getFillForegroundColor());
//                                    XSSFColor hColor2 = palette.getColor(cellStyle.getFont(workbook).getColor());
//                                    String bgColor = convertToStardColorOfXlsx(hColor);
//                                    String fontColor = convertToStardColorOfXlsx(hColor2);
//
//                                    if (bgColor != null && !"".equals(bgColor.trim())) {
//                                        tdStyle.append(" background-color:" + bgColor + "; ");
//                                    }
//                                    if (fontColor != null && !"".equals(fontColor.trim())) {
//                                        tdStyle.append(" color:" + fontColor + "; ");
//                                    }
//
//                                    tdStyle.append(" font-weight:" + boldWeight + "; ");
//                                    tdStyle.append(" font-size: " + fontHeight + "%;");
//                                    lsb.append(tdStyle + "\"");
//
//                                    width = (int) (sheet.getColumnWidth(cellNum) / 35.7);
//                                    widthList[cellNum] = width;
//                                    int cellReginCol = getMergerCellRegionColOfXlsx(sheet, rowNum, cellNum);
//                                    int cellReginRow = getMergerCellRegionRowOfXlsx(sheet, rowNum, cellNum);
//                                    String align = convertAlignToHtml(cellStyle.getAlignment());
//                                    String vAlign = convertVerticalAlignToHtml(cellStyle.getVerticalAlignment());
//                                    lsb.append(" align=\"" + align + "\" valign=\"" + vAlign + "\" width=\"" + width + "\" ");
//                                    lsb.append(" colspan=\"" + cellReginCol + "\" rowspan=\"" + cellReginRow + "\"");
//                                    lsb.append(">" + getCellValueOfXlsx(cell) + "</td>");
//                                } else {
//                                    lsb.append(tdStyle + "\"");
//                                    if (widthList[cellNum] == 0) {
//                                        lsb.append(" width=\"" + width + "\" ");//第一行
//                                        widthList[cellNum] = width;
//                                    } else {
//                                        lsb.append(" width=\"" + widthList[cellNum] + "\" ");//非第一行
//                                    }
//                                    lsb.append(">" + "</td>");
//                                }
//                            }
//                            lsb.append("</tr>");
//                        }
//                    }
//
//                }
//            }
//            fileOutput.write(lsb.toString().getBytes());
//            fis.close();
//        } catch (FileNotFoundException e) {
//            throw new Exception("文件 " + excelFileName + " 没有找到!");
//        } catch (IOException e) {
//            throw new Exception("文件 " + excelFileName + " 处理错误(" + e.getMessage() + ")!");
//        }
//    }

    /**
     * @param cell
     * @return
     * @throws IOException
     */
    private static Object getCellValueOfXls(HSSFCell cell) throws IOException {
        Object value = "";
        if (cell.getCellType() == HSSFCell.CELL_TYPE_STRING) {
            value = cell.getRichStringCellValue().toString();
        } else if (cell.getCellType() == HSSFCell.CELL_TYPE_NUMERIC) {
            if (HSSFDateUtil.isCellDateFormatted(cell)) {
                Date date = (Date) cell.getDateCellValue();
                SimpleDateFormat sdf = new SimpleDateFormat("yyyy-MM-dd");
                value = sdf.format(date);
            } else {
                double value_temp = (double) cell.getNumericCellValue();
                BigDecimal bd = new BigDecimal(value_temp);
                BigDecimal bd1 = bd.setScale(3, bd.ROUND_HALF_UP);
                value = bd1.doubleValue();

                DecimalFormat format = new DecimalFormat("#0.###");
                value = format.format(cell.getNumericCellValue());

            }
        }
        if (cell.getCellType() == HSSFCell.CELL_TYPE_BLANK) {
            value = "";
        }
        return value;
    }

//    private static Object getCellValueOfXlsx(XSSFCell cell) throws IOException {
//        Object value = "";
//        if (cell.getCellType() == HSSFCell.CELL_TYPE_STRING) {
//            value = cell.getRichStringCellValue().toString();
//        } else if (cell.getCellType() == HSSFCell.CELL_TYPE_NUMERIC) {
//            if (HSSFDateUtil.isCellDateFormatted(cell)) {
//                Date date = (Date) cell.getDateCellValue();
//                SimpleDateFormat sdf = new SimpleDateFormat("yyyy-MM-dd");
//                value = sdf.format(date);
//            } else {
//                double value_temp = (double) cell.getNumericCellValue();
//                BigDecimal bd = new BigDecimal(value_temp);
//                BigDecimal bd1 = bd.setScale(3, bd.ROUND_HALF_UP);
//                value = bd1.doubleValue();
//
//                DecimalFormat format = new DecimalFormat("#0.###");
//                value = format.format(cell.getNumericCellValue());
//
//            }
//        }
//        if (cell.getCellType() == HSSFCell.CELL_TYPE_BLANK) {
//            value = "";
//        }
//        return value;
//    }

    /**
     * @param sheet
     * @param cellRow жϵĵк
     * @param cellCol жϵĵк
     * @return
     * @throws IOException
     */
    private static int getMergerCellRegionColOfXls(HSSFSheet sheet, int cellRow, int cellCol) throws IOException {
        int retVal = 0;
        int sheetMergerCount = sheet.getNumMergedRegions();
        for (int i = 0; i < sheetMergerCount; i++) {
            CellRangeAddress cra = (CellRangeAddress) sheet.getMergedRegion(i);
            int firstRow = cra.getFirstRow();
            int firstCol = cra.getFirstColumn();
            int lastRow = cra.getLastRow();
            int lastCol = cra.getLastColumn();
            if (cellRow >= firstRow && cellRow <= lastRow) {
                if (cellCol >= firstCol && cellCol <= lastCol) {
                    retVal = lastCol - firstCol + 1;
                    break;
                }
            }
        }
        return retVal;
    }


//    private static int getMergerCellRegionColOfXlsx(XSSFSheet sheet, int cellRow, int cellCol) throws IOException {
//        int retVal = 0;
//        int sheetMergerCount = sheet.getNumMergedRegions();
//        for (int i = 0; i < sheetMergerCount; i++) {
//            CellRangeAddress cra = (CellRangeAddress) sheet.getMergedRegion(i);
//            int firstRow = cra.getFirstRow();
//            int firstCol = cra.getFirstColumn();
//            int lastRow = cra.getLastRow();
//            int lastCol = cra.getLastColumn();
//            if (cellRow >= firstRow && cellRow <= lastRow) {
//                if (cellCol >= firstCol && cellCol <= lastCol) {
//                    retVal = lastCol - firstCol + 1;
//                    break;
//                }
//            }
//        }
//        return retVal;
//    }

    /**
     * @param sheet   ?
     * @param cellRow жϵĵк
     * @param cellCol жϵĵк
     * @return
     * @throws IOException
     */
    private static int getMergerCellRegionRowOfXls(HSSFSheet sheet, int cellRow, int cellCol) throws IOException {
        int retVal = 0;
        int sheetMergerCount = sheet.getNumMergedRegions();
        for (int i = 0; i < sheetMergerCount; i++) {
            CellRangeAddress cra = (CellRangeAddress) sheet.getMergedRegion(i);
            int firstRow = cra.getFirstRow();
            int firstCol = cra.getFirstColumn();
            int lastRow = cra.getLastRow();
            int lastCol = cra.getLastColumn();
            if (cellRow >= firstRow && cellRow <= lastRow) {
                if (cellCol >= firstCol && cellCol <= lastCol) {
                    retVal = lastRow - firstRow + 1;
                    break;
                }
            }
        }
        return 0;
    }

//    /**
//     * @param sheet   ?
//     * @param cellRow жϵĵк
//     * @param cellCol жϵĵк
//     * @return
//     * @throws IOException
//     */
//    private static int getMergerCellRegionRowOfXlsx(XSSFSheet sheet, int cellRow, int cellCol) throws IOException {
//        int retVal = 0;
//        int sheetMergerCount = sheet.getNumMergedRegions();
//        for (int i = 0; i < sheetMergerCount; i++) {
//            CellRangeAddress cra = (CellRangeAddress) sheet.getMergedRegion(i);
//            int firstRow = cra.getFirstRow();
//            int firstCol = cra.getFirstColumn();
//            int lastRow = cra.getLastRow();
//            int lastCol = cra.getLastColumn();
//            if (cellRow >= firstRow && cellRow <= lastRow) {
//                if (cellCol >= firstCol && cellCol <= lastCol) {
//                    retVal = lastRow - firstRow + 1;
//                    break;
//                }
//            }
//        }
//        return 0;
//    }

    /**
     * @param hc
     * @return
     */
    private String convertToStardColorOfXls(HSSFColor hc) {
        StringBuffer sb = new StringBuffer("");
        if (hc != null) {
            int a = HSSFColor.AUTOMATIC.index;
            int b = hc.getIndex();
            if (a == b) {
                return null;
            }
            sb.append("#");
            for (int i = 0; i < hc.getTriplet().length; i++) {
                String str;
                String str_tmp = Integer.toHexString(hc.getTriplet()[i]);
                if (str_tmp != null && str_tmp.length() < 2) {
                    str = "0" + str_tmp;
                } else {
                    str = str_tmp;
                }
                sb.append(str);
            }
        }
        return sb.toString();
    }

//    /**
//     * @param hc
//     * @return
//     */
//    private String convertToStardColorOfXlsx(XSSFColor hc) {
//        StringBuffer sb = new StringBuffer("");
//        if (hc != null) {
//            int a = HSSFColor.AUTOMATIC.index;
//            int b = hc.getIndex();
//            if (a == b) {
//                return null;
//            }
//            sb.append("#");
//            for (int i = 0; i < hc.getTriplet().length; i++) {
//                String str;
//                String str_tmp = Integer.toHexString(hc.getTriplet()[i]);
//                if (str_tmp != null && str_tmp.length() < 2) {
//                    str = "0" + str_tmp;
//                } else {
//                    str = str_tmp;
//                }
//                sb.append(str);
//            }
//        }
//        return sb.toString();
//    }

    /**
     * @param alignment
     * @return
     */
    private String convertAlignToHtml(short alignment) {
        String align = "left";
        switch (alignment) {
            case HSSFCellStyle.ALIGN_LEFT:
                align = "left";
                break;
            case HSSFCellStyle.ALIGN_CENTER:
                align = "center";
                break;
            case HSSFCellStyle.ALIGN_RIGHT:
                align = "right";
                break;
            default:
                break;
        }
        return align;
    }

    /**
     * @param verticalAlignment
     * @return
     */
    private String convertVerticalAlignToHtml(short verticalAlignment) {
        String valign = "middle";
        switch (verticalAlignment) {
            case HSSFCellStyle.VERTICAL_BOTTOM:
                valign = "bottom";
                break;
            case HSSFCellStyle.VERTICAL_CENTER:
                valign = "center";
                break;
            case HSSFCellStyle.VERTICAL_TOP:
                valign = "top";
                break;
            default:
                break;
        }
        return valign;
    }

}
