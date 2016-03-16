package com.example.caimengyun.myreader;

import android.app.Activity;
import android.content.Intent;
import android.os.Bundle;
import android.util.Log;
import android.view.View;
import android.widget.AdapterView;
import android.widget.ListView;
import android.widget.SimpleAdapter;

import java.io.File;
import java.util.ArrayList;
import java.util.HashMap;

/**
 * Created by caimengyun on 16-3-9.
 */
public class FileRead extends Activity {

    private String LOG_TAG = "FileRead_";
    private int fileIcon[] = {R.drawable.dirr, R.drawable.text, R.drawable.docx_win, R.drawable.html, R.drawable.xlsx_win, R.drawable.pptx_win};
    private ListView fileList;//
    private ArrayList<HashMap<String, Object>> fileIconNameList;//图标列表
    private ArrayList<File> fileNameList;//文件列表

    @Override
    protected void onCreate(Bundle savedInstanceState) {
        super.onCreate(savedInstanceState);
        setContentView(R.layout.activity_filelist);
        fileList = (ListView) findViewById(R.id.filelist);
        File path = android.os.Environment.getExternalStorageDirectory();
        File[] f = path.listFiles();//取出所有文件夹以及子文件夹
        fillFileList(f);
    }

    /**
     * 将File转换为List，适配ListView
     *
     * @param files
     */
    private void fillFileList(File[] files) {
        fileNameList = new ArrayList<File>();
        fileIconNameList = new ArrayList<HashMap<String, Object>>();
        for (File file : files) {//files是一个文件数组，相当于每次将files的文件对象赋给file
            int flag = isValidFileOrDir(file);
            if (flag >= 0 && flag <= 5) {
                if (flag == 0) {
                    HashMap<String, Object> map = new HashMap<String, Object>();
                    map.put("Picture", fileIcon[0]);
                    map.put("Filename", file.getName());
                    fileNameList.add(file);
                    fileIconNameList.add(map);
                } else if (flag == 1) {
                    HashMap<String, Object> map = new HashMap<String, Object>();
                    map.put("Picture", fileIcon[1]);
                    map.put("Filename", file.getName());
                    fileNameList.add(file);
                    fileIconNameList.add(map);
                } else if (flag == 2) {
                    HashMap<String, Object> map = new HashMap<String, Object>();
                    map.put("Picture", fileIcon[2]);
                    map.put("Filename", file.getName());
                    fileNameList.add(file);
                    fileIconNameList.add(map);
                } else if (flag == 3) {
                    HashMap<String, Object> map = new HashMap<String, Object>();
                    map.put("Picture", fileIcon[3]);
                    map.put("Filename", file.getName());
                    fileNameList.add(file);
                    fileIconNameList.add(map);
                } else if (flag == 4) {
                    HashMap<String, Object> map = new HashMap<String, Object>();
                    map.put("Picture", fileIcon[4]);
                    map.put("Filename", file.getName());
                    fileNameList.add(file);
                    fileIconNameList.add(map);
                } else if (flag == 5) {
                    HashMap<String, Object> map = new HashMap<String, Object>();
                    map.put("Picture", fileIcon[5]);
                    map.put("Filename", file.getName());
                    fileNameList.add(file);
                    fileIconNameList.add(map);
                } else finish();
            }
        }
        //对应fileIconNameList的adapter，将文件名作为子项，会有很多testview，可以自由延伸
        SimpleAdapter adapter = new SimpleAdapter(this, fileIconNameList, R.layout.item_picname, new String[]{"Picture", "Filename"}, new int[]{R.id.picture, R.id.name});
        fileList.setAdapter(adapter);
        fileList.setOnItemClickListener(new Clicker());
    }

    /**
     * 筛选出文件夹和各种格式的文件
     *
     * @param file
     * @return
     */
    private int isValidFileOrDir(File file) {
        String fileName = file.getName().toLowerCase();
        if (file.isDirectory()) {
            return 0;
        } else if (fileName.toLowerCase().endsWith(".txt")) {
            return 1;
        } else if (fileName.toLowerCase().endsWith(".doc") || (fileName.toLowerCase().endsWith("docx"))) {
            return 2;
        } else if (fileName.toLowerCase().endsWith(".html")) {
            return 3;
        } else if (fileName.toLowerCase().endsWith(".xls") || (fileName.toLowerCase().endsWith(".xlsx"))) {
            return 4;
        } else if (fileName.toLowerCase().endsWith(".pptx") || (fileName.toLowerCase().endsWith("ppt"))) {
            return 5;
        }
        return 6;
    }

    /**
     * 单击事件响应
     */
    private class Clicker implements AdapterView.OnItemClickListener {

        public void onItemClick(AdapterView<?> arg0, View v, int position, long id) {
            Bundle bundle = new Bundle();
            File file = fileNameList.get(position);//定义文件操作指针的位置
            if (file.isDirectory()) {//判断是否是文件夹，是则继续打开
                File[] files = file.listFiles();
                fillFileList(files);
            } else {
                String filePath = file.getAbsolutePath();
                if (filePath.endsWith(".txt")) {
                    Intent intent = new Intent(FileRead.this, ViewTxt.class);
                    bundle.putString("filePath", filePath);//绝对路径
                    intent.putExtras(bundle);
                    startActivityForResult(intent, 0);
                } else if (filePath.endsWith(".doc") || filePath.endsWith(".html") || filePath.endsWith(".docx")) {
                    Intent intent = new Intent(FileRead.this, ViewWord.class);
                    bundle.putString("filePath", filePath);//绝对路径
                    intent.putExtras(bundle);
                    startActivityForResult(intent, 0);
                } else if (filePath.endsWith(".xls") || (filePath.endsWith(".xlsx"))) {
                    Intent intent = new Intent(FileRead.this, ViewExcel.class);
                    bundle.putString("filePath", filePath);//绝对路径
                    intent.putExtras(bundle);
                    startActivityForResult(intent, 0);
                } else {
                    Intent intent = new Intent(FileRead.this, ViewPPTX.class);
                    bundle.putString("filePath", filePath);//绝对路径
                    intent.putExtras(bundle);
                    startActivityForResult(intent, 0);
                }
                finish();
            }
        }
    }
}

