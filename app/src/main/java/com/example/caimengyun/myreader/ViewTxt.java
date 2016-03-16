package com.example.caimengyun.myreader;

import android.app.Activity;
import android.os.Bundle;
import android.widget.TextView;

import java.io.BufferedReader;
import java.io.File;
import java.io.FileInputStream;
import java.io.InputStreamReader;

/**
 * Created by caimengyun on 16-3-10.
 */
public class ViewTxt extends Activity {

    private String filenameString;
    private static final String code = "GB2312";

    @Override
    protected void onCreate(Bundle savedInstanceState) {
        super.onCreate(savedInstanceState);
        setContentView(R.layout.activity_viewfile);
        try {
            Bundle bundle = this.getIntent().getExtras();
            filenameString = bundle.getString("filePath");
            TextView viewFile = (TextView) findViewById(R.id.viewtxt);
            String fileContent = getStringFormFile();
            viewFile.setText(fileContent);
        } catch (Exception e) {
        }
    }

    /**
     * 取出文件内容
     * @return
     */
    private String getStringFormFile() {
        try {
            StringBuffer sBuffer = new StringBuffer();
            //读取文件的内容
            FileInputStream fileInputStream = new FileInputStream(filenameString);
            //将子节流转换为字符刘，要启用从字节到字符的有效转换
            InputStreamReader inputStreamReader = new InputStreamReader(fileInputStream, code);
            //创建字符缓冲区
            BufferedReader in = new BufferedReader(inputStreamReader);
            if (!new File(filenameString).exists()) {
                return null;
            }
            while (in.ready()) {
                sBuffer.append(in.readLine() + "\n");
            }
            in.close();

            return sBuffer.toString();
        } catch (Exception e) {
            e.printStackTrace();
        }
        return null;
    }

}
