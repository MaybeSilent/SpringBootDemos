package com.maybesilent.filesoperator.controller;

import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestMethod;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.bind.annotation.ResponseBody;
import org.springframework.web.multipart.MultipartFile;

import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import java.io.*;
import java.net.URLEncoder;

/**
 * 下载相应的模版文件
 */

/**
 * 也可以覆盖相应的模版文件
 */

@Controller
@RequestMapping("/file")
public class FileController {

    private static final String rootPath = Class.class.getClass().getResource("/").getPath();
    private static final String realPath = rootPath + "templates/";
    private static final String fileName = "template.xlsx";

    /**
     * 上传excel文件接口
     * @param file
     * @return 返回上传结果
     */
    @RequestMapping(value = "/upload/template", method = RequestMethod.POST)
    @ResponseBody
    public boolean uploadCCNameList(@RequestParam("nameList") MultipartFile file) {

        File target = new File(realPath,fileName);

        try {
            System.out.println("transfer to new file");
            file.transferTo(target);
        } catch (Exception e) {
            e.printStackTrace();
            return false;
        }
        return true;
    }

    @RequestMapping(value = "/download/template", method = RequestMethod.GET)
    public String downloadCCNameListTemplate(HttpServletRequest request, HttpServletResponse response) {
        try {
            File file = new File(realPath, fileName);
            if (file.exists()) {
                // 配置文件下载
                response.setHeader("content-type", "application/octet-stream");
                response.setContentType("application/octet-stream");
                // 下载文件能正常显示中文
                response.setHeader("Content-Disposition", "attachment;filename=" + URLEncoder.encode(fileName, "UTF-8"));

                // 实现文件下载
                byte[] buffer = new byte[1024];
                FileInputStream fis = null;
                BufferedInputStream bis = null;
                try {
                    fis = new FileInputStream(file);
                    bis = new BufferedInputStream(fis);
                    OutputStream os = response.getOutputStream();
                    int i = bis.read(buffer);
                    while (i != -1) {
                        os.write(buffer, 0, i);
                        i = bis.read(buffer);
                    }
                    //return Result.success("Download the ccNameList successfully!");
                    System.out.println("Download the ccNameList successfully!");
                } catch (Exception e) {
                    System.out.println("Download the ccNameList failed!");
                    //return Result.error(CodeMsg.DOWNLOAD_FAILED, "Download the ccNameList failed!");

                } finally {
                    if (bis != null) {
                        try {
                            bis.close();
                        } catch (IOException e) {
                            e.printStackTrace();
                        }
                    }
                    if (fis != null) {
                        try {
                            fis.close();
                        } catch (IOException e) {
                            e.printStackTrace();
                        }
                    }
                }
            }
        } catch (UnsupportedEncodingException ue) {
            System.out.println("文件名编码异常!");
        } catch (Exception e) {
            System.out.println("文件下载异常!");
        }
        return null;
    }
}
