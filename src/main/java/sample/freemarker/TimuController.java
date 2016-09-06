/*
 * Copyright 2012-2016 the original author or authors.
 *
 * Licensed under the Apache License, Version 2.0 (the "License");
 * you may not use this file except in compliance with the License.
 * You may obtain a copy of the License at
 *
 *      http://www.apache.org/licenses/LICENSE-2.0
 *
 * Unless required by applicable law or agreed to in writing, software
 * distributed under the License is distributed on an "AS IS" BASIS,
 * WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 * See the License for the specific language governing permissions and
 * limitations under the License.
 */

package sample.freemarker;

import com.google.common.collect.Lists;
import org.apache.commons.logging.Log;
import org.apache.commons.logging.LogFactory;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestMethod;

import javax.servlet.http.HttpServletRequest;
import java.io.File;
import java.util.List;
import java.util.Map;

@Controller
public class TimuController {
    Log log = LogFactory.getLog(TimuController.class);

    @RequestMapping(value = "/to_import", method = RequestMethod.GET)
	public String toImport(Map<String, Object> model) {
		return "import";
	}

    @RequestMapping(value = "/do_import", method = RequestMethod.POST)
    public String doImport(Map<String, Object> model,HttpServletRequest request) {
        String docPath = request.getParameter("docPath");
        String outDirPathTemp = docPath+"\\out-temp";
        String outDirPath = docPath+"\\out";
        List<String> docList = Lists.newArrayList();
        try{
            docList = FileUtil.getFileList(docPath+"\\doc",docList);
        } catch (Exception e) {
            String errormsg = "获取文件列表失败！";
            String extramsg = "请检查路径是否正确。文件夹：" + docPath + "\\doc";
            String desc = "";
            log.error(StringUtil.logmsg(errormsg,extramsg,desc), e);
            model.put("error_msg", StringUtil.htmlmsg(errormsg,extramsg,desc));
            return "error";
        }

        File outTempDir = new File(outDirPathTemp);
        if(!outTempDir.exists()){
            outTempDir.mkdirs();
        }
        File outDir = new File(outDirPath);
        if(!outDir.exists()){
            outDir.mkdirs();
        }
        for (String filepath : docList) {
            File file = new File(filepath);
            String filename = file.getName();
            if(filename.toLowerCase().endsWith(".docx")){
                continue;
            }

            String outHtmlFilePath = "";

            try {
                outHtmlFilePath = outDirPathTemp+"\\"+filename.replace(".doc",".html");
                FileUtil.docToHtml(filepath,outHtmlFilePath);
            } catch (Exception e) {
                String errormsg = "读取03版word文档失败！";
                String extramsg = "请另存为03版doc文档后重试。文件：" + filepath;
                String desc = "";
                log.error(StringUtil.logmsg(errormsg,extramsg,desc), e);
                model.put("error_msg", StringUtil.htmlmsg(errormsg,extramsg,desc));
                return "error";
            }

            try {
                String outDocFilePath = filepath.replace("\\doc\\","\\out\\");
                outDocFilePath = outDocFilePath.replace(".doc",".xls");
                File outfile = new File(outDocFilePath);
                if(!outfile.getParentFile().exists()){
                    outfile.getParentFile().mkdirs();
                }
                ImportTool.improtExcel(filepath,outHtmlFilePath,outDocFilePath);
            }
            catch (ReadDataException e){
                String errormsg = "导入失败！";
                String extramsg = "文件：" + filepath;
                String desc = e.getDesc();
                log.error(StringUtil.logmsg(errormsg,extramsg,desc), e);
                model.put("error_msg", StringUtil.htmlmsg(errormsg, extramsg,desc));
                return "error";
            }
            catch (Exception e){
                String errormsg = "导入失败！";
                String extramsg = "请检查格式是否正确后再试。文件：" + filepath;
                String desc = "";
                log.error(StringUtil.logmsg(errormsg,extramsg,desc), e);
                model.put("error_msg", StringUtil.htmlmsg(errormsg, extramsg,desc));
                return "error";
            }

        }

        String msg = "导入成功！";
        String extramsg = "";
        String desc = "";
        model.put("msg", StringUtil.htmlmsg(msg, extramsg,desc));
        return "success";
    }

    @RequestMapping(value = "/to_bq", method = RequestMethod.GET)
    public String toBiaoQian(Map<String, Object> model) {
        return "bq";
    }

    @RequestMapping(value = "/do_bq", method = RequestMethod.POST)
    public String doBiaoQian(Map<String, Object> model,HttpServletRequest request) {
        String docxPath = request.getParameter("docxPath");
        String outDirPathTemp = docxPath+"\\out-temp";
        String outDirPath = docxPath+"\\out";
        List<String> fileList = Lists.newArrayList();
        try {
            fileList = FileUtil.getFileList(docxPath+"\\docx",fileList);
        } catch (Exception e) {
            String errormsg = "获取文件列表失败！";
            String extramsg = "请检查路径是否正确。文件夹：" + docxPath + "\\docx";
            String desc = "";
            log.error(StringUtil.logmsg(errormsg,extramsg,desc), e);
            model.put("error_msg", StringUtil.htmlmsg(errormsg,extramsg,desc));
            return "error";
        }
        File outTempDir = new File(outDirPathTemp);
        if(!outTempDir.exists()){
            outTempDir.mkdirs();
        }
        File outDir = new File(outDirPath);
        if(!outDir.exists()){
            outDir.mkdirs();
        }
        for (String filepath : fileList) {
            File file = new File(filepath);
            String filename = file.getName();
            if(filename.toLowerCase().endsWith(".doc")){
                continue;
            }
            String outHtml1FilePath = "";
            try {
                outHtml1FilePath = outDirPathTemp+"\\"+filename+"-1.html";
                FileUtil.docxToHtml(filepath, outHtml1FilePath, outDirPathTemp);
            } catch (Exception e) {
                String errormsg = "读取07版word文档失败！";
                String extramsg = "请检查文档后后重试。文件：" + filepath;
                String desc = "";
                log.error(StringUtil.logmsg(errormsg,extramsg,desc), e);
                model.put("error_msg", StringUtil.htmlmsg(errormsg,extramsg,desc));
                return "error";
            }

            String outHtml2FilePath = "";
            try {
                outHtml2FilePath = outDirPathTemp+"\\"+filename+"-2.html";
                BqTool.parseHtml(outHtml1FilePath, outHtml2FilePath);
            } catch (Exception e) {
                String errormsg = "解析临时html-1失败！";
                String extramsg = "请检查文档后后重试。文件：" + outHtml1FilePath;
                String desc = "";
                log.error(StringUtil.logmsg(errormsg,extramsg,desc), e);
                model.put("error_msg", StringUtil.htmlmsg(errormsg,extramsg,desc));
                return "error";
            }

            try {
                String outDocxFilePath = filepath.replace("\\docx\\","\\out\\");
                BqTool.toDocx(outHtml2FilePath,outDocxFilePath);
            }catch (Exception e){
                String errormsg = "html-2转docx失败！";
                String extramsg = "请检查文档后后重试。文件：" + outHtml2FilePath;
                String desc = "";
                log.error(StringUtil.logmsg(errormsg,extramsg,desc), e);
                model.put("error_msg", StringUtil.htmlmsg(errormsg,extramsg,desc));
                return "error";
            }

        }
        String msg = "调整标签位置成功！";
        String extramsg = "";
        String desc = "";
        model.put("msg", StringUtil.htmlmsg(msg, extramsg,desc));
        return "success";
    }
}
