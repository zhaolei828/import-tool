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
import java.util.logging.Logger;

@Controller
public class ImportController {
    Log log = LogFactory.getLog(ImportController.class);

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
        docList = FileUtil.getFileList(docPath+"\\doc",docList);

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
            try {
                String outHtmlFilePath = outDirPathTemp+"\\"+filename.replace(".doc",".html");
                FileUtil.docToHtml(filepath,outHtmlFilePath);

                String outDocxFilePath = filepath.replace("\\doc\\","\\out\\");
                outDocxFilePath = outDocxFilePath.replace(".doc",".xls");
                File outfile = new File(outDocxFilePath);
                if(!outfile.getParentFile().exists()){
                    outfile.getParentFile().mkdirs();
                }
                ImportTool.improtExcel(filepath,outHtmlFilePath,outDocxFilePath);
            }catch (Exception e){
                log.error("Import error!\nFile:"+filename,e);
            }

        }
        return "import";
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
        fileList = FileUtil.getFileList(docxPath+"\\docx",fileList);
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
            try {
                String outHtml1FilePath = outDirPathTemp+"\\"+filename+"-1.html";
                FileUtil.docxToHtml(filepath,outHtml1FilePath,outDirPathTemp);
                String outHtml2FilePath = outDirPathTemp+"\\"+filename+"-2.html";
                BqTool.parseHtml(outHtml1FilePath,outHtml2FilePath);

                String outDocxFilePath = filepath.replace("\\docx\\","\\out\\");
                BqTool.toDocx(outHtml2FilePath,outDocxFilePath);
            }catch (Exception e){
                e.printStackTrace();
            }

        }
        return "bq";
    }
}
