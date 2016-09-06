package sample.freemarker;

import org.apache.poi.hwpf.HWPFDocumentCore;
import org.apache.poi.hwpf.converter.WordToHtmlConverter;
import org.apache.poi.hwpf.converter.WordToHtmlUtils;
import org.apache.poi.util.XMLHelper;
import org.docx4j.Docx4J;
import org.docx4j.Docx4jProperties;
import org.docx4j.convert.out.HTMLSettings;
import org.docx4j.convert.out.html.SdtToListSdtTagHandler;
import org.docx4j.convert.out.html.SdtWriter;
import org.docx4j.openpackaging.exceptions.Docx4JException;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.w3c.dom.Document;

import javax.xml.transform.Transformer;
import javax.xml.transform.TransformerFactory;
import javax.xml.transform.dom.DOMSource;
import javax.xml.transform.stream.StreamResult;
import java.io.*;
import java.util.List;
/**
 * Created with IntelliJ IDEA.
 * User: zhaolei
 * Date: 16-9-4
 * Time: 下午2:10
 */
public class FileUtil {
    public static List<String> getFileList(String dirPath,List<String> fileList){
        File dirFile = new File(dirPath);
        String[] sonFilePaths = dirFile.list();
        for (String sonFilePath : sonFilePaths) {
            File sonFile = new File(dirPath+"\\"+sonFilePath);
            if (sonFile.isDirectory()){
                getFileList(dirPath + "\\" + sonFilePath, fileList);
            }else {
                String filename = sonFile.getName();
                if(filename.toLowerCase().endsWith("doc") || filename.toLowerCase().endsWith("docx")){
                    fileList.add(dirPath+"\\"+sonFilePath);
                }
            }
        }
        return fileList;
    }

    public static void docxToHtml(String inputfilepath,String outfilepath,String outDirPath) throws Docx4JException, IOException {
        WordprocessingMLPackage wordMLPackage;
        if (inputfilepath==null) {
            wordMLPackage = WordprocessingMLPackage.createPackage();
        } else {
            wordMLPackage = Docx4J.load(new File(inputfilepath));
        }
        HTMLSettings htmlSettings = Docx4J.createHTMLSettings();

        String imgfilepath = outDirPath+"\\"+ MD5Util.md5(outfilepath)+"_files";
        htmlSettings.setImageDirPath(imgfilepath);
        htmlSettings.setImageTargetUri(imgfilepath);
        htmlSettings.setWmlPackage(wordMLPackage);

        SdtWriter.registerTagHandler("HTML_ELEMENT", new SdtToListSdtTagHandler());

        OutputStream os = new FileOutputStream(outfilepath);

        Docx4jProperties.setProperty("docx4j.Convert.Out.HTML.OutputMethodXML", true);

        Docx4J.toHTML(htmlSettings, os, Docx4J.FLAG_NONE);
        if (wordMLPackage.getMainDocumentPart().getFontTablePart()!=null) {
            wordMLPackage.getMainDocumentPart().getFontTablePart().deleteEmbeddedFontTempFiles();
        }
        // This would also do it, via finalize() methods
        htmlSettings = null;
        wordMLPackage = null;
        os.close();
    }

    public static void docToHtml(String filePath,String outFilePath) throws Exception {
        Document doc = process(new File(filePath));
        DOMSource domSource = new DOMSource(doc);
        OutputStream outputStream = new FileOutputStream(new File(outFilePath));
        StreamResult streamResult = new StreamResult(outputStream);
        TransformerFactory tf = TransformerFactory.newInstance();
        Transformer serializer = tf.newTransformer();
        serializer.setOutputProperty("encoding", "UTF-8");
        serializer.setOutputProperty("indent", "yes");
        serializer.setOutputProperty("method", "html");
        serializer.transform(domSource, streamResult);
        outputStream.close();
    }
    private static Document process(File docFile) throws Exception {
        HWPFDocumentCore wordDocument = WordToHtmlUtils.loadDoc(docFile);
        WordToHtmlConverter wordToHtmlConverter = new WordToHtmlConverter(XMLHelper.getDocumentBuilderFactory().newDocumentBuilder().newDocument());
        wordToHtmlConverter.processDocument(wordDocument);
        return wordToHtmlConverter.getDocument();
    }
}
