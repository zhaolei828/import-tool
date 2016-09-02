package sample.freemarker;

import com.google.common.collect.Lists;
import org.docx4j.Docx4J;
import org.docx4j.Docx4jProperties;
import org.docx4j.convert.in.xhtml.XHTMLImporterImpl;
import org.docx4j.convert.out.HTMLSettings;
import org.docx4j.convert.out.html.SdtToListSdtTagHandler;
import org.docx4j.convert.out.html.SdtWriter;
import org.docx4j.model.structure.PageSizePaper;
import org.docx4j.openpackaging.exceptions.Docx4JException;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.jsoup.Jsoup;
import org.jsoup.nodes.Document;
import org.jsoup.nodes.Element;
import org.jsoup.nodes.Entities;
import org.jsoup.parser.Tag;
import org.jsoup.select.Elements;

import java.io.*;
import java.util.List;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

/**
 * @author zhaolei
 * @create 2016-08-13 15:41
 */
public class BqTool {
    private static boolean save = true;
//    public static String docxPath;
//    public static String outDirPathTemp = docxPath+"\\out-temp";
//    public static String outDirPath = docxPath+"\\out";

    static void toHtml(String inputfilepath,String outfilepath,String outDirPath) throws Docx4JException, FileNotFoundException {
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

        OutputStream os;
        if (save) {
            os = new FileOutputStream(outfilepath);
        } else {
            os = new ByteArrayOutputStream();
        }

        Docx4jProperties.setProperty("docx4j.Convert.Out.HTML.OutputMethodXML", true);

		Docx4J.toHTML(htmlSettings, os, Docx4J.FLAG_NONE);
        if (wordMLPackage.getMainDocumentPart().getFontTablePart()!=null) {
            wordMLPackage.getMainDocumentPart().getFontTablePart().deleteEmbeddedFontTempFiles();
        }
        // This would also do it, via finalize() methods
        htmlSettings = null;
        wordMLPackage = null;
    }

    static void parseHtml(String htmlInputFilePath,String outfilepath) throws IOException {
        File inputFile = new File(htmlInputFilePath);
        Document doc = Jsoup.parse(inputFile, "UTF-8");
        /**
         * 题号、分类、学段、年级、一级知识点、二级知识点、三级知识点、四级知识点
         * 难度、能力结构、题型、题干、解答、解析
         *
         * 题号、分类、学段、年级、知识点、
         * 难度、题型、题干、解答、解析、小题
         *
         * 【答案】、【解析】、【题型】、【一级知识点】、【二级知识点】、【三级知识点】、【四级知识点】、【试题评价】和【能力结构】
         */
        Elements elements = doc.getElementsByClass("DocDefaults");//

        List<Element> bqList = Lists.newArrayList();
        List<List<Element>> tmList = Lists.newArrayList();
        for (Element element : elements) {
            if(isBiaoQian(element,"题号")){
                if (bqList.size()>0){
                    tmList.add(bqList);
                }
                bqList = Lists.newArrayList();
            }
            if(element.children().size() > 0){
                bqList.add(element);
            }
            if(elements.indexOf(element) == elements.size()-1){
                tmList.add(bqList);
            }
        }

        List<List<Element>> toElementList = Lists.newArrayList();
        Elements tempElements;

        for (List<Element> elementList : tmList) {
            //给每个子题的答案及解析加上编号
            setSubNo(elementList);

            String tmno = "";
            tempElements = new Elements();

            Element tiHaoElement = getBiaoQianElement(elementList,"题号");
            tempElements.add(tiHaoElement);

            List<Element> timuElements;
            timuElements = betweenThisAndNextBiaoQianElementList(elementList,"题干");

            List<Element> subElements = betweenThisAndNextBiaoQianElementList(elementList, "小题");
            timuElements.addAll(subElements);
            Elements timuTitles = new Elements(timuElements);
            tempElements.addAll(timuTitles);
            //end timu

            //daan
            List<Element> daanElementList = betweenThisAndNextBiaoQianElementList(elementList, "(解答|答案)");
            tempElements.addAll(daanElementList);
            //daan end

            //jiexi
            List<Element> jiexiElementList = betweenThisAndNextBiaoQianElementList(elementList, "解析");
            if(jiexiElementList.size() == 0){
                Element jieXiElement = createBiaoQianElement("解析");
                tempElements.add(jieXiElement);
            }else {
                tempElements.addAll(jiexiElementList);
            }

            //tixing
            Element tiXingElement = getBiaoQianElement(elementList,"题型");
            tempElements.add(tiXingElement);

            Element zsd1Element = getBiaoQianElement(elementList,"一级知识点");
            if(null == zsd1Element){
                zsd1Element = createBiaoQianElement("一级知识点");
            }
            tempElements.add(zsd1Element);

            Element zsd2Element = getBiaoQianElement(elementList,"二级知识点");
            if(null == zsd2Element){
                zsd2Element = createBiaoQianElement("二级知识点");
            }
            tempElements.add(zsd2Element);

            Element zsd3Element = getBiaoQianElement(elementList,"三级知识点");
            if(null == zsd3Element){
                zsd3Element = createBiaoQianElement("三级知识点");
            }
            tempElements.add(zsd3Element);

            Element zsd4Element = getBiaoQianElement(elementList,"四级知识点");
            if(null == zsd4Element){
                zsd4Element = createBiaoQianElement("四级知识点");
            }
            tempElements.add(zsd4Element);

            Element pingXiElement = getBiaoQianElement(elementList,"试题评价");
            if(null == pingXiElement){
                pingXiElement = createBiaoQianElement("试题评价");
            }
            tempElements.add(pingXiElement);

            Element nljgElement = getBiaoQianElement(elementList,"能力结构");
            if(null == nljgElement){
                nljgElement = createBiaoQianElement("能力结构");
            }
            tempElements.add(nljgElement);

            Element blankElement = new Element(Tag.valueOf("p"),"");
            tempElements.add(blankElement);

            toElementList.add(tempElements);
        }
        Elements toElements = new Elements();
        for (List<Element> elementList : toElementList) {
            String tiHaoNo = "";
            int hasDaanBiaoQian = 0;
            int hasJieXiBiaoQian = 0;
            for (Element element : elementList) {
                if(null == element){
                    continue;
                }
                String eText = element.text();
                if(isBiaoQian(element,"题号")){
                    tiHaoNo = getNumber(eText);
                    continue;
                }
                if(isBiaoQian(element,"题干")){
                    String tempHtml = element.html();
                    element.html(reTextBiaoQian(tempHtml, "题干", tiHaoNo + "."));
                }
                if(isBiaoQian(element,"小题")){
                    String tempHtml = element.html();
                    element.html(reTextBiaoQian(tempHtml, "小题", ""));
                }
                if(isBiaoQian(element,"答案")){
                    if(hasDaanBiaoQian == 0){
                        hasDaanBiaoQian = 1;
                    }else {
                        String tempHtml = element.html();
                        element.html(reTextBiaoQian(tempHtml, "答案", ""));
                    }
                }
                if(isBiaoQian(element,"解析")){
                    if(hasJieXiBiaoQian == 0){
                        hasJieXiBiaoQian = 1;
                    }else {
                        String tempHtml = element.html();
                        element.html(reTextBiaoQian(tempHtml, "解析", ""));
                    }
                }
                if(isBiaoQian(element,"题型")){
                    String tempHtml = element.html();
                    element.html(reTextBiaoQian(tempHtml, "题型", "【题型】"));
                }
                toElements.add(element);
            }
        }

        String html="<html><head><meta content=\"text/html; charset=GB2312\" http-equiv=\"Content-Type\" /></head><body>";
        html += toElements.outerHtml();
        html += "</body></html>";
        FileOutputStream fos = new FileOutputStream(outfilepath,false);
        OutputStreamWriter osw = new OutputStreamWriter(fos);
        osw.write(html);
        osw.close();
    }

    public static void toDocx(String inputHtmlFilePath,String outfilepath) throws IOException, Docx4JException {
        File inputFile = new File(inputHtmlFilePath);
        File outFile = new File(outfilepath);
        File outDirFilePath = outFile.getParentFile();
        if(!outDirFilePath.exists()){
            outDirFilePath.mkdirs();
        }
        Document doc = Jsoup.parse(inputFile,"GB2312");
        doc.outputSettings()
                .syntax(Document.OutputSettings.Syntax.xml)
                .escapeMode(Entities.EscapeMode.xhtml);
        WordprocessingMLPackage wordMLPackage = WordprocessingMLPackage.createPackage(PageSizePaper.valueOf("A4"), true); //A4纸，//横版:true

        XHTMLImporterImpl xhtmlImporter = new XHTMLImporterImpl(wordMLPackage);

        wordMLPackage.getMainDocumentPart().getContent().addAll( //导入 xhtml
                xhtmlImporter.convert(doc.html(), doc.baseUri()));

        wordMLPackage.save(new File(outfilepath));
    }

//    public static void doit(String[] args) throws IOException, Docx4JException {
//        fileList(args[0]+"\\docx");
//        outDirPathTemp = args[0]+"\\"+outDirPathTemp;
//        outDirPath = args[0]+"\\"+outDirPath;
//        File outTempDir = new File(outDirPathTemp);
//        if(!outTempDir.exists()){
//            outTempDir.mkdirs();
//        }
//        File outDir = new File(outDirPath);
//        if(!outDir.exists()){
//            outDir.mkdirs();
//        }
//        for (String filepath : fileList) {
//            File file = new File(filepath);
//            String filename = file.getName();
//            try {
//                String outHtml1FilePath = outDirPathTemp+"\\"+filename+"-1.html";
//                toHtml(filepath,outHtml1FilePath);
//                String outHtml2FilePath = outDirPathTemp+"\\"+filename+"-2.html";
//                parseHtml(outHtml1FilePath,outHtml2FilePath);
//
//                String outDocxFilePath = filepath.replace("\\docx\\","\\out\\");
//                toDocx(outHtml2FilePath,outDocxFilePath);
//            }catch (Exception e){
//                e.printStackTrace();
//            }
//
//        }
//    }

    static List<String> fileList = Lists.newArrayList();
    public static void fileList(String dirPath){
        File dirFile = new File(dirPath);
        String[] sonFilePaths = dirFile.list();
        for (String sonFilePath : sonFilePaths) {
            File sonFile = new File(dirPath+"\\"+sonFilePath);
            if (sonFile.isDirectory()){
                fileList(dirPath+"\\"+sonFilePath);
            }else {
                String filename = sonFile.getName();
                if(filename.toLowerCase().endsWith("docx")){
                    fileList.add(dirPath+"\\"+sonFilePath);
                }
            }
        }
    }

    public static String getNumber(String str){
        String regEx="[^0-9]";
        Pattern p = Pattern.compile(regEx);
        Matcher m = p.matcher(str);
        String numStr = m.replaceAll("").trim();
        return numStr;
    }

    public static String reTextBiaoQian(String str,String biaoQian,String now){
        String regEx="[〖【]"+biaoQian+"[〗】]";
        Pattern p = Pattern.compile(regEx);
        Matcher m = p.matcher(str);
        String res = m.replaceAll(now);
        return res;
    }

    public static boolean isBiaoQian(Element element,String bqName){
        String text = element.text();
        String regEx="^[〖【]"+bqName+"[〗】]";
        Pattern p = Pattern.compile(regEx);
        Matcher m = p.matcher(text);
        if(m.find()){
            return true;
        }else {
            return false;
        }
    }

    public static List<Element> timuElementList(List<Element> list){
        return null;
    }

    public static boolean hasSub(Elements elements){
        String text = elements.html();
        String regEx="[〖【]小题[〗】]";
        Pattern p = Pattern.compile(regEx);
        Matcher m = p.matcher(text);
        if(m.find()){
            return true;
        }else {
            return false;
        }
    }

    //获取某个标签的下一个标签
    public static Element getNextBiaoQianElement(List<Element> list,Element tarElement){
        /**
         * 题号、分类、学段、年级、一级知识点、二级知识点、三级知识点、四级知识点
         * 难度、能力结构、题型、题干、解答、解析
         *
         * 题号、分类、学段、年级、知识点、
         * 难度、题型、题干、解答、解析、小题
         *
         * 【答案】、【解析】、【题型】、【一级知识点】、【二级知识点】、【三级知识点】、【四级知识点】、【试题评价】和【能力结构】
         */
        String refStr = "(题号|分类|学段|年级|一级知识点|二级知识点|三级知识点|四级知识点|难度|能力结构|题型|题干|解答|解析|小题|知识点|答案|标记)";
        for (int i = list.indexOf(tarElement)+1; i < list.size() ; i++) {
            Element element = list.get(i);
            if(isBiaoQian(element,refStr)){
                return element;
            }
        }
        return null;
    }

    public static Element getBiaoQianElement(List<Element> list,String biaoQianName){
        for (Element element : list) {
            if(isBiaoQian(element,biaoQianName)){
                return element;
            }
        }
        return null;
    }

    public static List<Integer> sameNameBiaoQianIndexList(List<Element> list,String biaoQianName){
        List<Integer> resList = Lists.newArrayList();
        for (Element element : list) {
            if(isBiaoQian(element,biaoQianName)){
                resList.add(list.indexOf(element));
            }
        }
        return resList;
    }

    @Deprecated
    public static List<Element> betweenBiaoQianElementList(List<Element> list,String beginBiaoQian,String endBiaoQian){
        int index1 = 0;
        int index2;
        List<Element> resList = Lists.newArrayList();
        List<Element> tempList;
        for (Element element : list) {
            if(isBiaoQian(element,beginBiaoQian)){
                index1 = list.indexOf(element);
            }
            if(isBiaoQian(element,endBiaoQian)){
                index2 = list.indexOf(element);
                if(index1>0 && index2>0){
                    tempList = list.subList(index1,index2);
                    resList.addAll(tempList);
                    index1 = 0;
                }
            }
        }
        return resList;
    }

    public static List<Element> betweenThisAndNextBiaoQianElementList(List<Element> list,String beginBiaoQian){
        List<Integer> sameBiaoQianIndexList = sameNameBiaoQianIndexList(list, beginBiaoQian);
        if(null != sameBiaoQianIndexList && sameBiaoQianIndexList.size()>0){
            List<Element> resList = Lists.newArrayList();
            for (int index : sameBiaoQianIndexList) {
                Element thisBiaoQianElement = list.get(index);
                Element nextBiaoQianElement = getNextBiaoQianElement(list, thisBiaoQianElement);
                int nextIndex;
                if(null == nextBiaoQianElement){
                    nextIndex = list.size();
                }else {
                    nextIndex = list.indexOf(nextBiaoQianElement);
                }
                List<Element> tempList = list.subList(index,nextIndex);
                resList.addAll(tempList);
            }
            return resList;
        }
        return Lists.newArrayList();
    }

    public static Element createBiaoQianElement(String biaoQianName){
        Element element = new Element(Tag.valueOf("p"),"");
        element.attr("class","a DocDefaults");
        element.html("<span class=\"a0 \" style=\"\">【"+biaoQianName+"】</span>");
        return element;
    }

    public static void setSubNo(List<Element> list){
        String sunNo = "";
        for (Element element : list) {
            if (isBiaoQian(element,"小题")) {
                String text = element.text();
                sunNo = getSubNo(text);
            }
            if(isBiaoQian(element,"(解答|答案)")){
                String tempHtml = element.html();
                element.html(reTextBiaoQian(tempHtml, "(解答|答案)", "【答案】"+sunNo));
            }
            if(isBiaoQian(element,"解析")){
                String tempHtml = element.html();
                element.html(reTextBiaoQian(tempHtml, "解析", "【解析】"+sunNo));
            }
        }
    }

    public static String getSubNo(String str){
        String numStr="";
        String regEx="[〖【].+[〗】]\\(\\d+\\)";
        Pattern p = Pattern.compile(regEx);
        Matcher m = p.matcher(str);
        while (m.find()){
            numStr = m.group();
        }
        regEx="\\(\\d+\\)";
        p = Pattern.compile(regEx);
        m = p.matcher(numStr);
        while (m.find()){
            numStr = m.group();
        }
        return numStr;
    }
}
