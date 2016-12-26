package sample.freemarker;

import com.google.common.collect.Lists;
import com.google.common.collect.Maps;
import org.apache.commons.logging.Log;
import org.apache.commons.logging.LogFactory;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.jsoup.Jsoup;
import org.jsoup.nodes.Document;
import org.jsoup.nodes.Element;
import org.jsoup.parser.Tag;
import org.jsoup.select.Elements;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Arrays;
import java.util.List;
import java.util.Map;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

/**
 * Created with IntelliJ IDEA.
 * User: zhaolei
 * Date: 16-9-4
 * Time: 下午1:58
 */
public class ImportTool {
    public static void improtExcel(String docPath,String htmlPath,String excelPath) throws Exception{
        File htmlFile = new File(htmlPath);
        Document doc = Jsoup.parse(htmlFile, "UTF-8");
        Element body = doc.body();
        Elements elements = body.children();
        List<Element> pElementList = Lists.newArrayList();
        for (Element element : elements) {
            if(element.tag().getName().equals("p")){
                element.html(element.html().replaceAll("&nbsp;"," "));
                pElementList.add(element);
            }
            if(element.tag().getName().equals("div")){
                Elements divChildren = element.children();
                for (Element divChildElement:divChildren) {
                    if(divChildElement.tag().getName().equals("p")){
                        divChildElement.html(divChildElement.html().replaceAll("&nbsp;"," "));
                        pElementList.add(divChildElement);
                    }
                }
            }
        }

        List<List<Element>> reList = regroup(pElementList);
        List<Timu> timuList = Lists.newArrayList();
        int blankDaAnCount = 0;
        for (List<Element> elementList : reList) {
            if(!isTiGan(elementList.get(0))){
                continue;
            }
            Timu timu = new Timu();
            List<String> xxList = Lists.newArrayList();
            String tigan = "";
            boolean notDaAnFlag = true;
            boolean appenTiGanFlag = false;

            StringBuffer tiGanBuffer = new StringBuffer();

            StringBuffer xiaoTiBuffer = new StringBuffer();

            StringBuffer daAnBuffer = new StringBuffer();

            StringBuffer jieXiBuffer = new StringBuffer();
            for (Element element : elementList) {
                System.out.println("[*] "+element.text());
                //题干
                if (isTiGan(element)){
                    tigan = element.text();
                    try {
                        String tihao = getTihao(element);
                        timu.setTihao(tihao);
                        tigan = tigan.split("\\d+(\\.|．)")[1]+"\n";
                    }catch (Exception e) {
                        if (tigan.length() > 15) {
                            tigan = tigan.substring(0,10);
                        }
                        throw new ReadDataException("500","GetTiGanException","读取题干异常：请检查题干格式。：["+tigan+"]");
                    }
                    appenTiGanFlag = true;
                    continue;
                }

                //收集题干与小题之间的内容
                if (appenTiGanFlag && !isXiaoTi(element) && !isXuanxiang(element)) {
                    if (element.text().trim().length() > 0) {
                        tiGanBuffer.append(element.text()+"\n");
                    }
                }

                if(isXiaoTi(element) && notDaAnFlag){
                    if (isFirstXiaoTi(element)){
                        xiaoTiBuffer.append(tiGanBuffer);
                    }
                    xiaoTiBuffer.append(element.text()+"\n");
                    appenTiGanFlag = false;
                }

                //选项
                if(isXuanxiang(element)){
                    if (isFirstXuanxiang(element)){
                        xiaoTiBuffer.append(tiGanBuffer);
                    }
                    String[] xx = splitXuanxiang(element);
                    for (String s : xx) {
                        if(!s.trim().equals("")){
                            xxList.add(s.trim());
                        }
                    }
                    appenTiGanFlag = false;
                }
                //答案 or 答案里含解析
                if (isDaAn(element)){
                    StringBuffer daAnBufferTemp = new StringBuffer();
                    String daAnFull = getBiaoQianText(element,daAnBufferTemp);
//                    String daAn = element.text();
                    daAnBuffer.append(daAnFull+"\n");
                    notDaAnFlag = false;

                    String[] daAnJieXiArr = daAnFull.split("(解析|过程|分析)(:|：)?");
                    if (daAnJieXiArr.length>1) {
                        jieXiBuffer.append(daAnJieXiArr[1]);
                    }
                }

                //解析
                if (isJieXi(element)){
                    StringBuffer jiexiBufferTemp = new StringBuffer();
                    String jieXiFull = getBiaoQianText(element,jiexiBufferTemp);
                    try {
                        jieXiBuffer.append(jieXiFull.split("(〖|【)?(解析|过程|分析)(:|：)?(〗|】)?")[1]);
                    }catch (Exception e) {
                        if (jieXiFull.length() > 15) {
                            jieXiFull = jieXiFull.substring(0,10);
                        }
                        throw new ReadDataException("500","GetJieXiException","读取解析异常：请检查格式。：["+jieXiFull+"]");
                    }
                }

                //题型
                if (isTixing(element)) {
                    String tixingText = element.text();
                    String tixing = "";
                    try {
                        tixing = tixingText.substring(tixingText.indexOf("】")+1);
                    }catch (Exception e){
                        throw new ReadDataException("500","GetTiXingException","读取题型异常：请检查格式。：["+tixing+"]");
                    }
                    if(tixing.trim().equals("填空题")){
                        timu.setTixing("综合填空");
                    }else if (tixing.trim().equals("选择题")) {
                        timu.setTixing("单选题");
                    }else {
                        timu.setTixing(tixing);
                    }
                }
                //知识点1～5/考点
                if (isZsd(element)) {
                    String zsdText = element.text();
                    String zsd = "";
                    String[] zsdArr = null;
                    try {
                        int endIndex = zsdText.length();
                        if (zsdText.contains("答案")){
                            endIndex = zsdText.indexOf("答案");
                            String daanJieXi = zsdText.split("答案(:|：)?")[1];
                            String[] daanJiexiArr = daanJieXi.split("(解析|过程|分析)(:|：)?");
                            String daan = daanJiexiArr[0];
                            if(daanJiexiArr.length>1){
                                String jiexi = daanJiexiArr[1];
                                jieXiBuffer.append(jiexi);
                            }
                            daAnBuffer.append(daan);
                        }
                        zsd = zsdText.substring(zsdText.indexOf("】")+1,endIndex);
                    }catch (Exception e){
                        throw new ReadDataException("500","GetZsdException","读取三级知识点异常：请检查格式。：["+zsdText+"]");
                    }
                    if(zsd.trim().length()>0){
                        if(zsd.contains(" ")){
                            zsdArr = zsd.split(" ");
                        }else if(zsd.contains("；")){
                            zsdArr = zsd.split("；");
                        }else {
                            zsdArr = new String[]{zsd};
                        }
                    }
                    timu.setZsdArr(zsdArr);
                }

                if (isZsd4(element)) {
                    String zsd4_Text = element.text();
                    String zsd4 = "";
                    String[] zsd4_Arr = null;
                    try {
                        int endIndex = zsd4_Text.length();
                        if (zsd4_Text.contains("答案")){
                            endIndex = zsd4_Text.indexOf("答案");
                            String daanJieXi = zsd4_Text.split("答案(:|：)?")[1];
                            String[] daanJiexiArr = daanJieXi.split("(解析|过程|分析)(:|：)?");
                            String daan = daanJiexiArr[0];
                            if(daanJiexiArr.length>1){
                                String jiexi = daanJiexiArr[1];
                                jieXiBuffer.append(jiexi);
                            }
                            daAnBuffer.append(daan);
                        }
                        zsd4 = zsd4_Text.substring(zsd4_Text.indexOf("】")+1,endIndex);
                    }catch (Exception e){
                        throw new ReadDataException("500","GetZsdException","读取四级知识点异常：请检查格式。：["+zsd4_Text+"]");
                    }
                    if(zsd4.trim().length()>0){
                        if(zsd4.contains(" ")){
                            zsd4_Arr = zsd4.split(" ");
                        }else if(zsd4.contains("；")){
                            zsd4_Arr = zsd4.split("；");
                        }else {
                            zsd4_Arr = new String[]{zsd4};
                        }
                    }
                    timu.setZsd4_Arr(zsd4_Arr);
                }

                //能力结构
                if (isNengLiJieGou(element)) {
                    String nengliText = element.text();
                    String nengli = "";
                    try {
                        nengli = nengliText.substring(nengliText.indexOf("】")+1);
                    }catch (Exception e){
                        throw new ReadDataException("500","GetNljgException","读取能力结构异常：请检查格式。：["+nengliText+"]");
                    }
                    timu.setNljg(nengli);
                }

                //评价
                if (isPingJia(element)) {
                    String pingJiaText = element.text();
                    String pingjia = "";
                    try {
                        pingjia = pingJiaText.substring(pingJiaText.indexOf("】")+1);
                    }catch (Exception e){
                        throw new ReadDataException("500","GetPingJiaException","读取评价异常：请检查格式。：["+pingJiaText+"]");
                    }
                    timu.setPingjia(pingjia);
                }
            }
            timu.setTigan(tigan+xiaoTiBuffer.toString());
            timu.setXuanxiang(makeXuanxiang(xxList));
            String[] daAnArr = daAnBuffer.toString().split("(〖|【)?(答案|解|过程|分析)(:|：)?(〗|】)?");
            if (daAnArr.length>1){
                timu.setDaan(daAnArr[1].trim());
            }else {
                timu.setDaan(daAnArr[0].trim());
            }
            if(null == timu.getDaan() || timu.getDaan().length() == 0){
                blankDaAnCount++;
            }

            timu.setJiexi(jieXiBuffer.toString().trim());
            timuList.add(timu);
            System.out.println(timu.getXuanxiang());
            System.out.println("=============");
        }

        if(timuList.size() > 0 && blankDaAnCount/timuList.size() > 0.5){
            File ansFile = findAnsFile(docPath);
            File qHtmlParentDirFile = htmlFile.getParentFile();
            if(null != ansFile){
                String ansHtmlFilePath = qHtmlParentDirFile.getAbsolutePath() + "\\" + ansFile.getName() + "_.html";
                try{
                    FileUtil.docxToHtml(ansFile.getAbsolutePath(), ansHtmlFilePath, qHtmlParentDirFile.getAbsolutePath());
                } catch (Exception e) {
                    throw new LeiRuntimeException("500","DocxToHtmlException","转html异常。：["+ansFile.getAbsolutePath()+"]");
                }

                //解析答案文档
                org.jsoup.nodes.Document ansDoc = Jsoup.parse(new File(ansHtmlFilePath), "UTF-8");
                Elements ansElements = ansDoc.getElementsByClass("DocDefaults");
                int ansIndex = 0;
                for (Element ansElement : ansElements) {
                    ansElement.html(ansElement.html().replaceAll("&nbsp;"," "));
                }
                for (Element element : ansElements) {
                    if (isTiGan(element)) {
                        StringBuffer ansStringBuffer = new StringBuffer();
                        String ansContent = getBiaoQianText(element,ansStringBuffer);
                        try{
                            ansContent = ansContent.substring(ansContent.indexOf(".")+1);
                        } catch (Exception e) {
                            throw new ReadDataException("500","GetAnsDocsContentException","获取答案文档的内容异常。：["+ansContent+"]");
                        }
                        try{
                            timuList.get(ansIndex).setDaan(ansContent);
                        } catch (Exception e) {
                            throw new ReadDataException("500","GetAnsDocsContentException","题目数量和答案数量不匹配。：[题目数量："+timuList.size() +"， 当前正在读取："+ (ansIndex+1)+"]");
                        }
                        ansIndex++;
                    }
                }
            }
        }

        String htmlFullName = htmlFile.getName();
        String htmlName = htmlFullName.substring(0,htmlFullName.lastIndexOf("."));
        Map<String,String> extraMap = getExtraMap(htmlName);
        try {
            writeIntoExcel(timuList, excelPath, extraMap);
        } catch (Exception e) {
            throw new WriteIntoExcelException("500","WriteIntoExcelException","题目："+tgNow);
        }
    }

    static Map<String,String> getExtraMap(String fileName){
        String bianHao = "";
        try{bianHao = regMatchGetString(fileName,"(：|:)\\d+").substring(1);}catch (Exception e) {}
        String nianFen = "";
        try{nianFen = regMatchGetString(fileName,"\\d{4}年");}catch (Exception e) {}
        String xueKe = "";
        try{xueKe = regMatchGetString(fileName,"(语文|英语|数学|物理|化学)");}catch (Exception e) {}
        String shengFen = "";
        try{shengFen = regMatchGetString(fileName,"年([\\u4e00-\\u9fa5]+)省").substring(1);}catch (Exception e) {}
        String chengShi = "";
        try{chengShi = regMatchGetString(fileName,"省([\\u4e00-\\u9fa5]+)市").substring(1);}catch (Exception e) {}

        Map<String,String> extraMap = Maps.newHashMap();
        extraMap.put("BianHao",bianHao);
        extraMap.put("XueKe",xueKe);
        extraMap.put("ShengFen",shengFen);
        extraMap.put("ChengShi",chengShi);
        extraMap.put("NianFen",nianFen);
        return extraMap;
    }

    static File findAnsFile(String docPath){
        File docFile = new File(docPath);
        String docFileName = docFile.getName();
        if (docFileName.contains("试题")){
            String docName = docFileName.split("试题")[0];
            File parentDirFile = docFile.getParentFile();
            if (parentDirFile.isDirectory()){
                File[] sonFiles = parentDirFile.listFiles();
                for (File sonFile:sonFiles) {
                    String sonFileName = sonFile.getName();
                    String sonName = sonFileName.substring(0,sonFileName.lastIndexOf("."));
                    String regex = "(.*)"+docName+"(.*)答案(.*)";
                    Pattern p = Pattern.compile(regex);
                    Matcher m = p.matcher(sonName);
                    if(m.find()){
                        return sonFile;
                    }
                }
            }
        }
        return null;
    }

    public static List<List<Element>> regroup(List<Element> pElementList) throws BusinessException {
        List<List<Element>> returnList = Lists.newArrayList();
        List<Element> tiMuElementList = Lists.newArrayList();
        String txName = "";
        String txNameTemp = "";
        for (Element element : pElementList) {
            if(isDaTi(element)){
                txName = getDaTi(element);
            }
            if(isTiGan(element) || isDaTi(element)){
                if(tiMuElementList.size()>0){
                    if (txName.trim().length()>0){
                        Element txElement = createElement(txNameTemp);
                        tiMuElementList.add(txElement);
                    }
                    returnList.add(tiMuElementList);
                }
                txNameTemp = txName;
                tiMuElementList = Lists.newArrayList();
            }
            tiMuElementList.add(element);

            if(pElementList.indexOf(element) == pElementList.size()-1){
                if (txName.trim().length()>0){
                    Element txElement = createElement(txNameTemp);
                    tiMuElementList.add(txElement);
                }
                returnList.add(tiMuElementList);
            }
        }
        return returnList;
    }

    public static boolean isTiGan(Element element){
        String regEx="^(\\(.*\\))?(（.*）)?(\\d+(\\.|．)).*";
        Element preElement = element.previousElementSibling();
        String preElementText = "";
        if(null != preElement){
            preElementText = preElement.text().trim();
        }
        return isBiaoQian(element,regEx) && preElementText.length()==0;
    }

    public static String getTihao(Element element){
        String tihaoText="";
        String tihao="";
        String regEx="^\\d+(\\.|．)";
        Pattern p = Pattern.compile(regEx);
        Matcher m = p.matcher(element.text());
        while (m.find()){
            tihaoText = m.group();
        }

        Pattern p2 = Pattern.compile("\\d+");
        Matcher m2 = p2.matcher(tihaoText);
        while (m2.find()){
            tihao = m2.group();
        }
        return tihao;
    }

    public static boolean isXiaoTi(Element element){
        String regEx="^((（|\\()\\d+(）|\\))|(①|②|③|④|⑤|⑥|⑦|⑧|⑨|⑩|⑪|⑫|⑬|⑭|⑮|⑯|⑰|⑱|⑲|⑳)).+";
        return isBiaoQian(element,regEx);
    }

    public static boolean isFirstXiaoTi(Element element){
        String regEx="^(（|\\()1(）|\\))";
        return isBiaoQian(element,regEx);
    }

    public static boolean isFirstXuanxiang(Element element){
        String regEx="^A(\\.|．)+";
        return isBiaoQian(element,regEx);
    }

    public static boolean isDaTi(Element element){
        String regEx="^(一|二|三|四|五|六|七|八|九|十)、.+";
        return isBiaoQian(element,regEx);
    }

    public static String getDaTi(Element element) throws ReadDataException{
        //^[一|二|三|四|五|六|七|八|九|十]、.+
        String name = "";
        String text = element.text();
        String regEx="^(一|二|三|四|五|六|七|八|九|十)、.+";
        Pattern p = Pattern.compile(regEx);
        Matcher m = p.matcher(text);
        String nameText = "";
        while (m.find()){
            nameText = m.group();
        }
        if(!nameText.trim().equals("")){
            try {
                name = nameText.substring(nameText.indexOf("、")+1,nameText.indexOf("题")+1).trim();
            }catch (Exception e) {
                if (text.length() > 15) {
                    text = text.substring(0,10);
                }
                throw new ReadDataException("500","GetDaTiException","读取大题异常：请检查是否有“、和‘题’”。：["+text+"]");
            }
        }
        return name;
    }

    static String regMatchGetString(String inString,String regEx){
        String outString = "";
        Pattern p = Pattern.compile(regEx);
        Matcher m = p.matcher(inString);
        while (m.find()){
            outString = m.group();
        }
        return outString;
    }

    public static Element createElement(String elementName){
        Element element = new Element(Tag.valueOf("p"),"");
        element.attr("class","p3");
        element.html("<span class=\"s3\">【题型】</span><span class=\"s5\">"+elementName+"</span>");
        return element;
    }

    public static boolean isDaAn(Element element){
        String regEx="^(〖|【)?(答案|解:|解：)(〗|】)?.+";
        return isBiaoQian(element,regEx);
    }

    public static boolean isJieXi(Element element){
        String regEx="^(〖|【)?(解析|过程|分析)(〗|】)?.+";
        return isBiaoQian(element,regEx);
    }

    public static boolean isXuanxiang(Element element){
        String regEx="^(A|B|C|D|E|F)(\\.|．)+";
        return isBiaoQian(element,regEx);
    }

    public static boolean isTixing(Element element){
        String regEx="^(〖|【)?题型(〗|】)?.+";
        return isBiaoQian(element,regEx);
    }
    public static boolean isZsd(Element element){
        String regEx="(〖|【)?\\s*考点(〗|】)?.+";
        String regEx2="^(〖|【)?三级知识点(〗|】)?.+";
        boolean b1 = isBiaoQian(element,regEx);
        boolean b2 = isBiaoQian(element,regEx2);
        return b1 || b2;
    }

    public static boolean isZsd4(Element element){
        String regEx="(〖|【)?\\s*(四级知识点)(〗|】)?.+";
        return isBiaoQian(element,regEx);
    }

    public static boolean isNengLiJieGou(Element element){
        String regEx="^(〖|【)?能力结构(〗|】)?.+";
        return isBiaoQian(element,regEx);
    }

    public static boolean isPingJia(Element element){
        String regEx="^(〖|【)?难度等级(〗|】)?.+";
        return isBiaoQian(element,regEx);
    }

    public static String[] splitXuanxiang(Element element){
        String text = element.text();
        String regEx="(A|B|C|D|E|F)(\\.|．)+";
        String[]xx = text.split(regEx);
        return xx;
    }

    public static String makeXuanxiang(List<String> list){
        StringBuilder sb = new StringBuilder();
        for (int i = 0; i < list.size(); i++) {
            switch (i) {
                case 0:
                    sb.append("A::");
                    break;
                case 1:
                    sb.append("B::");
                    break;
                case 2:
                    sb.append("C::");
                    break;
                case 3:
                    sb.append("D::");
                    break;
                case 4:
                    sb.append("E::");
                    break;
                case 5:
                    sb.append("F::");
                    break;
                default:break;
            }
            sb.append(list.get(i));
            if(i != list.size()-1){
                sb.append("\n");
            }
        }
        return sb.toString();
    }

    public static boolean isBiaoQian(Element element,String regex){
        String text = element.text();
        Pattern p = Pattern.compile(regex);
        Matcher m = p.matcher(text);
        if(m.find()){
            return true;
        }else {
            return false;
        }
    }

    public static String getBiaoQianText(Element element,StringBuffer sb){
        if(element.text().trim().length() > 0){
            sb.append(element.text().trim()+"\n");
        }
        Element nextElement = element.nextElementSibling();
        if(null != nextElement){
            String regEx="^(〖|【)([\\u4e00-\\u9fa5]+)(〗|】)";
            if (!isTiGan(nextElement) && !isBiaoQian(nextElement,regEx)) {
                getBiaoQianText(nextElement,sb);
            }
        }
        return sb.toString();
    }

    static String tgNow = "";
    public static void  writeIntoExcel(List<Timu> list,String excelPath,Map<String,String> extraMap) throws IOException {
        Workbook wb = new XSSFWorkbook();
        Sheet s = wb.createSheet();


        Font font = wb.createFont();
        font.setFontHeightInPoints((short)14);
        font.setFontName("Courier New");
        font.setBold(true);
        CellStyle style = wb.createCellStyle();
        style.setFont(font);


        String[] head = new String[]{"题目类型*","编号*","学科*","省份","城市","年份","题型*","错误率","题号","题干","备选答案",
                "正确答案","解析"	,"试题评价","典型题","能力结构","来源","是否有视频","视频文件","视频质量","视频类型"
                ,"第三级知识点1","第三级知识点2","第三级知识点3","第三级知识点4","第三级知识点5"
                ,"第四级知识点1","第四级知识点2","第四级知识点3","第四级知识点4","第四级知识点5"};
        List<String> headList = Arrays.asList(head);
        Row r0 = s.createRow(0);
        for(int cellnum = 0; cellnum < head.length; cellnum ++) {
            Cell c = r0.createCell(cellnum);
            c.setCellValue(head[cellnum]);
            c.setCellStyle(style);
        }

        tgNow = "";
        for(int rownum = 1; rownum <= list.size(); rownum++) {
            Row r = s.createRow(rownum);
            Timu timu = list.get(rownum-1);
            tgNow = timu.getTigan();
            if(tgNow.length()>40){
                tgNow = tgNow.substring(0,30)+"……";
            }
            String tihao = timu.getTihao();
            if (null != tihao && tihao.length()==1){
                tihao = "0"+tihao;
            }
            if (null != extraMap) {
                Cell cTmlx = r.createCell(headList.indexOf("题目类型*"));
                cTmlx.setCellValue("普通");

                Cell cBianHao = r.createCell(headList.indexOf("编号*"));
                cBianHao.setCellValue(extraMap.get("BianHao") + tihao);

                Cell cXueKe = r.createCell(headList.indexOf("学科*"));
                cXueKe.setCellValue(extraMap.get("XueKe"));

                Cell cShengFen = r.createCell(headList.indexOf("省份"));
                cShengFen.setCellValue(extraMap.get("ShengFen"));

                Cell cChengShi = r.createCell(headList.indexOf("城市"));
                cChengShi.setCellValue(extraMap.get("ChengShi"));

                Cell cNianFen = r.createCell(headList.indexOf("年份"));
                cNianFen.setCellValue(extraMap.get("NianFen"));
            }

            Cell cTiHao = r.createCell(headList.indexOf("题号"));
            cTiHao.setCellValue(tihao);

            Cell cTiGan = r.createCell(headList.indexOf("题干"));
            cTiGan.setCellValue(timu.getTigan());

            Cell cXuanXiang = r.createCell(headList.indexOf("备选答案"));
            if (null != timu.getTixing() && timu.getTixing().equals("单选题") && timu.getXuanxiang().equals("")){
                cXuanXiang.setCellValue("A::\nB::\nC::\nD::\n");
            }else {
                cXuanXiang.setCellValue(timu.getXuanxiang());
            }

            Cell cDaAn = r.createCell(headList.indexOf("正确答案"));
            cDaAn.setCellValue(timu.getDaan());

            Cell cJieXi = r.createCell(headList.indexOf("解析"));
            cJieXi.setCellValue(timu.getJiexi());

            Cell cPingJia = r.createCell(headList.indexOf("试题评价"));
            cPingJia.setCellValue(timu.getPingjia());

            Cell cNljg = r.createCell(headList.indexOf("能力结构"));
            cNljg.setCellValue(timu.getNljg());

            Cell cTiXing = r.createCell(headList.indexOf("题型*"));
            cTiXing.setCellValue(timu.getTixing());

            String[] zsdArr = timu.getZsdArr();
            if (null != zsdArr && zsdArr.length > 0) {
                int zsdArr3Length = zsdArr.length;
                if (zsdArr3Length>5){
                    zsdArr3Length = 5;
                }
                for (int i = 1; i <= zsdArr3Length; i++) {
                    Cell cZsd = r.createCell(headList.indexOf("第三级知识点"+i));
                    cZsd.setCellValue(zsdArr[i-1]);
                }
            }

            String[] zsd4_Arr = timu.getZsd4_Arr();
            if (null != zsd4_Arr && zsd4_Arr.length > 0) {
                int zsdArr4Length = zsd4_Arr.length;
                if (zsdArr4Length>5){
                    zsdArr4Length = 5;
                }
                for (int i = 1; i <= zsdArr4Length; i++) {
                    Cell cZsd4 = r.createCell(headList.indexOf("第四级知识点"+i));
                    cZsd4.setCellValue(zsd4_Arr[i-1]);
                }
            }
        }

        String filename = excelPath;
        if(wb instanceof XSSFWorkbook) {
            filename = filename + "x";
        }
        FileOutputStream out = new FileOutputStream(filename);
        wb.write(out);
        out.close();
    }
}
