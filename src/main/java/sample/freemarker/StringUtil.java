package sample.freemarker;

/**
 * @author zhaolei
 * @create 2016-09-06 16:30
 */
public class StringUtil {
    public static String logmsg(String msg,String extramsg,String desc) {
        StringBuffer msgBuffer = new StringBuffer();
        msgBuffer.append("\n\n[*]"+msg);
        if(null != extramsg && !"".equals(extramsg)){
            msgBuffer.append("\n[*]"+extramsg);
        }
        if(null != desc && !"".equals(desc)){
            msgBuffer.append("\n[*]"+desc);
        }
        msgBuffer.append("\n");
        return msgBuffer.toString();
    }

    public static String htmlmsg(String msg,String extramsg,String desc) {
        StringBuffer msgBuffer = new StringBuffer();
        msgBuffer.append("[*]"+msg);
        if(null != extramsg && !"".equals(extramsg)){
            msgBuffer.append("<p>[*]"+extramsg+"</p>");
        }
        if(null != desc && !"".equals(desc)){
            msgBuffer.append("<p>[*]"+desc+"</p>");
        }
        return msgBuffer.toString();
    }
}
