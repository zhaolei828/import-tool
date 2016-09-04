package sample.freemarker;

import java.io.File;
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
                getFileList(dirPath+"\\"+sonFilePath,fileList);
            }else {
                String filename = sonFile.getName();
                if(filename.toLowerCase().endsWith("doc") || filename.toLowerCase().endsWith("docx")){
                    fileList.add(dirPath+"\\"+sonFilePath);
                }
            }
        }
        return fileList;
    }
}
