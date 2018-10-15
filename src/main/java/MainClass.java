import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.xwpf.extractor.XWPFWordExtractor;
import org.apache.poi.xwpf.model.XWPFHeaderFooterPolicy;
import org.apache.poi.xwpf.usermodel.XWPFDocument;

import java.io.BufferedInputStream;
import java.io.BufferedReader;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStreamReader;

/**
 * Created by Александр on 15.10.2018.
 */
public class MainClass {
    private String path;
    private BufferedReader consoleReader;
    private XWPFDocument docxFile;
    private String tempPath = "C:\\Users\\avsel\\Desktop\\Разворачивание веб-кабинета КД с ядром.docx";

    public static void main(String[] args) {
        new MainClass().start();
    }

    private void start() {
        consoleReader = new BufferedReader(new InputStreamReader(System.in));
        try {
            setPathFile();
            FileInputStream fileInputStream = new FileInputStream(path);

            // открываем файл и считываем его содержимое в объект XWPFDocument
            docxFile = new XWPFDocument(OPCPackage.open(fileInputStream));

            XWPFWordExtractor extractor = new XWPFWordExtractor(docxFile);
            System.out.println(extractor.getText());
        } catch (Exception e){
            e.printStackTrace();
        }


    }

    private void setPathFile() throws IOException{
        System.out.println("Введите путь к файлу.");
        path = consoleReader.readLine();
    }


}
