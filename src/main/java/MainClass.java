import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableCell;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;

import java.io.BufferedReader;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStreamReader;
import java.util.List;

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

            String fromText = getFromText();
            String toText = getToText();
            changeText(fromText, toText);

            //XWPFWordExtractor extractor = new XWPFWordExtractor(docxFile);
            //System.out.println(extractor.getText());
        } catch (Exception e) {
            e.printStackTrace();
        }


    }

    private String getFromText() throws Exception {
        System.out.println("Введите заменяемый текст");
        return consoleReader.readLine();
    }

    private String getToText() throws Exception {
        System.out.println("Введите текст На который нужно заменить");
        return consoleReader.readLine();
    }

    private void setPathFile() throws IOException {
        System.out.println("Введите путь к файлу.");
        path = consoleReader.readLine();
    }

    private void changeText(String fromText, String toText) {
        for (XWPFParagraph p : docxFile.getParagraphs()) {
            List<XWPFRun> runs = p.getRuns();
            if (runs != null) {
                for (XWPFRun r : runs) {
                    String text = r.getText(0);
                    if (text != null && text.contains(fromText)) {
                        text = text.replace(fromText, toText);
                        r.setText(text, 0);
                    }
                }
            }
        }
        for (XWPFTable tbl : docxFile.getTables()) {
            for (XWPFTableRow row : tbl.getRows()) {
                for (XWPFTableCell cell : row.getTableCells()) {
                    for (XWPFParagraph p : cell.getParagraphs()) {
                        for (XWPFRun r : p.getRuns()) {
                            String text = r.getText(0);
                            if (text != null && text.contains(fromText)) {
                                text = text.replace(fromText, toText);
                                r.setText(text, 0);
                            }
                        }
                    }
                }
            }
        }
        try {
            docxFile.write(new FileOutputStream(path));
        } catch (Exception e) {
            e.printStackTrace();
        }
    }


}
