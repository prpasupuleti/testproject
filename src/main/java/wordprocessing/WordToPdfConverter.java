package wordprocessing;

import fr.opensagres.poi.xwpf.converter.pdf.PdfConverter;
import fr.opensagres.poi.xwpf.converter.pdf.PdfOptions;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.net.URL;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

public class WordToPdfConverter {
    private XWPFDocument document = null;
    private String searchValue;
    private String replacement;

    public WordToPdfConverter(String inputFile) throws IOException {
        URL res = getClass().getClassLoader().getResource(inputFile);
        if (res != null) {
            InputStream inputStream = new FileInputStream(new File(res.getPath()));
            init(new XWPFDocument(inputStream));
        } else {
            System.err.println("Input file [" + inputFile + "] not found.");
        }
    }

    private void init(XWPFDocument xwpfDoc) {
        if (xwpfDoc == null) throw new NullPointerException();
        document = xwpfDoc;
    }

    public void replaceText(Map<String, String> inputMapper) {
        inputMapper.entrySet().stream().forEach(entry -> {
            replaceText(entry.getKey(), entry.getValue());
        });
    }

    public void replaceText(String searchValue, String replacement) {
        this.searchValue = searchValue;
        this.replacement = replacement;
        this.replace(document);
    }

    public void replace(XWPFDocument document) {
        this.document = document;
        List<XWPFParagraph> paragraphs = document.getParagraphs();

        for (XWPFParagraph xwpfParagraph : paragraphs) {
            replace(xwpfParagraph);
        }
    }

    private void replace(XWPFParagraph paragraph) {
        if (hasReplaceableItem(paragraph.getText())) {
            String replacedText = StringUtils.replace(paragraph.getText(), searchValue, replacement);

            removeAllRuns(paragraph);

            insertReplacementRuns(paragraph, replacedText);
        }
    }

    private void insertReplacementRuns(XWPFParagraph paragraph, String replacedText) {
        String[] replacementTextSplitOnCarriageReturn = StringUtils.split(replacedText, "\n");

        for (int j = 0; j < replacementTextSplitOnCarriageReturn.length; j++) {
            String part = replacementTextSplitOnCarriageReturn[j];

            XWPFRun newRun = paragraph.insertNewRun(j);
            newRun.setText(part);

            if (j + 1 < replacementTextSplitOnCarriageReturn.length) {
                newRun.addCarriageReturn();
            }
        }
    }

    private void removeAllRuns(XWPFParagraph paragraph) {
        int size = paragraph.getRuns().size();
        for (int i = 0; i < size; i++) {
            paragraph.removeRun(0);
        }
    }

    private boolean hasReplaceableItem(String runText) {
        return StringUtils.contains(runText, searchValue);
    }

    private File saveToFile(File file) throws Exception {
        FileOutputStream out = null;
        try {
            out = new FileOutputStream(file, false);
            document.write(out);
            document.close();
            return file;
        } catch (Exception e) {
            throw e;
        } finally {
            if (out != null) {
                out.flush();
                out.close();
            }
        }
    }

    public static void ConvertToPDF(String docPath, String pdfPath) {
        try {
            InputStream doc = new FileInputStream(new File(docPath));
            XWPFDocument document = new XWPFDocument(doc);
            PdfOptions options = PdfOptions.create();
            OutputStream out = new FileOutputStream(new File(pdfPath));
            PdfConverter.getInstance().convert(document, out, options);
            System.out.println("Done");
        } catch (FileNotFoundException ex) {
            System.out.println(ex.getMessage());
        } catch (IOException ex) {

            System.out.println(ex.getMessage());
        }
    }

    public static void main(String[] args) throws Exception {
        String inputfile = "sample.docx";
        WordToPdfConverter test = new WordToPdfConverter(inputfile);

        Map<String, String> placeholderMapper = new HashMap<>();
        placeholderMapper.put("<letter_date>", new SimpleDateFormat("dd-MM-yyyy").format(new Date()));
        placeholderMapper.put("<addressee_name>", "Test Addressee name");
        placeholderMapper.put("<addressee_title>", "Test Addressee title");
        placeholderMapper.put("<resource_legal_name>", "Test resource legal name");
        placeholderMapper.put("<addressee_greeting>", "Mr. Test");
        placeholderMapper.put("<AAID>", "Test AgentAccountId");
        placeholderMapper.put("<SCOR>", "Test SCOR number");
        placeholderMapper.put("<PREV_FYXXXX>", "FY2020");
        placeholderMapper.put("<PREV_FY_start>", "FY2020");
        placeholderMapper.put("<PREV_FY_end>", "FY2021");
        placeholderMapper.put("<RCM_signature>", "Test RCM Signature");
        placeholderMapper.put("<RCM_email>", "Test_RCM@Email.com");
        placeholderMapper.put("<RCM_print_name>", "Test RCM Print name");
        placeholderMapper.put("<RCM_title>", "Test RCM title");
        placeholderMapper.put("<CC1_Name>", "Test CC1 Name");
        placeholderMapper.put("<CC1_Title>", "Test CC1 title");
        placeholderMapper.put("<CC2_Name>", "Test CC2 Name");
        placeholderMapper.put("<CC2_Title>", "Test CC2 title");
        placeholderMapper.put("<CC3_Name>", "Test CC3 Name");
        placeholderMapper.put("<CC3_Title>", "Test CC3 title");

        test.replaceText(placeholderMapper);
        test.saveToFile(new File("abc.docx"));
        test.ConvertToPDF("abc.docx", "abc.pdf");
    }
}