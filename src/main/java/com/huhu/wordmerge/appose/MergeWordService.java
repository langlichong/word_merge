package com.huhu.wordmerge.appose;

import com.aspose.words.Document;
import com.aspose.words.ImportFormatMode;
import com.aspose.words.License;
import com.aspose.words.SaveFormat;
import org.springframework.stereotype.Service;

import java.io.InputStream;

@Service
public class MergeWordService {

    public  void mergeDocx() throws Exception {

        if (!getLicense()) {
            return;
        }

        // 文档内容追加式合并

        String source1 = "E:\\projects\\word_merge\\source1.docx";
        String source2 = "E:\\projects\\word_merge\\source2.docx";
        String merged = "E:\\projects\\word_merge\\merged.docx";

        Document doc = new Document(source1);
        doc.appendDocument(new Document(source2), ImportFormatMode.KEEP_SOURCE_FORMATTING);

        // 由于合并后文档中有水印字符（独占了一页），在使用getLicense()去除水印后，第一页内容为空白，此处删去第一页
        //doc.getSections().removeAt(0);

        doc.save(merged, SaveFormat.DOCX);
    }


    /**
     * 97-2003 版本
     * @throws Exception
     */
    public  void mergeDoc() throws Exception {

        if (!getLicense()) {
            return;
        }

        // 文档内容追加式合并

        String source1 = "E:\\projects\\word_merge\\source1.doc";
        String source2 = "E:\\projects\\word_merge\\source2.doc";
        String merged = "E:\\projects\\word_merge\\merged.doc";

        Document doc = new Document(source1);
        doc.appendDocument(new Document(source2), ImportFormatMode.KEEP_SOURCE_FORMATTING);

        // 由于合并后文档中有水印字符（独占了一页），在使用getLicense()去除水印后，第一页内容为空白，此处删去第一页
        //doc.getSections().removeAt(0);

        doc.save(merged, SaveFormat.DOCX);
    }


    public static boolean getLicense() {
        boolean result = false;
        try {
            InputStream is = Thread.currentThread().getContextClassLoader().getResourceAsStream("license.xml");
            License aposeLic = new License();
            aposeLic.setLicense(is);
            result = true;
        } catch (Exception e) {
            e.printStackTrace();
        }
        return result;
    }
}
