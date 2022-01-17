package com.huhu.wordmerge.doc4j;

import org.docx4j.jaxb.Context;
import org.docx4j.openpackaging.contenttype.ContentType;
import org.docx4j.openpackaging.io.SaveToZipFile;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.openpackaging.parts.PartName;
import org.docx4j.openpackaging.parts.WordprocessingML.AlternativeFormatInputPart;
import org.docx4j.openpackaging.parts.WordprocessingML.MainDocumentPart;
import org.docx4j.org.apache.poi.util.IOUtils;
import org.docx4j.relationships.Relationship;
import org.docx4j.wml.CTAltChunk;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.InputStream;
import java.io.OutputStream;

/*
*
 *  1、合并多个docx 格式文档，但是合并后的文档打开时候会提示报错，选择恢复选项（是） 仍然可以查看文档内容
 *  2、不支持 doc 格式
*/


public class MergeDocx4J {
    private static long chunk = 0;
    private static final String CONTENT_TYPE = "application/vnd.openxmlformats-officedocument.wordprocessingml.document";

    public static void mergeDocx(InputStream s1, InputStream s2, OutputStream os) throws Exception {
        WordprocessingMLPackage target = WordprocessingMLPackage.load(s1);
        insertDocx(target.getMainDocumentPart(), IOUtils.toByteArray(s2));
        SaveToZipFile saver = new SaveToZipFile(target);
        saver.save(os);
    }

    private static void insertDocx(MainDocumentPart main, byte[] bytes) throws Exception {
            AlternativeFormatInputPart afiPart = new AlternativeFormatInputPart(new PartName("/part" + (chunk++) + ".docx"));
            afiPart.setContentType(new ContentType(CONTENT_TYPE));
            afiPart.setBinaryData(bytes);
            Relationship altChunkRel = main.addTargetPart(afiPart);

            CTAltChunk chunk = Context.getWmlObjectFactory().createCTAltChunk();
            chunk.setId(altChunkRel.getId());

            main.addObject(chunk);
    }

    public static void main(String[] args)throws Exception {

        String source1 = "E:\\projects\\word_merge\\poi\\source1.docx";
        String source2 = "E:\\projects\\word_merge\\poi\\source2.docx";
        String output = "E:\\projects\\word_merge\\poi\\source12.docx";
        mergeDocx(new FileInputStream(source1),new FileInputStream(source2),new FileOutputStream(output));



        /*String source1 = "E:\\projects\\word_merge\\poi\\source1.doc";
        String source2 = "E:\\projects\\word_merge\\poi\\source2.doc";
        String output = "E:\\projects\\word_merge\\poi\\source12.doc";
        mergeDocx(new FileInputStream(source1),new FileInputStream(source2),new FileOutputStream(output));*/
    }
}
