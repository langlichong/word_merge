package com.huhu.wordmerge.poi;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.xwpf.usermodel.*;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTbl;

import java.io.*;
import java.nio.file.Files;
import java.nio.file.Paths;
import java.util.LinkedList;
import java.util.List;


/**
 *  拷贝word 2 中内容 到 word 1 中，并重新生产新文件 word 3
 */
public class MergeDocxByCopy {

    public static void main(String[] args) throws Exception {

        String source1 = "E:\\projects\\word_merge\\poi\\source1.docx";
        String source2 = "E:\\projects\\word_merge\\poi\\source2.docx";
        String source3 = "E:\\projects\\word_merge\\poi\\source3.docx";
        String copyRes = "E:\\projects\\word_merge\\poi\\files.docx";


        List<String> files = new LinkedList<String>(){{
            add(source1);
            add(source2);
            add(source3);
        }};

        mergeDocxFiles2NewOne(files,copyRes);

    }

    private static void mergeDocxFiles2NewOne(List<String> srcFiles,String mergeFilePath) throws IOException {

        // 默认取第一个docx 为开始文件：其他文件内容追加到第一个文件末尾，最后重新生成一个新docx文件
        String appendStartFile = srcFiles.get(0);
        CustomXWPFDocument dstDoc = new CustomXWPFDocument(Files.newInputStream(Paths.get(appendStartFile)));
        for(int i=1 ;i < srcFiles.size() ; i ++){

            String source = srcFiles.get(i);

            XWPFDocument srcDoc = new XWPFDocument(Files.newInputStream(Paths.get(source)));

            for (IBodyElement bodyElement : srcDoc.getBodyElements()) {

                BodyElementType elementType = bodyElement.getElementType();

                if(elementType == BodyElementType.PARAGRAPH){

                    XWPFParagraph srcPr = (XWPFParagraph) bodyElement;

                    //将该部分的样式添加到目标文档(追加方式)
                    copyStyle(srcDoc,dstDoc,srcDoc.getStyles().getStyle(srcPr.getStyleID()));

                    // 给目标文档新加一个段落
                    XWPFParagraph dstPr = dstDoc.createParagraph();

                    // word中某一element可能是图片，可能是普通段落，无法通过类型判别（不知道MS.word咋想的）
                    // 此处只能将图片 、普通文本段落当做是理想的上下平铺布局关系（不关心文字环绕图片布局）
                    boolean hasImage = false;

                    // 图片处理比较特殊，需要使用XWPFRun 类（万万想不到会有这么一个大拐弯，前后风格。。。）， 将原文档图片插入到新文档
                    for (XWPFRun srcRun : srcPr.getRuns()) {
                        //
                        dstPr.createRun();

                        if (srcRun.getEmbeddedPictures().size() > 0){
                            // 有图片则就不处理其他形式的段落
                            hasImage = true;
                        }

                        for (XWPFPicture pic : srcRun.getEmbeddedPictures()) {

                            XWPFPictureData picData = pic.getPictureData();
                            int picType = picData.getPictureType();
                            byte[] img = pic.getPictureData().getData();
                            long cx = pic.getCTPicture().getSpPr().getXfrm().getExt().getCx();
                            long cy = pic.getCTPicture().getSpPr().getXfrm().getExt().getCy();

                            try {
                                // 插入图片
                                String blipId = dstPr.getDocument().addPictureData(new ByteArrayInputStream(img),picType);
                                dstDoc.createPictureCxCy(blipId, dstDoc.getNextPicNameNumber(picType),cx, cy);

                            } catch (InvalidFormatException e1) {
                                e1.printStackTrace();
                            }
                        }
                    }

                    if (!hasImage){
                      /*  // 方法 1 ： 容易IndexOutOfBoundsException
                        int pos = dstDoc.getParagraphs().size() - 1;
                        dstDoc.setParagraph(srcPr, pos);*/

                        // 方法 2 ：
                        copyParagraph(srcPr,dstPr);
                    }

                }else if(elementType == BodyElementType.TABLE){ //将表格拷过去

                    XWPFTable table = (XWPFTable) bodyElement;

                    copyStyle(srcDoc, dstDoc, srcDoc.getStyles().getStyle(table.getStyleID()));

                    XWPFTable newTable = dstDoc.createTable();

                 /*
                 方法 1  容易 IndexOutOfBoundsException
                 int pos = dstDoc.getTables().size()-1;
                 dstDoc.setTable(pos, table);
                */

                    // 方法2
                    copyTable(table,newTable);

                }
            }
        }

        dstDoc.write(Files.newOutputStream(Paths.get(mergeFilePath)));
    }


    private static void copyParagraph(XWPFParagraph source, XWPFParagraph target) {
        target.getCTP().setPPr(source.getCTP().getPPr());
        for (int i=0; i<source.getRuns().size(); i++ ) {
            XWPFRun run = source.getRuns().get(i);
            XWPFRun targetRun = target.createRun();
            //copy formatting
            targetRun.getCTR().setRPr(run.getCTR().getRPr());
            //no images just copy text
            targetRun.setText(run.getText(0));
        }
    }


    private static void copyTable(XWPFTable source, XWPFTable target) {
        CTTbl sourceCTTbl = source.getCTTbl();
        CTTbl targetCTTbl = target.getCTTbl();
        targetCTTbl.setTblPr(sourceCTTbl.getTblPr());
        targetCTTbl.setTrArray(sourceCTTbl.getTrArray());
    }

    // Copy Styles of Table and Paragraph.
    private static void copyStyle(XWPFDocument srcDoc, XWPFDocument destDoc, XWPFStyle style)
    {
        if (destDoc == null || style == null)
            return;

        if (destDoc.getStyles() == null) {
            destDoc.createStyles();
        }

        List<XWPFStyle> usedStyleList = srcDoc.getStyles().getUsedStyleList(style);
        for (XWPFStyle xwpfStyle : usedStyleList) {
            destDoc.getStyles().addStyle(xwpfStyle);
        }
    }


    /**
     * 拷贝 段落及表格 但没法拷贝图片
     */
    private static void copyParagraphAndTables()throws Exception{

        OutputStream out = new FileOutputStream("Destination.docx");

        XWPFDocument doc = new XWPFDocument(new FileInputStream("source.docx"));
        XWPFDocument destDoc = new XWPFDocument();

        for (IBodyElement bodyElement : doc.getBodyElements()) {

            BodyElementType elementType = bodyElement.getElementType();

            if (elementType.name().equals("PARAGRAPH")) {

                XWPFParagraph pr = (XWPFParagraph) bodyElement;

                destDoc.createParagraph();

                int pos = destDoc.getParagraphs().size() - 1;

                destDoc.setParagraph(pr, pos);

            } else if( elementType.name().equals("TABLE") ) {

                XWPFTable table = (XWPFTable) bodyElement;

                destDoc.createTable();

                int pos = destDoc.getTables().size() - 1;

                destDoc.setTable(pos, table);
            }
        }

        destDoc.write(out);
    }
}
