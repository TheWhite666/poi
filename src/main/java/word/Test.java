package word;


import com.spire.doc.Document;
import com.spire.doc.FileFormat;
import com.spire.doc.Section;
import com.spire.doc.Table;
import com.spire.doc.documents.BorderStyle;
import com.spire.doc.documents.DefaultTableStyle;
import com.spire.doc.fields.TextRange;
import org.apache.poi.xwpf.usermodel.*;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.*;

import java.awt.*;
import java.io.File;
import java.io.FileOutputStream;
import java.math.BigInteger;


public class Test {

    @org.testng.annotations.Test
    public void Test() throws Exception {
        XWPFDocument document = new XWPFDocument();
        FileOutputStream out = new FileOutputStream(new File("Test.docx"));
        XWPFTable ComTable = document.createTable();
        CTTblWidth comTableWidth = ComTable.getCTTbl().addNewTblPr().addNewTblW();
        comTableWidth.setType(STTblWidth.DXA);
        comTableWidth.setW(BigInteger.valueOf(9072));

        //表格第一行
        XWPFTableRow comTableRowOne = ComTable.getRow(0);
        comTableRowOne.getCell(0).setText("专业一");
        comTableRowOne.addNewTableCell().setText("");
        comTableRowOne.addNewTableCell().setText("");
        comTableRowOne.addNewTableCell().setText("");
        comTableRowOne.addNewTableCell().setText("");

        //表格第二行
        XWPFTableRow comTableRowTwo = ComTable.createRow();
        comTableRowTwo.getCell(0).setText("");
        comTableRowTwo.getCell(1).setText("1级项目");
        comTableRowTwo.getCell(2).setText("");
        comTableRowTwo.getCell(3).setText("");

        //表格第三行
        XWPFTableRow comTableRowThree = ComTable.createRow();
        comTableRowThree.getCell(0).setText("");
        comTableRowThree.getCell(1).setText("勘验情况");
        comTableRowThree.getCell(2).setText("");
        comTableRowThree.getCell(3).setText("");
        document.write(out);
        out.close();
        System.out.println("成功");
    }

    @org.testng.annotations.Test
    public void hpTest() throws Exception {
        //创建Document类的对象
        Document doc = new Document();
        Section sec = doc.addSection();

        //添加一个4行4列的表格
        Table tb = sec.addTable(true);
        tb.resetCells(9, 5);

        /*Table tbb = sec.addTable(true);
        tbb.resetCells(5, 6);*/

        /*//调用方法纵向合并第3列中的第2、3个单元格
        tb.applyVerticalMerge(2, 1, 2);
        //调用方法纵向合并第2列中的第2、3个单元格
        tb.applyVerticalMerge(1, 1, 2);
        //调用方法横向合并第2行中的第2、3个单元格
        tb.applyHorizontalMerge(1, 1, 2);
        //调用方法横向合并第3行中的第2、3个单元格
        tb.applyHorizontalMerge(2, 1, 2);*/

        //调用方法获取第3行中的第3个单元格，拆分成5列5行
        //tb.getRows().get(2).getCells().get(2).splitCell(5,5);

        //doc.loadFromFile("Test.docx");

        //获取第一个section
        /*Section section = doc.getSections().get(0);

        //获取section中第一个表格
        Table table = section.getTables().get(0);

        //给表格应用样式
        table.applyStyle(DefaultTableStyle.Colorful_List);

        //设置表格的右边框
        table.getTableFormat().getBorders().getRight().setBorderType(BorderStyle.Hairline);
        table.getTableFormat().getBorders().getRight().setLineWidth(1.0F);
        table.getTableFormat().getBorders().getRight().setColor(Color.RED);

        //设置表格的顶部边框
        table.getTableFormat().getBorders().getTop().setBorderType(BorderStyle.Hairline);
        table.getTableFormat().getBorders().getTop().setLineWidth(1.0F);
        table.getTableFormat().getBorders().getTop().setColor(Color.RED);

        //设置表格的左边框
        table.getTableFormat().getBorders().getLeft().setBorderType(BorderStyle.Hairline);
        table.getTableFormat().getBorders().getLeft().setLineWidth(1.0F);
        table.getTableFormat().getBorders().getLeft().setColor(Color.RED);

        //设置表格的底部边框
        table.getTableFormat().getBorders().getBottom().setBorderType(BorderStyle.Hairline);
        table.getTableFormat().getBorders().getBottom().setLineWidth(1.0F);
        table.getTableFormat().getBorders().getBottom().setColor(Color.RED);

        //设置表格的水平和垂直边框
        table.getTableFormat().getBorders().getVertical().setBorderType(BorderStyle.Hairline);
        table.getTableFormat().getBorders().getVertical().setColor(Color.RED);
        table.getTableFormat().getBorders().getHorizontal().setBorderType(BorderStyle.Hairline);
        table.getTableFormat().getBorders().getHorizontal().setColor(Color.RED);*/

        doc.loadFromFile("Test.docx");
        Section section = doc.getSections().get(0);


        Table table = section.getTables().get(0);
        TextRange textRange = table.get(0, 0).addParagraph().appendText("一 专业一");
        textRange.getCharacterFormat().setFontName("宋体");
        textRange = table.get(1, 1).addParagraph().appendText("1级项目1");
        textRange.getCharacterFormat().setFontName("宋体");
        textRange = table.get(1, 2).addParagraph().appendText("勘验情况");
        textRange.getCharacterFormat().setFontName("宋体");
        textRange = table.get(1, 3).addParagraph().appendText("附表1");
        textRange.getCharacterFormat().setFontName("宋体");
        textRange = table.get(2, 1).addParagraph().appendText("2级项目1");
        textRange.getCharacterFormat().setFontName("宋体");
        textRange = table.get(2, 2).addParagraph().appendText("勘验情况");
        textRange.getCharacterFormat().setFontName("宋体");
        textRange = table.get(2, 3).addParagraph().appendText("附表1");
        textRange.getCharacterFormat().setFontName("宋体");
        textRange = table.get(2, 4).addParagraph().appendText("附表2");
        textRange.getCharacterFormat().setFontName("宋体");
        textRange = table.get(3, 1).addParagraph().appendText("2级项目2");
        textRange.getCharacterFormat().setFontName("宋体");
        textRange = table.get(3, 2).addParagraph().appendText("勘验情况");
        textRange.getCharacterFormat().setFontName("宋体");
        textRange = table.get(4, 1).addParagraph().appendText("1级项目2");
        textRange.getCharacterFormat().setFontName("宋体");
        textRange = table.get(4, 2).addParagraph().appendText("勘验情况");
        textRange.getCharacterFormat().setFontName("宋体");
        textRange = table.get(5, 0).addParagraph().appendText("二 专业二");
        textRange.getCharacterFormat().setFontName("宋体");

        doc.saveToFile("Test.docx", FileFormat.Docx_2010);
        System.out.println("成功");
    }


}
