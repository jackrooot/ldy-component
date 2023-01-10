package cn.lindianyu.component.excel.sax.xlsx;

import cn.lindianyu.component.excel.sax.base.BaseReadBuild;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.openxml4j.exceptions.OLE2NotOfficeXmlFileException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.xssf.eventusermodel.ReadOnlySharedStringsTable;
import org.apache.poi.xssf.eventusermodel.XSSFReader;
import org.apache.poi.xssf.eventusermodel.XSSFSheetXMLHandler;
import org.apache.poi.xssf.model.StylesTable;
import org.xml.sax.InputSource;
import org.xml.sax.SAXException;
import org.xml.sax.XMLReader;
import org.xml.sax.helpers.XMLReaderFactory;

import java.io.IOException;
import java.io.InputStream;
import java.util.List;

/**
 * @Author: ya
 * @Email: 1002019117@qq.com
 * @Date:create in  2022/11/16
 */
public class ExcelXlsxReadBuild<T> extends BaseReadBuild<T> {

    private XSSFReader r;

    private ExcelXlsxReadOfSax<T> excelXlsxReadOfSax;

    private XMLReader xmlReader;

    private Integer sheetNum = 0;

    public ExcelXlsxReadBuild<T> buildFile(InputStream inputStream, Class<? extends T> clazz){
        try {
            T o = clazz.newInstance();

            OPCPackage opcPackage = OPCPackage.open(inputStream);

            r = new XSSFReader(opcPackage);

            ReadOnlySharedStringsTable readOnlySharedStringsTable = new ReadOnlySharedStringsTable(opcPackage);

            xmlReader = XMLReaderFactory.createXMLReader();

            StylesTable styles = r.getStylesTable();

            DataFormatter dataFormatter = new DataFormatter();

            excelXlsxReadOfSax = new ExcelXlsxReadOfSax<>(o);

            XSSFSheetXMLHandler handler = new XSSFSheetXMLHandler(styles, readOnlySharedStringsTable, excelXlsxReadOfSax, dataFormatter, false);

            xmlReader.setContentHandler(handler);
        } catch (Exception e){
            if (e instanceof OLE2NotOfficeXmlFileException) {
                new OLE2NotOfficeXmlFileException("请检查是否为xlsx文件!").printStackTrace();
            } else {
                e.printStackTrace();
            }
        }
        return this;
    }
    public ExcelXlsxReadBuild<T> buildSheetNum(Integer sheetNum){
        this.sheetNum = sheetNum;
        return this;
    }
    public ExcelXlsxReadBuild<T> buildStartRow(Integer rowNum){
        this.excelXlsxReadOfSax.startRow = rowNum;
        return this;
    }
    public ExcelXlsxReadBuild<T> buildJumpBlankRow(Boolean b){
        excelXlsxReadOfSax.blankFlag = b;
        return this;
    }

    public List<T> build(){
        try {
            if (sheetNum == 0) {
                XSSFReader.SheetIterator sheetsData =(XSSFReader.SheetIterator)r.getSheetsData();
                while (sheetsData.hasNext()) {
                    InputStream next = sheetsData.next();
                    InputSource inputSource = new InputSource(next);
                    xmlReader.parse(inputSource);
                }
            } else {
                InputStream sheet = r.getSheet("rId" + sheetNum);
                InputSource inputSource1 = new InputSource(sheet);
                xmlReader.parse(inputSource1);
            }
        } catch (IOException | InvalidFormatException | SAXException e) {
            e.printStackTrace();
        }
        return excelXlsxReadOfSax.getDataList();
    }
}
