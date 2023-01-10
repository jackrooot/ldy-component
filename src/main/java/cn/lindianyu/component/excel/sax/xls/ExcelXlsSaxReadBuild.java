package cn.lindianyu.component.excel.sax.xls;

import cn.lindianyu.component.excel.sax.base.BaseReadBuild;
import org.apache.poi.hssf.eventusermodel.*;
import org.apache.poi.poifs.filesystem.OfficeXmlFileException;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;

import java.io.IOException;
import java.io.InputStream;
import java.util.List;

/**
 * @Author: ya
 * @Email: 1002019117@qq.com
 * @Date:create in  2022/11/4
 */
public class ExcelXlsSaxReadBuild<T> extends BaseReadBuild<T> {

    private POIFSFileSystem poifsFileSystem = null;

    private HSSFEventFactory factory = new HSSFEventFactory();

    private HSSFRequest request = new HSSFRequest();

    private ExcelXlsReaderOfSax<T> excelXlsReaderOfSax;

    private EventWorkbookBuilder.SheetRecordCollectingListener sheetRecordCollectingListener;

    @Override
    public ExcelXlsSaxReadBuild<T> buildFile(InputStream inputStream, Class<? extends T> clazz){
        try {
            T o = clazz.newInstance();

            poifsFileSystem = new POIFSFileSystem(inputStream);

            excelXlsReaderOfSax = new ExcelXlsReaderOfSax<>(o);

        } catch (Exception e){
          if(e instanceof OfficeXmlFileException){
              throw new OfficeXmlFileException("请检查文件是否为xls文件");
          }else{
              e.printStackTrace();
          }
        }
        return this;
    }

    @Override
    public ExcelXlsSaxReadBuild<T> buildSheetNum(Integer sheetNum){
        excelXlsReaderOfSax.sheetNum = sheetNum;
        return this;
    }

    @Override
    public ExcelXlsSaxReadBuild<T> buildStartRow(Integer rowNum){
        excelXlsReaderOfSax.rowStartIndex = rowNum;
        return this;
    }
    @Override
    public ExcelXlsSaxReadBuild<T> buildJumpBlankRow(Boolean b){
        excelXlsReaderOfSax.blankFlag = b;
        return this;
    }

    private void buildCenter(){
        MissingRecordAwareHSSFListener missingRecordAwareHSSFListener = new MissingRecordAwareHSSFListener(excelXlsReaderOfSax);
        FormatTrackingHSSFListener formatTrackingHSSFListener = new FormatTrackingHSSFListener(missingRecordAwareHSSFListener);
        sheetRecordCollectingListener = new EventWorkbookBuilder.SheetRecordCollectingListener(formatTrackingHSSFListener);
    }

    public List<T> build(){
        try {
            buildCenter();
            request.addListenerForAllRecords(sheetRecordCollectingListener);
            factory.processWorkbookEvents(request,poifsFileSystem);
            poifsFileSystem.close();
        } catch (IOException e) {
            e.printStackTrace();
        }
        List<T> list = excelXlsReaderOfSax.dataList;
        return list;
    }
}