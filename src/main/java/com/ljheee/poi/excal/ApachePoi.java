package com.ljheee.poi.excal;

import com.ljheee.poi.entity.ExcalEntity;
import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.*;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.math.BigDecimal;
import java.math.RoundingMode;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;

import static java.util.stream.Collectors.toList;


public class ApachePoi {


    public static void main(String[] args) throws EncryptedDocumentException, InvalidFormatException, IOException {


        List<ExcalEntity> list1 = new ArrayList<>();
        list1.add(new ExcalEntity(System.currentTimeMillis(), System.currentTimeMillis(), "dept1", "1", "1", "1", "1", "1", "1", "10%"));
        list1.add(new ExcalEntity(System.currentTimeMillis(), System.currentTimeMillis(), "dept2", "2", "1", "1", "1", "1", "1", "10%"));
        list1.add(new ExcalEntity(System.currentTimeMillis(), System.currentTimeMillis(), "dept3", "3", "1", "1", "1", "1", "1", "10%"));
        list1.add(new ExcalEntity(System.currentTimeMillis(), System.currentTimeMillis(), "dept4", "4", "1", "1", "1", "1", "1", "10%"));
        list1.add(new ExcalEntity(System.currentTimeMillis(), System.currentTimeMillis(), "dept5", "5", "1", "1", "1", "15", "1", "10%"));

        Workbook workbook = makeSheet0(list1, "AHC店");


        // 存储到  当前工程路径下
        FileOutputStream fileOut = new FileOutputStream(new File("export.xls"));
        workbook.write(fileOut);
        fileOut.close();

    }

    /**
     * 数据写入excal
     *
     * @param list1
     * @param poiName
     * @return
     * @throws IOException
     * @throws InvalidFormatException
     */
    private static Workbook makeSheet0(List<ExcalEntity> list1, String poiName) throws IOException, InvalidFormatException {
        InputStream inputStream = ClassLoader.getSystemClassLoader().getResourceAsStream("template_excal.xls");
        Workbook workbook = WorkbookFactory.create(inputStream);
        Sheet sheet0 = workbook.getSheetAt(0);
//        Sheet sheet1 = workbook.getSheetAt(1);

        CellStyle cellStyle = workbook.createCellStyle();
        cellStyle.setAlignment(HSSFCellStyle.ALIGN_RIGHT); // 居右

        SimpleDateFormat sdf = new SimpleDateFormat("yyyy.MM.dd hh:mm");
//        setTitle0(sheet0.getRow(0).getCell(0),"JD","2018-10-01 15:00","2018-10-31 15:00");
        setTitle0(sheet0.getRow(0).getCell(0), poiName, sdf.format(new Date(list1.get(0).beginDate)), sdf.format(new Date(list1.get(0).endDate)));

        // 初始化 N行，7列
        int totleDataRow = list1.size();
        for (int i = 2; i <= totleDataRow + 2; i++) {
            //插入行
            sheet0.shiftRows(i + 1, sheet0.getLastRowNum(), 1, true, false);
            sheet0.createRow(i);
            Row row = sheet0.getRow(i);
            for (int j = 0; j <= 7; j++) {
                row.createCell(j).setCellStyle(cellStyle);
            }
        }

//        sheet0.getRow(2).getCell(0).setCellValue("lllkk");

        // 跳过 表头，开始写数据
        for (int i = 2, j = 0; i <= totleDataRow + 1; i++, j++) {
            Row row = sheet0.getRow(i);
            row.getCell(0).setCellValue(list1.get(j).deptName);
            row.getCell(1).setCellValue(list1.get(j).startAmount);
            row.getCell(2).setCellValue(list1.get(j).procureCost);
            row.getCell(3).setCellValue(list1.get(j).endAmount);
            row.getCell(4).setCellValue(list1.get(j).saleCost);
            row.getCell(5).setCellValue(list1.get(j).saleAmount);
            row.getCell(6).setCellValue(list1.get(j).profit);
            row.getCell(7).setCellValue(list1.get(j).profitRate);
        }

        // 最后一行 总计
        Row lastRow = sheet0.getRow(totleDataRow + 2);
        setLastRow(lastRow, list1);
        return workbook;
    }

    /**
     * 计算 总计行
     *
     * @param lastRow
     * @param list1
     */
    private static void setLastRow(Row lastRow, List<ExcalEntity> list1) {
        Double startAmount = list1.stream().map(item -> Double.parseDouble(item.startAmount)).collect(toList()).stream().map(item -> item).reduce((sum, n) -> sum + n).get();
        Double endAmount = list1.stream().map(item -> Double.parseDouble(item.endAmount)).collect(toList()).stream().map(item -> item).reduce((sum, n) -> sum + n).get();
        Double procureCost = list1.stream().map(item -> Double.parseDouble(item.procureCost)).collect(toList()).stream().map(item -> item).reduce((sum, n) -> sum + n).get();
        Double saleCost = list1.stream().map(item -> Double.parseDouble(item.saleCost)).collect(toList()).stream().map(item -> item).reduce((sum, n) -> sum + n).get();
        Double saleAmount = list1.stream().map(item -> Double.parseDouble(item.saleAmount)).collect(toList()).stream().map(item -> item).reduce((sum, n) -> sum + n).get();
        Double profit = list1.stream().map(item -> Double.parseDouble(item.profit)).collect(toList()).stream().map(item -> item).reduce((sum, n) -> sum + n).get();
        Double profitRate = 100 * profit / saleAmount;
        BigDecimal decimal = new BigDecimal(profitRate);
        decimal.setScale(1, RoundingMode.HALF_UP);
        int dot = decimal.toPlainString().indexOf(".");
        String rate = decimal.toPlainString().substring(0, dot + 2);

        lastRow.getCell(0).setCellValue("总计");
        lastRow.getCell(1).setCellValue(String.valueOf(startAmount));
        lastRow.getCell(2).setCellValue(String.valueOf(procureCost));
        lastRow.getCell(3).setCellValue(String.valueOf(endAmount));
        lastRow.getCell(4).setCellValue(String.valueOf(saleCost));
        lastRow.getCell(5).setCellValue(String.valueOf(saleAmount));
        lastRow.getCell(6).setCellValue(String.valueOf(profit));
        lastRow.getCell(7).setCellType(HSSFCell.CELL_TYPE_STRING);
        lastRow.getCell(7).setCellValue(rate + "%");
    }


    /**
     * 设置 sheet0 的 title
     * 替换 抬头和日期
     *
     * @param cell
     * @param poiName
     * @param time1
     * @param time2
     */
    public static void setTitle0(Cell cell, String poiName, String time1, String time2) {
        String value = cell.getStringCellValue();

        value = value.replaceAll("N", poiName);
        value = value.replaceAll("F", time1);
        value = value.replaceAll("T", time2);
        cell.setCellValue(value);
    }


}
