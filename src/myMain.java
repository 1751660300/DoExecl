import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;

import java.awt.*;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.util.ArrayList;
import java.util.Iterator;


public class myMain {
    public static void main(String[] args) {
        try {
            File file = new File("static/template.xls");
            System.out.printf(file.getPath() + "\n");

            FileInputStream fis = new FileInputStream(file);//            读取模板
            HSSFWorkbook xssfWorkbook = new HSSFWorkbook(fis);
            System.out.println("xssfWorkbook对象：" + xssfWorkbook + "\n");

            HSSFSheet xssfSheet;
            int x = 0;
            String sheetName = "项目支出绩效自评表";
//            ArrayList l1 = new ArrayList<List>();

//
//            for (int i = 0; i < 4; i++) {
//                ArrayList<String> l2 = new ArrayList<String>();
//                for (int j = 0; j < 11; j++) {
//                    l2.add(String.valueOf(x++));
//                }
//                l1.add(l2);
//            }

            if (sheetName.equals("")) {
                // 默认取第一个子表
                xssfSheet = xssfWorkbook.getSheetAt(0);
            } else {
                xssfSheet = xssfWorkbook.getSheet(sheetName);
            }

            ArrayList<Integer> addRows = new ArrayList<Integer>();
            addRows.add(5);
            addRows.add(8);

            ArrayList<ArrayList<ArrayList<String>>> data = new ArrayList<ArrayList<ArrayList<String>>>();
            for (int i = 0; i < 2; i++){
                ArrayList rowData = new ArrayList<ArrayList<String>>();
                for (int z = 0; z < 4; z++){
                    ArrayList<String> cellData = new ArrayList<String>();
                    for (int j = 0; j < 11; j++) {
                        cellData.add(String.valueOf(x++));
                    }
                    rowData.add(cellData);
                }
                data.add(rowData);
            }

//            [[[]],     [[]]]
            ArrayList<ArrayList<ArrayList<Integer>>> hbCellNum = new ArrayList();


            for (int i = 0; i < 2; i++){
                ArrayList<ArrayList<Integer>> rowData = new ArrayList();
                rowData.add(new ArrayList<Integer>());
                rowData.add(new ArrayList<Integer>());
                hbCellNum.add(rowData);
            }

//            hbCellNum.get(0).get(0).add(2);
//            hbCellNum.get(0).get(0).add(3);
//            hbCellNum.get(0).get(1).add(8);
//            hbCellNum.get(0).get(1).add(9);
//
//            hbCellNum.get(1).get(0).add(3);
//            hbCellNum.get(1).get(0).add(4);
//            hbCellNum.get(1).get(1).add(9);
//            hbCellNum.get(1).get(1).add(10);

            System.out.printf("数据赋值完成");

            LoadEexcel2.setTemplateData(xssfSheet,addRows,data,hbCellNum);


            HSSFRow lastHssfRow = xssfSheet.getRow(xssfSheet.getLastRowNum()-1);

            System.out.printf(lastHssfRow.getCell(7).toString()+"\n");
            lastHssfRow.getCell(7).setCellValue(100);
            System.out.printf(lastHssfRow.getCell(8).toString()+"\n");
            lastHssfRow.getCell(8).setCellValue(100);

            File file1 = new File("static/1.xls");
            if (file1.exists()) {
                file1.createNewFile();
                System.out.printf("创建文件");
            }
            FileOutputStream FileOutPutStream = new FileOutputStream(file1);
            xssfWorkbook.write(FileOutPutStream);
            FileOutPutStream.close();
            fis.close();
        } catch (Exception e) {
            System.out.printf(e.getMessage());
        }
    }
}
