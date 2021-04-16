import jdk.nashorn.internal.runtime.regexp.joni.Region;
import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.awt.*;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.util.ArrayList;
import java.util.Iterator;

public class LoadExecl {

    public void doExcel() {
        try {
            File file = new File("static/template.xls");
            System.out.printf(file.getPath()+"\n");
            FileInputStream fis = new FileInputStream(file);//            读取模板

            HSSFWorkbook xssfWorkbook = new HSSFWorkbook(fis);
            System.out.println("xssfWorkbook对象：" + xssfWorkbook);
//定义工作表
            HSSFSheet xssfSheet;
            String sheetName = "项目支出绩效自评表";
            ArrayList l1 = new ArrayList<List>();
            int x = 0;

            for(int i = 0; i < 4; i++){
                ArrayList<String> l2 = new ArrayList<String>();
                for(int j = 0; j<7; j++){
                   l2.add(String.valueOf(x++));
                }
                l1.add(l2);
            }

            if (sheetName.equals("")) {
                // 默认取第一个子表
                xssfSheet = xssfWorkbook.getSheetAt(0);
            } else {
                xssfSheet = xssfWorkbook.getSheet(sheetName);
            }

            int max = xssfSheet.getLastRowNum();
            for (int i = 0; i<max; i++){
                try {
                    HSSFRow row = xssfSheet.getRow(i);
                    Iterator cell = row.cellIterator();
                    int j =0;
                    int z =0;
                    if(i == 5){
//                        System.out.printf(String.valueOf(i));
                        Iterator oldCell = xssfSheet.getRow(i).cellIterator();
                        xssfSheet.shiftRows(i+1, xssfSheet.getLastRowNum(), 1, true, false);
                        HSSFRow newRow = xssfSheet.createRow(i+1);
//                        HSSFCellStyle hct =  row.getRowStyle();
//                        newRow.setRowStyle(row.getRowStyle());
                        int itNum = 0;
                        int rx = 2;
                        while (true){
                            if(oldCell.hasNext()){
                                HSSFCell oc = (HSSFCell) oldCell.next();
                                HSSFCell newCell = newRow.createCell(itNum++);
                                newCell.setCellValue(oc.toString());
                                newCell.setCellStyle(oc.getCellStyle());
                                System.out.printf(i + " " + itNum + " "+ oc.toString() +"\n");
                                try{
                                    if(oc.toString().equals("") && itNum > 2){
                                        System.out.printf("坐标： "+(i+1)+" "+rx+"; "+(i+1)+" "+itNum+"\n");
                                        xssfSheet.addMergedRegion(new CellRangeAddress(i+1,i+1, rx,itNum-1));
                                    }else{
                                        if(itNum > 2){
                                            rx = itNum-1;
                                        }

                                    }
                                }catch (Exception e){
                                    continue;
                                }

                            }else{
                                break;
                            }
                        }

                        try {
                            xssfSheet.addMergedRegion(new CellRangeAddress(7, 7, 0, 1));
                        }catch (Exception e){
                            System.out.printf(e.toString());
                        }

                    }

//                    xssfSheet.addMergedRegion(new CellRangeAddress(6,6, 2,3));

//                    while (true){
//                        if (cell.hasNext()){
//                            HSSFCell c = (HSSFCell) cell.next();
//                            System.out.printf(i + " " + j + " "+ c.toString() +"\n");
////                            c.setCellValue("1");
//                        }else{
//                            break;
//                        }
//                        j++;
//                    }
                }catch (Exception e){
                    System.out.printf(e.toString());
                    continue;
                }

            }



            File file1 = new File("static/1.xls");
            if (file1.exists()){
                file1.createNewFile();
                System.out.printf("创建文件");
            }
            FileOutputStream FileOutPutStream = new FileOutputStream(file1);
            xssfWorkbook.write(FileOutPutStream);
            FileOutPutStream.close();
            fis.close();
        }catch (Exception e){
            System.out.printf(e.getMessage());
        }
    }
    public static void main(String[] args) {
        LoadExecl l = new LoadExecl();
        l.doExcel();
    }
}
