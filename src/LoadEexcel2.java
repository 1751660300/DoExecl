import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;

import java.util.ArrayList;
import java.util.Iterator;


public class LoadEexcel2 {
    public static void addRow(HSSFSheet xssfSheet, int startRow, ArrayList<String> data) {
//      获取原来的行
        Iterator oldCell = xssfSheet.getRow(startRow).cellIterator();
//      将下面的行都往下移动一行
        xssfSheet.shiftRows(startRow + 1, xssfSheet.getLastRowNum(), 1, true, false);
//      在原来行下创建新行
        HSSFRow newRow = xssfSheet.createRow(startRow + 1);
//      根据原来的行给新行赋值和样式
        int itNum = 0;   // 记录列数
        int rx = 2;      // 记录合并的开始列
        int x = 0;
        while (true) {
            // 遍历原始列并修改新列
            if (oldCell.hasNext()) {
                HSSFCell oc = (HSSFCell) oldCell.next();        // 获取原始行的每一列
                HSSFCell newCell = newRow.createCell(itNum);  // 新行创建一列
                if (!oc.toString().equals("")) {
                    System.out.printf("赋值" + x + "次\n");
                    newCell.setCellValue(data.get(x++));            // 为新创建的列赋值
                }

                newCell.setCellStyle(oc.getCellStyle());        // 为新创建的列附样式

                System.out.printf(startRow + " " + itNum + " " + data.get(itNum) + "\n");

                try {
                    if (oc.toString().equals("") && itNum > 2) {
                        System.out.printf("合并坐标： " + (startRow + 1) + " " + rx + "; " + (startRow + 1) + " " + itNum + "\n");
                        xssfSheet.addMergedRegion(new CellRangeAddress(startRow + 1, startRow + 1, rx, itNum));
                    } else {
                        if (itNum > 2) {
                            rx = itNum;
                        }

                    }
                } catch (Exception e) {
                    continue;
                }

            } else {
                break;
            }
            itNum++;
        }
    }


    // 添加行, 并指定行合并的列数
    public static void addRowHB(HSSFSheet xssfSheet, int startRow, ArrayList<String> data, ArrayList<ArrayList<Integer>> hbCellNum) {
        /*
         * xssfSheet:表格对象
         * startRow: 开始添加的行数,就是从哪一行开始添加
         * data: 添加行需要填充的数据
         * hbCellNum: 合并列的列数 如: 第5列与第6列合并, 第9列与第10列合并 [[4,5],[8,9]]
         *
         * */
//      获取原来的行,就是标题行
        Iterator oldCell = xssfSheet.getRow(startRow).cellIterator();

//      将下面的行都往下移动一行
        xssfSheet.shiftRows(startRow + 1, xssfSheet.getLastRowNum(), 1, true, false);

//      在原来行下创建新行
        HSSFRow newRow = xssfSheet.createRow(startRow + 1);

//      根据原来的行给新行赋值和样式
        int itNum = 0;   // 记录列数
        int x = 0;
        while (true) {
            // 遍历原始列并修改新列
            if (oldCell.hasNext()) {
                HSSFCell oc = (HSSFCell) oldCell.next();        // 获取原始行的每一列

                HSSFCell newCell = newRow.createCell(itNum);  // 新行创建一列

//              如果标题列不为空,则给新添加的列赋值
                if (!oc.toString().equals("")) {
                    System.out.printf("赋值" + x + "次\n");
                    newCell.setCellValue(data.get(x++));            // 为新创建的列赋值
                }

                newCell.setCellStyle(oc.getCellStyle());        // 为新创建的列附样式

                System.out.printf(startRow + " " + itNum + " " + data.get(itNum) + "\n");

            } else {
                break;
            }
            itNum++;
        }

//        合并该合并的列
        for (int i = 0; i < hbCellNum.size(); i++) {
            try {
                System.out.printf("合并坐标： " + (startRow + 1) + " " + hbCellNum.get(i).get(0) + "; " + (startRow + 1) + " " + hbCellNum.get(i).get(1) + "\n");
                xssfSheet.addMergedRegion(new CellRangeAddress(startRow + 1, startRow + 1, hbCellNum.get(i).get(0), hbCellNum.get(i).get(1)));
            } catch (Exception e) {
                continue;
            }
        }
    }


    public static void setTemplateData(HSSFSheet hssfSheet, ArrayList<Integer> addRows, ArrayList<ArrayList<ArrayList<String>>> data, ArrayList<ArrayList<ArrayList<Integer>>> hbCellNum) {
        /*
         * hssfSheet:获一个表格对象
         * addRows: 数据需要在模板中的第几行开始插入,其前面一行为表头标题
         *  data: 最外层list中的元素个数要与addRows中元素个数相同,第二层是行数,第三层是列数,列数的个数为最大列数,有效数据后可以为空值,有效数据按模板顺序添加
         * */

//        遍历要添加的行数
        int x = 0; // 记录添加的行数
        for (int i = 0; i < addRows.size(); i++) {
            int rowNum = addRows.get(i) + x;
            ArrayList<ArrayList<String>> rowDatas = data.get(i);
            for (int j = 0; j < rowDatas.size(); j++) {
                addRowHB(hssfSheet, rowNum + j, rowDatas.get(j), hbCellNum.get(i));
            }
            x += rowDatas.size();
        }
    }

    public static void hbMoreRows(HSSFSheet hssfSheet, ArrayList<Integer> datas){
        /*
        * hssfSheet: 表格对象
        * datas: 合并的左上单元格坐标,右下角单元格坐标,如第第1行,第1列到第2行,第2列 [0,0,1,1]
        * */
        hssfSheet.addMergedRegion(new CellRangeAddress(datas.get(0),datas.get(2),datas.get(1), datas.get(3)));
    }



}
