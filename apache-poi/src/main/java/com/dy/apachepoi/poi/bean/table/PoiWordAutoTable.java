package com.dy.apachepoi.poi.bean.table;

import java.util.ArrayList;
import java.util.List;

/**
 * 动态表格 基类
 * <p>
 *
 * </p>
 *
 * @author HuangDongYang<huangdy @ pvc123.com>
 * Create on 2018/11/30 10:17
 */
public class PoiWordAutoTable implements IPoiWordTable {

    private static final long serialVersionUID = 376077395317881991L;
    //表格总行数
    private int rows;
    //表格总列数
    private int cols;
    // 行列中对应的内容
    public List<List<String>> tableContext;

    public PoiWordAutoTable(int rows, int cols){
        this.rows = rows;
        this.cols = cols;

        tableContext = new ArrayList(rows);
        for (int row = 0; row < rows; row++){
            tableContext.add(new ArrayList(cols));
            for (int col = 0; col < cols; col++){
                tableContext.get(row).add("");
            }
        }
    }

    /**
     * 给表格某一个位置，设置内容
     * @param row 行下标 0开始计算
     * @param col 列下标 0开始计算
     * @param text 内容
     */
    public void setCell(int row, int col, String text){
        if (row >= rows){
            return;
        }
        if (col >= cols){
            return;
        }

        tableContext.get(row).set(col, text);
    }

    /**
     * 读取表格内容
     * @param row
     * @param col
     * @return
     */
    public String getCell(int row, int col){
        if (row >= rows){
            return "";
        }
        if (col >= cols){
            return "";
        }
        return tableContext.get(row).get(col);
    }

    public int getRows() {
        return rows;
    }

    public int getCols() {
        return cols;
    }


}
