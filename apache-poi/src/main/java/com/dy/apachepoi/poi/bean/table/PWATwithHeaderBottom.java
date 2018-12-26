package com.dy.apachepoi.poi.bean.table;

/**
 * 特殊动态表格（携带 题目 和 跟随文本）
 *
 * @author HuangDongYang<huangdy @ pvc123.com>
 * Create on 2018/12/1 11:25
 */
public class PWATwithHeaderBottom extends PoiWordAutoTable {
    private static final long serialVersionUID = 6775968049120051793L;

    /* 表格上方 段落内容 */
    private String title = "";
    /* 表格下方 段落内容 */
    private String bottom = "";

    public PWATwithHeaderBottom(int rows, int cols){
        super(rows, cols);
    }

    public String getTitle() {
        return title;
    }

    public void setTitle(String title) {
        this.title = title;
    }

    public String getBottom() {
        return bottom;
    }

    public void setBottom(String bottom) {
        this.bottom = bottom;
    }
}
