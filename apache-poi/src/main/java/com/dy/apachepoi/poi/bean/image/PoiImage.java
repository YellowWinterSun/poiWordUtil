package com.dy.apachepoi.poi.bean.image;

/**
 * 图片类
 *
 * @author HuangDongYang<huangdy @ pvc123.com>
 * Create on 2018/12/26 15:23
 */
public class PoiImage {

    /** 图片路径 */
    private String imgFilePath;

    /** 图片类型 XWPFDocument.PICTURE_TYPE_XXXX */
    private int PICTURE_TYPE;

    /** 宽px */
    private int imgWidthPx;

    /** 高px */
    private int imgHeightPx;

    public PoiImage(String imgFilePath, int PICTURE_TYPE, int imgWidthPx, int imgHeightPx) {
        this.imgFilePath = imgFilePath;
        this.PICTURE_TYPE = PICTURE_TYPE;
        this.imgWidthPx = imgWidthPx;
        this.imgHeightPx = imgHeightPx;
    }

    public String getImgFilePath() {
        return imgFilePath;
    }

    public void setImgFilePath(String imgFilePath) {
        this.imgFilePath = imgFilePath;
    }

    public int getPICTURE_TYPE() {
        return PICTURE_TYPE;
    }

    public void setPICTURE_TYPE(int PICTURE_TYPE) {
        this.PICTURE_TYPE = PICTURE_TYPE;
    }

    public int getImgWidthPx() {
        return imgWidthPx;
    }

    public void setImgWidthPx(int imgWidthPx) {
        this.imgWidthPx = imgWidthPx;
    }

    public int getImgHeightPx() {
        return imgHeightPx;
    }

    public void setImgHeightPx(int imgHeightPx) {
        this.imgHeightPx = imgHeightPx;
    }
}
