package com.dy.apachepoi;

import com.dy.apachepoi.poi.bean.image.PoiImage;
import com.dy.apachepoi.poi.bean.table.IPoiWordTable;
import com.dy.apachepoi.poi.bean.table.PWATwithHeaderBottom;
import com.dy.apachepoi.poi.bean.table.PoiWordAutoTable;
import com.dy.apachepoi.poi.PoiWordUtil;
import org.apache.poi.xwpf.usermodel.XWPFDocument;

import java.util.Arrays;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

/**
 * DEMO
 *
 * @author HuangDongYang<huangdy @ pvc123.com>
 * Create on 2018/12/5 11:12
 */
public class PoiDemoMain {

    /**
     * DEMO
     */
    public static void main(String[] args) {
        //模板文件地址
        String inputUrl = "./word/word-demo.docx";
        //新生产的模板文件
        String outputUrl = "./word/output.docx";

        // （一）文本替换数据
        Map<String, String> testMap = new HashMap<String, String>();
        testMap.put("t_author", "走在刀剑上的羊");
        testMap.put("t_email", "448241091@qq.com");
        testMap.put("t_company", "xxx");
        testMap.put("t_companyNumber", "COMPANY001");
        testMap.put("t_year", "2018");
        testMap.put("t_month", "11");
        testMap.put("t_day", "30");
        testMap.put("t_poi_cool", "【我不会影响左右两边的文本】");

        //（二）动态表格数据
        Map<String, List<IPoiWordTable>> autoTableMap = new HashMap();
        PoiWordAutoTable writeData = new PoiWordAutoTable(5,3);
        writeData.setCell(0, 0, "股东1");
        writeData.setCell(0, 1, "股东类型1");
        writeData.setCell(0, 2, "很有钱");

        writeData.setCell(1, 0, "股东2");
        writeData.setCell(1, 1, "股东类型2");
        writeData.setCell(1, 2, "很穷");

        writeData.setCell(2, 0, "股东3");
        writeData.setCell(2, 1, "股东类型3");
        writeData.setCell(2, 2, "很穷3");

        writeData.setCell(3, 0, "股东4");
        writeData.setCell(3, 1, "股东类型4");
        writeData.setCell(3, 2, "很穷4");

        writeData.setCell(4, 0, "股东5");
        writeData.setCell(4, 1, "5");
        writeData.setCell(4, 2, "5");

        autoTableMap.put("at_row_autoRow", Arrays.<IPoiWordTable>asList(writeData));

        // 动态表格(max01)
        PoiWordAutoTable data1 = new PoiWordAutoTable(2,2);
        data1.setCell(0, 0, "企业名称");
        data1.setCell(0, 1, "xxx");
        data1.setCell(1, 0, "注册号");
        data1.setCell(1, 1, "XXX123");

        PoiWordAutoTable data2 = new PoiWordAutoTable(2,2);
        data2.setCell(0, 0, "企业名称");
        data2.setCell(0, 1, "xxx");
        data2.setCell(1, 0, "注册号");
        data2.setCell(1, 1, "---x2---");
        autoTableMap.put("at_max01_auto", Arrays.<IPoiWordTable>asList(data1, data2));

        // 动态表格(max02)
        PWATwithHeaderBottom pwat1 = new PWATwithHeaderBottom(3,2);
        pwat1.setTitle("1.实际控制人：xxx（身份证号：441900XXXXXXX）查询日期：1995年11月23日");
        pwat1.setBottom("");
        pwat1.setCell(0, 1, "信用卡");
        pwat1.setCell(1, 0, "账户数");
        pwat1.setCell(1, 1, "2个");
        pwat1.setCell(2, 0, "未结清/未注销账户数");
        pwat1.setCell(2, 1, "2个");

        PWATwithHeaderBottom pwat2 = new PWATwithHeaderBottom(3,2);
        pwat2.setTitle("2.实际控制人：xxx（身份证号：xxx）查询日期：2018年11月22日");
        pwat2.setBottom("底部跟随文本");
        pwat2.setCell(0, 1, "信用卡");
        pwat2.setCell(1, 0, "账户数");
        pwat2.setCell(1, 1, "255个");
        pwat2.setCell(2, 0, "未结清/未注销账户数");
        pwat2.setCell(2, 1, "255个");
        autoTableMap.put("at_max02_auto", Arrays.<IPoiWordTable>asList(pwat1, pwat2));

        // noneTextMap
        Map<String, String> noneTableTextMap = new HashMap();
        // noneTableTextMap.put("at_static_demo", null);
        //noneTableTextMap.put("at_max02_auto", null);
        //noneTableTextMap.put("at_row_autoRow", "暂无相关信息");

        // （三）图片替换数据
        Map<String, PoiImage> imageMap = new HashMap<String, PoiImage>();
        imageMap.put("img00_pic1", new PoiImage("./word/4.jpg",
                XWPFDocument.PICTURE_TYPE_JPEG,
                100,
                100));

        PoiWordUtil.changWord(inputUrl, outputUrl, testMap, autoTableMap, noneTableTextMap, imageMap);
    }
}
