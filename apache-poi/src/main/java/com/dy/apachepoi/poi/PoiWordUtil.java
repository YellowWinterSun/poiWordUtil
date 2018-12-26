package com.dy.apachepoi.poi;

import com.dy.apachepoi.poi.bean.image.PoiImage;
import com.dy.apachepoi.poi.bean.table.IPoiWordTable;
import com.dy.apachepoi.poi.bean.table.PWATwithHeaderBottom;
import com.dy.apachepoi.poi.bean.table.PoiWordAutoTable;
import org.apache.poi.ooxml.POIXMLDocument;
import org.apache.poi.util.Units;
import org.apache.poi.xwpf.model.XWPFHeaderFooterPolicy;
import org.apache.poi.xwpf.usermodel.*;
import org.apache.xmlbeans.XmlCursor;

import java.io.*;
import java.util.*;

/**
 * <p>
 *     Apache Poi 高度封装操作Word文档
 *     版本号: V 1.0.0
 * </p>
 * <p>
 *     (*)占位符定义规则：
 *     (*) 文本替换的占位符，前后必须增加能作为独立 Run 的标识符 . <br>
 *          例: @${t_author}@ 。前后的 @ 与 ${} 的样式不一致，在POI中解析就会被识别为不同的Run， <br>
 *          因此这里会有3个Run，占位符独立占据一个Run，算法在替换文本后，会把前后的@去掉 <br>
 *
 *     文本: ${t_*}
 *     普通表格（文本替换）： ${t_*}
 *     动态表格（行动态）: ${at_row_*}
 *     动态表格（固定表格动态增减） : ${at_max??_*}   (?? 为了根据业务格式的不一致，适应不同的表格）
 * </p>
 * <p>
 *     功能描述：
 *     （1）允许对 占位符 进行文本替换
 *     （2）允许对 某个表格 的行 根据给定数据进行增减和写入
 *     （3）允许对 某个表格 的整个表格 根据给定的数据进行table的增减和写入
 * </p>
 * <p>
 *     工具特色：
 *     （1）占位符高度独立，一个占位符只会占有一个XMPFRun，因此可以保持原有的样式和位置，不会影响到${xxx}外的其他内容。
 *     （2）尽可能优化占位符在进行文本替换时的性能
 *     （3）动态表格增减，使得word更加灵活
 * </p>
 * <p>
 *     展望功能：
 *     （1）图片
 *     （2）富文本
 *     （3）占位符特殊样式设置
 * </p>
 *
 * @author HuangDongYang<huangdy @ pvc123.com>
 * Create on 2018/11/29 21:46
 */
public class PoiWordUtil {

    /**
     * 约定：
     * （1）文本：
     *      仅替换： @${t_xxx}@  （其中左右两边的@，可根据给定word文档不同的段落，自定义，但是其必须样式和周围不一致，其作为一个单独的XWPFRun <br>
     * （2）表格：
     *        1. 静态表格（仅替换表格中的 占位文本）：
     *        2. 动态表格（表头和列数固定，根据给定数据动态适应行数）  ：   <br>
     *        3. 动态表格（整个表格行数和列数都根据给定的 PoiWordAutoTableRow 固定，但是整个表格动态增减 ) :  <br>
     */

    /**
     * 根据模板生成新word文档
     * 判断表格是需要替换还是需要插入，判断逻辑有$为替换，表格无$为插入
     * @param inputUrl 模板word所在的路径
     * @param outputUrl 写出的文件路径（可以是不存在的文件路径）
     * @param textMap 需要替换的数据Map
     * @param autoTableRowMap 动态表格的数据Map
     * @param noneTableTextMap 没有表格时显示的文字内容 （key存在，则说明这个表格是不存在的，直接提示文字内容）
     * @return 成功返回true,失败返回false
     */
    public static boolean changWord(final String inputUrl, final String outputUrl,
                                    final Map<String, String> textMap,
                                    final Map<String, List<IPoiWordTable>> autoTableRowMap,
                                    final Map<String, String> noneTableTextMap,
                                    final Map<String, PoiImage> imageMap) {

        //模板转换默认成功
        boolean changeFlag = true;
        try {
            //获取docx解析对象
            XWPFDocument document = new XWPFDocument(POIXMLDocument.openPackage(inputUrl));

            //处理页眉页脚的占位符
            XWPFHeaderFooterPolicy xwpfHeaderFooterPolicy = document.getHeaderFooterPolicy();
            if (null != xwpfHeaderFooterPolicy.getDefaultHeader()) {
                changeDefaultHeaderFooter(xwpfHeaderFooterPolicy.getDefaultHeader().getParagraphs(), textMap);
            }
            if (null != xwpfHeaderFooterPolicy.getDefaultFooter()) {
                changeDefaultHeaderFooter(xwpfHeaderFooterPolicy.getDefaultFooter().getParagraphs(), textMap);
            }
            //解析替换表格对象
            PoiWordUtil.changeTable(document, textMap, autoTableRowMap, noneTableTextMap);
            //解析替换文本段落,图片对象
            PoiWordUtil.changeText(document, textMap, imageMap);

            //生成新的word
            File file = new File(outputUrl);
            FileOutputStream stream = new FileOutputStream(file);
            document.write(stream);
            stream.close();

        } catch (IOException e) {
            e.printStackTrace();
            changeFlag = false;
        }

        return changeFlag;
    }

    /**
     * 替换页眉页脚
     */
    private static void changeDefaultHeaderFooter(List<XWPFParagraph> listParagraphs,
                                                  Map<String, String> textMap){
        for (XWPFParagraph paragraph : listParagraphs){
            //判断此段落是否需要进行替换
            String text = paragraph.getText();

            if (PoiWordKeyMatchRule.checkText(text)){
                Set<Integer> treeSet = new TreeSet();
                List<XWPFRun> runs = paragraph.getRuns();
                for (int i = 0; i < runs.size(); i++){
                    XWPFRun run = runs.get(i);
                    if (PoiWordKeyMatchRule.checkText(run.toString())){
                        run.setText(changeValue(run.toString(), textMap), 0);
                        treeSet.add(i-1);
                        treeSet.add(i+1);
                    }
                }
                // 移除掉 独立占位符Run
                int totalRemove = 0;
                for (Integer removeIndex : treeSet){
                    paragraph.removeRun(removeIndex - totalRemove);
                    totalRemove++;
                }
            }
        }
    }

    /**
     * 替换段落文本
     * @param document docx解析对象
     * @param textMap 需要替换的信息集合
     * @param imageMap 需要替换的图片信息Map
     */
    private static void changeText(XWPFDocument document,
                                   Map<String, String> textMap,
                                   Map<String, PoiImage> imageMap){
        //获取段落集合
        List<XWPFParagraph> paragraphs = document.getParagraphs();

        for (XWPFParagraph paragraph : paragraphs) {
            //判断此段落时候需要进行替换
            String text = paragraph.getText();

            // 替换文本
            if(PoiWordKeyMatchRule.checkText(text)){
                Set<Integer> treeSet = new TreeSet();    //要删除的run, ${} 前后的 独立边界字符@
                List<XWPFRun> runs = paragraph.getRuns();
                for (int i = 0; i < runs.size(); i++){
                    XWPFRun run = runs.get(i);
                    if (PoiWordKeyMatchRule.checkText(run.toString())){
                        // 占位符的 run 必定是完全独立的
                        run.setText(changeValue(run.toString(), textMap), 0);
                        treeSet.add(i-1);
                        treeSet.add(i+1);
                    }
                }
                // 移除掉 独立占位符的Run
                int totalRemove = 0;
                for (Integer removeIndex : treeSet){
                    paragraph.removeRun(removeIndex - totalRemove);
                    totalRemove++;
                }
            }
            else if (PoiWordKeyMatchRule.checkFullText(text)){
                // 富文本
                Set<Integer> treeSet = new TreeSet();    //要删除的run, ${} 前后的 独立边界字符@
                List<XWPFRun> runs = paragraph.getRuns();
                for (int i = 0; i < runs.size(); i++){
                    XWPFRun run = runs.get(i);
                    if (PoiWordKeyMatchRule.checkFullText(run.toString())){
                        // 占位符的 run 必定是完全独立的
                        run.setText(changeValue(run.toString(), textMap), 0);
                        treeSet.add(i-1);
                        treeSet.add(i+1);
                    }
                }
                // 移除掉 独立占位符的Run
                int totalRemove = 0;
                for (Integer removeIndex : treeSet){
                    paragraph.removeRun(removeIndex - totalRemove);
                    totalRemove++;
                }
            }
            else if (PoiWordKeyMatchRule.checkImage00(text)){
                // 图片替换
                Set<Integer> treeSet = new TreeSet();    //要删除的run, ${} 前后的 独立边界字符@
                List<XWPFRun> runs = paragraph.getRuns();
                for (int i = 0; i < runs.size(); i++){
                    XWPFRun run = runs.get(i);
                    if (PoiWordKeyMatchRule.checkImage00(run.toString())){
                        // 占位符的 run 必定是完全独立的
                        //run.setText(changeValue(run.toString(), textMap), 0);
                        replaceImage(run, imageMap);//FIXME 图片替换
                        treeSet.add(i-1);
                        treeSet.add(i+1);
                    }
                }
                // 移除掉 独立占位符的Run
                int totalRemove = 0;
                for (Integer removeIndex : treeSet){
                    paragraph.removeRun(removeIndex - totalRemove);
                    totalRemove++;
                }
            }
        }
    }

    /**
     * 替换表格对象方法
     * @param document docx解析对象
     * @param textMap 需要替换的信息集合
     * @param autoTableRowMap 需要插入的表格信息集合
     */
    private static void changeTable(XWPFDocument document, Map<String, String> textMap,
                                    Map<String, List<IPoiWordTable>> autoTableRowMap,
                                    Map<String, String> noneTableTextMap){
        //获取表格对象集合
        List<XWPFTable> tables = document.getTables();
        int offerset = 0;   //偏移量
        for (int i = 0; i < tables.size(); i++) {

            XWPFTable table = tables.get(i);    //获取当前table
            if (null == table || table.getRows().size() == 0 ||
                    table.getRows().get(0) == null ||
                    table.getRows().get(0).getTableCells() == null ||
                    table.getRows().get(0).getTableCells().size() == 0){
                continue;
            }
            if (PoiWordKeyMatchRule.checkAutoTableRow(table.getRow(0).getCell(0).getText())){
                //判断是否为 特殊表格（动态增减行）
                String tableKey = table.getRow(0).getCell(0).getText().trim();
                tableKey = tableKey.substring(2, tableKey.length() - 1);
                if (noneTableTextMap.containsKey(tableKey)){
                    // 存在tableKey，说明这个表格无内容
                    tableIsNull(table, noneTableTextMap.get(tableKey));
                }
                else {
                    doAutoTableRow(table, autoTableRowMap);
                }
            }
            else if (PoiWordKeyMatchRule.checkAutoTableMax01(table.getRow(0).getCell(0).getText())){
                //判断是否为 特殊表格（动态增减整个表格）
                String tableKey = table.getRow(0).getCell(0).getText().trim();
                tableKey = tableKey.substring(2, tableKey.length() - 1);
                if (noneTableTextMap.containsKey(tableKey)){
                    // 存在tableKey，说明这个表格无内容
                    tableIsNull(table, noneTableTextMap.get(tableKey));
                }
                else {
                    doAutoTableMax(document, i, table, autoTableRowMap);
                }
            }
            else if (PoiWordKeyMatchRule.checkAutoTableMax02(table.getRow(0).getCell(0).getText())){
                //判断是否为 特殊表格（动态增减整个表格携带标题和随后内容）

                String tableKey = table.getRow(0).getCell(0).getText().trim();
                tableKey = tableKey.substring(2, tableKey.length() - 1);
                if (noneTableTextMap.containsKey(tableKey)){
                    // 存在tableKey，说明这个表格无内容
                    tableIsNull(table, noneTableTextMap.get(tableKey));
                }
                else {
                    doAutoTableMax02(document, i, table, autoTableRowMap);
                }
            }
            else if (PoiWordKeyMatchRule.checkStaticTable(table.getRow(0).getCell(0).getText())){
                // 获取表格头的key名
                String tableKey = table.getRow(0).getCell(0).getText().trim();
                tableKey = tableKey.substring(2, tableKey.length() - 1);

                if (noneTableTextMap.containsKey(tableKey)){
                    // noneTableTextMap存在tableKey，说明这个table没有内容，用无内容提示文字代替
                    tableIsNull(table, noneTableTextMap.get(tableKey));
                }
                else {
                    //删除第一行
                    table.removeRow(0);
                    //普通替换文本 表格
                    List<XWPFTableRow> rows = table.getRows();
                    // 遍历表格每一行，并替换文本
                    eachTable(rows, textMap);
                }
            }
        }
    }

    /**
     * 图片替换
     * @param run
     * @return
     */
    private static boolean replaceImage(XWPFRun run, Map<String, PoiImage> imageMap){
        String tableKey = run.getText(0);
        tableKey = tableKey.substring(2, tableKey.length() - 1);
        run.setText("", 0); //清空文字
        try {
            if (imageMap.containsKey(tableKey)) {
                PoiImage poiImage = imageMap.get(tableKey);
                if (null != poiImage) {
                    run.addPicture(new FileInputStream(poiImage.getImgFilePath()),
                            poiImage.getPICTURE_TYPE(),
                            poiImage.getImgFilePath(),
                            Units.toEMU(poiImage.getImgWidthPx()),
                            Units.toEMU(poiImage.getImgHeightPx()));
                }
            }
        } catch (Exception e) {
            e.printStackTrace();
        }
        return true;
    }

    /**
     * 特殊表格（动态增减行） 替换
     * @param table
     * @param autoTableRowMap
     */
    private static boolean doAutoTableRow(XWPFTable table, Map<String,
            List<IPoiWordTable>> autoTableRowMap){
        //获取行头的 key
        String tableKey = table.getRow(0).getCell(0).getText().trim();
        tableKey = tableKey.substring(2, tableKey.length() - 1);

        List<IPoiWordTable> list = autoTableRowMap.get(tableKey);
        if (null == list && list.isEmpty()){
            table.removeRow(0);
            table.removeRow(1);
            return false;
        }
        // 动态行，list有且只有一个对象
        PoiWordAutoTable writeData = (PoiWordAutoTable) list.get(0);

        // 由于模板引擎默认已经有一行空的，所以少创建一行
        for (int i = 0; i < writeData.getRows()-1; i++){
            XWPFTableRow xwpfTableRow = table.createRow();

            copyTableRow(xwpfTableRow, table.getRow(2));
        }

        //table - row
        List<XWPFTableRow> rows = table.getRows();
        for (int i = 2; i < rows.size(); i++){
            List<XWPFTableCell> cells = table.getRow(i).getTableCells();
            for (int j = 0; j < cells.size(); j++){
                XWPFTableCell cell = cells.get(j);

                cellSetTextOverride(cell, writeData.getCell(i-2, j));
                //cell.setText(writeData.getCell(i-2, j));
            }
        }

        table.removeRow(0);
        return true;
    }

    /**
     * 特殊表格（动态增减整个表格） 替换
     * @param table
     * @param autoTableRowMap
     * @return
     */
    private static int doAutoTableMax(XWPFDocument document,
                                      int tableIndex,
                                      XWPFTable table,
                                      Map<String, List<IPoiWordTable>> autoTableRowMap){
        //获取行头的 key
        String tableKey = table.getRow(0).getCell(0).getText().trim();
        tableKey = tableKey.substring(2, tableKey.length() - 1);

        List<IPoiWordTable> list = autoTableRowMap.get(tableKey);
        if (null == list || list.isEmpty()){
            table.removeRow(0);
            return -1;
        }
        // 先删除模板table的第一行
        table.removeRow(0);

        int tableOffset = 0; //table偏移量

        List<XWPFTable> listNewTable = new ArrayList();   //存放需要处理的新表格

        // 动态整个表格,list有多少个，就生成多少个表格
        for (int i = 0; i < list.size()-1; i++){

            XmlCursor cursor = table.getCTTbl().newCursor();
            cursor.toEndToken();
            while(cursor.toNextToken() != XmlCursor.TokenType.START);
            XWPFParagraph newParagraph = document.insertNewParagraph(cursor);
            newParagraph.createRun().setText("\n", 0);

            cursor = newParagraph.getCTP().newCursor();
            cursor.toEndToken();
            while(cursor.toNextToken() != XmlCursor.TokenType.START);

            XWPFTable newTbl = document.insertNewTbl(cursor);
            while (newTbl != null && !newTbl.getRows().isEmpty()){
                newTbl.removeRow(0);
            }
            //copyTable(newTbl, table);
            for (int z = 0; z < table.getRows().size(); z++) {
                copyTable(newTbl, table.getRow(z), z);
            }

            listNewTable.add(newTbl);
            tableOffset++;
        }

        // 处理第一个表格
        PoiWordAutoTable writeData = (PoiWordAutoTable) list.get(0);
        for (int i = 0;i < table.getRows().size(); i++){
            for (int j = 0; j < table.getRow(i).getTableCells().size(); j++){
                cellSetTextOverride(table.getRow(i).getCell(j), writeData.getCell(i, j));
            }
        }

        // 第二个，第三个... 表格（如果有的话）
        for (int i = 1; i < list.size(); i++){
            writeData = (PoiWordAutoTable) list.get(i);
            XWPFTable xwpfTable = listNewTable.get(i-1);
            for (int x = 0; x < xwpfTable.getRows().size(); x++){
                for (int y = 0; y < xwpfTable.getRow(x).getTableCells().size(); y++){
                    cellSetTextOverride(xwpfTable.getRow(x).getCell(y), writeData.getCell(x, y));
                }
            }
        }

        return tableOffset;
    }

    /**
     * 特殊表格（动态增减整个表格 携带标题和随后文本） 替换
     * @param table
     * @param autoTableRowMap
     * @return
     */
    private static int doAutoTableMax02(XWPFDocument document,
                                        int tableIndex,
                                        XWPFTable table,
                                        Map<String, List<IPoiWordTable>> autoTableRowMap){
        //获取行头的key
        String tableKey = table.getRow(0).getCell(0).getText().trim();
        tableKey = tableKey.substring(2, tableKey.length() - 1);

        List<IPoiWordTable> list = autoTableRowMap.get(tableKey);
        if (null == list || list.isEmpty()){
            table.removeRow(0);
            return -1;
        }
        //先删除模板table的第一行
        table.removeRow(0);

        int tableOffset = 0;    //偏移量

        List<XWPFTable> listNewTable = new ArrayList();   //记录先增加的表格

        //动态整个表格，list有多少个，就生成多少个（由于默认已经有1个了，所以-1）
        for (int i = 0; i < list.size()-1; i++){
            XmlCursor cursor = table.getCTTbl().newCursor();
            cursor.toEndToken();
            while(cursor.toNextToken() != XmlCursor.TokenType.START);
            XWPFParagraph newParagraph = document.insertNewParagraph(cursor);
            newParagraph.createRun().setText("\n", 0);

            cursor = newParagraph.getCTP().newCursor();
            cursor.toEndToken();
            while(cursor.toNextToken() != XmlCursor.TokenType.START);

            XWPFTable newTbl = document.insertNewTbl(cursor);
            while (newTbl != null && !newTbl.getRows().isEmpty()){
                newTbl.removeRow(0);
            }
            // 新表格样式跟原表格样式保持一致
            newTbl.getCTTbl().setTblPr(table.getCTTbl().getTblPr());
            // 外表格边框为 空线
            newTbl.setTopBorder(XWPFTable.XWPFBorderType.NONE, 0, 0, "000000");
            newTbl.setBottomBorder(XWPFTable.XWPFBorderType.NONE, 0, 0, "000000");
            newTbl.setLeftBorder(XWPFTable.XWPFBorderType.NONE, 0, 0, "000000");
            newTbl.setRightBorder(XWPFTable.XWPFBorderType.NONE, 0, 0, "000000");
            //newTbl.getCTTbl().setTblGrid(table.getCTTbl().getTblGrid());

            // 由于有内置表格，所以要对内置表格进行复制才对
            newTbl.insertNewTableRow(0);
            newTbl.getRow(0).addNewTableCell();
            XWPFTableCell masterCell = newTbl.getRow(0).getCell(0);

            masterCell.getParagraphs().get(0);
            cursor = masterCell.getParagraphs().get(0).getCTP().newCursor();
            // 插入 标题段落
            XWPFParagraph masterCellHeaderP = masterCell.insertNewParagraph(cursor);
            copyParagraph(masterCellHeaderP, table.getRow(0).getCell(0).getParagraphs().get(0));
            masterCellHeaderP.createRun().setText("", 0);
            cursor = masterCellHeaderP.getCTP().newCursor();
            cursor.toEndToken();
            while(cursor.toNextToken() != XmlCursor.TokenType.START);
            // cell内创建一个新表格
            XWPFTable masterCellInnerTable = masterCell.insertNewTbl(cursor);
            masterCellInnerTable.getCTTbl().setTblPr(table.getRow(0).getCell(0).getTables().get(0).getCTTbl().getTblPr());
            cursor = masterCellInnerTable.getCTTbl().newCursor();
            cursor.toEndToken();
            while(cursor.toNextToken() != XmlCursor.TokenType.START);
            // cell内在表格下方，再插入一个段落
            copyParagraph(masterCell.insertNewParagraph(cursor),
                    table.getRow(0).getCell(0).getParagraphs().get(1));
            while (masterCell.getParagraphs().size() > 2){
                masterCell.removeParagraph(2);
            }

            // copyTable
            for (int z = 0; z < table.getRow(0).getCell(0).getTables().get(0).getRows().size(); z++) {
                copyTable(masterCellInnerTable, table.getRow(0).getCell(0).getTables().get(0).getRow(z), z);
            }

            listNewTable.add(newTbl);
            tableOffset++;
        }

        // 处理第一个表格
        PWATwithHeaderBottom writeData = (PWATwithHeaderBottom) list.get(0);
        XWPFTable innerTable = table.getRow(0).getCell(0).getTables().get(0);
        //设置标题
        List<XWPFParagraph> listParagraphs = table.getRow(0).getCell(0).getParagraphs();
        if (null == writeData.getTitle()){
            table.getRow(0).getCell(0).removeParagraph(0);
        }
        else {
            paragraphSetText(writeData.getTitle(), listParagraphs.get(0));
        }
        //设置inner表格内容
        for (int i = 0; i < innerTable.getRows().size(); i++){
            for (int j = 0; j < innerTable.getRow(i).getTableCells().size(); j++){
                cellSetTextOverride(innerTable.getRow(i).getCell(j), writeData.getCell(i, j));
            }
        }
        //设置随后文本
        if (null == writeData.getBottom()) {
            table.getRow(0).getCell(0).removeParagraph(listParagraphs.size() - 1);
        }
        else {
            paragraphSetText(writeData.getBottom(),
                    listParagraphs.get(listParagraphs.size() - 1));
        }


        // 处理第二，第三...个表格
        for (int i = 1; i < list.size(); i++){
            writeData = (PWATwithHeaderBottom) list.get(i);
            XWPFTable otherTable = listNewTable.get(i-1);
            innerTable = otherTable.getRow(0).getCell(0).getTables().get(0);
            //设置标题
            listParagraphs = otherTable.getRow(0).getCell(0).getParagraphs();
            if (writeData.getTitle() == null){
                otherTable.getRow(0).getCell(0).removeParagraph(0);
            }
            else {
                paragraphSetText(writeData.getTitle(), listParagraphs.get(0));
            }
            //设置内表格内容
            for (int x = 0; x < innerTable.getRows().size();x++){
                for (int y = 0; y < innerTable.getRow(x).getTableCells().size(); y++){
                    cellSetTextOverride(innerTable.getRow(x).getCell(y),
                            writeData.getCell(x, y));
                }
            }
            //设置随后文本
            if (writeData.getBottom() == null) {
                otherTable.getRow(0).getCell(0).removeParagraph(listParagraphs.size()-1);
            }
            else {
                paragraphSetText(writeData.getBottom(), listParagraphs.get(listParagraphs.size()-1));
            }
//            while (listParagraphs.size() > 2){
//                otherTable.getRow(0).getCell(0).removeParagraph(2);
//            }
        }

        return tableOffset;
    }

    /**
     * 将段落清空，并set进一个新text,保留最后一个Run，以这个Run为样式基础
     * @param text
     * @param paragraph
     */
    private static void paragraphSetText(String text, XWPFParagraph paragraph){
        for (int i = 1; i < paragraph.getRuns().size(); ){
            paragraph.removeRun(1);
        }

        paragraph.getRuns().get(0).setText(text, 0);
    }

    /**
     * cell 覆盖更新文本 (用于普通cell，即cell里面只是显示文本的作用)
     * @param cell
     * @param text
     */
    private static void cellSetTextOverride(XWPFTableCell cell, String text){
        int paraSize = cell.getParagraphs().size();
        if (paraSize == 0){
            cell.addParagraph().createRun().setText(text, 0);
            return;
        }

        for (int i = 1; i < paraSize; ){
            cell.getParagraphs().remove(1);
        }

        XWPFParagraph paragraph = cell.getParagraphs().get(0);

        int runSize = paragraph.getRuns().size();
        if (runSize == 0){
            paragraph.createRun().setText(text, 0);
            return;
        }
        for (int i = 1; i < runSize; i++){
            paragraph.removeRun(0);
        }
        paragraph.getRuns().get(0).setText(text, 0);
    }

    /**
     * 遍历表格，替换 占位符文本
     * @param rows 表格行对象
     * @param textMap 需要替换的信息集合
     */
    private static void eachTable(List<XWPFTableRow> rows ,Map<String, String> textMap){
        //逐行遍历
        for (XWPFTableRow row : rows) {
            List<XWPFTableCell> cells = row.getTableCells();
            for (XWPFTableCell cell : cells) {
                //判断单元格是否需要替换
                if(PoiWordKeyMatchRule.checkText(cell.getText())){
                    List<XWPFParagraph> paragraphs = cell.getParagraphs();
                    for (XWPFParagraph paragraph : paragraphs) {
                        Set<Integer> treeSet = new TreeSet(); //要删除的 Run

                        List<XWPFRun> runs = paragraph.getRuns();
                        for (int i = 0; i < runs.size(); i++){

                            XWPFRun run = runs.get(i);
                            if (PoiWordKeyMatchRule.checkText(run.toString())){
                                // 占位符的 run 必定是完全独立的
                                run.setText(changeValue(run.toString(), textMap), 0);
                                treeSet.add(i-1);
                                treeSet.add(i+1);
                            }
                        }
                        // 移除掉 独立占位符Run
                        int totalRemove = 0;
                        for (Integer removeIndex : treeSet){
                            paragraph.removeRun(removeIndex - totalRemove);
                            totalRemove++;
                        }
                    }
                }
            }
        }
    }

    /**
     * 匹配传入信息集合与模板
     * @param value 模板需要替换的区域
     * @param textMap 传入信息集合
     * @return 模板需要替换区域信息集合对应值
     */
    private static String changeValue(String value, Map<String, String> textMap){
        // value = ${xxx} 因此第0,1 和 最后一个字符，都是非key内容
        String mapKey = value.substring(2, value.length() - 1); //获取key
        String mapValue = textMap.get(mapKey);
        if (null == mapValue){
            return "";  //占位符没有给定替换文本，则默认空字符串替换
        }
        else {
            return mapValue;
        }
    }

    /**
     * table是空，用一段文字代替table的位置
     * @param table
     */
    private static void tableIsNull(XWPFTable table, String text){
        while (!table.getRows().isEmpty()){
            table.removeRow(0);
        }
        if (null != text){
            XmlCursor cursor = table.getCTTbl().newCursor();
            cursor.toEndToken();
            while (cursor.toNextToken() != XmlCursor.TokenType.START) ;
            table.getBody().getXWPFDocument().insertNewParagraph(cursor).createRun().setText(text, 0);
        }
    }

    // 复制Run
    private static void CopyRun(XWPFRun target, XWPFRun source) {
        target.getCTR().setRPr(source.getCTR().getRPr());
        // 设置文本
        target.setText(source.text(), 0);
    }

    // 复制段落
    private static void copyParagraph(XWPFParagraph target, XWPFParagraph source) {
        // 设置段落样式
        target.getCTP().setPPr(source.getCTP().getPPr());
        // 删除目标所有 Run
        for (int pos = 0; pos < target.getRuns().size(); pos++) {
            target.removeRun(pos);
        }
        for (XWPFRun s : source.getRuns()) {
            XWPFRun targetrun = target.createRun();
            CopyRun(targetrun, s);
        }
    }

    // 复制Cell
    private static void copyTableCell(XWPFTableCell target, XWPFTableCell source) {
        // 列属性
        target.getCTTc().setTcPr(source.getCTTc().getTcPr());
        // 删除目标 targetCell 所有段落
        for (int pos = 0; pos < target.getParagraphs().size(); pos++) {
            target.removeParagraph(pos);
        }
        // 添加段落
        for (XWPFParagraph sp : source.getParagraphs()) {
            XWPFParagraph targetP = target.addParagraph();
            copyParagraph(targetP, sp);
        }
    }

    // 复制Row
    private static void copyTableRow(XWPFTableRow target, XWPFTableRow source) {
        // 复制样式
        target.getCtRow().setTrPr(source.getCtRow().getTrPr());

        // 复制单元格
        for (int i = 0; i < target.getTableCells().size(); i++) {
            copyTableCell(target.getCell(i), source.getCell(i));
        }
    }

    // 复制Table
    private static void copyTable(XWPFTable table,XWPFTableRow sourceRow,int rowIndex){
        //判断表格指定位置原来是否有数据
        if (rowIndex < table.getRows().size()){
            table.removeRow(rowIndex);
        }
        //在表格指定位置新增一行
        XWPFTableRow targetRow = table.insertNewTableRow(rowIndex);
        //复制行属性
        targetRow.getCtRow().setTrPr(sourceRow.getCtRow().getTrPr());
        List<XWPFTableCell> cellList = sourceRow.getTableCells();
        if (null == cellList) {
            return;
        }
        //复制列及其属性和内容
        XWPFTableCell targetCell = null;
        for (XWPFTableCell sourceCell : cellList) {
            targetCell = targetRow.addNewTableCell();

            copyTableCell(targetCell, sourceCell);
        }
    }

}
