package com.dy.apachepoi.poi;

/**
 * 占位符匹配校验器
 *
 * @author HuangDongYang<huangdy @ pvc123.com>
 * Create on 2018/12/1 10:44
 */
public class PoiWordKeyMatchRule {
    /**
     * <p>
     *     在这里改动自己的校验规则
     * </p>
     */

    /** 普通文本替换 */
    protected static final String replace_text = "*${t_*}*";

    /** 富文本 */
    protected static final String full_text = "*{ft_*}*";

    /** 静态表格(文本替换) */
    protected static final String auto_table_static = "*${at_static_*}*";

    /** 动态表格-行 */
    protected  static final String auto_table_row = "*${at_row_*}*";

    /** 动态表格-整个表格 类型01（普通整个表格） */
    protected  static final String auto_table_max01 = "*${at_max01_*}*";

    /** 动态表格-整个表格 类型02（携带标题和跟随文本的 整个表格） */
    protected  static final String auto_table_max02 = "*${at_max02_*}*";

    /** 普通图片替换 */
    protected static final String image_00 = "*${img00_*}*";

    /**
     * 判断是否为静态表格（仅文本填充）
     * @param text
     * @return
     */
    protected static boolean checkStaticTable(String text){
        return isMatch(text, auto_table_static);
    }

    /**
     * 判断是否为特殊表格（动态增减行）
     * @param text
     * @return
     */
    protected static boolean checkAutoTableRow(String text){
        return isMatch(text, auto_table_row);
    }

    /**
     * 判断是否为动态表格 01
     * @param text
     * @return
     */
    protected static boolean checkAutoTableMax01(String text){
        return isMatch(text, auto_table_max01);
    }

    /**
     * 判断是否为 动态表格02
     * @param text
     * @return
     */
    protected static boolean checkAutoTableMax02(String text){
        return isMatch(text, auto_table_max02);
    }

    /**
     * 判断文本中是否 替换文本
     * @param text 文本
     * @return 包含返回true,不包含返回false
     */
    protected static boolean checkText(String text){
        return isMatch(text, replace_text);
    }

    /**
     * 判断是否富文本
     * @param text
     * @return
     */
    protected static boolean checkFullText(String text){
        return isMatch(text, full_text);
    }

    protected static boolean checkImage00(String text){
        return isMatch(text, image_00);
    }

    /**
     * 动态规划实现的 通配符匹配算法
     * 允许 ?  * <br>
     *     ps: 不用indexOf是因为大计算的时候效率不及这套算法，其次使用该算法，可以允许用户在文本中存在不符合规则
     *     的${}存在,这些使用了占位符的原文档文本，不会被替换掉.
     * @param s 匹配字符
     * @param p 匹配规则
     * @author huangdy
     * @return 传入s，是否匹配规则p
     */
    private static boolean isMatch(String s, String p) {
        char[] str = s.toCharArray();
        char[] pattern = p.toCharArray();

        //replace multiple * with one *
        //e.g a**b --> a*b
        int writeIndex = 0;
        boolean isFirst = true;
        for (int i = 0; i < pattern.length; ++i){
            if (pattern[i] == '*'){
                if (isFirst){
                    pattern[writeIndex++] = pattern[i];
                    isFirst = false;
                }
            }
            else {
                pattern[writeIndex++] = pattern[i];
                isFirst = true;
            }
        }
        //code better faster than without this code in 10ms
        if (str.length > 0 && writeIndex > 0 && pattern[0] != '*' && pattern[0] != '?' && str[0] != pattern[0]){
            return false;
        }
        if (str.length > 0 && writeIndex > 0 && pattern[writeIndex-1] != '*' && pattern[writeIndex-1] != '?' && str[str.length-1] != pattern[writeIndex-1]){
            return false;
        }

        //begin the alo.
        boolean T[][] = new boolean[str.length+1][writeIndex+1];
        T[0][0] = true;     //because empty-string == empty-string
        if (writeIndex > 0 && pattern[0] == '*'){
            //because empty-string == *
            T[0][1] = true;
        }

        //DP deal with this problem
        for (int i = 1; i < T.length; ++i){
            for (int j = 1; j < T[0].length; ++j){
                if (pattern[j-1] == '?' || str[i-1] == pattern[j-1]){
                    T[i][j] = T[i-1][j-1];
                }
                else if (pattern[j-1] == '*'){
                    T[i][j] = T[i-1][j] || T[i][j-1];
                }
            }
        }

        return T[str.length][writeIndex];   //last one is answer
    }
}
