package com.itheima.handler;

import com.itheima.entity.PoiEntity;
import org.apache.poi.xssf.eventusermodel.XSSFSheetXMLHandler;
import org.apache.poi.xssf.usermodel.XSSFComment;

/**
 * 自定义事件处理器
 */
// 自定义Sheet基于Sax的解析处理器
public class SheetHandler implements XSSFSheetXMLHandler.SheetContentsHandler {

    // 封装实体对象
    private PoiEntity poiEntity;

    /**
     * 当开始解析某一行的时候触发
     *
     * @param i 行索引
     */

    public void startRow(int i) {
        // 实例化对象
        if (i > 0) {
            poiEntity = new PoiEntity();
        }
    }

    /**
     * 当结束解析某一行的时候触发
     *
     * @param i 行索引
     */

    public void endRow(int i) {
        // 使用对象进行业务操作
        System.out.println(poiEntity);
    }

    /**
     * 对行中的每一个表格进行处理
     * cellReference: 单元格名称
     * value：数据
     * xssfComment：批注
     */

    public void cell(String cellReference, String value, XSSFComment xssfComment) {
        // 对对象属性赋值
        if (poiEntity != null) {
            String pix = cellReference.substring(0, 1);
            if (pix.equals("A")) {
                poiEntity.setId(value);

            } else if (pix.equals("B")) {
                poiEntity.setBreast(value);

            } else if (pix.equals("C")) {
                poiEntity.setAdipocytes(value);

            } else if (pix.equals("D")) {
                poiEntity.setNegative(value);

            } else if (pix.equals("E")) {
                poiEntity.setStaining(value);

            } else if (pix.equals("F")) {
                poiEntity.setSupportive(value);

            } else {
            }
        }
    }
}