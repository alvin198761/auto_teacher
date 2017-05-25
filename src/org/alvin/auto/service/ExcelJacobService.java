/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package org.alvin.auto.service;

import com.jacob.activeX.ActiveXComponent;
import com.jacob.com.ComThread;
import com.jacob.com.Dispatch;
import com.jacob.com.Variant;
import java.io.IOException;
import java.util.HashMap;
import java.util.Map;
import java.util.logging.Level;
import java.util.logging.Logger;

/**
 *
 * @author tangzhichao
 */
public class ExcelJacobService extends AbstractJacobService {

    private static final String COLUMNS = "ABCDEFGHIGKLMNOPQRSTUVWXYZ";

    @Override
    public void initApplication() {
        ComThread.InitSTA();
        app = new ActiveXComponent("Excel.Application");
        app.setProperty("Visible", new Variant(false));
    }

    public void valuation(String baseDir, int startRow, int nameCol, int scroeCol, Map<String, String> map, String exclePath) {
        openDoc(exclePath);
        Dispatch sheets = Dispatch.get(documents, "Sheets").toDispatch();
        Dispatch.call(documents, "Activate");
        Dispatch sheet = Dispatch.invoke(sheets, "Item", Dispatch.Get, new Object[]{"成绩表"}, new int[0]).getDispatch();
        Object[] keys = map.keySet().toArray();
        for (int i = 0; i < keys.length; i++) {
            String cName = COLUMNS.charAt(nameCol - 1) + "" + startRow;
            //姓名
            Dispatch cell = Dispatch.invoke(sheet, "Range", Dispatch.Get, new Object[]{cName}, new int[1]).toDispatch();
            Dispatch.put(cell, "Value", keys[i].toString());
            //分数
            cName = COLUMNS.charAt(scroeCol - 1) + "" + startRow;
            cell = Dispatch.invoke(sheet, "Range", Dispatch.Get, new Object[]{cName}, new int[1]).toDispatch();
            Dispatch.put(cell, "Value", map.get(keys[i]));
            startRow++;
        }
    }

    @Override
    public void openDoc(String docPath) {
        closeDoc();
        Dispatch workbooks = app.getProperty("Workbooks").toDispatch();
        documents = Dispatch.invoke(workbooks, "Open", Dispatch.Method, new Object[]{docPath}, new int[0]).getDispatch();
    }

    @Override
    public void closeDoc() {
        if (documents != null) {
            Dispatch.call(documents, "Save");
            Dispatch.call(documents, "Close", new Variant(true));
            documents = null;
        }
    }

    public static void main(String[] args) {
        try (ExcelJacobService jacob = new ExcelJacobService()) {
            HashMap<String, String> map = new HashMap();
            map.put("aaa", "1000");
            jacob.valuation("", 5, 2, 3, map, "F:/test.xls");
        } catch (IOException ex) {
            Logger.getLogger(ExcelJacobService.class.getName()).log(Level.SEVERE, null, ex);
        }
    }

}
