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
import java.io.File;
import java.util.List;
import javax.swing.JTextArea;

/**
 *
 * @author tangzhichao
 */
public class WordJacobService extends AbstractJacobService {

    public void initApplication() {
        ComThread.InitSTA();
        app = new ActiveXComponent("Word.Application");
        app.setProperty("Visible", new Variant(false));
        documents = app.getProperty("Documents").toDispatch();
    }

    public void check(File paperFile, List<String> answerList, double score, JTextArea console) throws Exception {
        openDoc(paperFile.getAbsolutePath());
        Dispatch tables = Dispatch.get(doc, "Tables").toDispatch();
        Dispatch scoreTable = Dispatch.call(tables, "Item", new Variant(1)).toDispatch();
        Dispatch answerTable = Dispatch.call(tables, "Item", new Variant(2)).toDispatch();
        if (answerTable == null) {
            console.append("没有找到答案表格\n");
            throw new Exception("没有找到答案表格:" + paperFile.getAbsolutePath());
        }
        Dispatch rows = Dispatch.call(answerTable, "Rows").toDispatch();
        Dispatch columns = Dispatch.call(answerTable, "Columns").toDispatch();
        int rowCount = Dispatch.get(rows, "Count").getInt();
        int colCount = Dispatch.get(columns, "Count").getInt();

        double total = 0;
        for (int r = 1; r <= rowCount; r += 2) {
            for (int c = 2; c <= colCount; c++) {
                Dispatch cell = Dispatch.call(answerTable, "Cell", new Variant(r), new Variant(c)).toDispatch();
                Dispatch Range = Dispatch.get(cell, "Range").toDispatch();
                String text = Dispatch.get(Range, "Text").getString().trim();

                //获得题号码
                if (!text.trim().matches("\\d+")) {
                    continue;
                }
                int questionIndex = Integer.parseInt(text.trim());
                Dispatch aTr = Dispatch.call(answerTable, "Cell", new Variant(r + 1), new Variant(c)).toDispatch();
                Range = Dispatch.get(aTr, "Range").toDispatch();
                String stuAnswer = Dispatch.get(Range, "Text").getString().trim().replaceAll("[^A-Za-z]", "").toUpperCase();
                if (answerList.get(questionIndex - 1).equalsIgnoreCase(stuAnswer)) {
                    total += score;
                } else {
                    console.append(String.format("第%d题：正确答案 %s ,学生答案 %s\n", questionIndex, answerList.get(questionIndex - 1), stuAnswer));
                }
            }
        }
        console.append("得分：" + total + "\n");
        //分数写入
        if (scoreTable == null) {
            console.append("无法填入分数\n");
            throw new Exception("无法填入分数:" + paperFile.getAbsolutePath());
        }
        //写入和计算分数
        setScroe(scoreTable, total);
    }

    private void setScroe(Dispatch scoreTable, double total) {
        Dispatch rows = Dispatch.call(scoreTable, "Rows").toDispatch();
        Dispatch columns = Dispatch.call(scoreTable, "Columns").toDispatch();

        Dispatch firstCell = Dispatch.call(scoreTable, "Cell", new Variant(2), new Variant(2)).toDispatch();
        Dispatch Range = Dispatch.get(firstCell, "Range").toDispatch();
        Dispatch.call(Range, "InsertAfter", new Variant(total));

        Dispatch secondCell = Dispatch.call(scoreTable, "Cell", new Variant(2), new Variant(3)).toDispatch();
        Range = Dispatch.get(secondCell, "Range").toDispatch();
        String secendValue = Dispatch.get(Range, "Text").getString().trim();

        Dispatch thirdCell = Dispatch.call(scoreTable, "Cell", new Variant(2), new Variant(4)).toDispatch();
        Range = Dispatch.get(thirdCell, "Range").toDispatch();
        String thirdValue = Dispatch.get(Range, "Text").getString().trim();
        if (secendValue.matches("\\d+") && thirdValue.matches("\\d+")) {
            total += Integer.parseInt(secendValue);
            total += Integer.parseInt(thirdValue);
            //
            Dispatch totalCell = Dispatch.call(scoreTable, "Cell", new Variant(2), new Variant(5)).toDispatch();
            Range = Dispatch.get(totalCell, "Range").toDispatch();
            Dispatch.call(Range, "InsertAfter", new Variant(total));
        }

    }

    public static void main(String[] args) throws Exception {
        try (WordJacobService jacob = new WordJacobService()) {
            jacob.check(new File("F:\\test\\test.doc"), null, 0, new JTextArea());
        }
    }

    public String getTotalScroe(File paperFile, int col) {
        openDoc(paperFile.getAbsolutePath());
        Dispatch tables = Dispatch.get(doc, "Tables").toDispatch();
        Dispatch scoreTable = Dispatch.call(tables, "Item", new Variant(1)).toDispatch();
        Dispatch totalCell = Dispatch.call(scoreTable, "Cell", new Variant(2), new Variant(col)).toDispatch();
        Dispatch Range = Dispatch.get(totalCell, "Range").toDispatch();
        return Dispatch.get(Range, "Text").getString().trim();
    }

}
