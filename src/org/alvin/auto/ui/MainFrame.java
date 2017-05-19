/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package org.alvin.auto.ui;

import java.io.File;
import java.io.IOException;
import java.io.OutputStream;
import java.nio.file.Files;
import java.nio.file.Paths;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.LinkedList;
import java.util.List;
import java.util.logging.Level;
import java.util.logging.Logger;
import javax.swing.JFileChooser;
import javax.swing.JOptionPane;

/**
 *
 * @author tangzhichao
 */
public class MainFrame extends javax.swing.JFrame {

    /**
     * Creates new form MainFrame
     */
    public MainFrame() {
        initComponents();
        setTitle("选择题自动阅卷工具");
        setResizable(false);
        setLocationRelativeTo(null);
    }

    /**
     * This method is called from within the constructor to initialize the form.
     * WARNING: Do NOT modify this code. The content of this method is always
     * regenerated by the Form Editor.
     */
    @SuppressWarnings("unchecked")
    // <editor-fold defaultstate="collapsed" desc="Generated Code">//GEN-BEGIN:initComponents
    private void initComponents() {

        jLabel1 = new javax.swing.JLabel();
        jScrollPane1 = new javax.swing.JScrollPane();
        anwser = new javax.swing.JTextArea();
        jLabel2 = new javax.swing.JLabel();
        score = new javax.swing.JTextField();
        jLabel3 = new javax.swing.JLabel();
        paperDir = new javax.swing.JTextField();
        jButton1 = new javax.swing.JButton();
        jTabbedPane1 = new javax.swing.JTabbedPane();
        jScrollPane2 = new javax.swing.JScrollPane();
        console = new javax.swing.JTextArea();
        jButton2 = new javax.swing.JButton();
        jCheckBox1 = new javax.swing.JCheckBox();

        setDefaultCloseOperation(javax.swing.WindowConstants.EXIT_ON_CLOSE);

        jLabel1.setText("正确答案：");

        anwser.setColumns(20);
        anwser.setRows(5);
        jScrollPane1.setViewportView(anwser);

        jLabel2.setText("单题分数：");

        score.setText("2");

        jLabel3.setText("试卷目录：");

        paperDir.setEditable(false);

        jButton1.setText("浏览");
        jButton1.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                jButton1ActionPerformed(evt);
            }
        });

        console.setColumns(20);
        console.setRows(5);
        jScrollPane2.setViewportView(console);

        jTabbedPane1.addTab("控制台", jScrollPane2);

        jButton2.setText("开始阅卷");
        jButton2.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                jButton2ActionPerformed(evt);
            }
        });

        jCheckBox1.setSelected(true);
        jCheckBox1.setText("自动保存控制台信息");

        javax.swing.GroupLayout layout = new javax.swing.GroupLayout(getContentPane());
        getContentPane().setLayout(layout);
        layout.setHorizontalGroup(
            layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addGroup(layout.createSequentialGroup()
                .addGap(42, 42, 42)
                .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING, false)
                    .addComponent(jTabbedPane1)
                    .addGroup(layout.createSequentialGroup()
                        .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                            .addComponent(jLabel3)
                            .addComponent(jLabel2)
                            .addComponent(jLabel1))
                        .addGap(12, 12, 12)
                        .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING, false)
                            .addGroup(layout.createSequentialGroup()
                                .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING, false)
                                    .addComponent(paperDir, javax.swing.GroupLayout.PREFERRED_SIZE, 429, javax.swing.GroupLayout.PREFERRED_SIZE)
                                    .addGroup(layout.createSequentialGroup()
                                        .addComponent(score, javax.swing.GroupLayout.PREFERRED_SIZE, 266, javax.swing.GroupLayout.PREFERRED_SIZE)
                                        .addGap(27, 27, 27)
                                        .addComponent(jCheckBox1, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE)))
                                .addGap(18, 18, 18)
                                .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                                    .addComponent(jButton1)
                                    .addComponent(jButton2)))
                            .addComponent(jScrollPane1))))
                .addContainerGap(javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE))
        );

        layout.linkSize(javax.swing.SwingConstants.HORIZONTAL, new java.awt.Component[] {jButton1, jButton2});

        layout.setVerticalGroup(
            layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addGroup(layout.createSequentialGroup()
                .addGap(11, 11, 11)
                .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.BASELINE)
                    .addComponent(jLabel3)
                    .addComponent(paperDir, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE)
                    .addComponent(jButton1))
                .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
                .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.BASELINE)
                    .addComponent(score, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE)
                    .addComponent(jLabel2)
                    .addComponent(jButton2)
                    .addComponent(jCheckBox1))
                .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.UNRELATED)
                .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                    .addComponent(jScrollPane1, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE)
                    .addComponent(jLabel1))
                .addGap(29, 29, 29)
                .addComponent(jTabbedPane1, javax.swing.GroupLayout.DEFAULT_SIZE, 280, Short.MAX_VALUE)
                .addGap(29, 29, 29))
        );

        layout.linkSize(javax.swing.SwingConstants.VERTICAL, new java.awt.Component[] {jButton1, jButton2, jLabel1, jLabel2, jLabel3, paperDir, score});

        pack();
    }// </editor-fold>//GEN-END:initComponents

    private void jButton1ActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_jButton1ActionPerformed
        // TODO add your handling code here:
        JFileChooser jfc = new JFileChooser();
        jfc.setDialogType(JFileChooser.OPEN_DIALOG);
        jfc.setFileSelectionMode(JFileChooser.DIRECTORIES_ONLY);
        jfc.showOpenDialog(this);
        File f = jfc.getSelectedFile();
        if (f != null) {
            this.paperDir.setText(f.getAbsolutePath());
        }
    }//GEN-LAST:event_jButton1ActionPerformed

    private void jButton2ActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_jButton2ActionPerformed
        // TODO add your handling code here:
        int res = JOptionPane.showConfirmDialog(null, "1.请先备份需要检查的文件\n"
                + "2.确认所有的 word 程序已经关闭"
                + "\n3.过程中也不得打开word"
                + "\n4.以上三点都确认以后，点击【确定】", "提示", JOptionPane.OK_CANCEL_OPTION);
        if (res != JOptionPane.OK_OPTION) {
            return;
        }
        if (this.paperDir.getText().trim().isEmpty()) {
            JOptionPane.showMessageDialog(null, "请选择试卷目录！");
            this.paperDir.requestFocus();
            return;
        }
        if (!this.score.getText().trim().matches("\\d+")) {
            JOptionPane.showMessageDialog(null, "选择题分数请输入正确！");
            this.score.requestFocus();
            return;
        }
        if (this.anwser.getText().trim().isEmpty()) {
            JOptionPane.showMessageDialog(null, "参考答案不能为空！");
            this.score.requestFocus();
            return;
        }
        this.console.setText("");
        new Thread() {
            @Override
            public void run() {
                jButton2.setEnabled(false);
                doRun();
                if (jCheckBox1.isSelected()) {
                    try {
                        saveConsole(console.getText());
                    } catch (IOException ex) {
                        Logger.getLogger(MainFrame.class.getName()).log(Level.SEVERE, null, ex);
                    }
                }
                jButton2.setEnabled(true);
            }

        }.start();

    }//GEN-LAST:event_jButton2ActionPerformed

    /**
     * @param args the command line arguments
     */
    public static void main(String args[]) {
        /* Set the Nimbus look and feel */
        //<editor-fold defaultstate="collapsed" desc=" Look and feel setting code (optional) ">
        /* If Nimbus (introduced in Java SE 6) is not available, stay with the default look and feel.
         * For details see http://download.oracle.com/javase/tutorial/uiswing/lookandfeel/plaf.html 
         */
        try {
            for (javax.swing.UIManager.LookAndFeelInfo info : javax.swing.UIManager.getInstalledLookAndFeels()) {
                if ("Nimbus".equals(info.getName())) {
                    javax.swing.UIManager.setLookAndFeel(info.getClassName());
                    break;
                }
            }
        } catch (ClassNotFoundException ex) {
            java.util.logging.Logger.getLogger(MainFrame.class.getName()).log(java.util.logging.Level.SEVERE, null, ex);
        } catch (InstantiationException ex) {
            java.util.logging.Logger.getLogger(MainFrame.class.getName()).log(java.util.logging.Level.SEVERE, null, ex);
        } catch (IllegalAccessException ex) {
            java.util.logging.Logger.getLogger(MainFrame.class.getName()).log(java.util.logging.Level.SEVERE, null, ex);
        } catch (javax.swing.UnsupportedLookAndFeelException ex) {
            java.util.logging.Logger.getLogger(MainFrame.class.getName()).log(java.util.logging.Level.SEVERE, null, ex);
        }
        //</editor-fold>

        /* Create and display the form */
        java.awt.EventQueue.invokeLater(new Runnable() {
            public void run() {
                new MainFrame().setVisible(true);
            }
        });
    }

    // Variables declaration - do not modify//GEN-BEGIN:variables
    private javax.swing.JTextArea anwser;
    private javax.swing.JTextArea console;
    private javax.swing.JButton jButton1;
    private javax.swing.JButton jButton2;
    private javax.swing.JCheckBox jCheckBox1;
    private javax.swing.JLabel jLabel1;
    private javax.swing.JLabel jLabel2;
    private javax.swing.JLabel jLabel3;
    private javax.swing.JScrollPane jScrollPane1;
    private javax.swing.JScrollPane jScrollPane2;
    private javax.swing.JTabbedPane jTabbedPane1;
    private javax.swing.JTextField paperDir;
    private javax.swing.JTextField score;
    // End of variables declaration//GEN-END:variables
    private void saveConsole(String text) throws IOException {
        String date = new SimpleDateFormat("yyyy-MM-dd-HH-mm-ss").format(new Date());
        try (OutputStream out = Files.newOutputStream(Paths.get(date.concat(".txt")))) {
            out.write(text.getBytes());
        }
    }

    private void doRun() {
        String[] answers = this.anwser.getText().trim().split("\t|\\s|[^A-Za-z]");
        List<String> answerList = new ArrayList<>();
        int score = Integer.parseInt(this.score.getText().trim());
        for (String s : answers) {
            answerList.add(s.replaceAll("[^A-Za-z]", "").toUpperCase());
        }
        String text = String.join(",", answerList);
        this.console.append("正确答案：\n");
        this.console.append(text + "\n");
        this.console.append("开始阅卷\n");
        File f = new File(this.paperDir.getText().trim());
        if (!f.isDirectory()) {
            JOptionPane.showMessageDialog(null, "您选择的不是目录！");
            return;
        }
        doAutoRun(f, answerList, score);

    }

    private void doAutoRun(File f, List<String> answerList, int score) {
        LinkedList<File> list = new LinkedList<File>();
        list.add(f);
        while (!list.isEmpty()) {
            File tmpF = list.removeFirst();
            if (tmpF.isDirectory()) {
                for (File subF : tmpF.listFiles()) {
                    list.add(subF);
                }
                continue;
            }
            if (tmpF.getName().toLowerCase().endsWith(".doc") || tmpF.getName().toLowerCase().endsWith(".docx")) {
                try {
                    check(tmpF, answerList, score);
                    this.console.append("\n----------------------------分割线----------------------------\n");
                } catch (Exception ex) {
                    Logger.getLogger(MainFrame.class.getName()).log(Level.SEVERE, null, ex);
                }
            }
        }
    }

    private void check(File paperFile, List<String> answerList, int score) throws Exception {
        this.console.append("检查试卷：" + paperFile.getAbsolutePath() + "\n");
        try (JacobService jacob = new JacobService()) {
            jacob.check(paperFile, answerList, score, this.console);
        }

    }

}
