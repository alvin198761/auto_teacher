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
import java.io.Closeable;
import java.io.IOException;

/**
 *
 * @author tangzhichao
 */
public abstract class AbstractJacobService implements Closeable {

    protected Dispatch doc = null;
    protected ActiveXComponent app = null;
    protected Dispatch documents = null;

    public AbstractJacobService() {
        initApplication();
    }

    public abstract void initApplication();

    public void openDoc(String docPath) {
        closeDoc();
        doc = Dispatch.call(documents, "Open", docPath).toDispatch();
    }

    public void closeDoc() {
        if (doc != null) {
            Dispatch.call(doc, "Save");
            Dispatch.call(doc, "Close", new Variant(true));
            doc = null;
        }
    }

    @Override
    public void close() throws IOException {
        closeDoc();
        if (app != null) {
            Dispatch.call(app, "Quit");
            app = null;
        }
        documents = null;
        ComThread.Release();
        System.gc();
    }

}
