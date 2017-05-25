/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package org.alvin.auto.bean;

/**
 *
 * @author tangzhichao
 */
public class ItemBean {

    private String name;
    private int scroe;

    public ItemBean(String name, int scroe) {
        this.name = name;
        this.scroe = scroe;
    }

    public String getName() {
        return name;
    }

    public void setName(String name) {
        this.name = name;
    }

    public int getScroe() {
        return scroe;
    }

    public void setScroe(int scroe) {
        this.scroe = scroe;
    }

}
