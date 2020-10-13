package com.vike0906.be.help;
import java.lang.reflect.Field;

/**
 * @author: lsl
 * @createDate: 2020/10/12
 */
public class FieldSort {

    private Field field;

    private int index;

    private String title;

    public String getTitle() {
        return title;
    }

    public void setTitle(String title) {
        this.title = title;
    }

    public FieldSort(){

    }
    public FieldSort(Field field, int index) {
        this.field = field;
        this.index = index;
    }

    public FieldSort(Field field, int index, String title) {
        this.field = field;
        this.index = index;
        this.title = title;
    }

    public Field getField() {
        return field;
    }

    public void setField(Field field) {
        this.field = field;
    }

    public int getIndex() {
        return index;
    }

    public void setIndex(int index) {
        this.index = index;
    }
}
