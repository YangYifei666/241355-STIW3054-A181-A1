/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package com.mycompany.rtassignment1;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.jsoup.Jsoup;
import org.jsoup.nodes.Document;
import org.jsoup.nodes.Element;
import org.jsoup.select.Elements;

/**
 *
 * @author lenovo
 */

public class Try {
    public static void main(String[] args) throws FileNotFoundException, IOException {
        Document doc;
        Element table = null;
        Elements tr;
        HSSFWorkbook wb = new HSSFWorkbook();
        HSSFSheet sheet = wb.createSheet("Trivia");
        doc = Jsoup.connect("https://ms.wikipedia.org/wiki/Malaysia").get();
        table = doc.select("h2:has(span#Trivia)").next("table").first();
        tr = table.child(0).getElementsByTag("tr");
        for(Element row :tr){
            sheet.createRow(row.elementSiblingIndex());
            Elements cells = row.children();
            for(Element cell : cells){
                sheet.getRow(row.elementSiblingIndex()).createCell(cell.elementSiblingIndex()).setCellValue(cell.text());
            }
        }
        FileOutputStream output=new FileOutputStream("C:\\Users\\lenovo\\Desktop\\output.xls");
        wb.write(output); 
        output.flush();
    }
}
