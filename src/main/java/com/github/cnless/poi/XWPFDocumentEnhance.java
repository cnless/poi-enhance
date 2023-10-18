package com.github.cnless.poi;

import com.github.cnless.poi.sdt.StructuredDocumentTagEnhance;
import org.apache.poi.ooxml.POIXMLProperties;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.xwpf.usermodel.*;
import org.apache.xmlbeans.XmlCursor;
import org.apache.xmlbeans.XmlObject;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.*;

import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.*;


public class XWPFDocumentEnhance extends XWPFDocument {
    private final List<StructuredDocumentTagEnhance> sdtList = new ArrayList<>();
    public XWPFDocumentEnhance(InputStream is) throws IOException {
        super(is);
        sdtRead();
    }
    public XWPFDocumentEnhance(OPCPackage pkg) throws IOException {
        super(pkg);
        sdtRead();
    }
    public List<StructuredDocumentTagEnhance> getSdtList() {
        return sdtList;
    }

    public void writeFile(FileOutputStream outputStream) throws IOException {
        POIXMLProperties.CoreProperties coreProperties = getProperties().getCoreProperties();
        coreProperties.setCreator("poi-enhance");
        write(outputStream);
    }

    private void sdtRead(){
        XmlCursor docCursor = getDocument().newCursor();
        docCursor.selectPath("./*");
        while (docCursor.toNextSelection()) {
            XmlObject o = docCursor.getObject();
            if (o instanceof CTBody) {
                try (XmlCursor bodyCursor = o.newCursor()) {
                    bodyCursor.selectPath("./*");
                    while (bodyCursor.toNextSelection()) {
                        XmlObject xmlObject = bodyCursor.getObject();
                        if (xmlObject instanceof CTP) {
                            readInParagraph((CTP) xmlObject);
                        } else if (xmlObject instanceof CTTbl) {
                            readInTable((CTTbl) xmlObject);
                        } else if (xmlObject instanceof CTSdtBlock) {
                            sdtList.add(new StructuredDocumentTagEnhance((CTSdtBlock) xmlObject,this));
                        }
                    }
                }
            }
        }
        docCursor.close();
    }
    private void readInParagraph(CTP ctp){
        try (XmlCursor pCursor = ctp.newCursor()) {
            pCursor.selectPath("./*");
            while (pCursor.toNextSelection()) {
                XmlObject xmlObject = pCursor.getObject();
                if (xmlObject instanceof CTSdtBlock) {
                    sdtList.add(new StructuredDocumentTagEnhance((CTSdtBlock) xmlObject,this));
                } else if (xmlObject instanceof CTSdtRun) {
                    sdtList.add(new StructuredDocumentTagEnhance((CTSdtRun) xmlObject,this));
                }
            }
        }
    }
    private void readInTable(CTTbl ctTbl){
        try(XmlCursor tblCursor = ctTbl.newCursor()) {
            tblCursor.selectPath("./*");
            while (tblCursor.toNextSelection()){
                XmlObject xmlObject = tblCursor.getObject();
                if (xmlObject instanceof CTSdtBlock) {
                    sdtList.add(new StructuredDocumentTagEnhance((CTSdtBlock) xmlObject,this));
                } else if (xmlObject instanceof CTRow) {
                    readInRow((CTRow) xmlObject);
                }
            }
        }
    }

    private void readInRow(CTRow ctRow){
        try (XmlCursor rowCursor = ctRow.newCursor()){
            rowCursor.selectPath("./*");
            while (rowCursor.toNextSelection()){
                XmlObject xmlObject = rowCursor.getObject();
                if (xmlObject instanceof CTSdtBlock) {
                    sdtList.add(new StructuredDocumentTagEnhance((CTSdtBlock) xmlObject,this));
                } else if (xmlObject instanceof CTTbl) {
                    readInTable((CTTbl) xmlObject);
                } else if (xmlObject instanceof CTP) {
                    readInParagraph((CTP) xmlObject);
                } else if (xmlObject instanceof CTTc) {
                    readInCell((CTTc) xmlObject);
                } else if (xmlObject instanceof CTSdtCell) {
                    sdtList.add(new StructuredDocumentTagEnhance((CTSdtCell) xmlObject,this));
                }
            }
        }
    }
    private void readInCell(CTTc ctTc){
        try (XmlCursor cellCursor = ctTc.newCursor()){
            cellCursor.selectPath("./*");
            while (cellCursor.toNextSelection()){
                XmlObject xmlObject = cellCursor.getObject();
                if (xmlObject instanceof CTP){
                    readInParagraph((CTP) xmlObject);
                } else if (xmlObject instanceof CTTbl) {
                    readInTable((CTTbl) xmlObject);
                } else if (xmlObject instanceof CTSdtBlock) {
                    sdtList.add(new StructuredDocumentTagEnhance((CTSdtBlock) xmlObject,this));
                } else if (xmlObject instanceof CTSdtRun) {
                    sdtList.add(new StructuredDocumentTagEnhance((CTSdtRun) xmlObject,this));
                }
            }
        }
    }
}
