package com.github.cnless.poi.sdt;

import com.github.cnless.poi.XWPFDocumentEnhance;
import com.github.cnless.poi.util.ImageUtils;
import org.apache.poi.common.usermodel.PictureType;
import org.apache.poi.ooxml.POIXMLDocumentPart;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.util.StringUtil;
import org.apache.poi.xwpf.usermodel.XWPFPictureData;
import org.apache.xmlbeans.XmlCursor;
import org.apache.xmlbeans.XmlObject;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.*;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.impl.CTSdtPrImpl;

import javax.xml.namespace.QName;
import java.io.InputStream;
import java.text.SimpleDateFormat;
import java.util.Calendar;
import java.util.Date;
import java.util.List;
import java.util.Objects;


public class StructuredDocumentTagEnhance {
    private CTSdtRun ctSdtRun;
    private CTSdtBlock ctSdtBlock;
    private CTSdtCell ctSdtCell;
    private CTSdtPr ctSdtPr;
    private CTSdtEndPr ctSdtEndPr;
    private StructuredDocumentTagContentEnhance contentEnhance;
    private XWPFDocumentEnhance documentEnhance;
    private int sdtType;
    private XmlObject checked;
    private XmlObject checkedState;
    private XmlObject uncheckedState;

    public int getSdtType() {
        return sdtType;
    }

    public StructuredDocumentTagEnhance(CTSdtRun sdtRun,XWPFDocumentEnhance documentEnhance) {
        this.documentEnhance = documentEnhance;
        this.ctSdtRun = sdtRun;
        this.ctSdtPr = sdtRun.getSdtPr();
        matchSdtType();
        this.ctSdtEndPr = sdtRun.getSdtEndPr();
        this.contentEnhance = new StructuredDocumentTagContentEnhance(sdtRun.getSdtContent());
    }

    public StructuredDocumentTagEnhance(CTSdtBlock block,XWPFDocumentEnhance documentEnhance) {
        this.documentEnhance = documentEnhance;
        this.ctSdtBlock = block;
        this.ctSdtPr = block.getSdtPr();
        matchSdtType();
        this.ctSdtEndPr = block.getSdtEndPr();
        this.contentEnhance = new StructuredDocumentTagContentEnhance(block.getSdtContent());
    }

    public StructuredDocumentTagEnhance(CTSdtCell sdtCell,XWPFDocumentEnhance documentEnhance) {
        this.documentEnhance = documentEnhance;
        this.ctSdtCell = sdtCell;
        this.ctSdtPr = sdtCell.getSdtPr();
        matchSdtType();
        this.ctSdtEndPr = sdtCell.getSdtEndPr();
        this.contentEnhance = new StructuredDocumentTagContentEnhance(sdtCell.getSdtContent());
    }

    public void matchSdtType(){
        if (ctSdtPr.isSetText()) {
            sdtType = SdtType.PLAIN_TEXT;
        }else if (ctSdtPr.isSetPicture()){
            sdtType = SdtType.PICTURE;
        } else if (ctSdtPr.isSetComboBox()) {
            sdtType = SdtType.COMBO_BOX;
        } else if (ctSdtPr.isSetDropDownList()) {
            sdtType = SdtType.DROP_DOWN_LIST;
        } else if (ctSdtPr.isSetDate()) {
            sdtType = SdtType.DATE;
        }else if (isSetCheckBox((CTSdtPrImpl) ctSdtPr)){
            sdtType = SdtType.CHECKBOX;
            String declareNameSpaces = "declare namespace w14='http://schemas.microsoft.com/office/word/2010/wordml'";
            this.checked = ctSdtPr.selectPath(declareNameSpaces + ".//w14:checkbox/w14:checked")[0];
            this.checkedState = ctSdtPr.selectPath(declareNameSpaces + ".//w14:checkbox/w14:checkedState")[0];
            this.uncheckedState = ctSdtPr.selectPath(declareNameSpaces + ".//w14:checkbox/w14:uncheckedState")[0];
        } else if (isSetRepeatingSectionItem((CTSdtPrImpl) ctSdtPr)) {
            sdtType = SdtType.REPEATING_SECTION_ITEM;
        } else if (isSetRepeatingSection((CTSdtPrImpl) ctSdtPr)) {
            sdtType = SdtType.REPEATING_SECTION;
        } else {
            sdtType = SdtType.RICH_TEXT;
        }
    }
    private boolean isSetCheckBox(CTSdtPrImpl ctSdtPrImpl){
        synchronized (ctSdtPrImpl.monitor()) {
            return ctSdtPrImpl.get_store().count_elements(new QName("http://schemas.microsoft.com/office/word/2010/wordml", "checkbox")) != 0;
        }
    }

    private boolean isSetRepeatingSectionItem(CTSdtPrImpl ctSdtPrImpl){
        synchronized (ctSdtPrImpl.monitor()) {
            return ctSdtPrImpl.get_store().count_elements(new QName("http://schemas.microsoft.com/office/word/2012/wordml", "repeatingSectionItem")) != 0;
        }
    }

    private boolean isSetRepeatingSection(CTSdtPrImpl ctSdtPrImpl){
        synchronized (ctSdtPrImpl.monitor()) {
            return ctSdtPrImpl.get_store().count_elements(new QName("http://schemas.microsoft.com/office/word/2012/wordml", "repeatingSection")) != 0;
        }
    }

    public String getTextValue(){
        return contentEnhance.getText();
    }

    public void setRichText(String value){
        contentEnhance.setRichText(value);
    }

    public void setPlainText(String value){
        contentEnhance.setPlainText(value);
    }

    public CTSdtDropDownList getCTSdtDropDownList (){
        return ctSdtPr.getDropDownList();
    }

    public void removeAllCTSdtListItem (){
        CTSdtDropDownList ctSdtDropDownList = getCTSdtDropDownList();
        int size = ctSdtDropDownList.sizeOfListItemArray();
        for (int i = size-1; i >=0; i--) {
            ctSdtDropDownList.removeListItem(i);
        }
    }

    public void removeCTSdtListItem(CTSdtListItem ctSdtListItem){
        if (Objects.isNull(ctSdtListItem)) return;
        CTSdtDropDownList ctSdtDropDownList = getCTSdtDropDownList();
        List<CTSdtListItem> listItemList = ctSdtDropDownList.getListItemList();
        int index = 0;
        for (CTSdtListItem sdtListItem : listItemList) {
            if (sdtListItem.equals(ctSdtListItem)){
                break;
            }
            index++;
        }
        ctSdtDropDownList.removeListItem(index);
    }

    public CTSdtListItem addNewCTSdtListItem(String displayText,String valueText){
        CTSdtListItem newListItem = getCTSdtDropDownList().addNewListItem();
        newListItem.setDisplayText(displayText);
        newListItem.setValue(valueText);
        return newListItem;
    }

    public CTSdtListItem addNewCTSdtListItem(String displayText,String valueText,int index){
        CTSdtListItem newListItem = getCTSdtDropDownList().insertNewListItem(index);
        newListItem.setDisplayText(displayText);
        newListItem.setValue(valueText);
        return newListItem;
    }

    public void setSelectedCTSdtListItem(CTSdtListItem selectedItem){
        setPlainText(selectedItem.getDisplayText());
    }

    public CTSdtListItem getSelectedCTSdtListItem(){
        CTSdtDropDownList ctSdtDropDownList = getCTSdtDropDownList();
        String displayValue = getTextValue();
        for (CTSdtListItem ctSdtListItem : ctSdtDropDownList.getListItemList()) {
            if (ctSdtListItem.getDisplayText().equals(displayValue)) return ctSdtListItem;
        }
        return null;
    }

    public byte[] getPictureData(){
        String pictureRelationId = contentEnhance.getPictureRelationId();
        if (Objects.nonNull(pictureRelationId)){
            POIXMLDocumentPart relatedPart = documentEnhance.getRelationById(pictureRelationId);
            if (Objects.nonNull(relatedPart)){
                if (relatedPart instanceof XWPFPictureData) {
                    return ((XWPFPictureData) relatedPart).getData();
                }
            }
        }
        return null;
    }
    public void modifyPictureData(byte[] data){
        if (Objects.isNull(data)) return;
        PictureType pictureType = ImageUtils.suggestFileType(data);
        try {
            String relationId = documentEnhance.addPictureData(data, pictureType);
            contentEnhance.setPictureReference(relationId);
        } catch (InvalidFormatException e) {
            throw new RuntimeException(e);
        }
    }

    public void modifyPictureData(String url){
        if (StringUtil.isBlank(url)) return;
        byte[] data = ImageUtils.toByteArray(url);
        PictureType pictureType = ImageUtils.suggestFileType(data);
        try {
            String relationId = documentEnhance.addPictureData(data, pictureType);
            contentEnhance.setPictureReference(relationId);
        } catch (InvalidFormatException e) {
            throw new RuntimeException(e);
        }
    }

    public void modifyPictureData(InputStream inputStream){
        if (Objects.isNull(inputStream)) return;
        byte[] data = ImageUtils.toByteArray(inputStream);
        PictureType pictureType = ImageUtils.suggestFileType(data);
        try {
            String relationId = documentEnhance.addPictureData(data, pictureType);
            contentEnhance.setPictureReference(relationId);
        } catch (InvalidFormatException e) {
            throw new RuntimeException(e);
        }
    }

    public Date getFullDate(){
        Calendar fullDate = ctSdtPr.getDate().getFullDate();
        return fullDate.getTime();
    }

    public void modifyDate(Date date){
        if (Objects.isNull(date)) return;
        String formatString = ctSdtPr.getDate().getDateFormat().getVal();
        SimpleDateFormat dateFormat = new SimpleDateFormat(formatString);
        contentEnhance.setPlainText(dateFormat.format(date));
    }

    public boolean getChecked(){
        String val;
        try (XmlCursor xmlCursor = checked.newCursor()) {
            val = xmlCursor.getAttributeText(new QName("http://schemas.microsoft.com/office/word/2010/wordml", "val", "w14"));
        }
        return "1".equals(val) || "true".equals(val);
    }

    public void setChecked(boolean checked){
        try (XmlCursor checkedCursor = this.checked.newCursor();
             XmlCursor checkedStateCursor = this.checkedState.newCursor();
             XmlCursor uncheckedStateCursor = this.uncheckedState.newCursor()) {
            String val = (checked) ? "1" : "0";
            checkedCursor.setAttributeText(new QName("http://schemas.microsoft.com/office/word/2010/wordml", "val", "w14"), val);
            String checkedStateVal = checkedStateCursor.getAttributeText(new QName("http://schemas.microsoft.com/office/word/2010/wordml", "val", "w14"));
            String uncheckedStateVal = uncheckedStateCursor.getAttributeText(new QName("http://schemas.microsoft.com/office/word/2010/wordml", "val", "w14"));
            String content = checked?convertToUnicode(checkedStateVal):convertToUnicode(uncheckedStateVal);
            this.contentEnhance.setCheckboxText(content);
        }
    }
    private String convertToUnicode(String input) {
        int i = Integer.parseInt(input, 16);
        return String.valueOf((char) i);
    }

}
