package io.github.cnless.poi.sdt;

import org.apache.poi.ooxml.POIXMLException;
import org.apache.xmlbeans.XmlException;
import org.apache.xmlbeans.XmlObject;
import org.apache.xmlbeans.impl.values.XmlAnyTypeImpl;
import org.openxmlformats.schemas.drawingml.x2006.main.CTBlipFillProperties;
import org.openxmlformats.schemas.drawingml.x2006.picture.CTPicture;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.*;

import java.util.ArrayList;
import java.util.List;

import static org.apache.poi.ooxml.POIXMLTypeLoader.DEFAULT_XML_OPTIONS;

public class StructuredDocumentTagContentEnhance {
    private CTSdtContentBlock ctSdtContentBlock;
    private CTSdtContentRun ctSdtContentRun;
    private CTSdtContentCell ctSdtContentCell;
    private int levelType;

    public StructuredDocumentTagContentEnhance(CTSdtContentBlock ctSdtContentBlock) {
        this.ctSdtContentBlock = ctSdtContentBlock;
        this.levelType = LevelType.SDT_BLOCK;
    }

    public StructuredDocumentTagContentEnhance(CTSdtContentRun ctSdtContentRun) {
        this.ctSdtContentRun = ctSdtContentRun;
        this.levelType = LevelType.SDT_RUN;
    }

    public StructuredDocumentTagContentEnhance(CTSdtContentCell ctSdtContentCell) {
        this.ctSdtContentCell = ctSdtContentCell;
        this.levelType = LevelType.SDT_CELL;
    }
    protected String getText(){
        StringBuilder value = new StringBuilder();
        int flag = 0;
        if (LevelType.SDT_BLOCK == levelType){
            CTP[] pArray = ctSdtContentBlock.getPArray();
            for (CTP ctp : pArray) {
                CTR[] rArray = ctp.getRArray();
                for (CTR ctr : rArray) {
                    for (CTText ctText : ctr.getTArray()) {
                        value.append(ctText.getStringValue());
                    }
                }
                flag++;
                if (flag == pArray.length) break;
                value.append("\n");
            }
        } else if (LevelType.SDT_RUN == levelType) {
            CTR[] rArray = ctSdtContentRun.getRArray();
            for (CTR ctr : rArray) {
                CTText[] tArray = ctr.getTArray();
                for (CTText ctText : tArray) {
                    value.append(ctText.getStringValue());
                }
            }
        } else if (LevelType.SDT_CELL == levelType) {
            CTTc[] tcArray = ctSdtContentCell.getTcArray();
            for (CTTc ctTc : tcArray) {
                CTP[] pArray = ctTc.getPArray();
                for (CTP ctp : pArray) {
                    CTR[] rArray = ctp.getRArray();
                    for (CTR ctr : rArray) {
                        for (CTText ctText : ctr.getTArray()) {
                            value.append(ctText.getStringValue());
                        }
                    }
                    flag++;
                    if (flag == pArray.length) break;
                    value.append("\n");
                }
            }
        }
        return value.toString();
    }
    protected void setRichText(String value){
        String[] split = value.split("\n");
        if (LevelType.SDT_BLOCK == levelType){
            int length = ctSdtContentBlock.getPArray().length;
            for (int i = length-1; i >= 0; i--) {
                ctSdtContentBlock.removeP(i);
            }
            for (String text : split) {
                ctSdtContentBlock.addNewP().addNewR().addNewT().setStringValue(text);
            }
        } else if (LevelType.SDT_RUN == levelType) {
            int length = ctSdtContentRun.getRArray().length;
            for (int i = length-1; i >= 0; i--) {
                ctSdtContentRun.removeR(i);
            }
            ctSdtContentRun.addNewR().addNewT().setStringValue(value);
        } else if (LevelType.SDT_CELL == levelType) {
            CTTc tc = ctSdtContentCell.getTcArray(0);
            int length = tc.getPArray().length;
            for (int i = length-1; i >= 0; i--) {
                tc.removeP(i);
            }
            for (String text : split) {
                tc.addNewP().addNewR().addNewT().setStringValue(text);
            }
        }
    }
    protected void setPlainText(String value){
        if (LevelType.SDT_BLOCK == levelType){
            int length = ctSdtContentBlock.getPArray().length;
            for (int i = length-1; i >= 0; i--) {
                ctSdtContentBlock.removeP(i);
            }
            ctSdtContentBlock.addNewP().addNewR().addNewT().setStringValue(value);
        } else if (LevelType.SDT_RUN == levelType) {
            int length = ctSdtContentRun.getRArray().length;
            for (int i = length-1; i >= 0; i--) {
                ctSdtContentRun.removeR(i);
            }
            ctSdtContentRun.addNewR().addNewT().setStringValue(value);
        } else if (LevelType.SDT_CELL == levelType) {
            CTTc tc = ctSdtContentCell.getTcArray(0);
            int length = tc.getPArray().length;
            for (int i = length-1; i >= 0; i--) {
                tc.removeP(i);
            }
            tc.addNewP().addNewR().addNewT().setStringValue(value);
        }
    }

    protected String getPictureRelationId(){
        if (LevelType.SDT_BLOCK == levelType){
            CTDrawing drawing= ctSdtContentBlock.getPArray(0).getRArray(0).getDrawingArray(0);
            List<CTPicture> ctPictures = getCTPictures(drawing);
            CTPicture ctPicture = ctPictures.get(0);
            return getPictureEmbed(ctPicture);
        } else if (LevelType.SDT_RUN == levelType) {
            CTDrawing drawing = ctSdtContentRun.getRArray(0).getDrawingArray(0);
            List<CTPicture> ctPictures = getCTPictures(drawing);
            CTPicture ctPicture = ctPictures.get(0);
            return getPictureEmbed(ctPicture);
        } else if (LevelType.SDT_CELL == levelType) {
            CTDrawing drawing = ctSdtContentCell.getTcArray(0).getPArray(0).getRArray(0).getDrawingArray(0);
            List<CTPicture> ctPictures = getCTPictures(drawing);
            CTPicture ctPicture = ctPictures.get(0);
            return getPictureEmbed(ctPicture);
        }
        return null;
    }

    private String getPictureEmbed(CTPicture ctPicture){
        CTBlipFillProperties blipProps = ctPicture.getBlipFill();
        if (blipProps == null || !blipProps.isSetBlip()) {
            // return null if Blip data is missing
            return null;
        }
        return blipProps.getBlip().getEmbed();
    }

    private List<CTPicture> getCTPictures(XmlObject o) {
        List<CTPicture> pics = new ArrayList<>();
        String xquery = "declare namespace pic='" + CTPicture.type.getName().getNamespaceURI() + "' .//pic:pic";
        XmlObject[] picts = o.selectPath(xquery);
        for (XmlObject pict : picts) {
            if (pict instanceof XmlAnyTypeImpl) {
                // Pesky XmlBeans bug - see Bugzilla #49934
                try {
                    pict = CTPicture.Factory.parse(pict.toString(), DEFAULT_XML_OPTIONS);
                } catch (XmlException e) {
                    throw new POIXMLException(e);
                }
            }
            if (pict instanceof CTPicture) {
                pics.add((CTPicture) pict);
            }
        }
        return pics;
    }

    public void setPictureReference(String relationId){
        if (LevelType.SDT_BLOCK == levelType){
            CTDrawing drawing= ctSdtContentBlock.getPArray(0).getRArray(0).getDrawingArray(0);
            List<CTPicture> ctPictures = getCTPictures(drawing);
            CTPicture ctPicture = ctPictures.get(0);
            ctPicture.getBlipFill().getBlip().setEmbed(relationId);
        } else if (LevelType.SDT_RUN == levelType) {
            CTDrawing drawing = ctSdtContentRun.getRArray(0).getDrawingArray(0);
            List<CTPicture> ctPictures = getCTPictures(drawing);
            CTPicture ctPicture = ctPictures.get(0);
            ctPicture.getBlipFill().getBlip().setEmbed(relationId);
        } else if (LevelType.SDT_CELL == levelType) {
            CTDrawing drawing = ctSdtContentCell.getTcArray(0).getPArray(0).getRArray(0).getDrawingArray(0);
            List<CTPicture> ctPictures = getCTPictures(drawing);
            CTPicture ctPicture = ctPictures.get(0);
            ctPicture.getBlipFill().getBlip().setEmbed(relationId);
        }
    }

    public void setCheckboxText(String text){
        if (LevelType.SDT_BLOCK == levelType){
            ctSdtContentBlock.getPArray(0).getRArray(0).getTArray(0).setStringValue(text);
        } else if (LevelType.SDT_RUN == levelType) {
            ctSdtContentRun.getRArray(0).getTArray(0).setStringValue(text);
        } else if (LevelType.SDT_CELL == levelType) {
            ctSdtContentCell.getTcArray(0).getPArray(0).getRArray(0).getTArray(0).setStringValue(text);
        }
    }
}
