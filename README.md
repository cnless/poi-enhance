# what is poi-enhance？

Poi-enhance adds operations on SDT based on apache poi.

# usage



```java
//create a XWPFDocumentEnhance Object
            XWPFDocumentEnhance xwpfDocumentEnhance = new XWPFDocumentEnhance(fileInputStream);
            //get all StructuredDocumentTagEnhance,StructuredDocumentTagEnhance contains all method to operate StructuredDocumentTag
            List<StructuredDocumentTagEnhance> sdtList = xwpfDocumentEnhance.getSdtList();
            //save the file
            xwpfDocumentEnhance.writeFile(fileOutputStream);
            xwpfDocumentEnhance.close();
```

