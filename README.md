# what is poi-enhance？

Poi-enhance adds operations on SDT based on apache poi.

# dependency

```xml
<dependency>
  <groupId>io.github.cnless</groupId>
  <artifactId>poi-enhance</artifactId>
  <version>1.0.0</version>
</dependency>
```



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

