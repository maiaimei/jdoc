package cn.maiaimei.jdocs;

import java.awt.image.BufferedImage;
import java.io.BufferedWriter;
import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStreamWriter;
import java.nio.charset.StandardCharsets;
import java.util.Calendar;
import java.util.GregorianCalendar;
import java.util.Iterator;
import java.util.List;
import javax.imageio.ImageIO;
import org.apache.pdfbox.multipdf.PDFMergerUtility;
import org.apache.pdfbox.multipdf.Splitter;
import org.apache.pdfbox.pdmodel.PDDocument;
import org.apache.pdfbox.pdmodel.PDDocumentInformation;
import org.apache.pdfbox.pdmodel.PDPage;
import org.apache.pdfbox.pdmodel.PDPageContentStream;
import org.apache.pdfbox.pdmodel.encryption.AccessPermission;
import org.apache.pdfbox.pdmodel.encryption.StandardProtectionPolicy;
import org.apache.pdfbox.pdmodel.font.PDFont;
import org.apache.pdfbox.pdmodel.font.PDType0Font;
import org.apache.pdfbox.pdmodel.font.PDType1Font;
import org.apache.pdfbox.pdmodel.graphics.image.PDImageXObject;
import org.apache.pdfbox.rendering.PDFRenderer;
import org.apache.pdfbox.text.PDFTextStripper;
import org.fit.pdfdom.PDFDomTree;
import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Test;

public class PdfBoxTest {

  private static final String PATH = "C:/Users/lenovo/Desktop/tmp";

  @DisplayName("创建一个空的PDF文档")
  @Test
  public void testDocumentCreation() throws IOException {
    //Creating PDF document object 
    PDDocument document = new PDDocument();

    //Saving the document
    document.save(PATH + "/sample.pdf");

    System.out.println("PDF created");

    //Closing the document  
    document.close();
  }

  @DisplayName("将页面添加到PDF文档")
  @Test
  public void testAddingPages() throws IOException {
    //Creating PDF document object 
    PDDocument document = new PDDocument();

    for (int i = 0; i < 3; i++) {
      //Creating a blank page 
      PDPage blankPage = new PDPage();

      //Adding the blank page to the document
      document.addPage(blankPage);
    }

    //Saving the document
    document.save(PATH + "/sample.pdf");

    System.out.println("PDF created");

    //Closing the document
    document.close();
  }

  @DisplayName("加载现有的PDF文档")
  @Test
  public void testLoadingExistingDocument() throws IOException {
    // Loading an existing document
    File file = new File(PATH + "/sample.pdf");
    PDDocument document = PDDocument.load(file);

    System.out.println("PDF loaded");

    // Adding a blank page to the document
    document.addPage(new PDPage());

    // Saving the document
    document.save(PATH + "/sample-add-pages.pdf");

    // Closing the document
    document.close();
  }

  @DisplayName("从现有文档中删除页面")
  @Test
  public void testRemovingPages() throws IOException {
    //Loading an existing document
    File file = new File(PATH + "/sample.pdf");
    PDDocument document = PDDocument.load(file);

    //Listing the number of existing pages
    int noOfPages = document.getNumberOfPages();

    System.out.println(noOfPages);

    //Removing the pages
    document.removePage(2);

    System.out.println("page removed");

    //Saving the document
    document.save(PATH + "/sample-remove-pages.pdf");

    //Closing the document
    document.close();
  }

  @DisplayName("设置文档属性")
  @Test
  public void testAddingDocumentAttributes() throws IOException {
    //Creating PDF document object
    PDDocument document = new PDDocument();

    //Creating a blank page
    PDPage blankPage = new PDPage();

    //Adding the blank page to the document
    document.addPage(blankPage);

    //Creating the PDDocumentInformation object 
    PDDocumentInformation pdd = document.getDocumentInformation();

    //Setting the author of the document
    pdd.setAuthor("Yiibai.com");

    // Setting the title of the document
    pdd.setTitle("一个简单的文档标题");

    //Setting the creator of the document 
    pdd.setCreator("PDF Examples");

    //Setting the subject of the document 
    pdd.setSubject("文档标题");

    //Setting the created date of the document 
    Calendar date = new GregorianCalendar();
    date.set(2017, 11, 5);
    pdd.setCreationDate(date);
    //Setting the modified date of the document 
    date.set(2018, 10, 5);
    pdd.setModificationDate(date);

    //Setting keywords for the document 
    pdd.setKeywords("pdfbox, first example, my pdf");

    //Saving the document 
    document.save(PATH + "/sample-doc-attributes.pdf");

    System.out.println("Properties added successfully ");

    //Closing the document
    document.close();
  }

  @DisplayName("检索文档属性")
  @Test
  public void testRetrivingDocumentAttributes() throws IOException {
    //Loading an existing document 
    File file = new File(PATH + "/sample-doc-attributes.pdf");
    PDDocument document = PDDocument.load(file);

    //Getting the PDDocumentInformation object
    PDDocumentInformation pdd = document.getDocumentInformation();

    //Retrieving the info of a PDF document
    System.out.println("Author of the document is :" + pdd.getAuthor());
    System.out.println("Title of the document is :" + pdd.getTitle());
    System.out.println("Subject of the document is :" + pdd.getSubject());

    System.out.println("Creator of the document is :" + pdd.getCreator());
    System.out.println("Creation date of the document is :" + pdd.getCreationDate());
    System.out.println("Modification date of the document is :" +
        pdd.getModificationDate());
    System.out.println("Keywords of the document are :" + pdd.getKeywords());

    //Closing the document 
    document.close();
  }

  @DisplayName("将文本添加到现有的PDF文档")
  @Test
  public void testAddingContent() throws IOException {
    //Loading an existing document
    File file = new File(PATH + "/sample.pdf");
    PDDocument document = PDDocument.load(file);

    //Retrieving the pages of the document 
    PDPage page = document.getPage(0);
    PDPageContentStream contentStream = new PDPageContentStream(document, page);

    //Begin the Content stream 
    contentStream.beginText();
    // contentStream.set
    //Setting the font to the Content stream  
    contentStream.setFont(PDType1Font.HELVETICA_BOLD, 14);

    //Setting the position for the line 
    contentStream.newLineAtOffset(25, 500);

    String text = "This is the sample document and we are adding content to it. - By yiibai.com";

    //Adding text in the form of string 
    contentStream.showText(text);

    //Ending the content stream
    contentStream.endText();

    System.out.println("Content added");

    //Closing the content stream
    contentStream.close();

    //Saving the document
    document.save(new File(PATH + "/sample-adding-content.pdf"));

    //Closing the document
    document.close();
  }

  @DisplayName("将多行文本添加到现有的PDF文档")
  @Test
  public void testAddMultipleLines() throws IOException {
    //Loading an existing document
    File file = new File(PATH + "/sample.pdf");
    PDDocument doc = PDDocument.load(file);

    //Creating a PDF Document
    PDPage page = doc.getPage(1);
    PDFont font = PDType0Font.load(doc, new File("c:/windows/fonts/times.ttf"));

    PDPageContentStream contentStream = new PDPageContentStream(doc, page);

    //Begin the Content stream 
    contentStream.beginText();

    //Setting the font to the Content stream
    contentStream.setFont(PDType1Font.TIMES_ROMAN, 16);
    System.out.println(" getName => " + font.getName());
    // contentStream.setFont( font, 16 );

    //Setting the leading
    contentStream.setLeading(14.5f);

    //Setting the position for the line
    contentStream.newLineAtOffset(25, 725);

    String text1 = "This is an example of adding text to a page in the pdf document. we can add as many lines";
    String text2 = "as we want like this using the ShowText()  method of the ContentStream class";
    //Adding text in the form of string
    contentStream.showText(text1);
    contentStream.newLine();
    contentStream.newLine();
    contentStream.showText(text2);
    contentStream.newLine();
    //Ending the content stream
    contentStream.endText();

    System.out.println("Content added");

    //Closing the content stream
    contentStream.close();

    //Saving the document
    doc.save(new File(PATH + "/sample-multiple-lines.pdf"));

    //Closing the document
    doc.close();
  }

  @DisplayName("添加文本")
  @Test
  public void testWritingText() {
    // 创建空白文档
    try (PDDocument doc = new PDDocument()) {
      // 创建一个空白页面
      PDPage page = new PDPage();
      // 将页面添加到文档，第2、3...N页如此类推
      doc.addPage(page);

      PDFont font = PDType1Font.HELVETICA_BOLD;

      try (PDPageContentStream contents = new PDPageContentStream(doc, page)) {
        contents.beginText();
        contents.setFont(font, 12);
        contents.newLineAtOffset(100, 700);
        contents.showText("pdfbox sample");
        contents.endText();
      } catch (IOException e) {
        throw new RuntimeException(e);
      }

      // 保存文档
      doc.save(PATH + "/sample.pdf");
    } catch (IOException e) {
      throw new RuntimeException(e);
    }
  }

  @DisplayName("从现有的PDF文档中提取文本")
  @Test
  public void testReadingText() throws IOException {
    //Loading an existing document
    File file = new File(PATH + "/sample-multiple-lines.pdf");
    PDDocument document = PDDocument.load(file);

    //Instantiate PDFTextStripper class
    PDFTextStripper pdfStripper = new PDFTextStripper();

    //Retrieving text from PDF document
    String text = pdfStripper.getText(document);
    System.out.println(text);

    //Closing the document
    document.close();
  }

  @DisplayName("将图像插入PDF文档")
  @Test
  public void testInsertingImage() throws IOException {
    //Loading an existing document
    File file = new File(PATH + "/sample.pdf");
    PDDocument doc = PDDocument.load(file);

    //Retrieving the page
    PDPage page = doc.getPage(0);

    //Creating PDImageXObject object
    PDImageXObject pdImage = PDImageXObject.createFromFile(PATH + "/logo.png", doc);

    //creating the PDPageContentStream object
    PDPageContentStream contents = new PDPageContentStream(doc, page);

    //Drawing the image in the PDF document
    contents.drawImage(pdImage, 70, 250);

    System.out.println("Image inserted");

    //Closing the PDPageContentStream object
    contents.close();

    //Saving the document
    doc.save(PATH + "/sample-image.pdf");

    //Closing the document
    doc.close();
  }

  @DisplayName("加密PDF文档")
  @Test
  public void testEncryptingPDF() throws IOException {
    //Loading an existing document
    File file = new File(PATH + "/sample.pdf");
    PDDocument document = PDDocument.load(file);

    //Creating access permission object
    AccessPermission ap = new AccessPermission();

    //Creating StandardProtectionPolicy object
    StandardProtectionPolicy spp = new StandardProtectionPolicy("123456", "123456", ap);

    //Setting the length of the encryption key
    spp.setEncryptionKeyLength(128);

    //Setting the access permissions
    spp.setPermissions(ap);

    //Protecting the document
    document.protect(spp);

    System.out.println("Document encrypted");

    //Saving the document
    document.save(PATH + "/sample-encrypt.pdf");
    //Closing the document
    document.close();
  }

  @DisplayName("分割PDF文档中的页面")
  @Test
  public void testSplitPages() throws IOException {
    //Loading an existing PDF document
    File file = new File(PATH + "/sample.pdf");
    PDDocument document = PDDocument.load(file);

    //Instantiating Splitter class
    Splitter splitter = new Splitter();

    //splitting the pages of a PDF document
    List<PDDocument> Pages = splitter.split(document);

    //Creating an iterator 
    Iterator<PDDocument> iterator = Pages.listIterator();

    //Saving each page as an individual document
    int i = 1;
    while (iterator.hasNext()) {
      PDDocument pd = iterator.next();
      pd.save(PATH + "/sample-split-pages-" + i + ".pdf");
      i = i + 1;
    }
    System.out.println("Multiple PDF’s created");
    document.close();
  }

  @DisplayName("合并多个PDF文档")
  @Test
  public void testMergePDFs() throws IOException {
    //Loading an existing PDF document
    File file1 = new File(PATH + "/sample1.pdf");
    writePdf(file1.getAbsolutePath(), file1.getName());
    PDDocument doc1 = PDDocument.load(file1);

    File file2 = new File(PATH + "/sample2.pdf");
    writePdf(file2.getAbsolutePath(), file2.getName());
    PDDocument doc2 = PDDocument.load(file2);

    //Instantiating PDFMergerUtility class
    PDFMergerUtility PDFmerger = new PDFMergerUtility();

    //Setting the destination file
    PDFmerger.setDestinationFileName(PATH + "/sample-merged.pdf");

    //adding the source files
    PDFmerger.addSource(file1);
    PDFmerger.addSource(file2);

    //Merging the two documents
    PDFmerger.mergeDocuments();

    System.out.println("Documents merged");
    //Closing the documents
    doc1.close();
    doc2.close();
  }

  @DisplayName("PDF转Image")
  @Test
  public void testPdfToImage() throws IOException {
    //Loading an existing PDF document
    File file = new File(PATH + "/sample-merged.pdf");
    PDDocument document = PDDocument.load(file);

    //Instantiating the PDFRenderer class
    PDFRenderer renderer = new PDFRenderer(document);

    //Rendering an image from the PDF document
    BufferedImage image = renderer.renderImage(0);

    //Writing the image to a file
    ImageIO.write(image, "JPEG", new File(PATH + "/sample-pdf2image.jpg"));

    System.out.println("Image created");

    //Closing the document
    document.close();
  }

  @DisplayName("PDF转HTML")
  @Test
  public void pdfToHtmlTest() throws IOException {
    String pathname = PATH + "/sample-merged.pdf";
    String outputPath = PATH + "/sample-pdf2html.html";
    final File file = new File(outputPath);
    if (!file.exists()) {
      file.createNewFile();
    }
    try (BufferedWriter out = new BufferedWriter(new OutputStreamWriter(new FileOutputStream(
        outputPath), StandardCharsets.UTF_8))) {
      PDDocument document = PDDocument.load(new File(pathname));
      PDFDomTree pdfDomTree = new PDFDomTree();
      pdfDomTree.writeText(document, out);
    } catch (Exception e) {
      e.printStackTrace();
    }
  }

  private void writePdf(String pathname, String text) {
    try (PDDocument document = new PDDocument()) {
      final PDPage page = new PDPage();
      document.addPage(page);
      try (PDPageContentStream contentStream = new PDPageContentStream(document, page)) {
        contentStream.beginText();
        contentStream.setFont(PDType1Font.HELVETICA_BOLD, 14);
        contentStream.newLineAtOffset(25, 500);
        contentStream.showText(text);
        contentStream.endText();
      }
      document.save(pathname);
    } catch (IOException e) {
      throw new RuntimeException(e);
    }
  }
}
