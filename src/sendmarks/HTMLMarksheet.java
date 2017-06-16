/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package sendmarks;

import java.io.BufferedReader;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileWriter;
import java.io.IOException;
import java.io.InputStream;
import java.io.InputStreamReader;
import java.util.ArrayList;
import java.util.Properties;
import java.util.prefs.Preferences;
import java.util.regex.Matcher;
import java.util.regex.Pattern;
import javax.mail.Message;
import javax.mail.MessagingException;
import javax.mail.PasswordAuthentication;
import javax.mail.Session;
import javax.mail.Transport;
import javax.mail.internet.InternetAddress;
import javax.mail.internet.MimeMessage;
import javax.swing.JOptionPane;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.util.AreaReference;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 *
 * @author James M. Curran <j.curran@auckland.ac.nz>
 */
public class HTMLMarksheet {

  private String strFormattedHTMLSheet;
  private String strNetID, strEmail;
  private double finalMark;

  public HTMLMarksheet(File f1, SheetInfo sheetInfo) throws FilenameFormatException, FileNotFoundException, IOException {

    Pattern p = Pattern.compile("^(?<netid>[A-Za-z]{3,4}[0-9]{3})\\.xlsx");
    Matcher m = p.matcher(f1.getName());

    if (m.find()) {
      strNetID = m.group("netid");
      strEmail = strNetID + "@aucklanduni.ac.nz";
      
      InputStream excelFileToRead;
      XSSFWorkbook workbook = null;

      excelFileToRead = new FileInputStream(f1.getAbsolutePath());
      workbook = new XSSFWorkbook(excelFileToRead);

      /*XSSFSheet worksheet = workbook.getSheetAt(0);
        XSSFRow row; 
        XSSFCell cell;

        row = worksheet.getRow(0);
        cell = row.getCell(1);*/
 /*int nameRangeIdx = workbook.getNameIndex("Name");      
        int markRangeIdx = workbook.getNameIndex("Marks");
        int finalMarkRangeIdx = workbook.getNameIndex("FinalMark");
        
        Name nameRange = workbook.getNameAt(nameRangeIdx);
        Name markRange = workbook.getNameAt(markRangeIdx);
        Name finalMarkRange = workbook.getNameAt(finalMarkRangeIdx);*/
      String strSheetName = sheetInfo.getSheetName();
      String strNameRange = sheetInfo.getNameRange();
      String strMarkRange = sheetInfo.getMarkRange();
      String strFinalMarkRange = sheetInfo.getFinalMarkRange();

      // retrieve the cell at the named range and test its contents
      //AreaReference aref = new AreaReference(nameRange.getRefersToFormula());
      //AreaReference aref = new AreaReference("Sheet1!$B$1:$B$2");
      AreaReference aref = new AreaReference(strNameRange);
      CellReference[] crefs = aref.getAllReferencedCells();

      Sheet sheet = workbook.getSheet(crefs[0].getSheetName());
      Row r = sheet.getRow(crefs[0].getRow());
      Cell c = r.getCell(crefs[0].getCol());
      String strStudentName = c.getStringCellValue();

      sheet = workbook.getSheet(crefs[1].getSheetName());
      r = sheet.getRow(crefs[1].getRow());
      c = r.getCell(crefs[1].getCol());
      String strStudentUPI = c.getStringCellValue();

      //aref = new AreaReference(finalMarkRange.getRefersToFormula());
      //aref = new AreaReference("Sheet1!$C$25:$D$25");
      aref = new AreaReference(strFinalMarkRange);
      crefs = aref.getAllReferencedCells();
      double totalMark = 0;
      finalMark = 0;
      double percentMark = 0;

      for (int i = 0; i < 2; i++) {
        sheet = workbook.getSheet(crefs[i].getSheetName());
        r = sheet.getRow(crefs[i].getRow());
        c = r.getCell(crefs[i].getCol());

        if (i == 0) {
          totalMark = c.getNumericCellValue();
        } else {
          finalMark = c.getNumericCellValue();
        }
      }

      percentMark = 100 * finalMark / totalMark;

      //aref = new AreaReference(markRange.getRefersToFormula());
      //aref = new AreaReference("Sheet1!$A$4:$E$24");
      aref = new AreaReference(strMarkRange);
      crefs = aref.getAllReferencedCells();

      int numRows = aref.getLastCell().getRow() - aref.getFirstCell().getRow() + 1;
      int numCols = aref.getLastCell().getCol() - aref.getFirstCell().getCol() + 1;

      Object[][] markTable = new Object[numRows][];

      for (int row = 0; row < numRows; row++) {

        markTable[row] = new Object[numCols];

        for (int col = 0; col < numCols; col++) {
          int idx = row * numCols + col;

          CellReference cr = crefs[idx];
          sheet = workbook.getSheet(cr.getSheetName());
          r = sheet.getRow(cr.getRow());
          c = r.getCell(cr.getCol());

          if (c != null) {
            int cellType = c.getCellType();

            switch (cellType) {
              case XSSFCell.CELL_TYPE_STRING:
                markTable[row][col] = c.getStringCellValue();
                break;
              case XSSFCell.CELL_TYPE_NUMERIC:
              case XSSFCell.CELL_TYPE_FORMULA:
                double d = c.getNumericCellValue();
                if (d - Math.floor(d) < 0.01) {
                  markTable[row][col] = (int) d;
                } else {
                  markTable[row][col] = d;
                }
                break;
              case XSSFCell.CELL_TYPE_BLANK:
                markTable[row][col] = "";
                break;
              default:
                markTable[row][col] = "";
                break;
            }
          }
        }
      }

      Object[] studentData = {sheetInfo.getAssignmentNumber(),sheetInfo.getCourse(),
                              strStudentName, strStudentUPI,
                              finalMark, totalMark, percentMark};

      String strStudentDetails = readStudentDetailsTemplate();
      String strStudentDetailsTable = completeStudentDetails(strStudentDetails, studentData);

      StringBuilder sb = new StringBuilder();

      sb.append("<!DOCTYPE html>\n<html>\n<head>\n");
      sb.append("<title>");
      sb.append(sheetInfo.getCourse());
      sb.append(" Assignment ");
      sb.append(sheetInfo.getAssignmentNumber());
      sb.append("</title>\n");
      sb.append("<style>\n");
      sb.append(readCSS("\t\t"));
      sb.append("\t</style>\n");
      sb.append("</head>\n<body>\n");

      sb.append(strStudentDetailsTable);

      sb.append((new Marktable(markTable)).toString());

      sb.append("\t\t\t</table>\n");
      sb.append("\t\t</div>\n");
      sb.append("\t</body>\n");
      sb.append("</html>\n");

      strFormattedHTMLSheet = sb.toString();
    } else {
      StringBuilder sb = new StringBuilder();

      sb.append("The file ").append(f1.getName()).append(" does not conform to the format: UPI.xlsx");
      sb.append(" where UPI consists of 3 or 4 letters and exaxtly three numbers, e.g. jcur002");
      throw new FilenameFormatException(sb.toString());
    }
  }

  @Override
  public String toString() {
    return strFormattedHTMLSheet;
  }
  
  public String getNetId(){
    return strNetID;
  }
  
  public double getFinalMark(){
    return finalMark;
  }
 
  public void send(Session session, String strSubject, String from, boolean bDummyRun) {
    // Recipient's email ID needs to be mentioned.
    String to = strEmail;

    try {
      // Create a default MimeMessage object.
      Message message = new MimeMessage(session);

      // Set From: header field of the header.
      message.setFrom(new InternetAddress(from));

      // Set To: header field of the header.
      message.setRecipients(Message.RecipientType.TO,
        InternetAddress.parse(to));

      // Set Subject: header field
      message.setSubject(strSubject);

      // Send the actual HTML message, as big as you like
      message.setContent(strFormattedHTMLSheet, "text/html");

      // Send message
      if(!bDummyRun){
        Transport.send(message);
        System.out.println("Sent message successfully....");
      }else{
        System.out.println("DummyRun Sent message successfully....");
      }

    } catch (MessagingException e) {
      e.printStackTrace();
      throw new RuntimeException(e);
    }
  }

  public void write(String strPath) throws IOException {
    FileWriter f1 = new FileWriter(strPath + strNetID + ".html");

    f1.write(strFormattedHTMLSheet);
    f1.write("\n");
    f1.close();

  }

  private String completeStudentDetails(String strDetails, Object[] studentData) {
    String strStudentDetails = fillDetails("ASSNUM", (String) studentData[0], strDetails);
    
    strStudentDetails = fillDetails("COURSE", (String) studentData[1], strStudentDetails);
    
    
    strStudentDetails = fillDetails("STUDENTNAME", (String) studentData[2], strStudentDetails);
    strStudentDetails = fillDetails("STUDENTUPI", (String) studentData[3], strStudentDetails);
    strStudentDetails = fillDetails("STUDENTEMAIL", String.format("%s@aucklanduni.ac.nz", (String) studentData[3]), strStudentDetails);

    double finalMark = (Double) studentData[4];
    double totalMark = (Double) studentData[5];
    double percentMark = (Double) studentData[6];

    String strFinalMark = ((finalMark - Math.floor(finalMark)) > 0.1) ? String.format("%4.1f", Math.floor(finalMark))
      : String.format("%d", (int) finalMark);

    String strTotalMark = ((totalMark - Math.floor(totalMark)) > 0.1) ? String.format("%4.1f", Math.floor(totalMark))
      : String.format("%d", (int) totalMark);

    String strPercentMark = String.format("%5.2f", percentMark);

    strStudentDetails = fillDetails("YOURMARK", strFinalMark, strStudentDetails);
    strStudentDetails = fillDetails("TOTALMARK", strTotalMark, strStudentDetails);
    strStudentDetails = fillDetails("PERCENTMARK", strPercentMark, strStudentDetails);

    return strStudentDetails;
  }

  private String readCSS(String indent) {
    InputStream in = getClass().getResourceAsStream("HTMLResources/assignment.css");
    BufferedReader reader = new BufferedReader(new InputStreamReader(in));
    StringBuilder sb = new StringBuilder();

    try {
      String line = reader.readLine();

      while (line != null) {
        sb.append(indent).append(line).append('\n');
        line = reader.readLine();
      }
    } catch (IOException ex) {
      ex.printStackTrace();
    }

    return sb.toString();
  }

  private String fillDetails(String strTag, String strFmtData, String strDetails) {
    StringBuilder sb = new StringBuilder();

    sb.append("(?s)(.*)").append(strTag).append("(.*$)");

    String strPattern = sb.toString();
    Pattern p = Pattern.compile(strPattern);
    Matcher m = p.matcher(strDetails);

    if (!m.find()) {
      System.out.println("Couldn't find pattern: " + strPattern);
    }

    sb = new StringBuilder();
    sb.append("$1").append(strFmtData).append("$2");
    String strReplace = sb.toString();

    return strDetails.replaceAll(strPattern, strReplace);
  }

  private String readStudentDetailsTemplate() {
    InputStream in = getClass().getResourceAsStream("HTMLResources/studentDetails.html");
    BufferedReader reader = new BufferedReader(new InputStreamReader(in));
    StringBuilder sb = new StringBuilder();

    try {
      String line = reader.readLine();

      while (line != null) {
        sb.append(line).append('\n');
        line = reader.readLine();
      }
    } catch (IOException ex) {
      ex.printStackTrace();
    }

    return sb.toString();
  }
}
