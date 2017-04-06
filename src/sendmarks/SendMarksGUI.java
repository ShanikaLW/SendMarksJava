/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package sendmarks;

import java.awt.Toolkit;
import java.awt.datatransfer.Clipboard;
import java.awt.datatransfer.StringSelection;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.awt.event.MouseAdapter;
import java.awt.event.MouseEvent;
import java.awt.event.MouseListener;
import java.io.BufferedReader;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FilenameFilter;
import java.io.IOException;
import java.io.InputStream;
import java.io.InputStreamReader;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Iterator;
import java.util.logging.Level;
import java.util.logging.Logger;
import java.util.regex.Matcher;
import java.util.regex.Pattern;
import javax.swing.JFileChooser;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Name;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.util.AreaReference;
import org.apache.poi.ss.util.CellReference;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 *
 * @author James M. Curran <j.curran@auckland.ac.nz>
 */
public class SendMarksGUI extends javax.swing.JFrame {

  public class PopupListener extends MouseAdapter {

    public void mousePressed(MouseEvent e) {
      maybeShowPopup(e);
    }

    public void mouseReleased(MouseEvent e) {
      maybeShowPopup(e);
    }

    private void maybeShowPopup(MouseEvent e) {
      if (e.isPopupTrigger()) {
        jPopupMenuEdit.show(e.getComponent(),
          e.getX(), e.getY());
      }
    }
  }
  /**
   * Creates new form SendMarksGUI
   */
  public SendMarksGUI() {
    initComponents();
    
    jcbDummyRun.setSelected(true);
    bDummyRun = true;
    
    jPopupMenuEdit = new javax.swing.JPopupMenu();
    jPMICopy = new javax.swing.JMenuItem("Copy");

    jPMICopy.addActionListener(new ActionListener() {
      @Override
      public void actionPerformed(ActionEvent evt) {
        jPMICopyActionPerformed(evt);
      }
    });

    jPMICopy.setAccelerator(javax.swing.KeyStroke.getKeyStroke(java.awt.event.KeyEvent.VK_C, java.awt.event.InputEvent.META_MASK));


    jPopupMenuEdit.add(jPMICopy);

    MouseListener popupListener = new PopupListener();
    jtaActivityLog.addMouseListener(popupListener);
    //m_strCWD = "/Users/jcur002/curran/Work/2016/Teaching/779/Assignments/A3/Marks/MarksAss3";
    m_strCWD = "/Users/jcur002/Dropbox/Work/2017/Teaching/779/Assignments/A1/Marks";
    jlabCurrentDirectory.setText(m_strCWD);
    
    jtfNameRange.setInputVerifier(new RangeInputVerifier());
    jtfMarkRange.setInputVerifier(new RangeInputVerifier());
    jtfFinalMarkRange.setInputVerifier(new RangeInputVerifier());
    
    
    
    readCSS("\t\t");
    readStudentDetails();
  }
  
  private String fillDetails(String strTag, String strFmtData, String strDetails){
    StringBuilder sb = new StringBuilder();
  
    sb.append("(?s)(.*)").append(strTag).append("(.*$)");
    
    String strPattern = sb.toString();
    Pattern p = Pattern.compile(strPattern);
    Matcher m = p.matcher(strDetails);
    
    
    if(!m.find()){
      System.out.println("Couldn't find pattern: " + strPattern);
    }
    
    sb = new StringBuilder();
    sb.append("$1").append(strFmtData).append("$2");
    String strReplace = sb.toString();
    
    return strDetails.replaceAll(strPattern, strReplace);
  }
  
  private String completeStudentDetails(String strDetails, Object[] studentData){
    String strStudentDetails = fillDetails("ASSNUM", (String)studentData[0], strDetails);
    strStudentDetails = fillDetails("STUDENTNAME", (String)studentData[1], strStudentDetails);
    strStudentDetails = fillDetails("STUDENTUPI", (String)studentData[2], strStudentDetails);
    strStudentDetails = fillDetails("STUDENTEMAIL", String.format("%s@aucklanduni.ac.nz", (String)studentData[2]),strStudentDetails);

    double finalMark = (Double)studentData[3];
    double totalMark = (Double)studentData[4];
    double percentMark = (Double)studentData[5];
    
    String strFinalMark = ((finalMark - Math.floor(finalMark)) > 0.1) ? String.format("%4.1f", Math.floor(finalMark)) : 
                         String.format("%d", (int)finalMark);

    String strTotalMark = ((totalMark - Math.floor(totalMark)) > 0.1) ? String.format("%4.1f", Math.floor(totalMark)) : 
                         String.format("%d", (int)totalMark);

    String strPercentMark = String.format("%5.2f", percentMark);

    strStudentDetails = fillDetails("YOURMARK", strFinalMark, strStudentDetails);
    strStudentDetails = fillDetails("TOTALMARK", strTotalMark, strStudentDetails);
    strStudentDetails = fillDetails("PERCENTMARK", strPercentMark, strStudentDetails);
    
    return strStudentDetails;
  }
  
  void readCSS(String indent){
    InputStream in = getClass().getResourceAsStream("HTMLResources/assignment.css"); 
    BufferedReader reader = new BufferedReader(new InputStreamReader(in));
    StringBuilder sb = new StringBuilder();
    
    try{
      String line = reader.readLine();
      
      while(line != null){
        sb.append(indent).append(line).append('\n');
        line = reader.readLine();
      }
    }catch(IOException ex){
      ex.printStackTrace();
    }
    
    strCSS = sb.toString();
  }
  
  void readStudentDetails(){
    InputStream in = getClass().getResourceAsStream("HTMLResources/studentDetails.html"); 
    BufferedReader reader = new BufferedReader(new InputStreamReader(in));
    StringBuilder sb = new StringBuilder();
    
    try{
      String line = reader.readLine();
      
      while(line != null){
        sb.append(line).append('\n');
        line = reader.readLine();
      }
    }catch(IOException ex){
      ex.printStackTrace();
    }
    
    strStudentDetails = sb.toString();
  }

  /**
   * This method is called from within the constructor to initialize the form. WARNING: Do NOT modify this code. The
   * content of this method is always regenerated by the Form Editor.
   */
  @SuppressWarnings("unchecked")
  // <editor-fold defaultstate="collapsed" desc="Generated Code">//GEN-BEGIN:initComponents
  private void initComponents() {

    jtfSubjectLine = new javax.swing.JTextField();
    jLabel2 = new javax.swing.JLabel();
    jLabel3 = new javax.swing.JLabel();
    jcbAssignmentNumber = new javax.swing.JComboBox<>();
    jbChangeDir = new javax.swing.JButton();
    jbSendMarks = new javax.swing.JButton();
    jbScrapeGrades = new javax.swing.JButton();
    jcbColumn = new javax.swing.JComboBox<>();
    jspinRow = new javax.swing.JSpinner();
    jLabel4 = new javax.swing.JLabel();
    jLabel5 = new javax.swing.JLabel();
    jScrollPane2 = new javax.swing.JScrollPane();
    jtaActivityLog = new javax.swing.JTextArea();
    jLabel6 = new javax.swing.JLabel();
    jLabel7 = new javax.swing.JLabel();
    jlabCurrentDirectory = new javax.swing.JLabel();
    jLabel1 = new javax.swing.JLabel();
    jLabel8 = new javax.swing.JLabel();
    jLabel9 = new javax.swing.JLabel();
    jtfMarkRange = new javax.swing.JTextField();
    jLabel10 = new javax.swing.JLabel();
    jcbSheetNumber = new javax.swing.JComboBox<>();
    jtfNameRange = new javax.swing.JTextField();
    jtfFinalMarkRange = new javax.swing.JTextField();
    jcbDummyRun = new javax.swing.JCheckBox();
    jMenuBar1 = new javax.swing.JMenuBar();
    jmenuSetup = new javax.swing.JMenu();
    jMenuItemConfigure = new javax.swing.JMenuItem();
    jMenu2 = new javax.swing.JMenu();

    setDefaultCloseOperation(javax.swing.WindowConstants.EXIT_ON_CLOSE);

    jtfSubjectLine.setText("STATS 779 Assignment Marks");

    jLabel2.setText("Enter your subject line here");

    jLabel3.setText("Assignment Number");

    jcbAssignmentNumber.setModel(new javax.swing.DefaultComboBoxModel<>(new String[] { "1", "2", "3", "4", "5" }));
    jcbAssignmentNumber.addActionListener(new java.awt.event.ActionListener() {
      public void actionPerformed(java.awt.event.ActionEvent evt) {
        jcbAssignmentNumberActionPerformed(evt);
      }
    });

    jbChangeDir.setText("Change Directory");
    jbChangeDir.addActionListener(new java.awt.event.ActionListener() {
      public void actionPerformed(java.awt.event.ActionEvent evt) {
        jbChangeDirActionPerformed(evt);
      }
    });

    jbSendMarks.setText("Send Marks");
    jbSendMarks.addActionListener(new java.awt.event.ActionListener() {
      public void actionPerformed(java.awt.event.ActionEvent evt) {
        jbSendMarksActionPerformed(evt);
      }
    });

    jbScrapeGrades.setText("Scrape Grades");
    jbScrapeGrades.addActionListener(new java.awt.event.ActionListener() {
      public void actionPerformed(java.awt.event.ActionEvent evt) {
        jbScrapeGradesActionPerformed(evt);
      }
    });

    jcbColumn.setModel(new javax.swing.DefaultComboBoxModel<>(new String[] { "A", "B", "C", "D", "E" }));
    jcbColumn.setSelectedIndex(3);

    jspinRow.setModel(new javax.swing.SpinnerNumberModel(30, 1, 50, 1));

    jLabel4.setText("Column");

    jLabel5.setText("Row");

    jtaActivityLog.setColumns(20);
    jtaActivityLog.setRows(5);
    jScrollPane2.setViewportView(jtaActivityLog);

    jLabel6.setText("Activity Log");

    jLabel7.setText("Current Directory:");

    jlabCurrentDirectory.setText("(None)");

    jLabel1.setText("Name Range:");

    jLabel8.setText("Mark Range:");

    jLabel9.setText("Final Mark Range:");

    jtfMarkRange.setText("$A$4:$E$24");

    jLabel10.setText("Sheet:");

    jcbSheetNumber.setModel(new javax.swing.DefaultComboBoxModel<>(new String[] { "1", "2", "3" }));

    jtfNameRange.setText("$B$1:$B$2");
    jtfNameRange.addActionListener(new java.awt.event.ActionListener() {
      public void actionPerformed(java.awt.event.ActionEvent evt) {
        jtfNameRangeActionPerformed(evt);
      }
    });

    jtfFinalMarkRange.setText("$C$25:$D$25");

    jcbDummyRun.setSelected(true);
    jcbDummyRun.setText("Dummy Run");
    jcbDummyRun.addActionListener(new java.awt.event.ActionListener() {
      public void actionPerformed(java.awt.event.ActionEvent evt) {
        jcbDummyRunActionPerformed(evt);
      }
    });

    jmenuSetup.setText("Setup");

    jMenuItemConfigure.setText("Configure...");
    jMenuItemConfigure.addActionListener(new java.awt.event.ActionListener() {
      public void actionPerformed(java.awt.event.ActionEvent evt) {
        jMenuItemConfigureActionPerformed(evt);
      }
    });
    jmenuSetup.add(jMenuItemConfigure);

    jMenuBar1.add(jmenuSetup);

    jMenu2.setText("Edit");
    jMenuBar1.add(jMenu2);

    setJMenuBar(jMenuBar1);

    javax.swing.GroupLayout layout = new javax.swing.GroupLayout(getContentPane());
    getContentPane().setLayout(layout);
    layout.setHorizontalGroup(
      layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
      .addGroup(layout.createSequentialGroup()
        .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
          .addGroup(layout.createSequentialGroup()
            .addContainerGap()
            .addComponent(jLabel7)
            .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
            .addComponent(jlabCurrentDirectory, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE))
          .addGroup(layout.createSequentialGroup()
            .addGap(16, 16, 16)
            .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
              .addGroup(layout.createSequentialGroup()
                .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                  .addComponent(jLabel8, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE)
                  .addComponent(jLabel1, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE)
                  .addComponent(jLabel9, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE))
                .addGap(18, 18, 18)
                .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                  .addComponent(jtfMarkRange)
                  .addComponent(jtfNameRange)
                  .addComponent(jtfFinalMarkRange))
                .addGap(245, 245, 245))
              .addGroup(layout.createSequentialGroup()
                .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                  .addComponent(jtfSubjectLine, javax.swing.GroupLayout.PREFERRED_SIZE, 485, javax.swing.GroupLayout.PREFERRED_SIZE)
                  .addComponent(jLabel2)
                  .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                    .addGroup(layout.createSequentialGroup()
                      .addComponent(jLabel3)
                      .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
                      .addComponent(jcbAssignmentNumber, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE))
                    .addComponent(jbChangeDir)
                    .addGroup(layout.createSequentialGroup()
                      .addComponent(jLabel10, javax.swing.GroupLayout.PREFERRED_SIZE, 57, javax.swing.GroupLayout.PREFERRED_SIZE)
                      .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
                      .addComponent(jcbSheetNumber, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE))))
                .addGap(0, 26, Short.MAX_VALUE)))
            .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
              .addComponent(jLabel6)
              .addComponent(jScrollPane2, javax.swing.GroupLayout.PREFERRED_SIZE, 465, javax.swing.GroupLayout.PREFERRED_SIZE)
              .addGroup(layout.createSequentialGroup()
                .addComponent(jbSendMarks)
                .addGap(18, 18, 18)
                .addComponent(jcbDummyRun)))
            .addGap(39, 39, 39)))
        .addContainerGap())
      .addGroup(layout.createSequentialGroup()
        .addContainerGap()
        .addComponent(jbScrapeGrades)
        .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
        .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
          .addComponent(jLabel4)
          .addComponent(jcbColumn, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE))
        .addGap(18, 18, 18)
        .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
          .addComponent(jspinRow, javax.swing.GroupLayout.PREFERRED_SIZE, 54, javax.swing.GroupLayout.PREFERRED_SIZE)
          .addComponent(jLabel5))
        .addContainerGap(javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE))
    );
    layout.setVerticalGroup(
      layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
      .addGroup(layout.createSequentialGroup()
        .addGap(26, 26, 26)
        .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.BASELINE)
          .addComponent(jLabel2)
          .addComponent(jLabel6))
        .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.UNRELATED)
        .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.TRAILING)
          .addGroup(layout.createSequentialGroup()
            .addComponent(jScrollPane2, javax.swing.GroupLayout.PREFERRED_SIZE, 200, javax.swing.GroupLayout.PREFERRED_SIZE)
            .addGap(27, 27, 27)
            .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.BASELINE)
              .addComponent(jbSendMarks)
              .addComponent(jcbDummyRun)))
          .addGroup(layout.createSequentialGroup()
            .addComponent(jtfSubjectLine, javax.swing.GroupLayout.PREFERRED_SIZE, 32, javax.swing.GroupLayout.PREFERRED_SIZE)
            .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.UNRELATED)
            .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.BASELINE)
              .addComponent(jLabel3)
              .addComponent(jcbAssignmentNumber, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE))
            .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
            .addComponent(jbChangeDir)
            .addGap(12, 12, 12)
            .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.BASELINE)
              .addComponent(jLabel10)
              .addComponent(jcbSheetNumber, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE))
            .addGap(18, 18, 18)
            .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.BASELINE)
              .addComponent(jLabel1, javax.swing.GroupLayout.PREFERRED_SIZE, 16, javax.swing.GroupLayout.PREFERRED_SIZE)
              .addComponent(jtfNameRange, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE))
            .addGap(11, 11, 11)
            .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.BASELINE)
              .addComponent(jLabel8, javax.swing.GroupLayout.PREFERRED_SIZE, 16, javax.swing.GroupLayout.PREFERRED_SIZE)
              .addComponent(jtfMarkRange, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE))
            .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.UNRELATED)
            .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.BASELINE)
              .addComponent(jLabel9)
              .addComponent(jtfFinalMarkRange, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE))
            .addGap(6, 6, 6)))
        .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED, 8, Short.MAX_VALUE)
        .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.TRAILING)
          .addComponent(jbScrapeGrades)
          .addGroup(layout.createSequentialGroup()
            .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.BASELINE)
              .addComponent(jLabel4)
              .addComponent(jLabel5))
            .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
            .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.BASELINE)
              .addComponent(jcbColumn, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE)
              .addComponent(jspinRow, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE))))
        .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED, 15, Short.MAX_VALUE)
        .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING, false)
          .addComponent(jLabel7, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE)
          .addComponent(jlabCurrentDirectory, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE))
        .addContainerGap())
    );

    pack();
  }// </editor-fold>//GEN-END:initComponents

  private void jMenuItemCopyActionPerformed(java.awt.event.ActionEvent evt) {                                              
    String strFormatted = jtaActivityLog.getText();
    Clipboard clipboard = Toolkit.getDefaultToolkit().getSystemClipboard();
    clipboard.setContents(new StringSelection(strFormatted), null);
  }                                             

  private void jPMICopyActionPerformed(ActionEvent evt) {
    String strFormatted = jtaActivityLog.getText();
    Clipboard clipboard = Toolkit.getDefaultToolkit().getSystemClipboard();
    clipboard.setContents(new StringSelection(strFormatted), null);
  }
  
  private void jbChangeDirActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_jbChangeDirActionPerformed
    // TODO add your handling code here:
    JFileChooser fc = new JFileChooser();
    fc.setDialogTitle("Choose a case directory...");
    fc.setFileSelectionMode(JFileChooser.DIRECTORIES_ONLY);
    fc.setAcceptAllFileFilterUsed(false);

    int nResult = fc.showOpenDialog(this);

    if (nResult == JFileChooser.APPROVE_OPTION) {
      File selectedFile = fc.getSelectedFile();
      String strPath = selectedFile.getAbsolutePath();
      jlabCurrentDirectory.setText(strPath);
      m_strCWD = strPath;
    }
  }//GEN-LAST:event_jbChangeDirActionPerformed

  private void jbSendMarksActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_jbSendMarksActionPerformed
    // TODO add your handling code here:
    File f = new File(m_strCWD);
    
    FilenameFilter xlsxFilter = new FilenameFilter() {
      @Override
			public boolean accept(File dir, String name) {
				return name.toLowerCase().endsWith(".xlsx");
			}
		};
    
    ArrayList<File> Files = new ArrayList<File>(Arrays.asList(f.listFiles(xlsxFilter)));
    String strLog = "";
    
    for(File f1 : Files){
      Pattern p = Pattern.compile("^(?<netid>[A-Za-z]{3,4}[0-9]{3})\\.xlsx");
      Matcher m = p.matcher(f1.getName());
      
      if(m.find()){
        
        InputStream excelFileToRead;
        XSSFWorkbook workbook = null;
        
        try {
          excelFileToRead = new FileInputStream(f1.getAbsolutePath());
          workbook = new XSSFWorkbook(excelFileToRead);
        } catch (FileNotFoundException ex) {
          ex.printStackTrace();
          Logger.getLogger(SendMarksGUI.class.getName()).log(Level.SEVERE, null, ex);
        } catch (IOException ex){
          Logger.getLogger(SendMarksGUI.class.getName()).log(Level.SEVERE, null, ex);
        }
        
        
        
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
        
        String strSheetName = "Sheet" + jcbSheetNumber.getSelectedItem();
        String strNameRange = strSheetName + "!" + jtfNameRange.getText();
        String strMarkRange = strSheetName + "!" + jtfMarkRange.getText();
        String strFinalMarkRange = strSheetName + "!" + jtfFinalMarkRange.getText();

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
        double finalMark = 0;
        double percentMark = 0;
        
        for(int i = 0; i < 2; i++){
          sheet = workbook.getSheet(crefs[i].getSheetName());
          r = sheet.getRow(crefs[i].getRow());
          c = r.getCell(crefs[i].getCol());
          
          if(i == 0){
            totalMark = c.getNumericCellValue();
          }else{
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
  
        for(int row = 0; row  < numRows; row++){
          
          markTable[row] = new Object[numCols];
          
          for(int col = 0; col < numCols; col++){
            int idx = row * numCols + col;
           
            CellReference cr = crefs[idx];
            sheet = workbook.getSheet(cr.getSheetName());
            r = sheet.getRow(cr.getRow());
            c = r.getCell(cr.getCol());
            
            if(c != null){
              int cellType = c.getCellType();
              
              switch(cellType){
                case XSSFCell.CELL_TYPE_STRING:
                  markTable[row][col] = c.getStringCellValue();
                  break;
                case XSSFCell.CELL_TYPE_NUMERIC:
                case XSSFCell.CELL_TYPE_FORMULA:
                  double d = c.getNumericCellValue();
                  if(d - Math.floor(d) < 0.01)
                    markTable[row][col] = (int)d;
                  else
                    markTable[row][col] = d;
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
        
        
        Object[] studentData = {(String)jcbAssignmentNumber.getSelectedItem(),
                                strStudentName, strStudentUPI, 
                                finalMark, totalMark, percentMark};
        readStudentDetails();
        strStudentDetails = completeStudentDetails(strStudentDetails, studentData);
        HTMLMarksheet ms = new HTMLMarksheet((String)jcbAssignmentNumber.getSelectedItem(), 
                                             strCSS, strStudentDetails,
                                             markTable);
        
       
        try{
          ms.write(m_strCWD + "/test/" + strStudentUPI + ".html");
        }catch(IOException ex){
          ex.printStackTrace();
        }
        
        String strNetID = m.group("netid");
        String strEmail = strNetID + "@aucklanduni.ac.nz";
        if(!bDummyRun){
          strLog += "Sending mail to " + strEmail + "\n";  
          ms.send(jtfSubjectLine.getText(), strEmail, false);
        }else{
          strLog += "Pretending to send mail to " + strEmail + "\n";
        }
        jtaActivityLog.setText(strLog); 
        
        
      }
    }
    
  }//GEN-LAST:event_jbSendMarksActionPerformed

  private void jbScrapeGradesActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_jbScrapeGradesActionPerformed
    // TODO add your handling code here:
  }//GEN-LAST:event_jbScrapeGradesActionPerformed

  private void jcbAssignmentNumberActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_jcbAssignmentNumberActionPerformed
    StringBuilder sb = new StringBuilder();
    
    sb.append("STATS 779 Assignment ").append(jcbAssignmentNumber.getSelectedItem()).append(" Marks");
    jtfSubjectLine.setText(sb.toString());
  }//GEN-LAST:event_jcbAssignmentNumberActionPerformed

  private void jcbDummyRunActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_jcbDummyRunActionPerformed
    bDummyRun = jcbDummyRun.isSelected();
  }//GEN-LAST:event_jcbDummyRunActionPerformed

  private void jtfNameRangeActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_jtfNameRangeActionPerformed
    // TODO add your handling code here:
  }//GEN-LAST:event_jtfNameRangeActionPerformed

  private void jMenuItemConfigureActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_jMenuItemConfigureActionPerformed
    
    try{
      MailConfigDialog dlg = new MailConfigDialog(this, true);
      dlg.show(true);
      
      if(dlg.isSaved()){
        
      }
    }catch(Exception e){
      e.printStackTrace();
    }
    
  }//GEN-LAST:event_jMenuItemConfigureActionPerformed

  /**
   * @param args the command line arguments
   */
  public static void main(String args[]) {
    /* Set the Nimbus look and feel */
    //<editor-fold defaultstate="collapsed" desc=" Look and feel setting code (optional) ">
    /* If Nimbus (introduced in Java SE 6) is not available, stay with the default look and feel.
         * For details see http://download.oracle.com/javase/tutorial/uiswing/lookandfeel/plaf.html 
     */
    try {
      for (javax.swing.UIManager.LookAndFeelInfo info : javax.swing.UIManager.getInstalledLookAndFeels()) {
        if ("Nimbus".equals(info.getName())) {
          javax.swing.UIManager.setLookAndFeel(info.getClassName());
          break;
        }
      }
    } catch (ClassNotFoundException ex) {
      java.util.logging.Logger.getLogger(SendMarksGUI.class.getName()).log(java.util.logging.Level.SEVERE, null, ex);
    } catch (InstantiationException ex) {
      java.util.logging.Logger.getLogger(SendMarksGUI.class.getName()).log(java.util.logging.Level.SEVERE, null, ex);
    } catch (IllegalAccessException ex) {
      java.util.logging.Logger.getLogger(SendMarksGUI.class.getName()).log(java.util.logging.Level.SEVERE, null, ex);
    } catch (javax.swing.UnsupportedLookAndFeelException ex) {
      java.util.logging.Logger.getLogger(SendMarksGUI.class.getName()).log(java.util.logging.Level.SEVERE, null, ex);
    }
    //</editor-fold>

    /* Create and display the form */
    java.awt.EventQueue.invokeLater(new Runnable() {
      public void run() {
        new SendMarksGUI().setVisible(true);
      }
    });
  }
  
  private boolean bDummyRun;
  
  private String m_strCWD;
  private String strCSS;
  private String strStudentDetails;
  
  private javax.swing.JPopupMenu jPopupMenuEdit;
  private javax.swing.JMenuItem jPMIClear;
  private javax.swing.JMenuItem jPMICopy;
  private javax.swing.JMenuItem jPMIPaste;

  // Variables declaration - do not modify//GEN-BEGIN:variables
  private javax.swing.JLabel jLabel1;
  private javax.swing.JLabel jLabel10;
  private javax.swing.JLabel jLabel2;
  private javax.swing.JLabel jLabel3;
  private javax.swing.JLabel jLabel4;
  private javax.swing.JLabel jLabel5;
  private javax.swing.JLabel jLabel6;
  private javax.swing.JLabel jLabel7;
  private javax.swing.JLabel jLabel8;
  private javax.swing.JLabel jLabel9;
  private javax.swing.JMenu jMenu2;
  private javax.swing.JMenuBar jMenuBar1;
  private javax.swing.JMenuItem jMenuItemConfigure;
  private javax.swing.JScrollPane jScrollPane2;
  private javax.swing.JButton jbChangeDir;
  private javax.swing.JButton jbScrapeGrades;
  private javax.swing.JButton jbSendMarks;
  private javax.swing.JComboBox<String> jcbAssignmentNumber;
  private javax.swing.JComboBox<String> jcbColumn;
  private javax.swing.JCheckBox jcbDummyRun;
  private javax.swing.JComboBox<String> jcbSheetNumber;
  private javax.swing.JLabel jlabCurrentDirectory;
  private javax.swing.JMenu jmenuSetup;
  private javax.swing.JSpinner jspinRow;
  private javax.swing.JTextArea jtaActivityLog;
  private javax.swing.JTextField jtfFinalMarkRange;
  private javax.swing.JTextField jtfMarkRange;
  private javax.swing.JTextField jtfNameRange;
  private javax.swing.JTextField jtfSubjectLine;
  // End of variables declaration//GEN-END:variables
}
