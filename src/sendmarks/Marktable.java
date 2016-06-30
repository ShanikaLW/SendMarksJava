/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package sendmarks;

import java.util.ArrayList;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

/**
 *
 * @author James M. Curran <j.curran@auckland.ac.nz>
 */
public class Marktable {
  protected class Question {
    protected class QuestionRow{
      String strMarkItem;
      int availableMarks;
      double actualMarks;
      String strComments;
      boolean isSubtotal;
      
      protected QuestionRow(Object[] row, boolean isSubT){
        strMarkItem = (String)row[1];
        availableMarks = (int)row[2];

        if(row[3] instanceof Integer){
          actualMarks = 1.0 * (Integer)row[3];
        }else if(row[3] instanceof Double){
          actualMarks = (Double)row[3];
        }else{
          actualMarks = -1;
        }
        strComments = (String)row[4];
        isSubtotal = isSubT;
        
      }
      
      protected QuestionRow(Object[] row){
        this(row, false);
      }
      
      
      String toHTMLTableRow(){
        StringBuilder sb = new StringBuilder();
       
        if(isSubtotal)
          sb.append("\t<tr class=\"total\">\n");
        else
          sb.append("\t<tr>\n");
        
        sb.append("\t\t<td>");
        sb.append(strMarkItem);
        sb.append("</td>\n");
        
        sb.append("\t\t<td>");
        sb.append(availableMarks);
        sb.append("</td>\n");
          
        sb.append("\t\t<td>");
        if(actualMarks - Math.floor(actualMarks) < 0.01)
          sb.append((int)actualMarks);
        else
          sb.append(actualMarks);
        sb.append("\t\t</td>\n");
        
        sb.append("\t\t<td>");
        if(strComments != null)
          sb.append(strComments);
        sb.append("</td>\n");
        sb.append("\t</tr>\n");
       
        return sb.toString();
      }
    }
    
    final String tagStartHeading = "<tr class = \"heading\" ><td colspan=\"4\">Question ";
    final String tagEndHeading = "</td></tr>\n";
    final String tagQHeader = "<tr class = \"qheader\">\n" +
                              "\t<td>Item</td>\n" +
                              "\t<td>Mark</td>\n" +
                              "\t<td>Your Mark</td>\n" +
                              "\t<td class = \"qcomment\">Comments</td>\n" +
                              "</tr>";
    final String tagQuestionRow = "\t<tr class=\"\">\n";
    final String tagEndRow = "\t</tr>\n";
      
    protected int questionNumber;
    protected int numRows;
    protected ArrayList<QuestionRow> rows;
    
    
    protected Question(int number){
      questionNumber = number;
      numRows = 0;
      rows = new ArrayList<>();
    }
    
    protected void addRow(Object[] row){
      rows.add(new QuestionRow(row));
      numRows++;
    }
    
    protected void addRow(Object[] row, boolean bIsSubtotal){
      rows.add(new QuestionRow(row, bIsSubtotal));
      numRows++;
    }
    
    @Override
    public String toString(){
      StringBuilder sb = new StringBuilder();
      
      sb.append(tagStartHeading);
      sb.append(questionNumber);
      sb.append(tagEndHeading);
      sb.append(tagQHeader);

      
      for(QuestionRow row : rows){
        sb.append(row.toHTMLTableRow());
      }
      
      return sb.toString();
    }
  }
  
  private ArrayList<int[]> findQuestions(Object[][] markTable){
    ArrayList<int[]> questionPos = new ArrayList<>();
    
    int questionCtr = 0;
    int start = 0;
    
    for(int i = 0; i < markTable.length; i++){
      String entry = String.valueOf(markTable[i][0]);
      Pattern p = Pattern.compile("([1-9]|1[0-9]+)");
      Matcher m = p.matcher(entry);
      
      if(m.find()){
        int[] qPair = new int[2];
        qPair[0] = i;
        questionCtr++;
        
        questionPos.add(qPair);
      }
    }
    
    for(int i = 0; i < questionCtr - 1; i++){
      int[] qPair1 = questionPos.get(i);
      int[] qPair2 = questionPos.get(i + 1);
      
      qPair1[1] = qPair2[0] - 1;
    }
    
    int[] qPair = questionPos.get(questionCtr - 1);
    qPair[1] = markTable.length - 1;
    
    return questionPos;
  }
  
  private Question[] questions;
  private int numQuestions;
  
  public Marktable(Object[][] markTableData){
    ArrayList<int[]> questionPos = findQuestions(markTableData);
    numQuestions = questionPos.size();
    questions = new Question[numQuestions];
    
    int questionNumber = 0;
    
    for(int[] qPair : questionPos){
      Question q = new Question(questionNumber + 1);
      
      int start = qPair[0];
      int end = qPair[1];
      
      for(int i = start; i <= end; i++){
        if(i != end)
          q.addRow(markTableData[i]);
        else
          q.addRow(markTableData[i], true); // NOTE: always assumes last row is subtotal
                                            // Is this reasonable?
      }
      questions[questionNumber++] = q;
    }
  }
  
  public String toString(){
    StringBuilder sb = new StringBuilder();
    
    for(int i = 0; i < numQuestions; i++){
      Question q = questions[i];
      
      sb.append(q.toString());
      
    }
    
    return sb.toString();
  }
}
