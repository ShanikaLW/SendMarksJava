/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package sendmarks;

/**
 *
 * @author James M. Curran <j.curran@auckland.ac.nz>
 */
public class SheetInfo {
  private final String strSheetName;
  private final String strNameRange;
  private final String strMarkRange;
  private final String strFinalMarkRange;
  private final String strAssignmentNumber;
  private final String strCourse;
  
  public SheetInfo(String sheetName, String nameRange, String markRange, String finalMarkRange, String assignmentNumber, String course){
    strSheetName = sheetName;
    strNameRange = strSheetName + "!" + nameRange;
    strMarkRange = strSheetName + "!" + markRange;
    strFinalMarkRange = strSheetName + "!" + finalMarkRange;
    strAssignmentNumber = assignmentNumber;
    strCourse = course;
  }
  
  public String getSheetName(){
    return strSheetName;
  }
  
  public String getNameRange(){
    return strNameRange;
  }
  
  public String getMarkRange(){
    return strMarkRange;
  }
  public String getFinalMarkRange(){
    return strFinalMarkRange;
  }
  
  public String getAssignmentNumber(){
    return strAssignmentNumber;
  } 
  
  public String getCourse(){
    return strCourse;
  }
}
