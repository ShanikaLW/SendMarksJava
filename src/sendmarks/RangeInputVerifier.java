/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package sendmarks;

import java.util.regex.Matcher;
import java.util.regex.Pattern;
import javax.swing.InputVerifier;
import javax.swing.JComponent;
import javax.swing.JFrame;
import javax.swing.JOptionPane;
import javax.swing.JPanel;
import javax.swing.JTextField;

/**
 *
 * @author James M. Curran <j.curran@auckland.ac.nz>
 */
public class RangeInputVerifier extends InputVerifier{

  int rowLettersToNum(String strRow){
  
    int result;
    
    if(strRow.length() == 2){
      result = (strRow.charAt(0) - 'A' + 1) * 26  + (strRow.charAt(1) - 'A' + 1);
    }else{
      result = strRow.charAt(0) - 'A' + 1;
    }
    
    return result;
  }
  
  @Override
  public boolean verify(JComponent input) {
    try{
      Pattern p = Pattern.compile( 	"^\\$([A-Z]{1,2})\\$([0-9]{1,2}):\\$([A-Z]{1,2})\\$([0-9]{1,2})$");
      String strRangeTxt = ((JTextField)input).getText().toUpperCase();

      Matcher m = p.matcher(strRangeTxt);
      if(m.find()){
        
        int rowOne = rowLettersToNum(m.group(1));
        int colOne = Integer.parseInt(m.group(2));
        int rowTwo = rowLettersToNum(m.group(3));
        int colTwo = Integer.parseInt(m.group(4));
              
        if(colOne > colTwo || rowOne > rowTwo){
          StringBuilder sb = new StringBuilder();
          
          sb.append(strRangeTxt).append(" is not a valid cell range\n");
          if(colOne > colTwo){
            sb.append("The first column to the right of the second column number\n");
          }
          if(rowOne > rowTwo){
            sb.append("The first row must be above the row\n");
          }
          throw new SheetRangeException(sb.toString());
        }else{
          ((JTextField)input).setText(strRangeTxt); // make sure it's upper case
          return true;
        }
      }else{
        throw new SheetRangeException(strRangeTxt + " is not in the $R1$C1:$R2$C2 format");
      }
    }catch(SheetRangeException e){
      JPanel parent = (JPanel)input.getParent();
      JOptionPane.showMessageDialog(parent, e.getMessage());
      
      return false;
     }
  } 
}
