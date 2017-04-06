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
public class MailConfig {
  private String from;
  private String username;
  private String password;
  private String host;
  private int port;
  private boolean useTTLS;
  
  MailConfig(){
    port = 25;
    useTTLS = false;
  }
}
