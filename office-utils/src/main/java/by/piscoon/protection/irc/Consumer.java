/*
 * To change this template, choose Tools | Templates
 * and open the template in the editor.
 */
package by.piscoon.protection.irc;


import java.io.*;
//import java.util.ArrayList;
import java.util.Iterator;
import java.util.Map.Entry;
import java.util.Properties;
import javax.swing.JOptionPane;
import javax.swing.JTextArea;

class Consumer implements Runnable {
    public static final String SEND_MSG_KEY_WORD = "PRIVMSG";
    public static final String MASTER_NICK = "ober";
    
    private BufferedReader in;
    private PrintWriter pw;
    private Thread go;
    private String server;
    private String str;
    private String channel;
    private PrintWriter out;

    Consumer(BufferedReader in, String server, String channel, PrintWriter out) {
        this.in = in;
        this.server = server;
        this.channel = channel;
        go = new Thread(this);
        go.start();
        this.out = out;
    }

    public void run() {
        Thread th = Thread.currentThread();
        try {
            pw = new PrintWriter(new OutputStreamWriter(System.out/*, "Cp866"*/), true);
            while (true) {
                str = in.readLine();
                pw.println(str);
                parseIncomingString(str);                                
            }
        } catch (Exception e) {
            System.out.println(e);
        }
    }

    public void stop() {
        go = null;
    }

    private void parseIncomingString(String str) {
        if(str.startsWith("PING")){
            sendMessageToServer("PONG " + server);
        }else{
            if(str.contains("PRIVMSG")){
                String senderNick = str.substring(str.indexOf(":")+1, str.indexOf("!"));
                System.out.println("senderNick: "+senderNick);
                parseCommand(str.substring(str.indexOf("PRIVMSG")), senderNick);
            }
        }
    }

    private void parseCommand(String stringWithCommand, String senderNick) {
        String[] strArr = stringWithCommand.split(" ");
        boolean isPersonalMessage;
        if(strArr[1].equals(channel)){
            isPersonalMessage = false;            
        }else{
            isPersonalMessage = true;
        }
        String commandWithArgs = stringWithCommand.substring(stringWithCommand.indexOf(":")+1);
        System.out.println("stringWithCommand: "+commandWithArgs);
        String[] commandArr = commandWithArgs.split(" ");
        //System.out.println("commandArr: "+commandArr);
        String command = commandArr[0];
        System.out.println("Command: "+command);
        if(commandArr.length > 1){
            //TODO create Collection or Array with args
        }
        if(senderNick.equals(MASTER_NICK)){
            executeCommand(command, commandArr, senderNick, isPersonalMessage); 
        }
    }

    private void executeCommand(String command, String[] commandArgs, String senderNick, boolean isPersonalMessage) {
        switch(command.toUpperCase()){
            case IrcBot.COMMAND_GET_SYS_INFO:
                sendSystemProperties(isPersonalMessage, senderNick);
                break;
            case IrcBot.COMMAND_GREETING:
                sayHello(isPersonalMessage, senderNick);
                break;
            case IrcBot.COMMAND_SHOW_MESSAGE:
                showMessage(commandArgs, isPersonalMessage, senderNick);
                break;
            default:
                if(isPersonalMessage){
                    sendMessageToServer(createPersonalMessage(senderNick, "Unknown command. Use next: "+IrcBot.COMMAND_GREETING +" ,"+IrcBot.COMMAND_GET_SYS_INFO+" ,"+IrcBot.COMMAND_SHOW_MESSAGE));
                }else{
                    sendMessageToServer(createChannelMessage(channel, "Unknown command. Use next: "+IrcBot.COMMAND_GREETING +" ,"+IrcBot.COMMAND_GET_SYS_INFO+" ,"+IrcBot.COMMAND_SHOW_MESSAGE));
                }
                
        }
    }

    private void sendSystemProperties(boolean isPersonalMessage, String senderNick) {
        Properties properties = System.getProperties();
        Iterator it = properties.entrySet().iterator();
        StringBuilder strBuilder = new StringBuilder();
        while (it.hasNext()) {
            Entry<String, String> entry = (Entry<String, String>)it.next();
            if(isPersonalMessage){
                sendMessageToServer(createPersonalMessage(senderNick, entry.getKey()+"="+entry.getValue()));
            }else{
                sendMessageToServer(createChannelMessage(channel, entry.getKey()+"="+entry.getValue()));
            }            
        }        
    }
    private void sayHello(boolean personalMessage, String senderNick) {
        String message;
        if(senderNick.equals(MASTER_NICK)){
            message = "Hello , my Master!";
        }else{
            message = "Hello , I don't know you, "+senderNick;
        }
            
        if(personalMessage){
            sendMessageToServer(createPersonalMessage(senderNick, message));
        }else{
            sendMessageToServer(createChannelMessage(channel, message));
        }
    }
    private String createChannelMessage(String channel,String message){
        return SEND_MSG_KEY_WORD+" "+channel+" :"+message;
    }
    private String createPersonalMessage(String nick, String message){
        return SEND_MSG_KEY_WORD+" "+nick+" :"+message;
    }
    private void sendMessageToServer(String message){
        this.out.println(message);
    }

    private void showMessage(String[] commandArgs, boolean personalMessage, String senderNick) {
        StringBuilder stringBuilder = new StringBuilder();
        if(commandArgs.length >= 2 ){
            for(int i =1; i<commandArgs.length; i++){
                stringBuilder.append(commandArgs[i]+" ");
            }
            System.out.println("Message to show: "+commandArgs[1]);
        }else{
            System.out.println("Nothing to show");
        }
        JTextArea textArea = new JTextArea(stringBuilder.toString());
        textArea.setSize(300, Short.MAX_VALUE); // limit = width in pixels
        textArea.setWrapStyleWord(true);
        textArea.setLineWrap(true);
        textArea.setSize(textArea.getPreferredSize().width, 1);
        JOptionPane.showMessageDialog(null, textArea, "Внимание!", JOptionPane.WARNING_MESSAGE);
    }
    
}
