/*
 * To change this template, choose Tools | Templates
 * and open the template in the editor.
 */
package by.piscoon.protection.irc;

import java.net.*;
import java.io.*;
import java.util.logging.Level;
import java.util.logging.Logger;

public class IrcBot implements Runnable{
    public static final String COMMAND_GET_SYS_INFO = "GET_SYS_INFO";
    public static final String COMMAND_GREETING = "HELLO";
    public static final String COMMAND_SHOW_MESSAGE = "SHOW_MESSAGE";
    public static final String DEFAULT_NAME = "java_bot";
    public static final String DEFAULT_HOST = "java_host";
    
    private Socket socket;
    private PrintWriter out;
    private BufferedReader in;
    
    private String server;
    private int port;
    private String nick;
    private String channel;
    
    public IrcBot(String server, int port, String nick, String channel){
        this.server = server;
        this.port = port;
        this.channel = channel;
        this.nick = nick;
    }
    public static void main(String[] args) {
        Thread ircBotThread = new Thread(new IrcBot("irc.by", 6667, "bot", "#la2"));
        ircBotThread.start();        
        Thread ircBotThread1 = new Thread(new IrcBot("irc.by", 6667, "bot", "#la2"));
        ircBotThread1.start();   
        Thread ircBotThread2 = new Thread(new IrcBot("irc.by", 6667, "bot", "#la2"));
        ircBotThread2.start();   
        try {
            Thread.sleep(0);
        } catch (InterruptedException ex) {
            ex.printStackTrace();
        }
    }
    
    public void send_string(String str) {
        out.println(str);
    }    

    @Override
    public void run() {
        try {
            socket = new Socket(server, port);
            in = new BufferedReader(new InputStreamReader(socket.getInputStream()));
            out = new PrintWriter(new BufferedWriter(new OutputStreamWriter(socket.getOutputStream())), true);
            int counter = 0;
            int connectReturnedValue;
            do{
                connectReturnedValue = connect(counter);
                if(connectReturnedValue == -2){
                    counter++;
                }                
                Thread.sleep(1000);
            }while(connectReturnedValue!=0 && counter < Integer.MAX_VALUE);
            System.out.println("Try to join channel");
            out.write("JOIN " + channel + "\r\n");
            out.flush();


            Producer p = new Producer(out);
            Consumer c = new Consumer(in, server, channel, out);
        } catch (Exception e) {
            System.out.println(e);
        }
        try {
            Thread.sleep(0);
        } catch (InterruptedException ie) {
            ie.printStackTrace();
        }          
    }
    private int connect(int count) throws IOException{
        
        out.println("NICK " + nick+count);
        out.println("USER " + DEFAULT_HOST + " \"...\" \"...\" " + DEFAULT_NAME);
        out.flush();
        String line = null;
        while ((line = in.readLine()) != null) {
            if (line.indexOf("004") >= 0) {
                // We are now logged in.
                return 0;
            } else if (line.indexOf("433") >= 0) {
                System.out.println("Nickname is already in use.");               
                return -2;
            }
            System.out.println(line);               
        }  
        System.out.println("Server is not responding.");
        return -1;
    }
}
